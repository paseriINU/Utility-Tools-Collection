#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
================================================================================
WinRM Remote Batch Executor for Linux (Python版 - 標準ライブラリのみ)
================================================================================

【概要】
このプログラムは、LinuxからWindows ServerへWinRM (Windows Remote Management)
プロトコルを使用してリモート接続し、バッチファイルを実行するツールです。

【特徴】
- 標準ライブラリのみ使用（pip install不要）
- NTLM v2認証を自前実装（MD4を含む）
- IT制限環境でも動作可能（外部パッケージ不要）
- Windows側の設定変更不要（デフォルトのNTLM認証を使用）

【なぜ標準ライブラリのみで実装するのか】
企業のIT制限環境では、外部ライブラリのインストールが禁止されていることが多い。
このプログラムは、そのような環境でも確実に動作するよう設計されている。
- requests, pywinrm等の外部パッケージは使用しない
- MD4はPython標準ライブラリに含まれていないため自前実装

【NTLM認証の仕組み】
NTLM認証は3段階のハンドシェイクで構成される：
  1. Type 1 (Negotiate): クライアントが認証開始を宣言
  2. Type 2 (Challenge): サーバーがランダムなチャレンジを返送
  3. Type 3 (Authenticate): クライアントがチャレンジに対する応答を送信

【WinRMプロトコルの流れ】
  1. シェル作成 (Create): リモートシェルセッションを開始
  2. コマンド実行 (Command): バッチファイルを実行
  3. 出力取得 (Receive): 標準出力・標準エラーを取得
  4. シェル削除 (Delete): セッションをクリーンアップ

【必要な環境】
- Linux（Red Hat, CentOS, Ubuntu等）
- Python 3.6以降
- ネットワーク接続（ポート5985/HTTP または 5986/HTTPS）

【使い方】
  1. このソースファイル内の設定セクションを編集
  2. 実行: python3 winrm_exec.py ENV

  環境を引数で指定（必須）:
    python3 winrm_exec.py TST1T
    python3 winrm_exec.py TST2T

  コマンドライン引数で詳細設定も指定可能:
    python3 winrm_exec.py TST1T --host 192.168.1.100 --user Administrator --password Pass123

【セキュリティに関する注意】
- パスワードはソースコード内に記載するため、適切なファイル権限を設定すること
- 本番環境ではコマンドライン引数での上書きを推奨

================================================================================
"""

# ==============================================================================
# インポート（すべて標準ライブラリ - 外部パッケージ不要）
# ==============================================================================

import sys              # システム関連: コマンドライン引数、終了コード
import argparse         # コマンドライン引数のパース
import logging          # ログ出力
import base64           # Base64エンコード/デコード（NTLMメッセージ用）
import uuid             # UUID生成（SOAPメッセージID用）
import socket           # ソケット通信（HTTP接続用）
import struct           # バイナリデータのパック/アンパック
import time             # 時間関連（タイムスタンプ、スリープ）
import hashlib          # MD5ハッシュ（HMAC-MD5用）
import hmac             # HMAC計算（NTLMv2認証用）
import os               # OS機能（乱数生成: os.urandom）
from xml.etree import ElementTree as ET  # XML解析（SOAPレスポンス用）

# ==============================================================================
# 設定セクション（ユーザー編集エリア）
# ==============================================================================
#
# 【使用方法】
# 1. 以下の設定値を環境に合わせて編集してください
# 2. 実行: python3 winrm_exec.py TST1T
#
# 【設定の優先順位】
# コマンドライン引数 > ソースコード内の設定
# 例: python3 winrm_exec.py TST1T --host 10.0.0.1
#     → コマンドライン引数の値が優先される
# ==============================================================================

# --- Windows接続情報 ---
WINDOWS_HOST = "192.168.1.100"      # Windows ServerのIPアドレスまたはホスト名
WINDOWS_PORT = 5985                  # WinRMポート: HTTP=5985, HTTPS=5986
WINDOWS_USER = "Administrator"       # Windowsのログインユーザー名
WINDOWS_PASSWORD = "YourPassword"    # Windowsのログインパスワード
WINDOWS_DOMAIN = ""                  # ドメイン名（空文字列 = ローカル認証）

# --- 利用可能な環境のリスト ---
# コマンドライン引数で指定可能な環境名を定義
# 新しい環境を追加する場合はこのリストに追加してください
ENVIRONMENTS = ["TST1T", "TST2T"]

# --- 実行するバッチファイル ---
# {ENV} プレースホルダは実行時に環境名（TST1T等）に置換されます
# 例: "C:\Scripts\{ENV}\test.bat" → "C:\Scripts\TST1T\test.bat"
# 注: {ENV}は複数箇所に使用可能（すべて置換される）
BATCH_FILE_PATH = r"C:\Scripts\{ENV}\test.bat"

# --- 直接コマンド実行 ---
# バッチファイルではなく直接コマンドを実行する場合に設定
# Noneの場合はBATCH_FILE_PATHが使用される
DIRECT_COMMAND = None  # 例: "echo Hello from WinRM"

# --- タイムアウト設定 ---
# コマンド実行の最大待機時間（秒）
# バッチ処理が長時間かかる場合は増やしてください
TIMEOUT = 300

# --- ログレベル ---
# DEBUG: 詳細なデバッグ情報（トラブルシューティング用）
# INFO:  通常の動作情報（推奨）
# WARNING: 警告のみ
# ERROR: エラーのみ
LOG_LEVEL = "INFO"

# ==============================================================================

# ==============================================================================
# XML名前空間の定義
# ==============================================================================
# WinRMはSOAP/XMLプロトコルを使用するため、名前空間の定義が必要
# これらはWS-Management (WS-Man) 仕様で定義されている
# ==============================================================================
NAMESPACES = {
    's': 'http://www.w3.org/2003/05/soap-envelope',           # SOAPエンベロープ
    'a': 'http://schemas.xmlsoap.org/ws/2004/08/addressing',  # WS-Addressing
    'w': 'http://schemas.dmtf.org/wbem/wsman/1/wsman.xsd',    # WS-Management
    'rsp': 'http://schemas.microsoft.com/wbem/wsman/1/windows/shell',  # WinRS (シェル)
    'p': 'http://schemas.microsoft.com/wbem/wsman/1/wsman.xsd'  # WS-Man (Microsoft)
}


# ==============================================================================
# MD4ハッシュアルゴリズム実装
# ==============================================================================
#
# 【なぜMD4を自前実装するのか】
# - MD4はNTLM認証でパスワードハッシュの計算に必須
# - Python標準ライブラリ(hashlib)にはMD4が含まれていない
# - OpenSSLのMD4は多くのディストリビューションで非推奨/削除済み
# - 外部ライブラリに依存せずIT制限環境で動作させるため
#
# 【MD4アルゴリズムの概要】
# - 入力: 任意長のデータ
# - 出力: 128ビット（16バイト）のハッシュ値
# - RFC 1320で定義
# - 3ラウンド×16ステップ = 48ステップの処理
#
# 【セキュリティ上の注意】
# MD4は暗号学的に破られており、新規開発には推奨されない。
# ただし、NTLMプロトコルとの互換性のために使用が必要。
# ==============================================================================


class MD4:
    """
    MD4ハッシュアルゴリズムの実装

    NTLM認証でパスワードのNTハッシュを計算するために使用。
    計算式: NT_Hash = MD4(UTF-16LE(Password))
    """

    def __init__(self):
        """
        MD4ハッシュ計算の初期化

        A, B, C, Dは128ビットの状態を表す4つの32ビットレジスタ
        初期値はRFC 1320で定義されている
        """
        self.A = 0x67452301  # 状態レジスタA
        self.B = 0xefcdab89  # 状態レジスタB
        self.C = 0x98badcfe  # 状態レジスタC
        self.D = 0x10325476  # 状態レジスタD
        self.count = 0       # 処理済みバイト数
        self.buffer = b''    # 未処理のデータバッファ

    @staticmethod
    def _left_rotate(x, n):
        """左循環シフト（ローテート）: ビットを左にn個回転"""
        return ((x << n) | (x >> (32 - n))) & 0xffffffff

    @staticmethod
    def _F(x, y, z):
        """Round 1用: 条件選択関数 - xが1のビットはyを、0のビットはzを選択"""
        return (x & y) | (~x & z)

    @staticmethod
    def _G(x, y, z):
        """Round 2用: 多数決関数 - 3つの入力のうち2つ以上が1なら1"""
        return (x & y) | (x & z) | (y & z)

    @staticmethod
    def _H(x, y, z):
        """Round 3用: パリティ関数 - XOR演算"""
        return x ^ y ^ z

    def _process_block(self, block):
        """
        64バイト（512ビット）のブロックを処理

        Args:
            block: 処理する64バイトのデータ

        処理内容:
        1. ブロックを16個の32ビットワードに分割
        2. Round 1: F関数を使用した16ステップ
        3. Round 2: G関数を使用した16ステップ
        4. Round 3: H関数を使用した16ステップ
        5. 結果を状態レジスタに加算
        """
        # リトルエンディアンで16個の32ビット整数としてアンパック
        M = struct.unpack('<16I', block)

        A, B, C, D = self.A, self.B, self.C, self.D

        # Round 1
        for i in range(16):
            k = i
            s = [3, 7, 11, 19][i % 4]
            A = self._left_rotate((A + self._F(B, C, D) + M[k]) & 0xffffffff, s)
            A, B, C, D = D, A, B, C

        # Round 2
        for i in range(16):
            k = [0, 4, 8, 12, 1, 5, 9, 13, 2, 6, 10, 14, 3, 7, 11, 15][i]
            s = [3, 5, 9, 13][i % 4]
            A = self._left_rotate((A + self._G(B, C, D) + M[k] + 0x5a827999) & 0xffffffff, s)
            A, B, C, D = D, A, B, C

        # Round 3
        for i in range(16):
            k = [0, 8, 4, 12, 2, 10, 6, 14, 1, 9, 5, 13, 3, 11, 7, 15][i]
            s = [3, 9, 11, 15][i % 4]
            A = self._left_rotate((A + self._H(B, C, D) + M[k] + 0x6ed9eba1) & 0xffffffff, s)
            A, B, C, D = D, A, B, C

        self.A = (self.A + A) & 0xffffffff
        self.B = (self.B + B) & 0xffffffff
        self.C = (self.C + C) & 0xffffffff
        self.D = (self.D + D) & 0xffffffff

    def update(self, data):
        """
        ハッシュ計算にデータを追加

        Args:
            data: 追加するデータ（bytes or str）

        データをバッファに追加し、64バイト以上になったら処理を実行
        """
        if isinstance(data, str):
            data = data.encode('utf-8')

        self.buffer += data
        self.count += len(data)

        # バッファが64バイト以上あれば処理
        while len(self.buffer) >= 64:
            self._process_block(self.buffer[:64])
            self.buffer = self.buffer[64:]

    def digest(self):
        """
        最終的なハッシュ値を取得

        Returns:
            bytes: 16バイトのMD4ハッシュ値

        処理内容:
        1. パディングの追加（1ビットの1 + 0の列 + 64ビットの長さ）
        2. 残りのブロックを処理
        3. 状態レジスタから結果を生成
        """
        # パディング
        msg = self.buffer
        msg_len = self.count
        msg += b'\x80'
        msg += b'\x00' * ((55 - msg_len) % 64)
        msg += struct.pack('<Q', msg_len * 8)

        # 一時的な状態を保存
        A, B, C, D = self.A, self.B, self.C, self.D

        # 残りのブロックを処理
        for i in range(0, len(msg), 64):
            self._process_block(msg[i:i+64])

        result = struct.pack('<4I', self.A, self.B, self.C, self.D)

        # 状態を復元
        self.A, self.B, self.C, self.D = A, B, C, D

        return result


def md4_hash(data):
    """
    MD4ハッシュを計算するヘルパー関数

    Args:
        data: ハッシュ対象のデータ（bytes or str）

    Returns:
        bytes: 16バイトのMD4ハッシュ値
    """
    hasher = MD4()
    hasher.update(data)
    return hasher.digest()


# ==============================================================================
# NTLM認証実装
# ==============================================================================
#
# 【NTLMv2認証の流れ】
#
# 1. NT Hashの生成
#    NT_Hash = MD4(UTF-16LE(Password))
#
# 2. NTLMv2 Hashの生成
#    NTLMv2_Hash = HMAC-MD5(NT_Hash, UTF-16LE(Username.upper() + Domain))
#
# 3. Type 1メッセージ（Negotiate）の送信
#    - 使用可能な認証方式を通知
#    - サポートするフラグを送信
#
# 4. Type 2メッセージ（Challenge）の受信
#    - サーバーから8バイトのチャレンジを受信
#    - TargetInfo構造体を受信（オプション）
#
# 5. Type 3メッセージ（Authenticate）の送信
#    - NTProofStr = HMAC-MD5(NTLMv2_Hash, ServerChallenge + Blob)
#    - NTResponse = NTProofStr + Blob
#    - Blobにはタイムスタンプ、クライアントチャレンジ等を含む
#
# 【セキュリティ考慮】
# - NTLMv2は、NTLMv1より安全（リプレイ攻撃耐性あり）
# - タイムスタンプにより時間制限付きの認証が可能
# - クライアントチャレンジにより、サーバー側のなりすましを防止
# ==============================================================================


class NTLMAuth:
    """
    NTLMv2認証の実装

    Windows環境でデフォルトで有効なNTLM認証を実装。
    外部ライブラリを使用せず、標準ライブラリのみで動作。
    """

    # NTLMネゴシエーションフラグ
    # クライアントとサーバー間で使用する機能を交渉するためのビットフラグ
    NTLMSSP_NEGOTIATE_UNICODE = 0x00000001           # Unicode文字列を使用
    NTLMSSP_NEGOTIATE_NTLM = 0x00000200              # NTLM認証を使用
    NTLMSSP_NEGOTIATE_ALWAYS_SIGN = 0x00008000       # 常に署名を使用
    NTLMSSP_NEGOTIATE_EXTENDED_SESSIONSECURITY = 0x00080000  # NTLMv2セッションセキュリティ
    NTLMSSP_REQUEST_TARGET = 0x00000004              # ターゲット情報を要求
    NTLMSSP_NEGOTIATE_TARGET_INFO = 0x00800000       # TargetInfo構造体を含む

    def __init__(self, username, password, domain=''):
        """
        NTLMAuthの初期化

        Args:
            username: Windowsユーザー名
            password: パスワード
            domain: ドメイン名（ローカル認証時は空文字列）
        """
        self.username = username
        self.password = password
        self.domain = domain

    @staticmethod
    def _utf16le(s):
        """
        文字列をUTF-16LE（Little Endian）に変換

        Args:
            s: 変換する文字列

        Returns:
            bytes: UTF-16LEエンコードされたバイト列

        WindowsはUTF-16LEを内部文字コードとして使用するため、
        NTLM認証でもこの形式でエンコードする必要がある。
        """
        return s.encode('utf-16-le')

    def _nt_hash(self):
        """
        NTハッシュを計算（パスワードのMD4ハッシュ）

        Returns:
            bytes: 16バイトのNTハッシュ

        計算式: NT_Hash = MD4(UTF-16LE(Password))
        """
        return md4_hash(self._utf16le(self.password))

    def _ntlmv2_hash(self):
        """
        NTLMv2ハッシュを計算

        Returns:
            bytes: 16バイトのNTLMv2ハッシュ

        計算式: NTLMv2_Hash = HMAC-MD5(NT_Hash, UTF-16LE(Username.upper() + Domain))
        注: ユーザー名は大文字に変換されるが、ドメイン名はそのまま使用
        """
        nt_hash = self._nt_hash()
        user_domain = (self.username.upper() + self.domain).encode('utf-16-le')
        return hmac.new(nt_hash, user_domain, hashlib.md5).digest()

    def create_type1_message(self):
        """
        Type 1メッセージ（Negotiate）を生成

        Returns:
            str: Base64エンコードされたType 1メッセージ

        Type 1メッセージ構造:
        - Signature (8バイト): "NTLMSSP\0"
        - MessageType (4バイト): 0x00000001
        - NegotiateFlags (4バイト): サポートする機能フラグ
        - DomainNameFields (8バイト): ドメイン名（空）
        - WorkstationFields (8バイト): ワークステーション名（空）
        """
        # サポートする機能をフラグで指定
        flags = (self.NTLMSSP_NEGOTIATE_UNICODE |
                 self.NTLMSSP_NEGOTIATE_NTLM |
                 self.NTLMSSP_NEGOTIATE_ALWAYS_SIGN |
                 self.NTLMSSP_NEGOTIATE_EXTENDED_SESSIONSECURITY |
                 self.NTLMSSP_REQUEST_TARGET)

        message = b'NTLMSSP\x00'  # Signature: 固定文字列
        message += struct.pack('<I', 1)  # Type: 1 (Negotiate)
        message += struct.pack('<I', flags)  # Flags: ネゴシエーションフラグ
        message += struct.pack('<HHI', 0, 0, 0)  # Domain: 長さ, 最大長, オフセット (空)
        message += struct.pack('<HHI', 0, 0, 0)  # Workstation: 長さ, 最大長, オフセット (空)

        # HTTPヘッダで送信するためBase64エンコード
        return base64.b64encode(message).decode('ascii')

    def parse_type2_message(self, type2_b64):
        """
        Type 2メッセージ（Challenge）を解析

        Args:
            type2_b64: Base64エンコードされたType 2メッセージ

        Returns:
            tuple: (challenge, flags, target_info)
                - challenge: 8バイトのサーバーチャレンジ
                - flags: ネゴシエートフラグ
                - target_info: TargetInfo構造体

        Type 2メッセージ構造:
        - Signature (8バイト): "NTLMSSP\0"
        - MessageType (4バイト): 0x00000002
        - TargetNameFields (8バイト)
        - NegotiateFlags (4バイト)
        - ServerChallenge (8バイト): ← 重要: 認証に使用
        - Reserved (8バイト)
        - TargetInfoFields (8バイト)
        - [TargetInfo]: サーバー情報
        """
        message = base64.b64decode(type2_b64)

        # 署名の検証
        if message[:8] != b'NTLMSSP\x00':
            raise ValueError("Invalid NTLM signature")

        # メッセージタイプの検証
        msg_type = struct.unpack('<I', message[8:12])[0]
        if msg_type != 2:
            raise ValueError(f"Expected Type 2 message, got Type {msg_type}")

        # サーバーチャレンジを取得（オフセット24から8バイト）
        # これが認証の核となる値
        challenge = message[24:32]

        # ネゴシエートフラグを取得
        flags = struct.unpack('<I', message[20:24])[0]

        # TargetInfoを取得（フラグで含まれている場合のみ）
        target_info = b''
        if flags & self.NTLMSSP_NEGOTIATE_TARGET_INFO:
            ti_len = struct.unpack('<H', message[40:42])[0]    # 長さ
            ti_offset = struct.unpack('<I', message[44:48])[0]  # オフセット
            if ti_offset + ti_len <= len(message):
                target_info = message[ti_offset:ti_offset + ti_len]

        return challenge, flags, target_info

    def create_type3_message(self, challenge, target_info):
        """
        Type 3メッセージ（Authenticate）を生成

        Args:
            challenge: サーバーから受信した8バイトのチャレンジ
            target_info: サーバーから受信したTargetInfo

        Returns:
            str: Base64エンコードされたType 3メッセージ

        Type 3メッセージ構造:
        - Signature (8バイト): "NTLMSSP\0"
        - MessageType (4バイト): 0x00000003
        - LmChallengeResponseFields (8バイト): LMv2レスポンス（空）
        - NtChallengeResponseFields (8バイト): NTLMv2レスポンス
        - DomainNameFields (8バイト)
        - UserNameFields (8バイト)
        - WorkstationFields (8バイト)
        - EncryptedRandomSessionKeyFields (8バイト)
        - NegotiateFlags (4バイト)
        - [データ部分]: 各フィールドの実データ

        NTLMv2レスポンスの計算:
        - NTProofStr = HMAC-MD5(NTLMv2Hash, ServerChallenge + Blob)
        - NTResponse = NTProofStr + Blob
        """
        ntlmv2_hash = self._ntlmv2_hash()

        # クライアントチャレンジ（ランダム8バイト）
        client_challenge = os.urandom(8)

        # タイムスタンプ（Windows FILETIME形式）
        timestamp = int((time.time() + 11644473600) * 10000000)
        timestamp_bytes = struct.pack('<Q', timestamp)

        # NTLMv2 blob
        blob = b'\x01\x01'  # Blob signature
        blob += b'\x00\x00'  # Reserved
        blob += b'\x00\x00\x00\x00'  # Reserved
        blob += timestamp_bytes
        blob += client_challenge
        blob += b'\x00\x00\x00\x00'  # Reserved
        blob += target_info
        blob += b'\x00\x00\x00\x00'  # Reserved

        # NTProofStr = HMAC-MD5(NTLMv2Hash, ServerChallenge + Blob)
        nt_proof_str = hmac.new(ntlmv2_hash, challenge + blob, hashlib.md5).digest()

        # NTLMv2 Response = NTProofStr + Blob
        nt_response = nt_proof_str + blob

        # LM Response（空）
        lm_response = b''

        # セッションキー
        session_key = hmac.new(ntlmv2_hash, nt_proof_str, hashlib.md5).digest()

        # UTF-16LEに変換
        domain_utf16 = self._utf16le(self.domain)
        user_utf16 = self._utf16le(self.username)
        workstation_utf16 = b''

        # フラグ
        flags = (self.NTLMSSP_NEGOTIATE_UNICODE |
                 self.NTLMSSP_NEGOTIATE_NTLM |
                 self.NTLMSSP_NEGOTIATE_ALWAYS_SIGN |
                 self.NTLMSSP_NEGOTIATE_EXTENDED_SESSIONSECURITY)

        # オフセット計算
        offset = 88  # ヘッダサイズ

        lm_offset = offset
        offset += len(lm_response)

        nt_offset = offset
        offset += len(nt_response)

        domain_offset = offset
        offset += len(domain_utf16)

        user_offset = offset
        offset += len(user_utf16)

        workstation_offset = offset
        offset += len(workstation_utf16)

        session_key_offset = offset

        # Type 3メッセージ構築
        message = b'NTLMSSP\x00'  # Signature
        message += struct.pack('<I', 3)  # Type 3

        # LM Response
        message += struct.pack('<HHI', len(lm_response), len(lm_response), lm_offset)

        # NT Response
        message += struct.pack('<HHI', len(nt_response), len(nt_response), nt_offset)

        # Domain
        message += struct.pack('<HHI', len(domain_utf16), len(domain_utf16), domain_offset)

        # User
        message += struct.pack('<HHI', len(user_utf16), len(user_utf16), user_offset)

        # Workstation
        message += struct.pack('<HHI', len(workstation_utf16), len(workstation_utf16), workstation_offset)

        # Encrypted Session Key
        message += struct.pack('<HHI', 0, 0, session_key_offset)

        # Flags
        message += struct.pack('<I', flags)

        # 88バイトまでパディング
        message += b'\x00' * (88 - len(message))

        # データ部分
        message += lm_response
        message += nt_response
        message += domain_utf16
        message += user_utf16
        message += workstation_utf16

        return base64.b64encode(message).decode('ascii')


# ==============================================================================
# HTTP通信（ソケット直接使用）
# ==============================================================================
#
# 【なぜrequestsライブラリを使わないのか】
# - IT制限環境では外部パッケージのインストールが禁止されていることが多い
# - 標準ライブラリのsocketモジュールのみで実装
# - NTLM認証のハンドシェイクを完全に制御可能
#
# 【HTTP通信の流れ】
# 1. TCPソケットを作成・接続
# 2. HTTPリクエストを送信（POST /wsman）
# 3. HTTPレスポンスを受信・解析
# 4. ソケットをクローズ
#
# 【NTLM認証のHTTPでの流れ】
# 1. Type 1メッセージを含むリクエスト送信 → 401応答受信
# 2. 401のWWW-AuthenticateヘッダからType 2メッセージ取得
# 3. Type 3メッセージを含むリクエスト送信 → 200応答受信（認証成功）
# ==============================================================================


class HTTPClient:
    """
    ソケットを使用したHTTPクライアント

    requestsライブラリを使用せず、標準ライブラリのsocketのみで実装。
    NTLM認証のための複数リクエスト処理をサポート。
    """

    def __init__(self, host, port, timeout=300):
        """
        HTTPClientの初期化

        Args:
            host: 接続先ホスト名/IPアドレス
            port: 接続先ポート番号
            timeout: ソケットタイムアウト（秒）
        """
        self.host = host
        self.port = port
        self.timeout = timeout

    def _create_socket(self):
        """
        TCPソケットを作成して接続

        Returns:
            socket: 接続済みのソケットオブジェクト
        """
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.settimeout(self.timeout)
        sock.connect((self.host, self.port))
        return sock

    def _send_request(self, sock, method, path, headers, body):
        """
        HTTPリクエストを送信

        Args:
            sock: ソケットオブジェクト
            method: HTTPメソッド（POST等）
            path: リクエストパス（/wsman等）
            headers: HTTPヘッダの辞書
            body: リクエスト本文
        """
        request = f"{method} {path} HTTP/1.1\r\n"
        request += f"Host: {self.host}:{self.port}\r\n"
        for key, value in headers.items():
            request += f"{key}: {value}\r\n"
        request += f"Content-Length: {len(body)}\r\n"
        request += "\r\n"

        sock.sendall(request.encode('utf-8') + body.encode('utf-8'))

    def _recv_response(self, sock):
        """HTTPレスポンスを受信"""
        response = b''
        headers_done = False
        content_length = 0

        while True:
            chunk = sock.recv(4096)
            if not chunk:
                break
            response += chunk

            if not headers_done:
                if b'\r\n\r\n' in response:
                    headers_done = True
                    header_end = response.find(b'\r\n\r\n')
                    headers = response[:header_end].decode('utf-8', errors='replace')

                    # Content-Lengthを取得
                    for line in headers.split('\r\n'):
                        if line.lower().startswith('content-length:'):
                            content_length = int(line.split(':')[1].strip())
                            break

                    # 本文を十分受信したか確認
                    body_start = header_end + 4
                    if len(response) >= body_start + content_length:
                        break
            else:
                header_end = response.find(b'\r\n\r\n')
                body_start = header_end + 4
                if len(response) >= body_start + content_length:
                    break

        return response.decode('utf-8', errors='replace')

    def _parse_response(self, response):
        """HTTPレスポンスを解析"""
        lines = response.split('\r\n')
        status_line = lines[0]
        status_code = int(status_line.split()[1])

        headers = {}
        body_start = 0
        for i, line in enumerate(lines[1:], 1):
            if line == '':
                body_start = i + 1
                break
            if ':' in line:
                key, value = line.split(':', 1)
                headers[key.strip().lower()] = value.strip()

        body = '\r\n'.join(lines[body_start:])

        return status_code, headers, body

    def request_with_ntlm(self, path, body, username, password, domain=''):
        """NTLM認証付きHTTPリクエストを送信"""
        ntlm = NTLMAuth(username, password, domain)

        # Step 1: Type 1メッセージを送信
        sock = self._create_socket()
        try:
            type1 = ntlm.create_type1_message()
            headers = {
                'Authorization': f'NTLM {type1}',
                'Content-Type': 'application/soap+xml;charset=UTF-8',
                'Connection': 'keep-alive'
            }
            self._send_request(sock, 'POST', path, headers, body)
            response = self._recv_response(sock)
        finally:
            sock.close()

        status_code, resp_headers, _ = self._parse_response(response)

        if status_code != 401:
            raise Exception(f"Expected 401, got {status_code}")

        # WWW-AuthenticateヘッダからType 2メッセージを取得
        auth_header = resp_headers.get('www-authenticate', '')
        if not auth_header.upper().startswith('NTLM '):
            raise Exception("No NTLM challenge in response")

        type2_b64 = auth_header[5:].strip()

        # Step 2: Type 2メッセージを解析
        challenge, flags, target_info = ntlm.parse_type2_message(type2_b64)

        # Step 3: Type 3メッセージを送信
        sock = self._create_socket()
        try:
            type3 = ntlm.create_type3_message(challenge, target_info)
            headers = {
                'Authorization': f'NTLM {type3}',
                'Content-Type': 'application/soap+xml;charset=UTF-8',
                'Connection': 'close'
            }
            self._send_request(sock, 'POST', path, headers, body)
            response = self._recv_response(sock)
        finally:
            sock.close()

        status_code, resp_headers, resp_body = self._parse_response(response)

        if status_code == 401:
            raise Exception("Authentication failed (HTTP 401)")
        elif status_code == 500:
            raise Exception(f"Server error (HTTP 500): {resp_body}")
        elif status_code != 200:
            logging.warning(f"Unexpected status code: {status_code}")

        return resp_body


# ==============================================================================
# WinRMクライアント
# ==============================================================================
#
# 【WinRM (Windows Remote Management) とは】
# Microsoftが開発したリモート管理プロトコル。
# WS-Management（Web Services for Management）仕様に基づく。
#
# 【WinRMセッションの流れ】
# 1. Create: リモートシェル（cmd.exe）を作成
#    → ShellIdを取得
#
# 2. Command: シェル上でコマンドを実行
#    → CommandIdを取得
#
# 3. Receive: コマンドの出力（stdout/stderr）を取得
#    → 出力はBase64エンコードされて返される
#    → CommandState/Doneになるまでポーリング
#
# 4. Delete: シェルを削除（リソース解放）
#
# 【SOAPメッセージ形式】
# <?xml version="1.0" encoding="UTF-8"?>
# <s:Envelope xmlns:s="..." xmlns:a="..." ...>
#   <s:Header>
#     <a:Action>...</a:Action>
#     <w:ResourceURI>...</w:ResourceURI>
#     ...
#   </s:Header>
#   <s:Body>
#     ... 操作固有のXML ...
#   </s:Body>
# </s:Envelope>
# ==============================================================================


class WinRMClient:
    """
    WinRMプロトコルを実装したクライアント（NTLM認証版）

    pywinrmライブラリを使用せず、標準ライブラリのみで実装。
    SOAP/XMLメッセージを直接構築してWinRM操作を実行。
    """

    def __init__(self, host, port, username, password, domain='', timeout=300):
        """
        WinRMクライアントの初期化

        Args:
            host: Windows ServerのIPアドレスまたはホスト名
            port: WinRMポート（HTTP: 5985, HTTPS: 5986）
            username: Windowsユーザー名
            password: Windowsパスワード
            domain: ドメイン名（空文字列でローカル認証）
            timeout: タイムアウト（秒）
        """
        self.host = host
        self.port = port
        self.username = username
        self.password = password
        self.domain = domain
        self.timeout = timeout
        self.endpoint = f"http://{host}:{port}/wsman"  # WinRMエンドポイントURL
        self.http_client = HTTPClient(host, port, timeout)  # HTTP通信クライアント

        logging.info(f"WinRMエンドポイント: {self.endpoint}")

    def _send_soap_request(self, soap_envelope):
        """
        SOAPリクエストを送信（NTLM認証付き）

        Args:
            soap_envelope: 送信するSOAP XMLエンベロープ

        Returns:
            str: レスポンスの本文（XML）
        """
        logging.debug("SOAPリクエスト送信")
        return self.http_client.request_with_ntlm(
            '/wsman',
            soap_envelope,
            self.username,
            self.password,
            self.domain
        )

    def _create_shell(self):
        """
        リモートシェル（cmd.exe）を作成

        Returns:
            str: ShellId（後続のコマンド実行に使用）

        WS-Transfer Createアクションを使用してリモートシェルを作成。
        """
        action = "http://schemas.xmlsoap.org/ws/2004/09/transfer/Create"
        resource_uri = "http://schemas.microsoft.com/wbem/wsman/1/windows/shell/cmd"

        soap_envelope = f'''<?xml version="1.0" encoding="UTF-8"?>
<s:Envelope xmlns:s="http://www.w3.org/2003/05/soap-envelope"
            xmlns:a="http://schemas.xmlsoap.org/ws/2004/08/addressing"
            xmlns:w="http://schemas.dmtf.org/wbem/wsman/1/wsman.xsd"
            xmlns:rsp="http://schemas.microsoft.com/wbem/wsman/1/windows/shell">
  <s:Header>
    <a:To>{self.endpoint}</a:To>
    <a:ReplyTo>
      <a:Address s:mustUnderstand="true">http://schemas.xmlsoap.org/ws/2004/08/addressing/role/anonymous</a:Address>
    </a:ReplyTo>
    <a:Action s:mustUnderstand="true">{action}</a:Action>
    <w:MaxEnvelopeSize s:mustUnderstand="true">153600</w:MaxEnvelopeSize>
    <a:MessageID>uuid:{uuid.uuid4()}</a:MessageID>
    <w:Locale xml:lang="ja-JP" s:mustUnderstand="false"/>
    <w:OperationTimeout>PT{self.timeout}S</w:OperationTimeout>
    <w:ResourceURI s:mustUnderstand="true">{resource_uri}</w:ResourceURI>
    <w:OptionSet>
      <w:Option Name="WINRS_NOPROFILE">FALSE</w:Option>
      <w:Option Name="WINRS_CODEPAGE">65001</w:Option>
    </w:OptionSet>
  </s:Header>
  <s:Body>
    <rsp:Shell>
      <rsp:InputStreams>stdin</rsp:InputStreams>
      <rsp:OutputStreams>stdout stderr</rsp:OutputStreams>
    </rsp:Shell>
  </s:Body>
</s:Envelope>'''

        logging.info("シェル作成中...")
        response = self._send_soap_request(soap_envelope)

        root = ET.fromstring(response)
        shell_id = root.find('.//rsp:ShellId', NAMESPACES)
        if shell_id is None:
            raise Exception("シェルIDの取得に失敗しました")

        shell_id_value = shell_id.text
        logging.info(f"シェル作成成功: {shell_id_value}")
        return shell_id_value

    def _run_command(self, shell_id, command):
        """
        シェル上でコマンドを実行

        Args:
            shell_id: 対象のShellId
            command: 実行するコマンド文字列

        Returns:
            str: CommandId（出力取得に使用）

        WinRS Commandアクションを使用してコマンドを実行。
        """
        action = "http://schemas.microsoft.com/wbem/wsman/1/windows/shell/Command"
        resource_uri = "http://schemas.microsoft.com/wbem/wsman/1/windows/shell/cmd"

        command_escaped = command.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')

        soap_envelope = f'''<?xml version="1.0" encoding="UTF-8"?>
<s:Envelope xmlns:s="http://www.w3.org/2003/05/soap-envelope"
            xmlns:a="http://schemas.xmlsoap.org/ws/2004/08/addressing"
            xmlns:w="http://schemas.dmtf.org/wbem/wsman/1/wsman.xsd"
            xmlns:rsp="http://schemas.microsoft.com/wbem/wsman/1/windows/shell">
  <s:Header>
    <a:To>{self.endpoint}</a:To>
    <a:ReplyTo>
      <a:Address s:mustUnderstand="true">http://schemas.xmlsoap.org/ws/2004/08/addressing/role/anonymous</a:Address>
    </a:ReplyTo>
    <a:Action s:mustUnderstand="true">{action}</a:Action>
    <w:MaxEnvelopeSize s:mustUnderstand="true">153600</w:MaxEnvelopeSize>
    <a:MessageID>uuid:{uuid.uuid4()}</a:MessageID>
    <w:Locale xml:lang="ja-JP" s:mustUnderstand="false"/>
    <w:OperationTimeout>PT{self.timeout}S</w:OperationTimeout>
    <w:ResourceURI s:mustUnderstand="true">{resource_uri}</w:ResourceURI>
    <w:SelectorSet>
      <w:Selector Name="ShellId">{shell_id}</w:Selector>
    </w:SelectorSet>
  </s:Header>
  <s:Body>
    <rsp:CommandLine>
      <rsp:Command>{command_escaped}</rsp:Command>
    </rsp:CommandLine>
  </s:Body>
</s:Envelope>'''

        logging.info("コマンド実行中...")
        response = self._send_soap_request(soap_envelope)

        root = ET.fromstring(response)
        command_id = root.find('.//rsp:CommandId', NAMESPACES)
        if command_id is None:
            raise Exception("コマンドIDの取得に失敗しました")

        command_id_value = command_id.text
        logging.info(f"コマンド実行開始: {command_id_value}")
        return command_id_value

    def _get_command_output(self, shell_id, command_id):
        """
        コマンドの出力を取得

        Args:
            shell_id: 対象のShellId
            command_id: 対象のCommandId

        Returns:
            tuple: (exit_code, stdout, stderr)

        WinRS Receiveアクションを使用して出力を取得。
        CommandState/Doneになるまでポーリングを繰り返す。
        出力はBase64エンコードされているためデコードが必要。
        """
        action = "http://schemas.microsoft.com/wbem/wsman/1/windows/shell/Receive"
        resource_uri = "http://schemas.microsoft.com/wbem/wsman/1/windows/shell/cmd"

        stdout_parts = []
        stderr_parts = []
        exit_code = 0
        command_done = False

        logging.info(f"コマンド出力取得中...（最大{self.timeout}秒待機）")

        while not command_done:
            soap_envelope = f'''<?xml version="1.0" encoding="UTF-8"?>
<s:Envelope xmlns:s="http://www.w3.org/2003/05/soap-envelope"
            xmlns:a="http://schemas.xmlsoap.org/ws/2004/08/addressing"
            xmlns:w="http://schemas.dmtf.org/wbem/wsman/1/wsman.xsd"
            xmlns:rsp="http://schemas.microsoft.com/wbem/wsman/1/windows/shell">
  <s:Header>
    <a:To>{self.endpoint}</a:To>
    <a:ReplyTo>
      <a:Address s:mustUnderstand="true">http://schemas.xmlsoap.org/ws/2004/08/addressing/role/anonymous</a:Address>
    </a:ReplyTo>
    <a:Action s:mustUnderstand="true">{action}</a:Action>
    <w:MaxEnvelopeSize s:mustUnderstand="true">153600</w:MaxEnvelopeSize>
    <a:MessageID>uuid:{uuid.uuid4()}</a:MessageID>
    <w:Locale xml:lang="ja-JP" s:mustUnderstand="false"/>
    <w:OperationTimeout>PT{self.timeout}S</w:OperationTimeout>
    <w:ResourceURI s:mustUnderstand="true">{resource_uri}</w:ResourceURI>
    <w:SelectorSet>
      <w:Selector Name="ShellId">{shell_id}</w:Selector>
    </w:SelectorSet>
  </s:Header>
  <s:Body>
    <rsp:Receive>
      <rsp:DesiredStream CommandId="{command_id}">stdout stderr</rsp:DesiredStream>
    </rsp:Receive>
  </s:Body>
</s:Envelope>'''

            response = self._send_soap_request(soap_envelope)
            root = ET.fromstring(response)

            # 標準出力の取得
            for stream in root.findall('.//rsp:Stream[@Name="stdout"]', NAMESPACES):
                if stream.text:
                    decoded = base64.b64decode(stream.text).decode('utf-8', errors='replace')
                    stdout_parts.append(decoded)

            # 標準エラー出力の取得
            for stream in root.findall('.//rsp:Stream[@Name="stderr"]', NAMESPACES):
                if stream.text:
                    decoded = base64.b64decode(stream.text).decode('utf-8', errors='replace')
                    stderr_parts.append(decoded)

            # コマンド完了状態の確認
            state = root.find('.//rsp:CommandState', NAMESPACES)
            if state is not None and 'Done' in state.get('State', ''):
                command_done = True
                exit_code_elem = root.find('.//rsp:ExitCode', NAMESPACES)
                if exit_code_elem is not None:
                    exit_code = int(exit_code_elem.text)

            if not command_done:
                time.sleep(0.5)

        stdout = ''.join(stdout_parts)
        stderr = ''.join(stderr_parts)

        logging.info(f"コマンド完了 (終了コード: {exit_code})")
        return exit_code, stdout, stderr

    def _delete_shell(self, shell_id):
        """
        リモートシェルを削除

        Args:
            shell_id: 削除対象のShellId

        WS-Transfer Deleteアクションを使用してシェルを削除。
        リソース解放のため、コマンド完了後は必ず呼び出すこと。
        """
        action = "http://schemas.xmlsoap.org/ws/2004/09/transfer/Delete"
        resource_uri = "http://schemas.microsoft.com/wbem/wsman/1/windows/shell/cmd"

        soap_envelope = f'''<?xml version="1.0" encoding="UTF-8"?>
<s:Envelope xmlns:s="http://www.w3.org/2003/05/soap-envelope"
            xmlns:a="http://schemas.xmlsoap.org/ws/2004/08/addressing"
            xmlns:w="http://schemas.dmtf.org/wbem/wsman/1/wsman.xsd">
  <s:Header>
    <a:To>{self.endpoint}</a:To>
    <a:ReplyTo>
      <a:Address s:mustUnderstand="true">http://schemas.xmlsoap.org/ws/2004/08/addressing/role/anonymous</a:Address>
    </a:ReplyTo>
    <a:Action s:mustUnderstand="true">{action}</a:Action>
    <w:MaxEnvelopeSize s:mustUnderstand="true">153600</w:MaxEnvelopeSize>
    <a:MessageID>uuid:{uuid.uuid4()}</a:MessageID>
    <w:Locale xml:lang="ja-JP" s:mustUnderstand="false"/>
    <w:OperationTimeout>PT{self.timeout}S</w:OperationTimeout>
    <w:ResourceURI s:mustUnderstand="true">{resource_uri}</w:ResourceURI>
    <w:SelectorSet>
      <w:Selector Name="ShellId">{shell_id}</w:Selector>
    </w:SelectorSet>
  </s:Header>
  <s:Body/>
</s:Envelope>'''

        logging.info("シェル削除中...")
        self._send_soap_request(soap_envelope)
        logging.info("シェル削除完了")

    def execute_command(self, command):
        """
        コマンドを実行する（メインAPI）

        Args:
            command: 実行するコマンド文字列

        Returns:
            tuple: (exit_code, stdout, stderr)

        処理フロー:
        1. シェルを作成（_create_shell）
        2. コマンドを実行（_run_command）
        3. 出力を取得（_get_command_output）
        4. シェルを削除（_delete_shell）- finally句で確実に実行
        """
        shell_id = None
        try:
            shell_id = self._create_shell()
            command_id = self._run_command(shell_id, command)
            exit_code, stdout, stderr = self._get_command_output(shell_id, command_id)
            return exit_code, stdout, stderr
        finally:
            # エラー発生時もシェルを削除してリソース解放
            if shell_id:
                try:
                    self._delete_shell(shell_id)
                except Exception as e:
                    logging.warning(f"シェル削除時にエラー: {e}")

    def execute_batch_file(self, batch_path):
        """
        バッチファイルを実行する

        Args:
            batch_path: 実行するバッチファイルのWindowsパス

        Returns:
            tuple: (exit_code, stdout, stderr)

        cmd.exe /c を使用してバッチファイルを実行。
        """
        command = f'cmd.exe /c "{batch_path}"'
        return self.execute_command(command)


# ==============================================================================
# メイン処理
# ==============================================================================


def setup_logging(level):
    """
    ロギングの設定

    Args:
        level: ログレベル文字列（DEBUG, INFO, WARNING, ERROR）
    """
    numeric_level = getattr(logging, level.upper(), None)
    if not isinstance(numeric_level, int):
        raise ValueError(f'Invalid log level: {level}')

    logging.basicConfig(
        level=numeric_level,
        format='%(asctime)s [%(levelname)s] %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )


def main():
    """
    メインエントリーポイント

    処理フロー:
    1. コマンドライン引数のパース
    2. ログ設定
    3. 環境名の有効性チェック
    4. {ENV}プレースホルダの置換
    5. WinRMクライアント作成・接続
    6. コマンド/バッチファイル実行
    7. 結果表示

    Returns:
        int: 終了コード（成功時0、エラー時1以上）
    """
    parser = argparse.ArgumentParser(
        description='Linux to Windows WinRM Batch Executor (NTLM認証版)',
        usage='%(prog)s ENV [オプション]'
    )
    parser.add_argument('env', metavar='ENV',
                        help='環境名 (例: TST1T, TST2T)')
    parser.add_argument('--host', default=WINDOWS_HOST,
                        help='Windows ServerのIPアドレスまたはホスト名')
    parser.add_argument('--port', type=int, default=WINDOWS_PORT,
                        help='WinRMポート')
    parser.add_argument('--user', default=WINDOWS_USER,
                        help='Windowsユーザー名')
    parser.add_argument('--password', default=WINDOWS_PASSWORD,
                        help='Windowsパスワード')
    parser.add_argument('--domain', default=WINDOWS_DOMAIN,
                        help='ドメイン名（ローカル認証の場合は空）')
    parser.add_argument('--batch', default=BATCH_FILE_PATH,
                        help='実行するバッチファイル（Windows側のパス）')
    parser.add_argument('--command', default=DIRECT_COMMAND,
                        help='直接実行するコマンド')
    parser.add_argument('--timeout', type=int, default=TIMEOUT,
                        help='タイムアウト（秒）')
    parser.add_argument('--log-level', default=LOG_LEVEL,
                        choices=['DEBUG', 'INFO', 'WARNING', 'ERROR'],
                        help='ログレベル')

    args = parser.parse_args()

    setup_logging(args.log_level)

    # タイトル表示
    print("")
    print("=" * 72)
    print("  WinRM Remote Batch Executor (Python)")
    print("  標準ライブラリのみ - NTLM認証版")
    print("=" * 72)
    print("")

    # 環境の有効性チェック
    if args.env not in ENVIRONMENTS:
        logging.error(f"無効な環境が指定されました: {args.env}")
        logging.error(f"利用可能な環境: {', '.join(ENVIRONMENTS)}")
        return 1

    selected_env = args.env
    logging.info(f"指定された環境: {selected_env}")

    # {ENV} プレースホルダーを置換
    if args.batch and '{ENV}' in args.batch:
        args.batch = args.batch.replace('{ENV}', selected_env)
    if args.command and '{ENV}' in args.command:
        args.command = args.command.replace('{ENV}', selected_env)

    logging.info(f"接続先: http://{args.host}:{args.port}/wsman")
    logging.info(f"ユーザー: {args.user}")

    try:
        client = WinRMClient(
            host=args.host,
            port=args.port,
            username=args.user,
            password=args.password,
            domain=args.domain,
            timeout=args.timeout
        )

        if args.command:
            logging.info(f"直接コマンド実行: {args.command}")
            exit_code, stdout, stderr = client.execute_command(args.command)
        elif args.batch:
            logging.info(f"バッチファイル実行: {args.batch}")
            exit_code, stdout, stderr = client.execute_batch_file(args.batch)
        else:
            logging.error("実行するコマンドまたはバッチファイルが指定されていません")
            return 1

        # 結果の表示
        print("\n" + "=" * 60)
        print("実行結果")
        print("=" * 60)

        if stdout:
            print("\n[標準出力]")
            print(stdout)

        if stderr:
            print("\n[標準エラー出力]")
            print(stderr)

        print(f"\n終了コード: {exit_code}")
        print("=" * 60)

        if exit_code == 0:
            logging.info("完了")
        else:
            logging.error(f"コマンドが失敗しました (終了コード: {exit_code})")

        return exit_code

    except Exception as e:
        logging.error(f"エラーが発生しました: {e}")
        import traceback
        logging.debug(traceback.format_exc())
        return 1


if __name__ == "__main__":
    sys.exit(main())

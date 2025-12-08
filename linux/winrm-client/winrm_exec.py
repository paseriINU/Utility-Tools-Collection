#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
WinRM Remote Batch Executor for Linux (標準ライブラリのみ - NTLM認証版)
Linux（Red Hat等）からWindows Server 2022へWinRM接続してバッチを実行

必要な環境:
    Python 3.6以降（標準ライブラリのみ使用、追加パッケージ不要）
    外部ライブラリ不要 - MD4/NTLM認証を自前実装

使い方:
    # 環境を引数で指定（必須）
    python3 winrm_exec_ntlm.py TST1T
    python3 winrm_exec_ntlm.py TST2T

    # またはコマンドライン引数で詳細設定も指定
    python3 winrm_exec_ntlm.py TST1T --host 192.168.1.100 --user Administrator --password Pass123
"""

import sys
import argparse
import logging
import base64
import uuid
import socket
import struct
import time
import hashlib
import hmac
import os
from xml.etree import ElementTree as ET

# ==================== 設定セクション ====================
# ここを編集して使用してください

# Windows接続情報
WINDOWS_HOST = "192.168.1.100"      # Windows ServerのIPアドレスまたはホスト名
WINDOWS_PORT = 5985                  # WinRMポート（HTTP: 5985, HTTPS: 5986）
WINDOWS_USER = "Administrator"       # Windowsユーザー名
WINDOWS_PASSWORD = "YourPassword"    # Windowsパスワード
WINDOWS_DOMAIN = ""                  # ドメイン（空文字列でローカル認証）

# 環境フォルダ名のリスト（実行時に選択可能）
# 新しい環境を追加する場合は、このリストに追加してください
ENVIRONMENTS = ["TST1T", "TST2T"]    # 利用可能な環境のリスト

# 実行するバッチファイル（Windows側のパス）
# {ENV} は選択した環境フォルダ名に置換されます
BATCH_FILE_PATH = r"C:\Scripts\{ENV}\test.bat"

# または直接コマンドを指定
DIRECT_COMMAND = None  # 例: "echo Hello from WinRM"

# タイムアウト（秒）
TIMEOUT = 300

# ログレベル（DEBUG, INFO, WARNING, ERROR）
LOG_LEVEL = "INFO"

# =========================================================

# XML名前空間の定義
NAMESPACES = {
    's': 'http://www.w3.org/2003/05/soap-envelope',
    'a': 'http://schemas.xmlsoap.org/ws/2004/08/addressing',
    'w': 'http://schemas.dmtf.org/wbem/wsman/1/wsman.xsd',
    'rsp': 'http://schemas.microsoft.com/wbem/wsman/1/windows/shell',
    'p': 'http://schemas.microsoft.com/wbem/wsman/1/wsman.xsd'
}


# ==================== MD4実装（標準ライブラリにないため自前実装） ====================

class MD4:
    """MD4ハッシュアルゴリズムの実装"""

    def __init__(self):
        self.A = 0x67452301
        self.B = 0xefcdab89
        self.C = 0x98badcfe
        self.D = 0x10325476
        self.count = 0
        self.buffer = b''

    @staticmethod
    def _left_rotate(x, n):
        return ((x << n) | (x >> (32 - n))) & 0xffffffff

    @staticmethod
    def _F(x, y, z):
        return (x & y) | (~x & z)

    @staticmethod
    def _G(x, y, z):
        return (x & y) | (x & z) | (y & z)

    @staticmethod
    def _H(x, y, z):
        return x ^ y ^ z

    def _process_block(self, block):
        """64バイトブロックを処理"""
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
        """データを追加"""
        if isinstance(data, str):
            data = data.encode('utf-8')

        self.buffer += data
        self.count += len(data)

        while len(self.buffer) >= 64:
            self._process_block(self.buffer[:64])
            self.buffer = self.buffer[64:]

    def digest(self):
        """ハッシュ値を取得"""
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
    """MD4ハッシュを計算"""
    hasher = MD4()
    hasher.update(data)
    return hasher.digest()


# ==================== NTLM認証実装 ====================

class NTLMAuth:
    """NTLMv2認証の実装"""

    NTLMSSP_NEGOTIATE_UNICODE = 0x00000001
    NTLMSSP_NEGOTIATE_NTLM = 0x00000200
    NTLMSSP_NEGOTIATE_ALWAYS_SIGN = 0x00008000
    NTLMSSP_NEGOTIATE_EXTENDED_SESSIONSECURITY = 0x00080000
    NTLMSSP_REQUEST_TARGET = 0x00000004
    NTLMSSP_NEGOTIATE_TARGET_INFO = 0x00800000

    def __init__(self, username, password, domain=''):
        self.username = username
        self.password = password
        self.domain = domain

    @staticmethod
    def _utf16le(s):
        """文字列をUTF-16LEに変換"""
        return s.encode('utf-16-le')

    def _nt_hash(self):
        """NTハッシュを計算（パスワードのMD4ハッシュ）"""
        return md4_hash(self._utf16le(self.password))

    def _ntlmv2_hash(self):
        """NTLMv2ハッシュを計算"""
        nt_hash = self._nt_hash()
        user_domain = (self.username.upper() + self.domain).encode('utf-16-le')
        return hmac.new(nt_hash, user_domain, hashlib.md5).digest()

    def create_type1_message(self):
        """Type 1メッセージ（Negotiate）を生成"""
        flags = (self.NTLMSSP_NEGOTIATE_UNICODE |
                 self.NTLMSSP_NEGOTIATE_NTLM |
                 self.NTLMSSP_NEGOTIATE_ALWAYS_SIGN |
                 self.NTLMSSP_NEGOTIATE_EXTENDED_SESSIONSECURITY |
                 self.NTLMSSP_REQUEST_TARGET)

        message = b'NTLMSSP\x00'  # Signature
        message += struct.pack('<I', 1)  # Type 1
        message += struct.pack('<I', flags)  # Flags
        message += struct.pack('<HHI', 0, 0, 0)  # Domain (empty)
        message += struct.pack('<HHI', 0, 0, 0)  # Workstation (empty)

        return base64.b64encode(message).decode('ascii')

    def parse_type2_message(self, type2_b64):
        """Type 2メッセージ（Challenge）を解析"""
        message = base64.b64decode(type2_b64)

        if message[:8] != b'NTLMSSP\x00':
            raise ValueError("Invalid NTLM signature")

        msg_type = struct.unpack('<I', message[8:12])[0]
        if msg_type != 2:
            raise ValueError(f"Expected Type 2 message, got Type {msg_type}")

        # Challengeを取得（オフセット24から8バイト）
        challenge = message[24:32]

        # フラグを取得
        flags = struct.unpack('<I', message[20:24])[0]

        # TargetInfoを取得（フラグにNTLMSSP_NEGOTIATE_TARGET_INFOが含まれている場合）
        target_info = b''
        if flags & self.NTLMSSP_NEGOTIATE_TARGET_INFO:
            ti_len = struct.unpack('<H', message[40:42])[0]
            ti_offset = struct.unpack('<I', message[44:48])[0]
            if ti_offset + ti_len <= len(message):
                target_info = message[ti_offset:ti_offset + ti_len]

        return challenge, flags, target_info

    def create_type3_message(self, challenge, target_info):
        """Type 3メッセージ（Authenticate）を生成"""
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


# ==================== HTTP通信（ソケット直接使用） ====================

class HTTPClient:
    """ソケットを使用したHTTPクライアント"""

    def __init__(self, host, port, timeout=300):
        self.host = host
        self.port = port
        self.timeout = timeout

    def _create_socket(self):
        """ソケットを作成"""
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.settimeout(self.timeout)
        sock.connect((self.host, self.port))
        return sock

    def _send_request(self, sock, method, path, headers, body):
        """HTTPリクエストを送信"""
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


# ==================== WinRMクライアント ====================

class WinRMClient:
    """WinRMプロトコルを実装したクライアント（NTLM認証版）"""

    def __init__(self, host, port, username, password, domain='', timeout=300):
        """
        WinRMクライアントの初期化

        Args:
            host: WindowsサーバのIPアドレスまたはホスト名
            port: WinRMポート
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
        self.endpoint = f"http://{host}:{port}/wsman"
        self.http_client = HTTPClient(host, port, timeout)

        logging.info(f"WinRMエンドポイント: {self.endpoint}")

    def _send_soap_request(self, soap_envelope):
        """SOAPリクエストを送信（NTLM認証付き）"""
        logging.debug("SOAPリクエスト送信")
        return self.http_client.request_with_ntlm(
            '/wsman',
            soap_envelope,
            self.username,
            self.password,
            self.domain
        )

    def _create_shell(self):
        """リモートシェルを作成"""
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
        """シェル上でコマンドを実行"""
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
        """コマンドの出力を取得"""
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
        """シェルを削除"""
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
        """コマンドを実行"""
        shell_id = None
        try:
            shell_id = self._create_shell()
            command_id = self._run_command(shell_id, command)
            exit_code, stdout, stderr = self._get_command_output(shell_id, command_id)
            return exit_code, stdout, stderr
        finally:
            if shell_id:
                try:
                    self._delete_shell(shell_id)
                except Exception as e:
                    logging.warning(f"シェル削除時にエラー: {e}")

    def execute_batch_file(self, batch_path):
        """バッチファイルを実行"""
        command = f'cmd.exe /c "{batch_path}"'
        return self.execute_command(command)


def setup_logging(level):
    """ロギングの設定"""
    numeric_level = getattr(logging, level.upper(), None)
    if not isinstance(numeric_level, int):
        raise ValueError(f'Invalid log level: {level}')

    logging.basicConfig(
        level=numeric_level,
        format='%(asctime)s [%(levelname)s] %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )


def main():
    """メイン処理"""
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

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
WinRM Remote Batch Executor for Linux (標準ライブラリのみ版)
Linux（Red Hat等）からWindows Server 2022へWinRM接続してバッチを実行

必要な環境:
    Python 3.6以降（標準ライブラリのみ使用、追加パッケージ不要）

使い方:
    # スクリプト内の設定を編集してから実行
    python3 winrm_exec.py

    # またはコマンドライン引数で指定
    python3 winrm_exec.py --host 192.168.1.100 --user Administrator --password Pass123
"""

import sys
import argparse
import logging
import base64
import uuid
import socket
import ssl
from urllib.request import Request, urlopen, HTTPPasswordMgrWithDefaultRealm, HTTPBasicAuthHandler, build_opener
from urllib.error import URLError, HTTPError
from xml.etree import ElementTree as ET

# ==================== 設定セクション ====================
# ここを編集して使用してください

# Windows接続情報
WINDOWS_HOST = "192.168.1.100"      # Windows ServerのIPアドレスまたはホスト名
WINDOWS_PORT = 5985                  # WinRMポート（HTTP: 5985, HTTPS: 5986）
WINDOWS_USER = "Administrator"       # Windowsユーザー名
WINDOWS_PASSWORD = "YourPassword"    # Windowsパスワード

# 実行するバッチファイル（Windows側のパス）
BATCH_FILE_PATH = r"C:\Scripts\test.bat"

# または直接コマンドを指定
DIRECT_COMMAND = None  # 例: "echo Hello from WinRM"

# HTTPS接続を使用する場合はTrueに設定
USE_HTTPS = False

# 証明書検証を無効にする場合はTrue（自己署名証明書の場合）
DISABLE_CERT_VALIDATION = True

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


class WinRMClient:
    """WinRMプロトコルを実装したクライアント（標準ライブラリのみ使用）"""

    def __init__(self, host, port, username, password, use_https=False,
                 disable_cert_validation=True, timeout=300):
        """
        WinRMクライアントの初期化

        Args:
            host: WindowsサーバのIPアドレスまたはホスト名
            port: WinRMポート
            username: Windowsユーザー名
            password: Windowsパスワード
            use_https: HTTPS接続を使用するか
            disable_cert_validation: 証明書検証を無効にするか
            timeout: タイムアウト（秒）
        """
        self.host = host
        self.port = port
        self.username = username
        self.password = password
        self.timeout = timeout
        self.disable_cert_validation = disable_cert_validation

        # プロトコルとエンドポイントの設定
        protocol = "https" if use_https else "http"
        self.endpoint = f"{protocol}://{host}:{port}/wsman"

        # Basic認証の設定
        password_mgr = HTTPPasswordMgrWithDefaultRealm()
        password_mgr.add_password(None, self.endpoint, username, password)
        auth_handler = HTTPBasicAuthHandler(password_mgr)
        self.opener = build_opener(auth_handler)

        logging.info(f"WinRMエンドポイント: {self.endpoint}")

    def _send_soap_request(self, soap_envelope):
        """
        SOAPリクエストを送信

        Args:
            soap_envelope: SOAP XMLエンベロープ

        Returns:
            レスポンスのXML文字列
        """
        headers = {
            'Content-Type': 'application/soap+xml;charset=UTF-8',
            'User-Agent': 'Python-WinRM-Client/1.0'
        }

        # Basic認証のヘッダーを追加
        credentials = f"{self.username}:{self.password}"
        encoded_credentials = base64.b64encode(credentials.encode()).decode()
        headers['Authorization'] = f'Basic {encoded_credentials}'

        request = Request(
            self.endpoint,
            data=soap_envelope.encode('utf-8'),
            headers=headers,
            method='POST'
        )

        try:
            # SSL証明書検証の無効化（必要な場合）
            if self.disable_cert_validation:
                import ssl
                context = ssl._create_unverified_context()
                response = urlopen(request, timeout=self.timeout, context=context)
            else:
                response = urlopen(request, timeout=self.timeout)

            return response.read().decode('utf-8')

        except HTTPError as e:
            error_body = e.read().decode('utf-8', errors='replace')
            logging.error(f"HTTP Error {e.code}: {e.reason}")
            logging.debug(f"Error body: {error_body}")
            raise Exception(f"WinRM HTTP Error: {e.code} {e.reason}")
        except URLError as e:
            logging.error(f"URL Error: {e.reason}")
            raise Exception(f"WinRM Connection Error: {e.reason}")
        except socket.timeout:
            logging.error("接続タイムアウト")
            raise Exception("WinRM Connection Timeout")

    def _create_shell(self):
        """
        リモートシェルを作成

        Returns:
            シェルID
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

        logging.debug("シェル作成リクエスト送信")
        response = self._send_soap_request(soap_envelope)

        # レスポンスからシェルIDを抽出
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
            shell_id: シェルID
            command: 実行するコマンド

        Returns:
            コマンドID
        """
        action = "http://schemas.microsoft.com/wbem/wsman/1/windows/shell/Command"
        resource_uri = "http://schemas.microsoft.com/wbem/wsman/1/windows/shell/cmd"

        # コマンドをXMLエスケープ
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

        logging.debug("コマンド実行リクエスト送信")
        response = self._send_soap_request(soap_envelope)

        # レスポンスからコマンドIDを抽出
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
            shell_id: シェルID
            command_id: コマンドID

        Returns:
            (終了コード, 標準出力, 標準エラー出力)
        """
        action = "http://schemas.microsoft.com/wbem/wsman/1/windows/shell/Receive"
        resource_uri = "http://schemas.microsoft.com/wbem/wsman/1/windows/shell/cmd"

        stdout_parts = []
        stderr_parts = []
        exit_code = 0
        command_done = False

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

            logging.debug("出力取得リクエスト送信")
            response = self._send_soap_request(soap_envelope)

            # レスポンスをパース
            root = ET.fromstring(response)

            # 標準出力の取得
            for stream in root.findall('.//rsp:Stream[@Name="stdout"]', NAMESPACES):
                if stream.text:
                    decoded = base64.b64decode(stream.text).decode('cp932', errors='replace')
                    stdout_parts.append(decoded)

            # 標準エラー出力の取得
            for stream in root.findall('.//rsp:Stream[@Name="stderr"]', NAMESPACES):
                if stream.text:
                    decoded = base64.b64decode(stream.text).decode('cp932', errors='replace')
                    stderr_parts.append(decoded)

            # コマンド完了状態の確認
            state = root.find('.//rsp:CommandState', NAMESPACES)
            if state is not None and state.get('State') == 'http://schemas.microsoft.com/wbem/wsman/1/windows/shell/CommandState/Done':
                command_done = True
                exit_code_elem = root.find('.//rsp:ExitCode', NAMESPACES)
                if exit_code_elem is not None:
                    exit_code = int(exit_code_elem.text)

        stdout = ''.join(stdout_parts)
        stderr = ''.join(stderr_parts)

        logging.info(f"コマンド完了: 終了コード={exit_code}")
        return exit_code, stdout, stderr

    def _delete_shell(self, shell_id):
        """
        シェルを削除

        Args:
            shell_id: シェルID
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

        logging.debug("シェル削除リクエスト送信")
        self._send_soap_request(soap_envelope)
        logging.info("シェル削除完了")

    def execute_command(self, command):
        """
        コマンドを実行（シェル作成→実行→出力取得→シェル削除）

        Args:
            command: 実行するコマンド

        Returns:
            (終了コード, 標準出力, 標準エラー出力)
        """
        shell_id = None
        try:
            # シェル作成
            shell_id = self._create_shell()

            # コマンド実行
            command_id = self._run_command(shell_id, command)

            # 出力取得
            exit_code, stdout, stderr = self._get_command_output(shell_id, command_id)

            return exit_code, stdout, stderr

        finally:
            # シェル削除（必ず実行）
            if shell_id:
                try:
                    self._delete_shell(shell_id)
                except Exception as e:
                    logging.warning(f"シェル削除時にエラー: {e}")

    def execute_batch_file(self, batch_path):
        """
        バッチファイルを実行

        Args:
            batch_path: バッチファイルのパス（Windows側）

        Returns:
            (終了コード, 標準出力, 標準エラー出力)
        """
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
    # コマンドライン引数のパース
    parser = argparse.ArgumentParser(
        description='Linux to Windows WinRM Batch Executor (標準ライブラリのみ版)'
    )
    parser.add_argument('--host', default=WINDOWS_HOST,
                        help='Windows ServerのIPアドレスまたはホスト名')
    parser.add_argument('--port', type=int, default=WINDOWS_PORT,
                        help='WinRMポート')
    parser.add_argument('--user', default=WINDOWS_USER,
                        help='Windowsユーザー名')
    parser.add_argument('--password', default=WINDOWS_PASSWORD,
                        help='Windowsパスワード')
    parser.add_argument('--batch', default=BATCH_FILE_PATH,
                        help='実行するバッチファイル（Windows側のパス）')
    parser.add_argument('--command', default=DIRECT_COMMAND,
                        help='直接実行するコマンド')
    parser.add_argument('--https', action='store_true', default=USE_HTTPS,
                        help='HTTPS接続を使用')
    parser.add_argument('--no-cert-check', action='store_true',
                        default=DISABLE_CERT_VALIDATION,
                        help='証明書検証を無効化')
    parser.add_argument('--timeout', type=int, default=TIMEOUT,
                        help='タイムアウト（秒）')
    parser.add_argument('--log-level', default=LOG_LEVEL,
                        choices=['DEBUG', 'INFO', 'WARNING', 'ERROR'],
                        help='ログレベル')

    args = parser.parse_args()

    # ロギング設定
    setup_logging(args.log_level)

    logging.info("=== WinRM Remote Batch Executor (標準ライブラリ版) ===")
    logging.info(f"接続先: {args.host}:{args.port}")
    logging.info(f"ユーザー: {args.user}")

    try:
        # WinRMクライアントの作成
        client = WinRMClient(
            host=args.host,
            port=args.port,
            username=args.user,
            password=args.password,
            use_https=args.https,
            disable_cert_validation=args.no_cert_check,
            timeout=args.timeout
        )

        # コマンドまたはバッチファイルの実行
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

        return exit_code

    except Exception as e:
        logging.error(f"エラーが発生しました: {e}")
        import traceback
        logging.debug(traceback.format_exc())
        return 1


if __name__ == "__main__":
    sys.exit(main())

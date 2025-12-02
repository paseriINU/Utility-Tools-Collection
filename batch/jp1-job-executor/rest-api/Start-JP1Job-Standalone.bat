<# :
@echo off
setlocal
chcp 65001 >nul
powershell -NoProfile -ExecutionPolicy Bypass -Command "iex ((gc '%~f0') -join \"`n\")"
exit /b %ERRORLEVEL%
: #> | sv -name _ > $null

#Requires -Version 5.1
<#
.SYNOPSIS
    JP1/AJS3 REST APIを使用してジョブネットを起動する（スタンドアローン版）

.DESCRIPTION
    設定を内部に記述できるスタンドアローン版です。
    ファイル内の「設定セクション」を編集して、ダブルクリックで実行できます。

.NOTES
    必要なバージョン: JP1/AJS3 - Manager 10以降
    PowerShell 5.1以降
    作成日: 2025-12-02
    バージョン: 2.0

    使い方:
    1. 下記の「■ 設定セクション」を編集
    2. このファイルをダブルクリックで実行
#>

# ==============================================================================
# ■ 設定セクション（ここを編集してください）
# ==============================================================================

$Config = @{
    # JP1/AJS3マネージャーのホスト名またはIPアドレス
    JP1Host = "192.168.1.100"

    # JP1/AJS3 REST APIのポート番号
    JP1Port = 22250

    # JP1ユーザー名
    JP1User = "jp1admin"

    # JP1パスワード（空の場合は実行時に入力を求められます）
    JP1Password = ""

    # 起動するジョブネットのフルパス
    JobnetPath = "/main_unit/jobgroup1/daily_batch"

    # HTTPS接続を使用する場合は $true（通常は $false）
    UseSSL = $false
}

# ==============================================================================
# ■ メイン処理（以下は編集不要）
# ==============================================================================

# ========================================
# 初期設定
# ========================================

$ErrorActionPreference = "Stop"

# UTF-8出力設定
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# 証明書検証を無効化（自己署名証明書対応）
if ($Config.UseSSL) {
    Add-Type @"
        using System.Net;
        using System.Security.Cryptography.X509Certificates;
        public class TrustAllCertsPolicy : ICertificatePolicy {
            public bool CheckValidationResult(
                ServicePoint srvPoint, X509Certificate certificate,
                WebRequest request, int certificateProblem) {
                return true;
            }
        }
"@
    [System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
}

# ========================================
# パスワード取得
# ========================================

if ([string]::IsNullOrEmpty($Config.JP1Password)) {
    $JP1PasswordSecure = Read-Host "JP1パスワードを入力してください" -AsSecureString
} else {
    $JP1PasswordSecure = ConvertTo-SecureString $Config.JP1Password -AsPlainText -Force
}

# SecureStringを平文に変換（API送信用）
$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($JP1PasswordSecure)
$PlainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

# ========================================
# REST API エンドポイント構築
# ========================================

$protocol = if ($Config.UseSSL) { "https" } else { "http" }
$baseUrl = "${protocol}://$($Config.JP1Host):$($Config.JP1Port)/ajs3web/api"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "JP1ジョブネット起動（REST API版）" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "JP1ホスト      : $($Config.JP1Host)"
Write-Host "ポート番号      : $($Config.JP1Port)"
Write-Host "プロトコル      : $protocol"
Write-Host "JP1ユーザー    : $($Config.JP1User)"
Write-Host "ジョブネットパス: $($Config.JobnetPath)"
Write-Host ""

# ========================================
# 認証トークン取得
# ========================================

Write-Host "認証トークンを取得中..." -ForegroundColor Cyan

$authUrl = "$baseUrl/auth/login"
$authBody = @{
    userName = $Config.JP1User
    password = $PlainPassword
} | ConvertTo-Json

try {
    $authResponse = Invoke-RestMethod -Uri $authUrl -Method Post -Body $authBody -ContentType "application/json"
    $token = $authResponse.token
    Write-Host "認証成功" -ForegroundColor Green
    Write-Host ""
} catch {
    Write-Host ""
    Write-Host "[エラー] 認証に失敗しました。" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host ""
    Write-Host "以下を確認してください：" -ForegroundColor Yellow
    Write-Host "- JP1ホスト名、ポート番号が正しいか" -ForegroundColor Yellow
    Write-Host "- JP1ユーザー名、パスワードが正しいか" -ForegroundColor Yellow
    Write-Host "- JP1/AJS3のREST APIサービスが起動しているか" -ForegroundColor Yellow
    Write-Host "- ネットワーク接続が正常か" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Enterキーを押して終了..." -ForegroundColor Gray
    $null = Read-Host
    exit 1
}

# ========================================
# ジョブネット起動
# ========================================

Write-Host "ジョブネットを起動中..." -ForegroundColor Cyan

# ジョブネットパスをエンコード
Add-Type -AssemblyName System.Web
$encodedPath = [System.Web.HttpUtility]::UrlEncode($Config.JobnetPath)

$startUrl = "$baseUrl/jobnets/$encodedPath/executions"

$headers = @{
    "Authorization" = "Bearer $token"
    "Content-Type" = "application/json"
}

$startBody = @{
    executionType = "immediate"
} | ConvertTo-Json

try {
    $startResponse = Invoke-RestMethod -Uri $startUrl -Method Post -Headers $headers -Body $startBody

    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "ジョブネットの起動に成功しました" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "ジョブネット: $($Config.JobnetPath)"
    Write-Host "実行ID      : $($startResponse.executionId)"
    Write-Host "ホスト      : $($Config.JP1Host)"
    Write-Host ""

} catch {
    Write-Host ""
    Write-Host "[エラー] ジョブネットの起動に失敗しました。" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host ""
    Write-Host "以下を確認してください：" -ForegroundColor Yellow
    Write-Host "- ジョブネットパスが正しいか" -ForegroundColor Yellow
    Write-Host "- ジョブネットが存在するか" -ForegroundColor Yellow
    Write-Host "- ジョブネットの実行権限があるか" -ForegroundColor Yellow

    # ログアウト
    try {
        $logoutUrl = "$baseUrl/auth/logout"
        Invoke-RestMethod -Uri $logoutUrl -Method Post -Headers $headers | Out-Null
    } catch {
        # ログアウト失敗は無視
    }

    Write-Host ""
    Write-Host "Enterキーを押して終了..." -ForegroundColor Gray
    $null = Read-Host
    exit 1
}

# ========================================
# ログアウト
# ========================================

Write-Host "ログアウト中..." -ForegroundColor Cyan

try {
    $logoutUrl = "$baseUrl/auth/logout"
    Invoke-RestMethod -Uri $logoutUrl -Method Post -Headers $headers | Out-Null
    Write-Host "ログアウト完了" -ForegroundColor Green
} catch {
    Write-Host "ログアウトに失敗しました（無視します）" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "処理が完了しました。" -ForegroundColor Cyan
Write-Host ""
Write-Host "Enterキーを押して終了..." -ForegroundColor Gray
$null = Read-Host

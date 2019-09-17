##########################################################################
#
#
# pram
# return sting message
##########################################################################
#
$cmdResult = @{
    $userGroup =
}

#
$config = @{
    printServer = @{
        ipddr = @(
            192.168.100.1
            192.168.100.2
            192.168.100.3
            192.168.100.4
        )
    }
    outputFile = Join-Path $env:USERPROFILE 'Desktop¥ネットワークプリンター情報.csv'
}

#
if(-not($userGroup -Contain('domainGroup¥Domain Admins'))) {
    Wite-Output ''
    exit 1
}

# ネットワークプリンター情報取得
# Destscription:
# 各プリンターサーバの情報を取得する
$config | ForEach-Object {
    Get-WmiObject Win32_Printer -ComputerName $_.printerServer.ipddr
} `
    | Select-Object -Property Name, PrinterStatus, SystemName, ShareName, AccessName, DriverName, Location, PortName, Comment `
    | Add-Members @{
    } `
        | Export-Csv -Path $config.outputFile -Encording UTF8 -NoTypeInformation

# ヘッダー置換
# Destscription:
# 可読性を上げるために日本語に変換する
Get-Content $config.outputFile | ForEach-Object {
    $_ -Creplace 'PrinterStatus', 'ステータス'
    $_ -Creplace 'SystemName', '格納先サーバ'
    $_ -Creplace 'ShareName', '共有名'
    $_ -Creplace 'AccessName', 'アクセス名'
    $_ -Creplace 'DriverName', 'ドライバー名'
    $_ -Creplace 'Location', '設置場所'
    $_ -Creplace 'PortName', 'プリンターポート'
    $_ -Creplace 'Comment', 'コメント'
    $_ -Creplace 'Name', '管理番号'
} `
    | Out-File $config.outputFile

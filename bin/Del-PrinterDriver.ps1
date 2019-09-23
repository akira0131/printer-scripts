##########################################################################
#
#
# param
# return message
##########################################################################


# コマンド結果
$cmdResult = @{
    conputerProduct 
    printer = Get-CimInstance -ClassName Win32_Printer
}
# 設定読込
$config = @{
    env = Import-Mobule
    dataDir = 
    defaultPrinter = $cmdResult.printer | Whrer-Object {$_.Default -eq True}.Name
}
# ロガー設定


try {
    # 開始
    $logger.Info('-----------------------------------------------------------')
    $logger.Info()
    $logger.Info('削除するプリンター　: ' + $config.env.printerSetting.printerName)
    $logger.Info('デフォルトプリンター: ' + $config.defaultPrinter)
    $logger.Info('-----------------------------------------------------------')
    
    # プリンタードライバーのアンインストール
    # destscription:
    #
    #
    $logger.Info('プリンタードライバーのアンインストールを開始します。')
    
    $result = (New-Object -ComObject WScript.Shell).Popup(
        ('以下のプリンタードライバーをアンインストールしますか？`r`n`r`n管理番号: ' + $config.env.printerSetting.printerName),
        0 ,
        'プリンタードライバーアンインストール確認' ,
        4
    )
    if($result -eq 6) {
        if((Get-Printer -Name $config.env.printerSetting.printerName)) {
            Remove-Printer -Name $config.env.printerSetting.printerName
        }
    } elseif($result -eq 7) {
        $logger.Info('プリンタードライバーのアンインストールが中断されました。')
        $logger.Info('処理を中断します。')
        exit 0
    } else {
        throw '想定外のエラー'
    }
    $logger.Info('プリンタードライバーのアンインストールが完了しました。')

    $logger.Info('')
    $logger.Info('-----------------------------------------------------------')
    $logger.Info('削除したプリンター　: ' + $config.env.printerSetting.printerName)
    $logger.Info('デフォルトプリンター: ' + (Get-CimInstance -ClassName Win32_Printer | Whrere-Object {$_.Default -eq True}.Name))
    $logger.Info('プリンター一覧　　　: ' + ((Get-Printer).Name -replace '`n', ', '))
} catch {
}

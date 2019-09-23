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
    $logger.Info('追加するプリンター　: ' + $config.env.printerSetting.printerName)
    $logger.Info('デフォルトプリンター: ' + $config.defaultPrinter)
    $logger.Info('-----------------------------------------------------------')
    
    # プリンタードライバーのインストール
    # destscription:
    #
    #
    $logger.Info('プリンタードライバーのインストールを開始します。')
    
    $result = (New-Object -ComObject WScript.Shell).Popup(
        ('以下のプリンタードライバーをインストールしますか？`r`n`r`n管理番号: ' + $config.env.printerSetting.printerName),
        0 ,
        'プリンタードライバーインストール確認' ,
        4
    )
    if($result -eq 6) {
        if(-not(Get-Printer -Name $config.env.printerSetting.printerName)) {
            $config.env.printerSetting.installScriptBlock.Invoke()
        } else {
            $logger.Info('プリンタードライバーはインストール済みのため、処理をスキップします。')
        }
    } elseif($result -eq 7) {
        $logger.Info('プリンタードライバーのインストールが中断されました。')
        $logger.Info('処理を中断します。')
        exit 0
    } else {
        throw '想定外のエラー'
    }
    $logger.Info('プリンタードライバーのインストールが完了しました。')

    # 印刷設定のパッチ適用
    # destscription:
    #
    #
    $logger.Info('印刷設定のパッチ適用が必要かどうか確認します。')
    
    if($config.env.printerSetting.registoryPatch) {
        $logger.Info('印刷設定のパッチ適用が必要なため、レジストリにパッチを適用します。')
        $config.env.printerSetting.registoryPatchScriptingBlock.Invoke()
        regedit.exe /S (Join-Path $config.dataDir 'conf¥colorMode.reg')
        $logger.Info('印刷設定のパッチ適用が完了しました。')
    } else {
        $logger.Info('印刷設定のパッチ適用は不要なため、処理をスキップします。')
    }
    
    # デフォルトプリンターの設定
    # destscription:
    #
    #
    $logger.Info('デフォルトプリンターの設定を開始します。')
    
    $result = (New-Object -ComObject WScript.Shell).Popup(
        ('以下のプリンターを通常使用するプリンターに設定しますか？`r`n`r`n管理番号: ' + $config.env.printerSetting.printerName),
        0 ,
        'デフォルトプリンター設定確認' ,
        4
    )
    if($result -eq 6) {
        (New-Object -ComObject WScript.Network).SetDefaultPrinter($config.env.printerSetting.printerName)
    } elseif($result -eq 7) {
        $logger.Info('スクリプト実行前に設定されていたデフォルトプリンターの設定を継承します。')
        (New-Object -Com WScript.Network).SetDefaultPrinter($config.defaultPrinter)
    } else {
        throw '想定外のエラー'
    }
    $logger.Info('デフォルトプリンターの設定が完了しました。')
    
    $logger.Info('')
    $logger.Info('-----------------------------------------------------------')
    $logger.Info('追加したプリンター　: ' + $config.env.printerSetting.printerName)
    $logger.Info('デフォルトプリンター: ' + (Get-CimInstance -ClassName Win32_Printer | Whrere-Object {$_.Default -eq True}.Name))
    $logger.Info('プリンター一覧　　　: ' + ((Get-Printer).Name -replace '`n', ', '))
} catch {
}

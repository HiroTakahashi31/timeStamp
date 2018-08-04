# ログイン時のメールアドレス
# パスワード
# 同じ階層にある日報のExcel名

###############################################
Get-Content setting.ini | %{$hash += ConvertFrom-StringData $_}

$email = $hash.user
$pass = $hash.pass
$excelName = "\" + $hash.excelName
###############################################

# 処理
python .\dateTime.py $email $pass $excelName.Replace("\","./")

$scriptPath = $MyInvocation.MyCommand.Path
$path = Split-Path -Parent $scriptPath

# Excelを操作する為の宣言
$excel = New-Object -ComObject Excel.Application

# 可視化しない
$excel.Visible = $false

# コピー元のExcel
$SourceBook = $excel.Workbooks.open($path + "\AAA.xls")

# コピー先のExcel
$TargetBook = $excel.Workbooks.open($path + $excelName)

# コピー先シートの指定
$Targetsheet = $TargetBook.worksheets.item(1)

##### ループ処理
$row = 11
$startCell = ""
$endCell = ""
for ($i = 0; $i -lt 6; $i++) {
    for ($j = 0; $j -lt 7; $j++) {

        switch ($j) {
            0 {
                $startCell = "C" + [string]$row
                $endCell = "E" + [string]$row
            }
            1 {
                $startCell = "J" + [string]$row
                $endCell = "L" + [string]$row
            }
            2 {
                $startCell = "Q" + [string]$row
                $endCell = "S" + [string]$row
            }
            3 {
                $startCell = "X" + [string]$row
                $endCell = "Z" + [string]$row
            }
            4 {
                $startCell = "AE" + [string]$row
                $endCell = "AG" + [string]$row
            }
            5 {
                $startCell = "AL" + [string]$row
                $endCell = "AN" + [string]$row
            }
            6 {
                $startCell = "AS" + [string]$row
                $endCell = "AU" + [string]$row
            }
        }

        function copy_paste ($cell) {
            # コピー範囲の指定
            $SourceRange = $SourceBook.WorkSheets.item(1).Range($cell)

            # コピー
            $SourceRange.copy()

            # 貼り付け開始位置の指定
            $Range = $Targetsheet.Range($cell)

            # 貼り付け
            $Targetsheet.paste($Range)

            $Range.NumberFormatLocal = '@'
        }

        copy_paste $startCell
        copy_paste $endCell  
    }
    $row = $row + 20
}

# 上書き保存
$SourceBook.Save()
$TargetBook.Save()

# Excelを閉じる
$excel.Quit()

# プロセスを解放する
$excel = $null
[GC]::Collect()

Remove-Item .\AAA.xls
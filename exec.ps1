Param(
    [string]$command,
    [string]$inputDir,
    [string]$outputDir
)
# .\exec.ps1 -command combine -inputDir src -outputDir out

# xlsxにVBコンポーネントをインポートしxlxmを出力
function _combine([string]$inputDir, [string]$outputDir) {
    $xlsxFiles = Get-ChildItem "$inputDir\*.*" -Include *.xlsx
    $vbcompFIles = Get-ChildItem "$inputDir\*.*" -Include *.bas,*.frm,*.cls
    [System.Console]::WriteLine("input xlsx file : $($xlsxFiles)")
    [System.Console]::WriteLine("input vbcomponents file : $($vbcompFiles)")

    if ($xlsxFiles.Count -eq 0) {
        [System.Console]::WriteLine("No excel file in $inputDir")
        return
    }
    if ($vbcompFiles.Count -eq 0) {
        [System.Console]::WriteLine("No vbcomponent file in $inputDir")
        return
    }

    # Excelオブジェクトを取得
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false

    # Excelブックを開く
    $book = $excel.Workbooks.Open($xlsxFiles[0])

    # ブックにVBコンポーネントをインポート
    foreach ($vbComponent in $vbcompFiles) {
        $book.VBProject.VBComponents.Import($vbComponent)
    }

    $destPath = Resolve-Path $outputDir
    $destFile = [System.IO.Path]::GetFileNameWithoutExtension($xlsxFiles[0]);
    # ブックを別名保存
    $book.SaveAs("$destPath\$destFile", 52)

    # Excelオブジェクトを破棄
    $excel.Quit()
    $excel = $null
    [GC]::Collect()
}

# xlsmからVBコンポーネントをエクスポート
function _decombine([string]$inputDir, [string]$outputDir) {
    $xlsmFiles = Get-ChildItem -Path $inputDir -Filter "*.xlsm"
    if ($xlsmFiles.Count -eq 0) {
        [System.Console]::WriteLine("No file in $inputDir")
        return
    }
    [System.Console]::WriteLine("input xlsm file : $inputDir\$($xlsmFiles[0])")

    # Excelオブジェクトを取得
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false

    # Excelブックを開く
    $absPath = Resolve-Path "$inputDir\$($xlsmFiles[0])"
    $book = $excel.Workbooks.Open($absPath)

    $destPath = Resolve-Path $outputDir
    # VBコンポーネントをエクスポートする
    foreach($module in $book.VBProject.VBComponents) {
        $ext = ""
        switch($module.Type) {
            1 { $ext = "bas"}
            2 { $ext = "frm" }
            3 { $ext = "cls" }
        }
        if ($ext -ne "") {
            $module.Export("$destPath\$($module.Name).$ext")
        }
    }

    # Excelオブジェクトを破棄
    $excel.Quit()
    $excel = $null
    [GC]::Collect()
}

# xlsmからVBコンポーネントをクリア
function _clear([string]$inputDir, [string]$outputDir) {
    $xlsmFiles = Get-ChildItem -Path $inputDir
    if ($xlsmFiles.Count -eq 0) {
        [System.Console]::WriteLine("No file in $inputDir")
        return
    }
    [System.Console]::WriteLine("input xlsm file : $inputDir\$($xlsmFiles[0])")

    # Excelオブジェクトを取得
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false

    # Excelブックを開く
    $absPath = Resolve-Path "$inputDir\$($xlsmFiles[0])"
    $book = $excel.Workbooks.Open($absPath)

    $destPath = Resolve-Path $outputDir
    # VBコンポーネントをリムーブする
    foreach($module in $book.VBProject.VBComponents) {
        $ext = ""
        switch($module.Type) {
            1 { $ext = "bas"}
            2 { $ext = "frm" }
            3 { $ext = "cls" }
        }
        if ($ext -ne "") {
            $book.VBProject.VBComponents.Remove($module)
        }
    }

    $destFile = [System.IO.Path]::GetFileNameWithoutExtension($absPath);
    # ブックを別名保存
    $book.SaveAs("$destPath\$destFile", 61)

    # Excelオブジェクトを破棄
    $excel.Quit()
    $excel = $null
    [GC]::Collect()
}

# 処理本体
switch ($command) {
    "Combine" {
        [System.Console]::WriteLine("Combine from $inputDir to $outputDir")
        _combine $inputDir $outputDir
    }
    "Decombine" {
        [System.Console]::WriteLine("Decombine from $inputDir to $outputDir")
        _decombine $inputDir $outputDir
        _clear $inputDir $outputDir
    }
    "Clear" {
        [System.Console]::WriteLine("Clear from $inputDir to $outputDir")
        _clear $inputDir $outputDir
    }
    Default {}
}

Read-Host "Press the Enter key to exit"

param([string]$JsonPath)
# Grava o resultado da automacao numa COPIA da planilha usando o proprio Excel
# (preserva formulas, validacoes, links e formatacao).
$ErrorActionPreference = "Stop"
function S($v) { if ($null -eq $v -or "$v" -eq "") { return $null } return [string]$v }
$d = Get-Content -Raw -Encoding UTF8 $JsonPath | ConvertFrom-Json
Copy-Item -LiteralPath $d.origem -Destination $d.saida

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.AskToUpdateLinks = $false
try {
    $wb = $xl.Workbooks.Open($d.saida, 0)   # 0 = nao atualizar links externos
    $xl.Calculation = -4135                 # manual durante a escrita
    $ws = $wb.Worksheets.Item("RELAT" + [char]0x00D3 + "RIO")

    $linhas = @($d.linhas)
    $n = $linhas.Count
    $first = $linhas[0].linha
    $last = $linhas[$n - 1].linha
    $rows = $last - $first + 1

    # cabecalhos O:T
    $hdr = @(("A ORIGINAL (antes da automa" + [char]0x00E7 + [char]0x00E3 + "o)"), "AUTO - CATEGORIA", ("AUTO - DECIS" + [char]0x00C3 + "O"), "AUTO - MOTIVO", "AUTO - PROCESSO ACHADO NO TEOR", "AUTO - CLIENTE DO PROCESSO DO TEOR")
    for ($c = 0; $c -lt 6; $c++) { $ws.Cells.Item(1, 15 + $c).Value2 = $hdr[$c] }

    $arrA = New-Object 'object[,]' $rows, 1
    $arrO = New-Object 'object[,]' $rows, 6
    # mantem o que ja existe em A para linhas sem processo
    $oldA = $ws.Range("A$first`:A$last").Value2
    for ($i = 0; $i -lt $rows; $i++) { $v = $oldA[($i + 1), 1]; if ($null -ne $v) { $arrA[$i, 0] = [string]$v } }

    $campos = @("A_orig", "cat", "dec", "motivo", "p_teor", "c_teor")
    foreach ($l in $linhas) {
        $i = [int]($l.linha - $first)
        $v = $l.A
        if ($null -ne $v -and "$v" -ne "") { $arrA[$i, 0] = [string]$v } else { $arrA[$i, 0] = $null }
        for ($c = 0; $c -lt 6; $c++) {
            $v = $l.($campos[$c])
            if ($null -ne $v -and "$v" -ne "") { $arrO[$i, $c] = [string]$v }
        }
        if ($l.D) { $ws.Cells.Item([int]$l.linha, 4).Value2 = [string]$l.D }
    }
    # grava celula a celula (o PowerShell 5.1 falha ao passar o bloco inteiro para esta planilha)
    $xl.ScreenUpdating = $false
    $cells = $ws.Cells
    for ($i = 0; $i -lt $rows; $i++) {
        $r = $first + $i
        $cells.Item($r, 1).Value2 = $arrA[$i, 0]
        for ($c = 0; $c -lt 6; $c++) {
            $v = $arrO[$i, $c]
            if ($null -ne $v) { $cells.Item($r, 15 + $c).Value2 = $v }
        }
    }
    $ws.Range("O1:T1").Font.Bold = $true
    $ws.Range("O1:T1").Interior.Color = 0xCCE5FF

    # aba de resumo
    $rs = $wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count))
    $rs.Name = "RESUMO AUTOMACAO"
    $rs.Cells.Item(1, 1).Value2 = "CATEGORIA"
    $rs.Cells.Item(1, 2).Value2 = "DECISAO"
    $rs.Cells.Item(1, 3).Value2 = "QTD"
    $r = 2
    foreach ($x in $d.resumo) {
        $rs.Cells.Item($r, 1).Value2 = $x[0]; $rs.Cells.Item($r, 2).Value2 = $x[1]; $rs.Cells.Item($r, 3).Formula = [string]$x[2]; $r++
    }
    $rs.Range("A1:C1").Font.Bold = $true
    [void]$rs.Columns.Item("A:C").AutoFit()

    $xl.Calculation = -4105                 # automatico
    $xl.Calculate()
    $wb.Save()
    $wb.Close($false)
}
finally {
    # em caso de erro, fecha a copia sem salvar
    if ($wb) { try { $wb.Close($false) } catch {} }
    # esta instancia foi criada pelo script (New-Object abre um Excel separado), entao e seguro fechar
    $xl.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}


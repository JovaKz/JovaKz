@echo off
setlocal
chcp 65001 >nul
title Conversor Word para PDF

echo ========================================================
echo          CONVERSOR DE WORD PARA PDF EM LOTE
echo ========================================================
echo.

if not "%~1"=="" (
    set "DOCX2PDF_PATH=%~1"
) else (
    set /p "DOCX2PDF_PATH=Digite ou cole o caminho da pasta com os arquivos Word: "
)

if not defined DOCX2PDF_PATH (
    echo.
    echo ERRO: Nenhuma pasta foi informada.
    pause
    exit /b 1
)

echo.
echo Iniciando a conversao...
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "& { $ErrorActionPreference = 'Stop'; $folder = [Environment]::GetEnvironmentVariable('DOCX2PDF_PATH'); $folder = [IO.Path]::GetFullPath($folder.Trim().Trim([char]34)); if (-not (Test-Path -LiteralPath $folder -PathType Container)) { throw ('A pasta informada nao existe: ' + $folder) }; $files = @(Get-ChildItem -LiteralPath $folder -File | Where-Object { ($_.Extension -in '.doc', '.docx') -and ($_.Name -notlike '~$*') }); if ($files.Count -eq 0) { Write-Host 'Nenhum arquivo .doc ou .docx foi encontrado.' -ForegroundColor Yellow; exit 0 }; $word = $null; $converted = 0; $failed = 0; try { $word = New-Object -ComObject Word.Application; $word.Visible = $false; $word.DisplayAlerts = 0; foreach ($file in $files) { $document = $null; try { Write-Host ('Convertendo: ' + $file.Name); $document = $word.Documents.Open($file.FullName, $false, $true); $pdf = [IO.Path]::ChangeExtension($file.FullName, '.pdf'); $document.ExportAsFixedFormat($pdf, 17); $converted++ } catch { $failed++; Write-Host ('Falha em ' + $file.Name + ': ' + $_.Exception.Message) -ForegroundColor Red } finally { if ($null -ne $document) { $document.Close($false); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($document) } } } } finally { if ($null -ne $word) { $word.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word) }; [GC]::Collect(); [GC]::WaitForPendingFinalizers() }; Write-Host ''; Write-Host ('Conversao concluida. PDFs criados: ' + $converted) -ForegroundColor Green; if ($failed -gt 0) { Write-Host ('Arquivos com falha: ' + $failed) -ForegroundColor Yellow }; Write-Host ('Pasta de destino: ' + $folder) -ForegroundColor Cyan }"

if errorlevel 1 (
    echo.
    echo A conversao nao foi concluida. Verifique se o Microsoft Word esta instalado
    echo e se o caminho informado esta correto.
)

echo.
pause
endlocal

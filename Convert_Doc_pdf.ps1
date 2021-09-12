#requires -version 2
param
(
    [string]$path = $null
)

# Если путь не задан ищем в текущем каталоге
if (!$path)
{
    $path = Get-Location
    Write-Warning "Путь не задан, по умолчанию используется текущий путь: $path"
}
# Ищем .doc файлы
$docs = Get-ChildItem $path -Filter "*.doc"
if (!$docs)
{
    Write-Error -Category ObjectNotFound `
        -Message "В указанном каталоге .doc файлы не найдены" `
        -TargetObject $path
    exit
}
$pdf_path = $path + "\PDF"
if (!(Test-Path $pdf_path))
{
    Write-Verbose "Создаём каталог $pdf_path`n" -Verbose
    New-Item -ItemType directory -Path $pdf_path | Out-Null
}
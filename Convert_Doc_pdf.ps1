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
# Количество .doc файлов
$count = $docs.Count

# Создаём объект Word
$word = New-Object -ComObject Word.Application
$word.Visible = $false

foreach ($file in $docs) {
    # Номер и имя обрабатываемого файла
    $name = $file.Name
    ++$number
    Write-Verbose "$name `t`t (файл $number из $count)" -Verbose

    # Открываем файл
    $doc = $word.Documents.Open($file.FullName)
    
    # Задаём имя и путь к pdf файлу
    $pdf_file = $pdf_path + "\" + $doc.Name
    
    # Меняем расширение (иначе будет pdf-файл с расширением doc)
    $pdf_file = [System.IO.Path]::ChangeExtension($pdf_file, '.pdf')
    
    # Конвертируем
    $doc.ExportAsFixedFormat($pdf_file, 'wdExportFormatPDF', $false, 1)
    
    # Закрываем Файл
    $doc.Close()
}

# Закрываем Word
$word.Quit()
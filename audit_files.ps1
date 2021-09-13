#Задаем период, в течении которого мы будем запускать один раз скрипт, и искать нужные нам события. Здесь период задан - 1 час. Т.е. проверяются все события за последний час.
$time =  (get-date) - (new-timespan -min 60)

#$BodyL - переменная для записи в лог-файл
$BodyL = ""

#$Body - переменная, в которую записываются ВСЕ события с нужным ID. 
$Body = Get-WinEvent -FilterHashtable @{LogName="Security";ID=4663;StartTime=$Time}|where{ ([xml]$_.ToXml()).Event.EventData.Data |where {$_.name -eq "ObjectName"}|where {($_.'#text') -notmatch ".*tmp"} |where {($_.'#text') -notmatch ".*~lock*"}|where {($_.'#text') -notmatch ".*~$*"}} |select TimeCreated, @{n="Файл_";e={([xml]$_.ToXml()).Event.EventData.Data | ? {$_.Name -eq "ObjectName"} | %{$_.'#text'}}},@{n="Пользователь_";e={([xml]$_.ToXml()).Event.EventData.Data | ? {$_.Name -eq "SubjectUserName"} | %{$_.'#text'}}} |sort TimeCreated

#Далее в цикле проверяем содержит ли событие определенное слово (к примеру название шары, например: Secret)	
foreach ($bod in $body){
	if ($Body -match ".*Secret*"){

#Если содержит, то пишем в переменную $BodyL данные в первую строчку: время, полный путь файла, имя пользователя. И #в конце строчки переводим каретку на новую строчку, чтобы писать следующую строчку с данными о новом файле. И так #до тех пор, пока переменная $BodyL не будет содержать в себе все данные о доступах к файлам пользователей.

		$BodyL=$BodyL+$Bod.TimeCreated+"`t"+$Bod.Файл_+"`t"+$Bod.Пользователь_+"`n"
	}
}

#Т.к. записей может быть очень много (в зависимости от активности использования общего ресурса), то лучше разбить лог #на дни. Каждый день - новый лог. Имя лога состоит из Названия AccessFile и даты: день, месяц, год.
$Day = $time.day
$Month = $Time.Month
$Year = $Time.Year
$name = "AccessFile-"+$Day+"-"+$Month+"-"+$Year+".txt"
$Outfile = "\serverServerLogFilesAccessFileLog"+$name

#Пишем нашу переменную со всеми данными за последний час в лог-файл.
$BodyL | out-file $Outfile -append
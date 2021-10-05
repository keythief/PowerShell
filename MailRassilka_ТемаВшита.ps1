#Выбор файлов для письма
Add-Type -AssemblyName System.Windows.Forms
$Explorer = New-Object System.Windows.Forms.OpenFileDialog

$Explorer.Title = "Выбор файла с адресами"
#$Explorer.filter = "Txt (*.txt)| *.txt"
[void]$Explorer.ShowDialog()
$AddressList = $Explorer.FileName; # выбираем файл с адресами

$Explorer.Title = "Выбор файла для вложения"
$Explorer.Multiselect = $True;
[void]$Explorer.ShowDialog()
$Attachments = $Explorer.FileNames; # выбираем файл для вложения

$Explorer.Title = "Выбор файл тела письма"
[void]$Explorer.ShowDialog()
$Bodu = $Explorer.FileName; # выбираем файл для тела письма
$Explorer.Dispose() # удаляем обьекты из рабочего пространства


#Обработка выбранных файлов и отправка
Get-Content $AddressList| Foreach-Object {
 Send-MailMessage `
-From kireenko-as@norebo.ru `
-To $_ `
-Subject "Тема письма тут пишется"  `
-Body ([string](Get-Content $Bodu)) `
-SmtpServer 'mail.norebo.ru' `
-Attachments $Attachments.Split(' ') `
-BodyAsHTML `
-Encoding utf8 `
-DeliveryNotificationOption OnFailure, OnSuccess
}
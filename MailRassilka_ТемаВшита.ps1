#����� ������ ��� ������
Add-Type -AssemblyName System.Windows.Forms
$Explorer = New-Object System.Windows.Forms.OpenFileDialog

$Explorer.Title = "����� ����� � ��������"
#$Explorer.filter = "Txt (*.txt)| *.txt"
[void]$Explorer.ShowDialog()
$AddressList = $Explorer.FileName; # �������� ���� � ��������

$Explorer.Title = "����� ����� ��� ��������"
$Explorer.Multiselect = $True;
[void]$Explorer.ShowDialog()
$Attachments = $Explorer.FileNames; # �������� ���� ��� ��������

$Explorer.Title = "����� ���� ���� ������"
[void]$Explorer.ShowDialog()
$Bodu = $Explorer.FileName; # �������� ���� ��� ���� ������
$Explorer.Dispose() # ������� ������� �� �������� ������������


#��������� ��������� ������ � ��������
Get-Content $AddressList| Foreach-Object {
 Send-MailMessage `
-From kireenko-as@norebo.ru `
-To $_ `
-Subject "���� ������ ��� �������"  `
-Body ([string](Get-Content $Bodu)) `
-SmtpServer 'mail.norebo.ru' `
-Attachments $Attachments.Split(' ') `
-BodyAsHTML `
-Encoding utf8 `
-DeliveryNotificationOption OnFailure, OnSuccess
}
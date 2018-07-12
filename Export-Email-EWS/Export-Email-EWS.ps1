param (
[Switch]$report=$False,
[Switch]$exportEML=$false,
[Switch]$ExportMsgFormatTypeTXT=$false,
$mailbox="Sunil.chauhan@Domain.com" ,
$itemsView=1000 ,                                       			                  #No of Items to Export from the Folder
$dir="D:\WindowsPowerShell\Download-Email\" ,                                   #Folder Path to save items in
$userName = "userid@Domain.com" ,
$password ="password"
)

$uri=[system.URI] "https://outlook.office365.com/ews/exchange.asmx"
$dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"
Import-Module $dllpath

#Setup EWS Service Client
$ExchangeVersion=[Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP2
$service=New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)
$service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials -ArgumentList $userName, $password
$service.url = $uri

#Entering into Mailbox
$service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId `
([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SMTPAddress,$Mailbox);  

$MailboxRootid= new-object Microsoft.Exchange.WebServices.Data.FolderId `
([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,$mailbox)
$MailboxRoot=[Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$MailboxRootid)

#Loading All Folders Under Recoverable Items Root
$FolderList = new-object Microsoft.Exchange.WebServices.Data.FolderView(50)
$FolderList.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep
$findFolderResults = $MailboxRoot.FindFolders($FolderList)
$findFolderResults | ? {$_.FolderClass -match "IPF.Note"} |ft DisplayName, TotalCount, FolderClass
Write-Host "Type the folder name to export" -f yellow
$folder=read-host "Folder to Export:"
$folderID = ($findFolderResults | ? {$_.DisplayName -eq $folder}).ID

if ($ExportMsgFormatTypeTXT) {
#Setup property set for email
$psPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet `
([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
$psPropset.RequestedBodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::text

#Loading Email Items
$view = New-Object Microsoft.Exchange.WebServices.Data.ItemView -ArgumentList $ItemsView
$view.PropertySet = $propertyset
$items=$service.FindItems($folderid,$View)
$items.load($psPropset)

$items.items | select Subject,LastModifiedTime,DateTimeReceived,Sender,{$_.ToRecipients}
	foreach ($item in $items.items)      
			{
			Write-Host "Writing Email to File:" $item.Subject
			$fileName = $Dir + $($item.Subject).replace(":","-") + ".txt"
			$psPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.ItemSchema]::MimeContent)
			$item.load($psPropset)
			$Email = new-object System.IO.FileStream($fileName, [System.IO.FileMode]::Create)
			$Email.Write($Item.MimeContent.Content, 0,$Item.MimeContent.Content.Length)
			$Email.Close()
			}
} else {
#Setup property set for email
$psPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet `
([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
$psPropset.RequestedBodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::HTML

#Loading Email Items
$view = New-Object Microsoft.Exchange.WebServices.Data.ItemView -ArgumentList $ItemsView
$view.PropertySet = $propertyset
$items=$service.FindItems($folderid,$View)
$items.load($psPropset)
     }
	 
if ($report) {
		$items.items | ft Subject,LastModifiedTime,DateTimeReceived,Sender,{$_.ToRecipients}
             }
	   
if ($exportEML) {
        foreach ($item in $items.items)
        {
        Write-Host "Writing Email to File:" $item.Subject
        $fileName = $Dir + $($item.Subject).replace(":","-") + ".eml"
        $psPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.ItemSchema]::MimeContent)
        $item.load($psPropset)
        $Email = new-object System.IO.FileStream($fileName, [System.IO.FileMode]::Create)
        $Email.Write($Item.MimeContent.Content, 0,$Item.MimeContent.Content.Length)
        $Email.Close()
        }
 }

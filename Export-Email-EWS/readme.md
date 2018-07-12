Export or report email from a specific folder.
Below CMD export email in EML format, available Folder list will be show

[D:\WindowsPowerShell\Download-Email]
SUNIL:65 >.\Download-Email-EWS.ps1 -exportEML -userName admin@domain.com -password Password -mailbox user1@domain.com

|DisplayName |  TotalCount | FolderClass
|----------- | ---------- | -----------
|Archive             |   0 |IPF.Note
|Deleted Items       |   1  |IPF.Note
|Drafts             |    0 |IPF.Note
|Inbox             |    60 |IPF.Note
|Junk Email       |      0 |IPF.Note
|Outbox          |       0 |IPF.Note
|Sent Items     |        0 |IPF.Note

Type the folder name to export
Folder to Export:: Inbox

Below CMD will exprot the msg in txt format.

[D:\WindowsPowerShell\Download-Email]
SUNIL:65 >.\Download-Email-EWS.ps1 -ExportMsgFormatTypeTXT -userName admin@domain.com -password Password -mailbox user1@domain.com

DisplayName   TotalCount FolderClass
-----------   ---------- -----------
Archive                0 IPF.Note
Deleted Items          1 IPF.Note
Drafts                 0 IPF.Note
Inbox                 60 IPF.Note
Junk Email             0 IPF.Note
Outbox                 0 IPF.Note
Sent Items             0 IPF.Note

Type the folder name to export
Folder to Export:: Inbox

Below CMD will generate a report of the available msg in the mailbox and will show the display the same in the powershell.

[D:\WindowsPowerShell\Download-Email]
SUNIL:65 >.\Download-Email-EWS.ps1 -report -userName admin@domain.com -password Password -mailbox user1@domain.com

DisplayName   TotalCount FolderClass
-----------   ---------- -----------
Archive                0 IPF.Note
Deleted Items          1 IPF.Note
Drafts                 0 IPF.Note
Inbox                 60 IPF.Note
Junk Email             0 IPF.Note
Outbox                 0 IPF.Note
Sent Items             0 IPF.Note

Type the folder name to export
Folder to Export:: Inbox


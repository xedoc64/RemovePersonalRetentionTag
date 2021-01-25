## Synopsis

RemovePersonalRetentionTag.exe will remove personal retention tags from all folders in an Exchange Server mailbox.

## Why should I use this

If you had run into a MRM issue or need to remove personal tags from a mailbox you could only remove them if the policy tags
will be removed from the entire Exchange Org. This utility can remove one or more personal tags from a mailbox without 
touching the policy tags in the Org.

This utility is designed for Exchange administrators.

## Installation

Simply copy all files to a location where you are allowed to run it and of course the Exchange servers are reachable.

## Requirements
* Exchange Server 2016/2019 (Tested with Exchange 2016 CU16, maybe it will work with Exchange 2007/2010 as well. It should work in Exchange 2013.)
* Application Impersonation Rights if you want to change items on other mailboxes than yours
* Microsoft.Exchange.WebServices.dll, log4net.dll (are provided in the repository and also in the binaries)
* For Exchange Online you need to enable Basic Authentification because of EWS (Set the retention tags via Graph isn't currently possible)

## Usage
```
RemovePersonalRetentionTag.exe -mailbox "user@example.com" [-logonly] [-foldername "Inbox"]  [-ignorecertificate] [-url "https://server/EWS/Exchange.asmx"] [-user "user@example.com"] [-password "Pa$$w0rd"] [-impersonate] [-retentionid "a7966968-dadf-4df7-ae87-4482686b4634" [-archive]
```

## Examples
```
RemovePersonalRetentionTag.exe -mailbox "user@example.com" -logonly -impersonate
```
This will log all folder which have an retention tag set. It will be logged per default into the log folder. You can change the logging behaviour in the file "RemovePersonalRetentionTag.exe.config"


```
RemovePersonalRetentionTag.exe -mailbox "user@example.com" -impersonate
```
This will remove all personal tags from folders inside the mailbox.

```
RemovePersonalRetentionTag.exe -mailbox "user@example.com" -impersonate -retentionid "a7966968-dadf-4df7-ae87-4482686b4634"
```
This will remove the tag with the retention id "a7966968-dadf-4df7-ae87-4482686b4634" from a folder. Other tags will not be removed.



# Parameters
* mandatory: -mailbox "user@domain.com"

Mailbox which you want to alter.

* optional: -logonly

Items will only be logged.

* optional: -foldername "Inbox"

Will filter the items to the Folderpath. Uses Contains, so "Inbox" would also include "Inbox\bla".

* optional: -ignorecertificate

Ignore certificate errors. Interesting if you connect to a lab config with self signed certificate.

* optional: -impersonate

If you want to alter a other mailbox than yours set this parameter.

* optional: -user "user@domain.com"

If set together with -password this credentials would be used. Elsewhere the credentials from your session will be used.

* optional: -password "Pa$$w0rd"

Password for the user you specified with -user

* optional: -url "https://server/EWS/Exchange.asmx"

If you set an specific URL this URL will be used instead of autodiscover. Should be used with -ignorecertificate if your CN is not in the certficate.

* optional: -allowredirection

If your autodiscover redirects you the default behaviour is to quit the connection. With this parameter you will be connected anyhow (Hint: O365)

* optional: -archive

Search for folders inside the archive instead of the mailbox

* optional: -retentionid

One or more retention ids separated with a ",". Limit the all actions only to this ids. You can get the ids with the Exchange powershell


## License

MIT License

# Scripts
**DeleteItemsMCv2.ps1** - Uses Exchange Web Services to look up items of a specific Message Class and delete them. The script uses Users.csv to read the mailbox to run against.

Examles:

With Impersonation:
```
#Service Account Credentials
$cred =Get-Credential
.\DeleteItemsMCv2.ps1 -Credentials $cred -EwsUrl "https://mail.contoso.com/ews/exchange.asmx" -IgnoreSSLCertificate -MessageClass IPM.Note.xyz -Impersonate
```

Without Impersonation:
```
#Mailbox Credentials
$cred =Get-Credential
.\DeleteItemsMCv2.ps1 -Credentials $cred -EwsUrl "https://mail.contoso.com/ews/exchange.asmx" -IgnoreSSLCertificate -MessageClass IPM.Note.xyz
```

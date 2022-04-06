# SPO User Profile Bulk update
This is sample to demonstrate how to bulk update User Profile service properties

## PowerShell
"PS" folder includes a script which used PnP PowerShell

## CSharp
"SPOUserProfileBulkUpdate" folder includes a sample which used PnP.Core, PnP.Core.Auth and PnP.Framework

### Instruction
Update the following code insdie program.cs
```CSharp
    string certFile = @"[pfx-path]";
    string clientId = "[client-id]";
    string tenantId = "[tenant-id]";
    X509Certificate2 certficate = new X509Certificate2(certFile, "[password]");

```
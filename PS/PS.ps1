#
# Script.ps1
#
Connect-PnPOnline -Url "[site-url]" `
    -ClientId [client-id] `
    -Tenant [tenant-id] `
    -CertificatePath "[certificate-file-path]" `
    -CertificatePassword (ConvertTo-SecureString -String "[password]" -AsPlainText -Force)

# Queue a job
$jsonFile = "https://m365x725618.sharepoint.com/sites/FrankCommunication1/Shared Documents/UserProfileValues.json";
New-PnPUPABulkImportJob -Url $jsonFile -IdProperty "IdName" -UserProfilePropertyMapping @{"EmployeeId" = "EmployeeId" }

# Check the status
Get-PnPUPABulkImportStatus
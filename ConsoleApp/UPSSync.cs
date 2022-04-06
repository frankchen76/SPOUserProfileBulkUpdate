using Microsoft.Online.SharePoint.TenantManagement;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SPOUserProfileBulkUpdate
{
    internal class UPSSync
    {
        public void QueueUserProfileJob(ClientContext csomContext, string jsonFileUrl)
        {
            // Create an instance of the Office 365 Tenant object. Loading this object is not technically needed for this operation.
            Office365Tenant tenant = new Office365Tenant(csomContext);
            csomContext.Load(tenant);
            csomContext.ExecuteQuery();

            // Type of user identifier ["PrincipalName", "Email", "CloudId"] in the
            // user profile service. In this case, we use Email as the identifier at the UPA storage
            ImportProfilePropertiesUserIdType userIdType =
                  ImportProfilePropertiesUserIdType.PrincipalName;

            // Name of the user identifier property within the JSON file
            var userLookupKey = "IdName";

            var propertyMap = new System.Collections.Generic.Dictionary<string, string>();

            // The key is the property in the JSON file
            // The value is the user profile property Name in the user profile service
            // Notice that we have 2 custom properties in UPA called 'City' and 'OfficeCode'
            propertyMap.Add("EmployeeId", "EmployeeId");

            // Returns a GUID that can be used to check the status of the execution and the end results
            var workItemId = tenant.QueueImportProfileProperties(userIdType, userLookupKey, propertyMap, jsonFileUrl);

            Console.WriteLine($"Job Id: {workItemId}");

            csomContext.ExecuteQuery();
        }
    }
}

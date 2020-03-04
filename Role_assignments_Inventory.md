# Exchange Role Assignments Inventory

The script lists all Exchange management roles assigned in your organization. Only roles with active assignments are returned by default, along with details about the corresponding user or group object. The following parameters are supported:
 
* __IncludeRoleGroups__ - used to include "parent" role group entries in the output. Without this parameter, only the "expanded" membership of each role group is returned. 
* __IncludeUnassignedRoleGroups__  - used to include role groups that have no corresponding role assignments in the output. 
* __IncludeDelegatingAssingments__ - used to include delegating role assignments in the output. Any delegating role assignments for the "Organization management" role group are ignored. 

The script does not handle connectivity to Exchange Online, due to the variety of methods now available. If any existing session is detected, it will be used to run the script, otherwise an error will be thrown. Exchange on-premises is also supported, as is running the script in the EMS, but those scenarios are not extensively tested. The script requires PowerShell v3 at minimum.  

By default, the script will generate a separate entry for each user or group and write the output to a CSV file. You can then sort, filter or group the entries via Excel or similar tools. In addition, the script will also provide a "transformed" version of the output, which is grouped by each security principal, and will aim to provide an unique identifier for each object. Multiple role assignments for the same user are concatenated together in the "Roles" field. This "transformed" output is returned to the console window and is stored in the $varRoleAssignments global variable.
 
Here is an example on how to invoke the script with all parameters:
```PowerShell
.\Role_assignments_Inventory.ps1 -IncludeDelegatingAssingments -IncludeRoleGroups -IncludeUnassignedRoleGroups
```

Additional information about the script can be found in [this article](https://practical365.com/exchange-online/how-to-report-on-exchange-rbac-assignments/).
 
Questions and feedback are welcome.

# Exchange Online Group membership inventory with support for nested groups

The script enumerates all groups of the specified type(s) in the organization and returns their membership. The following parameters are supported:
 
* __IncludeDGs__: Specify whether to include Distribution Groups and Mail-enabled Security Groups 
* __IncludeDynamicDGs__: Specify whether to include Dynamic Distribution Groups 
* __IncludeO365Groups__: Specify whether to include Office 365 Groups 
* __IncludeAll__: Use to include all of the above 
* __RecursiveOutput__: Specify whether to expand the membership of any nested groups 
* __RecursiveOutputListGroups__: Specify whether to include an entry for any nested groups in the output, or just their expanded member objects 

As the full list of groups in the organization needs to be cycled in order to generate the report, the script will use Invoke-Command in order to get a minimum set of attributes returned. In case you run into throttling or connectivity errors, consider adjusting the artifical delay added on line 21.

The script does not handle connectivity to Exchange Online, due to the variety of methods now available. If any existing session is detected, it will be used to run the script. Otherwise an attempt to connect to ExO via Basic auth will be performed, and failing that an error will be thrown, halting the script execution. Exchange on-premises is also supported, as is running the script in the EMS, but those scenarios are not extensively tested. The script requires PowerShell v3 at minimum.
 
Here are some example uses of the script. To get a list of all Distribution Groups and Mail-Enabled Security Groups in the company and their membership, simply run the script without any parameters. If you want to include membership of any nested groups as well, use the -RecursiveOutput parameter:
```PowerShell
.\DG_members_recursive.ps1 -RecursiveOutput
```
To include all group types recognized by Exchange Online, use the -IncludeAll parameter:
```PowerShell
.\DG_members_recursive.ps1 -IncludeAll
```
To get a report for all group objects, including membership of any nested groups and also return entries for the nested groups themselves, use:
```PowerShell
.\DG_members_recursive.ps1 -IncludeAll -RecursiveOutput -RecursiveOutputListGroups
```
By default the script outputs the results to a CSV file and also stores them in the $varGroupMembership variable. If you want to generate a CSV file for each group individually, uncomment line 86 (NOT tested with complex scenarios, use at your own risk!).

The screenshot below illustrates the different types of output you can get depending on the parameters used. The top example list just direct members of the DG group, the middle one includes any members of the nested “empty” group as well, since the -RecursiveOutput switch was used. The bottom example was run with both the -RecursiveOutput and -RecursiveOutputListGroups switches, and thus includes the membership of any nested groups, as well as an entry that lists the address (or identifier) for the actual nested group.



  
Additional information about the script can be found in the built-in help or in [this article](https://practical365.com/blog/how-to-inventory-membership-of-exchange-groups-recursively/).
  
Questions and feedback are welcome.

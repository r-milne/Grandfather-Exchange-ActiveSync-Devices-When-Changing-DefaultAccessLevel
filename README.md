# Grandfather Exchange ActiveSync Devices When Changing DefaultAccessLevel
 Grandfather Exchange ActiveSync Devices When Changing DefaultAccessLevel
 
.SYNOPSIS

Purpose of this script to to assist when changing the default Exchange 2010/2013 ActiveSync DefaultAccessLevel setting from it's default value of Allow to either Quarrantined or Blocked.
This is an issue as devices which were allowed in, will be blocked or quarantined if there are no other device rules that permit them to connect to Exchange. 

One way is to create all of the device rules in advance, prior to making this change. 

Another is to get a listing of all the ActiveSync Devices and grandfather existing devices in.  
This is the purpose of this script. All devices will be considered as valid, and will be set as allowed.  No exceptions. 
Be sure to understand this, and test in your lab prior to running in production. 

 

.DESCRIPTION
This is to grandfather in *ALL* existing devices.

Script will consider devices that are synchronising in the last 30 days as valid.  
If a device has not synchronised in this timeperiod it is not carried over.

You can change the timespan if needed to suit your business requirements. 

    
See this blog post for more details:

https://blog.rmilne.ca/2015/02/25/exchange-activesync-script-to-grandfather-existing-devices

 

You may want to filter out any CAS Test  mailboxes
CAS test mailboxes are created when the in-box script new-TestCasConnectivityUser.ps1 has been executed and then the Test-ActiveSyncConnectivity was used to verify Exchange at some point
However if you choose to do this, then the Test-ActiveSyncConnectivity cmdlets will fail as device type that the cmdlet uses will be blocked. 

An example of a test device would be the following from Get-CASMailbox:

SamAccountName    : extest_4ca5fda1c3994
ServerName        : tail-ca-exch-2
DisplayName       : extest_4ca5fda1c3994
Name              : extest_4ca5fda1c3994
DistinguishedName : CN=extest_4ca5fda1c3994,CN=Users,DC=Tailspintoys,DC=ca

 

.ASSUMPTIONS
 

Script is being executed from an existing Exchange Management Shell sesion.

Script is being executed with sufficient permissions to access Exchange.

You can live with the Write-Host cmdlets :)

You can add your error handling if you need it. 

.VERSION

 
1.0   20-1-2015 -- Initial version

   
This Sample Code is provided for the purpose of illustration only and is not intended to be used in a production environment. 
THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED,
INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE. 
We grant You a nonexclusive, royalty-free right to use and modify the Sample Code and to reproduce and distribute the object code form of the Sample Code,
provided that You agree:
(i) to not use Our name, logo, or trademarks to market Your software product in which the Sample Code is embedded;
(ii) to include a valid copyright notice on Your software product in which the Sample Code is embedded; and
(iii) to indemnify, hold harmless, and defend Us and Our suppliers from and against any claims or lawsuits, including attorneys’ fees, that arise or result from the use or distribution of the Sample Code.
Please note: None of the conditions outlined in the disclaimer above will supercede the terms and conditions contained within the Premier Customer Services Description.
This posting is provided "AS IS" with no warranties, and confers no rights.

Use of included script samples are subject to the terms specified at http://www.microsoft.com/info/cpyright.htm.

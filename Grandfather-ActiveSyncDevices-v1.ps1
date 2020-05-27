<# 

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
    http://blogs.technet.com/b/rmilne/archive/2015/02/25/exchange-activesync-script-to-grandfather-existing-devices.aspx


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
  
	1.0  	20-1-2015 -- Initial version

    
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

#>

# Declare an empty array to hold the output
$Output = @()

# Declare a custom PS object. This is the template that will be copied multiple times. 
# This is used to allow easy manipulation of data from potentially different sources
$TemplateObject = New-Object PSObject | Select DisplayName, AllowedDeviceIDs

# Controls the maximum age of devices that are to be carried over.  If the device has not synchronised sucessfully within this period, it will not be carried over. 
# This is to be specified in days
$IntDeviceAgeLimit = 30 

# Capture the current time information.  Remove $IntDeviceAgeLimit value.
$DeviceAgelimit = (Get-Date).AddDays(-$IntDeviceAgeLimit)

Write-Host "Processing devices that have synched after: $DeviceAgelimit" -ForegroundColor Green # Enable for debugging
Write-Host 

# Add logic to say that only devices in the last 30 days will be proted over.  This is found using the Get-ActiveSyncStatistics cmdlet, will look at the LastSuccessSync attribute.  

# add a conrirm Yes/No for each device, or add logic for if the device type contains....

# Get a list of Mailboxes that have an ActiveSync deviced attached to them.  



# Get a list of all ActiveSync devices that are currently connected in the organisation. 
# In large organisations you may wish to process on a mailbox database by mailbox database basis.  The choice is yours.  The below query can be easily modified.
# For examples on different query syntax, please see this post
# http://blogs.technet.com/b/rmilne/archive/2014/03/24/powershell-filtering-examples.aspx
# 

# Firstly get a list of all mailboxes that have one or more ActiveSync devices associated with them
$EASMailboxes = Get-CASMailbox -Filter {hasactivesyncdevicepartnership -eq $true -and -not displayname -like "CAS_{*"} -ResultSize Unlimited;



# Step through each mailbox and process the devices that are currently in use  
FOREACH ($EASMailbox in $EASMailboxes)
{
    # Make a copy of the TemplateObject.  Then work with the copy...
    $WorkingObject = $TemplateObject | Select-Object * 

    Write-Host "Processing Mailbox: $EASMailbox" -ForegroundColor  Magenta 
    Write-Host 
        # Create null array to store current user's devices
        $EASDeviceIDs = @() 
        # Creat null array to store current user's device statistics.  Needed to work out devices that have connected in the specified time period.
        # Initialise it to zero for each user 
        $EASDevices = @()

        # Retrieve the ActiveSync Device Statistics for the associated user mailbox.  This may be multivalued, hence is stored in an array.  
        # Need to se the .identity attribute else Get-ActiveSyncDeviceStatistics will not have the expected input object and you will get the below error. 
        # Cannot process argument transformation on parameter 'Mailbox'. Cannot convert the "Tailspintoys.ca/Users/user-50" value of type "Deserialized.Microsoft.Exchange.Data.Directory.Management.CASMailbox" to type 
        # "Microsoft.Exchange.Configuration.Tasks.MailboxIdParameter".
        # + CategoryInfo          : InvalidData: (:) [Get-ActiveSyncDeviceStatistics], ParameterBindin...mationException
        # + FullyQualifiedErrorId : ParameterArgumentTransformationError,Get-ActiveSyncDeviceStatistics
        # + PSComputerName        : tail-exch-1.tailspintoys.ca

        $EASDevices = Get-ActiveSyncDeviceStatistics -Mailbox $EASMailbox.Identity 
        
        # Use the information retrieved above to store information one by one about each ActiveSync Device
        FOREACH ($EASDevice in $EASDevices)
        {
            # Write-Host "Processing Device: " $EASDevice.DeviceID " Last Sync Time:" $EASDevice.LastSuccessSync  # Enable for debugging 
            # Logic to carry over or discard particular devices.  
            # First item to evaluate is LastSuccessSync
            IF  ($DeviceAgelimit  -LT $EASDevice.LastSuccessSync)
            {
                   Write-Host "DeviceID: " $EASDevice.DeviceID " has synchronised in the last 30 days, on " $EASDevice.LastSuccessSync
                         

                         $EASDeviceIDs += $EASDevice.DeviceID
            }

             
        }
        Write-Host "For User: $EASMailbox Found " ($EASDeviceIDs).count  " EAS Devices" 

      # Write the collection of devices as allowed for the given user 
      Set-CasMailbox $EASMailbox.Identity -ActiveSyncAllowedDeviceIDs $EASDeviceIDs



    # Build me up buttercup
    # Populate the TemplateObject with the necessary details.
    $WorkingObject.DisplayName      = $EASMailbox
    $WorkingObject.AllowedDeviceIDs =  $EASDeviceIDs
  
    
    # Display output to screen.  REM out if not reqired/wanted 
    # $WorkingObject

    # Append  current results to final output
    $Output += $WorkingObject
    


}


Write-host
Write-Host
Write-host "Processing complete" 
Write-Host 

# Echo to screen
$Output

# Or output to a file.  The below is an example of going toa  CSV file
# The Output.csv file is located in the same folder as the script.  This is the $PWD or Present Working Directory. 
$Output | Export-Csv -Path $PWD\Output.csv -NoTypeInformation
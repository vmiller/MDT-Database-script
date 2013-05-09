# Script to import computer information from a .csv file and
# load the info into the MDT database
#
# Script requires the MDTDB.psm1 writen by Michael Niehaus and assumes
# that it is located in the same directory.
#
# Written by Vaughn Miller

#Import the module that allows us to interact with MDT
Import-Module -name .\MDTDB.psm1

#Open a connection to our database
Connect-MDTDatabase -sqlServer deploy.example.com -instance ADSQL -database msdeploy

#Read in the csv file
$machines = Import-Csv .\computers.csv

#Set up the log file information and initialize counters

$date = get-date -format MMddyyHHmmss
$logfile = ".\logs\Import"+$date+".log"
$validrecords = 0
$invalidrecords = 0


#Loop through each record from the file and process the information

For ($i=1; $i -le $machines.count; $i++)
  {
       $machines[$i-1].mac=$machines[$i-1].mac.ToUpper()
       $macvalid=$TRUE
       $temp = $machines[$i-1].mac.Split(":")
       if (($temp.count -ne 6) -or ($machines[$i-1].mac.length -ne 17))
         { $macvalid = $FALSE
         }  

	$namevalid=$TRUE
	if ($machines[$i-1].name.length -gt 15)
	  { $namevalid=$FALSE
	  }
	
	if ($namevalid -and $macvalid)
	  { 
		$validrecords=$validrecords+1
		$machineid=Get-MDTComputer -macAddress $machines[$i-1].mac
		if ($machineid.id -gt 0) 
			{
			Remove-MDTComputer $machineid.id
			}
		New-MDTComputer -macAddress $machines[$i-1].mac -description $machines[$i-1].name -settings @{
			OSInstall='YES';
	 		OSDComputerName=$machines[$i-1].name;
            AdminPassword=$machines[$i-1].password;
			}
		$machineid=Get-MDTComputer -macAddress $machines[$i-1].mac
		Set-MDTComputerRole $machineid.id $machines[$i-1].role
	   }
	else
	  { #log the invalid record
		$invalidrecords=$invalidrecords+1
		$text = "Invalid Record : "+$i+"  "+$machines[$i-1].mac+" "+$machines[$i-1].name
		$text >> $logfile
	   }
   }
#Some closing information for the logfile

" " >> $logfile
$text="Total records processed = "+($invalidrecords+$validrecords)
$text>>$logfile
$text="Invalid records         = "+$invalidrecords
$text>>$logfile
$text="Valid records           = "+$validrecords
$text>>$logfile

#Let the user know the script is finished and where the log file is
" "
" "
"Script execution Complete"
"Check "+$logfile+" "+"for invalid records"
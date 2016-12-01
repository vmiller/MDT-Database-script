MDT-Database-script
===================
This is an example of a Powershell script that builds on the MDT Powershell module made by Michael Neihaus.  Module was obtained from [Michael's blog post](https://blogs.technet.microsoft.com/mniehaus/2009/05/14/manipulating-the-microsoft-deployment-toolkit-database-using-powershell/) and is included in this repo for convenience.

This scirpt is designed to read in a csv file containing headers and the computer information.  A template csv file is included in this repo.   

This example loads Mac address, Computer name, Admin Password, and Role. It could easily be adjusted for any set of fields.  

Obviously the line of code that sets up the connection to the database will need to be modified to match the environment being used.   

The DeleteComputers.ps1 script needs some clean up as it was created by simply deleting some parts out of the original LoadComputers.ps1 script.
# powershell-refresh-excel
PowerShell script that refreshes a list of excel files.

This script will go through a list of Excel files, access them in the background, resfresh the entire file, save, and then close. My intent for this script was to automatically refresh excel files that were connected to SQL databases. This can be particularly helpful if: 

<ol>
<li>Your file is accessed by multiple users, and those users don't have access to the SQL database to refresh the file themselves.</li>
<li>The query takes a while to run, therefore refreshing it during off-hours can save time.</li>
<li>You have many such files, so refreshing them manually is a pain.</li>
</ol>



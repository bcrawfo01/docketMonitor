docketMonitor
===========================
automated case alerts using Google Apps Script


## Overview
This script is designed to help attorneys keep track of case events using [public CourtConnect](https://caseinfo.aoc.arkansas.gov/cconnect/PROD/public/ck_public_qry_main.cp_main_idx). The script monitors cases on CourtConnect and sends email alerts when entries are added or removed from a docket page.
 * It checks for any new cases once a day and sends an email alert when cases are added
 * It checks every case for an individual attorney twice a day and sends an email alert if new entries are added or old entries are modified or removed
 * Any scanned documents added to a docket page are attached to alert emails


## Limitations
The script depends on Google's servers so if [Google experiences performance issues](https://www.google.com/appsstatus), the script will experience errors. 
Apps Script has [quotas](https://developers.google.com/apps-script/guides/services/quotas). Attorneys with a substantial number of cases might hit quota limitations. You might also hit quota limitations if you use a free Google account. 



## Installation
1. Create a [new Google Sheet](http://spreadsheets.google.com/ccc?new), and open the script editor (Tools > Script editor). If you see a new script dialog, select the Blank Project option.
2. Delete the default code. Copy the code from [docketMonitor.gs](https://raw.githubusercontent.com/bcrawfo01/docketMonitor/master/docketMonitor.gs) and paste it into the script editor. Save the file. Use any name you wish when prompted to name the script.
3. Select Run > setup and grant the script authorization.
4. Close the script editor.
5. Refresh the Google Sheet. Once you refresh the sheet, the script will add a custom menu named "Docket Monitor" to the menu bar.
6. Review the help file (Docket Monitor > help) and add your information to the settings sheet.
7. Update your case list (Docket Monitor > update case list).

That’s it. You can close the sheet. Everything else is automated.



## Removing the script
The script creates a folder named "Docket Monitor" to store files without cluttering Drive. To stop the script, delete the folder. 


***
<strong>Released under Creative Commons</strong>

<a rel="license" href="http://creativecommons.org/licenses/by-sa/4.0/"><img alt="Creative Commons License" style="border-width:0" src="https://i.creativecommons.org/l/by-sa/4.0/88x31.png" /></a><br /><span xmlns:dct="http://purl.org/dc/terms/" property="dct:title">Docket Monitor</span> by <a xmlns:cc="http://creativecommons.org/ns#" href="https://www.dynamicpractices.com/" property="cc:attributionName" rel="cc:attributionURL">Brandon Crawford</a> is licensed under a <a rel="license" href="http://creativecommons.org/licenses/by-sa/4.0/">Creative Commons Attribution-ShareAlike 4.0 International License</a>.

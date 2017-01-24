docketMonitor
===========================
automated case alerts using Google Apps Script


## Overview
This script is designed to help attorneys keep track of case events using [public CourtConnect](https://caseinfo.aoc.arkansas.gov/cconnect/PROD/public/ck_public_qry_main.cp_main_idx). The script monitors cases for an individual attorney on CourtConnect and sends email alerts when entries are added or removed from a docket page.
 * It checks for any new cases once a day and sends an email alert when cases are added
 * It checks every case twice a day and sends an email alert if new entries are added or old entries are modified or removed
 * Any scanned documents added to a docket page are attached to alert emails


## Limitations
The script depends on Google's servers and CourtConnect’s servers so if either [Google](https://www.google.com/appsstatus) or CourtConnect experience performance or maintenance issues, the script will generate errors.<br> 
Apps Script has [quotas](https://developers.google.com/apps-script/guides/services/quotas). Attorneys with a substantial number of cases might hit quota limitations. You might also hit quota limitations if you use a free Google account. 


## Changelog & Updates
[v3.0](https://github.com/bcrawfo01/docketMonitor/tree/91dc74d95e1aaab8913cd18548b4c15e085027fe/docketMonitor.gs)<br>
&nbsp; &bull; Modified the text used for comparison to avoid false updates<br>
&nbsp; &bull; Reduced max run time to prevent "Exceeded maximum execution time" error<br>
&nbsp; &bull; Updated some formatting for update emails<br>
&nbsp; &bull; Added link to Docket Monitor Google Sheet to update emails<br>
&nbsp; &bull; Modified getSettings function<br>

[v2.0](https://github.com/bcrawfo01/docketMonitor/tree/b73b72b71555fd56000fb9ea0b4804915589875c/docketMonitor.gs)<br>
&nbsp; &bull; Added ability to stop unwanted email updates for certain cases<br>
&nbsp; &bull; Added ability to specify additional cases to monitor<br>

[v1.0](https://github.com/bcrawfo01/docketMonitor/blob/1a43ff79b9cf75b26a8d8cc7b8abc9c5ebc57e2e/docketMonitor.gs)<br>
&nbsp; &bull; Initial release


## Installation
<p><i>Method 1 (easier method):</i></p>
1. Open [this Google sheet](https://docs.google.com/spreadsheets/d/1_20QFJNNWEYpGvjbX8UjWZnuAYRCp-QE3Xc20rJJjVk/edit?usp=sharing), and make a copy (File > Make a copy...).
2. Refresh the new Google Sheet. Once you refresh the sheet, the script will add a custom menu named "Docket Monitor" to the menu bar.
3. Run setup (Docket Monitor > setup) and grant the script authorization.
4. Review the help file (Docket Monitor > help) and add your information to the settings sheet.
5. Update your case list (Docket Monitor > update case list).

<p><i>Method 2:</i></p>
1. Create a [new Google Sheet](http://spreadsheets.google.com/ccc?new), and open the script editor (Tools > Script editor). If you see a new script dialog, select the Blank Project option.
2. Delete the default code. Copy the code from [docketMonitor.gs](https://raw.githubusercontent.com/bcrawfo01/docketMonitor/master/docketMonitor.gs) and paste it into the script editor. Save the file. Use any name you wish when prompted to name the script.
3. Select Run > setup and grant the script authorization.
4. Close the script editor.
5. Refresh the Google Sheet. Once you refresh the sheet, the script will add a custom menu named "Docket Monitor" to the menu bar.
6. Review the help file (Docket Monitor > help) and add your information to the settings sheet.
7. Update your case list (Docket Monitor > update case list).


[YouTube video of installation process](https://youtu.be/Pf-myw_do9w)


That’s it. You can close the sheet. Everything else is automated.


Optional: [Sign up](http://github-file-watcher.com/?repository=bcrawfo01/docketMonitor&glob=*) to receive email alerts when this repo is updated.


## Updating
1. Open the Docket Monitor Google Sheet, and open the script editor (Tools > Script editor).
2. Delete the existing code.
3. Copy the code from docketMonitor.gs and paste it into the script editor.
4. Save the file.


## Removing the script
The script creates a folder named "Docket Monitor" to store files without cluttering Drive. To stop the script, delete the folder. 


***
<strong>Released under Creative Commons</strong>

<a rel="license" href="http://creativecommons.org/licenses/by-sa/4.0/"><img alt="Creative Commons License" style="border-width:0" src="https://i.creativecommons.org/l/by-sa/4.0/88x31.png" /></a><br /><span xmlns:dct="http://purl.org/dc/terms/" property="dct:title">Docket Monitor</span> by <a xmlns:cc="http://creativecommons.org/ns#" href="https://www.dynamicpractices.com/" property="cc:attributionName" rel="cc:attributionURL">Brandon Crawford</a> is licensed under a <a rel="license" href="http://creativecommons.org/licenses/by-sa/4.0/">Creative Commons Attribution-ShareAlike 4.0 International License</a>.

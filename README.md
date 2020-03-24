# Microsoft Teams self-service app
Contains Powerhshell script for Teams self-service app. Requests are stored in a SharePoint list through PowerApps and processed in Powershell. PowerAutomate formats and sends email messages, and then deletes SharePoint list items when a request is processed.

Benefits:
Eliminates duplicate display name limitation on the Teams client.
Ensures that the Team email address is unique without random numbers appended to it.
Informs the requestor when a name is already in use and provides the name of the current Team owner.
Blocks restricted words from being used. More restricted words can be added in the csv file. 
Gives administrators more control over how teams are created. The script can be modified to include unique naming conventions.

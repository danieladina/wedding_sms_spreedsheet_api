# wedding_sms_spreedsheet_api


1.1.1-Steps to Run the Code:



1.2.1-Setting Up Google Sheets API:



1.3.1-Google Cloud Platform (GCP) Setup:


1.3.2-Create or use an existing GCP project.
1.3.3-Enable the Google Sheets API for the chosen project.


1.4.1-Generate Credentials:


1.4.2-Create a service account in the GCP project.
1.4.3-Download the service account's credentials as a JSON file.


1.5.1-Prepare Python Environment:


1.5.2-Install Required Packages:


1.5.3-Use pip to install the necessary Python packages:


1.5.4-pip install google-auth google-auth-oauthlib google-auth-httplib2 



1.5.5-Adjust Code Variables:


1.6.1-Set File Paths:


1.6.2-Modify the FOLDER_PATH variable to point to the directory where the downloaded credentials JSON file is located.


1.7.1-Customize Data:


1.7.2-Update placeholders such as SHEET_NAME, values_a4, values_b4, etc., with the desired content or leave them as is based on your needs.


2.1.1-Running the Code:


2.2.1-Execute the Script:


2.2.2-Save the modifications made to the script.


2.3.1-Run the Python script via the preferred Python interpreter:
python your_script_name.py 


2.4.1Code Functionality Overview:


2.4.2-The provided code performs the following actions:

-Authentication and Service Creation:

-Imports necessary modules and sets up authentication using Google Sheets API credentials.

-Creating a New Spreadsheet:

-Utilizes the Google Sheets service to create a new spreadsheet.

-Retrieves the URL and ID of the newly created spreadsheet.

-Manipulating Spreadsheet Data:

--Formats and merges cells within the created spreadsheet using batch update requests.
--Defines border styles, merges cells in specific ranges, and applies formatting to cells.


-Updating Cell Values:
--Updates specific cells in the spreadsheet with predefined values (values_a4, values_b4, etc.).
--Utilizes the service.spreadsheets().values().update() method for different cell ranges and values.



********This code interacts with the Google Sheets API, performing tasks like creating a spreadsheet, formatting cells, merging ranges, and updating cell values. It's crucial to set up the Google Sheets API properly and modify the variables within the script to match the desired spreadsheet structure and content before execution.

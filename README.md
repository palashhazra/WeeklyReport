Weekly Asset Report
==========
It generates excel report on different assets and also prints count of individual assets and their subtypes 
in console which can be copied to send reporting mail to MS SAP team every Monday morning IST. 



## Features

- Get count of all assets and their sub types.
- Fetch all asset ID and their types and created date in excel sheet.

## Loading and configuring the module
- In paramsList.json file need to replace token string from Postman application in PROD environment.
- Also change startDate and endDate values.

## Common Usage

- Run the application by typing any one of the below commands in Terminal tab in VS code:

npm start
node app.js

#### Handling exceptions
After getting error, console will print the error message appended with the segment where error occured.
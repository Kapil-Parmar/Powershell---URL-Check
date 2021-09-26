# Description
This process is for checking if web URLs are up using Powershell

# High level process flow
* Read excel sheet having URLs to be checked.
* Make HTTP request for each URL and update status code in "Status Code" column of input sheet.
* Save a copy of updated input sheet in report sheet folder.
* Status code of 200 will imply URL check is successful. Execute macro which will make cell colour of status codes which are not 200 as red.

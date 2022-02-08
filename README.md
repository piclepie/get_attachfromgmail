# get_attachfromgmail
Get Attachment from Gmail to Spreadsheet, Trigger can be set for daily running. v1
v1 using Intergreated Gmail API to access Gmail threads and message , then get the attacchment.
use Google Advanced file service to insert a new file. Then return new file id. Using Spreadsheet service to access this file and get data from this sheet and pass the data to the existing sheet.
It will take about 60-80 seconds to run the scipt function. 



V2. 
Instead of using xlsx file as attachment export from IBM cognos studio. I used csv file format as the attachment to transfer data. 
Using get the blob of the csv , get as datastring. Use Utilities.parseCsv(csvString) service to convert string as 2d array. 
Then pass the data to existing spreadsheet.

Using this way, the script function running time is reduced to 4-8 seconds. Much fast and efficiency. 


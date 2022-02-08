function attachmentfromGmail(attachmentname,destsheet_id,destsheetname) {
  const folderid ='';
  /* https://drive.google.com/drive/folders/1jAym-_qehMg-SyKEt8hXECXD-rQZQHaS attachment from Gmail */
  const folder = DriveApp.getFolderById(folderid);
  //destsheet_id = '1XtF8ciE75YgeU2C8brD8hXd_Vi8kynEVgD6-Q6672uU';
  var threads = GmailApp.getInboxThreads(0, 50);
  var today_date = Utilities.formatDate(new Date(),"GMT-6","MM/dd/yyyy");
  var fileid_FromGmail = "";
  //get message date and today's date. using more efficiency filter call back function to get all rows in array. 
  var today_messges = threads.filter(r => {
      r.getMessages().filter(o => {
      var dates = Utilities.formatDate(o.getDate(),"GMT-6","MM/dd/yyyy");
        if (dates == today_date){
             o.getAttachments().filter(k => {
              if( k.getName() == attachmentname){
                  o.markRead();
                  //o.reply("Attachment received and updated. Thank you! ");
                  var blob = k.copyBlob();
                  var file_FromGmail = Drive.Files.insert(
                    {title: attachmentname, parents: [{"id": folder.getId()}]},
                    blob,
                    {convert: true}
                  );   
                 fileid_FromGmail = file_FromGmail.id ;
                 o.moveToTrash(); 
              }
           });
        }
     });
    });
    var created_file = SpreadsheetApp.openById(fileid_FromGmail);
    var values = created_file.getDataRange().getDisplayValues();
    var destfiles_sheet = SpreadsheetApp.openById(destsheet_id).getSheetByName(destsheetname);
    destfiles_sheet.clear();
    destfiles_sheet.getRange(1,1,created_file.getLastRow(), created_file.getLastColumn())
    .setValues(values);
    DriveApp.getFileById(fileid_FromGmail).setTrashed(true);  
 }




//using csv format file to transfer data from IBM cognos report studio.
//compare to v1, v2 improves the code running time. And does not need to create files and delete file on the google drive. 
//average code running time. 4-8 secons.
//v2 of data sync function
function attachmentfromGmail_csv(attachmentname, destsheet_id, destsheetname) {
   //using custom label will significantly improve the code running time. GmailApp service do not need to go through all the threads
   //and the messages.  It need you to filfer the message on the gmail app. 
   const labels = GmailApp.getUserLabelByName('Yourlabelname').getThreads(0,10);
   const threads = GmailApp.getMessagesForTHreads(labels);
   const today_date = Utilities.formatDate(new Date(), "GMT-6", "MM/dd/yyyy");
   let values;
   threads.forEach(threads => {
      threads.filter(message => {
         if (Utilities.formatDate(message.getDate(), "GMT-6", "MM/dd/yyyy") == today_date) {
            message.getAttachments().filter(attchments => {
               if (attchments.getName() == attachmentname && message.isUnread() === true) {
                  const blob = attchments.copyBlob().getDataAsString();
                  values = Utilities.parseCsv(blob);
                  message.markRead();
               }
            })
         }
      })
   });
   const destfiles_sheet = SpreadsheetApp.openById(destsheet_id).getSheetByName(destsheetname);
   destfiles_sheet.clear();
   destfiles_sheet.getRange(1, 1, values.length, values[0].length).setValues(values);
}












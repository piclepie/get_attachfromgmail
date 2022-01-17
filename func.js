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


// this function is set for removing duplicate data on tobe hire report. 
  function filter_duplicate(){
   
    var main_ssSheet = SpreadsheetApp.openById(id);
    var rawDataTBH = main_ssSheet.getSheetByName(sheetname)
                             .getDataRange()
                             .getValues();    
    var sheeet = main_ssSheet.getSheetByName(sheetname);

    var unique = rawDataTBH.slice().filter((v,i,a) => a.findIndex(t=>(t[0] === v[0]))===i);

      sheeet.clear();
      sheeet.getRange(1,1,unique.length,unique[1].length).setValues(unique);

     
  }


  function filter_admin_ppty(){
   
    const main_ssSheet = SpreadsheetApp.openById(id);
    const all_properties = main_ssSheet.getSheetByName(sheetname)
                             .getDataRange()
                             .getValues();    
    const sheeet = main_ssSheet.getSheetByName(sheetname);

    const active_values = all_properties.filter(r => { return (r[2] == '' || r[2] == '')&& r[7]!=='').filter(r => {return r[10].includes('')});
    
      sheeet.getRange(2,1,sheeet.getLastRow(),sheeet.getLastColumn()).clear()
      //sheeet.clear();
      sheeet.getRange(2,1,active_values.length,active_values[0].length).setValues(active_values);
    

     
  }













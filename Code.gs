function iPull() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet Details");
  var analyst = ss.getRange(7,3).getValue(); // finds the email address of the analyst on the Sheet Details
  
  if (analyst ==""){
    SpreadsheetApp.getUi().alert('Please enter the email address of the analyst');
      } else if(analyst !== "") {
  
      ss.getRange(6,3).clearContent(); // clears previous details
      ss.getRange(8,3).clearContent(); // clears previous details
      ss.getRange(9,3).clearContent(); // clears previous details
      ss.getRange(10,3).clearContent(); // clears previous details
      ss.getRange(11,3).clearContent(); // clears previous details
      ss.getRange(12,3).clearContent(); // clears previous details
  
  
  var user = ss.getRange(5,3).getValue(); // finds the name of the customer on the Sheet Details
  var ssI = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Input Sheet"); // defines the data sheet 
  ssI.getRange("b2:bs260").clearContent();
  var source = SpreadsheetApp.openById("1PDSr53kxFwWGDk9CdEu3KGu8mysSMxtExj7VRXV13nY"); // finds the iDatabase sheet
  var newdata = source.getSheetByName(user);  // finds the users data in iDatabase
  var trans = newdata.getRange("A1:bs260").getValues(); // gets the users data
  ssI.getRange("A1:bs260").setValues(trans); // pastes it to this sheet
        
  var work1 = ssI.getRange(1,2,ssI.getLastRow()); // creates a range of the accounts pulled from iMaster, including headers
  var work2 = work1.getLastRow();  // counts the length of the range including headers
  var report = source.getSheetByName("Batches");
  var loc = report.getRange(1,1).getValue();
        var load = ssI.getRange("A1:A").getValues();  // measure size of the user's account
   var work3 = load.filter(String).length;  // used this logic https://stackoverflow.com/questions/17632165/determining-the-last-row-in-a-single-column
  var tDate = new Date(); // date stamps the process
  
        report.getRange(loc+2,3).setValue(analyst); // writes the batch info to batches in iDatabase
        report.getRange(loc+2,4).setValue(user);
        report.getRange(loc+2,5).setValue(work3-1);
        report.getRange(loc+2,6).setValue(tDate);
  
        
  var uData = source.getSheetByName("UserList");  // locates UserList in iDatabase
  var detail = uData.getRange("b2:w100").getValues();
  
  
  for (var x = 0; x< 97; x++){
    if(detail[x][0] === user){
      ss.getRange(6,3).setValue(detail[x][3]); // writes the company name
      ss.getRange(8,3).setValue(detail[x][7]); // writes the user domain
      ss.getRange(9,3).setValue(detail[x][8]); // writes the company size
      ss.getRange(10,3).setValue(detail[x][9]); // writes the product type
      ss.getRange(11,3).setValue(detail[x][4]); // writes the user email
      ss.getRange(12,3).setValue(detail[x][10]); // writes the user note
    }
  }
  
  var iMaster = SpreadsheetApp.openById("1sEjzhq96me6aaQLIBqY6Wfgy9D6VrKhtHL9eUoqyT2Q"); // finding the master spreadsheet currently set to Robustified Index Master
  var ePfull = iMaster.getSheetByName("Event ID").getRange("a3:bj1999").getValues(); // take all Events and makes them an array
  var lu = ssI.getRange("B3:b253").getValues(); // makes a range of the Contacts
   
 
  for (var j=0; j < 1996; j++){
    
  if (ePfull[j][1] == user && ePfull[j][28] == "PASS" ){    
      
    var Con = ePfull[j][5];
    var Note = ePfull[j][13]; // takes the event note
    var URL = ePfull[j][15]; // take the event URL
    var date = ePfull[j][16]; // take the event date
    var label = ePfull[j][14]; // take the event label
    var copy = ePfull[j][50]; // copies the text into the copy variable
      
      for (var k=0; k <250; k++){
        
        if (lu[k][0] == Con){
          
          ssI.getRange(k+3,54).setValue(date); // returns the event date
          ssI.getRange(k+3,55).setValue(label);  // returns the event label
          ssI.getRange(k+3,56).setValue(Note);  // returns the event note
          ssI.getRange(k+3,57).setValue(URL);  // returns the event URL
          ssI.getRange(k+3,58).setValue(copy);  // returns the event copy
        
        }
      }
  } else if(ePfull[j][1] == user){
    
    var Con = ePfull[j][5];
    var Note = ePfull[j][13]; // takes the event note
    var URL = ePfull[j][15]; // take the event URL
    var date = ePfull[j][16]; // take the event date
    var label = ePfull[j][14]; // take the event label
    
    for (var m=0; m <250; m++){
        
        if (lu[m][0] == Con){
          
          ssI.getRange(m+3,59).setValue(date); // returns the event 
          ssI.getRange(m+3,60).setValue(label);  // returns the event
          ssI.getRange(m+3,61).setValue(Note);  // returns the event
          ssI.getRange(m+3,62).setValue(URL); 
     
  }
    }
    }
  }
}
  
}

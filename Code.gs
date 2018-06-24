// SHEET IDS

// Robustified Index Master: 1-5Vf4LbGOI29eVabBluk8WoHg5-8qJkzhgazLdLtVDE
// Test Rig iQA: 1-5Vf4LbGOI29eVabBluk8WoHg5-8qJkzhgazLdLtVDE



function iPull() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet Details");
  var eventSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Event List");
  var analyst = ss.getRange(7,3).getValue(); // finds the email address of the analyst on the Sheet Details
  var ssI = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Input Sheet"); // defines the data sheet
  
  
   for(var i = 3; i < 253; i++){
    
    var event1 = ssI.getRange(i,41).getValue(); // column 28 is where the label is 
    var event2 = ssI.getRange(i,47).getValue(); // column 33 is where the second label is
    

    if(event2 !== "") { 
      
    var actRow = ssI.getActiveCell().getRow(); //this is the row co-ordinate of the change to the Event 1 Category **not sure this is used anymore? and line below.
    eventSheet.getRange(1,4).setValue(ssI.getActiveCell().getRow());//prints the row of the on Edited cell Q why do you need to define the var within the if clause not above it?
    
    
    var range = eventSheet.getRange(1,1).getValue(); // Defines the range of the events down to the last event currently on the list
   
    var newPrint = range+2; // Row co-ordinate of the new event
    var nP2 = newPrint+1; // Row below for the second layer
    var Contact = ssI.getRange(i,2).getValue(); // Returns the name of the contact
    var userOld = ssI.getRange(1,1).getValue(); // User's name
    var userComp = ss.getRange(6,3).getValue(); // User's company name
    // var userLink = ssI.getRange(7,47).getValue(); // User's sheet link
    var userMail = ss.getRange(11,3).getValue(); // User's email
    var contactmail = ssI.getRange(i,11).getValue(); // Returns the name of the contact's work email
    var contactpmail = ssI.getRange(i,12).getValue(); // Returns the name of the contact's private email
    var eventNote1 = ssI.getRange(i,40).getValue(); // Returns event note 1
    var eventLabel1 = ssI.getRange(i,41).getValue(); // event label 1
    var eventURL1 = ssI.getRange(i,42).getValue(); // event URL 1
    var eventBAU1 = ssI.getRange(i,43).getValue(); // event BAU 1
    var eventIMP1 = ssI.getRange(i,44).getValue(); // event IMP 1
    var eventDate1 = ssI.getRange(i,45).getValue(); // event date 1
    var userNote = ss.getRange(12,3).getValue(); // Returns the user note
    var contactComp = ssI.getRange(i,1).getValue(); // contact's company
    var contactLI = ssI.getRange(i,30).getValue(); // contact's LinkedIn
    var contactRole = ssI.getRange(i,13).getValue(); // contacts's role
    var eventNote2 = ssI.getRange(i,46).getValue(); // Returns event note 2
    var eventLabel2 = ssI.getRange(i,47).getValue(); // event label 2
    var eventURL2 = ssI.getRange(i,48).getValue(); // event URL 2
    var eventBAU2 = ssI.getRange(i,49).getValue(); // event BAU 2
    var eventIMP2 = ssI.getRange(i,50).getValue(); // event IMP 2
    var eventDate2 = ssI.getRange(i,51).getValue(); // event date 2
    var eventhstat = ssI.getRange(i,16).getValue(); // event home static
      
      Logger.log(user);
    
    eventSheet.getRange(newPrint,5).setValue(i);
    eventSheet.getRange(newPrint,6).setValue(Contact);
    eventSheet.getRange(newPrint,7).setValue(contactComp); // Contact's company
    eventSheet.getRange(newPrint,8).setValue(contactLI); // Individual LinkedIn profile
    eventSheet.getRange(newPrint,9).setValue(contactmail);
    eventSheet.getRange(newPrint,10).setValue(contactpmail);
    eventSheet.getRange(newPrint,11).setValue(userMail);
    eventSheet.getRange(newPrint,2).setValue(userOld); // User name
    eventSheet.getRange(newPrint,3).setValue(userComp); // User's company
    // eventSheet.getRange(newPrint,4).setValue(userLink); // User's sheet link
    eventSheet.getRange(newPrint,12).setValue(userNote); // User note
    eventSheet.getRange(newPrint,13).setValue(contactRole); // Contact's role
    eventSheet.getRange(newPrint,14).setValue(eventNote1); // event 1 note 
    eventSheet.getRange(newPrint,15).setValue(eventLabel1); // event 1 label
    eventSheet.getRange(newPrint,16).setValue(eventURL1); // event 1 URL
    eventSheet.getRange(newPrint,17).setValue(eventDate1); // event 1 date
    eventSheet.getRange(newPrint,20).setValue(eventBAU1); // event 1 BAU
    eventSheet.getRange(newPrint,21).setValue(eventIMP1); // event 1 IMP
    eventSheet.getRange(newPrint,22).setValue(eventhstat); // event 1 home static
      
    eventSheet.getRange(nP2,5).setValue(i);
    eventSheet.getRange(nP2,6).setValue(Contact);
    eventSheet.getRange(nP2,7).setValue(contactComp); // Contact's company
    eventSheet.getRange(nP2,8).setValue(contactLI); // Individual LinkedIn profile
    eventSheet.getRange(nP2,9).setValue(contactmail);
    eventSheet.getRange(nP2,10).setValue(contactpmail);
    eventSheet.getRange(nP2,11).setValue(userMail);
    eventSheet.getRange(nP2,2).setValue(userOld); // User name
    eventSheet.getRange(nP2,3).setValue(userComp); // User's company
    // eventSheet.getRange(nP2,4).setValue(userLink); // User's sheet link
    eventSheet.getRange(nP2,12).setValue(userNote); // User note
    eventSheet.getRange(nP2,13).setValue(contactRole); // Contact's role
    eventSheet.getRange(nP2,14).setValue(eventNote2); // event 2 note 
    eventSheet.getRange(nP2,15).setValue(eventLabel2); // event 2 label
    eventSheet.getRange(nP2,16).setValue(eventURL2); // event 2 URL
    eventSheet.getRange(nP2,17).setValue(eventDate2); // event 2 date
    eventSheet.getRange(nP2,20).setValue(eventBAU2); // event 2 BAU
    eventSheet.getRange(nP2,21).setValue(eventIMP2); // event 2 IMP
    eventSheet.getRange(nP2,21).setValue(eventhstat); // event 2 home static
      
    } else if(event1 !== "") {
      
    var range = eventSheet.getRange(1,1).getValue(); // Defines the range of the events down to the last event currently on the list
   
    
    var newPrint = range+2; // Row co-ordinate of the new event
    var nP2 = newPrint+1; // Row below for the second layer
    var Contact = ssI.getRange(i,2).getValue(); // Returns the name of the contact
    var userOld = ssI.getRange(1,1).getValue(); // User's name
    var userComp = ss.getRange(6,3).getValue(); // User's company name
    // var userLink = ssI.getRange(7,47).getValue(); // User's sheet link
    var userMail = ss.getRange(11,3).getValue(); // User's email
    var contactmail = ssI.getRange(i,11).getValue(); // Returns the name of the contact's work email
    var contactpmail = ssI.getRange(i,12).getValue(); // Returns the name of the contact's private email
    var eventNote1 = ssI.getRange(i,40).getValue(); // Returns event note 1
    var eventLabel1 = ssI.getRange(i,41).getValue(); // event label 1
    var eventURL1 = ssI.getRange(i,42).getValue(); // event URL 1
    var eventBAU1 = ssI.getRange(i,43).getValue(); // event BAU 1
    var eventIMP1 = ssI.getRange(i,44).getValue(); // event IMP 1
    var eventDate1 = ssI.getRange(i,45).getValue(); // event date 1
    var userNote = ss.getRange(12,3).getValue(); // Returns the user note
    var contactComp = ssI.getRange(i,1).getValue(); // contact's company
    var contactLI = ssI.getRange(i,30).getValue(); // contact's LinkedIn
    var contactRole = ssI.getRange(i,13).getValue(); // contacts's role
    var eventhstat = ssI.getRange(i,16).getValue(); // event home static  
    
    eventSheet.getRange(newPrint,5).setValue(i);
    eventSheet.getRange(newPrint,6).setValue(Contact);
    eventSheet.getRange(newPrint,7).setValue(contactComp); // Contact's company
    eventSheet.getRange(newPrint,8).setValue(contactLI); // Individual LinkedIn profile
    eventSheet.getRange(newPrint,9).setValue(contactmail);
    eventSheet.getRange(newPrint,10).setValue(contactpmail);
    eventSheet.getRange(newPrint,11).setValue(userMail);
    eventSheet.getRange(newPrint,2).setValue(userOld); // User name
    eventSheet.getRange(newPrint,3).setValue(userComp); // User's company
    // eventSheet.getRange(newPrint,4).setValue(userLink); // User's sheet link
    eventSheet.getRange(newPrint,12).setValue(userNote); // User note
    eventSheet.getRange(newPrint,13).setValue(contactRole); // Contact's role
    eventSheet.getRange(newPrint,14).setValue(eventNote1); // event 1 note 
    eventSheet.getRange(newPrint,15).setValue(eventLabel1); // event 1 label
    eventSheet.getRange(newPrint,16).setValue(eventURL1); // event 1 URL
    eventSheet.getRange(newPrint,17).setValue(eventDate1); // event 1 date
    eventSheet.getRange(newPrint,20).setValue(eventBAU1); // event 1 BAU
    eventSheet.getRange(newPrint,21).setValue(eventIMP1); // event 1 IMP  
    eventSheet.getRange(newPrint,22).setValue(eventhstat); // event 1 home static  
      
    }
  }
  
  
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
  
  ssI.getRange("b2:bs260").clearContent();
  var source = SpreadsheetApp.openById("1PDSr53kxFwWGDk9CdEu3KGu8mysSMxtExj7VRXV13nY"); // finds the iDatabase sheet
  var newdata = source.getSheetByName(user);  // finds the users data in iDatabase
  var trans = newdata.getRange("A1:bs260").getValues(); // gets the users data
  ssI.getRange("A1:bs260").setValues(trans); // pastes it to this sheet
  ssI.getRange(1,1).setValue(user); // adds the user name to the top left for future collection to Event List (see line 33 etc..)
        
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
  
  var iMaster = SpreadsheetApp.openById("1-5Vf4LbGOI29eVabBluk8WoHg5-8qJkzhgazLdLtVDE"); // finding the master spreadsheet currently set to Robustified Index Master
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

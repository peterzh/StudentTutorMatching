function flatten(arrayOfArrays){
  return [].concat.apply([], arrayOfArrays);
}

function getProgrammes() {
  //function returns the names of the MA programmes and tutors
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  var rows = doc.getSheetByName("Names").getDataRange().getValues();
  var MA_Programmes = [];
  
  for (i=1;i<rows.length;i++) {
    if (rows[i][0]!=="") {
      MA_Programmes.push(rows[i][0]);
    } 
  }
  return MA_Programmes;
}

function getTutors() {
  //function returns the names of the MA programmes and tutors
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  var rows = doc.getSheetByName("Names").getDataRange().getValues();
  var Tutors = [];
  
  for (i=1;i<rows.length;i++) {
    if (rows[i][1]!=="") {
      Tutors.push(rows[i][1]);
    }
  }
  return Tutors;
}

function createStudentForm()
{
  
  //check if another form is already there and warning that they need to unlink/delete first
  
  
  var confirm = Browser.msgBox('Confirmation','Are you sure you want to create a new form? The old URL will be overwritten!', Browser.Buttons.OK_CANCEL);
  if(confirm=='ok'){ 
    
    //get list of tutors and MA programmes
    tutors = getTutors();
    programmes = getProgrammes();
    
    //Create form for students to rank tutors
    var form = FormApp.create('Student Preferences')
    .setCollectEmail(true).setAllowResponseEdits(true).setRequireLogin(false).setShowLinkToRespondAgain(false);
    
    //add field for entering their name
    var item = form.addTextItem();
    item.setTitle('Full name').setRequired(true);
    
    //add field for selecting their MA programme
    var name = form.addListItem();
    name.setTitle('MA Programme')
    .setChoiceValues(programmes).setRequired(true);
    
    //add matrix for selecting preferences
    var item = form.addSectionHeaderItem();
    item.setTitle('Please rank your preferences');
    
    var num=[];
    for (i = 1; i <= tutors.length; i++) {
      num[i-1] = i;
    }
    var item = form.addGridItem();
    item.setRows(tutors)
    .setColumns(num)
    .setRequired(true);
    // .setHelpText('You may need to scroll horizontally to see all tutors');
    
    var gridValidation = FormApp.createGridValidation()
    .setHelpText('Select one tutor per rank.')
    .requireLimitOneResponsePerColumn()
    .build();
    item.setValidation(gridValidation);
    
    //place links in the main page
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("START HERE");  
    sheet.getRange('D6').setValue(form.getPublishedUrl()).setFontColor('green');
    
    //link form to new spreadsheet and place link to spreadsheet
    var responses_sheet = SpreadsheetApp.getActiveSpreadsheet();
    form.setDestination(FormApp.DestinationType.SPREADSHEET, responses_sheet.getId());
    sheet.getRange('D7').setValue(responses_sheet.getUrl()).setFontColor('green');
    
  };
}

function getResponses() {
  //get list of tutors and MA programmes
  tutors = getTutors();
  programmes = getProgrammes();
  
  //Load student preferences from the form responses
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = doc.getSheetByName("START HERE");  
  var studentResponseURL = sheet.getRange('D7').getValue();
  studentResponses = SpreadsheetApp.openByUrl(studentResponseURL).getDataRange().getValues();
  studentResponses.splice(0,1); //remove headers
  
  return studentResponses;
}

function runAlgorithm() {
  tutors = getTutors();
  programmes = getProgrammes();
  responses = getResponses();
  
  Logger.log(responses);
  
  
  var results = [];
  var unassigned_students = [];
  //split the responses by programme
  for (p=0;p<programmes.length;p++) {
    
    //get responses from one MA programme only
    var respProg = responses.filter(
      function(row) { return row[3]==programmes[p]})
    
    if (respProg.length>0) {
      
      var names = respProg.map(function(row){return row[2]});
      var prefs = respProg.map(function(row){return row.slice(4)});
      
      //calculate the number of slots given to this MA programme for each tutor
      var numSlots = Math.ceil(respProg.length/tutors.length); //number of slots given to each tutor is ceiling(numStudents/numTutors)
      
      //pad the prefs by repeating each rank by numSlots so the assignment problem can be solved for multiple 'slots' per tutor
      var prefs_padded = prefs; //copy of array            
      for (i=0;i<prefs.length;i++) {        
        for (j=0;j<tutors.length;j++) {
          prefs_padded[i][j] = Array.apply(null, Array(numSlots)).map(function(){return prefs[i][j] });
        }
        prefs_padded[i] = flatten(prefs_padded[i]);
      }
      //pad the tutor names as well for extracting assignments later
      var tutors_padded = [];
      for (j=0;j<tutors.length;j++) {
        tutors_padded[j] = Array.apply(null, Array(numSlots)).map(function(){return tutors[j] });
      }
      tutors_padded = flatten(tutors_padded);
      
      
      //run algorithm on padded preferences
      Logger.log(programmes[p]);
      Logger.log('    Number of responses: ' + respProg.length.toString());
      Logger.log('    Number of slots given per tutor: ' + numSlots.toString());
      Logger.log('    Number of total slots given to this programme: ' + (numSlots*tutors.length).toString());
      
      var h = new Hungarian(prefs_padded);
      var assign_idx = h.execute();
      var assign_tutor = assign_idx.map(function(idx){return tutors_padded[idx]});
      
      //write out to dataset
      for (i=0;i<names.length;i++) {
        if (assign_idx[i]==-1) {
          unassigned_students.push([names[i]]);
        } else {
          var rankOfPreference = prefs_padded[i][tutors_padded.indexOf(assign_tutor[i])];
          results.push([names[i], programmes[p], assign_tutor[i], rankOfPreference]); 
        }
      }
    }
  }
  
  
  //write out assignments to sheet  
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  
  doc.getSheetByName("Assignments (List)").getRange("A2:E").clear();
  range = doc.getSheetByName("Assignments (List)").getRange(2, 1, results.length, 4)
  range.setValues(results);
  
  Logger.log(unassigned_students);
  if (unassigned_students.length>0) {
    Logger.log(unassigned_students);
    
    range = doc.getSheetByName("Assignments (List)").getRange(2, 6, unassigned_students.length)
    range.setValues(unassigned_students);    
  }
  
}

function checkStudentPrefsForDuplicates() {
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  
  //Load student preferences from the form responses
  var sheet = doc.getSheetByName("START HERE");  
  var studentResponseURL = sheet.getRange('J4').getValue();
  studentPrefsSheet = SpreadsheetApp.openByUrl(studentResponseURL).getSheetByName('Form Responses 1');
  studentPreferences = studentPrefsSheet.getDataRange().getValues();
  var numResponses = studentPreferences.length - 1; 
  
  var anyDuplicates = false;
  for (i=0;i<numResponses-1;i++) { //go through each entry
    for (j=i+1;j<numResponses;j++) { //for all other subsequent entries
      if (studentPreferences[1+i][2]==studentPreferences[1+j][2]) {
        
        studentPrefsSheet.getRange(1+i+1,3,1,1).setBackground("red");
        studentPrefsSheet.getRange(1+j+1,3,1,1).setBackground("red");
        
        anyDuplicates = true;
      }
    }
  }
  
  return anyDuplicates;
}

function checkTutorPrefsForDuplicates() {
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  
  //Load tutor preferences from the form responses
  var sheet = doc.getSheetByName("START HERE");  
  var tutorResponseURL = sheet.getRange('J7').getValue();
  tutorPrefsSheet = SpreadsheetApp.openByUrl(tutorResponseURL).getSheetByName('Form Responses 1');
  tutorPreferences = tutorPrefsSheet.getDataRange().getValues();
  var numResponses = tutorPreferences.length - 1;
  
  var anyDuplicates = false;
  for (i=0;i<numResponses-1;i++) { //go through each entry
    for (j=i+1;j<numResponses;j++) { //for all other subsequent entries
      if (tutorPreferences[1+i][2]==tutorPreferences[1+j][2]) {
        
        tutorPrefsSheet.getRange(1+i+1,3,1,1).setBackground("red");
        tutorPrefsSheet.getRange(1+j+1,3,1,1).setBackground("red");
        
        anyDuplicates = true;
      }
    }
  }
  
  return anyDuplicates;
}

function createStudentForm()
{
  var confirm = Browser.msgBox('Confirmation','Are you sure you want to create new forms? The old URLs will be overwritten!', Browser.Buttons.OK_CANCEL);
  if(confirm=='ok'){ 
    
    //get list of tutors and students
    var doc = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = doc.getSheetByName("Names (Students)");  
    var students = sheet.getDataRange().getValues();
    numStudents = students.length;
    
    var sheet = doc.getSheetByName("Names (Tutors)");  
    var tutors = sheet.getDataRange().getValues();
    numTutors = tutors.length;
    
    //Create form for students to rank tutors
    var form = FormApp.create('Student Preferences')
    .setCollectEmail(true).setAllowResponseEdits(true).setRequireLogin(false);
    
    var name = form.addListItem();
    name.setTitle('Name')
    .setChoiceValues(students).setRequired(true);
    
    var name = form.addListItem();
    name.setTitle('Year')
    .setChoiceValues(['One','Two']).setRequired(true);
    
    var item = form.addSectionHeaderItem();
    item.setTitle('Please enter your preferences');
    
    var num=[];
    for (i = 1; i <= numTutors; i++) {
      num[i-1] = 'Rank ' + i.toString();
    }
    var item = form.addGridItem();
    item.setRows(num)
    .setColumns(tutors)
    .setRequired(true)
    .setHelpText('You may need to scroll horizontally to see all tutors');
    
    var gridValidation = FormApp.createGridValidation()
    .setHelpText('Select one rank per tutor.')
    .requireLimitOneResponsePerColumn()
    .build();
    item.setValidation(gridValidation);
    
    /*
    //loop through and create dropdown for each rank
    for (i = 1; i <= numTutors; i++) {
    var item = form.addListItem().setTitle('Rank '+ i.toString())
    .setChoiceValues(tutors)
    .setRequired(true);
    }
    */
    
    //place links in the main page
    var sheet = doc.getSheetByName("START HERE");  
    sheet.getRange('J2').setValue(form.getPublishedUrl()).setFontColor('green');
    sheet.getRange('J3').setValue(form.getEditUrl()).setFontColor('green');
    
    //link form to new spreadsheet and place link to spreadsheet
    var responses_sheet = SpreadsheetApp.create("StudentPreferences_Responses");
    form.setDestination(FormApp.DestinationType.SPREADSHEET, responses_sheet.getId());
    sheet.getRange('J4').setValue(responses_sheet.getUrl()).setFontColor('green');
  };
}

function createTutorForm()
{
  var confirm = Browser.msgBox('Confirmation','Are you sure you want to create new forms? The old URLs will be overwritten!', Browser.Buttons.OK_CANCEL);
  if(confirm=='ok'){ 
    
    //get list of tutors and students
    var doc = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = doc.getSheetByName("Names (Students)");  
    var students = sheet.getDataRange().getValues();
    numStudents = students.length;
    
    var sheet = doc.getSheetByName("Names (Tutors)");  
    var tutors = sheet.getDataRange().getValues();
    numTutors = tutors.length;
    
    //Create form for tutors to choose top students
    var form = FormApp.create('Tutor Preferences')
    .setCollectEmail(true).setAllowResponseEdits(true).setRequireLogin(false);;
    
    var name = form.addListItem();
    name.setTitle('Name')
    .setChoiceValues(tutors).setRequired(true);
    
    /*
    //set capacity
    var textItem = form.addTextItem().setTitle('Student capacity').setRequired(true);
    var textValidation = FormApp.createTextValidation()
    .setHelpText('Input was not a number between 1 and 300.')
    .requireNumberBetween(1, 300)
    .build();
    textItem.setValidation(textValidation);
    */
    
    form.addPageBreakItem();
    
    var item = form.addSectionHeaderItem();
    item.setTitle('Please enter your preferences');
    //loop through and create dropdown for 50 preferences
    for (i = 1; i <= 100; i++) {
      var item = form.addListItem().setTitle('Rank '+i.toString()).setChoiceValues(students);
      
      if ((i % 51) == 0) {
        form.addPageBreakItem();
      }
    }
    
    //place links in the main page
    var sheet = doc.getSheetByName("START HERE");  
    sheet.getRange('J5').setValue(form.getPublishedUrl()).setFontColor('green');
    sheet.getRange('J6').setValue(form.getEditUrl()).setFontColor('green');
    
    //link form to new spreadsheet and place link to spreadsheet
    var responses_sheet = SpreadsheetApp.create("TutorPreferences_Responses");
    form.setDestination(FormApp.DestinationType.SPREADSHEET, responses_sheet.getId());
    sheet.getRange('J7').setValue(responses_sheet.getUrl()).setFontColor('green');
  };
}


function toggleStudentForm()
{
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = doc.getSheetByName("START HERE");  
  var formURL = sheet.getRange('J3').getValue();
  var form = FormApp.openByUrl(formURL);
  
  if (form.isAcceptingResponses()){
    form.setAcceptingResponses(false);
    sheet.getRange('J2:J4').setFontColor('red');   
  } else {
    form.setAcceptingResponses(true);
    sheet.getRange('J2:J4').setFontColor('green');
  }
}


function toggleTutorForm()
{
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = doc.getSheetByName("START HERE");  
  var formURL = sheet.getRange('J6').getValue();
  var form = FormApp.openByUrl(formURL);
  
  if (form.isAcceptingResponses()){
    form.setAcceptingResponses(false);
    sheet.getRange('J5:J8').setFontColor('red');   
  } else {
    form.setAcceptingResponses(true);
    sheet.getRange('J5:J8').setFontColor('green');
  }
}

function getPreferences() {
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  
  var query_data = {};
  query_data['student_prefs'] = {};
  query_data['college_prefs'] = {};
  query_data['college_capacity'] = {};
  
  //get Tutor Names
  var sheet = doc.getSheetByName("Names (Tutors)");  
  var tutors = sheet.getDataRange().getValues();
  numTutors = tutors.length;
  
  //Load student preferences from the form responses
  var sheet = doc.getSheetByName("START HERE");  
  var studentResponseURL = sheet.getRange('J4').getValue();
  studentPreferences = SpreadsheetApp.openByUrl(studentResponseURL).getSheetByName('Form Responses 1').getDataRange().getValues();
  var numResponses = studentPreferences.length - 1;
  
  //Put student preferences into query_data
  for(var s=0; s<numResponses;s++) { //for each student
    var studentName = studentPreferences[1+s][2];
    
    query_data['student_prefs'][studentName] = []
    for(var t=0; t<numTutors; t++) { //from 1st to last place 
      query_data['student_prefs'][studentName][t] = studentPreferences[1+s][4+t];
    }
  }
  
  //Load tutor preferences from the form responses
  var sheet = doc.getSheetByName("START HERE");  
  var tutorResponseURL = sheet.getRange('J7').getValue();
  tutorPreferences = SpreadsheetApp.openByUrl(tutorResponseURL).getSheetByName('Form Responses 1').getDataRange().getValues();
  var numResponses = tutorPreferences.length - 1;
  
  var students = doc.getSheetByName("Names (Students)").getDataRange().getValues();  
  numStudents = students.length;
  var tutors = doc.getSheetByName("Names (Tutors)").getDataRange().getValues();  
  numTutors = tutors.length;
  
  //Put tutor preferences into query_data
  for(var t=0; t<numResponses;t++) { //for each tutor
    var tutorName = tutorPreferences[1+t][2];
    
    query_data['college_prefs'][tutorName] = []
    //query_data['college_capacity'][tutorName] = tutorPreferences[1+t][3];
    query_data['college_capacity'][tutorName] = Math.ceil(numStudents/numTutors)+2;
    
    for(var s=0; s<100; s++) { 
      if (tutorPreferences[1+t][4+s] == "") {
        break;
      } //break if empty column
      
      query_data['college_prefs'][tutorName][s] = tutorPreferences[1+t][3+s];
    }
  }
  Logger.log(query_data);
  return query_data;
}

function computeAssignments() {
  //check that there are no duplicates
  if (checkStudentPrefsForDuplicates() | checkTutorPrefsForDuplicates()) {
    Browser.msgBox('Duplicate names in preferences','There are duplicate entries in the student and/or tutor preferences sheets (highlighted red). Please fix these and try again.', Browser.Buttons.OK);
    
  } else {
    
    var url = 'https://api.matchingtools.org/hri/demo'
    var username = 'mannheim'
    var password = 'Exc3llence!'
    
    // Make a POST request with form data.
    /*
    var query_data = {
    "student_prefs": [
    {
    "Student1": [
    "TutorA",
    "TutorB",
    "TutorC"
    ],
    "Student2": [
    "TutorA",
    "TutorB",
    "TutorC"
    ],
    "Student3": [
    "TutorA",
    "TutorB",
    "TutorC"
    ],
    "Student4": [
    "TutorA",
    "TutorB",
    "TutorC"
    ]
    }
    ],
    "college_prefs": [
    {
    "TutorA": [
    "Student1",
    "Student2",
    "Student3"
    ],
    "TutorB": [
    "Student1",
    "Student2",
    "Student3"
    ],
    "TutorC": [
    "Student1",
    "Student2",
    "Student4"
    ]
    }
    ],
    "college_capacity": [
    {
    "TutorA": 1,
    "TutorB": 2,
    "TutorC": 1
    }
    ]
    };
    */
    var query_data = getPreferences();
    
    var options = {
      'method' : 'post',
      'contentType': 'application/json',
      'headers': {"Authorization": "Basic " + Utilities.base64Encode(username + ":" + password)},
      'payload' : JSON.stringify(query_data)
    };
    var response = UrlFetchApp.fetch(url, options);
    var json = response.getContentText(); // get the response content as text
    var dataSet = JSON.parse(json); //parse text into json
    dataSet = dataSet.hri_matching;
    
    //write data to Output sheet
    var doc = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = doc.getSheetByName("Assignments (List)");  
    var range = sheet.getRange("A2:B300");
    range.clearContent();
    
    
    var rows = [],
        data;
    
    for (i = 0; i < dataSet.length; i++) {
      data = dataSet[i];    
      rows.push([data['student.y'], data['college.y']]);
    }
    
    dataRange = sheet.getRange(2, 1, rows.length, 2);
    dataRange.setValues(rows);
    dataRange.sort([{column: 1, ascending: true}]);
    
    doc.toast('Assignments computed');
  }
}



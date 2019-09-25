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
    .setCollectEmail(true).setAllowResponseEdits(true);
    
    var name = form.addListItem();
    name.setTitle('Name')
    .setChoiceValues(students).setRequired(true);
    
    //create matrix for choosing tutor rankings
    var nums = [];
    for (i = 1; i <= numTutors; i++) {
      nums.push('Rank '+ i.toString());
    }
    
    var checkboxGridItem = form.addCheckboxGridItem();
    checkboxGridItem.setTitle('Preference')
    .setRows(nums)
    .setColumns(tutors)
    .setRequired(true)
    .setHelpText('Note: you may have to scroll horizontally to see more tutors');
    
    var checkboxGridValidation = FormApp.createCheckboxGridValidation()
    .setHelpText('Select one item per column.')
    .requireLimitOneResponsePerColumn()
    .build();
    checkboxGridItem.setValidation(checkboxGridValidation);
    
    //place links in the main page
    var sheet = doc.getSheetByName("START HERE");  
    sheet.getRange('J2').setValue(form.getPublishedUrl()).setFontColor('green');
    sheet.getRange('J3').setValue(form.getEditUrl()).setFontColor('green');
    
    //link form to new spreadsheet and place link to spreadsheet
    var responses_sheet = SpreadsheetApp.create("StudentPreferences_Responses");
    form.setDestination(FormApp.DestinationType.SPREADSHEET, responses_sheet.getId());
    sheet.getRange('J4').setValue(responses_sheet.getUrl());
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
    .setCollectEmail(true).setAllowResponseEdits(true);
    
    var name = form.addListItem();
    name.setTitle('Name')
    .setChoiceValues(tutors).setRequired(true);
    
    //set capacity
    var textItem = form.addTextItem().setTitle('Student capacity').setRequired(true);
    var textValidation = FormApp.createTextValidation()
    .setHelpText('Input was not a number between 1 and 300.')
    .requireNumberBetween(1, 300)
    .build();
    textItem.setValidation(textValidation);
    
    
    var pageBreak = form.addPageBreakItem();
    var item = form.addSectionHeaderItem();
    item.setTitle('Please enter your preferences');
    //loop through and create dropdown for 50 preferences
    for (i = 1; i <= 50; i++) {
      var item = form.addListItem().setTitle(i.toString()).setChoiceValues(students);
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

function getPreferences() {
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  
  var query_data = {};
  query_data['student_prefs'] = {};
  query_data['college_prefs'] = {};
  query_data['college_capacity'] = {};
  
  //get Student Names
  var sheet = doc.getSheetByName("Names (Students)");  
  var students = sheet.getDataRange().getValues();
  numStudents = students.length;
  
  //get Tutor Names
  var sheet = doc.getSheetByName("Names (Tutors)");  
  var tutors = sheet.getDataRange().getValues();
  numTutors = tutors.length;
  
  //get student preferences
  var sheet = doc.getSheetByName("Student preferences");  
  var rows = sheet.getRange(2,2,numStudents,numTutors).getValues();
  
  //Put student preferences into query_data
  for(var s=0; s<numStudents;s++) { //for each student
    var studentName = students[s];
    query_data['student_prefs'][studentName] = []
    
    for(var t=0; t<numTutors; t++) { //from 1st to last place 
      query_data['student_prefs'][studentName][t] = rows[s][t];
    }
  }
  
  //get tutor preferences
  var sheet = doc.getSheetByName("Tutor preferences");  
  var rows = sheet.getRange(2,2,numTutors,numStudents).getValues();
  
  for(var t=0; t<numTutors;t++) { //for each tutor
    var tutorName = tutors[t];
    query_data['college_prefs'][tutorName] = []
    query_data['college_capacity'][tutorName] = rows[t][0];
    
    for(var s=0; s<numStudents; s++) { 
      if (rows[t][1+s] == "") {
        break;
      } //break if empty column
      
      query_data['college_prefs'][tutorName][s] = rows[t][1+s];
    }
  }
  return query_data;
}

function run() {
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
  var sheet = doc.getSheetByName("__assignments");  
  
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
}


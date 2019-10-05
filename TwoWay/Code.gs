function transpose(a)
{
  return Object.keys(a[0]).map(function (c) { return a.map(function (r) { return r[c]; }); });
}

// Takes an Array that contains only Strings and Numbers and produces a column vector
Object.defineProperty(Object.prototype, "toColumnVector", {value: function(){
  if (this.constructor !== Array){
    throw("typeError", "The object is not an array");
  }
  var output = [];
  for (var i = 0; i < this.length; i++){
    if (typeof this[i] !== "number" && typeof this[i] !== "string"){
      throw("typeError", "The element is not a number or string");
    } else {
      output.push([this[i]]); 
    }
  }
  return output;
}});

// Takes an Array that contains only Strings and Numbers and produces a row vector
Object.defineProperty(Object.prototype, "toRowVector", {value: function(){
  if (this.constructor !== Array){
    throw("typeError", "The object is not an array");
  }
  for (var i = 0; i < this.length; i++){
    if (typeof this[i] !== "number" && typeof this[i] !== "string"){
      throw("typeError", "The element is not a number or string");
    }
  }
  return [this];  
}});

function columnToRowVector(columnVector){
  var rowVector = [];
  rowVector.push([]);
  for (var row = 0; row < columnVector.length; ++row){
    rowVector[0].push(columnVector[row][0])
  }
  return rowVector;
}

function createInterviewSchedule(){
  
  
  var confirm = Browser.msgBox('Confirmation','Are you sure you want to create the schedule? This will overwrite the old schedule and will mean having to re-do the conflicts', Browser.Buttons.OK_CANCEL);
  if(confirm=='ok'){ 
    
    
    var doc = SpreadsheetApp.getActiveSpreadsheet();
    
    //get tutor names & student preferences
    var tutors = doc.getSheetByName("Names (Tutors)").getDataRange().getValues();  
    var scheduleSheet = doc.getSheetByName("Interview Schedule");
    scheduleSheet.getRange('B2:O').clearContent();
    scheduleSheet.getRange('B3:O').clearFormat();
    scheduleSheet.getRange(2, 2, 1, tutors.length).setValues(transpose(tutors));
    
    
    var sheet = doc.getSheetByName("START HERE");  
    var studentResponseURL = sheet.getRange('J4').getValue();
    studentPrefsSheet = SpreadsheetApp.openByUrl(studentResponseURL).getSheetByName('Form Responses 1');
    studentPreferences = studentPrefsSheet.getDataRange().getValues();
    var numResponses = studentPreferences.length - 1; 
    
    
    //get preferences for ADS0
    for (tut=0;tut<tutors.length;tut++) {
      
      
      var prev_length = 0;
      for (rank=0;rank<10;rank++) {
        
        if (prev_length<60){
          var NameList = [];
          
          for (resp=0;resp<numResponses;resp++) {
            var studentName = studentPreferences[1+resp][2];
            var studentYear = studentPreferences[1+resp][3];
            var tutorChoice = studentPreferences[1+resp][4+rank];
            
            if (tutorChoice==tutors[tut]) {
              NameList.push('('+ (studentYear.toString()=="One" ? 1 : 2) +') ' + studentName);
            }
          }
          
          if (NameList.length>0) 
          {
            
            range = scheduleSheet.getRange(3+prev_length, 2+tut, NameList.length);
            range.setValues(NameList.toColumnVector());
            prev_length += NameList.length;
            range.setBackgroundRGB(255, 245 - 20*rank, 245 - 20*rank);
          }
          
        }
      }
      
    }
  }
}

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
    var doc = SpreadsheetApp.getActiveSpreadsheet();
    
    //get student preferences
    var sheet = doc.getSheetByName("START HERE");  
    var studentResponseURL = sheet.getRange('J4').getValue();
    studentPreferences = SpreadsheetApp.openByUrl(studentResponseURL).getSheetByName('Form Responses 1').getDataRange().getValues();
    var numResponses = studentPreferences.length - 1;
    
    var sheet = doc.getSheetByName("Names (Tutors)");  
    var tutors = sheet.getDataRange().getValues();
    numTutors = tutors.length;
    
    //split student names by year group
    var YearOneStudents = [];
    var YearTwoStudents = [];
    for (s=0;s<numResponses;s++) {
      if (studentPreferences[1+s][3]=="One") { 
        YearOneStudents.push(studentPreferences[1+s][2]);
      } else { 
        YearTwoStudents.push(studentPreferences[1+s][2]);
      }  
    }
    
    //Create form for tutors to choose top students
    var form = FormApp.create('Tutor Preferences')
    .setCollectEmail(true).setAllowResponseEdits(true).setRequireLogin(false);;
    
    var name = form.addListItem();
    name.setTitle('Name')
    .setChoiceValues(tutors).setRequired(true);
    
    
    form.addPageBreakItem();
    
    var item = form.addSectionHeaderItem();
    item.setTitle('Enter your preferences for YEAR ONE ');
    for (i = 1; i <= 30; i++) {
      var item = form.addListItem().setTitle('Y1 Rank '+i.toString()).setChoiceValues(YearOneStudents);
    }
    
    form.addPageBreakItem();
    
    var item = form.addSectionHeaderItem();
    item.setTitle('Enter your preferences for YEAR TWO ');
    for (i = 1; i <= 30; i++) {
      var item = form.addListItem().setTitle('Y2 Rank '+i.toString()).setChoiceValues(YearTwoStudents);
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
  
  var year_one_data = {};
  year_one_data['student_prefs'] = {};
  year_one_data['college_prefs'] = {};
  year_one_data['college_capacity'] = {};
  
  var year_two_data = {};
  year_two_data['student_prefs'] = {};
  year_two_data['college_prefs'] = {};
  year_two_data['college_capacity'] = {};
  
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
    var studentYear = studentPreferences[1+s][3];
    
    if (studentYear=="One") {
      year_one_data['student_prefs'][studentName] = []
      for(var t=0; t<numTutors; t++) { //from 1st to last place 
        year_one_data['student_prefs'][studentName][t] = studentPreferences[1+s][4+t];
      }
    } else if (studentYear=="Two") {
      year_two_data['student_prefs'][studentName] = []
      for(var t=0; t<numTutors; t++) { //from 1st to last place 
        year_two_data['student_prefs'][studentName][t] = studentPreferences[1+s][4+t];
      }
    }
  }
  
  //Load tutor preferences from the form responses
  var sheet = doc.getSheetByName("START HERE");  
  var tutorResponseURL = sheet.getRange('J7').getValue();
  tutorPreferences = SpreadsheetApp.openByUrl(tutorResponseURL).getSheetByName('Form Responses 1').getDataRange().getValues();
  var numResponses = tutorPreferences.length - 1;
  
  var sheet = doc.getSheetByName("Assignments (List)");  
  var YearOneCapacity = sheet.getRange('M2').getValue();
  var YearTwoCapacity = sheet.getRange('M3').getValue();
  
  //Put tutor preferences into query_data
  for(var t=0; t<numResponses;t++) { //for each tutor
    var tutorName = tutorPreferences[1+t][2];
    
    year_one_data['college_prefs'][tutorName] = []
    year_one_data['college_capacity'][tutorName] = YearOneCapacity;
    year_two_data['college_prefs'][tutorName] = []
    year_two_data['college_capacity'][tutorName] = YearTwoCapacity;
    
    for(var s=0; s<30; s++) { 
      if (tutorPreferences[1+t][3+s] == "") {
        break;
      } //break if empty column
      
      year_one_data['college_prefs'][tutorName][s] = tutorPreferences[1+t][3+s];
    }
    
    for(var s=0; s<30; s++) { 
      if (tutorPreferences[1+t][33+s] == "") {
        break;
      } //break if empty column
      
      year_two_data['college_prefs'][tutorName][s] = tutorPreferences[1+t][33+s];
    }
  }
  
  return [year_one_data, year_two_data];
  
}

function computeAssignments() {
  //check that there are no duplicates
  if (checkStudentPrefsForDuplicates() | checkTutorPrefsForDuplicates()) {
    Browser.msgBox('Duplicate names in preferences','There are duplicate entries in the student and/or tutor preferences sheets (highlighted red). Please fix these and try again.', Browser.Buttons.OK);
    
  } else {
    
    //get Tutor Names
    var doc = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = doc.getSheetByName("Names (Tutors)");  
    var tutors = sheet.getDataRange().getValues();
    numTutors = tutors.length;
    
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
    var data = getPreferences();
    var year_one_data = data[0];
    var year_two_data = data[1];
    
    //year 1 assignments
    var options = {
      'method' : 'post',
      'contentType': 'application/json',
      'headers': {"Authorization": "Basic " + Utilities.base64Encode(username + ":" + password)},
      'payload' : JSON.stringify(year_one_data)
    };
    var response = UrlFetchApp.fetch(url, options);
    var json = response.getContentText(); // get the response content as text
    var dataSet = JSON.parse(json); //parse text into json
    dataSet = dataSet.hri_matching;
    
    //write data to Output sheet
    var sheet = doc.getSheetByName("Assignments (List)");  
    var range = sheet.getRange("A2:C300");
    range.clearContent();
    
    var rows = [], data;
    for (i = 0; i < dataSet.length; i++) {
      data = dataSet[i];
      var studentName = data['student.y'];
      var tutorName = data['college.y'];
      
      //identify the rank of the assignment
      for (t=0;t<numTutors;t++) {
        
        if (year_one_data['student_prefs'][studentName][t] == tutorName) {
          var rank = t+1;
        }
      }
      rows.push([studentName, tutorName, rank]);
    }
    dataRange = sheet.getRange(2, 1, rows.length, 3);
    dataRange.setValues(rows);
    dataRange.sort([{column: 1, ascending: true}]);
    
    //year 2 assignments
    var options = {
      'method' : 'post',
      'contentType': 'application/json',
      'headers': {"Authorization": "Basic " + Utilities.base64Encode(username + ":" + password)},
      'payload' : JSON.stringify(year_two_data)
    };
    var response = UrlFetchApp.fetch(url, options);
    var json = response.getContentText(); // get the response content as text
    var dataSet = JSON.parse(json); //parse text into json
    dataSet = dataSet.hri_matching;
    
    //write data to Output sheet
    var doc = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = doc.getSheetByName("Assignments (List)");  
    var range = sheet.getRange("E2:G300");
    range.clearContent();
    
    var rows = [], data;
    for (i = 0; i < dataSet.length; i++) {
      data = dataSet[i];
      var studentName = data['student.y'];
      var tutorName = data['college.y'];
      
      //identify the rank of the assignment
      for (t=0;t<numTutors;t++) {
        
        if (year_two_data['student_prefs'][studentName][t] == tutorName) {
          var rank = t+1;
        }
      }
      rows.push([studentName, tutorName, rank]);
    }
    dataRange = sheet.getRange(2, 5, rows.length, 3);
    dataRange.setValues(rows);
    dataRange.sort([{column: 5, ascending: true}]);
    
    doc.toast('Assignments computed');
  }
}


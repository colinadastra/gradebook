const checkmark = "☑";
const uncheckmark = "☐"
var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
var mpSheet = spreadsheet.getSheetByName("Marking Periods");
var documentProperties = PropertiesService.getDocumentProperties();
const categoryNames = ["Assessments","Classwork","Ungraded"];
const mpNamesRaw = mpSheet.getRange("A2:A").getValues();
const mpNames = [];
for (var mp of mpNamesRaw) mpNames.push(mp[0]);

var mpValues = mpSheet.getRange("A:C").getValues();

/** 
 * COLORS
 */

class AccentColor {
  constructor(dark, medium, light, name) {
    this.dark = dark;
    this.medium = medium;
    this.light = light;
    this.name = name;
  }
}

var rawColorInfo = [
  ["#B71C1C","#D32F2F","#F44336","Red"],
  ["#BF360C","#D84315","#FF5722","Rust"],
  ["#E65100","#F57C00","#FF9800","Orange"],
  ["#FFD600","#FFEA00","#FFFF00","Yellow"],
  ["#33691E","#558B2F","#7CB342","Pear"],
  ["#1B5E20","#2E7D32","#43A047","Emerald"],
  ["#004D40","#00695C","#00897B","Teal"],
  ["#006064","#007781","#0097A7","Cyan"],
  ["#0D47A1","#1565C0","#1E88E5","Blue"],
  ["#1A237E","#283593","#3F51B5","Indigo"],
  ["#311B92","#4527A0","#673AB7","Violet"],
  ["#6A1B9A","#8E24AA","#AB47BC","Purple"],
  ["#880E4F","#AD1457","#D81B60","Fuschia"],
  ["#4E342E","#5D4037","#795548","Brown"],
  ["#212121","#424242","#616161","Gray"]
];

var colors = [];

for (var row of rawColorInfo) {
  colors.push(new AccentColor(row[0],row[1],row[2],row[3]));
}

function setColor(targetColor) {
  var courseSheet = SpreadsheetApp.getActiveSheet();
  if (courseSheet.getName() == "Marking Periods") {
    return;
  }
  for (var color of colors) {
    if (color.name == targetColor) {
      courseSheet.getRange("1:1").setBackground(color.dark).setBorder(true, true, true, true, true, true, color.dark, SpreadsheetApp.BorderStyle.SOLID);
      courseSheet.getRange("2:2").setBackground(color.medium).setBorder(false, true, true, true, true, true, color.medium, SpreadsheetApp.BorderStyle.SOLID);
      courseSheet.getRange("3:3").setBackground(color.light).setBorder(false, true, true, true, true, true, color.light, SpreadsheetApp.BorderStyle.SOLID);
      courseSheet.setTabColor(color.light);
      if (color.name == "Yellow") {
        courseSheet.getRange("1:3").setFontColor('black');
      } else {
        courseSheet.getRange("1:3").setFontColor('white');
      }
      return;
    }
  }
}

/** 
 * CLASSROOM SYNC
 */

function createClassroomMappings() { 
  var response = Classroom.Courses.list({courseStates:"ACTIVE", teacherId:"me"});
  var classes = response.courses;
  while (response.nextPageToken != null) {
    response = Classroom.Courses.list({courseStates:"ACTIVE", teacherId:"me", pageToken:response.nextPageToken});
    classes.concat(response.courses);
  }
  var mappings = [];
  for (var classroom of classes) {
    Logger.log("Course Id: %s\t Course Name: %s", classroom.id, classroom.name);
    var classSheet = spreadsheet.getSheetByName(classroom.name);
    if (classSheet == null) {
      classSheet = createCourseSheet_(classroom.id, classroom.name);
    }
    mappings.push({"courseId": classroom.id, "courseName": classroom.name, "sheetId": classSheet.getSheetId()})
  }
  documentProperties.setProperty("classroomMappings",JSON.stringify(mappings));
}

function createCourseSheet_(classroomId, classroomName) {
  var template = spreadsheet.getSheetByName("Template");
  var courseSheet = template.copyTo(spreadsheet).setName(classroomName);
  syncStudents_(courseSheet,classroomId);
  courseSheet.getRange("A1").setValue(classroomName);
  var url = Classroom.Courses.get(classroomId).alternateLink;
  courseSheet.getRange("A3").setValue("=hyperlink(\""+url+"\", \"OPEN CLASSROOM\")");
  return courseSheet;
}

function exists_(array, searchValue) {
  return array.some(row => row.includes(searchValue));
}

function syncStudents_(courseSheet, classroomId) {
  var response = Classroom.Courses.Students.list(classroomId);
  var students = response.students;
  while (response.nextPageToken != null) {
    response = Classroom.Courses.Students.list(classroomId,{pageToken:response.nextPageToken});
    students.concat(response.students);
  }
  var existingStudentIds = courseSheet.getRange("D7:D").getValues();
  for (var student of students) {
    if (!exists_(existingStudentIds, student.profile.emailAddress)) {
      courseSheet.insertRowAfter(courseSheet.getMaxRows());
      var oldRow = courseSheet.getRange(courseSheet.getMaxRows()-1,1,1,8);
      var newRow = courseSheet.getRange(courseSheet.getMaxRows(),1,1,8);
      oldRow.copyTo(newRow);
      newRow = courseSheet.getRange(courseSheet.getMaxRows(),1,1,4);
      newRow.setValues([[student.profile.name.familyName, student.profile.name.givenName, "", student.profile.emailAddress]]);
    }
  }
  var firstStudent = courseSheet.getRange(7,1,1).getValues()[0];
  if (firstStudent[0] == "LastName") {
    courseSheet.deleteRow(7);
  }
  courseSheet.sort(1);
}

function syncRoster() {
  var courseSheet = spreadsheet.getActiveSheet();
  var mappings = JSON.parse(documentProperties.getProperty("classroomMappings"));
  for (var course of mappings) {
    if (courseSheet.getName() == course.courseName && courseSheet.getSheetId() == course.sheetId) {
      syncStudents_(courseSheet, course.courseId);
      break;
    }
  }
}

function syncAllRosters() {
  var mappings = JSON.parse(documentProperties.getProperty("classroomMappings"));
  var sheetList = spreadsheet.getSheets();
  for (var course of mappings) {
    for (var sheet of sheetList) {
      if (sheet.getName() == course.courseName && sheet.getSheetId() == course.sheetId) {
        spreadsheet.setActiveSheet(sheet);
        syncStudents_(sheet, course.courseId);
        break;
      }
    }
  }
}

function syncCoursework_(courseSheet, classroomId) {
  var response = Classroom.Courses.CourseWork.list(classroomId);
  var assignments = response.courseWork;
  while (response.nextPageToken != null) {
    response = Classroom.Courses.CourseWork.list(classromId,{pageToken: response.nextPageToken});
    assignments.concat(response.courseWork);
  }
  var existingAssignmentNames = courseSheet.getRange("J2:2").getValues()[0];
  var pulledAssignmentNames = [];
  for (var assignment of assignments) {
    pulledAssignmentNames.push(assignment.title);
    var category;
    try {
      category = assignment.gradeCategory.name;
    } catch (e) {
      category = "";
    }
    var dueDate = formatDate_(assignment.dueDate);
    var column;
    /** TODO: handle duplicated assignments */
    if (!existingAssignmentNames.includes(assignment.title)) {
      courseSheet.insertColumnBefore(10);
      var oldColumn = courseSheet.getRange("K1:K6");
      var newColumn = courseSheet.getRange("J1:J6");
      oldColumn.copyTo(newColumn);
      newColumn.setValues([[dueDate],
        ["=hyperlink(\""+assignment.alternateLink+"\",\""+assignment.title+"\")"],
        [assignment.maxPoints],
        [category],
        [(assignment.maxPoints>0)?guessMarkingPeriod_(dueDate,assignment.updateTime):""],
        [(assignment.maxPoints>0)?"=iferror(AVERAGE(BB7:BB),\"\")":""]
      ]);
      courseSheet.getRange("J2").setFontLine('none');
      newColumn.setFontLine("none");
      column = 10;
    } else {
      column = existingAssignmentNames.indexOf(assignment.title)+10;
      courseSheet.getRange(1,column,4,1).setValues([[dueDate],
        ["=hyperlink(\""+assignment.alternateLink+"\",\""+assignment.title+"\")"],
        [assignment.maxPoints],
        [category]
      ]);
    }
    if (assignment.maxPoints <= 0 || assignment.maxPoints == null) {
      courseSheet.getRange(7,column,courseSheet.getMaxRows()-7,1).setBackground("#d9d9d9");
      courseSheet.getRange(4,column).setValue("Ungraded");
    } else {
      courseSheet.getRange(7,column,courseSheet.getMaxRows()-7,1).setBackground(null);
    }
    /** TODO: individual student assignees */
  }
  var newExistingAssignments = courseSheet.getRange("J2:2").getValues()[0];
  for (var i = 0; i < newExistingAssignments.length; i++) {
    if (!pulledAssignmentNames.includes(newExistingAssignments[i])) {
      courseSheet.getRange(1,i+10,courseSheet.getMaxRows(),1).setFontLine('line-through');
      courseSheet.getRange(2,i+10).setNote("Assignment not found in Classroom. Consider removing from this sheet; otherwise, it will still count towards the averages."+courseSheet.getRange(2,i+10).getNote());
    }
    if (newExistingAssignments[i] == "Assignment Title") {
      courseSheet.deleteColumn(i+10);
    }
  }
}

function syncAssignments() {
  var courseSheet = spreadsheet.getActiveSheet();
  var mappings = JSON.parse(documentProperties.getProperty("classroomMappings"));
  for (var course of mappings) {
    if (courseSheet.getName() == course.courseName && courseSheet.getSheetId() == course.sheetId) {
      syncCoursework_(courseSheet, course.courseId);
      break;
    }
  }
}

function syncAllAssignments() {
  var mappings = JSON.parse(documentProperties.getProperty("classroomMappings"));
  var sheetList = spreadsheet.getSheets();
  for (var course of mappings) {
    for (var sheet of sheetList) {
      if (sheet.getName() == course.courseName && sheet.getSheetId() == course.sheetId) {
        spreadsheet.setActiveSheet(sheet);
        syncCoursework_(sheet, course.courseId);
        break;
      }
    }
  }
}

function guessMarkingPeriod_(dueDate,updateTime) {
  if (dueDate == null || dueDate == "") {
    dueDate = updateTime;
  }
  for (var i = 1; i < mpValues.length; i++) {
    if (Date.parse(dueDate) > Date.parse(mpValues[i][1]) && Date.parse(dueDate) < Date.parse(mpValues[i][2])) {
      return mpValues[i][0];
    }
  }
  return "";
}

function formatDate_(classroomDate) {
  try {
    return classroomDate.month + "/" + classroomDate.day + "/" + classroomDate.year;
  } catch (e) {
    return "";
  }
}

/** TODO: these bad bois */
function syncGrades_(classroomId, courseWorkId, courseSheet, column) {
  var response = JSON.parse(Classroom.Courses.CourseWork.StudentSubmissions.list(classroomId,courseWorkId));
  var submissions = response.studentSubmissions;
  while (response.nextPageToken != null) {
    response = JSON.parse(Classroom.Courses.CourseWork.StudentSubmissions.list(classroomId,courseWorkId,{pageToken: response.nextPageToken}));
    submissions.concat(response.studentSubmissions);
  }
  var studentEmails = courseSheet.getRange("D7:D").getValues();
  var gradeRange = courseSheet.getRange(7,column,courseSheet.getMaxRows()-6,1);
  var grades = gradeRange.getValues();
  var backgrounds = gradeRange.getBackgrounds();
  var styles = gradeRange.getFontStyles();
  rowLoop: for (var row = 0; row < studentEmails.length; row++) {
    var userId = Classroom.UserProfiles.get(studentEmails[row]).id;
    grades[row] = [null];
    backgrounds[row] = ["#d9d9d9"];
    styles[row] = ['normal'];
    for (var i = 0; i < submissions.length; i++) {
      if (submissions[i].userId == userId) {
        grades[row] = [submissions[i].assignedGrade];
        if (submissions[i].assignedGrade == null && submissions[i].draftGrade != null) {
          grades[row] = [submissions[i].draftGrade];
          styles[row] = ['italic'];
        } else {
          styles[row] = ['normal'];
        }
        backgrounds[row] = [null];
        continue rowLoop;
      }
    }
  }
  gradeRange.setValues(grades);
  gradeRange.setBackgrounds(backgrounds);
  gradeRange.setFontStyles(styles);
  SpreadsheetApp.flush();
}

function syncAssignmentGrades() {
  var courseSheet = spreadsheet.getActiveSheet();
  var column = spreadsheet.getCurrentCell().getColumn();
  var selectedRanges = spreadsheet.getSelection().getActiveRangeList().getRanges();
  for (var range of selectedRanges) {
    for (var column = range.getColumn(); column <= range.getLastColumn(); column++) {
      var targetAssignment = courseSheet.getRange(2,column).getValue();
      var mappings = JSON.parse(documentProperties.getProperty("classroomMappings"));
      for (var course of mappings) {
        if (courseSheet.getName() == course.courseName && courseSheet.getSheetId() == course.sheetId) {
          var courseworkId;
          var response = Classroom.Courses.CourseWork.list(course.courseId);
          var assignments = response.courseWork;
          while (response.nextPageToken != null) {
            response = Classroom.Courses.CourseWork.list(course.courseId,{pageToken: response.nextPageToken});
            assignments.concat(response.courseWork);
          }
          for (var assignment of assignments) {
            if (assignment.title == targetAssignment) {
              courseworkId = assignment.id;
            }
          }
          if (courseworkId == null) {
            Logger.log("Assignment %s not found in spreadsheet %s.", assignment.title, courseSheet.getName());
            continue;
          }
          syncGrades_(course.courseId,courseworkId,courseSheet,column);
        }
      }
    }
  }
  
}

function syncAllGrades(courseSheet) {
  if (courseSheet == null) {
    var courseSheet = spreadsheet.getActiveSheet();
  }
  for (var column = 10; column <= courseSheet.getMaxRows(); column++) {
    courseSheet.setCurrentCell(courseSheet.getRange(1, column));
    syncAssignmentGrades();
  }
}

function syncAllGradesAllClasses() {
  var courseSheets = spreadsheet.getSheets();
  for (var courseSheet of courseSheets) {
    if (courseSheet.getName() == "Template" || courseSheet.getName() == "Marking Periods") continue;
    spreadsheet.setActiveSheet(courseSheet);
    syncAllGrades(courseSheet);
  }
}

/** 
 * FILTERING AND MENUS
 */

class DynamicMenu {
  constructor() {
    const categoryParams = categoryNames;
    const mpParams = mpNames;
    const colorParams = [];
    for (var color of colors) {
      colorParams.push(color.name);
    }

    this.createMenu = (ui) => {
      const menu = ui.createMenu('Gradebook');
      const categoryMenu = ui.createMenu('Filter by category');
      const mpMenu = ui.createMenu('Filter by marking period');
      const colorMenu = ui.createMenu('Set gradebook color')
      categoryParams.forEach(param => {
        const functionName = `toggleCat${param}`;
        const entryName = (getCategoryVisibility_(`${param}`) ? checkmark : uncheckmark) + ` ${param}`;
        categoryMenu.addItem(entryName, `menuActions.${functionName}`);
      })
      mpParams.forEach(param => {
        const functionName = `toggleMp${param}`;
        const entryName = (getMpVisibility_(`${param}`) ? checkmark : uncheckmark) + ` ${param}`;
        mpMenu.addItem(entryName, `menuActions.${functionName}`);
      })
      colorParams.forEach(param => {
        const functionName = `setColor${param}`;
        const entryName = `${param}`;
        colorMenu.addItem(entryName, `menuActions.${functionName}`);
      })
      const syncStudentsMenu = ui.createMenu('Sync student roster')
        .addItem('For this class','syncRoster')
        .addItem('For all classes','syncAllRosters');
      const syncGradesMenu = ui.createMenu('Sync grades')
        .addItem('For selected assignment(s)','syncAssignmentGrades')
        .addItem('For all assignments in this class','syncAllGrades')
        .addItem('For all classes', 'syncAllGradesAllClasses');
      const syncAssignmentsMenu = ui.createMenu('Sync assignments')
        .addItem('For this class','syncAssignments')
        .addItem('For all classes','syncAllAssignments');
      menu.addSubMenu(categoryMenu);
      menu.addSubMenu(mpMenu);
      menu.addSeparator();
      menu.addSubMenu(syncAssignmentsMenu);
      menu.addSubMenu(syncGradesMenu);
      menu.addSubMenu(syncStudentsMenu);
      menu.addSeparator();
      menu.addItem('Initial gradebook setup','createClassroomMappings');
      menu.addSubMenu(colorMenu);
      menu.addToUi();
    }

    this.createActions = () => {
      const menuActions = {};
      categoryParams.forEach(param => {
        const functionName = `toggleCat${param}`;
        menuActions[functionName] = function () { toggleCategory(param) };
      });
      mpParams.forEach(param => {
        const functionName = `toggleMp${param}`;
        menuActions[functionName] = function () { toggleMp(param) };
      });
      colorParams.forEach(param => {
        const functionName = `setColor${param}`;
        menuActions[functionName] = function () { setColor(param) };
      });
      return menuActions;
    }
  }
}

const menu = new DynamicMenu();
const menuActions = menu.createActions();

function onOpen() {
  setupCategoryVisibility_();
  setupMpVisibility_();
  menu.createMenu(SpreadsheetApp.getUi());
}

function readDocumentProperties() {
  Logger.log(documentProperties.getProperties());
}

function getCategoryVisibility_(name) {
  return JSON.parse(documentProperties.getProperty("categoryVisibility"))[name];
}

function getMpVisibility_(name) {
  return JSON.parse(documentProperties.getProperty("mpVisibility"))[name];
}

function resetVisibilities() {
  documentProperties.deleteProperty("categoryVisibility");
  documentProperties.deleteProperty("mpVisibility");
}

function setupCategoryVisibility_() {
  var categoryVisibility = documentProperties.getProperty("categoryVisibility");
  if (categoryVisibility == null) {
    categoryVisibility = {};
    for (var category of categoryNames) {
      categoryVisibility[category] = true;
    }
    documentProperties.setProperty("categoryVisibility",JSON.stringify(categoryVisibility));
    return categoryVisibility;
  } else return categoryVisibility;
}

function setupMpVisibility_() {
  var mpVisibility = documentProperties.getProperty("mpVisibility");
  if (mpVisibility == null) {
    mpVisibility = {};
    for (var mp of mpNames) {
      mpVisibility[mp] = true;
    }
    documentProperties.setProperty("mpVisibility",JSON.stringify(mpVisibility));
    return mpVisibility;
  } else return mpVisibility;
}

function toggleCategory(category) {
  var categoryVisibility = JSON.parse(documentProperties.getProperty("categoryVisibility"));
  categoryVisibility[category] = !categoryVisibility[category];
  documentProperties.setProperty("categoryVisibility",JSON.stringify(categoryVisibility));
  updateAssignmentVisibility_();
}

function toggleMp(mp) {
  var mpVisibility = JSON.parse(documentProperties.getProperty("mpVisibility"));
  mpVisibility[mp] = !mpVisibility[mp];
  documentProperties.setProperty("mpVisibility",JSON.stringify(mpVisibility));
  updateAssignmentVisibility_();
}

function updateAssignmentVisibility_() {
  var categoryVisibile = JSON.parse(documentProperties.getProperty("categoryVisibility"));
  var mpVisible = JSON.parse(documentProperties.getProperty("mpVisibility"));
  var classSheets = spreadsheet.getSheets();
  for (var classSheet of classSheets) {
    if (classSheet.getName() == "Marking Periods") continue;
    var assignmentCategories = classSheet.getRange("J4:4").getValues()[0];
    var assignmentMps = classSheet.getRange("J5:5").getValues()[0];
    for (var i = 0; i < assignmentCategories.length; i++) {
      if (categoryVisibile[assignmentCategories[i]] && (mpVisible[assignmentMps[i]] || assignmentMps[i]=="")) {
        classSheet.showColumns(i+10);
      } else {
        classSheet.hideColumns(i+10);
      }
    }
  }
  menu.createMenu(SpreadsheetApp.getUi());
}

/** 
 * Calculates the marking period averages for the given student.
 * @param {array} numerators The student's scores on the assignments.
 * @param {array} denominators The maximum scores for the assignments.
 * @param {array} categories The categories of the assignments.
 * @param {array} markingPeriods The marking periods of the assignments.
 * @param {string} targetMarkingPeriods The marking periods for which the average is desired.
 * @return The marking period average (weighted by category).
 * @customFunction
 */
function MPAVERAGES(numerators,denominators,categories,markingPeriods,targetMarkingPeriods) {
  try {
    var averages = [];
    for (var markingPeriod of targetMarkingPeriods[0]) {
      averages.push(MPAVERAGE(numerators,denominators,categories,markingPeriods,markingPeriod))
    }
    return [averages];
  } catch (e) {
    return MPAVERAGE(numerators,denominators,categories,markingPeriods,targetMarkingPeriods);
  }
}

/** 
 * Calculates the marking period average for the given student.
 * @param {array} numerators The student's scores on the assignments.
 * @param {array} denominators The maximum scores for the assignments.
 * @param {array} categories The categories of the assignments.
 * @param {array} markingPeriods The marking periods of the assignments.
 * @param {string} targetMarkingPeriod The marking period for which the average is desired.
 * @return The marking period average (weighted by category).
 * @customFunction
 */
function MPAVERAGE(numerators,denominators,categories,markingPeriods,targetMarkingPeriod) {
  numerators = numerators[0];
  denominators = denominators[0];
  categories = categories[0];
  markingPeriods = markingPeriods[0];
  var assessmentsNumerator = 0;
  var assessmentsDenominator = 0;
  var classworkNumerator = 0;
  var classworkDenominator = 0;
  for (var i = 0; i< numerators.length; i++) {
    if (markingPeriods[i]==targetMarkingPeriod && numerators[i] > 0) {
      if (categories[i]=="Assessments") {
        assessmentsNumerator += numerators[i];
        assessmentsDenominator += denominators[i];
      } else if (categories[i]=="Classwork") {
        classworkNumerator += numerators[i];
        classworkDenominator += denominators[i];
      }
    }
  }
  if (assessmentsDenominator == 0 && classworkDenominator == 0) {
    return "";
  } else if (assessmentsDenominator == 0) {
    return (classworkNumerator/classworkDenominator);
  } else if (classworkDenominator == 0) {
    return (assessmentNumerator/assessmentsDenominator);
  } else {
    return 0.6*(classworkNumerator/classworkDenominator) + 0.4*(assessmentsNumerator/assessmentsDenominator);
  }
}

/** 
 * Calculates the marking period averages for the given student in specific categories.
 * @param {array} numerators The student's scores on the assignments.
 * @param {array} denominators The maximum scores for the assignments.
 * @param {string} targetCategories The categories for which the average is desired.
 * @param {array} categories The categories of the assignments.
 * @param {array} markingPeriods The marking periods of the assignments.
 * @param {string} targetMarkingPeriod The marking period for which the average is desired.
 * @return The marking period average (weighted by category).
 * @customFunction
 */
function CATEGORYAVERAGES(numerators,denominators,categories,targetCategories,markingPeriods,targetMarkingPeriod) {
  try {
    var averages = [];
    for (var category of targetCategories[0]) {
      averages.push(CATEGORYAVERAGE(numerators,denominators,categories,category,markingPeriods,targetMarkingPeriod))
    }
    return [averages];
  } catch (e) {
    return CATEGORYAVERAGE(numerators,denominators,categories,targetCategories,markingPeriods,targetMarkingPeriod);
  }
}

/** 
 * Calculates the marking period average for the given student in a specific category.
 * @param {array} numerators The student's scores on the assignments.
 * @param {array} denominators The maximum scores for the assignments.
 * @param {string} targetCategory The category for which the average is desired.
 * @param {array} categories The categories of the assignments.
 * @param {array} markingPeriods The marking periods of the assignments.
 * @param {string} targetMarkingPeriod The marking period for which the average is desired.
 * @return The marking period average (weighted by category).
 * @customFunction
 */
function CATEGORYAVERAGE(numerators,denominators,categories,targetCategory,markingPeriods,targetMarkingPeriod) {
  numerators = numerators[0];
  denominators = denominators[0];
  categories = categories[0];
  markingPeriods = markingPeriods[0];
  var categoryNumerator = 0;
  var categoryDenominator = 0;
  for (var i = 0; i< numerators.length; i++) {
    if (markingPeriods[i]==targetMarkingPeriod && categories[i]==targetCategory && numerators[i] > 0) {
      categoryNumerator += numerators[i];
      categoryDenominator += denominators[i];
    }
  }
  if (categoryDenominator == 0) {
    return "";
  } else {
    return (categoryNumerator/categoryDenominator);
  }
}

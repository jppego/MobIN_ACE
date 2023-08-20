// http://wafflebytes.blogspot.com/2016/10/google-script-create-drop-down-list.html



// ----------------------------------------------------------------------------------------------------------
/*
Project: Data to populate FEUP's mobility IN changes to LA requests forms
Function: fetchFormID()
Description: Fetches the form IDs. 
Copyright: https://github.com/jppego
*/
function fetchFormID() {

    //-------------------------------------------------
    // Selects the current google sheets
    var ss = SpreadsheetApp.getActive(); // if the script is running from the spreadsheet with the info

    //-------------------------------------------------
    // Fetches the ID of the forms 

    // select the scripts sheet
    var scriptsSheet = ss.getSheetByName("SCRIPTS");


    // ID of the two forms (PT and EN)
    var formID_EN = scriptsSheet.getRange(1, 3).getValue(); // EN form
    var formID_PT = scriptsSheet.getRange(2, 3).getValue(); // PT form


    // https://stackoverflow.com/questions/39585021/return-multiple-values-and-access-them
    var results = [formID_EN, formID_PT];
    return results;

}


// ----------------------------------------------------------------------------------------------------------
/*
Project: Data to populate FEUP's mobility IN changes to LA requests forms
Function: updateCourses()
Description: Updates the course units lists in the google forms. 
Copyright: https://github.com/jppego
*/

function updateCourses() {

    //-------------------------------------------------
    // Fetches the ID of the forms 

    results = fetchFormID();

    formID_EN = results[0];
    formID_PT = results[1];

    //-------------------------------------------------
    // Selects the current google sheets
    var ss = SpreadsheetApp.getActive(); // if the script is running from the spreadsheet with the info


    // select the scripts sheet
    var scriptsSheet = ss.getSheetByName("SCRIPTS");

    // number of courses to populate the forms
    var opt_n_courses = scriptsSheet.getRange(3, 3).getValue(); // number of courses to insert in the dropdown lists

    //  if (opt_n_courses == "") 
    //  {
    //    Browser.msgBox(opt_n_courses, Browser.Buttons.OK_CANCEL);
    //  }

    // Turns OFF the google forms
    form_Responses_OFF()


    //http://wafflebytes.blogspot.com/2016/10/google-script-create-drop-down-list.html
    // ID of the two forms (PT and EN)
    // var formID_EN = "1nm_1cy3sSdUNofgiXhKI_NbBuL-oF2xl55zyCae03go";
    // var formID_PT=  "1qounq0H0I39N2iszx6zII07bJwM6B59EdO6aUhpeQaI";


    // call the forms 
    var form_EN = FormApp.openById(formID_EN);
    var form_PT = FormApp.openById(formID_PT);


    // fetches the lists in the form
    var coursesListItemArray_EN = form_EN.getItems(FormApp.ItemType.LIST);
    var coursesListItemArray_PT = form_PT.getItems(FormApp.ItemType.LIST);
    //var coursesList =coursesListItemArray_EN[0].asListItem(); //its the first list i want//


    //-------------------------------------------------
    // Fetches the course units values from the spreadsheet

    // select the course units sheet
    var coursesData = ss.getSheetByName("UC_IN");


    // grab the values from the sheet - use 2 to skip header row
    var courseColumn_PT = 11; // the courses list in PT is in the courseColumn_PT column
    var courseColumn_EN = 12; // the courses list in EN is in the courseColumn_EN column
    var coursesPT_Values = coursesData.getRange(2, courseColumn_PT, coursesData.getMaxRows() - 1).getValues();
    var coursesEN_Values = coursesData.getRange(2, courseColumn_EN, coursesData.getMaxRows() - 1).getValues();

    var courseNames_PT = [];
    var courseNames_EN = [];

    // convert the array ignoring empty cells
    var j = 0; // index to courseNames_EN array

    // https://stackoverflow.com/questions/10843768/in-apps-script-how-to-include-optional-arguments-in-custom-functions
    // number of courses to populate the forms
    if (opt_n_courses == "") {
        var n_courses = coursesEN_Values.length; //if you want to read the whole list
        // In fact, google form has a limit of 1000 options per dropdown list  https://issuetracker.google.com/issues/63395462
    } else {
        var n_courses = opt_n_courses; // fixed number of lines to read from the course units list
    }


    for (var i = 0; i < n_courses; i++) //include all lines in the spreadsheet 
        if (coursesEN_Values[i][0] != "") {
            courseNames_EN[j] = coursesEN_Values[i][0];
            courseNames_PT[j] = coursesPT_Values[i][0];
            j++; // increments the courseNames_EN array index
        }



    //-------------------------------------------------
    // Inserts the course units values as options for the form lists



    // There are 3 sets of lists (approved, deleted and added) in this form, with 10 lists each. 

    // Defines the number of the first list in each set and the number of list per set
    var courseList_1 = 4; // the first approved courses list is #4
    var ncourseList_1 = 10; // the number of approved courses list is 10

    var courseList_2 = courseList_1 + ncourseList_1; // the first deleted courses list is #14
    var ncourseList_2 = 10; // the number of deleted courses list is 10

    var courseList_3 = courseList_2 + ncourseList_2; // the first added courses list is #24
    var ncourseList_3 = 10; // the number of added courses list is 10


    // LA course units lists
    var courseList_counter = courseList_1 - 1; // sets the counter for course list of approved courses.  The array begins at 0, hence the -1

    // Copies the course units value to each list repeats ncourseList_1 x

    for (i = 0; i < ncourseList_1; i++) {

        // EN 
        // fetches the corresponding list from the form
        var coursesList = coursesListItemArray_EN[i + courseList_counter].asListItem();
        // populate the drop-down with the array data
        coursesList.setChoiceValues(courseNames_EN);

        // PT 
        // fetches the corresponding list from the form
        coursesList = coursesListItemArray_PT[i + courseList_counter].asListItem();
        // populate the drop-down with the array data
        coursesList.setChoiceValues(courseNames_PT);
    }


    // deleted course units lists
    courseList_counter = courseList_2 - 1; // sets the counter for course list of LA course.  The array begins at 0, hence the -1

    // Copies the course units value to each list repeats ncourseList_2 x

    for (i = 0; i < ncourseList_2; i++) {

        //EN
        // fetches the corresponding list from the form
        coursesList = coursesListItemArray_EN[i + courseList_counter].asListItem();
        // populate the drop-down with the array data
        coursesList.setChoiceValues(courseNames_EN);
        // PT 
        // fetches the corresponding list from the form
        coursesList = coursesListItemArray_PT[i + courseList_counter].asListItem();
        // populate the drop-down with the array data
        coursesList.setChoiceValues(courseNames_PT);
    }


    // added course units lists
    courseList_counter = courseList_3 - 1; // sets the counter for course list of LA course.  The array begins at 0, hence the -1

    // Copies the course units value to each list repeats ncourseList_3 x

    for (i = 0; i < ncourseList_3; i++) {

        //EN
        // fetches the corresponding list from the form
        coursesList = coursesListItemArray_EN[i + courseList_counter].asListItem();
        // populate the drop-down with the array data
        coursesList.setChoiceValues(courseNames_EN);
        // PT 
        // fetches the corresponding list from the form
        coursesList = coursesListItemArray_PT[i + courseList_counter].asListItem();
        // populate the drop-down with the array data
        coursesList.setChoiceValues(courseNames_PT);
    }



    // Turns ON the google forms
    form_Responses_ON()


}



// ----------------------------------------------------------------------------------------------------------
/*
Project: Data to populate FEUP's mobility IN changes to LA requests forms
Function: updateCountries()
Description: Updates the counties list in the google forms. 
Copyright: https://github.com/jppego
*/

function updateCountries() {

    //-------------------------------------------------
    // Fetches the ID of the forms 

    results = fetchFormID();

    formID_EN = results[0];
    formID_PT = results[1];

    // call the forms 
    var form_EN = FormApp.openById(formID_EN);
    var form_PT = FormApp.openById(formID_PT);

    // fetches the lists in the form
    var coursesListItemArray_EN = form_EN.getItems(FormApp.ItemType.LIST);
    var coursesListItemArray_PT = form_PT.getItems(FormApp.ItemType.LIST);

    // fetches the list for the countries names
    var countriesList_EN = coursesListItemArray_EN[0].asListItem(); // the countries list is #1
    var countriesList_PT = coursesListItemArray_PT[0].asListItem(); // the countries list is #1


    // identify the sheet where the data needed to populate the drop-down resides
    var ss = SpreadsheetApp.getActive(); // it assumes the countries list is in the current spreadsheet


    //-------------------------------------------------
    // Populate countries list

    // select the countries sheet
    var countriesData = ss.getSheetByName("COUNTRIES");


    // grab the values in the first column of the sheet - use 2 to skip header row
    var countriesEN_Values = countriesData.getRange(2, 2, countriesData.getMaxRows() - 1).getValues();
    var countriesPT_Values = countriesData.getRange(2, 1, countriesData.getMaxRows() - 1).getValues();


    var countriesNames_EN = [];
    var countriesNames_PT = [];


    // convert the array ignoring empty cells
    var j = 0; // index to courseNames array
    for (var i = 0; i < countriesEN_Values.length; i++)
        //for(var i = 0; i < 25; i++)   
        if (countriesEN_Values[i][0] != "") {
            countriesNames_EN[j] = countriesEN_Values[i][0];
            countriesNames_PT[j] = countriesPT_Values[i][0];
            j++;
        }

    // populate the drop-down with the array data
    countriesList_EN.setChoiceValues(countriesNames_EN);
    countriesList_PT.setChoiceValues(countriesNames_PT);

}


// ----------------------------------------------------------------------------------------------------------
/*
Project: Data to populate FEUP's mobility IN changes to LA requests forms
Function: form_Responses_OFF()
Description: Closes the forms
Copyright: https://github.com/jppego
*/
function form_Responses_OFF() {

    //https://developers.google.com/apps-script/reference/forms/form

    // Open a form by ID and create a new spreadsheet.

    //-------------------------------------------------
    // Fetches the ID of the forms 

    results = fetchFormID();

    formID_EN = results[0];
    formID_PT = results[1];

    // call the forms 
    var form_EN = FormApp.openById(formID_EN);
    var form_PT = FormApp.openById(formID_PT);

    // Update form properties via chaining.
    form_EN.setAcceptingResponses(false);
    form_PT.setAcceptingResponses(false);
}


// ----------------------------------------------------------------------------------------------------------
/*
Project: Data to populate FEUP's mobility IN changes to LA requests forms
Function: form_Responses_OFF()
Description: Opens the forms
Copyright: https://github.com/jppego
*/
function form_Responses_ON() {

    //https://developers.google.com/apps-script/reference/forms/form

    // Open a form by ID and create a new spreadsheet.

    //-------------------------------------------------
    // Fetches the ID of the forms 

    results = fetchFormID();

    formID_EN = results[0];
    formID_PT = results[1];

    // call the forms 
    var form_EN = FormApp.openById(formID_EN);
    var form_PT = FormApp.openById(formID_PT);

    // Update form properties via chaining.
    form_EN.setAcceptingResponses(true);
    form_PT.setAcceptingResponses(true);
}
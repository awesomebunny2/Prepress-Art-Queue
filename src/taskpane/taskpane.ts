/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

//#region IMAGE REFERENCES --------------------------------------------------------------------------------------
import { ContextReplacementPlugin } from "webpack";
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";
//#endregion ----------------------------------------------------------------------------------------------------

//#region GLOBAL VARIABLES --------------------------------------------------------------------------------------
var artistColumn = "S";
var moveEvent;
var sortEvent;
var sortColumn = "Priority";
var projectTypeColumn = "H";
var productColumn = "G";
//#endregion ----------------------------------------------------------------------------------------------

//#region TASKPANE BUTTONS ---------------------------------------------------------------------------------------
window.onload = function() { //Wait for the window to load, then do the following:
  document.getElementById("queue-btn").onclick = function addRequest() {
    document.getElementById("home").style.display = "none";
    document.getElementById("add-to-queue").style.display = "block";
  };

  document.getElementById("back-btn").onclick = function backToHome() {
    document.getElementById("add-to-queue").style.display = "none";
    document.getElementById("home").style.display = "block";
  };

  // Working here

  // var theDate = new Date();
  // theDate.setHours(17);
  // theDate.setMinutes(29);
  // theDate.setSeconds(0);
  


  // console.log(timeBetweenNowAndTomorrowMorning)

};
//#endregion ------------------------------------------------------------------------------------------------------

//#region CHECKBOX SETUP ________________________________________________________________________________________
/** When the checkbox is CHANGED */
$("#set-behavior").on("change", function() {

  // Is this set to checked?
  var checked = $(this).prop("checked");

  if (checked == true) { // Set the startup behavior!
    Office.addin.setStartupBehavior(Office.StartupBehavior.load); //when document opens, references startup behavioir in manifest, which automatically opens the taskpane
  } else { // Turn off the startup behavior!
    Office.addin.setStartupBehavior(Office.StartupBehavior.none); //when document opens, references startup behavioir in manifest, which automatically opens the taskpane
  }
})
//#endregion ----------------------------------------------------------------------------------------------------

//#region STARTUP BEHAVIOR --------------------------------------------------------------------------------------
Office.onReady((info) => {
  // Load on Startup
  // setStartupBehavior is **document level**
  var currentBehavior = Office.addin.getStartupBehavior().then(function(returned) {
    if (returned == "Load") {
      /* Check the checkbox */
      $("#set-behavior").prop("checked", true);
    } else {
      /* Uncheck the checkbox */
      $("#set-behavior").prop("checked", false);
    }
    console.log(returned);
  });
    if (info.host === Office.HostType.Excel) { //If application is Excel
      document.getElementById("sideload-msg").style.display = "none"; //Don't show side-loading message
      document.getElementById("app-body").style.display = "flex"; //Keep content in taskpane flexible to scaling, I think...
        
      Excel.run(async context => { //Do while Excel is running
        
        moveEvent = context.workbook.tables.onChanged.add(onTableChanged);

        // sortEvent = context.workbook.tables.onChanged.add(sortDate);

        return context.sync().then(function() { //Commits changes to document and then returns the console.log
          // console.log("Event handlers have been successfully registered");
        });
      });
    };
});
//#endregion ------------------------------------------------------------------------------------------------

//#region MOVES DATA BETWEEN WORKSHEETS ------------------------------------------------------------------------
async function onTableChanged(eventArgs: Excel.TableChangedEventArgs) { //This function will be using event arguments to collect data from the workbook

  await Excel.run(async (context) => {

    //#region EVENT VARIABLES -----------------------------------------------------------------------------------
    var details = eventArgs.details; //Loads the values before and after the event
    var address = eventArgs.address; //Loads the cell's address where the event took place
    var sheet = context.workbook.worksheets.getActiveWorksheet().load("name");
    var changedTable = context.workbook.tables.getItem(eventArgs.tableId).load("name"); //Returns tableId of the table where the event occured
    var regexStr = address.match(/[a-zA-Z]+|[0-9]+(?:\.[0-9]+|)/g); //Separates the column letter(s) from the row number for the address: presented as a string
    var changedColumn = regexStr[0]; //The first instance of the separated address array, being the column letter(s)
    var changedRow = Number(regexStr[1]) - 2; //The second instance of the separated address array, being the row, converted into a number and subtracted by 2
    var myRow = changedTable.rows.getItemAt(changedRow).load("values"); //loads the values of the changed row in the table where the event was fired      
    //#endregion ------------------------------------------------------------------------------------------------

    //#region SPECIFIC TABLE VARIABLES --------------------------------------------------------------------------
      //#region UNASSIGNED PROJECTS VARIABLES ------------------------------------------------------------
      var unassignedSheet = context.workbook.worksheets.getItem("Unassigned Projects");
      var unassignedTable = unassignedSheet.tables.getItem("UnassignedProjects");
      //#endregion --------------------------------------------------------------------------
      //#region MATT VARIABLES --------------------------------------------------------
      var mattSheet = context.workbook.worksheets.getItem("Matt");
      var mattTable = mattSheet.tables.getItem("MattProjects");
      //#endregion --------------------------------------------------------------------------
      //#region ALAINA VARIABLES ------------------------------------------------------
      var alainaSheet = context.workbook.worksheets.getItem("Alaina");
      var alainaTable = alainaSheet.tables.getItem("AlainaProjects");
      //#endregion --------------------------------------------------------------------------
      //#region BERTO VARIABLES ------------------------------------------------------
      var bertoSheet = context.workbook.worksheets.getItem("Berto");
      var bertoTable = bertoSheet.tables.getItem("BertoProjects");
      //#endregion ---------------------------------------------------------------------------
      //#region BRE B. VARIABLES ------------------------------------------------------
      var breBSheet = context.workbook.worksheets.getItem("Bre B.");
      var breBTable = breBSheet.tables.getItem("BreBProjects");
      //#endregion ---------------------------------------------------------------------------
      //#region CHRISTIAN VARIABLES ------------------------------------------------------
      var christianSheet = context.workbook.worksheets.getItem("Christian");
      var christianTable = christianSheet.tables.getItem("ChristianProjects");
      //#endregion ---------------------------------------------------------------------------
      //#region EMILY VARIABLES ------------------------------------------------------
      var emilySheet = context.workbook.worksheets.getItem("Emily");
      var emilyTable = emilySheet.tables.getItem("EmilyProjects");
      //#endregion ---------------------------------------------------------------------------
      //#region IAN VARIABLES ------------------------------------------------------
      var ianSheet = context.workbook.worksheets.getItem("Ian");
      var ianTable = ianSheet.tables.getItem("IanProjects");
      //#endregion ---------------------------------------------------------------------------
      //#region JEFF VARIABLES ------------------------------------------------------
      var jeffSheet = context.workbook.worksheets.getItem("Jeff");
      var jeffTable = jeffSheet.tables.getItem("JeffProjects");
      //#endregion ---------------------------------------------------------------------------
      //#region JOSH VARIABLES ------------------------------------------------------
      var joshSheet = context.workbook.worksheets.getItem("Josh");
      var joshTable = joshSheet.tables.getItem("JoshProjects");
      //#endregion ---------------------------------------------------------------------------
      //#region KRISTEN VARIABLES ------------------------------------------------------
      var kristenSheet = context.workbook.worksheets.getItem("Kristen");
      var kristenTable = kristenSheet.tables.getItem("KristenProjects");
      //#endregion ---------------------------------------------------------------------------
      //#region NICHOLE VARIABLES ------------------------------------------------------
      var nicholeSheet = context.workbook.worksheets.getItem("Nichole");
      var nicholeTable = nicholeSheet.tables.getItem("NicholeProjects");
      //#endregion ---------------------------------------------------------------------------
      //#region LUKE VARIABLES ------------------------------------------------------
      var lukeSheet = context.workbook.worksheets.getItem("Luke");
      var lukeTable = lukeSheet.tables.getItem("LukeProjects");
      //#endregion ---------------------------------------------------------------------------
      //#region LISA VARIABLES ------------------------------------------------------
      var lisaSheet = context.workbook.worksheets.getItem("Lisa");
      var lisaTable = lisaSheet.tables.getItem("LisaProjects");
      //#endregion ---------------------------------------------------------------------------
      //#region LUIS VARIABLES ------------------------------------------------------
      var luisSheet = context.workbook.worksheets.getItem("Luis");
      var luisTable = luisSheet.tables.getItem("LuisProjects");
      //#endregion ---------------------------------------------------------------------------
      //#region PETER VARIABLES ------------------------------------------------------
      var peterSheet = context.workbook.worksheets.getItem("Peter");
      var peterTable = peterSheet.tables.getItem("PeterProjects");
      //#endregion ---------------------------------------------------------------------------
      //#region RITA VARIABLES ------------------------------------------------------
      var ritaSheet = context.workbook.worksheets.getItem("Rita");
      var ritaTable = ritaSheet.tables.getItem("RitaProjects");
      //#endregion ---------------------------------------------------------------------------
      //#region ETHAN VARIABLES ------------------------------------------------------
      var ethanSheet = context.workbook.worksheets.getItem("Ethan");
      var ethanTable = ethanSheet.tables.getItem("EthanProjects");
      //#endregion ---------------------------------------------------------------------------
      //#region BRE Z. VARIABLES ------------------------------------------------------
      var breZSheet = context.workbook.worksheets.getItem("Bre Z.");
      var breZTable = breZSheet.tables.getItem("BreZProjects");
      //#endregion ---------------------------------------------------------------------------
      //#region JOE VARIABLES ------------------------------------------------------
      var joeSheet = context.workbook.worksheets.getItem("Joe");
      var joeTable = joeSheet.tables.getItem("JoeProjects");
      //#endregion ---------------------------------------------------------------------------
      //#region JORDAN VARIABLES ------------------------------------------------------
      var jordanSheet = context.workbook.worksheets.getItem("Jordan");
      var jordanTable = jordanSheet.tables.getItem("JordanProjects");
      //#endregion ---------------------------------------------------------------------------
      //#region HAZEL-RAH VARIABLES ------------------------------------------------------
      var hazelSheet = context.workbook.worksheets.getItem("Hazel-Rah");
      var hazelTable = hazelSheet.tables.getItem("HazelProjects");
      //#endregion ---------------------------------------------------------------------------
      //#region TODD VARIABLES ------------------------------------------------------
      var toddSheet = context.workbook.worksheets.getItem("Todd");
      var toddTable = toddSheet.tables.getItem("ToddProjects");
      //#endregion ---------------------------------------------------------------------------
    //#endregion ------------------------------------------------------------------------------------------------



    //#region INITIATING THE MOVE EVENT -------------------------------------------------------------------------  
    var theChange = eventArgs.changeType; //Kind of change that was made
    if (theChange == "RangeEdited" && eventArgs.details !== undefined ) {
      
      // Ignore the moved-to table's on change event 


      
      console.log("The move data event has been initiated!!");
      
      if (eventArgs.details.valueAfter == eventArgs.details.valueBefore) {
        console.log("No values have changed. Exiting move data event...")
        return;
      };
    //#endregion ------------------------------------------------------------------------------------------------

    //#region MOVE CONDITIONS -----------------------------------------------------------------------------------
        
      await context.sync().then(function () { // WHAT IS LOOPING
        // console.log("Promise Fulfilled!");

        // console.log(myRow.values);

        var rowValues = myRow.values;

        // console.log("The active worksheet is " + sheet.name);



        // if (changedColumn == projectTypeColumn || productColumn) { //if updated data was in Project Type column, run the lookupStart function
          var projectTypeHours = lookupStart(rowValues, changedRow); //adds hours to turn-around time based on Project Type
          // console.log("The projectTypeHours are " + projectTypeHours + " hours");
          var productHours = preLookupWork(rowValues, projectTypeHours); //adds hours based on Product and adds to lookupStart output
          // console.log("The productHours are " + productHours + " hours");
          var workHoursAdjust = lookupWork(projectTypeHours, productHours); //takes prelookupWork variable and divides by 3 if lookupStart was equal to 2. Otherwise remains the same.
          // console.log("The workHoursAdjust are " + workHoursAdjust + " hours");
          var myDate = receivedAdjust(rowValues, changedRow); //grabs values from Added column and converts into date object in EST.
          // console.log("The added date within office hours is " + myDate);
          var override = startPreAdjust(rowValues, projectTypeHours, myDate); //adds manual override start hours to adjusted start time. Adjusts for office hours and weekends.
          // console.log("The date including the projectTypeHours and Start Override values is " + override);
          var startedPickedUpBy = startedBy(changedRow, sheet, override); //Prints the value of override to the Picked Up / Started By column and formats the date in a readible format.
          // console.log("The Started / Picked Up time is " + startedPickedUpBy);
          var workOverride = workPrePreAdjust(rowValues, workHoursAdjust, override); //Finds the value of Work Override in the changed row and adds it to workHoursAdjust, then adds that new number as hours to startedPickedUpBy. Formats to be within office hours and on a weekday if needed.
          // console.log("The started date adjusted with the Work Override is " + workOverride);
          var proofToClient = toClient(changedRow, sheet, workOverride); //Prints the value of workOverride to the Proof to Client column and formats the date in a readible format.
          // console.log("The date for Proof to Client is " + proofToClient);
        // }


        if (changedColumn == artistColumn) { //if updated data was in the Artist column, run the following code

          if (details.valueAfter == "Unassigned") {
            unassignedTable.rows.add(null, myRow.values); //Adds empty row to bottom of GreenBasket Table, then inserts the changed values into this empty row
            myRow.delete(); //Deletes the changed row from the original sheet
            console.log("Data was moved to the Unassigned Projects Table!");
            return;

          } else if (details.valueAfter == "Matt") {
            mattTable.rows.add(null, myRow.values); //Adds empty row to bottom of GreenBasket Table, then inserts the changed values into this empty row
            myRow.delete(); //Deletes the changed row from the original sheet
            console.log("Data was moved to the Matt Table!");
            return;
            
          } else if (details.valueAfter == "Alaina") {
            alainaTable.rows.add(null, myRow.values); //Adds empty row to bottom of GreenBasket Table, then inserts the changed values into this empty row
            myRow.delete(); //Deletes the changed row from the original sheet
            console.log("Data was moved to the Alaina Table!");
            return;            
          } else if (details.valueAfter == "Berto") {
            bertoTable.rows.add(null, myRow.values); //Adds empty row to bottom of GreenBasket Table, then inserts the changed values into this empty row
            myRow.delete(); //Deletes the changed row from the original sheet
            console.log("Data was moved to the Berto Table!");
            return;
          } else if (details.valueAfter == "Bre B.") {
            breBTable.rows.add(null, myRow.values); //Adds empty row to bottom of GreenBasket Table, then inserts the changed values into this empty row
            myRow.delete(); //Deletes the changed row from the original sheet
            console.log("Data was moved to the Bre B. Table!");
            return;
          } else if (details.valueAfter == "Christian") {
            christianTable.rows.add(null, myRow.values); //Adds empty row to bottom of GreenBasket Table, then inserts the changed values into this empty row
            myRow.delete(); //Deletes the changed row from the original sheet
            console.log("Data was moved to the Christian Table!");
            return;
          } else if (details.valueAfter == "Emily") {
            emilyTable.rows.add(null, myRow.values); //Adds empty row to bottom of GreenBasket Table, then inserts the changed values into this empty row
            myRow.delete(); //Deletes the changed row from the original sheet
            console.log("Data was moved to the Emily Table!");
            return;
          } else if (details.valueAfter == "Ian") {
            ianTable.rows.add(null, myRow.values); //Adds empty row to bottom of GreenBasket Table, then inserts the changed values into this empty row
            myRow.delete(); //Deletes the changed row from the original sheet
            console.log("Data was moved to the Ian Table!");
            return;
          } else if (details.valueAfter == "Jeff") {
            jeffTable.rows.add(null, myRow.values); //Adds empty row to bottom of GreenBasket Table, then inserts the changed values into this empty row
            myRow.delete(); //Deletes the changed row from the original sheet
            console.log("Data was moved to the Jeff Table!");
            return;
          } else if (details.valueAfter == "Josh") {
            joshTable.rows.add(null, myRow.values); //Adds empty row to bottom of GreenBasket Table, then inserts the changed values into this empty row
            myRow.delete(); //Deletes the changed row from the original sheet
            console.log("Data was moved to the Josh Table!");
            return;
          } else if (details.valueAfter == "Kristen") {
            kristenTable.rows.add(null, myRow.values); //Adds empty row to bottom of GreenBasket Table, then inserts the changed values into this empty row
            myRow.delete(); //Deletes the changed row from the original sheet
            console.log("Data was moved to the Kristen Table!");
            return;
          } else if (details.valueAfter == "Nichole") {
            nicholeTable.rows.add(null, myRow.values); //Adds empty row to bottom of GreenBasket Table, then inserts the changed values into this empty row
            myRow.delete(); //Deletes the changed row from the original sheet
            console.log("Data was moved to the Nichole Table!");
            return;
          } else if (details.valueAfter == "Luke") {
            lukeTable.rows.add(null, myRow.values); //Adds empty row to bottom of GreenBasket Table, then inserts the changed values into this empty row
            myRow.delete(); //Deletes the changed row from the original sheet
            console.log("Data was moved to the Luke Table!");
            return;
          } else if (details.valueAfter == "Lisa") {
            lisaTable.rows.add(null, myRow.values); //Adds empty row to bottom of GreenBasket Table, then inserts the changed values into this empty row
            myRow.delete(); //Deletes the changed row from the original sheet
            console.log("Data was moved to the Lisa Table!");
            return;
          } else if (details.valueAfter == "Luis") {
            luisTable.rows.add(null, myRow.values); //Adds empty row to bottom of GreenBasket Table, then inserts the changed values into this empty row
            myRow.delete(); //Deletes the changed row from the original sheet
            console.log("Data was moved to the Luis Table!");
            return;
          } else if (details.valueAfter == "Peter") {
            peterTable.rows.add(null, myRow.values); //Adds empty row to bottom of GreenBasket Table, then inserts the changed values into this empty row
            myRow.delete(); //Deletes the changed row from the original sheet
            console.log("Data was moved to the Peter Table!");
            return;
          } else if (details.valueAfter == "Rita") {
            ritaTable.rows.add(null, myRow.values); //Adds empty row to bottom of GreenBasket Table, then inserts the changed values into this empty row
            myRow.delete(); //Deletes the changed row from the original sheet
            console.log("Data was moved to the Rita Table!");
            return;
          } else if (details.valueAfter == "Ethan") {
            ethanTable.rows.add(null, myRow.values); //Adds empty row to bottom of GreenBasket Table, then inserts the changed values into this empty row
            myRow.delete(); //Deletes the changed row from the original sheet
            console.log("Data was moved to the Ethan Table!");
            return;
          } else if (details.valueAfter == "Bre Z.") {
            breZTable.rows.add(null, myRow.values); //Adds empty row to bottom of GreenBasket Table, then inserts the changed values into this empty row
            myRow.delete(); //Deletes the changed row from the original sheet
            console.log("Data was moved to the Bre Z. Table!");
            return;
          } else if (details.valueAfter == "Joe") {
            joeTable.rows.add(null, myRow.values); //Adds empty row to bottom of GreenBasket Table, then inserts the changed values into this empty row
            myRow.delete(); //Deletes the changed row from the original sheet
            console.log("Data was moved to the Joe Table!");
            return;
          } else if (details.valueAfter == "Jordan") {
            jordanTable.rows.add(null, myRow.values); //Adds empty row to bottom of GreenBasket Table, then inserts the changed values into this empty row
            myRow.delete(); //Deletes the changed row from the original sheet
            console.log("Data was moved to the Jordan Table!");
            return;
          } else if (details.valueAfter == "Hazel-Rah") {
            hazelTable.rows.add(null, myRow.values); //Adds empty row to bottom of GreenBasket Table, then inserts the changed values into this empty row
            myRow.delete(); //Deletes the changed row from the original sheet
            console.log("Data was moved to the Hazel-Rah Table!");
            return;
          } else if (details.valueAfter == "Todd") {
            toddTable.rows.add(null, myRow.values); //Adds empty row to bottom of GreenBasket Table, then inserts the changed values into this empty row
            myRow.delete(); //Deletes the changed row from the original sheet
            console.log("Data was moved to the Todd Table!");
            return;
          } else {
            console.log("Looks like there wasn't an Artist change this time. No data was moved...");
          } return;
        } else {

          console.log("The artist column was not updated, so nothing was moved!");
          // context.sync();
          return;
        }
        // context.sync();

      }).catch(function (error) {
        console.log("Promise Rejected");
      });
    //#endregion ------------------------------------------------------------------------------------------------
    };
  });
}
//#endregion ----------------------------------------------------------------------------------------------------

//#region SORT BY DATE ------------------------------------------------------------------------------------------
async function sortDate(eventArgs: Excel.TableChangedEventArgs) { //This function will be using event arguments to collect data from the workbook
  // console.log("SORT FUNCTION FIRED!");
  // console.log(eventArgs);

  var theChange = eventArgs.changeType; //Kind of change that was made
  var theDetails = eventArgs.details;

  // console.log("args ");

  
  if (theChange == "RangeEdited" && (theDetails == undefined || theDetails.valueTypeAfter == "String")) { //&& theDetails == undefined) {
    console.log("The sorting event has been initiated!!"); //Prevents an event from being triggered when a new row is inserted into the other sheet, thus causing duplicate runs

    //#region SORTING VARIABLES ---------------------------------------------------------------------------------
    Excel.run(async context => {
      var changedTable = context.workbook.tables.getItem(eventArgs.tableId); //Returns tableId of the table where the event occured
      var tableRange = changedTable.getRange(); //Gets the range of the changed table
      var sortHeader = tableRange.find(sortColumn, {}); //Gets the range of the entire sortColumn (the "Date" column) from the changed table
      sortHeader.load("columnIndex");
      sortHeader.load("addressLocal")
      // var sortTag = ["Urgent", "Semi-Urgent", "Not Urgent", "Eventual", "Downtime"];
      // const list = [
      //   { Tag: 'Urgent'},
      //   { Tag: 'Semi-Urgent'},
      //   { Tag: 'Not Urgent'},
      //   { Tag: 'Eventual'},
      //   { Tag: 'Downtime'},
      // ]
      //#endregion --------------------------------------------------------------------------------------------------

      //#region SORTING CONDITIONS --------------------------------------------------------------------------------
      return context.sync().then(function() {
        console.log("Sync completed...Ready to sort")
        // console.log(sortHeader.addressLocal);
        // console.log(list);

        // if (sortHeader.columnIndex == 14) {
        //   list.sort((a, b) => (a.Tag < b.Tag) ? 1 : -1);
        //   console.log(list);
        // }

        tableRange.sort.apply(
          [
            { //list of conditions to sort on
              key: sortHeader.columnIndex, //sorts based on data in Date column
              sortOn: Excel.SortOn.value, //sorts based on cell vlaues
              ascending: true
              // subField: Excel.subField, //sorts based on cell vlaues
              // subField: String(sortTag)
            }
          ],
          false, //will not impact string ordering
          true, //table has headers
          Excel.SortOrientation.rows //sorts the rows based on previous conditions
        );

        // const myArray = [1, 2, 3, 4, 5, 6];
        // let filteredArray = list.filter((x) => {
        //   return x % 2 === 0;
        // });
        

     

        // Queue a command to apply a filter on the Category column
        // var filter = changedTable.columns.getItem("Tags").filter;
        // filter.apply({
        //     filterOn: Excel.FilterOn.values,
        //     values: ["Urgent", "Semi-Urgent", "Not Urgent", "Eventual", "Downtime"]
        // });



        console.log("Sorting is completed.")
      }); 
      //#endregion --------------------------------------------------------------------------------------------------
    }).catch(tryCatch); // CATCH EXCEL.RUN
  
  }; // END IF  
} // END SORTDATE()
//#endregion ----------------------------------------------------------------------------------------------------

//#region TRY CATCH ---------------------------------------------------------------------------------------------
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    console.error(error);
  }
}
//#endregion ---------------------------------------------------------------------------------------------------




//#region AUTOFILL FUNCTIONS -------------------------------------------------------------------------------------

  //#region PROJECT TYPE HOURS -----------------------------------------------------------------------------------
  /**
   * Finds the value of Project Type in the changed row and returns a number of hours depending on the project type
   * @param {Array} rowValues loads the values of the changed row
   * @param {Number} changedRow loads the row number of the changed row
   * @returns A Number
   */   
  function lookupStart(rowValues, changedRow) { //loads these variables from another function to use in this function
    var address = "H" + (changedRow + 2); //takes the row that was updated and locates the address from the Project Type column.
    // console.log("The address of the new Project Type is " + address);
    var input = rowValues[0][7]; //assigns input the cell value in the changed row and the Project Type column (a nested array of values)
    // console.log(input);

    var a = ["Brand New Build", "Special Request"];
    var b = ["Brand New Build from Other Product Natives", "Brand New Build From Template", "Changes to Exisiting Natives", "Specification Check", "WeTransfer Upload to MS"];
    var output;

    if (a.includes(input)) { //if value in column H includes any input from var a...
      output = 4; //adds 4 hours
    } else if(b.includes(input)) { //if value in column H includes any input from var b...
      output = 2; //adds 2 hours
    } else { //everything else...
      output = 24; //adds 24 hours
    }
    return output;
  };
  //#endregion ---------------------------------------------------------------------------------------------------


  //#region PICKED UP / STARTED BY -------------------------------------------------------------------------------

    //References the Project Type column (H), Added column (J), and the Start Override column (U) to return a specific date and time for the project to by picked up by. This value is returned in the Picked Up / Started By column (M).

    //#region MY DATE ----------------------------------------------------------------------------------------------
    /**
     * Finds the value of Date Added in the changed row and converts it to be a date object in EST.
     * @param rowValues loads the values of the changed row
     * @param changedRow loads the row number of the changed row
     * @returns Date
     */
    function receivedAdjust(rowValues, changedRow) {
      var address = "J" + (changedRow + 2); //takes the row that was updated and locates the address from the Added column.
      // console.log("The address of the new Product is " + address);
      var dateTime = rowValues[0][9]; //assigns input the cell value in the changed row and the Added column (a nested array of values)

      //the below code basically is just converting the serial number in dateTime to a date object, and then adjusting to read in EST.
      var date = new Date(Math.round((dateTime - 25569)*86400*1000)); //convert serial number to date object
      date.setHours(date.getHours() + 4); //adjusting from GMT to EST (adds 4 hours)
      // console.log(`Date() ::  Convert Excel serial to Date():
      // ${date}`)
      return date;
    };
    //#endregion ---------------------------------------------------------------------------------------------------

    //#region OVERRIDE ---------------------------------------------------------------------------------------------
    /**
     * Finds the value of Start Override in the changed row and adds it to projectTypeHours, then adds that new number as hours to myDate. Adjusts for office hours and weekends.
     * @param {Array} rowValues loads the values of the changed row
     * @param {Number} projectTypeHours lookupStart returned number
     * @param {Date} myDate receivedAdjust returned date
     * @return {Date}
     */

    function startPreAdjust(rowValues, projectTypeHours, myDate) {
      var startOverride = rowValues[0][20]; //gets values of Start Orverride cell
      var startManualOverride = projectTypeHours + startOverride; //adds start override value to the number of hours for the project type
      var myDateCopy = new Date(myDate); //sets myDateCopy to myDate as a new date variable (so the old date doesnt get changed)
      var datePreHoursAdjust = new Date(myDateCopy);
      datePreHoursAdjust.setHours(datePreHoursAdjust.getHours() + startManualOverride);; //adds startManualOverride hours to myDate
      // console.log(datePreHoursAdjust);
      var adjustedDateTime = officeHours(datePreHoursAdjust, startManualOverride, myDateCopy); //converts to be within office hours if it already isn't
      var dateWeekendAdjusted = weekendAdjust(adjustedDateTime); //converts to be a weekday if it already isn't
      return dateWeekendAdjusted;
    }
    /*
    function startPreAdjust(rowValues, projectTypeHours, myDate) {
      var startOverride = rowValues[0][20]; //gets values of Start Orverride cell
      var snail = projectTypeHours + startOverride; //adds start override value to the number of hours for the project type
      var snailFail = new Date(myDate); //sets snailFail to myDate as a new date variable (so the old date doesnt get changed)
      snailFail.setHours(snailFail.getHours() + snail);; //adds snail hours to myDate
      // console.log(snailFail);
      var snailMail = officeHoursStart(snailFail); //converts to be within office hours if it already isn't
      var snailFlail = weekendAdjust(snailMail); //converts to be a weekday if it already isn't
      return snailFlail;
    }
    */
    //#endregion ----------------------------------------------------------------------------------------------------

    //#region STARTED PICKED UP BY ---------------------------------------------------------------------------------
    /**
     * Prints the value of override to the Picked Up / Started By column and formats the date in a readible format
     * @param {Number} changedRow loads the row number of the changed row
     * @param {Object} sheet the active worksheet
     * @param {Date} weekendHoursAdjust date adjusted to not land on a weekend
     * @returns date
     */
    function startedBy(changedRow, sheet, override) { //loads these variables from another function to use in this function
      var address = "M" + (changedRow + 2); //takes the row that was updated and locates the address from the Picked Up / Started By column.
      var range = sheet.getRange(address); //assigns the cell from the address variable to range
      // console.log(range);

      var formatDate = override.toLocaleDateString("en-us", { //formats the date to display correctly
          weekday:'short',
          month:'numeric',
          day: 'numeric',
          year: '2-digit'
      });

      var formatTime = override.toLocaleTimeString("en-us", { //formats the time to display correctly
        hour: '2-digit',
        minute:'2-digit'
      });

      var squeekday = formatDate + " " + formatTime; //adds the correctly displayed date and time together

      range.values = [[squeekday]]; //assigns the returned date value to the cell

      return range.values; //commits changes and exits the function
    };
    //#endregion ----------------------------------------------------------------------------------------------------

  //#endregion ------------------------------------------------------------------------------------------------------


  //#region PROOF TO CLIENT --------------------------------------------------------------------------------------

    //References the Project Type column (H), Product column (G), and the Work Override column (V) to return a specific date and time for a proof to be sent to the client. This value is returned in the Proof to Client column (N).

    //#region PRODUCT HOURS ----------------------------------------------------------------------------------------
    /**
     * Finds the value of Product in the changed row, returns a number of hours based on the product, and adds this number to projectTypeHours
     * @param {Array} rowValues loads the values of the changed row
     * @param {Number} projectTypeHours lookupStart returned number
     * @returns A Number
     */
    function preLookupWork(rowValues, projectTypeHours) {
      var input = rowValues[0][6]; //assigns input the cell value in the changed row and the Product column (a nested array of values)
      // console.log(input);
      var a = ["Menu", "Brochure", "Coupon Booklet", "Jumbo Postcard"];
      var b = ["MenuXL", "BrochureXL", "Folded Magnet", "Colossal Postcard", "Large Plastic"];
      var c = ["Small Menu", "Small Brochure", "Flyer", "Letter", "Envelope Mailer", "Postcard", "Magnet", "Door Hanger", "New Mover", "Birthday??", "Logo Creation"];
      var d = ["2SBT", "Box Topper", "Logo Recreation", "Business Cards"];
      var e = ["Scratch-Off Postcard", "Peel-A-Box Postcard", "Small Plastic", "Plastic New Mover", "Wide Format", "Artwork Only"];
      var f = ["Medium Plastic"];
      var output;

      if (a.includes(input)) { //if value in column G includes any input from var a...
        output = 10; //adds 10 hours
      } else if (b.includes(input)) { //if value in column G includes any input from var b...
        output = 18; //adds 18 hours
      } else if (c.includes(input)) { //if value in column G includes any input from var c...
        output = 4; //adds 4 hours
      } else if (d.includes(input)) { //if value in column G includes any input from var d...
        output = 2; //adds 2 hours
      } else if (e.includes(input)) { //if value in column G includes any input from var e...
        output = 7; //adds 7 hours
      } else if (f.includes(input)) { //if value in column G includes any input from var f...
        output = 15; //adds 15 hours
      } else { //everything else...
        output = 96; //adds 96 hours
      } //console.log(output);
      var newOutput = output + projectTypeHours; //adds hours from lookupStart to output and assigns new output to global variable
      // console.log(newOutput);
      return newOutput;
    };
    //#endregion --------------------------------------------------------------------------------------------------

    //#region WORK HOURS ADJUST ------------------------------------------------------------------------------------
    /**
     * if lookupStart number is 2, divide the preLookupWork number by 3. Otherwise, returns preLookupWork number
     * @param {Number} projectTypeHours lookupStart returned number
     * @param {Number} productHours preLookupWork returned number
     * @returns A Number
     */
    function lookupWork(projectTypeHours, productHours) {
      if(projectTypeHours == 2) { //if lookupStart number was 2...
        return (productHours / 3) //returns the productHours number divided by 3
      }
      return productHours; //otherwise returns the productHours number unaltered
    }
    //#endregion ---------------------------------------------------------------------------------------------------

    //#region WORKOVERRIDE --------------------------------------------------------------------------------------------
    /**
     * Finds the value of Work Override in the changed row and adds it to workHoursAdjust, then adds that new number as hours to startedPickedUpBy. Formats to be within office hours and on a weekday if needed.
     * @param {Array} rowValues loads the values of the changed row
     * @param {Number} workHoursAdjust loads the values of lookupWork
     * @param {Date} startedPickedUpBy loads the date that the project should be picked up by
     * @returns Date
     */
    function workPrePreAdjust (rowValues, workHoursAdjust, override) {
      var workOverride = rowValues[0][21]; //gets values of Work Orverride cell
      // console.log(workOverride);
      var workManualAdjust = workHoursAdjust + workOverride; //adds start override value to the number of hours for the project type
      var overrideCopy = new Date(override); //sets overrideCopy to startedPickedUpBy as a new date variable (so the old date doesnt get changed)
      var datePreHoursAdjust = new Date(overrideCopy);
      datePreHoursAdjust.setHours(datePreHoursAdjust.getHours() + workManualAdjust); //adds workManualAdjust hours to startedPickedUpBy date

      var adjustedDateTime = officeHours(datePreHoursAdjust, workManualAdjust, overrideCopy);
      // How many hours between adjustedDateTime and the next day at 8:30am 

      console.log(adjustedDateTime)

      // if (adjustedDateTime > 15 && adjustedDateTime < 24) {
      //   // It's before 5:30PM & 8:30AM
      //   console.log(`workManualAdjustFake is between 5:30PM & 8:30AM`)
      // } else if (adjustedDateTime < 15) {
      //   // It's past 5:30
      //   console.log(`workManualAdjustFake is  past 5:30PM`)
      // }



      // var sharkBait = officeHoursWork(datePreHoursAdjust, workManualAdjust, override); //converts to be within office hours if it already isn't
      var dateWeekendAdjusted = weekendAdjust(adjustedDateTime); //converts to be a weekday if it already isn't
      return dateWeekendAdjusted;
    };
    //#endregion --------------------------------------------------------------------------------------------------

    //#region PROOF TO CLIENT ---------------------------------------------------------------------------------
    /**
     * Prints the value of workOverride to the Proof to Client column and formats the date in a readible format
     * @param {Number} changedRow loads the row number of the changed row
     * @param {Object} sheet the active worksheet
     * @param {Date} workOverride proof to client date found in the workPreAdjust function (after converted to be within office hours and on a weekday)
     * @returns date
     */
    function toClient(changedRow, sheet, workOverride) { //loads these variables from another function to use in this function
      var address = "N" + (changedRow + 2); //takes the row that was updated and locates the address from the Proof to Client column.
      var range = sheet.getRange(address); //assigns the cell from the address variable to range
      // console.log(range);

      var formatDate = workOverride.toLocaleDateString("en-us", { //formats the date to display correctly
          weekday:'short',
          month:'numeric',
          day: 'numeric',
          year: '2-digit'
      });

      var formatTime = workOverride.toLocaleTimeString("en-us", { //formats the time to display correctly
        hour: '2-digit',
        minute:'2-digit'
      });

      var squeekday = formatDate + " " + formatTime; //adds the correctly displayed date and time together

      range.values = [[squeekday]]; //assigns the returned date value to the cell

      return range.values; //commits changes and exits the function
    };
    //#endregion ----------------------------------------------------------------------------------------------------

  //#endregion ------------------------------------------------------------------------------------------------------



  //#region OFFICE HOURS START ---------------------------------------------------------------------------------------
  /**
   * Adjusts a date object so that it falls into office hours
   * @param {Date} date Date to be adjusted to the start of office hours
   * @returns Date
   */
  /*
  function officeHoursStart(date) {
    var h = date.getHours(); // 14
    var m = date.getMinutes(); // 17
    var dayOfWeek = date.getDay(); //get day of week from date - 1
    var remainderHour;
    var remainderMinutes;

      // Morning
      if (h < 8) { //if hours is before 8, set hours to 8 and minutes to 30.
        date.setHours(8);
        date.setMinutes(30);
      }
      if (h == 8 && m < 30) { //if hours is 8 and minutes is before 30, set minutes to 30.
        date.setMinutes(30);
      };
      // Evening
      if (dayOfWeek == 5) { //if the day of the week is Friday...
        if (h > 13 || (h == 13 && m > 30)) { //if hours is greater than 15 (1:00pm) OR if hours is 13 and minutes are greater than 30...
          remainderHour = h - 13; //gets the humber of hours over 1:00pm
          remainderMinutes = m - 30; //gets the number of minutes over 30 minutes (a negative number if it is under 30 minutes)
          // console.log("remainderHour = " + remainderHour);
          // console.log("remainderMinutes = " + remainderMinutes);
          date.setDate(date.getDate() + 1); //add 1 day to the date (pushing this into the weekend, thus triggering the weekend adjust)
          date.setHours(8 + remainderHour); //set hour to 8:00AM + remainderHours
          date.setMinutes(30 + remainderMinutes); //set minutes to 30 + remainderMinutes
          // console.log("officeHours adjustment date is " + date);
        };
      } else if (h > 17 || (h == 17 && m > 30)) { //if hours is greater than 17 (5:00pm) OR if hours is 17 and minutes are greater than 30...
        remainderHour = h - 17 //gets the numer of hours over 5:00pm
        remainderMinutes = m - 30 //gets the number of minutes over 30 minutes (a negative number if it is under 30 minutes)
        // console.log("remainderHour = " + remainderHour);
        // console.log("remainderMinutes = " + remainderMinutes);
        date.setDate(date.getDate() + 1); //add 1 day to the date
        date.setHours(8 + remainderHour); //set hour to 8:00AM + remainderHours
        date.setMinutes(30 + remainderMinutes); //set minutes to 30 + remainderMinutes
        // console.log("officeHours adjustment date is " + date);
      };
      // console.log(date);
      return date;
  };
  //#endregion --------------------------------------------------------------------------------------------------
  */

  // Gets the day of the week
  function dayOfWeek(d) { //finds the day of the week
    var day = d.getDay();
    return day;
  }

  //#region OFFICE HOURS ---------------------------------------------------------------------------------------
  /**
   * Adjusts a date object so that it falls into office hours
   * @param {Date} date Date to be adjusted to the start of office hours
   * @returns Date
   */
  function officeHours(date, hoursAdjust, originalDate) {
    var endHour;
    var endMinute;
    var startHour;
    var startMinute;

    // function dayOfWeek(d) { //finds the day of the week
    //   var day = d.getDay();
    //   return day;
    // }

    var h = date.getHours(); // 12
    var m = date.getMinutes(); // 30
    var originalH = originalDate.getHours(); //9
    var originalM = originalDate.getMinutes(); //30
    var remainderHour;
    var remainderMinutes;

    var aNum = 0;


    var withinOfficeHours = true;


    while (withinOfficeHours == true) {

      if (aNum > 0) {
        // Not first pass ↓
        originalDate = date;
      };

        var adjustedDayOfWeek = dayOfWeek(date);

        if (adjustedDayOfWeek == 5) { //if day of week is Friday, set end of office hours to 1:30PM, otherwise office hours end at 5:30PM
          endHour = 13;
          endMinute = 30;
        } else {
          endHour = 17;
          endMinute = 30;
        };

        var hoursToEnd = 24 - endHour; //number of hours between end of work day and end of actual day (12:00 [24] - 5:00 [17] on most days) //7
        var minutesToEnd = 0; //amount of minutes will be calculated later since this affects the hours in most cases

        if (endMinute > 0) { //if office hours end at anytime other than on the hour, do this...
          hoursToEnd = hoursToEnd - 1; //subracts 1 hour from hoursToEnd (since we are adding in the minutes and are essentially counting backwards at this point)
          minutesToEnd = minutesToEnd + endMinute; //adds endMinute [30] to minutesToEnd [0]
        }; //hoursToEnd = 6 & minutesToEnd = 30

        var hoursToNextDay = hoursToEnd + 8; //assuming start time everyday is at hour[8], adds the amount of time between office hours end and end of day to end of day and beginning of office hours
        var minutesToNextDay = minutesToEnd + 30; //assuming start time everyday is at minute[30], add the minutes to end to the start minutes [30]

        if (minutesToNextDay > 59) { //if minutes goes into the hours terittory, we need to covert the hours and minutes to make visual sense
          hoursToNextDay = hoursToNextDay + 1;
          minutesToNextDay = minutesToNextDay -60;
        } //hoursToNextDay = 15 & minutesToNextDay = 0

        var hoursToAdd = hoursToNextDay + hoursAdjust; //adds hoursToNextDay (time between end of office and beginning of office next day) to adjustment hours
        var minutesToAdd = minutesToNextDay; //retuns the minutesToNextDay variable (if we anticipate the hoursAdjust ever returning minutes as well, we will add to this variable)
        date = new Date(originalDate);

        date.setHours(date.getHours() + hoursToAdd); //14 (2:00PM) + 31 (9:00PM next day)
        date.setMinutes(date.getMinutes() + minutesToAdd); //17 (2:17) + 0 = 17 (9:17PM next day)



        if (((date.getHours() < endHour) || (date.getHours() == endHour && date.getMinutes() < 30)) && ((date.getHours() > 8) || (date.getHours() == 8 && date.getMinutes() > 30))) {
            withinOfficeHours = false;
        }

        aNum++;


      /*
      var adjustedDate = new Date(originalDate);

      adjustedDate.setHours(adjustedDate.getHours() + hoursToAdd); //14 (2:00PM) + 31 (9:00PM next day)
      adjustedDate.setMinutes(adjustedDate.getMinutes() + minutesToAdd); //17 (2:17) + 0 = 17 (9:17PM next day)


      if (((adjustedDate.getHours() < endHour) || (adjustedDate.getHours() == endHour && adjustedDate.getMinutes() < 30)) && ((adjustedDate.getHours() > 8) || (adjustedDate.getHours() == 8 && adjustedDate.getMinutes() > 30)))
        {
          withinOfficeHours = false;
        } else 

      */

      
    };


    // var theRemainders = findingRemainders(h, m, endHour, endMinute, originalH, originalM); //finds the remainder from the end of day to the pick up time
    // // theRemainders = [remainderHour, remainderMinutes]


    // var newHoursAdjust = adj(theRemainders[0], theRemainders[1], hoursAdjust); //subtracts theRemainders from the hoursAdjust
    // // adjustedRemainder = [remainingHours, remainingMinutes];
    // var addToStartHours = hoursToNextDay + newHoursAdjust[0]; //15+12 = 27
    // var addToStartMinutes = 30 + newHoursAdjust[1]; //30+47 = 77

    // if (addToStartMinutes > 59) {
    //   addToStartHours + 1;
    //   addToStartMinutes - 60;
    // }; //startHours = 28 & startMinutes = 17

    // var adjustedDate = new Date(originalDate);

    // adjustedDate.setHours(originalDate.getHours() + addToStartHours); //14 (2:00) + 28 = 6:00PM next day
    // adjustedDate.setMinutes(originalDate.getMinutes() + addToStartMinutes); //17 (2:17) + 17 = 34 (6:34PM next day)



    // date.setDate(date.getDate() + 1); //add 1 day to the date
    // date.setHours(8 + ); //set hour to 8:00AM + remainderHours
    // date.setMinutes(30 + remainderMinutes); //set minutes to 30 + remainderMinutes
    // var newHours = date.getHours();
    // var newMinutes = date.getMinutes();

    // for (var i = 0; i < ?; i++) {
    // findingRemainders(newHours, newMinutes, endHour, endMinute, h, m);
    // remainders(remainderHour, remainderMinutes, hoursAdjust);
    // // afterHoursRepeat(newHours, newMinutes); //if the new date is still 
    // };



    // if (h < 8 || (h == 8 && m < 30)) { //if the adjusted time falls before office hours...

    // }




    //need to find the remainder from the end of day to the pick up time = var remainder
    //need to subtract the remainder time from the hoursAdjust = newHoursAdjust
    //need to find the amount of time between the end of day and the beginning of next day = var outOfOfficeTime
    //need to add add the newHoursAdjust time to the outOfOfficeTime = totalTimeToAdd
    //need to add totalTimeToAdd to pick up time






      // // Morning
      // if (h < 8) { //if hours is before 8, set hours to 8 and minutes to 30.
      //   date.setHours(8 + remainingHours);
      //   date.setMinutes(30 + remainingMinutes);
      // }
      // if (h == 8 && m < 30) { //if hours is 8 and minutes is before 30, set minutes to 30.
      //   date.setMinutes(30);
      // };
      // // Evening
      // if (dayOfWeek == 5) { //if the day of the week is Friday...
      //   if (h > 13 || (h == 13 && m > 30)) { //if hours is greater than 15 (1:00pm) OR if hours is 13 and minutes are greater than 30...
      //     remainderHour = h - 13; //gets the humber of hours over 1:00pm
      //     remainderMinutes = m - 30; //gets the number of minutes over 30 minutes (a negative number if it is under 30 minutes)
      //     console.log("remainderHour = " + remainderHour);
      //     console.log("remainderMinutes = " + remainderMinutes);
      //     date.setDate(date.getDate() + 1); //add 1 day to the date (pushing this into the weekend, thus triggering the weekend adjust)
      //     date.setHours(8 + remainderHour); //set hour to 8:00AM + remainderHours
      //     date.setMinutes(30 + remainderMinutes); //set minutes to 30 + remainderMinutes
      //     console.log("officeHours adjustment date is " + date);
      //   };
      // } else if (h > 17 || (h == 17 && m > 30)) { //if hours is greater than 17 (5:00pm) OR if hours is 17 and minutes are greater than 30...
      //   remainderHour = h - 17 //gets the numer of hours over 5:00pm
      //   remainderMinutes = m - 30 //gets the number of minutes over 30 minutes (a negative number if it is under 30 minutes)
      //   console.log("remainderHour = " + remainderHour);
      //   console.log("remainderMinutes = " + remainderMinutes);
      //   date.setDate(date.getDate() + 1); //add 1 day to the date
      //   date.setHours(8 + remainderHour); //set hour to 8:00AM + remainderHours
      //   date.setMinutes(30 + remainderMinutes); //set minutes to 30 + remainderMinutes
      //   console.log("officeHours adjustment date is " + date);
      // };
      // // console.log(date);
      // return date;
  };
//#endregion --------------------------------------------------------------------------------------------------

// function findingRemainders (newH, newM, endHour, endMinute, oldH, oldM) {

//   if (newH > endHour || (newH == endHour && newM > endMinute)) { //if the adjusted time falls after office hours...
//     //I need to find out how much time I need to add to the next day based on the amount of hours that were added in the workOverride and the amount of time that has already passed from the pick-up/start date to the end of office hours
//     var remainderHour = endHour - oldH; //gets the number of hours before 5:00pm or 1:00pm from the start by time
//     //remainderHour = 3
//     var remainderMinutes = endMinute - oldM; //gets the number of minutes under 30 minutes from the start by time (a negative number if it is above 30 minutes)
//     //remainderMinutes = 13

//     //So far, basically, we have the end of office hours time - the start by time (which will always be within office hours), giving us the amount of time that has already passed from the hoursAdjust number that we are adding to the start time of the next day
//   };

//   return [remainderHour, remainderMinutes];

// };

// function adj(rH, rM, hoursAdjust) { //Basically, this will return the value of hoursAdjust - remainder hours and minutes, giving us an adjusted hoursAdjust time that we will end up adding to the new start time to accuractly reflect a turn-around time within office hours
//   if (rM > 0) { //if the remainderMinutes is greater than 0 (and therefore has minutes)...
//     var remainingHours = hoursAdjust - rH; //subtracts the remainderHours from hoursAdjust (# of hours added to turn around time based on previously run functions)
//     //hoursAdjust = 16
//     //remainingHours = 16 - 3 = 13
//     remainingHours - 1; //subracts one, since minutes will be added to this number (remember, we are going backwards)
//     var remainingMinutes = 60 - rM; //subtracts 60 from the remainderMinutes (since hoursAdjust should always be a whole number...hopefully)
//     //remainingMinutes = 60 - 13 = 47
//   } else if (rM == 0) {
//     var remainingHours = hoursAdjust - rH; //subtracts the remainderHours from hoursAdjust (# of hours added to turn around time based on previously run functions)
//     //hoursAdjust = 16
//     //remainingHours = 16 - 3 = 13
//     var remainingMinutes = 0;
//   };
//   return [remainingHours, remainingMinutes];
// };


// function afterHoursRepeat(newHours, newMinutes) {
//   if (newHours > endHour || (newHours == endHour && newMinutes > endMinute)) { //if the adjusted time falls after office hours again...
//     remainderHour = newHours - endHour; //gets the number of hours after 5:00pm or 1:00pm from the start by time
//     remainderMinutes = newMinutes - endMinute; //gets the number of minutes under 30 minutes from the start by time (a negative number if it is above 30 minutes)
//     remainders(remainderHour, remainderMinutes);
//     date.setDate(date.getDate() + 1); //add 1 day to the date
//     date.setHours(8 + remainderHour); //set hour to 8:00AM + remainderHours
//     date.setMinutes(30 + remainderMinutes); //set minutes to 30 + remainderMinutes
//   };
// };

//#region WEEKEND HOURS ADJUST ---------------------------------------------------------------------------------
  /**
   * Finds the day of the week from override and if it is a weekend, changes it to be Monday at 8:30am
   * @param {Date} date Date to be adjusted to a weekday
   * @returns Date
   */

    function weekendAdjust(date) {
    var h = date.getHours(); // 4
    var m = date.getMinutes(); // 30
    var dayOfWeek = date.getDay(); //get day of week from date
    var adjustedDayOfWeek = dayOfWeek(date);

    if (dayOfWeek == 6) { //if weekday = Saturday
      var newDate = new Date(date);
      newDate.setDate(newDate.getDate() + 2);
      return newDate;
    }

    if (dayOfWeek == 0) { //if weekday = Sunday
      var newDate = new Date(date);
      newDate.setDate(newDate.getDate() + 1);
      return newDate;
    }
  }
/*
  function weekendAdjust(date) {
    var h = date.getHours(); // 4
    var m = date.getMinutes(); // 30
    var dayOfWeek = date.getDay(); //get day of week from date

    if (dayOfWeek == 5 && (h > 13 || h == 13 && m > 30)) { //if weekday = Friday and hours is greater than 13 OR if hours is equal to 13 and minutes is greater than 30, add three days and set time to 8:30am + remainders
      var newDate = new Date(date);
      newDate.setDate(newDate.getDate() + 3);
      newDate.setHours(8) //+ remainderHour);
      newDate.setMinutes(30) // + remainderMinutes);
      // console.log(newDate);
      return newDate;
    }

    if (dayOfWeek == 0) { //if weekday = Sunday, add one day and set time to 8:30am + remainders
      var newDate = new Date(date);
      newDate.setDate(newDate.getDate() + 1);
      newDate.setHours(8) //+ remainderHour);
      newDate.setMinutes(30) // + remainderMinutes);
      // console.log(newDate);
      return newDate;
    }

    if (dayOfWeek == 6) { //if weekday = Saturday, add 2 days and set time to 8:30am + remainders
      var newDate = new Date(date);

      newDate.setDate(newDate.getDate() + 2);
      newDate.setHours(8) // + remainderHour);
      newDate.setMinutes(30) // + remainderMinutes);
      // console.log(newDate);
      return newDate;
    } else { //if not a weekend, use original date
      return date;
    }
  }
  */
  //#endregion ------------------------------------------------------------------------------------------------------

//#endregion -----------------------------------------------------------------------------------------------------
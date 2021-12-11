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
var loop = true;

  //#region WEEKDAY VARIABLES ----------------------------------------------------------------------------------------------------------------

    var sunday = {
      dayID: 0,
      startHour: 8,
      startMinute: 30,
      endHour: 17,
      endMinute: 30,
      workDay: 0,
    }
    var monday = {
      dayID: 1,
      startHour: 8,
      startMinute: 0,
      endHour: 17,
      endMinute: 0,
      workDay: 0,
    }
    var tuesday = {
      dayID: 2,
      startHour: 8,
      startMinute: 30,
      endHour: 17,
      endMinute: 30,
      workDay: 0,
    }
    var wednesday = {
      dayID: 3,
      startHour: 8,
      startMinute: 30,
      endHour: 17,
      endMinute: 30,
      workDay: 0,
    }
    var thursday = {
      dayID: 4,
      startHour: 8,
      startMinute: 0,
      endHour: 18,
      endMinute: 0,
      workDay: 0,
    }
    var friday = {
      dayID: 5,
      startHour: 8,
      startMinute: 30,
      endHour: 13,
      endMinute: 30,
      workDay: 0,
    }
    var saturday = {
      dayID: 6,
      startHour: 8,
      startMinute: 30,
      endHour: 17,
      endMinute: 30,
      workDay: 0,
    }

    var weekdayList = [sunday, monday, tuesday, wednesday, thursday, friday, saturday];

  //#endregion --------------------------------------------------------------------------------------------------------------------------------

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
        
          var productHours = preLookupWork(rowValues, projectTypeHours); //adds hours based on Product and adds to lookupStart output
         
          var workHoursAdjust = lookupWork(projectTypeHours, productHours); //takes prelookupWork variable and divides by 3 if lookupStart was equal to 2. Otherwise remains the same.
     
          var myDate = receivedAdjust(rowValues, changedRow); //grabs values from Added column and converts into date object in EST.
        
          var override = startPreAdjust(rowValues, projectTypeHours, myDate); //adds manual override start hours to adjusted start time. Adjusts for office hours and weekends.
        
          var startedPickedUpBy = startedBy(changedRow, sheet, override); //Prints the value of override to the Picked Up / Started By column and formats the date in a readible format.
     
          var workOverride = workPrePreAdjust(rowValues, workHoursAdjust, override); //Finds the value of Work Override in the changed row and adds it to workHoursAdjust, then adds that new number as hours to startedPickedUpBy. Formats to be within office hours and on a weekday if needed.
       
          var proofToClient = toClient(changedRow, sheet, workOverride); //Prints the value of workOverride to the Proof to Client column and formats the date in a readible format.
        
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
      //var datePreHoursAdjust = new Date(myDateCopy);
      //datePreHoursAdjust.setHours(datePreHoursAdjust.getHours() + startManualOverride);; //adds startManualOverride hours to myDate
      // console.log(datePreHoursAdjust);
      var adjustedDateTime = officeHours(myDateCopy, startManualOverride); //converts to be within office hours if it already isn't
      //var dateWeekendAdjusted = weekendAdjust(adjustedDateTime); //converts to be a weekday if it already isn't
      return adjustedDateTime;
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
      var workManualAdjust = workHoursAdjust + workOverride; //adds start override value to the number of hours for the project type
      var overrideCopy = new Date(override); //sets overrideCopy to a new date variable (so the old date doesnt get changed)
      var adjustedDateTime = officeHours(overrideCopy, workManualAdjust);
      return adjustedDateTime;
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



  //#region OFFICE HOURS ---------------------------------------------------------------------------------------
    /**
     * Adjusts a date object so that it falls into office hours
     * @param {Date} date Date to be adjusted to the start of office hours
     * @returns Date
     */
    function officeHours(day, number) {

      //#region SETTING WORKDAY HOURS IN THE WEEKDAY VARIABLES -------------------------------------------------------------------------------------

        //loops through my weekday variables, finds returns the proper variable title for it's index in the array, and then runs it through the findWorkDay function
        for (var i = 0; i < weekdayList.length; i++) {
          var weekdayReplacement = findWorkDay(weekdayList[i]);
        };

      //#endregion --------------------------------------------------------------------------------------------------------------------------------

      //var aNum = 0

      while (loop == true) {
      var officeHours = withinOfficeHours(day, number);
      day = officeHours.date;
      number = officeHours.adjustmentNumber;
      loop = officeHours.loop;
      //aNum++
      };
      console.log("The correct date & time is: " + day);
      loop = true;
      return day;
    };

      //#region FUNCTIONS -------------------------------------------------------------------------------------------------------------------------


        //#region WITHIN OFFICE HOURS FUNCTION -------------------------------------------------------------------------------------------------

              function withinOfficeHours(date, adjustmentNumber) {

                //#region VARIABLES ------------------------------------------------------------------------------------------------------------

                    //#region SETS DATE VARIABLES ----------------------------------------------------------------------------------------------

                        //converts our input variables into milliseconds
                        var dateMilli = date.getTime();
                        var adjustmentNumberMilli = adjustmentNumber * 3600000;

                        //gets day of the week attributes for the date variable
                        var dateDayOfWeek = dayOfWeek(date); //returns a dayID (0-6) for the day of the week of the date object
                        var dayTitle = titleDOW(dateDayOfWeek); //returns a day title based on the dayID of the dateDayOfWeek variable

                        //retrives workday variables associated with the weekday of the date variable
                        var bookendVars = startEndMidnight(date, dayTitle);

                            //#region ADJUSTS DATES IN CASE REQUEST WAS SUBMITTED OUTSIDE OF OFFICE HOURS ---------------------------------------

                                if (date < bookendVars.startOfWorkDayMilli) { //if date is between 12AM and start time, adjust hours to be the start time
                                    date.setHours(dayTitle.startHour);
                                    date.setMinutes(dayTitle.startMinute);
                                    date.setSeconds(0);
                                    dateMilli = date.getTime();
                                    bookendVars = startEndMidnight(date, dayTitle);
                                };

                                if (date > bookendVars.endOfWorkDayMilli) { //if date is after end time and before 12AM, go to next day and adjust hours to be the start time of that next day
                                    date.setDate(date.getDate() + 1);
                                    dateDayOfWeek = dayOfWeek(date);
                                    dayTitle = titleDOW(dateDayOfWeek);
                                    date.setHours(dayTitle.startHour);
                                    date.setMinutes(dayTitle.startMinute);
                                    date.setSeconds(0);
                                    dateMilli = date.getTime();
                                    bookendVars = startEndMidnight(date, dayTitle);
                                };
                            
                            //#endregion ------------------------------------------------------------------------------------------------------------

                            //#region ADJUSTS DATES IN CASE REQUEST WAS SUBMITTED ON WEEKEND ----------------------------------------------------

                              if ((dateDayOfWeek == 6) || (dateDayOfWeek == 0)) { //if date was submitted on a weekend...
                                  date = weekendAdjust(date, dateDayOfWeek);
                                  dateDayOfWeek = dayOfWeek(date);
                                  dayTitle = titleDOW(dateDayOfWeek);
                                  date.setHours(dayTitle.startHour);
                                  date.setMinutes(dayTitle.startMinute);
                                  date.setSeconds(0);
                                  dateMilli = date.getTime();
                                  bookendVars = startEndMidnight(date, dayTitle);
                              };
                    
                        //#endregion ------------------------------------------------------------------------------------------------------------

                    //#endregion ----------------------------------------------------------------------------------------------------------------

                    //#region SETS ADJUSTMENT DATE VARIABLES -----------------------------------------------------------------------------------

                        //adds adjustmentNumber to date to get an adjustedDate value that will be used in later checks and calculations
                        var adjustedDate = new Date(date);
                        var adjustedDateMilli = adjustedDate.getTime();
                        adjustedDateMilli = adjustedDateMilli + adjustmentNumberMilli;
                        adjustedDate = new Date(adjustedDateMilli);

                    //#endregion ---------------------------------------------------------------------------------------------------------------

                    //#region SETS ADD A DAY VARIABLES -----------------------------------------------------------------------------------------

                        //gets day of the week attributes for the day after the date variable




                        /** --------------------------------------------------------
                         * .-. .-. .-. .-.   . . .-. .-. .-. .-. 
                         *  |  | | |  )|  )  |\| | |  |  |-  `-. 
                         *  '  `-' `-' `-'   ' ` `-'  '  `-' `-'
                         *  You are using way to many "nextDay"s here.
                         *  I fixed this error by creating a new variable
                         *  "newNextDay", and assigning that to the output of
                         *  getNextDay();
                         --------------------------------------------------------
                         ORIGINAL CODE:
                         -------------------------------------------------------- */
                        /*
                            var nextDay = new Date(date);
                            nextDay = getNextDay(nextDay); //also sets this variable to the start time of the next day
                            var addADay = nextDay.nextDay;
                            var addADayTitle = nextDay.nextDayTitle;
                            var addADayMilli = addADay.getTime();
                            
                            //retrives workday variables associated with the weekday of the addADay variable
                            var bookendAddedDate = startEndMidnight(addADay, addADayTitle);
                        */
                        /*  --------------------------------------------------------
                          FIXED CODE:
                          -------------------------------------------------------- */
                          var nextDay = new Date(date);

                          var newNextDay = getNextDay(nextDay); //also sets this variable to the start time of the next day
                          var addADay = newNextDay.nextDay;
                          var addADayTitle = newNextDay.nextDayTitle;
                          var addADayMilli = addADay.getTime();
                          
                          //retrives workday variables associated with the weekday of the addADay variable
                          var bookendAddedDate = startEndMidnight(addADay, addADayTitle);
                        /*  -------------------------------------------------------- */




                    //#endregion ----------------------------------------------------------------------------------------------------------------

                //#endregion --------------------------------------------------------------------------------------------------------------------

                //#region ACTION: SETS ADJUSTED DATE TO BE WITHIN OFFICE HOURS ------------------------------------------------------------------

                    //if adjustedDate falls outside of office hours, do this...
                    if (adjustedDateMilli < bookendVars.startOfWorkDayMilli || adjustedDateMilli > bookendVars.endOfWorkDayMilli) { //since the bookendVars is in reference to the date variable, this function will still trigger if adjustedDate is technically within office hours, but on a different day

                        //#region SETS ADJUSTMENT NUMBER VALUES ---------------------------------------------------------------------------------

                            var dayRemainder = (((bookendVars.endOfWorkDayMilli - dateMilli) / 1000) / 60) / 60; //time between end of work day and the original date time
                            var remainingAdjust = adjustmentNumber - dayRemainder; //gives us the remaining adjustment hours based off of what was already used to get to the end of the work day
                            var remainingAdjustMilli = remainingAdjust * 3600000;

                        //#endregion ------------------------------------------------------------------------------------------------------------

                        //#region NEW DAY CALCULATIONS ------------------------------------------------------------------------------------------

                            var newDay = new Date(addADay);

                            //adds remaining adjustment hours to the beginning of the work day the next day after date (addADay)
                            var dateTimeAdjusted = newDay.setMilliseconds((newDay.getMilliseconds() + remainingAdjustMilli));

                            var dateTimeAdjustedConvert = new Date(dateTimeAdjusted); //convert serial number to date object

                            date = dateTimeAdjustedConvert; //not sure if it should be date or something else yet. Need to make sure that the function works with this

                        //#endregion ------------------------------------------------------------------------------------------------------------

                        //#region SET LOOP VARIABLES IF STILL NOT WITHIN OFFICE HOURS OR EXCEEDS OFFICE HOURS OF NEXT DAY -----------------------

                            //if the new date exceeds the office hours of addADay, then do this...
                            if (dateTimeAdjusted > bookendAddedDate.endOfWorkDayMilli) {
                                adjustmentNumber = (remainingAdjust - addADayTitle.workDay) //subtracts remainingAdjust hours from the total workDay hours in the addADay variable
                                

                              /** --------------------------------------------------------
                              *  .-. .-. .-. .-.   . . .-. .-. .-. .-. 
                              *   |  | | |  )|  )  |\| | |  |  |-  `-. 
                              *   '  `-' `-' `-'   ' ` `-'  '  `-' `-'
                              *  Fixed by assigning the output of getNextDay to
                              *  a new variable "newDayAfterTomorrow".
                              * 
                              *  This line seems a little silly to me.
                              *  Reassigning an object to it's own property?
                              *     
                              *     dayAfterTomorrow = dayAfterTomorrow.nextDay;
                              *     
                              --------------------------------------------------------
                              ORIGINAL CODE:
                              -------------------------------------------------------- */
                              /*
                                var dayAfterTomorrow = new Date(addADay);
                                dayAfterTomorrow = getNextDay(dayAfterTomorrow);
                                dayAfterTomorrow = dayAfterTomorrow.nextDay;
                                date = new Date(dayAfterTomorrow);
                              */
                              /*  --------------------------------------------------------
                                FIXED CODE:
                                -------------------------------------------------------- */
                                var dayAfterTomorrow = new Date(addADay);
                                var newDayAfterTomorrow = getNextDay(dayAfterTomorrow);
                                // dayAfterTomorrow = dayAfterTomorrow.nextDay;
                                date = new Date(newDayAfterTomorrow.nextDay);
                              /* -------------------------------------------------------- */


                                loop = true;
                                return {
                                    date,
                                    adjustmentNumber,
                                    loop
                                };
                            } else {
                                loop = false;
                                return {
                                    date,
                                    //adjustmentNumber,
                                    loop
                                };
                            };

                        //#endregion -------------------------------------------------------------------------------------------------------------
                    
                    } else {
                        date = adjustedDate;
                        loop = false;
                        return {
                            date,
                            adjustmentNumber,
                            loop
                        };
                    };
                
                //#endregion --------------------------------------------------------------------------------------------------------------------

            };

          //#endregion ---------------------------------------------------------------------------------------------------------------------------


        //#region FIND WORK DAY FUNCTION -------------------------------------------------------------------------------------------------------

          function findWorkDay(weekday) {

              //sets start time for weekday variable to a date for calculations
              var start = new Date(0); //69, baby
              start.setHours(weekday.startHour);
              start.setMinutes(weekday.startMinute);
              start.setSeconds(0);

              //sets end time for weekday variable to a date for calculations
              var end = new Date(0); //seriously though, just making sure the dates for both variables will always be the same
              end.setHours(weekday.endHour);
              end.setMinutes(weekday.endMinute);
              end.setSeconds(0);

              
              /** --------------------------------------------------------
              *  .-. .-. .-. .-.   . . .-. .-. .-. .-. 
              *   |  | | |  )|  )  |\| | |  |  |-  `-. 
              *   '  `-' `-' `-'   ' ` `-'  '  `-' `-'
              * 
              *  Found insight here:
              *  https://stackoverflow.com/questions/36560806/the-left-hand-side-of-an-arithmetic-operation-must-be-of-type-any-number-or
              *  
              *  "You can fix by explicitly making the operands number (bigint) types so the - works."
              *  Fixed Example:
              *  new Date("2020-03-15T00:47:38.813Z").valueOf() - new Date("2020-03-15T00:47:24.676Z").valueOf()
              *  
              --------------------------------------------------------
              ORIGINAL CODE:
              -------------------------------------------------------- */
              /*
                  var workDayTime = (((end - start) / 1000) / 60) / 60; //subtracts end of day from start of day to get total work day hours for that weekday, then converts the milliseconds into hours (with decimal for minutes, if any)
              */
              /*  --------------------------------------------------------
                FIXED CODE:
                -------------------------------------------------------- */
                  var workDayTime = (((end.valueOf() - start.valueOf()) / 1000) / 60) / 60; //subtracts end of day from start of day to get total work day hours for that weekday, then converts the milliseconds into hours (with decimal for minutes, if any)
              /*  -------------------------------------------------------- */


              weekday.workDay = workDayTime; //sets our number to the variable 

              return weekday.workDay //returns our number to the actual object variable outside of the function

          };

        //#endregion ----------------------------------------------------------------------------------------------------------------------------


        //#region DAY OF WEEK FUNCTION ---------------------------------------------------------------------------------------------------------

          function dayOfWeek(d) { //finds the day of the week
          var day = d.getDay();
          return day;
          };

        //#endregion ----------------------------------------------------------------------------------------------------------------------------------


        //#region TITLE DAY OF WEEK FUNCTION ---------------------------------------------------------------------------------------------------

        /*  
        function titleDOW(d) { //returns the day of the week (refered to directly in another variable) based on the dayID index number
            if (d == 0) {
              var sunday = {
                dayID: 0,
                startHour: 8,
                startMinute: 30,
                endHour: 17,
                endMinute: 30,
                workDay: 0,
              }
              return sunday;
            } else if (d == 1) {
              var monday = {
                dayID: 1,
                startHour: 8,
                startMinute: 0,
                endHour: 17,
                endMinute: 0,
                workDay: 0,
              }
              return monday;
            } else if (d == 2) {
              var tuesday = {
                dayID: 2,
                startHour: 8,
                startMinute: 30,
                endHour: 17,
                endMinute: 30,
                workDay: 0,
              }
              return tuesday;
            } else if (d == 3) {
              var wednesday = {
                dayID: 3,
                startHour: 8,
                startMinute: 30,
                endHour: 17,
                endMinute: 30,
                workDay: 0,
              }
              return wednesday;
            } else if (d == 4) {
              var thursday = {
                dayID: 4,
                startHour: 8,
                startMinute: 0,
                endHour: 18,
                endMinute: 0,
                workDay: 0,
              }
              return thursday;
            } else if (d == 5) {
              var friday = {
                dayID: 5,
                startHour: 8,
                startMinute: 30,
                endHour: 13,
                endMinute: 30,
                workDay: 0,
              }
              return friday;
            } else if (d == 6) {
              var saturday = {
                dayID: 6,
                startHour: 8,
                startMinute: 30,
                endHour: 17,
                endMinute: 30,
                workDay: 0,
              }
              return saturday;
            };
          };
          */

          function titleDOW(d) { //returns the day of the week (refered to directly in another variable) based on the dayID index number
            if (d == 0) {
              return sunday;
            } else if (d == 1) {
              return monday;
            } else if (d == 2) {
              return tuesday;
            } else if (d == 3) {
              return wednesday;
            } else if (d == 4) {
              return thursday;
            } else if (d == 5) {
              return friday;
            } else if (d == 6) {
              return saturday;
            };
          };

        //#endregion ----------------------------------------------------------------------------------------------------------------------------------


        //#region START/END/MIDNIGHT FUNCTION --------------------------------------------------------------------------------------------------

          function startEndMidnight(originalDate, weekday) {

              var startOfWorkDay = new Date(originalDate); //adjusts start time of work day based on the day of the week
              startOfWorkDay.setHours(weekday.startHour);
              startOfWorkDay.setMinutes(weekday.startMinute);
              startOfWorkDay.setSeconds(0);
              var startOfWorkDayMilli = startOfWorkDay.getTime();

              var endOfWorkDay = new Date(originalDate); //adjusts end time of work day based on the day of the week
              endOfWorkDay.setHours(weekday.endHour);
              endOfWorkDay.setMinutes(weekday.endMinute);
              endOfWorkDay.setSeconds(0);
              var endOfWorkDayMilli = endOfWorkDay.getTime();

              var midnight = new Date(originalDate);
              midnight.setDate(midnight.getDate() + 1);
              midnight.setHours(0);
              midnight.setMinutes(0);
              midnight.setSeconds(0);
              var midnightMilli = midnight.getTime();

              return {
                  startOfWorkDay,
                  startOfWorkDayMilli,
                  endOfWorkDay,
                  endOfWorkDayMilli,
                  midnight,
                  midnightMilli
              };
          };

        //#endregion ----------------------------------------------------------------------------------------------------------------------------------


        //#region GET NEXT DAY FUNCTION --------------------------------------------------------------------------------------------------------

          function getNextDay(date) {

              /** --------------------------------------------------------
              *  .-. .-. .-. .-.   . . .-. .-. .-. .-. 
              *   |  | | |  )|  )  |\| | |  |  |-  `-. 
              *   '  `-' `-' `-'   ' ` `-'  '  `-' `-'
              * 
              *  Created a new variable "newNextDay" and assigned it to the
              *  output of nextDay.setDate(nextDay.getDate() + 1);
              *  Then made nextDay the new Date() from newNextDay
              * 
              * This code can be cleaned up I think. Lots of Date object
              * floating around and being coerced into other Date objects
              *  
              --------------------------------------------------------
              ORIGINAL CODE:
              -------------------------------------------------------- */
              /*
                var nextDay = new Date(date);
                nextDay = nextDay.setDate(nextDay.getDate() + 1); //returns the day after the original date
                nextDay = new Date(nextDay);
                var nextDayDayOfWeek = dayOfWeek(nextDay);
                var nextDayTitle = titleDOW(nextDayDayOfWeek); //returns a day title based on the dayID of the addADay variable              */
              /*  --------------------------------------------------------
                FIXED CODE:
                -------------------------------------------------------- */
                var nextDay = new Date(date);
                var newNextDay = nextDay.setDate(nextDay.getDate() + 1); //returns the day after the original date
                nextDay = new Date(newNextDay);
                var nextDayDayOfWeek = dayOfWeek(nextDay);
                var nextDayTitle = titleDOW(nextDayDayOfWeek); //returns a day title based on the dayID of the addADay variable
              /*  -------------------------------------------------------- */

              

              if ((nextDayDayOfWeek == 6) || (nextDayDayOfWeek == 0)) { //checks if nextDay falls on a weekend
                  nextDay = weekendAdjust(nextDay, nextDayDayOfWeek); //adjusts nextDay output to not fall on a weekend
                  nextDayDayOfWeek = dayOfWeek(nextDay);
                  nextDayTitle = titleDOW(nextDayDayOfWeek);
              };

              nextDay.setHours(nextDayTitle.startHour);
              nextDay.setMinutes(nextDayTitle.startMinute);
              nextDay.setSeconds(0);
              return {
                  nextDay,
                  nextDayTitle
              };
          };

        //#endregion ----------------------------------------------------------------------------------------------------------------------------------


        //#region WEEKEND ADJUST FUNCTION ------------------------------------------------------------------------------------------------------
          
function weekendAdjust(date, dateWeekday) {
  if (dateWeekday == 6) {
      var weekend = new Date(date);
      weekend.setDate(weekend.getDate() + 2);
      return weekend;
  } else if (dateWeekday == 0) {
      var weekend = new Date(date);
      weekend.setDate(weekend.getDate() + 1);
      return weekend;
  };
};

//#endregion ------------------------------------------------------------------------------------------------------------------------------


      //#endregion -------------------------------------------------------------------------------------------------------------------------------

  //#endregion ---------------------------------------------------------------------------------------------------







//OLD OFFICE HOURS REFERENCE!!!!!!!!!!!!!!!!!!!!!!!! _down _down 

//#region OFFICE HOURS ---------------------------------------------------------------------------------------
  /**
   * Adjusts a date object so that it falls into office hours
   * @param {Date} date Date to be adjusted to the start of office hours
   * @returns Date
   */

/*
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
    var startHour;
    var startMinutes;

    var aNum = 0;


    var withinOfficeHours = true;


    while (withinOfficeHours == true) {

      if (aNum > 0) {
        // Not first pass ↓
        originalDate = date;
      };

        var adjustedDayOfWeek = dayOfWeek(date);

        if (adjustedDayOfWeek == 5) { //if day of week is Friday, set office hours to 8:30 - 1:30
          startHour = 8;
          startMinutes = 30;
          endHour = 13;
          endMinute = 30;
        } else if (adjustedDayOfWeek == 1) { //if day of week is Monday, set office hours to 8:00 - 5:00
          startHour = 8;
          startMinutes = 0;
          endHour = 17;
          endMinute = 0;
        } else if (adjustedDayOfWeek == 4) { //if day of week is Thursday, set office hours to 8:00 - 6:00
          startHour = 8;
          startMinutes = 0;
          endHour = 18;
          endMinute = 0;
        } else { //for all other days of the week, set office hours to 8:30 - 5:30
          startHour = 8;
          startMinutes = 30;
          endHour = 17;
          endMinute = 30;
        }

        var monday = {
          dayID: 1,
          startHour: 8,
          startMinute: 0,
          endHour: 17,
          endMinute: 0,
          workDay: endHour - startHour
        }



        var h = date.getHours(); // 12
        var m = date.getMinutes(); // 30

        //if time of date falls in the evening, do this...
        if (h > endHour || h == endHour && m > endMinute) { 
          //calculates amount of time between end of office hours and end of day
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
        } 

        //if time of date falls in the morning, do this...
        else if (h < 8 || h == 8 && m < 30) {
          var hourToNextDay = 8 - h; //assuming start time everyday is at hour[8], subtracts the amount of time between beginning of office hours and the number of hours returned from the date variable
          var minuteToNextDay = m; //returns the minutes from the date variable
          startMinute = 30; //REMOVE ONCE START TIME VARIABLES ARE INPLAMENTED

          if (startMinute > 0) { //if time from date variable is at anytime other than on the hour, do this...
            hourToNextDay = hourToNextDay - 1; //6
            minuteToNextDay = 60 - minuteToNextDay; //60-17 = 43

            // minuteToNextDay = minuteToNextDay + 30 + m; //43+30+17=90
          }



        }

      



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

      
   // };

    //return date;



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
  //};

  /*
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
/*
    function weekendAdjust(date) {
    var dayOfWeek = date.getDay(); //get day of week from date

    if (dayOfWeek == 6) { //if weekday = Saturday
      var newDate = new Date(date);
      newDate.setDate(newDate.getDate() + 2);
      return newDate;
    } else if (dayOfWeek == 0) { //if weekday = Sunday
      var newDate = new Date(date);
      newDate.setDate(newDate.getDate() + 1);
      return newDate;
    } else {
      return date;
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
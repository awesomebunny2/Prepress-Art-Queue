/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

//#region IMAGE REFERENCES --------------------------------------------------------------------------------------
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";
//#endregion ----------------------------------------------------------------------------------------------------

//#region GLOBAL VARIABLES --------------------------------------------------------------------------------------
var artistColumn = "Q";
var moveEvent;
var eventsOff;
var eventsOn;
var sortEvent;
var sortColumn = "Date";
//#endregion ----------------------------------------------------------------------------------------------

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
  console.log("Ready!");
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
        // turnEventsOn();

        // console.log("Events are turned on!");
        
        // moveEvent = context.workbook.tables.onChanged.add(onTableChanged);

        // sortEvent = context.workbook.tables.onChanged.add(sortDate);

        return context.sync().then(function() { //Commits changes to document and then returns the console.log
          console.log("Event handlers have been successfully registered");
        });
      });
    };
});
//#endregion ------------------------------------------------------------------------------------------------

//#region TOGGLE EVENTS ----------------------------------------------------------------------------------------

  //#region EVENTS OFF -----------------------------------------------------------------------------------------
  async function turnEventsOff() {
    Excel.run(function (context) {
      context.runtime.load("enableEvents");
      return context.sync()
          .then(function () {
            eventsOff = !context.runtime.enableEvents;
            return;
          });
    });
  };
  //#endregion ----------------------------------------------------------------------------------------------

  //#region EVENTS ON ------------------------------------------------------------------------------------------
  async function turnEventsOn() {
    Excel.run(function turnEventsOn(context) {
      context.runtime.load("enableEvents");
      return context.sync()
          .then(function () {
            eventsOn = context.runtime.enableEvents;
            return;
          });
    });
  };
  //#endregion ------------------------------------------------------------------------------------------------

//#endregion ---------------------------------------------------------------------------------------------------

//#region MOVES DATA BETWEEN WORKSHEETS ------------------------------------------------------------------------
async function onTableChanged(eventArgs: Excel.TableChangedEventArgs) { //This function will be using event arguments to collect data from the workbook
  await Excel.run(async (context) => {
  
    var theChange = eventArgs.changeType; //Kind of change that was made
    // console.log("args ");
    // console.log(eventArgs);
    // if (theChange == "RowInserted" || eventArgs.details == undefined || (eventArgs.details.valueTypeBefore == "Empty" && eventArgs.details.valueTypeAfter == "Double")) {
    //   return; //Prevents an event from being triggered when a new row is inserted into the other sheet, thus causing duplicate runs
    // }
    if (theChange == "RangeEdited" && eventArgs.details !== undefined) {
    console.log("The move data event has been initiated!!");
      if (eventArgs.details.valueAfter == eventArgs.details.valueBefore) {
        console.log("No values have changed. Exiting move data event...")
        return;
      }

      //#region EVENT VARIABLES -----------------------------------------------------------------------------------
      var details = eventArgs.details; //Loads the values before and after the event
      var address = eventArgs.address; //Loads the cell's address where the event took place
      var changedTable = context.workbook.tables.getItem(eventArgs.tableId).load("name"); //Returns tableId of the table where the event occured
      // var newRange = changedTable.getDataBodyRange().load("address");
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

      //#region MOVE CONDITIONS -----------------------------------------------------------------------------------

      await context.sync() //Commits data to variables above

      .then(function moveData() {

        console.log(`
          changedColumn: ${changedColumn}
          artistColumn: ${artistColumn}
          details.valueAfter: ${details.valueAfter}
          myRow: ${myRow.values}
        `);

        

          // if (changedColumn == artistColumn && details.valueAfter == "Red Basket") { //If changed column = B & the updated value = "Red Basket", then...    
          //   redTable.rows.add(null, myRow.values); //Adds empty row to bottom of RedBasket Table, then inserts the changed values into this empty row
          //   myRow.delete(); //Deletes the changed row from the original sheet
          //   console.log("Data was moved to the Red Basket!");
          //   return;
          // } else if (changedColumn == artistColumn && details.valueAfter == "Green Basket") { //If changed column = B & the updated value = "Green Basket", then... 
          //   greenTable.rows.add(null, myRow.values); //Adds empty row to bottom of GreenBasket Table, then inserts the changed values into this empty row
          //   myRow.delete(); //Deletes the changed row from the original sheet
          //   console.log("Data was moved to the Green Basket!");
          //   return;
          // } else if (changedColumn == artistColumn && details.valueAfter == "Orange Basket") { //If changed column = B & the updated value = "Orange Basket", then... 
          //   orangeTable.rows.add(null, myRow.values); //Adds empty row to bottom of OrangeBasket Table, then inserts the changed values into this empty row
          //   myRow.delete(); //Deletes the changed row from the original sheet
          //   console.log("Data was moved to the Orange Basket!");
          //   return;
          // } else if (changedColumn == artistColumn && details.valueAfter == "Yellow Basket") { //If changed column = B & the updated value = "Yellow Basket", then... 
          //   yellowTable.rows.add(null, myRow.values); //Adds empty row to bottom of YellowBasket Table, then inserts the changed values into this empty row
          //   myRow.delete(); //Deletes the changed row from the original sheet
          //   console.log("Data was moved to the Yellow Basket!");
          //   return;
          // } else if (changedColumn == artistColumn && details.valueAfter == "Ground") { //If changed column = B & the updated value = "Ground", then... 
          //   fruitsTable.rows.add(null, myRow.values); //Adds empty row to bottom of Fruits Table, then inserts the changed values into this empty row
          //   myRow.delete(); //Deletes the changed row from the original sheet
          //   console.log("Data was moved to the Friuts!");
          //   return;
          //   } else {
          //     console.log("Looks like there wasn't a Basket change this time. No data was moved...")
          //   };
          });
    //#endregion ------------------------------------------------------------------------------------------------
    };
  });
};
//#endregion ----------------------------------------------------------------------------------------------------

//#region SORT BY DATE ------------------------------------------------------------------------------------------
async function sortDate(eventArgs: Excel.TableChangedEventArgs) { //This function will be using event arguments to collect data from the workbook

  var theChange = eventArgs.changeType; //Kind of change that was made
  var theDetails = eventArgs.details;

  // console.log("args ");
  // console.log(eventArgs);

  if (theChange == "RangeEdited" && (theDetails == undefined || theDetails.valueTypeAfter == "Double")) { //&& theDetails == undefined) {
    console.log("The sorting event has been initiated!!");

    //#region SORTING VARIABLES ---------------------------------------------------------------------------------
    Excel.run(async context => {
      var changedTable = context.workbook.tables.getItem(eventArgs.tableId); //Returns tableId of the table where the event occured
      var tableRange = changedTable.getRange(); //Gets the range of the changed table
      var sortHeader = tableRange.find(sortColumn, {}); //Gets the range of the entire sortColumn (the "Date" column) from the changed table
      sortHeader.load("columnIndex");
      //#endregion --------------------------------------------------------------------------------------------------

      //#region SORTING CONDITIONS --------------------------------------------------------------------------------
      return context.sync().then(function() {
        // console.log("Sync completed...Ready to sort")
        // console.log(sortHeader.columnIndex);
        tableRange.sort.apply(
          [
            { //list of conditions to sort on
              key: sortHeader.columnIndex, //sorts based on data in Date column
              sortOn: Excel.SortOn.value //sorts based on cell vlaues
            }
          ],
          false, //will not impact string ordering
          true, //table has headers
          Excel.SortOrientation.rows //sorts the rows based on previous conditions
        );
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


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
var x;
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
        
        moveEvent = context.workbook.tables.onChanged.add(onTableChanged);

        // sortEvent = context.workbook.tables.onChanged.add(sortDate);

        return context.sync().then(function() { //Commits changes to document and then returns the console.log
          console.log("Event handlers have been successfully registered");
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

      // console.log(eventArgs)
      
      console.log("The move data event has been initiated!!");
      
      if (eventArgs.details.valueAfter == eventArgs.details.valueBefore) {
        console.log("No values have changed. Exiting move data event...")
        return;
      };
    //#endregion ------------------------------------------------------------------------------------------------

    //#region MOVE CONDITIONS -----------------------------------------------------------------------------------
        
      await context.sync().then(function () {
        console.log("Promise Fulfilled!");

        var rowValues = myRow.values;

        if (changedColumn == projectTypeColumn || productColumn) { //if updated data was in Project Type column, run the lookupStart function
          lookupStart(rowValues, changedTable, changedRow); //inserts the new data as the function's input
          preLookupWork(rowValues, changedTable, changedRow, x);
          // prioritySort(rowValues, changedTable, changedRow, x);
        }

        // if (changedColumn == productColumn) { //if updated data was in Project Type column, run the lookupStart function
        //   preLookupWork(details.valueAfter); //inserts the new data as the function's input
        // } //NEED TO ADD RETURNED OUTPUT FROM LOOKUPSTART TO THE RETURNED VALUE OF THIS FUNCTION!!!!!!!!!!!----------------------


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
            console.log("Looks like there wasn't an Artist change this time. No data was moved...")
          } return context.sync();
        };
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
  console.log(eventArgs);

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
        console.log(sortHeader.addressLocal);
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

   
// function prioritySort(changedTable, changedRow, x) {
//   Excel.run(async context => { //Do while Excel is running

    function lookupStart(rowValues, changedTable, changedRow) {
      var address = "H" + (changedRow + 2);
      console.log("The address of the new Project Type is " + address);
      console.log(rowValues);
      console.log(rowValues[7]);
      // var input = changedTable.rows.getItemAt(address).load("values");
      var input = rowValues[7];
      console.log(input);


      // context.sync()

      console.log("The values of Project Type: " + input.values);

      var a = ["Brand New Build", "Special Request"];
      var b = ["Brand New Build from Other Product Natives", "Brand New Build From Template", "Changes to Exisiting Natives", "Specification Check", "WeTransfer Upload to MS"];
      var output;

      if (a.includes(input.values)) {
        output = 4;
      } else if(b.includes(input.values)) {
        output = 2;
      } else {
        output = 24;
      } console.log(output);
      x = output;
      return output;
    };

    function preLookupWork(rowValues, changedTable, changedRow, x) {
      var address = "G" + (changedRow + 2);
      console.log("The address of the new Product is " + address);
      var input = changedTable.rows.getItemAt(address).load("values");
      var a = ["Menu", "Brochure", "Coupon Booklet", "Jumbo Postcard"];
      var b = ["MenuXL", "BrochureXL", "Folded Magnet", "Colossal Postcard", "Large Plastic"];
      var c = ["Small Menu", "Small Brochure", "Flyer", "Letter", "Envelope Mailer", "Postcard", "Magnet", "Door Hanger", "New Mover", "Birthday??", "Logo Creation"];
      var d = ["2SBT", "Box Topper", "Logo Recreation", "Business Cards"];
      var e = ["Scratch-Off Postcard", "Peel-A-Box Postcard", "Small Plastic", "Plastic New Mover", "Wide Format", "Artwork Only"];
      var f = ["Medium Plastic"];
      var output;

      if (a.includes(input.values)) {
        output = 10;
      } else if (b.includes(input.values)) {
        output = 18;
      } else if (c.includes(input.values)) {
        output = 4;
      } else if (d.includes(input.values)) {
        output = 2;
      } else if (e.includes(input.values)) {
        output = 7;
      } else if (f.includes(input.values)) {
        output = 15;
      } else {
        output = 96;
      } console.log(output);
      return output;
    };
//   });
// };
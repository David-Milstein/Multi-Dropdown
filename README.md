/* 
** 
** This function will add the ability to have multiple values in a single cell while using the data validation drop-down.
** 
** First, set up data validation in the cells and enter the list of values (or point to the list of values) so that the drop-
** down displays. Make sure to set up the validation as Show Warning and NOT as Reject Output. There will be a red flag in
** the corner of the cell when there are multiple values, ignore the warning.
** 
** Save and then run the code to trigger the authorization pop-up, once authorized the code will be live (ignore the error)
** 
** Once live, select a value from the drop-down to add the first value, select a second value from the drop-down and the
** cell will refresh and return both values delimited with a ', ' (e.g. 'value1, value2'), add more values as needed.
** 
** Since the data validation is set to Show Warning, incorrect data can be added to the cell if typed, make sure to only
** enter values by using the drop-down, or start typing in the cell the first part of the value and select it from the
** drop-down, do NOT click enter to auto-fill the remainder of the value as done when using the Reject Output setting.
** 
** The Code is set to not allow the same value multiple times, it will ignore the value if it is selected twice.
** 
** The function will work even without setting up data validation, however, there will not be a drop-down.
** 
** As this is an onEdit function, actions cannot be undone (including clearing the cell), when the wrong value is entered, 
** the cell needs to be cleared and all the values need to be reentered.
**  
*/

function onEdit(e){

  //Declaring Variables

  var oldValue;
  var newValue;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activeCell = ss.getActiveCell();
  
  //This is an onEdit function, meaning the range of cells this applies to needs to be set
  //There are 4 versions of this function that apply to different range types

  //#1 - Apply to 1 cell (e.g. A1)
  //#2 - Apply to 1 column (e.g. A:A)
  //#3 - Apply to specific rows in a column (e.g. A2:A1000)
  //#4 - Apply to all rows in column after specific row (e.g A2:A)

  //To apply this function to additional rows, columns, or sheets add the || (OR) operator to the if statement, make sure to
  //put the current set of conditions and the new set of conditions into their own parantheses with the OR operator in between
  //For example: if ((Condition && Condition && Condition) || (Condition && Condition && Condition))
  //The data in the if conditions will need to be modified accordingly for the different values of the rows, columns, or sheets\
  //See the below example using the OR operator

  //The first part of the if statement below (testing the conditions) is what changes between versions
  //Here are the 4 versions, currently version #4 is being used, swap out the if statments as needed
  //Between versions #3 and #4 is an example using the OR operator using versions #3 and #4

  //#1 - Apply to a single cell (e.g. A1)

  /*
  if (activeCell.getColumn() == 4 //Enter the number value of the column (A = 1, B = 2, etc.)
  && activeCell.getRow() == 2 //Enter the row number
  && ss.getActiveSheet().getName() == 'Sheet1') { //Enter the name of the sheet (not the workbook)
  */

  //#2 - Apply to a single column column (e.g. A:A)

  /*
  if (activeCell.getColumn() == 4 //Enter the number value of the column (A = 1, B = 2, etc.)
  && ss.getActiveSheet().getName() == 'Sheet1') { //Enter the name of the sheet (not the workbook)
  */
  
  //#3 - Apply to specific rows in a column (e.g. A2:A1000)

  /*
  if (activeCell.getColumn() == 4 //Enter the number value of the column (A = 1, B = 2, etc.)
  && activeCell.getRow() >= 2 //Enter the first row number
  && activeCell.getRow() <= 1000 //Enter the last row number
  && ss.getActiveSheet().getName() == 'Sheet1') { //Enter the name of the sheet (not the workbook)
  */

  //Example using OR operator using versions #3 (first set) and version #4 (second set)

  /*
  if (
  //Enter the variables for the first set of conditions
  (activeCell.getColumn() == 4 //Enter the number value of the column (A = 1, B = 2, etc.)
  && activeCell.getRow() >= 2 //Enter the first row number
  && activeCell.getRow() <= 1000 //Enter the last row number
  && ss.getActiveSheet().getName() == 'Sheet1') //Enter the name of the sheet (not the workbook)
  || //the OR operator
  //Enter the variables for the second set of conditions
  (activeCell.getColumn() == 3 //Enter the number value of the column (A = 1, B = 2, etc.)
  && activeCell.getRow() >= 2 //Enter the first row number
  && ss.getActiveSheet().getName() == 'Sheet1') //Enter the name of the sheet (not the workbook)
  // Add additional OR operators and sets of conditions here
  ) {
  */

  //#4 - Apply to all rows in column after specific row (e.g A2:A)

  if (activeCell.getColumn() == 4 //Enter the number value of the column (A = 1, B = 2, etc.)
  && activeCell.getRow() >= 2 //Enter the first row number
  && ss.getActiveSheet().getName() == 'Sheet1') { //Enter the name of the sheet (not the workbook)
    
    //newValue is the value added to the cell with the edit
    newValue = e.value;
    //oldValue is the value that was in the cell before the edit
    oldValue = e.oldValue;

    //Converts the oldValue into an array of single values to check for duplicates
    var oldValueArray = oldValue.split(', ');
    var oldValueCount = oldValueArray.length;
    var oldMatch = 0;

    //Loops through individual values of oldValue to check for duplicates
    for (i = 0; i < oldValueCount; i++){

      //Checks if newValue is equal to any of the values in the oldValue array
      //If there is a duplicate, increase oldMatch value

      if (oldValueArray[i] == newValue) {
        oldMatch++;
      }
    }

    //If cell is blank, leave the cell blank

    if (!e.value) {
      activeCell.setValue("");
    }
    else {
      
      //If there was no value before, paste the new value to the cell

      if (!e.oldValue) {
        activeCell.setValue(newValue);
      }
      else {
        
        //If oldMatch is less than 1 it means there are no duplicates, append the new value to the string delimited with ', '

        if (oldMatch < 1){
          activeCell.setValue(oldValue+', '+newValue);
        }

        //If there is a duplicate, leave the oldValue as is

        else {
            activeCell.setValue(oldValue);
        }
      }
    }
  }
}

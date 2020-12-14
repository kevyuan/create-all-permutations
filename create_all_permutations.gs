/* 
 * This file is part of the kevyuan/create-all-permutations distribution 
 * (https://github.com/kevyuan/create-all-permutations).
 * Copyright (c) 2020 Kevin Yuan.
 * 
 * This program is free software: you can redistribute it and/or modify  
 * it under the terms of the GNU General Public License as published by  
 * the Free Software Foundation, version 3.
 *
 * This program is distributed in the hope that it will be useful, but 
 * WITHOUT ANY WARRANTY; without even the implied warranty of 
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU 
 * General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License 
 * along with this program. If not, see <http://www.gnu.org/licenses/>.
 */

function printPermutations() {
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var inputSheet = activeSpreadsheet.getSheetByName("INPUT")
  var lastColumn = inputSheet.getLastColumn()
  var firstRow = 1 // To-do: Handle column headers
  var lastRow = inputSheet.getLastColumn()
  var multiplier = 1  // Declare variable to temporary store the multiplier for a column before writing it back into the array
  var rangeToCopy = "" // Declare variable to use as a cursor on the INPUT sheet
  var numberOfPermutations = 1 // Declare variable to store number of Permutations; used for creating number of rows in OUTPUT sheet
  
  // What is the position of the last column? 
//  console.log("Last Column: " + lastColumn)
  
  // Set up an array to store number of elements from each column; array size is based on last column
  var numberOfElements = [lastColumn]

  // Populate the array with the number of elements from each column  
  for (var column = 0; column < lastColumn; column++) {
    
    // Fancy trick to find number of populated rows
    // https://stackoverflow.com/questions/17632165/determining-the-last-row-in-a-single-column
    numberOfElements[column] = inputSheet.getRange(firstRow, column+1, lastRow).getValues().filter(String).length
   
//    console.log("numberOfElements["+column+"]: " + numberOfElements[column])    
  }
  
  // Set up an array to store the multipliers for each column; array size is based on last column
  var multiplierPerColumn = [lastColumn]

  // Use second array to store the multiplier for each output column
  for (var column = 0; column < lastColumn; column++) {
    
    // Write the multiplier into the position in the array
    multiplierPerColumn[column] = multiplier
    
    // Calculate the next multiplier
    multiplier *= numberOfElements[lastColumn - column - 1]

//    console.log("multiplierPerColumn["+column+"]: " + multiplierPerColumn[column])   
    
    // Calculate the number of permutations
    numberOfPermutations *= numberOfElements[column]

  }

//  console.log("numberOfPermutations: " + numberOfPermutations)      
  
  // Reverse the multiplierPerColumn array so we can build the columns from left-to-right
  multiplierPerColumn = multiplierPerColumn.reverse()
//  console.log("Array multiplierPerColumn is reversed")   
  
  // Create a new tab with label "OUTPUT"; if it already exists, delete it and create new sheet
  var outputSheet = activeSpreadsheet.getSheetByName("OUTPUT")
  if (outputSheet != null) {
    activeSpreadsheet.deleteSheet(outputSheet)
  }
  outputSheet = activeSpreadsheet.insertSheet("OUTPUT")
  outputSheet.insertRows(1, numberOfPermutations)
  
  
  
  var startingRowNumber = 1 // Declare variable to manage the row where the pasting is to start
  var elementNumber = 1     // Declare variable to manage which element from the INPUT sheet is going to be copied from
  var currentColumn = 1     // Declare variable to manage the location of the current column
  var numberOfLoops = 1     // Declare variable to manage how many times the set of elements has to be pasted
  
  // Loop to move onto the next column
  for (var k = 0; k < lastColumn; k++) {

    console.log("Currently working on column: " + currentColumn) 
//    console.log("k: " + k)
    
    // Loop to copy values again, depending on the number of elements in the previous column
    for (var j = 0; j < numberOfLoops; j++) {
//      console.log("j: " + j)
      
      // Loop to copy values across the entire column
      for (var i = 0; i < numberOfElements[k]; i++) {
//        console.log("i: " + i)
      
        rangeToCopy = inputSheet.getRange(elementNumber, currentColumn)
//        console.log("Value at rangeToCopy: " + rangeToCopy.getValue())   
      
        rangeToCopy.copyValuesToRange(outputSheet, currentColumn, currentColumn, startingRowNumber, startingRowNumber+multiplierPerColumn[k]-1)
        startingRowNumber += multiplierPerColumn[k]
        elementNumber++
      }  
      elementNumber = 1        
    }
  
    currentColumn++
    startingRowNumber = 1
    numberOfLoops *= numberOfElements[k]
    
  }
  console.log("Permutations are completed!") 
}


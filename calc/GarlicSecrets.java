
// GarlicSecrets.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, September 2015

/* Open a spreadsheet and display it on-screen.

   (Freeze the top row.)
   Search for Garlic produce, and change each one's cost/pound value.
   Label the changes in red and bold.
   
   Calculate the sum of the totals column before and after the changes.

   Add a large label at the end of the spreadsheet -- black text in a red
   box created by merging several cells. Make its row invisible.

   Split the pane, and move the selection to the first row of the 
   top pane. Add the large label as a new first row.

   The basic idea for this example comes from the book 
   "Automate the Boring Stuff with Python" by Al Sweigart, 
   along with the test spreadsheet.
*/

import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.frame.*;

import com.sun.star.sheet.*;
import com.sun.star.table.*;

import com.sun.star.util.*;
import com.sun.star.container.*;


import com.sun.star.view.*;
import com.sun.star.awt.*;
import com.sun.star.beans.*;




public class GarlicSecrets
{
  private static final String FNM = "produceSales.xlsx";
  private static final String OUT_FNM = "garlicSecrets.ods";


  public static void main(String args[])
  {  
    XComponentLoader loader = Lo.loadOffice();
    XSpreadsheetDocument doc = Calc.openDoc(FNM, loader);
    if (doc == null) {
      System.out.println("Could not open " + FNM);
      Lo.closeOffice();
      return;
    }

    GUI.setVisible(doc, true);
    XSpreadsheet sheet = Calc.getSheet(doc, 0);
    Calc.gotoCell(doc, "A1");

    // freeze one row of view
    // Calc.freezeRows(doc, 1); 


    // find total for the "Total" column
    XCellRange totalRange = Calc.getColRange(sheet, 3);
    double total = Calc.computeFunction(GeneralFunction.SUM, totalRange);
    System.out.printf("Total before change: %.2f\n", total);

    increaseGarlicCost(doc, sheet);    // takes 15 secs

    // recalculate total
    total = Calc.computeFunction(GeneralFunction.SUM, totalRange);
    System.out.printf("Total after change:  %.2f\n", total);


    // add a label at the bottom of the data, and hide it
    int emptyRowNum = findEmptyRow(sheet);
    addGarlicLabel(doc, sheet, emptyRowNum);
    Lo.delay(1000);       // wait a bit before hiding last row
    XCellRange rowRange = Calc.getRowRange(sheet, emptyRowNum);
    Props.setProperty(rowRange, "IsVisible", false);  // property is in TableRow


    // split window into 2 view panes
    String cellName = Calc.getCellStr(0, emptyRowNum-2);
    System.out.println("Spliting at: " + cellName);
    Calc.splitWindow(doc, cellName);   // doesn't work with Calc.freeze() 


    // access panes; make top pane show first row
    XViewPane[] panes = Calc.getViewPanes(doc);
    System.out.println("No of panes: " + panes.length);
    panes[0].setFirstVisibleRow(0);

    // display view properties
    XSpreadsheetView ssView = Calc.getView(doc);
    Props.showObjProps("Spreadsheet view", ssView);

    // show view data 
    System.out.println("View data: " + Calc.getViewData(doc));

    // show sheet states
    ViewState[] states = Calc.getViewStates(doc);
    for(ViewState s : states)
      s.report();

    // make top pane the active one in the first sheet
    states[0].movePaneFocus(ViewState.MOVE_UP);
    Calc.setViewStates(doc, states);
    Calc.gotoCell(doc, "A1");    // move selection to top cell

    // show revised sheet states
    states = Calc.getViewStates(doc);
    for(ViewState s : states)
      s.report();


    // add a new first row, and label that as at the bottom
    Calc.insertRow(sheet, 0);
    addGarlicLabel(doc, sheet, 0);

    //XCellRange blanks = Calc.getCellRange(sheet, "A4999:B5001");
    //Calc.insertCells(sheet, blanks, true);  // shift right

    Lo.saveDoc(doc, OUT_FNM);

    Lo.waitEnter();
    Lo.closeDoc(doc);
    Lo.closeOffice();
  }  // end of main()



  private static int increaseGarlicCost(XSpreadsheetDocument doc, XSpreadsheet sheet)
  /* Iterate down the "Produce" column. If the text in the current cell is
     "Garlic" then change the corresponding "Cost Per Pound" cell by
     multiplying by 1.05, and changing its text to bold red.

     Return the "Produce" row index which is first empty.
  */
  {
    int row = 0;
    XCell prodCell = Calc.getCell(sheet, 0, row);   // produce column
    XCell costCell;
    while (prodCell.getType() !=  CellContentType.EMPTY) {
      // iterate down produce column until an empty cell is reached
      if (prodCell.getFormula().equals("Garlic")) {
        Calc.gotoCell(doc, Calc.getCellStr(0, row));   // show the cell in-screen
        costCell = Calc.getCell(sheet, 1, row);    // change cost/pound column
        costCell.setValue( 1.05* costCell.getValue());
        Props.setProperty(costCell, "CharWeight", com.sun.star.awt.FontWeight.BOLD);
        Props.setProperty(costCell, "CharColor", 0xFF0000);  // red
      }
      row++;
      prodCell = Calc.getCell(sheet, 0, row);   // get next cell
    }
    return row;
  }  // end of increaseGarlicCost()



  public static int findEmptyRow(XSpreadsheet sheet)
  /* Return the index of the first empty row by finding all the empty cell ranges in 
     the first column, and return the smallest row index in those ranges.
  */
  {
    // create a ranges query for the first column of the sheet
    XCellRange cellRange = Calc.getColRange(sheet, 0);
    Calc.printAddress(cellRange);
    XCellRangesQuery crQuery = Lo.qi(XCellRangesQuery.class, cellRange);

    XSheetCellRanges scRanges = crQuery.queryEmptyCells();
    CellRangeAddress[] addrs = scRanges.getRangeAddresses();
    Calc.printAddresses(addrs);

    // find smallest row index
    int row = -1;
    if (addrs != null) {
      row  = addrs[0].StartRow;
      for (int i = 1; i < addrs.length; i++) {
        if (row < addrs[i].StartRow)
          row = addrs[i].StartRow;
      }
      System.out.println("First empty row is at position: " + row);
    }
    else
      System.out.println("Could not find an empty row");
    return row;
  }  // end of findEmptyRow()



  private static void addGarlicLabel(XSpreadsheetDocument doc, 
                                             XSpreadsheet sheet, int emptyRowNum)
  /* Add a large text string ("Top Secret Garlic Changes") to the first cell
     in the empty row. Make the cell bigger by merging a few cells, and taller
     The text is black and bold in a red cell, and is centered.
  */
  {
    Calc.gotoCell(doc, Calc.getCellStr(0, emptyRowNum));

    // Merge first few cells of the last row
    XCellRange cellRange = Calc.getCellRange(sheet, 0, emptyRowNum,  
                                                   3, emptyRowNum);
    XMergeable xMerge = Lo.qi(XMergeable.class, cellRange);
    xMerge.merge(true);

    Calc.setRowHeight(sheet, emptyRowNum, 18);    // make the row taller
    XCell cell = Calc.getCell(sheet, 0, emptyRowNum);
    cell.setFormula("Top Secret Garlic Changes");
    Props.setProperty(cell, "CharWeight", com.sun.star.awt.FontWeight.BOLD);
    Props.setProperty(cell, "CharHeight", 24);
    Props.setProperty(cell, "CellBackColor", 0xFF0000);  // red
    Props.setProperty(cell, "HoriJustify", CellHoriJustify.CENTER);
    Props.setProperty(cell, "VertJustify", CellVertJustify.CENTER);
  }  // end of addGarlicLabel()

}  // end of GarlicSecrets class

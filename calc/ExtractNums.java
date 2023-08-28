
// ExtractNums.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, September 2015

/* Open a spreadsheet and display it on-screen.
   Extract all the numerical cell ranges as 2D arrays

   Usage:
      run ExtractNums totals.ods
      run ExtractNums sorted.csv
      run ExtractNums "small totals.ods"
*/

import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.frame.*;

import com.sun.star.sheet.*;
import com.sun.star.table.*;

import com.sun.star.util.*;


public class ExtractNums
{

  public static void main(String args[])
  {  
    String outFnm = null;
    if (args.length != 1) {
      System.out.println("Usage: run ExtractNums fnm");
      return;
    }
    
    XComponentLoader loader = Lo.loadOffice();
    XSpreadsheetDocument doc = Calc.openDoc(args[0], loader);
    if (doc == null) {
      System.out.println("Could not open " + args[0]);
      Lo.closeOffice();
      return;
    }
    GUI.setVisible(doc, true);

    XSpreadsheet sheet = Calc.getSheet(doc, 0);


    // basic data extraction
    // this code assumes the input file is "small totals.ods"
    System.out.println("\nA1 string: " + Calc.getVal(sheet, "A1"));   // string

    XCell cell = Calc.getCell(sheet, "A2");
    System.out.println("A2 type: " + Calc.getTypeString(cell));
    System.out.println("A2 value: " + Calc.getNum(sheet, "A2"));   // double

    cell = Calc.getCell(sheet, "E2");
    System.out.println("E2 type: " + Calc.getTypeString(cell));
    System.out.println("E2 value: " + Calc.getVal(sheet, "E2"));   // formula string
    System.out.println();

    Object[][] data = Calc.getArray(sheet, "A1:E10");
    Calc.printArray(data);

    double[][] ids = Calc.getDoublesArray(sheet, "A2:A7");
    Calc.printArray(ids);

    double[] projs = Calc.convertToDoubles( Calc.getCol(sheet, "B2:B7"));
    System.out.println("Project scores");
    for(double proj : projs)
      System.out.println("  " + proj);

    double[] stud = Calc.convertToDoubles( Calc.getRow(sheet, "A4:E4"));
    System.out.println("\nStudent scores");
    for(double v : stud)
      System.out.println("  " + v);


    // create a cell range that spans the used area of the sheet
    XCellRange usedCellRange = Calc.findUsedRange(sheet);
    System.out.println("\nThe used area is: " + Calc.getRangeStr(usedCellRange) + "\n");

    // find cell ranges that cover all the specified data types
    XCellRangesQuery crQuery = Lo.qi(XCellRangesQuery.class, usedCellRange);
    XSheetCellRanges cellRanges =  crQuery.queryContentCells(
                                       (short) CellFlags.VALUE);
                                    // (short) (CellFlags.VALUE | CellFlags.FORMULA));   
                                    // (short) CellFlags.STRING);

    // process each of the cell ranges
    //  -- extract each range as a 2D array of doubles
    if (cellRanges == null)
      System.out.println("No cell ranges found");
    else {
      System.out.println("Found cell ranges: " +
                      cellRanges.getRangeAddressesAsString() + "\n");
      CellRangeAddress[] addrs = cellRanges.getRangeAddresses();
      System.out.println("Cell ranges (" + addrs.length + "):");
      for(CellRangeAddress addr : addrs) {
        Calc.printAddress(addr);
        double[][] vals = Calc.getDoublesArray(sheet, Calc.getRangeStr(addr));
        Calc.printArray(vals);
      }
    }

    Lo.waitEnter();
    Lo.closeDoc(doc);
    Lo.closeOffice();
  }  // end of main()



}  // end of ExtractNums class

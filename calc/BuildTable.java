
// BuildTable.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, September 2015

/*  Based on the SCalc.java example in the Dev Guide
      * create two new cell styles
      * insert some data into the first sheet of a new spreadsheet
      * Print some examples of cell and cellrange names <--> position conversion.

      * apply the cell styles
      * display a chart
      * add a picture
      * split the view window
*/


import com.sun.star.lang.*;
import com.sun.star.frame.*;
import com.sun.star.beans.*;

import com.sun.star.sheet.*;
import com.sun.star.table.*;

import com.sun.star.style.*;
import com.sun.star.drawing.*;
import com.sun.star.chart2.*;   // instead of chart module




public class BuildTable
{
  private static String HEADER_STYLE_NAME = "My HeaderStyle";
  private static String DATA_STYLE_NAME = "My DataStyle";



  public static void main(String args[])
  {
    XComponentLoader loader = Lo.loadOffice();
    XSpreadsheetDocument doc = Calc.createDoc(loader);
    if (doc == null) {
      System.out.println("Document creation failed");
      Lo.closeOffice();
      return;
    }
    GUI.setVisible(doc, true);
    XSpreadsheet sheet = Calc.getSheet(doc, 0);

    convertAddresses(sheet);

    // buildCells(sheet);
    // buildRows(sheet);
    // buildCols(sheet);
    buildArray(sheet);

    // addPic(sheet, doc);

/*
    // add a chart
    CellRangeAddress rangeAddr =  Calc.getAddress(sheet, "A1:N4");
             // assumes buildArray() has filled the spreadsheet with data
    Chart2.insertChart(sheet, rangeAddr, "D6", 21, 11, "Column");
*/

    createStyles(doc);
    applyStyles(sheet);

    // Calc.splitSheet(doc, "E5");

    Lo.saveDoc(doc, "buildTable.ods");

    Lo.waitEnter();
    Lo.closeDoc(doc);
    Lo.closeOffice();
  }  // end of main()




  private static void buildCells(XSpreadsheet sheet)
  {
    Calc.setVal(sheet, 1, 0, "JAN");     // column -- row
    Calc.setVal(sheet, 2, 0, "FEB");
    Calc.setVal(sheet, 3, 0, "MAR");
    Calc.setVal(sheet, 4, 0, "APR");
    Calc.setVal(sheet, 5, 0, "MAY");
    Calc.setVal(sheet, 6, 0, "JUN");
    Calc.setVal(sheet, 7, 0, "JUL");
    Calc.setVal(sheet, 8, 0, "AUG");
    Calc.setVal(sheet, 9, 0, "SEP");
    Calc.setVal(sheet, 10, 0, "OCT");
    Calc.setVal(sheet, 11, 0, "NOV");
    Calc.setVal(sheet, 12, 0, "DEC");

    Calc.setVal(sheet, "B2", 31.45);    // name
    Calc.setVal(sheet, "C2", -20.9);
    Calc.setVal(sheet, "D2", -117.5);
    Calc.setVal(sheet, "E2", 23.4);
    Calc.setVal(sheet, "F2", -114.5);
    Calc.setVal(sheet, "G2", 115.3);
    Calc.setVal(sheet, "H2", -171.3);
    Calc.setVal(sheet, "I2", 89.5);
    Calc.setVal(sheet, "J2", 41.2);
    Calc.setVal(sheet, "K2", 71.3);
    Calc.setVal(sheet, "L2", 25.4);
    Calc.setVal(sheet, "M2", 38.5);
  }  // end of buildCells()


  private static void buildRows(XSpreadsheet sheet)
  {
    Calc.setRow(sheet, "B1", new Object[] {"JAN", "FEB", "MAR", "APR", "MAY", "JUN",
                                               "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"});
    Calc.setVal(sheet, "N1", "SUM");

    Calc.setVal(sheet, "A2", "Smith");
    Calc.setRow(sheet, "B2",  new Object[] {42, 58.9, -66.5, 43.4, 44.5, 45.3, 
                                           -67.3, 30.5, 23.2, -97.3, 22.4, 23.5});
    Calc.setVal(sheet, "N2", "=SUM(B2:M2)");

    Calc.setVal(sheet, 0, 2, "Jones");
    Calc.setRow(sheet, 1, 2,  new Object[] {21, 40.9, -57.5, -23.4, 34.5, 59.3, 
                                           27.3, -38.5, 43.2, 57.3, 25.4, 28.5});
    Calc.setVal(sheet, 13, 2, "=SUM(B3:M3)");

    Calc.setVal(sheet, 0, 3, "Brown");
    Calc.setRow(sheet, 1, 3,  new Object[] {31.45, -20.9, -117.5, 23.4, -114.5, 115.3, 
                                           -171.3, 89.5, 41.2, 71.3, 25.4, 38.5});
    Calc.setVal(sheet, 13, 3, "=SUM(A4:L4)");
  }  // end of buildRows()



  private static void buildCols(XSpreadsheet sheet)
  {
    Calc.setCol(sheet, "A2", new Object[] {"JAN", "FEB", "MAR", "APR", "MAY", "JUN",
                                               "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"});
    Calc.setVal(sheet, "A14", "SUM");

    Calc.setVal(sheet, "B1", "Smith");
    Calc.setCol(sheet, "B2",  new Object[] {42, 58.9, -66.5, 43.4, 44.5, 45.3, 
                                           -67.3, 30.5, 23.2, -97.3, 22.4, 23.5});
    Calc.setVal(sheet, "B14", "=SUM(B2:M2)");

    Calc.setVal(sheet, 2, 0, "Jones");
    Calc.setCol(sheet, 2, 1,  new Object[] {21, 40.9, -57.5, -23.4, 34.5, 59.3, 
                                           27.3, -38.5, 43.2, 57.3, 25.4, 28.5});
    Calc.setVal(sheet, 2, 13, "=SUM(B3:M3)");

    Calc.setVal(sheet, 3, 0, "Brown");
    Calc.setCol(sheet, 3, 1,  new Object[] {31.45, -20.9, -117.5, 23.4, -114.5, 115.3, 
                                           -171.3, 89.5, 41.2, 71.3, 25.4, 38.5});
    Calc.setVal(sheet, 3, 13, "=SUM(A4:L4)");
  }  // end of buildCols()




  private static void buildArray(XSpreadsheet sheet)
  {
    Object[][] vals = {
                 {"", "JAN", "FEB", "MAR", "APR", "MAY", "JUN",
                       "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"},
                 {"Smith", 42, 58.9, -66.5, 43.4, 44.5, 45.3, 
                          -67.3, 30.5, 23.2, -97.3, 22.4, 23.5},
                 {"Jones", 21, 40.9, -57.5, -23.4, 34.5, 59.3, 
                           27.3, -38.5, 43.2, 57.3, 25.4, 28.5},
                 {"Brown", 31.45, -20.9, -117.5, 23.4, -114.5, 115.3, 
                           -171.3, 89.5, 41.2, 71.3, 25.4, 38.5}
              };
    Calc.setArray(sheet, "A1:M4", vals);    // or just "A1"

    Calc.setVal(sheet, "N1", "SUM");
    Calc.setVal(sheet, "N2", "=SUM(B2:M2)");
    Calc.setVal(sheet, "N3", "=SUM(B3:M3)");
    Calc.setVal(sheet, "N4", "=SUM(A4:L4)");
  }  // end of buildArray()




  private static void convertAddresses(XSpreadsheet sheet)
  {
    // cell name <--> position
    java.awt.Point pos = Calc.getCellPosition("AA2");
    System.out.println("Position of AA2: (" + pos.x + ", " +  pos.y +  ")");

    XCell cell = Calc.getCell(sheet, pos.x, pos.y);
    Calc.printCellAddress(cell);

    System.out.println("AA2: " + Calc.getCellStr(pos.x, pos.y));
    System.out.println();

    // cell range name <--> position
    java.awt.Point[] range = Calc.getCellRangePositions("A1:D5");
    System.out.println("Range of A1:D5: (" + range[0].x + ", " +  range[0].y +  
                                  ") -- (" + range[1].x + ", " +  range[1].y +  ")");

    XCellRange cellRange = Calc.getCellRange(sheet, range[0].x, range[0].y,  
                                                    range[1].x, range[1].y);
    Calc.printAddress(cellRange);
    System.out.println("A1:D5: " + Calc.getRangeStr(range[0].x, range[0].y,  
                                                      range[1].x, range[1].y));
    System.out.println();
  }  // end of convertAddresses()



  public static void addPic(XSpreadsheet sheet, XSpreadsheetDocument doc)
  {
    // add a picture to the draw page for this sheet
    XDrawPageSupplier dpSupp = Lo.qi(XDrawPageSupplier.class, sheet);
    XDrawPage page = dpSupp.getDrawPage();
    Draw.drawImage(page, "skinner.png", 50, 40);

    // look at all the draw pages
    XDrawPagesSupplier supplier = Lo.qi(XDrawPagesSupplier.class, doc);
    XDrawPages pages = supplier.getDrawPages();
    System.out.println("1. No of draw pages: " + pages.getCount());

    XComponent compDoc = Lo.qi(XComponent.class, doc);
    System.out.println("2. No of draw pages: " + Draw.getSlidesCount(compDoc));
  }  // end of addPic()



  public static void createStyles(XSpreadsheetDocument doc)
  // create HEADER_STYLE_NAME and DATA_STYLE_NAME cell styles
  {
    try {
      XStyle style1 = Calc.createCellStyle(doc, HEADER_STYLE_NAME);

      XPropertySet props1 = Lo.qi(XPropertySet.class, style1);
      props1.setPropertyValue("IsCellBackgroundTransparent", false);
      props1.setPropertyValue("CellBackColor", 0x6699FF);   // dark blue
      props1.setPropertyValue("CharColor", 0xFFFFFF);       // white

      // Center cell content horizontally and vertically in the cell
      props1.setPropertyValue("HoriJustify", CellHoriJustify.CENTER); 
      props1.setPropertyValue("VertJustify", CellVertJustify.CENTER); 


      XStyle style2 = Calc.createCellStyle(doc, DATA_STYLE_NAME);

      XPropertySet props2 = Lo.qi(XPropertySet.class, style2);
      props2.setPropertyValue("IsCellBackgroundTransparent", false);
      props2.setPropertyValue("CellBackColor", 0xC2EBFF);  // light blue
      props2.setPropertyValue("ParaRightMargin", 500);  // move away from right edge
    }
    catch (com.sun.star.uno.Exception e) {
      System.out.println(e);
    }
  }  // end of createStyles()


  private static void applyStyles(XSpreadsheet sheet)
  {
    // apply cell styles
    //Calc.changeStyle(sheet, 1, 0, 13, 0, HEADER_STYLE_NAME);
    //Calc.changeStyle(sheet, 0, 1, 0, 3, HEADER_STYLE_NAME); 
    //Calc.changeStyle(sheet, 1, 1, 13, 3, DATA_STYLE_NAME); 
    Calc.changeStyle(sheet, "B1:N1", HEADER_STYLE_NAME);   
    Calc.changeStyle(sheet, "A2:A4", HEADER_STYLE_NAME);  
    Calc.changeStyle(sheet, "B2:N4", DATA_STYLE_NAME); 

    Calc.addBorder(sheet, "A4:N4", Calc.BOTTOM_BORDER, 0x6699FF);   // dark blue
    Calc.addBorder(sheet, "N1:N4", Calc.LEFT_BORDER | Calc.RIGHT_BORDER,
                                                          0x6699FF);   // dark blue
  }  // end of applyStyles()


}  // end of BuildTable class

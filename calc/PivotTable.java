
// PivotTable.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, September 2015

/*  Create pivot table (data pilot) which changes after 4 secs

   Based on code in Dev Guide's SpreadSheetSample.java example.
*/

import com.sun.star.lang.*;
import com.sun.star.frame.*;
import com.sun.star.beans.*;
import com.sun.star.container.*;

import com.sun.star.sheet.*;
import com.sun.star.table.*;



public class PivotTable
{

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
    createTable(sheet);
    createPivotTable(sheet);

    Lo.saveDoc(doc, "pivotExample.ods");
    Lo.waitEnter();
    Lo.closeDoc(doc);
    Lo.closeOffice();
  }  // end of main()



  private static void createTable(XSpreadsheet sheet)
  {
    Calc.highlightRange(sheet, "A1:C22", "Data Used by Pivot Table");
    Object[][] vals = {  { "Name", "Fruit", "Quantity" }, 
                { "Alice", "Apples", 3 },
                { "Alice", "Oranges", 7 }, { "Bob", "Apples", 3 },
                { "Alice", "Apples", 9 }, { "Bob", "Apples", 5 },
                { "Bob", "Oranges", 6 }, { "Alice", "Oranges", 3 },
                { "Alice", "Apples", 8 }, { "Alice", "Oranges", 1 },
                { "Bob", "Oranges", 2 }, { "Bob", "Oranges", 7 },
                { "Bob", "Apples", 1 }, { "Alice", "Apples", 8 },
                { "Alice", "Oranges", 8 }, { "Alice", "Apples", 7 },
                { "Bob", "Apples", 1 }, { "Bob", "Oranges", 9 },
                { "Bob", "Oranges", 3 }, { "Alice", "Oranges", 4 },
                { "Alice", "Apples", 9 }
              };
    Calc.setArray(sheet, "A2:C22", vals);    // or just "A1"
  }  // end of createTable()



  private static void createPivotTable(XSpreadsheet sheet)
  {
    Calc.highlightRange(sheet, "E1:H1", "Pivot Table");
    Calc.setColWidth(sheet, 4, 40);   // E column; in mm

    XDataPilotTables dpTables = Calc.getPilotTables(sheet);
    XDataPilotDescriptor dpDesc = dpTables.createDataPilotDescriptor();

    // set source range (use data range from CellRange test)
    CellRangeAddress srcAddr = Calc.getAddress(sheet, "A2:C22");
    dpDesc.setSourceRange(srcAddr);

    // settings for fields
    XIndexAccess fields = dpDesc.getDataPilotFields();

    // properties defined in DataPilotField
    // use first column as column field
    XPropertySet props = Lo.findContainerProps(fields, "Name");
    Props.setProperty(props, "Orientation", DataPilotFieldOrientation.COLUMN);

    // use second column as row field
    props = Lo.findContainerProps(fields, "Fruit");
    Props.setProperty(props, "Orientation", DataPilotFieldOrientation.ROW);

    // use third column as data field, calculating the sum
    props = Lo.findContainerProps(fields, "Quantity");
    Props.setProperty(props, "Orientation", DataPilotFieldOrientation.DATA);
    Props.setProperty(props, "Function", GeneralFunction.SUM);

    // select output position
    CellAddress destAddr = Calc.getCellAddress(sheet, "E3");
    dpTables.insertNewByName("DataPilotExample", destAddr, dpDesc);

/*
    Lo.delay(4000);
    System.out.println("Modifying Data field from sum to average...");
    try {
      dpDesc = Lo.qi(XDataPilotDescriptor.class, 
                             dpTables.getByName("DataPilotExample"));
      fields = dpDesc.getDataPilotFields();

      // change the data field to calculate the average of Quantity
      props = Lo.findContainerProps(fields, "Quantity");
      Props.setProperty(props, "Orientation", DataPilotFieldOrientation.DATA);
      Props.setProperty(props, "Function", GeneralFunction.AVERAGE);
    }
    catch(com.sun.star.uno.Exception e)
    {  System.out.println("Could not access pivot table fields: " + e);  }
*/
  }  // end of createPivotTable()



}  // end of PivotTable class


// PivotSheet1.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, September 2015

/*  Create new sheet containing pivot table in the document
    The data comes from "pivottable1.ods"
    Column field names: 
         Year	Period	Category	Revenue

    Usage:
      run PivotSheet1
*/

import com.sun.star.lang.*;
import com.sun.star.frame.*;
import com.sun.star.beans.*;
import com.sun.star.container.*;

import com.sun.star.sheet.*;
import com.sun.star.table.*;



public class PivotSheet1
{

  public static void main(String args[])
  {
    XComponentLoader loader = Lo.loadOffice();
    XSpreadsheetDocument doc = Calc.openDoc("pivottable1.ods", loader);
    if (doc == null) {
      System.out.println("Could not open pivottable1.ods");
      Lo.closeOffice();
      return;
    }
    GUI.setVisible(doc, true);

    XSpreadsheet sheet = Calc.getSheet(doc, 0);
    XSpreadsheet dpSheet = Calc.insertSheet(doc, "Pivot Table", (short)1);
    
    createPivotTable(sheet, dpSheet);
    Calc.setActiveSheet(doc, dpSheet);

    Lo.saveDoc(doc, "pivotExample1.ods");
    Lo.waitEnter();
    Lo.closeDoc(doc);
    Lo.closeOffice();
  }  // end of main()



  private static void createPivotTable(XSpreadsheet sheet, XSpreadsheet dpSheet)
  {
    XCellRange cellRange = Calc.findUsedRange(sheet);
    System.out.println("The used area is: " + Calc.getRangeStr(cellRange) + "\n");

    XDataPilotTables dpTables = Calc.getPilotTables(sheet);
    XDataPilotDescriptor dpDesc = dpTables.createDataPilotDescriptor();
    dpDesc.setSourceRange( Calc.getAddress(cellRange) );

    // XIndexAccess fields = dpDesc.getDataPilotFields();
    XIndexAccess fields = dpDesc.getHiddenFields();
    String[] fieldNames = Lo.getContainerNames(fields);
    System.out.println("Field Names (" + fieldNames.length + "):");
    for(String name : fieldNames)
      System.out.println("  " + name);

    // properties defined in DataPilotField
    XPropertySet props;
    // set page field
    //props = Lo.findContainerProps(fields, "Year");
    //Props.setProperty(props, "Orientation", DataPilotFieldOrientation.PAGE);

    // set column field
    props = Lo.findContainerProps(fields, "Category");
    Props.setProperty(props, "Orientation", DataPilotFieldOrientation.COLUMN);

    // set row field
    props = Lo.findContainerProps(fields, "Period");
    Props.setProperty(props, "Orientation", DataPilotFieldOrientation.ROW);


    // set data field, calculating the sum
    props = Lo.findContainerProps(fields, "Revenue");
    Props.setProperty(props, "Orientation", DataPilotFieldOrientation.DATA);
    Props.setProperty(props, "Function", GeneralFunction.SUM);

    // place onto sheet
    CellAddress destAddr = Calc.getCellAddress(dpSheet, "A1");
    dpTables.insertNewByName("PivotTableExample", destAddr, dpDesc);
    Calc.setColWidth(dpSheet, 0, 60);   // A column; in mm

/*
    Usually the table is not fully updated. The cells are often
    drawn with #VALUE! contents (?).

    This can be fixed by explicitly refreshing the table, but it has to
    be accessed via the sheet or the tables container is considered
    empty, and the table is not found.
*/
   // refreshTable(dpTables, "PivotTableExample");
          // no tables found; refresh cannot be done

    XDataPilotTables dpTables2 = Calc.getPilotTables(dpSheet);
    refreshTable(dpTables2, "PivotTableExample");
  }  // end of createPivotTable()




  private static void refreshTable(XDataPilotTables dpTables, String tableName)
  {
    String[] nms = dpTables.getElementNames();
    System.out.println("No. of DP tables: " + nms.length);
    for (String nm : nms)
      System.out.println("  " + nm);

    XDataPilotTable dpTable = Calc.getPilotTable(dpTables, tableName);
    if (dpTable != null)
      dpTable.refresh();
  }  // end of refreshTable()



}  // end of PivotSheet1 class

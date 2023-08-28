
// Validator.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, September 2015

/* Validata data entry so between 0 and 5.

   Based on code in Dev Guide's SpreadSheetSample.java example.
   and see the dev guide "Data Validation" section:
     https://wiki.openoffice.org/w/index.php?title=Documentation/
                           DevGuide/Spreadsheets/Other_Table_Operations
*/

import com.sun.star.lang.*;
import com.sun.star.frame.*;
import com.sun.star.beans.*;

import com.sun.star.sheet.*;
import com.sun.star.table.*;


public class Validator
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
    Calc.highlightRange(sheet, "A1:C3", "Validation Example");
    Calc.setVal(sheet, "A2", "Insert 3 values between 0.0 and 5.0:");

    XCellRange validRange = Calc.getCellRange(sheet, "A3:C3");
    validate(validRange, 0, 5);

    Lo.saveDoc(doc, "validator.ods");
    Lo.waitEnter();
    Lo.closeDoc(doc);
    Lo.closeOffice();
  }  // end of main()



  private static void validate(XCellRange cellRange, double start, double end)
  {
    try {
      // get cell range properties
      XPropertySet crProps = Lo.qi(XPropertySet.class, cellRange);

      // change validation properties of cell range
      XPropertySet vProps = Lo.qi(
             XPropertySet.class, crProps.getPropertyValue("Validation"));
                                 // "Validation" is defined inside SheetCellRange
      Props.showProps("Validation", vProps);
              // see "Data Validation" in the Dev Guide
              // these props are defined in the TableValidation service

      vProps.setPropertyValue("Type", ValidationType.DECIMAL);
      vProps.setPropertyValue("ShowErrorMessage", true);
      vProps.setPropertyValue("ErrorMessage", "This is an invalid value!");
      vProps.setPropertyValue("ErrorAlertStyle", ValidationAlertStyle.STOP);

      // set condition
      XSheetCondition cond = Lo.qi(XSheetCondition.class, vProps);
      cond.setOperator(ConditionOperator.BETWEEN);
      cond.setFormula1(""+start);
      cond.setFormula2(""+end);

      // store updated validation props in cell range
      crProps.setPropertyValue("Validation", vProps);
    }
    catch(com.sun.star.uno.Exception e)
    {  System.out.println(e);  }
  }  // end of validate()



}  // end of Validator class

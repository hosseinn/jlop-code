
// CopyPasteCalc.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, July 2016

/* Copy a row from the spreadsheet, and paste it
   back in further down the same sheet.

   Use Clip/JClip functions and "Copy"/"Paste" dispatches.

   "Copy" stores the data in multiple formats on the clipboard
    which are read bu the Clip.getXXX() methods.

   Based on GarlicSecrets.java in Calc Tests/
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




public class CopyPasteCalc
{
  private static final String FNM = "Addresses.ods";


  public static void main(String args[])
  {  
    XComponentLoader loader = Lo.loadOffice();
    XSpreadsheetDocument doc = Calc.openDoc(FNM, loader);
    if (doc == null) {
      System.out.println("Could not open " + FNM);
      Lo.closeOffice();
      return;
    }

    // useClipUtils(doc);
    useDispatches(doc);

    Lo.closeDoc(doc);
    Lo.closeOffice();
  }  // end of main()



  private static void useClipUtils(XSpreadsheetDocument doc)
  {
    XSpreadsheet sheet = Calc.getSheet(doc, 0);

    XCellRange cellRange = Calc.getCellRange(sheet, 0, 2, 6, 2); // row 3
    // Calc.printAddress(cellRange);
    Object[][] data = Calc.getCellRangeArray(cellRange);
    // Calc.printArray(data);
    JClip.setArray(data);   // copy array to clipboard

    Lo.wait(2000);

    XCellRange toCellRange = Calc.getCellRange(sheet, 0, 8, 6, 8);  // row 9
    Object[][] cbData = JClip.getArray();      // paste in array
    Calc.setCellRangeArray(toCellRange, cbData);

    GUI.setVisible(doc, true);   // so can see the paste
    Lo.waitEnter();
  }  // end of useClipUtils()




  private static void useDispatches(XSpreadsheetDocument doc)
  {
    GUI.setVisible(doc, true);   // Office *must* be visible
    XSpreadsheet sheet = Calc.getSheet(doc, 0);

    XCellRange cellRange = Calc.getCellRange(sheet, 0, 2, 6, 2); // row 3
    Calc.printAddress(cellRange);

    XSelectionSupplier selSupp = GUI.getSelectionSupplier(doc);
    selSupp.select(cellRange);
    Lo.dispatchCmd("Copy");

    Clip.listFlavors();

    // save embedded sheet as ODS; 
    FileIO.saveBytes("AddressesFrag.ods", Clip.getEmbedSource());

    // save sheet as SYLK file
    FileIO.saveString("AddressesFrag.slk", Clip.getSylk());


    XCellRange toCellRange = Calc.getCellRange(sheet, 0, 8, 6, 8);  // row 9
                 // GUI dialog asks about insertion if range is too small
                 // and data repeats if range is too big
    selSupp.select(toCellRange);
    Lo.dispatchCmd("Paste");

    Lo.waitEnter();
  }  // end of useDispatches()



}  // end of CopyPasteCalc class

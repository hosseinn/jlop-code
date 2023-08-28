
// MakeSheet.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, Jan 2017

/* Make a new spreadhsheet containing on sheet with some data.
*/

import org.odftoolkit.simple.*;
import org.odftoolkit.simple.table.*;


public class MakeSheet 
{
  public static void main(String[] args) 
  {
    try {
      SpreadsheetDocument doc = SpreadsheetDocument.newSpreadsheetDocument();
      Table sheet = doc.getSheetByIndex(0);
      sheet.getCellByPosition(0, 0).setStringValue("Hello");
      for (int row = 0; row < 5; row++)
        sheet.getCellByPosition(1, row).setDoubleValue(row*2.0);

      System.out.println("Saving document to makeSheet.ods");
      doc.save("makeSheet.ods");
    } 
    catch (Exception e) 
    {  System.out.println(e); }
  }  // end of main()

}  // end of MakeSheet class

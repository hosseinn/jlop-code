
// CombineSheets.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, Jan. 2017


/* Combine two spreadsheets together. 
   Save the resulting document in a new file.
*/

import org.odftoolkit.simple.*;
import org.odftoolkit.simple.table.*;


public class CombineSheets 
{
  public static void main(String[] args) 
  {
    try {
      SpreadsheetDocument doc1 = SpreadsheetDocument.loadDocument("ss1.ods");
      SpreadsheetDocument doc2 = SpreadsheetDocument.loadDocument("ss2.ods");
      int numSheets2 = doc2.getSheetCount();

      // add sheets of second document to end of first doc
      for(int i=0; i < numSheets2; i++) {
        Table t = doc2.getSheetByIndex(i);
        doc1.appendSheet(t, t.getTableName());
      }
      System.out.println("Saving combination to combined.ods");
      doc1.save("combined.ods");
      doc1.close();
      doc2.close();
    } 
    catch (Exception e) 
    {  System.out.println(e); }
  }  // end of main()

}  // end of CombineSheets class

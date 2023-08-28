
// ReplaceAll.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, September 2015

/* Search cells and replace all ocurrences of 
   the specified word by another randomly chosen word.

   Usage:
      run ReplaceAll dog
*/

import java.util.*;

// import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.frame.*;
import com.sun.star.beans.*;

import com.sun.star.sheet.*;
import com.sun.star.table.*;
import com.sun.star.text.*;
import com.sun.star.util.*;

import com.sun.star.uno.Exception;
// import com.sun.star.container.*;



public class ReplaceAll
{
  private static final String[] animals =
       { "ass", "cat", "cow", "cub", "doe", "dog", "elk", 
         "ewe", "fox", "gnu", "hog", "kid", "kit", "man",
         "orc", "pig", "pup", "ram", "rat", "roe", "sow", "yak" };


  public static void main(String args[])
  {
    if (args.length != 1) {
      System.out.println("Usage: run ReplaceAll <animal>");
      return;
    }


    XComponentLoader loader = Lo.loadOffice();
    XSpreadsheetDocument doc = Calc.createDoc(loader);
    if (doc == null) {
      System.out.println("Document creation failed");
      Lo.closeOffice();
      return;
    }
    GUI.setVisible(doc, true);

    XSpreadsheet sheet = Calc.getSheet(doc, 0);

    Random rand = new Random();
    for (int i=0; i < 6; i++)
      for (int j=0; j < 15; j++)
        Calc.setVal(sheet, i, j, animals[ rand.nextInt(animals.length)] );
                                // add text for the search

    XCellRange cellRange = Calc.getCellRange(sheet, 0, 0, 5, 14);  // A1:F15
                               // colStart, rowStart, colEnd, rowEnd

    searchIter(sheet, cellRange, args[0]);
    //searchAll(sheet, cellRange, args[0]);

    replaceAll(cellRange, args[0],
                          // "[a-z]+",
                          animals[ rand.nextInt(animals.length)] );

    Lo.saveDoc(doc, "replace.ods");
    Lo.waitEnter();
    Lo.closeDoc(doc);
    Lo.closeOffice();
  }  // end of main()




  private static void searchIter(XSpreadsheet sheet,
                                XCellRange cellRange, String srchStr)
  { 
    System.out.println("Searching for all occurrences of \"" + srchStr + "\"");
    try {
      XSearchable srch = Lo.qi(XSearchable.class, cellRange);
      XSearchDescriptor sd = srch.createSearchDescriptor();

      sd.setSearchString(srchStr);
      sd.setPropertyValue("SearchWords", true);
                           // only complete words will be found

      XCellRange cr = Lo.qi(XCellRange.class, srch.findFirst(sd));
      int count = 0;
      while (cr != null) {
        highlight(cr);
        System.out.println("  Match " + (count+1) + ": " +  
                                           Calc.getRangeStr(cr));
        cr = Lo.qi(XCellRange.class, srch.findNext(cr, sd));
              // occasionally XCell convert returns null ??
        count++;
      }
    }
    catch(Exception e)
    {  System.out.println(e);  }
  }  // end of searchIter()


  private static void highlight(XCellRange cr)
  {
    Props.setProperty(cr, "CharWeight", com.sun.star.awt.FontWeight.BOLD);
    Props.setProperty(cr, "CharColor", Calc.DARK_BLUE);
    Props.setProperty(cr, "CellBackColor", Calc.LIGHT_BLUE);
  }  // end of highlight()




  private static void searchAll(XSpreadsheet sheet,
                                XCellRange cellRange, String srchStr)
  { 
    System.out.println("Searching for all occurrences of \"" + srchStr + "\"");
    try {
      XSearchable srch = Lo.qi(XSearchable.class, cellRange);
      XSearchDescriptor sd = srch.createSearchDescriptor();

      sd.setSearchString(srchStr);
      sd.setPropertyValue("SearchWords", true);

      XCellRange[] matchCrs = Calc.findAll(srch, sd);
      if (matchCrs == null) 
        return;

      System.out.println("Search text found " + matchCrs.length + " times");
      for (int i=0; i < matchCrs.length; i++) {
        highlight(matchCrs[i]);
        System.out.println("  Index " + i + ": " +  
                                    Calc.getRangeStr(matchCrs[i]));
      }
    }
    catch(Exception e)
    {  System.out.println(e);  }
  }  // end of searchAll()



  private static void replaceAll(XCellRange cellRange,
                                         String srchStr, String replStr)
  { 
    System.out.println("Replacing \"" + srchStr + "\" with \"" + replStr + "\"");
    Lo.delay(2000);   // wait a bit before search & replace
    try {
      XReplaceable repl = Lo.qi(XReplaceable.class, cellRange);
      XReplaceDescriptor rd = repl.createReplaceDescriptor();

      rd.setSearchString(srchStr);
      rd.setReplaceString(replStr);
      rd.setPropertyValue("SearchWords", true);
      // rd.setPropertyValue("SearchRegularExpression", true);

      int count = repl.replaceAll(rd);
      System.out.println("Search text replaced " + count + " times\n");
    }
    catch(Exception e)
    {  System.out.println(e);  }
  }  // end of replaceAll()

}  // end of ReplaceAll class

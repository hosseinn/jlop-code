
// DataSort.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, September 2015

/* Sort a complete sheet of data, based on putting its second column
   into ascending order.
*/ 

import com.sun.star.lang.*;
import com.sun.star.frame.*;
import com.sun.star.beans.*;

import com.sun.star.sheet.*;
import com.sun.star.table.*;
import com.sun.star.util.*;


public class DataSort 
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

    // create the table that needs sorting
    Object[][] vals = {
      { "Level", "Code", "No.", "Team", "Name" },
      { "BS", 20, 4, "B", "Elle" }, 
      { "BS", 20, 6, "C", "Sweet" },
      { "BS", 20, 2, "A", "Chcomic" }, 
      { "CS", 30, 5, "A", "Ally" },
      { "MS", 10, 1, "A", "Joker" }, 
      { "MS", 10, 3, "B", "Kevin" },
      { "CS", 30, 7, "C", "Tom" } };
    Calc.setArray(sheet, "A1:E8", vals);    // or just "A1"


    // 1. obtain an XSortable interface for the cell range
    XCellRange sourceRange = Calc.getCellRange(sheet, "A1:E8");
    XSortable xSort = Lo.qi(XSortable.class, sourceRange);

    // 2. specify the sorting criteria as a TableSortField array
    TableSortField[] sortFields = new TableSortField[2];
    sortFields[0] = makeSortAsc(1, true);     // sort by second column
    sortFields[1] = makeSortAsc(2, true);     // then sort by third column

    // 3. define a sort descriptor
    PropertyValue[] props = Props.makeProps("SortFields", sortFields,
                                            "ContainsHeader", true);

    Lo.wait(2000);   // wait so user can see original before it is sorted
    System.out.println("Sorting..."); 
    xSort.sort(props);   // 4. do the sort

    Lo.saveDoc(doc, "dataSort.ods");
    Lo.waitEnter();
    Lo.closeDoc(doc);
    Lo.closeOffice();
  }  // end of main()


  private static TableSortField makeSortAsc(int index, boolean isAscending)
  {
    TableSortField sf = new TableSortField();
    sf.Field = index;
    sf.IsAscending = isAscending;
    sf.IsCaseSensitive = false;
    return sf;
  }  // end of makeSortAsc()

}  // end of DataSort class

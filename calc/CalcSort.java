
// CalcSort.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, September 2015


/* Read in a sequence of doubles from a text file (one number
   per line), and sort them using a LibreOffice Calc spreadsheet.
   The sorted numbers are saved and also printed out.

   Usage:
      run CalcSort unsorted.txt

*/

import java.io.*;
import java.util.*;

import com.sun.star.lang.*;
import com.sun.star.frame.*;
import com.sun.star.beans.*;

import com.sun.star.sheet.*;
import com.sun.star.table.*;
import com.sun.star.util.*;




public class CalcSort
{

  public static void main(String args[])
  {
    if (args.length != 1) {
      System.out.println("Usage: CalcSort fnm");
      return;
    }

    ArrayList<Double> doubles = loadDoubles(args[0]);
    if ((doubles == null) || (doubles.size() == 0)) {
      System.out.println("No doubles read in");
      return;
    }
    int numDoubles = doubles.size();
    System.out.println("Numbe of doubles read in: " + numDoubles);


    XComponentLoader loader = Lo.loadOffice();
    XSpreadsheetDocument doc = Calc.createDoc(loader);
    if (doc == null) {
      System.out.println("doc creation failed");
      Lo.closeOffice();
      return;
    }
    GUI.setVisible(doc, true);

    XSpreadsheet sheet = Calc.getSheet(doc, 0);

    System.out.println("Storing the doubles in a spreadsheet");
    for (int i=0; i < numDoubles; i++)
      Calc.setVal(sheet, 0, i, doubles.get(i));

    
    Lo.wait(2000);
    sortDoubles(sheet, numDoubles);
    Lo.wait(2000);

    System.out.println("Sorted doubles:");
    for (int i=0; i < numDoubles; i++) {
      double d = Calc.getNum(sheet, 0, i);
      System.out.println("  " + d);
    }

    Lo.saveDoc(doc, "sorted.csv");    // i.e. save as text
      // save in CSV format, so the file can be read as text

    Lo.waitEnter();
    Lo.closeDoc(doc);
    Lo.closeOffice();
  }  // end of main()



  public static ArrayList<Double> loadDoubles(String fnm)
  // load doubles from a file, returning them in a list;
  // we assume one double per line
  {
    System.out.println("Reading doubles from " + fnm);
    ArrayList<Double> doubles = new ArrayList<Double>();
    try { 
      BufferedReader reader = new BufferedReader(new FileReader(fnm));
      String line = null;
      while((line = reader.readLine()) != null) {
        Double d = Double.valueOf(line);
        doubles.add(d);
        System.out.println("  Read: " + d);
      }
      reader.close();
    }
    catch (IOException e) 
    {  System.out.println("Problem loading doubles from " + fnm);  }
    return doubles;
  }  // end of loadDoubles()




  public static void sortDoubles(XSpreadsheet sheet, int numDoubles)
  /* sort the specified number of doubles which are assumed to be
     stored in the first column of the spreadsheet
  */
  {
    // sort range is numDoubles cells in the first column
    XCellRange sortRange = Calc.getCellRange(sheet, "A1:A"+numDoubles);
    if (sortRange == null) {
      System.out.println("Sort range incorrect");
      return;
    }

    // sort by first column
    TableSortField[] sortFields = new TableSortField[1];
    sortFields[0] = new TableSortField();
    sortFields[0].Field = 0;
    sortFields[0].IsAscending = true;
    sortFields[0].IsCaseSensitive = false;
    
    // define the sort descriptor
    PropertyValue[] sortDesc = 
        Props.makeProps("SortFields", sortFields);

    // perform the sorting
    XSortable xSort = Lo.qi(XSortable.class, sortRange);
    xSort.sort(sortDesc);
  }  // end of sortDoubles()



}  // end of CalcSort class

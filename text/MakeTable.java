
// MakeTable.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, June 2015

/* Read tabbed text from an input file of Bond movies (bondMovies.txt)
   and store as a blue table in "table.odt".

   Usage:
      run MakeTable
*/

import java.io.*;
import java.util.*;

import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.frame.*;
import com.sun.star.text.*;



public class MakeTable
{
  public static void main(String args[])
  {

    ArrayList<String[]> rowsList = readTable("bondMovies.txt");
    if (rowsList.size() == 0) {
      System.out.println("No data read from bondMovies.txt");
      return;
    }

    XComponentLoader loader = Lo.loadOffice();
    XTextDocument doc = Write.createDoc(loader);
    if (doc == null) {
      System.out.println("Writer doc creation failed");
      Lo.closeOffice();
      return;
    }

    XTextCursor cursor = Write.getCursor(doc);


    Write.appendPara(cursor, "Table of Bond Movies");
    Write.stylePrevParagraph(cursor, "Heading 1");

    Write.appendPara(cursor, "The following table comes from \"bondMovies.txt\":\n");

    Write.addTable(cursor, rowsList);
    Write.endParagraph(cursor);

    Write.append(cursor, "Timestamp: " + Lo.getTimeStamp() + "\n");

    Lo.saveDoc(doc, "table.odt");    // colors don't appear in doc/docx output; ok in odt, pdf
                   // doc, docx, rtf, odt, pdf, txt
    Lo.closeDoc(doc);
    Lo.closeOffice();
  } // end of main()




  private static ArrayList<String[]> readTable(String fnm)
  /* fnm contains table information as a series of lines.
     The first line is assumed to be text for the table header
     and the other lines are for the table data rows.
     Each line consists of text separated by tabs, which become
     the data in the row cells.
     Blank lines, and lines starting with "//" are ignored.
  */
  {
    ArrayList<String[]> rowsList = new ArrayList<String[]>();
    BufferedReader br = null;
    try {
      br = new BufferedReader( new FileReader(fnm));
      System.out.println("Reading table from: " + fnm + "\n");

      boolean isHeader = true;
      int numColumns = 0;
      int rowIdx = 1;    // first row is index 1
      String line;
      while ((line = br.readLine()) != null) {
        line = line.trim();
        //System.out.println("<" + line + ">");
        if (line.length() == 0)
          continue;
        else if (line.startsWith("//"))
          continue;
        else {
          String[] vals = line.split("\t");
          if ((vals == null) || (vals.length == 0))
            System.out.println("Data line is empty");
          else {  // there is some data
            if (isHeader) {
              numColumns = vals.length;
              isHeader = false;
            }
            else {
              if (vals.length != numColumns)
                System.out.println("**** No. of data columns incorrect in row " + rowIdx);
            }
            rowsList.add(vals);
            System.out.print(rowIdx + ". ");
            for(String val : vals)
              System.out.print(val + "  ");
            System.out.println();
            rowIdx++;
          }
        }
      }
    }
    catch (FileNotFoundException ex)
    {  System.out.println("Could not open: " + fnm); }
    catch (IOException ex) 
    {  System.out.println("Read error: " + ex); }
    finally {
      try {
        br.close();
      }
      catch (IOException ex) 
      {  System.out.println("Problem closing " + fnm); }
    }
    System.out.println();
    return rowsList;
  }  // end of readTable()

}  // end of MakeTable class

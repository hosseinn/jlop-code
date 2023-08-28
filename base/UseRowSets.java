
// UseRowSets.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, April 2016

/* Use a Rowset to query the "liangTables.odb" file.

   Use a Rowset listener to report on changes to the result set.

   Needs a call to Lo.killOffice() to properly terminate Office.
*/


import com.sun.star.beans.*;
import com.sun.star.container.*;
import com.sun.star.frame.*;
import com.sun.star.lang.*;
import com.sun.star.uno.*;

import com.sun.star.sdb.*;
import com.sun.star.sdbc.*;



public class UseRowSets
{


  public static void main(String[] args)
  {
    XComponentLoader loader = Lo.loadOffice();
    execQuery();
    Lo.closeOffice();
    Lo.killOffice();    // needed to clean-up after XRowSet
  }  // end of main()



  private static void execQuery()
  {
    try {
      System.out.println("Initializing rowset...");
      XRowSet xRowSet = Base.rowSetQuery("liangTables.odb",
                             "SELECT \"firstName\", \"lastName\" FROM \"Student\"");

      // add rowset listener
      xRowSet.addRowSetListener( new XRowSetListener() {
        public void cursorMoved(EventObject e) 
        {  System.out.println("Cursor moved");  }

        public void rowChanged(EventObject e) 
        {  System.out.println("Row changed");  }

        public void rowSetChanged(EventObject e) 
        {  System.out.println("RowSet changed");  }

        public void disposing(EventObject e)
        {  System.out.println("RowSet being destroyed"); }
      });

      System.out.println("Executing rowset...");
      xRowSet.execute();

      BaseTablePrinter.printResultSet(xRowSet);  
            // possible since XRowSet is a subclass of XResultSet

      // dispose of row set
      XComponent xComp = Lo.qi(XComponent.class, xRowSet);
      xComp.dispose();
    }
    catch(com.sun.star.uno.Exception e)
    {  System.out.println(e);  }
  }  // end of execQuery()


}  // end of UseRowSets class

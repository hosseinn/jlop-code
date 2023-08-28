
// FancyRS.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, April 2016


/* Shows how scrollable and updateable result sets could be utilized 
   if they were available. 

   Begin by creating a sales.odb document which can either hold an 
   embedded HSQLDB or Firebird database with a small SALES table.

     useScrollable() creates a scrollable result set which is neither 
     sensitive to database changes nor can update the underlying database. 
     
     useUpdatable() tries to create a result set that is insensitive to 
     database changes but can write its own changes to the database.  
     
     Since neither of the embedded databases support updating result sets, 
     the code raises an exception when XResultSetUpdate methods are called.

   An embedded HSQLDB database only supports:
    * forward only + read only
    * scrollable; db insensitive + read only

   An embedded Firebird database only supports:
    * forward only + read only
*/

import com.sun.star.frame.*;
import com.sun.star.lang.*;
import com.sun.star.uno.*;
import com.sun.star.util.*;

import com.sun.star.sdb.*;
import com.sun.star.sdbc.*;



public class FancyRS
{
  private static final String FNM = "sales.odb";


  public static void main(String[] args)
  {
    XComponentLoader loader = Lo.loadOffice();
    XOfficeDatabaseDocument dbDoc = Base.createBaseDoc(FNM, // Base.HSQLDB, loader);
                                                            Base.FIREBIRD, loader);
    if (dbDoc == null) {
      Lo.closeOffice();
      return;
    }

    XConnection conn = null;
    try {
      XDataSource dataSource = dbDoc.getDataSource();
      conn = dataSource.getConnection("", "");
    
      Base.exec("CREATE TABLE SALES " +
                   "(SALENR INTEGER NOT NULL, " +
                   " COS_NR INTEGER,  SNR INTEGER, NAME VARCHAR(50)," +
                   " SALEDATE DATE, PRICE FLOAT(10), PRIMARY KEY(SALENR) )", conn);

      Base.exec("INSERT INTO SALES " +
                 "VALUES (1, '100', '1','Linux','2016-02-12',15)", conn);
      Base.exec("INSERT INTO SALES " +
                "VALUES (2, '101', '2','Beef','2016-10-18',15.78)", conn);
      Base.exec("INSERT INTO SALES " +
                "VALUES (3, '104', '4','Juice','2016-08-09',1.5)", conn);
      
     /* Date is wrongly imported into an embedded Firebird database
          e.g.  2016-02-12  ==> 0116-01-12
        This is a known bug:
           http://libreoffice-bugs.freedesktop.narkive.com/fPfsZXpa/
                   bug-91324-new-embedded-firebird-current-date-gives-wrong-date-back
     */

      XFlushable flusher = Lo.qi(XFlushable.class, dataSource);
      flusher.flush(); 
      System.out.println();

      System.out.println("DataSource type: " + Base.getDataSourceType(dbDoc));
      Base.reportDBInfo(conn);
      Base.reportResultSetSupport(conn);
      System.out.println();

      XResultSet rs = Base.executeQuery("SELECT * FROM SALES", conn);
      BaseTablePrinter.printResultSet(rs);  
      
      useScrollable(conn);
      useUpdatable(conn);
    }
    catch(SQLException e) {
      System.out.println(e);
    }
    
   Base.closeConnection(conn);
   Base.closeBaseDoc(dbDoc);
   Lo.closeOffice();
  }  // end of main()




  private static void useScrollable(XConnection conn)
  /* use a scrollable, none-updating resultset 
     see https://wiki.openoffice.org/wiki/Documentation/DevGuide/Database/Scrollable_Result_Sets
      -- only works for HSQLDB
  */
  {
    try {
      XStatement stmt = conn.createStatement();
      Props.setProperty(stmt, "ResultSetType", ResultSetType.SCROLL_INSENSITIVE);
                      // does not update if there are changes to the underlying data
      Props.setProperty(stmt, "ResultSetConcurrency", ResultSetConcurrency.READ_ONLY);
                      // can not update the database
      
      XResultSet srs = stmt.executeQuery("SELECT NAME, PRICE FROM SALES");
      System.out.println("\nReport name and price backwards:");
      // report in reverse order (works for HSQLDB, but not supported by Firebird)
      XRow row = Lo.qi(XRow.class, srs);
      srs.afterLast(); 
      while (srs.previous())   // name: price printed
        System.out.println("  " + row.getString(1) + ": " + row.getFloat(2));
      System.out.println();
    }
    catch(SQLException e) {
      System.out.println("useScrollable(): " + e.getMessage() + "\n");
    }
  }  // end of useScrollable()




  private static void useUpdatable(XConnection conn)
  /* create a scrollable, updating resultset 
     see https://wiki.openoffice.org/wiki/Documentation/DevGuide/Database/Modifiable_Result_Sets
        -- HSQLDB and Firebird do not support updateable result sets
  */
  {
    try {
      XStatement stmt = conn.createStatement();
      Props.setProperty(stmt, "ResultSetType", ResultSetType.SCROLL_INSENSITIVE);
                      // does not update if there are changes to the underlying data
      Props.setProperty(stmt, "ResultSetConcurrency", ResultSetConcurrency.UPDATABLE);
                      // tries to update the database
      
      /* HSQLDB in version 1.8 does not support updating result sets (1.9 may support them).
         The result set type and concurrency are silently downgraded, which 
         is allowed by both the JDBC and the SDBC API.
      */
      
      XResultSet srs = stmt.executeQuery("SELECT NAME, PRICE FROM SALES");
      srs.next();
      XRowUpdate updateRow = Lo.qi(XRowUpdate.class,srs);
      updateRow.updateFloat(2, 25);
          // not possible since UPDATABLE downgraded to read-only; exception time!

      XResultSetUpdate updateRs = Lo.qi(XResultSetUpdate.class, srs);
      updateRs.updateRow();      
         // this call tries to update the data in DBMS; 
         // an exception would occur if execution reached here

      XResultSet rs = Base.executeQuery("SELECT * FROM SALES", conn);
      BaseTablePrinter.printResultSet(rs);  
    }
    catch(SQLException e) {
      System.out.println("useUpdatable(): " + e.getMessage() + "\n");
    }
  }  // end of useUpdatable()



}  // end of FancyRS class


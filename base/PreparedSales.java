
// PreparedSales.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, April 2016


/* Shows how prepared statements are created and used. 

   It starts by creating a salesman.odb file containing a 
   SALESMAN table with 5 rows of data.

   updatePs() updates the table in 4 different ways using a 
   single with a prepared UPDATE statement.

*/

import com.sun.star.frame.*;
import com.sun.star.lang.*;
import com.sun.star.uno.*;
import com.sun.star.util.*;

import com.sun.star.sdb.*;
import com.sun.star.sdbc.*;




public class PreparedSales 
{
  private static final String FNM = "salesman.odb";


  public static void main(String[] args)
  {
    XComponentLoader loader = Lo.loadOffice();
    XOfficeDatabaseDocument dbDoc = Base.createBaseDoc(FNM,  //Base.HSQLDB, loader);
                                                          Base.FIREBIRD, loader);
    if (dbDoc == null) {
      Lo.closeOffice();
      return;
    }

    XConnection conn = null;
    try {
      XDataSource dataSource = dbDoc.getDataSource();
      conn = dataSource.getConnection("", "");

      // create a table in the database
      Base.exec("CREATE TABLE SALESMAN " +
                    "(SNR INTEGER NOT NULL, " +
                    " FIRSTNAME VARCHAR(50),  LASTNAME VARCHAR(100)," +
                    " STREET VARCHAR(50), STATE VARCHAR(50)," +
                    " ZIP INTEGER, BIRTHDATE DATE," +
                    " PRIMARY KEY(SNR)  )", conn);

      Base.exec("INSERT INTO SALESMAN " +
                       "VALUES (1, 'Joseph', 'Smith','Bond Street','CA',95460,"
                       + "'1946-07-02')", conn);
      Base.exec("INSERT INTO SALESMAN " +
                       "VALUES (2, 'Frank', 'Jones','Lake Silver','CA',95460,"
                       + "'1963-12-24')", conn);
      Base.exec("INSERT INTO SALESMAN " +
                       "VALUES (3, 'Jane', 'Esperansa','23 Hollywood drive','CA',95460,"
                       + "'1972-04-01')", conn);
      Base.exec("INSERT INTO SALESMAN " +
                       "VALUES (4, 'George', 'Flint','12 Washington street','CA',95460,"
                       + "'1953-02-13')", conn);
      Base.exec("INSERT INTO SALESMAN " +
                       "VALUES (5, 'Bob', 'Meyers','2 Moon way','CA',95460,"
                       + "'1949-09-07')", conn);

      updatePs(conn);

      XFlushable flusher = Lo.qi(XFlushable.class, dataSource);
      flusher.flush();   // needed or data not saved to file; can be called only once

      //Base.refreshTables(conn);
      //Base.showTables(dbDoc);   // may not close cleanly after showing tables
      //Lo.waitEnter();

      //Base.displayDatabase(conn);

      XResultSet rs = Base.executeQuery("SELECT FIRSTNAME, STREET FROM SALESMAN", conn);
      BaseTablePrinter.printResultSet(rs);  
    }
    catch(SQLException e) {
      System.out.println(e);
    }
    Base.closeConnection(conn);
    Base.closeBaseDoc(dbDoc);
    Lo.closeOffice();
  }  // end of main()





  private static void updatePs(XConnection conn) throws SQLException
  // update the table with a prepared statement; 4 changes made
  {
    XPreparedStatement preStmt = conn.prepareStatement(
                               "UPDATE SALESMAN SET STREET = ? WHERE SNR = ?");

    // change STREET value for salesman Joseph (SNR == 1) to 34 Main Road
    XParameters ps = Lo.qi(XParameters.class, preStmt);
    ps.setString(1, "34 Main Road");
    ps.setInt(2, 1);
    preStmt.executeUpdate();

    // change STREET value for salesman George (SNR == 4) to Marryland
    ps.setString(1, "Marryland");
    ps.setInt(2, 4);
    preStmt.executeUpdate();

    // 2nd change of STREET column of salesman George to Michigan road (since SNR == 4)
    ps.setString(1, "Michigan road");
    preStmt.executeUpdate();

    // change STREET value for salesman Jane (SNR == 3) to Bond Street
    ps.setString(1, "Bond Street");
    ps.setInt(2, 3);
    int numRowsChanged = preStmt.executeUpdate();
    System.out.println("No. of rows changed by executeUpdate(): " + numRowsChanged);  // == 1
  }  // end of updatePs()


}  // end of PreparedSales class


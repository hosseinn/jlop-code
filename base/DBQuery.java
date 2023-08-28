
// DBQuery.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, April 2016

/* Use SQL commands to query a database.

   Base.openBaseDoc() opens the database document. 

   XFlushable.flush() is needed to store any changed table data 
   in the database.

   Base.refreshTables() is needed if Base's GUI is visible and you
   want the Tables view to be updated.

   Base.showTables() open the tables in the Tables view, but if
   these windows are open when the programs exits, then Office hangs.

   Base.displayDatabase() is a better way to show all the tables in
   a series of JFrames.

   This example also reports database metadata using:

      Base.printDataSourceInfo(dataSource);  // can produce large output
      Base.isPasswordRequired(dataSource);
      Base.isReadOnly(dataSource);

      Base.reportDBInfo(conn);
      Base.getTablesNames(conn);

      Base.displayTablesInfo(conn);      // can produce large output
      Base.reportSQLTypes(conn);         // large output
      Base.reportFunctionSupport(conn);   // large output
         -- these last three are commented out to reduce output

   Usage:
     > run DBQuery liangTables.odb    // embedded HSQLDB

     > run DBQuery sales.odb    // embedded firebird DB

     > run DBQuery Books.odb    // link to Books.mdb

     > run DBQuery NewBooks.odb    // conversion of Books.mdb to embedded HSQLDB


   DBCmdsQuery.java is a more complex example of querying a database.
*/

import java.util.*;

import com.sun.star.frame.*;
import com.sun.star.lang.*;
import com.sun.star.uno.*;

import com.sun.star.sdb.*;
import com.sun.star.sdbc.*;




public class DBQuery
{

  public static void main(String[] args)
  {
    if (args.length != 1) {
      System.out.println("Usage: run DBQuery <fnm>");
      return;
    }

    XComponentLoader loader = Lo.loadOffice();
    XOfficeDatabaseDocument dbDoc = Base.openBaseDoc(args[0], loader);
    if (dbDoc == null) {
      System.out.println("Could not open database " + args[0]);
      Lo.closeOffice();
      return;
    }

    System.out.println("Database type: " + Base.getDataSourceType(dbDoc));
    System.out.println("Is embedded? " + Base.isEmbedded(dbDoc));


    XConnection conn = null;
    try {
      XDataSource dataSource = dbDoc.getDataSource();

      System.out.println();
      // Base.printDataSourceInfo(dataSource);
      System.out.println("Is password required: " + 
                                Base.isPasswordRequired(dataSource));
      System.out.println("Is read only: " + Base.isReadOnly(dataSource));

      conn = dataSource.getConnection("", ""); // no login/password
/*
      Base.showTables(dbDoc);    
          // Office may not close cleanly while showing tables
      Lo.waitEnter();
*/
      // Base.printDatabase(conn);
      Base.displayDatabase(conn);

      System.out.println();
      Base.reportDBInfo(conn);

      ArrayList<String> tableNames = Base.getTablesNames(conn);
                              // getTablesNamesMD(conn)
      System.out.println("No. of tables: " + tableNames.size());
      System.out.println( Arrays.toString(tableNames.toArray()));

      // Base.displayTablesInfo(conn);   // can be large
      // Base.reportSQLTypes(conn);   // large
      // Base.reportFunctionSupport(conn);   // large

      System.out.println();
      XResultSet rs = Base.executeQuery("SELECT * FROM \"" + tableNames.get(0) + "\"", conn);

      // Base.printResultSet(rs);
      BaseTablePrinter.printResultSet(rs);  
               // based on https://github.com/htorun/BaseTablePrinter 
               // by Hami Galip Torun
      // Base.displayResultSet(rs);
    }
    catch(SQLException e) {
      System.out.println(e);
    }

    Base.closeConnection(conn);
    Base.closeBaseDoc(dbDoc);
    Lo.closeOffice();
  }  // end of main()



}  // end of DBQuery class

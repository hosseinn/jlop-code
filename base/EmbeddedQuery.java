
// EmbeddedQuery.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, April 2016

// http://www.firebirdsql.org/en/firebird-2-5-3-upd1/

/* Treat the ODB file as a zipped folder and extract its database
   files to baseTmp/. Then execute the database using JDBC with
   Office's HSQLDB driver or the downloaded Firebird driver called
   Jaybird.
 
   Usage:
     > compile EmbeddedQuery.java

     > run EmbeddedQuery sales.odb          // embedded firebird DB
        // added Jaybird to classpath and library path in run.bat

     > run EmbeddedQuery liangTables.odb    // embedded HSQLDB
         // uses <LO>\program\classes\hsqldb.jar

     > run EmbeddedQuery Books.odb          // link to Books.mdb

     > run EmbeddedQuery NewBooks.odb       // conversion of Books.mdb to embedded HSQLDB
*/


import java.io.*;
import java.util.*;
import java.sql.*;




public class EmbeddedQuery 
{

  public static void main(String[] args)
  {
    if (args.length != 1) {
      System.out.println("Usage: run EmbeddedQuery <embedded HSQL/Firebird fnm>");
      return;
    }

    ArrayList<String> dbFnms = Base.extractEmbedded(args[0]);
    if (dbFnms == null)
      return;
    System.out.println();
    // System.out.println( Arrays.toString(dbFnms.toArray()));

    System.out.println("Is this a Firebird embedded database? " + 
                                        Jdbc.isFirebirdEmbedded(dbFnms));
    System.out.println("Is this an HSQLDB embedded database? " + 
                                        Jdbc.isHSQLEmbedded(dbFnms));

    Connection conn = null;
    try {
      conn = Jdbc.connectToDB(dbFnms);
      if (conn == null)
        return;
      System.out.println();
      ArrayList<String> tableNames = Jdbc.getTablesNames(conn);
      System.out.println("No. of tables: " + tableNames.size());
      System.out.println( Arrays.toString(tableNames.toArray()));

      ResultSet rs = Jdbc.executeQuery("SELECT * FROM \"" + tableNames.get(0) + "\"", conn);

      // Jdbc.printResultSet(rs);
      DBTablePrinter.printResultSet(rs);  
      // Jdbc.displayResultSet(rs);

      conn.close();
    }
    catch(SQLException e) {
      System.out.println(e);
    }
    // FileIO.deleteFiles(dbFnms);
  }  // end of main()


}  // end of EmbeddedQuery class

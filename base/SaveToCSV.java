
// SaveToCSV.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, April 2016

/*  Save the supplied ODB file as CSV files, one for each table
    in the database.

    The hard work is done by Base.saveDatabase()

    Usage:
     > run SaveToCSV liangTables.odb    // embedded HSQLDB

     > run SaveToCSV sales.odb    // embedded firebird DB

     > run SaveToCSV Books.odb    // link to Books.mdb

     > run SaveToCSV NewBooks.odb    // conversion of Books.mdb to embedded HSQLDB

*/

import com.sun.star.frame.*;
import com.sun.star.lang.*;
import com.sun.star.uno.*;

import com.sun.star.sdb.*;
import com.sun.star.sdbc.*;



public class SaveToCSV
{

  public static void main(String[] args)
  {
    if (args.length != 1) {
      System.out.println("Usage: run SaveToCSV <fnm>");
      return;
    }

    XComponentLoader loader = Lo.loadOffice();
    XOfficeDatabaseDocument dbDoc = Base.openBaseDoc(args[0], loader);
    if (dbDoc == null) {
      System.out.println("Could not open database " + args[0]);
      Lo.closeOffice();
      return;
    }

    XConnection conn = null;
    try {
      XDataSource dataSource = dbDoc.getDataSource();
      conn = dataSource.getConnection("", "");
      Base.displayTablesSchema(conn);

      // Base.printDatabase(conn);
      // Base.displayDatabase(conn);

      Base.saveDatabase(conn);
    }
    catch(SQLException e) {
      System.out.println(e);
    }

    Base.closeConnection(conn);
    Base.closeBaseDoc(dbDoc);
    Lo.closeOffice();
  }  // end of main()


}  // end of SaveToCSV class

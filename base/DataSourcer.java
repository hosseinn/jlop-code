
// DataSourcer.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, April 2016


/* Use the database context and a data source.

   Shows how "liangTables.odb" can be queried without having 
   to first instantiate XOfficeDatabaseDocument 
   (i.e. without opening the document).

   Requires a call to Lo.killOffice() to properly terminate.

   Shows how to call Base.printRegisteredDataSources()
*/

import java.util.*;

import com.sun.star.beans.*;
import com.sun.star.frame.*;
import com.sun.star.lang.*;
import com.sun.star.uno.*;

import com.sun.star.sdb.*;
import com.sun.star.sdbc.*;



public class DataSourcer
{

  public static void main(String[] args)
  {
    XComponentLoader loader = Lo.loadOffice();

    Base.printRegisteredDataSources();

    XConnection conn = null;
    try {
      XDataSource dataSource = Base.getFileDataSource("liangTables.odb");
      // no office document created or opened

      conn = dataSource.getConnection("", ""); 

      // Base.displayTablesSchema(conn);
      XResultSet rs = Base.executeQuery("SELECT \"firstName\", \"lastName\" FROM \"Student\"", conn);

      BaseTablePrinter.printResultSet(rs);  
    }
    catch(SQLException e) {
      System.out.println(e);
    }

    Base.closeConnection(conn);
    Lo.closeOffice();
    Lo.killOffice();   // needed when Base.getFileDataSource() used 
  }  // end of main()



}  // end of DataSourcer class

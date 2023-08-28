
// CSVQuery.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, April 2016

/* Use the DriverManager to access a CSV file.

   Shows the use of Base.getSupportedDrivers() to
   list supported drivers.

   Shows the use of Base.getTablesNames() and
   Base.getTablesNamesMD() -- the first fails, and the
   second returns too many answers. See the chapter for
   a discussion.
*/

import java.util.*;

import com.sun.star.beans.*;
import com.sun.star.frame.*;
import com.sun.star.lang.*;
import com.sun.star.uno.*;
import com.sun.star.util.*;

import com.sun.star.sdb.*;
import com.sun.star.sdbc.*;
import com.sun.star.sdbcx.*;

import com.sun.star.uno.Exception;



public class CSVQuery
{
  private static final String FNM = "us-500.csv"; 



  public static void main(String[] args)
  {
    XComponentLoader loader = Lo.loadOffice();

    XDriverManager dm = Base.getDriverManager();
    if (dm == null) {
      System.out.println("Could not access Driver manager");
      Lo.closeOffice();
      return;
    }

    ArrayList<String> driveNms = Base.getSupportedDrivers();
    System.out.println("Drivers (" + driveNms.size() + "):");
    for(String driveNm : driveNms)
      System.out.println("  " + driveNm);
    System.out.println();
    /*
      com.sun.star.comp.sdbc.ODBCDriver
      com.sun.star.comp.sdbc.JDBCDriver
      com.sun.star.comp.sdbc.MozabDriver
      com.sun.star.comp.sdbc.ado.ODriver
      com.sun.star.comp.sdbc.calc.ODriver
      com.sun.star.comp.sdbc.dbase.ODriver
      com.sun.star.comp.sdbc.firebird.Driver
      com.sun.star.comp.sdbc.flat.ODriver
      com.sun.star.sdbcx.comp.hsqldb.Driver
      org.openoffice.comp.connectivity.pq.Driver.noext
      org.openoffice.comp.drivers.MySQL.Driver
    */

    String url = "sdbc:flat:" + FileIO.fnmToURL(FNM);
    System.out.println("Using URL: " + url);

    XDriver driver = Base.getDriverByURL(url);
    Base.printDriverProperties(driver, url);

    // set up properties for a CSV file with no password
    PropertyValue[] props = Props.makeProps(
          new String[] { "user", "password", "JavaDriverClass", "Extension",
                         "HeaderLine", "FieldDelimiter", "StringDelimiter" },
          new Object[] { "", "", "com.sun.star.comp.sdbc.flat.ODriver", "csv",
                         true, ",", "\"\"" }
          ); 
    
    XConnection conn = null;
    try {
      conn = dm.getConnectionWithInfo(url, props);

      // ArrayList<String> tableNames = Base.getTablesNames(conn);  // returns null
      ArrayList<String> tableNames = Base.getTablesNamesMD(conn);
      System.out.println("No. of tables: " + tableNames.size());
      System.out.println( Arrays.toString(tableNames.toArray()));
                  // reports all CSV tables ??

      XResultSet rs = Base.executeQuery("SELECT \"first_name\" FROM \"us-500\" WHERE \"City\" = 'New York'", conn);
      BaseTablePrinter.printResultSet(rs);  
    }
    catch(Exception e) {
      System.out.println(e);
    }

    Base.closeConnection(conn);
    Lo.closeOffice();
  }  // end of main()



}  // end of CSVQuery class

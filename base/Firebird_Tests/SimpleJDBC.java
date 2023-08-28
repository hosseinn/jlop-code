
// SimpleJDBC.java
// Andrew Davison, March 2016, ad@fivedots.coe.psu.ac.th

/* A basic JDBC example, using Firebird and its Jaybird
   JDBC driver.

   This example does not use the Office API.

   This program displays the contents of the crossrate table 
   in the crossrate.fdb database. Its format is:

     FROMCURRENCY TOCURRENCY       CONVRATE  UPDATEDATE
     ============ ========== ============== ===========
     Dollar       CdnDlr          1.3273000 22-NOV-2016
     Dollar       FFranc          5.9116001 22-NOV-2016
     Dollar       D-Mark          1.7038000 22-NOV-2016
     Dollar       Lira            1680.0000 22-NOV-2016
     Dollar       Yen             108.43000 22-NOV-2016
     Dollar       Guilder         1.9115000 22-NOV-2016
     Dollar       SFranc          1.4945000 22-NOV-2016
     Dollar       Pound          0.67773998 22-NOV-2016
     Pound        FFranc          8.7340002 22-NOV-2016
     Pound        Yen             159.99001 22-NOV-2016
     Yen          Pound        0.0062500001 22-NOV-2016
     CdnDlr       Dollar         0.75340998 22-NOV-2016
     CdnDlr       FFranc          4.4597001 22-NOV-2016


  Usage:
     > compile SimpleJDBC.java
     > run SimpleJDBC
*/

import java.sql.*;
import java.text.*;   // for timestamp formatting


public class SimpleJDBC 
{

  public static void main(String[] args)
  {
    try {
      Class.forName("org.firebirdsql.jdbc.FBDriver");    
          // requires Jaybird and Firebird
          // this line can be left out in JDBC 4 (present since Java SE 6)

      // connect to the database
      Connection conn = DriverManager.getConnection("jdbc:firebirdsql:embedded:crossrate.fdb", "sysdba", "masterkeys"); 
      // printDriverInfo(conn);

      Statement statement = conn.createStatement();
       
      // Execute a SQL query
      ResultSet rs = statement.executeQuery("SELECT * FROM Crossrate" );

      // Print the result set
      SimpleDateFormat sdf = new SimpleDateFormat("dd-MM-yyyy");
      System.out.println("FromCurrency \tToCurrency \tConvRate \tUpdateDate");
      System.out.println("===========================================================");
      while(rs.next())
        System.out.println( rs.getString("FromCurrency") + ",   \t" +
                            rs.getString("ToCurrency") + ",   \t" +
                            rs.getFloat("ConvRate") + ",   \t" +
                            sdf.format( rs.getTimestamp("UpdateDate")) );
                        // 4th column returns java.sql.Timestamp 
      System.out.println("===========================================================");

      // Close down (should really be in a finally block)
      rs.close();
      statement.close();
      conn.close();
    } 
    catch (ClassNotFoundException e)   // for Class.forName() 
    {  System.out.println("Failed to load driver: \n  " + e); }
    catch (SQLException e) 
    {  for (Throwable t : e)
         System.out.println(t); // to handle a 'chain' of SQLExceptions
    }
  } // end of main()



  private static void printDriverInfo(Connection conn) 
  {
    try {
      DatabaseMetaData meta = conn.getMetaData();
      System.out.println("Driver: " + meta.getDriverName() + ", " + meta.getDriverVersion());

      int jdbcMajor = meta.getJDBCMajorVersion();  // not always supported
      int jdbcMinor = meta.getJDBCMinorVersion();
      System.out.println("JDBC version: " + jdbcMajor + "." + jdbcMinor);
    } 
    catch (Exception e) 
    {  System.out.println("Unsupported feature");  }
    System.out.println();
  }  // end of printDriverInfo()

} // end of SimpleJDBC class


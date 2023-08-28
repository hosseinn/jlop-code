
// SimpleQuery.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, April 2016

/* Two queries of the HSQLDB database embedded inside an ODB file called 
   liangTables.odb. It has three tables, called Course, Enrollment, 
   and Student.

   The ODB file is based on an example in Chapter 32 of
   "Intro to Java Programming", Y. Daniel Liang, 10th ed.
*/

import java.util.*;

import com.sun.star.frame.*;
import com.sun.star.lang.*;
import com.sun.star.uno.*;

import com.sun.star.sdb.*;
import com.sun.star.sdbc.*;




public class SimpleQuery
{
  private static final String FNM = "liangTables.odb";


  public static void main(String[] args)
  {

    XComponentLoader loader = Lo.loadOffice();
    XOfficeDatabaseDocument dbDoc = Base.openBaseDoc(FNM, loader);
    if (dbDoc == null) {
      System.out.println("Could not open database " + FNM);
      Lo.closeOffice();
      return;
    }

    XConnection conn = null;
    try {
      XDataSource dataSource = dbDoc.getDataSource();
      conn = dataSource.getConnection("", "");  // no login/password

      XStatement statement = conn.createStatement();

      // first query
      XResultSet rs = statement.executeQuery("SELECT * FROM \"Course\"");

      XRow xRow = Lo.qi(XRow.class, rs);

      System.out.println("CourseID \tSubjectID \tCourseNumber \tTitle             \tNumOfCredits");
      System.out.println("===================================================================================");
      while(rs.next())
        System.out.println( xRow.getString(1) + ",   \t" + xRow.getString(2) + ",   \t" +
                            xRow.getInt(3) + ",   \t" + xRow.getString(4) + ",   " +
                            xRow.getInt(5) );
      System.out.println("===================================================================================");

      // second query
      rs = statement.executeQuery("SELECT \"courseNumber\", \"title\" FROM \"Course\"");
      xRow = Lo.qi(XRow.class, rs);
      XColumnLocate xLoc = Lo.qi( XColumnLocate.class, rs);

      System.out.println("CourseNumber \tTitle");
      System.out.println("========================================");
      while(rs.next())
        System.out.println( xRow.getString( xLoc.findColumn("courseNumber")) + 
                            ",   \t" + 
                            xRow.getString( xLoc.findColumn("title")) );
      System.out.println("========================================");

    }
    catch(SQLException e) {
      System.out.println(e);
    }

    Base.closeConnection(conn);
    Base.closeBaseDoc(dbDoc);
    Lo.closeOffice();
  }  // end of main()



}  // end of SimpleQuery class

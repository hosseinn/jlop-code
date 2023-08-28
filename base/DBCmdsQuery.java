
// DBCmdsQuery.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, April 2016

/* Use SQL commands in a text file to query a database.

   Base.openBaseDoc() opens the database document. 

   XFlushable.flush() is needed to store any changed table data 
   in the database.

   Base.refreshTables() is needed if Base's GUI is visible and you
   want the Tables view to be updated.

   Base.showTables() open the tables in the Tables view, but if
   these windows are open when the programs exits, then Office hangs.

   Base.displayDatabase() is a better way to show all the tables in
   a series of JFrames.


   Usage:
     run DBCmdsQuery liangTables.odb liangTests.txt

   The ODB file is based on an example in Chapter 32 of
   "Intro to Java Programming", Y. Daniel Liang, 10th ed.

   DBQuery.java is a simpler example of querying a database.
*/


import java.util.*;

import com.sun.star.frame.*;
import com.sun.star.lang.*;
import com.sun.star.uno.*;
import com.sun.star.util.*;

import com.sun.star.sdb.*;
import com.sun.star.sdbc.*;




public class DBCmdsQuery
{

  public static void main(String[] args)
  {
    if (args.length != 2) {
      System.out.println("Usage: java DBCmdsQuery <DB fnm> <cmds fnm>");
      return;
    }

    XComponentLoader loader = Lo.loadOffice();
    XOfficeDatabaseDocument dbDoc = Base.openBaseDoc(args[0], loader);
    if (dbDoc == null) {
      System.out.println("Could not open " + args[0]);
      Lo.closeOffice();
      return;
    }


    ArrayList<String> cmds = Base.readCmds(args[1]);
    if (cmds == null) {
      System.out.println("No commands to process");
      Lo.closeOffice();
      return;
    }
    System.out.println("Read in " + cmds.size() + " commands");
    //for (String cmd : cmds)
    //   System.out.println("\"" + cmd + "\"\n");


    XConnection conn = null;
    try {
      XDataSource dataSource = dbDoc.getDataSource();
      conn = dataSource.getConnection("", "");

      processCmds(cmds, conn);

      XFlushable flusher = Lo.qi(XFlushable.class, dataSource);
      flusher.flush(); // needed or data not saved to file; can only be called once

      // Base.refreshTables(conn);
            /* must refresh or new table(s) will not be visible 
               inside the Tables window of the Base app; this must
               be done after the call to flush() */
      // Base.showTablesView(dbDoc)
      // Lo.waitEnter();
  
      // Base.displayDatabase(conn);
    }
    catch(SQLException e) {
      System.out.println(e);
    }

    Base.closeConnection(conn);
    Base.closeBaseDoc(dbDoc);
    Lo.closeOffice();
  }  // end of main()



  private static void processCmds(ArrayList<String> cmds, XConnection conn)
  // filter out CREATE commands since this example is only querying the db
  {
    for(String cmd : cmds) {
      if (cmd.startsWith("CREATE"))
        System.out.println("Ignoring create commands: \"" + cmd + "\"");
      else
        Base.exec(cmd, conn);
    }
  }  // end of processCmds()


}  // end of DBCmdsQuery class

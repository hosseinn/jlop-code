
// DBCmdsCreate.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, April 2016

/* Use SQL commands in a text file to create an embedded HSQLDB
   or Firebird database.

   Base.createBaseDoc() creates a new file containing an empty
   database. 

   XFlushable.flush() is needed to store table data in the database.

   Base.refreshTables() is needed if Base's GUI is visible and you
   want the Tables view to be updated.

   Base.showTables() open the tables in the Tables view, but if
   these windows are open when the programs exits, then Office hangs.

   Base.displayDatabase() is a better way to show all the tables in
   a series of JFrames.

   Usage:
     run DBCmdsCreate liangTables.txt
                  // creates liangTables.odb

   The ODB file is based on an example in Chapter 32 of
   "Intro to Java Programming", Y. Daniel Liang, 10th ed.


   DBCreate.java is a simpler example of creating a database.
*/


import java.util.*;

import com.sun.star.frame.*;
import com.sun.star.lang.*;
import com.sun.star.uno.*;
import com.sun.star.util.*;

import com.sun.star.sdb.*;
import com.sun.star.sdbc.*;



public class DBCmdsCreate
{

  public static void main(String[] args)
  {
    if (args.length != 1) {
      System.out.println("Usage: java DBCmdsCreate <cmds fnm>");
      return;
    }
    String fnm = Info.getName(args[0]) + ".odb";

    XComponentLoader loader = Lo.loadOffice();
    XOfficeDatabaseDocument dbDoc = Base.createBaseDoc(fnm, Base.HSQLDB, loader);
                                                        // Base.FIREBIRD, loader);
    if (dbDoc == null) {
      Lo.closeOffice();
      return;
    }

    ArrayList<String> cmds = Base.readCmds(args[0]);
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
      conn = dataSource.getConnection("", "");  // no login/password
      processCmds(cmds, conn);

      XFlushable flusher = Lo.qi(XFlushable.class, dataSource);
      flusher.flush();   // needed or data not saved to file

      // must refresh the connection or the table will not be visible inside Base app
      Base.refreshTables(conn);
      Base.showTables(dbDoc);   
            // Office may not close cleanly while showing tables
      Lo.waitEnter();

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
  // filter out SELECT commands since only creating in this example
  {
    for(String cmd : cmds) {
      if (cmd.startsWith("SELECT"))
        System.out.println("Ignoring select commands: \"" + cmd + "\"");
      else
        Base.exec(cmd, conn);
    }
  }  // end of processCmds()



}  // end of DBCmdsCreate class

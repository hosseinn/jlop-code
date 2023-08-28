
// DBRels.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, April 2016

/*  Generate a relation design diagram programmatically  
    via dispatch commands, Java Robot, and the clipboard.
    The diagram is saved as an image in the "relations.png" 
    file.

   Usage:
     > run DBRels liangTables.odb    // embedded HSQLDB

     > run DBRels sales.odb    // embedded firebird DB  
           // relation design not supported, so nothing generated

     > run DBRels Books.odb    // link to Books.mdb

     > run DBRels NewBooks.odb    // conversion of Books.mdb to embedded HSQLDB


*/

import java.util.*;

import com.sun.star.frame.*;
import com.sun.star.lang.*;
import com.sun.star.uno.*;

import com.sun.star.sdb.*;
import com.sun.star.sdbc.*;

import java.awt.image.*;



public class DBRels
{

  public static void main(String[] args)
  {
    if (args.length != 1) {
      System.out.println("Usage: run DBRels <fnm>");
      return;
    }

    XComponentLoader loader = Lo.loadOffice();
    XOfficeDatabaseDocument dbDoc = Base.openBaseDoc(args[0], loader);
    if (dbDoc == null) {
      System.out.println("Could not open database " + args[0]);
      Lo.closeOffice();
      return;
    }

    if (Base.isFirebirdEmbedded(dbDoc))
      System.out.println("Embedded Firebird does not support relation design");
    else {
      GUI.setVisible(dbDoc, true);
      Lo.delay(500);  // wait for GUI to appear

      Lo.dispatchCmd("DBRelationDesign");
      Lo.delay(1000);

      JNAUtils.shootWindow();
      Lo.delay(1000);

      // access the clipboard
      BufferedImage im = Clip.readImage();
      if (im != null)
        Images.saveImage(im, "relations.png");
      
      Lo.dispatchCmd("CloseWin");  // close the relation design window
      Lo.delay(500);
    }

    // Lo.waitEnter();
    Base.closeBaseDoc(dbDoc);
    Lo.closeOffice();
  }  // end of main()



}  // end of DBRels class

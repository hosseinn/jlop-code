
// PoemCreator.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, September 2016

/* Create a poem, save as "poem.doc"

   Usage:
     > compileExt RandomSents PoemCreator.java

     > runExt RandomSents PoemCreator

    -- must supply the service name of the UNO component extension
       as well as Java filename

*/

import java.io.*;

import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.frame.*;

import com.sun.star.text.*;
import com.sun.star.style.*;
import com.sun.star.container.*;
import com.sun.star.beans.*;

import org.openoffice.randomsents.XRandomSents;



public class PoemCreator
{
  public static void main(String args[])
  {

    XComponentLoader loader = Lo.loadOffice();
    XTextDocument doc = Write.createDoc(loader);
    if (doc == null) {
      System.out.println("Writer doc creation failed");
      Lo.closeOffice();
      return;
    }

    Info.listExtensions();

    GUI.setVisible(doc, true);

    Write.setHeader(doc, "Muse of the Office");
    Write.setA4PageFormat(doc);
    Write.setPageNumbers(doc);

    // use the generated default create() method to instantiate a new RandomSents object
    // XRandomSents rs = RandomSents.create( Lo.getContext() );   // RandomSents not found

    XRandomSents rs = Lo.createInstanceMCF(XRandomSents.class,
                                "org.openoffice.randomsents.RandomSents");
    String[] sents = rs.getSentences(5);

    XTextCursor cursor = Write.getCursor(doc);
    for(String sent : sents)
       Write.appendPara(cursor, sent+"\n");

    rs.setisAllCaps(true);
    Write.appendPara(cursor, rs.getParagraph(2)+"\n");

    Write.appendPara(cursor, Lo.getTimeStamp());

    Lo.waitEnter();
    Lo.saveDoc(doc, "poem.doc");

    Lo.closeDoc(doc);
    Lo.closeOffice();
  } // end of main()

}  // end of PoemCreator class



// FormOutliner.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, June 2016

/* Open a text document containing a form, and display information
   about its controls.

   Usage:
       run FormOutliner form1.odt
                // simple form

       run FormOutliner scratchExample.odt
               // more realistic form; longer

       run FormOutliner build.odt
               // saved version of form created by BuildForm.java
*/

import java.util.ArrayList;

import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.frame.*;

import com.sun.star.container.*;
import com.sun.star.awt.*;




public class FormOutliner
{

  public static void main(String args[])
  {
    if (args.length != 1) {
      System.out.println("Usage: run FormOutliner <ODT file>");
      return;
    }
    XComponentLoader loader = Lo.loadOffice();
    XComponent doc = Lo.openDoc(args[0], loader);
    if (doc == null) {
      System.out.println("Could not open " + args[0]);
      Lo.closeOffice();
      return;
    }

    Forms.listForms(doc);    
             // print info about all the forms and their components

    Lo.closeDoc(doc);
    Lo.closeOffice();
  }  // end of main()


}  // end of FormOutliner class

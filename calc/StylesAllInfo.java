
// StylesAllInfo.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, June 2015

/* Report all the style family names used in the document.
   For a Calc document there are two style family names:
        CellStyles
        PageStyles
   List the container names used by each style family.

    Usage:
      run StylesAllInfo totals.ods
*/


import com.sun.star.lang.*;
import com.sun.star.frame.*;
import com.sun.star.beans.*;

import com.sun.star.text.*;
import com.sun.star.style.*;



public class StylesAllInfo
{

  public static void main(String args[])
  {
    if (args.length < 1) {
      System.out.println("Usage: run StylesAllInfo <fnm>");
      return;
    }

    XComponentLoader loader = Lo.loadOffice();
    XComponent doc = Lo.openDoc(args[0], loader);
    if (doc == null) {
      System.out.println("Could not open " + args[0]);
      Lo.closeOffice();
      return;
    }


    // get all the style families for this document
    String[] styleFamilies = Info.getStyleFamilyNames(doc);
    System.out.println("Style Family Names (" + styleFamilies.length + "): ");
    for(String styleFamily : styleFamilies)
      System.out.println("  " + styleFamily);
    System.out.println();


    // list all the style names for each style family
    for(int i=0; i < styleFamilies.length; i++) {
      String styleFamily = styleFamilies[i];
      System.out.println((i+1) + ". \"" + styleFamily + "\" Family styles:");
      String[] styleNames = Info.getStyleNames(doc, styleFamily);
      Lo.printNames(styleNames);
    }


    // --- Report the Default CellStyles and PageStyles (long) ----------
/*
    Props.showProps("CellStyles Default", 
                  Info.getStyleProps(doc, "CellStyles", "Default") );

    Props.showProps("PageStyles Default", 
                 Info.getStyleProps(doc, "PageStyles", "Default") );
*/

/*
    // create new style
    XStyle style =  Lo.createInstanceMSF(XStyle.class, 
                             "com.sun.star.style.CellStyle");
                                          // no DisplayName, and more
                            // "com.sun.star.style.PageStyle");    //ok
                                          // DisplayName == "Default Style"
    Props.showObjProps("New Cell style", style);   // long
    System.out.println();
*/
    Lo.closeDoc(doc);
    Lo.closeOffice();
  }  // end of main()


}  // end of StylesAllInfo class


// StylesInfo.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, June 2015

/* Report all the style family names used in the document.
   For a text document there are five style family names.
   List the container names used by each style family.

   Print text document style family + "Standard" container property info on:
       page styles
       paragraph styles    -- lots of container names (e.g. "Header 1", "Numbering 1 Start")

    Usage:
      run StylesInfo story.doc
*/


import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.frame.*;
import com.sun.star.text.*;


public class StylesInfo
{

  public static void main(String args[])
  {
    if (args.length < 1) {
      System.out.println("Usage: run StylesInfo <fnm>");
      return;
    }

    XComponentLoader loader = Lo.loadOffice();
    XTextDocument doc = Write.openDoc(args[0], loader);
    if (doc == null) {
      System.out.println("Could not open " + args[0]);
      Lo.closeOffice();
      return;
    }

    // get all the style families for this document
    String[] styleFamilies = Info.getStyleFamilyNames(doc);
    System.out.println("No. of Style Family Names: " + styleFamilies.length);
    for(String styleFamily : styleFamilies)
      System.out.println("  " + styleFamily);
    System.out.println();

    // list all the style names for each style family
    for(int i=0; i < styleFamilies.length; i++) {
      String styleFamily = styleFamilies[i];
      System.out.println((i+1) + ". \"" + styleFamily + "\" Style Family contains containers:");
      String[] styleNames = Info.getStyleNames(doc, styleFamily);
      Lo.printNames(styleNames);
    }


    /* Report the properties for the paragraph styles family
       under the "Standard" name */
    Props.showProps("ParagraphStyles \"Standard\"", 
                 Info.getStyleProps(doc, "ParagraphStyles", "Header") );
            // long

    // access other style families, other names...
    //Props.showProps("FrameStyles \"Graphics\"", 
    //             Info.getStyleProps(doc, "FrameStyles", "Graphics") );


    //Props.showProps("NumberingStyles \"List 1\"", 
    //             Info.getStyleProps(doc, "NumberingStyles", "List 1") );

    //Props.showProps("PageStyles \"Envelope\"", 
    //             Info.getStyleProps(doc, "PageStyles", "Standard") );

    System.out.println();


    Lo.closeDoc(doc);
    Lo.closeOffice();
  }  // end of main()


}  // end of StylesInfo class

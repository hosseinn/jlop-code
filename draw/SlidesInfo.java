
// SlidesInfo.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, July 2015

/*  Features: 
      - open a draw/slides document
      - resize page in application window
      - report on layers
      - report style information: familiies, styles, properties
      - show no. of slides and slide size


   Usage:
      run SlidesInfo points.odp
      run SlidesInfo algs.odp
      run SlidesInfo cover.odg    // works with drawings too

*/

import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.frame.*;

import com.sun.star.drawing.*;

import com.sun.star.container.*;
import com.sun.star.awt.*;
import com.sun.star.beans.*;


public class SlidesInfo
{

  public static void main(String args[])
  {
    if (args.length != 1) {
      System.out.println("Usage: SlidesInfo fnm");
      return;
    }

    XComponentLoader loader = Lo.loadOffice();
    XComponent doc = Lo.openDoc(args[0], loader);
    if (doc == null) {
      System.out.println("Could not open the file: " + args[0]);
      Lo.closeOffice();
      return;
    }

    if (!Draw.isShapesBased(doc)) {
      System.out.println("-- not a drawing or slides presentation");
      Lo.closeOffice();
      return;
    }

    GUI.setVisible(doc, true);
    Lo.delay(1000);    // need delay or zoom may not occur

    GUI.zoom(doc, GUI.ENTIRE_PAGE);
           // GUI.OPTIMAL, GUI.PAGE_WIDTH, GUI.ENTIRE_PAGE
    // GUI.zoomValue(doc, 60);   // percentage size


    System.out.println("\nNo of slides: " + Draw.getSlidesCount(doc) + "\n");

    XDrawPage currSlide = Draw.getSlide(doc, 0);   // access first page

    Size slideSize = Draw.getSlideSize(currSlide);
    System.out.println("Size of slide (mm): (" + slideSize.Width + ", " +
                                                 slideSize.Height + ")\n");

    reportLayers(doc);
    reportStyles(doc);

    System.out.println("Waiting a bit before closing...");
    Lo.delay(3000);
    Lo.closeDoc(doc);
    Lo.closeOffice();
  } // end of main()



  private static void reportLayers(XComponent doc)
  {
    XLayerManager lm = Draw.getLayerManager(doc);
    for(int i=0; i < lm.getCount(); i++)
      try {
        Props.showObjProps(" Layer " + i, lm.getByIndex(i));
      }
      catch(com.sun.star.uno.Exception e) {}

    XLayer layer = Draw.getLayer(doc, "backgroundobjects");
         // "layout", "background", "backgroundobjects", "controls", "measurelines"
    Props.showObjProps("Background Object Props", layer);
  }  // end of reportLayers()



  private static void reportStyles(XComponent doc)
  // report style information: families, styles, properties
  {
    String[] styleNames = Info.getStyleFamilyNames(doc);
    System.out.println("Style Families in this document:");
    Lo.printNames(styleNames);
         // usually: "Default"  "cell"  "graphics"  "table"
         // Default is the name of the default Master Page template inside Office

    for (String styleName : styleNames) {
      String[] conNames = Info.getStyleNames(doc, styleName);
      System.out.println("Styles in the \"" + styleName + "\" style family:");
      Lo.printNames(conNames);
    }

    // XPropertySet props = Info.getStyleProps(doc, "Inspiration", "outline1");   //Default
    // XPropertySet props = Info.getStyleProps(doc, "graphics", "objectwitharrow");
               // accessing a style causes Office to crash upon exiting

    // Props.showProps("Object Arrow Properties", props);     // long
  }  // end of reportStyles()



}  // end of SlidesInfo class

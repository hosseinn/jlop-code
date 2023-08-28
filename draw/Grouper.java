
// Grouper.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, July 2015

/* Features:
      - create draw document
      - resize page in application window

      - connect two rectangles (style changes causes crash on exit)
      - grouping/binding/composition of shapes
*/

import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.frame.*;
import com.sun.star.drawing.*;

import com.sun.star.beans.*;
import com.sun.star.awt.*;
import com.sun.star.container.*;



public class Grouper
{

  public static void main (String args[])
  {
    XComponentLoader loader = Lo.loadOffice();
    XComponent doc = Draw.createDrawDoc(loader);
    if (doc == null) {
      System.out.println("Draw doc creation failed");
      Lo.closeOffice();
      return;
    }

    GUI.setVisible(doc, true);
    Lo.delay(1000);    // need delay or zoom may not occur
    GUI.zoom(doc, GUI.ENTIRE_PAGE);

    XDrawPage currSlide = Draw.getSlide(doc, 0);   // access first page

    System.out.println("\nConnecting rectangles ...");
    XNameContainer gStyles = Info.getStyleContainer(doc, "graphics");
    //Props.showObjProps("Objects with Arrow Graphics Style", 
    //        Info.getStyleProps(doc, "graphics", "objectwitharrow"));
    connectRectangles(currSlide, gStyles);

    
    // create two ellipses
    Size slideSize = Draw.getSlideSize(currSlide);
    int width = 40;
    int height = 20;
    int x = (slideSize.Width*3)/4 - width/2;
    int y1 = 20;
    int y2 = slideSize.Height/2 - (y1 + height);  // so separated
    // int y2 = 30;     // so overlapping

    XShape s1 = Draw.drawEllipse(currSlide, x, y1, width, height);
    XShape s2 = Draw.drawEllipse(currSlide, x, y2, width, height);

    Draw.showShapesInfo(currSlide);

    // group, bind, or combine the ellipses
    System.out.println("\nGrouping (or binding) ellipses ...");
    // groupEllipses(currSlide, s1, s2);
    // bindEllipses(currSlide, s1, s2);
    combineEllipses(currSlide, s1, s2);
    Draw.showShapesInfo(currSlide);

    // combine some rectangles
    XShape compShape = combineRects(doc, currSlide);
    Draw.showShapesInfo(currSlide);

    System.out.println("Waiting a bit before splitting...");
    Lo.delay(3000);   // delay so user can see previous composition

    // split the combination into component shapes
    System.out.println("\nSplitting the combination ...");
    XShapeCombiner combiner = Lo.qi(XShapeCombiner.class, currSlide);
    combiner.split(compShape);
    Draw.showShapesInfo(currSlide);

    Lo.saveDoc(doc, "grouper.odg");

    System.out.println("Wait for enter...");
    Lo.waitEnter();
    Lo.closeDoc(doc);
    Lo.closeOffice();
  } // end of main()




  private static void connectRectangles(XDrawPage currSlide,
                                                  XNameContainer gStyles)
  /* draw two two labelled rectangles, one green, one blue, and
     connect them. Changing the connector to an arrow
     causes a crash.  */
  {
    // dark green rectangle with shadow and text
    XShape greenRect = Draw.drawRectangle(currSlide, 70, 180, 50, 25);
    Props.setProperty(greenRect, "FillColor", 
                              Lo.getColorInt(java.awt.Color.GREEN.darker()));
    Props.setProperty(greenRect, "Shadow", true);
    Draw.addText(greenRect, "Green Rect");

    // (blue, the default color) rectangle with shadow and text
    XShape blueRect = Draw.drawRectangle(currSlide, 140, 220, 50, 25);
    Props.setProperty(blueRect, "Shadow", true);
    Draw.addText(blueRect, "Blue Rect");

    
    // connect the two rectangles; from the first shape to the second
    XShape connShape = Draw.addConnector(currSlide, greenRect, Draw.CONNECT_BOTTOM, 
                                                    blueRect, Draw.CONNECT_TOP);

    Draw.setStyle(connShape, gStyles, "objectwitharrow");
        /* arrow added at the 'from' end of the connector shape
           and it thickens line and turns it black */
        // causes a crash at Office close time ??
    Props.setProperty(connShape, "LineWidth", 50);  // 0.5 mm
    Props.setProperty(connShape, "FillColor", 7512015);  // dark blue; does not work
    // Props.showObjProps("Connector Shape", connShape);

    // change the arrow style and have it finish at the end of the line
    Props.setProperty(connShape, "LineStartName", "Short line arrow");
    Props.setProperty(connShape, "LineStartCenter", false);


    // report the glue points for the blue rectangle
    GluePoint2[] gps = Draw.getGluePoints(blueRect);
    if (gps != null) {
      System.out.println("Glue Points for blue rectangle");
      for(int i=0; i < gps.length; i++) {
        Point pos = gps[i].Position;
        System.out.println("  Glue point " + i + ": (" + pos.X + ", " + pos.Y + ")");
      }
    }
  }  // end of connectRectangles()




  private static void groupEllipses(XDrawPage currSlide, XShape s1, XShape s2)
  {
    XShape shapeGroup = Draw.addShape(currSlide, "GroupShape",  0, 0, 0, 0);
    XShapes shapes = Lo.qi(XShapes.class, shapeGroup);
    //XShapeGrouper xShapeGrouper = Lo.qi(XShapeGrouper.class, currSlide);
        // XShapeGrouper is deprecated; use GroupShape instead
    shapes.add(s1);
    shapes.add(s2);
  }  // end of groupEllipses()



  private static void bindEllipses(XDrawPage currSlide, XShape s1, XShape s2)
  {
    XShapes shapes = 
         Lo.createInstanceMCF(XShapes.class, "com.sun.star.drawing.ShapeCollection");
    shapes.add(s1);
    shapes.add(s2);
    XShapeBinder binder = Lo.qi(XShapeBinder.class, currSlide);
    binder.bind(shapes);
  }  // end of bindEllipses()


  private static void combineEllipses(XDrawPage currSlide, 
                                        XShape s1, XShape s2)
  {
    XShapes shapes =  Lo.createInstanceMCF(XShapes.class, 
                            "com.sun.star.drawing.ShapeCollection");
    shapes.add(s1);
    shapes.add(s2);
    XShapeCombiner combiner = Lo.qi(XShapeCombiner.class, currSlide);
    combiner.combine(shapes);
  }  // end of combineEllipses()



  private static XShape combineRects(XComponent doc, XDrawPage currSlide)
  // create a single PolyPolygon shape made from two rectangles
  {
    System.out.println("\nCombining rectangles ...");
    XShape r1 = Draw.drawRectangle(currSlide, 50, 20, 40, 20);
    XShape r2 = Draw.drawRectangle(currSlide, 70, 25, 40, 20);

    XShapes shapes = 
       Lo.createInstanceMCF(XShapes.class, "com.sun.star.drawing.ShapeCollection");
    shapes.add(r1);
    shapes.add(r2);
    return Draw.combineShape(doc, shapes, Draw.COMBINE);
                        // Draw.MERGE, Draw.SUBTRACT, Draw.INTERSECT, Draw.COMBINE
  }  // end of combineRects()



} // end of Grouper class


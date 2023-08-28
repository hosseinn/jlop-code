
// AnimBicycle.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, July 2015

/*  Create a page containing:
      - two polgons
      - a PolyLineShape
      - part of an ellipse
      - bicycle graphic from the Gallery, animated to move right
        and rotate counter-clockwise.

*/

import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.frame.*;

import com.sun.star.text.*;
import com.sun.star.drawing.*;
import com.sun.star.container.*;

import com.sun.star.util.*;
import com.sun.star.awt.*;

import com.sun.star.gallery.*;


public class AnimBicycle
{

  public static void main (String args[])
  {
    XComponentLoader loader = Lo.loadOffice();

    // create Impress page or Draw slide
    //XComponent doc = Draw.createImpressDoc(loader);
    XComponent doc = Draw.createDrawDoc(loader);
    if (doc == null) {
      System.out.println("Doc creation failed");
      Lo.closeOffice();
      return;
    }

    XDrawPage currSlide = Draw.getSlide(doc, 0);
    if (Draw.isImpress(doc))
      Draw.titleOnlySlide(currSlide, "Bicycle in Motion");

    GUI.setVisible(doc, true);

    Size slideSize = Draw.getSlideSize(currSlide);

    XShape square = Draw.drawPolygon(currSlide, 125, 125, 25, 4);
                                        // (x,y), radius, no. of sides
    Props.setProperty(square, "FillColor", 0x3fe694);  // pale green
              // obtained colors from http://www.color-hex.com/

    XShape pentagon = Draw.drawPolygon(currSlide, 150, 75, 5);
                                        // (x,y), default radius of 20, no. of sides
    Props.setProperty(pentagon, "FillColor", 0xe7b9c7);  // purple


    int[] xs = {10, 30, 10, 30};
    int[] ys = {10, 100, 100, 10};
    Draw.drawLines(currSlide, xs, ys);

    XShape pie = Draw.drawEllipse(currSlide, 30, slideSize.Width-100, 40, 20);
    Props.setProperty(pie, "CircleStartAngle", 9000);   // 90 degrees ccw
    Props.setProperty(pie, "CircleEndAngle", 36000);    // 360 degrees ccw
    Props.setProperty(pie, "CircleKind", CircleKind.SECTION);
                     // CircleKind.SECTION, CircleKind.CUT, CircleKind.ARC

    animateBike(currSlide);

    Lo.saveDoc(doc, "bicycle.odg");

    System.out.println("Waiting a bit before closing...");
    Lo.delay(3000);
    Lo.closeDoc(doc);
    Lo.closeOffice();
  } // end of main()


  private static void animateBike(XDrawPage currSlide)
  // bike moves to the right and rotates ccw
  {
    String fnm = Info.getGalleryDir() + "/transportation/Bicycle-Blue.png";
    //XGalleryItem item = Gallery.findGalleryItem("Transportation", "bicycle");
    //String fnm = Gallery.getGalleryPath(item);

    XShape shape = Draw.drawImage(currSlide, fnm, 60, 100, 90, 50);
                                                  // (x,y), width, height
    if (shape == null) {
      System.out.println("Bike shape not created");
      return;
    }

    Point pt = Draw.getPosition(shape);
    int angle = Draw.getRotation(shape);
    System.out.println("Start Angle: " + angle);
    for (int i=0; i <= 18; i++) {
       Draw.setPosition(shape, pt.X+(i*5), pt.Y); // move right
       Draw.setRotation(shape, angle+(i*5));   // rotates ccw
       Lo.delay(200);
    }

    System.out.println("Final Angle: " + Draw.getRotation(shape));
    Draw.printMatrix( Draw.getTransformation(shape));
  }  // end of animateBike()



} // end of AnimBicycle class



// DrawPicture.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, July 2015

/* Features:
      - create draw document
      - resize page in application window
      - draw various shapes: line, ellipse, rectangle, text, circle, polar line
      - different fills: solid, gradients, hatching, bitmaps
      - create an OLE shape (a math formulae)
      - animate a circle and a line
*/

import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.frame.*;
import com.sun.star.drawing.*;

import com.sun.star.beans.*;
import com.sun.star.awt.*;
import com.sun.star.container.*;


public class DrawPicture
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
    drawShapes(currSlide);

    Draw.drawFormula(currSlide, "func e^{i %pi} + 1 = 0",  23, 20, 20, 40);

    animShapes(currSlide);

    XShape s = Draw.findShapeByName(currSlide, "text1");
    Draw.reportPosSize(s);

    Lo.saveDoc(doc, "picture.odg");

    System.out.println("Waiting a bit before closing...");
    Lo.delay(3000);
    Lo.closeDoc(doc);
    Lo.closeOffice();
  } // end of main()



  private static void drawShapes(XDrawPage currSlide)
  // Draw: line, ellipse, rectangle, text, circle, polar line
  {
    // black dashed line; uses (x1,y1) (x2,y2)
    XShape line1 = Draw.drawLine(currSlide, 50, 50, 200, 200);
    Props.setProperty(line1, "LineColor", 0x000000);   // black
    Draw.setDashedLine(line1, true);
    // Lo.delay(2000);
    // Draw.setDashedLine(line1, false);

    // red ellipse; uses (x, y) width, height
    XShape circle1 = Draw.drawEllipse(currSlide, 100, 100, 75, 25);
    Props.setProperty(circle1, "FillColor", 0xFF0000);

    // rectangle with different fills; uses (x, y) width, height
    XShape rect1 = Draw.drawRectangle(currSlide, 70, 70, 25, 50);
    Props.setProperty(rect1, "FillColor", 0x00FF00);
    // Draw.setGradientColor(rect1, "Gradient 4");
                            // "Radial red/yellow");
    // Draw.setGradientColor(rect1, java.awt.Color.GREEN, java.awt.Color.RED);

    // Draw.setHatchingColor(rect1, "Red crossed 45 degrees");
    // Draw.setBitmapColor(rect1, "Roses");
    // Draw.setBitmapFileColor(rect1, "crazy_blue.jpg");

    // (x, y), width, height [size]
    XShape text1 = Draw.drawText(currSlide, "Hello LibreOffice", 120, 120, 60, 30, 24);
    Props.setProperty(text1, "Name", "text1");
    // Props.showProps("TextShape's Text Properties",  Draw.getTextProperties(text1));


    // gray transparent circle; uses (x,y), radius
    XShape circle2 = Draw.drawCircle(currSlide, 40, 150, 20);
    Props.setProperty(circle2, "FillColor", Lo.getColorInt(java.awt.Color.GRAY));
    Draw.setTransparency(circle2, 25);

    // thick line; uses (x,y), angle clockwise from x-axis, length
    XShape line2 = Draw.drawPolarLine(currSlide, 60, 200, 45, 100);
    Props.setProperty(line2, "LineWidth", 300);    // 3mm
  }  // end of drawShapes()



  private static void animShapes(XDrawPage currSlide)
  // two animations of a circle and a line
  /* the animation loop is: 
           redraw shape, delay, update shape position/size
  */
  {
    // reduce circle size and move to the right
    int xc = 40;
    int yc = 150;
    int radius = 40;
    XShape circle = null;
    for (int i=0; i < 20; i++) {  // move right
      if (circle != null)
        currSlide.remove(circle);
      circle = Draw.drawCircle(currSlide, xc, yc, radius);
                                 // (x-center, y-center, radius)
      Lo.delay(200);
      xc += 5;
      radius *= 0.95;
    }

    // rotate line counter-clockwise, and change length
    int x2 = 140;
    int y2 = 110;
    XShape line = null;
    // Props.setProperty(line, "Name", "animLine");
    // Draw.reportPosSize(line);
    for (int i=0; i <= 25; i++) {
      if (line != null)
        currSlide.remove(line);
      line = Draw.drawLine(currSlide, 40, 100, x2, y2);
      Lo.delay(200);
      x2 -= 4;
      y2 -= 4;
    }
  }  // end of animShapes()


} // end of DrawPicture class


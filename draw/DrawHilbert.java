
// DrawHilbert.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, July 2015

/* Generate a Hilbert curve of the specified level.
   Created using a series of rounded blue lines.
   Save as hilbert.png

   Position/size the window, resize the page view

   Usage:
     run DrawHilbert 4
   
   Using '6' takes about 2 minutes to fully draw. It's
   fun to try once.

   Using '7' causes the code to mis-calculate incr, so the 
   line drawing goes off the left side of the canvas. And it
   takes forever to do it!
*/

import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.frame.*;

import com.sun.star.text.*;
import com.sun.star.drawing.*;
import com.sun.star.awt.*;



public class DrawHilbert
{
  private static XDrawPage slide;

  private static int x, y, level, delay;
  private static int incr;


  public static void main (String args[])
  {
    if (args.length != 1) {
      System.out.println("Usage: java Hilbert <level>");
      return;
    }

    XComponentLoader loader = Lo.loadOffice();
    XComponent doc = Draw.createDrawDoc(loader);
    if (doc == null) {
      System.out.println("Draw doc creation failed");
      Lo.closeOffice();
      return;
    }

    Rectangle rect = GUI.getScreenSize();
    GUI.setPosSize(doc, 0, 0, rect.Width/2, rect.Height-40);
    GUI.setVisible(doc, true);
    Lo.delay(1000);            // need delay or zoom may not occur
    GUI.zoomValue(doc, 75);   // percentage size

    slide = Draw.getSlide(doc, 0);
    startHilbert(args[0], Draw.getSlideSize(slide));

    Lo.saveDoc(doc, "hilbert.png");

    System.out.println("Waiting a bit before closing...");
    Lo.delay(3000);
    Lo.closeDoc(doc);
    Lo.closeOffice();
  } // end of main()



  private static void startHilbert(String levelStr, Size slideSize)
  // init global params and then start Hilbert
  {
    level = 1;
    try {
      level = Integer.parseInt(levelStr);
    }
    catch(NumberFormatException e) {
      System.out.println("Level is not an integer; using 1");
    }
    if (level < 1) {
      System.out.println("Level must be >= 1; using 1");
      level = 1;
    }

    // store smallest mm dimension
    int sqWidth = (slideSize.Width > slideSize.Height) ? slideSize.Height :
                                                         slideSize.Width;
    x = sqWidth - 10;
    y = 10;
    incr = (int) Math.round((sqWidth - 20)/((Math.pow(2, level)) - 1));
    if (incr == 0) {
      System.out.println("Hilbert increment is 0");
      System.exit(1);
    }
    else
      System.out.println("Line increment: " + incr);

    delay = 1000/(level*level);
    System.out.println("Delay: " + delay);

    a(level);
  } // end of startHilbert()



  private static void a(int i) 
  {
    if (i > 0) {
      d(i-1);  moveBy(-incr, 0);
      a(i-1);  moveBy(0, incr);
      a(i-1);  moveBy(incr, 0);
      b(i-1);
      Lo.delay(delay);
    }
  }  // end of a()



  private static void b(int i) 
  {
    if (i > 0) {
      c(i-1);  moveBy(0, -incr);
      b(i-1);  moveBy(incr, 0);
      b(i-1);  moveBy(0, incr);
      a(i-1);
      Lo.delay(delay);
    }
  }  // end of b()


  private static void c(int i) 
  {
    if (i > 0) {
      b(i-1);  moveBy(incr, 0);
      c(i-1);  moveBy(0, -incr);
      c(i-1);  moveBy(-incr, 0);
      d(i-1);
      Lo.delay(delay);
    }
  }  // end of c()


  private static void d(int i) 
  {
    if (i > 0) {
      a(i-1);  moveBy(0, incr);
      d(i-1);  moveBy(-incr, 0);
      d(i-1);  moveBy(0, -incr);
      c(i-1);
      Lo.delay(delay);
    }
  }  // end of d()


  private static void moveBy(int xStep, int yStep) 
  {
    int xOld = x;
    int yOld = y;
    x += xStep;
    y += yStep;

    XShape line = Draw.drawLine(slide, xOld, yOld, x, y);
    Props.setProperty(line, "LineColor", 0x0000FF);  // blue
    Props.setProperty(line, "LineWidth", 110);        // in 1/100 mm units
    Props.setProperty(line, "LineCap", LineCap.ROUND);  // round-off the line
  }  // end of moveBy()




} // end of DrawHilbert class



// BezierBuilder.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, July 2015

/* Read Bezier points from a text file and build an open or closed 
   Bezier shape, which is displayed in a Draw page.

   The Bezier points originally come from a Draw file of a simple Bezier shape,
   exported in SVG format. The file was openned in a text editor, and the
   OpenBezierShape data extracted. An example fragment of a SVG file:

   <g class="com.sun.star.drawing.OpenBezierShape">
         <g id="id3">
          <path fill="none" stroke="rgb(0,0,0)" d="M 4500,15100 
              C 5400,14100 3600,3500 10000,7200 16400,10900 13200,4300 13200,4300 Z"/>
         </g>
    </g>

   The extracted data is the "M", "C", and optional "Z" data, which are saved as
   separate lines in a text file. An example text file:

     M 4500,15100 
     C 5400,14100 3600,3500 10000,7200 16400,10900 13200,4300 13200,4300 
     Z

   This data defines a simple SVG path. For full details, see
   the SVG path specification at:
      http://www.w3.org/TR/SVG/paths.html

   It is this kind of text file that is read in by this program.

   Usage:
      > run BezierBuilder bpts0.txt   
                  // also bpts1.txt and bpts2.txt
*/

import java.io.*;
import java.util.*;

import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.frame.*;

import com.sun.star.text.*;
import com.sun.star.drawing.*;
import com.sun.star.container.*;

import com.sun.star.util.*;
import com.sun.star.awt.*;



public class BezierBuilder
{

  public static void main (String args[])
  {
    if (args.length != 1) {
      System.out.println("Usage: BezierBuilder <bezier points fnm>");
      return;
    }

    XComponentLoader loader = Lo.loadOffice();

    // create Impress page or Draw slide
    XComponent doc = Draw.createDrawDoc(loader);
    if (doc == null) {
      System.out.println("Doc creation failed");
      Lo.closeOffice();
      return;
    }

    XDrawPage currSlide = Draw.getSlide(doc, 0);

    GUI.setVisible(doc, true);

    drawCurve(currSlide);

    Point startPt = new Point();
    ArrayList<Point> curvePts = new ArrayList<Point>();
    boolean isOpen = readPoints(args[0], startPt, curvePts);
    XShape bezierShape = createBezier(currSlide, startPt, curvePts, isOpen);

    Lo.saveDoc(doc, "bezier.odg");


    System.out.println("Waiting a bit before closing...");
    Lo.delay(3000);
    Lo.closeDoc(doc);
    Lo.closeOffice();
  } // end of main()




  private static XShape drawCurve(XDrawPage currSlide)
  {
    Point[] pathPts = new Point[4];
    PolygonFlags[] pathFlags = new PolygonFlags[4];

    pathPts[0] = new Point(1000, 2500);
    pathFlags[0] = PolygonFlags.NORMAL;
    
    pathPts[1] = new Point(1000,1000);    // control point
    pathFlags[1] = PolygonFlags.CONTROL;
    
    pathPts[2] = new Point(4000,1000);    // control point
    pathFlags[2] = PolygonFlags.CONTROL;
    
    pathPts[3] = new Point(4000,2500);
    pathFlags[3] = PolygonFlags.NORMAL;


    return Draw.drawBezier(currSlide, pathPts, pathFlags, true);
  }  // end of drawCurve()




  private static boolean readPoints(String fnm, Point startPt, ArrayList<Point> curvePts)
  // Read in a text file of bezier points
  { 
    boolean isOpen = true;
    try {
      BufferedReader br = new BufferedReader( new FileReader(fnm));
      String line;
      while((line = br.readLine()) != null) {
        if (line.length() == 0)  // blank line
          continue;
        if (line.startsWith("//"))   // comment
          continue;
        char ch = line.charAt(0);
        if (ch == 'M') {
          System.out.println("Reading start point");
          Point pt = extractPoint(line.substring(1).trim());
          if (pt != null) {
            startPt.X = pt.X;
            startPt.Y = pt.Y;
            Draw.printPoint(startPt);
          }
        }
        else if (ch == 'C')  {
          System.out.println("Reading curve points");
          setCurve(curvePts, line.substring(1).trim());
        }
        else if (ch == 'Z')  {
          System.out.println("Read closedpath flag");
          isOpen = false;
        }
        else
          System.out.println("Do not recognize: " + ch);
      }
      br.close();
      System.out.println("Read in points file: " + fnm);
    } 
    catch (IOException e) 
    { System.out.println("Error reading points file: " + fnm); }

    return isOpen;
  }  // end of readPoints()



  private static Point extractPoint(String ptStr)
  // convert a string of the form "16400,10900" into a Point
  {
    Point startPt = new Point();
    String[] vals = ptStr.split(",");
    if (vals.length != 2)
      System.out.println("Could not parse the point string into 2 parts");
    else {
      startPt.X = getInteger(vals[0]);
      startPt.Y = getInteger(vals[1]);
    }
    return startPt;
  }  // end of extractPoint()



  private static int getInteger(String s)
  {
    int val = 0;
    try {
      val = Integer.parseInt(s);
    }
    catch(NumberFormatException e)
    {  System.out.println("Value is not a integer: " + s);  }
    return val;
  }  // end of getInteger()



  private static void setCurve(ArrayList<Point> curvePts, String ptsStr)
  // convert a string of the form "5400,14100 3600,3500 10000,7200" into points
  {
    String[] vals = ptsStr.split("\\s+");
    for (int i=0; i < vals.length; i++)
      curvePts.add( extractPoint(vals[i]) );
  }  // end of setCurve()




  private static XShape createBezier(XDrawPage currSlide, 
                              Point startPt, ArrayList<Point> curvePts, boolean isOpen)
  /* The assumption is that the curve points use the C format from the SVG path
     specification. Namely (x1 y1 x2 y2 x y)+,	where each tuple defines a cubic
     Bézier curve. It starts at the startPt and ends at (x,y), using (x1,y1) as the 
     control point at the beginning of the curve and (x2,y2) as the control point 
     at the end of the curve. 

     The (x,y) becomes the new startPt for the next cubic curve.
  */
  {
    ArrayList<Point> bezPts = new ArrayList<Point>();
    ArrayList<PolygonFlags> flags = new ArrayList<PolygonFlags>();

    if (curvePts.size() % 3 != 0) {
      System.out.println("Number of points must be a multiple of 3");
      return null;
    }

    int pathStep = curvePts.size()/3;
    for (int i = 0; i < pathStep; i++) {
      // fill in the points and flags for 1 step
      bezPts.add(startPt);
      flags.add(PolygonFlags.NORMAL);
      
      bezPts.add( curvePts.get(i*3) );    // (x1,y1) control point
      flags.add(PolygonFlags.CONTROL);
      
      bezPts.add( curvePts.get(i*3+1) );  // (x2,y2) control point
      flags.add(PolygonFlags.CONTROL);
      
      bezPts.add( curvePts.get(i*3+2) );  // (x,y) end point
      flags.add(PolygonFlags.NORMAL);
      startPt = curvePts.get(i*3+2);      // and the next start point
    }

    int numBezPts = bezPts.size();
    Point[] pathPts = new Point[numBezPts];
    PolygonFlags[] pathFlags = new PolygonFlags[numBezPts];
    for (int i=0; i < numBezPts; i++) {
      pathPts[i] = bezPts.get(i);
      pathFlags[i] = flags.get(i);
    }

    return Draw.drawBezier(currSlide, pathPts, pathFlags, isOpen);
  }  // end of createBezier()



} // end of BezierBuilder class


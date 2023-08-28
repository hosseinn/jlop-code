
// CopySlide.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, July 2015

/* Copy the slide from the first index position to the
   position after the second index, and save the changes.

   Usage:
      run CopySlide points.odp 1 4
*/

import java.util.*;
import java.awt.Point;
import java.awt.Rectangle;

import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.frame.*;

import com.sun.star.text.*;
import com.sun.star.drawing.*;




public class CopySlide
{


  public static void main(String args[])
  {
    if (args.length != 3) {
      System.out.println("Usage: CopySlide fnm from-index to-index");
      return;
    }

    XComponentLoader loader = Lo.loadOffice();
    XComponent doc = Lo.openDoc(args[0], loader);
    if (doc == null) {
      System.out.println("Could not open the file: " + args[0]);
      Lo.closeOffice();
      return;
    }

    int fromIdx = Lo.parseInt(args[1]);
    int toIdx = Lo.parseInt(args[2]);
    int numSlides = Draw.getSlidesCount(doc);
    if ((fromIdx < 0) || (toIdx < 0) || 
        (fromIdx >= numSlides) || (toIdx >= numSlides)) {
      System.out.println("One or both indicies are out of range");
      Lo.closeOffice();
      return;
    }

    GUI.setVisible(doc, true);

    copyTo(doc, fromIdx, toIdx);
    // Draw.deleteSlide(doc, fromIdx);   
        // a problem if the pasting changes the indicies

    // Draw.duplicate(doc, 1);
            // places copy after original

    Lo.save(doc);   // overwrite
    Lo.closeDoc(doc);
    Lo.closeOffice();
  } // end of main()




  private static void copyTo(XComponent doc, int fromIdx, int toIdx)
  /* Copy fromIdx slide to the clipboard in slide-sorter mode,
     then paste it to after the toIdx slide. 
  */
  {
    XController ctrl = GUI.getCurrentController(doc);

    Lo.dispatchCmd("DiaMode");
        // Switch to slide sorter view so that slides can be pasted
    Lo.delay(5000);   // give Office plenty of time to do it

    XDrawPage fromSlide = Draw.getSlide(doc, fromIdx);
    XDrawPage toSlide = Draw.getSlide(doc, toIdx);

    Draw.gotoPage(ctrl, fromSlide);  // select this slide
    Lo.dispatchCmd("Copy");
    System.out.println("Copied " + fromIdx);

    Draw.gotoPage(ctrl, toSlide);   
            // select this slide; 'doc' version of gotoPage() pastes incorrectly
    Lo.dispatchCmd("Paste");
    System.out.println("Pasted to after " + toIdx);

    // Lo.dispatchCmd("PageMode");  // back to normal mode (not working)
    Lo.dispatchCmd("DrawingMode"); 
  }  // end of copyTo()


}  // end of CopySlide class

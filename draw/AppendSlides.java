
// AppendSlides.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, July 2015

/* Copy the slides in the second, third, fourth, etc files
   onto the end of the slides in the first file.

   Open the first file as a document. Change to slide sorter
   view so that slides can be pasted into it.

   Open each of the next files in turn. Change to slide sorter
   view so that slides can be copied from it.

   JNA is needed to press a Yes button in the paste dialog.
   This dialog only appears if the slides being appended have
   a different master page than the destination slides.

   This is slow (and hacky), but it means the copies slides 
   retain their original formatting.
 
   Usage:
      run AppendSlides points.ppt algs.ppt
         -->  points_Append.ppt
      run AppendSlides points.odp algsSmall.ppt    // quicker!
*/

import java.util.*;
import java.awt.Point;
import java.awt.Rectangle;

import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.frame.*;

import com.sun.star.text.*;
import com.sun.star.drawing.*;

import com.sun.jna.platform.win32.WinDef.HWND;



public class AppendSlides
{
  // controller and frame for document being appended to
  private static XController toCtrl;
  private static XFrame toFrame;


  public static void main(String args[])
  {
    if (args.length < 2) {
      System.out.println("Usage: AppendSlides fnm1 fnm2 ...");
      return;
    }

    XComponentLoader loader = Lo.loadOffice();
    XComponent doc = Lo.openDoc(args[0], loader);
    if (doc == null) {
      System.out.println("Could not open the first file: " + args[0]);
      Lo.closeOffice();
      return;
    }

    GUI.setVisible(doc, true);

    toCtrl = GUI.getCurrentController(doc);
    toFrame = GUI.getFrame(doc);

    Lo.dispatchCmd(toFrame, "DiaMode", null);
        // Switch to slide sorter view so that slides can be pasted

    XDrawPages toSlides = Draw.getSlides(doc);

    // process the other files on the command line
    for(int i=1; i < args.length; i++) {  // start at 1
      XComponent appDoc = Lo.openDoc(args[i], loader);
      if (appDoc == null)
        System.out.println("Could not open the file: " + args[i]);
      else
        appendDoc(toSlides, appDoc);
    }

    Lo.saveDoc(doc, Info.getName(args[0]) + "_Append." + Info.getExt(args[0]));

    Lo.dispatchCmd(toFrame, "PageMode", null);  // back to normal mode (does not work)
    Lo.closeDoc(doc);
    Lo.closeOffice();
  } // end of main()




  private static void appendDoc(XDrawPages toSlides, XComponent doc)
  /* Append doc to the end of  toSlides.
     Access the slides in the document, and the document's controller and frame refs.
     Switch to slide sorter view so that slides can be copied.
  */
  {
    GUI.setVisible(doc, true);

    XController fromCtrl = GUI.getCurrentController(doc);
    XFrame fromFrame = GUI.getFrame(doc);
    Lo.dispatchCmd(fromFrame, "DiaMode", null);   // slide sorter view

    XDrawPages fromSlides = Draw.getSlides(doc);
    if (fromSlides == null)
      System.out.println("- No slides found");
    else {
      System.out.println("- Adding slides");
      appendSlides(toSlides, fromSlides, fromCtrl, fromFrame);
    }

    Lo.dispatchCmd(fromFrame, "PageMode", null);  
            // change back to norrmal slide view (does not work)
    Lo.closeDoc(doc);
    System.out.println();
  }  // end of appendDoc()




  private static void appendSlides(XDrawPages toSlides, XDrawPages fromSlides,
                          XController fromCtrl, XFrame fromFrame)
  /* Append fromSlides to the end of toSlides
     Loop through the fromSlides, copying each one.
  */
  {
    for (int i=0; i < fromSlides.getCount(); i++) {
      XDrawPage fromSlide = Draw.getSlide(fromSlides, i);
      XDrawPage toSlide = Draw.getSlide(toSlides, toSlides.getCount()-1);
           // the copy will be placed *after* this slide

      copyTo(fromSlide, fromCtrl, fromFrame, toSlide, toCtrl, toFrame);
    }
  }  // end of appendSlides()





  private static void copyTo(XDrawPage fromSlide,
                             XController fromCtrl, XFrame fromFrame,
                             XDrawPage toSlide,
                             XController toCtrl, XFrame toFrame)
  /* Copy fromSlide to the clipboard, and
     then paste it to after the toSlide. Unfortunately, the
     paste requires a "Yes" button to be pressed, and so JNA is needed
     to access the dialog.
  */
  {
    Draw.gotoPage(fromCtrl, fromSlide);  // select this slide
    System.out.print("-- Copy --> ");
    Lo.dispatchCmd(fromFrame, "Copy", null);
    Lo.delay(1000); 

    Draw.gotoPage(toCtrl, toSlide);   // select this slide
    System.out.println("Paste");

    // wait for "Adaption" dialog and click it
    Thread monitorThread = new Thread() {
      public void run() {
        Lo.delay(500); 
        clickWindow("LibreOffice");  // full title is "LibreOffice 4...."
      }
    };
    monitorThread.start();

    Lo.dispatchCmd(toFrame, "Paste", null);
  }  // end of copyTo()




  private static void clickWindow(String windowTitle)
  {
    HWND handle = JNAUtils.findTitledWin(windowTitle);
    if (handle == null)
      return;

    // System.out.println("Found \"" + windowTitle + "\"");
    Rectangle bounds = JNAUtils.getBounds(handle);
    // System.out.println("Button bounds: " + bounds);

    int xCenter = bounds.x + 64;     
           // hard-wired location for "Yes" (using MS Paint on screenshot)
    int yCenter = bounds.y + 91;
    JNAUtils.doClick( new Point(xCenter, yCenter));

  }  // end of clickWindow()

}  // end of AppendSlides class

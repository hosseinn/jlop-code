
// CopySlide.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, July 2016

/* Copy a slide in two possible ways:

    * copySave() copies a specified slide to the clipboard

    * copyTo() copies a specified slide to another index
      position in the slide deck
        - this is a simplified version of CopySlide.java 
          in Draw Tests/

    Note: slide indicies start at 0 for the first slide.

   Office crashes at the end, but there are no zombie
   processes left over, so the error can be ignored.
*/

import java.util.*;

import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.frame.*;
import com.sun.star.view.*;

import com.sun.star.text.*;
import com.sun.star.drawing.*;




public class CopySlide
{
  private static final String FNM = "algs.odp";


  public static void main(String args[])
  {
    XComponentLoader loader = Lo.loadOffice();
    XComponent doc = Lo.openDoc(FNM, loader);
    if (doc == null) {
      System.out.println("Could not open the file: " + FNM);
      Lo.closeOffice();
      return;
    }

    GUI.setVisible(doc, true);   
    Lo.wait(2000);

    copySave(doc, 3);   // save copy of 4th slide 

    // copyTo(doc, 0, 4);
      // copy 1st slide; insert after 5th slide

    Lo.waitEnter();
    Lo.closeDoc(doc);
    Lo.closeOffice();
  } // end of main()



  private static void copySave(XComponent doc, int fromIdx)
  /* Copy fromIdx slide to the clipboard in slide-sorter mode.
  */
  {
    XController ctrl = GUI.getCurrentController(doc);

    Lo.dispatchCmd("DiaMode");
        // Switch to slide sorter view so that slide copies on clipboard
        // will have embedded, xml, and image versions
    Lo.delay(5000);   // give Office plenty of time to do it

    XDrawPage fromSlide = Draw.getSlide(doc, fromIdx);
    Draw.gotoPage(ctrl, fromSlide);  // select this slide

    Lo.dispatchCmd("Copy");
    System.out.println("Copied slide no. " + (fromIdx+1));

    Clip.listFlavors();

    // save embedded slide as ODP; 
    // double click to open in Office
    FileIO.saveBytes("slide" + (fromIdx+1) +".odp", Clip.getEmbedSource());

    // save slide as XML ('flat' ODP)
    FileIO.saveString("slide" + (fromIdx+1) +".fodp", Clip.getXMLDraw());

    // save slide as PNG image
    Images.saveImage(Clip.getImage(), "slide" + (fromIdx+1) +".png");
          // for another approach see Slide2Image.java in Draw Tests/

    // save slide as a Bitmap 
    FileIO.saveBytes("slide" + (fromIdx+1) +".bmp", Clip.getBitmap());

    Lo.dispatchCmd("DrawingMode"); 
  }  // end of copySave()





  private static void copyTo(XComponent doc, int fromIdx, int toIdx)
  /* Copy fromIdx slide to the clipboard in slide-sorter mode,
     then paste it to after the toIdx slide. Start counting slides
     from 0.
     Same as CopySlide.java in Draw Tests/
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
    System.out.println("Copied slide no. " + (fromIdx+1));

    Draw.gotoPage(ctrl, toSlide);   
            // select this slide; 'doc' version of gotoPage() pastes incorrectly
    Lo.dispatchCmd("Paste");
    System.out.println("Pasted to after slide no. " + (toIdx+1));

    Lo.dispatchCmd("DrawingMode"); 
  }  // end of copyTo()


}  // end of CopySlide class

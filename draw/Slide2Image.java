
// Slide2Image.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, July 2015

/* Save a given page of a slide presentation (e.g. ppt, odp)
   as an image file (e.g. "gif", "png", "jpeg", "wmf", "bmp", "svg")

   Usage
     run Slide2Image algs.ppt 2 png
             --> creates the image file algs2.png

   Since images are represented as a single slide in Office, you
   can also convert an image into another format using this code.

   e.g.
     run Slide2Image cover.png 0 jpeg
            --> creates the image file cover0.jpeg

   Note that the page index starts at 0.
*/


import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.frame.*;

import com.sun.star.drawing.*;



public class Slide2Image
{

  public static void main(String args[])
  {
    if (args.length != 3) {
      System.out.println("Usage: Slide2Image fnm index imageFormat");
      return;
    }
    String fnm = args[0];
    int idx = Lo.parseInt(args[1]);
    String imFormat = args[2];

    XComponentLoader loader = Lo.loadOffice();
    XComponent doc = Lo.openDoc(fnm, loader);
    if (doc == null) {
      System.out.println("Could not open " + fnm);
      Lo.closeOffice();
      return;
    }

    XDrawPage slide = Draw.getSlide(doc, idx);
    if (slide == null) {
      Lo.closeOffice();
      return;
    }

    String[] names = Images.getMimeTypes();
    System.out.println("Known GraphicExportFilter mime types:");
    for (int i=0; i < names.length; i++) 
      System.out.println("  " + names[i]); 

    String outFnm = Info.getName(fnm) + idx + "." + imFormat;
    System.out.println("Saving page " + idx + " to \"" + outFnm + "\"");
    String mime = Images.changeToMime(imFormat);
    Draw.savePage(slide, outFnm, mime);
    Lo.closeDoc(doc);
    Lo.closeOffice();
  }  // end of main()


}  // end of Slide2Image class

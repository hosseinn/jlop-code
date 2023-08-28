
// ModifySlides.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, July 2015

/* Add two new slides to the input document:
    -- add a title-only slide with a graphic at the end
    -- add a title/subtitle slide at the start;
       XDrawPage.insertNewByIndex() cannot add a slide at the front
        of the slide deck, so this slide is inserted at index 1

   Usage:
      run ModifySlides algsSmall.ppt
         -->  algsSmall_Mod.ppt
      run ModifySlides algsMedium.ppt
*/


import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.frame.*;
import com.sun.star.container.*;

import com.sun.star.drawing.*;



public class ModifySlides
{

  public static void main(String args[])
  {
    if (args.length != 1) {
      System.out.println("Usage: ModifySlides fnm");
      return;
    }

    XComponentLoader loader = Lo.loadOffice();
    XComponent doc = Lo.openDoc(args[0], loader);

    if (!Info.isDocType(doc, Lo.IMPRESS_SERVICE)) {
      System.out.println("-- not a slides presentation");
      Lo.closeOffice();
      return;
    }

    // Lo.setVisible(doc, true);

    XDrawPages slides = Draw.getSlides(doc);
    int numSlides = slides.getCount();
    System.out.println("No. of slides: " + numSlides);

    // add a title-only slide with a graphic at the end
    //XDrawPage lastPage = Draw.insertSlide(doc, numSlides);
    XDrawPage lastPage = slides.insertNewByIndex(numSlides);
    Draw.titleOnlySlide(lastPage, "Any Questions?");
    Draw.drawImage(lastPage, "questions.png");

    // add a title/subtitle slide at the start
    XDrawPage firstPage = slides.insertNewByIndex(0);    // is added after first slide!
    Draw.titleSlide(firstPage, "Interesting Slides", "Andrew Davison");

    Lo.saveDoc(doc, Info.getName(args[0]) + "_Mod." + Info.getExt(args[0]));

    Lo.closeDoc(doc);
    Lo.closeOffice();
  } // end of main()


}  // end of ModifySlides class

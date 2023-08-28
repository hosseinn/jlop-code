
// AutoShow.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, July 2015

/* Once the slide show is started, the slides automatically 
   transition to the next one.

   The program detects when the last slide is reached, will sleep
   for a short time, and then close and exit the  presentation.

   To make the slide show repeat until the user exits from it,
   uncomment the two property sets, and replace Draw.waitLast()
   by Draw.waitEnded().

   Usage:
      run AutoShow algs.odp
*/

import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.frame.*;

import com.sun.star.drawing.*;
import com.sun.star.container.*;
import com.sun.star.presentation.*;



public class AutoShow
{
  public static void main(String args[])
  {
    XComponentLoader loader = Lo.loadOffice();
    XComponent doc = Lo.openDoc(args[0], loader);
    if (doc == null) {
      System.out.println("Could not open "+ args[0]);
      Lo.closeOffice();
      return;
    }

    GUI.setVisible(doc, true);
       // slideshow start() crashes if the doc is not visible

    // set up a fast automatic change between all the slides
    XDrawPage[] slides = Draw.getSlidesArr(doc);
    for(XDrawPage slide : slides)
      Draw.setTransition(slide, FadeEffect.NONE, AnimationSpeed.FAST, 
                                       Draw.AUTO_CHANGE, 1);

    XPresentation2 show = Draw.getShow(doc);
    Props.showObjProps("Slide show", show);
    //Props.setProperty(show, "IsEndless", true);
    //Props.setProperty(show, "Pause", 0);   // no pause before repeating
    show.start();

    XSlideShowController sc = Draw.getShowController(show);

    Draw.waitLast(sc, 3000);
    //Draw.waitEnded(sc);
    System.out.println("Ending the slide show");
    sc.deactivate();
    show.end();

    Lo.closeDoc(doc);
    Lo.closeOffice();
  }  // end of main()


}  // end of AutoShow class

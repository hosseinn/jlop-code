
// BasicShow.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, July 2015

/* Start a slide show that the user must advance by clicking,
   pressing the right arrow key, or using the space bar. Other
   useful keys are 'b' for making toggling the current slide
   to black, and the left arrow key for moving back to the
   previous slide.

   When the user clicks on the black
   "click to exit" slide, the program will detect the slide show's
   finish, and then close and exit the presentation.

   Usage:
      run BasicShow algs.odp
*/

import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.frame.*;

import com.sun.star.presentation.*;



public class BasicShow
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

    XPresentation2 show = Draw.getShow(doc);
    Props.showObjProps("Slide show", show);

/*
    Props.setProperty(show, "IsAutomatic", true);
    Props.setProperty(show, "IsTransitionOnClick", false);
    Props.setProperty(show, "StartWithNavigator", false);
    Props.setProperty(show, "IsAlwaysOnTop", true);
    Props.setProperty(show, "IsEndless", true);
    Props.setProperty(show, "Pause", 0);
*/
    // PropertyValue[] props = Props.makeProps("IsAutomatic", true);
            // these properties override the ones set in Presentation
    // show.startWithArguments(props);

    show.start();
    XSlideShowController sc = Draw.getShowController(show);
    Draw.waitEnded(sc);
    
    Lo.closeDoc(doc);
    Lo.closeOffice();
  }  // end of main()


}  // end of BasicShow class

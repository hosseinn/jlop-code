
// CustomShow.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, July 2015

/* Build a custom show using an array of slide indicies, and then
   start the show. The user must click to move between slides.

   When the user clicks on the black
   "click to exit" slide, the program will detect the slide show's
   finish, and then close and exit the presentation.

   Usage:
      run CustomShow algs.odp
*/


import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.frame.*;

import com.sun.star.container.*;
import com.sun.star.presentation.*;



public class CustomShow
{
  public static void main(String args[])
  {
    XComponentLoader loader = Lo.loadOffice();
    XComponent doc = Lo.openDoc("algs.odp", loader);
    if (doc == null) {
      System.out.println("Could not open algs.odp");
      Lo.closeOffice();
      return;
    }

    GUI.setVisible(doc, true);
       // slideshow start() crashes if the doc is not visible

    int[] slideIdxs = new int[] {5, 6, 7, 8, 56};   // 56 is too big
    XNameContainer playList = Draw.buildPlayList(doc, slideIdxs, "ShortPlay");

    XPresentation2 show = Draw.getShow(doc);
    Props.setProperty(show, "CustomShow", "ShortPlay");
    Props.showObjProps("Slide show", show);

    show.start();
    XSlideShowController sc = Draw.getShowController(show);
    Draw.waitEnded(sc);

    Lo.closeDoc(doc);
    Lo.closeOffice();
  }  // end of main()



}  // end of CustomShow class

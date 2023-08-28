
// GalleryInfo.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, July 2015

/*  Demonstrates the use of the gallery API.
    Note that there's no need to open a document.

    There's another example in AnimBicycle.java, for finding
    the blue bike.

    The gallery is located in <OFFICE>/share/gallery/
*/

import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.frame.*;

import com.sun.star.gallery.*;



public class GalleryInfo
{

  public static void main(String args[])
  {
    XComponentLoader loader = Lo.loadOffice();

    // list all the gallery themes (i.e. the sub-directories below gallery/)
    Gallery.reportGallerys();
    System.out.println();

    // list all the items for the Transportation theme
    // i.e. list all the files in gallery/transportation
    Gallery.reportGalleryItems("Transportation");
    System.out.println();

    // find an item that has "bicycle" as part of its name in the Transportation theme
    XGalleryItem gItem = Gallery.findGalleryItem("Transportation", "bicycle");
    System.out.println();

    // print out the item's properties
    Gallery.reportGalleryItem(gItem);

    Lo.closeOffice();
  } // end of main()



}  // end of GalleryInfo class

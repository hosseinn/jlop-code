
// MoveSlide.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, Jan 2017

/* Move first slide to be the last one.
   Save the modified presentation as a new document.
*/


import org.odftoolkit.simple.presentation.*;
import org.odftoolkit.simple.PresentationDocument;


public class MoveSlide 
{
  public static void main(String[] args) 
  {
    try {
      PresentationDocument doc = PresentationDocument.loadDocument("algs.odp");
      int numSlides = doc.getSlideCount();
      System.out.println("Moving first slide to the end");
      doc.moveSlide(0, numSlides);   // why not numSlides-1?

      System.out.println("Saving document to algsMoved.odp");
      doc.save("algsMoved.odp");
      doc.close();
    } 
    catch (Exception e) 
    {  System.out.println(e); }
  }  // end of main()


}  // end of MoveSlide class

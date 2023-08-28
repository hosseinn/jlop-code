
// CombineDecks.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, Jan. 2017

/* Combine two slide decks together. 
   Save the resulting document in a new file.

   Note: the call to PresentationDocument.appendPresentation() 
         raises a NullPointerException when it calls getCopyStyleList(), 
         but the combined deck is still created
*/


import org.odftoolkit.simple.*;
import org.odftoolkit.simple.presentation.*;


public class CombineDecks 
{
  public static void main(String[] args) 
  {
    try {
      PresentationDocument doc1 = PresentationDocument.loadDocument("deck1.odp");
      PresentationDocument doc2 = PresentationDocument.loadDocument("deck2.odp");
      doc1.appendPresentation(doc2);
      System.out.println("Saving combination to combined.odp");
      doc1.save("combined.odp");
      doc1.close();
      doc2.close();
    } 
    catch (Exception e) 
    {  System.out.println(e); }
  }  // end of main()

}  // end of CombineDecks class

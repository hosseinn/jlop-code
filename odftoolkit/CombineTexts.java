
// CombineTexts.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, Jan. 2017

/* Combine two text documents together. Put the second after
   a page break, and save the resulting document in a new file.
*/

import org.odftoolkit.simple.*;
import org.odftoolkit.simple.text.*;


public class CombineTexts 
{
  public static void main(String[] args) 
  {
    try {
      TextDocument doc1 = TextDocument.loadDocument("doc1.odt");
      TextDocument doc2 = TextDocument.loadDocument("doc2.odt");

      doc1.addPageBreak();
      Paragraph lastPara = doc1.getParagraphByReverseIndex(0, false);

      // insert contents at end and copy styles
      doc1.insertContentFromDocumentAfter(doc2, lastPara, true);

      System.out.println("Saving combination to combined.odt");
      doc1.save("combined.odt");
      doc1.close();
      doc2.close();
    } 
    catch (Exception e) 
    {  System.out.println(e); }
  }  // end of main()

}  // end of CombineTexts class

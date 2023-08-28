
// MakeTextDoc.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, Jan. 2017

/* Make a new text document containing an image, some text,
   a list, and a table.

   Based on HelloWorld.java from:
     http://incubator.apache.org/odftoolkit/simple/gettingstartguide.html
*/

import java.net.*;

import org.odftoolkit.simple.*;
import org.odftoolkit.simple.table.*;
import org.odftoolkit.simple.text.list.List;



public class MakeTextDoc
{
  public static void main(String[] args)
  {
    try {
      TextDocument doc = TextDocument.newTextDocument();
      doc.newImage(new URI("odf-logo.png"));

      // add paragraphs and list
      doc.addParagraph("Hello World, Hello Simple ODF!");
      doc.addParagraph("The following is a list.");
      List list = doc.addList();
      String[] items = {"item1", "item2", "item3"};
      list.addItems(items);

      // add table
      Table table = doc.addTable(2, 2);
      Cell cell = table.getCellByPosition(0, 0);
      cell.setStringValue("Hello World!");

      System.out.println("Creating MakeTextDoc.odt");
      doc.save("MakeTextDoc.odt");
    }
    catch (Exception e) {
      System.out.println(e);
    }
  }  // end of main()

} // end of MakeTextDoc class

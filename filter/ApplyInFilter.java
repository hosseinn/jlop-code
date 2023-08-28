
// ApplyInFilter.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, Dec. 2016

/* Convert XML to an Office document in two steps:
     1) use the supplied XSLT to convert the XML
        into Flat XML understood by Office;

     2) Use the correct Flat XML import filter to load
        the flat XML data into Office

  Usage:
     run ApplyInFilter pay.xml payImport.xsl payment.ods
          -- pay.xml is imported into a spreadsheet

     run ApplyInFilter clubs.xml clubsImport.xsl clubs.odt
          -- clubs.xml is imported in a text doc
*/

import java.io.*;

import com.sun.star.lang.*;
import com.sun.star.frame.*;



public class ApplyInFilter
{

  public static void main(String[] args) 
  {
    if (args.length != 3) {
      System.out.println("Usage: java ApplyInFilter <XML fnm> <Flat XML import filter> <new ODF>");
      return;
    }

    // convert the data to Flat XML format
    String xmlStr = XML.applyXSLT(args[0], args[1]);
    if (xmlStr == null) {
      System.out.println("Filtering failed");
      return;
    }
    // String xmlStr1 = XML.indent2Str(xmlStr);
    // System.out.println( xmlStr1);

    // save flat XML data to temp file
    String flatFnm = FileIO.createTempFile("xml");
    FileIO.saveString(flatFnm, xmlStr);

    // open temp file using Office's correct Flat XML filter
    String odfFnm = args[2];
    XComponentLoader loader = Lo.loadOffice();

    String docType = Lo.ext2DocType( Info.getExt(odfFnm));
    System.out.println("Doc type: " + docType);
    XComponent doc = Lo.openFlatDoc(flatFnm, docType, loader);
    if (doc == null)
      System.out.println("Document creation failed");
    else {
      GUI.setVisible(doc, true);

      Lo.waitEnter();
      Lo.saveDoc(doc, "flat.xml");   // save Flat XML for debugging
      Lo.saveDoc(doc, odfFnm);
      Lo.closeDoc(doc);
    }

    Lo.closeOffice();
  }   // end of main()



}  // end of ApplyInFilter class

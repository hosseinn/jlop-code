
// ApplyOutFilter.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, Dec. 2016

/* Convert an Office document to simple XML in two steps:
     1) Use Office's built in Flat XML export filter to
        save the document as Flat XML data (in flat.xml)

     2) Use the supplied export filter to convert
        Flat XML into simple XML, which is saved to the
        specified file.

  Usage:
     run ApplyOutFilter payment.ods payExport.xsl payEx.xml

     run ApplyOutFilter clubs.odt clubsExport.xsl clubsEx.xml
        -- clubs.odt must be formatted using clubsTemplate.ott
           or the XSLT will fail
*/

import java.io.*;

import com.sun.star.lang.*;
import com.sun.star.frame.*;


public class ApplyOutFilter
{


  public static void main(String[] args) 
  {
    if (args.length != 3) {
     System.out.println("Usage: java ApplyOutFilter <XML file> <Flat XML export filter> <new XML file>");
      return;
    }

    XComponentLoader loader = Lo.loadOffice();
    XComponent doc = Lo.openDoc(args[0], loader);
    if (doc == null) {
      System.out.println("Could not open document: " + args[0]);
      Lo.closeOffice();
      return;
    }

    // save flat XML data
    String flatFnm = FileIO.createTempFile("xml");
    Lo.saveDoc(doc, flatFnm);
    Lo.closeDoc(doc);


    // use XSLT to convert Flat XML into simple XML
    String filteredXML = XML.applyXSLT(flatFnm, args[1]);
    if (filteredXML == null)
      System.out.println("Filtering failed");
    else {
      String xmlStr = XML.indent2Str(filteredXML);
      System.out.println(xmlStr);
      FileIO.saveString(args[2], xmlStr);
    }

    Lo.closeOffice();
  }   // end of main()


}  // end of ApplyOutFilter class

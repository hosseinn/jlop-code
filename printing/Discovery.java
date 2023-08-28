
// Discovery.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, July 2016

/* Several examples of searching for print services with
   specific DocFlavor or printer attributes.

*/


import java.io.*;

import javax.print.*;
import javax.print.attribute.*;
import javax.print.attribute.standard.*;



public class Discovery
{
  public static void main(String[] args)
  {
    // find printers that supports a MIME file type
    DocFlavor flavor = DocFlavor.INPUT_STREAM.JPEG;
    // DocFlavor flavor = DocFlavor.INPUT_STREAM.PDF;

    PrintService[] psa =  PrintServiceLookup.lookupPrintServices(flavor, null);
    System.out.println("\nServices that support " + flavor);
    String[] pNames = JPrint.getPrinterNames(psa);
    if (pNames != null)
      for (String pName : pNames)
        System.out.println("  " + pName);


    // get a named print service
    PrintService ps = JPrint.findService("HP LaserJet 1200 Series PCL 6");
    if (ps != null)
      System.out.println("\nObtained HP LaserJet 1200 Series PCL 6");


    // print default printer name
    System.out.println("\nDefault printer: " + JPrint.getDefaultPrinterName());

    // find print services using a partial name fragment
    pNames = JPrint.findPrinterNames("HP");
    System.out.println("\nServices that use the name \"HP\"");
    if (pNames != null)
      for (String pName : pNames)
        System.out.println("  " + pName);



    // find HP print services that support color printing
    AttributeSet attrs = new HashAttributeSet();
       // see javax.print.attribute.Attribute for attributes
    System.out.println("\nHP Services that support color:");
    for (String pName : pNames) {
      attrs.clear();
      attrs.add(new PrinterName(pName, null));   // must be included
      attrs.add(ColorSupported.SUPPORTED);
      psa = PrintServiceLookup.lookupPrintServices(null, attrs);
      if (psa.length > 0) {
        System.out.println("  " + pName);
      }
    }

    System.out.println("\nServices that support A4 duplex and " + flavor);
    attrs.clear();    // reset
    // attrs.add(MediaSizeName.ISO_A4);
    attrs.add(Sides.DUPLEX);    // comment out this attribute to get some matches
    psa = PrintServiceLookup.lookupPrintServices(flavor, attrs);
    pNames = JPrint.getPrinterNames(psa);
    if (pNames != null)
      for (String pName : pNames)
         System.out.println("  " + pName);


    // use a dialog to choose a print service
    ps = JPrint.dialogSelect();
    if (ps != null) {
      System.out.println("\nYou selected " + ps.getName());
      JPrint.listService(ps, true);
    }
  }
}  // end of Discovery class

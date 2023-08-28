
// ListPrinters.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, July 2016

/* Print basic information about all the available printers; 
   fullinformation about the default printer; and a quick
   list of printer names.

   Uses support functions in the JPrint.java Utils/ file.
*/

import javax.print.*;


public class ListPrinters 
{

  public static void main(String [] args)
  {
    JPrint.listServices();         // basic info
    // JPrint.listServices(true);  // full info, but slow to generate

    PrintService ps = PrintServiceLookup.lookupDefaultPrintService();
    System.out.println("Default printer \"" + ps.getName() + "\":");
    JPrint.listService(ps, true);  // full info

    String[] pNames = JPrint.getPrinterNames();
    System.out.println("Printer names (" + pNames.length + "):");
    for(String pName : pNames)
      System.out.println("  " + pName);

  }  // end of main()


}  // end of ListPrinters class
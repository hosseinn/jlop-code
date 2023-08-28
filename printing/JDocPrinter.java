
// JDocPrinter.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, July 2016

/* Print a document to the specified printer, or use the
   default printer using only Java. Office is not employed.

   Usage:
      run JDocPrinter skinner.png "fineprint"
      run JDocPrinter skinner.png "1200"
      run JDocPrinter skinner.png

      run JDocPrinter python-outline.pdf "fineprint"
*/


import javax.print.*;



public class JDocPrinter 
{

  public static void main(String[] args)
  {
    if ((args.length < 1) || (args.length > 2)) {
      System.out.println("Usage: java JDocPrinter fnm [(partial)printer-name]");
      return;
    }

    String fnm = args[0];
    String pName = null;
    if (args.length == 2) { 
      String[] pNames = JPrint.findPrinterNames(args[1]);
      if (pNames != null) {
        System.out.println("Using first matching printer: \"" + pNames[0] + "\"");
        pName = pNames[0];
      }
    }

    if (pName != null)
      JPrint.print(pName, fnm);
    else {
      PrintService ps = JPrint.dialogSelect();
      if (ps != null)
        JPrint.printMonitorFile(ps, fnm);
        // JPrint.printFile(ps, fnm);

    }
  }  // end of main()

}  // end of JDocPrinter class

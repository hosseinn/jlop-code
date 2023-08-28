
// DocUnzip.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, Jan 2017

/*  Extract file from inside the specified
    document by unzipping.

    Usage:
      run DocUnzip algs.odp content.xml
      run DocUnzip algs.odp manifest.xml
      run DocUnzip algs.odp settings.xml
      run DocUnzip algs.odp mimetype
*/

import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.frame.*;

import com.sun.star.packages.zip.*;


public class DocUnzip
{

  public static void main(String args[])
  {
    if ((args.length < 1) || (args.length > 2)) {
      System.out.println("Usage: run DocUnzip <fnm> [<ExtractFnm>]");
      return;
    }

    XComponentLoader loader = Lo.loadOffice();

    FileIO.zipList(args[0]);
    // FileIO.zipListUno(args[0]);   // only names listed

    // get zip access to the document
    XZipFileAccess zfa = FileIO.zipAccess(args[0]);

    String mimeType = FileIO.getMimeType(zfa);
    System.out.println("MIME type: " + mimeType);
    System.out.println("Other MIME type approach: " + Info.getMIMEType(args[0]));

    int docType = Info.mimeDocType(mimeType);
    System.out.println("Doc Type: " + docType + "; " + Lo.docTypeStr(docType));


    if (args.length == 2)
      FileIO.unzipFile(zfa, args[1]);

    Lo.closeOffice();
  }  // end of main()



}  // end of DocUnzip class


JLOP Chapter 8. Printing

From the website:

  Java LibreOffice Programming
  http://fivedots.coe.psu.ac.th/~ad/jlop

  Dr. Andrew Davison
  Dept. of Computer Engineering
  Prince of Songkla University
  Hat yai, Songkhla 90112, Thailand
  E-mail: ad@fivedots.coe.psu.ac.th


If you use this code, please mention my name, and include a link
to the website.

Thanks,
  Andrew

============================

This directory contains 9 Java files:

 * ListPrinters.java, JDocPrinter.java
      // these only use Java's JPS, not Office

 * Discovery.java, ShowPrintProps.java, DocPrinter.java
   TextPrinter.java, ImpressPrinter.java, SheetPrinter.java,
   PrintSheet.java


There are 9 test files for the Java code:

  * algs.odp, Books.odb, classic_letter.ott,
    cover.odg, python-outline.pdf, skinner.png
    storyStart.doc, tables.ods, totals.ods


There are printing-related 5 batch files:

 * printersList.bat, printerStatus.bat,
   printerJobs.bat, printerClean.bat
       // these use Windows VBScripts

 * loPrint.bat
       // uses soffice.exe

  -- see the chapter for examples of how to use them


There are Office-related 7 batch files:
  * compile.bat
  * run.bat
     - make sure they refer to the correct locations for your
       Java, JNA, and my Utils classes.
       For details, read:
         "Installing the code for "Java LibreOffice Programming"
          http://fivedots.coe.psu.ac.th/~ad/jlop/install.html

    compile.bat and run.bat file both call:

  * lofind.bat
     - this tries to find LibreOffice (v4 or v5) on your machine,
       and checks its "bitness" with that of Java. They should
       both either be 32-bit or 64-bit. A warning is printed if 
       they are not.

  * loKill.bat        // kills the Office process
  * loList.bat        // lists the Office process (if it is running)

  * loDoc.bat         // searches the online LibreOffice documentation
  * loGuide.bat       // searches the online OpenOffice Dev Guide



----------------------------
Compilation

> compile *.java
    // you must have LibreOffice, Java, JNA, and my Utils classes installed

===================================
Execution:

// You must have LibreOffice, Java, JNA, and my Utils classes installed


> run ListPrinter


> run JDocPrinter fnm [(partial)printer-name]
e.g.
> run JDocPrinter skinner.png HP


> run Discovery


> run ShowPrintProps fnm
e.g.
> run ShowPrintProps classic_letter.ott


> run DocPrinter fnm [(partial)printer-name [no-of-pages]]
e.g.
> run DocPrinter skinner.png
> run DocPrinter storyStart.doc HP 3-4


> run TextPrinter fnm  [(partial)printer-name]
e.g.
> run TextPrinter classic_letter.ott HP


> run ImpressPrinter fnm  [(partial)printer-name]
e.g.
> run ImpressPrinter algs.odp HP


> run SheetPrinter fnm  [(partial)printer-name]
e.g.
> run SheetPrinter totals.ods HP


> run PrintSheet
 -- this prints "Sheet2" of totals.ods to the
    FinePrint printer

=================================
Last updated: 26th June 2016

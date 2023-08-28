
JLOP Chapter 17. Simple ODF

From the website:

  Java LibreOffice Programming
  http://fivedots.coe.psu.ac.th/~ad/jlop

  Dr. Andrew Davison
  Dept. of Computer Engineering
  Prince of Songkla University
  Hat Yai, Songkhla 90112, Thailand
  E-mail: ad@fivedots.coe.psu.ac.th


If you use this code, please mention my name, and include a link
to the website.

Thanks,
  Andrew

============================
Java files:

  * DocInfo.java, DocSecret.java, DocUnzip.java

ODFToolkit Examples:
  * MakeSheet.java, MakeSlides.java, MakeTextDoc.java
  * MoveSlide.java
  * CombineDecks.java, CombineSheets.java, CombineTexts.java

----------------------------
Support files for the examples: 

  *  algs.odp
     deck1.odp, deck2.odp
     doc1.odt, doc2.odt
     odf-logo.png, skinner.png
     ss1.ods, ss2.ods

----------------------------
Batch files for compiling/running ODFToolkit examples:

  * ocompile.bat
  * orun.bat
     - make sure they refer to the correct locations for
       ODFToolkit *and* the many other libraries it needs.

       For details, see README.txt in the "ODFToolkit Libs"
       folder.

----------------------------
7 Office-related batch files:

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


=================================
Compilation:

For the three "Doc" examples:
  * DocInfo.java, DocSecret.java, DocUnzip.java

> compile *.java
     - make sure compile.bat refers to the correct locations for your
       Java, JNA, and my Utils classes.

-----
For the ODFToolkit Examples:
  * MakeSheet.java, MakeSlides.java, MakeTextDoc.java
  * MoveSlide.java
  * CombineDecks.java, CombineSheets.java, CombineTexts.java

> ocompile *.java
     - make sure they refer to the correct locations for
       ODFToolkit *and* the many other libraries it needs.

=================================
Usage:

> run DocInfo algs.odp

> run DocUnzip algs.odp content.xml
          -- adds "Copy" to the extracted filename

> run DocSecret algs.odp "Made in Thailand"


ODFToolkit Examples:

> orun MakeTextDoc
         -- creates "MakeTextDoc.odt"

> orun MakeSheet
         -- creates "makeSheet.ods"

> orun MakeSlides
         -- creates "makeSlides.odp"

> orun MoveSlide
         -- rearranged slide deck is saved to "algsMoved.odp"

> orun CombineTexts
         -- "doc1.odt" and "doc2.odt" are saved to "combined.odt"

> orun CombineSheets
         -- "ss1.ods" and "ss2.ods" are saved to "combined.ods"

> orun CombineDecks
         -- "deck1.odp" and "deck2.odp" are saved to "combined.odp"
         -- note: the call to PresentationDocument.appendPresentation() 
            raises a NullPointerException when it calls getCopyStyleList(), 
            but the combined deck is still created
   
=================================
Last updated: 13th January 2017

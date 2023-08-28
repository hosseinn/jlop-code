
JLOP Chapter 10. Using the Clipboard

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

This directory contains 6 Java files:
 * CPTests.java, JCPTests.java
 * CopyPasteText.java, CopyPasteCalc.java,
   CopySlide.java, CopyResultSet.java



There are 6 test files for the Java code:
  * Addresses.ods, algs.odp, liangTables.odb,
    skinner.jpg, skinner.png, storyStart.doc



There are Office-related 7 batch files:
  * compile.bat
  * run.bat
     - make sure they refer to the correct locations for your
       Java, JNA, JavaMail, and my Utils classes.
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

It's a good idea to download a clipboard viewer.
I use ClCl:
  http://www.nakka.com/soft/clcl/index_eng.html


----------------------------
Compilation

> compile *.java
    // you must have LibreOffice, Java, JNA, and my Utils classes installed

===================================
Execution:

// You must have LibreOffice, Java, JNA, and my Utils classes installed

> run CPTests
> run JCPTests
    
> run CopyPasteText
> run CopyPasteCalc
     // uncomment the call to useClipUtils() or useDispatches()

> run CopySlide
> run CopyResultSet


=================================
Last updated: 21st July 2016

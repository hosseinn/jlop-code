
JLOP Chapter 11. Office as a GUI Component

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

This directory contains 4 Java files:
 * SwingViewer.java, OffViewer.java,
   TBViewer.java, MenuViewer.java


There are 5 test files for the Java code:
  * algs.odp, classic_letter.ott, skinner.png,
    tables.ods, liangTables.odb


One image file used by TBViewer.java and MenuViewer.java:
  * H.png              


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

/* Note: this version of run.bat sets the UNO_PATH environment variable
   used by OOoBean.
*/

> run SwingViewer <fnm>
e.g.
    run SwingViewer classic_letter.ott
    run SwingViewer algs.odp


> run OffViewer <fnm>
e.g.
    run OffViewer tables.ods
    run OffViewer liangTables.odb


> run TBViewer <fnm>
e.g.
   run TBViewer classic_letter.ott


> run MenuViewer <fnm>
e.g.
    run MenuViewer classic_letter.ott

/* This creates a new menu bar for the menu.
   Uncomment the code at the end of MenuViewer() to
   use the existing Office menu bar.
*/


=================================
Last updated: 29th August 2016

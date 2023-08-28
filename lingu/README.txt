
JLOP Chapter 2.1. The Linguistics API

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

This directory contains 2 Java files:
 * Lingo.java, LingoFile.java


There are 2 test files for the Java code:
  * badGrammar.odt, bigStory.doc


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

> run Lingo
    
> run LingoFile badGrammar.odt

> run LingoFile bigStory.doc > spellInfo.txt
          // redirect the output because it's so long


=================================
Last updated: 2nd August 2016

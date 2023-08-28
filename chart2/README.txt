
JLOP Chapter 4. Spreadsheet Processing

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

This directory contains 3 Java files:
  * Chart2Views.java,
    SlideChart.java, TextChart.java


A spreadsheet full of different tables:
  * chartsData.ods
      - this is used by all the Java examples


There are 7 batch files:
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

All the Java files will compile with the compile.bat batch file:

> compile *.java
    // you must have LibreOffice, Java, JNA, and my Utils classes installed


===================================
Execution:

// You must have LibreOffice, Java, JNA, and my Utils classes installed
// call them with run:

> run Chart2Views
    // Chart2Views.java contains many alternative chart creation function,
       but all but one of them is commented out. Edit the main() function
       to call the function you want to use.

> run SlideChart
    // creates a slide doc with a chart, called makeslide.odp
    // also saves the chart image to chartImage.png

> run TextChart
    // creates a text doc with a chart, called hello.odt


=================================
Last updated: 7th November 2015

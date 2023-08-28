
JLOP Chapter 3. Drawings and Slides Processing

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

This directory contains 20 Java files:
  * AnimBicycle.java, AnimationDemo.java, AppendSlides.java,
    AutoShow.java, BasicShow.java, BezierBuilder.java,
    CheckLinks.java, CopySlide.java, CustomShow.java,
    DrawHilbert.java, DrawPicture.java, GalleryInfo.java,
    Grouper.java, MakeSlides.java, MastersUse.java,
    ModifySlides.java, PointsBuilder.java, Slide2Image.java,
    SlideShow.java, SlidesInfo.java


There are 5 batch files:
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



Assorted test files for the Java code:
  * algs.odp, algs.ppt,
    algs2.png, algs2.svg,
    algsMedium.ppt, algsSmall.ppt,
    bpts0.txt, bpts1.txt, bpts2.txt, bpts3.txt,
    clock.avi,
    cover.odg, cover.png,
    crazy_blue.jpg,
    points.odp, points.ppt,
    pointsInfo.txt,
    questions.png,
    skinner.png


----------------------------
Compilation

All the Java files will compile with the compile.bat batch file:

> compile *.java
    // you must have LibreOffice, Java, JNA, and my Utils classes installed


===================================
Execution:

// You must have LibreOffice, Java, JNA, and my Utils classes installed

---------------------------
AnimBicycle.java

> run AnimBicycle
--> bicycle.odg

  // loads the image from <GALLERY>/transportation/Bicycle-Blue.png
---------------------------
AnimationDemo.java

> run AnimationDemo
  // not explained in chapter
---------------------------
AppendSlides.java

> run AppendSlides points.ppt algs.ppt
-->  points_Append.ppt

  // don't change the desktop focus during the execution
  // hands off the mouse and keyboard!
---------------------------
AutoShow.java

> run AutoShow algs.odp
  // user doesn't need to do anything
---------------------------
BasicShow.java

> run BasicShow algs.odp
  // user must click to advance slides; ESC to exit
---------------------------
BezierBuilder.java

> run BezierBuilder bpts0.txt
--> bezier.odg
---------------------------
CheckLinks.java

> run CheckLinks algs.ppt
  // not explained in chapter
---------------------------
CopySlide.java

> run CopySlide points.odp 1 4
--> overwrites the input file, points.odp

  // don't change the desktop focus during the execution
  // hands off the mouse and keyboard!
---------------------------
CustomShow.java

> run CustomShow algs.odp
  // user must click to advance slides; ESC to exit
---------------------------
DrawHilbert.java

> run DrawHilbert 4
--> hilbert.png

  // not explained in chapter
---------------------------
DrawPicture.java

  // loads the image from crazy_blue.png

> run DrawPicture
--> picture.odg
---------------------------
GalleryInfo.java

> run GalleryInfo
  // not explained in chapter
---------------------------
Grouper.java

> run Grouper
--> grouper.odg

  // style change to arrows causes crash on exit
     -- click on "Close the program"
     -- no LibreOffice process left running
---------------------------
MakeSlides.java

  // loads the image from skinner.png
  // links to videos in clock.avi and wildlife.wmv
  // you need to download wildlife.wmv separately from the webpage

  // fourth slide creation uses Draw.drawMedia() which causes a crash on exit
     -- click on "Close the program"
     -- no LibreOffice process left running

  // fifth slide creation is commented out since it's slow and hacky
     -- if you uncomment it, don't change the desktop focus during the execution
     -- hands off the mouse and keyboard!

> run MakeSlides
--> makeslides.odp

  // if the fifth slide is not added, then Office does not make the document 
     visible during its creation

  // to show the video and execute the buttons requires that "makeSlides.odp"
     be openned as a Slide Show in Office.
---------------------------
MastersUse.java

  // loads the image from questions.png

--> muSlides.odp

  // Office does not make the document visible during its creation
---------------------------
ModifySlides.java

> run ModifySlides algsSmall.ppt
-->  algsSmall_Mod.ppt

  // Office does not make the document visible during its creation
---------------------------
PointsBuilder.java

> run PointsBuilder pointsInfo.txt
--> points.odp

  // Office does not make the document visible during its creation
---------------------------
Slide2Image.java

> run Slide2Image algs.ppt 2 png
--> algs2.png

  // Office does not make the document visible during its creation
---------------------------
SlideShow.java

> run SlideShow
  // not explained in chapter
---------------------------
SlidesInfo.java

> run SlidesInfo points.odp

=================================
Last updated: 15th August 2015

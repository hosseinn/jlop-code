
JLOP Chapter 2. Text Processing

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

This directory contains 17 Java files:
  * BuildDoc.java, DocInfo.java, DocsAppend.java
    ExtractGraphics.java, ExtractText.java, HelloText.java
    HighlightText.java, ItalicsStyler.java
    MakeTable.java, MathQuestions.java, ShowBookText.java
    ShuffleWords.java, Speaker.java, StoryCreator.java
    StylesInfo.java, TalkingBook.java, TextReplace.java



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

  * compileTalk.bat
  * runTalk.bat
      - these are extended versions of the compile.bat and run.bat
        files which also use FreeTTS (http://freetts.sourceforge.net/) 
        and its Mbrola voices.


Assorted test files for the Java code:
  * oneLiner.odt, bondMovies.txt
    scandal.txt, scandalStart.txt
    skinner.png, story.doc, storyStart.doc



----------------------------
Compilation

All the Java files will compile with the compile.bat batch file 
except for Speaker.java, which requires compileTalk.bat:

> compileTalk Speaker.java
    // it uses FreeTTS
    // The details for installing FreeTTS, the Mbrola binary, and three voices
       are explained in Speaker.java
    // TalkingBook.java uses Speaker 

> ren Speaker.java Speaker.txt

> compile *.java
    // you must have LibreOffice, Java, JNA, and my Utils classes installed

> ren Speaker.txt Speaker.java


----------------------------
Execution:

The programs are listed roughly in the order you will encounter them
when reading Chapter 2:

// You must have LibreOffice, Java, JNA, and my Utils classes installed


> run ExtractText <fnm>
e.g.
> run ExtractText storyStart.doc


> run HighlightText <fnm>
e.g.
> run HighlightText storyStart.doc


> run HelloText
  // creates "hello.odt"


> runTalk TalkingBook <fnm>
e.g.
> runTalk TalkingBook storyStart.doc
    // uses Speaker.java which requires FreeTTS, Mbrola and 3 US voices
    // See inside Speaker.java for details about installing and using these


> run ShuffleWords <fnm>
e.g.
> run ShuffleWords storyStart.doc
  // creates "shuffled.doc"


> run StylesInfo <fnm>
e.g.
> run StylesInfo story.doc

> run StoryCreator <fnm>
e.g.
> run StoryCreator scandal.txt
  // creates "bigStory.doc"


> run BuildDoc
  // loads "skinner.png"; creates "build.odt"


> run MathQuestions
  // creates "mathQuestions.pdf"


> run ExtractGraphics <fnm>
e.g.
> run ExtractGraphics build.odt
  // creates PNG files called "graphics" + <NUM> + ".png"


> run MakeTable
  // loads "bondMovies.txt"; creates "table.odt"


> run ShowBookText <fnm>
e.g.
> run ShowBookText storyStart.doc


> run TextReplace <fnm>
e.g.
> run TextReplace story.doc
  // creates "replaced.doc"


> run ItalicsStyler <fnm> <string to italicize>
e.g.
> run ItalicsStyler story.doc scandal
  // creates "italicized.doc"


> run DocsAppend <fnm1> <fnm2> <fnm3>*
e.g.
> run DocsAppend oneLiner.odt story.doc
  // creates <fnm1>_APPENDED
     e.g. oneLiner_APPENDED.odt


> run DocInfo <fnm>
e.g.
> run DocInfo build.odt


----------------------------
Last updated: 15th August 2015


JLOP Chapter 9. Sending E-mail

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
 * JMailer.java, LoMailer.java, MailMerge.java


There are 4 test files for the Java code:
  * Addresses.odb, addresses.ods, 
    formLetter.ott, skinner.png


There is one e-mail related batch file:
 * TBExec.bat
     -- this calls Thunderbird's command line


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
Installing JavaMail

JavaMail is needed when using JMail.sendEmail(), which is used for
several examples in JMailer.java.

Download javax.mail.jar from:
  https://java.net/projects/javamail/pages/Home

Place it in 
  D:\JavaMail 
to match the paths used in compile.java and run.java


----------------------------
Compilation

> compile *.java
    // you must have LibreOffice, Java, JNA, JavaMail, and my Utils classes installed

===================================
Execution:

// You must have LibreOffice, Java, JNA, Javamail, and my Utils classes installed


> run LoMailer <password>
   -- there are 4 e-mail examples, aimed at my fivedots and Gmail accounts,
      with 3 commented out
   -- you will need to change the examples to use your e-mail account(s)


> run JMailer <password>
   -- there are 5 e-mail examples, aimed at my fivedots and Gmail accounts,
      with 4 commented out
   -- you will need to change the examples to use your e-mail account(s)


> run MailMerge [ <password> ]
   -- there are 3 mail merge examples, with 2 commented out
   -- the example that sends e-mail uses my Writer e-mail settings 
      which needs a password


=================================
Last updated: 8th July 2016

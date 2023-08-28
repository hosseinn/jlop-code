
// FindExtJar.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, September 2016

/* Use the module name to find the corresponding extension 
   installation folder, and save the path to its JAR file in
   lofindTemp.txt

   This program is called by the compileExt.bat and runExt.bat
   batch files.

   Usage:
     run FindExtJar org.openoffice.randomsents

   This will save the path to RandomSentsImpl.jar inside lofindTemp.txt

*/

import java.io.*;
import java.net.*;


public class FindExtJar
{
  public static void main(String args[])
  {
    if (args.length != 1) {
      System.out.println("Usage: run FindExtJar <ID>");
      return;
    }

    Lo.loadOffice();

    FileIO.saveString("lofindTemp.txt", "\"xx\"");  // clears temp file
    // Info.listExtensions();

    String extDir = Info.getExtensionLoc(args[0]);
    if ((extDir == null) || extDir.equals("")) {
      System.out.println("Could not find extension: " + args[0]);
      Lo.closeOffice();
      return;
    }
    // System.out.println("Location of " + args[0] + ": \n  " + extDir);

    // look in folder for JAR filename
    try {
      FilenameFilter filter = new FilenameFilter() {
           public boolean accept(File dir, String name) 
           {  return name.endsWith(".jar"); }
      };
      File dir = new File(new URI(extDir));
      String[] fnms = dir.list(filter);
      
      if (fnms == null)
        System.out.println("No jars found");
      else {
        String extPath = dir.getAbsolutePath();
        String jarNm = "\"" + extPath + "/" + fnms[0] + "\"";
        // System.out.println("JAR: " + jarNm);
        FileIO.saveString("lofindTemp.txt", jarNm);
      }
    }
    catch(java.lang.Exception e)
    {  System.out.println(e);  }

    Lo.closeOffice();
  } // end of main()


}  // end of FindExtJar class


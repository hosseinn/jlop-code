
// FindMacros.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, Nov. 2016

/* Demonstrates the use of Macros.findScripts()

  Usage:
     run FindMacros hello
     run FindMacros Logo

*/

import java.util.*;

import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.frame.*;
import com.sun.star.container.*;
import com.sun.star.beans.*;


public class FindMacros
{

  public static void main(String[] args)
  {
    if (args.length != 1) {
      System.out.println("Usage: run FindMacros <substring>");
      return;
    }

    XComponentLoader loader = Lo.loadOffice();

    ArrayList<String> scriptURIs = Macros.findScripts(args[0]);
    System.out.println("Matching Macros in Office: (" + scriptURIs.size() + ")");
    for(String scriptURI : scriptURIs)
      System.out.println("  " + scriptURI);

    Lo.closeOffice();
  } // end of main()



}  // end of FindMacros class


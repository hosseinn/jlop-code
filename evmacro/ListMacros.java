
// ListMacros.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, Nov. 2016

/* Demonstrates the use of Macros.getLangScripts()
  Usage:
     run ListMacros    or   run ListMacros Java
     run ListMacros Python
     run ListMacros BeanShell
     run ListMacros Basic
     run ListMacros JavaScript
*/

import java.util.*;

import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.frame.*;
import com.sun.star.container.*;
import com.sun.star.beans.*;


public class ListMacros
{

  public static void main(String[] args)
  {
    String lang = "Java";
    if (args.length != 1) {
      System.out.println("Usage: run ListMacros [Java | Python | BeanShell | Basic | JavaScript]");
      System.out.println("Using \"Java\"");
    }
    else 
      lang = args[0];

    XComponentLoader loader = Lo.loadOffice();

    ArrayList<String> scriptURIs = Macros.getLangScripts(lang);
    if (scriptURIs != null) {
      System.out.println(lang + " Macros in Office: (" + scriptURIs.size() + ")");
      for(String scriptURI : scriptURIs)
        System.out.println("  " + scriptURI);
    }

    Lo.closeOffice();
  } // end of main()



}  // end of ListMacros class


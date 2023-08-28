
// DocEvents.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, Nov. 2016

/* Manipulate document events, which are visible in Office's
   GUI via Tools > Customize > Events

   List all the Office and document events.

   Attach ShowEvent.show to the "OnStartApp" and "OnLoad" Office events

   Load "build.odt"
   Attach ShowEvent.show to its "OnPageCountChange" doc event; save doc

  Usage:
     run DocEvents
*/

import java.util.*;

import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.frame.*;
import com.sun.star.container.*;
import com.sun.star.document.*;
import com.sun.star.beans.*;

import com.sun.star.text.*;


public class DocEvents
{

  public static void main(String[] args)
  {
    XComponentLoader loader = Lo.loadOffice();

    Macros.listOfficeEvents();

    // list the "OnStartApp" and "OnLoad" Office event properties
    PropertyValue[] osaProps = Macros.getEventProps("OnStartApp");
    Props.showProps("OnStartApp Event", osaProps);

    PropertyValue[] olProps = Macros.getEventProps("OnLoad");
    Props.showProps("OnLoad Event", olProps);

    // attach macros to those event if they do not have macros already
    if (Lo.isNullOrEmpty( (String)Props.getProp(osaProps, "Script")))
      Macros.setEventScript("OnStartApp", 
         "vnd.sun.star.script:ShowEvent.ShowEvent.show?language=Java&location=share");

    if (Lo.isNullOrEmpty( (String)Props.getProp(olProps, "Script")))
      Macros.setEventScript("OnLoad", 
         "vnd.sun.star.script:ShowEvent.ShowEvent.show?language=Java&location=share");


    XTextDocument doc = Write.openDoc("build.odt", loader);
    if (doc == null) {
      System.out.println("Could not open build.odt");
      Lo.closeOffice();
      return;
    }

    GUI.setVisible(doc, true);
    Lo.wait(2000);

    Macros.listDocEvents(doc);

    // list the "OnPageCountChange" doc event properties
    PropertyValue[] opccProps = Macros.getDocEventProps(doc, "OnPageCountChange");
    Props.showProps("OnPageCountChange Event", opccProps);

    if (Lo.isNullOrEmpty( (String)Props.getProp(opccProps, "Script"))) {
      Macros.setDocEventScript(doc, "OnPageCountChange", 
         "vnd.sun.star.script:ShowEvent.ShowEvent.show?language=Java&location=share");

      Lo.save(doc);  // must save doc after event macro change
    }


    Lo.waitEnter();

    Lo.closeDoc(doc);
    Lo.closeOffice();
  } // end of main()



}  // end of DocEvents class


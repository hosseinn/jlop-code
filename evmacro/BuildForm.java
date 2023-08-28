
// BuildForm.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, Nov. 2016

/* Generate a simple form at run time, and assign the Java 
   ShowEvent macro as a listener to all its components.

   After the user presses enter, the form is saved as "build.odt"
*/


import com.sun.star.uno.*;
import com.sun.star.beans.*;
import com.sun.star.container.*;
import com.sun.star.lang.*;
import com.sun.star.view.*;
import com.sun.star.frame.*;

import com.sun.star.awt.*;
import com.sun.star.form.*;
import com.sun.star.text.*;



public class BuildForm                             
{


  public BuildForm()
  {
    XComponentLoader loader = Lo.loadOffice();
    XTextDocument doc = Write.createDoc(loader);

    if (doc == null) {
      System.out.println("Writer doc creation failed");
      Lo.closeOffice();
      System.exit(1);
    }

    GUI.setVisible(doc, true);
    Lo.delay(1000);

    XTextViewCursor tvc = Write.getViewCursor(doc);
    Write.append(tvc, "Building a Form\n");
    Write.endParagraph(tvc);

    createForm(doc);
    Lo.dispatchCmd("SwitchControlDesignMode");
               // switch from form design/editing mode to text (live mode)
    Lo.waitEnter();

    Lo.saveDoc(doc, "build.odt");
    Lo.closeDoc(doc);
    Lo.closeOffice();
  }  // end of BuildForm()




  private void createForm(XTextDocument doc)
  {
    XPropertySet props =
         Forms.addLabelledControl(doc, "FIRSTNAME", "TextField", 11);
    textEvents(props);    // only the first textfield has a listener

    Forms.addLabelledControl(doc, "LASTNAME", "TextField",  19);

    props = Forms.addLabelledControl(doc, "AGE", "NumericField", 43);
    Props.setProperty(props, "DecimalAccuracy", (short) 0);  // no decimal place

    Forms.addLabelledControl(doc, "BIRTHDATE", "FormattedField", 51);


    // buttons, all with listeners
    props = Forms.addButton(doc, "first", "<<", 2, 63, 8);
    buttonEvent(props);

    props = Forms.addButton(doc, "prev", "<", 12, 63, 8);
    buttonEvent(props);

    props = Forms.addButton(doc, "next", ">", 22, 63, 8);
    buttonEvent(props);

    props = Forms.addButton(doc, "last", ">>", 32, 63, 8);
    buttonEvent(props);

    props = Forms.addButton(doc, "new", ">*", 42, 63, 8);
    buttonEvent(props);

    props = Forms.addButton(doc, "reload", "reload", 58, 63, 13);
    buttonEvent(props);
  }  // end of createForm()



  public void textEvents(XPropertySet props)
  {
    Forms.assignScript(props, "XTextListener", "textChanged", 
                         "ShowEvent.ShowEvent.show", "share"); 
                          // obtained by calling ListMacros

    //Forms.assignScript(props, "XItemListener", "itemStateChanged", 
    //                     "ShowEvent.ShowEvent.show", "share"); 

    Forms.assignScript(props, "XFocusListener", "focusLost", 
                         "ShowEvent.ShowEvent.show", "share"); 

    Forms.assignScript(props, "XKeyListener", "keyPressed", 
                         "ShowEvent.ShowEvent.show", "share"); 
  }  // end of textEvents()



  public void buttonEvent(XPropertySet props)
  { Forms.assignScript(props, "XActionListener", "actionPerformed", 
                         "ShowEvent.ShowEvent.show", "share"); 
  }


  // --------------------------------------------------

  public static void main(String args[])
  {  new BuildForm();  }

}  // end of BuildForm class


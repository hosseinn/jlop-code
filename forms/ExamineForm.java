
// ExamineForm.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, June 2016

/* Open a text document containing a form, and attach button
   and text field listeners to every button and text field
   component.

   The text field listener monitors both text changes and 
   focus changes to the control.

   Usage:
       run ExamineForm form1.odt
                // simple form

       run ExamineForm scratchExample.odt
               // more realistic form; longer

       run ExamineForm build.odt
               // saved version of form created by BuildForm.java
*/

import java.util.ArrayList;

import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.frame.*;

import com.sun.star.container.*;
import com.sun.star.awt.*;




public class ExamineForm
{


  public static void main(String args[])
  {
    if (args.length != 1) {
      System.out.println("Usage: run ExamineForm <ODT file>");
      return;
    }
    XComponentLoader loader = Lo.loadOffice();
    XComponent doc = Lo.openDoc(args[0], loader);
    if (doc == null) {
      System.out.println("Could not open " + args[0]);
      Lo.closeOffice();
      return;
    }

    ArrayList<XControlModel> models = Forms.getModels(doc);
    System.out.println("No. of control models in form: " + models.size());

    // examine each model
    for(XControlModel model : models) {
      System.out.println("  " + Forms.getName(model) + ": " +
                                Forms.getTypeStr(model));
      // System.out.println("  Belongs to Form: \"" + Forms.getFormName(model) + "\"");
      // Props.showObjProps("Model", model);

      // look at the control for each model
      XControl ctrl = Forms.getControl(doc, model);
      if (ctrl == null) 
        System.out.println("   No control found");
      else {   // attach listener if the control is a button or text field
        if (Forms.isButton(model))
          attachButtonListener(ctrl, model);
        else if (Forms.isTextField(model))
          attachTextFieldListeners(ctrl);
      }
      // Info.showInterfaces("Component", ctrl);
      // Info.showServices("Component", ctrl);
    }
    System.out.println();
    GUI.setVisible(doc, true);
    Lo.waitEnter();

    Lo.closeDoc(doc);
    Lo.closeOffice();
  }  // end of main()



  private static void attachButtonListener(XControl ctrl, XControlModel cModel)
  {
    XButton xButton = Lo.qi(XButton.class, ctrl);
    // XControlModel cModel = ctrl.getModel();
    xButton.setActionCommand( Forms.getLabel(cModel));

    xButton.addActionListener( new XActionListener() {
      public void disposing(com.sun.star.lang.EventObject ev) {}

      public void actionPerformed(ActionEvent ev)
      {  System.out.println("Pressed \"" + ev.ActionCommand + "\""); }
    });
  }  // end of attachButtonListener()



  private static void attachTextFieldListeners(XControl ctrl)
  // listen for text changes and focus changes in the control
  {
    
    XTextComponent tc = Lo.qi(XTextComponent.class, ctrl);
    tc.addTextListener( new XTextListener() {
      public void textChanged(TextEvent ev)
      {
	      XControlModel cModel = Forms.getEventControlModel(ev);
        System.out.println( Forms.getName(cModel) + 
                             " text: " + Props.getProperty(cModel, "Text"));
      }  // end of textChanged()

      public void disposing(EventObject ev) {}

    });

    XWindow tfWindow = Lo.qi(XWindow.class, ctrl);
    tfWindow.addFocusListener( new XFocusListener() {

      public void focusLost(FocusEvent ev)
      { XControlModel model = Forms.getEventControlModel(ev);
        System.out.println("Leaving text: " + 
                            Props.getProperty(model, "Text"));
      }

      public void disposing(EventObject ev) {}
      public void focusGained(FocusEvent ev) {}
    });

  }  // end of attachTextFieldListeners()




}  // end of ExamineForm class

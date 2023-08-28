
// BuildForm.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, June 2016

/* Generate a form at run-time along with various listeners that
   print information

   After the user presses enter, the form (minus the listeners)
   is saved as "build.odt"

   There are 3 database form components -- two list boxes and a grid
   They use the tables in the DB_FNM database.
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

import com.sun.star.sdbc.*;


public class BuildForm implements XEventListener, 
                                  XPropertyChangeListener, XActionListener,
                                  XTextListener, XFocusListener, XItemListener,
                                  XMouseListener, XSelectionChangeListener,
                                  XGridControlListener
{
  private static final String DB_FNM = "liang.odb";    // database 

  private XTextDocument doc;  // for use by the listeners



  public BuildForm()
  {
    XComponentLoader loader = Lo.loadOffice();
    doc = Write.createDoc(loader);

    if (doc == null) {
      System.out.println("Writer doc creation failed");
      Lo.closeOffice();
      System.exit(1);
    }

    doc.addEventListener(this);  
        // for showing disposing of document (and controls)
    GUI.setVisible(doc, true);


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
    try {
      XPropertySet props =
           Forms.addLabelledControl(doc, "FIRSTNAME", "TextField", 11);
      listenToTextField(props);    // only the first textfield has a listener

      Forms.addLabelledControl(doc, "LASTNAME", "TextField",  19);

      props = Forms.addLabelledControl(doc, "AGE", "NumericField", 43);
      Props.setProperty(props, "DecimalAccuracy", (short) 0);  // no decimal place

      Forms.addLabelledControl(doc, "BIRTHDATE", "FormattedField", 51);


      // buttons, all with listeners
      props = Forms.addButton(doc, "first", "<<", 2, 63, 8);
      listenToButton(props);

      props = Forms.addButton(doc, "prev", "<", 12, 63, 8);
      listenToButton(props);

      props = Forms.addButton(doc, "next", ">", 22, 63, 8);
      listenToButton(props);

      props = Forms.addButton(doc, "last", ">>", 32, 63, 8);
      listenToButton(props);

      props = Forms.addButton(doc, "new", ">*", 42, 63, 8);
      listenToButton(props);

      props = Forms.addButton(doc, "reload", "reload", 58, 63, 13);
      listenToButton(props);
      listenToMouse(props);


      // some fixed text; no listener
      Forms.addControl(doc,  "text-1", "show only sales since", 
                                               "FixedText",  2, 80, 35, 6);


      // radio buttons inside a group box; use a property change listener
      Forms.addControl(doc, "Options", "Options", "GroupBox", 103, 5, 56, 25);

      // these three radio buttons have the same name ("Option"), and
      // so only one can be on at a time
      props = Forms.addControl(doc, "Option",  "No automatic generation", // "NoAuto
                                                       "RadioButton", 106, 11, 50, 6);
      props.addPropertyChangeListener("State", this);
               // can not deselect radio buttons (?)
      
      props = Forms.addControl(doc, "Option", "Before inserting a record",   // "Before"
                                                  "RadioButton", 106, 17, 50, 6);
      props.addPropertyChangeListener("State", this);

      props = Forms.addControl(doc, "Option", "When moving to a new record",  // "Move"
                                                  "RadioButton", 106, 23, 50, 6);
      props.addPropertyChangeListener("State", this);


      // check boxes inside another group box
      // use the same property change listener
      Forms.addControl(doc, "Misc", "Miscellaneous", "GroupBox", 103, 35, 56, 25);
     
      props = Forms.addControl(doc, "DefaultDate", 
                "Default sales date to \"today\"", "CheckBox", 106, 39, 60, 6);
      Props.setProperty(props, "HelpText",
                "When checked, newly entered sales records are pre-filled");
      props.addPropertyChangeListener("State", this);
     
      props = Forms.addControl(doc, "Protect", 
                "Protect key fields from editing", "CheckBox", 106, 45, 60, 6);
      Props.setProperty(props, "HelpText",
                "When checked, you cannot modify the values");
      props.addPropertyChangeListener("State", this);
     
      props = Forms.addControl(doc, "Empty", 
                 "Check for empty sales names", "CheckBox", 106, 51, 60, 6);
      Props.setProperty(props, "HelpText",
                 "When checked, you cannot enter empty values");
      props.addPropertyChangeListener("State", this);


      // a list using simple text
      String[] fruits = { "apple", "orange", "pear", "grape"};
      props = Forms.addList(doc, "Fruits", fruits, 2, 90, 20, 6);
      listenToList(props);


      // image button; commented out because it occasionally hangs (?)
/*
      props = Forms.addButton(doc, "Smiley", null, 90, 80, 10, 10);
      //props = Forms.addControl(doc, "Smiley", null,
      //                                         "ImageButton", 90, 80, 10, 10);
      props.setPropertyValue("ImageURL", FileIO.fnmToURL("smiley.png"));
      listenToButton(props);
*/

      // set Form's data source to be the DB_FNM database
      XForm defForm = Forms.getForm(doc, "Form");
      Forms.bindFormToTable(defForm, FileIO.fnmToURL(DB_FNM), "Course");

      // a list filled using an SQL query on the form's data source
      props = Forms.addDatabaseList(doc,
                 "CourseNames", "SELECT \"title\" FROM \"Course\"", 90, 90, 20, 6);
      listenToList(props);


      // another list filled using a different SQL query on the form's data source
      props = Forms.addDatabaseList(doc,
             "StudNames", "SELECT \"lastName\" FROM \"Student\"",  140, 90, 20, 6);
      listenToList(props);


      // ------------------------ set up database grid/table ----------------

      // create a new form, gridForm, 
      XNameContainer gridCon = Forms.insertForm("GridForm", doc);
      XForm gridForm = Lo.qi(XForm.class, gridCon);

      // which uses an SQL query as its data source
      Forms.bindFormToSQL(gridForm, FileIO.fnmToURL(DB_FNM), 
                        "SELECT \"firstName\", \"lastName\" FROM \"Student\"");

      // create the grid/table component and set its columns
      props = Forms.addControl(doc, "SalesTable", null, 
                                 "GridControl", 2, 100, 100, 40, gridCon);

      XControlModel gridModel = Lo.qi(XControlModel.class, props);
      Forms.createGridColumn(gridModel, "firstName", "TextField", 25);
      Forms.createGridColumn(gridModel, "lastName", "TextField", 25);

      listenToGrid(gridModel);
      // listenToMouse(props);
    }
    catch (com.sun.star.uno.Exception e) {
      System.out.println(e);
    }
  }  // end of createForm()



  // --------- attach various listeners to the components ---------------

  /* all these functions work in a similar way: the component properties
     are cast into a model, then the model is converted into a general
     control (of type XControl), and then cast to a specific control,
     such as a button or text field. A listener is attached to this
     specific control/component.
  */


  public void listenToButton(XPropertySet props)
  {
    XControlModel cModel = Lo.qi(XControlModel.class, props);
    XControl control = Forms.getControl(doc, cModel);
    XButton xButton = Lo.qi(XButton.class, control);

    xButton.setActionCommand(Forms.getName(cModel));
    xButton.addActionListener(this);
  }  // end of listenToButton()




  public void listenToTextField(XPropertySet props)
  {
    XControlModel cModel = Lo.qi(XControlModel.class, props);
    XControl control = Forms.getControl(doc, cModel);

    // convert the control into two components, so two different
    // listeners can be attached
    XTextComponent tc = Lo.qi(XTextComponent.class, control);
    tc.addTextListener(this);

    XWindow tfWindow = Lo.qi(XWindow.class, control);
    tfWindow.addFocusListener(this);
  }  // end of listenToTextField()




  public void listenToList(XPropertySet props)
  {
    XControlModel cModel = Lo.qi(XControlModel.class, props);
    XControl control = Forms.getControl(doc, cModel);

    XListBox listBox = Lo.qi(XListBox.class, control);
    listBox.addItemListener(this);
  }  // end of listenToList()



  public void listenToGrid(XControlModel gridModel)
  {
    // Info.showServices("Model", gridModel);
    XControl control = Forms.getControl(doc, gridModel);
    XGridControl gc = Lo.qi(XGridControl.class, control);
    gc.addGridControlListener(this);

    XSelectionSupplier gridSelection = 
                        Lo.qi(XSelectionSupplier.class, gc);
    gridSelection.addSelectionChangeListener(this);
  }  // end of listenToGrid()




  public void listenToMouse(XPropertySet props)
  {
    XControlModel cModel = Lo.qi(XControlModel.class, props);
    XControl control = Forms.getControl(doc, cModel);
    XWindow xWindow = Lo.qi(XWindow.class, control);
    xWindow.addMouseListener(this);
  }  // end of listenToMouse()



  // define the listener methods for ...


  /* ------------ XEventListener ---------------- */


  public void disposing(EventObject ev)
  {
    String implName = Info.getImplementationName(ev.Source);
    System.out.println("Disposing: " + implName);
  }  // end of disposing()




  /* ------------ XActionListener ---------------- */

  public void actionPerformed(ActionEvent ev)
  // called when a button has been pressed
  {  System.out.println("Pressed \"" + ev.ActionCommand + "\"");  }  





  /* ------------ XTextListener ---------------- */

  public void textChanged(TextEvent ev)
  // used by the text field
  { XControlModel cModel = Forms.getEventControlModel(ev);
    System.out.println(Forms.getName(cModel) + 
                       " text: " + Props.getProperty(cModel, "Text"));
  }  // end of textChanged()




  /* ------------ XFocusListener ---------------- */
  // also used by the text field

  public void focusGained(FocusEvent ev)
  { XControlModel cModel = Forms.getEventControlModel(ev);
    System.out.println("Into: " + Forms.getName(cModel));
  }


  public void focusLost(FocusEvent ev)
  { XControlModel cModel = Forms.getEventControlModel(ev);
    System.out.println("Left: " + Forms.getName(cModel));
  }




  /* ------------ XPropertyChangeListener ---------------- */


  public void propertyChange(PropertyChangeEvent ev)
  // used by the radio buttons and checkboxes
  {
    System.out.println("Property Change detected");
    if (ev.PropertyName.equals("State")) {
      String name = (String) Props.getProperty(ev.Source, "Name");
      String label = (String) Props.getProperty(ev.Source, "Label");
      short classID = ((Short) Props.getProperty(ev.Source, "ClassId")).shortValue();
      boolean isEnabled = (((Short) ev.NewValue).shortValue() != 0);

      // did it come from a radio button or checkbox?
      if (classID == FormComponentType.RADIOBUTTON)
        System.out.println("\"" + label + "\" radio button: " + isEnabled);
            // use the label since all my radio buttons have the same name
      else if (classID == FormComponentType.CHECKBOX)
        System.out.println("\"" + name + "\" checkbox: " + isEnabled);
    }
  }  // end of propertyChange()




  /* ------------ XItemListener ---------------- */


  public void itemStateChanged(ItemEvent ev)
  // used by the list boxes
  {
    short selection = (short) ev.Selected;
    String selectedItem = null;

    XControlModel cModel = Forms.getEventControlModel(ev);
    XListBox listBox = Lo.qi(XListBox.class, ev.Source);
    if (listBox != null)
      selectedItem = listBox.getItem(selection);
    System.out.println("List: \"" + Forms.getName(cModel) + 
                                  "\" selected: " + selectedItem);
  }  // end of itemStateChanged()



  /* ------------ XMouseListener ---------------- */


  public void mousePressed(MouseEvent ev)
  {
    XControlModel cModel = Forms.getEventControlModel(ev);
    System.out.println("Pressed " + Forms.getName(cModel) + 
                         " at (" + ev.X + ", " + ev.Y + ")");
  }

  public void mouseReleased(MouseEvent ev)
  {  // System.out.println("Released");
  }

  public void mouseEntered(MouseEvent ev)
  { XControlModel cModel = Forms.getEventControlModel(ev);
    System.out.println("Entered " + Forms.getName(cModel));  
  }

  public void mouseExited(MouseEvent ev)
  {  // System.out.println("Exited");
  }



  // -------------------- XSelectionChangeListener ----------------------


  public void selectionChanged(EventObject ev)
  // used by the grid/table control
  {
    XControlModel cModel = Forms.getEventControlModel(ev);
    XGridControl gc =  Lo.qi(XGridControl.class, ev.Source);  
             // in the form module
             // not the one in com.sun.star.awt.grid
    System.out.println("Grid " + Forms.getName(cModel) + 
                       " column: " + gc.getCurrentColumnPosition());
            // no way to get the current row with XGridControl; see form.XGrid

    // must access the result set inside the form for row info
    String formName = Forms.getFormName(cModel);
    // System.out.println("Parent form: " + formName);
    XForm gForm = Forms.getForm(doc, formName);
    XResultSet rs = Lo.qi(XResultSet.class, gForm);
    try {
      System.out.println("    row: " + rs.getRow());
    }
    catch (com.sun.star.uno.Exception e) {
      System.out.println(e);
    }
  }  // end of selectionChanged()


  // -------------------- XGridControlListener ------------------

  public void columnChanged(EventObject ev)
  // used by the grid/table control
  {  System.out.println("Grid Column change");  } 



  // --------------------------------------------------

  public static void main(String args[])
  {  new BuildForm();  }

}  // end of BuildForm class


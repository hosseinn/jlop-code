
// ${className}.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, ${.now}
// Generated using CreateAddonImpl.java

/* Code for processing an Add-on with a dialog interface.

   The dialog's initialization (in initDialog()) assumes there
   is a button (called "CommandButton1") and a textfield (called
   "TextField1"). Listeners are attached to both:
   actionPerformed() and keyPressed(), which should be extended to
   use the info string they access in the text field.

   A Console window is used for printing debug info during the
   execution of the Addon. This can be removed when the coding
   is finished. Delete the Console construction in this class's
   constructor and its visibility changes in processCmd(). The
   System.out.println() calls can be left in, but will not
   appear anywhere.

   More info at:
      https://wiki.openoffice.org/wiki/Documentation/DevGuide/WritingUNO/Implementation
*/

import com.sun.star.uno.*;
import com.sun.star.lib.uno.helper.*;
import com.sun.star.lang.*;
import com.sun.star.frame.*;
import com.sun.star.awt.*;
import com.sun.star.util.*;
import com.sun.star.beans.*;
import com.sun.star.registry.*;

import com.sun.star.text.*;


public class ${className} extends WeakBase implements
                    XInitialization, XServiceInfo, XDispatchProvider, XDispatch,   // for Addon
                    XActionListener, XTopWindowListener, XKeyListener              // for dialog
{
  // used to get an instance of this service
  private static final String[] serviceNames = {"com.sun.star.frame.ProtocolHandler"};

  private static final String myName = ${className}.class.getName();

  private Console console;    // for debugging output
  private int printCount = 1;

  // gives access to the service manager and registered services
  private XComponentContext xcc;

  // UNO awt components of the dialog
  private XDialog dialog = null;
  private XTextComponent textBox;   // the text in the dialog's text field



  public ${className}(XComponentContext context)
  {
    xcc = context;
    console = new Console();
    System.out.println("${className} called");
  }  // end of ${className}()



  // -------------------------- XInitialization method ----------------------


  public void initialize(Object[] objects) throws com.sun.star.uno.Exception
  {   System.out.println("initialize() called");  }


  // -------------------------- XServiceInfo methods -------------------------


  public String[] getSupportedServiceNames()
  {  return serviceNames;  }

  public static String[] getServiceNames()   // static version for outer class
  {   return serviceNames;  }


  public boolean supportsService(String service)
  {
    for (int i = 0; i < serviceNames.length; i++) {
      if (service.equals(serviceNames[i]))
        return true;
    }
    return false;
  }  // end of supportsService()


  public String getImplementationName()
  {  return ${className}.class.getName();  }



  // ----------------------- XDispatchProvider methods ------------------------


  public XDispatch queryDispatch(URL commandURL, String targetFrameName, int searchFlags)
  // return dispatch object which provides the functionality for the command URL, or null
  {
    if (commandURL.Protocol.compareTo("${extensionID}:") == 0) {
<#list cmdNames as cmd>
      if (commandURL.Path.compareTo("${cmd}") == 0) {
        System.out.println("queryDispatch() called for \"${cmd}\"");
        return this;
      }
</#list>
    }
    return null;
  }  // end of queryDispatch()



  public XDispatch[] queryDispatches(DispatchDescriptor[] descrs)
  // returns multiple dispatch objects for all the specified descriptors
  {
    int count = descrs.length;
    XDispatch[] dispatches = new XDispatch[count];
    for (int i = 0; i < count; i++)
      dispatches[i] = queryDispatch(descrs[i].FeatureURL,
                          descrs[i].FrameName, descrs[i].SearchFlags);
    return dispatches;
  }  // end of queryDispatches()



  // ------------------------ XDispatch methods --------------------------


  public void dispatch(URL commandURL, PropertyValue[] props)
  // process the command URL
  {
    if (commandURL.Protocol.compareTo("${extensionID}:") == 0) {
<#list cmdNames as cmd>
      if (commandURL.Path.compareTo("${cmd}") == 0)
        processCmd("${cmd}");   
</#list>
    }
  }  // end of dispatch()



  public void addStatusListener(XStatusListener xControl, URL commandURL)
  // registers a listener for a specific control and command URL
  {
    if (commandURL.Protocol.compareTo("${extensionID}:") == 0) {
<#list cmdNames as cmd>
      if (commandURL.Path.compareTo("${cmd}") == 0)
        System.out.println("status listener added for \"${cmd}\"");
</#list>
    }
  }  // end of addStatusListener()



  public void removeStatusListener(XStatusListener xControl, URL commandURL)
  {
    if (commandURL.Protocol.compareTo("${extensionID}:") == 0) {
<#list cmdNames as cmd>
      if (commandURL.Path.compareTo("${cmd}") == 0)
        System.out.println("status listener removed for \"${cmd}\"");
</#list>
    }
  } // end of removeStatusListener()



  // -------------- methods used by Office's service manager ----------------

  public static XSingleComponentFactory __getComponentFactory(String implName)
  {
    XSingleComponentFactory xFactory = null;
    if (implName.equals(myName))
      xFactory = Factory.createComponentFactory(${className}.class, serviceNames);
    return xFactory;
  }


  public static boolean __writeRegistryServiceInfo(XRegistryKey xRegistryKey)
  {  return Factory.writeRegistryServiceInfo(myName, serviceNames, xRegistryKey);  }


  // ---------------------- command processing --------------------------



  private void processCmd(String cmd)
  {
    XComponent doc = Lo.addonInitialize(xcc);   // so my utils can be used safely

    System.out.println("Window title: " + GUI.getTitleBar());
    System.out.println(printCount++ + ". dispatch() called for \"" + cmd + "\"");

    if (cmd.equals("help")) {
      GUI.showMessageBox("Add-on Help", "Type in the text, then press return or click the Ok button.");
      return;
    }

    // any other command is processed by the following code...
    console.setVisible(true);

    dialog = Dialogs.loadAddonDialog("${extensionID}",
                                      "dialogLibrary/" + cmd + ".xdl");
    if (dialog == null) {
      System.out.println("Could not load " + cmd + " dialog");
      return;
    }

    XControl dialogControl = Dialogs.getDialogControl(dialog);
    initDialog(dialogControl);
    Dialogs.execute(dialogControl);

    // Lo.delay(3000); 
    console.setVisible(false);
  }  // end of processCmd()



  private void initDialog(XControl dialogControl)
  {
    // listen to the dialog window
    XTopWindow topWin = Dialogs.getDialogWindow(dialogControl);
    topWin.addTopWindowListener(this);

    // Dialogs.showControlInfo(dialogControl);

    // set listener for Ok button
    XButton button = Lo.qi(XButton.class,
                        Dialogs.findControl(dialogControl, "CommandButton1"));
    button.addActionListener(this);

    // set listener for text box
    textBox = Lo.qi(XTextComponent.class, 
                              Dialogs.findControl(dialogControl, "TextField1"));

    XWindow xTFWindow = (XWindow) Lo.qi(XWindow.class, textBox);
    xTFWindow.addKeyListener(this);
    xTFWindow.setFocus();
  }  // end of initDialog()



  // ----------------------- dialog listener methods -----------------------


  public void actionPerformed(ActionEvent e)
  // for processing the Ok button
  {
    // System.out.println("Event received from control : " + Dialogs.getEventSourceName(e));

    String info = textBox.getText();
    if (info.equals(""))
      return;
    System.out.println("Info: \"" + info +"\"");
    textBox.setText("");

    // ADD CODE HERE using info
  }  // end of actionPerformed()



  public void windowClosing(EventObject event)
  // for processing the closure of the dialog window
  {  // System.out.println("Window closing");
    dialog.endExecute();
    //console.closeDown();
  }


  public void keyPressed(KeyEvent event)
  // for dealing with text entry in the text box
  {  
    // System.out.println("Key char / code: " + event.KeyChar + " / " + event.KeyCode);
    if (event.KeyCode == Key.RETURN) {  // https://www.openoffice.org/api/docs/common/ref/com/sun/star/awt/Key.html

      String info = textBox.getText();
      if (info.equals(""))
        return;
      System.out.println("Info: \"" + info +"\"");
      textBox.setText("");

      // ADD CODE HERE using info
    }
  }  // end of keyPressed()



  // ------------------ other listener methods -------------------------------

  public void windowOpened(EventObject event) {}

  public void windowActivated(EventObject event) {}

  public void windowDeactivated(EventObject event) {}

  public void windowMinimized(EventObject event) {}

  public void windowNormalized(EventObject event) {}

  public void windowClosed(EventObject event) {}

  public void disposing(EventObject event) {}

  public void keyReleased(KeyEvent event) {}

}  // end of ${className} class


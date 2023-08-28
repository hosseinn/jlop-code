
// NumActionListener.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, Nov. 2016

/* ActionListener used by GetNumber

   Listener for whether the user presses "ok" or "Cancel" in 
   the dialog, and update the form's text field (cModel) accordingly.
*/

import com.sun.star.lang.*;
import com.sun.star.awt.*;



public class NumActionListener implements XActionListener
{
  private XDialog dialog;
  private XControlModel cModel;
  private int val;


  public NumActionListener(XDialog dialog, XControlModel cModel, int val)
  {
    this.dialog = dialog;
    this.cModel = cModel;
    this.val = val;
  }  // end of NumActionListener()



  public void actionPerformed(ActionEvent e)
  {
    String buttonName = Dialogs.getEventSourceName(e);
    System.out.println("Event received from : " + buttonName);

    if (buttonName.equals("CommandButton1"))        // "OK" button
      Props.setProperty(cModel, "Text", "" + val);  // put val in the text field
    else if (buttonName.equals("CommandButton2"))   // "Cancel" button
      Props.setProperty(cModel, "Text", "");        // clear the text field

    dialog.endExecute();
  }  // end of actionPerformed()


  public void disposing(EventObject e) { } 

}  // end of NumActionListener class



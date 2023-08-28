
// SwingViewer.java
// Andrew Davison, August 2016, ad@fivedots.coe.psu.ac.th

/* Embeds Office inside a JPanel using OOoBean.

   Added button controls and a textfield for changing page,
   and zooming. The page changes only work for multi-page
   Writer and Impress docs.


  Usage:
    run SwingViewer classic_letter.ott

    run SwingViewer skinner.png
    run SwingViewer algs.odp
        -- read-only mode yellow banner appears when page
           number typed in

    run SwingViewer tables.ods
    run SwingViewer liangTables.odb

*/


import javax.swing.*;
import java.awt.*;
import java.awt.event.*;

import com.sun.star.comp.beans.*;
import com.sun.star.beans.*;
import com.sun.star.lang.*;
import com.sun.star.text.*;
import com.sun.star.uno.*;
import com.sun.star.frame.*;
import com.sun.star.view.*;
import com.sun.star.drawing.*;
// import com.sun.star.awt.*;


public class SwingViewer extends JFrame
{
  private static final int ZOOM_STEP = 5;

  private OBeanPanel officePanel;

  private XComponent doc = null;
  private int docType = Lo.UNKNOWN;

  private XTextViewCursor tvCursor = null;   // for text docs

  private XDrawView xDrawView = null;     // for slide docs
  private XDrawPages xDrawPages = null;

  // GUI
  private JTextField currPageTF, gotoPageTF;
  private JLabel pageCountLabel;

  private int currentZoom = 100;
  private int pageCount = -1;



  public SwingViewer(String fnm)
  {
    super("SwingViewer");
    Container c = getContentPane();

    officePanel = new OBeanPanel(850, 600);
                     // OOoBean inside a JPanel
    c.add(officePanel, BorderLayout.CENTER);
    if (officePanel.getBean() == null) {
      System.out.println("OOoBean not created");
      System.exit(1);
    }

    c.add(makeButtons(), BorderLayout.EAST);
    c.add(makePageControls(), BorderLayout.SOUTH);

    addWindowListener(new WindowAdapter() {
      public void windowClosing(WindowEvent e){
        officePanel.closeDown();
        System.exit(0);
      }
    });

    pack();
    setLocationRelativeTo(null);  // center the window
    setResizable(false);
    setVisible(true);
    Lo.delay(200);    // give time for the OOBean to appear

    doc = officePanel.loadDoc(fnm);

    docType = Info.reportDocType(doc);
    if (doc == null) {
      officePanel.closeDown();
      System.exit(0);
    }

   // initialize loaded document
    if (docType == Lo.WRITER)
      initTextDoc();
    else if (docType == Lo.IMPRESS)
      initDrawDoc();

    // GUI.printUIs(doc);
  }  // end of SwingViewer()



  // -------------------- build GUI -----------------------


  private JPanel makeButtons()
  /* 6 buttons in a panel: jump to start, move up one page,
     move down one page, jump to end, zoom-in, zoom-out
  */
  {
    JButton startButton = new JButton("To Start");
    startButton.addActionListener(new ActionListener() {
      public void actionPerformed(ActionEvent ev)
      { if (docType == Lo.WRITER)
          textJump(1);
        else if (docType == Lo.IMPRESS)
          slideJump(1);
      }
    });

    JButton upButton = new JButton("Up");
    upButton.addActionListener(new ActionListener() {
      public void actionPerformed(ActionEvent ev)
      {  pageChange(false);  }   // isDown is false, i.e. move up
    });

    JButton downButton = new JButton("Down");
    downButton.addActionListener(new ActionListener() {
      public void actionPerformed(ActionEvent ev)
      {  pageChange(true);  }   // isDown is true, i.e. move down
    });

    JButton endButton = new JButton("To End");
    endButton.addActionListener(new ActionListener() {
      public void actionPerformed(ActionEvent ev)
      { if (docType == Lo.WRITER)
          textJump(pageCount);
        else if (docType == Lo.IMPRESS)
          slideJump(pageCount);
      }
    });

    JButton zInButton = new JButton("Zoom In");
    zInButton.addActionListener(new ActionListener() {
      public void actionPerformed(ActionEvent ev)
      { zoom(true);  }
    });

    JButton zOutButton = new JButton("Zoom Out");
    zOutButton.addActionListener(new ActionListener() {
      public void actionPerformed(ActionEvent ev)
      {  zoom(false);  }
    });

    JPanel p = new JPanel();
    p.setLayout(new BoxLayout(p, BoxLayout.Y_AXIS));
    p.add(Box.createRigidArea(new Dimension(0, 15)));
    p.add(startButton);
    p.add(Box.createRigidArea(new Dimension(0, 15)));
    p.add(upButton);
    p.add(Box.createRigidArea(new Dimension(0, 15)));
    p.add(downButton);
    p.add(Box.createRigidArea(new Dimension(0, 15)));
    p.add(endButton);
    p.add(Box.createRigidArea(new Dimension(0, 15)));
    p.add(zInButton);
    p.add(Box.createRigidArea(new Dimension(0, 15)));
    p.add(zOutButton);

    return p;
  }  // end of makeButtons()



  private JPanel makePageControls()
  /* panel holding a page count label, and a textfield
     for jumping to a page based on its page number.
  */
  {
    JPanel p = new JPanel();

    pageCountLabel = new JLabel("Page Count: ??");
    p.add(pageCountLabel);

    p.add( new JLabel("   Current Page: "));
    currPageTF = new JTextField(5);
    currPageTF.setEditable(false);
    currPageTF.setText("1");
    p.add(currPageTF);

    p.add( new JLabel("   Go to Page: "));

    gotoPageTF = new JTextField(10);
    p.add(gotoPageTF);
    gotoPageTF.addActionListener( new ActionListener() {
      public void actionPerformed(ActionEvent e)
      {  // System.out.println(e.getActionCommand());
         pageJump(e.getActionCommand());
      }
    });

    return p;
  }  // end of makePageControls()


  private void setCurrentPage(int i)
  {  currPageTF.setText("" + i);  }


  private void setLastPage(int i)
  {  pageCountLabel.setText("Last Page: " + i);  }




  // ------------------ office doc setup ---------------------



  private void initTextDoc()
  {
    // report the number of pages in the doc
    XTextDocument textDoc = Lo.qi(XTextDocument.class, doc);
    pageCount = Write.getPageCount(textDoc);
    setLastPage(pageCount);

    try {  // get the document's visible cursor
      XController xc = officePanel.getController();
      XTextViewCursorSupplier tvcSupplier =
                            Lo.qi(XTextViewCursorSupplier.class, xc);
      tvCursor = tvcSupplier.getViewCursor();
    }
    catch(java.lang.Exception e) {
      System.out.println("Could not access document cursor");
    }
  }  // end of initTextDoc()



  private void initDrawDoc()
  {
    try {
      // report the number of slides in the doc
      xDrawPages = Draw.getSlides(doc);
      pageCount = Draw.getSlidesCount(doc);
      setLastPage(pageCount);

      // get a slide view
      xDrawView = Lo.qi(XDrawView.class, officePanel.getController());
    }
    catch(java.lang.Exception e)
    {  System.out.println("Could not access document pages");  }
  }  // end of initDrawDoc()



  // -------------------------- page change ---------------


  public void pageChange(boolean isDown)
  // goto next/previous page/slide in the doc
  {
    if (docType == Lo.WRITER)
      textChange(isDown);
    else if (docType == Lo.IMPRESS)
      slideChange(isDown);
  }  // end of pageChange()



  private void textChange(boolean isDown)
  // goto next/previous page in the doc
  {
    if (tvCursor == null)
      return;
    int pageNum = Write.getCurrentPage(tvCursor);
    if (isDown)  // move down
      textJump(pageNum+1);
    else
      textJump(pageNum-1);
  }  // end of textChange()



  private void slideChange(boolean isDown)
  // goto next/previous slide in the doc
  {
    if (xDrawView == null)
      return;
    int slideNum = Draw.getSlideNumber(xDrawView);
    if (slideNum == -1)
      return;
    if (isDown && (slideNum == pageCount))   // at end of slides
      return;
    if (!isDown && (slideNum == 1))   // at start of slides
      return;

    int newSlideNum = ((isDown) ? (slideNum + 1) : (slideNum - 1));
    slideJump(newSlideNum);
  }  // end of slideChange()




  // ----------------------------- page jump ---------------------


  private void pageJump(String pageStr)
  // jump to the numbered page/slide in the doc
  {
    int pgNum = Lo.parseInt(pageStr);
    if (docType == Lo.WRITER)
      textJump(pgNum);
    else if (docType == Lo.IMPRESS)
      slideJump(pgNum);
    else
      System.out.println("Incompatible document format");
  }  // end of pageJump()



  private void textJump(int pgNum)
  // jump to the numbered page in the doc
  {
    if (pgNum < 1) {
      System.out.println("Smallest page number is 1");
      pgNum = 1;
    }
    else if (pgNum > pageCount) {
      System.out.println("Largest page number is " + pageCount);
      pgNum = pageCount;
    }

    XPageCursor pageCursor = Lo.qi(XPageCursor.class, tvCursor);
    if (pageCursor == null) {
      System.out.println("Could not create a page cursor");
      return;
    }

    boolean hasJumped = pageCursor.jumpToPage((short)pgNum);
    if (!hasJumped)
      System.out.println("Jump to page " + pgNum + " failed");
    else
      setCurrentPage(pgNum);
  }  // end of textJump()



  private void slideJump(int slideNum)
  // jump to the numbered slide in the doc
  {
    if (xDrawPages == null)
      return;
    try {
      XDrawPage newPage = Lo.qi(XDrawPage.class,
                              xDrawPages.getByIndex(slideNum-1));
      xDrawView.setCurrentPage(newPage);
      setCurrentPage(slideNum);
    }
    catch(java.lang.Exception e)
    {  System.out.println("Could not go to slide " + slideNum);  }
  }  // end of slideJump()




  // ----------------------- zooming ---------------------------


  public void zoom(boolean zoomIn)
  { if (zoomIn)
      zoomTo(currentZoom + ZOOM_STEP);
    else   // zoom out
      zoomTo(currentZoom - ZOOM_STEP);
  }  // end of zoom();



  private void zoomTo(int zoomValue)
  {
    PropertyValue[] props = Props.makeProps("Zoom.Value", zoomValue,
                                            "Zoom.Type", 3);
    Lo.dispatchCmd(officePanel.getFrame(), "Zoom", props);
    currentZoom = zoomValue;
  }  // end of zoomTo()



  // ---------------------------------------------------

  public static void main(String args[])
  {
    if (args.length == 1)
      new SwingViewer(args[0]);
    else
      System.out.println("Usage: run SwingViewer [<fnm>]");
  }  // end of main()

} // end of SwingViewer


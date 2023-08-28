
// AnimationDemo.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, July 2015

// Based on an example by Ariel, January 2009

import com.sun.star.uno.*;
import com.sun.star.frame.*;
import com.sun.star.lang.*;
import com.sun.star.ucb.XSimpleFileAccess;
import com.sun.star.comp.helper.Bootstrap;

import com.sun.star.awt.*;
import com.sun.star.beans.*;
import com.sun.star.drawing.*;
import com.sun.star.animations.*;

import com.sun.star.presentation.*;



public class AnimationDemo
{

  public static void main(String[] args)
  {
    XComponentLoader loader = Lo.loadOffice();
    XComponent doc = Draw.createImpressDoc(loader);
    if (doc == null) {
      System.out.println("Impress doc creation failed");
      Lo.closeOffice();
      return;
    }

    XDrawPage slide = Draw.getSlide(doc, 0);   // access first page
    Draw.blankSlide(slide);

    // add an ellipse to the center of the slide
    Size slideSize = Draw.getSlideSize(slide);
    int width = 50;
    int height = 50;
    int x = slideSize.Width/2 - width/2;
    int y = slideSize.Height/2 - height/2;
    XShape s1 = Draw.drawEllipse(slide, x, y, width, height);

    /* animation goes from top-left to middle-bottom and then to top-right,
       playing the applause sound at the same time
    */
    try {
      XAnimationNode root = Draw.getAnimationNode(slide);
      setUserData(root, EffectNodeType.AFTER_PREVIOUS, EffectPresetClass.MOTIONPATH, null);

      // root --> seq --> par
      XTimeContainer rootTime = Lo.qi(XTimeContainer.class, root);

      XTimeContainer seqTime = Lo.createInstanceMCF(XTimeContainer.class,
                                       "com.sun.star.animations.SequenceTimeContainer");
      rootTime.appendChild(seqTime);

      XTimeContainer parTime = Lo.createInstanceMCF(XTimeContainer.class,
                                       "com.sun.star.animations.ParallelTimeContainer");
      parTime.setAcceleration(0.05);
      parTime.setDecelerate(0.05);
      parTime.setFill(AnimationFill.HOLD);
      seqTime.appendChild(parTime);


      // set animation of ellipse to execute in parallel
      XAnimateMotion motion = Lo.createInstanceMCF(XAnimateMotion.class,
                                               "com.sun.star.animations.AnimateMotion");
      motion.setDuration(2);
      motion.setFill(AnimationFill.HOLD);
      motion.setTarget(s1);
      motion.setPath("m -0.5 -0.5 0.5 1 0.5 -1");
      parTime.appendChild(motion);


      // create audio playing in parallel
      XAudio audio = Lo.createInstanceMCF(XAudio.class, "com.sun.star.animations.Audio");
      String fnm = Info.getGalleryDir() + "/sounds/applause.wav";   
      audio.setSource(FileIO.fnmToURL(fnm));
      audio.setVolume(1.0);     // does setting the volume work?
      parTime.appendChild(audio);
    }
    catch (com.sun.star.uno.Exception e) {
      e.printStackTrace();
    }

    GUI.setVisible(doc, true);
       // slideshow start() crashes if the doc is not visible

    XPresentation2 show = Draw.getShow(doc);
    show.start();
    XSlideShowController sc = Draw.getShowController(show);
    Draw.waitEnded(sc);

    Lo.closeDoc(doc);
    Lo.closeOffice();
  }  // end of main()



  public static void setUserData(XAnimationNode node, short effectNodeType,
                              short effectPresetClass, String presetID)
  {
    boolean hasID = ((presetID != null) && (presetID.length() > 0));

    NamedValue[] userData = new NamedValue[hasID ? 3 : 2];
    userData[0] = new NamedValue("node-type", Short.valueOf(effectNodeType));
    userData[1] = new NamedValue("preset-class", Short.valueOf(effectPresetClass));
    if (hasID) 
      userData[2] = new NamedValue("preset-id", presetID);

    node.setUserData(userData);
  }  // end of setUserData()




  // ---------------- not used --------------------

  public static XIterateContainer createIterateContainer()
  {  return Lo.createInstanceMCF(XIterateContainer.class, "com.sun.star.animations.IterateContainer"); }


  public static XAnimateSet createAnimateSet()
  {  return Lo.createInstanceMCF(XAnimateSet.class, "com.sun.star.animations.AnimateSet" ); }


  public static XAnimateColor createAnimateColor()
  {  return Lo.createInstanceMCF(XAnimateColor.class, "com.sun.star.animations.AnimateColor"); }


  public static XAnimateTransform createAnimateTransform()
  {  return Lo.createInstanceMCF(XAnimateTransform.class, "com.sun.star.animations.AnimateTransform");  }


  public static XTransitionFilter createTransitionFilter()
  {  return Lo.createInstanceMCF(XTransitionFilter.class, "com.sun.star.animations.TransitionFilter"); }


  public static XCommand createAnimCommand()
  {  return Lo.createInstanceMCF(XCommand.class, "com.sun.star.animations.Command"); }


}  // end of AnimationDemo.java

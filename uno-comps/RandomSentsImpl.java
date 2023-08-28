
// RandomSentsImpl.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, September 2016


/* Based on code in RandomSentences.java

*/


import com.sun.star.uno.XComponentContext;
import com.sun.star.lang.XSingleComponentFactory;
import com.sun.star.registry.XRegistryKey;

// helper documentation is part of "Java UNO Reference"
import com.sun.star.lib.uno.helper.WeakBase;
   // the base class for UNO components
   // http://api.libreoffice.org/docs/java/ref/
   //            com/sun/star/lib/uno/helper/WeakBase.html
   // implements XWeak, XTypeProvider
      /* no need for this class to implement XTypeProvider
         methods, unlike the example in the dev guide */

import com.sun.star.lib.uno.helper.Factory;
   // implements XServiceInfo and XSingleComponentFactory
   /* replaces old-style com.sun.star.comp.loader.FactoryHelper
      used in dev guide */


public final class RandomSentsImpl extends WeakBase
                      implements com.sun.star.lang.XServiceInfo,
                                 org.openoffice.randomsents.XRandomSents
{
  private final XComponentContext xcc;

  private static final String implName = RandomSentsImpl.class.getName();

  private static final String[] serviceNames = 
                       { "org.openoffice.randomsents.RandomSents" };


  // ----------- random sentences data structures  -------------------

  /* The grammar and data come from:
        http://funprogramming.org/57-A-random-sentence-generator-writes-nonsense.html
  */

  private static final int MAX_SENTENCES = 100;

  private static String[] articles = {
     "the", "my", "your", "our", "that", "this", "every", "one", "the only", 
     "his", "her" 
  };

  private static String[] adjs = {
     "happy", "rotating", "red", "fast", "elastic", "smily", "unbelievable", 
     "infinite", "surprising", "mysterious", "glowing", "green", "blue", 
     "tired", "hard", "soft", "transparent", "long", "short", "excellent", 
     "noisy", "silent", "rare", "normal", "typical", "living", "clean", 
     "glamorous", "fancy", "handsome", "lazy", "scary", "helpless", "skinny", 
     "melodic", "silly", "kind", "brave", "nice", "ancient", "modern", 
     "young", "sweet", "wet", "cold", "dry", "heavy", "industrial", 
     "complex", "accurate", "awesome", "shiny", "cool", "glittering", "fake", 
     "unreal", "naked", "intelligent", "smart", "curious", "strange", 
     "unique", "empty", "gray", "saturated", "blurry" 
  };

  private static String[] nouns = {
     "forest", "tree", "flower", "sky", "grass", "mountain", "car", 
     "computer", "man", "woman", "dog", "elephant", "ant", "road", 
     "butterfly", "phone", "program", "grandma", "school", "bed", 
     "mouse", "keyboard", "bicycle", "spaghetti", "drink", "cat", "t-shirt", 
     "carpet", "wall", "poster", "airport", "bridge", "road", "river", 
     "beach", "sculpture", "piano", "guitar", "fruit", "banana", "apple", 
     "strawberry", "rubber band", "saxophone", "window", "linux OS", 
     "skateboard", "piece of paper", "photograph", "painting", "hat", 
     "space", "fork", "mission", "goal", "project", "tax", "windmill", 
     "light bulb", "microphone", "cpu", "hard drive", "screwdriver" 
  }; 

  private static String[] preps = {
     "under", "in front of", "above", "behind", "near", "following", 
     "inside", "besides", "unlike", "like", "beneath", "against", "into", 
     "beyond", "considering", "without", "with", "towards" 
  };

  private static String[] verbs = {
     "sings", "dances", "was dancing", "runs", "will run", "walks", "flies", 
     "moves", "moved", "will move", "glows", "glowed", "spins", "promised", 
     "hugs", "cheated", "waits", "is waiting", "is studying", "swims", 
     "travels", "traveled", "plays", "played", "enjoys", "will enjoy", 
     "illuminates", "arises", "eats", "drinks", "calculates", "kissed", 
     "faded", "listens", "navigated", "responds", "smiles", "will smile", 
     "will succeed", "is wondering", "is thinking", "is", "was", "will be", 
     "might be", "was never"
  };


  private boolean isAllCaps = false;


  // ------------------ boiler-plate ------------------


  public RandomSentsImpl(XComponentContext context)
  {  xcc = context; }


  public static XSingleComponentFactory __getComponentFactory(String name)
  /* this method replaces the old style __getServiceFactory()
     which returns a XSingleServiceFactory
  */ 
  { XSingleComponentFactory xFactory = null;
    if (name.equals(implName))
      xFactory = Factory.createComponentFactory(RandomSentsImpl.class, serviceNames);
    return xFactory;   // can create a service instance
  }


  // this method is still needed or unopkg fails
  public static boolean __writeRegistryServiceInfo(XRegistryKey regKey) 
  // writes implementation info into the registry key
  {  return Factory.writeRegistryServiceInfo(implName, serviceNames, regKey);   }


  // ----- XServiceInfo methods -----

  public String getImplementationName() 
  {  return implName;  }


  public boolean supportsService(String sService) 
  {
    for( int i=0; i < serviceNames.length; i++)
      if (sService.equals(serviceNames[i]))
        return true;
    return false;
  }


  public String[] getSupportedServiceNames() 
  {  return serviceNames;  }



  // ------------------- random sentences methods -----------------

  public boolean getisAllCaps()
  {  return isAllCaps;  }

  public void setisAllCaps(boolean b)
  {  isAllCaps = b; }



  public String getParagraph(int numSents)
  {
    String[] sents = getSentences(numSents);
    StringBuilder sb = new StringBuilder();
    for (int i=0; i < sents.length; i++)
      sb.append( sents[i]);
    return sb.toString();
  }  // end of getParagraph()



  public String[] getSentences(int numSents) 
  {
    if (numSents < 1)
      numSents = 1;
    else if (numSents > MAX_SENTENCES)
      numSents = MAX_SENTENCES;

    String[] sents = new String[numSents];
    StringBuilder sb = new StringBuilder();
    for (int i=0; i < numSents; i++) {
      sb.setLength(0);    // empty builder
      sb.append( capitalize(pickWord(articles)) + " ");
      sb.append( pickWord(adjs) + " ");
      sb.append( pickWord(nouns) + " ");
      
      sb.append( pickWord(verbs) + " ");
      sb.append( pickWord(preps) + " ");
      
      sb.append( pickWord(articles) + " ");
      sb.append( pickWord(adjs) + " ");
      sb.append( pickWord(nouns) + ". ");

      sents[i] = isAllCaps ? sb.toString().toUpperCase() : sb.toString();
    }
    return sents;
  }  // end of getSentences()



  private String pickWord(String[] words) 
  {  return words[ (int)(Math.random()*words.length) ];  }


  private String capitalize(String word) 
  {  return word.substring(0,1).toUpperCase() + word.substring(1);  }

}  // end of RandomSentsImpl class

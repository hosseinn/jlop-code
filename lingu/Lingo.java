
// Lingo.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, July 2016

/* Examples using the Linguistics API (the com.sun.star.linguistic2 module).
   See Dev Guide, ch.6, p.706, "Linguistics".
   https://wiki.openoffice.org/wiki/Documentation/DevGuide/OfficeDev/Linguistics

   Use the spell checker, thesaurus, proof reader (grammar checker), and
   language guesser. The hypernator isn't used.

   The proof reader uses LanguageTool (https://www.languagetool.org/)
   not the default LightProof on one of my test machines.
   (http://extensions.libreoffice.org/extension-center/lightproof-editor;
    http://libreoffice.hu/2011/12/08/grammar-checking-in-libreoffice/)

   LanguageTool was separately installed as an extension to Office, and configured using
   Office's GUI. It needs Java 8 unlike LightProof.
*/


import com.sun.star.beans.*;
import com.sun.star.lang.*;
import com.sun.star.container.*;

import com.sun.star.linguistic2.*;
import com.sun.star.uno.*;

import java.util.Arrays;


public class Lingo
{

  public static void main(String args[])
  {
    Lo.loadOffice();

    // print linguistics info
    Write.dictsInfo();    

    XLinguProperties linguProps = Write.getLinguProperties();
    Props.showProps("Linguistic Manager", linguProps);
                   // see lodoc LinguProperties

    Info.listExtensions();   // these include linguistic extensions

    // get lingo manager
    XLinguServiceManager2 lingoMgr = 
             Lo.createInstanceMCF(XLinguServiceManager2.class, 
                     "com.sun.star.linguistic2.LinguServiceManager");
    if (lingoMgr == null) {
      System.out.println("No linguistics manager found");
      Lo.closeOffice();
      return;
    }

    Write.printServicesInfo(lingoMgr);

    // load spell checker & thesaurus
    XSpellChecker speller = lingoMgr.getSpellChecker();
    XThesaurus thesaurus = lingoMgr.getThesaurus();


    // use spell checker
    // Write.printLocales("Spell checker", speller.getLocales());
    Write.spellWord("horseback", speller); 
    Write.spellWord("ceurse", speller);
    // Write.spellWord("CEURSE", speller);
    Write.spellWord("magisian", speller);
    Write.spellWord("ellucidate", speller);

    // use thesaurus
    Write.printMeaning("magician", thesaurus);
    Write.printMeaning("elucidate", thesaurus);

/*
    Write.setConfiguredServices(lingoMgr, "Proofreader", "org.languagetool.openoffice.Main");
    // Write.setConfiguredServices(lingoMgr, "Proofreader", "org.libreoffice.comp.pyuno.Lightproof.en");

    Locale usLoc = new Locale("en","US","");
    Write.printConfigServiceInfo(lingoMgr, "Proofreader", usLoc);
*/

    // load & use proof reader (Lightproof or LanguageTool)
    XProofreader proofreader = Write.loadProofreader();
    System.out.println("Proofing...");
    int numErrs = Write.proofSentence("i dont have one one dogs.", proofreader);
    System.out.println("No. of proofing errors: " + numErrs + "\n");


    // guess the language
    Locale loc = Write.guessLocale("The rain in Spain stays mainly on the plain.");
    Write.printLocale(loc);
    if (loc != null)
      System.out.println("Guessed language: " + loc.Language);

    loc = Write.guessLocale("A vaincre sans péril, on triomphe sans gloire.");
                // To win without risk is a triumph without glory.
    if (loc != null)
      System.out.println("Guessed language: " + loc.Language);


    Lo.closeOffice();
  }  // end of main()


}  // end of Lingo class


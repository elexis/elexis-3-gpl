package com.jsigle.msword_js;

//TODO: REMOVE AS MANY REFERENCES / PARTS FROM agIon as possible and replace them by lean local ones,
//TODO: e.g. the agIonDoc.setModified() which fails during clear() during save on close() WordEventHandler,
//TODO: and rather give us a local setModified() if that should not break anything else (around the panel / view logic);
//TODO  BUT KEEP A panel/view Window, even if empty (or with the filename) within Elexis.

//TODO: ProbablyUsingTarmed_xxTemplateSoDoNOTAccessHeaderRangeToAvoidGenerationOfEmptyHeaderLines is a workaround to keep Header completely absent and to set Shapes ZORDER to bringToFront for Tarmedrechnung_xx templates.  

//TODO: Prüfen: Braucht es bei den findOrReplace Implementationen im orig -> replace repl Abschnitt noch den Support für das Einfügen
//TODO: von Tabellen, wenn von cb als repl eine Tabelle zurückgeliefert wird? Der Code war im NoaText/OpenOffice basierten plugin vorhanden,
//TODO: aktuell auskommentiert, weil es ja weiter unten ein insertTable gibt, welches eben solches tut, z.B. für [Liste].

//TODO: Der Code für findOrReplace in Headers könnte auch für Footers dupliziert werden.
//TODO: Der eigentliche Suchen-Ersetzen-Code in den findOrReplace-Abschnitten ist möglicherweise so refactorierbar,
//TODO: dass die Hauptteile als Methoden ausgegliedert und für die verschiedenartigen Dokumentanteile dann identisch verwendet werden.
//TODO: Das verbraucht dann aber wiederum etwas Verständlichkeit, weil ja die betroffenen Dispatches etc. andere sind -
//TODO: und das eben NICHT nur Textbausteines sind, die leicht und einheitlich übergegben werden könnten. 

//TODO: vermutlich via Close-Listener (den existierenden anpassen/übernehmen): Save wenn das Word-Fenster ODER das Panel in Elexis geschlossen wird.
//TODO: AutoSave alle paar Minuten: geht das in Word, wenn von VBA aus aufgerufen?
//TODO: Derzeit wird wohl in File, dann in Datenbak gespeichert, wenn ich von Word heraus irgendwo in Elexis hin klicke.
//TODO: Noch nicht gespeichert wird, wenn ich in Word hineinklicke, dort ändere, und dann via WORD close Button das Fenster schliesse.
//TODO: Unklar: Was passiert, wenn ich z.B. die Briefe View schliesse?

//TODO: insertTextAt (oder wie es heisst, für TextFrame/Shape-Texte, offenbar bei der Rechnungsschreibung notwendig).

//TODO: setFont / setFormat Support etc.

//TODO: Bei meinen Etikettenvorlagen funktioniert nur für 1/3 das findOrReplace; bei den anderen beiden wird maximal die Patientennummer ersetzt! WARUM??? -> Verbessern.
//TODO: Wahrscheinlich liegt es daran, dass findOrRestore in Shapes = TextFrames nur den ersten Platzhalter ersetzt - genauso wie beim TarmedRechnungsdruck;
//TODO:    und das hängt wahrscheinlich mit der fehlenden MoveRight; MoveLeft; eben nach dieser ersten Ersetzung zusammen.

//TODO: Nach dem Etikettendrucken bleibt das Word-Fenster offen. Sollte nicht sein.
//TODO: print(): waitUntilCompleted Support hinzufügen. testen ob funktioniert: toPrinter; toTray. 

//TODO: Sidenote: js ch.elexis.text/TextContainer.java createFromTemplate(): aboout to return brief...  (2x o in aboout!)

//TODO: Refactor: Use the added methods close, quit, etc. when closing, quitting, instead of calls to Dispatch() or whatever (if possible), in removeMe, dispose, clean, etc.

//TODO: Refactor the code in order to: Give all (at least all important) variables and objects proper names, indicating their type.
//ToDo: Especially: ActiveXComponent vs. Dispatch objects (like: jacobObjWord vs jacobDocuments, jacobDocument etc.). Allenfalls jacobSelection noch innerhalb des Files hoch zu jacobObjWord level nehmen, siehe auch bei dispose(). 
//ToDo: Check whether we could homogenize the use of ActiveXComponent vs. Dispatch objects.
//ToDo: Remove unnecessary code left over from the NoaText/OpenOffice based implementation.

//TODO: In META-INF/MANIFEST.MF, currently, the plugin ID is com.jsigle.MSWord_js,
//and the Activator is: ag.ion.bion.workbench.office.editor.core.EditorCorePlugin -
//this should probably be changed in the course of replacing
//NOAText / OpenOffice Interface specifics by JaCoB MS Word Interface specifics.

//DIES IST EINE FASSUNG, welche ich nun in 20160921 wieder ins elexis-2.1.7-...js...20130605based hereingeholt habe,
//mit der jetzt aktuellen Version von Jacob 1.18. Und sie funktiniert soweit.
//Hier ist auch noch viel mehr Code von NOA enthalten, als später nötig sein wird - siehe unten.



//201609221323js: Commenting out     System.setProperty(IOfficeApplication.NOA_NATIVE_LIB_PATH,getLibrariesLocation());
//in com.jsigle.msword_js/src/ag/ion/bion/workbench/office/editor/core/EditorCorePlugin.java
//made a number of errors in the Elexis startup plugin registration phase disappear,
//and also resulted in the Einstellungen - MSWord_js configuration page to appear correctly (still with references to OpenOffice, anyway...),
//and made the error message in the Briefe view disappear and let it appear with gray space (much more correct),
//i.e.: transformed the plugin from: NON-WORKING, NON REGISTRABLE on startup, to WORKING, REGISTRABLE on startup.
//
//************************************************************************************************************
//YEP, AND FINALLY, I CAN NOW OPEN LETTERS FROM Briefauswahl IN MSWORD VIA THIS PLUGIN (AGAIN). 201609221330js
//SO NOW, THE msword_js PLUGIN (proof of feasibility prototype) is back working, in elexis 2.1.7js :-) :-) :-) 
//************************************************************************************************************
//
//The line has been commented out now, and comprehensive console log printout has been included there. 
//
//ToDo: We'll have to review if any similar functionality is needed for the msword_js variant,
//ToDo: or only for the noa/ag.ion environment, and thereafter either strip it out completely, or adopt it.


//201609221234js: Das Starten des Plugins hat weiterhin mit einigen Fehlermeldungen NICHT funktioniert,
//BIS ich (1) die *.dlls aus dem /lib/ Subdirectory im Project in Eclipse entfernt habe
//(siehe unten zu alternativen Positionen), was die ZipFile-related-Fehler dazu verschwinden lies,
//
//UND dann: in META-INF/MANIFEST.MF: die ID: geändert nach com.jsigle.MSWord_js (GROSSbuchstaben am Anfang!) statt com.jsigle.msword_js
//Also, nicht die package-ID muss dot stehen offenbar, sondern der Name des constructors unten... Hmpf.
//
//ToDo: Prüfen: Stimmt das so wirklich??? Wenn ich die drastisch reduzierten Fehler beim startup anschaue, dann wohl schon...?
//NEIN, absolut nicht. Denn jetzt wird das Plugin einfach GAR NICHT MEHR initialisiert,
//taucht auch nicht mehr in den Einstellungen unter Textverarbeitung oder mit eigener Settings-Seite auf.

//201609210652js Das Starten des Plugins hat zunächst NICHT funktioniert -
//BIS ich die beiden jacob-1.18-x86.dll und jacob-1.18-x64.dll kopiert habe.
//von L:\Elexis\jacob (wo sie hätten gefunden werden müssen, da in Run Configuration VMWare Arguments
//der Pfad mit angegeben, wie unten beschrieben, und früher funktioniert - aber vielleicht verhindert
//das aktuelle Win 7 das Ausführen von Programmcode von L: ???),
//nach:
//C:\Program Files (x86)\Java\jdk1.7.0_71\bin
//und
//C:\Program Files (x86)\Java\jre1.8.0_77\bin
//
//Ich hab beide Files an beide Orte gelegt, weil ich mir jetzt keine Gedanken darüber machen will,
//was wann verwendet wird (vermutlich nur nötig: x86 im jre, da ich 32-Bit Java verwende und
//zum Laufenlassen wohl auch innerhalb von Eclipse, aber vor allem ausserhalb, eher das jre als das jdk
//verwendet wird).
//
//Wenn die beiden *.dlls im Ordner lib unterhalb des Projekts liegen (danach allenfalls mit F5 Refresh des Projects in Eclipse),
//dann gibt es seitenweise extra Fehlermeldungen: error opening zip file jacob-1.18-x86.dll etc.
//
//Die sind wenigstens verschwunden, nachdem ich die beiden *.dll dort rausgenommen habe.
//Trotzdem noch Fehler beim Initialisieren des Plugins.





//201609210202js Trying to bring this project back into my 2.1.7 version of Elexis and complete it.
//I want to do this (finally), to provide a truly *stable* integration of MS Word,
//after I had put aside the prototype made for feasibility testing when offering Medelexis to
//do it for them (within a funded project, where funding had been available), and they still gave the
//task to the colleagues from Australia (Thomas Huster et all) - so I didn't want to duplicate something
//that others would provide anyway (and would be paid for doing, anyway).
//What they did, however: (a) requires Word 2007 or newer; and (b) modifies the *.docx file directly,
//supposedly bypassing any API, and (c) is reported NOT to be completely stable, still, and (d)
//run into some problems from direct file tampering during my tests, i.e. not processing some SQL placeholders,
//supposedly when they were very long and split by the conversion from odt to docx into multiple XML tags,
//or whatever. Before, I had done intense testing and identified several issues as a quality check for Thomas Huster,
//and they fixed several things as a result of that - both in their plugin, and in Elexis - but I didn't find it
//perfect up to my last tests (about a year or two ago) and didn't want Juerg to migrate to something that might
//be of unknown reliability, and closed, and only in Elexis 3.x, where others of my extensions are still missing etc.
//WELL. NOW, wo get it clear, I'm trying to get this approach here back into the workbench and look where I can get...


//WARNING: AS I'M GETTING CODE BACK OUT FROM 2012, THE INCLUDED NOATEXT AND JS CONTENT MAY BE SEVERELY OUTDATED.
//ToDo: Update the Noatext and js portions to match the latest 2.1.7 20130605based versions.
//ToDo: Backport the Niklaus Giger & Thomas Huster addons since 20130605, PRESERVING MY NEWER/OWN MODS, to that 20130605based version.
//      (That has only been done in the 201311xx based version *completely*, and partially for the other one
//       because I did it to that one erroneously, instead of the 20130605based, which Jürg actually uses.)
//ToDo: LATER: Update the Noatext and js portions to match the latest 2.1.7 201311xx based with document housekeeping.
//ToDo: LATER: Update it all to the Elexis 3.x environment, given that Gerry told me he adopted my noatext_js to there,
//      and that had not been difficult, in the beginning of this year.



//20120304js Combined ch.elexis.NOAText_js.NOAText.java (based upon noa-libre, noa4e 2.0.14)
//and JACOB sample WordDocumentProperties.java so that this file shows both the example implementation
//of an MS Word interface via JACOB application (no editor, only a properties lookup thing),
//and the original (js improved) OpenOffice/LibreOffice interface via NOA/noa4e/OfficePanel/EditorPlugin...,
//illustrating which methods must be provided in this file and how they have been implemented in the NOAText world.

//This stage of the development actually does run (even if the connection from the plugin.xml file
//to the class implementing ITextPlugin does NOT point to the original method NOAText (further below, commented out),
//but to the method MSWord - which has been enhanced by all variables from NOAText -
//but at the beginning, it still provides only the OO/LO functionality via NOA.

//In order to achieve this, I have copied all the methods and variables/objects from NOAText
//originally implementing the abstract methods of ITextPlugin over into the new MSWord class.
//This achieves a collection of code that can formally be compiled at any time,
//although functionality is only the old one / a very rudimentary portion of the new one at the beginning.

//During further development, more and more methods will be changed over from implementation
//based upon NOA and not really functional in MSWord_jsText, to implementation based upon JACOB and functional.

//I will keep this starting state for this development (maybe in closed older versions of this project) because:
//The JACOB WordDocumentProperties.java example quite straightforwardly uses calls to Word to do something.
//The NOAText ITextPlugin.java implementation, however, relies heavily on (use com.jsigle. instead of ch.elexis. now):
///ch.elexis.noatext_jsl/src/ag/ion/noa4e/ui/widgets/OfficePanel.java,
//ch.elexis.noatext_jsl/src/ag/ion/noa4e/ui/wizards/application/LocalApplicationWizard.java,
///ch.elexis.noatext_jsl/src/ag/ion/noa4e/ui/wizards/application/LocalApplicationWizardDefinePage.java
//and on various other files below the ag.ion structure. 

//So apparently, NOAText.java uses (at least) one more wrapping level than MSWord_js.java might require.
//This means that either, I should adopt OfficePanel.java rather than NOAText.java to meet the requirements of Elexis,
//or that I may require to implement/pull up some controlling of frames and user interactivity
//up in(to) the NOAText.java=MSWord_jsText.java level.

//And as I cannot perfectly say which level would be better to embed the adoption,
//I want to keep a way back to the start readily available.

import java.awt.Frame;
import java.io.BufferedInputStream;
import java.io.Closeable;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.LinkedList;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.eclipse.core.runtime.CoreException;
import org.eclipse.core.runtime.FileLocator;
import org.eclipse.core.runtime.IConfigurationElement;
import org.eclipse.core.runtime.Path;
import org.eclipse.core.runtime.Platform;
import org.eclipse.swt.SWT;
import org.eclipse.swt.widgets.Composite;
import org.osgi.framework.Bundle;

import ag.ion.bion.officelayer.application.IOfficeApplication;
import ag.ion.bion.officelayer.application.OfficeApplicationException;
import ag.ion.bion.officelayer.document.DocumentDescriptor;
import ag.ion.bion.officelayer.document.DocumentException;
import ag.ion.bion.officelayer.event.ICloseEvent;
import ag.ion.bion.officelayer.event.ICloseListener;
import ag.ion.bion.officelayer.event.IEvent;
import ag.ion.bion.officelayer.form.IFormComponent;
import ag.ion.bion.officelayer.form.IFormService;
import ag.ion.bion.officelayer.text.ITextDocument;
import ag.ion.bion.officelayer.text.ITextRange;
import ag.ion.bion.officelayer.text.ITextTable;
import ag.ion.bion.officelayer.text.table.ITextTablePropertyStore;
import ag.ion.bion.workbench.office.editor.core.EditorCorePlugin;
import ag.ion.noa.NOAException;
import ag.ion.noa.search.ISearchResult;
import ag.ion.noa.search.SearchDescriptor;
import ag.ion.noa4e.ui.widgets.OfficePanel;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComException;
import com.jacob.com.Dispatch;
import com.jacob.com.DispatchEvents;
import com.jacob.com.Variant;

import com.sun.star.awt.FontWeight;
import com.sun.star.awt.Size;
import com.sun.star.awt.XTextComponent;
import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.beans.XPropertySetInfo;
import com.sun.star.drawing.XShape;
import com.sun.star.form.FormComponentType;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.style.ParagraphAdjust;
import com.sun.star.text.HoriOrientation;
import com.sun.star.text.RelOrientation;
import com.sun.star.text.TextContentAnchorType;
import com.sun.star.text.VertOrientation;
import com.sun.star.text.XText;
import com.sun.star.text.XTextCursor;
import com.sun.star.text.XTextDocument;
import com.sun.star.text.XTextFrame;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.view.PrintableState;
import com.sun.star.view.XPrintable;

import com.jsigle.msword_js.MSWord_jsPrinter;
import com.jsigle.msword_js.MSWord_jsPrinter.MyXPrintJobListener;
import com.jsigle.msword_js.MSWord_jsText;
import com.jsigle.msword_js.MSWord_jsText.closeListener;

import ch.elexis.core.ui.text.ITextPlugin;
import ch.elexis.core.ui.util.SWTHelper;
import ch.rgw.io.FileTool;
import ch.rgw.tools.ExHandler;
import ch.rgw.tools.Log;
import ch.rgw.tools.StringTool;
import ch.rgw.tools.TimeTool;
import ch.elexis.core.data.interfaces.text.ReplaceCallback;
/**
 * Submitted to the Jacob SourceForge web site as a sample 3/2005
 * <p>
 * Added a quit method and done other things - 2012, 2016 js
 * Freundlicherweise lässt das quit das Word tatsächlich laufen, wenn z.B. aus voriger Session noch Dokumente offen sind, d.h. die werden nicht mit geschlossen :-) :-) :-)
 * 
 * @author Date Created Description Jason Twist 04 Mar 2005 Code opens a locally
 *         stored Word document and extracts the Built In properties and Custom
 *         properties from it. This code just gives an intro to JACOB and there
 *         are sections that could be enhanced
 */
public class MSWord_jsText implements ITextPlugin {
	//Please note: Upon close() and quit(), I do also set jacobObjWord = null; jacobDocument = null; etc. - So you need to re-allocate these if needed again.

	//ToDo: This is a workaround. Solve it properly sometimes!
	//As we can't make Word completely remove and hide a completely empty header programmatically yet
	//(maybe we should try to delete *all* kinds of headers?)
	//https://msdn.microsoft.com/en-us/library/office/ff197738.aspx
	//https://msdn.microsoft.com/en-us/library/office/aa196645(v=office.11).aspx
	//we now try to recognize the Tarmed_xx templates, so that we do NOT access the header.range at all in these.
	//Recognition: In findOrReplace, if it's called with search pattern \[Titel\] and that produces a hit. Yep, that's it :-)
	//See: findOrReplace.
	private Boolean ProbablyUsingTarmed_xxTemplateSoDoNOTAccessHeaderRangeToAvoidGenerationOfEmptyHeaderLines = false;
	
	// Declare word object
	private ActiveXComponent jacobObjWord;

	// Declare Word Properties
	private Dispatch jacobCustDocprops;

	private Dispatch jacobBuiltInDocProps;

	// The jacobDocuments object is important in any real app but this demo doesn't use it
	
	private Dispatch jacobDocuments;

	private Dispatch jacobDocument;

	private Dispatch jacobWordObject;

	public WordEventHandler	jacobWordEventHandler = null;		//For a MS Word window close listener
	
	//The following variables are from the NOAText plugin;
	//I copy them here so that the previous (=now default, to be adopted) implementations
	//of methods inherited from ITextPlugin and ICloseListener can work before they are
	//adopted into the MSWord_jsText world. Some of these variables/objects most probably have to be adopted as well.

	private final Log log = Log.get("MSWord_jsText");
	
	public static final String MIMETYPE_OO2 = "application/vnd.oasis.opendocument.text";
	public static LinkedList<MSWord_jsText> agIonNoas = new LinkedList<MSWord_jsText>();

	OfficePanel			agIonPanel;							//type from ag.ion...
	ITextDocument		agIonDoc;							//type from ag.ion...
	ICallback			textHandler;						//type from ch.elexis.text.ITextPluginCallback callback interface for save operations
	IOfficeApplication	agIonOffice;						//type from ag.ion...
	
	File myFile;											//type from java.io
	
	private String font;
	private float hi = 0;
	private int stil = -1;
	
	//Here come the methods from js
	/**
	 * Prints the MSWord_jsText interface status
	 */
	public void debug_print_status() {
		System.out.println("MSWord_jsText JACOB interface status:");
		
		if (jacobObjWord==null)	System.out.println("WARNING: jacobObjWord==null");
		else					System.out.println("jacobObjWord="+jacobObjWord.toString());
		
		if (jacobCustDocprops==null)	System.out.println("WARNING: jacobCustDocprops==null");
		else							System.out.println("jacobCustDocprops="+jacobCustDocprops.toString());

		if (jacobBuiltInDocProps==null)	System.out.println("WARNING: builtInDocProps==null");
		else							System.out.println("builtInDocProps="+jacobBuiltInDocProps.toString());

		if (jacobWordObject==null)	System.out.println("WARNING: jacobWordObject==null");
		else						System.out.println("jacobWordObject="+jacobWordObject.toString());

		if (jacobDocuments==null)	System.out.println("WARNING: jacobDocuments==null");
		else						System.out.println("jacobDocuments="+jacobDocuments.toString());	
		
		if (jacobDocument==null)	System.out.println("WARNING: jacobDocument==null");
		else						System.out.println("jacobDocument="+jacobDocument.toString());	

		if (jacobWordEventHandler==null)	System.out.println("WARNING: jacobWordEventHandler==null");
		else								System.out.println("jacobWordEventHandler="+jacobWordEventHandler.toString());	
	}
	
	
	//Here come the methods from the sample WordDocumentProperties.java
	

	/**
	 * Empty Constructor
	 * 
	 */
	public MSWord_jsText() {
		System.out.println("MSWord_jsText: Constructor - This is the MSWord_jsText plugin speaking :-)");
		System.out.println("MSWord_jsText: The NOAText interface would obtain the path to OO here, but we don't need that now.");
		System.out.println("MSWord_jsText: At the most, we might get the path to the jacob*.dll, but this also should not be necessary.");
		System.out.println("MSWord_jsText: OR WE MIGHT configure the name of the ActiveXComponent to run,");
		System.out.println("MSWord_jsText: mainly if \"Word.Application\" should not work.");
	}

	/**
	 * Opens a document
	 * 
	 * @param filename, visible
	 */
	public void openDocInWord(String filename, Boolean visible) {
		System.out.println("MSWord_jsText: openDocInWord("+filename+", "+visible+") begin");
		
		//201203042135js:
		//To get a list of known applications from cmd:
		//
		//Entweder (dauert dann einige Zeit):
		//wmic
		//product get name
		//
		//Oder:
		//psinfo -s > software.txt
		//psinfo -s -c > software.csv
		//
		//For more info & possibilities:
		//http://superuser.com/questions/68611/get-list-of-installed-apps-from-windows-command-line
		//
		//Oder auch:
		//Start - Ausführen - dcomcnfg
		
		// Instantiate jacobObjWord
		//jacobObjWord = new ActiveXComponent("Word.Application");
		
		System.out.println("MSWord_jsText: openDocInWord(): Trying to instantiate jacobObjWord - please put that into a try..catch STA later on!");
		System.out.println("See: http://www.java-forum.org/allgemeine-java-themen/20775-datenuebergabe-java-ms-word-vorlage.html for an example.");
		System.out.println("   Falls das Programm hier stehenbleibt, fehlt der JavaVM der Pfad zu den jacob*.dll Dateien, z.B. -Djava.library.path=l:/Elexis/jacob"); 
		System.out.println("   Ich hab es auch beobachtet, nachdem ich die jacob-1.16-*.dll durch -1.18-*.dll ersetzt habe,"); 
		System.out.println("   beim Versuch, com.jsigle.MSWord_js wieder ins Elexis 2.1.7 20130605based einzubinden, und ebendiese"); 
		System.out.println("   Files im Projektordner com.jsigle.msword_js/lib/ ausgetauscht habe, obwohl /lib im classpath und die"); 
		System.out.println("   beiden dlls in den Dependencies drinstehen bzw. angekreuzt sind.");
		System.out.println("   Nachdem ich manuell in den bin-includes die Zeile lib/,\\ ergänzt habe, compilierte und startete es dann wieder."); 
		System.out.println("   ...UND BEIM NAECHSTEN START SCHON WIEDER NICHT. :-(   - hab's also wieder rausgetan."); 
		System.out.println("   Nun hab ich in MANIFEST.XML mal /lib durch /lib/jacob-1.18-x86.dll ersetzt. Damit geht's auch wieder..."); 
		System.out.println("   Ausserdem läuft es wahrscheinlich NICHT mit der Office Starter Version, die kann keine Automation/OLE"); 
		System.out.println("   Hier auch zum Office Version lookup: http://social.msdn.microsoft.com/Forums/nl/worddev/thread/2e235d2f-4fca-4d99-98c2-1dc8d92561cb"); 
		
		System.out.println("");

		System.out.println("MSWord_jsText: openDocInWord(): INFO INFO INFO INFO INFO INFO INFO INFO INFO INFO INFO INFO INFO INFO INFO INFO INFO INFO INFO INFO INFO INFO INFO INFO INFO INFO INFO INFO INFO INFO INFO INFO INFO");  
		System.out.println("MSWord_jsText: openDocInWord(): Man kann im VBA Editor (aus Word heraus) den Objektkatalog anzeigen, um eine Menge Infos über die verfügbaren Objekte, Felder etc. zu bekommen.");
		System.out.println("MSWord_jsText: openDocInWord(): Und dort auch Aktionen aufzeichnen, um den zugehörigen VBA Macro-Code zu sehen - der gibt zumindest Hinweise auf das, was hier entsprechend stehen muss.");
		System.out.println("MSWord_jsText: openDocInWord(): INFO INFO INFO INFO INFO INFO INFO INFO INFO INFO INFO INFO INFO INFO INFO INFO INFO INFO INFO INFO INFO INFO INFO INFO INFO INFO INFO INFO INFO INFO INFO INFO INFO");  

		System.out.println("");

		System.out.println("MSWord_jsText: openDocInWord(): TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO");
		System.out.println("MSWord_jsText: openDocInWord(): ToDo: Probably set boolean tSaveOnExit=true (check JACOB doc pages for proper Syntax and usage), and some NOTIFY CALLER ON EXIT (in time to allow re-importing...), too, and NOT deleteOnExit... - http://www.land-of-kain.de/docs/jacob/");
		System.out.println("MSWord_jsText: openDocInWord(): ToDo: Possibly include s.th. like: activeXApp.setProperty(\"DisplayAlerts\", new Variant(false)); to suppress annoying Normal.dot can't be saved messages etc. - http://readerim.iteye.com/blog/183005");
		System.out.println("MSWord_jsText: openDocInWord(): TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO");
		
		System.out.println("");

		System.out.println("MSWord_jsText: openDocInWord(): TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO");
		System.out.println("MSWord_jsText: openDocInWord(): ToDo: USE REASONABLE FILENAMES!!! / Configurable from MSWord_js Settings dialog, like in noatext_jsl in the meantime.");
		System.out.println("MSWord_jsText: openDocInWord(): TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO");

		System.out.println("");

		//ToDo: Warum wirft die Ersetzung von [Konsultation.Diagnose] (oder Diagnosen?) noch einen Fehler - Methode nicht definiert oder so?? in findOrReplace()
		System.out.println("MSWord_jsText: openDocInWord(): TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO");
		System.out.println("MSWord_jsText: openDocInWord(): ToDo: Warum wirft die Ersetzung von [Konsultation.Diagnose] (oder Diagnosen?) noch einen Fehler - Methode nicht definiert oder so?? in findOrReplace()");
		System.out.println("MSWord_jsText: openDocInWord(): ToDo: AHA. Das ist [Konsultation.Diagnose] -> ???Konsultation.Diagnose???, und das gibt es wohl auch früher schon nicht - sondern nur die [Patient.Diagnosen]");
		System.out.println("MSWord_jsText: openDocInWord(): TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO");
		
		System.out.println("MSWord_jsText: openDocInWord(): TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO");
		System.out.println("MSWord_jsText: openDocInWord(): ToDo: Am Ende der Ersetzerei in Briefen: Cursor bitte NICHT im Datumsfeld stehen lassen/dort alles markiert ist sowieso ungünstig,");
		System.out.println("MSWord_jsText: openDocInWord(): ToDo:    sondern in den Haupttext oder allenfalls an Anfang des Docs gehen!");
		System.out.println("MSWord_jsText: openDocInWord(): TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO");
		
		//Das vorinstallierte WordStarter reicht mit den folgenden beiden nicht:
		//jacobObjWord = new ActiveXComponent("WordStarter.Application");
		//jacobObjWord = new ActiveXComponent("Office.Application");
		jacobObjWord = new ActiveXComponent("Word.Application");
		//DAS LAG ABER wohl am fehlenden Pfad zu den jacob DLLs.
		//Habe jetzt beide DLLs nach L:\Elexis\jacob gelegt, und
		//in Run Configurations - Arguments - VM Arguments eine Zeile hinzugefügt:
		//-Djava.library.path=l:/Elexis/jacob
		//Nun läuft es nach obigem Befehl weiter, davor blieb es dort stehen.
		
		if (jacobObjWord==null)	System.out.println("MSWord_jsText: openDocInWord(): WARNING: jacobObjWord==null  (may be ok, because we're about to open() and assign stuff first...)");
		else					System.out.println("MSWord_jsText: openDocInWord(): jacobObjWord="+jacobObjWord.toString());


		
		// Assign a local word object

		System.out.println("MSWord_jsText: openDocInWord(): About to instantiate wordObject: wordObject = jacobObjWord.getObject();...");
		jacobWordObject = jacobObjWord.getObject();
		if (jacobWordObject==null)	System.out.println("MSWord_jsText: openDocInWord(): WARNING: wordObject==null");
		else						System.out.println("MSWord_jsText: openDocInWord(): wordObject="+jacobWordObject.toString());

		System.out.println("MSWord_jsText: openDocInWord(): About to Dispatch.put(wordObject, \"Visible\", new Variant(visible));...");
		// Create a Dispatch Parameter to hide the document that is opened
		Dispatch.put(jacobWordObject, "Visible", new Variant(visible));
		
		// Instantiate the jacobDocuments Property
		System.out.println("MSWord_jsText: openDocInWord(): Trying to instantiate the (List of all) jacobDocuments...");
		System.out.println("MSWord_jsText: openDocInWord(): About to Dispatch jacobDocuments = jacobObjWord.getProperty(\"Documents\").toDispatch();...");
		Dispatch jacobDocuments = jacobObjWord.getProperty("Documents").toDispatch();
		if (jacobDocuments==null)	System.out.println("MSWord_jsText: openDocInWord(): WARNING: jacobDocuments==null");
		else						System.out.println("MSWord_jsText: openDocInWord(): jacobDocuments="+jacobDocuments.toString());

		// Open a word jacobDocument, Current Active Document
		System.out.println("MSWord_jsText: openDocInWord(): Trying to load the jacobDocument: "+filename+"into jacobDocument...");
		System.out.println("MSWord_jsText: openDocInWord(): About to jacobDocument = Dispatch.call(jacobDocuments, \"Open\", filename).toDispatch();...");
		jacobDocument = Dispatch.call(jacobDocuments, "Open", filename).toDispatch();
		if (jacobDocument==null)	System.out.println("MSWord_jsText: openDocInWord(): ERROR: jacobDocument==null");
		else {
			System.out.println("MSWord_jsText: openDocInWord(): jacobDocument == "+jacobDocument.toString());

			System.out.println("MSWord_jsText: openDocInWord(): Trying to attach jacobWordEventHandler to jacobDocument...");
			System.out.println("MSWord_jsText: openDocInWord(): About to jacobWordEventHandler = new WordEventHandler();");
			jacobWordEventHandler = new WordEventHandler();
			System.out.println("MSWord_jsText: openDocInWord(): About to new DispatchEvents(jacobDocument, jacobWordEventHandler);");
			new DispatchEvents(jacobDocument, jacobWordEventHandler);

			if (jacobWordEventHandler==null)	System.out.println("MSWord_jsText: openDocInWord(): ERROR: jacobWordEventHandler==null");
			else 								System.out.println("MSWord_jsText: openDocInWord(): jacobWordEventHandler == "+jacobWordEventHandler.toString());
		}
		
	
		System.out.println("MSWord_jsText: openDocInWord() end");
	}

	
	
	
	
	
	/**
	 * Creates an instance of the VBA CustomDocumentProperties property
	 * 
	 */
	public void selectCustomDocumentProperitiesMode() {
		// Create CustomDocumentProperties and BuiltInDocumentProperties
		// properties
		System.out.println("MSWord_jsText: selectCustomDocumentProperitiesMode() begins");
		System.out.println("MSWord_jsText: selectCustomDocumentProperitiesMode(): About to custDocprops = Dispatch.get(jacobDocument, \"CustomDocumentProperties\").toDispatch();");
		jacobCustDocprops = Dispatch.get(jacobDocument, "CustomDocumentProperties").toDispatch();
		System.out.println("MSWord_jsText: selectCustomDocumentProperitiesMode() ends");
	}

	
	
	
	
	/**
	 * Creates an instance of the VBA BuiltInDocumentProperties property
	 * 
	 */
	public void selectBuiltinPropertiesMode() {
		// Create CustomDocumentProperties and BuiltInDocumentProperties
		// properties
		System.out.println("MSWord_jsText: selectBuiltinPropertiesMode() begins");
		System.out.println("MSWord_jsText: selectBuiltinPropertiesMode(): About to builtInDocProps = Dispatch.get(jacobDocument, \"BuiltInDocumentProperties\").toDispatch();");
		jacobBuiltInDocProps = Dispatch.get(jacobDocument, "BuiltInDocumentProperties").toDispatch();
		System.out.println("MSWord_jsText: selectBuiltinPropertiesMode() ends");
	}

	
	
	
	
	/**
	 * Closes a document
	 * 
	 */
	public void close() {
		System.out.println("MSWord_jsText: close() begins");
		// Close object
		System.out.println("MSWord_jsText: close(): About to Dispatch.call(jacobDocument, \"Close\");");
		Dispatch.call(jacobDocument, "Close");

		//Todo: close(): (a): Is it useful and SAFE to null jacobDocument here?
		System.out.println("MSWord_jsText: close(): About to jacobDocument = null");
		jacobDocument = null;
		System.out.println("MSWord_jsText: close() ends");
	}

	
	
	
	
	/**
	 * Quit word
	 * 
	 */
	public void quit() {
		System.out.println("MSWord_jsText: quit() begins");
		// Close object
		System.out.println("MSWord_jsText: quit(): About to jacobObjWord.invoke(\"Quit\", new Variant(false));");
		jacobObjWord.invoke("Quit", new Variant(false));
		System.out.println("MSWord_jsText: quit(): About to jacobObjWord = null");
		jacobObjWord = null;	
		
		//Todo: quit(): (a): Is it useful and SAFE to null these jacob... objects here?
		//Todo: quit(): (b): If jacobSelection were MSWord_jsText.java-global, we should also jacobSelction = null");
		System.out.println("MSWord_jsText: quit(): ToDo: TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO ");
	    System.out.println("MSWord_jsText: quit(): ToDo: If jacobSelection were MSWord_jsText.java-global, we should also jacobSelction = null");
	    System.out.println("MSWord_jsText: quit(): ToDo: TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO ");
	    //jacobSelection = null;
		
	    System.out.println("MSWord_jsText: quit(): About to jacobDocuments = null");
	    jacobDocuments = null; 

		System.out.println("MSWord_jsText: quit() ends");
	}

	
	
	
	
	/**
	 * Custom Property Name is passed in
	 * 
	 * @param cusPropName
	 * @return String - Custom property value
	 */
	public String getCustomProperty(String cusPropName) {
		System.out.println("MSWord_jsText: getCustomProperty(String cusPropName) begins");
		
		try {
			cusPropName = Dispatch.call(jacobCustDocprops, "Item", cusPropName)
					.toString();
		} catch (ComException e) {
			// Do nothing
			cusPropName = null;
		}

		System.out.println("MSWord_jsText: getCustomProperty(String cusPropName) ends, returning cusPropName...");
		return cusPropName;
	}

	
	
	
	
	/**
	 * Built In Property Name is passed in
	 * 
	 * @param builtInPropName
	 * @return String - Built in property value
	 */
	public String getBuiltInProperty(String builtInPropName) {
		System.out.println("MSWord_jsText: getBuiltInProperty(String builtInPropName) begins");
		
		try {
			builtInPropName = Dispatch.call(jacobBuiltInDocProps, "Item",
					builtInPropName).toString();
		} catch (ComException e) {
			// Do nothing
			builtInPropName = null;
		}

		System.out.println("MSWord_jsText: getBuiltInProperty(String builtInPropName) ends, returning builtInPropName...");
		return builtInPropName;
	}

	
	
	
	
	
	/**
	 * simple main program that gets some properties and prints them out
	 * (I have added some more actions for testing: opening a document, closing, opening another, cursor movement, locate Shapes/Tables/etc., insert table, quit word / close window...)
	 * 
	 * @param args
	 */
	public static void main(String[] args) {
		System.out.println("MSWord_jsText: main(String[] args): Simple main program DEMO begins - this should probably NOT run within Elexis plugin.");
		try {
			

			
			
			
			
			
			
			
			
			
			
			
			
/* MAIN TEST ROUTINES COMMENTED OUT	        
			
			//Open a document, in a visible window, close it again, and quit word (if no other window is still open)    
			
			// Instantiate the class
			MSWord_jsText jacobTest = new MSWord_jsText();

			// Open the word doc, visibly (second parameter in the jacTest.open)
			File doc = new File("samples/com/jacob/samples/office/TestDocument.doc");
			jacobTest.openDocInWord(doc.getAbsolutePath(), true);

			// Set Custom Properties
			jacobTest.selectCustomDocumentProperitiesMode();

			// Set Built In Properties
			jacobTest.selectBuiltinPropertiesMode();

			// Get custom Property Value
			String custValue = jacobTest.getCustomProperty("Information Source");

			// Get built in prroperty Property Value
			String builtInValue = jacobTest.getBuiltInProperty("Author");
			
			// Close Word Doc
			jacobTest.close();	//this includes: jacobDocument = null;
			
			// Quit Word
			jacobTest.quit();	//this includes: jacobObjWord = null; jacobDocuments = null;

			// Output data
			System.out.println("Document Val One: " + custValue);
			System.out.println("Document Author: " + builtInValue);

			
			System.out.println("");
			System.out.println("");
			
MAIN TEST ROUTINES COMMENTED OUT */	        

			
			
			
			
			
/* MAIN TEST ROUTINES COMMENTED OUT	        
			
			
		    //Try to move the cursor to the end, then to the beginning of the document - insert some text at either position
		    
			// Put the cursor to the top of the document (js added this)
			ActiveXComponent oWord = new ActiveXComponent("Word.Application");
		    oWord.setProperty("Visible", true);
		    ActiveXComponent oDocuments = oWord.getPropertyAsComponent("Documents");
					    					    
			//String sDir = "c:\\java\\jacob\\";
		    //String sInputDoc = sDir + "file_in.doc";
		    //String sInputDoc = "samples/com/jacob/samples/office/TestDocument.doc";
		    //String sInputDoc = "L:\\home\\jsigle\\workspace\\elexis-2.1.7-20130523\\elexis-bootstrap-js\\jsigle\\com.jsigle.msword_js\\samples\\com\\jacob\\samples\\office\\TestDocument.doc";
		    String sInputDoc = "L:/home/jsigle/workspace/elexis-2.1.7-20130523/elexis-bootstrap-js/jsigle/com.jsigle.msword_js/doc/Vorlage-scratch.doc";
			ActiveXComponent oDocument = oDocuments.invokeGetComponent("Open", new Variant(sInputDoc)); 
		    ActiveXComponent oSelection = oWord.getPropertyAsComponent("Selection");
		    ActiveXComponent oFind = oSelection.getPropertyAsComponent("Find");
		    
			oSelection.setProperty("Text", "InsertStuffatFileOpenCursorPosition");

			//Ok, now the following finally works. From: http://www.programering.com/a/MDN5YzMwATM.html 201609230504js
			Dispatch mySelection = Dispatch.get(oWord, "Selection").toDispatch();
			Dispatch.call(mySelection, "EndKey", new Variant(6));
			//Dispatch.call(mySelection, "setProperty", "Text", "InsertStuffAtDocumentEnd");
			oSelection.setProperty("Text", "InsertStuffatCursorPosAfterEndKey");

			Dispatch.call(mySelection, "HomeKey", new Variant(6));
			//Dispatch.call(mySelection, "setProperty", "Text", "InsertStuffAtDocumentHomeAgain");
			oSelection.setProperty("Text", "InsertStuffatCursorPosAfterHomeKey");
			
						
			System.out.println("");
			System.out.println("");
			
/* MAIN TEST ROUTINES COMMENTED OUT */	        

			
			
			
			
			
			//Textfelder von Word stecken in den Shapes...: 
			
			//Select a text field (i.e. not the main text, but the Adress-Field or Date-Field of the letter.
			//A normal programmed Find through JACOB will apparently NOT go through all fields
			//(and maybe neither through tables, buttons, whatever etc.) but only through the currently selected one.
			//An interactively called find will go through all fields.
			
	        
	        //Mal ein Makro aufgezeichnet, wie ich so ein Textfeld lösche:
	        /*
	        Sub Makro1()
	        '
	        ' Makro1 Makro
	        ' Makro aufgezeichnet am 23.09.2016 von Jörg M. Sigle
	        '
	            Selection.ShapeRange.Delete
	            Selection.Delete Unit:=wdCharacter, Count:=1
	        End Sub
	        */

/* MAIN TEST ROUTINES COMMENTED OUT	        
    
	        //So. Und Shape/Shapes sollte es dann wohl sein:
	        
	    	System.out.println("");
	    	System.out.println("Shapes test: About to list Shapes:");
	        try
	        {
	            Dispatch oShapes = Dispatch.get((Dispatch) oDocument, "Shapes").toDispatch();
	            int shpcnt = Dispatch.get(oShapes, "Count").getInt();
	            System.out.println("Shapes test: shpcnt="+shpcnt);
		        
	            for (int i = 0; i < shpcnt; i++)
	            {
	                Dispatch oShape = Dispatch.call(oShapes, "Item", new Variant(i + 1)).toDispatch(); //Existence of Shapes.Item confirmed via Word VBA Object Catalog
	                //Variant oRes = Dispatch.get(oshp, "Name");                
	                System.out.println("Shapes test: About to String oShape_text = Dispatch.get(otbl, \"Name\").toString();");
	                String oShapeName = Dispatch.get(oShape, "Name").toString();						//Existence of Shape.Name etc. confirmed via Word VBA Object Catalog

	                if(oShapeName != null)
	                {
	                	System.out.println("Shapes test: oShapeName "+ i+1 +": "+oShapeName);
	                    //Dispatch.put(oshp, "Range", new Variant(oShape_text));
	                }

	                Dispatch oShapeTextFrame = Dispatch.call(oShape, "TextFrame").toDispatch();

	                Integer oShapeTextFrameHasText = Dispatch.get(oShapeTextFrame, "HasText").toInt();
    
	                System.out.println("Shapes test: oShapeTextFrameHasText: "+oShapeTextFrameHasText);
	                
	                if (oShapeTextFrameHasText == -1) {
	                
	                	Dispatch oShapeTextFrameTextRange = Dispatch.call(oShapeTextFrame, "TextRange").toDispatch();
	                
	                	String oShapeTextFrameTextRangeText = Dispatch.get(oShapeTextFrameTextRange, "Text").toString();
	                    
	                	if(oShapeTextFrameTextRangeText   != null)
		                {
	                		//DAS LIEFERT ENDLICH DEN TEXT DES Shapes.Shape.TextFrame.TextRange.Text...
	                		
	                		System.out.println("Shapes test: oShapeTextFrameTextRangeText : "+oShapeTextFrameTextRangeText);
		                    //Dispatch.put(oshp, "Range", new Variant(oShape_text));
		                }
	                }                

	            }
	        }
	        catch (Exception e)
	        {
	            System.out.println(e);
	        }
		    
MAIN TEST ROUTINES COMMENTED OUT */	        

	        

			
/* MAIN TEST ROUTINES COMMENTED OUT	        
			
	        //Das funktioniert auch: Tables: - 2 Tabellen im Scratch-Dokument werden gefunden und berichtet.
	        
	    	System.out.println("");
	    	System.out.println("Tables test: About to list Tables:");
	        try
	        {
	            Dispatch oTables = Dispatch.get((Dispatch) oDocument, "Tables").toDispatch();
	            int tblcnt = Dispatch.get(oTables, "Count").getInt();
	            System.out.println("Tables test: tblcnt="+tblcnt);
		            for (int i = 0; i < tblcnt; i++)
	            {
	                Dispatch otbl = Dispatch.call(oTables, "Item", new Variant(i + 1)).toDispatch();	//Existence of Tables.Item confirmed via Word VBA Object Catalog
	                //Variant oRes = Dispatch.get(otbl, "Name");           	                
	                System.out.println("Tables test: About to String oTable_text = Dispatch.get(otbl, \"Columns\").toString();");
	                String oTable_cols = Dispatch.get(otbl, "Columns").toString();						//Existence of Table.Columns confirmed via Word VBA Object Catalog 	           
	                System.out.println("Tables test: Table.Columns "+i+": "+oTable_cols);
	                System.out.println("Tables test: About to String oTable_text = Dispatch.get(otbl, \"Rows\").toString();");
	                String oTable_rows = Dispatch.get(otbl, "Rows").toString();							//Existence of Table.Rows confirmed via Word VBA Object Catalog 	           
	                System.out.println("Tables test: Table.Columns "+i+": "+oTable_rows);
	            }
	        }
	        catch (Exception e)
	        {
	            System.out.println(e);
	        }

		
	        
	        
	        
	        
	     //Jetzt kommt noch alles mögliche an Objekten, was ich kurzfristig NICHT verarbeiten werde: 
	        
			//Ursprünglich in die richtige Richtung gekommen bin ich via:
			//https://www.tutorials.de/threads/mit-jacob-textmarken-von-word-vorlage-fuellen.291346/
			//dort erst mal adaptiert zum automatisierten Handling vorhandener Bookmarks:
			
	    	System.out.println("");
        	System.out.println("Bookmarks test: About to list Bookmarks:");
	        try
	        {
	            Dispatch oBookmarks = Dispatch.get((Dispatch) oDocument, "Bookmarks").toDispatch();
	            int bkmcnt = Dispatch.get(oBookmarks, "Count").getInt();
	            System.out.println("Bookmarks test: bkmcnt="+bkmcnt);
		            for (int i = 0; i < bkmcnt; i++)
	            {
	                Dispatch oBkm = Dispatch.call(oBookmarks, "Item", new Variant(i + 1)).toDispatch(); //Existence of Bookmarks.Item confirmed via Word VBA Object Catalog
	                //Variant oRes = Dispatch.get(oBkm, "Name");            

	                System.out.println("Bookmarks test: About to String oBookmark_text = Dispatch.get(otbl, \"Name\").toString();");
	                String oBookmark_text = Dispatch.get(oBkm, "Name").toString();						//Existence of Bookmark.Name confirmed via Word VBA Object Catalog
		               
	                if(oBookmark_text != null)
	                {
	                	System.out.println("Bookmarks test: Bookmark "+i+": "+oBookmark_text);
	                    //Dispatch.put(oBkm, "Range", new Variant(oBookmark_text));
	                }
	                
	            }
	        }
	        catch (Exception e)
	        {
	            System.out.println(e);
	        }
		
	        //Offenbar sind Fields NICHT das gleiche wie Textfelder = Textrahmen
	        
	    	System.out.println("");
	    	System.out.println("Fields test: About to list Fields:");
	        try
	        {
	            Dispatch oFields = Dispatch.get((Dispatch) oDocument, "Fields").toDispatch();
	            int fldcnt = Dispatch.get(oFields, "Count").getInt();
	            System.out.println("Fields test: fldcnt="+fldcnt);
		            for (int i = 0; i < fldcnt; i++)
	            {
	                Dispatch oFld = Dispatch.call(oFields, "Item", new Variant(i + 1)).toDispatch(); //Existence of Fields.Item confirmed via Word VBA Object Catalog
	                //Variant oRes = Dispatch.get(oFld, "Name");                
	                System.out.println("Fields test: About to String oField_text = Dispatch.get(otbl, \"Data\").toString();");
	                String oField_text = Dispatch.get(oFld, "Data").toString();						//Existence of Field.Data and Field.Code etc. confirmed via Word VBA Object Catalog
	               
	                if(oField_text != null)
	                {
	                	System.out.println("Fields test: Field "+i+": "+oField_text);
	                    //Dispatch.put(oFld, "Range", new Variant(oField_text));
	                }
	                
	            }
	        }
	        catch (Exception e)
	        {
	            System.out.println(e);
	        }
		
		
	        
	        
	        
	        //Offenbar sind FormFields NICHT das gleiche wie Textfelder = Textrahmen
	        
	    	System.out.println("");
	    	System.out.println("FormFields test: About to list FormFields:");
	        try
	        {
	            Dispatch oFormFields = Dispatch.get((Dispatch) oDocument, "FormFields").toDispatch();
	            int frmfldcnt = Dispatch.get(oFormFields, "Count").getInt();
	            System.out.println("FormFields test: frmfldcnt="+frmfldcnt);
		            for (int i = 0; i < frmfldcnt; i++)
	            {
	                Dispatch ofrmfld = Dispatch.call(oFormFields, "Item", new Variant(i + 1)).toDispatch(); //Existence of FormFields.Item confirmed via Word VBA Object Catalog
	                //Variant oRes = Dispatch.get(ofrmfld, "Name");                
	                System.out.println("FormFields test: About to String oFormField_text = Dispatch.get(otbl, \"Data\").toString();");
	                String oFormField_text = Dispatch.get(ofrmfld, "Data").toString();						//Existence of FormField.Data and FormField.Code etc. confirmed via Word VBA Object Catalog
	               
	                if(oFormField_text != null)
	                {
	                	System.out.println("FormFields test: FormField "+i+": "+oFormField_text);
	                    //Dispatch.put(ofrmfld, "Range", new Variant(oFormField_text));
	                }
	                
	            }
	        }
	        catch (Exception e)
	        {
	            System.out.println(e);
	        }
		
	        //Offenbar sind Frames NICHT das gleiche wie Textfelder = Textrahmen
	        
	    	System.out.println("");
	    	System.out.println("Frames test: About to list Frames:");
	        try
	        {
	            Dispatch oFrames = Dispatch.get((Dispatch) oDocument, "Frames").toDispatch();
	            int framcnt = Dispatch.get(oFrames, "Count").getInt();
	            System.out.println("Frames test: framcnt="+framcnt);
		            for (int i = 0; i < framcnt; i++)
	            {
	                Dispatch ofram = Dispatch.call(oFrames, "Item", new Variant(i + 1)).toDispatch(); //Existence of Frames.Item confirmed via Word VBA Object Catalog
	                //Variant oRes = Dispatch.get(ofram, "Name");                
	                System.out.println("Frames test: About to String oFrame_text = Dispatch.get(otbl, \"Data\").toString();");
	                String oFrame_text = Dispatch.get(ofram, "Data").toString();						//Existence of Frame.Data and Frame.Code etc. confirmed via Word VBA Object Catalog
	               
	                if(oFrame_text != null)
	                {
	                	System.out.println("Frames test: Frame "+i+": "+oFrame_text);
	                    //Dispatch.put(ofram, "Range", new Variant(oFrame_text));
	                }
	                
	            }
	        }
	        catch (Exception e)
	        {
	            System.out.println(e);
	        }
		        
	       
	        
	        
	        
	        //Offenbar sind TextColumns NICHT das gleiche wie Textfelder = Textrahmen
	        
	    	System.out.println("");
	    	System.out.println("TextColumns test: About to list TextColumns:");
	        try
	        {
	            Dispatch oTextColumns = Dispatch.get((Dispatch) oDocument, "TextColumns").toDispatch();
	            int txtcolcnt = Dispatch.get(oTextColumns, "Count").getInt();
	            System.out.println("TextColumns test: txtcolcnt="+txtcolcnt);
		            for (int i = 0; i < txtcolcnt; i++)
	            {
	                Dispatch otxtcol = Dispatch.call(oTextColumns, "Item", new Variant(i + 1)).toDispatch(); //Existence of TextColumns.Item confirmed via Word VBA Object Catalog
	                //Variant oRes = Dispatch.get(otxtcol, "Name");                
	                System.out.println("TextColumns test: About to String oTextColumn_text = Dispatch.get(otbl, \"Application\").toString();");
	                String oTextColumn_text = Dispatch.get(otxtcol, "Application").toString();						//Existence of TextColumn.Data and TextColumn.Code etc. confirmed via Word VBA Object Catalog
	               
	                if(oTextColumn_text != null)
	                {
	                	System.out.println("TextColumns test: TextColumn "+i+": "+oTextColumn_text);
	                    //Dispatch.put(otxtcol, "Range", new Variant(oTextColumn_text));
	                }
	                
	            }
	        }
	        catch (Exception e)
	        {
	            System.out.println(e);
	        }
		        
	        
	        //Mal schauen, ob AutoTextEntries besser passt...
	        
	    	System.out.println("");
	    	System.out.println("AutoTextEntries test: About to list AutoTextEntries:");
	        try
	        {
	            Dispatch oAutoTextEntries = Dispatch.get((Dispatch) oDocument, "AutoTextEntries").toDispatch();
	            int atecnt = Dispatch.get(oAutoTextEntries, "Count").getInt();
	            System.out.println("AutoTextEntries test: atecnt="+atecnt);
		            for (int i = 0; i < atecnt; i++)
	            {
	                Dispatch oate = Dispatch.call(oAutoTextEntries, "Item", new Variant(i + 1)).toDispatch(); //Existence of AutoTextEntries.Item confirmed via Word VBA Object Catalog
	                //Variant oRes = Dispatch.get(oate, "Name");                
	                System.out.println("AutoTextEntries test: About to String oAutoTextEntry_text = Dispatch.get(otbl, \"Name\").toString();");
	                String oAutoTextEntry_text = Dispatch.get(oate, "Name").toString();						//Existence of AutoTextEntry.Data and AutoTextEntry.Code etc. confirmed via Word VBA Object Catalog
	               
	                if(oAutoTextEntry_text != null)
	                {
	                	System.out.println("AutoTextEntries test: AutoTextEntry "+i+": "+oAutoTextEntry_text);
	                    //Dispatch.put(oate, "Range", new Variant(oAutoTextEntry_text));
	                }
	                
	            }
	        }
	        catch (Exception e)
	        {
	            System.out.println(e);
	        }
		
	        
	        
	        
	        
	        
	        
	        
	        
	        
	        
	        //Mal probieren, eine Tabelle einzufügen... - Das funktioniert. Sogar INNERHALB einer Tabelle,
	        //in deren erstem Feld von den vorigen Versuchen her der Cursor noch steht... 
	        
			//First, compute required numbers of rows and columns - as not all lines need to have the same length in Java, step through all of them...
			
	        String[][] contents = { 
                    {"Hupf", "Dupf"},
                    {"Wonz", "Donk"},
                    {"Doink", "Boink"},
                   };
	        
			int iRowCount = contents.length;
			int iColCount = 0;
			for (int row = 0; row < iRowCount; row++) {
				if (contents[row].length > iColCount) { iColCount = contents[row].length; }  
			}
				
			//Allocate a table of suitable size
			
			ActiveXComponent oTables = oDocument.getPropertyAsComponent("Tables");
			Variant oSelectionRange = oSelection.getProperty("Range");
			ActiveXComponent oTable = oTables.invokeGetComponent("Add", oSelectionRange, new Variant(iRowCount), new Variant(iColCount));
			oTable.invoke("AutoFormat", 16);
	        
	        
	        
/* MAIN TEST ROUTINES COMMENTED OUT */	        

			
			
			
			
			
			
			
			
			/* LEIDER KANN ICH DIE FUNKTIONALITÄT EINER METHODE MAL WIEDER NICHT EFFIZIENT TESTEN/WEITERENTWICKELN, WEIL - NATÜRLICH -
			 * findOrReplace() nicht einfach auf irgendein beliebiges schnell ad-hoc geladenes Dokument angewendet werden kann - 
			 * sondern dafür extra erst ein MSWord_jsText instantiiert werden muss - (n.B.: Wir mit diesem main(), was via Ctrl-Run
			 * schnell ausführbar wäre, aber leider schon innendrin!!!).
			 * 
			 * Wenn ich aber Einzelteile nicht in einem einfachen, schnell aufstartbaren Testsetting testen, verifizieren oder verbessern kann,
			 * dann kann ich auch irgendwelche grossartigen Konstrukte völlig vergessen, quo ad Durchschaubarkeit und Stabilität, und genau das
			 * ist ein Problem von Java und Elexis.
			 * 	
			 * Also dupliziere ich hier eine MENGE code aus findOrReplace(), was wieder fehlerträchtig und redundant ist.
			 * Und das ganze in mehreren Abschnitten, weil ich ja mehrere Aspekte von findOrReplace funktionierend machen muss...
			 * [...] 
			 */

			
			
/* TESTCODE FÜR SUCHEN mit findOrReplace etc.: COMMON PREPARATIONS, PREPARES A DOCUMENT FOR MULTIPLE PORTIONS LATER ON... */
			
			
		    //Try findOrReplace in a multi-placeholder-Shape
		    
			// Duplicate a bunch of fields and code from MSWord_jsText inside main() for quick testing only,
			// so that a LOCAL COPY!!! of portions of code from findOrReplace can be used for quick r&d unaltered below here.
			// That may be removed some time later, and yes, this surely adds to intransparency and confusion...
			final ActiveXComponent jacobObjWord;
			final Dispatch jacobCustDocprops;
			final Dispatch jacobBuiltInDocProps;
			final Dispatch jacobDocuments;
			final Dispatch jacobDocument;
			final Dispatch jacobWordObject;
			jacobObjWord = new ActiveXComponent("Word.Application");
		    jacobObjWord.setProperty("Visible", true);
		    jacobDocuments = jacobObjWord.getPropertyAsComponent("Documents");
		    String myInputDoc = "L:/home/jsigle/workspace/elexis-2.1.7-20130523/elexis-bootstrap-js/jsigle/com.jsigle.msword_js/doc/Vorlage-scratch2.doc";
			jacobDocument = ((ActiveXComponent) jacobDocuments).invokeGetComponent("Open", new Variant(myInputDoc)); 

			/*
			 * Hier hätte ich gerne zum Testen eingefügt: findOrReplace(searchpattern,null);
			 * Geht aber wieder mal nicht. Also baue ich halt alles nötige daraus nach, und bin mir bewusst,
			 * dass die schnell erreichbare Testumgebung hier mal wieder Abweichungen vom echten Betriebsumfeld haben wird,
			 * die die Ergebnisse fraglich übertragbar machen. Mal abgesehen von tonnenweise Codezeilen, die dadurch
			 * verdoppelt werden und die Unübersichtlichkeit weiter steigern. 
			 */
			
			if (jacobDocument == null) {
				System.out.println("MSWord_jsText: findOrReplace (test in main): TODO: please review the text of the error message, in German and English...");
				SWTHelper.showError("findOrReplace (test in main): doc IS NULL", "Fehler:","findOrReplace (test in main): Statt eines Dokuments wurde NULL übergeben - möglicherweise fehlt die Dokumentenvorlage, z.B. Rechnungsvorlage.");
				System.out.println("MSWord_jsText: findOrReplace (test in main): ERROR: doc IS NULL. THE FOLLOWING R&D DEBUGGING CODE WILL NOT RUN AS EXPECTED.");
				System.out.println("");
			}
			
			String pattern2="\\[?*\\]";		//Please note that this minimal R&D search pattern may locate *sub-portions* of large placeholders combined from SQL-queries and direct placeholders.
											//That's not a malfunction of the findOrReplace() code, but a limitation of this simple search pattern.
											//Here, I want it that way however, to reliably find [Liste] and the like, and reliably show whether Java-Jacob code works the right way, in the right document section.
			
			Integer jacobSearchResultInt = 0;
			Integer numberOfHits = 0;


			
			
/* TESTCODE FÜR SUCHEN IN KOPFZEILEN / analog: FUSSZEILEN - bitte auch vorgängiges allgeimeines freischalten... 
 *  			OUTDATED, RESULTS AND MORE ADVANCED MORE COMPLETE VERSION IN ACTUAL findOrResearch()


 			System.out.println("");
			System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to find/replace in Headers: ActiveDocument.Sections(i).Headers(wdHeaderFooterPrimary).Range.Text...");
			System.out.println("");
			
			ActiveXComponent jacobSelection = jacobObjWord.getPropertyAsComponent("Selection");
			ActiveXComponent jacobFind = jacobSelection.getPropertyAsComponent("Find");

			try {
	            Dispatch jacobSections = Dispatch.get((Dispatch) jacobDocument, "Sections").toDispatch();
	            int sectionsCount = Dispatch.get(jacobSections , "Count").getInt();
	            System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): sectionsCount="+sectionsCount);
		        
	            for (int i = 0; i < sectionsCount; i++) {
		            System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): i=="+i);

		            /*
		            //THIS FAILS:
		            //com.jacob.com.ComFailException: Can't map name to dispid: Headers
		            System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to Variant jacobSectionHeadersVariant = Dispatch.call(jacobSections, \"Item\", new Variant(i + 1));");
		            Variant jacobSectionHeadersVariant = Dispatch.call(jacobSections, "Headers", new Variant(i + 1)); //Sorry, das liefert tatsächlich Zugang zum Haupttext
		            */
/* Continued
 		            Variant jacobSectionVariant = Dispatch.call(jacobSections, "Item", new Variant(i + 1));
 
	                if (jacobSectionVariant == null)	System.out.println("MSWord_jsText: findOrReplace (Headers): WARNING: jacobSectionVariant IS NULL");
	                else 								System.out.println("MSWord_jsText: findOrReplace (Headers): jacobSectionVariant="+jacobSectionVariant.toString());

	                System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to Dispatch jacobSection = jacobSectionVariant.toDispatch();");
	                Dispatch jacobSection = jacobSectionVariant.toDispatch();

	                /*
		            System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to Variant jacobSectionHeaderVariant = Dispatch.call(jacobSections, \"Item\", new Variant(i + 1));");
		            Variant jacobSectionHeaderVariant = Dispatch.call(jacobSections, "Item", new Variant(i + 1)); //Sorry, das liefert tatsächlich letzendlich Zugang zum Haupttext
		            
		            if (jacobSectionHeaderVariant == null)	System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): WARNING: jacobSectionHeaderVariant IS NULL");
	                else 							System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): jacobSectionHeaderVariant="+jacobSectionHeaderVariant.toString());

		            System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to Dispatch jacobSectionHeader = jacobSectionHeaderVariant.toDispatch();");
	                Dispatch jacobSectionHeader = jacobSectionHeaderVariant.toDispatch();
	                
	                System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to Dispatch jacobSectionHeaderRange = Dispatch.call(jacobSectionHeader, \"Range\").toDispatch();");
	                Dispatch jacobSectionHeaderRange = Dispatch.call(jacobSectionHeader, "Range").toDispatch();
		            
		            String jacobSectionHeaderRangeText = Dispatch.get(jacobSectionHeaderRange, "Text").toString();
	                */

	                
	                //THIS FAILS:
	                //--------------Exception--------------
	                //MSWord_jsText: findOrReplace (SectionHeaders): Fehler bei enumeration von jacobSections...
	                //com.jacob.com.ComFailException: Can't map name to dispid: Item
	                //Variant jacobSectionHeaders = Dispatch.call(jacobSection, "Item", "wdHeaderFooterPrimary");
/* CONTINUED...
					Dispatch jacobSectionHeaders = Dispatch.get((Dispatch) jacobSection, "Headers").toDispatch();		               	           
	                if (jacobSectionHeaders == null)	System.out.println("MSWord_jsText: findOrReplace (Headers): WARNING: jacobSectionHeaders IS NULL");
		            int sectionHeadersCount = Dispatch.get(jacobSectionHeaders , "Count").getInt();
		            System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): sectionHeadersCount="+sectionHeadersCount);
			            
		            
		            //Wegen sectionHeadersCount = 3 wird der for Block 3x durchlaufen, das bewirkt aber wiederholt Ersetzungen im gleichen SectionHeaders block.
		            //VIELLEICHT Wirkt sich das nur dann nützlich aus, wenn "Erste Seite anders" oder "Linke / Rechte Seite anders" gewählt wurde, ich lasse es mal so.
		            for (int j = 0; j < sectionHeadersCount; j++) {
			            System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): Section i=="+i+"; SectionHeader j=="+j);			            
           
	                	do {
	                		//Hier muss das ganze Procedere, schon beginnend mit Dispatch jacobSectionHeaderVariant = ... in den do...while block (also noch mehr als bei (Shapes))
	                		//damit mehrere Platzhalter innerhalb eines SectionHeaders ersetzt werden.
	                		//Hab's ausprobiert, wenn erst ab jacobSectionHeaderRangeText = ... hier drin stand, wurden nur 3 Ersetzungen ausgeführt,
	                		//und zwar weil der for-Block bei sectionHeadersCount = 3 auch 3x durchlaufen wird.
	                		
				            System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to Variant jacobSectionHeadersVariant = Dispatch.call(jacobSectionHeaders, \"Item\", new Variant(i + 1));");
				            Variant jacobSectionHeaderVariant = Dispatch.call(jacobSectionHeaders, "Item", new Variant(i + 1)); //Sorry, das liefert tatsächlich letzendlich Zugang zum Haupttext
				            
				            if (jacobSectionHeaderVariant == null)	System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): WARNING: jacobSectionHeaderVariant IS NULL");
			                else 									System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): jacobSectionHeaderVariant="+jacobSectionHeaderVariant.toString());

				            System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to Dispatch jacobSectionHeader = jacobSectionHeaderVariant.toDispatch();");
			                Dispatch jacobSectionHeader = jacobSectionHeaderVariant.toDispatch();
			                
			                System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to Dispatch jacobSectionHeaderRange = Dispatch.call(jacobSectionHeader, \"Range\").toDispatch();");
			                Dispatch jacobSectionHeaderRange = Dispatch.call(jacobSectionHeader, "Range").toDispatch();

			                System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to String jacobSectionHeaderRangeText = Dispatch.get(jacobSectionHeaderRange, \"Text\").toString();");
		                	String jacobSectionHeaderRangeText = Dispatch.get(jacobSectionHeaderRange, "Text").toString();
		                    
		                    if (jacobSectionHeaderRangeText == null)	System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): WARNING: Section["+(i+1)+"].SectionHeader["+(j+1)+"].RangeText IS NULL");
		                    else {
		                    	System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): Section["+(i+1)+"].SectionHeader["+(j+1)+"].Range.Text="+jacobSectionHeaderRangeText);
	
		                    	//THIS WORKS, and causes the text in the SectionHeader to become selected            
		                    	System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to Variant jacobSectionHeaderRangeSelectVariant = Dispatch.call(jacobSectionHeaderRange, \"Select\");");
		                        Variant jacobSectionHeaderRangeSelectVariant = Dispatch.call(jacobSectionHeaderRange, "Select");
		                        if (jacobSectionHeaderRangeSelectVariant == null)	System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): WARNING: jacobSectionHeaderRangeSelectVariant IS NULL");
		                        else 	System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): jacobSectionHeaderRangeTextSelectVariant="+jacobSectionHeaderRangeSelectVariant.toString());                        
		                        
		                        //DAS HIER SCHEINT ZU GEHEN!!!!! ENDLICH!!!
		                        //(Analog der Methode, die über insertText()... cur ... pos hinwegegeholfen hat,
		                        // wobei ich mich einfach nicht darum kümmere, eine Selection als Eigenschaft des aktuellen Textfeldes anzusprechen -
		                        // sondern einfach eine Selection als Eigenschaft des ganzen Dokuments!)
		                        //
		                        //Object cur = jacobSelection.getObject();
		                        //Dispatch.call((Dispatch) cur, "MoveLeft");
		                        //
		                    	// UND DAS HIER AUCH - ist effektiv eine Kurzform davon:
		                        //Dispatch.call(jacobSelection, "MoveLeft");
	
		                    	System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to Variant jacobSectionHeaderRangeTextFindVariant = Dispatch.call(jacobSectionHeaderRangeText, \"Find\");");
		                        Variant jacobSectionHeaderRangeTextFindVariant = Dispatch.get(jacobSectionHeaderRange, "Find");
		                        if (jacobSectionHeaderRangeTextFindVariant == null)	System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): WARNING: jacobSectionHeaderRangeTextFindVariant IS NULL");
		                        else 								System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): jacobSectionHeaderRangeTextFindVariant="+jacobSectionHeaderRangeTextFindVariant.toString());                        
	
		                        System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to Dispatch jacobSectionHeaderRangeTextFind = jacobSectionHeaderRangeTextFindVariant.toDispatch();");
		                        Dispatch jacobSectionHeaderRangeTextFind = jacobSectionHeaderRangeTextFindVariant.toDispatch();
		                        if (jacobSectionHeaderRangeTextFind  == null)	System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): WARNING: jacobSectionHeaderRangeTextFind IS NULL");
		                        else 							System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): jacobSectionHeaderRangeTextFind="+jacobSectionHeaderRangeTextFind.toString());
	
		                       
		                        
		                        System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to try Dispatch.call(jacobSectionHeaderRangeTextFind, \"Text\", pattern2);... etc.");
		                		try {
		                			Dispatch.put(jacobSectionHeaderRangeTextFind, "Text", pattern2);
		                			Dispatch.put(jacobSectionHeaderRangeTextFind, "Forward", "True");
		                			Dispatch.put(jacobSectionHeaderRangeTextFind, "Format", "False");
		                			Dispatch.put(jacobSectionHeaderRangeTextFind, "MatchCase", "False");
		                			Dispatch.put(jacobSectionHeaderRangeTextFind, "MatchWholeWord", "False");
		                			Dispatch.put(jacobSectionHeaderRangeTextFind, "MatchByte", "False");
		                			Dispatch.put(jacobSectionHeaderRangeTextFind, "MatchAllWordForms", "False");
		                			Dispatch.put(jacobSectionHeaderRangeTextFind, "MatchSoundsLike", "False");
		                			Dispatch.put(jacobSectionHeaderRangeTextFind, "MatchWildcards", "True");	
		                		} catch (Exception ex) {
		                			ExHandler.handle(ex);
		                			System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): Fehler beim Ersetzen: Dispatch.put(jacobSectionHeaderRangeTextFind...);");
		                		}
		 
		                    	
		                		//The following block performs search-and-replace for each SectionHeaders text portion of the document.
		                		//An almost identical block is further above for the main text (without updating all the comments),
		                		//and will probably be added further below, for tables. 

		                        System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to try ... the actual search and replace block...");
		                		try {	 
	                				jacobSearchResultInt = 0;			//Reset this to 0 in each loop iteration, so that even when jacobFind.invoke() should NOT put a valid result into this variable
	                													//(might this happen? if an exception was throuwn?), we do definitely NOT get an endless loop, NOR a misguided attempt to replace text. 	
	                				if (pattern2 == null)	System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): ERROR: pattern2 IS NULL!");
	                				else 					System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): pattern2="+pattern2);
	                				
	                				try {
	                					System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to jacobSearchResultInt = Dispatch.call(jacobSectionHeaderRangeTextFind,\"Execute\").toInt();...");
	                					//jacobSearchResultInt = jacobFind.invoke("Execute").toInt();
	                					jacobSearchResultInt = Dispatch.call(jacobSectionHeaderRangeTextFind,"Execute").toInt();
	                					
	                					//Please note: Wenn ich erste jacobSearchResultInt = jacobSearchresultVariant.toInt() verwende, ist nacher auch der string = ""-1", sonst "true"" +
	                					System.out.println("MSWord_jsText: findOrReplace (test in main): jacobSearchResultInt="+jacobSearchResultInt);
	                					//System.out.println("Result: jacobInvokeResult.toString()="+jacobInvokeResult.toString());	//Returns true if match found, false if no match found
	                					//System.out.println("Result: jacobInvokeResult.toInt()="+jacobInvokeResult.toInt());		//Returns -1 if match found, 0 if no match found
	                					//System.out.println("Result: jacobInvokeResult.toError()="+jacobInvokeResult.toError());	//Throws java.lang.IllegalStateException: getError() only legal on Variants of type VariantError, not 3
	                				} catch (Exception ex) {
	                					ExHandler.handle(ex);
	                					//ToDo: Add precautions for pattern==null or pattern2==null...
	                					System.out.println("MSWord_jsText: findOrReplace (SectionHeaders):\nException caught.\n"+
	                					"Das verwendete Suchpattern funktioniert möglicherweise NICHT mit MS Word.\n"+
	                					"pattern2\t(Von msword_js für MS Word):\t"+pattern2+"\n"+		
	                					"Falls neue Suchpatterns hinzugefügt wurden, muss möglicherweise eine on-the-fly-Konversion in com.jsigle.msword_js.MSWord_jsText.java ergänzt werden.");
	                				}
	                				
	                				if (jacobSearchResultInt == -1) {		//Do ONLY attempt a replacement if there was a search hit. Otherwise, replacement text might be simply inserted at cursor position.
                						numberOfHits += 1;
                						System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): numberOfHits="+numberOfHits);
                						
                			/* SIMPLE GET, PRINTLN, REPLACE WITH CONSTANT STRING FOR TESTING ONLY - THIS WORKS :-) :-) :-)
                						//and very fine: especially, in Bern, [Brief.Datum] (or similar), only the [Brief.Datum] portion is replaced.
                						//This means, that the "Find" stuff actually works and controls the range that is influenced by the following commands. PUH.
                			*/
/* CONTINUED...
										System.out.println(Dispatch.get(jacobSectionHeaderRange, "Text").toString());
                						System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to Dispatch.put(jacobSectionHeaderRangeTextFind, \"Text\", new Variant(\"XRplX\");");
                						Dispatch.put(jacobSectionHeaderRange, "Text", new Variant("XRplX"));
                						
                						//GETESTET: OFFENBAR NICHT NÖTIG FÜR ERSETZUNGEN IM SectionHeader (=Kopfzeile)
                						
                						//THIS APPARENTLY WORKS!!!
                                        //System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to Dispatch.call(jacobSelection, \"MoveRight\");");
            	                        //Dispatch.call(jacobSelection, "MoveRight");
                                        //System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to Dispatch.call(jacobSelection, \"MoveLeft\");");
            	                        //Dispatch.call(jacobSelection, "MoveLeft");
            	                        //Nun. Das verschiebt zwar den Cursor ans Ende des Textes im SectionHeader - aber immer noch wird nur EIN Platzhalter ersetzt...
                                        //for (int j = 0; j < 100; j++) { Dispatch.call(jacobSelection, "MoveLeft"); }
                                        //Das verschiebt den Cursor weit nach oben zum Anfang des Textfeldes - aber trotzdem wird nur EIN Platzhalter ersetzt...
                						
                                        System.out.println("");
	                				}		                				
		                		} catch (Exception ex) {
		                			ExHandler.handle(ex);
		                			//ToDo: Add precautions for pattern==null or pattern2==null...
		                			System.out.println("MSWord_jsText: findOrReplace (SectionHeaders):\nFehler beim Suchen und Ersetzen im Header:\n"+"" +
		                					"Exception caught für:\n"+
		                					"pattern2\t(Von msword_js für MS Word):\t"+pattern2+"\n"+		
		                					"numberOfHits="+numberOfHits);
		                		}
		 
			                }
            			} while (jacobSearchResultInt == -1); 	                                
		         	} //for (int j = 0; i < SectionHeadersCount; i++)
		        } //for (int i = 0; i < SectionsCount; i++)
			} catch (Exception ex) {
				ExHandler.handle(ex);
				System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): Fehler bei enumeration von jacobSections...");
			}

			

			//We have three tasks here:
			//(a) switch back to page layout view from normal view
			//(b) bring back cursor from Header area to main text area
			//(c) ensure that a completely empty header line is NOT visible, taking NO y-space at all.
			//    which would usually be the case when a new Word file, or a template without header line, would be opened.
			//    But this is situation is (quite irreveribly) lost, when ...header.range is used to check whether it contains anything, or to find/replace placeholders therein. 
			//NOW TRYING HOW TO GET an empty KOPFZEILE COMPLETELY ABSENT AGAIN, NOT TAKING y-SPACE ANY MORE, not even for its single new-paragraph character:
			

			//Das funktioniert zwar, und löscht sogar vorher existierende Kopfzeilen mit Inhalt komplett - hilft jedoch nicht, den Platz der Kopfzeilen auch völlig freizugeben
			//(simplified, handling only sections(1), sectionHeaders(1):

            Dispatch jacobSections = Dispatch.get((Dispatch) jacobDocument, "Sections").toDispatch();
	        Variant jacobSectionVariant = Dispatch.call(jacobSections, "Item", new Variant(1));
            Dispatch jacobSection = jacobSectionVariant.toDispatch();
            Dispatch jacobSectionHeaders = Dispatch.get((Dispatch) jacobSection, "Headers").toDispatch();		               	           
            Variant jacobSectionHeaderVariant = Dispatch.call(jacobSectionHeaders, "Item", new Variant(1)); //Sorry, das liefert tatsächlich letzendlich Zugang zum Haupttext
            Dispatch jacobSectionHeader = jacobSectionHeaderVariant.toDispatch();
            Dispatch jacobSectionHeaderRange = Dispatch.call(jacobSectionHeader, "Range").toDispatch();
            Dispatch.call(jacobSectionHeaderRange, "Delete");
                       
			//Find/Replace in the Headers section switches Word to Normal view, Splits the MS Word Window into two sections, with the Headers Section in the lower one and the cursor there.
			//Now, we want to get back to Page Layout View:
			
			//THIS FAILS (all of them)
			//com.jacob.com.ComFailException: Can't map name to dispid: View
			//Variant jacobViewVariant = Dispatch.get(jacobObjWord, "View");
			//Variant jacobViewVariant = Dispatch.get(jacobDocument, "View");
			//Variant jacobViewVariant = Dispatch.get(jacobSelection, "View");
			//jacobObjWord.setProperty("View", 0);
			//jacobObjWord.getPropertyAsComponent("View");
			//jacobSelection.getPropertyAsComponent("View");
			
			//THIS FAILS, so we actually must go down step by step...
			//ActiveXComponent jacobActiveWindowActivePaneViewAXC = jacobObjWord.getPropertyAsComponent("ActiveWindow.ActivePane");

			System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to ActiveXComponent jacobDocumentAXC = jacobSelection.getPropertyAsComponent(\"Document\");");
			ActiveXComponent jacobDocumentAXC = jacobSelection.getPropertyAsComponent("Document");
			
			System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to ActiveXComponent jacobActiveWindowAXC = jacobObjWord.getPropertyAsComponent(\"ActiveWindow\");");
			ActiveXComponent jacobActiveWindowAXC = jacobObjWord.getPropertyAsComponent("ActiveWindow");
			
			System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to ActiveXComponent jacobActiveWindowViewAXC = jacobActiveWindowAXC.getPropertyAsComponent(\"View\");");
			ActiveXComponent jacobActiveWindowViewAXC = jacobActiveWindowAXC.getPropertyAsComponent("View");

			//If ActiveWindow.View.SplitSpecial <> wdPaneNone Then
	        //  ActiveWindow.Panes(2).Close
			//End If
			
			System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to ActiveXComponent jacobActiveWindowPanesAXC = jacobActiveWindowAXC.getPropertyAsComponent(\"Panes\");");
			ActiveXComponent jacobActiveWindowPanesAXC = jacobActiveWindowAXC.getPropertyAsComponent("Panes");

			try {			
			System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): jacobActiveWindowViewAXC.getProperty(\"SplitSpecial\") == "+jacobActiveWindowViewAXC.getProperty("SplitSpecial"));
			if (jacobActiveWindowViewAXC.getProperty("SplitSpecial").toInt() != 0) {
			
				System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to Variant jacobActiveWindowPane2Variant = Dispatch.call(jacobActiveWindowPanesAXC, \"Item\", new Variant(2));");
				Variant jacobActiveWindowPane2Variant = Dispatch.call(jacobActiveWindowPanesAXC, "Item", new Variant(2));
				System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to Dispatch.call(jacobActiveWindowPane2Variant.toDispatch(), \"Close\");");
	            Dispatch.call(jacobActiveWindowPane2Variant.toDispatch(), "Close");
			}
			
			} catch (Exception ex) {
				ExHandler.handle(ex);
				System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): Fehler bei ActiveWindow.Panes(2).Close...");
			}

			//System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to ActiveXComponent jacobActiveWindowActivePaneAXC = jacobActiveWindowAXC.getPropertyAsComponent(\"ActivePane\");");
			//ActiveXComponent jacobActiveWindowActivePaneAXC = jacobActiveWindowAXC.getPropertyAsComponent("ActivePane");
			
			//System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to ActiveXComponent jacobActiveWindowActivePaneViewAXC = jacobActiveWindowActivePaneAXC.getPropertyAsComponent(\"View\");");
			//ActiveXComponent jacobActiveWindowActivePaneViewAXC = jacobActiveWindowActivePaneAXC.getPropertyAsComponent("View");

			
			
			//This returns 1 when the normal view is active, with the Headers in the lower portion of the window and the main document text in the upper portion of the window.
			//This returns 3 when the page layout view is active
			System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): jacobActiveWindowViewAXC.getPropertyAsComponent(\"Type\") == "+jacobActiveWindowViewAXC.getPropertyAsInt("Type"));
			//System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): jacobActiveWindowActivePaneViewAXC.getPropertyAsComponent(\"Type\") == "+jacobActiveWindowActivePaneViewAXC.getPropertyAsInt("Type"));			
			
			//THIS ALL SHOULD WORK BUT DOESN'T. IT works INTERACTIVELY, THOUGH! (truly Microsoft...) 
			
			try {
				//THIS FAILS:
				//com.jacob.com.ComFailException: Invoke of: Type
				//Source: Microsoft Word
				//Description: Diese Eigenschaft oder Methode ist auf diesem System nicht verfügbar.
				//jacobActiveWindowActivePaneViewAXC.setProperty("Type",3);
				
				//THIS SUCCEEDS (i.e. we're back in Page Layout View afterwards, with the Headers section active, and the cursor in the Headers section.
				//
				//ONLY IF WE HAD ALLOCATED ActiveXComponent jacobActiveWindowActivePaneAXC = ... jacobActiveWindowActivePaneViewAXC = ... above, this THROWS AN EXCEPTION, however NOT RED, but black, probably informative:
				//com.jacob.com.ComFailException: Invoke of: Type
				//Source: Microsoft Word
				//Description: Objekt wurde gelöscht.
				//
				//So I don't allocate the components ActivePane and below, and it all works without any error :-)
				System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to jacobActiveWindowViewAXC.setProperty(\"Type\",3);");
				jacobActiveWindowViewAXC.setProperty("Type",3);
			} catch (Exception ex) {
				ExHandler.handle(ex);
				System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): Fehler bei Setzen von View.Type=3...");
			}
			
			
			System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): jacobActiveWindowViewAXC.getPropertyAsComponent(\"Type\") == "+jacobActiveWindowViewAXC.getPropertyAsInt("Type"));
			//System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): jacobActiveWindowActivePaneViewAXC.getPropertyAsComponent(\"Type\") == "+jacobActiveWindowActivePaneViewAXC.getPropertyAsInt("Type"));

			
			//JETZT ist Word zwar wieder im Page Layout View (bzw.: wdPrintView), aber immer noch sind Kopf-/Fusszeilen aktiv und der Cursor im Kopfzeilenbereich.
			//Weiteres Suchen-/Ersetzen findet dann ebenfalls nur dort statt.
			
			//N.B.: Must be in Page Layout View = wdPrintView in order to change SeekView Setting below.
			
			//Ein aufgezeichnetes Makro, welches in dieser Situation Menü: Ansicht - Kopf-/Fusszeilen (aus) entspricht, enthält:
			/*
			Sub Makro4()
			'
			' Makro4 Makro
			' Makro aufgezeichnet am 28.09.2016 von Jörg M. Sigle
			'
			    If ActiveWindow.View.SplitSpecial <> wdPaneNone Then
			        ActiveWindow.Panes(2).Close
			    End If
			    If ActiveWindow.ActivePane.View.Type = wdNormalView Or ActiveWindow. _
			        ActivePane.View.Type = wdOutlineView Then
			        ActiveWindow.ActivePane.View.Type = wdPrintView
			    End If
			    ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
			End Sub
			*/
			
			/*
			//Das bringt mich zwar wieder an den Haupttext-Anfang zurück.
			//Allerdings ist die Kopfzeile auch in der Layout-Ansicht immer noch oben angezeigt, auch wenn sie leer ist,
			//mit einem New-Paragraph-Zeichen drin, und y-Platzverbrauch.
			try {
				System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to jacobActiveWindowViewAXC.setProperty(\"SeekView\",0);");
				jacobActiveWindowViewAXC.setProperty("SeekView",0);
			} catch (Exception ex) {
				ExHandler.handle(ex);
				System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): Fehler bei Setzen von View.SeekView=0...");
			}

			//HILFT NICHTS
			//Dispatch.call(jacobSelection, "MoveRight");
			//Dispatch.call(jacobSelection, "MoveLeft");

			//Das bringt mich in die Kopfzeile.
			try {
				System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to jacobActiveWindowViewAXC.setProperty(\"SeekView\",1);");
				jacobActiveWindowViewAXC.setProperty("SeekView",1);
			} catch (Exception ex) {
				ExHandler.handle(ex);
				System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): Fehler bei Setzen von View.SeekView=1...");
			}

			//Thread.sleep(200);

            //Das bringt mich in die Kopfzeile.
			try {
				System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to jacobActiveWindowViewAXC.setProperty(\"SeekView\",0);");
				jacobActiveWindowViewAXC.setProperty("SeekView",0);
			} catch (Exception ex) {
				ExHandler.handle(ex);
				System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): Fehler bei Setzen von View.SeekView=0...");
			}
			
            //Thread.sleep(200);
            
			//Das bringt mich in die Kopfzeile.
			try {
				System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to jacobActiveWindowViewAXC.setProperty(\"SeekView\",1);");
				jacobActiveWindowViewAXC.setProperty("SeekView",1);
			} catch (Exception ex) {
				ExHandler.handle(ex);
				System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): Fehler bei Setzen von View.SeekView=1...");
			}

			//Thread.sleep(200);

			/* Mal im Word VBA kurz nachgesehen, ob die Konstanten stimmen, und gefunden: */
/*			int wdSeekCurrentPageHeader=9;
			int wdSeekMainDocument=0;
			int wdNormalView=1;
			int wdOutlineView=2;
			int wdPrintView=3;
			
			//Das resultierende Verhalten ist aber genau gleich wie bei 1 und 0.
			
			//Das bringt mich in die Kopfzeile.
			try {
				System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to jacobActiveWindowViewAXC.setProperty(\"SeekView\",wdSeekCurrentPageHeader);  i.e. = 9");
				jacobActiveWindowViewAXC.setProperty("SeekView",wdSeekCurrentPageHeader);
			} catch (Exception ex) {
				ExHandler.handle(ex);
				System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): Fehler bei Setzen von View.SeekView = wdSeekCurrentPageHeader = 9...");
			}

			//Thread.sleep(200);

			//Das bringt mich in die Kopfzeile.
			try {
				System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to jacobActiveWindowViewAXC.setProperty(\"SeekView\",wdSeekMainDocument);  i.e. = 0");
				jacobActiveWindowViewAXC.setProperty("SeekView",wdSeekMainDocument);
			} catch (Exception ex) {
				ExHandler.handle(ex);
				System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): Fehler bei Setzen von View.SeekView = wdSeekMainDocument = 0...");
			}

			//Thread.sleep(200);
			
			/*
			Dispatch jacobSections = Dispatch.get((Dispatch) jacobDocument, "Sections").toDispatch();
	        Variant jacobSectionVariant = Dispatch.call(jacobSections, "Item", new Variant(1));
            Dispatch jacobSection = jacobSectionVariant.toDispatch();
            Dispatch jacobSectionHeaders = Dispatch.get((Dispatch) jacobSection, "Headers").toDispatch();		               	           
            Variant jacobSectionHeaderVariant = Dispatch.call(jacobSectionHeaders, "Item", new Variant(1)); //Sorry, das liefert tatsächlich letzendlich Zugang zum Haupttext
            Dispatch jacobSectionHeader = jacobSectionHeaderVariant.toDispatch();
            Dispatch jacobSectionHeaderRange = Dispatch.call(jacobSectionHeader, "Range").toDispatch();
            Dispatch.call(jacobSectionHeaderRange, "Delete");
            */

			//Das Folgende hilft auch alles nichts...
			/*
			jacobSectionVariant.safeRelease();
			jacobSection.safeRelease();
			jacobSectionHeaderVariant.safeRelease();
            jacobSectionHeader.safeRelease();
            jacobSectionHeaderRange.safeRelease();
			
			jacobSectionVariant=null;
			jacobSection=null;
			jacobSectionHeaderVariant=null;
            jacobSectionHeader=null;
            jacobSectionHeaderRange=null;
            */
			
			//HILFT NICHTS
			//Dispatch.call(jacobSelection, "MoveRight");
			//Dispatch.call(jacobSelection, "MoveLeft");

			//N.B.: Bei Auswahl von 2 kommt:
			//MSWord_jsText: findOrReplace (SectionHeaders): Fehler bei Setzen von View.SeekView=2...
			//		com.jacob.com.ComFailException: Invoke of: SeekView
			//		Source: Microsoft Word
			//		Description: Die angeforderte Ansicht ist nicht verfügbar.
			//try {
			//	System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to jacobActiveWindowViewAXC.setProperty(\"SeekView\",2);");
			//	jacobActiveWindowViewAXC.setProperty("SeekView",2);
			//} catch (Exception ex) {
			//	ExHandler.handle(ex);
			//	System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): Fehler bei Setzen von View.SeekView=2...");
			//}
			
			//SplitSpecial ist nun 0, obwohl die leere Kopfzeile noch über dem Haupttext angezeigt wird.
			//Witzigerweise: Wenn ich manuell im Menü Ansicht - Kopf-und-Fusszeile und nochmal Ansicht - Kopf-und-Fusszeile wähle - ist sie verschwunden...
            
            /*
			try {			
			System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): jacobActiveWindowViewAXC.getProperty(\"SplitSpecial\") == "+jacobActiveWindowViewAXC.getProperty("SplitSpecial"));
			if (jacobActiveWindowViewAXC.getProperty("SplitSpecial").toInt() != 0) {
			
				System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to Variant jacobActiveWindowPane2Variant = Dispatch.call(jacobActiveWindowPanesAXC, \"Item\", new Variant(2));");
				Variant jacobActiveWindowPane2Variant = Dispatch.call(jacobActiveWindowPanesAXC, "Item", new Variant(2));
				System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to Dispatch.call(jacobActiveWindowPane2Variant.toDispatch(), \"Close\");");
	            Dispatch.call(jacobActiveWindowPane2Variant.toDispatch(), "Close");
			}
			
			} catch (Exception ex) {
				ExHandler.handle(ex);
				System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): Fehler bei ActiveWindow.Panes(2).Close...");
			}
			*/

			/*
			//Also brauche ich DANACH NOCMALS die Umschaltung zu Page Layout View:
			//N.B.: Ich hab auch versucht, die SeekView = 0 Umschaltung oben VOR das erstmalige Type = 3 zu setzen. Wirft eine Exception und funktioniert gar nicht.
			try {
				System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to jacobActiveWindowViewAXC.setProperty(\"Type\",3);");
				jacobActiveWindowViewAXC.setProperty("Type",3);
			} catch (Exception ex) {
				ExHandler.handle(ex);
				System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): Fehler bei Setzen von View.Type=3...");
			}
			*/
				
			//YEP. JETZT sind wir wieder vollständig zurück... :-)
			


			
			
			
			//During R&D and testing, we may want to close the Word Window again and only look at the console log output for Errors and Exceptions: 
			//Dispatch.call(jacobObjWord, "Quit");
			
			
			
/*ENDE SUCHEN IN HEADER SECTION UND VERSUCH: WIEDERHERSTELLEN EINER NICHT SICHTBAREN KOPFZEILE DANACH (geht nicht (skriptbar)) */ 			
			
			
			
			
			
			
			
/* TESTCODE FÜR SUCHEN IM HAUPTTEIL - insbesondere auch zum Feststellen, ob das MoveLeft/MoveRight nach einem searchReplace nötig ist			
			ActiveXComponent jacobSelection = jacobObjWord.getPropertyAsComponent("Selection");
			ActiveXComponent jacobFind = jacobSelection.getPropertyAsComponent("Find");

			try {
				jacobFind.setProperty("Text", pattern2);
				jacobFind.setProperty("Forward", "True");
				jacobFind.setProperty("Format", "False");
				jacobFind.setProperty("MatchCase", "False");
				jacobFind.setProperty("MatchWholeWord", "False");
				jacobFind.setProperty("MatchByte", "False");
				jacobFind.setProperty("MatchAllWordForms", "False");
				jacobFind.setProperty("MatchSoundsLike", "False");
				jacobFind.setProperty("MatchWildcards", "True");	
			} catch (Exception ex) {
				ExHandler.handle(ex);
				System.out.println("MSWord_jsText: findOrReplace (test in main): Fehler bei jacobSelection.setProperty(\"...\", \"...\");");
			}
		
			try {	 
				do {
					jacobSearchResultInt = 0;			//Reset this to 0 in each loop iteration, so that even when jacobFind.invoke() should NOT put a valid result into this variable
														//(might this happen? if an exception was throuwn?), we do definitely NOT get an endless loop, NOR a misguided attempt to replace text. 	
					if (pattern2 == null)	System.out.println("MSWord_jsText: findOrReplace (test in main): ERROR: pattern2 IS NULL!");
					else 					System.out.println("MSWord_jsText: findOrReplace (test in main): pattern2="+pattern2);
					
					try {
						System.out.println("MSWord_jsText: findOrReplace (test in main): About to jacobFind.invoke(\"Execute\");");
						jacobSearchResultInt = jacobFind.invoke("Execute").toInt();
						System.out.println("MSWord_jsText: findOrReplace (test in main): jacobSearchResultInt="+jacobSearchResultInt);
					} catch (Exception ex) {
						ExHandler.handle(ex);
						//ToDo: Add precautions for pattern==null or pattern2==null...
						System.out.println("MSWord_jsText: findOrReplace (Haupttext):\nException caught.\n"+
						"Das verwendete Suchpattern funktioniert möglicherweise NICHT mit MS Word.\n"+
						"pattern2\t(Von msword_js für MS Word):\t"+pattern2+"\n"+		
						"Falls neue Suchpatterns hinzugefügt wurden, muss möglicherweise eine on-the-fly-Konversion in com.jsigle.msword_js.MSWord_jsText.java ergänzt werden.");
					}
					
					if (jacobSearchResultInt == -1) {		//Do ONLY attempt a replacement if there was a search hit. Otherwise, replacement text might be simply inserted at cursor position.
							numberOfHits += 1;
							System.out.println("MSWord_jsText: findOrReplace (test in main): numberOfHits="+numberOfHits);
							
                			/* SIMPLE GET, PRINTLN, REPLACE WITH CONSTANT STRING FOR TESTING ONLY - THIS WORKS :-) :-) :-)
    						//and very fine: especially, in Bern, [Brief.Datum] (or similar), only the [Brief.Datum] portion is replaced.
    						//This means, that the "Find" stuff actually works and controls the range that is influenced by the following commands. PUH.
    			*/
/* CONTINUED...
							System.out.println("MSWord_jsText: findOrReplace (test in main): Found: "+jacobSelection.getProperty("Text").toString());						
							System.out.println("MSWord_jsText: findOrReplace (test in main): About to jacobSelection.setProperty(\"Text\", \"XRplX\");");						
							jacobSelection.setProperty("Text", "XRplX");		//Das sollte "Replaced" anstelle des Suchtexts einfügen.

							//THIS CODE WOULD BE NEEDED FOR REPLACEMENT IN Shapes:
							//System.out.println(Dispatch.get(jacobShapeTextFrameTextRange, "Text").toString());
    						//Dispatch.put(jacobShapeTextFrameTextRange, "Text", new Variant("XRplX"));
							
							//GETESTET: DAS MoveRight; MoveLeft; IST WIRKLICH NÖTIG, UM IM HAUPTTEXT ZUVERLÄSSIG ALLE PLATZHALTER ZU ERSETZEN. NICHT NÖTIG IN SHAPES.
							
							//Moving right removes the highlighting and places the cursor to the right of the replaced text.
							//This is required, as otherwise, successive find/replace occurances may become confused.
							System.out.println("MSWord_jsText: findOrReplace (test in main): About to jacobSelection.invoke(\"MoveRight\");");
							jacobSelection.invoke("MoveRight");
							
							//However, it's also necessary to go back to the left by one step afterwards,
							//or otherwise, a seamlessly following [placeholders][seamlesslyFollowingPlaceholder] will NOT be found.
							//The MoveRight - MoveLeft sequence has the effect that the selection = highlighting is removed from the inserted text.
							System.out.println("MSWord_jsText: findOrReplace (test in main): About to jacobSelection.invoke(\"MoveLeft\");");
							jacobSelection.invoke("MoveLeft");

							System.out.println("");
					}
					
				} while (jacobSearchResultInt == -1); 
			} catch (Exception ex) {
				ExHandler.handle(ex);
				//ToDo: Add precautions for pattern==null or pattern2==null...
				System.out.println("MSWord_jsText: findOrReplace:\nFehler beim Suchen und Ersetzen im Haupttext:\n"+"" +
						"Exception caught für:\n"+
						"pattern2\t(Von msword_js für MS Word):\t"+pattern2+"\n"+		
						"numberOfHits="+numberOfHits);
			}
			
/* Ende TESTCODE FÜR SUCHEN IM HAUPTTEIL - insbesondere auch zum Feststellen, ob das MoveLeft/MoveRight nach einem searchReplace nötig ist*/			

			
			
			
			
			
			
			
			
			
			
/* TESTCODE FÜR SUCHEN IN SHAPES, esp. wie man dort MoveLeft oder MoveRight implementieren kann
 * und mehrere [Platzhalter] innerhalb desselben Shapes verarbeiten kann (gelöst): 
 * THIS IS HIGHLY INCOMPLETE, MUCH FINAL CODE REMOVED FOR FURTHER TESTING & DEVELOPMENT. */
 
		//The following code performs search-and-replace in text within all Shapes (Textfelder),
		//e.g. the Adressfeld and Datumsfeld in my medical letter templates:
        
		System.out.println("");
		System.out.println("MSWord_jsText: findOrReplace (Shapes): About to list Shapes, and find/replace in each Shape.TextFrame.TextRange.Text:");
		System.out.println("");

		//Identify and process each available Shape...
		
		try {
            Dispatch jacobShapes = Dispatch.get((Dispatch) jacobDocument, "Shapes").toDispatch();
            int shapesCount = Dispatch.get(jacobShapes , "Count").getInt();
            System.out.println("MSWord_jsText: findOrReplace (Shapes): shapesCount="+shapesCount);
	        
            
            //ADD A SHAPE JUST FOR TESTING
            
			System.out.println("MSWord_jsText: insertTextAt: About to Variant jacobShapeVariant = Dispatch.call(jacobShapes, \"AddTextBox\", 1, "+100+" pt, "+200+" pt, "+200+" pt, "+200+" pt); (MS Word coordinates)");
            Variant jacobShapeVariant = Dispatch.call(jacobShapes, "AddTextBox", 1, 100, 200, 200, 200);
            
            System.out.println("MSWord_jsText: insertTextAt: About to Dispatch jacobShape = jacobShapeVariant.toDispatch();");
            Dispatch jacobShape = jacobShapeVariant.toDispatch();

            
            System.out.println("MSWord_jsText: findOrReplace (Shapes): INFO: Workaround for Tarmedrechnung_xx templates.");
        	//Dispatch.call(jacobShape, "ZOrder", 0);
            //Dispatch.call(jacobShapeTextFrame, "ZOrder", 0);
            
            System.out.println("MSWord_jsText: findOrReplace (Shapes): INFO: Workaround for Tarmedrechnung_xx templates. About to Dispatch jacobShapeWrapFormat = Dispatch.get(jacobShape, \"WrapFormat\").toDispatch();");
            Dispatch jacobShapeWrapFormat = Dispatch.get(jacobShape, "WrapFormat").toDispatch();
            System.out.println("MSWord_jsText: findOrReplace (Shapes): INFO: Workaround for Tarmedrechnung_xx templates. About to Dispatch.put(jacobShapeWrapFormat, \"AllowOverlap\", new Variant (true));");
            Dispatch.put(jacobShapeWrapFormat, "AllowOverlap", new Variant (true)); //Das WrapFormat.AllowOverlap bezieht sich auf Overlapping durch andere Shapes.
            
            int wdWrapSquare = 0;		//Das WrapFormat.Type bezieht sich darauf, wie Text das Shape um- oder über- oder unter-fliesst.
            int wdWrapTight = 1;
            int wdWrapThrough = 2;
            int wdWrapNone = 3;
            int wdWrapTopBottom = 4;
            int wdWrapInline = 7;
            Dispatch.put(jacobShapeWrapFormat, "Type", wdWrapNone); //wdWrapNone ist am ehesten, was ich für Text-Shapes am Beginn der Tarmedrechnung_Sx etc. brauche,
            														//damit die Leerzeilen zwischen [Titel] und Leistungs-Tabelle nicht mehr UNTER, sondern HINTER
            														//den Rechtecken ganz oben stehen.
            
            
            //Dispatch.call(jacobShape,"ConvertToInlineShape");
            //Dispatch.call(jacobShapeWrapFormat,"ConvertToInlineShape");
             
            //System.out.println("MSWord_jsText: findOrReplace (Shapes): INFO: Workaround for Tarmedrechnung_xx templates. About to Dispatch.call(jacobShape, \"ZOrder\", 4);");
            //Dispatch.call(jacobShape, "ZOrder", 4);

            
            /*
            for (int i = 0; i < shapesCount; i++) {
            	//Dispatch jacobShape = Dispatch.call(jacobShapes, "Item", new Variant(i + 1)).toDispatch();
            	//The above one-step call + conversion used to fail, a two step approach using a Variant type intermediate storage appears to work. 20160924js
            	//Same behaviour observed for another step further below.
                System.out.println("MSWord_jsText: findOrReplace (Shapes): About to Variant jacobShapeVariant = Dispatch.call(jacobShapes, \"Item\", new Variant("+(i + 1)+"));");
                jacobShapeVariant = Dispatch.call(jacobShapes, "Item", new Variant(i + 1));
                if (jacobShapeVariant == null)	System.out.println("MSWord_jsText: findOrReplace (Shapes): WARNING: jacobShapeVariant IS NULL");
                else 							System.out.println("MSWord_jsText: findOrReplace (Shapes): jacobShapeVariant="+jacobShapeVariant.toString());
                System.out.println("MSWord_jsText: findOrReplace (Shapes): About to Dispatch jacobShape = jacobShapeVariant.toDispatch();");
                jacobShape = jacobShapeVariant.toDispatch();        
                
                System.out.println("MSWord_jsText: findOrReplace (Shapes): About to String jacobShapeName = Dispatch.get(jacobShape, \"Name\").toString();");
                String jacobShapeName = Dispatch.get(jacobShape, "Name").toString();
                if (jacobShapeName == null)	System.out.println("MSWord_jsText: findOrReplace (Shapes): WARNING: Shape["+(i+1)+"].Name IS NULL");
                else 						System.out.println("MSWord_jsText: findOrReplace (Shapes): Shape["+(i+1)+"].Name="+jacobShapeName);
                
                
                
                //ToDo: This is a Workaround. Should probably make new Tarmedrechnung_xx templates for Word anyway.
                //For Tarmedrechnung_xx templates: Set the ZORDER of existing shapes to msoBringInFrontOfText = 4 (or: msoBringToFront = 0).
                //In these templates, a title line is followed by numerous newlines that shall go BEHIND the boxes below the title,
                //and define its distance from the first row of the bill positions table. Shape ZORDER is interpreted as inline with the text,
                //when these templates are imported from their OpenOffice format (for whatever reason).
                //As we've already done work to recognize these templates, we can just as well change that problem on the fly.

// DIS FOR TEST 				if ( ProbablyUsingTarmed_xxTemplateSoDoNOTAccessHeaderRangeToAvoidGenerationOfEmptyHeaderLines ) {
 				if ( true ) {
                    System.out.println("MSWord_jsText: findOrReplace (Shapes): INFO: Workaround for Tarmedrechnung_xx templates.");
                	//Dispatch.call(jacobShape, "ZOrder", 0);
                    //Dispatch.call(jacobShapeTextFrame, "ZOrder", 0);
                    
                    System.out.println("MSWord_jsText: findOrReplace (Shapes): INFO: Workaround for Tarmedrechnung_xx templates. About to Dispatch jacobShapeWrapFormat = Dispatch.get(jacobShape, \"WrapFormat\").toDispatch();");
                    Dispatch jacobShapeWrapFormat = Dispatch.get(jacobShape, "WrapFormat").toDispatch();
                    System.out.println("MSWord_jsText: findOrReplace (Shapes): INFO: Workaround for Tarmedrechnung_xx templates. About to Dispatch.put(jacobShapeWrapFormat, \"AllowOverlap\", new Variant (true));");
                    Dispatch.put(jacobShapeWrapFormat, "AllowOverlap", new Variant (true));
                    
                    System.out.println("MSWord_jsText: findOrReplace (Shapes): INFO: Workaround for Tarmedrechnung_xx templates. About to Dispatch.call(jacobShape, \"ZOrder\", 4);");
                    Dispatch.call(jacobShape, "ZOrder", 4);
                    
                }
                
                
            	
                
                
                
                
                
                System.out.println("MSWord_jsText: findOrReplace (Shapes): About to Dispatch jacobShapeTextFrame = Dispatch.call(jacobShape, \"TextFrame\").toDispatch();");
                Dispatch jacobShapeTextFrame = Dispatch.call(jacobShape, "TextFrame").toDispatch();

                System.out.println("MSWord_jsText: findOrReplace (Shapes): About to Integer jacobShapeTextFrameHasText = Dispatch.get(jacobShapeTextFrame, \"HasText\").toInt();");
                Integer jacobShapeTextFrameHasText = Dispatch.get(jacobShapeTextFrame, "HasText").toInt();
                System.out.println("MSWord_jsText: findOrReplace (Shapes): Shape["+(i+1)+"].TextFrame.HasText="+jacobShapeTextFrameHasText);
                
                if (jacobShapeTextFrameHasText == -1) {
                	do {
        				jacobSearchResultInt = 0;	//Reset this to 0 in each loop iteration, so that even when jacobFind.invoke() should NOT put a valid result into this variable
        											//(or when it should not even get invoked!),
													//we do definitely NOT get an endless loop, NOR a misguided attempt to replace text. 	

                    	//Hier muss das ganze Procedere, schon beginnend mit Dispatch jacobShapeTextFrameTextRange = ... (!!!) in den do...while block,
	            		//damit mehrere Platzhalter innerhalb eines Shapes ersetzt werden. Hab's probiert, nichts anderes hilft.
	            		//Insbesondere auch nicht das Verschieben des Cursors nach Durchführen einer Ersetzung - und das,
	            		//nachdem ich sehr lange gebraucht habe, um herauszufinden, wie das wirklich ausführbar codiert werden kann. 201609271147js
	                	Dispatch jacobShapeTextFrameTextRange = Dispatch.call(jacobShapeTextFrame, "TextRange").toDispatch();
	                	String jacobShapeTextFrameTextRangeText = Dispatch.get(jacobShapeTextFrameTextRange, "Text").toString();
	                    
	                    if (jacobShapeTextFrameTextRangeText == null)	System.out.println("MSWord_jsText: findOrReplace (Shapes): WARNING: Shape["+(i+1)+"].TextFrame.TextRange.Text IS NULL");
	                    else {
	                    	System.out.println("MSWord_jsText: findOrReplace (Shapes): Shape["+(i+1)+"].TextFrame.TextRange.Text="+jacobShapeTextFrameTextRangeText);
	                		//JETZT HABEN WIR ENDLICH DEN TEXT DES Shapes.Shape.TextFrame.TextRange.Text...
	
	                    	
	                    	//THIS WORKS, and causes the *shape!* to become selected, but *not* the text inside.
	                    	/*                    	
	                    	System.out.println("MSWord_jsText: findOrReplace: About to Variant jacobShapeSelectVariant = Dispatch.call(jacobShape, \"Select\");");                        
	                    	Variant jacobShapeSelectVariant = Dispatch.call(jacobShape, "Select");
	                        if (jacobShapeSelectVariant == null)	System.out.println("MSWord_jsText: findOrReplace: WARNING: jacobShapeSelectVariant IS NULL");
	                        else 									System.out.println("MSWord_jsText: findOrReplace: jacobShapeSelectVariant="+jacobShapeSelectVariant.toString());
	                        */
	
	
	                    	//THIS WORKS, and causes the text in the shape to become selected            
/*	                    	System.out.println("MSWord_jsText: findOrReplace (Shapes): About to Variant jacobShapeTextFrameTextRangeSelectVariant = Dispatch.call(jacobShapeTextFrameTextRange, \"Select\");");
	                        Variant jacobShapeTextFrameTextRangeSelectVariant = Dispatch.call(jacobShapeTextFrameTextRange, "Select");
	                        if (jacobShapeTextFrameTextRangeSelectVariant == null)	System.out.println("MSWord_jsText: findOrReplace (Shapes): WARNING: jacobShapeTextFrameTextRangeSelectVariant IS NULL");
	                        else 	System.out.println("MSWord_jsText: findOrReplace (Shapes): jacobShapeTextFrameTextRangeSelectVariant="+jacobShapeTextFrameTextRangeSelectVariant.toString());                        
	                        
	
	
	                        /*
	                    	//THIS WORKS, and causes the text in the shape to remain selected
	                        //But later on, when I want to  use jacobShapeFind, that throws this error
	                        //com.jacob.com.ComFailException: A COM exception has been encountered:
	                        //At Invoke of: Text
	                        //Description: 80020011 / Does not support a collection.
	                        */
/*	                    	System.out.println("MSWord_jsText: findOrReplace (Shapes): About to Variant jacobShapeTextFrameTextRangeFindVariant = Dispatch.call(jacobShapeTextFrameTextRange, \"Find\");");
	                        Variant jacobShapeTextFrameTextRangeFindVariant = Dispatch.get(jacobShapeTextFrameTextRange, "Find");
	                        if (jacobShapeTextFrameTextRangeFindVariant == null)	System.out.println("MSWord_jsText: findOrReplace (Shapes): WARNING: jacobShapeTextFrameTextRangeFindVariant IS NULL");
	                        else 								System.out.println("MSWord_jsText: findOrReplace (Shapes): jacobShapeTextFrameTextRangeFindVariant="+jacobShapeTextFrameTextRangeFindVariant.toString());                        
	
	                        System.out.println("MSWord_jsText: findOrReplace (Shapes): About to Dispatch jacobShapeTextFrameTextRangeFind = jacobShapeTextFrameTextRangeFindVariant.toDispatch();");
	                        Dispatch jacobShapeTextFrameTextRangeFind = jacobShapeTextFrameTextRangeFindVariant.toDispatch();
	                        if (jacobShapeTextFrameTextRangeFind  == null)	System.out.println("MSWord_jsText: findOrReplace (Shapes): WARNING: jacobShapeTextFrameTextRangeFind IS NULL");
	                        else 							System.out.println("MSWord_jsText: findOrReplace (Shapes): jacobShapeTextFrameTextRangeFind="+jacobShapeTextFrameTextRangeFind.toString());
	
	                        
	
	                        
	                        
	                        System.out.println("MSWord_jsText: findOrReplace (Shapes): About to try Dispatch.call(jacobShapeTextFrameTextRangeFind, \"Text\", pattern2);... etc.");
	                		try {
	                			Dispatch.put(jacobShapeTextFrameTextRangeFind, "Text", pattern2);
	                			Dispatch.put(jacobShapeTextFrameTextRangeFind, "Forward", "True");
	                			Dispatch.put(jacobShapeTextFrameTextRangeFind, "Format", "False");
	                			Dispatch.put(jacobShapeTextFrameTextRangeFind, "MatchCase", "False");
	                			Dispatch.put(jacobShapeTextFrameTextRangeFind, "MatchWholeWord", "False");
	                			Dispatch.put(jacobShapeTextFrameTextRangeFind, "MatchByte", "False");
	                			Dispatch.put(jacobShapeTextFrameTextRangeFind, "MatchAllWordForms", "False");
	                			Dispatch.put(jacobShapeTextFrameTextRangeFind, "MatchSoundsLike", "False");
	                			Dispatch.put(jacobShapeTextFrameTextRangeFind, "MatchWildcards", "True");	
	                		} catch (Exception ex) {
	                			ExHandler.handle(ex);
	                			System.out.println("MSWord_jsText: findOrReplace (Shapes): Fehler beim Ersetzen: Dispatch.put(jacobShapeTextFrameTextRangeFind...);");
	                			
	                			//SWTHelper.showError("No doc in bill", "Fehler:","Es ist keine Rechnungsvorlage definiert");
	                			SWTHelper.showError("findOrReplace (Shapes): Fehler beim Ersetzen: ", "Fehler:","Fehler beim Ersetzen: Dispatch.put(jacobShapeTextFrameTextRangeFind...);");
	                		}
	 
	                    	
	                		//The following block performs search-and-replace for each Shapes text portion of the document.
	                		//An almost identical block is further above for the main text (without updating all the comments),
	                		//and will probably be added further below, for tables. 
	                		
	                        System.out.println("MSWord_jsText: findOrReplace (Shapes): About to try ... the actual search and replace block...");
	                		try {	 	            				
//DIS FOR TEST	            				if (pattern == null)	System.out.println("MSWord_jsText: findOrReplace (Shapes): ERROR: pattern IS NULL!");
//DIS FOR TEST	            				else 					System.out.println("MSWord_jsText: findOrReplace (Shapes): pattern="+pattern);
	            				
	            				if (pattern2 == null)	System.out.println("MSWord_jsText: findOrReplace (Shapes): ERROR: pattern2 IS NULL!");
	            				else 					System.out.println("MSWord_jsText: findOrReplace (Shapes): pattern2="+pattern2);
	            				
	            				try {
	            					System.out.println("MSWord_jsText: findOrReplace (Shapes): About to jacobSearchResultInt = Dispatch.call(jacobShapeTextFrameTextRangeFind,\"Execute\").toInt();...");
	            					//jacobSearchResultInt = jacobFind.invoke("Execute").toInt();
	            					jacobSearchResultInt = Dispatch.call(jacobShapeTextFrameTextRangeFind,"Execute").toInt();
	            					
	            					//Please note: Wenn ich erste jacobSearchResultInt = jacobSearchresultVariant.toInt() verwende, ist nacher auch der string = ""-1", sonst "true"" +
	            					System.out.println("MSWord_jsText: findOrReplace: jacobSearchResultInt="+jacobSearchResultInt);
	            					//System.out.println("Result: jacobInvokeResult.toString()="+jacobInvokeResult.toString());	//Returns true if match found, false if no match found
	            					//System.out.println("Result: jacobInvokeResult.toInt()="+jacobInvokeResult.toInt());		//Returns -1 if match found, 0 if no match found
	            					//System.out.println("Result: jacobInvokeResult.toError()="+jacobInvokeResult.toError());	//Throws java.lang.IllegalStateException: getError() only legal on Variants of type VariantError, not 3
	            				} catch (Exception ex) {
	            					ExHandler.handle(ex);
	            					//ToDo: Add precautions for pattern==null or pattern2==null...
	            					System.out.println("MSWord_jsText: findOrReplace (Shapes):\nException caught.\n"+
	            					"Das verwendete Suchpattern funktioniert möglicherweise NICHT mit MS Word.\n"+
//DIS FOR TEST	            					"pattern \t(Von Elexis für OpenOffice):\t"+pattern+"\n"+
	            					"pattern2\t(Von msword_js für MS Word):\t"+pattern2+"\n"+		
	            					"Falls neue Suchpatterns hinzugefügt wurden, muss möglicherweise eine on-the-fly-Konversion in com.jsigle.msword_js.MSWord_jsText.java ergänzt werden.");
	            					//ToDo: Add precautions for pattern==null or pattern2==null...
	            					SWTHelper.showError(
	            							"MSWord_jsText: findOrReplace (Shapes):", 
	            							"Fehler:",
	            							"Exception caught.\n"+
	            							"Das verwendete Suchpattern funktioniert möglicherweise NICHT mit MS Word.\n"+
//DIS FOR TEST	            							"pattern \t(Von Elexis für OpenOffice):    \t"+pattern+"\n"+	//spaces needed for tab alignment in proportional font
	            							"pattern2\t(Von msword_js für MS Word):\t"+pattern2+"\n"+		
	            							"Falls neue Suchpatterns hinzugefügt wurden, muss möglicherweise eine on-the-fly-Konversion in com.jsigle.msword_js.MSWord_jsText.java ergänzt werden.");
	            				}
	            				
	            				if (jacobSearchResultInt == -1) {		//Do ONLY attempt a replacement if there was a search hit. Otherwise, replacement text might be simply inserted at cursor position.
            						//the following line should NOT produce an error - but it's flagged in Eclipse's editor with a red cross (x):
            						//The local variable ... might not have been initialized - well, it's *defined* before and outside both try... blocks?!
            						//If I actually initialize the variable up there to = null, the code error notice disappears.
            						
            						numberOfHits += 1;
            						System.out.println("MSWord_jsText: findOrReplace (Shapes): numberOfHits="+numberOfHits);
            						
            						//System.out.println("MSWord_jsText: findOrReplace: About to jacobSelection.setProperty(\"Text\", \"Replaced\");");						
            						//jacobSelection.setProperty("Text", "Replaced");		//Das sollte "Replaced" anstelle des Suchtexts einfügen.
            						
            						            						
            			/* SIMPLE GET, PRINTLN, REPLACE WITH CONSTANT STRING FOR TESTING ONLY - THIS WORKS :-) :-) :-)
            						//and very fine: especially, in Bern, [Brief.Datum] (or similar), only the [Brief.Datum] portion is replaced.
            						//This means, that the "Find" stuff actually works and controls the range that is influenced by the following commands. PUH.
            			*/
  /*          						System.out.println(Dispatch.get(jacobShapeTextFrameTextRange, "Text").toString());
            						Dispatch.put(jacobShapeTextFrameTextRange, "Text", new Variant("Replaced"));
            						                						

/* A LOT OF CODE REMOVED FOR THE TEST SETTING */            						
            						
        							//GETESTET: DAS MoveRight; MoveLeft; IST WIRKLICH NÖTIG, UM IM HAUPTTEXT ZUVERLÄSSIG ALLE PLATZHALTER ZU ERSETZEN. NICHT NÖTIG IN SHAPES.

                                    //Kommentare zur Info von oben übernommen:
                                    //
        	                        //DAS HIER SCHEINT ZU GEHEN!!!!! ENDLICH!!!
        	                        //(Analog der Methode, die über insertText()... cur ... pos hinwegegeholfen hat,
        	                        // wobei ich mich einfach nicht darum kümmere, eine Selection als Eigenschaft des aktuellen Textfeldes anzusprechen -
        	                        // sondern einfach eine Selection als Eigenschaft des ganzen Dokuments!)
        	                        //
        	                        //Object cur = jacobSelection.getObject();
        	                        //Dispatch.call((Dispatch) cur, "MoveRight");
        	                        //
        	                    	// UND DAS HIER AUCH - ist effektiv eine Kurzform davon:
                                    //Dispatch.call(jacobSelection, "MoveLeft");
            						//
                                    //THIS APPARENTLY WORKS!!!
                                    //System.out.println("MSWord_jsText: findOrReplace (Shapes): About to Dispatch.call(jacobSelection, \"MoveRight\");");
        	                        //Dispatch.call(jacobSelection, "MoveRight");
                                    //System.out.println("MSWord_jsText: findOrReplace (Shapes): About to Dispatch.call(jacobSelection, \"MoveLeft\");");
        	                        //Dispatch.call(jacobSelection, "MoveLeft");
        	                        //Nun. Das verschiebt zwar den Cursor ans Ende des Textes im Shape - aber immer noch wird nur EIN Platzhalter ersetzt...
                                    //for (int j = 0; j < 100; j++) { Dispatch.call(jacobSelection, "MoveLeft"); }
                                    //Das verschiebt den Cursor weit nach oben zum Anfang des Textfeldes - aber trotzdem wird nur EIN Platzhalter ersetzt...
            						
/*            						System.out.println("");
	            				}
	                		} catch (Exception ex) {
	                			ExHandler.handle(ex);
	                			//ToDo: Add precautions for pattern==null or pattern2==null...
	                			System.out.println("MSWord_jsText: findOrReplace (Shapes):\nFehler beim Suchen und Ersetzen im Haupttext:\n"+"" +
	                					"Exception caught für:\n"+
//DIS FOR TEST	                					"pattern \t(Von Elexis für OpenOffice):\t"+pattern+"\n"+
	                					"pattern2\t(Von msword_js für MS Word):\t"+pattern2+"\n"+		
	                					"numberOfHits="+numberOfHits);
	                			//SWTHelper.showError("No doc in bill", "Fehler:","Es ist keine Rechnungsvorlage definiert");
	                			//ToDo: Add precautions for pattern==null or pattern2==null...
	                			SWTHelper.showError(
	                					"MSWord_jsText: findOrReplace:"+ 
	                					"Fehler:",
	                					"MSWord_jsText: findOrReplace (Shapes):\nFehler beim Suchen und Ersetzen im Haupttext:\n"+
	                					"Exception caught für:\n"+
//DIS FOR TEST	                					"pattern \t(Von Elexis für OpenOffice):    \t"+pattern+"\n"+	//spaces needed for tab alignment in proportional font
	                					"pattern2\t(Von msword_js für MS Word):\t"+pattern2+"\n"+		
	                					"numberOfHits="+numberOfHits);
	                		}
	 
		                }
                	} while (jacobSearchResultInt == -1);
                } // (jacobShapeTextFrameHasText == -1)
            } //for (int i = 0; i < shapesCount; i++)            
*/  
        
        }
        catch (Exception ex) {
			ExHandler.handle(ex);
			//ToDo: Add precautions for pattern==null or pattern2==null...
			System.out.println("MSWord_jsText: findOrReplace:\nFehler beim Suchen und Ersetzen in Shapes:\n"+"" +
					"Exception caught für:\n"+
//DIS FOR TEST					"pattern \t(Von Elexis für OpenOffice):\t"+pattern+"\n"+
					"pattern2\t(Von msword_js für MS Word):\t"+pattern2+"\n"+		
					"numberOfHits="+numberOfHits);
			//SWTHelper.showError("No doc in bill", "Fehler:","Es ist keine Rechnungsvorlage definiert");
			//ToDo: Add precautions for pattern==null or pattern2==null...
			SWTHelper.showError(
					"MSWord_jsText: findOrReplace:"+ 
					"Fehler:",
					"MSWord_jsText: findOrReplace:\nFehler beim Suchen und Ersetzen in Shapes:\n"+
					"Exception caught für:\n"+
//DIS FOR TEST					"pattern \t(Von Elexis für OpenOffice):    \t"+pattern+"\n"+	//spaces needed for tab alignment in proportional font
					"pattern2\t(Von msword_js für MS Word):\t"+pattern2+"\n"+		
					"numberOfHits="+numberOfHits);
        }
        		
		
		

/* Ende TESTCODE FÜR SUCHEN IN SHAPES, esp. wie man dort MoveLeft oder MoveRight implementieren kann
* und mehrere [Platzhalter] innerhalb desselben Shapes verarbeiten kann (gelöst). */
			
			
		

			
			
			
			
/* TEST CODE FOR INSERTING TEXTFRAME SHAPES... 	   	THIS IS HIGHLY OUTDATED AND INCOMPLETE NOW,
 * 													SEE THE ACTUAL IMPLEMENTATION BELOW AROUND AddTextShape
 * 													IT HAS SUPPORT FOR CORRECT POSITIONING, INVISIBLE FRAMES, AUTOSIZEADJUST, NOWRAP, ADJUST, SETFORMAT ETC.PP.
			
		    //Try findOrReplace in a multi-placeholder-Shape
		    
			// Duplicate a bunch of fields and code from MSWord_jsText inside main() for quick testing only,
			// so that a LOCAL COPY!!! of portions of code from findOrReplace can be used for quick r&d unaltered below here.
			// That may be removed some time later, and yes, this surely adds to intransparency and confusion...
			final ActiveXComponent jacobObjWord;
			final Dispatch jacobCustDocprops;
			final Dispatch jacobBuiltInDocProps;
			final Dispatch jacobDocuments;
			final Dispatch jacobDocument;
			final Dispatch jacobWordObject;
			jacobObjWord = new ActiveXComponent("Word.Application");
		    jacobObjWord.setProperty("Visible", true);
		    jacobDocuments = jacobObjWord.getPropertyAsComponent("Documents");
		    String myInputDoc = "L:/home/jsigle/workspace/elexis-2.1.7-20130523/elexis-bootstrap-js/jsigle/com.jsigle.msword_js/doc/Vorlage-scratch2.doc";
			jacobDocument = ((ActiveXComponent) jacobDocuments).invokeGetComponent("Open", new Variant(myInputDoc)); 
			
			System.out.println("MSWord_jsText: About to ActiveXComponent jacobSelection = jacobObjWord.getPropertyAsComponent(\"Selection\");");
			ActiveXComponent jacobSelection = jacobObjWord.getPropertyAsComponent("Selection");
			System.out.println("MSWord_jsText: About to Dispatch jacobShapes = Dispatch.get((Dispatch) jacobDocument, \"Shapes\").toDispatch();");
			Dispatch jacobShapes = Dispatch.get((Dispatch) jacobDocument, "Shapes").toDispatch();
            
			int x=0;
			int y=200;
			int w=100;
			int h=50;
			
			String text = "Dies ist ein Teststring";
			
			try {
				
				//THIS FAILS:
				//com.jacob.com.ComFailException: Invoke of: AddShape
				//Source: 
				//Description: Der angegebene Wert ist außerhalb des zulässigen Bereichs.
				//
				//Variant jacobShapeVariant = Dispatch.call(jacobShapes, "AddShape", 0, x, y, w, h);
				//
				//Das Problem ist die 0 bei Type. Ab 1 geht es.

				//Typen:
				//0:
				//com.jacob.com.ComFailException: Invoke of: AddShape
				//Source: 
				//Description: Der angegebene Wert ist außerhalb des zulässigen Bereichs.
				//1 = Rahmen, kein Text drin
				//2 = Parallelogramm
				//3 = Trapez
				//4 = Raute / Rhombus
				//5 = Kasten, runde Ecken
				//6 = Kasten, abgeschrägte Ecken
				//...
				//System.out.println("MSWord_jsText: About to Variant jacobShapeVariant = Dispatch.call(jacobShapes, \"AddShape\", 1, "+x+", "+y+", "+w+", "+h+");");
			    //Variant jacobShapeVariant = Dispatch.call(jacobShapes, "AddShape", 6, x, y, w, h);

				//Typen:
				//0:
				//com.jacob.com.ComFailException: Invoke of: AddTextBox
				//Source: 
				//Description: Der angegebene Wert ist außerhalb des zulässigen Bereichs.
				//1 = Textbox, Ränder schwarz, horizontal
				//2 = Textbox, Ränder schwarz, text nach oben (um 90° gegen Uhrzeigersinn gedreht)
				//3 = Textbox, Ränder schwarz, text nach unten (um 90° im Uhrzeigersinn gedreht)
				//4 = Textbox, Ränder schwarz, text nach unten (um 90° im Uhrzeigersinn gedreht)
				//5 = Textbox, Ränder schwarz, text nach unten (um 90° im Uhrzeigersinn gedreht)
				//...
				System.out.println("MSWord_jsText: About to Variant jacobShapeVariant = Dispatch.call(jacobShapes, \"AddTextBox\", 1, "+x+", "+y+", "+w+", "+h+");");
			    Variant jacobShapeVariant = Dispatch.call(jacobShapes, "AddTextBox", 1, x, y, w, h);
			    
			    System.out.println("MSWord_jsText: About to Dispatch jacobShape = jacobShapeVariant.toDispatch();");
	            Dispatch jacobShape = jacobShapeVariant.toDispatch();

				System.out.println("MSWord_jsText: insertTextAt: About to Dispatch jacobShapeTextFrame = Dispatch.call(jacobShape, \"TextFrame\").toDispatch();");
	            Dispatch jacobShapeTextFrame = Dispatch.call(jacobShape, "TextFrame").toDispatch();
	            
	            System.out.println("MSWord_jsText: insertTextAt: About to Integer jacobShapeTextFrameHasText = Dispatch.get(jacobShapeTextFrame, \"HasText\").toInt();");
	            Integer jacobShapeTextFrameHasText = Dispatch.get(jacobShapeTextFrame, "HasText").toInt();
	            
	            System.out.println("MSWord_jsText: insertTextAt: Shape[added].TextFrame.HasText="+jacobShapeTextFrameHasText);
	            
	            if (jacobShapeTextFrameHasText == -1) {
	            	Dispatch jacobShapeTextFrameTextRange = Dispatch.call(jacobShapeTextFrame, "TextRange").toDispatch();
	            	String jacobShapeTextFrameTextRangeText = Dispatch.get(jacobShapeTextFrameTextRange, "Text").toString();
	                
	                if (jacobShapeTextFrameTextRangeText == null)	System.out.println("MSWord_jsText: insertTextAt: WARNING: Shape[added].TextFrame.TextRange.Text IS NULL");
	                else {
	                	System.out.println("MSWord_jsText: insertTextAt: Shape[added].TextFrame.TextRange.Text="+jacobShapeTextFrameTextRangeText);
	
	                	//THIS WORKS, and causes the text in the shape to become selected            
	                	System.out.println("MSWord_jsText: insertTextAt: About to Variant jacobShapeTextFrameTextRangeSelectVariant = Dispatch.call(jacobShapeTextFrameTextRange, \"Select\");");
	                    Variant jacobShapeTextFrameTextRangeSelectVariant = Dispatch.call(jacobShapeTextFrameTextRange, "Select");
	                    if (jacobShapeTextFrameTextRangeSelectVariant == null)	System.out.println("MSWord_jsText: insertTextAt: WARNING: jacobShapeTextFrameTextRangeSelectVariant IS NULL");
	                    else 	System.out.println("MSWord_jsText: insertTextAt: jacobShapeTextFrameTextRangeSelectVariant="+jacobShapeTextFrameTextRangeSelectVariant.toString());                        
	                    }
	
			    	System.out.println("MSWord_jsText: TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO ");
			    	System.out.println("MSWord_jsText: insertTextAt(): ToDo: Support for Font control.");
			    	System.out.println("MSWord_jsText: TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO ");
		                
		                
					ORIGINAL CODE FROM NOATEXT/OPENOFFICE`*/
					/* ORIGINAL CODE FROM NOATEXT/OPENOFFICE
					com.sun.star.beans.XPropertySet charProps = setFormat(xtc);
					Das folgende ersetzt die entsprechende Prozedur - 
					I'm doing this inline (in three occasions in this file) to avoid all the complications
					Java usually brings up when you just would want to move a little bit of code into a simple procedure.  
					ORIGINAL CODE FROM NOATEXT/OPENOFFICE`*/
					
/* CONTINUED...
					System.out.println("MSWord_jsText: insertTextAt: About to Dispatch fontDispatch = Dispatch.get(jacobSelection, \"Font\").toDispatch();");
		            Dispatch fontDispatch = Dispatch.get(jacobSelection, "Font").toDispatch();
			    	if ( font != null )		{ 
				    	System.out.println("MSWord_jsText: insertTextAt: About to Dispatch.put(fontDispatch, \"Name\", new Variant(font));");
			    		Dispatch.put(fontDispatch, "Name", new Variant(font));
			    	} else {
			    		System.out.println("MSWord_jsText: insertTextAt: WARNING: font IS NULL.");
			    	}
			        if (hi > 0)				{ 
			        	System.out.println("MSWord_jsText: insertTextAt: Dispatch.put(fontDispatch, \"Size\", new Float(hi));");
			    		//Dispatch.put(fontDispatch, "CharHeight", new Float(hi)); 	//OpenOffice: Height of the character in point
			    		Dispatch.put(fontDispatch, "Size", new Float(hi)); 
			        }
			        if (stil > -1) {
		        		System.out.println("MSWord_jsText: WARNING: The MS Word FONT property does apparently NOT support numeric font weight, so we have fewer steps available. { SWT.MIN = SWT.NORMAL; SWT.BOLD }"); 
			        	switch (stil) {
			        	case SWT.MIN:		{ 
			        		System.out.println("MSWord_jsText: insertTextAt: Mapping: OpenOffice SWT.MIN -> Dispatch.put(fontDispatch, \"Bold\", false);");
			        		//Dispatch.put(fontDispatch, "CharWeight", 15f); break; 
				        	Dispatch.put(fontDispatch, "Bold", false);
			        	}
			        	case SWT.NORMAL:	{ 
				        	System.out.println("MSWord_jsText: insertTextAt: Mapping: OpenOffice SWT.NORMAL -> Dispatch.put(fontDispatch, \"Bold\", false);");
			        		//Dispatch.put(fontDispatch, "CharWeight", FontWeight.LIGHT); break;
				        	Dispatch.put(fontDispatch, "Bold", false);
			        	}
			        	case SWT.BOLD:		{ 
				        	System.out.println("MSWord_jsText: insertTextAt: Mapping: OpenOffice SWT.BOLD -> Dispatch.put(fontDispatch, \"Bold\", true);");
			        		//Dispatch.put(fontDispatch, "CharWeight", FontWeight.BOLD); break;
				        	Dispatch.put(fontDispatch, "Bold", true);
			        	}
				        }
			        }
					        	
			        /*
			         Oben sind die Parameter von OpenOffice wie in NoaText_jsl verwendet.
			         Ansonsten gäbe es laut Infoseiten für MS Word VBA wohl diese Parameter:
			         
			         	Dispatch.put(fontDispatch, "Size", new Float(hi)); 
			        	Dispatch.put(fontDispatch, "Bold", new Variant(bold)
			            Dispatch.put(fontDispatch, "Italic", new Variant(italic));
			            Dispatch.put(fontDispatch, "Underline", new Variant(underLine));
			            Dispatch.put(fontDispatch, "Color", colorSize);
			        */
/* CONTINUED...					
			    	System.out.println("MSWord_jsText: TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO ");
			    	System.out.println("MSWord_jsText: insertTextAt(): ToDo: Support for ParagraphAdjust SWT.LEFT SWT.RIGHT default");
			    	System.out.println("MSWord_jsText: TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO ");
					
			        /* ORIGINAL FROM NOATEXT/OPENOFFICE
					ParagraphAdjust paradj;
					switch (adjust) {
					case SWT.LEFT:
						paradj = ParagraphAdjust.LEFT;
						break;
					case SWT.RIGHT:
						paradj = ParagraphAdjust.RIGHT;
						break;
					default:
						paradj = ParagraphAdjust.CENTER;
					}
					
					charProps.setPropertyValue("ParaAdjust", paradj);
					xFrameText.insertString(xtc, text, false);
					
					ORIGINAL FROM NOATEXT/OPENOFFICE */
					
					/*
					String orig = Dispatch.get(jacobShapeTextFrameTextRange, "Text").toString();
					if (orig  == null)	System.out.println("MSWord_jsText: insertTextAt: ERROR: orig IS NULL!");
					else 				System.out.println("MSWord_jsText: insertTextAt: orig="+orig);
					*/
/* CONTINUED...					
			    	System.out.println("MSWord_jsText: insertTextAt: text == "+text.toString());
					System.out.println("MSWord_jsText: insertTextAt: About to Dispatch.put(jacobShapeTextFrameTextRange, \"Text\", text);");
					Dispatch.put(jacobShapeTextFrameTextRange, "Text", text);
	            }
			
			} catch (Exception e) {
				System.out.println(e);
			}
			
			
END OF TESTCODE FOR INSERTING TEXTFRAME SHAPES*/			
    		
			
            //Close the word window directly after the test; we only want to see the console log output.
            //jacobObjWord.invoke("Quit", new Variant(false));
			
			
			
			
			
			
	        
		} catch (Exception e) {
			System.out.println(e);
		}
	

        
        System.out.println("");
		System.out.println("");
		System.out.println("MSWord_jsText: main(String[] args): Simple main program DEMO ends.");
		System.out.println("");
		System.out.println("MSWord_jsText: main(): IF YOU SEE THIS in console log output, and only a little bit more,");
		System.out.println("you've probably run the plugin msword_js instead of Elexis ... from Run - Configurations or Run - History...");
		System.out.println("You can probably see already whether contacting MS Word basically works.");
		System.out.println("");
		System.out.println("But please re-check what you're doing! - e.g. try: Run - Run History - Elexis Praxis Windows with JaCoB DLLs...");
		System.out.println("201609221831js");
}
	
	

	
	
	
	
	/*
	 * The following methods are taken over from NOAText,
	 * they need to be implemented foran ITextPlugin or ICloseListener object.
	 */
	
	/*
	public class NOAText implements ITextPlugin {
		public static final String MIMETYPE_OO2 = "application/vnd.oasis.opendocument.text";
		public static LinkedList<NOAText> noas = new LinkedList<NOAText>();
		OfficePanel panel;
		ITextDocument doc;
		ICallback textHandler;
		File myFile;
		private final Log log = Log.get("NOAText");
		IOfficeApplication office;
		private String font;
		private float hi = 0;
		private int stil = -1;
		
		public NOAText(){
			System.out.println("NOAText msword_js: NOAText msword_js: noa loaded");
			System.out.println("NOAText msword_js: NOAText msword_js: computing defaultbase...");
<<<<<<< Updated upstream
			File base = new File(Hub.getBasePath());
			File fDef = new File(base.getParentFile().getParent() + "/ooo");
			System.out.println("NOAText msword_js: NOAText msword_js: Hub.getBasePath():"+Hub.getBasePath());
=======
			File base = new File(CoreHub.getBasePath());
			File fDef = new File(base.getParentFile().getParent() + "/ooo");
			System.out.println("NOAText msword_js: NOAText msword_js: CoreHub.getBasePath():"+CoreHub.getBasePath());
>>>>>>> Stashed changes
			System.out.println("NOAText msword_js: NOAText msword_js: base.getParentFile().getParent() + \"/ooo\":"+base.getParentFile().getParent() + "/ooo");
			String defaultbase;
			if (fDef.exists()) {
				defaultbase = fDef.getAbsolutePath();
<<<<<<< Updated upstream
				//Hub.localCfg.set(PreferenceConstants.P_OOBASEDIR, defaultbase);   //20160921js
				Hub.localCfg.set(PreferenceConstants.P_MSWORDBASEDIR, defaultbase);
			} else {
				//defaultbase = Hub.localCfg.get(PreferenceConstants.P_OOBASEDIR, ".");
				defaultbase = Hub.localCfg.get(PreferenceConstants.P_MSWORDBASEDIR, ".");
=======
				//CoreHub.localCfg.set(PreferenceConstants.P_OOBASEDIR, defaultbase);   //20160921js
				CoreHub.localCfg.set(PreferenceConstants.P_MSWORDBASEDIR, defaultbase);
			} else {
				//defaultbase = CoreHub.localCfg.get(PreferenceConstants.P_OOBASEDIR, ".");
				defaultbase = CoreHub.localCfg.get(PreferenceConstants.P_MSWORDBASEDIR, ".");
>>>>>>> Stashed changes
			}
			System.out.println("NOAText msword_js: NOAText msword_js: computed defaultbase=openoffice.path.name:"+defaultbase);
			System.setProperty("openoffice.path.name", defaultbase);
		}
	*/

	/*
	 * We keep track on opened office windows
	 */
	private void createMe(){
		System.out.println("MSWord_jsText: createMe begin");
	
		//System.out.println("MSWord_jsText: createMe THIS IS CURRENTLY A DUMMY.");
		
		//System.out.println("MSWord_jsText: TODO: create an MS Word document and keep track of it.");
		//System.out.println("MSWord_jsText: TODO: However, I've deleted references to the ag.ion stuff doing that for OpenOffice.");
		//System.out.println("MSWord_jsText: TODO: So we need to re-define office, panel , and doc objects suitable for the MS Word Jacob Wrapper.");

		
		//This is what happened in NOAText_jsl - I could comment this out and the msword_js plugin basic functionality still worked;
		//i.e. the office object is not needed to get an empty panel and to retrieve a document from the database into a file, and open it in MS Word via JACOB. 
		/*
		if (office == null) {
			System.out.println("MSWord_jsText: Please note: createMe: office==null");
			office = EditorCorePlugin.getDefault().getManagedLocalOfficeApplication();
		}

		if (office == null)	System.out.println("MSWord_jsText: createMe: WARNING: still, office==null");
		else 				System.out.println("MSWord_jsText: createMe: office="+office.toString());		
		*/

		//This is the replacement for msword_js:
		if (jacobObjWord == null) {
			System.out.println("MSWord_jsText: createMe: INFO: jacobObjWord IS NULL.");
			System.out.println("MSWord_jsText: createMe: About to jacobObjWord = new ActiveXComponent(\"Word.Application\");...");
			jacobObjWord = new ActiveXComponent("Word.Application");
		}
		
		if (jacobObjWord==null)	System.out.println("MSWord_jsText: open(): WARNING: jacobObjWord==null ");
		else					System.out.println("MSWord_jsText: open(): jacobObjWord="+jacobObjWord.toString());

		
		
		
		
		
		
		//ToDo: Provide a replacement for getting the word document into a panel (if strongly desired,
		//ToDo:   with support for multiple panels in same or mult perspectives and same or mult instances of Elexis...) for msword_js...
		
		System.out.println("");
		System.out.println("MSWord_jsText: createMe: TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO");
		System.out.println("MSWord_jsText: createMe: ToDo: Provide a replacement for getting the word document into a panel (if strongly desired,");
		System.out.println("MSWord_jsText: createMe: ToDo:   with support for multiple panels in same or mult perspectives and same or mult instances of Elexis...) for msword_js.");
		System.out.println("MSWord_jsText: createMe: TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO");
		
		System.out.println("");
		System.out.println("MSWord_jsText: createMe: TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO");
		System.out.println("MSWord_jsText: createMe: ToDo: agIon related code is left here preliminarily, so that all original operations may continue to work"); 
		System.out.println("MSWord_jsText: createMe: ToDo:       until all methods have been transformed to msword_js JaCoB based implementation.");
		System.out.println("MSWord_jsText: createMe: ToDo:       Thereafter, all agIon related code should be removed. We will probably NOT continue to use");
		System.out.println("MSWord_jsText: createMe: ToDo:       a panel inside Elexis for MS Word documents, as it's:");
		System.out.println("MSWord_jsText: createMe: ToDo:       (a) more difficult to implement and maintain, and");
		System.out.println("MSWord_jsText: createMe: ToDo:       (b) rather limiting than useful during operation, and");
		System.out.println("MSWord_jsText: createMe: ToDo:       (c) only useful in the context of perspectives, i.e. beyond normal users' demand and understanding.");
		System.out.println("MSWord_jsText: createMe: TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO");
		System.out.println("MSWord_jsText: createMe: ToDo:       PLEASE NOTE: DO STILL KEEP A Briefe etc. Panel around, to keep its PulldownMenu available!!!!");
		System.out.println("MSWord_jsText: createMe: TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO");
		System.out.println("");
		
		
		if (agIonPanel == null)	System.out.println("MSWord_jsText: createMe: WARNING: panel==null");
		else 					System.out.println("MSWord_jsText: createMe: panel="+agIonPanel.toString());

		agIonDoc = (ITextDocument) agIonPanel.getDocument();
		
		if (agIonDoc == null)	System.out.println("MSWord_jsText: createMe: WARNING: doc==null, so we won't be able to doc.addCloseListener() or noas.add(this).");
		else 					System.out.println("MSWord_jsText: createMe: doc="+agIonDoc.toString());

		//ToDo: Provide a replacement for closeListener and noas-keeping-track-of-opened-documents for msword_js...

		System.out.println("MSWord_jsText: createMe: TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO");
		System.out.println("MSWord_jsText: createMe: ToDo: Provide a replacement for closeListener and noas-keeping-track-of-opened-documents for msword_js...");
		System.out.println("MSWord_jsText: createMe: TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO");
		
		if (agIonDoc != null) {
			System.out.println("MSWord_jsText: createMe: doc.addCloseListener()...");
		
			agIonDoc.addCloseListener(new closeListener(agIonOffice));
			
			System.out.println("MSWord_jsText: createMe: noas.add(this)...");
					
			agIonNoas.add(this);

			if (agIonNoas != null)	System.out.println("MSWord_jsText: createMe: noas = "+agIonNoas.toString());
			else					System.out.println("MSWord_jsText: createMe: WARNING: noas IS NULL, even though we should have added something.");
		}

		
		System.out.println("MSWord_jsText: createMe ends");
	}
	
	/**
	 * We deactivate the office application IF the user has closed the last office window:
	 * textHandler.save(); noas.remove(this); doc.setModified(false); doc.close(); if (noas.isEmpty()) { office.deactivate() }...
	 */
	private void removeMe(){
		System.out.println("MSWord_jsText: removeMe begin");
		
		System.out.println("MSWord_jsText: removeMe: TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO");
		System.out.println("MSWord_jsText: removeMe: ToDo: Provide a replacement for document closing and removal from monitoring list noas....");
		System.out.println("MSWord_jsText: removeMe: ToDo: Maybe JaCoB/MS Word is simpler to use: Just close/quit after each document,");
		System.out.println("MSWord_jsText: removeMe: ToDo: and it will close each window when asked, but truly quit Word only after the last window using it has been closed.");
		System.out.println("MSWord_jsText: removeMe: ToDo: TO REVIEW: !!!!!!!!!!!! So I'm essentially re-using the code from dispose() here. !!!!!!!!!!!!"); 
		System.out.println("MSWord_jsText: removeMe: ToDo: TO REVIEW: !!!!!!!!!!!! Maybe should CHECK IF NOAS ARE EMPTY - DON'T want to jacobObjWord=null; jacobDocuments=null; when inappropriate! !!!!!!!!!!!!"); 
		System.out.println("MSWord_jsText: removeMe: TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO");
		
		if (jacobDocument != null) {
			System.out.println("MSWord_jsText: removeMe: About to Dispatch.call(jacobDocument, \"Save\");");
		    Dispatch.call(jacobDocument, "Save");
		    //Dispatch.call(jacobDocument, "Close", new Variant(saveOnExit));

		    System.out.println("MSWord_jsText: removeMe: textHandler.save()...");
			textHandler.save();
   
			System.out.println("MSWord_jsText: removeMe: About to close()...");
		    close();	//this includes: jacobDocument = null;
		}
		else {System.out.println("MSWord_jsText: removeMe: WARNING: jacobDocument already WAS NULL.");}
		
		if (jacobObjWord != null) {
			System.out.println("MSWord_jsText: removeMe: About to quit()...");
			quit();		//this includes: jacobObjWord = null; jacobDocuments = null;
		}
		else {System.out.println("MSWord_jsText: removeMe: WARNING: jacobObjWord already WAS NULL.");}
		
		/* ORIGINAL CODE FROM THE NOATEXT/OPENOFFICE IMPLEMENTATION
		//System.out.println("MSWord_jsText: removeMe THIS IS CURRENTLY A DUMMY.");
		//System.out.println("MSWord_jsText: removeMe TODO: Should remove the office application after the user closes the last office window");
		
		try {
			System.out.println("MSWord_jsText: removeMe: trying 1...");
				if (textHandler != null) {
				System.out.println("MSWord_jsText: removeMe: textHandler.save()...");
				textHandler.save();
				System.out.println("MSWord_jsText: removeMe: noas.remove(this)...");;
				agIonNoas.remove(this);
				if (agIonDoc != null) {
					agIonDoc.setModified(false);
					System.out.println("MSWord_jsText: removeMe: doc.close()...");
					agIonDoc.close();
				}
			}
		} catch (Exception ex) {
			System.out.println("MSWord_jsText: removeMe: WARNING: caught Exception");
			ExHandler.handle(ex);
		}
		if (agIonNoas.isEmpty()) {
			System.out.println("MSWord_jsText: removeMe: noas.isEmpty()");
			try {
				System.out.println("MSWord_jsText: removeMe: trying office.deactivate()...");
				agIonOffice.deactivate();
				log.log("Office deactivated", Log.INFOS);
			} catch (OfficeApplicationException e) {
				System.out.println("MSWord_jsText: removeMe: WARNING: caught Exception");
				ExHandler.handle(e);
				log.log("Office deactivation failed", Log.ERRORS);
			}
		}
		ORIGINAL CODE FROM THE NOATEXT/OPENOFFICE IMPLEMENTATION */
	
		System.out.println("MSWord_jsText: removeMe ends");
	}
	
	/**
	 * This apparently merely saves the document and then clears the modified flag via: if (textHandler != null) { try { textHandler().save; doc.setModified(false); return true; } ... else return false;}
	 */
	public boolean clear(){
		System.out.println("MSWord_jsText: clear begins");
		
		//System.out.println("MSWord_jsText: clear THIS IS CURRENTLY A DUMMY.");
		//System.out.println("MSWord_jsText: clear TODO: Should save somthing open (to be understood...), then set.Modified(false)"); 
		

		//THIS WORKS NOW SINCE I NOTED DOWN textHandler = handler (supplied as ICallback handler in CreateMe) :-)
		//I did NOT need to write any implementation for the save() or saveAs() methods here for this.
		System.out.println("MSWord_jsText: clear: TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO");
		System.out.println("MSWord_jsText: clear: ToDo: Provide a replacement for document saving and (!) for the setting and clearing of agIonDoc.setModified() flag");
		System.out.println("MSWord_jsText: clear: TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO");

		if (textHandler != null) {
			try {
				System.out.println("MSWord_jsText: clear: textHandler is NOT null.");

				System.out.println("MSWord_jsText: clear: about to textHandler.save()...");
				textHandler.save();
			
				System.out.println("MSWord_jsText: clear: about to agIonDoc.setModified(false)... (PLEASE NOTE: agIon: is outdated)");
				if (agIonDoc != null)				//201701030912js to avoid an exception
					agIonDoc.setModified(false);
				
				System.out.println("MSWord_jsText: clear: about to end - returning true...");
				return true;
			} catch (DocumentException e) {
				ExHandler.handle(e);
			}
		} else {
			System.out.println("MSWord_jsText: clear: TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO");
			System.out.println("MSWord_jsText: clear: WARNING: ToDo: texthandler IS NULL when trying to use it to save document!");
			System.out.println("MSWord_jsText: clear: TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO");			
		}

		System.out.println("MSWord_jsText: clear: EITHER textHandler IS NULL -- OR DocumentException caught...");
	
		System.out.println("MSWord_jsText: clear ends, about to end - returning false");
		return false;
	}
	
	
	
	
	
	/**
	 * Create the OOo-Container that will appear inside the view or dialog for Text-Display.
	 * Here we use a slightly adapted OfficePanel from NOA4e (www.ubion.org)
	 */
	public Composite createContainer(final Composite parent, final ICallback handler){
		System.out.println("MSWord_jsText: createContainer begins");

		System.out.println("MSWord_jsText: createContainer: TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO");
		System.out.println("MSWord_jsText: createContainer: ToDo: Provide a replacement for createContainer()... / agIonPanel = new OfficePanel()");
		System.out.println("MSWord_jsText: createContainer: TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO");

		//System.out.println("MSWord_jsText: createContainer THIS IS CURRENTLY A DUMMY.");
		//System.out.println("MSWord_jsText: createContainer TODO: Should Probably open a new container = panel, i.e. frame (or window) (details to be understood) and return that");
		
		
		//Write down the ICallback handler for save operations passed from the calling TextContainer object.
		textHandler=handler;

		
		System.out.println("MSWord_jsText: AT THE MOMENT, A NEW WINDOW WILL APPEAR FOR EACH DOCUMENT.");
		
		System.out.println("MSWord_jsText: createContainer: About to new Frame()...");
		new Frame();
		
		System.out.println("MSWord_jsText: createContainer: About to TODO: THE FOLLOWING PROBABLY STILL USES OpenOffice, SHOULD USE MS Word instead!!!");

		System.out.println("MSWord_jsText: createContainer: About to panel = new OfficePanel()...");
		agIonPanel = new OfficePanel(parent, SWT.NONE);

		if (agIonPanel == null)	System.out.println("MSWord_jsText: createContainer: WARNING: panel IS NULL!");
		else 				System.out.println("MSWord_jsText: createContainer: panel="+agIonPanel.toString());
		
		agIonPanel.setBuildAlwaysNewFrames(false);
		
		System.out.println("MSWord_jsText: createContainer: About to office = EditorCorePlugin.getdefault().getmanagedLocalOfficeApplication()...");

		agIonOffice = EditorCorePlugin.getDefault().getManagedLocalOfficeApplication();

		if (agIonOffice == null)	System.out.println("MSWord_jsText: createContainer: WARNING: office IS NULL!");
		else 				System.out.println("MSWord_jsText: createContainer: office="+agIonOffice.toString());
		
		System.out.println("MSWord_jsText: createContainer ends, about to return panel");
		return agIonPanel;
	}
	
	
	
	
	
	/**
	 * Create an empty text document. We simply use an empty template and save it immediately into a
	 * temporary file to avoid OOo's complaints when we close the Container or overwrite its contents.
	 */
	public boolean createEmptyDocument(){
		System.out.println("MSWord_jsText: createEmptyDocument begins");

		System.out.println("MSWord_jsText: createEmptyDocument TODO: Review, understand");

		try {
			System.out.println("MSWord_jsText: createEmptyDocument: try...");
			System.out.println("MSWord_jsText: createEmptyDocument: About to clean()...");
			clean();
			
			System.out.println("MSWord_jsTextNTI: *** WARNING: THE plugin.xml Overview General Information ID must match this string,");
			System.out.println("MSWord_jsTextNTI: *** WARNING: otherwise the empty.docx / empty.doc / empty.odt is NOT found.");
			System.out.println("MSWord_jsTextNTI: *** WARNING: THIS MAY BE THE REASON FOR NEUES DOKUMENT NOT TO WORK IN NOATEXT JS exported SO FAR.");
						
			System.out.println("MSWord_jsText: createEmptyDocument: About to Bundle bundle = Platform.getBundle(\"com.jsigle.msword_js\")...");
			Bundle bundle = Platform.getBundle("com.jsigle.msword_js");
			
			System.out.println("MSWord_jsTextNTI: *** MS Word 2010 does NOT want to open the empty.odt document,");
			System.out.println("MSWord_jsTextNTI: *** most probably because it is an OpenDocument 1.2 document,");
			System.out.println("MSWord_jsTextNTI: *** and Word would only handle 1.1 - so we may need to call a proper");
			System.out.println("MSWord_jsTextNTI: *** Office program (like LibreOffice) to convert old *.odt documents");
			System.out.println("MSWord_jsTextNTI: *** to *.docx before opening them with word. Or we use *.doc to support");
			System.out.println("MSWord_jsTextNTI: *** older versions of word as well, if all else works with them.");
			
			System.out.println("MSWord_jsTextNTI: *** JUST NOW: added plugin.xml - Build - Binary Build - rsc/empty.docx and empty.odt");
			System.out.println("MSWord_jsTextNTI: *** JUST NOW: Maybe the missing latter was the reason for Neues Dokument not to work in NOAText js 2012-02-2x exported.");
			System.out.println("MSWord_jsTextNTI: *** JUST NOW: This didn't suffice to get past InputStream is ... rsc/empty.docx,");
			System.out.println("MSWord_jsTextNTI: *** JUST NOW: so I (also) added rsc/ in plugin.xml - Runtime - Classpath");
				
			System.out.println("MSWord_jsTextNTI: *** JUST NOW: Didn't work either in Eclipse. So AD HOC I USE: l:/Elexis/empty.docx");
			System.out.println("MSWord_jsTextNTI: *** JUST NOW: Didn't work either in Eclipse. So AD HOC I USE: empty.docx");
			System.out.println("MSWord_jsTextNTI: *** JUST NOW: Nope.");
			System.out.println("MSWord_jsTextNTI: *** JUST NOW: Putting it in the current user dir did not help either...");
			
			System.out.println("MSWord_jsTextNTI: !!!! WARNING: For MS Word 2003/2007 to open *.odt files, you need a Word plugin.");
			System.out.println("MSWord_jsTextNTI: !!!! WARNING: http://www.poweraxess.com/uncategorized/open-odt-files-ms-word");
			System.out.println("MSWord_jsTextNTI: !!!! WARNING: http://www.poweraxess.com/uncategorized/open-ms-word-2007-docx-in-ms-word-2003-or-older-versions");
			System.out.println("MSWord_jsTextNTI: !!!! WARNING: Mit dem Word 2007 SP 3 scheint es nun auch sonst problemlos zu gehen.");
					
			System.out.println("MSWord_jsTextNTI: !!!! INFO: Changed the code and ressource to use/provide rsc/empty.doc instead of rsc/empty.odt.");

			Path path = new Path("rsc/empty.doc");		//201611131641js Umstellung von *.odt auf *.doc für MS-Word
			////Path path = new Path("rsc/empty.odt");
			//Path path = new Path("rsc/empty.docx");
			
			System.out.println("Current user directory is: "+System.getProperty("user.dir"));
						
			System.out.println("InputStream is -- "+path);
			//InputStream is = FileLocator.openStream(bundle, path, false);
			InputStream is = FileLocator.openStream(bundle, path, true);
			
			System.out.println("FileOutputStream os -- "+myFile);
			FileOutputStream fos = new FileOutputStream(myFile);
		
			System.out.println("MSWord_jsTextNTI: *** WARNING: VERY FUNNY: When WORD OPENS, IT SHOWS *.ODT now...");
			System.out.println("MSWord_jsTextNTI: ***          (maybe due to ODT Plugin being installed? Or MIMETYPE elsewhere herein?)");
			System.out.println("MSWord_jsTextNTI: *** JUST NOW: Putting it in the current user dir did not help either...");

			System.out.println("copyStreams...");
			FileTool.copyStreams(is, fos);
			
			System.out.println("is.close()");
			is.close();
			
			System.out.println("os.close()");
			fos.close();
			
			//Beginning of MSWord_jsText adoption.

			System.out.println("MSWord_jsTextNTI: *** TO DO: We would now call panel.loadDocument()");
			System.out.println("   to load a new document into the panel created before.");
			System.out.println("   In the MS Word variant, currently, we do not load a document into a panel,");
			System.out.println("   We can just so instruct Word to open the document.");
			System.out.println("   The open() routine will read the file into the Dispatch document.");
			System.out.println("   Commented out the original code, replaced by open(myFile.getAbsolutePath(),true);");
			
			//panel.loadDocument(false, myFile.getAbsolutePath(), DocumentDescriptor.DEFAULT);
			
			debug_print_status();
			
			//The open method instantiates an jacobObjWord,
			//and then assigns a local wordObject,
			//and then instantiates the documents property to Documents,
			//and then opens the document and associates it with the object document.
			//Desired visibility is true.
			System.out.println("Calling openDocInWord("+myFile.getAbsolutePath()+",true)...");
			openDocInWord(myFile.getAbsolutePath(),true);
			
			debug_print_status();
			
			//End of MSWord_jsText adoption.
			
			System.out.println("MSWord_jsText: createEmptyDocument: About to createMe()...");
			createMe();

			System.out.println("MSWord_jsText: createEmptyDocument ends, about to return true");
			return true;
			
			/* ORIGINAL CODE FROM NOATEXT/OPENOFFICE IMPLEMENTATION:
			 * doc=(ITextDocument)office.getDocumentService().constructNewDocument(IDocument.WRITER,
			 * DocumentDescriptor.DEFAULT); if(doc!=null){
			 * doc.getPersistenceService().store(myFile.getAbsolutePath()); doc.close();
			 * panel.loadDocument(false, myFile.getAbsolutePath(), DocumentDescriptor.DEFAULT);
			 * doc=(ITextDocument)panel.getDocument(); return true; }
			 */

		} catch (Exception e) {
			ExHandler.handle(e);
			
		}

		System.out.println("MSWord_jsText: createEmptyDocument: try ... failed and/or caught Exception.");		
		System.out.println("MSWord_jsText: createEmptyDocument ends, about to return false");
		return false;
	}
	
	
	
	
	
	/**
	 * Load a file from a byte array. Again, we store it first into a temporary disk file because
	 * OOo does not like documents that have no representation on disk.
	 */
	public boolean loadFromByteArray(final byte[] bs, final boolean asTemplate){
		System.out.println("MSWord_jsText: loadFromByteArray(final byte[] bs, final boolean asTemplate) begins...");

		System.out.println("MSWord_jsText: loadFromByteArray: asTemplate == "+asTemplate);

		if (bs == null) { System.out.println("MSWord_jsText: loadFromByteArray: WARNING: bs IS NULL!");}
		else		    { System.out.println("MSWord_jsText: loadFromByteArray: bs == "+bs);}
			
		if (bs == null) {
			System.out.println("MSWord_jsText: loadFromByteArray: ERROR: bs IS NULL!");

			log.log("Null-Array zum speichern!", Log.ERRORS);
			
			System.out.println("MSWord_jsText: loadFromByteArray: about to end (early), returning false...");
			return false;
		}
		
		try {
			System.out.println("MSWord_jsText: loadFromByteArray: trying...");

			System.out.println("MSWord_jsText: loadFromByteArray: about to clean()...");
			clean();
			
			System.out.println("MSWord_jsText: loadFromByteArray: about to FileOutputStream fout = new FileOutputStream(myFile)...");
			
			FileOutputStream fout = new FileOutputStream(myFile);
			System.out.println("MSWord_jsText: loadFromByteArray: about to fout.write(bs)...");
			fout.write(bs);
			fout.close();
			
			
			//Beginning of MSWord_jsText adoption.

			System.out.println("MSWord_jsText: loadFromByteArray: TODO TODO TODO TODO TODO TODO TODO ");

			System.out.println("MSWord_jsTextNTI: *** TO DO: We would now call panel.loadDocument()");
			System.out.println("  to load a new document into the panel created before.");
			System.out.println("  FOR A START in the MS Word variant, we do not load a document INTO A PANEL,");
			System.out.println("  but we instruct Word to simply OPEN the document.");
			System.out.println("  The open() routine will read the file into the Dispatch document.");
			System.out.println("  Commented out the original code.");			
			
			System.out.println("MSWord_jsText: loadFromByteArray: TODO TODO TODO TODO TODO TODO TODO ");
			
			System.out.println("MSWord_jsText: loadFromByteArray: Commented out: panel.loadDocument(false, myFile.getAbsolutePath(), DocumentDescriptor.DEFAULT);");
			System.out.println("MSWord_jsText: loadFromByteArray: Replaced by:   open(myFile.getAbsolutePath(),true);");

			//panel.loadDocument(false, myFile.getAbsolutePath(), DocumentDescriptor.DEFAULT);
			
			debug_print_status();
			
			//The open method instantiates an jacobObjWord,
			//and then assigns a local wordObject,
			//and then instantiates the documents property to Documents,
			//and then opens the document and associates it with the object document.
			//Desired visibility is true.
			
			System.out.println("MSWord_jsText: loadFromByteArray: Calling openDocInWord("+myFile.getAbsolutePath()+",true)...");
			openDocInWord(myFile.getAbsolutePath(),true);
			
			debug_print_status();

			System.out.println("MSWord_jsText: loadFromByteArray: TODO TODO TODO TODO TODO TODO TODO ");
			
			//End of MSWord_jsText adoption.			
			
			System.out.println("MSWord_jsText: loadFromByteArray: about to CreateMe()...");
			createMe();
			
			System.out.println("MSWord_jsText: loadFromByteArray: about to end, returning true...");
			return true;
		} catch (Exception ex) {
			ExHandler.handle(ex);
			System.out.println("MSWord_jsText: loadFromByteArray: about to end, returning true...");
			return false;
		}
	}
	
	
	
	
	/**
	 * Load a file from an input stream. Explanations
	 * @see loadFromByteArray()
	 */
	public boolean loadFromStream(final InputStream is, final boolean asTemplate){
		System.out.println("MSWord_jsText: loadFromStream begins...");

		System.out.println("MSWord_jsText: TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO ");
		System.out.println("MSWord_jsText: loadFromStream: ToDo: Adopt to msword_js...");
		System.out.println("MSWord_jsText: This would probably require obtaining a word document from InputStream is, saving that to file, closing, reopening.");
		System.out.println("MSWord_jsText: TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO ");
		
		try {
			System.out.println("MSWord_jsText: loadFromStream: try...");
			System.out.println("MSWord_jsText: loadFromStream: About to clean()...");
		
			clean();

			/* ORIGINAL CODE FROM NOATEXT/OPENOFFICE
			//(try to) load a document from InputStream is
			System.out.println("MSWord_jsText: loadFromStream: About to doc = ... office.getDocumentService().loadDocument()...");
			agIonDoc = (ITextDocument) agIonOffice.getDocumentService().loadDocument(is, DocumentDescriptor.DEFAULT_HIDDEN);

			//If a document has been loaded, then store it to myFile.getAbsolutePath(); then close it; then load it again from myFile.getAbsolutePath()
			if (agIonDoc != null) {
				System.out.println("MSWord_jsText: loadFromStream: doc is not null - ok.");
				System.out.println("MSWord_jsText: loadFromStream: About to doc.getPersistenceService().store("+myFile.getAbsolutePath()+")...");
				agIonDoc.getPersistenceService().store(myFile.getAbsolutePath());
				System.out.println("MSWord_jsText: loadFromStream: About to doc.close()...");
				agIonDoc.close();
				
				System.out.println("MSWord_jsText: loadFromStream: About to panel.loadDocument(false, "+myFile.getAbsolutePath()+", DocumentDescriptor.DEFAULT)...");
				agIonPanel.loadDocument(false, myFile.getAbsolutePath(), DocumentDescriptor.DEFAULT);
			ORIGINAL CODE FROM NOATEXT/OPENOFFICE */
			
			//SO I don't really understand why this code is needed in addition to the other load_from...
			//and I don't really know whether some agIonDoc = (ITextDocument) agIonOffice.getDocumentService().loadDocument(is, DocumentDescriptor.DEFAULT_HIDDEN);
			//equivalent is available somewhere deep in Word's VBA object collection via JaCoB...
			//nor do I even know how to test this (can only see it's called from ch.elexis.views/TextView/makeActions()/new Action()/run()... 
			//But still I try to produce something that formally and maybe technically does in msword_js what the above lines did in noatext_jsl -
			//by first redirecting is into a file, then loading that file in Word - from similar code in createEmptyDocument().
			
			//PLEASE NOTE: MAYBE THIS IS THE IMPLEMENTATION BEHIND "Dokument Importieren" in the Briefe View???
			//And it produces 866+ pages of nonsense, when I use it to open an open office document (which I just randomly selected) into Word. 
			//BUT IT WORKED VERY WELL when I imported a little Word document that I just prepared for testing through the same way,
			//however, in the console output log, I see this error message now - whatever that means:
			
			/*
			MSWord_jsText: createMe: panel=OfficePanel {}
			OfficePanel: getDocument
			OfficePanel: WARNING: Please note: will return document==null
			MSWord_jsText: createMe: WARNING: doc==null, so we won't be able to doc.addCloseListener() or noas.add(this).
			MSWord_jsText: createMe: TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO
			MSWord_jsText: createMe: ToDo: Provide a replacement for closeListener and noas-keeping-track-of-opened-documents for msword_js...
			MSWord_jsText: createMe: TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO
			MSWord_jsText: createMe ends
			MSWord_jsText: TODO: CHECK: WHY ARE WE RETURNING FALSE HERE IN ANY CASE???
			MSWord_jsText: loadFromStream ends, about to return false
			js ch.elexis.views/TextView.java makeActions().importAction(): catching Throwable ex...
			--------------Exception--------------
			java.lang.NullPointerException
				at ch.elexis.views.TextView$7.run(TextView.java:536)
				at org.eclipse.jface.action.Action.runWithEvent(Action.java:498)
				at org.eclipse.jface.action.ActionContributionItem.handleWidgetSelection(ActionContributionItem.java:584)
				at org.eclipse.jface.action.ActionContributionItem.access$2(ActionContributionItem.java:501)
				at org.eclipse.jface.action.ActionContributionItem$5.handleEvent(ActionContributionItem.java:411)
				at org.eclipse.swt.widgets.EventTable.sendEvent(EventTable.java:84)
				at org.eclipse.swt.widgets.Widget.sendEvent(Widget.java:1053)
				at org.eclipse.swt.widgets.Display.runDeferredEvents(Display.java:4169)
				at org.eclipse.swt.widgets.Display.readAndDispatch(Display.java:3758)
				at org.eclipse.ui.internal.Workbench.runEventLoop(Workbench.java:2701)
				at org.eclipse.ui.internal.Workbench.runUI(Workbench.java:2665)
				at org.eclipse.ui.internal.Workbench.access$4(Workbench.java:2499)
				at org.eclipse.ui.internal.Workbench$7.run(Workbench.java:679)
				at org.eclipse.core.databinding.observable.Realm.runWithDefault(Realm.java:332)
				at org.eclipse.ui.internal.Workbench.createAndRunWorkbench(Workbench.java:668)
				at org.eclipse.ui.PlatformUI.createAndRunWorkbench(PlatformUI.java:149)
				at ch.elexis.Desk.start(Desk.java:175)
				at org.eclipse.equinox.internal.app.EclipseAppHandle.run(EclipseAppHandle.java:196)
				at org.eclipse.core.runtime.internal.adaptor.EclipseAppLauncher.runApplication(EclipseAppLauncher.java:110)
				at org.eclipse.core.runtime.internal.adaptor.EclipseAppLauncher.start(EclipseAppLauncher.java:79)
				at org.eclipse.core.runtime.adaptor.EclipseStarter.run(EclipseStarter.java:353)
				at org.eclipse.core.runtime.adaptor.EclipseStarter.run(EclipseStarter.java:180)
				at sun.reflect.NativeMethodAccessorImpl.invoke0(Native Method)
				at sun.reflect.NativeMethodAccessorImpl.invoke(NativeMethodAccessorImpl.java:57)
				at sun.reflect.DelegatingMethodAccessorImpl.invoke(DelegatingMethodAccessorImpl.java:43)
				at java.lang.reflect.Method.invoke(Method.java:606)
				at org.eclipse.equinox.launcher.Main.invokeFramework(Main.java:629)
				at org.eclipse.equinox.launcher.Main.basicRun(Main.java:584)
				at org.eclipse.equinox.launcher.Main.run(Main.java:1438)
				at org.eclipse.equinox.launcher.Main.main(Main.java:1414)
			-----------End Exception handler-----
			js ch.elexis.views/TextView.java makeActions().importAction.run(): end
			
			
			OR, in another attempt:
			
			
			MSWord_jsText: createMe: panel=OfficePanel {}
			OfficePanel: getDocument
			OfficePanel: WARNING: Please note: will return document==null
			MSWord_jsText: createMe: WARNING: doc==null, so we won't be able to doc.addCloseListener() or noas.add(this).
			MSWord_jsText: createMe: TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO
			MSWord_jsText: createMe: ToDo: Provide a replacement for closeListener and noas-keeping-track-of-opened-documents for msword_js...
			MSWord_jsText: createMe: TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO
			MSWord_jsText: createMe ends
			MSWord_jsText: TODO: CHECK: WHY ARE WE RETURNING FALSE HERE IN ANY CASE???
			MSWord_jsText: loadFromStream ends, about to return false
			js ch.elexis.views/TextView.java makeActions().importAction(): catching Throwable ex...
			--------------Exception--------------
			java.lang.NullPointerExceptionjs ch.elexis.views/TextView.java makeActions().importAction.run(): end
			
			
				at ch.elexis.views.TextView$7.run(TextView.java:536)
				at org.eclipse.jface.action.Action.runWithEvent(Action.java:498)
				at org.eclipse.jface.action.ActionContributionItem.handleWidgetSelection(ActionContributionItem.java:584)
				at org.eclipse.jface.action.ActionContributionItem.access$2(ActionContributionItem.java:501)
				at org.eclipse.jface.action.ActionContributionItem$5.handleEvent(ActionContributionItem.java:411)
				at org.eclipse.swt.widgets.EventTable.sendEvent(EventTable.java:84)
				at org.eclipse.swt.widgets.Widget.sendEvent(Widget.java:1053)
				at org.eclipse.swt.widgets.Display.runDeferredEvents(Display.java:4169)
				at org.eclipse.swt.widgets.Display.readAndDispatch(Display.java:3758)
				at org.eclipse.ui.internal.Workbench.runEventLoop(Workbench.java:2701)
				at org.eclipse.ui.internal.Workbench.runUI(Workbench.java:2665)
				at org.eclipse.ui.internal.Workbench.access$4(Workbench.java:2499)
				at org.eclipse.ui.internal.Workbench$7.run(Workbench.java:679)
				at org.eclipse.core.databinding.observable.Realm.runWithDefault(Realm.java:332)
				at org.eclipse.ui.internal.Workbench.createAndRunWorkbench(Workbench.java:668)
				at org.eclipse.ui.PlatformUI.createAndRunWorkbench(PlatformUI.java:149)
				at ch.elexis.Desk.start(Desk.java:175)
				at org.eclipse.equinox.internal.app.EclipseAppHandle.run(EclipseAppHandle.java:196)
				at org.eclipse.core.runtime.internal.adaptor.EclipseAppLauncher.runApplication(EclipseAppLauncher.java:110)
				at org.eclipse.core.runtime.internal.adaptor.EclipseAppLauncher.start(EclipseAppLauncher.java:79)
				at org.eclipse.core.runtime.adaptor.EclipseStarter.run(EclipseStarter.java:353)
				at org.eclipse.core.runtime.adaptor.EclipseStarter.run(EclipseStarter.java:180)
				at sun.reflect.NativeMethodAccessorImpl.invoke0(Native Method)
				at sun.reflect.NativeMethodAccessorImpl.invoke(NativeMethodAccessorImpl.java:57)
				at sun.reflect.DelegatingMethodAccessorImpl.invoke(DelegatingMethodAccessorImpl.java:43)
				at java.lang.reflect.Method.invoke(Method.java:606)
				at org.eclipse.equinox.launcher.Main.invokeFramework(Main.java:629)
				at org.eclipse.equinox.launcher.Main.basicRun(Main.java:584)
				at org.eclipse.equinox.launcher.Main.run(Main.java:1438)
				at org.eclipse.equinox.launcher.Main.main(Main.java:1414)
			-----------End Exception handler-----
			
			
			The problem may be located in the public Composite createContainer(final Composite parent, final ICallback handler){...} further above.
			*/
			
			Path path = new Path("rsc/empty.doc");		//201611131641js Umstellung von *.odt auf *.doc für MS-Word
			////Path path = new Path("rsc/empty.odt");
			//Path path = new Path("rsc/empty.docx");
			
			System.out.println("Current user directory is: "+System.getProperty("user.dir"));
						
			System.out.println("FileOutputStream os -- "+myFile);
			FileOutputStream fos = new FileOutputStream(myFile);
		
			System.out.println("copyStreams: InputStream is to FileOutputStream fos...");
			FileTool.copyStreams(is, fos);
			
			System.out.println("is.close()");
			is.close();
			
			System.out.println("os.close()");
			fos.close();
	
			System.out.println("Calling openDocInWord("+myFile.getAbsolutePath()+",true)...");
			openDocInWord(myFile.getAbsolutePath(),true);
			
			if (jacobDocument != null) {
				System.out.println("MSWord_jsText: loadFromStream: About to createMe()...");
				createMe();
			}
			else {
				System.out.println("MSWord_jsText: loadFromStream: WARNING: jacobDocument IS NULL! NOTHING to store or load...");				
			}
			
		} catch (Exception e) {
			ExHandler.handle(e);
			
		}

		System.out.println("MSWord_jsText: TODO: CHECK: WHY ARE WE RETURNING FALSE HERE IN ANY CASE???");
		System.out.println("MSWord_jsText: loadFromStream ends, about to return false");
		return false;
	}
	
	/**
	 * Store the contents of the OOo-Frame into a byte array. We save it into a temporary disk file
	 * first to ensure OOo, that the file ist really saved. That way OOo will not complain about
	 * corrupted or lost files.
	 */
	public byte[] storeToByteArray(){
		System.out.println("MSWord_jsText: storeToByteArray begins");
		
		//if (agIonDoc == null) {
		if (jacobDocument == null) {
			System.out.println("MSWord_jsText: storeToByteArray: WARNING: jacobDocument IS NULL!");
			System.out.println("MSWord_jsText: storeToByteArray: about to return null...");
			return null;
		} else {
			System.out.println("MSWord_jsText: storeToByteArray: INFO: jacobDocument="+jacobDocument.toString());
		}

		try {
			System.out.println("MSWord_jsText: storeToByteArray: try...");
			System.out.println("MSWord_jsText: storeToByteArray: jacobDocument is not null - ok.");
			
			/*
			System.out.println("MSWord_jsText: storeToByteArray: About to doc.getPersistenceService().store("+myFile.getAbsolutePath()+")...");
			agIonDoc.getPersistenceService().store(myFile.getAbsolutePath());
			*/
			
			String myFilename=myFile.getAbsolutePath();
			
			System.out.println("MSWord_jsText: storeToByteArray(): Trying to save the jacobDocument: "+myFilename+" from jacobDocument...");
			
			/*
			//THIS FAILS
			com.jacob.com.ComFailException: A COM exception has been encountered:
			At Invoke of: Save
			Description: 80020005 / Type mismatch.

			System.out.println("MSWord_jsText: open(): About to jacobDocument = Dispatch.call(jacobDocument, \"Save\", myFilename);...");
			Dispatch.call(jacobDocument, "Save", myFilename);
			*/
			
			//THIS WORKS
			System.out.println("MSWord_jsText: storeToByteArray(): About to Dispatch.call( (Dispatch) Dispatch.call(jacobObjWord, \"WordBasic\").getDispatch(),\"FileSaveAs\", myFilename);");
			Dispatch.call( (Dispatch) Dispatch.call(jacobObjWord, "WordBasic").getDispatch(),"FileSaveAs", myFilename); 

			System.out.println("MSWord_jsText: storeToByteArray: Now re-loading file content from "+myFilename+" into byte[] ret...");
			System.out.println("MSWord_jsText: storeToByteArray: BufferedInputStream bis = new BufferedInputStream(new FileInputStream(myFile));");
			BufferedInputStream bis = new BufferedInputStream(new FileInputStream(myFile));
			byte[] ret = new byte[(int) myFile.length()];
			System.out.println("MSWord_jsText: storeToByteArray: Reading begins. myFile.length == "+myFile.length());
			int pos = 0, len = 0;
			System.out.println("MSWord_jsText: storeToByteArray: about to read file via while (pos + len = bis.read(ret)) != ret.legnth) {}...");
			while (pos + (len = bis.read(ret)) != ret.length) {
				pos += len;
			}
			System.out.println("MSWord_jsText: storeToByteArray: Reading complete. Result: byte[] ret.length == "+ret.length);
			
			
			System.out.println("MSWord_jsText: storeToByteArray: about to end, returning byte[] ret... (to Elexis.TextView.java for storage as BLOB in DBMS)");
			return ret;
		} catch (Exception ex) {
			ExHandler.handle(ex);
			System.out.println("MSWord_jsText: storeToByteArray: Exception caught. About to return null");
			return null;
		}
	}
	
	
	
	/**
	 * Destroy the Panel with the OOo frame
	 */
	public void dispose(){
		System.out.println("MSWord_jsText: dispose begins - this should destroy the window with the MS Word document...");
		
		System.out.println("MSWord_jsText: removeMe: TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO");
		System.out.println("MSWord_jsText: removeMe: ToDo: TO REVIEW: !!!!!!!!!!!! Maybe should CHECK IF NOAS ARE EMPTY - DON'T want to jacobObjWord=null; jacobDocuments=null; when inappropriate! !!!!!!!!!!!!"); 
		System.out.println("MSWord_jsText: removeMe: ToDo: TO REVIEW: !!!!!!!!!!!! Cross-Check: RemoveMe(); Dispose(); !!!!!!!!!!!!"); 
		System.out.println("MSWord_jsText: removeMe: TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO");

		
		/* ORIGINAL CODE FROM NOA/OPENOFFICE
		System.out.println("MSWord_jsText: dispose begins - this should destroy the panel with the OOo frame...");
		if (agIonDoc != null) {
			System.out.println("MSWord_jsText: dispose: about to doc.close(); doc = null...");
			agIonDoc.close();
			agIonDoc = null;
		}
		else {System.out.println("MSWord_jsText: dispose: WARNING: doc already WAS NULL.");}
		ORIGINAL CODE FROM NOA/OPENOFFICE */
		
		if (jacobDocument != null) {
			System.out.println("MSWord_jsText: dispose: About to Dispatch.call(jacobDocument, \"Save\");");
		    Dispatch.call(jacobDocument, "Save");
		    //Dispatch.call(jacobDocument, "Close", new Variant(saveOnExit));
			System.out.println("MSWord_jsText: dispose: About to Dispatch.call(jacobDocument, \"Close\", new Variant(false)); jacobDocument = null;");
		    Dispatch.call(jacobDocument, "Close", new Variant(false));
		    jacobDocument = null; 
		}
		else {System.out.println("MSWord_jsText: dispose: WARNING: jacobDocument already WAS NULL.");}
		
		/* ORIGINAL CODE FROM NOA/OPENOFFICE
		if (agIonPanel != null) {
			System.out.println("MSWord_jsText: dispose: about to panel.dispose()...");
			agIonPanel.dispose();
		}
		else {System.out.println("MSWord_jsText: dispose: WARNING: panel already WAS NULL.");}
		ORIGINAL CODE FROM NOA/OPENOFFICE */

		if (jacobObjWord != null) {
			System.out.println("MSWord_jsText: dispose: About to panelDispatch.call(jacobObjWord, \"Quit\"); jacobObjWord = null; jacobDocuments = null;");
			Dispatch.call(jacobObjWord, "Quit");
			jacobObjWord = null;
			//jacobSelection = null;
	        jacobDocuments = null; 
		}
		else {System.out.println("MSWord_jsText: dispose: WARNING: jacobObjWord already WAS NULL.");}
			
		
		System.out.println("MSWord_jsText: dispose ends");
	}
	
	
	
	// The replaced noatext_jsl version (already updated re error message; please check vs current noatext_jsl content...
	/*
	public boolean findOrReplace(final String pattern, final ReplaceCallback cb){
		System.out.println("MSWord_jsText: findOrReplace begins");

		if (pattern == null)	System.out.println("MSWord_jsText: findOrReplace: WARNING: pattern IS NULL!");
		else					System.out.println("MSWord_jsText: findOrReplace: pattern="+pattern);

		SearchDescriptor search = new SearchDescriptor(pattern);
		search.setUseRegularExpression(true);

		if (agIonDoc == null) {
			//ToDo: Review the text of the error message, in German and English... - refers to No doc in bill, or Es ist keine Rechnungsvorlage definiert - may apply to other doc templates, too.
			System.out.println("MSWord_jsText: findOrReplace: TODO: please review the text of the error message, in German and English...");
		
			//SWTHelper.showError("No doc in bill", "Fehler:","Es ist keine Rechnungsvorlage definiert");
			SWTHelper.showError("findOrReplace: doc IS NULL", "Fehler:","findOrReplace: Statt eines Dokuments wurde NULL übergeben - möglicherweise fehlt die Dokumentenvorlage, z.B. Rechnungsvorlage.");
			
			System.out.println("MSWord_jsText: findOrReplace: ERROR: doc IS NULL. About to end, returning false");
			return false;
		}
			
		System.out.println("MSWord_jsText: findOrReplace: about to start (rather complex) replacement code...");

		// *** START support for replacement of placeholders inside Forms/Controls
		String cWrongNumOfArgs       = "*** Wrong number of arguments: Allowed number of arguments for this type of control: ";
		String cWrongNumOfArgs_2     = " ***";
		
		IFormService formService = agIonDoc.getFormService();
		IFormComponent[] formComponents;
		try {
			formComponents = formService.getFormComponents();
			for (int i = 0; i < formComponents.length; i++){	
				try	{
					IFormComponent formComponent = formComponents[i];
					// *** read control name - this may contain a replacement instruction
					XPropertySet xPSet = formComponent.getXPropertySet();
					int componentType = getFormComponentType(xPSet);
					try {
						String controlName = (String) xPSet.getPropertyValue("Name");
					} catch (UnknownPropertyException e) {
						break; // don't process if this can't be found
					} catch (WrappedTargetException e) {
						break; // don't process if this can't be found
					}
					
					// *** get the replacement specification
					String replacement = (String) xPSet.getPropertyValue("Tag");
					
					// *** do the replacement
					if (cb != null) {
						Pattern p = Pattern.compile(pattern, Pattern.CASE_INSENSITIVE);
						Matcher m = p.matcher(replacement);
						StringBuffer sb = new StringBuffer(replacement.length() * 4);
						while (m.find()) {
							int start = m.start();
							int end   = m.end();
							String orig = replacement.substring(start, end);
							Object replace = cb.replace(orig);
							if (replace == null) {
								m.appendReplacement(sb, "??Auswahl??");
							} else if (replace instanceof String) {
								String repl = ((String) replace).replaceAll("\\r", "\n");
								repl = repl.replaceAll("\\n\\n+", "\n");
								m.appendReplacement(sb, repl);
							} else {
								m.appendReplacement(sb, "Not a String");
							}
						}
						m.appendTail(sb);
						replacement = sb.toString();
					}
					
					// *** must save into Tag field because called repeatedly with different replacements
					xPSet.setPropertyValue("Tag", replacement);
					
					// *** split into parts
					String[] replacementParts = replacement.split("@@@");
					String replacement1 = replacementParts[0];
					String replacement2 = replacementParts.length >=2 ? replacementParts[1] : null;
					String replacement3 = replacementParts.length >=3 ? replacementParts[2] : null;
					String replacement4 = replacementParts.length >=4 ? replacementParts[3] : null;
					String replacement5 = replacementParts.length >=5 ? replacementParts[4] : null;
					
					if (StringTool.isNothing(replacement1)) replacement1 = null;
					if (StringTool.isNothing(replacement2)) replacement2 = null;
					if (StringTool.isNothing(replacement3)) replacement3 = null;
					if (StringTool.isNothing(replacement4)) replacement4 = null;
					if (StringTool.isNothing(replacement5)) replacement5 = null;
					
					//test if number of params ok,
					//if error break, show error info in tag field for debugging purposes
					String[] argumentsMapping = {
							FormComponentType.PATTERNFIELD  + ":" + 1,
							FormComponentType.FILECONTROL   + ":" + 1,
							FormComponentType.RADIOBUTTON   + ":" + 1,
							FormComponentType.CHECKBOX      + ":" + 1,
							FormComponentType.COMMANDBUTTON + ":" + 1,
							FormComponentType.FIXEDTEXT     + ":" + 1,
							FormComponentType.GROUPBOX      + ":" + 1,
							FormComponentType.IMAGEBUTTON   + ":" + 1,
							FormComponentType.IMAGECONTROL  + ":" + 1,
							FormComponentType.COMBOBOX      + ":" + 2,
							FormComponentType.LISTBOX       + ":" + 2,
							FormComponentType.DATEFIELD     + ":" + 3,
							FormComponentType.TIMEFIELD     + ":" + 3,
							FormComponentType.NUMERICFIELD  + ":" + 4,
							FormComponentType.SPINBUTTON    + ":" + 4,
							FormComponentType.CURRENCYFIELD + ":" + 4,
							FormComponentType.SCROLLBAR     + ":" + 5
					};
					for (int argi = 0; argi < argumentsMapping.length; argi++)	{
						String argMap = argumentsMapping[argi];
						int argType      = Integer.parseInt(argMap.split(":")[0]);
						int argNumOfArgs = Integer.parseInt(argMap.split(":")[1]);
						if (componentType == argType)	{
							String controlDefaultControl = (String) xPSet.getPropertyValue("DefaultControl");
							if (controlDefaultControl.equalsIgnoreCase("com.sun.star.form.control.FormattedField"))	{
								// *** special case FormattedField which is a text field
								if (replacementParts.length > argNumOfArgs) xPSet.setPropertyValue("Tag", cWrongNumOfArgs + 3 + cWrongNumOfArgs_2);
							} else	{
								// *** "normal" fields
								if (replacementParts.length > argNumOfArgs) xPSet.setPropertyValue("Tag", cWrongNumOfArgs + argNumOfArgs + cWrongNumOfArgs_2);
							}
							break;
						}
					}
					
					// if ComboBox or ListBox, then set list items if specified
					if ((componentType == FormComponentType.COMBOBOX) || (componentType == FormComponentType.LISTBOX))	{
						// *** if delimited by returns (coming from SQL-Select)
						replacement2 = replacement2.replaceAll("\\n", ";");
						if (replacement2 != null) xPSet.setPropertyValue("StringItemList",  replacement2.split(";"));
					}
					
					switch (componentType)	{
						case (FormComponentType.TEXTFIELD):
						case (FormComponentType.COMBOBOX):
						case (FormComponentType.PATTERNFIELD):
						case (FormComponentType.FILECONTROL):
							String controlDefaultControl = (String) xPSet.getPropertyValue("DefaultControl");
							if (controlDefaultControl.equalsIgnoreCase("com.sun.star.form.control.FormattedField"))	{
								// *** FormattedField
								if (isInteger(replacement1)) xPSet.setPropertyValue("EffectiveValue", new Short((short) Integer.parseInt(replacement1)));
								if (isInteger(replacement2)) xPSet.setPropertyValue("EffectiveMin",   new Short((short) Integer.parseInt(replacement2)));
								if (isInteger(replacement3)) xPSet.setPropertyValue("EffectiveMax",   new Short((short) Integer.parseInt(replacement3)));
							} else	{
								// *** simple text field
								XTextComponent xTextComponent = formComponent.getXTextComponent();
								if (replacement1 != null) xTextComponent.setText(replacement1);
							}
							break;
						case (FormComponentType.DATEFIELD):
							TimeTool timeTool = new TimeTool();
							// *** set date
							if (timeTool.set(replacement1))	{
								String yyyymmddDate = timeTool.toString(TimeTool.DATE_COMPACT);
								if (!StringTool.isNothing(yyyymmddDate)) xPSet.setPropertyValue("Date", Integer.parseInt(yyyymmddDate));
							}
							// *** set DateMin
							if (timeTool.set(replacement2))	{
								String yyyymmddDate = timeTool.toString(TimeTool.DATE_COMPACT);
								if (!StringTool.isNothing(yyyymmddDate)) xPSet.setPropertyValue("DateMin", Integer.parseInt(yyyymmddDate));
							}
							// *** set DateMax
							if (timeTool.set(replacement3))	{
								String yyyymmddDate = timeTool.toString(TimeTool.DATE_COMPACT);
								if (!StringTool.isNothing(yyyymmddDate)) xPSet.setPropertyValue("DateMax", Integer.parseInt(yyyymmddDate));
							}
							break;
						case (FormComponentType.TIMEFIELD):
							TimeTool timeTool2 = new TimeTool();
							// *** set time
							if (timeTool2.set(replacement1))	{
								String hhmmssTime = timeTool2.toString(TimeTool.TIME_FULL);
								hhmmssTime = hhmmssTime.replaceAll(":", "") + "00";
								if (!StringTool.isNothing(hhmmssTime)) xPSet.setPropertyValue("Time", Integer.parseInt(hhmmssTime));
							}
							// *** set TimeMin
							if (timeTool2.set(replacement2))	{
								String hhmmssTime = timeTool2.toString(TimeTool.TIME_FULL);
								hhmmssTime = hhmmssTime.replaceAll(":", "") + "00";
								if (!StringTool.isNothing(hhmmssTime))xPSet.setPropertyValue("TimeMin", Integer.parseInt(hhmmssTime));
							}
							// *** set TimeMax
							if (timeTool2.set(replacement3))	{
								String hhmmssTime = timeTool2.toString(TimeTool.TIME_FULL);
								hhmmssTime = hhmmssTime.replaceAll(":", "") + "00";
								if (!StringTool.isNothing(hhmmssTime)) xPSet.setPropertyValue("TimeMax", Integer.parseInt(hhmmssTime));
							}
							break;
						case (FormComponentType.NUMERICFIELD):
						case (FormComponentType.CURRENCYFIELD):
							if (isInteger(replacement1)) xPSet.setPropertyValue("Value",     new Short((short) Integer.parseInt(replacement1)));
							if (isInteger(replacement2)) xPSet.setPropertyValue("ValueMin",  new Short((short) Integer.parseInt(replacement2)));
							if (isInteger(replacement3)) xPSet.setPropertyValue("ValueMax",  new Short((short) Integer.parseInt(replacement3)));
							if (isInteger(replacement4)) xPSet.setPropertyValue("ValueStep", new Short((short) Integer.parseInt(replacement4)));
							break;
						case (FormComponentType.RADIOBUTTON):
						case (FormComponentType.CHECKBOX):
							if (isInteger(replacement1))  xPSet.setPropertyValue("State", new Short((short) Integer.parseInt(replacement1)));
							break;
						case (FormComponentType.COMMANDBUTTON):
						case (FormComponentType.FIXEDTEXT):
						case (FormComponentType.GROUPBOX):
							if (replacement1 != null) xPSet.setPropertyValue("Label", replacement1);
							break;
						case (FormComponentType.LISTBOX):
							// *** if delimited by returns (coming from SQL-Select)
							replacement1 = replacement1.replaceAll("\\n", ";");
							// *** create short[] from replacement1
							String[] splittedArgs = replacement1.split(";");
							short[] shortList = new short[splittedArgs.length];
							for (int argsi = 0; argsi < splittedArgs.length; argsi++)	{
								String argStr = splittedArgs[argsi];
								if (isInteger(argStr))	{
									short arg = (short) Integer.parseInt(argStr);
									shortList[argsi] = arg;
								}
							}
							if (replacement1 != null) xPSet.setPropertyValue("SelectedItems", shortList);
							break;
						case (FormComponentType.SPINBUTTON):
							if (isInteger(replacement3)) xPSet.setPropertyValue("SpinValueMax",  new Short((short) Integer.parseInt(replacement3)));
							if (isInteger(replacement2)) xPSet.setPropertyValue("SpinValueMin",  new Short((short) Integer.parseInt(replacement2)));
							if (isInteger(replacement4)) xPSet.setPropertyValue("SpinIncrement", new Short((short) Integer.parseInt(replacement4)));
							if (isInteger(replacement1)) xPSet.setPropertyValue("SpinValue",     new Short((short) Integer.parseInt(replacement1)));
							break;
						case (FormComponentType.SCROLLBAR):
							if (isInteger(replacement1)) xPSet.setPropertyValue("ScrollValue",    new Short((short) Integer.parseInt(replacement1)));
							if (isInteger(replacement2)) xPSet.setPropertyValue("ScrollValueMin", new Short((short) Integer.parseInt(replacement2)));
							if (isInteger(replacement3)) xPSet.setPropertyValue("ScrollValueMax", new Short((short) Integer.parseInt(replacement3)));
							if (isInteger(replacement4)) xPSet.setPropertyValue("LineIncrement",  new Short((short) Integer.parseInt(replacement4)));
							if (isInteger(replacement5)) xPSet.setPropertyValue("BlockIncrement", new Short((short) Integer.parseInt(replacement5)));
							break;
						case (FormComponentType.IMAGEBUTTON):
						case (FormComponentType.IMAGECONTROL):
							// *** doesn't work correctly... hmmmm... can anyone tell me how to get this to work???
							//     anyway: embedding into doc doesn't work in OO < 3.1
							//     so: more or less useless this way - and waiting for new OO in Elexis
 							if (replacement1 != null) xPSet.setPropertyValue("ImageURL", replacement1);
							break;
					}
				} catch (NOAException e) {
					e.printStackTrace();
				} catch (UnknownPropertyException e) {
					e.printStackTrace();
				} catch (PropertyVetoException e) {
					e.printStackTrace();
				} catch (IllegalArgumentException e) {
					e.printStackTrace();
				} catch (WrappedTargetException e) {
					e.printStackTrace();
				} catch (Exception e)	{
					// *** catch just everything so that the proc is going on...
				}
			}
		} catch (NOAException e1) {
			e1.printStackTrace();
		} catch (Exception e1) {
			// *** catch just everything so that the proc is going on...
			e1.printStackTrace();
		}
		// *** END support for replacement of placeholders inside Forms/Controls
		
		ISearchResult searchResult = agIonDoc.getSearchService().findAll(search);
		if (!searchResult.isEmpty()) {
			ITextRange[] textRanges = searchResult.getTextRanges();
			if (cb != null) {
				for (ITextRange r : textRanges) {
					String orig = r.getXTextRange().getString();
					Object replace = cb.replace(orig);
					if (replace == null) {
						r.setText("??Auswahl??");
					} else if (replace instanceof String) {
						// String repl=((String)replace).replaceAll("\\r\\n[\\r\\n]*", "\n")
						String repl = ((String) replace).replaceAll("\\r", "\n");
						repl = repl.replaceAll("\\n\\n+", "\n");
						r.setText(repl);
					} else if (replace instanceof String[][]) {
						String[][] contents = (String[][]) replace;
						try {
							ITextTable textTable =
								agIonDoc.getTextTableService().constructTextTable(contents.length,
									contents[0].length);
							agIonDoc.getTextService().getTextContentService().insertTextContent(r,
								textTable);
							r.setText("");
							ITextTablePropertyStore props = textTable.getPropertyStore();
							// long w=props.getWidth();
							// long percent=w/100;
							for (int row = 0; row < contents.length; row++) {
								String[] zeile = contents[row];
								for (int col = 0; col < zeile.length; col++) {
									textTable.getCell(col, row).getTextService().getText().setText(
										zeile[col]);
								}
							}
							textTable.spreadColumnsEvenly();
							
						} catch (Exception ex) {
							ExHandler.handle(ex);
							r.setText("Fehler beim Ersetzen");
						}
						
					} else {
						r.setText("Not a String");
					}
				}
			}
			System.out.println("MSWord_jsText: findOrReplace: about to end, returning true...");
			return true;
		}
		System.out.println("MSWord_jsText: findOrReplace: about to end, returning false...");
		return false;
	}
	*/
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	// The replacing msword_js version (already updated re error message; please check vs current noatext_jsl content...
	public boolean findOrReplace(final String pattern, final ReplaceCallback cb){
		System.out.println("MSWord_jsText: findOrReplace begins");
		
		Boolean debugSysPrnFindOrReplaceDetails = false;	//20170115js Das kann weiter unten bei Bedarf vorübergehend auf true umgestellt werden, habe ich z.B. im Bereich der findOrReplace (SectionHeaders) vorübergehend so benutzt und Codezeilen dafür noch dringelassen.
		if (!debugSysPrnFindOrReplaceDetails)
			System.out.println("MSWord_jsText: findOrReplace: INFO: debugSysPrnFindOrReplaceDetails==false, so DEBUG OUTPUT WILL BE LIMITED. This be changed in MSWord_jsText.java"); //201701021952js
		
		if (pattern == null)	System.out.println("MSWord_jsText: findOrReplace: WARNING: pattern IS NULL!");
		else					System.out.println("MSWord_jsText: findOrReplace: pattern="+pattern);

		if (jacobDocument == null) {
			//ToDo: Review the text of the error message, in German and English... - refers to No doc in bill, or Es ist keine Rechnungsvorlage definiert - may apply to other doc templates, too.
			System.out.println("MSWord_jsText: findOrReplace: TODO: please review the text of the error message, in German and English...");
		
			//SWTHelper.showError("No doc in bill", "Fehler:","Es ist keine Rechnungsvorlage definiert");
			SWTHelper.showError("findOrReplace: doc IS NULL", "Fehler:","findOrReplace: Statt eines Dokuments wurde NULL übergeben - möglicherweise fehlt die Dokumentenvorlage, z.B. Rechnungsvorlage.");
			
			System.out.println("MSWord_jsText: findOrReplace: ERROR: doc IS NULL. About to end, returning false");
			return false;
		}
		
		
		
		//In order to minimize flicker, we first set the WordObj to invisible...
		//It still flickers, because we have about 6 calls through findOrReplace() from Elexis.
		//So... we evaluate the patterns supplied from Elexis to decide whether we shall set Word to invisible or not :-)
		//This is the 1st pattern that Elexis sends, so we're at the beginning of the 1st pass through findOrReplace and switch Word back to invisible :-)
		if (pattern.equals("\\[[*]?[-a-zA-ZäöüÄÖÜéàè_ ]+\\.[-a-zA-Z0-9äöüÄÖÜéàè_ ]+\\]")) {
			//ToDo: Allenfalls wieder einschalten, wenn danach das MSWord Dokumentenwindow wieder nach vorne geholt werden kann: //20170201js commented out
			System.out.println("MSWord_jsText: findOrReplace: FOLGENDER CODE COMMENTED OUT, DA WORD DABEI LEICHT NACH HINTEN RUTSCHT:");		
			System.out.println("MSWord_jsText: findOrReplace: BITTE ERST DANN WIEDER EINFUEGEN, WENN WORD ANSCHLIESSEND NACH VORNE GEHOLT WERDEN KANN.");		
			System.out.println("MSWord_jsText: findOrReplace: SIEHE AUCH KORRESPONDIERENDEN CODE WEITER UNTEN, mit Visible/Variant(true).");		
			System.out.println("MSWord_jsText: findOrReplace: COMMENTED OUT: About to jacobObjWord.setProperty(\"Visible\", new Variant(false));");		
			//jacobObjWord.setProperty("Visible", new Variant(false));	 //20170201js commented out
		}

		
		if (debugSysPrnFindOrReplaceDetails) 
			System.out.println("MSWord_jsText: findOrReplace: On-the-fly translation of OO RegExp into MS Word Platzhalter-Suche compatible search patterns...");

		//On-the-fly translation der Wildcard-Suchen von OpenOffice RegExp nach MS Word Platzhalter-Suche.
		//Die \ müssen in Java als \\ escaped werden, damit sie während des Programmlaufs als \ erscheinen und nicht als Escape für das nachfolgende Zeichen.
		String pattern2 = null;
		
		//Korrekt wäre hier: "\\[[*]{0,1}[-a-zA-ZäöüÄÖÜéàè_ ]@.[-a-zA-Z0-9äöüÄÖÜéàè_ ]@\\]"; oder "\\[[*]{0;1}[-a-zA-ZäöüÄÖÜéàè_ ]@.[-a-zA-Z0-9äöüÄÖÜéàè_ ]@\\]";
		//je nach Ländereinstellung; aber beides lässt Word leider nicht zu. Ein {1} etc. geht übrigens, nur der Bereich geht nicht. Auch {;1} geht nicht.
		//Ebenso wird ^ unten im 3. bis 6. Pattern durch ? ersetzt...
		
		//OO:	\[[*]?[-a-zA-ZäöüÄÖÜéàè_ ]+\.[-a-zA-Z0-9äöüÄÖÜéàè_ ]+\]
		//Word:	\[?[\-a-zA-ZäöüÄÖÜéàè_ ]@.[\-a-zA-Z0-9äöüÄÖÜéàè_ ]@\]		//[\*]{0,1} oder [\*]{0;1} wäre korrekt für [*]? wird aber nicht akzeptiert.
		if (pattern.equals("\\[[*]?[-a-zA-ZäöüÄÖÜéàè_ ]+\\.[-a-zA-Z0-9äöüÄÖÜéàè_ ]+\\]"))
			pattern2 = "\\[?[\\-a-zA-ZäöüÄÖÜéàè_ ]@.[\\-a-zA-Z0-9äöüÄÖÜéàè_ ]@\\]";
		else
		//OO:	\[[*]?[-a-zA-ZäöüÄÖÜéàè_ ]+(\.[-a-zA-Z0-9äöüÄÖÜéàè_ ]+)+\]
		//Word:	\[?[\-a-zA-ZäöüÄÖÜéàè_ ]@(.[\-a-zA-Z0-9äöüÄÖÜéàè_ ]@)@\]	//[\*]{0,1} oder [\*]{0;1} wäre korrekt für [*]? wird aber nicht akzeptiert.
		if (pattern.equals("\\[[*]?[-a-zA-ZäöüÄÖÜéàè_ ]+(\\.[-a-zA-Z0-9äöüÄÖÜéàè_ ]+)+\\]"))
			pattern2 = "\\[?[\\-a-zA-ZäöüÄÖÜéàè_ ]@(.[\\-a-zA-Z0-9äöüÄÖÜéàè_ ]@)@\\]";	 
		else
		//OO:	\[[*]?[a-zA-Z]+:mwn?:[^\[]+\]
		//Word:	\[?[a-zA-Z]@:[mwn]@:[a-zA-Z/\-.,_0-9 ]@\]		//s.o. und: "^ nicht erlaubt" und ohnehin ging es erst nach vielen Versuchen.
		if (pattern.equals("\\[[*]?[a-zA-Z]+:mwn?:[^\\[]+\\]"))
			pattern2 = "\\[?[a-zA-Z]@:[mwn]@:[a-zA-Z/\\-.,_0-9 ]@\\]";
		else
		//OO:	\[[*]?[-_a-zA-Z0-9]+:[-a-zA-Z0-9]+:[-a-zA-Z0-9\.]+:[-a-zA-Z0-9\.]:?[^\]]*\]		//Dafür hab ich in meinem Brief kein Testsubstrat.
		//Word:	\[?[\-_a-zA-Z0-9]@:[\a-zA-Z0-9]@:[\-a-zA-Z0-9.]@:[\-a-zA-Z0-9.]:?[?\]]*\]
		if (pattern.equals("\\[[*]?[-_a-zA-Z0-9]+:[-a-zA-Z0-9]+:[-a-zA-Z0-9\\.]+:[-a-zA-Z0-9\\.]:?[^\\]]*\\]"))
			pattern2 = "\\[?[\\-_a-zA-Z0-9]@:[\\a-zA-Z0-9]@:[\\-a-zA-Z0-9.]@:[\\-a-zA-Z0-9.]:?[?\\]]*\\]";	 
		else
		//OO:	\[[*]?SQL[^:]*:[^\[]+\]					//Da findet auch das Original in meinem Test-Brief nichts, obwohl es eigentlich etwas finden sollte!
		//Word:	\[SQL*:*\]								//Das ist nun grob vereinfacht, geht aber eher nicht anders. Freundlicherweise ist Word nicht greedy.
		if (pattern.equals("\\[[*]?SQL[^:]*:[^\\[]+\\]"))	//Please note: My final SQL statement begins with: [SQL|\n|\n \n:select concat_ws(', ',concat(ifn...
			pattern2 = "\\[SQL*:*\\]"; 						//Therefore, we must accept characters between [SQL and :
		else
		//OO:	\[SCRIPT:[^\[]+\]
		//Word:	\[SCRIPT:[?\[]@\]
		if (pattern.equals("\\[SCRIPT:[^\\[]+\\]")) 	//Dafür hab ich in meinem Brief kein Testsubstrat.
			pattern2 = "\\[SCRIPT:[?\\[]@\\]";	 
		else
		pattern2 = pattern;
		
				

		if (debugSysPrnFindOrReplaceDetails) 
			System.out.println("MSWord_jsText: findOrReplace: About to start (rather complex) replacement code...");

		
		
		// *** START support for replacement of placeholders inside Forms/Controls
		
		System.out.println("MSWord_jsText: findOrReplace: TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO");
		System.out.println("MSWord_jsText: findOrReplace: ToDo: Review and find out whether we still need search and replace code for additional areas: e.g. footers (like: headers) etc.");
		System.out.println("MSWord_jsText: findOrReplace: ToDo: Check/Review Wildcard interpretation compatibility after on-the-fly translation OpenOffice -> Word");
		System.out.println("MSWord_jsText: findOrReplace: ToDo: Prüfen, ob Word dieselben Wildcards wie Elexis = Java/OpenOffice/UNO/NOA verwendet - andernfalls die o.g. Suchpatterns on-the-fly hier drin durch angepasste andere ersetzen!");
		System.out.println("MSWord_jsText: findOrReplace: ToDo: Allenfalls ein sanitizing von [Placeholder] und [Replacement] Text - so dass im SQL kein drop table... vorkommen darf, und im Replacement kein [...]... (Und insbesondere: Kein [SQL:drop table ...]... oder [SCRIPT: ...]");
		System.out.println("MSWord_jsText: findOrReplace: TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO");
	

		//Initialize total search results counters for this pattern: 
		
		Integer jacobSearchResultInt = 0;
		Integer numberOfHits = 0;

		


		
		
		
		
		
		
		
		

		
		//The following block performs search-and-replace for the main text portion of the document.
		//An almost identical block is further below for shapes (without updating all the comments),
		//and will probably be added further below, for tables. 
		
		if (debugSysPrnFindOrReplaceDetails) { 
			System.out.println("MSWord_jsText: findOrReplace: ");
			System.out.println("MSWord_jsText: findOrReplace: About to find/replace in the main text block:");
			System.out.println("MSWord_jsText: findOrReplace: ");
		}
		
		System.out.println("MSWord_jsText: findOrReplace: ");
		System.out.println("MSWord_jsText: findOrReplace: TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO ");
		System.out.println("MSWord_jsText: findOrReplace: ToDo: ARE WE REALLY SURE HERE, WHICH OBJECT WE SHOULD PROCESS (i.e. that we're processing the correct document in Word if multiple docs are open???)");
		System.out.println("MSWord_jsText: findOrReplace: ToDo: Inability to use ActiveXComponents instead of Dispatch... might arise below from the fact that we have NOT used s.th. like:");
		System.out.println("MSWord_jsText: findOrReplace: ToDo: ActiveXComponent jacobDocuments = ...");
		System.out.println("MSWord_jsText: findOrReplace: ToDo: ActiveXComponent jacobDocument = ...");
		System.out.println("MSWord_jsText: findOrReplace: ToDo: and just work on \"something\" instead.");
		System.out.println("MSWord_jsText: findOrReplace: TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO ");
		System.out.println("MSWord_jsText: findOrReplace: ");
		
		
		
		//For the main block of non-form documents, begin searching at the top of the document.
		
		//Die Variablen wurden oben schon angelegt, ich mach aber die Zuweisungen frisch, falls die inzwischen geändert worden wären.
		ActiveXComponent jacobSelection = jacobObjWord.getPropertyAsComponent("Selection");
		ActiveXComponent jacobFind = jacobSelection.getPropertyAsComponent("Find");
		
		//Put the cursor (back) to the document top. This necessary, or all searches but the first one will most probably NOT find anything replacable: 
		//N.B.: We want to support multiple replacements within another; i.e. for an [SQL:select...] Query that has [Patient.Name] (or similar) as an argument.
		
		if (debugSysPrnFindOrReplaceDetails) 
			System.out.println("MSWord_jsText: findOrReplace: About to Dispatch.call(jacobSelection, \"HomeKey\", \"wdStory\", \"wdMove\");");
	    //Dispatch.call(jacobSelection, "HomeKey", "wdStory", "wdMove");
		Dispatch.call(jacobSelection, "HomeKey", new Variant(6));


	    //The following works (developed and tested above in main() ):
	    //Dispatch mySelection = Dispatch.get(oWord, "Selection").toDispatch();
		//Dispatch.call(mySelection, "HomeKey", new Variant(6));
	    //oSelection.setProperty("Text", "InsertStuffatCursorPosAfterHomeKey");

		//Dispatch.put(jacobWordObject, "Visible", new Variant(visible));
		
		//Tatsächlich fügt eine (erste..., besser: siehe unten :-)  ) Suche nach pattern und Ersetzen durch "Replaced" ganz links oben das Wort "Replaced" ein,
		//der eigentlich angestrebte Suchtext (mehrfaches Vorkommen von [...] mit Wildcards zwischendrin wird NICHT gefunden oder ersetzt.
		
		//Wenn ich als Suchtext aber "Mandant" übergebe, dann wird das erste Auftreten von ebendiesem Suchtext gefunden und dort ersetzt. :-)
		//Nein, nach dem Hinzufügen von MoveRight nach der Ersetzung werden sogar ALLE Auftreten von "Mandant" durch "Replaced" ersetzt. :-)
		
		//Laut log output wird die Methode von Elexis mit folgenden patterns aufgerufen, und zwar in aufeinanderfolgenden Aufrufen:
		//pattern=\[[*]?[-a-zA-ZäöüÄÖÜéàè_ ]+\.[-a-zA-Z0-9äöüÄÖÜéàè_ ]+\]
		//pattern=\[[*]?[-a-zA-ZäöüÄÖÜéàè_ ]+(\.[-a-zA-Z0-9äöüÄÖÜéàè_ ]+)+\]
		//pattern=\[[*]?[a-zA-Z]+:mwn?:[^\[]+\]
		//pattern=\[[*]?[-_a-zA-Z0-9]+:[-a-zA-Z0-9]+:[-a-zA-Z0-9\.]+:[-a-zA-Z0-9\.]:?[^\]]*\]
		//pattern=\[[*]?SQL[^:]*:[^\[]+\]
		//pattern=\[SCRIPT:[^\[]+\]
		
		//D.h. wenn ich testweise mit einem fixen pattern suche (während der Entwicklung, siehe unten),
		//darf ich schon mal bis zu 6 Hits und Replacements erwarten, wenn jeder Durchlauf hier nur einen Hit produziert.
		
		//N.B.: mehrstufige Ersetzungen durch Platzhalter: Sollen die Funktionieren?
		//Ich denke eher nicht, sonst könnte man auch endlose Loops erreichen, indem man z.B. als Name in der Kontaktdatenbank [Adressat.Name] angibt...
		//Wenn man das wollte, bräuchte man eher: MoveLeft (or possibly: MoveToDocumentTop oder so ähnlich) statt MoveRight. 
		
		//Simplified serach patterns for testing, research and development:
		//jacobFind.setProperty("Text", "Mandant");				//Das findet den Suchtext "Mandant", tritt oben links mehrfach auf, funktioniert
		//jacobFind.setProperty("Text", "Adressat");			//Das findet den Suchtext "Adressat", funktioniert im Text (z.B. Anrede), ABER (mag auch an der Zahl der Aufrufe vs. Vorkommen liegen) NICHT IM Adressfeld (eigenes Textfeld) !!!
		//jacobFind.setProperty("Text", "[");					//Das findet den Suchtext "[", funktioniert
		
		//jacobFind.setProperty("Text", "\\[?*\\]");			//Das soll den Suchtext "\[?*\]" finden (also: "[", dann: beliebig viele beliebige Zeichen, dann "]").
																//In Word interaktiv funktioniert das, wenn man Pattern-Matching dort einschaltet. Hier nicht ohne weiteres.
																//Es kann aber sein, dass Word die Patterns anders als Java/OpenOffice/UNO/NOA RegEx interpretiert!
		
		
		/* Ein in Word aufgezeichnetes Makro lautet so:
		 * Sub SucheMitWildcards()
'
' SucheMitWildcards Makro
' Makro aufgezeichnet am 22.09.2016 von Jörg M. Sigle
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "\[?@\]"
        .Replacement.Text = "FLUP"
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub

		 */
				
		//Das setze ich jetzt wie folgt für Elexis - Java - JACOB um -
		//dabei hatte ich auch falsche übernommene Identifiers mit drin, nämlich diese hier:
		//
		//jacobFind.setProperty("MatchCase", "False");
		//jacobFind.setProperty("MatchAllowWordForms", "False");
		//
		//Dabei finde ich, dass bei nicht existierenden Feldern jeweils eine Java Exception geworfen wird,
		//mit rotem Output im console log, hier für das zweite Beispiel:
		
		/*		
		--------------Exception--------------
		com.jacob.com.ComFailException: Can't map name to dispid: MatchAllowWordForms
			at com.jacob.com.Dispatch.invokev(Native Method)
			at com.jacob.com.Dispatch.invokev(Dispatch.java:625)
			at com.jacob.com.Dispatch.invoke(Dispatch.java:498)
			at com.jacob.com.Dispatch.put(Dispatch.java:580)
			at com.jacob.activeX.ActiveXComponent.setProperty(ActiveXComponent.java:239)
			at com.jacob.activeX.ActiveXComponent.setProperty(ActiveXComponent.java:261)
			at com.jsigle.msword_js.MSWord_jsText.findOrReplace(MSWord_jsText.java:1565)
			at ch.elexis.text.TextContainer.createFromTemplate(TextContainer.java:276)
			at ch.elexis.views.TextView.createDocument(TextView.java:406)
			at ch.elexis.views.BriefAuswahl$4.run(BriefAuswahl.java:323)
			at org.eclipse.jface.action.Action.runWithEvent(Action.java:498)
			at org.eclipse.jface.action.ActionContributionItem.handleWidgetSelection(ActionContributionItem.java:584)
			at org.eclipse.jface.action.ActionContributionItem.access$2(ActionContributionItem.java:501)
			at org.eclipse.jface.action.ActionContributionItem$6.handleEvent(ActionContributionItem.java:452)
			at org.eclipse.swt.widgets.EventTable.sendEvent(EventTable.java:84)
			at org.eclipse.swt.widgets.Widget.sendEvent(Widget.java:1053)
			at org.eclipse.swt.widgets.Display.runDeferredEvents(Display.java:4169)
			at org.eclipse.swt.widgets.Display.readAndDispatch(Display.java:3758)
			at org.eclipse.ui.internal.Workbench.runEventLoop(Workbench.java:2701)
			at org.eclipse.ui.internal.Workbench.runUI(Workbench.java:2665)
			at org.eclipse.ui.internal.Workbench.access$4(Workbench.java:2499)
			at org.eclipse.ui.internal.Workbench$7.run(Workbench.java:679)
			at org.eclipse.core.databinding.observable.Realm.runWithDefault(Realm.java:332)
			at org.eclipse.ui.internal.Workbench.createAndRunWorkbench(Workbench.java:668)
			at org.eclipse.ui.PlatformUI.createAndRunWorkbench(PlatformUI.java:149)
			at ch.elexis.Desk.start(Desk.java:175)
			at org.eclipse.equinox.internal.app.EclipseAppHandle.run(EclipseAppHandle.java:196)
			at org.eclipse.core.runtime.internal.adaptor.EclipseAppLauncher.runApplication(EclipseAppLauncher.java:110)
			at org.eclipse.core.runtime.internal.adaptor.EclipseAppLauncher.start(EclipseAppLauncher.java:79)
			at org.eclipse.core.runtime.adaptor.EclipseStarter.run(EclipseStarter.java:353)
			at org.eclipse.core.runtime.adaptor.EclipseStarter.run(EclipseStarter.java:180)
			at sun.reflect.NativeMethodAccessorImpl.invoke0(Native Method)
			at sun.reflect.NativeMethodAccessorImpl.invoke(NativeMethodAccessorImpl.java:57)
			at sun.reflect.DelegatingMethodAccessorImpl.invoke(DelegatingMethodAccessorImpl.java:43)
			at java.lang.reflect.Method.invoke(Method.java:606)
			at org.eclipse.equinox.launcher.Main.invokeFramework(Main.java:629)
			at org.eclipse.equinox.launcher.Main.basicRun(Main.java:584)
			at org.eclipse.equinox.launcher.Main.run(Main.java:1438)
			at org.eclipse.equinox.launcher.Main.main(Main.java:1414)
		-----------End Exception handler-----
		*/

		
		
		
		
		//The following block performs search-and-replace for the main text portion of the document.
		//An almost identical block is further below for shapes (without updating all the comments),
		//and will probably be added further below, for tables. 
		
		if (debugSysPrnFindOrReplaceDetails) { 
			System.out.println("MSWord_jsText: findOrReplace: ");
			System.out.println("MSWord_jsText: findOrReplace: About to find/replace in the main text block:");
			System.out.println("MSWord_jsText: findOrReplace: ");
		}
		
		System.out.println("MSWord_jsText: findOrReplace: ");
		System.out.println("MSWord_jsText: findOrReplace: TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO ");
		System.out.println("MSWord_jsText: findOrReplace: ToDo: ARE WE REALLY SURE HERE, WHICH OBJECT WE SHOULD PROCESS (i.e. that we're processing the correct document in Word if multiple docs are open???)");
		System.out.println("MSWord_jsText: findOrReplace: ToDo: Inability to use ActiveXComponents instead of Dispatch... might arise below from the fact that we have NOT used s.th. like:");
		System.out.println("MSWord_jsText: findOrReplace: ToDo: ActiveXComponent jacobDocuments = ...");
		System.out.println("MSWord_jsText: findOrReplace: ToDo: ActiveXComponent jacobDocument = ...");
		System.out.println("MSWord_jsText: findOrReplace: ToDo: and just work on \"something\" instead.");
		System.out.println("MSWord_jsText: findOrReplace: TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO ");
		System.out.println("MSWord_jsText: findOrReplace: ");
		
		//For the main block of non-form documents, begin searching at the top of the document.
		
		//Die Variablen wurden oben schon angelegt, ich mach aber die Zuweisungen frisch, falls die inzwischen geändert worden wären.
		//jacobSelection = jacobObjWord.getPropertyAsComponent("Selection");
		//ActiveXComponent jacobFind = jacobSelection.getPropertyAsComponent("Find");
		//jacobSelection = jacobObjWord.getPropertyAsComponent("Selection");
		//jacobFind = jacobSelection.getPropertyAsComponent("Find");

		//Put the cursor (back) to the document top. This necessary, or all searches but the first one will most probably NOT find anything replacable: 
		//N.B.: We want to support multiple replacements within another; i.e. for an [SQL:select...] Query that has [Patient.Name] (or similar) as an argument.
		
		if (debugSysPrnFindOrReplaceDetails) 
			System.out.println("MSWord_jsText: findOrReplace: About to Dispatch.call(jacobSelection, \"HomeKey\", \"wdStory\", \"wdMove\");");
	    //Dispatch.call(jacobSelection, "HomeKey", "wdStory", "wdMove");
		Dispatch.call(jacobSelection, "HomeKey", new Variant(6));


	    //The following works (developed and tested above in main() ):
	    //Dispatch mySelection = Dispatch.get(oWord, "Selection").toDispatch();
		//Dispatch.call(mySelection, "HomeKey", new Variant(6));
	    //oSelection.setProperty("Text", "InsertStuffatCursorPosAfterHomeKey");

		//Dispatch.put(jacobWordObject, "Visible", new Variant(visible));
		
		//Tatsächlich fügt eine (erste..., besser: siehe unten :-)  ) Suche nach pattern und Ersetzen durch "Replaced" ganz links oben das Wort "Replaced" ein,
		//der eigentlich angestrebte Suchtext (mehrfaches Vorkommen von [...] mit Wildcards zwischendrin wird NICHT gefunden oder ersetzt.
		
		//Wenn ich als Suchtext aber "Mandant" übergebe, dann wird das erste Auftreten von ebendiesem Suchtext gefunden und dort ersetzt. :-)
		//Nein, nach dem Hinzufügen von MoveRight nach der Ersetzung werden sogar ALLE Auftreten von "Mandant" durch "Replaced" ersetzt. :-)
		
		//Laut log output wird die Methode von Elexis mit folgenden patterns aufgerufen, und zwar in aufeinanderfolgenden Aufrufen:
		//pattern=\[[*]?[-a-zA-ZäöüÄÖÜéàè_ ]+\.[-a-zA-Z0-9äöüÄÖÜéàè_ ]+\]
		//pattern=\[[*]?[-a-zA-ZäöüÄÖÜéàè_ ]+(\.[-a-zA-Z0-9äöüÄÖÜéàè_ ]+)+\]
		//pattern=\[[*]?[a-zA-Z]+:mwn?:[^\[]+\]
		//pattern=\[[*]?[-_a-zA-Z0-9]+:[-a-zA-Z0-9]+:[-a-zA-Z0-9\.]+:[-a-zA-Z0-9\.]:?[^\]]*\]
		//pattern=\[[*]?SQL[^:]*:[^\[]+\]
		//pattern=\[SCRIPT:[^\[]+\]
		
		//D.h. wenn ich testweise mit einem fixen pattern suche (während der Entwicklung, siehe unten),
		//darf ich schon mal bis zu 6 Hits und Replacements erwarten, wenn jeder Durchlauf hier nur einen Hit produziert.
		
		//N.B.: mehrstufige Ersetzungen durch Platzhalter: Sollen die Funktionieren?
		//Ich denke eher nicht, sonst könnte man auch endlose Loops erreichen, indem man z.B. als Name in der Kontaktdatenbank [Adressat.Name] angibt...
		//Wenn man das wollte, bräuchte man eher: MoveLeft (or possibly: MoveToDocumentTop oder so ähnlich) statt MoveRight. 
		
		//Simplified serach patterns for testing, research and development:
		//jacobFind.setProperty("Text", "Mandant");				//Das findet den Suchtext "Mandant", tritt oben links mehrfach auf, funktioniert
		//jacobFind.setProperty("Text", "Adressat");			//Das findet den Suchtext "Adressat", funktioniert im Text (z.B. Anrede), ABER (mag auch an der Zahl der Aufrufe vs. Vorkommen liegen) NICHT IM Adressfeld (eigenes Textfeld) !!!
		//jacobFind.setProperty("Text", "[");					//Das findet den Suchtext "[", funktioniert
		
		//jacobFind.setProperty("Text", "\\[?*\\]");			//Das soll den Suchtext "\[?*\]" finden (also: "[", dann: beliebig viele beliebige Zeichen, dann "]").
																//In Word interaktiv funktioniert das, wenn man Pattern-Matching dort einschaltet. Hier nicht ohne weiteres.
																//Es kann aber sein, dass Word die Patterns anders als Java/OpenOffice/UNO/NOA RegEx interpretiert!
		
		
		/* Ein in Word aufgezeichnetes Makro lautet so:
		 * Sub SucheMitWildcards()
'
' SucheMitWildcards Makro
' Makro aufgezeichnet am 22.09.2016 von Jörg M. Sigle
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "\[?@\]"
        .Replacement.Text = "FLUP"
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub

		 */
				
		//Das setze ich jetzt wie folgt für Elexis - Java - JACOB um -
		//dabei hatte ich auch falsche übernommene Identifiers mit drin, nämlich diese hier:
		//
		//jacobFind.setProperty("MatchCase", "False");
		//jacobFind.setProperty("MatchAllowWordForms", "False");
		//
		//Dabei finde ich, dass bei nicht existierenden Feldern jeweils eine Java Exception geworfen wird,
		//mit rotem Output im console log, hier für das zweite Beispiel:
		
		/*		
		--------------Exception--------------
		com.jacob.com.ComFailException: Can't map name to dispid: MatchAllowWordForms
			at com.jacob.com.Dispatch.invokev(Native Method)
			at com.jacob.com.Dispatch.invokev(Dispatch.java:625)
			at com.jacob.com.Dispatch.invoke(Dispatch.java:498)
			at com.jacob.com.Dispatch.put(Dispatch.java:580)
			at com.jacob.activeX.ActiveXComponent.setProperty(ActiveXComponent.java:239)
			at com.jacob.activeX.ActiveXComponent.setProperty(ActiveXComponent.java:261)
			at com.jsigle.msword_js.MSWord_jsText.findOrReplace(MSWord_jsText.java:1565)
			at ch.elexis.text.TextContainer.createFromTemplate(TextContainer.java:276)
			at ch.elexis.views.TextView.createDocument(TextView.java:406)
			at ch.elexis.views.BriefAuswahl$4.run(BriefAuswahl.java:323)
			at org.eclipse.jface.action.Action.runWithEvent(Action.java:498)
			at org.eclipse.jface.action.ActionContributionItem.handleWidgetSelection(ActionContributionItem.java:584)
			at org.eclipse.jface.action.ActionContributionItem.access$2(ActionContributionItem.java:501)
			at org.eclipse.jface.action.ActionContributionItem$6.handleEvent(ActionContributionItem.java:452)
			at org.eclipse.swt.widgets.EventTable.sendEvent(EventTable.java:84)
			at org.eclipse.swt.widgets.Widget.sendEvent(Widget.java:1053)
			at org.eclipse.swt.widgets.Display.runDeferredEvents(Display.java:4169)
			at org.eclipse.swt.widgets.Display.readAndDispatch(Display.java:3758)
			at org.eclipse.ui.internal.Workbench.runEventLoop(Workbench.java:2701)
			at org.eclipse.ui.internal.Workbench.runUI(Workbench.java:2665)
			at org.eclipse.ui.internal.Workbench.access$4(Workbench.java:2499)
			at org.eclipse.ui.internal.Workbench$7.run(Workbench.java:679)
			at org.eclipse.core.databinding.observable.Realm.runWithDefault(Realm.java:332)
			at org.eclipse.ui.internal.Workbench.createAndRunWorkbench(Workbench.java:668)
			at org.eclipse.ui.PlatformUI.createAndRunWorkbench(PlatformUI.java:149)
			at ch.elexis.Desk.start(Desk.java:175)
			at org.eclipse.equinox.internal.app.EclipseAppHandle.run(EclipseAppHandle.java:196)
			at org.eclipse.core.runtime.internal.adaptor.EclipseAppLauncher.runApplication(EclipseAppLauncher.java:110)
			at org.eclipse.core.runtime.internal.adaptor.EclipseAppLauncher.start(EclipseAppLauncher.java:79)
			at org.eclipse.core.runtime.adaptor.EclipseStarter.run(EclipseStarter.java:353)
			at org.eclipse.core.runtime.adaptor.EclipseStarter.run(EclipseStarter.java:180)
			at sun.reflect.NativeMethodAccessorImpl.invoke0(Native Method)
			at sun.reflect.NativeMethodAccessorImpl.invoke(NativeMethodAccessorImpl.java:57)
			at sun.reflect.DelegatingMethodAccessorImpl.invoke(DelegatingMethodAccessorImpl.java:43)
			at java.lang.reflect.Method.invoke(Method.java:606)
			at org.eclipse.equinox.launcher.Main.invokeFramework(Main.java:629)
			at org.eclipse.equinox.launcher.Main.basicRun(Main.java:584)
			at org.eclipse.equinox.launcher.Main.run(Main.java:1438)
			at org.eclipse.equinox.launcher.Main.main(Main.java:1414)
		-----------End Exception handler-----
		*/

		
		
		//Before we do anything else, quickly try to recognize whether this is a Tarmedrechnung_xx template
		//with a [Titel] Placeholder, that must be protected from messing with its Header.Range at all.
		//We must do this in every pass if the identification has not succeeded yet,
		//because searches for other patterns from Elexis would otherwise reach the Header area and
		//cause the Tarmedrechnung_xx templates to become messed up.
		
		if ( !ProbablyUsingTarmed_xxTemplateSoDoNOTAccessHeaderRangeToAvoidGenerationOfEmptyHeaderLines ) {
			if (debugSysPrnFindOrReplaceDetails) 
				System.out.println("MSWord_jsText: findOrReplace: INFO: Checking whether this might be a Tarmedrechnung_xx template whose Header.Range MUST NOT BE ACCESSED AT ALL...");
			jacobSearchResultInt = 0;
			
			try {
				jacobFind.setProperty("Text","\\[Titel\\]");
				jacobFind.setProperty("Forward", "True");
				jacobFind.setProperty("Format", "False");
				jacobFind.setProperty("MatchCase", "False");			//Hatte zuvor hier: MatchCase=True; MatchWildcards=False. Eigentlich passend, aber damit keinen Hit bekommen?!?
				jacobFind.setProperty("MatchWholeWord", "False");
				jacobFind.setProperty("MatchByte", "False");
				jacobFind.setProperty("MatchAllWordForms", "False");
				jacobFind.setProperty("MatchSoundsLike", "False");
				jacobFind.setProperty("MatchWildcards", "True");	
			} catch (Exception ex) {
				ExHandler.handle(ex);
				System.out.println("MSWord_jsText: findOrReplace: Fehler bei jacobSelection.setProperty(\"...\", \"...\");");
				SWTHelper.showError("findOrReplace:", "Fehler:","findOrReplace: Fehler bei bei jacobSelection.setProperty(\"...\", \"...\");");
			}
						
			try {
				jacobSearchResultInt = jacobFind.invoke("Execute").toInt();
				if (debugSysPrnFindOrReplaceDetails)
					System.out.println("MSWord_jsText: findOrReplace: jacobSearchResultInt="+jacobSearchResultInt);
			} catch (Exception ex) {
				ExHandler.handle(ex);
				System.out.println("MSWord_jsText: findOrReplace (Haupttext): Fehler bei Suche nach \\[Titel\\].");
				SWTHelper.showError(
						"MSWord_jsText: findOrReplace (Haupttext):", 
						"Fehler:",
						"Exception caught.\n"+
						"Bei Suche nach \\[Titel\\]");
			}
			
			if (jacobSearchResultInt == -1) { 
				ProbablyUsingTarmed_xxTemplateSoDoNOTAccessHeaderRangeToAvoidGenerationOfEmptyHeaderLines = true; 
				System.out.println("MSWord_jsText: findOrReplace: INFO: \\[Titel\\] FOUND in main text of document. This might be a Tarmedrechnung_xx template whose Header.Range MUST NOT BE ACCESSED AT ALL!");
			} else {
				System.out.println("MSWord_jsText: findOrReplace: INFO: \\[Titel\\] NOT found in main text of document. The document Header may be accessed in this pass.");
			}
		}
		
		
		
		
	
		
		
		
		
		try {
			//Please note: We must ensure, that NO prior path disrupts placeholders that would be valid and recognized in a subsequent pass!
			//jacobFind.setProperty("Text", "\\[[*]?[-a-zA-ZäöüÄÖÜéàè_ ]+\\.[-a-zA-Z0-9äöüÄÖÜéàè_ ]+\\]");  //The pattern supplied by Elexis in the first pass... 

			jacobFind.setProperty("Text", pattern2);
			
			//jacobFind.setProperty("Text", pattern);			//This would set the actually supplied Elexis search pattern. 
			//jacobFind.setProperty("Text", "\\[?*\\]");		//This searches for [...] placeholder markers of any Elexis-Placeholder-Type, used for development.
																//Please note: This produces hits *WITHIN PORTIONS* of certain [SQL...] placeholders, ripping those into parts. 
			//jacobFind.setProperty("Text", "Humpfidumpfi");	//This will most probably produce a search miss. Used to obtain the result of that situation.
			jacobFind.setProperty("Forward", "True");
			jacobFind.setProperty("Format", "False");
			jacobFind.setProperty("MatchCase", "False");
			jacobFind.setProperty("MatchWholeWord", "False");
			jacobFind.setProperty("MatchByte", "False");
			jacobFind.setProperty("MatchAllWordForms", "False");
			jacobFind.setProperty("MatchSoundsLike", "False");
			jacobFind.setProperty("MatchWildcards", "True");	
		} catch (Exception ex) {
			ExHandler.handle(ex);
			System.out.println("MSWord_jsText: findOrReplace: Fehler bei jacobSelection.setProperty(\"...\", \"...\");");
			
			//SWTHelper.showError("No doc in bill", "Fehler:","Es ist keine Rechnungsvorlage definiert");
			SWTHelper.showError("findOrReplace:", "Fehler:","findOrReplace: Fehler bei bei jacobSelection.setProperty(\"...\", \"...\");");
		}
					
		try {	 
			do {
				jacobSearchResultInt = 0;	//Reset this to 0 in each loop iteration, so that even when jacobFind.invoke() should NOT put a valid result into this variable
											//(or when it should not even get invoked!),
											//we do definitely NOT get an endless loop, NOR a misguided attempt to replace text. 	
				
				if (debugSysPrnFindOrReplaceDetails) {
					if (pattern == null)	System.out.println("MSWord_jsText: findOrReplace: ERROR: pattern IS NULL!");
					else 					System.out.println("MSWord_jsText: findOrReplace: pattern="+pattern);
					if (pattern2 == null)	System.out.println("MSWord_jsText: findOrReplace: ERROR: pattern2 IS NULL!");
					else 					System.out.println("MSWord_jsText: findOrReplace: pattern2="+pattern2);
				}
				
				try {
					if (debugSysPrnFindOrReplaceDetails)
						System.out.println("MSWord_jsText: findOrReplace: About to jacobFind.invoke(\"Execute\");");
					jacobSearchResultInt = jacobFind.invoke("Execute").toInt();
					//Please note: Wenn ich erste jacobSearchResultInt = jacobSearchresultVariant.toInt() verwende, ist nacher auch der string = ""-1", sonst "true"" +
					if (debugSysPrnFindOrReplaceDetails)
						System.out.println("MSWord_jsText: findOrReplace: jacobSearchResultInt="+jacobSearchResultInt);
					//System.out.println("Result: jacobInvokeResult.toString()="+jacobInvokeResult.toString());	//Returns true if match found, false if no match found
					//System.out.println("Result: jacobInvokeResult.toInt()="+jacobInvokeResult.toInt());		//Returns -1 if match found, 0 if no match found
					//System.out.println("Result: jacobInvokeResult.toError()="+jacobInvokeResult.toError());	//Throws java.lang.IllegalStateException: getError() only legal on Variants of type VariantError, not 3
				} catch (Exception ex) {
					ExHandler.handle(ex);
					//ToDo: Add precautions for pattern==null or pattern2==null...
					System.out.println("MSWord_jsText: findOrReplace (Haupttext):\nException caught.\n"+
					"Das verwendete Suchpattern funktioniert möglicherweise NICHT mit MS Word.\n"+
					"pattern \t(Von Elexis für OpenOffice):\t"+pattern+"\n"+
					"pattern2\t(Von msword_js für MS Word):\t"+pattern2+"\n"+		
					"Falls neue Suchpatterns hinzugefügt wurden, muss möglicherweise eine on-the-fly-Konversion in com.jsigle.msword_js.MSWord_jsText.java ergänzt werden.");
					//ToDo: Add precautions for pattern==null or pattern2==null...
					SWTHelper.showError(
							"MSWord_jsText: findOrReplace (Haupttext):", 
							"Fehler:",
							"Exception caught.\n"+
							"Das verwendete Suchpattern funktioniert möglicherweise NICHT mit MS Word.\n"+
							"pattern \t(Von Elexis für OpenOffice):    \t"+pattern+"\n"+	//spaces needed for tab alignment in proportional font
							"pattern2\t(Von msword_js für MS Word):\t"+pattern2+"\n"+		
							"Falls neue Suchpatterns hinzugefügt wurden, muss möglicherweise eine on-the-fly-Konversion in com.jsigle.msword_js.MSWord_jsText.java ergänzt werden.");
				}
				
			
				if (jacobSearchResultInt == -1) {		//Do ONLY attempt a replacement if there was a search hit. Otherwise, replacement text might be simply inserted at cursor position.
						//the following line should NOT produce an error - but it's flagged in Eclipse's editor with a red cross (x):
						//The local variable ... might not have been initialized - well, it's *defined* before and outside both try... blocks?!
						//If I actually initialize the variable up there to = null, the code error notice disappears.						
						
						numberOfHits += 1;
						if (debugSysPrnFindOrReplaceDetails)
							System.out.println("MSWord_jsText: findOrReplace: numberOfHits="+numberOfHits);
						
						
						//ToDo: This is an adhoc workaround to protect Tarmedrechnung_xx templates so that their header lines are NOT made appear.
						//ToDo:   We should rather find out how Word can completely do away with an empty header line again,
						//ToDo:   after that has been displayed by accessing ...Section.Header.Range (below in the SectionHeaders portion of findOrReplace).
						if (pattern.equals("\\[Titel\\]")) { ProbablyUsingTarmed_xxTemplateSoDoNOTAccessHeaderRangeToAvoidGenerationOfEmptyHeaderLines = true; }
						
						
						
						
						//System.out.println("MSWord_jsText: findOrReplace: About to jacobSelection.setProperty(\"Text\", \"Replaced\");");						
						//jacobSelection.setProperty("Text", "Replaced");		//Das sollte "Replaced" anstelle des Suchtexts einfügen.
											
						//Callback an Elexis den gefundenen [Platzhalter] übergeben und um replacement bitten
						
						String orig = jacobSelection.getProperty("Text").toString();
						if (debugSysPrnFindOrReplaceDetails) {
							if (orig  == null)	System.out.println("MSWord_jsText: findOrReplace: ERROR: orig IS NULL!");
							else 				System.out.println("MSWord_jsText: findOrReplace: orig="+orig);
						}

						Object replace = null;
						if (cb == null)		{
							System.out.println("MSWord_jsText: findOrReplace: ERROR: cb IS NULL! - static replacement ??Auswahl?? will be used - may be helpful for debugging.");
						}
						else 				{
							if (debugSysPrnFindOrReplaceDetails) { 
								System.out.println("MSWord_jsText: findOrReplace: cb="+cb.toString());
								System.out.println("MSWord_jsText: findOrReplace: About to Object replace = cb.replace(orig);...");
							}
							replace = cb.replace(orig);
						}

						if (debugSysPrnFindOrReplaceDetails) { 
							if (replace == null)	System.out.println("MSWord_jsText: findOrReplace: ERROR: replace IS NULL!");
							else 					System.out.println("MSWord_jsText: findOrReplace: replace="+replace.toString());
						}
						
						if (replace == null) {									//Falls nichts brauchbares zurückkommt: Fehlermeldung in den Text setzen.
							if (debugSysPrnFindOrReplaceDetails)
								System.out.println("MSWord_jsText: findOrReplace: About to jacobSelection.setProperty(\"Text\", \"??Auswahl??\");");
							jacobSelection.setProperty("Text", "??Auswahl??");	//ToDo: Das ist ein wenig dämlich: besser wäre es, NICHTS zu tun, dann könnte man den Platzhalter noch sehen.
						
						} else if (replace instanceof String) {
							if (debugSysPrnFindOrReplaceDetails) 
								System.out.println("MSWord_jsText: findOrReplace: About to replace \\r and \\n and their combinations by (traditionally) suitable linebreaks...");
							// String repl=((String)replace).replaceAll("\\r\\n[\\r\\n]*", "\n")
							String repl = ((String) replace).replaceAll("\\r", "\n");
							repl = repl.replaceAll("\\n\\n+", "\n");
							
							if (debugSysPrnFindOrReplaceDetails) { 
								if (repl == null)	System.out.println("MSWord_jsText: findOrReplace: ERROR: repl IS NULL!");
								else 				System.out.println("MSWord_jsText: findOrReplace: repl="+repl);							
								if (repl != null) {
									System.out.println("MSWord_jsText: findOrReplace: repl.length()="+repl.length());						
									/* No, we do NOT need the special handling, the [SQL in question] was simply not recognized because it succeeded
									 * the preceeding [SQL ...] immediately, and the MoveRight (without an immediate MoveLeft, back then) put the
									 * cursor = find starting point into that second [SQL placeholder] already, so it was not found.
									 * The preceding one that also returned an empty repl, was replaced away perfectly without any special handling.
									 * After having added MoveLeft after the MoveRight (see below) after the replacement action, it all works fine.   
									 if (repl.length() == 0) {
									    System.out.println("MSWord_jsText: findOrReplace: Special handling: repl.length()=0 so to ensure orig is replaced away:");
										System.out.println("MSWord_jsText: findOrReplace: About to jacobSelection.setProperty(\"Text\", chr(0));"); 
										jacobSelection.setProperty("Text", " ");  //  \"\" won't suffice, neither will "\0" - an alternative would be to put a " " and then a backspace.
									}
									else {								 
										System.out.println("MSWord_jsText: findOrReplace: About to jacobSelection.setProperty(\"Text\", repl);");
										jacobSelection.setProperty("Text", repl);
									}	
									*/															
								}
							}
							
							if (debugSysPrnFindOrReplaceDetails)
								System.out.println("MSWord_jsText: findOrReplace: About to jacobSelection.setProperty(\"Text\", repl);");
							jacobSelection.setProperty("Text", repl);
						} else if (replace instanceof String[][]) {
							String[][] contents = (String[][]) replace;
							//ToDo: Handler für Tabellen-Einfügung als String-Ersetzung noch hinzufügen.
							System.out.println("MSWord_jsText: findOrReplace: TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO");
							System.out.println("MSWord_jsText: findOrReplace: ToDo Tabellen-Einfügung als String-Ersetzung noch hinzufügen!");
							System.out.println("MSWord_jsText: findOrReplace: TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO");

							if (debugSysPrnFindOrReplaceDetails)
								System.out.println("MSWord_jsText: findOrReplace: About to jacobSelection.setProperty(\"Text\", \"MSWord_jsText.java ToDo: Tabellen-Einfügung als String-Ersetzung noch hinzufügen!\");");
							jacobSelection.setProperty("Text", "MSWord_jsText.java ToDo: Tabellen-Einfügung als String-Ersetzung noch hinzufügen!");

							/*
							try {
								ITextTable textTable =
									agIonDoc.getTextTableService().constructTextTable(contents.length,
										contents[0].length);
								agIonDoc.getTextService().getTextContentService().insertTextContent(r,
									textTable);
								r.setText("");
								
								ITextTablePropertyStore props = textTable.getPropertyStore();
								// long w=props.getWidth();
								// long percent=w/100;
								for (int row = 0; row < contents.length; row++) {
									String[] zeile = contents[row];
									for (int col = 0; col < zeile.length; col++) {
										textTable.getCell(col, row).getTextService().getText().setText(
											zeile[col]);
									}
								}
								textTable.spreadColumnsEvenly();
								
							} catch (Exception ex) {
								ExHandler.handle(ex);
								r.setText("Fehler beim Ersetzen");
							}
							*/
								
						} else {
							if (debugSysPrnFindOrReplaceDetails)
								System.out.println("MSWord_jsText: findOrReplace: About to jacobSelection.setProperty(\"Text\", \"Not a String\");");
							jacobSelection.setProperty("Text", "Not a String");
						}
						
						//GETESTET: DAS MoveRight; MoveLeft; IST WIRKLICH NÖTIG, UM IM HAUPTTEXT ZUVERLÄSSIG ALLE PLATZHALTER ZU ERSETZEN. NICHT NÖTIG IN SHAPES.

						//Moving right removes the highlighting and places the cursor to the right of the replaced text.
						//This is required, as otherwise, successive find/replace occurances may become confused.
						if (debugSysPrnFindOrReplaceDetails)
							System.out.println("MSWord_jsText: findOrReplace: About to jacobSelection.invoke(\"MoveRight\");");
						jacobSelection.invoke("MoveRight");
						
						//However, it's also necessary to go back to the left by one step afterwards,
						//or otherwise, a seamlessly following [placeholders][seamlesslyFollowingPlaceholder] will NOT be found.
						//The MoveRight - MoveLeft sequence has the effect that the selection = highlighting is removed from the inserted text.
						if (debugSysPrnFindOrReplaceDetails)
							System.out.println("MSWord_jsText: findOrReplace: About to jacobSelection.invoke(\"MoveLeft\");");
						jacobSelection.invoke("MoveLeft");
						
						if (debugSysPrnFindOrReplaceDetails)
							System.out.println("");
				}
			} while (jacobSearchResultInt == -1); 
		} catch (Exception ex) {
			ExHandler.handle(ex);
			//ToDo: Add precautions for pattern==null or pattern2==null...
			System.out.println("MSWord_jsText: findOrReplace:\nFehler beim Suchen und Ersetzen im Haupttext:\n"+"" +
					"Exception caught für:\n"+
					"pattern \t(Von Elexis für OpenOffice):\t"+pattern+"\n"+
					"pattern2\t(Von msword_js für MS Word):\t"+pattern2+"\n"+		
					"numberOfHits="+numberOfHits);
			//SWTHelper.showError("No doc in bill", "Fehler:","Es ist keine Rechnungsvorlage definiert");
			//ToDo: Add precautions for pattern==null or pattern2==null...
			SWTHelper.showError(
					"MSWord_jsText: findOrReplace:"+ 
					"Fehler:",
					"MSWord_jsText: findOrReplace:\nFehler beim Suchen und Ersetzen im Haupttext:\n"+
					"Exception caught für:\n"+
					"pattern \t(Von Elexis für OpenOffice):    \t"+pattern+"\n"+	//spaces needed for tab alignment in proportional font
					"pattern2\t(Von msword_js für MS Word):\t"+pattern2+"\n"+		
					"numberOfHits="+numberOfHits);
		}
					
	
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		//The following code performs search-and-replace in text within all Shapes (Textfelder),
		//e.g. the Adressfeld and Datumsfeld in my medical letter templates:
        
		if (debugSysPrnFindOrReplaceDetails) {
			System.out.println("MSWord_jsText: findOrReplace (Shapes): ");
			System.out.println("MSWord_jsText: findOrReplace (Shapes): About to list Shapes, and find/replace in each Shape.TextFrame.TextRange.Text:");
			System.out.println("MSWord_jsText: findOrReplace (Shapes): ");
		}

		//Identify and process each available Shape...
		
		try {
            Dispatch jacobShapes = Dispatch.get((Dispatch) jacobDocument, "Shapes").toDispatch();
            int shapesCount = Dispatch.get(jacobShapes , "Count").getInt();
            if (debugSysPrnFindOrReplaceDetails)
            	System.out.println("MSWord_jsText: findOrReplace (Shapes): shapesCount="+shapesCount);
	        
            for (int i = 0; i < shapesCount; i++) {
            	//Dispatch jacobShape = Dispatch.call(jacobShapes, "Item", new Variant(i + 1)).toDispatch();
            	//The above one-step call + conversion used to fail, a two step approach using a Variant type intermediate storage appears to work. 20160924js
            	//Same behaviour observed for another step further below.
            	if (debugSysPrnFindOrReplaceDetails)
            		System.out.println("MSWord_jsText: findOrReplace (Shapes): About to Variant jacobShapeVariant = Dispatch.call(jacobShapes, \"Item\", new Variant("+(i + 1)+"));");
                Variant jacobShapeVariant = Dispatch.call(jacobShapes, "Item", new Variant(i + 1));
                if (debugSysPrnFindOrReplaceDetails) {
                	if (jacobShapeVariant == null)	System.out.println("MSWord_jsText: findOrReplace (Shapes): WARNING: jacobShapeVariant IS NULL");
                	else 							System.out.println("MSWord_jsText: findOrReplace (Shapes): jacobShapeVariant="+jacobShapeVariant.toString());
                }
                if (debugSysPrnFindOrReplaceDetails) 
                	System.out.println("MSWord_jsText: findOrReplace (Shapes): About to Dispatch jacobShape = jacobShapeVariant.toDispatch();");
                Dispatch jacobShape = jacobShapeVariant.toDispatch();        
                
                if (debugSysPrnFindOrReplaceDetails) 
                	System.out.println("MSWord_jsText: findOrReplace (Shapes): About to String jacobShapeName = Dispatch.get(jacobShape, \"Name\").toString();");
                String jacobShapeName = Dispatch.get(jacobShape, "Name").toString();
                if (debugSysPrnFindOrReplaceDetails) {
                    if (jacobShapeName == null)	System.out.println("MSWord_jsText: findOrReplace (Shapes): WARNING: Shape["+(i+1)+"].Name IS NULL");
                    else 						System.out.println("MSWord_jsText: findOrReplace (Shapes): Shape["+(i+1)+"].Name="+jacobShapeName);
                }
                
                
                
                //ToDo: This is a Workaround. Should probably make new Tarmedrechnung_xx templates for Word anyway.
                //For Tarmedrechnung_xx templates: Set the ZORDER of existing shapes to msoBringInFrontOfText = 4 (or: msoBringToFront = 0).
                //In these templates, a title line is followed by numerous newlines that shall go BEHIND the boxes below the title,
                //and define its distance from the first row of the bill positions table. Shape ZORDER is interpreted as inline with the text,
                //when these templates are imported from their OpenOffice format (for whatever reason).
                //As we've already done work to recognize these templates, we can just as well change that problem on the fly.
                if ( ProbablyUsingTarmed_xxTemplateSoDoNOTAccessHeaderRangeToAvoidGenerationOfEmptyHeaderLines ) {
                	if (debugSysPrnFindOrReplaceDetails)
                        System.out.println("MSWord_jsText: findOrReplace (Shapes): INFO: Workaround for Tarmedrechnung_xx templates.");
                	//Dispatch.call(jacobShape, "ZOrder", 0);
                    //Dispatch.call(jacobShapeTextFrame, "ZOrder", 0);	//ZOrder wirkt wohl nur zwischen Shapes (?)
                    //System.out.println("MSWord_jsText: findOrReplace (Shapes): INFO: Workaround for Tarmedrechnung_xx templates. About to Dispatch.call(jacobShape, \"ZOrder\", 4);");
                    //Dispatch.call(jacobShape, "ZOrder", 4);
                    
                	if (debugSysPrnFindOrReplaceDetails)
                		System.out.println("MSWord_jsText: findOrReplace (Shapes): INFO: Workaround for Tarmedrechnung_xx templates. About to Dispatch jacobShapeWrapFormat = Dispatch.get(jacobShape, \"WrapFormat\").toDispatch();");
                    Dispatch jacobShapeWrapFormat = Dispatch.get(jacobShape, "WrapFormat").toDispatch();
                    if (debugSysPrnFindOrReplaceDetails)
                    	System.out.println("MSWord_jsText: findOrReplace (Shapes): INFO: Workaround for Tarmedrechnung_xx templates. About to Dispatch.put(jacobShapeWrapFormat, \"AllowOverlap\", new Variant (true));");
                    Dispatch.put(jacobShapeWrapFormat, "AllowOverlap", new Variant (true));  //Bezieht sich auf andere Shapes
                                     
                    int wdWrapSquare = 0;		//Das WrapFormat.Type bezieht sich darauf, wie Text das Shape um- oder über- oder unter-fliesst.
                    int wdWrapTight = 1;
                    int wdWrapThrough = 2;
                    int wdWrapNone = 3;
                    int wdWrapTopBottom = 4;
                    int wdWrapInline = 7;
                    Dispatch.put(jacobShapeWrapFormat, "Type", wdWrapNone); //wdWrapNone ist am ehesten, was ich für Text-Shapes am Beginn der Tarmedrechnung_Sx etc. brauche,
                    														//damit die Leerzeilen zwischen [Titel] und Leistungs-Tabelle nicht mehr UNTER, sondern HINTER
                    														//den Rechtecken ganz oben stehen - danach ist dann Word auch mit den dazu unter OpenOffice
                    														//erstellten Druckvorlagen ausreichend kompatibel, um diese verwenden zu können.
                    														//Andernfalls würde der Block mit den Leistungen sehr weit an den unteren Seitenrand
                    														//und hinter die ESR-Vordruck-Textframes geschoben, da die Leerzeilen, die normalerweise HINTER
                    														//den Rechtecken im oberen Bereich mit den Meta-Daten liegen sollen, stattdessen UNTERHALB
                    														//von diesen versetzt würden.
                    
                }
                
                
                
            	
                
                
                
                
                
                if (debugSysPrnFindOrReplaceDetails)
                	System.out.println("MSWord_jsText: findOrReplace (Shapes): About to Dispatch jacobShapeTextFrame = Dispatch.call(jacobShape, \"TextFrame\").toDispatch();");
                Dispatch jacobShapeTextFrame = Dispatch.call(jacobShape, "TextFrame").toDispatch();

                if (debugSysPrnFindOrReplaceDetails)
                	System.out.println("MSWord_jsText: findOrReplace (Shapes): About to Integer jacobShapeTextFrameHasText = Dispatch.get(jacobShapeTextFrame, \"HasText\").toInt();");
                Integer jacobShapeTextFrameHasText = Dispatch.get(jacobShapeTextFrame, "HasText").toInt();
                if (debugSysPrnFindOrReplaceDetails)
                	System.out.println("MSWord_jsText: findOrReplace (Shapes): Shape["+(i+1)+"].TextFrame.HasText="+jacobShapeTextFrameHasText);
                
                if (jacobShapeTextFrameHasText == -1) {
                	do {
        				jacobSearchResultInt = 0;	//Reset this to 0 in each loop iteration, so that even when jacobFind.invoke() should NOT put a valid result into this variable
        											//(or when it should not even get invoked!),
													//we do definitely NOT get an endless loop, NOR a misguided attempt to replace text. 	

                    	//Hier muss das ganze Procedere, schon beginnend mit Dispatch jacobShapeTextFrameTextRange = ... (!!!) in den do...while block,
	            		//damit mehrere Platzhalter innerhalb eines Shapes ersetzt werden. Hab's probiert, nichts anderes hilft.
	            		//Insbesondere auch nicht das Verschieben des Cursors nach Durchführen einer Ersetzung - und das,
	            		//nachdem ich sehr lange gebraucht habe, um herauszufinden, wie das wirklich ausführbar codiert werden kann. 201609271147js
	                	Dispatch jacobShapeTextFrameTextRange = Dispatch.call(jacobShapeTextFrame, "TextRange").toDispatch();
	                	String jacobShapeTextFrameTextRangeText = Dispatch.get(jacobShapeTextFrameTextRange, "Text").toString();
	                    
	                    if (jacobShapeTextFrameTextRangeText == null) {
	                    	if (debugSysPrnFindOrReplaceDetails)
	                    		System.out.println("MSWord_jsText: findOrReplace (Shapes): WARNING: Shape["+(i+1)+"].TextFrame.TextRange.Text IS NULL");
	                    	}
	                    else {
	                    	if (debugSysPrnFindOrReplaceDetails)
	                    		System.out.println("MSWord_jsText: findOrReplace (Shapes): Shape["+(i+1)+"].TextFrame.TextRange.Text="+jacobShapeTextFrameTextRangeText);
	                		//JETZT HABEN WIR ENDLICH DEN TEXT DES Shapes.Shape.TextFrame.TextRange.Text...
	
	                    	//Hier KÖNNTE ich innerhalb von Java einfach eine regex search, wiederholend, innerhalb des Shape-Textes durchführen.
	                    	//UND also etwas wie das hier schreiben:
	                    	//
	                    	//do-one-complete-global-regexp-replacement-including-multiple-search-hits-per ~s/../../g or similar call-in-java
	                		//Dispatch.put(jacobShapeTextFrameTextRange, "Text", jacobShapeTextFrameTextRangeTextALLTOGETHERREGEXPREPLACED).toString();
	                    	//
	                    	//ABER:
	                    	//(0) + Vermutlich wäre das schneller.
	                    	//(1) - Würde das allfällige Formatierungen im Ziel vermutlich homogenisieren, beim gemeinsamen zurückschreiben es Ergebnisses.
	                    	//(2) - Könnte es auch bei Umbrüchen etc. Probleme geben. Also lass ich das lieber durch Word erledigen, wie oben im Haupttext auch.
	                    	
	
	                		//For each Shape, begin searching at the top of its own TextRange
	                		
	                    	//!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
	                    	//Wegen der anderen Ausgangssituation ist hier alles als Dispatch.call... formuliert,
	                    	//statt wie weiter oben als ActiveXComponent...
	                    	//!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
	                    	                    
	                    	//The original code from the main portion, variables were not allocated up there:
	                		//ActiveXComponent jacobSelection = jacobShapeTextFrameTextRange.getPropertyAsComponent("Selection");
	                		//ActiveXComponent jacobFind = jacobSelection.getPropertyAsComponent("Find");
	                    	
	                    	
	                    	//jacobShapeSelection = Dispatch.call(jacobShape, "Select").toDispatch();
	                    	//jacobShapeSelection = Dispatch.call(jacobShapeTextFrame, "Select").toDispatch();
	                    	//jacobShapeSelection = Dispatch.call(jacobShapeTextFrameTextRange, "Select").toDispatch();
	                    		
	                    	//jacobShapeSelection = Dispatch.call(jacobShapeTextFrameTextRange, "Select").toDispatch();
	                    	
	                    	//THIS WORKS, and causes the *shape!* to become selected, but *not* the text inside.
	                    	/*                    	
	                    	System.out.println("MSWord_jsText: findOrReplace: About to Variant jacobShapeSelectVariant = Dispatch.call(jacobShape, \"Select\");");                        
	                    	Variant jacobShapeSelectVariant = Dispatch.call(jacobShape, "Select");
	                        if (jacobShapeSelectVariant == null)	System.out.println("MSWord_jsText: findOrReplace: WARNING: jacobShapeSelectVariant IS NULL");
	                        else 									System.out.println("MSWord_jsText: findOrReplace: jacobShapeSelectVariant="+jacobShapeSelectVariant.toString());
	                        */
	
	                    	
	                        //THIS FAILS, NO MATTER WHETHER THE PRECEDING ONE WAS COMMENTED OUT OR NOT
	                        /*
	                        com.jacob.com.ComFailException: Can't map name to dispid: Select
	                    	at com.jacob.com.Dispatch.invokev(Native Method)
	                    	at com.jacob.com.Dispatch.invokev(Dispatch.java:625)
	                    	at com.jacob.com.Dispatch.callN(Dispatch.java:453)
	                    	at com.jacob.com.Dispatch.call(Dispatch.java:529)
	                    	at com.jsigle.msword_js.MSWord_jsText.findOrReplace(MSWord_jsText.java:2372)
	                    	at ch.elexis.text.TextContainer.createFromTemplate(TextContainer.java:276)
	                    	at ch.elexis.views.TextView.createDocument(TextView.java:406)
	                    	
	                    	System.out.println("MSWord_jsText: findOrReplace: About to Variant jacobShapeTextFrameSelectVariant = Dispatch.call(jacobShapeTextFrame, \"Select\");");
	                        Variant jacobShapeTextFrameSelectVariant = Dispatch.call(jacobShapeTextFrame, "Select");
	                        if (jacobShapeTextFrameSelectVariant == null)	System.out.println("MSWord_jsText: findOrReplace: WARNING: jacobShapeTextFrameSelectVariant IS NULL");
	                        else 									System.out.println("MSWord_jsText: findOrReplace: jacobShapeTextFrameSelectVariant="+jacobShapeTextFrameSelectVariant.toString());
	                        */
	
	                    	//THIS WORKS, and causes the text in the shape to become selected            
	                    	if (debugSysPrnFindOrReplaceDetails)
	                    		System.out.println("MSWord_jsText: findOrReplace (Shapes): About to Variant jacobShapeTextFrameTextRangeSelectVariant = Dispatch.call(jacobShapeTextFrameTextRange, \"Select\");");
	                        Variant jacobShapeTextFrameTextRangeSelectVariant = Dispatch.call(jacobShapeTextFrameTextRange, "Select");
	                        if (debugSysPrnFindOrReplaceDetails) {
	                        	if (jacobShapeTextFrameTextRangeSelectVariant == null)	System.out.println("MSWord_jsText: findOrReplace (Shapes): WARNING: jacobShapeTextFrameTextRangeSelectVariant IS NULL");
	                        	else 	System.out.println("MSWord_jsText: findOrReplace (Shapes): jacobShapeTextFrameTextRangeSelectVariant="+jacobShapeTextFrameTextRangeSelectVariant.toString());
	                        }
	                        
	                        /*
	                        //THIS FAILS
	                        /com.jacob.com.ComFailException: VariantChangeType failed
	                        System.out.println("MSWord_jsText: findOrReplace: About to Dispatch jacobShapeTextFrameTextRangeSelect = jacobShapeTextFrameTextRangeSelectVariant.toDispatch();");
	                        Dispatch jacobShapeTextFrameTextRangeSelect = jacobShapeTextFrameTextRangeSelectVariant.toDispatch();
	                        */
	                        
	                        /*
	                        System.out.println("MSWord_jsText: findOrReplace: About to Dispatch jacobShapeSelection = jacobShapeSelectionVariant.toDispatch();");
	                        Dispatch jacobShapeSelection = jacobShapeSelectionVariant.toDispatch();
	                        */
	                        
	                        /* THIS FAILS
	                        com.jacob.com.ComFailException: VariantChangeType failed
	                        System.out.println("MSWord_jsText: findOrReplace: About to Dispatch jacobShapeTextFrameTextRangeSelect = jacobShapeTextFrameTextRangeSelectVariant.toDispatch();");
	                        Dispatch jacobShapeTextFrameTextRangeSelect = jacobShapeTextFrameTextRangeSelectVariant.toDispatch();
	                        */
	                       
	                        /*THIS FAILS
	                        com.jacob.com.ComFailException: VariantChangeType failed
	                        System.out.println("MSWord_jsText: findOrReplace: About to Variant jacobShapeFindVariant = Dispatch.call(jacobShapeSelection, \"Find\");");                		
	                    	Variant jacobShapeFindVariant = Dispatch.call(jacobShapeTextFrameTextRangeSelectVariant.toDispatch(), "Find");
	                        if (jacobShapeFindVariant == null)	System.out.println("MSWord_jsText: findOrReplace: WARNING: jacobShapeFindVariant IS NULL");
	                        else 								System.out.println("MSWord_jsText: findOrReplace: jacobShapeFindVariant="+jacobShapeFindVariant.toString());
	                         */                   	
	
	                        
	    /* THIS WORKS HERE, BUT FAILS LATER ON.
	                    	//THIS WORKS, and causes the text in the shape to remain selected
	                        //But later on, when I want to  use jacobShapeFind, that throws this error
	                        //com.jacob.com.ComFailException: A COM exception has been encountered:
	                        //At Invoke of: Text
	                        //Description: 80020011 / Does not support a collection.
	                        */
	                        if (debugSysPrnFindOrReplaceDetails)
	                        	System.out.println("MSWord_jsText: findOrReplace (Shapes): About to Variant jacobShapeTextFrameTextRangeFindVariant = Dispatch.call(jacobShapeTextFrameTextRange, \"Find\");");
	                        Variant jacobShapeTextFrameTextRangeFindVariant = Dispatch.get(jacobShapeTextFrameTextRange, "Find");
	                        if (debugSysPrnFindOrReplaceDetails) {
	                        	if (jacobShapeTextFrameTextRangeFindVariant == null)	System.out.println("MSWord_jsText: findOrReplace (Shapes): WARNING: jacobShapeTextFrameTextRangeFindVariant IS NULL");
	                        	else 								System.out.println("MSWord_jsText: findOrReplace (Shapes): jacobShapeTextFrameTextRangeFindVariant="+jacobShapeTextFrameTextRangeFindVariant.toString());
	                        }
	
	                        if (debugSysPrnFindOrReplaceDetails)
	                        	System.out.println("MSWord_jsText: findOrReplace (Shapes): About to Dispatch jacobShapeTextFrameTextRangeFind = jacobShapeTextFrameTextRangeFindVariant.toDispatch();");
	                        Dispatch jacobShapeTextFrameTextRangeFind = jacobShapeTextFrameTextRangeFindVariant.toDispatch();
	                        if (debugSysPrnFindOrReplaceDetails) {
	                        	if (jacobShapeTextFrameTextRangeFind  == null)	System.out.println("MSWord_jsText: findOrReplace (Shapes): WARNING: jacobShapeTextFrameTextRangeFind IS NULL");
	                        	else 							System.out.println("MSWord_jsText: findOrReplace (Shapes): jacobShapeTextFrameTextRangeFind="+jacobShapeTextFrameTextRangeFind.toString());
	                        }
	
	                        
	
	                        
	                        /*
	                        //THIS FAILS
	                        com.jacob.com.ComFailException: Can't map name to dispid: Find
	                    	System.out.println("MSWord_jsText: findOrReplace: About to Variant jacobShapeTextFrameFindVariant = Dispatch.call(jacobShapeTextFrame, \"Find\");");
	                        Variant jacobShapeTextFrameFindVariant = Dispatch.call(jacobShapeTextFrame, "Find");
	                        if (jacobShapeTextFrameFindVariant == null)	System.out.println("MSWord_jsText: findOrReplace: WARNING: jacobShapeTextFrameFindVariant IS NULL");
	                        else 								System.out.println("MSWord_jsText: findOrReplace: jacobShapeTextFrameFindVariant="+jacobShapeTextFrameFindVariant.toString());                        
	                        System.out.println("MSWord_jsText: findOrReplace: About to Dispatch jacobShapeTextFrameFind = jacobShapeTextFrameFindVariant.toDispatch();");
	                        Dispatch jacobShapeTextFrameFind = jacobShapeTextFrameFindVariant.toDispatch();
	                        if (jacobShapeTextFrameFind  == null)	System.out.println("MSWord_jsText: findOrReplace: WARNING: jacobShapeTextFrameFind IS NULL");
	                        else 							System.out.println("MSWord_jsText: findOrReplace: jacobShapeTextFrameFind="+jacobShapeTextFrameFind.toString());
							*/
	                                                
	                        /*
	                        //THIS FAILS
	                        com.jacob.com.ComFailException: VariantChangeType failed
	                        System.out.println("MSWord_jsText: findOrReplace: About to Dispatch jacobShapeTextFrameTextRangeTextSelect = jacobShapeTextFrameTextRangeTextSelectVariant.toDispatch();");                    
	                        Dispatch jacobShapeTextFrameTextRangeTextSelect = jacobShapeTextFrameTextRangeTextSelectVariant.toDispatch();
	                        */
	                        
	                        /*
	                    	System.out.println("MSWord_jsText: findOrReplace: About to Variant jacobShapeFindVariant = Dispatch.call(jacobShapeTextFrameTextRange, \"Find\");");
	                        Variant jacobShapeFindVariant = Dispatch.call(jacobShapeTextFrameTextRange, "Find");
	                        if (jacobShapeFindVariant == null)	System.out.println("MSWord_jsText: findOrReplace: WARNING: jacobShapeFindVariant IS NULL");
	                        else 								System.out.println("MSWord_jsText: findOrReplace: jacobShapeFindVariant="+jacobShapeFindVariant.toString());                        
	
	                        */
	                        
	                        
	                        /*
	                    	//jacobShapeFind = Dispatch.call(jacobShapeSelection, "Find").toDispatch();
	                    	System.out.println("MSWord_jsText: findOrReplace: About to Variant jacobShapeFindVariant = Dispatch.call(jacobShapeSelection, \"Find\");");                		
	                    	Variant jacobShapeFindVariant = Dispatch.call(jacobShapeSelection, "Find");
	                        if (jacobShapeFindVariant == null)	System.out.println("MSWord_jsText: findOrReplace: WARNING: jacobShapeFindVariant IS NULL");
	                        else 								System.out.println("MSWord_jsText: findOrReplace: jacobShapeFindVariant="+jacobShapeFindVariant.toString());
	                    	
	                        System.out.println("MSWord_jsText: findOrReplace: About to Dispatch jacobShapeFind = jacobShapeFindVariant.toDispatch();");
	                        Dispatch jacobShapeFind = jacobShapeFindVariant.toDispatch();
	                        */
	                    	
	                		//jacobSelection = jacobShapeTextFrameTextRange.getPropertyAsComponent("Selection");
	                		//jacobFind = jacobSelection.getPropertyAsComponent("Find");
	
	                    	//Put the cursor (back) to the document top. This necessary, or all searches but the first one will most probably NOT find anything replacable: 
	                		//N.B.: We want to support multiple replacements within another; i.e. for an [SQL:select...] Query that has [Patient.Name] (or similar) as an argument.
	                			                	
	                        
	                        /*
	                        //System.out.println("MSWord_jsText: findOrReplace: About to Dispatch.call(jacobShapeSelection, \"HomeKey\", \"wdStory\", \"wdMove\");");
		                	//Dispatch.call(jacobSelection, "HomeKey", new Variant(6));
		                	//Dispatch.call(jacobShapeSelection, "HomeKey", new Variant(6));	              
		                	 */
	                        	                       
	                        System.out.println("MSWord_jsText: findOrReplace (Shapes): ");
	                        System.out.println("MSWord_jsText: findOrReplace (Shapes): TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO");
	                        System.out.println("MSWord_jsText: findOrReplace (Shapes): ToDo: Find a method to send HomeKey or MoveLeft or MoveRight, because we might need this for multiple replacements in one Text field.");
	                        System.out.println("MSWord_jsText: findOrReplace (Shapes): ToDo: This is definitely needed e.g. for Tarmed Bills and for Label printing.");
	                        System.out.println("MSWord_jsText: findOrReplace (Shapes): ToDo: Actually, it's NOT needed HERE (find/replace perfectly works without that for the first [Placeholder] in this Shape, maybe even for more),");
	                        System.out.println("but probably several lines further below, it might become necessary:");
	                        System.out.println("I've experienced difficulties (i.e. mixup of the replacements?) in the main text block processing multiple [PLaceholders], when I did not have MoveRight, MoveLeft after each replacement,");
	                        System.out.println("even though that should not have happened. So we need to test whether it's actually required by placeing multiple [Placeholders], with or without characters in between, with different placeholder mechanisms, into a Shape, some time.");
	                        System.out.println("Currently, the whole text within the currently processed shape remains selected before and after the successful replacement, and at least for one placeholder therein (as in my actually used medical letter template), that works,");
	                        System.out.println("for two different shapes/texts, when that placeholder is all of the text, or when there's additional non-placeholder text in the same shape before the placeholder."); 
	                        System.out.println("MSWord_jsText: findOrReplace (Shapes): TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO");
	                        System.out.println("MSWord_jsText: findOrReplace (Shapes): ");
	                                                
	                        /* THIS FAILS                     
	                        com.jacob.com.ComFailException: Can't map name to dispid: HomeKey
	                        System.out.println("MSWord_jsText: findOrReplace: About to Dispatch.call(jacobShapeTextFrameTextRange, \"HomeKey\", new Variant(6));");
	                        Dispatch.call(jacobShapeTextFrameTextRange, "HomeKey", new Variant(6));
	                        */
	
	                        /* THIS FAILS
	                        com.jacob.com.ComFailException: Can't map name to dispid: HomeKey
	                        System.out.println("MSWord_jsText: findOrReplace: About to Dispatch.call(jacobShapeTextFrame, \"HomeKey\", new Variant(6));");
	                        Dispatch.call(jacobShapeTextFrame, "HomeKey", new Variant(6));
	                        */
		                	
	                        /*
	                        System.out.println("MSWord_jsText: findOrReplace: About to Dispatch.call(jacobShapeSelection, \"HomeKey\", \"wdStory\", \"wdMove\");");
		                	//Dispatch.call(jacobSelection, "HomeKey", new Variant(6));
		                	Dispatch.call(jacobShapeSelection, "HomeKey", new Variant(6));
		                	*/
	                		
	/* COMMENTED OUT FOR DEBUGGING 2                        
	                        System.out.println("MSWord_jsText: findOrReplace: About to Dispatch.call(jacobShapeSelection, \"MoveLeft\");");
		                	Dispatch.call(jacobShapeSelection, "MoveLeft");
	 COMMENTED OUT FOR DEBUGGING 2 */                        
	                       
	                        
	                        if (debugSysPrnFindOrReplaceDetails)
	                        	System.out.println("MSWord_jsText: findOrReplace (Shapes): About to try Dispatch.call(jacobShapeTextFrameTextRangeFind, \"Text\", pattern2);... etc.");
	                		try {
	                			//Please note: We must ensure, that NO prior path disrupts placeholders that would be valid and recognized in a subsequent pass!
	                			//jacobFind.setProperty("Text", pattern2);
	                			//jacobFind.setProperty("Forward", "True");
	                			//jacobFind.setProperty("Format", "False");
	                			//jacobFind.setProperty("MatchCase", "False");
	                			//jacobFind.setProperty("MatchWholeWord", "False");
	                			//jacobFind.setProperty("MatchByte", "False");
	                			//jacobFind.setProperty("MatchAllWordForms", "False");
	                			//jacobFind.setProperty("MatchSoundsLike", "False");
	                			//jacobFind.setProperty("MatchWildcards", "True");
	                			Dispatch.put(jacobShapeTextFrameTextRangeFind, "Text", pattern2);
	                			Dispatch.put(jacobShapeTextFrameTextRangeFind, "Forward", "True");
	                			Dispatch.put(jacobShapeTextFrameTextRangeFind, "Format", "False");
	                			Dispatch.put(jacobShapeTextFrameTextRangeFind, "MatchCase", "False");
	                			Dispatch.put(jacobShapeTextFrameTextRangeFind, "MatchWholeWord", "False");
	                			Dispatch.put(jacobShapeTextFrameTextRangeFind, "MatchByte", "False");
	                			Dispatch.put(jacobShapeTextFrameTextRangeFind, "MatchAllWordForms", "False");
	                			Dispatch.put(jacobShapeTextFrameTextRangeFind, "MatchSoundsLike", "False");
	                			Dispatch.put(jacobShapeTextFrameTextRangeFind, "MatchWildcards", "True");	
	                		} catch (Exception ex) {
	                			ExHandler.handle(ex);
	                			System.out.println("MSWord_jsText: findOrReplace (Shapes): Fehler beim Ersetzen: Dispatch.put(jacobShapeTextFrameTextRangeFind...);");
	                			
	                			//SWTHelper.showError("No doc in bill", "Fehler:","Es ist keine Rechnungsvorlage definiert");
	                			SWTHelper.showError("findOrReplace (Shapes): Fehler beim Ersetzen: ", "Fehler:","Fehler beim Ersetzen: Dispatch.put(jacobShapeTextFrameTextRangeFind...);");
	                		}
	 
	                    	
	                		//The following block performs search-and-replace for each Shapes text portion of the document.
	                		//An almost identical block is further above for the main text (without updating all the comments),
	                		//and will probably be added further below, for tables. 
	                		
	                		if (debugSysPrnFindOrReplaceDetails)
	                			System.out.println("MSWord_jsText: findOrReplace (Shapes): About to try ... the actual search and replace block...");
	                		try {	 	
	                			if (debugSysPrnFindOrReplaceDetails) {
		            				if (pattern == null)	System.out.println("MSWord_jsText: findOrReplace (Shapes): ERROR: pattern IS NULL!");
		            				else 					System.out.println("MSWord_jsText: findOrReplace (Shapes): pattern="+pattern);		            				
		            				if (pattern2 == null)	System.out.println("MSWord_jsText: findOrReplace (Shapes): ERROR: pattern2 IS NULL!");
		            				else 					System.out.println("MSWord_jsText: findOrReplace (Shapes): pattern2="+pattern2);
	                			}
	            				
	            				try {
	            					if (debugSysPrnFindOrReplaceDetails)
	            						System.out.println("MSWord_jsText: findOrReplace (Shapes): About to jacobSearchResultInt = Dispatch.call(jacobShapeTextFrameTextRangeFind,\"Execute\").toInt();...");
	            					//jacobSearchResultInt = jacobFind.invoke("Execute").toInt();
	            					jacobSearchResultInt = Dispatch.call(jacobShapeTextFrameTextRangeFind,"Execute").toInt();
	            					
	            					//Please note: Wenn ich erste jacobSearchResultInt = jacobSearchresultVariant.toInt() verwende, ist nacher auch der string = ""-1", sonst "true"" +
	            					if (debugSysPrnFindOrReplaceDetails)
	            						System.out.println("MSWord_jsText: findOrReplace: jacobSearchResultInt="+jacobSearchResultInt);
	            					//System.out.println("Result: jacobInvokeResult.toString()="+jacobInvokeResult.toString());	//Returns true if match found, false if no match found
	            					//System.out.println("Result: jacobInvokeResult.toInt()="+jacobInvokeResult.toInt());		//Returns -1 if match found, 0 if no match found
	            					//System.out.println("Result: jacobInvokeResult.toError()="+jacobInvokeResult.toError());	//Throws java.lang.IllegalStateException: getError() only legal on Variants of type VariantError, not 3
	            				} catch (Exception ex) {
	            					ExHandler.handle(ex);
	            					//ToDo: Add precautions for pattern==null or pattern2==null...
	            					System.out.println("MSWord_jsText: findOrReplace (Shapes):\nException caught.\n"+
	            					"Das verwendete Suchpattern funktioniert möglicherweise NICHT mit MS Word.\n"+
	            					"pattern \t(Von Elexis für OpenOffice):\t"+pattern+"\n"+
	            					"pattern2\t(Von msword_js für MS Word):\t"+pattern2+"\n"+		
	            					"Falls neue Suchpatterns hinzugefügt wurden, muss möglicherweise eine on-the-fly-Konversion in com.jsigle.msword_js.MSWord_jsText.java ergänzt werden.");
	            					//ToDo: Add precautions for pattern==null or pattern2==null...
	            					SWTHelper.showError(
	            							"MSWord_jsText: findOrReplace (Shapes):", 
	            							"Fehler:",
	            							"Exception caught.\n"+
	            							"Das verwendete Suchpattern funktioniert möglicherweise NICHT mit MS Word.\n"+
	            							"pattern \t(Von Elexis für OpenOffice):    \t"+pattern+"\n"+	//spaces needed for tab alignment in proportional font
	            							"pattern2\t(Von msword_js für MS Word):\t"+pattern2+"\n"+		
	            							"Falls neue Suchpatterns hinzugefügt wurden, muss möglicherweise eine on-the-fly-Konversion in com.jsigle.msword_js.MSWord_jsText.java ergänzt werden.");
	            				}
	            				
	            				if (jacobSearchResultInt == -1) {		//Do ONLY attempt a replacement if there was a search hit. Otherwise, replacement text might be simply inserted at cursor position.
            						//the following line should NOT produce an error - but it's flagged in Eclipse's editor with a red cross (x):
            						//The local variable ... might not have been initialized - well, it's *defined* before and outside both try... blocks?!
            						//If I actually initialize the variable up there to = null, the code error notice disappears.
            						
            						numberOfHits += 1;
            						if (debugSysPrnFindOrReplaceDetails)
            							System.out.println("MSWord_jsText: findOrReplace (Shapes): numberOfHits="+numberOfHits);
            						
            						//System.out.println("MSWord_jsText: findOrReplace: About to jacobSelection.setProperty(\"Text\", \"Replaced\");");						
            						//jacobSelection.setProperty("Text", "Replaced");		//Das sollte "Replaced" anstelle des Suchtexts einfügen.
            						
            						
            						
            						/*
            						//THIS FAILS
            						/com.jacob.com.ComFailException: VariantChangeType failed
            						Dispatch.call(jacobShapeTextFrameTextRangeSelectVariant.toDispatch(), "Text", "Replaced");
            						*/
            						
            			/* SIMPLE GET, PRINTLN, REPLACE WITH CONSTANT STRING FOR TESTING ONLY - THIS WORKS :-) :-) :-)
            						//and very fine: especially, in Bern, [Brief.Datum] (or similar), only the [Brief.Datum] portion is replaced.
            						//This means, that the "Find" stuff actually works and controls the range that is influenced by the following commands. PUH.
            						System.out.println(Dispatch.get(jacobShapeTextFrameTextRange, "Text").toString());
            						Dispatch.put(jacobShapeTextFrameTextRange, "Text", new Variant("Replaced"));
            			*/
            						                						
            						//Callback an Elexis den gefundenen [Platzhalter] übergeben und um replacement bitten
            						
            						//String orig = jacobSelection.getProperty("Text").toString();
            						String orig = Dispatch.get(jacobShapeTextFrameTextRange, "Text").toString();
            						if (debugSysPrnFindOrReplaceDetails) {
            							if (orig  == null)	System.out.println("MSWord_jsText: findOrReplace (Shapes): ERROR: orig IS NULL!");
            							else 				System.out.println("MSWord_jsText: findOrReplace (Shapes): orig="+orig);
            							if (cb == null)		System.out.println("MSWord_jsText: findOrReplace (Shapes): ERROR: cb IS NULL!");
            							else 				System.out.println("MSWord_jsText: findOrReplace (Shapes): cb="+cb.toString());
            						}

            						if (debugSysPrnFindOrReplaceDetails)
            							System.out.println("MSWord_jsText: findOrReplace (Shapes): About to Object replace = cb.replace(orig);...");
            						Object replace = cb.replace(orig);

            						if (debugSysPrnFindOrReplaceDetails) {
            							if (replace == null)	System.out.println("MSWord_jsText: findOrReplace (Shapes): ERROR: replace IS NULL!");
            							else 					System.out.println("MSWord_jsText: findOrReplace (Shapes): replace="+replace.toString());
            						}
            						
            						if (replace == null) {									//Falls nichts brauchbares zurückkommt: Fehlermeldung in den Text setzen.
            							if (debugSysPrnFindOrReplaceDetails)
            								System.out.println("MSWord_jsText: findOrReplace (Shapes): About to Dispatch.put(jacobShapeTextFrameTextRange, \"Text\", \"??Auswahl??\");");
            							
            							//jacobSelection.setProperty("Text", "??Auswahl??");	//ToDo: Das ist ein wenig dämlich: besser wäre es, NICHTS zu tun, dann könnte man den Platzhalter noch sehen.
            							Dispatch.put(jacobShapeTextFrameTextRange, "Text", "??Auswahl??");
            						} else if (replace instanceof String) {
            							if (debugSysPrnFindOrReplaceDetails)
            								System.out.println("MSWord_jsText: findOrReplace (Shapes): About to replace \\r and \\n and their combinations by (traditionally) suitable linebreaks...");
            							// String repl=((String)replace).replaceAll("\\r\\n[\\r\\n]*", "\n")
            							String repl = ((String) replace).replaceAll("\\r", "\n");
            							repl = repl.replaceAll("\\n\\n+", "\n");
            							
            							if (debugSysPrnFindOrReplaceDetails) {
            								if (repl == null)	System.out.println("MSWord_jsText: findOrReplace (Shapes): ERROR: repl IS NULL!");
            								else 				System.out.println("MSWord_jsText: findOrReplace (Shapes): repl="+repl);
          							
	            							if (repl != null) {
	            								System.out.println("MSWord_jsText: findOrReplace (Shapes): repl.length()="+repl.length());														
	            							}
            							}
            							
            							if (debugSysPrnFindOrReplaceDetails)
            								System.out.println("MSWord_jsText: findOrReplace (Shapes): About to Dispatch.put(jacobShapeTextFrameTextRange, \"Text\", repl);");
            							//jacobSelection.setProperty("Text", repl);
            							Dispatch.put(jacobShapeTextFrameTextRange, "Text", repl);
            						
            						} else if (replace instanceof String[][]) {
            							String[][] contents = (String[][]) replace;
            							//ToDo: Handler für Tabellen-Einfügung als String-Ersetzung noch hinzufügen.
            							System.out.println("MSWord_jsText: findOrReplace (Shapes): TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO");
            							System.out.println("MSWord_jsText: findOrReplace (Shapes): ToDo Tabellen-Einfügung als String-Ersetzung noch hinzufügen! (only if Tables in Shapes shall be supported...)");
            							System.out.println("MSWord_jsText: findOrReplace (Shapes): TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO");

            							if (debugSysPrnFindOrReplaceDetails)
            								System.out.println("MSWord_jsText: findOrReplace (Shapes): About to Dispatch.put(jacobShapeTextFrameTextRange, \"Text\", \"MSWord_jsText.java ToDo: Tabellen-Einfügung als String-Ersetzung noch hinzufügen!\");");
            							//jacobSelection.setProperty("Text", "MSWord_jsText.java ToDo: Tabellen-Einfügung als String-Ersetzung noch hinzufügen!");
            							Dispatch.put(jacobShapeTextFrameTextRange, "Text", "MSWord_jsText.java ToDo:  (Shapes) Tabellen-Einfügung als String-Ersetzung noch hinzufügen!");

            							/*
            							try {
            								ITextTable textTable =
            									agIonDoc.getTextTableService().constructTextTable(contents.length,
            										contents[0].length);
            								agIonDoc.getTextService().getTextContentService().insertTextContent(r,
            									textTable);
            								r.setText("");
            								
            								ITextTablePropertyStore props = textTable.getPropertyStore();
            								// long w=props.getWidth();
            								// long percent=w/100;
            								for (int row = 0; row < contents.length; row++) {
            									String[] zeile = contents[row];
            									for (int col = 0; col < zeile.length; col++) {
            										textTable.getCell(col, row).getTextService().getText().setText(
            											zeile[col]);
            									}
            								}
            								textTable.spreadColumnsEvenly();
            								
            							} catch (Exception ex) {
            								ExHandler.handle(ex);
            								r.setText("Fehler beim Ersetzen");
            							}               
            							*/ 							
            						} else {
            							if (debugSysPrnFindOrReplaceDetails)
            								System.out.println("MSWord_jsText: findOrReplace (Shapes): About to Dispatch.put(jacobShapeTextFrameTextRange, \"Text\", \"Not a String\");");
            							//jacobSelection.setProperty("Text", "Not a String");
            							Dispatch.put(jacobShapeTextFrameTextRange, "Text", "Not a String");
            						}
            						
                                    System.out.println("MSWord_jsText: findOrReplace (Shapes): ");
                                    System.out.println("MSWord_jsText: findOrReplace (Shapes): TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO");
                                    System.out.println("MSWord_jsText: findOrReplace (Shapes): ToDo: Find a method to send HomeKey or MoveLeft or MoveRight, because we might need this for multiple replacements in one Text field.");                                        System.out.println("MSWord_jsText: findOrReplace (Shapes): ToDo: This is definitely needed e.g. for Tarmed Bills and for Label printing.");
                                    System.out.println("I've experienced difficulties (i.e. mixup of the replacements?) in the main text block processing multiple [PLaceholders], when I did not have MoveRight, MoveLeft after each replacement,");
                                    System.out.println("even though that should not have happened. So we need to test whether it's actually required by placeing multiple [Placeholders], with or without characters in between, with different placeholder mechanisms, into a Shape, some time.");
                                    System.out.println("Currently, the whole text within the currently processed shape remains selected before and after the successful replacement, and at least for one placeholder therein (as in my actually used medical letter template), that works,");
                                    System.out.println("for two different shapes/texts, when that placeholder is all of the text, or when there's additional non-placeholder text in the same shape before the placeholder."); 
                                    System.out.println("I HAVE MADE VARIOUS ATTEMPTS TO GET MoveRight; MoveLeft; TO WORK HERE, BUT DIDN'T SUCCEED - SO I'LL POSTPONE THIS TO LATER, COMPLETING OTHER DEFINITELY REQUIRED PORTIONS OF CODE FIRST."); 
                                    System.out.println("MSWord_jsText: findOrReplace (Shapes): TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO");
                                    System.out.println("MSWord_jsText: findOrReplace (Shapes): ");

                                    /*
                                    //THIS FAILS
            						System.out.println("MSWord_jsText: findOrReplace (Shapes): About to Dispatch.call((Dispatch) jacobShapeTextFrameTextRange, \"MoveRight\");");
                                    Dispatch.call((Dispatch) jacobShapeTextFrameTextRange, "MoveRight");
            						System.out.println("MSWord_jsText: findOrReplace (Shapes): About to Dispatch.call((Dispatch) jacobShapeTextFrameTextRange, \"MoveLeft\");");
                                    Dispatch.call((Dispatch) jacobShapeTextFrameTextRange, "MoveLeft");
                                    */
                                    
                                    /*
                                    //THIS FAILS
            						System.out.println("MSWord_jsText: findOrReplace (Shapes): About to Dispatch.call((Dispatch) jacobShapeTextFrame, \"MoveRight\");");
                                    Dispatch.call((Dispatch) jacobShapeTextFrame, "MoveRight");
            						System.out.println("MSWord_jsText: findOrReplace (Shapes): About to Dispatch.call((Dispatch) jacobShapeTextFrame, \"MoveLeft\");");
                                    Dispatch.call((Dispatch) jacobShapeTextFrame, "MoveLeft");
            						*/
                                    
                                    /* THIS FAILS
                                    System.out.println("MSWord_jsText: findOrReplace (Shapes): About to Dispatch.call((Dispatch) jacobShape, \"MoveRight\");");
                                    Dispatch.call((Dispatch) jacobShape, "MoveRight");
            						System.out.println("MSWord_jsText: findOrReplace (Shapes): About to Dispatch.call((Dispatch) jacobShape, \"MoveLeft\");");
                                    Dispatch.call((Dispatch) jacobShape, "MoveLeft");
                                    */

                                    /* THIS FAILS
                                    System.out.println("MSWord_jsText: findOrReplace (Shapes): About to Dispatch.call((Dispatch) jacobShape, \"MoveRight\");");
                                    Dispatch.call(jacobShapeTextFrameTextRangeSelectVariant.getDispatch(), "Invoke", "MoveRight");
            						System.out.println("MSWord_jsText: findOrReplace (Shapes): About to Dispatch.call((Dispatch) jacobShape, \"MoveLeft\");");
                                    Dispatch.call(jacobShapeTextFrameTextRangeSelectVariant.getDispatch(), "Invoke", "MoveLeft");
                                    */		                                     
                                    
            						/*
                                    //THIS FAILS
                                    //Moving right removes the highlighting and places the cursor to the right of the replaced text.
            						//This is required, as otherwise, successive find/replace occurances may become confused.
            						System.out.println("MSWord_jsText: findOrReplace (Shapes): About to Dispatch.call(jacobShapeTextFrameTextRange, \"MoveRight\");");
            						//jacobSelection.invoke("MoveRight");
            						Dispatch.call(jacobShapeTextFrameTextRange, "MoveRight");
            						
            						//However, it's also necessary to go back to the left by one step afterwards,
            						//or otherwise, a seamlessly following [placeholders][seamlesslyFollowingPlaceholder] will NOT be found.
            						//The MoveRight - MoveLeft sequence has the effect that the selection = highlighting is removed from the inserted text.
            						System.out.println("MSWord_jsText: findOrReplace (Shapes): About to Dispatch.call(jacobShapeTextFrameTextRange, \"MoveLeft\");");
            						//jacobSelection.invoke("MoveLeft");
            						Dispatch.call(jacobShapeTextFrameTextRange, "MoveLeft");
            						*/
                    
                                    /*
                                    //THIS FAILS
                                    com.jacob.com.ComFailException: Can't map name to dispid: MoveRight
                                    
                                    //Moving right removes the highlighting and places the cursor to the right of the replaced text.
            						//This is required, as otherwise, successive find/replace occurances may become confused.
            						System.out.println("MSWord_jsText: findOrReplace (Shapes): About to Dispatch.call(jacobShapeTextFrameTextRangeFind, \"MoveRight\");");
            						//jacobSelection.invoke("MoveRight");
            						Dispatch.call(jacobShapeTextFrameTextRangeFind, "MoveRight");
            						*/

                                    /*
                                    //THIS FAILS
                                    com.jacob.com.ComFailException: Can't map name to dispid: MoveRight
                                    
                                    //Moving right removes the highlighting and places the cursor to the right of the replaced text.
            						//This is required, as otherwise, successive find/replace occurances may become confused.
            						System.out.println("MSWord_jsText: findOrReplace (Shapes): About to Dispatch.call(jacobShapeTextFrame, \"MoveRight\");");
            						//jacobSelection.invoke("MoveRight");
            						Dispatch.call(jacobShapeTextFrame, "MoveRight");
            						*/

                                    /*
                                    //THIS FAILS
                                    com.jacob.com.ComFailException: Can't map name to dispid: MoveRight

                                    //Moving right removes the highlighting and places the cursor to the right of the replaced text.
            						//This is required, as otherwise, successive find/replace occurances may become confused.
            						System.out.println("MSWord_jsText: findOrReplace (Shapes): About to Dispatch.call(jacobShape, \"MoveRight\");");
            						//jacobSelection.invoke("MoveRight");
            						Dispatch.call(jacobShape, "MoveRight");
            						*/
                                    
                                    /*
                                    //THIS FAILS (tried this just for completeness, probably at a level too high anyway)
                                    com.jacob.com.ComFailException: Can't map name to dispid: MoveRight
                                    //Moving right removes the highlighting and places the cursor to the right of the replaced text.
            						//This is required, as otherwise, successive find/replace occurances may become confused.
            						System.out.println("MSWord_jsText: findOrReplace (Shapes): About to Dispatch.call(jacobShapes, \"MoveRight\");");
            						//jacobSelection.invoke("MoveRight");
            						Dispatch.call(jacobShapes, "MoveRight");
            						*/

                                    /*
            						//TECHNICALLY, THIS WORKS - BUT IT DOES NOT MOVE THE CURSOR / CHANGE THE SELECTION WITHIN IN THE PROCESSED SHAPE.
            						//I.E. THE WHOLE TEXT STAYS SELECTED AFTER THE SUCCESSFUL REPLACEMENT. 
                                    
                                    //This is an attempt to re-use the jacobSelection provided for the main text portion find/replace,
                                    //it's probably the wrong target, but those on or below the Shape level won't accept the MoveRight so far: 
            						//Moving right removes the highlighting and places the cursor to the right of the replaced text.
            						//This is required, as otherwise, successive find/replace occurances may become confused.
            						System.out.println("MSWord_jsText: findOrReplace (test in main): About to jacobSelection.invoke(\"MoveRight\");");
            						jacobSelection.invoke("MoveRight");
            						*/										
                        			
        							//GETESTET: DAS MoveRight; MoveLeft; IST WIRKLICH NÖTIG, UM IM HAUPTTEXT ZUVERLÄSSIG ALLE PLATZHALTER ZU ERSETZEN. NICHT NÖTIG IN SHAPES.

                                    //Kommentare zur Info von oben übernommen:
                                    //
        	                        //DAS HIER SCHEINT ZU GEHEN!!!!! ENDLICH!!!
        	                        //(Analog der Methode, die über insertText()... cur ... pos hinwegegeholfen hat,
        	                        // wobei ich mich einfach nicht darum kümmere, eine Selection als Eigenschaft des aktuellen Textfeldes anzusprechen -
        	                        // sondern einfach eine Selection als Eigenschaft des ganzen Dokuments!)
        	                        //
        	                        //Object cur = jacobSelection.getObject();
        	                        //Dispatch.call((Dispatch) cur, "MoveRight");
        	                        //
        	                    	// UND DAS HIER AUCH - ist effektiv eine Kurzform davon:
                                    //Dispatch.call(jacobSelection, "MoveLeft");
            						//
                                    //THIS APPARENTLY WORKS!!!
                                    //System.out.println("MSWord_jsText: findOrReplace (Shapes): About to Dispatch.call(jacobSelection, \"MoveRight\");");
        	                        //Dispatch.call(jacobSelection, "MoveRight");
                                    //System.out.println("MSWord_jsText: findOrReplace (Shapes): About to Dispatch.call(jacobSelection, \"MoveLeft\");");
        	                        //Dispatch.call(jacobSelection, "MoveLeft");
        	                        //Nun. Das verschiebt zwar den Cursor ans Ende des Textes im Shape - aber immer noch wird nur EIN Platzhalter ersetzt...
                                    //for (int j = 0; j < 100; j++) { Dispatch.call(jacobSelection, "MoveLeft"); }
                                    //Das verschiebt den Cursor weit nach oben zum Anfang des Textfeldes - aber trotzdem wird nur EIN Platzhalter ersetzt...
            						
            						System.out.println("");
	            				}
	                		} catch (Exception ex) {
	                			ExHandler.handle(ex);
	                			//ToDo: Add precautions for pattern==null or pattern2==null...
	                			System.out.println("MSWord_jsText: findOrReplace (Shapes):\nFehler beim Suchen und Ersetzen im Haupttext:\n"+"" +
	                					"Exception caught für:\n"+
	                					"pattern \t(Von Elexis für OpenOffice):\t"+pattern+"\n"+
	                					"pattern2\t(Von msword_js für MS Word):\t"+pattern2+"\n"+		
	                					"numberOfHits="+numberOfHits);
	                			//SWTHelper.showError("No doc in bill", "Fehler:","Es ist keine Rechnungsvorlage definiert");
	                			//ToDo: Add precautions for pattern==null or pattern2==null...
	                			SWTHelper.showError(
	                					"MSWord_jsText: findOrReplace:"+ 
	                					"Fehler:",
	                					"MSWord_jsText: findOrReplace (Shapes):\nFehler beim Suchen und Ersetzen im Haupttext:\n"+
	                					"Exception caught für:\n"+
	                					"pattern \t(Von Elexis für OpenOffice):    \t"+pattern+"\n"+	//spaces needed for tab alignment in proportional font
	                					"pattern2\t(Von msword_js für MS Word):\t"+pattern2+"\n"+		
	                					"numberOfHits="+numberOfHits);
	                		}
	 
		                }
                	} while (jacobSearchResultInt == -1);
                } // (jacobShapeTextFrameHasText == -1)
            } //for (int i = 0; i < shapesCount; i++)            
        }
        catch (Exception ex) {
			ExHandler.handle(ex);
			//ToDo: Add precautions for pattern==null or pattern2==null...
			System.out.println("MSWord_jsText: findOrReplace:\nFehler beim Suchen und Ersetzen in Shapes:\n"+"" +
					"Exception caught für:\n"+
					"pattern \t(Von Elexis für OpenOffice):\t"+pattern+"\n"+
					"pattern2\t(Von msword_js für MS Word):\t"+pattern2+"\n"+		
					"numberOfHits="+numberOfHits);
			//SWTHelper.showError("No doc in bill", "Fehler:","Es ist keine Rechnungsvorlage definiert");
			//ToDo: Add precautions for pattern==null or pattern2==null...
			SWTHelper.showError(
					"MSWord_jsText: findOrReplace:"+ 
					"Fehler:",
					"MSWord_jsText: findOrReplace:\nFehler beim Suchen und Ersetzen in Shapes:\n"+
					"Exception caught für:\n"+
					"pattern \t(Von Elexis für OpenOffice):    \t"+pattern+"\n"+	//spaces needed for tab alignment in proportional font
					"pattern2\t(Von msword_js für MS Word):\t"+pattern2+"\n"+		
					"numberOfHits="+numberOfHits);
        }
        		
		
		
		/*
		ISearchResult searchResult = agIonDoc.getSearchService().findAll(search);
		if (!searchResult.isEmpty()) {
			ITextRange[] textRanges = searchResult.getTextRanges();
			if (cb != null) {
				for (ITextRange r : textRanges) {
					String orig = r.getXTextRange().getString();
					Object replace = cb.replace(orig);
					if (replace == null) {
						r.setText("??Auswahl??");
					} else if (replace instanceof String) {
						// String repl=((String)replace).replaceAll("\\r\\n[\\r\\n]*", "\n")
						String repl = ((String) replace).replaceAll("\\r", "\n");
						repl = repl.replaceAll("\\n\\n+", "\n");
						r.setText(repl);
					} else if (replace instanceof String[][]) {
						String[][] contents = (String[][]) replace;
						try {
							ITextTable textTable =
								agIonDoc.getTextTableService().constructTextTable(contents.length,
									contents[0].length);
							agIonDoc.getTextService().getTextContentService().insertTextContent(r,
								textTable);
							r.setText("");
							ITextTablePropertyStore props = textTable.getPropertyStore();
							// long w=props.getWidth();
							// long percent=w/100;
							for (int row = 0; row < contents.length; row++) {
								String[] zeile = contents[row];
								for (int col = 0; col < zeile.length; col++) {
									textTable.getCell(col, row).getTextService().getText().setText(
										zeile[col]);
								}
							}
							textTable.spreadColumnsEvenly();
							
						} catch (Exception ex) {
							ExHandler.handle(ex);
							r.setText("Fehler beim Ersetzen");
						}
						
					} else {
						r.setText("Not a String");
					}
				}
			}
		
			
			System.out.println("MSWord_jsText: findOrReplace: about to end, returning true...");
			return true;
		}
		*/
		
		
		
		
		
		
		
		//ToDo: Lokales umschalten der debugSysPrnFindOrReplaceDetails nach Debugging von findOrReplace (SectionHeaders) vorher und nacher wieder entfernen.	//201701120109js
		//Siehe debugSysPrnFindOrReplaceDetails=true/=false; vor bzw. nach dem findOrReplace (SectionHeaders) Teil.												//201701120109js
		System.out.println("MSWord_jsText: debugSysPrnFindOrReplaceDetails=false (true kann hier TEMPORAER EINGESCHALTET werden für Debugging von findOrReplace(SectionHeaders);");
		debugSysPrnFindOrReplaceDetails=false;																													//201701120109js
		
		System.out.println("");
		System.out.println("MSWord_jsText: ProbablyUsingTarmed_xxTemplateSoDoNOTAccessHeaderRangeToAvoidGenerationOfEmptyHeaderLines=="+ProbablyUsingTarmed_xxTemplateSoDoNOTAccessHeaderRangeToAvoidGenerationOfEmptyHeaderLines.toString());
		if ( !ProbablyUsingTarmed_xxTemplateSoDoNOTAccessHeaderRangeToAvoidGenerationOfEmptyHeaderLines ) {
			System.out.println("MSWord_jsText: CONSEQUENTLY, I WILL ACCESS the header section to search for placeholders there and process them if found.");	 		
			System.out.println("MSWord_jsText:               This may cause the createion of an empty HEADER section on page 1 containing ONLY ONE NEWLINE symbol,");	 		
		    System.out.println("MSWord_jsText:               and this otherwise empty HEADER can NOT be removed again programmatically, and only uncertainly manually.");	 		
		} else {
			System.out.println("MSWord_jsText: CONSEQUENTLY, I WILL NOT ACCESS the header section to avoid the creation of an unwanted HEADER section");
			System.out.println("MSWord_jsText: 		         on page 1 that could afterwards NOT reliably be removed again.");
		}	
		System.out.println("");
		
		if ( !ProbablyUsingTarmed_xxTemplateSoDoNOTAccessHeaderRangeToAvoidGenerationOfEmptyHeaderLines ) { 
		
		//The following block performs search-and-replace for the header area (of all sections) of the document.
		//An almost identical block is further below for shapes (without updating all the comments).
		
		if (debugSysPrnFindOrReplaceDetails) {
			System.out.println("");
			System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to find/replace in Headers: ActiveDocument.Sections(i).Headers(wdHeaderFooterPrimary).Range.Text...");
			System.out.println("");
		}
		
		//ActiveXComponent jacobSelection = jacobObjWord.getPropertyAsComponent("Selection");
		//ActiveXComponent jacobFind = jacobSelection.getPropertyAsComponent("Find");
		jacobSelection = jacobObjWord.getPropertyAsComponent("Selection");
		jacobFind = jacobSelection.getPropertyAsComponent("Find");

		try {
            Dispatch jacobSections = Dispatch.get((Dispatch) jacobDocument, "Sections").toDispatch();
            int sectionsCount = Dispatch.get(jacobSections , "Count").getInt();
            if (debugSysPrnFindOrReplaceDetails)
        		System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): sectionsCount="+sectionsCount);
	        
            
            
            for (int i = 0; i < sectionsCount; i++) {
            	if (debugSysPrnFindOrReplaceDetails)
            		System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): i=="+i);

	            Variant jacobSectionVariant = Dispatch.call(jacobSections, "Item", new Variant(i + 1));
	            if (debugSysPrnFindOrReplaceDetails) {
	            	if (jacobSectionVariant == null)	System.out.println("MSWord_jsText: findOrReplace (Headers): WARNING: jacobSectionVariant IS NULL");
	            	else 								System.out.println("MSWord_jsText: findOrReplace (Headers): jacobSectionVariant="+jacobSectionVariant.toString());
	            }

	            if (debugSysPrnFindOrReplaceDetails)
	            	System.out.println("MSWord_jsText: findOrReplace (Headers): About to Dispatch jacobSection = jacobSectionVariant.toDispatch();");
                Dispatch jacobSection = jacobSectionVariant.toDispatch();
                
                Dispatch jacobSectionHeaders = Dispatch.get((Dispatch) jacobSection, "Headers").toDispatch();
                
                if (jacobSectionHeaders == null)	System.out.println("MSWord_jsText: findOrReplace (Headers): WARNING: jacobSectionHeaders IS NULL");
	            int sectionHeadersCount = Dispatch.get(jacobSectionHeaders , "Count").getInt();
	            if (debugSysPrnFindOrReplaceDetails)
	            	System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): sectionHeadersCount="+sectionHeadersCount);
		            
	            
	            //Wegen sectionHeadersCount = 3 wird der for Block 3x durchlaufen, das bewirkt aber wiederholt Ersetzungen im gleichen SectionHeaders block.
	            //VIELLEICHT Wirkt sich das nur dann nützlich aus, wenn "Erste Seite anders" oder "Linke / Rechte Seite anders" gewählt wurde, ich lasse es mal so.
	            for (int j = 0; j < sectionHeadersCount; j++) {
	            	if (debugSysPrnFindOrReplaceDetails)
	            		System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): Section i=="+i+"; SectionHeader j=="+j);			            

		            do {
        				jacobSearchResultInt = 0;	//Reset this to 0 in each loop iteration, so that even when jacobFind.invoke() should NOT put a valid result into this variable
													//(or when it should not even get invoked!),
													//we do definitely NOT get an endless loop, NOR a misguided attempt to replace text. 	

        				//Hier muss das ganze Procedere, schon beginnend mit Dispatch jacobSectionHeaderVariant = ... in den do...while block (also noch mehr als bei (Shapes))
                		//damit mehrere Platzhalter innerhalb eines SectionHeaders ersetzt werden.
                		//Hab's ausprobiert, wenn erst ab jacobSectionHeaderRangeText = ... hier drin stand, wurden nur 3 Ersetzungen ausgeführt,
                		//und zwar weil der for-Block bei sectionHeadersCount = 3 auch 3x durchlaufen wird.
                		
        				if (debugSysPrnFindOrReplaceDetails)
        					System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to Variant jacobSectionHeadersVariant = Dispatch.call(jacobSectionHeaders, \"Item\", new Variant(i + 1));");
			            Variant jacobSectionHeaderVariant = Dispatch.call(jacobSectionHeaders, "Item", new Variant(i + 1)); //Sorry, das liefert tatsächlich letzendlich Zugang zum Haupttext
			            
			            if (debugSysPrnFindOrReplaceDetails) {
			            	if (jacobSectionHeaderVariant == null)	System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): WARNING: jacobSectionHeaderVariant IS NULL");
			            	else 									System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): jacobSectionHeaderVariant="+jacobSectionHeaderVariant.toString());
			            }

			            
			            if (debugSysPrnFindOrReplaceDetails)
			            	System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to Dispatch jacobSectionHeader = jacobSectionHeaderVariant.toDispatch();");
		                Dispatch jacobSectionHeader = jacobSectionHeaderVariant.toDispatch();
		                if (debugSysPrnFindOrReplaceDetails)
		                	System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): jacobSectionHeader.toString() == "+jacobSectionHeader.toString());

//Hier (oder früher) per if (false) ... ausblenden erhält nicht-existierende Kopfzeile nicht existierend.
		                		               
		                //Das liefert leider immer true, selbst wenn noch *gar* kein Inhalt (und auch keine y-Platz-verbrauchende Absatzmarke) in der Kopfzeile existiert.
		                if (debugSysPrnFindOrReplaceDetails)
		                	System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): Dispatch.get(jacobSectionHeader, \"Exists\") == "+Dispatch.get(jacobSectionHeader, "Exists"));
			            //Das liefert leider immer true, selbst wenn noch *gar* kein Inhalt (und auch keine y-Platz-verbrauchende Absatzmarke) in der Kopfzeile existiert.
		                if (debugSysPrnFindOrReplaceDetails)
		                	System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): Dispatch.get(jacobSectionHeaderVariant.toDispatch(), \"Exists\") == "+Dispatch.get(jacobSectionHeaderVariant.toDispatch(), "Exists"));

		                if (debugSysPrnFindOrReplaceDetails) {
		                	System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): Dispatch.get(jacobSectionHeader, \"Application\") == "+Dispatch.get(jacobSectionHeader, "Application"));
		                	System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): Dispatch.get(jacobSectionHeader, \"Creator\") == "+Dispatch.get(jacobSectionHeader, "Creator"));
		                	System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): Dispatch.get(jacobSectionHeader, \"Index\") == "+Dispatch.get(jacobSectionHeader, "Index"));
		                	System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): Dispatch.get(jacobSectionHeader, \"IsHeader\") == "+Dispatch.get(jacobSectionHeader, "IsHeader"));
		                	System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): Dispatch.get(jacobSectionHeader, \"Parent\") == "+Dispatch.get(jacobSectionHeader, "Parent"));
		                }

			            //WARNUNG: DIESER AUFRUF SORGT DAFÜR, DASS ANSCHLIESSEND SICHER EINE KOPFZEILE EXISTIERT.
		                //WENN ZUVOR KEINE DA WAR, DANN ERSCHEINT EINE, DIE EINE ABSATZMARKE ENTHÄLT.
		                //DIESE IST KAUM WIEDER WEGZUBEKOMMEN UND VERDRÄNGT DEN INHALT DES EIGENTLICHEN DOKUMENT-TEXT-BEREICHS NACH UNTEN;
		                //z.B. IN DEN VORLAGEN Tarmedrechnung_xx WIRD DADURCH DER AUFBAU DES LAYOUTS KOMPLETT GESTÖRT, WEIL DIE TITELZEILE UNTER DIE TEXTFRAMES VERSETZT WIRD.
		                //DIESER AUFRUF SOLLTE EIGENTLICH NUR ERFOLGEN, WENN DIE KOPFZEILE SCHON EINEN NICHT LEEREN INHALT HAT - ABER WIE DAS VORHER HERAUSBEKOMMEN?!?
		                //...Header.Exists und ...HeaderVariant.Exists ist auch bei leeren, nicht sichtbaren, als nicht-existent erscheinenden Kopfzeilen = true (getestet).
			            //Siehe direkt darüber - auch alle anderen Eigenschaften von ... Header helfen nicht beim Erkennen, ob die Kopfzeile schon/noch_nicht sichtbar ist.

		                if (debugSysPrnFindOrReplaceDetails) {
		                	System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to Dispatch jacobSectionHeaderRange = Dispatch.call(jacobSectionHeader, \"Range\").toDispatch();");
		                	System.out.println("MSWord_jsText:   THIS CALL WILL BRING A HEADER SECTION INTO EXISTENCE (with one newline symbol) WHERE NONE WAS THERE BEFORE.");
		                	System.out.println("MSWord_jsText:   THIS UNWANTED HEADER SECTION CAN NOT BE RELIABLY REMOVED PROGRAMMATICALLY, AND ONLY WITH LUCK IF TRIED MANUALLY.");
		                }
			            Dispatch jacobSectionHeaderRange = Dispatch.call(jacobSectionHeader, "Range").toDispatch();
			            if (debugSysPrnFindOrReplaceDetails)
			            	System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): jacobSectionHeaderRange.toString() == "+jacobSectionHeaderRange.toString());

			            //Das funktioniert zwar, und löscht sogar vorher existierende Kopfzeilen mit Inhalt komplett - hilft jedoch nicht, den Platz der Kopfzeilen auch völlig freizugeben:
			            //Dispatch.call(jacobSectionHeaderRange, "Delete");

//Hier per if (false) ... ausblenden reicht noch aus, um nicht-existierende Kopfzeile nicht existierend zu erhalten.				//201701120109js
//Dabei wird allerdings jeglicher Support für das Ersetzen von Platzhaltern in Kopfzeilen entfernt.			            			//201701120109js

System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): *** BEGINN DES OPTIONAL PER if (false)... ausgeblendeten aktiven Teils.");	//201701120109js
/* if (false) */ {			            																								//201701120109js

						if (debugSysPrnFindOrReplaceDetails)
							System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to String jacobSectionHeaderRangeText = Dispatch.get(jacobSectionHeaderRange, \"Text\").toString();");
	                	String jacobSectionHeaderRangeText = Dispatch.get(jacobSectionHeaderRange, "Text").toString();

	                    
	                	//Falls die Kopfzeile ganz leer ist, NICHT DRAN RUMFUMMELN!
	                	//Wenn man in Word einmal die Kopfzeile öffnet (welche allenfalls vorher in der Vorlage nicht sichtbar und nicht Platz verbrauchend vorhanden war),
	                	//dann besteht die sehr hohe Gefahr, dass nacher permanent eine Absatzmarke (und sonst nichts) in der Kopfzeile verbleibt, und diese braucht Platz.
	                	//Der Haupttext wird dann um eine Zeile nach unten gerückt - und in den Tarmed-Vorlagen verschiebt das den [Titel] unter die umfangreichen Blöcke am Anfang.
	                	//Das ist kaum reversibel (ich hab's mal bei ein paar Versuchen hinbekommen, durch Ansicht - Kopf-und-Fusszeile wählen,
	                	//aber nicht sicher reproduzierbar und schon gar nicht VBAJava+Jacob-skriptbar.
	                	
	                    if (jacobSectionHeaderRangeText == null)			{ System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): WARNING: Section["+(i+1)+"].SectionHeader["+(j+1)+"].RangeText IS NULL. Skipping further find/replace in Section header."); }  //201701120138js corrected position info w/ Section/SectionHeader and indizes. 
	                    else if (jacobSectionHeaderRangeText.length() < 3)	{ System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): WARNING: Section["+(i+1)+"].SectionHeader["+(j+1)+"].RangeText.length() == "+jacobSectionHeaderRangeText.length()+". Skipping further find/replace in Section header."); }	//201701120138js corrected position info w/ Section/SectionHeader and indizes.	                
	                    else {
	                    	if (debugSysPrnFindOrReplaceDetails)
	                    		System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): Section["+(i+1)+"].SectionHeader["+(j+1)+"].Range.Text="+jacobSectionHeaderRangeText);	//201701151710js korr debug output

	                    	//THIS WORKS, and causes the text in the SectionHeader to become selected            
	                    	if (debugSysPrnFindOrReplaceDetails)
	                    		System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to Variant jacobSectionHeaderRangeSelectVariant = Dispatch.call(jacobSectionHeaderRange, \"Select\");");
	                        Variant jacobSectionHeaderRangeSelectVariant = Dispatch.call(jacobSectionHeaderRange, "Select");
	                        if (debugSysPrnFindOrReplaceDetails) {
	                        	if (jacobSectionHeaderRangeSelectVariant == null)	System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): WARNING: jacobSectionHeaderRangeSelectVariant IS NULL");
	                        	else 	System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): jacobSectionHeaderRangeTextSelectVariant="+jacobSectionHeaderRangeSelectVariant.toString());
	                        }
	                        
	                        //DAS HIER SCHEINT ZU GEHEN!!!!! ENDLICH!!!
	                        //(Analog der Methode, die über insertText()... cur ... pos hinwegegeholfen hat,
	                        // wobei ich mich einfach nicht darum kümmere, eine Selection als Eigenschaft des aktuellen Textfeldes anzusprechen -
	                        // sondern einfach eine Selection als Eigenschaft des ganzen Dokuments!)
	                        //
	                        //Object cur = jacobSelection.getObject();
	                        //Dispatch.call((Dispatch) cur, "MoveLeft");
	                        //
	                    	// UND DAS HIER AUCH - ist effektiv eine Kurzform davon:
	                        //Dispatch.call(jacobSelection, "MoveLeft");

	                        if (debugSysPrnFindOrReplaceDetails)
	                        	System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to Variant jacobSectionHeaderRangeTextFindVariant = Dispatch.call(jacobSectionHeaderRangeText, \"Find\");");
	                        Variant jacobSectionHeaderRangeTextFindVariant = Dispatch.get(jacobSectionHeaderRange, "Find");
	                        if (debugSysPrnFindOrReplaceDetails) {
	                        	if (jacobSectionHeaderRangeTextFindVariant == null)	System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): WARNING: jacobSectionHeaderRangeTextFindVariant IS NULL");
	                        	else 								System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): jacobSectionHeaderRangeTextFindVariant="+jacobSectionHeaderRangeTextFindVariant.toString());                        
	                        }
	                        
	                        if (debugSysPrnFindOrReplaceDetails)
	                        	System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to Dispatch jacobSectionHeaderRangeTextFind = jacobSectionHeaderRangeTextFindVariant.toDispatch();");
	                        Dispatch jacobSectionHeaderRangeTextFind = jacobSectionHeaderRangeTextFindVariant.toDispatch();
	                        if (debugSysPrnFindOrReplaceDetails) {
	                        	if (jacobSectionHeaderRangeTextFind  == null)	System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): WARNING: jacobSectionHeaderRangeTextFind IS NULL");
	                        	else 							System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): jacobSectionHeaderRangeTextFind="+jacobSectionHeaderRangeTextFind.toString());
	                        }
	                        
	                        if (debugSysPrnFindOrReplaceDetails)
	                        	System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to try Dispatch.call(jacobSectionHeaderRangeTextFind, \"Text\", pattern2);... etc.");
	                		try {
	                			Dispatch.put(jacobSectionHeaderRangeTextFind, "Text", pattern2);
	                			Dispatch.put(jacobSectionHeaderRangeTextFind, "Forward", "True");
	                			Dispatch.put(jacobSectionHeaderRangeTextFind, "Format", "False");
	                			Dispatch.put(jacobSectionHeaderRangeTextFind, "MatchCase", "False");
	                			Dispatch.put(jacobSectionHeaderRangeTextFind, "MatchWholeWord", "False");
	                			Dispatch.put(jacobSectionHeaderRangeTextFind, "MatchByte", "False");
	                			Dispatch.put(jacobSectionHeaderRangeTextFind, "MatchAllWordForms", "False");
	                			Dispatch.put(jacobSectionHeaderRangeTextFind, "MatchSoundsLike", "False");
	                			Dispatch.put(jacobSectionHeaderRangeTextFind, "MatchWildcards", "True");	
	                		} catch (Exception ex) {
	                			ExHandler.handle(ex);
	                			System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): Fehler beim Ersetzen: Dispatch.put(jacobSectionHeaderRangeTextFind...);");
	                		
	                			//SWTHelper.showError("No doc in bill", "Fehler:","Es ist keine Rechnungsvorlage definiert");
	                			SWTHelper.showError("findOrReplace (SectionHeaders): Fehler beim Ersetzen: ", "Fehler:","Fehler beim Ersetzen: Dispatch.put(jacobSectionHeaderRangeTextFind...);");
	                		}
	                    	
	                		//The following block performs search-and-replace for each SectionHeaders text portion of the document.
	                		//An almost identical block is further above for the main text (without updating all the comments),
	                		//and will probably be added further below, for tables. 
	                		
	                		if (debugSysPrnFindOrReplaceDetails)
	                			System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to try ... the actual search and replace block...");
	                		try {	 
	                			if (debugSysPrnFindOrReplaceDetails) {
	                				if (pattern2 == null)	System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): ERROR: pattern2 IS NULL!");
	                				else 					System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): pattern2="+pattern2);
	                			}
                				
                				try {
                					if (debugSysPrnFindOrReplaceDetails)
                						System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to jacobSearchResultInt = Dispatch.call(jacobSectionHeaderRangeTextFind,\"Execute\").toInt();...");
                					//jacobSearchResultInt = jacobFind.invoke("Execute").toInt();
                					jacobSearchResultInt = Dispatch.call(jacobSectionHeaderRangeTextFind,"Execute").toInt();
                					
                					//Please note: Wenn ich erste jacobSearchResultInt = jacobSearchresultVariant.toInt() verwende, ist nacher auch der string = ""-1", sonst "true"" +
                					if (debugSysPrnFindOrReplaceDetails)
                						System.out.println("MSWord_jsText: findOrReplace  jacobSearchResultInt="+jacobSearchResultInt);
                					//System.out.println("Result: jacobInvokeResult.toString()="+jacobInvokeResult.toString());	//Returns true if match found, false if no match found
                					//System.out.println("Result: jacobInvokeResult.toInt()="+jacobInvokeResult.toInt());		//Returns -1 if match found, 0 if no match found
                					//System.out.println("Result: jacobInvokeResult.toError()="+jacobInvokeResult.toError());	//Throws java.lang.IllegalStateException: getError() only legal on Variants of type VariantError, not 3
                				} catch (Exception ex) {
                					ExHandler.handle(ex);
                					//ToDo: Add precautions for pattern==null or pattern2==null...
                					System.out.println("MSWord_jsText: findOrReplace (SectionHeaders):\nException caught.\n"+
                					"Das verwendete Suchpattern funktioniert möglicherweise NICHT mit MS Word.\n"+
	            					"pattern \t(Von Elexis für OpenOffice):\t"+pattern+"\n"+
                					"pattern2\t(Von msword_js für MS Word):\t"+pattern2+"\n"+		
                					"Falls neue Suchpatterns hinzugefügt wurden, muss möglicherweise eine on-the-fly-Konversion in com.jsigle.msword_js.MSWord_jsText.java ergänzt werden.");

                					//ToDo: Add precautions for pattern==null or pattern2==null...
	            					SWTHelper.showError(
	            							"MSWord_jsText: findOrReplace (SectionHeaders):", 
	            							"Fehler:",
	            							"Exception caught.\n"+
	            							"Das verwendete Suchpattern funktioniert möglicherweise NICHT mit MS Word.\n"+
	            							"pattern \t(Von Elexis für OpenOffice):    \t"+pattern+"\n"+	//spaces needed for tab alignment in proportional font
	            							"pattern2\t(Von msword_js für MS Word):\t"+pattern2+"\n"+		
	            							"Falls neue Suchpatterns hinzugefügt wurden, muss möglicherweise eine on-the-fly-Konversion in com.jsigle.msword_js.MSWord_jsText.java ergänzt werden.");
                				}
                				
                				if (jacobSearchResultInt == -1) {		//Do ONLY attempt a replacement if there was a search hit. Otherwise, replacement text might be simply inserted at cursor position.
            						numberOfHits += 1;
            						if (debugSysPrnFindOrReplaceDetails)
            							System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): numberOfHits="+numberOfHits);
            						
            			/* SIMPLE GET, PRINTLN, REPLACE WITH CONSTANT STRING FOR TESTING ONLY - THIS WORKS :-) :-) :-)
            						//and very fine: especially, in Bern, [Brief.Datum] (or similar), only the [Brief.Datum] portion is replaced.
            						//This means, that the "Find" stuff actually works and controls the range that is influenced by the following commands. PUH.
            						System.out.println(Dispatch.get(jacobSectionHeaderRange, "Text").toString());
            						System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to Dispatch.put(jacobSectionHeaderRangeTextFind, \"Text\", new Variant(\"XRplX\");");
            						Dispatch.put(jacobSectionHeaderRange, "Text", new Variant("XRplX"));
            			*/
									//Callback an Elexis den gefundenen [Platzhalter] übergeben und um replacement bitten
									
									String orig = Dispatch.get(jacobSectionHeaderRange, "Text").toString();
									if (debugSysPrnFindOrReplaceDetails) {
										if (orig  == null)	System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): ERROR: orig IS NULL!");
										else 				System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): orig="+orig);
									
										if (cb == null)		System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): ERROR: cb IS NULL!");
										else 				System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): cb="+cb.toString());
									}
									
									if (debugSysPrnFindOrReplaceDetails)
										System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to Object replace = cb.replace(orig);...");
									Object replace = cb.replace(orig);
									
									if (debugSysPrnFindOrReplaceDetails) {
										if (replace == null)	System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): ERROR: replace IS NULL!");
										else 					System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): replace="+replace.toString());
									}
										
									if (replace == null) {									//Falls nichts brauchbares zurückkommt: Fehlermeldung in den Text setzen.
										if (debugSysPrnFindOrReplaceDetails)
											System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to Dispatch.put(jacobSectionHeaderRange, \"Text\", \"??Auswahl??\");");							
										//jacobSelection.setProperty("Text", "??Auswahl??");	//ToDo: Das ist ein wenig dämlich: besser wäre es, NICHTS zu tun, dann könnte man den Platzhalter noch sehen.
										Dispatch.put(jacobSectionHeaderRange, "Text", "??Auswahl??");
									} else if (replace instanceof String) {
										if (debugSysPrnFindOrReplaceDetails)
											System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to replace \\r and \\n and their combinations by (traditionally) suitable linebreaks...");
										// String repl=((String)replace).replaceAll("\\r\\n[\\r\\n]*", "\n")
										String repl = ((String) replace).replaceAll("\\r", "\n");
										repl = repl.replaceAll("\\n\\n+", "\n");
										if (debugSysPrnFindOrReplaceDetails) {
											if (repl == null)	System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): ERROR: repl IS NULL!");
											else 				System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): repl="+repl);
											if (repl != null) {
												System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): repl.length()="+repl.length());														
											}
										}
										if (debugSysPrnFindOrReplaceDetails)
											System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to Dispatch.put(jacobSectionHeaderRange, \"Text\", repl);");
										//jacobSelection.setProperty("Text", repl);
										Dispatch.put(jacobSectionHeaderRange, "Text", repl);
										
									} else if (replace instanceof String[][]) {
										String[][] contents = (String[][]) replace;
										//ToDo: Handler für Tabellen-Einfügung als String-Ersetzung noch hinzufügen.
										System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO");
										System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): ToDo Tabellen-Einfügung als String-Ersetzung noch hinzufügen! (only if Tables in jacobSectionHeaderRange shall be supported...)");
										System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO");
										
										if (debugSysPrnFindOrReplaceDetails)
											System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to Dispatch.put(jacobShapeTextFrameTextRange, \"Text\", \"MSWord_jsText.java ToDo: Tabellen-Einfügung als String-Ersetzung noch hinzufügen!\");");
										//jacobSelection.setProperty("Text", "MSWord_jsText.java ToDo: Tabellen-Einfügung als String-Ersetzung noch hinzufügen!");
										Dispatch.put(jacobSectionHeaderRange, "Text", "MSWord_jsText.java ToDo:  (SectionHeaders) Tabellen-Einfügung als String-Ersetzung noch hinzufügen!");
										
										/*
										try {
											ITextTable textTable =
												agIonDoc.getTextTableService().constructTextTable(contents.length,
													contents[0].length);
											agIonDoc.getTextService().getTextContentService().insertTextContent(r,
												textTable);
											r.setText("");
											
											ITextTablePropertyStore props = textTable.getPropertyStore();
											// long w=props.getWidth();
											// long percent=w/100;
											for (int row = 0; row < contents.length; row++) {
												String[] zeile = contents[row];
												for (int col = 0; col < zeile.length; col++) {
													textTable.getCell(col, row).getTextService().getText().setText(
														zeile[col]);
												}
											}
											textTable.spreadColumnsEvenly();
											
										} catch (Exception ex) {
											ExHandler.handle(ex);
											r.setText("Fehler beim Ersetzen");
										}               
										*/ 							
									} else {
										if (debugSysPrnFindOrReplaceDetails)
											System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to Dispatch.put(jacobSectionHeaderRange, \"Text\", \"Not a String\");");
										//jacobSelection.setProperty("Text", "Not a String");
										Dispatch.put(jacobSectionHeaderRange, "Text", "Not a String");
									}
            						
            						//GETESTET: OFFENBAR NICHT NÖTIG FÜR ERSETZUNGEN IM SectionHeader (=Kopfzeile)
            						
            						//THIS APPARENTLY WORKS!!!
                                    //System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to Dispatch.call(jacobSelection, \"MoveRight\");");
        	                        //Dispatch.call(jacobSelection, "MoveRight");
                                    //System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to Dispatch.call(jacobSelection, \"MoveLeft\");");
        	                        //Dispatch.call(jacobSelection, "MoveLeft");
        	                        //Nun. Das verschiebt zwar den Cursor ans Ende des Textes im SectionHeader - aber immer noch wird nur EIN Platzhalter ersetzt...
                                    //for (int j = 0; j < 100; j++) { Dispatch.call(jacobSelection, "MoveLeft"); }
                                    //Das verschiebt den Cursor weit nach oben zum Anfang des Textfeldes - aber trotzdem wird nur EIN Platzhalter ersetzt...
            						
									if (debugSysPrnFindOrReplaceDetails)
										System.out.println("");
                				}		                				
	                		} catch (Exception ex) {
	                			ExHandler.handle(ex);
	                			//ToDo: Add precautions for pattern==null or pattern2==null...
	                			System.out.println("MSWord_jsText: findOrReplace (SectionHeaders):\nFehler beim Suchen und Ersetzen im Header:\n"+"" +
	                					"Exception caught für:\n"+
	                					"pattern \t(Von Elexis für OpenOffice):\t"+pattern+"\n"+
	                					"pattern2\t(Von msword_js für MS Word):\t"+pattern2+"\n"+		
	                					"numberOfHits="+numberOfHits);
	                			//SWTHelper.showError("No doc in bill", "Fehler:","Es ist keine Rechnungsvorlage definiert");
	                			//ToDo: Add precautions for pattern==null or pattern2==null...
	                			SWTHelper.showError(
	                					"MSWord_jsText: findOrReplace:"+ 
	                					"Fehler:",
	                					"MSWord_jsText: findOrReplace (SectionHeaders):\nFehler beim Suchen und Ersetzen im Haupttext:\n"+
	                					"Exception caught für:\n"+
	                					"pattern \t(Von Elexis für OpenOffice):    \t"+pattern+"\n"+	//spaces needed for tab alignment in proportional font
	                					"pattern2\t(Von msword_js für MS Word):\t"+pattern2+"\n"+		
	                					"numberOfHits="+numberOfHits);
	                		}
	 
		                }
} //if (false) ...																													//201701120109js
System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): *** ENDE DES OPTIONAL PER if (false)... ausgeblendeten aktiven Teils.");		//201701120109js

		            } while (jacobSearchResultInt == -1); 	                                
	         	} //for (int j = 0; j < SectionHeadersCount; j++)
	        } //for (int i = 0; i < SectionsCount; i++)
		} catch (Exception ex) {
			ExHandler.handle(ex);
			System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): Fehler bei enumeration von jacobSections...");
		}	//201701151242js Korrektur: diese Klammer muss hier hoch, sonst wird der nachfolgende Code zur Rückkehr aus dem HeaderViewTeil NIE ausgeführt, der gehört nicht zum Exception Handler!
			//201701151242js Wenn die Klammer hier oben ist, und der nachfolgende Code ausgeführt wird,
			//201701151242js dann funktioniert mein nachfolgender Code für die Rückkehr in Seitenlayout und HauptDokument nach dem Ersetzen in den Kopfzeilen auch :-)
			

		
		
		//N.B.: Nachfolgend steht zu mehreren Ansätzen, dass diese nicht funktioniert hätten,
		//um wieder zurück in die Seitenansicht und in den Haupttext zu kommen.
		//Ich hab nun vor 201701151242 gefunden, dass ich allen folgenden Code durch eine zu tief stehende } Klammer
		//in dem Exception handler hatte; d.h. dieser wurde typischerweise gar nicht ausgeführt.
		//Ich weiss nicht, ob das durchgängig während der Entwicklung des Rückkehr-Codes so war (kann's mir aber kaum vorstellen,
		//da ja auch die Monitoring-Console-Ausgaben dann komplett abwesend gewesen wären). Falls doch, kännten darin mglw. auch
		//noch einfachere (bzw. direktere) Wege zur Rückkehr in die gewünschte Ansicht sein.
		//Nachdem ich das jetzt alles durchgesehen und dann die Klammer korrigiert habe,
		//funktioniert die ganze Lösung (incl. des Prüfens auf Länge der Header, Vermeidens der Berührung leerer Header
		//und damit Vermeidens der Erzeugung von fast-leeren-nicht-löschbaren-Headern mit nur einem Zeilenumbruch drin,
		//Ersetzen von Platzhaltern in den Headern, und Rücckehr in das Hauptdokument und zur Seitenansicht gut und schnell,
		//so dass ich am nachfolgenden Code jetzt auch nichts sogleich wieder ändern will.	//201701151242js
		
		
		
		
		//DO NOT MESS AROUND WITH HEADER-/FOOTER-VIEW if we are not in "Normal" view now.
		//Because activating the HEADER section may introduce an empty Header line that was not there before,
		//and contains only a paragraph mark, but needs space, and can not (or hardly) be got rid of again. (In MS Word, that is.)
		//This will mess up the layout for Tarmedrechnung_xx form templates, because it sends the [Titel]... line down below the textboxes,
		//and thereafter, the tarmed bill table far far down towards the foot of the form.
		//Once there, this (empty) Header line can hardly be made go away, my few attempts that succeeded were not reproducable, and not scriptable at all. 
				
		
		//Find/Replace in the Headers section switches Word to Normal view, Splits the MS Word Window into two sections, with the Headers Section in the lower one and the cursor there.
		//Now, we want to get back to Page Layout View:
		
		//THIS FAILS (all of them)
		//com.jacob.com.ComFailException: Can't map name to dispid: View
		//Variant jacobViewVariant = Dispatch.get(jacobObjWord, "View");
		//Variant jacobViewVariant = Dispatch.get(jacobDocument, "View");
		//Variant jacobViewVariant = Dispatch.get(jacobSelection, "View");
		//jacobObjWord.setProperty("View", 0);
		//jacobObjWord.getPropertyAsComponent("View");
		//jacobSelection.getPropertyAsComponent("View");
		
		//THIS FAILS, so we actually must go down step by step...
		//ActiveXComponent jacobActiveWindowActivePaneViewAXC = jacobObjWord.getPropertyAsComponent("ActiveWindow.ActivePane");

		//ToDo: Is the following necessary for anything in MSWord_js?
		System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): TODO: Is the following necessary for anything in MSWord_js?");
		
		System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to ActiveXComponent jacobDocumentAXC = jacobSelection.getPropertyAsComponent(\"Document\");");
		ActiveXComponent jacobDocumentAXC = jacobSelection.getPropertyAsComponent("Document");
		
		System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to ActiveXComponent jacobActiveWindowAXC = jacobObjWord.getPropertyAsComponent(\"ActiveWindow\");");
		ActiveXComponent jacobActiveWindowAXC = jacobObjWord.getPropertyAsComponent("ActiveWindow");
		
		System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to ActiveXComponent jacobActiveWindowViewAXC = jacobActiveWindowAXC.getPropertyAsComponent(\"View\");");
		ActiveXComponent jacobActiveWindowViewAXC = jacobActiveWindowAXC.getPropertyAsComponent("View");

		//If ActiveWindow.View.SplitSpecial <> wdPaneNone Then
        //  ActiveWindow.Panes(2).Close
		//End If
		
		System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to ActiveXComponent jacobActiveWindowPanesAXC = jacobActiveWindowAXC.getPropertyAsComponent(\"Panes\");");
		ActiveXComponent jacobActiveWindowPanesAXC = jacobActiveWindowAXC.getPropertyAsComponent("Panes");

		try {			
		System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): jacobActiveWindowViewAXC.getProperty(\"SplitSpecial\") == "+jacobActiveWindowViewAXC.getProperty("SplitSpecial"));
		System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): jacobActiveWindowViewAXC.getProperty(\"SplitSpecial\").toInt() == "+jacobActiveWindowViewAXC.getProperty("SplitSpecial").toInt());
		if (jacobActiveWindowViewAXC.getProperty("SplitSpecial").toInt() != 0) {
		
			System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to Variant jacobActiveWindowPane2Variant = Dispatch.call(jacobActiveWindowPanesAXC, \"Item\", new Variant(2));");
			Variant jacobActiveWindowPane2Variant = Dispatch.call(jacobActiveWindowPanesAXC, "Item", new Variant(2));
			System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to Dispatch.call(jacobActiveWindowPane2Variant.toDispatch(), \"Close\");");
            Dispatch.call(jacobActiveWindowPane2Variant.toDispatch(), "Close");
		}
		
		} catch (Exception ex2) {
			ExHandler.handle(ex2);
			System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): Fehler bei ActiveWindow.Panes(2).Close...");
		}

		//System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to ActiveXComponent jacobActiveWindowActivePaneAXC = jacobActiveWindowAXC.getPropertyAsComponent(\"ActivePane\");");
		//ActiveXComponent jacobActiveWindowActivePaneAXC = jacobActiveWindowAXC.getPropertyAsComponent("ActivePane");
		
		//System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to ActiveXComponent jacobActiveWindowActivePaneViewAXC = jacobActiveWindowActivePaneAXC.getPropertyAsComponent(\"View\");");
		//ActiveXComponent jacobActiveWindowActivePaneViewAXC = jacobActiveWindowActivePaneAXC.getPropertyAsComponent("View");

		
		
		//This returns 1 when the normal view is active, with the Headers in the lower portion of the window and the main document text in the upper portion of the window.
		//This returns 3 when the page layout view is active
		System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): jacobActiveWindowViewAXC.getPropertyAsComponent(\"Type\") == "+jacobActiveWindowViewAXC.getPropertyAsInt("Type"));
		//System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): jacobActiveWindowActivePaneViewAXC.getPropertyAsComponent(\"Type\") == "+jacobActiveWindowActivePaneViewAXC.getPropertyAsInt("Type"));			
		
		//THIS ALL SHOULD WORK BUT DOESN'T. IT works INTERACTIVELY, THOUGH! (truly Microsoft...) 
		
		try {
			//THIS FAILS:
			//com.jacob.com.ComFailException: Invoke of: Type
			//Source: Microsoft Word
			//Description: Diese Eigenschaft oder Methode ist auf diesem System nicht verfügbar.
			//jacobActiveWindowActivePaneViewAXC.setProperty("Type",3);
			
			//THIS SUCCEEDS (i.e. we're back in Page Layout View afterwards, with the Headers section active, and the cursor in the Headers section.
			//
			//ONLY IF WE HAD ALLOCATED ActiveXComponent jacobActiveWindowActivePaneAXC = ... jacobActiveWindowActivePaneViewAXC = ... above, this THROWS AN EXCEPTION, however NOT RED, but black, probably informative:
			//com.jacob.com.ComFailException: Invoke of: Type
			//Source: Microsoft Word
			//Description: Objekt wurde gelöscht.
			//
			//So I don't allocate the components ActivePane and below, and it all works without any error :-)
			System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to jacobActiveWindowViewAXC.setProperty(\"Type\",3);");
			jacobActiveWindowViewAXC.setProperty("Type",3);
		} catch (Exception ex2) {
			ExHandler.handle(ex2);
			System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): Fehler bei Setzen von View.Type=3...");
		}
		
		
		System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): jacobActiveWindowViewAXC.getPropertyAsComponent(\"Type\") == "+jacobActiveWindowViewAXC.getPropertyAsInt("Type"));
		//System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): jacobActiveWindowActivePaneViewAXC.getPropertyAsComponent(\"Type\") == "+jacobActiveWindowActivePaneViewAXC.getPropertyAsInt("Type"));

		
		//JETZT ist Word zwar wieder im Page Layout View (bzw.: wdPrintView), aber immer noch sind Kopf-/Fusszeilen aktiv und der Cursor im Kopfzeilenbereich.
		//Weiteres Suchen-/Ersetzen findet dann ebenfalls nur dort statt.
		
		//N.B.: Must be in Page Layout View = wdPrintView in order to change SeekView Setting below.
		
		//Ein aufgezeichnetes Makro, welches in dieser Situation Menü: Ansicht - Kopf-/Fusszeilen (aus) entspricht, enthält:
		/*
		Sub Makro4()
		'
		' Makro4 Makro
		' Makro aufgezeichnet am 28.09.2016 von Jörg M. Sigle
		'
		    If ActiveWindow.View.SplitSpecial <> wdPaneNone Then
		        ActiveWindow.Panes(2).Close
		    End If
		    If ActiveWindow.ActivePane.View.Type = wdNormalView Or ActiveWindow. _
		        ActivePane.View.Type = wdOutlineView Then
		        ActiveWindow.ActivePane.View.Type = wdPrintView
		    End If
		    ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
		End Sub
		*/
		
		/*
		//Das bringt mich zwar wieder an den Haupttext-Anfang zurück.
		//Allerdings ist die Kopfzeile auch in der Layout-Ansicht immer noch oben angezeigt, auch wenn sie leer ist,
		//mit einem New-Paragraph-Zeichen drin, und y-Platzverbrauch.
		try {
			System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to jacobActiveWindowViewAXC.setProperty(\"SeekView\",0);");
			jacobActiveWindowViewAXC.setProperty("SeekView",0);
		} catch (Exception ex) {
			ExHandler.handle(ex);
			System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): Fehler bei Setzen von View.SeekView=0...");
		}

		//HILFT NICHTS
		//Dispatch.call(jacobSelection, "MoveRight");
		//Dispatch.call(jacobSelection, "MoveLeft");

		//Das bringt mich in die Kopfzeile.
		try {
			System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to jacobActiveWindowViewAXC.setProperty(\"SeekView\",1);");
			jacobActiveWindowViewAXC.setProperty("SeekView",1);
		} catch (Exception ex) {
			ExHandler.handle(ex);
			System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): Fehler bei Setzen von View.SeekView=1...");
		}

		//Thread.sleep(200);

        //Das bringt mich in die Kopfzeile.
		try {
			System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to jacobActiveWindowViewAXC.setProperty(\"SeekView\",0);");
			jacobActiveWindowViewAXC.setProperty("SeekView",0);
		} catch (Exception ex) {
			ExHandler.handle(ex);
			System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): Fehler bei Setzen von View.SeekView=0...");
		}
		
        //Thread.sleep(200);
        
		//Das bringt mich in die Kopfzeile.
		try {
			System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to jacobActiveWindowViewAXC.setProperty(\"SeekView\",1);");
			jacobActiveWindowViewAXC.setProperty("SeekView",1);
		} catch (Exception ex) {
			ExHandler.handle(ex);
			System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): Fehler bei Setzen von View.SeekView=1...");
		}

		//Thread.sleep(200);

		/* Mal im Word VBA kurz nachgesehen, ob die Konstanten stimmen, und gefunden: */
		int wdSeekCurrentPageHeader=9;
		int wdSeekMainDocument=0;
		int wdNormalView=1;
		int wdOutlineView=2;
		int wdPrintView=3;
		
		//Das resultierende Verhalten ist aber genau gleich wie bei 1 und 0.
		
		//Das bringt mich in die Kopfzeile.
		try {
			if (debugSysPrnFindOrReplaceDetails)
				System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to jacobActiveWindowViewAXC.setProperty(\"SeekView\",wdSeekCurrentPageHeader);  i.e. = 9");
			jacobActiveWindowViewAXC.setProperty("SeekView",wdSeekCurrentPageHeader);
		} catch (Exception ex2) {
			ExHandler.handle(ex2);
			System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): Fehler bei Setzen von View.SeekView = wdSeekCurrentPageHeader = 9...");
		}

		//Thread.sleep(200);

		//Das bringt mich zurück ins Hauptdokument. (Und funktioniert offenbar sogar, wenn der Code wirklich ausgeführt wird...)
		try {
			if (debugSysPrnFindOrReplaceDetails)
				System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to jacobActiveWindowViewAXC.setProperty(\"SeekView\",wdSeekMainDocument);  i.e. = 0");
			jacobActiveWindowViewAXC.setProperty("SeekView",wdSeekMainDocument);
		} catch (Exception ex2) {
			ExHandler.handle(ex2);
			System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): Fehler bei Setzen von View.SeekView = wdSeekMainDocument = 0...");
		}

		//Thread.sleep(200);
		
		/*
		Dispatch jacobSections = Dispatch.get((Dispatch) jacobDocument, "Sections").toDispatch();
        Variant jacobSectionVariant = Dispatch.call(jacobSections, "Item", new Variant(1));
        Dispatch jacobSection = jacobSectionVariant.toDispatch();
        Dispatch jacobSectionHeaders = Dispatch.get((Dispatch) jacobSection, "Headers").toDispatch();		               	           
        Variant jacobSectionHeaderVariant = Dispatch.call(jacobSectionHeaders, "Item", new Variant(1)); //Sorry, das liefert tatsächlich letzendlich Zugang zum Haupttext
        Dispatch jacobSectionHeader = jacobSectionHeaderVariant.toDispatch();
        Dispatch jacobSectionHeaderRange = Dispatch.call(jacobSectionHeader, "Range").toDispatch();
        Dispatch.call(jacobSectionHeaderRange, "Delete");
        */

		//Das Folgende hilft auch alles nichts...
		/*
		jacobSectionVariant.safeRelease();
		jacobSection.safeRelease();
		jacobSectionHeaderVariant.safeRelease();
        jacobSectionHeader.safeRelease();
        jacobSectionHeaderRange.safeRelease();
		
		jacobSectionVariant=null;
		jacobSection=null;
		jacobSectionHeaderVariant=null;
        jacobSectionHeader=null;
        jacobSectionHeaderRange=null;
        */
		
		//HILFT NICHTS
		//Dispatch.call(jacobSelection, "MoveRight");
		//Dispatch.call(jacobSelection, "MoveLeft");

		//N.B.: Bei Auswahl von 2 kommt:
		//MSWord_jsText: findOrReplace (SectionHeaders): Fehler bei Setzen von View.SeekView=2...
		//		com.jacob.com.ComFailException: Invoke of: SeekView
		//		Source: Microsoft Word
		//		Description: Die angeforderte Ansicht ist nicht verfügbar.
		//try {
		//	System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to jacobActiveWindowViewAXC.setProperty(\"SeekView\",2);");
		//	jacobActiveWindowViewAXC.setProperty("SeekView",2);
		//} catch (Exception ex) {
		//	ExHandler.handle(ex);
		//	System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): Fehler bei Setzen von View.SeekView=2...");
		//}
		
		//SplitSpecial ist nun 0, obwohl die leere Kopfzeile noch über dem Haupttext angezeigt wird.
		//Witzigerweise: Wenn ich manuell im Menü Ansicht - Kopf-und-Fusszeile und nochmal Ansicht - Kopf-und-Fusszeile wähle - ist sie verschwunden...
        
        /*
		try {			
		System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): jacobActiveWindowViewAXC.getProperty(\"SplitSpecial\") == "+jacobActiveWindowViewAXC.getProperty("SplitSpecial"));
		if (jacobActiveWindowViewAXC.getProperty("SplitSpecial").toInt() != 0) {
		
			System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to Variant jacobActiveWindowPane2Variant = Dispatch.call(jacobActiveWindowPanesAXC, \"Item\", new Variant(2));");
			Variant jacobActiveWindowPane2Variant = Dispatch.call(jacobActiveWindowPanesAXC, "Item", new Variant(2));
			System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to Dispatch.call(jacobActiveWindowPane2Variant.toDispatch(), \"Close\");");
            Dispatch.call(jacobActiveWindowPane2Variant.toDispatch(), "Close");
		}
		
		} catch (Exception ex) {
			ExHandler.handle(ex);
			System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): Fehler bei ActiveWindow.Panes(2).Close...");
		}
		*/

		/*
		//Also brauche ich DANACH NOCMALS die Umschaltung zu Page Layout View:
		//N.B.: Ich hab auch versucht, die SeekView = 0 Umschaltung oben VOR das erstmalige Type = 3 zu setzen. Wirft eine Exception und funktioniert gar nicht.
		try {
			System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): About to jacobActiveWindowViewAXC.setProperty(\"Type\",3);");
			jacobActiveWindowViewAXC.setProperty("Type",3);
		} catch (Exception ex) {
			ExHandler.handle(ex);
			System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): Fehler bei Setzen von View.Type=3...");
		}
		*/
		
		//YEP. JETZT sind wir wieder vollständig zurück... :-)
		//DAS Funktioniert jetzt auch wirklich gut - incl. des Nicht-Berührens des Header-Bereichs der ersten Brief-Seite (ohne existierende Kopfzeile) :-)
		
		} //if ( !ProbablyUsingTarmed_xxTemplateSoDoNOTAccessHeaderRangeToAvoidGenerationOfEmptyHeaderLines )
		else {
			System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): INFO: Completely skipping processing of Headers (=Kopfzeilen) because this might be a Tarmedrechnung_xx template.");			
			System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): INFO: We do NOT want to access ...Section.Header.Range there, to avoid making empty header lines appear with a paragraph mark,");
			System.out.println("MSWord_jsText: findOrReplace (SectionHeaders): INFO: because this would cause the [Titel] line to go down below the boxes and destroy the complete layout of these templates.");
		}
		
		//ToDo: Lokales umschalten der debugSysPrnFindOrReplaceDetails nach Debugging von findOrReplace (SectionHeaders) vorher und nacher wieder entfernen.	//201701120109js
		//Siehe debugSysPrnFindOrReplaceDetails=true/=false; vor bzw. nach dem findOrReplace (SectionHeaders) Teil.												//201701120109js
		System.out.println("MSWord_jsText: debugSysPrnFindOrReplaceDetails=false WIEDER AUSGESCHALTET nach Debugging von findOrReplace(SectionHeaders);");
		debugSysPrnFindOrReplaceDetails=false;																													//201701120109js
		
		
		
		
		
		
		
		
		
		
		System.out.println("MSWord_jsText: findOrReplace: Final numberOfHits="+numberOfHits);
		
		// DIE FOLGENDE FEHLERMELDUNG BLENDE ICH KOMPLETT AUS, WEIL SIE SONST Z.B. BEI Kontakte -> Adressliste Drucken auftaucht,
		// so lange dort die Elemente im Titel nicht ersetzt werden - OBWOHL später [LISTE] ersetzt wird - was allerdings NICHT
		// via insertOrReplace und Suche nach Platzhalter-mit-Wildcards gefunden und ersetzt wird, sondern offenbar via insertText
		// vermutlich mit direkt aus dem Programmcode veranlasstem Suchen/Ersetzen erledigt wird. Die Platzhalter der ersten Suche
		// mit Wildcards brauchen ja auf jeden Fall einen Punkt, beim zweiten bin ich gerade unsicher, alle anderen brauchen aber
		// definitiv mindestens irgendeinen Punkt oder Doppelpunkt drin wenn ich gerade nicht irre... 
		
		//Laut log output wird die Methode von Elexis mit folgenden patterns aufgerufen, und zwar in aufeinanderfolgenden Aufrufen:
		//pattern=\[[*]?[-a-zA-ZäöüÄÖÜéàè_ ]+\.[-a-zA-Z0-9äöüÄÖÜéàè_ ]+\]
		//pattern=\[[*]?[-a-zA-ZäöüÄÖÜéàè_ ]+(\.[-a-zA-Z0-9äöüÄÖÜéàè_ ]+)+\]
		//pattern=\[[*]?[a-zA-Z]+:mwn?:[^\[]+\]
		//pattern=\[[*]?[-_a-zA-Z0-9]+:[-a-zA-Z0-9]+:[-a-zA-Z0-9\.]+:[-a-zA-Z0-9\.]:?[^\]]*\]
		//pattern=\[[*]?SQL[^:]*:[^\[]+\]
		//pattern=\[SCRIPT:[^\[]+\]

		/*
		//NUR beim ersten Pass (mit dem ersten Pattern) wird eine Fehlermeldung ausgegeben, wenn keine Hits erzielt wurden.
		//Denn das sollte typischerweise irgendeinen Hit liefern. Alle anderen Patterns müssen nicht unbedingt verwendet worden sein. 
		if ((numberOfHits == 0) && (pattern.equals("\\[[*]?[-a-zA-ZäöüÄÖÜéàè_ ]+\\.[-a-zA-Z0-9äöüÄÖÜéàè_ ]+\\]")))	{
			//ToDo: Add precautions for pattern==null or pattern2==null...
			System.out.println("MSWord_jsText: findOrReplace:\nFehler beim Ersetzen:\nKein Treffer für:\n"+
					"pattern \t(Von Elexis für OpenOffice):\t"+pattern+"\n"+
					"pattern2\t(Von msword_js für MS Word):\t"+pattern2+"\n"+		
					"Möglicherweise fehlt eine Dokument- oder Rechnungsvorlage, oder diese enthält keine ersetzbaren Platzhalter,\n"+
					"oder das verwendete Such-Pattern ist nicht mit Word kompatibel?\nIm letzteren Fall muss in MSWord_jsText.java "+
					"eine on-the-Fly-Ersetzung ergänzt werden.");
			//SWTHelper.showError("No doc in bill", "Fehler:","Es ist keine Rechnungsvorlage definiert");
			//ToDo: Add precautions for pattern==null or pattern2==null...
			SWTHelper.showError(
					"MSWord_jsText: findOrReplace:"+ 
					"Fehler:",
					"Fehler beim Suchen und Ersetzen:\nKein Treffer für:\n"+
					"pattern \t(Von Elexis für OpenOffice):    \t"+pattern+"\n"+	//spaces needed for tab alignment in proportional font
					"pattern2\t(Von msword_js für MS Word):\t"+pattern2+"\n"+		
					"\nMöglicherweise fehlt eine Dokument- oder Rechnungsvorlage, oder diese enthält keine ersetzbaren Platzhalter,\n"+
					"oder das verwendete Such-Pattern ist nicht mit Word kompatibel?\nIm letzteren Fall muss in MSWord_jsText.java "+
					"eine on-the-Fly-Ersetzung ergänzt werden.");
		
			System.out.println("MSWord_jsText: findOrReplace: about to end, returning false...");
			return false;
		}
		else
		*/
		
		//After all replacements are done (and Switches too and back from the Header section), we set the WordObj to visible again...
		//It still flickers, because we have about 6 calls through findOrReplace() from Elexis.
		//So... we evaluate the patterns supplied from Elexis to decide whether we shall set Word to invisible or not :-)
		//This is the 6th pattern that Elexis sends, so we're after the 6th pass through findOrReplace and switch Word back to visible :-)
		if (pattern.equals("\\[SCRIPT:[^\\[]+\\]")) {
			//ToDo: Allenfalls wieder einschalten, wenn danach das MSWord Dokumentenwindow wieder nach vorne geholt werden kann: //20170201js commented out
			System.out.println("MSWord_jsText: findOrReplace: FOLGENDER CODE COMMENTED OUT, DA WORD DABEI LEICHT HINTER ELEXIS RUTSCHT:");		
			System.out.println("MSWord_jsText: findOrReplace: BITTE ERST DANN WIEDER EINFUEGEN, WENN WORD ANSCHLmusterIESSEND NACH VORNE GEHOLT WERDEN KANN.");		
			System.out.println("MSWord_jsText: findOrReplace: SIEHE AUCH KORRESPONDIERENDEN CODE WEITER OBEN, mit Visible/Variant(false).");		
			System.out.println("MSWord_jsText: findOrReplace: COMMENTED OUT: About to jacobObjWord.setProperty(\"Visible\", new Variant(true));");		
			//jacobObjWord.setProperty("Visible", new Variant(true));	//20170201js commented out
		}
		
		
		//Especially after hiding the document, it needs to be brought to the front

		//Ich probiere jetzt doch nochmals, Word zu aktivieren und nach vorne zu bringen...
		
		//FollowUp: "objWord.Activate" is recommended, and said to be unreliable especially if other Word Document Windows are open at the same time.
		//Anyway, it works on my laptop, but does hardly ever work at any of Jürg's PCs - why I don't know.

		if (jacobObjWord != null) {	//Das ist ziemlich unwahrscheinlich.												//201701151805js
			System.out.println("MSWord_jsText: findOrReplace: About to jacobObjWord.setProperty(\"Visible\", new Variant(true));");	//201701151805js
			jacobObjWord.setProperty("Visible", new Variant(true));														//201701151805js
			
			//Reicht das vielleicht aus, zusätzlich zum jacobDocument auch noch das jacobObjWord zu aktivieren???
			System.out.println("MSWord_jsText: findOrReplace: About to Dispatch.call(jacobObjWord, \"Activate\");");	//201701151805js			
			Dispatch.call(jacobObjWord, "Activate");																	//201701151805js
		} else {																										//201701151805js
			System.out.println("MSWord_jsText: findOrReplace: WARNING: jacobObjWord IS NULL, I'm not trying to jacobObjWord.setProperty(\"Visible\", new Variant(true);");	//201701151805js								
		}																												//201701151805js
		if (jacobDocument != null) {																					//201701151805js
			//Alter Kommentar: Das funktioniert so nicht direkt, möglicherweise ist das jacobDocument hier nicht definiert?	//201701151805js
			System.out.println("MSWord_jsText: findOrReplace: About to Dispatch.call(jacobDocument, \"Activate\");");	//201701151805js			
			Dispatch.call(jacobDocument, "Activate");																	//201701151805js
			
			/*
			Dispatch jDWinState = null;
			jacobDocument.get(jDWinState, "Windows.Application");
			Variant jDWinstateVariant = get(jDWinState, "wdWindowState");
			jDWinState.setProperty("wdWindowState", new Variant("wdWindowStateMaximize"));
			"Windows.Application","wdWindowState", new Variant("wdWindowStateMaximize"));
			 */

			//HMM. Wie kann ich dieses wdWindowState erreichen? (um es z.B. zu minimieren und dann zu restoren?)
			//Die folgende Sequenz ist von weiter oben übernommen, so entferne ich den Debug-Output hier;
			//Aber ich komme nicht bis zu wdWindowState durch.
			
			//ActiveXComponent jacobDocumentAXC = jacobSelection.getPropertyAsComponent("Document");

			//com.jacob.com.ComFailException: Can't map name to dispid: wdWindowState
			//ActiveXComponent jacobActiveWindowAXC = jacobObjWord.getPropertyAsComponent("ActiveWindow");
			//jacobActiveWindowAXC.setProperty("wdWindowState", "wdWindowStateMaximize");

			//ActiveXComponent jacobActiveWindowViewAXC = jacobActiveWindowAXC.getPropertyAsComponent("View");
			//Fehler: com.jacob.com.ComFailException: Can't map name to dispid: wdWindowState
			//jacobActiveWindowViewAXC.setProperty("wdWindowState", "wdWindowStateMaximize");
			
		
	
		} else {																																				//201701151805js
			System.out.println("MSWord_jsText: findOrReplace: WARNING: jacobDocument IS NULL, I'm not trying to Dispatch.call(jacobDocument,\"Activate\");");	//201701151805js					
		}																																						//201701151805js

		//
		//Das hier wirft gleich eine Fehlermeldung:
		//--------------Exception--------------
		//com.jacob.com.ComFailException: Invoke of: Activate
		//Source: Microsoft Word
		//Description: Anwendung kann nicht aktiviert werden.
		//
		//System.out.println("MSWord_jsText: findOrReplace: About to Dispatch.call(jacobObjWord, \"Activate\");");		
		//Dispatch.call(jacobObjWord, "Activate");
		
		{
			System.out.println("MSWord_jsText: findOrReplace: About to end, returning true...\n");
			return true;
		}
	}
	
	
	
	
	public boolean isInteger(String input)	{  
		System.out.println("MSWord_jsText: isInteger begins");
		try	{  
			Integer.parseInt(input);
			System.out.println("MSWord_jsText: isInteger about to return true...");
			return true;  
		} catch(Exception e) {
			System.out.println("MSWord_jsText: isInteger about to return false...");
			return false;
			}
		}  
	
	
	
	/** retrieves the type of a form component.
	*/
	static public int getFormComponentType(XPropertySet xComponent)	{
		System.out.println("MSWord_jsText: getFormComponentType begins");
		
		System.out.println("MSWord_jsText: getFormComponentType: TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO");
		System.out.println("MSWord_jsText: getFormComponentType: ToDo: Adopt to msword_js...");
		System.out.println("MSWord_jsText: getFormComponentType: ToDo: This would probably only be used by a closer adoption of the original findOrReplace() above.");
		System.out.println("MSWord_jsText: getFormComponentType: TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO");
		
		XPropertySetInfo xPSI = null;
	    if (null != xComponent)
	        xPSI = xComponent.getPropertySetInfo();
	    
	    if ((null != xPSI) && xPSI.hasPropertyByName("ClassId")) {
	        // get the ClassId property
	        XPropertySet xCompProps = (XPropertySet)UnoRuntime.queryInterface(XPropertySet.class, xComponent);
			try {
				System.out.println("MSWord_jsText: getFormComponentType tries to return (Short)xCompProps.getPropertyValue(\"ClassId\");");
			    return (Short)xCompProps.getPropertyValue("ClassId");
			} catch (UnknownPropertyException e) {
				e.printStackTrace();
			} catch (WrappedTargetException e) {
				e.printStackTrace();
			}
	     }
		System.out.println("MSWord_jsText: getFormComponentType ends, about to return 0");
	    return 0;
		}	
	
	
	
	
	
	public PageFormat getFormat(){
		System.out.println("MSWord_jsText: getFormat begins and will return ITextPlugin.PageFormat.USER");
		return ITextPlugin.PageFormat.USER;
	}
	
	
	
	
	
	public String getMimeType(){
		System.out.println("MSWord_jsText: getMimeType begins and will return MIMETYPE_OO2");
		System.out.println("MSWord_jsText: with MIMETYPE_OO2="+MIMETYPE_OO2);
		return MIMETYPE_OO2;
	}
	
	
	
	
	
	/**
	 * Insert a table.
	 * 
	 * @param place
	 *            A string to search for and replace with the table
	 * @param properties
	 *            properties for the table
	 * @param contents
	 *            An Array of String[]s describing each line of the table
	 * @param columnsizes
	 *            int-array describing the relative width of each column (all columns together are
	 *            taken as 100%). May be null, in that case the columns will bhe spread evenly
	 */
	public boolean insertTable(final String place, final int properties, final String[][] contents,
		final int[] columnSizes){
		System.out.println("MSWord_jsText: insertTable begins");

		System.out.println("");
		System.out.println("MSWord_jsText: insertTable: TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO");
		System.out.println("MSWord_jsText: insertTable: Info: We only search for table placeholders in the main text of the document!");
		System.out.println("MSWord_jsText: insertTable: ToDo: ONLY IF NEEDED, add code to do the same thing within all word Shapes etc. - see above in findOrReplace().");
		System.out.println("MSWord_jsText: insertTable: TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO");
		System.out.println("");
		
		if (place == null)	{ System.out.println("MSWord_jsText: insertTable: ERROR: place IS NULL! - returning early, returning false"); return false; }
		else 				  System.out.println("MSWord_jsText: insertTable: place="+place);
		System.out.println("");
		
		int offset = 0;
		if ((properties & ITextPlugin.FIRST_ROW_IS_HEADER) == 0) {
			offset = 1;
		}
		
		/* 
		LINES FROM THE NOA / OpenOffife Implementation, completely replaced by JaCoB based code, CAN BE REMOVED:

		SearchDescriptor search = new SearchDescriptor(place);
		search.setIsCaseSensitive(true);

		ISearchResult searchResult = agIonDoc.getSearchService().findFirst(search);
		if (!searchResult.isEmpty()) {
			ITextRange r = searchResult.getTextRanges()[0];
			
			try {
				ITextTable textTable =
					agIonDoc.getTextTableService().constructTextTable(contents.length + offset,
						contents[0].length);
				agIonDoc.getTextService().getTextContentService().insertTextContent(r, textTable);
				r.setText("");
				ITextTablePropertyStore props = textTable.getPropertyStore();
				long w = props.getWidth();
				long percent = w / 100;
				for (int row = 0; row < contents.length; row++) {
					String[] zeile = contents[row];
					for (int col = 0; col < zeile.length; col++) {
						textTable.getCell(col, row + offset).getTextService().getText().setText(
							zeile[col]);
					}
				}
				if (columnSizes == null) {
					textTable.spreadColumnsEvenly();
				} else {
					for (int col = 0; col < contents[0].length; col++) {
						textTable.getColumn(col).setWidth((short) (columnSizes[col] * percent));
					}
					
				}
				
				System.out.println("MSWord_jsText: insertTable ends, about returning true");
				return true;
			} catch (Exception ex) {
				ExHandler.handle(ex);
			}
		}
		System.out.println("MSWord_jsText: insertTable ends, about returning false");
		return false;
		
		*/

		//Adopted the following search-replace related code from findOrReplace() above, left out comments and r&d attempts...
		
		//insertTable only searches for ONE FIRST occurence of the "place" string, and only replaces it ONCE (i.e. no looping over the whole document):
		
		Integer jacobSearchResultInt = 0;
		Integer numberOfHits = 0;

		ActiveXComponent jacobSelection = jacobObjWord.getPropertyAsComponent("Selection");	    
		ActiveXComponent jacobFind = jacobSelection.getPropertyAsComponent("Find");
		
		System.out.println("MSWord_jsText: insertTable: About to Dispatch.call(jacobSelection, \"HomeKey\", \"wdStory\", \"wdMove\");");
		Dispatch.call(jacobSelection, "HomeKey", new Variant(6));

		try {
			jacobFind.setProperty("Text", place);
			jacobFind.setProperty("Forward", "True");
			jacobFind.setProperty("Format", "False");
			jacobFind.setProperty("MatchCase", "True");		      // <- !!! There are differences between findOrReplace() and insertTable() !!!
			jacobFind.setProperty("MatchWholeWord", "False");
			jacobFind.setProperty("MatchByte", "False");
			jacobFind.setProperty("MatchAllWordForms", "False");
			jacobFind.setProperty("MatchSoundsLike", "False");
			jacobFind.setProperty("MatchWildcards", "False");     // <- !!! There are differences between insertTable() and insertTable() !!!
		} catch (Exception ex) {
			ExHandler.handle(ex);
			System.out.println("MSWord_jsText: insertTable: Fehler bei jacobSelection.setProperty(\"...\", \"...\");");
			
			//SWTHelper.showError("No doc in bill", "Fehler:","Es ist keine Rechnungsvorlage definiert");
			SWTHelper.showError("insertTable:", "Fehler:","insertTable: Fehler bei bei jacobSelection.setProperty(\"...\", \"...\");");
		}
		
		try {	 
				jacobSearchResultInt = 0;			//Reset this to 0 in each loop iteration, so that even when jacobFind.invoke() should NOT put a valid result into this variable
													//(might this happen? if an exception was throuwn?), we do definitely NOT get an endless loop, NOR a misguided attempt to replace text. 	
				try { 
					System.out.println("MSWord_jsText: insertTable: About to jacobFind.invoke(\"Execute\");");
					jacobSearchResultInt = jacobFind.invoke("Execute").toInt();

					//Please note: Wenn ich erste jacobSearchResultInt = jacobSearchresultVariant.toInt() verwende, ist nacher auch der string = ""-1", sonst "true"" +
					System.out.println("MSWord_jsText: insertTable: jacobSearchResultInt="+jacobSearchResultInt);
					
				} catch (Exception ex) {
					ExHandler.handle(ex);
					//ToDo: Add precautions for pattern==null or pattern2==null...
					System.out.println("MSWord_jsText: insertTable (Haupttext):\nException caught.\n"+
					"Das verwendete Suchpattern funktioniert möglicherweise NICHT mit MS Word.\n"+
					"place:\t"+place+"\n"+		
					"Falls neue Suchpatterns hinzugefügt wurden, muss möglicherweise eine on-the-fly-Konversion in com.jsigle.msword_js.MSWord_jsText.java ergänzt werden.");
					//ToDo: Add precautions for pattern==null or pattern2==null...
					SWTHelper.showError(
							"MSWord_jsText: insertTable (Haupttext):", 
							"Fehler:",
							"Exception caught.\n"+
							"Das verwendete Suchpattern funktioniert möglicherweise NICHT mit MS Word.\n"+
							"place:\t"+place+"\n"+		
							"Falls neue Suchpatterns hinzugefügt wurden, muss möglicherweise eine on-the-fly-Konversion in com.jsigle.msword_js.MSWord_jsText.java ergänzt werden.");
				}
				
			
				if (jacobSearchResultInt == -1) {		//Do ONLY attempt a replacement if there was a search hit. Otherwise, replacement text might be simply inserted at cursor position.					
						numberOfHits += 1;
						System.out.println("MSWord_jsText: insertTable: numberOfHits="+numberOfHits);
																
						//Obtain the found text portion - contents should be equal to "place"
						
						System.out.println("MSWord_jsText: insertTable: About to String orig = jacobSelection.getProperty(\"Text\").toString();");
						String orig = jacobSelection.getProperty("Text").toString();
						if (orig  == null)	System.out.println("MSWord_jsText: insertTable: ERROR: orig IS NULL!");
						else 				System.out.println("MSWord_jsText: insertTable: orig="+orig);

						//Replace the found text portion by "", hopefully maintaining the cursor position in the document
						System.out.println("MSWord_jsText: insertTable: About to jacobSelection.setProperty(\"Text\", \".InsertTableGOESHERE.\");");
						jacobSelection.setProperty("Text", ".InsertTableGOESHERE.");
						
						//First, compute required numbers of rows and columns - as not all lines need to have the same length in Java, step through all of them...
						
						System.out.println("MSWord_jsText: insertTable: Computing table size...");
						int iTableRowCount = contents.length;	
						int iTableColCount = 0;
						for (int row = 0; row < iTableRowCount; row++) {
							if (contents[row].length > iTableColCount) { iTableColCount = contents[row].length; }  
						}
						iTableRowCount= iTableRowCount + offset; //offset: 1 row reserved for header, if ((properties & ITextPlugin.FIRST_ROW_IS_HEADER) == 0) {...}, see above
						System.out.println("MSWord_jsText: insertTable: counted iRowCount="+iTableRowCount+" including offset="+offset+" for TableHeader");
						System.out.println("MSWord_jsText: insertTable: counted iColCount="+iTableColCount);
						
						//We can't allocate a table if iRowCount or iColCount is 0!
						//This would fail with:
						//	com.jacob.com.ComFailException: Invoke of: Add
						//	Source: Microsoft Word
						//	Description: Die Zahl muss zwischen 1 und 63 liegen.
						//So we clamp iTableRowCount, iTableColCount to a minimum value of 1, and iTableColCount to a maximum of 63.
						
						if ( iTableRowCount <= 0 ) {
							System.out.println("MSWord_jsText: insertTable: WARNING: Setting iTableRowCount to minimum allowed value of 1!");
							iTableRowCount = 1;}
						if ( iTableColCount <= 0 ) {
							System.out.println("MSWord_jsText: insertTable: WARNING: Setting iTableColCount to minimum allowed value of 1!");
							iTableColCount = 1;}
						if ( iTableColCount > 63 ) {
							System.out.println("MSWord_jsText: insertTable: WARNING: Setting iTableColCount to maximum allowed value of 63!");
							iTableColCount = 63;
							jacobSelection.setProperty("Text", "WARNING: iTableColCount limited to maximum value of 63 columns!\nContent of columns 64ff will not be shown.\n\n");
							}
					
						//Allocate a table of suitable size
						
						System.out.println("MSWord_jsText: insertTable: About to ActiveXComponent jacobTables = jacobSelection.getPropertyAsComponent(\"Tables\");");						
						ActiveXComponent jacobTables = jacobSelection.getPropertyAsComponent("Tables");

						//System.out.println("MSWord_jsText: insertTable: About to ActiveXComponent jacobTables = jacobObjWord.getPropertyAsComponent(\"Tables\");");						
						//ActiveXComponent jacobTables = jacobDocument.getPropertyAsComponent("Tables");

						System.out.println("MSWord_jsText: insertTable: About to Variant jacobSelectionRange = jacobSelection.getProperty(\"Range\");");
					    Variant jacobSelectionRange = jacobSelection.getProperty("Range");
						System.out.println("MSWord_jsText: insertTable: About to ActiveXComponent jacobTable = jacobTables.invokeGetComponent(\"Add\", jacobSelectionRange, new Variant(iRowCount), new Variant(iColCount));");
					    ActiveXComponent jacobTable = jacobTables.invokeGetComponent("Add", jacobSelectionRange, new Variant(iTableRowCount), new Variant(iTableColCount));
						System.out.println("MSWord_jsText: insertTable: About to jacobTable.invoke(\"AutoFormat\", 0);");
					    //jacobTable.invoke("AutoFormat", 16);	//schwarzer Rahmen
					    //jacobTable.invoke("AutoFormat", 1);	//grüner Rahmen
					    jacobTable.invoke("AutoFormat", 0);		//KEIN Rahmen - wie gewünscht
				
					    System.out.println("MSWord_jsText: insertTable: About to Variant jacobTableRange = jacobTable.getProperty(\"Range\");");
					    ActiveXComponent jacobTableRange = jacobTable.getPropertyAsComponent("Range");
					    System.out.println("MSWord_jsText: insertTable: About to Variant jacobCells = jacobTableRange.getPropertyAsComponent(\"Cells\");");
					    ActiveXComponent jacobCells = jacobTableRange.getPropertyAsComponent("Cells");
					    
						for (int row = 0; row < contents.length; row++) {
							String[] zeile = contents[row];
							int tableRow = row+1 + offset;
							for (int col = 0; col < zeile.length; col++) {
								int tableCol = col+1;
								if (( tableRow > 0 ) && ( tableCol > 0 ) && ( tableCol <= 63 )) {
								    System.out.println("MSWord_jsText: insertTable: About to ActiveXComponent jacobCell = jacobTable.invokeGetComponent(\"Cell\", new Variant("+tableRow+"), new Variant("+tableCol+"));");
									ActiveXComponent jacobCell = jacobTable.invokeGetComponent("Cell", new Variant(tableRow), new Variant(tableCol));
								    System.out.println("MSWord_jsText: insertTable: About to XComponent jacobCellRange = jacobCell.getPropertyAsComponent(\"Range\");");
									ActiveXComponent jacobCellRange = jacobCell.getPropertyAsComponent("Range");
									//jacobCellRange.invoke("Delete");
									//jacobCellRange.invoke("InsertAfter", "ContentForColRow("+tableRow+","+tableCol+")");
								    System.out.println("MSWord_jsText: insertTable: About to jacobCellRange.invoke(\"InsertAfter, zeile[col]); with zeile["+col+"]="+zeile[col]);
									jacobCellRange.invoke("InsertAfter",zeile[col]);
								} else {
								    System.out.println("MSWord_jsText: insertTable: WARNING: Table content cannot be output for:\n"+
								    				   "   tableRow="+tableRow+"; tableCol="+tableCol+"; content="+zeile[col]+"\n");
								}
							}
						}

						//ToDo: insertTable in msword_js does not support control of relative column sizes yet, instead uses MS Word AutoFormat feature. 
						System.out.println("");
						System.out.println("MSWord_jsText: insertTable: TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO");
						System.out.println("MSWord_jsText: insertTable: ToDo: THE NOATEXT ORIGINAL CAN SET PROVIDED RELATIVE TABLE COLUMN SIZES, i.e. WIDTHS"); 
						System.out.println("MSWord_jsText: insertTable: ToDo: At the moment, we're simply using Word's AutoFormat feature instead");
						System.out.println("MSWord_jsText: insertTable:       (because I have not yet figured out how to control the Word table format manually,");
						System.out.println("MSWord_jsText: insertTable: 	    nor perfectly understood the existing code from Elexis/NOAText, and other functionality is more important).");
						System.out.println("MSWord_jsText: insertTable: ToDo: There's also: AutoFit, which we might try, too. http://www.cnblogs.com/hzj-/articles/1732567.html");
						System.out.println("MSWord_jsText: insertTable: TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO");
						System.out.println("");

						jacobTable.invoke("AutoFormat", 0);		//KEIN Rahmen - wie gewünscht
					    
						//ToDo: THE NOATEXT ORIGINAL CAN SET PROVIDED RELATIVE TABLE COLUMN SIZES, i.e. WIDTHS
						
/* THE NOATEXT ORIGINAL CAN SET PROVIDED RELATIVE TABLE COLUMN SIZES, i.e. WIDTHS 						
						try {
							ITextTablePropertyStore props = textTable.getPropertyStore();
							long w = props.getWidth();
							long percent = w / 100;
							
							for (int row = 0; row < contents.length; row++) {
								String[] zeile = contents[row];
								for (int col = 0; col < zeile.length; col++) {
									textTable.getCell(col, row + offset).getTextService().getText().setText(zeile[col]);
								}
							}
							
							
							if (columnSizes == null) {
								textTable.spreadColumnsEvenly();
							} else {
								for (int col = 0; col < contents[0].length; col++) {
									textTable.getColumn(col).setWidth((short) (columnSizes[col] * percent));
								}
								
							}
						
*/
					    
/*													
						//GETESTET: DAS MoveRight; MoveLeft; IST WIRKLICH NÖTIG, UM IM HAUPTTEXT ZUVERLÄSSIG ALLE PLATZHALTER ZU ERSETZEN. NICHT NÖTIG IN SHAPES.

						//Move the cursor in the document to the right of the table just inserted
						
						//Moving right removes the highlighting and places the cursor to the right of the replaced text.
						//This is required, as otherwise, successive find/replace occurances may become confused.
						System.out.println("MSWord_jsText: insertTable: About to jacobSelection.invoke(\"MoveRight\");");
						jacobSelection.invoke("MoveRight");
						
						//However, it's also necessary to go back to the left by one step afterwards,
						//or otherwise, a seamlessly following [placeholders][seamlesslyFollowingPlaceholder] will NOT be found.
						//The MoveRight - MoveLeft sequence has the effect that the selection = highlighting is removed from the inserted text.
						System.out.println("MSWord_jsText: insertTable: About to jacobSelection.invoke(\"MoveLeft\");");
						jacobSelection.invoke("MoveLeft");
						
						System.out.println("");
*/
				} // if (jacobSearchResultInt == -1) 				
		} catch (Exception ex) {
			ExHandler.handle(ex);
			//ToDo: Add precautions for pattern==null or pattern2==null...
			System.out.println("MSWord_jsText: insertTable:\nFehler bei insertTable:\n"+"" +
					"Exception caught für:\n"+
					"place:\t"+place+"\n"+		
					"numberOfHits="+numberOfHits);
			//SWTHelper.showError("No doc in bill", "Fehler:","Es ist keine Rechnungsvorlage definiert");
			//ToDo: Add precautions for pattern==null or pattern2==null...
			SWTHelper.showError(
					"MSWord_jsText: insertTable:"+ 
					"Fehler:",
					"MSWord_jsText: insertTable:\nFehler bei insertTable:\n"+
					"Exception caught für:\n"+
					"place:\t"+place+"\n"+		
					"numberOfHits="+numberOfHits);
		}

		if (numberOfHits > 0) {
			System.out.println("MSWord_jsText: insertTable ends, returning true");
			return true;
		} else {
		System.out.println("MSWord_jsText: insertTable ends, returning false");
		return false;
		}		
	}
	
	
	
	
	
	/**
	 * Insert Text and return a cursor describing the position We can not avoid using UNO here,
	 * because NOA does not give us enough control over the text cursor
	 */
	public Object insertText(final String marke, final String text, final int adjust){
		System.out.println("MSWord_jsText: insertText(final String marke, final String text, final int adjust) begins");
		System.out.println("");
		
	//ToDo: msword_js insertText(): This does not support the adjust parameter it accepts.
	//ToDo: msword_js insertText(): Only insertTextAt() actually implements adjust support. Same as in the noatext original, so I'll leave it that way. 

		System.out.println("");
	System.out.println("MSWord_jsText: TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO ");
	System.out.println("MSWord_jsText: insertText() shall return an implementation specific cursor that allows another text insertion after that point.");
	System.out.println("MSWord_jsText: insertText()   Hopefully, returning cur = jacobSelection.GetObject should fulfill this requirement...");
	System.out.println("MSWord_jsText: TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO ");
	System.out.println("");
		
	/* ORIGINAL CODE FROM NOATEXT
		SearchDescriptor search = new SearchDescriptor(marke);
		search.setIsCaseSensitive(true);
		ISearchResult searchResult = agIonDoc.getSearchService().findFirst(search);
		XText myText = agIonDoc.getXTextDocument().getText();
		XTextCursor cur = myText.createTextCursor();
		// ITextCursor cur=doc.getTextService().getCursorService().getTextCursor();
		if (!searchResult.isEmpty()) {
			ITextRange r = searchResult.getTextRanges()[0];
			cur = myText.createTextCursorByRange(r.getXTextRange());
			cur.setString(text);
			try {
				setFormat(cur);
			} catch (Exception e) {
				ExHandler.handle(e);
			}
			
			cur.collapseToEnd();
		}
	ORIGINAL CODE FROM NOATEXT */
		
		if (marke == null)	{ System.out.println("MSWord_jsText: insertText: ERROR: marke IS NULL! - returning early, returning NULL"); return null; }
		else 				  System.out.println("MSWord_jsText: insertText: place="+marke);
		if (text == null)	{ System.out.println("MSWord_jsText: insertText: ERROR: text IS NULL!");}
		else 				  System.out.println("MSWord_jsText: insertText: text="+text);
		System.out.println("MSWord_jsText: insertText: adjust="+adjust+" (unsupported in insertText, only supported in insertTextAt, same as in NoaText)");
		System.out.println("");
		
		Integer jacobSearchResultInt = 0;
		Integer numberOfHits = 0;

		ActiveXComponent jacobSelection = jacobObjWord.getPropertyAsComponent("Selection");
		ActiveXComponent jacobFind = jacobSelection.getPropertyAsComponent("Find");
		
		System.out.println("MSWord_jsText: insertText: About to Dispatch.call(jacobSelection, \"HomeKey\", \"wdStory\", \"wdMove\");");
		Dispatch.call(jacobSelection, "HomeKey", new Variant(6));

		try {
			jacobFind.setProperty("Text", marke);
			jacobFind.setProperty("Forward", "True");
			jacobFind.setProperty("Format", "False");
			jacobFind.setProperty("MatchCase", "True");		      // <- !!! There are differences between findOrReplace() and insertText() !!!
			jacobFind.setProperty("MatchWholeWord", "False");
			jacobFind.setProperty("MatchByte", "False");
			jacobFind.setProperty("MatchAllWordForms", "False");
			jacobFind.setProperty("MatchSoundsLike", "False");
			jacobFind.setProperty("MatchWildcards", "False");     // <- !!! There are differences between insertTable() and insertText() !!!
		} catch (Exception ex) {
			ExHandler.handle(ex);
			System.out.println("MSWord_jsText: insertText: Fehler bei jacobSelection.setProperty(\"...\", \"...\");");
			
			//SWTHelper.showError("No doc in bill", "Fehler:","Es ist keine Rechnungsvorlage definiert");
			SWTHelper.showError("insertText:", "Fehler:","insertText: Fehler bei bei jacobSelection.setProperty(\"...\", \"...\");");
		}
		
		try {	 
				jacobSearchResultInt = 0;			//Reset this to 0 in each loop iteration, so that even when jacobFind.invoke() should NOT put a valid result into this variable
													//(might this happen? if an exception was throuwn?), we do definitely NOT get an endless loop, NOR a misguided attempt to replace text. 	
				try { 
					System.out.println("MSWord_jsText: insertText: About to jacobFind.invoke(\"Execute\");");
					jacobSearchResultInt = jacobFind.invoke("Execute").toInt();

					//Please note: Wenn ich erste jacobSearchResultInt = jacobSearchresultVariant.toInt() verwende, ist nacher auch der string = ""-1", sonst "true"" +
					System.out.println("MSWord_jsText: insertText: jacobSearchResultInt="+jacobSearchResultInt);
					
				} catch (Exception ex) {
					ExHandler.handle(ex);
					//ToDo: Add precautions for pattern==null or pattern2==null...
					System.out.println("MSWord_jsText: insertTable (Haupttext):\nException caught.\n"+
					"Das verwendete Suchpattern funktioniert möglicherweise NICHT mit MS Word.\n"+
					"marke:\t"+marke+"\n"+		
					"Falls neue Suchpatterns hinzugefügt wurden, muss möglicherweise eine on-the-fly-Konversion in com.jsigle.msword_js.MSWord_jsText.java ergänzt werden.");
					//ToDo: Add precautions for pattern==null or pattern2==null...
					SWTHelper.showError(
							"MSWord_jsText: insertTable (Haupttext):", 
							"Fehler:",
							"Exception caught.\n"+
							"Das verwendete Suchpattern funktioniert möglicherweise NICHT mit MS Word.\n"+
							"marke:\t"+marke+"\n"+		
							"Falls neue Suchpatterns hinzugefügt wurden, muss möglicherweise eine on-the-fly-Konversion in com.jsigle.msword_js.MSWord_jsText.java ergänzt werden.");
				}
				
			
				if (jacobSearchResultInt == -1) {		//Do ONLY attempt a replacement if there was a search hit. Otherwise, replacement text might be simply inserted at cursor position.					
						numberOfHits += 1;
						System.out.println("MSWord_jsText: insertText: numberOfHits="+numberOfHits);
																
						//Obtain the found text portion - contents should be equal to "place"
						
						System.out.println("MSWord_jsText: insertText: About to String orig = jacobSelection.getProperty(\"Text\").toString();");
						String orig = jacobSelection.getProperty("Text").toString();
						if (orig  == null)	System.out.println("MSWord_jsText: insertText: ERROR: orig IS NULL!");
						else 				System.out.println("MSWord_jsText: insertText: orig="+orig);

						//Replace the found text portion by "", hopefully maintaining the cursor position in the document
				    	System.out.println("MSWord_jsText: insertText: text == "+text.toString());
						
						//Text replacements adopted from insertAt() (continued) below,
						//which are needed *there* to make MS Word understand incoming text from Elexis correctly
						//so that tarmed bill table appears correctly.
				    	//
						//I'll adopt the same replacements for the initial insertAt() method here, to make both insertAt methods
						//behave homogeneously, though it would probably not be needed right here.
						//
						//For detailed explanations: See below.
				    	
				    	System.out.println("MSWord_jsText: insertText (continued): Replacing 9632 squareBullet by 13 NewParagraph + 9632 suitable for MS Word.");
				    	//String text2 = text.replace( (char) 9632, (char) 13 );
				    	//String text2 = text.replace( "\u9632", "\u0013\u9632" );	//THIS DOES NOT WORK AT ALL.
				    	//String text2 = text.replace( "\u25A0", "\n\u25A0" );		//THIS WORKS.
				    	//String text2 = text.replace( "a", "AA" );					//THIS WORKS.
				    	//OR: Rather try adding a \n after *each* incoming text (due to what happens in the TextFrames in other parts of Tarmed bill creation)
				    	String text2 = text+"\n";
				    	//YEP. The results in the table become correct, and those in the text boxes look reasonable, too, after this change. 

				    	System.out.println("MSWord_jsText: insertText (continued): text2 == "+text2);
				    	if (text2.length()>2) System.out.println("MSWord_jsText: insertText (continued): ord(text2.toString()[0-2]) == "+(int) text2.charAt(0)+" "+(int) text2.charAt(1)+" "+(int) text2.charAt(2)+"... ");

				    	//This might not be perfectly necessary any more after the addition of spaces as implemented a few lines above, but I'll leave it in for now.
				    	System.out.println("MSWord_jsText: insertText (continued): Replacing 32 Space by 160 NonBreakingSpace to keep things like \" + Konsultation,...\" together in one line.");
				    	String text3 = text2.replace((char) 32, (char) 160); 
				    	
						System.out.println("MSWord_jsText: insertText (continued): About to Dispatch.put((Dispatch) pos, \"Text\", text3);");
						
						System.out.println("MSWord_jsText: insertText: About to jacobSelection.setProperty(\"Text\", text3);");						
						jacobSelection.setProperty("Text", text3);
											    
						
						/* ORIGINAL CODE FROM NOATEXT/OPENOFFICE
						com.sun.star.beans.XPropertySet charProps = setFormat(xtc);
						Das folgende ersetzt die entsprechende Prozedur - 
						I'm doing this inline (in three occasions in this file) to avoid all the complications
						Java usually brings up when you just would want to move a little bit of code into a simple procedure.  
						ORIGINAL CODE FROM NOATEXT/OPENOFFICE`*/
						
				    	System.out.println("MSWord_jsText: insertTextAt: About to Dispatch fontDispatch = Dispatch.get(jacobSelection, \"Font\").toDispatch();");
			            Dispatch fontDispatch = Dispatch.get(jacobSelection, "Font").toDispatch();
				    	if ( font != null )		{ 
					    	System.out.println("MSWord_jsText: insertTextAt: About to Dispatch.put(fontDispatch, \"Name\", new Variant(font));");
				    		Dispatch.put(fontDispatch, "Name", new Variant(font));
				    	} else {
				    		System.out.println("MSWord_jsText: insertTextAt: WARNING: font IS NULL.");
				    	}
				        if (hi > 0)				{ 
				        	System.out.println("MSWord_jsText: insertTextAt: Dispatch.put(fontDispatch, \"Size\", new Float(hi));");
				    		//Dispatch.put(fontDispatch, "CharHeight", new Float(hi)); 	//OpenOffice: Height of the character in point
				    		Dispatch.put(fontDispatch, "Size", new Float(hi)); 
				        }
				        if (stil > -1) {
			        		System.out.println("MSWord_jsText: WARNING: The MS Word FONT property does apparently NOT support numeric font weight, so we have fewer steps available. { SWT.MIN = SWT.NORMAL; SWT.BOLD }"); 
				        	switch (stil) {
				        	case SWT.MIN:		{ 
				        		System.out.println("MSWord_jsText: insertTextAt: Mapping: OpenOffice SWT.MIN -> Dispatch.put(fontDispatch, \"Bold\", false);");
				        		//Dispatch.put(fontDispatch, "CharWeight", 15f); break; 
					        	Dispatch.put(fontDispatch, "Bold", false);
				        	}
				        	case SWT.NORMAL:	{ 
					        	System.out.println("MSWord_jsText: insertTextAt: Mapping: OpenOffice SWT.NORMAL -> Dispatch.put(fontDispatch, \"Bold\", false);");
				        		//Dispatch.put(fontDispatch, "CharWeight", FontWeight.LIGHT); break;
					        	Dispatch.put(fontDispatch, "Bold", false);
				        	}
				        	case SWT.BOLD:		{ 
					        	System.out.println("MSWord_jsText: insertTextAt: Mapping: OpenOffice SWT.BOLD -> Dispatch.put(fontDispatch, \"Bold\", true);");
				        		//Dispatch.put(fontDispatch, "CharWeight", FontWeight.BOLD); break;
					        	Dispatch.put(fontDispatch, "Bold", true);
				        	}
					        }
				        }
				        
						//GETESTET: DAS MoveRight; MoveLeft; IST WIRKLICH NÖTIG, UM IM HAUPTTEXT ZUVERLÄSSIG ALLE PLATZHALTER ZU ERSETZEN. NICHT NÖTIG IN SHAPES.
						
						//Move the cursor in the document to the right of the text just inserted
						
						//Moving right removes the highlighting and places the cursor to the right of the replaced text.
						//This is required, as otherwise, successive find/replace occurances may become confused.
						System.out.println("MSWord_jsText: insertText: About to jacobSelection.invoke(\"MoveRight\");");
						jacobSelection.invoke("MoveRight");
						
						//However, it's also necessary to go back to the left by one step afterwards,
						//or otherwise, a seamlessly following [placeholders][seamlesslyFollowingPlaceholder] will NOT be found.
						//The MoveRight - MoveLeft sequence has the effect that the selection = highlighting is removed from the inserted text.
						System.out.println("MSWord_jsText: insertText: About to jacobSelection.invoke(\"MoveLeft\");");
						jacobSelection.invoke("MoveLeft");
						
						System.out.println("");
				} // if (jacobSearchResultInt == -1) 				
		} catch (Exception ex) {
			ExHandler.handle(ex);
			//ToDo: Add precautions for pattern==null or pattern2==null...
			System.out.println("MSWord_jsText: insertText:\nFehler beim Suchen und Ersetzen im Haupttext:\n"+"" +
					"Exception caught für:\n"+
					"marke:\t"+marke+"\n"+		
					"numberOfHits="+numberOfHits);
			//SWTHelper.showError("No doc in bill", "Fehler:","Es ist keine Rechnungsvorlage definiert");
			//ToDo: Add precautions for pattern==null or pattern2==null...
			SWTHelper.showError(
					"MSWord_jsText: insertText:"+ 
					"Fehler:",
					"MSWord_jsText: insertText:\nFehler beim Suchen und Ersetzen im Haupttext:\n"+
					"Exception caught für:\n"+
					"marke:\t"+marke+"\n"+		
					"numberOfHits="+numberOfHits);
		}

		System.out.println("MSWord_jsText: insertText ends, returning jacobSelection (in lieu of cursor cur)");
		
		//cur.collapseToEnd; = sets start position to end (of current selection)
		jacobSelection.invoke("MoveRight");
		jacobSelection.invoke("MoveLeft");		
		
		//This works here with ActiveXComponent jacobSelection...
		//jacobSelection.setProperty("Text", text);
		
		Object cur = jacobSelection.getObject();
		
		//This compiles (and hopefully works) here with Object cur = jacobSelection.getObject();
		//and thus, it is a way to return an Object that contains information equivalent to jacobSelection
		//and hopefully can be used in a subsequent method - to put another text at the same position - like this: 
		//  Dispatch.put((Dispatch) cur, "Text", "Hallo");
		//Please note: Just using jacobSelection.setProperty("Text", text) down there would not compile.
		return cur;
	}
	
	
	
	
	
	/**
	 * Insert text at a position returned by insertText(String,text,adjust)
	 */
	public Object insertText(final Object pos, final String text, final int adjust){
		System.out.println("MSWord_jsText(final Object pos, final String text, final int adjust) begins");
		
	//ToDo: msword_js insertText(): This does not support the adjust parameter it accepts.
	//ToDo: msword_js insertText(): Only insertTextAt() actually implements adjust support. Same as in the noatext original, so I'll leave it that way. 
		
	/* ORIGINAL CODE FROM NOATEXT
				
		XTextCursor cur = (XTextCursor) pos;
		if (cur != null) {
			cur.setString(text);
			try {
				setFormat(cur);
			} catch (Exception e) {
				ExHandler.handle(e);
			}
			cur.collapseToEnd();
		}
	ORIGINAL CODE FROM NOATEXT */
			
		try {	 
			//Put text to the jacobSelection which has been passed as Object pos - hopefully that works...
	    	System.out.println("MSWord_jsText: insertText (continued): text == "+text);
	    	if (text.length()>2) System.out.println("MSWord_jsText: insertText (continued): ord(text.toString()[0-2]) == "+(int) text.charAt(0)+" "+(int) text.charAt(1)+" "+(int) text.charAt(2)+"... ");
	    	
	    	//The following incoming text from Elexis may be not be understood in MS Word as expected,
	    	//and will lead to a desynchronization between the prepared columns and the bill positions going inside.
	    	//Specifically, incoming text lines include a character 9632 for a square bullet at the beginning of the line, but NO newline understood by MS Word.
	    	//
	    	//This means that a tarmed position entry, normally consisting of Line1: cleartext; Line2: figures,
	    	//will have figures following cleartext on THE SAME line, and from there on, everything will look messed up.
	    	//
	    	//From my detailed logs:
	    	//
	    	//MSWord_jsText: insertText (continued): text == 	Konsultation, erste 5 Min. (Grundkonsultation)
	    	//MSWord_jsText: insertText (continued): ord(text.toString()[0-2]) == 9 75 111... 
	    	//
	    	//MSWord_jsText: insertText (continued): text == ■ 22.07.2016	001	00.0010	 	1	 	1.0	9.57	1.0	0.86	8.19		0.86	1	1	0	0	15.27
	    	//MSWord_jsText: insertText (continued): ord(text.toString()[0-2]) == 9632 32 50... 
	    	//
	    	//So I have to replace the 9632 by some newline character that MS Word understands
	    	//(I would prefer: New line, NOT New paragraph because the two lines here belong together -
	    	// but 13 results in New paragraph in Word, and maybe it happens for other constellations
	    	// where lines so separated do NOT belong together (or does 9632 mean newline in OpenOffice?),
	    	// so I'm leaving it in anyway...):

	    	System.out.println("MSWord_jsText: insertText (continued): Replacing 9632 squareBullet by 13 NewParagraph + 9632 suitable for MS Word.");
	    	//String text2 = text.replace( (char) 9632, (char) 13 );	//THIS WORKS.
	    	//String text2 = text.replace( "\u9632", "\u0013\u9632" );	//THIS DOES NOT WORK AT ALL.
	    	//String text2 = text.replace( "\u25A0", "\n\u25A0" );		//THIS WORKS.
	    	//String text2 = text.replace( "a", "AA" );					//THIS WORKS.
	    	//OR: Rather try adding a \n after *each* incoming text (due to what happens in the TextFrames in other parts of Tarmed bill creation)
	    	String text2 = text+"\n";
	    	//YEP. The results in the table become correct, and those in the text boxes look reasonable, too, after this change. 

	    	System.out.println("MSWord_jsText: insertText (continued): text2 == "+text2);
	    	if (text2.length()>2) System.out.println("MSWord_jsText: insertText (continued): ord(text2.toString()[0-2]) == "+(int) text2.charAt(0)+" "+(int) text2.charAt(1)+" "+(int) text2.charAt(2)+"... ");

	    	//This might not be perfectly necessary any more after the addition of spaces as implemented a few lines above, but I'll leave it in for now:

	    	//Now, another problem:
	    	//AFTER the figures for one tarmed bill entry, Elexis simply outputs a TAB (really...)
	    	//before it sends the cleartext line of the NEXT tarmed bill entry.
	    	//That should normalle work - but in my setup, with tarmed bill templates taken over form a previous OpenOffice environment,
	    	//and the apparent fact that MS Word vs. OpenOffice specified font sizes may result in slightly different actual character sizes or word lengths,
	    	//it just so happens, that when the next tarmed entry has a cleartext of " + Konsultation, jede weitere 5 Minuten...":
	    	//the TAB" + " will remain in the upper line (= be appended to the figures of the preceding bill entry),
	    	//and only "Konsultation, jede weitere..." will appear on the following line, actually starting at its beginning.

	    	//I considered, replacing spaces by non-breaking-spaces due to this, but:
	    	//MIGHT THIS ACTUALLY BE A RESULT OF JUSTIFICATION CONTROL NOT BEING CORRECTLY PASSED TO MS WORD YET?
	    	//NO. The TABSTOPs in the tarmed bill template define what's left or right justified completely, and that works.
	    	//And... voila, after the following replacement, the tarmed bill table actually appears in a correctly looking way. :-)

	    	//I'll adopt the same replacements for the initial insertAt() method above, to make them both behave homogeneously, though it would probably not be needed right here.
	    	
	    	//This might not be perfectly necessary any more after the addition of spaces as implemented a few lines above, but I'll leave it in for now.
	    	System.out.println("MSWord_jsText: insertText (continued): Replacing 32 Space by 160 NonBreakingSpace to keep things like \" + Konsultation,...\" together in one line.");
	    	String text3 = text2.replace((char) 32, (char) 160); 
	    	
			System.out.println("MSWord_jsText: insertText (continued): About to Dispatch.put((Dispatch) pos, \"Text\", text3);");
			Dispatch.put((Dispatch) pos, "Text", text3);

			/* ORIGINAL CODE FROM NOATEXT/OPENOFFICE
			com.sun.star.beans.XPropertySet charProps = setFormat(xtc);
			Das folgende ersetzt die entsprechende Prozedur - 
			I'm doing this inline (in three occasions in this file) to avoid all the complications
			Java usually brings up when you just would want to move a little bit of code into a simple procedure.  
			ORIGINAL CODE FROM NOATEXT/OPENOFFICE`*/

			System.out.println("");
			System.out.println("MSWord_jsText: insertText (continued): TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO ");
			System.out.println("MSWord_jsText: insertText (continued): Ist das folgende korrekt?					ActiveXComponent jacobSelection = (ActiveXComponent) pos;");
			System.out.println("MSWord_jsText: insertText (continued): Oder müsste/dürfte auch hier einfach stehen: ActiveXComponent jacobSelection = jacobObjWord.getPropertyAsComponent(\"Selection\");");
			System.out.println("MSWord_jsText: insertText (continued): TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO ");
			System.out.println("");
					
	    	System.out.println("MSWord_jsText: insertTextAt: About to ActiveXComponent jacobSelection = (ActiveXComponent) pos;");
			ActiveXComponent jacobSelection = (ActiveXComponent) pos;
			//ActiveXComponent jacobSelection = jacobObjWord.getPropertyAsComponent("Selection");

	    	System.out.println("MSWord_jsText: insertTextAt: About to Dispatch fontDispatch = Dispatch.get(jacobSelection, \"Font\").toDispatch();");
            Dispatch fontDispatch = Dispatch.get(jacobSelection, "Font").toDispatch();
	    	if ( font != null )		{ 
		    	System.out.println("MSWord_jsText: insertTextAt: About to Dispatch.put(fontDispatch, \"Name\", new Variant(font));");
	    		Dispatch.put(fontDispatch, "Name", new Variant(font));
	    	} else {
	    		System.out.println("MSWord_jsText: insertTextAt: WARNING: font IS NULL.");
	    	}
	        if (hi > 0)				{ 
	        	System.out.println("MSWord_jsText: insertTextAt: Dispatch.put(fontDispatch, \"Size\", new Float(hi));");
	    		//Dispatch.put(fontDispatch, "CharHeight", new Float(hi)); 	//OpenOffice: Height of the character in point
	    		Dispatch.put(fontDispatch, "Size", new Float(hi)); 
	        }
	        if (stil > -1) {
        		System.out.println("MSWord_jsText: WARNING: The MS Word FONT property does apparently NOT support numeric font weight, so we have fewer steps available. { SWT.MIN = SWT.NORMAL; SWT.BOLD }"); 
	        	switch (stil) {
	        	case SWT.MIN:		{ 
	        		System.out.println("MSWord_jsText: insertTextAt: Mapping: OpenOffice SWT.MIN -> Dispatch.put(fontDispatch, \"Bold\", false);");
	        		//Dispatch.put(fontDispatch, "CharWeight", 15f); break; 
		        	Dispatch.put(fontDispatch, "Bold", false);
	        	}
	        	case SWT.NORMAL:	{ 
		        	System.out.println("MSWord_jsText: insertTextAt: Mapping: OpenOffice SWT.NORMAL -> Dispatch.put(fontDispatch, \"Bold\", false);");
	        		//Dispatch.put(fontDispatch, "CharWeight", FontWeight.LIGHT); break;
		        	Dispatch.put(fontDispatch, "Bold", false);
	        	}
	        	case SWT.BOLD:		{ 
		        	System.out.println("MSWord_jsText: insertTextAt: Mapping: OpenOffice SWT.BOLD -> Dispatch.put(fontDispatch, \"Bold\", true);");
	        		//Dispatch.put(fontDispatch, "CharWeight", FontWeight.BOLD); break;
		        	Dispatch.put(fontDispatch, "Bold", true);
	        	}
		        }
	        }
			
			
			//Move the cursor in the document to the right of the text just inserted

			//Moving right removes the highlighting and places the cursor to the right of the replaced text.
			//This is required, as otherwise, successive find/replace occurances may become confused.
			System.out.println("MSWord_jsText: insertText (continued): About to Dispatch.call((Dispatch) pos, \"MoveRight\");");
			//jacobSelection.invoke("MoveRight");
			//Dispatch.call((Dispatch) pos, "Invoke", "MoveRight");
			Dispatch.call((Dispatch) pos, "MoveRight");
			
			//However, it's also necessary to go back to the left by one step afterwards,
			//or otherwise, a seamlessly following [placeholders][seamlesslyFollowingPlaceholder] will NOT be found.
			//The MoveRight - MoveLeft sequence has the effect that the selection = highlighting is removed from the inserted text.
			System.out.println("MSWord_jsText: insertText (continued): About to Dispatch.call((Dispatch) pos, \"MoveLeft\");");
			//jacobSelection.invoke("MoveLeft");
			//Dispatch.call((Dispatch) pos, "Invoke", "MoveLeft");
			Dispatch.call((Dispatch) pos, "MoveLeft");
			
			System.out.println("");			
	} catch (Exception ex) {
		ExHandler.handle(ex);
		//ToDo: Add precautions for pattern==null or pattern2==null...
		System.out.println("MSWord_jsText: insertText:\nFehler bei insertText(pos,...)");
		//ToDo: Add precautions for pattern==null or pattern2==null...
		SWTHelper.showError(
			"MSWord_jsText: insertText:"+ 
			"Fehler:",
			"MSWord_jsText: insertText:\nFehler bei insertText(pos,...)");
	}		
		
	ActiveXComponent jacobSelection = jacobObjWord.getPropertyAsComponent("Selection");

	//cur.collapseToEnd; = sets start position to end (of current selection)
	jacobSelection.invoke("MoveRight");
	jacobSelection.invoke("MoveLeft");		
	
	Object cur = jacobSelection.getObject();
		
	System.out.println("MSWord_jsText: insertText ends, returning cur...");
	return cur;
	}	
	
	
	
	
	
	/**
	 * Insert Text inside a rectangular area. Again we need UNO to get access to a Text frame.
	 */
	public Object insertTextAt(final int x, final int y, final int w, final int h,
		final String text, final int adjust){
		
		System.out.println("MSWord_jsText: insertTextAt begins");		
		
		try {
			/* ORIGINAL CODE FROM NOATEXT/OPENOFFICE
			XTextDocument myDoc = agIonDoc.getXTextDocument();
			com.sun.star.lang.XMultiServiceFactory documentFactory =
				(com.sun.star.lang.XMultiServiceFactory) UnoRuntime.queryInterface(
					com.sun.star.lang.XMultiServiceFactory.class, myDoc);

			Object frame = documentFactory.createInstance("com.sun.star.text.TextFrame");
			
			XText docText = myDoc.getText();
			XTextFrame xFrame = (XTextFrame) UnoRuntime.queryInterface(XTextFrame.class, frame);
			ORIGINAL CODE FROM NOATEXT/OPENOFFICE */
			
			System.out.println("MSWord_jsText: insertTextAt: About to ActiveXComponent jacobSelection = jacobObjWord.getPropertyAsComponent(\"Selection\");");
			ActiveXComponent jacobSelection = jacobObjWord.getPropertyAsComponent("Selection");
			System.out.println("MSWord_jsText: insertTextAt: About to Dispatch jacobShapes = Dispatch.get((Dispatch) jacobDocument, \"Shapes\").toDispatch();");
			Dispatch jacobShapes = Dispatch.get((Dispatch) jacobDocument, "Shapes").toDispatch();
            
			//From the observed positions, knowledge of the Application's Setup Dialog, and VBA Documentation I assume:
			//Quite probably, OpenOffice expects coordinates in mm, and MS Word in Point (or DTPPoint, as Google just called them...,
			//Well, they may have been there before the age of DTP (that I've seen come and go...). :-) (DTP- = DesktopPublishing-) 
			System.out.println("MSWord_jsText: insertTextAt: Incoming coordinates from Elexis: \"AddTextBox\", 1, "+x+" mm, "+y+" mm, "+w+" mm, "+h+" mm); (OpenOffice coordinates)");
			
			double mmToPoint = 2.83465;
			long xpt = Math.round(x * mmToPoint);
			long ypt = Math.round(y * mmToPoint);
			long wpt = Math.round(w * mmToPoint);
			long hpt = Math.round(h * mmToPoint);
						
			//N.B.: Orientation muss mehr als 0 sein (der erste Parameter nach "AddShape") - Shapes enthalten *hier* keinen Text.
			//System.out.println("MSWord_jsText: insertTextAt: About to Variant jacobShapeVariant = Dispatch.call(jacobShapes, \"AddShape\", 1, "+xpt+" pt, "+ypt+" pt, "+wpt+" pt, "+h+" pt);  (MS Word coordinates)");
            //Variant jacobShapeVariant = Dispatch.call(jacobShapes, "AddShape", 1, xpt, ypt, wpt, hpt);

			System.out.println("MSWord_jsText: insertTextAt: About to Variant jacobShapeVariant = Dispatch.call(jacobShapes, \"AddTextBox\", 1, "+xpt+" pt, "+ypt+" pt, "+wpt+" pt, "+hpt+" pt); (MS Word coordinates)");
            Variant jacobShapeVariant = Dispatch.call(jacobShapes, "AddTextBox", 1, xpt, ypt, wpt, hpt);
            
            System.out.println("MSWord_jsText: insertTextAt: About to Dispatch jacobShape = jacobShapeVariant.toDispatch();");
            Dispatch jacobShape = jacobShapeVariant.toDispatch();

			//Apparently, coordinates are interpreted by MS Word with respect from the page corner.
			//I would have to add 1 cm to bring it in line with the starting x-position of other shapes and text of the column
			//for Tarmedrechnung_S1 etc. and 0.5 cm for Tarmedrechnung_EZ. So I CANNOT simply use a fixed offset value.
			//The original NoaText/OpenOffice code uses: TextContentAnchorType.AT_PAGE, VertOrientationRelation, RelOrientation.PAGE_FRAME etc.
			//Now I have to look up how to position shapes relativ to inner page margins etc. ...
			
			//https://msdn.microsoft.com/en-us/library/office/ff196943.aspx
			//
			//Every Shape object is anchored to a range of text.
			//A shape is anchored to the beginning of the first paragraph that contains the anchoring range.
			//The shape will always remain on the same page as its anchor.
			//
			//You can view the anchor itself by setting the ShowObjectAnchors property to True.
			//The shape's Top and Left properties determine its vertical and horizontal positions.
			//
			//The shape's Top and Left properties determine its vertical and horizontal positions.
			//The shape's RelativeHorizontalPosition and RelativeVerticalPosition properties determine whether the position is measured
			//from the anchoring paragraph, the column that contains the anchoring paragraph, the margin, or the edge of the page.
            int wdRelativeHorizontalPositionMargin = 0;
            int wdRelativeHorizontalPositionPage = 1;
    	    int wdRelativeHorizontalPositionColumn = 2;
    	    int wdRelativeHorizontalPositionCharacter = 3;
    	    
    	    int wdRelativeVerticalPositionMargin = 0;
    	    int wdRelativeVerticalPositionPage = 1;
    	    int wdRelativeVerticalPositionParagraph = 2;
    	    int wdRelativeVerticalPositionLine = 3;
            
    	    System.out.println("MSWord_jsText: insertTextAt: About to Dispatch.put(jacobShape, \"RelativeHorizontalPosition\",wdRelativeHorizontalPositionMargin);");
            Dispatch.put(jacobShape, "RelativeHorizontalPosition",wdRelativeHorizontalPositionMargin);
    	    //System.out.println("MSWord_jsText: insertTextAt: About to Dispatch.put(jacobShape, \"RelativeHorizontalPosition\",wdRelativeHorizontalPositionColumn);");
            //Dispatch.put(jacobShape, "RelativeHorizontalPosition",wdRelativeHorizontalPositionColumn);
    	    
            //Even though wdRelativeVerticalPositionMargin MIGHT be a correct representation of what happens in NoaText_jsl/OpenOffice,
            //this causes the bottom ESR-Line to be clipped by the bottom page margin. So I use: wdRelativeVerticalPositionPage to get it a little bit higher.
            //NOPE, IT'S NOT THE FIELD POSITION VS. PAGE MARGINS, IT'S THE FIELD HEIGHT WHICH IS A BIT TOO LOW.
            
            //ToDo: CHECK IF THIS FITS IN WITH ACTUAL PRE-PRINTED ESR FORM PAPER.
            
            System.out.println("MSWord_jsText: insertTextAt: TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO");  
            System.out.println("MSWord_jsText: insertTextAt: ToDo: Check whether generated Tarmed bill AddTextBox jacobShape RelativeHorizontalPosition RelativeVerticalPosition etc.");  
            System.out.println("MSWord_jsText: insertTextAt: ToDo: matches coordinates of actual pre-printed ESR forms (if possible, with alignment unchanged from NoaText_jsl.");  
            System.out.println("MSWord_jsText: insertTextAt: TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO");  
            
            System.out.println("MSWord_jsText: insertTextAt: About to Dispatch.put(jacobShape, \"RelativeVerticalPosition\",wdRelativeVerticalPositionMargin);");
            Dispatch.put(jacobShape, "RelativeVerticalPosition",wdRelativeVerticalPositionMargin);
            //System.out.println("MSWord_jsText: insertTextAt: About to Dispatch.put(jacobShape, \"RelativeVerticalPosition\",wdRelativeVerticalPositionPage);");
            //Dispatch.put(jacobShape, "RelativeVerticalPosition",wdRelativeVerticalPositionPage);
            
            //Sadly, when I change the RelativeHorizontalPosition, RelativeVerticalPosition,
            //the already defined x,y coordinates are changed at the same time,
            //so that the Shape/TextBox actually KEEPS its position on the paper unchanged.
            //As I can apparently NOT set the Relative... before AddTextBox (or at least don't know how to do that yet...)
            //I need to update the top, left coordinates AFTER updating the reference point. 
            Dispatch.put(jacobShape, "Left",xpt);
            Dispatch.put(jacobShape, "Top",ypt);
            //YEP, that finally puts them into the right-looking positions (and the same ones they appear in with NoaText_jsl / OpenOffice 3.x)
           
            //Siehe weiter unten für AutoSize = True und WordWrap = False. Damit wirklich kein Text mehr durchgeschnitten wird.
            
            
            
            
            //Now, remove the black surrounding lines...
            Dispatch jacobShapeLine = Dispatch.get((Dispatch) jacobShape, "Line").toDispatch();
            Dispatch.put(jacobShapeLine, "Visible", 0);
            
            
			/*
			XShape xWriterShape = (XShape) UnoRuntime.queryInterface(XShape.class, xFrame);
			
			xWriterShape.setSize(new Size(w * 100, h * 100));
			
			XPropertySet xFrameProps =
				(XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xFrame);
			
			// Setting the vertical position
			xFrameProps.setPropertyValue("AnchorPageNo", new Short((short) 1));
			xFrameProps.setPropertyValue("VertOrientRelation", RelOrientation.PAGE_FRAME);
			xFrameProps.setPropertyValue("AnchorType", TextContentAnchorType.AT_PAGE);
			xFrameProps.setPropertyValue("HoriOrient", HoriOrientation.NONE);
			xFrameProps.setPropertyValue("VertOrient", VertOrientation.NONE);
			xFrameProps.setPropertyValue("HoriOrientPosition", x * 100);
			xFrameProps.setPropertyValue("VertOrientPosition", y * 100);
			*/
			
            /*
			XTextCursor docCursor = docText.createTextCursor();
			docCursor.gotoStart(false);
			// docText.insertControlCharacter(docCursor,ControlCharacter.PARAGRAPH_BREAK,false);
			docText.insertTextContent(docCursor, xFrame, false);
			*/
            
			// get the XText from the shape
						
            /* ORIGINAL NOATEXT/OPENOFFICE IMPLEMENTATION
			// XText xShapeText = ( XText ) UnoRuntime.queryInterface( XText.class, writerShape );

			XText xFrameText = xFrame.getText();
			XTextCursor xtc = xFrameText.createTextCursor();
			ORIGINAL NOATEXT/OPENOFFICE IMPLEMENTATION */
			
			System.out.println("MSWord_jsText: insertTextAt: About to Dispatch jacobShapeTextFrame = Dispatch.call(jacobShape, \"TextFrame\").toDispatch();");
            Dispatch jacobShapeTextFrame = Dispatch.call(jacobShape, "TextFrame").toDispatch();
            
            
            //DAS FOLGENDE REDUZIERT ZWAR DAS CLIPPING FUER MANCHE SHAPES (insb. auf der ersten Seite), ABER GERADE DIE ABSCHLIESSENDE ESR-ZEILE WIRD TROTZDEM HEFTIG GECLIPPT.
            /*
            //Ensure that the Shape has sufficient y-space to display the contained lines of text (i.e. round its height upwards to multiples of text height, if that is available).
            if ( hi > 0 ) {
            	long hRoundedUp = (((int) hpt / (int) hi) + 1) * hpt;
            	Dispatch.put(jacobShape, "Height", hRoundedUp);
            }
            */
            //STATTDESSEN ERLAUBE ICH GROESSENANPASSUNG NACH BEDARF DES TEXTS, UND VERBIETE WRAPPING DES TEXTS:
            System.out.println("MSWord_jsText: insertTextAt: About to Dispatch.put(jacobShapeTextFrame, \"WordWrap\", false);");
            Dispatch.put(jacobShapeTextFrame, "WordWrap", false);
            System.out.println("MSWord_jsText: insertTextAt: About to Dispatch.put(jacobShapeTextFrame, \"AutoSize\", true);");
            Dispatch.put(jacobShapeTextFrame, "AutoSize", true);
            
            
            
            System.out.println("MSWord_jsText: insertTextAt: About to Integer jacobShapeTextFrameHasText = Dispatch.get(jacobShapeTextFrame, \"HasText\").toInt();");
            Integer jacobShapeTextFrameHasText = Dispatch.get(jacobShapeTextFrame, "HasText").toInt();
            
            System.out.println("MSWord_jsText: insertTextAt: Shape[added].TextFrame.HasText="+jacobShapeTextFrameHasText);
            
            if (jacobShapeTextFrameHasText == -1) {
            	Dispatch jacobShapeTextFrameTextRange = Dispatch.call(jacobShapeTextFrame, "TextRange").toDispatch();
            	String jacobShapeTextFrameTextRangeText = Dispatch.get(jacobShapeTextFrameTextRange, "Text").toString();
                
                if (jacobShapeTextFrameTextRangeText == null)	System.out.println("MSWord_jsText: insertTextAt: WARNING: Shape[added].TextFrame.TextRange.Text IS NULL");
                else {
                	System.out.println("MSWord_jsText: insertTextAt: Shape[added].TextFrame.TextRange.Text="+jacobShapeTextFrameTextRangeText);

                	//THIS WORKS, and causes the text in the shape to become selected            
                	System.out.println("MSWord_jsText: insertTextAt: About to Variant jacobShapeTextFrameTextRangeSelectVariant = Dispatch.call(jacobShapeTextFrameTextRange, \"Select\");");
                    Variant jacobShapeTextFrameTextRangeSelectVariant = Dispatch.call(jacobShapeTextFrameTextRange, "Select");
                    if (jacobShapeTextFrameTextRangeSelectVariant == null)	System.out.println("MSWord_jsText: insertTextAt: WARNING: jacobShapeTextFrameTextRangeSelectVariant IS NULL");
                    else 	System.out.println("MSWord_jsText: insertTextAt: jacobShapeTextFrameTextRangeSelectVariant="+jacobShapeTextFrameTextRangeSelectVariant.toString());                        
                    }

			/* ORIGINAL CODE FROM NOATEXT/OPENOFFICE
			com.sun.star.beans.XPropertySet charProps = setFormat(xtc);
			Das folgende ersetzt die entsprechende Prozedur:
			ORIGINAL CODE FROM NOATEXT/OPENOFFICE`*/
			
	    	System.out.println("MSWord_jsText: insertTextAt: About to Dispatch fontDispatch = Dispatch.get(jacobSelection, \"Font\").toDispatch();");
            Dispatch fontDispatch = Dispatch.get(jacobSelection, "Font").toDispatch();
	    	if ( font != null )		{ 
		    	System.out.println("MSWord_jsText: insertTextAt: About to Dispatch.put(fontDispatch, \"Name\", new Variant(font));");
	    		Dispatch.put(fontDispatch, "Name", new Variant(font));
	    	} else {
	    		System.out.println("MSWord_jsText: insertTextAt: WARNING: font IS NULL.");
	    	}
	        if (hi > 0)				{ 
	        	System.out.println("MSWord_jsText: insertTextAt: Dispatch.put(fontDispatch, \"Size\", new Float(hi));");
	    		//Dispatch.put(fontDispatch, "CharHeight", new Float(hi)); 	//OpenOffice: Height of the character in point
	    		Dispatch.put(fontDispatch, "Size", new Float(hi)); 
	        }
	        if (stil > -1) {
        		System.out.println("MSWord_jsText: WARNING: The MS Word FONT property does apparently NOT support numeric font weight, so we have fewer steps available. { SWT.MIN = SWT.NORMAL; SWT.BOLD }"); 
	        	switch (stil) {
	        	case SWT.MIN:		{ 
	        		System.out.println("MSWord_jsText: insertTextAt: Mapping: OpenOffice SWT.MIN -> Dispatch.put(fontDispatch, \"Bold\", false);");
	        		//Dispatch.put(fontDispatch, "CharWeight", 15f); break; 
		        	Dispatch.put(fontDispatch, "Bold", false);
	        	}
	        	case SWT.NORMAL:	{ 
		        	System.out.println("MSWord_jsText: insertTextAt: Mapping: OpenOffice SWT.NORMAL -> Dispatch.put(fontDispatch, \"Bold\", false);");
	        		//Dispatch.put(fontDispatch, "CharWeight", FontWeight.LIGHT); break;
		        	Dispatch.put(fontDispatch, "Bold", false);
	        	}
	        	case SWT.BOLD:		{ 
		        	System.out.println("MSWord_jsText: insertTextAt: Mapping: OpenOffice SWT.BOLD -> Dispatch.put(fontDispatch, \"Bold\", true);");
	        		//Dispatch.put(fontDispatch, "CharWeight", FontWeight.BOLD); break;
		        	Dispatch.put(fontDispatch, "Bold", true);
	        	}
		        }
	        }
	        	
	        /*
	         Oben sind die Parameter von OpenOffice wie in NoaText_jsl verwendet.
	         Ansonsten gäbe es laut Infoseiten für MS Word VBA wohl diese Parameter:
	         
	         	Dispatch.put(fontDispatch, "Size", new Float(hi)); 
	        	Dispatch.put(fontDispatch, "Bold", new Variant(bold)
	            Dispatch.put(fontDispatch, "Italic", new Variant(italic));
	            Dispatch.put(fontDispatch, "Underline", new Variant(underLine));
	            Dispatch.put(fontDispatch, "Color", colorSize);
	        */
			
	        /* ORIGINAL FROM NOATEXT/OPENOFFICE
			ParagraphAdjust paradj;
			switch (adjust) {
			case SWT.LEFT:
				paradj = ParagraphAdjust.LEFT;
				break;
			case SWT.RIGHT:
				paradj = ParagraphAdjust.RIGHT;
				break;
			default:
				paradj = ParagraphAdjust.CENTER;
			}
			
			charProps.setPropertyValue("ParaAdjust", paradj);
			xFrameText.insertString(xtc, text, false);
			
			ORIGINAL FROM NOATEXT/OPENOFFICE */
	    	int wdAlignParagraphLeft = 0;
	    	int wdAlignParagraphCenter = 1;
			int wdAlignParagraphRight = 2;
			int wdAlignParagraphJustify = 3;
	    	int wdAlignParagraphDistribute = 4;
			int wdAlignParagraphJustifyMed = 5;
			int wdAlignParagraphJustifyHi = 7;
			int wdAlignParagraphJustifyLow = 8;
			int wdAlignParagraphThaiJustify=9;
			
			System.out.println("MSWord_jsText: insertTextAT: About to ActiveXComponent jacobSelectionParagraphFormatAXC = jacobSelection.getPropertyAsComponent(\"ParagraphFormat\");");
			ActiveXComponent jacobSelectionParagraphFormatAXC = jacobSelection.getPropertyAsComponent("ParagraphFormat");
			switch (adjust) {
			case SWT.LEFT:
				//paradj = ParagraphAdjust.LEFT;
				System.out.println("MSWord_jsText: adjust=SWT.LEFT -> insertTextAT: About to Dispatch.put(jacobSelectionParagraphFormatAXC, \"Alignment\", wdAlignParagraphLeft);");
				Dispatch.put(jacobSelectionParagraphFormatAXC, "Alignment", wdAlignParagraphLeft);
				break;
			case SWT.RIGHT:
				//paradj = ParagraphAdjust.RIGHT;
				System.out.println("MSWord_jsText: adjust=SWT.RIGHT -> insertTextAT: About to Dispatch.put(jacobSelectionParagraphFormatAXC, \"Alignment\", wdAlignParagraphRight);");
				Dispatch.put(jacobSelectionParagraphFormatAXC, "Alignment", wdAlignParagraphRight);
				break;
			default:
				//paradj = ParagraphAdjust.CENTER;
				System.out.println("MSWord_jsText: ajdust=default -> insertTextAT: About to Dispatch.put(jacobSelectionParagraphFormatAXC, \"Alignment\", wdAlignParagraphCenter);");
				Dispatch.put(jacobSelectionParagraphFormatAXC, "Alignment", wdAlignParagraphCenter);
			}
			
	    		    	
			/*
			String orig = Dispatch.get(jacobShapeTextFrameTextRange, "Text").toString();
			if (orig  == null)	System.out.println("MSWord_jsText: insertTextAt: ERROR: orig IS NULL!");
			else 				System.out.println("MSWord_jsText: insertTextAt: orig="+orig);
			*/
			
	    	System.out.println("MSWord_jsText: insertTextAt: text == "+text.toString());
			System.out.println("MSWord_jsText: insertTextAt: About to Dispatch.put(jacobShapeTextFrameTextRange, \"Text\", text);");
			Dispatch.put(jacobShapeTextFrameTextRange, "Text", text);
            } // if ...shape... hasText
			            
            //THIS FAILS LATER:
            //...insertText():
            //--------------Exception--------------
            //com.jacob.com.ComFailException: Can't map name to dispid: Text
            //
            //Object xtc = jacobShape;
            Object xtc = jacobSelection.getObject();
            
			System.out.println("MSWord_jsText: insertTextAt ends, returning xtc");
			return xtc;
		} catch (Exception ex) {
			ExHandler.handle(ex);
			System.out.println("MSWord_jsText: insertTextAt caught Exception, ends, returning false");
			return false;
		}
		
	}
	
	
	
	
	
	/**
	 * Print the contents of the panel. NOA does no allow us to select printer and tray, so we do it
	 * with UNO again.
	 */
	public boolean print(final String toPrinter, final String toTray,
		final boolean waitUntilFinished){
		System.out.println("MSWord_jsText: print begins");

		System.out.println("MSWord_jsText: print: TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO");
		System.out.println("MSWord_jsText: print: ToDo: Test/Review printer selection, printer tray selection.");
		System.out.println("MSWord_jsText: print: ToDo: Add waitUntilFinished support.");
		System.out.println("MSWord_jsText: print: ToDo: Die unit MSWord_jsPrinter gibt's ja auch noch. Hier nur verwendet, um das Tray zu setzen?!? In MSWord_jsText gar nicht mehr nötig?");
		System.out.println("MSWord_jsText: print: TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO");

		if ( jacobDocument != null ) { 
			
			try {
				
				/*
				PropertyValue[] pprops;
				if (StringTool.isNothing(toPrinter)) {
					pprops = new PropertyValue[1];
					pprops[0] = new PropertyValue();
					pprops[0].Name = "Pages";
					pprops[0].Value = "1-";
				} else {
					pprops = new PropertyValue[2];
					pprops[0] = new PropertyValue();
					pprops[0].Name = "Pages";
					pprops[0].Value = "1-";
					pprops[1] = new PropertyValue();
					pprops[1].Name = "Name";
					pprops[1].Value = toPrinter;
				}
				if (!StringTool.isNothing(toTray)) {
					XTextDocument myDoc = agIonDoc.getXTextDocument();
					// XTextDocument myDoc=(XTextDocument)
					// UnoRuntime.queryInterface(com.sun.star.text.XTextDocument.class,
					// doc);
					if (!MSWord_jsPrinter.setPrinterTray(myDoc, toTray)) {
						return false;
					}
				}
				XPrintable xPrintable =
					(XPrintable) UnoRuntime.queryInterface(com.sun.star.view.XPrintable.class, agIonDoc
						.getXTextDocument());
				
				com.sun.star.view.XPrintJobBroadcaster selection =
					(com.sun.star.view.XPrintJobBroadcaster) UnoRuntime.queryInterface(
						com.sun.star.view.XPrintJobBroadcaster.class, xPrintable);
				
				MyXPrintJobListener myXPrintJobListener = new MyXPrintJobListener();
				selection.addPrintJobListener(myXPrintJobListener);

				// bean.getDocument().print(pprops);
				xPrintable.print(pprops);
				*/
				
				/* From:
				   https://sourceforge.net/p/jacob-project/discussion/375946/thread/0dff623a/
				   http://www.cnblogs.com/hzj-/articles/1732567.html
				   
				 	int wdPrinterDefaultBin = 0;
					int wdPrinterOnlyBin = 1;
					int wdPrinterUpperBin = 1;
					int wdPrinterLowerBin = 2;
					int wdPrinterMiddleBin = 3;
					int wdPrinterManualFeed = 4;
					int wdPrinterEnvelopeFeed = 5;
					int wdPrinterManualEnvelopeFeed = 6;
					int wdPrinterAutomaticSheetFeed = 7;
					int wdPrinterTractorFeed = 8;
					int wdPrinterSmallFormatBin = 9;
					int wdPrinterLargeFormatBin = 10;
					int wdPrinterLargeCapacityBin = 11;
					int wdPrinterPaperCassette = 14;
					int wdPrinterFormSource = 15;
				 */
				
				Variant oldPrinter = null;																							//201701020956js
				
				if (StringTool.isNothing(toPrinter)) {
					System.out.println("MSWord_jsText: print: toPrinter.isNothing==true");
					System.out.println("MSWord_jsText: print: Leaving existing printer settings unchanged.");
				} else {
					System.out.println("MSWord_jsText: print: toPrinter=="+toPrinter);
										
					//Remember the current printer before changing it.																//201701020956js
					System.out.println("MSWord_jsText: print: About to oldPrinter = jacobObjWord.getProperty(\"ActivePrinter\");");	//201701020956js
					oldPrinter = jacobObjWord.getProperty("ActivePrinter"); 														//201701020956js
					System.out.println("MSWord_jsText: print: oldPrinter =="+oldPrinter.toString());								//201701020956js
					
					//This will actually change the system default printer!															//201701020956js
					System.out.println("MSWord_jsText: print: About to jacobObjWord.setProperty(\"ActivePrinter\",new Variant(toPrinter));");
					jacobObjWord.setProperty("ActivePrinter",new Variant(toPrinter));  //toPrinter should be like: \\\\server\\printername on NE05:" 
				}
				
				if (StringTool.isNothing(toTray)) { 
					System.out.println("MSWord_jsText: print: toTray.isNothing==true");
					System.out.println("MSWord_jsText: print: Leaving existing tray settings unchanged.");
				} else {
					System.out.println("MSWord_jsText: print: toTray=="+toTray);
					System.out.println("MSWord_jsText: print: About to Dispatch.put(Dispatch.call(jacobObjWord, \"PageSetup\").toDispatch(), \"FirstPageTray\", toTray);");
					Dispatch.put(Dispatch.call(jacobObjWord, "PageSetup").toDispatch(), "FirstPageTray", toTray);
					System.out.println("MSWord_jsText: print: About to Dispatch.put(Dispatch.call(jacobObjWord, \"PageSetup\").toDispatch(), \"OtherPagesTray\", toTray);");
					Dispatch.put(Dispatch.call(jacobObjWord, "PageSetup").toDispatch(), "OtherPagesTray", toTray);
				}

				System.out.println("MSWord_jsText: print: About to Dispatch.call(jacobDocument, \"PrintOut\");");
				Variant printOutCallResult = Dispatch.call(jacobDocument, "PrintOut");

				System.out.println("MSWord_jsText: print: printOutCallResult=="+printOutCallResult);

				//Restore the previously active default printer again.																//201701020956js
				if (oldPrinter != null) {																							//201701020956js
					System.out.println("MSWord_jsText: print: About to jacobObjWord.setProperty(\"ActivePrinter\",jacobObjWord.oldPrinter);");	//201701020956js
					jacobObjWord.setProperty("ActivePrinter",oldPrinter);															//201701020956js 
				} else {																											//201701020956js
					System.out.println("MSWord_jsText: print: WARNING: oldPrinter IS NULL. So I'm not trying to set ActivePrinter (back) to that.");	//201701020956js					
				}																													//201701020956js
				
				/*
				long timeout = System.currentTimeMillis();
				while ((myXPrintJobListener.getStatus() == null)
					|| (myXPrintJobListener.getStatus() == PrintableState.JOB_STARTED)) {
					Thread.sleep(100);
					long to = System.currentTimeMillis();
					if ((to - timeout) > 10000) {
						break;
					}
				}
				*/

				

				//ToDo: Is this correct? - Even if yes: WHY?
				//Documents that are printed from out of Elexis shall usually be closed right after printing
				//They closed themselves automatically in the NoaText_jsl environment,
				//but apparently stay open in the MSWord_js environment (WHY???)
				//I've now added this call to ensure that Etiketten (all three kinds),
				//and Rechnungen (ESR Formular und A4 Tarmed Rechnung) are automatically closed after printing.
				//Otherwise they stayed open forever.
				//An open problem/question is
				//that Etiketten should DEFINITELY not be saved as documents in the DB after printing,
				//and Rechnunge should PROBABLY not be saved as documents in the DB after printing (I assume).
				//ToDo: Or should they, so that we could review them as complete MS Word documents???
				System.out.println("MSWord_jsText: print: ToDo: Is this correct? - Even if yes: WHY?");														//201701022040js
				System.out.println("MSWord_jsText: print: ToDo: Should Bills (Rechnungen) go to the DB as Word documents? (Cave: Numerous documents!)");	//201701022040js
				
				//System.out.println("MSWord_jsText: print: About to Dispatch.call(jacobDocument, \"Close\"); (after PrintOut)");
				//Variant closeCallResult = Dispatch.call(jacobDocument, "Close");		//201702012019js added this
				//System.out.println("MSWord_jsText: print: closeCallResult=="+closeCallResult);
				//
				//ToDo: / PleaseNote: Ich sehe, dass das MS Word Dokument geschlossen wird, aber MS Word bleibt offen
				//(nach dem Etikettendruck). Das ist soweit wohl brauchbar. Nach dem Etikettendruck bleiben rote
				//Einträge im Log-Output, wie auch vor Einfügen dieses Close Calls hier (das gibt's nicht bei Rechnungen,
				//und kam auch zuvor nicht jedesmal aber meistens? bei Etiketten, ich kenne die Ursache nicht.)
				//ToDo: BÄH. Da bleibt wirklich ein leeres Word - Fenster, und: Das nächste Etikett macht TROTZDEM
				//ein NEUES Word-Fenster auf, nach dem Schliessen dieses Etiketts geht das AUCH nicht weg - so sammeln
				//sich statt vieler offener Dokumente viele offene leere Word Fenster... :-(
				
				//Das wirft auch einen Fehler:
				//Variant closeCallResult = Dispatch.call(jacobDocument, "Quit");		//201702012019js added this
				//System.out.println("MSWord_jsText: print: closeCallResult=="+closeCallResult);
				
				//Das erzeugt nur eine Fehlermeldung:
				//Variant closeCallResult2 = Dispatch.call(jacobObjWord, "Close");		//201702012019js added this
				//System.out.println("MSWord_jsText: print: closeCallResult2=="+closeCallResult2);
				
				System.out.println("MSWord_jsText: print: About to dispose();");														//201701022040js
				System.out.println("(Schliesst das MS Word-Fenster des ad-hoc für Ausdruck erzeugten Dokuments nach dem PrintOut.)");	//201701022040js
				dispose();																												//201701022040js
				//OK, SO ist es gut. Das schliesst das Etikett, UND das umgebende MS Word-Fenster nach dem Ausdruck.
				//Auch habe ich getestet:
				//- Einen Brief vom selben Patienten davor aufmachen.
				//- Etwas reintippen
				//- Mit geöffnetem Brief ein Etikett drucken: Das funktioniert,
				//  und das Etiketten-Word-Fenster wird danach geschlossen,
				//  und das Word-Fenster mit dem Brief wird nicht betroffen (wandert nur hinter Elexis).
				//- Danach etwas in den Brief tippen
				//- Brief-Window schliessen
				//- Brief wieder öffnen: Alle Änderungen wurden korrekt gespeichert :-)
				//
				//Dasselbe hab ich gerade mit dem Tarmed-Rechnungsdruck probiert, auch da funktioniert es.

				
				
				System.out.println("MSWord_jsText: print: TODO: Man könnte das dispose() (und/oder das Auto-Printout, oder ein Pause davor) noch schaltbar machen, um allenfalls vor dem (erneuten) Ausdruck nachkorrigieren zu können."); //201701022040js
				//ToDo: Man könnte das dispose() (und/oder das Auto-Printout, oder ein Pause davor) noch schaltbar machen, um allenfalls vor dem (erneuten) Ausdruck nachkorrigieren zu können. 
				
				//ToDo: FRAGE: Warum brauche ich das dispose() hier überhaupt?
				//ToDo: Funktioniert allenfalls ein übergeordnetes Äquivalent hier nicht mehr, das in noatext_jsl noch ging?
				System.out.println("MSWord_jsText: print: TODO: ***********************************************************************************************************");	//201701022040js
				System.out.println("MSWord_jsText: print: TODO: FRAGE: Warum brauche ich das dispose() hier überhaupt?");														//201701022040js
				System.out.println("MSWord_jsText: print: TODO: FRAGE: Funktioniert allenfalls ein übergeordnetes Äquivalent hier nicht mehr, das in noatext_jsl noch ging?");	//201701022040js
				System.out.println("MSWord_jsText: print: TODO: FRAGE: JEP, das könnte sein, weil ein Close() mit SetModified o.ä. mglw. nicht ging... - siehe logs.");			//201701022040js	
				System.out.println("MSWord_jsText: print: TODO: FRAGE: Und ich hoffe mal, dass ein dispose(); hier keine andere Funktionalität zerbricht...");					//201701022040js
				System.out.println("MSWord_jsText: print: TODO: ***********************************************************************************************************");	//201701022040js
				System.out.println("MSWord_jsText: print: TODO: N.B.: Man findet die Dokumente jetzt übrigens notfalls in der letzten Fassung im %AppData%...\\Temp Ordner.");	//201701022040js
				System.out.println("MSWord_jsText: print: TODO: ***********************************************************************************************************");	//201701022040js
				
				System.out.println("MSWord_jsText: print ends, returns true");
				return true;
			} catch (Exception ex) {
				ExHandler.handle(ex);
				System.out.println("MSWord_jsText: print: caught Exception");
			}
		} // if ( jacobDocument != null )
		else {
			System.out.println("MSWord_jsText: print: ERROR: jacobDocument IS NULL!");			
		}

	System.out.println("MSWord_jsText: print ends, returning false");
	return false;
	}
	
	
	
	public void setFocus(){
		System.out.println("MSWord_jsText: setFocus stub");
		// TODO Auto-generated method stub
	
	}
	
	
	
	public void setFormat(final PageFormat f){
		System.out.println("MSWord_jsText: setFormat(final PageFormat f) stub");
		// TODO Auto-generated method stub
	
	}
	
	
	
	public void setSaveOnFocusLost(final boolean bSave){
		System.out.println("MSWord_jsText: setSaveOnFocusLost stub");
		// TODO Auto-generated method stub
	
	}
	
	
	
	public void showMenu(final boolean b){
		System.out.println("MSWord_jsText: showMenu stub");
		// TODO Auto-generated method stub
	
	}
	
	
	
	public void showToolbar(final boolean b){
		System.out.println("MSWord_jsText: showToolbar stub");
		// TODO Auto-generated method stub
	
	}
	
	
	
	public void setInitializationData(final IConfigurationElement config,
		final String propertyName, final Object data) throws CoreException{
		System.out.println("MSWord_jsText: setInitializationData stub");
		// TODO Auto-generated method stub
	
	}
	
	
	
	/**
	 * basically: ensure that OpenOffice is happy closing the document and create a new temporary file
	 */
	private void clean(){
		System.out.println("MSWord_jsText: clean begins");
		
		/* ORIGINAL CODE FROM NOATEXT/OPENOFFICE IMPLEMENTATION
		try {
			if (agIonDoc != null) {
				System.out.println("MSWord_jsText: clean: about to doc.getPersistenceService().store("+myFile.getAbsolutePath()+")...");
				agIonDoc.getPersistenceService().store(myFile.getAbsolutePath());

				System.out.println("MSWord_jsText: clean: TODO / INFO: //doc.close() is commented out here!!");
				// doc.();

				System.out.println("MSWord_jsText: clean: about to myfile.delete()...");
				myFile.delete();
			}
			else { System.out.println("MSWord_jsText: clean: WARNING: doc IS NULL!"); };

			
			System.out.println("MSWord_jsText: clean: about to myFile = File.createTempFile(\"noa\", \".odt\");");
			myFile = File.createTempFile("noa", ".odt");

			System.out.println("MSWord_jsText: clean: about to myFile.deleteOnExit();");
			myFile.deleteOnExit();
			
		} catch (Exception ex) {
			ExHandler.handle(ex);
		}
		ORIGINAL CODE FROM NOATEXT/OPENOFFICE IMPLEMENTATION */
		
		
		
		try {
			if (jacobDocument != null) {
				String myFilename=myFile.getAbsolutePath();
				System.out.println("MSWord_jsText: clean: Trying to save the jacobDocument: "+myFilename+" from jacobDocument == "+jacobDocument.toString()+"...");
				System.out.println("MSWord_jsText: clean: About to Dispatch.call( (Dispatch) Dispatch.call(jacobObjWord, \"WordBasic\").getDispatch(),\"FileSaveAs\", myFilename);");
				Dispatch.call( (Dispatch) Dispatch.call(jacobObjWord, "WordBasic").getDispatch(),"FileSaveAs", myFilename); 

				System.out.println("");
				System.out.println("MSWord_jsText: clean: TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO");
				System.out.println("MSWord_jsText: clean: TODO WHY is doc.close() commented out here in the NOATEXT/OPENOFFICE Implementation? In which versions of that?!!");
				System.out.println("MSWord_jsText: clean: TODO I am making it active in the msword_js version of the plugin, because it feels plausible to do here.");
				System.out.println("MSWord_jsText: clean: TODO PLEASE CHECK whether that's correct, and if needed, backport to NoaText_jsl or other versions.");
				System.out.println("MSWord_jsText: clean: TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO");
				System.out.println("MSWord_jsText: clean: TODO   N.B.: It would NOT be necessary to close the old brief when a new one is opened, when we use detached MS Word Windows.");
				System.out.println("MSWord_jsText: clean: TODO   		That's only necessary when we have only 1 Frame to display e.g. 1 Brief at the same time.");				
				System.out.println("MSWord_jsText: clean: TODO   		So we could also provide contemporary-multibrief-viewing capability through mword_js; which might actually be helpful");				
				System.out.println("MSWord_jsText: clean: TODO   		e.g. when a course over time shall be reviewed, or elements from mulitple old documents combined to a new one.");				
				System.out.println("MSWord_jsText: clean: TODO   IF we close() however, we should also quit() or we would leave an MS Word Window open every time we change the displayed Brief content.");				
				System.out.println("MSWord_jsText: clean: TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO");
				System.out.println("");
				
				//System.out.println("MSWord_jsText: clean: About to Dispatch.call(jacobDocument, \"Close\", new Variant(false)); jacobDocument = null;");
			    //Dispatch.call(jacobDocument, "Close", new Variant(false));
			    //jacobDocument = null;
				close();	//this includes: jacobDocument = null;
				quit();		//this includes: jacobObjWord = null; jacobDocuments = null;	//Otherwise, an open empty MS Word Window would remain, e.g. each time another Brief would be dblclicked in Briefauswahl. 

				System.out.println("MSWord_jsText: clean: about to myfile.delete()...");
				myFile.delete();
			}
			else { System.out.println("MSWord_jsText: clean: WARNING: jacobDocument IS NULL!"); };

			
			//201611131641js Umstellung von *.odt auf *.doc für MS-Word
			System.out.println("MSWord_jsText: clean: about to myFile = File.createTempFile(\"MSWord_jsText\", \".doc\");");
			myFile = File.createTempFile("MSWord_jsText", ".doc");
			//System.out.println("MSWord_jsText: clean: about to myFile = File.createTempFile(\"noa\", \".odt\");");
			//myFile = File.createTempFile("noa", ".odt");

			System.out.println("MSWord_jsText: clean: TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO");
			System.out.println("MSWord_jsText: clean: TODO myFile.deleteOnExit(); commented out. Seems safer to do so.");
			System.out.println("MSWord_jsText: clean: TODO PLEASE CHECK whether that's truly preferred, and if needed, backport to NoaText_jsl or other versions.");
			System.out.println("MSWord_jsText: clean: TODO PLEASE CHECK whether deleteOnExit() is used/implemented anywhere else, because I thought it WAS already disabled.");
			System.out.println("MSWord_jsText: clean: TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO");
			//ToDo: Check/Review whether MSWord_js.clean().myFile.deleteOnExit() shall remain COMMENTED OUT. 
			//System.out.println("MSWord_jsText: clean: about to myFile.deleteOnExit();");
			//myFile.deleteOnExit();
			
		} catch (Exception ex) {
			ExHandler.handle(ex);
		}
		
		
		System.out.println("MSWord_jsText: clean ends");
	}
	
	
	
	public boolean setFont(final String name, final int style, final float size){
		System.out.println("MSWord_jsText: setFont begins");
		
		font = name;
		hi = size;
		stil = style;
		
		System.out.println("MSWord_jsText: setFont returning true");
		return true;
	}
	
	
	
	public boolean setStyle(final int style){
		System.out.println("MSWord_jsText: setStyle begins");
		
		stil = style;
		
		System.out.println("MSWord_jsText: setStyle returning true");
		return true;
	}
	
	
	
	//I FIRST ADOPTED THIS THEN I COMMENTED IT OUT LARGELY, BECAUSE THERE ARE NO REFERENCES AT ALL FROM THROUGHOUT THE WORKSPACE.
	
	//THIS IS ONLY CALLED FROM about three positions, e.g. insertText(int, int, int, int, String, int) above.
	//I added replacements inline right there, to avoid Java complications when moving code over here.
	//(probably could have replaced XTextCursor xtc by some reference to selection, but I've had sufficient
	//bad experiences with Java introduced code complications and don't want to spend more time on them. 
	//For now, the inline replacement does to Word via Jacob some close equivalent of what this did to OpenOffice via Noatext.
	//
	//I keep the procedure stub mainly to catch yet unknown stray calls to it, and to be a starting point for future refactoring.
	//If you have the time...
				
	private com.sun.star.beans.XPropertySet setFormat(final XTextCursor xtc)
		throws UnknownPropertyException, PropertyVetoException, IllegalArgumentException,
		WrappedTargetException{

		System.out.println("MSWord_jsText: com.sun.star.beans.XPropertySet setFormat begins");
		
		System.out.println("MSWord_jsText: com.sun.star.beans.XPropertySet setFormat: TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO");
		System.out.println("MSWord_jsText: com.sun.star.beans.XPropertySet setFormat: -----------------------------------------------------------------------------------------------------------------------");
		System.out.println("MSWord_jsText: com.sun.star.beans.XPropertySet setFormat: THIS METHOD HAS BEEN SHORT CIRCUITED TO RETURN NULL BECAUSE IT IS PROBABLY NOT USED AT ALL THROUGHOUT THIS PLUGIN.");
		System.out.println("MSWord_jsText: com.sun.star.beans.XPropertySet setFormat: PLEASE CONFIRM WHETHER THIS IS CORRECT. - YOU SHOULD NEVER SEE THIS MESSAGE IN THE LOGS.");
		System.out.println("MSWord_jsText: com.sun.star.beans.XPropertySet setFormat: -----------------------------------------------------------------------------------------------------------------------");
		System.out.println("MSWord_jsText: com.sun.star.beans.XPropertySet setFormat: TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO ");
		
		SWTHelper.showError("MSWord_jsText: com.sun.star.beans.XPropertySet setFormat: ", "Oops!","Diese Nachricht stammt aus einer Methode,\nvon der ich erwartet hatte, dass sie obsolet sei.\nBitte informieren Sie mich, bei welcher Aktion sie erschienen ist -\nich schaue dann den Code dann bei Bedarf nochmals durch.\njoerg.sigle@jsigle.com");

		return null;
		}
		
		
/* COMMENTED OUT - I SAW NO REFERENCES TO THIS METHOD.
		com.sun.star.beans.XPropertySet charProps =
			(com.sun.star.beans.XPropertySet) UnoRuntime.queryInterface(
				com.sun.star.beans.XPropertySet.class, xtc);
		if (font != null) {
			charProps.setPropertyValue("CharFontName", font);
		}
		if (hi > 0)
		{
			charProps.setPropertyValue("CharHeight", new Float(hi));
		}
		if (stil > -1)
		{
			switch (stil) {
			case SWT.MIN:
				charProps.setPropertyValue("CharWeight", 15f /* FontWeight.ULTRALIGHT */ /* COMMENTED OUT CONTINUED);
				break;
			case SWT.NORMAL:
				charProps.setPropertyValue("CharWeight", FontWeight.LIGHT);
				break;
			case SWT.BOLD:
				charProps.setPropertyValue("CharWeight", FontWeight.BOLD);
				break;
			}
		}
	
		System.out.println("MSWord_jsText: setFormat returning charProps");
		return charProps;
COMMENTED OUT */


/*
 *ADDED FOR R&D&UNDERSTANDING; PROBABLY NOT NEEDED FOR ANYTHING,
 *some text plugins use similar stuff maybe to modify what happens during saving
 *at least I had this impression, not researched into it any further.
 *It's sufficient to note down in textHandler the ICallback handler supplied to CreateMe,
 *thereafter we cann call textHandler.save() from inside MSWord_jsText.java
 *to trigger a save-file-and-store-contents-as-BLOB-to-DB. 201610021213js
 * 
	public boolean initiateSave(final ICallback saveNow){
		System.out.println("MSWord_jsText: initiateSave.saveNow begins");
		System.out.println("MSWord_jsText: initiateSave.saveNow *");
		System.out.println("MSWord_jsText: initiateSave.saveNow *");
		System.out.println("MSWord_jsText: initiateSave.saveNow *");
		System.out.println("MSWord_jsText: initiateSave.saveNow *");
		System.out.println("MSWord_jsText: initiateSave.saveNow *");
		System.out.println("MSWord_jsText: initiateSave.saveNow *");
		System.out.println("MSWord_jsText: initiateSave.saveNow *");
		System.out.println("MSWord_jsText: initiateSave.saveNow *");
		System.out.println("MSWord_jsText: initiateSave.saveNow *");
		System.out.println("MSWord_jsText: initiateSave.saveNow *");
		System.out.println("MSWord_jsText: initiateSave.saveNow *");
		System.out.println("MSWord_jsText: initiateSave.saveNow ends, about to return true");
	return true;
	}

*/	
	
	
	
	class closeListener implements ICloseListener {
		
		private IOfficeApplication officeAplication = null;
		
		// ----------------------------------------------------------------------------
		/**
		 * Constructs a new SnippetDocumentCloseListener
		 * 
		 * @author Sebastian Rösgen
		 * @date 17.03.2006
		 */
		public closeListener(final IOfficeApplication officeAplication){
			System.out.println("MSWord_jsText: IOfficeApplication noatext/OpenOffice/Elexis-panel/original closeListener: closeListener: about to: this.officeAplication = officeAplication");
				this.officeAplication = officeAplication;
		}
		
		// ----------------------------------------------------------------------------
		/**
		 * Is called when someone tries to close a listened object. Not needed in here.
		 * 
		 * @param closeEvent
		 *            close event
		 * @param getsOwnership
		 *            information about the ownership
		 * 
		 * @author Sebastian Rösgen
		 * @date 17.03.2006
		 */
		public void queryClosing(final ICloseEvent closeEvent, final boolean getsOwnership){
			System.out.println("MSWord_jsText: IOfficeApplication noatext/OpenOffice/Elexis-panel/original closeListener: queryClosing nop");
			// nothing to do in here
		}
		
		// ----------------------------------------------------------------------------
		/**
		 * Is called when the listened object is closed really.
		 * 
		 * @param closeEvent
		 *            close event
		 * 
		 * @author Sebastian Rösgen
		 * @date 17.03.2006
		 */
		public void notifyClosing(final ICloseEvent closeEvent){
			System.out.println("MSWord_jsText: IOfficeApplication noatext/OpenOffice/Elexis-panel/original closeListener: notifyClosing");
			try {
				System.out.println("MSWord_jsText: closeListener: try... about to: removeMe()");
				removeMe();
			} catch (Exception exception) {
				System.err.println("MSWord_jsText: closeListener: Error closing office application!");
				exception.printStackTrace();
			}
		}
		
		// ----------------------------------------------------------------------------
		/**
		 * Is called when the broadcaster is about to be disposed.
		 * 
		 * @param event
		 *            source event
		 * 
		 * @author Sebastian Rösgen
		 * @date 17.03.2006
		 */
		public void disposing(final IEvent event){
			System.out.println("MSWord_jsText: IOfficeApplication noatext/OpenOffice/Elexis-panel/original closeListener: disposing nop");
			// nothing to do in here
		}
		// ----------------------------------------------------------------------------
		
	}
	
	
	public class WordEventHandler {
	    public void BeforeClose(Variant[] arguments) {
	    	//https://msdn.microsoft.com/en-us/library/microsoft.office.tools.word.document.beforeclose.aspx
	    	//The event occurs before the document closes. To keep the document from closing, set the Cancel argument of the provided CancelEventArgs object to true.
	        System.out.println("WordEventHandler: JaCoB MSWord BeforeClose() event occured...");
	        System.out.println("WordEventHandler: provided arguments[] are:");
	        for (int i = 0; i < arguments.length; i++) System.out.println("WordEventHandler:   arguments["+i+"].toString() == "+arguments[i].toString());	        
	    }

	    public void Close(Variant[] arguments) {
	    	//https://msdn.microsoft.com/en-us/library/microsoft.office.tools.word.document.closeevent.aspx
	    	//Occurs when the document is closed.
	        System.out.println("WordEventHandler: JaCoB MSWord Close() event occured...");
	        System.out.println("WordEventHandler: provided arguments[] are:");
	        for (int i = 0; i < arguments.length; i++) System.out.println("WordEventHandler:   arguments["+i+"].toString() == "+arguments[i].toString());
	        
	        System.out.println("WordEventHandler: JaCoB MSWord Close(): super.getClass() == "+super.getClass().toString());
	        
	        System.out.println("About to initiateSave()...");
	        //try to texthandler.save() and clear the modified flag...
	        clear();
	    }
	}
	

	@Override
	public boolean isDirectOutput() {
		System.out.println("MSWord_jsText: isDirectOutput - always returns false");
		return false;
	}

	@Override
	public void setParameter(Parameter parameter) {
		// TODO Auto-generated method stub
		
	}


	@Override
	public void initTemplatePrintSettings(String template) {
		// TODO Auto-generated method stub
		
	}
}
/****************************************************************************
 * ubion.ORS - The Open Report Suite                                        *
 *                                                                          *
 * ------------------------------------------------------------------------ *
 *                                                                          *
 * Subproject: Office Editor Core                                           *
 *                                                                          *
 *                                                                          *
 * The Contents of this file are made available subject to                  *
 * the terms of GNU Lesser General Public License Version 2.1.              *
 *                                                                          * 
 * GNU Lesser General Public License Version 2.1                            *
 * ======================================================================== *
 * Copyright 2003-2005 by IOn AG                                            *
 *                                                                          *
 * This library is free software; you can redistribute it and/or            *
 * modify it under the terms of the GNU Lesser General Public               *
 * License version 2.1, as published by the Free Software Foundation.       *
 *                                                                          *
 * This library is distributed in the hope that it will be useful,          *
 * but WITHOUT ANY WARRANTY; without even the implied warranty of           *
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU        *
 * Lesser General Public License for more details.                          *
 *                                                                          *
 * You should have received a copy of the GNU Lesser General Public         *
 * License along with this library; if not, write to the Free Software      *
 * Foundation, Inc., 59 Temple Place, Suite 330, Boston,                    *
 * MA  02111-1307  USA                                                      *
 *                                                                          *
 * Contact us:                                                              *
 *  http://www.ion.ag                                                       *
 *  info@ion.ag                                                             *
 *                                                                          *
 ****************************************************************************/
 
/*
 * Last changes made by $Author: markus $, $Date: 2008-09-30 19:59:51 +0200 (Di, 30 Sep 2008) $
 */
package ag.ion.bion.workbench.office.editor.core;

import ag.ion.bion.officelayer.application.IOfficeApplication;
import ag.ion.bion.officelayer.application.OfficeApplicationRuntime;
import ag.ion.bion.officelayer.runtime.IOfficeProgressMonitor;

import org.eclipse.core.runtime.FileLocator;
import org.eclipse.core.runtime.IStatus;
import org.eclipse.core.runtime.Platform;
import org.eclipse.core.runtime.Plugin;
import org.eclipse.core.runtime.Status;

import org.osgi.framework.BundleContext;

import java.io.File;

import java.net.URL;

import java.util.HashMap;
import java.util.ResourceBundle;
import java.util.MissingResourceException;

import java.awt.Frame;

/**
 * The main plugin class to be used in the desktop.
 * 
 * @author Andreas Br�ker
 * @version $Revision: 11647 $
 */
public class EditorCorePlugin extends Plugin {
  
  /** ID of the plugin. */
  public static final String PLUGIN_ID = "ag.ion.bion.workbench.office.editor.core";
  
  //The shared instance.
  private static EditorCorePlugin plugin;
  //Resource bundle.
  private ResourceBundle resourceBundle;
  
  private IOfficeApplication localOfficeApplication = null;
  
  private String librariesLocation = null;
	
  //----------------------------------------------------------------------------
  /**
   * The constructor.
   * 
   * @author Andreas Br�ker
   */
  public EditorCorePlugin() {
    super();
    System.out.println("com.jsigle.msword_js: EditorCorePlugin: EditorCorePlugin() constructor just past super() begins...");
    System.out.println("com.jsigle.msword_js: EditorCorePlugin: EditorCorePlugin() constructor: about to plugin = this;");
		plugin = this;
		try {
		  System.out.println("com.jsigle.msword_js: EditorCorePlugin: EditorCorePlugin() constructor: try...");
		  System.out.println("com.jsigle.msword_js: EditorCorePlugin: EditorCorePlugin() constructor: about to resourceBundle = ResourceBundle.getBundle(\"ag.ion.bion.workbench.office.editor.core.CorePluginResources\");");
		  resourceBundle = ResourceBundle.getBundle("ag.ion.bion.workbench.office.editor.core.CorePluginResources");
		} 
    catch (MissingResourceException missingResourceException) {
	  System.out.println("com.jsigle.msword_js: EditorCorePlugin: EditorCorePlugin() constructor: catching MissingRessourceException; about to resourceBundle = null;");
      resourceBundle = null;
    }
	System.out.println("com.jsigle.msword_js: EditorCorePlugin: EditorCorePlugin() constructor ends");
	}
  //----------------------------------------------------------------------------
  /**
   * This method is called upon plug-in activation.
   * 
   * @param context context to be used
   * 
   * @throws Exception if the bundle can not be started
   * 
   * @author Andreas Br�ker
   */
  public void start(BundleContext context) throws Exception {
    super.start(context);
    System.out.println("com.jsigle.msword_js: EditorCorePlugin.java: start(BundleContext context) begins, just past super.start(context) now...");
        
    if (IOfficeApplication.LOCAL_APPLICATION == null) { System.out.println("com.jsigle.msword_js: EditorCorePlugin.java: start(): WARNING: IOfficeApplication.LOCAL_APPLICATION IS NULL!"); }
    else { System.out.println("com.jsigle.msword_js: EditorCorePlugin.java: start(): IOfficeApplication.LOCAL_APPLICATION == "+IOfficeApplication.LOCAL_APPLICATION ); }

    if (IOfficeApplication.NOA_NATIVE_LIB_PATH == null) { System.out.println("com.jsigle.msword_js: EditorCorePlugin.java: start(): WARNING: IOfficeApplication.NOA_NATIVE_LIB_PATH IS NULL!"); }
    else { System.out.println("com.jsigle.msword_js: EditorCorePlugin.java: start(): IOfficeApplication.NOA_NATIVE_LIB_PATH == "+IOfficeApplication.NOA_NATIVE_LIB_PATH ); }
    	
    System.out.println("com.jsigle.msword_js: EditorCorePlugin.java: start(BundleContext context): about to  System.setProperty(IOfficeApplication.NOA_NATIVE_LIB_PATH,getLibrariesLocation())...");

    //ToDo: This code from ag.ion NOA This will probably throw an error if nativeunix.zip nativeview.dll are NOT present in the lib subdirectory. However, they (and this code?) are only needed for NOA OpenOffice interfaceing, NOT for JaCoB to MS Word!
    //ToDo: Oh well, it ALSO fails IF the two libraries are there. :-(
    
    //ToDo: For JaCoB to MS Word, we should rather look for the jacob native libraries (jacob-1.18-x86.dll and optionally jacob-1.18-x64.dll or more recent versions), if anything at all..
    
    System.out.println("com.jsigle.msword_js: EditorCorePlugin.java: start(BundleContext context): about to check the return value of: getLibrariesLocation() in advance...");
    String dummystring = getLibrariesLocation();
    if (dummystring  == null) { System.out.println("com.jsigle.msword_js: EditorCorePlugin.java: start(): WARNING: getLibrariesLocation() has returned NULL!"); }
    else { System.out.println("com.jsigle.msword_js: EditorCorePlugin.java: start(): getLibrariesLocation() has returned: "+dummystring ); }
    
    System.out.println("com.jsigle.msword_js: EditorCorePlugin.java: start(BundleContext context): about to  System.setProperty(IOfficeApplication.NOA_NATIVE_LIB_PATH,getLibrariesLocation())...");

    //ToDo: Check whether we need to set System.setProperty(IOfficeApplication.NOA_NATIVE_LIB_PATH,getLibrariesLocation()); !!!!
    
    System.out.println("com.jsigle.msword_js: EditorCorePlugin.java: start(BundleContext context): TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO");
    
    System.out.println("com.jsigle.msword_js: EditorCorePlugin.java: start(BundleContext context): WE'RE ABOUT TO SKIP SETTING IOfficeApplication.NOA_NATIVE_LIB_PATH from getLIbrariesLocation()");
    System.out.println("com.jsigle.msword_js: EditorCorePlugin.java: start(BundleContext context): BECAUSE THAT WOULD ONLY RETURN NULL!!!! And throw a bunch of errors, and thereafter,");
    System.out.println("com.jsigle.msword_js: EditorCorePlugin.java: start(BundleContext context):   the \"Briefe\" view in Elexis will only show an error:");
    System.out.println("com.jsigle.msword_js: EditorCorePlugin.java: start(BundleContext context):   \"Texterstellung nicht möglich - Es konnte keine Verbindung mit einem Textplugin hergestellt werden\"...");
    System.out.println("com.jsigle.msword_js: EditorCorePlugin.java: start(BundleContext context):   and the \"Einstellungen\" - \"MSWord_js\" configuration page will show an error like \"Bundle initialization failed (158)\".");
    System.out.println("com.jsigle.msword_js: EditorCorePlugin.java: start(BundleContext context):   All this even when I made sure that the NOA native libs are under com.jsigle.msword_js/lib,");
    System.out.println("com.jsigle.msword_js: EditorCorePlugin.java: start(BundleContext context):   which is similar (at least: one place) to where they (at least: also) are in com.jsigle.noatext_jsl, where it works.");
    System.out.println("com.jsigle.msword_js: EditorCorePlugin.java: start(BundleContext context): IF WE SKIP THAT, on the other hand, NO ERRORS will appear, and the plugin will register,");
    System.out.println("com.jsigle.msword_js: EditorCorePlugin.java: start(BundleContext context):   and the \"Briefe\" view in Elexis will show an empty grey space (which is more correct as long as nothing's been opened there).");
    System.out.println("com.jsigle.msword_js: EditorCorePlugin.java: start(BundleContext context):   and the \"Einstellungen\" - \"MSWord_js\" configuration page will appear correctly");
    System.out.println("com.jsigle.msword_js: EditorCorePlugin.java: start(BundleContext context):   (although with labels mentioning to OpenOffice, and the Path to OpenOffice 3 in the respective field, yet.)");
    System.out.println("com.jsigle.msword_js: EditorCorePlugin.java: start(BundleContext context): SO: The following line may work in the original com.jsigle.noatext_jsl or ch.elexis.noatext projects, but NOT in com.jsigle.msword_js");
    System.out.println("com.jsigle.msword_js: EditorCorePlugin.java: start(BundleContext context): Might be interesting to find out why, or just leave it unresearched,");
    System.out.println("com.jsigle.msword_js: EditorCorePlugin.java: start(BundleContext context):   because we might not need it in its current form anyway for the JaCoB based MS Word integration.");
    System.out.println("com.jsigle.msword_js: EditorCorePlugin.java: start(BundleContext context):");
    System.out.println("com.jsigle.msword_js: EditorCorePlugin.java: start(BundleContext context): ************************************************************************************************************");
    System.out.println("com.jsigle.msword_js: EditorCorePlugin.java: start(BundleContext context): YEP, AND FINALLY, I CAN NOW OPEN LETTERS FROM Briefauswahl IN MSWORD VIA THIS PLUGIN (AGAIN).");
    System.out.println("com.jsigle.msword_js: EditorCorePlugin.java: start(BundleContext context): SO NOW, THE msword_js PLUGIN (proof of feasibility prototype) is back working, in elexis 2.1.7js :-) :-) :-)");
    System.out.println("com.jsigle.msword_js: EditorCorePlugin.java: start(BundleContext context): ************************************************************************************************************");
    System.out.println("com.jsigle.msword_js: EditorCorePlugin.java: start(BundleContext context):");     
    System.out.println("com.jsigle.msword_js: EditorCorePlugin.java: start(BundleContext context): TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO TODO");

    /*
    
    System.setProperty(IOfficeApplication.NOA_NATIVE_LIB_PATH,getLibrariesLocation());
    
    if (IOfficeApplication.NOA_NATIVE_LIB_PATH == null) { System.out.println("com.jsigle.msword_js: EditorCorePlugin.java: start(): WARNING: IOfficeApplication.NOA_NATIVE_LIB_PATH IS NULL!"); }
    else { System.out.println("com.jsigle.msword_js: EditorCorePlugin.java: start(): IOfficeApplication.NOA_NATIVE_LIB_PATH == "+IOfficeApplication.NOA_NATIVE_LIB_PATH ); }
	*/


    /**
     * Workaround in order to integrate the OpenOffice.org window into a AWT frame
     * on Linux based systems. 
     */
    try {
      System.out.println("com.jsigle.msword_js: EditorCorePlugin.java: start(): try...");
      System.out.println("com.jsigle.msword_js: EditorCorePlugin.java: start(): about to new Frame()");
      new Frame();
    }
    catch(Throwable throwable) {
      //only occurs in headless mode, where it doesn't matter
    }
    System.out.println("com.jsigle.msword_js: EditorCorePlugin.java: start() ends");
  }
  //----------------------------------------------------------------------------
  /**
   * This method is called when the plug-in is stopped.
   * 
   * @param context context to be used
   * 
   * @throws Exception if the bundle can not be stopped
   * 
   * @author Andreas Br�ker
   */
  public void stop(BundleContext context) throws Exception {
    super.stop(context);
  }
  //----------------------------------------------------------------------------
  /**
   * Returns the shared instance.
   * 
   * @return shared instance
   * 
   * @author Andreas Br�ker
   */
  public static EditorCorePlugin getDefault() {
    return plugin;
  }
  //----------------------------------------------------------------------------
  /**
   * Returns the string from the plugin's resource bundle,
   * or 'key' if not found.
   * 
   * @param key key to be used
   * 
   * @return string from the plugin's resource bundle,
   * or 'key' if not found
   * 
   * @author Andreas Br�ker
   */
  public static String getResourceString(String key) {
    ResourceBundle bundle = EditorCorePlugin.getDefault().getResourceBundle();
		try {
		  return (bundle != null) ? bundle.getString(key) : key;
		} 
    catch (MissingResourceException missingResourceException) {
      return key;
    }
  }
  //----------------------------------------------------------------------------
  /**
   *  Returns the plugin's resource bundle.
   * 
   * @return plugin's resource bundle
   * 
   * @author Andreas Br�ker
   */
  public ResourceBundle getResourceBundle() {
    return resourceBundle;
  }
  //----------------------------------------------------------------------------
  /**
   * Returns local office application. The instance of the application
   * will be managed by this plugin.
   * 
   * @return local office application
   * 
   * @author Andreas Br�ker
   */
  public synchronized IOfficeApplication getManagedLocalOfficeApplication() {
    if(localOfficeApplication == null) {
      HashMap configuration = new HashMap(1);
      configuration.put(IOfficeApplication.APPLICATION_TYPE_KEY, IOfficeApplication.LOCAL_APPLICATION);
      try {
        localOfficeApplication = OfficeApplicationRuntime.getApplication(configuration);
      }
      catch(Throwable throwable) {
        //can not be - this code must work
        Platform.getLog(getBundle()).log(new Status(IStatus.ERROR, EditorCorePlugin.PLUGIN_ID,
            IStatus.ERROR, throwable.getMessage(), throwable));
      }
    }
    return localOfficeApplication;
  }
  //----------------------------------------------------------------------------
  /**
   * Returns location of the libraries of the plugin. Returns null if the location
   * can not be provided.
   * 
   * @return location of the libraries of the plugin or null if the location
   * can not be provided
   * 
   * @author Andreas Br�ker
   */
  public String getLibrariesLocation() {
    if(librariesLocation == null) {
      try {
        URL url = Platform.getBundle("ag.ion.noa").getEntry("/");
        url  = FileLocator.toFileURL(url);
        String bundleLocation = url.getPath();
        File file = new File(bundleLocation);
        bundleLocation = file.getAbsolutePath();
        bundleLocation = bundleLocation.replace('/', File.separatorChar) + File.separator + "lib";
        librariesLocation = bundleLocation;
      }
      catch(Throwable throwable) {
        return null;
      }
    }
    return librariesLocation;
  }  
  //----------------------------------------------------------------------------
  
}
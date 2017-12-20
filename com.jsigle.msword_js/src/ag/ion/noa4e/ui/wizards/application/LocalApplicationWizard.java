/****************************************************************************
 *                                                                          *
 * NOA (Nice Office Access)                                     						*
 * ------------------------------------------------------------------------ *
 *                                                                          *
 * The Contents of this file are made available subject to                  *
 * the terms of GNU Lesser General Public License Version 2.1.              *
 *                                                                          * 
 * GNU Lesser General Public License Version 2.1                            *
 * ======================================================================== *
 * Copyright 2003-2006 by IOn AG                                            *
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
 *  http://www.ion.ag																												*
 *  http://ubion.ion.ag                                                     *
 *  info@ion.ag                                                             *
 *                                                                          *
 ****************************************************************************/

/*
 * Last changes made by $Author: markus $, $Date: 2008-11-18 14:07:54 +0100 (Di, 18 Nov 2008) $
 */
package ag.ion.noa4e.ui.wizards.application;

import org.eclipse.jface.wizard.IWizard;
import org.eclipse.jface.wizard.Wizard;

import ag.ion.bion.officelayer.application.ILazyApplicationInfo;
import ag.ion.noa4e.ui.NOAUIPluginImages;

/**
 * Wizard in order to define the path of a local OpenOffice.org application.
 * 
 * @author Andreas Br�ker
 * @version $Revision: 11685 $
 */
public class LocalApplicationWizard extends Wizard implements IWizard {

  private LocalApplicationWizardDefinePage localApplicationWizardDefinePage = null;

  private ILazyApplicationInfo[]           applicationInfos                 = null;

  private String                           homePath                         = null;

  //----------------------------------------------------------------------------
  /**
   * Constructs new LocalApplicationWizard.
   * 
   * @author Andreas Br�ker
   */
  public LocalApplicationWizard() {
	this(null);
	System.out.println("LOAW: LocalApplicationWizard() - just constructed new LocalApplicationWizard");
  	setNeedsProgressMonitor(true);
  }

  //----------------------------------------------------------------------------
  /**
   * Constructs new LocalApplicationWizard.
   * 
   * @param applicationInfos application infos to be used (can be null)
   * 
   * @author Andreas Br�ker
   */
  public LocalApplicationWizard(ILazyApplicationInfo[] applicationInfos) {
	System.out.println("LOAW: LocalApplicationWizard(applicationInfos) - constructs new LocalApplicationWizard");
	if (applicationInfos==null) System.out.println("LOAW: Please note: applicationInfos==null");
	else System.out.println("LOAW: Please note: applicationInfos: "+applicationInfos.toString());
	
	this.applicationInfos = applicationInfos;

    setWindowTitle(Messages.LocalApplicationWizard_title);

    System.out.println("LOAW: !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!");
	System.out.println("LOAW: LocalApplicationWizard(applicationInfos) WARNING: SKIPPED FOR DEBUGGING: setDefaultPageImageDescriptor()");
	System.out.println("LOAW: This actually makes the Einstellun - NOAText mod by js - Define - code work without throwing an error.");
	System.out.println("LOAW: !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!");
	//js setDefaultPageImageDescriptor(NOAUIPluginImages.getImageDescriptor(NOAUIPluginImages.IMG_WIZBAN_APPLICATION));
  }

  //----------------------------------------------------------------------------  
  /**
   * Sets home path to be edited.
   * 
   * @param homePath home path to be edited
   * 
   * @author Andreas Br�ker
   */
  public void setHomePath(String homePath) {
	  System.out.println("LOAW: setHomePath to "+homePath);
		this.homePath = homePath;
  }

  //----------------------------------------------------------------------------
  /**
   * Returns selected home path of an local office application. Returns null
   * if a home path is not available.
   * 
   * @return selected home path of an local office application or null
   * if a home path is not available
   * 
   * @author Andreas Br�ker
   */
  public String getSelectedHomePath() {
	  System.out.println("LOAW: getSelectedHomePath");
	  if (localApplicationWizardDefinePage != null)
      return localApplicationWizardDefinePage.getSelectedHomePath();
    return null;
  }

  //----------------------------------------------------------------------------
  /**
   * Performs any actions appropriate in response to the user 
   * having pressed the Finish button, or refuse if finishing
   * now is not permitted.
   *
   * @return <code>true</code> to indicate the finish request
   *   was accepted, and <code>false</code> to indicate
   *   that the finish request was refused
   * 
   * @author Andreas Br�ker
   */
  public boolean performFinish() {
	System.out.println("LOAW: performFinish");
	if (localApplicationWizardDefinePage.getSelectedHomePath() != null)
      return true;
    return false;
  }

  //----------------------------------------------------------------------------
  /**
   * Adds any last-minute pages to this wizard.
   * 
   * @author Andreas Br�ker
   */
  public void addPages() {
	System.out.println("LOAW: addPages");
	localApplicationWizardDefinePage = new LocalApplicationWizardDefinePage(homePath,
        applicationInfos);
    addPage(localApplicationWizardDefinePage);
  }
  //----------------------------------------------------------------------------

}
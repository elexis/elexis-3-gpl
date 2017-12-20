/*******************************************************************************
 * Copyright (c) 2007, G. Weirich and Elexis
 * All rights reserved. This program and the accompanying materials
 * are made available under the terms of the Eclipse Public License v1.0
 * which accompanies this distribution, and is available at
 * http://www.eclipse.org/legal/epl-v10.html
 *
 * Contributors:
 *    G. Weirich - initial implementation
 *    
 *  $Id$
 *******************************************************************************/

package com.jsigle.msword_js;

import javax.print.PrintService;
import javax.print.PrintServiceLookup;

import ch.rgw.tools.Log;
import ch.rgw.tools.ExHandler;

import com.sun.star.beans.XPropertySet;
import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.container.XNameAccess;
import com.sun.star.container.XNameContainer;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.style.XStyle;
import com.sun.star.style.XStyleFamiliesSupplier;
import com.sun.star.text.XText;
import com.sun.star.text.XTextCursor;
import com.sun.star.text.XTextDocument;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import com.sun.star.view.PrintJobEvent;
import com.sun.star.view.PrintableState;
import com.sun.star.view.XPrintJobListener;

/**
 * Some helpers for printing
 * @author Gerry
 *
 */
public class MSWord_jsPrinter {
	private com.sun.star.frame.XComponentLoader xCompLoader = null;
	XComponentContext xContext;
	XMultiComponentFactory xMCF;
	static Log log=Log.get("MSWord_jsPrinter");
	
	public boolean init(){
		//ToDo: Replace this by MSWord_js compatible stuff or remove
		System.out.println("MSWord_jsPrinter.MSWord_jsPrinter: init() begins...");
		
		System.out.println("MSWord_jsPrinter.MSWord_jsPrinter: TODO: **********************************************************");
		System.out.println("MSWord_jsPrinter.MSWord_jsPrinter: TODO: trying to get xCompLoader context");
		System.out.println("MSWord_jsPrinter.MSWord_jsPrinter: TODO: - replace this by MSWord_js compatible stuff or remove it.");
		System.out.println("MSWord_jsPrinter.MSWord_jsPrinter: TODO: **********************************************************");
		try{
			xContext=Bootstrap.bootstrap();
			xMCF=xContext.getServiceManager();
			if(xMCF!=null){
				xCompLoader = (XComponentLoader) UnoRuntime.queryInterface(XComponentLoader.class, 
						xMCF.createInstanceWithContext("com.sun.star.frame.Desktop",xContext));
				System.out.println("MSWord_jsPrinter.MSWord_jsPrinter: init() about to return (xCompLoader!=null)");
				return (xCompLoader!=null);
			}
			System.out.println("MSWord_jsPrinter.MSWord_jsPrinter: init() about to return false");
			return false;
		}catch(Exception ex){
			System.out.println("MSWord_jsPrinter.MSWord_jsPrinter: Caught exception, about to ExHandler.handle(ex)...");
			ExHandler.handle(ex);
			System.out.println("MSWord_jsPrinter.MSWord_jsPrinter: ...ExHandler completed; init() about to return false");
			return false;
		}
	}
	
	
	public static boolean setPrinterTray(XTextDocument doc, String tray) throws Exception{
		System.out.println("MSWord_jsPrinter.setPrinterTray() begins...");

		XText xText = doc.getText();
		XTextCursor cr = xText.createTextCursor();
	
		//ToDo: Replace by something MSWord_js compatible or remove.
		System.out.println("MSWord_jsPrinter.setPrinterTray(): TODO: Replace by something MSWord_js compatible or remove.");
		System.out.println("MSWord_jsPrinter.setPrinterTray(): About to PropertySet...");
		
		XPropertySet xTextCursorProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, cr);
		
		String pageStyleName = xTextCursorProps.getPropertyValue("PageStyleName").toString();
		
//		Get the StyleFamiliesSupplier interface of the document
		XStyleFamiliesSupplier xSupplier = (XStyleFamiliesSupplier) UnoRuntime.queryInterface(XStyleFamiliesSupplier.class, doc);
//		Use the StyleFamiliesSupplier interface to get the XNameAccess interface of the
//		actual style families
		XNameAccess xFamilies = (XNameAccess) UnoRuntime.queryInterface(XNameAccess.class, xSupplier.getStyleFamilies());
//		Access the 'PageStyles' Family
		XNameContainer xFamily = (XNameContainer) UnoRuntime.queryInterface(XNameContainer.class, xFamilies.getByName("PageStyles"));
		
		XStyle xStyle = (XStyle) UnoRuntime.queryInterface(XStyle.class, xFamily.getByName(pageStyleName));
//		Get the property set of the cell's TextCursor
		XPropertySet xStyleProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xStyle);
//		my PageStyleSetting ...
		try{
			System.out.println("MSWord_jsPrinter.setPrinterTray(): trying: About to xStypleProps.setPropertyValue(\"PrinterPaperTray\", tray)...");
			xStyleProps.setPropertyValue("PrinterPaperTray", tray);
			System.out.println("MSWord_jsPrinter.setPrinterTray(): xStyleProps completed, about to return true");
			return true;
		}catch(Exception ex){
			System.out.println("MSWord_jsPrinter.setPrinterTray(): xStyleProps exception caught");
			String possible=(String)xStyleProps.getPropertyValue("PrinterPaperTray");
			log.log("Could not set Tray to "+tray+" try "+possible, Log.ERRORS);
			System.out.println("MSWord_jsPrinter.setPrinterTray(): xStyleProps exception: Could not set Tray, about to return false");
			return false;
		}
	}
	
	 public static boolean checkExistsPrinter(String printer) {
		System.out.println("MSWord_jsPrinter.checkExistsPrinter(): begins...");
			
		boolean exists = true;
		
//		Look up all services
		PrintService[] services = PrintServiceLookup.lookupPrintServices(null, null);
		for (int i = 0; i < services.length; i++) {
			if (services[i].getName().trim().equals(printer.trim())) {
				System.out.println("MSWord_jsPrinter.checkExistsPrinter(): about to return exists (=true)");
				return exists;
			}
		}
		System.out.println("MSWord_jsPrinter.checkExistsPrinter(): about to return !exists (=false)");
		return !exists;
	}
	
	
	static class MyXPrintJobListener implements XPrintJobListener {
		private PrintableState status = null;
		public PrintableState getStatus() {
			return status;
		}
		
		public void setStatus(PrintableState status) {
			this.status = status;
		}
		
		/**
		 * The print job event: has to be called when the action is triggered.
		 */
		public void printJobEvent(PrintJobEvent printJobEvent) {
			if(printJobEvent.State == PrintableState.JOB_COMPLETED)
			{
				System.out.println("JOB_COMPLETED");
				this.setStatus(PrintableState.JOB_COMPLETED);
			}
			if(printJobEvent.State == PrintableState.JOB_ABORTED)
			{
				System.out.println("JOB_ABORTED");
				this.setStatus(PrintableState.JOB_ABORTED);
			}
			if(printJobEvent.State == PrintableState.JOB_FAILED)
			{
				System.out.println("JOB_FAILED");
				this.setStatus(PrintableState.JOB_FAILED);
				return;
			}
			if(printJobEvent.State == PrintableState.JOB_SPOOLED)
			{
				System.out.println("JOB_SPOOLED");
				this.setStatus(PrintableState.JOB_SPOOLED);
			}
			if(printJobEvent.State == PrintableState.JOB_SPOOLING_FAILED)
			{
				System.out.println("JOB_SPOOLING_FAILED");
				this.setStatus(PrintableState.JOB_SPOOLING_FAILED);
				return;
			}
			if(printJobEvent.State == PrintableState.JOB_STARTED)
			{
				System.out.println("JOB_STARTED");
				this.setStatus(PrintableState.JOB_STARTED);
				return;
			}
		}
		
		/**
		 * Disposing event: ignore.
		 */
		public void disposing(com.sun.star.lang.EventObject eventObject) {
			System.out.println("MSWord_jsPrinter.MyXPrintJobListener.disposing(): doing nothing but System.out.println.disposing");
			System.out.println("MSWord_jsPrinter.MyXPrintJobListener.disposing(): TODO: *******************************************");
			System.out.println("MSWord_jsPrinter.MyXPrintJobListener.disposing(): TODO: MAYBE WE SHOULD CLOSE THE WORD WINDOW HERE?");
			System.out.println("MSWord_jsPrinter.MyXPrintJobListener.disposing(): TODO: *******************************************");
			
			System.out.println("disposing");
		}
	}
	
}


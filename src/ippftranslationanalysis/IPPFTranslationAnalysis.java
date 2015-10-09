/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package ippftranslationanalysis;

import com.sun.star.uno.UnoRuntime;

import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.XPropertySet;

import com.sun.star.frame.XComponentLoader;

import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.lang.XMultiComponentFactory;

import com.sun.star.uno.XComponentContext;


import com.sun.star.sheet.XCellRangeAddressable;
import com.sun.star.sheet.XSpreadsheet;
import com.sun.star.sheet.XSpreadsheets;
import com.sun.star.sheet.XSpreadsheetDocument;

/**
 *
 * @author yehster
 */
public class IPPFTranslationAnalysis {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        // TODO code application logic here
        XComponentContext xContext = null;
        
        try {
            xContext = com.sun.star.comp.helper.Bootstrap.bootstrap();
            System.out.println("Connected to a running office ...");
        } catch( Exception e) {
            e.printStackTrace(System.err);
            System.exit(1);
        }

        XSpreadsheetDocument openemr_dutch = null;
        XSpreadsheetDocument ippf_spanish = null;
        
//        XCell oCell = null;

        System.out.println("Opening OpenEMR Translations");
        openemr_dutch = openCalc(xContext,"file:///c:\\Users\\yehster\\Downloads\\openemr_language_table.xlsx");  
        System.out.println("Opening IPPF Spanish Translations");
        ippf_spanish = openCalc(xContext,"file:///c:\\Users\\yehster\\Downloads\\OEMR_SpanishTranslation_MASTER_CURRENTXG.xls");  
    }

    
    public static XSpreadsheetDocument openCalc(XComponentContext xContext,String url)
    {
        //define variables
        XMultiComponentFactory xMCF = null;
        XComponentLoader xCLoader;
        XSpreadsheetDocument xSpreadSheetDoc = null;
        XComponent xComp = null;

        try {
            // get the servie manager rom the office
            xMCF = xContext.getServiceManager();

            // create a new instance of the desktop
            Object oDesktop = xMCF.createInstanceWithContext(
                "com.sun.star.frame.Desktop", xContext );

            // query the desktop object for the XComponentLoader
            xCLoader = UnoRuntime.queryInterface(
                XComponentLoader.class, oDesktop );

            PropertyValue [] szEmptyArgs = new PropertyValue [0];
            String strDoc = "private:factory/scalc";

            xComp = xCLoader.loadComponentFromURL(url, "_blank", 0, szEmptyArgs );
            xSpreadSheetDoc = UnoRuntime.queryInterface(
                XSpreadsheetDocument.class, xComp);

        } catch(Exception e){
            System.err.println(" Exception " + e);
            e.printStackTrace(System.err);
        }

        return xSpreadSheetDoc;
    }    
}

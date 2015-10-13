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

import com.sun.star.frame.XStorable;

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
        openemr_dutch = openCalc(xContext,"file:///c:\\Users\\yehster\\java\\IPPFTranslationAnalysis\\data\\openemr_language_table.xlsx");  
        String spanish_file_name="file:///c:\\Users\\yehster\\java\\IPPFTranslationAnalysis\\data\\OEMR_Translations_MASTER_CURRENTXG.xls";
//        spanish_file_name="file:///c:\\Users\\yehster\\java\\IPPFTranslationAnalysis\\data\\OEMR_SpanishTranslation_DEDUPE2.xls";
        System.out.println("Opening IPPF Spanish Translations");
        ippf_spanish = openCalc(xContext,spanish_file_name);  
        String spanish_first_sheet=((ippf_spanish.getSheets().getElementNames())[0]);
        String dutch_first_sheet=((openemr_dutch.getSheets().getElementNames())[0]);
        XSpreadsheet spanish_sheet=null;
        XSpreadsheet dutch_sheet=null;
        try
        {
            spanish_sheet=UnoRuntime.queryInterface(
                XSpreadsheet.class,
                ippf_spanish.getSheets().getByName(spanish_first_sheet));
            dutch_sheet=UnoRuntime.queryInterface(
                XSpreadsheet.class,
                openemr_dutch.getSheets().getByName(dutch_first_sheet));

            SpreadsheetScannerIPPF scanner = new SpreadsheetScannerIPPF();
           scanner.scan(true,spanish_sheet,1,1,2,5,null,0);
            
           scanner.scan(false,dutch_sheet,5,1,5,0,spanish_sheet,2);
//           scanner.scan(false,dutch_sheet,5,1,7,0,spanish_sheet,8);
           
           System.out.println("done!");
        }
        catch(Exception e)
        {
            e.printStackTrace();
        }
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

/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package ippftranslationanalysis;

import com.sun.star.sheet.XSpreadsheet;
import com.sun.star.table.XCell;

/**
 *
 * @author yehster
 */
public class SpreadsheetScannerOEMR {
    protected XSpreadsheet mSpreadsheet = null;
    public SpreadsheetScannerOEMR(XSpreadsheet spreadsheet)
    {
        mSpreadsheet = spreadsheet;
        
    }
    
    public void scan() throws Exception
    {
        int row=0;
        Boolean cont=true;
        while(cont)
        {
            XCell curCell=mSpreadsheet.getCellByPosition(0, row);
            
            System.out.println(Integer.toString(row)+":"+curCell.getType().toString());
            if((row>0)&& (curCell.getType().equals(com.sun.star.table.CellContentType.EMPTY)))
            {
                cont=false;
            }
            row++;
        }
    }
}

/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package ippftranslationanalysis;
import com.sun.star.uno.UnoRuntime;

import com.sun.star.sheet.XSpreadsheet;
import com.sun.star.table.XCell;
import com.sun.star.table.XTableRows;
import com.sun.star.table.XTable;
import com.sun.star.container.XElementAccess;
import com.sun.star.table.XColumnRowRange;
import java.util.HashMap;
import java.util.ArrayList;

/**
 *
 * @author yehster
 */
public class SpreadsheetScannerIPPF {
    public SpreadsheetScannerIPPF()
    {
        
    }
    protected HashMap<String,TranslationData> mMappedData = new HashMap();
    protected ArrayList<TranslationData> mDataList = new ArrayList();
    
    public void scan(boolean initialize, XSpreadsheet sheet,int StartingRow,int EnglishColumn,int ForeignColumn, int NotesColumn) throws Exception
    {
        int row=StartingRow;
        Boolean cont=true;
        TranslationData curTranslationData=null;
        XColumnRowRange table=null;
        try
        {
           table = UnoRuntime.queryInterface(XColumnRowRange.class, sheet.getSpreadsheet());
            
        }
        catch(Exception e)
        {
            e.printStackTrace();
        }
        String foreignData="";
        int unmatched = 0;
        int duplicates=0;
        while(cont)
        {
            XCell curCell=sheet.getCellByPosition(0, row);
            
            if(!curCell.getType().equals(com.sun.star.table.CellContentType.EMPTY))
            {
                XCell English_Cell=sheet.getCellByPosition(EnglishColumn, row);
                String EnglishData=English_Cell.getFormula();
                XCell Language_Cell=null;
                if(ForeignColumn!=0)
                {
                    Language_Cell=sheet.getCellByPosition(ForeignColumn, row);
                    foreignData=Language_Cell.getFormula();
                }
                XCell Notes_Cell=sheet.getCellByPosition(NotesColumn, row);
                if(initialize)
                {
                    if(English_Cell.getType().equals(com.sun.star.table.CellContentType.EMPTY))
                    {
                        if(curCell.getValue()<5000)
                        {
                            System.out.println("Removing Empty content at:"+curCell.getValue());
                            table.getRows().removeByIndex(row, 1);
                            XCell counterCell=sheet.getCellByPosition(0,row);
                            counterCell.setFormula("=A"+Integer.toString(row)+"+1");
                            row--;
                            table = UnoRuntime.queryInterface(XColumnRowRange.class, sheet.getSpreadsheet());                        
                            
                        }
                        else
                        {
                            System.out.println("Found Empty content at:"+curCell.getValue());
                        }
                        
                    }
                    curTranslationData=new TranslationData(
                        (int)row,
                        EnglishData,
                        foreignData,
                        Notes_Cell.getFormula()
                        );
                    TranslationData existingData = mMappedData.get(EnglishData);
                    if(existingData==null)
                    {
                        mMappedData.put(EnglishData, curTranslationData);
                        mDataList.add(curTranslationData);         
                    }
                    else
                    {
                        if(!Language_Cell.getType().equals(com.sun.star.table.CellContentType.EMPTY))
                                {
                                    duplicates++;
                                    System.out.println("Duplicate constant:"+":"+Notes_Cell.getFormula()+":"+row+":"+EnglishData+":"+foreignData.equals(existingData.getForeign())+":"+foreignData+"||"+existingData.toString());
                                    XCell prevForeign =sheet.getCellByPosition(ForeignColumn,existingData.getID());
                                    prevForeign.setFormula(foreignData);
                                    XCell prevNote = sheet.getCellByPosition(NotesColumn,existingData.getID());
                                    prevNote.setFormula(prevNote.getFormula() + "|"+Notes_Cell.getFormula());
                                    table.getRows().removeByIndex(row, 1);
                                    XCell counterCell=sheet.getCellByPosition(0,row);
                                    counterCell.setFormula("=A"+Integer.toString(row)+"+1");
                                    row--;
                                    table = UnoRuntime.queryInterface(XColumnRowRange.class, sheet.getSpreadsheet());                        
                                    
                                }
                    }
                    
                }
                else
                {
                    curTranslationData=mMappedData.get(EnglishData);
                    if(curTranslationData==null)
                    {
                        System.out.println(foreignData +"|     |" +EnglishData+" has no match");
                        unmatched++;
                    }
                    else
                    {
                        curTranslationData.setForeign(foreignData);
//                        System.out.println(curTranslationData.toString());
                    }
                    
                }
                        
            }
            if((row>0)&& (curCell.getType().equals(com.sun.star.table.CellContentType.EMPTY)))
            {

                cont=false;
            }
            row++;
        }
        System.out.println("Duplicates :"+Integer.toString(duplicates));
        System.out.println("Unmatched :"+Integer.toString(unmatched));
    }
}

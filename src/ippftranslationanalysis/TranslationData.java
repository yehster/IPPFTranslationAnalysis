/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package ippftranslationanalysis;

/**
 *
 * @author yehster
 */
public class TranslationData {
    String mEnglish;
    String mForeign;
    String mNotes;
    int mID;
    
    public TranslationData(int ID,String English, String Foreign, String Notes)
    {
        this.mEnglish=English;
        this.mForeign=Foreign;
        this.mNotes=Notes;
        this.mID=ID;
    }
    
    public String toString()
    {
        return Integer.toString(this.mID)+":"+this.mEnglish+":"+this.mForeign+":"+this.mNotes;
    }
    
    public void setForeign(String foreign)
    {
        this.mForeign=foreign;
    }
    
    public String getForeign()
    {
        return this.mForeign;
    }
    
    public int getID()
    {
        return this.mID;
    }
}

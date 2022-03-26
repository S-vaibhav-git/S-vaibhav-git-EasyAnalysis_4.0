/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.project.rough_work;

import com.project.clipboard_handling.ClipboardHandling;
import com.project.easyanalysis.ExcelForm;
import com.project.easyanalysis.Exp1;
import com.project.easyanalysis.HomeFrame;
import static com.project.easyanalysis.HomeFrame.tcSheetEval;
import com.project.utilities.PropertiesFile;
import com.project.utilities.AlertHandling;
import java.awt.datatransfer.UnsupportedFlavorException;
import java.io.IOException;
import java.text.ParseException;

import java.util.logging.Level;
import java.util.logging.Logger;

/**
 *
 * @author TechAnim
 *   
    1. capture copied TC num
    2. compare to excel sheet (tcNumMatched)
    3. if found in previous execution, return all row values
    4. set all values to the excel form class
    5. populate all values to the excel form frame
 * 
 */

public class EventListener2 implements Runnable {
    
 //   public Thread t1;
    public static int start_flag=1;     // a is start flag for 
    int status=0;
    private static String clipboardValue;
    private static boolean tcNumMatched; 
    private Object[] excelValues;
    
    public EventListener2() {
    //    t1=new Thread(this);        
    }
    
    public void run() {
                          
        while(true){
            status=Exp1.listener();
            if(status==EventListener2.start_flag){    // in this if conditoin catches most synchrinized start_flag value
                try { 
                    
                    // 1. capture copied TC num
                        clipboardValue=ClipboardHandling.readClipboard();
                        System.out.println(clipboardValue);                    
                    
                 /*       synchronized (Thread.currentThread()){
                            try{
                              Thread.currentThread().wait(3000);
                            } catch (InterruptedException e) {
                            }
                         }
                */
                } catch (UnsupportedFlavorException | IOException ex) {
                    Logger.getLogger(EventListener2.class.getName()).log(Level.SEVERE, null, ex);
                }
            }             
        }  
    }    
}
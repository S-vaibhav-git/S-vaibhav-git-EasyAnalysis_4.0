/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.project.rough_work;

import com.project.clipboard_handling.ClipboardHandling;
import com.project.easyanalysis.Exp1;
import java.awt.datatransfer.UnsupportedFlavorException;
import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;

/**
 *
 * @author TechAnim
 */
public class test2 {    
     
    public static volatile int start_flag=1;     // a is start flag for 
    int status=0;
    private static String clipboardValue;
    private static boolean tcNumMatched; 
    private Object[] excelValues;
    
    
        public synchronized void callListener(){
            
            while(true){
                int status = Exp1.listener();
                if(status==EventListener3.start_flag){    // in this if conditoin catches most synchrinized start_flag value
                    try { 

                        // 1. capture copied TC num
                        String clipboardValue = ClipboardHandling.readClipboard();
                            System.out.println(clipboardValue);

                    } catch (UnsupportedFlavorException | IOException ex) {
                        Logger.getLogger(EventListener3.class.getName()).log(Level.SEVERE, null, ex);
                    }
                }             
            }  
         }
    }    


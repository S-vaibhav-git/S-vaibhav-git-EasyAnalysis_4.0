/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.project.listeners;

import com.project.clipboard_handling.ClipboardHandling;
import com.project.easyanalysis.ExcelForm;
import com.project.easyanalysis.Exp1;
import com.project.easyanalysis.HomeFrame;
import static com.project.easyanalysis.HomeFrame.tcSheetEval;
import com.project.utilities.PropertiesFile;
import com.project.excel_handling.TestCaseSheetEvaluation;
import com.project.utilities.AlertHandling;
import java.awt.datatransfer.UnsupportedFlavorException;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

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

public class EventListener implements Runnable {
    
    public Thread t1;
    public static volatile int start_flag=1;     // a is start flag for 
    int status=0;
    private static String clipboardValue;
    private static boolean tcNumMatched; 
    private Object[] excelValues;
    
    public EventListener() {
        t1=new Thread(this);        
    }
    
    public void tcFind(){
        
        System.out.println(Thread.currentThread().getName());
        while(true){
            status=Exp1.listener();
            if(status==start_flag){    // in this if conditoin catches most synchrinized start_flag value
                System.out.println("Entered in IF condition EventListener class");
                // 1. capture copied TC num
                
                try{
                    clipboardValue=ClipboardHandling.readClipboard();
                }catch(Exception e){
                    AlertHandling.homeFrameAlerts(PropertiesFile.alertMsg.getProperty("clipboardWorking_failed"));
                }
                System.out.println(clipboardValue);  
                start_flag=0;
                
                /*
                //  Test below code, also comment code from point 2 below then test below code
                
                HomeFrame.setVisibility(false);
                ExcelForm.setVisibility(true);
                tcNumMatched(clipboardValue);
                */
                //2. compare to excel sheet (tcNumMatched)
                //make different class- TestCaseSheetEvaluation with method searchTestCase in excel_handling package to find tc in excel
                tcNumMatched=tcSheetEval.searchTestCase(clipboardValue);
                //3. if found in previous execution, return all row values to here
                if(tcNumMatched){
                    excelValues=tcSheetEval.readExcelData();
                }
                HomeFrame.setVisibility(false);
                //4. set all values to the excel form class, if tc not found then set alert value
                ExcelForm.setVisibility(true);          //ExcelForm is initialized through the setVisibility method
                if(tcNumMatched){
                    ExcelForm.setExcelFormValues(excelValues);
                    //5. populate all values to the excel form frame
                    try {
                        ExcelForm.populateExcelFormValues(tcNumMatched);
                    } catch (ParseException ex) {
                        Logger.getLogger(EventListener.class.getName()).log(Level.SEVERE, null, ex);
                    }
                    //ExcelForm.alertMessage.setText(PropertiesFile.alertMsg.getProperty("TC_found"));
                    AlertHandling.excelFrameAlerts(PropertiesFile.alertMsg.getProperty("TC_found"));
                }else{
                    //ExcelForm.alertMessage.setText(PropertiesFile.alertMsg.getProperty("TC_notfound"));
                    //when tc does not match populate all other data empty and keep test case number and current date in the excel form
                    /*  ExcelForm.tc_num_value=clipboardValue;
                    ExcelForm.date_value=new SimpleDateFormat("dd/MM/yy").format(new Date());
                    ExcelForm.populateExcelFormValues(! tcNumMatched);
                    
                    instead this chunk of code we can below method call also
                    */
                    ExcelForm.noTCFound_populateDefaultValues(clipboardValue);
                    // ExcelForm.date.setDate(new Date());
                    AlertHandling.excelFrameAlerts(PropertiesFile.alertMsg.getProperty("TC_notfound"));
                }
            }             
        } 
    }
     
    @Override
    public void run() {
         tcFind();
    }    
}
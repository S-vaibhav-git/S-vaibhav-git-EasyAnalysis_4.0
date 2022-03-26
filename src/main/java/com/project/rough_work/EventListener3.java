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
import com.project.excel_handling.TestCaseSheetEvaluation;
import com.project.utilities.AlertHandling;
import java.awt.datatransfer.UnsupportedFlavorException;
import java.io.IOException;
import java.text.ParseException;

import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.bcel.generic.AALOAD;

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

public class EventListener3 extends Thread {
    
   // public Thread t1;
    public static volatile int start_flag=1;     // a is start flag for 
    int status=0;
    private static String clipboardValue;
    private static boolean tcNumMatched; 
    private Object[] excelValues;
    test2 test;
    public EventListener3() {
            test=new test2();
    }
     
   
    @Override
    public void run() {
        test.callListener();
    }    
}
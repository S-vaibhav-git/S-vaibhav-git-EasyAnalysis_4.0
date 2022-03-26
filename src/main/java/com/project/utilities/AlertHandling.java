/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.project.utilities;

import com.project.easyanalysis.ExcelForm;
import com.project.easyanalysis.HomeFrame;

/**
 *
 * @author TechAnim
 */
public class AlertHandling {
    
    public static void homeFrameAlerts(String AlertMsg){
        HomeFrame.alertMessage.setText("Alert: "+AlertMsg);
    }
    public static void excelFrameAlerts(String AlertMsg){
        ExcelForm.alertMessage.setText("Alert: "+AlertMsg);    
    }
    
}

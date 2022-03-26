/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.project.excel_handling;

import com.project.utilities.AlertHandling;
import com.project.utilities.PropertiesFile;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.commons.io.FileUtils;

/**
 *
 * @author TechAnim
 */
public class DataSheetValidation {
    
    public static final String DataSheetPath= PropertiesFile.ProjectPath+PropertiesFile.filePath.getProperty("datasheetpath");
    private static final String BasicDataSheetPath= PropertiesFile.ProjectPath+PropertiesFile.filePath.getProperty("basicdatasheetpath");
    private static File dataSheet;
    private static File source;
    private static File dest;
    
    public static void validateSheet() throws FileNotFoundException{        
        dataSheet=new File(DataSheetPath);
        if(dataSheet.isFile()){             //if exist() s used, it wont allow to open or read the file
            System.out.println("Given File Exist");            
            AlertHandling.homeFrameAlerts(PropertiesFile.alertMsg.getProperty("DataSheetFile_exist"));
        }
        else{
            System.out.println("File not present");
            AlertHandling.homeFrameAlerts(PropertiesFile.alertMsg.getProperty("DataSheetFile_notExist"));
            DataSheetValidation.createBasicDataSheet();
        }      
    }
    
    private static void createBasicDataSheet(){        
        try {
            if(dataSheet.createNewFile()){
                source=new File(BasicDataSheetPath);
                dest=new File(DataSheetPath);
                FileUtils.copyFile(source, dest);
                AlertHandling.homeFrameAlerts(PropertiesFile.alertMsg.getProperty("DataSheetFile_created"));
            }
        } 
        catch (IOException ex) {
            Logger.getLogger(DataSheetValidation.class.getName()).log(Level.SEVERE, null, ex);
            AlertHandling.homeFrameAlerts(PropertiesFile.alertMsg.getProperty("DataSheetFile_notCreated"));
        }
    }
    
}

/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.project.utilities;

import com.project.easyanalysis.HomeFrame;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Properties;
import java.util.logging.Level;
import java.util.logging.Logger;
/**
 *
 * @author TechAnim
 */
public class PropertiesFile {
    
    public static Properties alertMsg,filePath;     // referenced in below method
    public static final String ProjectPath=System.getProperty("user.dir");          // referenced in below method
    private static final String AlertMsgFile_Path="\\src\\main\\java\\com\\project\\resource\\AlertMessages.properties";        // referenced in below method
    private static final String FilePaths_Path="\\src\\main\\java\\com\\project\\resource\\FilePaths.properties";               // referenced in below method 
    private static FileInputStream fis;
    
    public static void loadPropertiesFile() throws IOException{
         //load alert message properties
        alertMsg=new Properties();
        try {
            fis=new FileInputStream(ProjectPath+AlertMsgFile_Path);
        } catch (FileNotFoundException ex) {
            Logger.getLogger(HomeFrame.class.getName()).log(Level.SEVERE, null, ex);
        }
        alertMsg.load(fis);
        
        //load filepaths properties
        filePath=new Properties();
         filePath=new Properties();
        try {
            fis=new FileInputStream(ProjectPath+FilePaths_Path);
        } catch (FileNotFoundException ex) {
            Logger.getLogger(HomeFrame.class.getName()).log(Level.SEVERE, null, ex);
        }
        filePath.load(fis);       
    }
    
}

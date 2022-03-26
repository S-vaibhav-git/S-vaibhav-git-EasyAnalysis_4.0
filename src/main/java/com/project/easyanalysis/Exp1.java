/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.project.easyanalysis;

import com.project.utilities.PropertiesFile;

/**
 *
 * @author TechAnim
 */
public class Exp1 {    
    static{
        
        System.load(System.getProperty("user.dir")+"\\src\\main\\java\\com\\project\\resource\\keylistener.dll");             //dll file path
    }
    public static native int listener();  
}

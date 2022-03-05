/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.project.listeners;

import com.project.clipboard_handling.ClipboardHandling;
import com.project.easyanalysis.Exp1;
import com.project.easyanalysis.HomeFrame;
import java.awt.datatransfer.UnsupportedFlavorException;
import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;

/**
 *
 * @author TechAnim
 */
public class EventListener implements Runnable {
    
    public Thread t1;
    public int start_flag=1;     // a is stop flag
    int status=0;
  
    public EventListener() {
        t1=new Thread(this);
    }
    
    @Override
    public void run() {
        
        while(true){     
            status=Exp1.listener();
            if(status==this.start_flag){
                try {                    
                    System.out.println(ClipboardHandling.readClipboard());
                } catch (UnsupportedFlavorException | IOException ex) {
                    Logger.getLogger(EventListener.class.getName()).log(Level.SEVERE, null, ex);
                }
            }
        }  
    }
    
}
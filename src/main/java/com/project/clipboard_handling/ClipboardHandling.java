/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.project.clipboard_handling;

/**
  * @author TechAnim
 */
  
import java.awt.Toolkit;
import java.awt.datatransfer.Clipboard;
import java.awt.datatransfer.DataFlavor;
import java.awt.datatransfer.Transferable;
import java.awt.datatransfer.UnsupportedFlavorException;
import java.io.IOException;
 
public class ClipboardHandling {
    
    static String data;  
    public static String readClipboard() throws UnsupportedFlavorException, IOException{
    
        Clipboard clipboard = Toolkit.getDefaultToolkit().getSystemClipboard();
        Transferable contents = clipboard.getContents(null);
        boolean hasTransferableText = (contents != null) && contents.isDataFlavorSupported(DataFlavor.stringFlavor);
        if (hasTransferableText) {
                String result = "";
                result = (String) contents.getTransferData(DataFlavor.stringFlavor);
                return result;
        }
        return "Copy Again";        
    }
    
}

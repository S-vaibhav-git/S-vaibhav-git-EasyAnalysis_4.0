/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.project.rough_work;

import com.project.excel_handling.DataSheetValidation;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.bcel.generic.AALOAD;
import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author TechAnim
 */
public class test1{
       // public static String dest_path;        
        public static String src_path; 
	public static FileInputStream fis_src=null;    
       // public static FileInputStream fis_dest=null;
	public static FileOutputStream fos=null;	 
       // private static XSSFWorkbook workbook_dest=null;
        private static XSSFWorkbook workbook_src=null;
       // private static XSSFSheet sheet_dest=null;
        private static XSSFSheet sheet_src=null;
      //private static XSSFSheet sheet_dest_2;
	private static XSSFRow row_src=null;
       // private static XSSFRow row_dest=null;
        private static XSSFCell cell_src =null;
	//private static XSSFCell cell_dest =null;
	         
    public static void main(String[] args) {
        
        src_path="H:\\Test Folder\\Basic Data Sheet.xlsx";
     //   dest_path="H:\\Test Folder\\Analysis Data Sheet.xlsx";      
           
        try {
            //source input stream
            fis_src=new FileInputStream(src_path);                
            workbook_src=new XSSFWorkbook(fis_src);
            fis_src.close();    

//            //dest input stream
//            fis_dest=new FileInputStream(dest_path);
//            workbook_dest=new XSSFWorkbook(fis_dest); 
//            fis_dest.close();                                
        }
        catch(Exception e){
                e.printStackTrace();
        } 
        
        //Source sheet setup
        sheet_src=workbook_src.getSheetAt(0);
        row_src=sheet_src.getRow(0);
        int colCount=row_src.getLastCellNum();
        System.out.println("Column count: "+colCount);
        XSSFCellStyle base_style=workbook_src.createCellStyle();        
        workbook_src.createSheet("Demo2");
        workbook_src.getSheet("Demo2").createRow(0);
        
        //Dest sheet setup
//        workbook_dest.createSheet("17-03-22_QA");
//        sheet_dest=workbook_dest.getSheet("17-03-22_QA");     
//        if(sheet_dest.getRow(0)==null){
//             sheet_dest.createRow(0);
//        } 
//        row_dest=sheet_dest.getRow(0);
//        XSSFCellStyle set_style=workbook_dest.createCellStyle();       
       
        for(int i=0;i<colCount;i++){ 
            cell_src=row_src.getCell(i);
            base_style=cell_src.getCellStyle();
            
            if(workbook_src.getSheet("Demo2").getRow(0).getCell(i) == null){
                workbook_src.getSheet("Demo2").getRow(0).createCell(i); 
            }
            workbook_src.getSheet("Demo2").getRow(0).getCell(i).setCellStyle(base_style);
            workbook_src.getSheet("Demo2").getRow(0).getCell(i).setCellValue(cell_src.getStringCellValue());
            System.out.println("Cell style of src for Column: "+i+" is: "+base_style.toString());
//            if(row_dest.getCell(i)==null){
//                row_dest.createCell(i);
//            }                   
//            set_style.cloneStyleFrom(base_style);
//            
//            cell_dest=row_dest.getCell(i);
//            cell_dest.setCellStyle(set_style);
//            System.out.println("Cell style of Dest for Column: "+i+" is: "+cell_dest.getCellStyle().toString());
            
        } 
        
        try {
//                fos=new FileOutputStream(dest_path);
//                workbook_dest.write(fos);
//                fos.close();
                
                fos=new FileOutputStream(src_path);
                workbook_src.write(fos);                
                fos.close();
            } catch (FileNotFoundException ex) {
                Logger.getLogger(test1.class.getName()).log(Level.SEVERE, null, ex);
            } catch (IOException ex) {
                Logger.getLogger(test1.class.getName()).log(Level.SEVERE, null, ex);
            }
       
    }
}

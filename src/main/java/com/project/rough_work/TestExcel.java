/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.project.rough_work;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author TechAnim
 */
public class TestExcel {
    
    
    public static String path;
	public FileInputStream fis=null;
	public FileOutputStream fos=null;
	private XSSFWorkbook workbook=null;
	private XSSFSheet sheet=null;
	private XSSFRow row=null;
	private XSSFCell cell =null;
	
	private int rowCount;
	private int colCount;
	private Object data=new Object();
        
        public int getRowCount(String sheetName) {		
		int index=workbook.getSheetIndex(sheetName);
		if(index == -1) {
			return 0;
		}
		sheet=workbook.getSheetAt(index);
		int endRow=sheet.getLastRowNum();
		rowCount=endRow+1;
		return rowCount;
	}
	
	// Return the number of columns in sheet
	public int getColumnCount(String sheetName) {
		sheet=workbook.getSheet(sheetName);
		row=sheet.getRow(0);
		colCount = row.getLastCellNum();		
	
		return colCount;
	}
        
        public TestExcel(String path){
            this.path=path;		
		try {
			fis=new FileInputStream(path);
			workbook=new XSSFWorkbook(fis);
			sheet=workbook.getSheetAt(0);
			fis.close();
		}
		catch(Exception e){
			e.printStackTrace();
		}
        }
        
        public static void main(String args[]){
            
            TestExcel excel=new TestExcel("H:\\Test Folder\\Analysis Data Sheet.xlsx");
            System.out.println(excel.getRowCount("13-03-22_Stg"));
            
        }
    
}

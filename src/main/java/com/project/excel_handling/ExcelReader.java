package com.project.excel_handling;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Iterator;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCreationHelper;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


/**
 * 
 * @author TechAnim
 * 
 * following are the methods are basically used in ExcelReader class
 * 1.	getRowCount(String)		D
 * 2.	getColCount(String)		D
 * 
 * 3.   getCellData(String, String, int)		D
 * 4.   getCellData(String, int, int)			D
 * 5. 	setCellData(String, String, int, String)	D	 
 * 
 * 6.   createSheet(String)		D
 * 7.   removeSheet(String)		D
 * 
 * 8.   addColumn(String, String)	D
 * 9.   removeColumn(String, int)	D
 * 
 * 10.	isSheetExist(String)		D
 * 11.	getCellRowNum(String, String, String)	D	
 *
 */

public class ExcelReader {
	
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
	
	public ExcelReader(String path)
	{
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

	// Return the number of rows in sheet
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
	
	public boolean isSheetExist(String sheetName) {		
		int index=workbook.getSheetIndex(sheetName);
		if(index==-1) {			
			index=workbook.getSheetIndex(sheetName.toUpperCase());
			if(index==-1) {
				return false;
			}
			else
				return true;
		}
		else
			return true;		
	}
	
	public Object getCellData(String sheetName, int colNum, int rowNum) {		 
		sheet=workbook.getSheet(sheetName);		 
		row=sheet.getRow(rowNum);		 	
		cell=row.getCell(colNum);
                
		switch(cell.getCellType()) {		
			case STRING: data=cell.getStringCellValue();	// OR data[r-1][c].toString();
			 break;					
						
			case NUMERIC: data=cell.getNumericCellValue();  //OR data[r-1][c]=Integer.parseInt((String) data[r-1][c]);
			 break;				
			
			case BOOLEAN:  data=cell.getBooleanCellValue();
			 break;				
			
			default:
				break;
		}		
		return data;
	}
	
	public Object getCellData(String sheetName, String colName, int rowNum) {
		
		Boolean exit=false;
		sheet=workbook.getSheet(sheetName);
		row=sheet.getRow(sheet.getFirstRowNum());
		int i=0;
		String colName_Obtained;
		int colNum=0;
		while(!exit) {
			cell=row.getCell(i);
			colName_Obtained = cell.getStringCellValue();
			if(colName_Obtained.equals(colName)) {
				colNum=cell.getColumnIndex();
				exit=true;
			}
			i++;
		}		
		row=sheet.getRow(rowNum);
		cell=row.getCell(colNum);
		switch(cell.getCellType()) {

                    case STRING:	data=cell.getStringCellValue(); break;
                    case NUMERIC:	data=cell.getNumericCellValue(); break;
                    case BOOLEAN:	data=cell.getBooleanCellValue(); break;

                    default:
                            break;
		}	
		return data;
		
	}
	
        public boolean setCellData( String sheetName, String colName, int rowNum1, String data) {  //need to add Strig url parameter as well
		
		int colNum=0;		
		try {	
			//get sheet
			sheet = workbook.getSheet(sheetName);
			if(sheet == null) {
				return false;
			}			
			// find the column number from the given colName
			row=sheet.getRow(0);
			int totalCols=row.getLastCellNum();
			 
			for(int i=0; i< totalCols;i++) {				
                            if(row.getCell(i) != null) {					
                                if(colName.equalsIgnoreCase((String) this.getCellData(sheetName, i, 0))) {
                                        colNum = i;						
                                        //test colNum value
                                        System.out.println("This is the colNumn "+colNum);

                                        //test rowNum
                                        System.out.println("This is the rowNum "+row.getRowNum());

                                        //test column name 
                                        System.out.println("This is the ColName at "+colNum+", "+row.getRowNum()+"--> "+this.getCellData(sheetName, i, 0).toString());						
                                        break;
                                }
                                else 
                                        if(i == totalCols - 1) {
                                        System.out.println("Sheet not present");
                                        return false;
                                }
                            }				
			}
			
			// set data to the cell
			
                        row=sheet.getRow(rowNum1);
                        //if row not present then create the row.
                        if(row == null) {
                                row = sheet.createRow(rowNum1);
                                System.out.println("New row created");
                        }
                        cell=row.getCell(colNum);
                        //if cell not present or empty data then create the cell.
                        if(cell == null) {
                                cell=row.createCell(colNum);
                        }			
                        cell.setCellValue(data);
			 
			//write data to the workbook
			fos=new FileOutputStream(path);
			workbook.write(fos);;
			fos.close();
			
			if(this.getCellData(sheetName, colNum, rowNum1).toString().equalsIgnoreCase(data) ) {				
				//test newly set cell data 
				System.out.println("Newly set cell data : "+this.getCellData(sheetName, colNum, rowNum1).toString());
				return true;
			}
			else {
				//test already present cell data 
				System.out.println("Previous cell data (not changed): "+this.getCellData(sheetName, colNum, rowNum1).toString());
				return false;
			}			 
		}
		catch(Exception e) {
			e.printStackTrace();
			return false;
		}	 	
	}

	public boolean createSheet(String sheetName) {		
		try {			
			fos=new FileOutputStream(path);
			if(!isSheetExist(sheetName)) {
				workbook.createSheet(sheetName);
				workbook.write(fos);
				fos.close();
				return true;
			}
			else {
				return false;
			}			
		}
		catch(Exception e) {
			e.printStackTrace();
			return false;
		}
		
	}
	
	public boolean removeSheet(String sheetName) {
		try {
			fos=new FileOutputStream(path);
			int sheetIndex=workbook.getSheetIndex(sheetName);
			workbook.removeSheetAt(sheetIndex);
			workbook.write(fos);
			fos.close();
			return true;
		}
		catch (Exception e) {
			// TODO: handle exception
			e.printStackTrace();
			return false;
		}
	}
		
	public boolean addColumn(String sheetName, String colName) {	
		try {		
			// get row
			sheet=workbook.getSheet(sheetName);
			row=sheet.getRow(0);			
			if(row==null) {
				row=sheet.createRow(0);
			}			
			// get Cell 
			if(row.getLastCellNum() == -1) {
				
				cell = row.createCell(row.getRowNum());
			}else {
				
				cell = row.createCell(row.getLastCellNum());
			}
			
			cell.setCellValue(colName);
			
			FileOutputStream fos= new FileOutputStream(path);
			workbook.write(fos);
			fos.close();
			
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		}		
		return true;
	}
	
	public boolean removeColumn(String sheetName, String colName) {
		try {			
			sheet = workbook.getSheet(sheetName);
			
			//get ColNum from given colName
			row=sheet.getRow(0);
			
			int totalCols=row.getLastCellNum();
			int colNum = 0;
			
			for(int i=0; i< totalCols;i++) {
				cell=row.getCell(i);
				if(cell != null) {
					if(colName.equalsIgnoreCase(this.getCellData(sheetName, i, 0).toString())) {
						colNum = i;
						break;
					}
					else if (i+1 == totalCols) {
						System.out.println(this.getCellData(sheetName, i, 0).toString());
						return false;
					}
				}				
			}
			
			//remove cells from the given sheet
			int rowCount=sheet.getLastRowNum();
			
			for(int i=0; i< rowCount; i++) {
				
				row=sheet.getRow(i);
				if(row != null) {					
					cell=row.getCell(colNum);
					if(cell != null) {
						row.removeCell(cell);
					}
				}
			}
			
			FileOutputStream fos=new FileOutputStream(path);
			workbook.write(fos);
			fos.close();
			 
		} catch (Exception e) {
			// TODO: handle exception
			e.printStackTrace();
			return false;
		}
		
		return true;
	}

	public int getCellRowNum(String sheetName, String colName, String colValue) {
		int rowNum=0,colNum=0;
		try {
			sheet=workbook.getSheet(sheetName);
			row=sheet.getRow(0);
			int totalCols=row.getLastCellNum();
			for(int i=0;i<totalCols;i++) {
				if(row.getCell(i)!=null) {
					if(colName.equalsIgnoreCase(this.getCellData(sheetName, i, 0).toString())) {
						colNum=i;
						break;
					}
					else if (i+1 == totalCols) {
						return -1;
					}
				}
			}
			
			int rowCount= this.getRowCount(sheetName);
			for (int i = 0; i < rowCount ; i++) {
				row=sheet.getRow(i);
				if(row!=null && row.getCell(colNum)!=null) {
					if(colValue.equalsIgnoreCase(this.getCellData(sheetName, colNum, i).toString())) {
						rowNum=i;
						break;
					}
					else if (i+1 == rowCount) {
						return -1;
					}
				}
			}
			
		} catch (Exception e) {
			// TODO: handle exception
			e.printStackTrace();
			return -1;
		}
		return rowNum;
	}
}

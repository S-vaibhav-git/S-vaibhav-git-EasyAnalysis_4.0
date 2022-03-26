/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.project.excel_handling;

import com.project.easyanalysis.ExcelForm;
import static com.project.easyanalysis.ExcelForm.ADD_BASIC_DATA_FAILED;
import static com.project.easyanalysis.ExcelForm.CREATE_NEW_SHEET_FAILED;
import static com.project.easyanalysis.ExcelForm.TC_ADDED_TO_EXISTING_SHEET;
import static com.project.easyanalysis.ExcelForm.TC_ADDED_TO_EXISTING_SHEET_FAILED;
import static com.project.easyanalysis.ExcelForm.TC_ADDED_TO_NEW_SHEET;
import static com.project.easyanalysis.ExcelForm.TC_ADDED_TO_NEW_SHEET_FAILED;
import static com.project.easyanalysis.ExcelForm.TC_DATA_WRITE_FAILED;
import static com.project.easyanalysis.ExcelForm.TC_UPDATED;
import static com.project.easyanalysis.ExcelForm.TC_UPDATE_FAILED;
import static com.project.easyanalysis.HomeFrame.tcSheetEval;
import com.project.utilities.AlertHandling;
import com.project.utilities.PropertiesFile;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author TechAnim
 * 
 * 1. get row data
 * 2. sheetexist
 * 3. createColumn
 * 4. formatColumn
 * 5. selectSheet
 * 6. select
 */
public class TestCaseSheetEvaluation {
    
    private static String sheetName;
      
    //Excel essentials
    public String path;
    public FileInputStream fis=null;
    public FileOutputStream fos=null;
    private XSSFWorkbook workbook=null;
    private XSSFSheet sheet=null;
    private XSSFRow row=null;
    private XSSFCell cell =null;

    private int rowCount,rowNum;
    private int colCount,colNum;
    private Object data=new Object();
    
    public TestCaseSheetEvaluation(){
        //this.path=path;	
        path=DataSheetValidation.DataSheetPath;
        try {
                fis=new FileInputStream(path);
                workbook=new XSSFWorkbook(fis);
               // sheet=workbook.getSheetAt(0);
                fis.close();
        }
        catch(Exception e){
                e.printStackTrace();
        }
    } 
    
    public boolean searchTestCase(String tc_num){
        //logic here:
        
        int sheetsCount=tcSheetEval.getTotalSheetsCount();
        System.out.println("Sheet Count: "+sheetsCount);
        boolean checkStatus=false;
        for (int i = sheetsCount; i >= 0; i--){
             sheet=workbook.getSheetAt(i);
             System.out.println("Sheet selected at: "+i+" is "+ sheet.getSheetName());
             checkStatus=findCellDataExist(tc_num);
             if(checkStatus){
                 break;
             }
        }         
        //if found set sheetName and RowNum values, then return true else return false        
        return checkStatus;
    }
    
    public boolean findCellDataExist(String tc_num){
        rowCount=getRowCount();   
        System.out.println("Row Count "+rowCount);
        for(int i=1;i<rowCount;i++){
            row=sheet.getRow(i);        //we can reuse this row object
            System.out.println(row.getSheet().getSheetName()+" "+Thread.currentThread().getName() );            
            cell=row.getCell(0);        // finding in tc num Column only so passing 0 to get cell
         //   System.out.println(String.valueOf((int)cell.getNumericCellValue())+"     -- >   "+tc_num.trim());
         
            //this is done in this way because stored values of TC Num can be string or numeric, so doing try catch will not break the code for any string or numeric value of the tc_num
            boolean cellValueValid;
            try{
                cellValueValid=String.valueOf((int)cell.getNumericCellValue()).trim().equalsIgnoreCase(tc_num.trim());
                System.out.println("Test case matched: "+cellValueValid);
            }
            catch(IllegalStateException e){
               //  e.printStackTrace();
                 cellValueValid=cell.getStringCellValue().trim().equalsIgnoreCase(tc_num.trim());
                 System.out.println("Test case matched: "+cellValueValid);
            }
            catch(NullPointerException e_Nptr){
                //e_Nptr.printStackTrace();
                AlertHandling.homeFrameAlerts(PropertiesFile.alertMsg.getProperty("basicSheetData_notFound"));
                return false;
            }
            
            if( cellValueValid ){
                rowNum=i; 
                row=sheet.getRow(rowNum);
                System.out.println("Data found at: "+rowNum);
                return true;
            }
        }
        return false;
    }    
    
    public int getRowCount(){
        int endRow=sheet.getPhysicalNumberOfRows();
        System.out.println(endRow);       
        return endRow;
    }
    
    public int getColumnCount() {
        row=sheet.getRow(0);
        colCount = row.getLastCellNum();
        return colCount;
    }
    
    public int getTotalSheetsCount(){
         int i=workbook.getNumberOfSheets();
        return i-1;
    }
    
    public String search_flag;
    public boolean searchTestCase(String tc_num, String sheetName_mod){
        //logic here: finding tc in the sheet of previous date of  given date
        //int sheetsCount=tcSheetEval.getTotalSheetsCount();
        boolean checkStatus=false;
        System.out.println("Just passed Previous Sheet Name: "+sheetName_mod);
        sheet=workbook.getSheet(sheetName_mod);
        
        //to handle null pointer exception
        if(sheet==null){
            return false;
        }
         System.out.println("After set sheet object, Previous Sheet Name: "+sheet.getSheetName());
        //boolean sheetFound=searchSheet(date.replace("/", "-"));
        while(!checkStatus){
            checkStatus=findCellDataExist(tc_num);
            if(!checkStatus){
                if(search_flag.equals(ExcelForm.PREVIOUS)){
                    sheetName_mod=tcSheetEval.getPreviousSheetName(sheetName_mod);
                }
                else if(search_flag.equals(ExcelForm.NEXT)){
                     sheetName_mod=tcSheetEval.getNextSheetName(sheetName_mod);
                }               
            }
            if(sheetName_mod.equals("No_Previous_Sheet") || sheetName_mod.equals("No_Next_Sheet")){
                checkStatus=false;
                break;
            }
        }                  
        //if found set sheetName and RowNum values, then return true else return false        
        return checkStatus;       
    }
    
    String sheet_Name[];
    String date;
    String execution;
    public Object[] excelValues;      
    public Object[] readExcelData(){
        //logic here
        colCount=getColumnCount();
        excelValues=new Object[colCount+2];  
        colNum=0;
        row=sheet.getRow(rowNum);
        //once found the TC then read the row from already set values of sheetName and RowNum and convert last modification string value to the boolean value and return all data
        System.out.println("Col count: "+colCount);
        int i;
        for(i=0;i<colCount;i++){
            cell=row.getCell(i);            
           // excelValues[i]=cell.getStringCellValue();
            switch(cell.getCellType()) {
                    case STRING:	excelValues[i]=cell.getStringCellValue(); 
                                        System.out.println(cell.getStringCellValue());
                                        break;
                    case NUMERIC:	excelValues[i]=String.valueOf((int)cell.getNumericCellValue()); 
                                        System.out.println(String.valueOf((int)cell.getNumericCellValue()));
                                        break;
                    case BOOLEAN:	excelValues[i]=cell.getBooleanCellValue(); 
                                        System.out.println(cell.getBooleanCellValue());
                                        break;
                                        
                    default:            break;                            
		}
        }  
        //set last two values from the sheet name
        
        sheet_Name=row.getSheet().getSheetName().split("_", 2);
        date=sheet_Name[0].replace("-", "/");
        System.out.println("Date value read from excel: "+date);
        
        execution=sheet_Name[1];
        excelValues[i]=date;        //second from the last is date
        excelValues[i+1]=execution; //last value is execution type qa/stg
        
        return excelValues; 
    }
    
    public boolean writeExcelData(Object[] values){
        //logic here
        //all required variables for sheet data
        String tc_num;
        String testSuite;
        String execution;
        String date;
        String comment;
        String defect;
        boolean modification;
        
        //other required variables
        String sheet_name;
        boolean sheet_found;
        boolean tc_found;
        boolean row_empty;
        boolean dataWriteDone;
        boolean newSheetCreated;
        boolean basicDataCreated;
        
        //set values to these required variables        
        tc_num=(String)values[0];
        testSuite=(String) values[1];       
        comment=(String) values[2];
        defect=(String) values[3];
        modification=(boolean)values[4];
        date=(String) values[5];              
        System.out.println("Date value passed writeexceldata method: "+date);
        execution=(String) values[6];       
        
        //firstly we need to find sheet for the date and execution mentioned in the excel form
        // 1. format sheet name from given date and execution                
       
        sheet_name=getFormattedSheetName(date,execution);
        System.out.println("Sheet Name formated :" +sheet_name);
        
        sheet_found=searchSheet(sheet_name);
        
        // 2. if sheet found then overwrite the row data to that particular row otherwise create new sheet with the sheetname and feed the data
        if(sheet_found){
            tc_found=findCellDataExist(tc_num);
            if(tc_found){
               dataWriteDone= writeRowData(values);
               if(dataWriteDone){
                   System.out.println("Data written successfully");         //add alert message
                   ExcelForm.writeData_AlertFlag=TC_UPDATED;
                   return true;
               }
               else{
                  ExcelForm.writeData_AlertFlag=TC_UPDATE_FAILED;
               }
            }
            else 
                if(!tc_found){      // No need to add else part since this if is just for everification of value of tc_found
                    //then add data to the new empty row, 
                    row_empty=getLastEmptyRow(sheet_name);
                    if(row_empty){      // No need to write else part or alert msg in the else part of this if part, as row empty is returning only true value
                        dataWriteDone=writeRowData(values);
                        if(dataWriteDone){
                            //add alert value
                            ExcelForm.writeData_AlertFlag=TC_ADDED_TO_EXISTING_SHEET;
                            return true;
                        }
                        else{
                            ExcelForm.writeData_AlertFlag=TC_ADDED_TO_EXISTING_SHEET_FAILED;
                        }
                    }                
                }
        }else{
            System.out.println("Need to create sheet, Sheet not found: "+sheet_name);
            newSheetCreated = createNewSheet(sheet_name);
            if(newSheetCreated){
                basicDataCreated = createBasicData_NewSheet(sheet); 
                if(basicDataCreated){
                    rowNum=1;
                    row=sheet.getRow(rowNum);
                    dataWriteDone=writeRowData(values);
                    if(dataWriteDone){
                        //add alert value here
                        ExcelForm.writeData_AlertFlag=TC_ADDED_TO_NEW_SHEET;
                        return true;
                    }
                    else{
                        ExcelForm.writeData_AlertFlag=TC_ADDED_TO_NEW_SHEET_FAILED;
                    }
                }else{
                    ExcelForm.writeData_AlertFlag=ADD_BASIC_DATA_FAILED;
                }                 
            }else{
                 ExcelForm.writeData_AlertFlag=CREATE_NEW_SHEET_FAILED;
            }            
        }        
        return false;
    }  
    
    public boolean createNewSheet(String sheet_name){
       
        try {
            workbook.createSheet(sheet_name);
        } catch (Exception e) {
           // e.printStackTrace();
            return false;
        }
        sheet=workbook.getSheet(sheet_name);  
        return true;
    }
    
    public boolean createBasicData_NewSheet(XSSFSheet new_sheet){
        
        //basic sheet setup
        sheet=workbook.getSheetAt(1);
        row=sheet.getRow(0);
        int col_count=row.getLastCellNum();
        System.out.println("Column count: "+col_count);
        
        //new sheet setup
        if(new_sheet.getRow(0)==null){
            new_sheet.createRow(0);
        }
        
        XSSFCellStyle base_style=workbook.createCellStyle();          
        
        for(int i=0;i<colCount;i++){ 
             
            //fetch data and style of basic sheet 
            cell=row.getCell(i);
            base_style=cell.getCellStyle();           
            
            //set data for new sheet
            if(new_sheet.getRow(0).getCell(i) == null){
                new_sheet.getRow(0).createCell(i); 
            }
            new_sheet.getRow(0).getCell(i).setCellStyle(base_style);
            new_sheet.getRow(0).getCell(i).setCellValue(cell.getStringCellValue());
            System.out.println("Cell value from sheet 0: "+cell.getStringCellValue());
            new_sheet.setColumnWidth(i,4500 );
         }
         
        try {        
           fos=new FileOutputStream(path);
           workbook.write(fos);                
           fos.close(); 
           
           //re-provide the previous value sof object sheet
           sheet=new_sheet;
           return true;
           
        } catch (IOException ex) {
            Logger.getLogger(TestCaseSheetEvaluation.class.getName()).log(Level.SEVERE, null, ex);
            ex.printStackTrace();
            //re-provide the previous value sof object sheet
            sheet=new_sheet;
            return false;
        }
        
    }
    
    public boolean writeRowData(Object[] values){
        int colNum=0;		
        try {	
            //get sheet
            //sheet = workbook.getSheet(sheetName);
            if(sheet == null) {
                System.out.println("Sheet value NULL in writeRowData method");
                return false;                    
            }			
            // find the column number from the given colName
            //row=sheet.getRow(0);
            int totalCols=getColumnCount();

            //write all data to respective cells
            System.out.println("Value set to the rowNum variable: "+rowNum);
            row=sheet.getRow(rowNum);
            if(row==null){
                System.out.println("Row object is Null in writeRowData method");
                sheet.createRow(rowNum);
                row=sheet.getRow(rowNum);
               // return false; 
            }
            for(int i=0; i< totalCols;i++) {	
               
                 cell=row.getCell(i); 
                 if(cell==null){
                    System.out.println("Cell object is Null in for loop writeRowData method");
                    row.createCell(i);
                    cell=row.getCell(i);
                    //return false;
                 }
                 if(values[i] instanceof String){
                      cell.setCellValue((String)values[i]);
                      System.out.println("value:"+(String)values[i]+" set at: "+row.getRowNum()+" and column: "+i);
                 }else if(values[i] instanceof Boolean){
                      cell.setCellValue((boolean)values[i]);
                      System.out.println("value:"+(boolean)values[i]+"  set at: "+row.getRowNum()+" and column: "+i);
                 }else if(values[i] instanceof Integer){
                      cell.setCellValue((Integer)values[i]);
                      System.out.println("value:"+(Integer)values[i]+"  set at: "+row.getRowNum()+" and column: "+i);
                 }                               
            }	

            //write data to the workbook
            System.out.println("path passed to fos: "+path);
            fos=new FileOutputStream(path);
            workbook.write(fos);
            fos.close(); 
            return true;              
        }
        catch(IOException e) {
            e.printStackTrace();
            return false;
        }	 	
    }
    
    public boolean getLastEmptyRow(String sheet_name){
       // String tc_num="not null";
       //write proper logic to find next empty row
        int i=0;   
        //while(!(tc_num==null) && !(tc_num.equals("")) ){     //if tc_num is not empty continue the loop till it gets empty value
            rowCount=getRowCount();
            rowNum=rowCount;
            row=sheet.getRow(rowNum);
//            try {
//                row=sheet.getRow(++i);
//                cell=row.getCell(0);
//                System.out.println(" for row index: "+i+" TC Col values: "+tc_num);
//            } catch (NullPointerException e) {
//                System.out.println("Null pointer exception");
//                e.printStackTrace();                
//            }          
//            
//            try{
//                tc_num=String.valueOf(cell.getNumericCellValue());
//            }
//            catch(Exception e){                  
//                try {
//                    tc_num= cell.getStringCellValue();
//                } catch (Exception ex) {
//                    return false;
//                }
//            }            
       // }
//        rowNum=i;
//        row=sheet.getRow(rowNum);
        
        return true;
    }
    
    public boolean createCurrentSheet(String sheet_name){           //need to complete furhter steps
        
      //  excelreader.createSheet(renameCurrentSheet(sheetName, sheetName))
        try {			
            fos=new FileOutputStream(path);
            if(!searchSheet(sheet_name)) {
                workbook.createSheet(sheet_name);
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
    
    public String[] splitSheetName(){
     /*   sheet_Name=row.getSheet().getSheetName().split("_", 2);
        sheet_Name[0]=sheet_Name[0].replace("-", "/");       
     
        return sheet_Name;
        */
        return sheet_Name;
    }
    
    public String getFormattedSheetName(String date, String execution){
        String modifiedDate=date.replace("/", "-");
        sheetName=modifiedDate+"_"+execution;
        return sheetName;
    }
    
    public boolean searchSheet(String sheet_name){        
        
        int sheetsCount=tcSheetEval.getTotalSheetsCount();
        System.out.println("Sheet Count: "+sheetsCount);
        //boolean checkStatus=false;
        for (int i = sheetsCount; i >= 0; i--){
            sheet=workbook.getSheetAt(i);
            //checkStatus=findCellDataExist(tc_num);
            if(sheet.getSheetName().contains(sheet_name)){
                System.out.println("Sheet found in the excel: "+sheetName);
                return true;
            }
        } 
         return false;      
    }   
    
    public String getPreviousSheetName(String sheetName){
        int sheetIndex=workbook.getSheetIndex(sheetName);
        if((sheetIndex-1) >= 0){
            sheet=workbook.getSheetAt(sheetIndex-1);
        }
        else{
            return "No_Previous_Sheet";
        }
        sheetName=sheet.getSheetName();         
        return sheetName;
    }
    
    public String getNextSheetName(String sheetName){
        int sheetIndex=workbook.getSheetIndex(sheetName);
        int sheetCount=getTotalSheetsCount();
        if((sheetIndex+1) <= sheetCount){
            sheet=workbook.getSheetAt(sheetIndex+1);
        } 
        else{
            return "No_Next_Sheet";
        }
        sheetName=sheet.getSheetName();         
        return sheetName;
    }
}

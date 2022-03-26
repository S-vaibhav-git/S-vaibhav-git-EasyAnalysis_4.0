/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/GUIForms/JFrame.java to edit this template
 */
package com.project.easyanalysis;


import static com.project.easyanalysis.HomeFrame.tcSheetEval;
import com.project.utilities.PropertiesFile;
import com.project.listeners.EventListener;
import com.project.utilities.AlertHandling;
import java.awt.Font;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JOptionPane;
 
import javax.swing.ScrollPaneConstants;
/**
 *
 * @author TechAnim
 */
public class ExcelForm extends javax.swing.JFrame {
    /**
     * Creates new form ExcelForm 
     */        
    public static String tc_num_value;
    private static String testSuite_value;
    private static String execution_value;
    public static String date_value;
    private static String comment_value;
    private static String defect_value;
    private static boolean modification_value;
    public final static String PREVIOUS="previous";
    public final static String NEXT="next";

    //public static ExcelForm excelform;
    public ExcelForm() {
        initComponents(); 
        //set horizontal and veritical scroll bar hidden
        commentScrollPane.setHorizontalScrollBarPolicy(ScrollPaneConstants.HORIZONTAL_SCROLLBAR_NEVER);
        commentScrollPane.setVerticalScrollBarPolicy(ScrollPaneConstants.VERTICAL_SCROLLBAR_NEVER);
        
        //set label font
        defaultLabelFont();
        //By default all text field uneditable
        defaultTextFieldEditable(false);   
        EventListener.start_flag=0; 
        
//        this.addWindowListener(new WindowAdapter(){
//                   
//                });        
    }    
    
    static Object[] values=new Object[7];   
  /*  public static Object[] getExcelFormValues(){        
        values=excelform.extractExcelFormValues();
        return values;
    }
  */  
    public static void populateExcelFormValues(boolean tcNumMatched) throws ParseException{
        if(tcNumMatched){
            excelform.populateExcelFormValues();
        }        
    }
    
    //this method used to set local class values, but not used to set values in the frame
    public static void setExcelFormValues(Object[] excelValues){
        tc_num_value=(String)excelValues[0];
        testSuite_value=(String) excelValues[1];       
        comment_value=(String) excelValues[2];
        defect_value=(String) excelValues[3];
        modification_value=(boolean)excelValues[4];
        date_value=(String) excelValues[5];             //.....date needs to come from the sheet name in excel sheet
        System.out.println("Date value set in Excel form class: "+ExcelForm.date_value);
        
        execution_value=(String) excelValues[6];       //execution type needs to come from the sheet name in excel sheet
    }
    
    private void extractExcelFormValues(){        
        values[0]=TC_Num.getText();
        values[1]=testSuite.getText();        
        values[2]=comment.getText();
        values[3]=defect.getText();
        values[4]=modification.isSelected(); 
        //validate date
       // Date =date.getDate().toString();
        values[5]=new SimpleDateFormat("dd/MM/yy").format(date.getDate());          // date comming from the excel form is in full date format with time, so need to convert it to sinple date format
        
        values[6]=execution.getText();
       // return values;
    }
    
    private void populateExcelFormValues() throws ParseException{
        TC_Num.setText(tc_num_value);
        testSuite.setText(testSuite_value);
        execution.setText(execution_value);
        date.setDate(new SimpleDateFormat("dd/MM/yy").parse(date_value));
        System.out.println("Date value from populateExcelForm class: "+date_value);
        comment.setText(comment_value);
        defect.setText(defect_value);
        modification.setSelected(modification_value);
    }
    
    private void defaultLabelFont(){
        Font font=new Font("Calibri", Font.BOLD, 15);
        TC_Num_Label.setFont(font);
        execution_Label.setFont(font);
        comment_Label.setFont(font);
        defect_Label.setFont(font);
        modification_Label.setFont(font);
    }    
     
    private static boolean editCheckbox;        //associated with below method //used in mouse listener action for edit checkbox to handle default enability of checkbox 
    private void defaultTextFieldEditable(boolean flag){
        TC_Num.setEditable(false);
        testSuite.setEditable(flag);
        execution.setEditable(flag);         
        date.setEditable(flag);             
        comment.setEditable(flag);
        defect.setEditable(flag);        
        //set checkbox read only
         editCheckbox=flag;
        /*  OR
            modification.setSelected(true);
            modification.setEnabled(false);
        */        
        //disable save button by default, enable only after click on edit button
        
        ExcelForm.saveButton.setEnabled(flag);
        /*    try {
                populateExcelFormValues();
            } catch (ParseException ex) {
                Logger.getLogger(ExcelForm.class.getName()).log(Level.SEVERE, null, ex);
            }
        */
    }
    public static void noTCFound_populateDefaultValues(String tc_num){
        try {
            tc_num_value=tc_num;
            date_value=new SimpleDateFormat("dd/MM/yy").format(new Date());
            excelform.populateExcelFormValues();
            excelform.defaultTextFieldEditable(false);
        } catch (ParseException ex) {
            Logger.getLogger(ExcelForm.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
    private void setDefaultVal_TxtFieldsAndVariables(){
       
        //set text variables to null when we close the excelform window
        tc_num_value=null;
        testSuite_value=null;
        comment_value=null;
        defect_value=null;
        modification_value=false;
    //    date_value=new SimpleDateFormat("dd/MM/yyyy").toString();       //this is set from the listener itself    
        execution_value=null;       
                
       //set text textfields to null when we close the excelform window
        TC_Num.setText("");
        testSuite.setText("");
        execution.setText("");
    //  date.setDate(new Date());           //this is set from the listener itself    
        comment.setText("");
        defect.setText("");
        modification.setSelected(false);  
    }
    
    public static ExcelForm excelform=null;     //associated with below method
    public static void setVisibility(boolean visibility){
        if(excelform == null){
                excelform=new ExcelForm();
                excelform.addWindowListener(new WindowAdapter() {
                    @Override
                    public void windowClosing(WindowEvent e) {                     
                        super.windowClosing(e);              //default on_close activity is hide, which hides this excel window on click on close button in window
                        excelform.setDefaultVal_TxtFieldsAndVariables();   //on before every closing set text fields to the empty values
                        EventListener.start_flag=1;
                        HomeFrame.setVisibility(true);
                    }    
                   
                    @Override
                    public void windowClosed(WindowEvent e) {
                             
                          // setDefaultVal_TxtFieldsAndVariables();                           
                    }                  
                });
        }       
        if(visibility){    
           //  EventListener.start_flag=0;
             excelform.setVisible(visibility);
             excelform.setAlwaysOnTop(true);            
        }
        else{
            excelform.setVisible(visibility);
            excelform=null;
        }
    }

    /**
     * This method is called from within the constructor to initialize the form.
     * WARNING: Do NOT modify this code. The content of this method is always
     * regenerated by the Form Editor.
     */
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        TC_Num_Label = new javax.swing.JLabel();
        execution_Label = new javax.swing.JLabel();
        comment_Label = new javax.swing.JLabel();
        defect_Label = new javax.swing.JLabel();
        modification_Label = new javax.swing.JLabel();
        modification = new javax.swing.JCheckBox();
        TC_Num = new javax.swing.JTextField();
        execution = new javax.swing.JTextField();
        defect = new javax.swing.JTextField();
        javax.swing.JButton editButton = new javax.swing.JButton();
        commentScrollPane = new javax.swing.JScrollPane();
        comment = new javax.swing.JTextArea();
        previousButton = new javax.swing.JButton();
        nextButton = new javax.swing.JButton();
        saveButton = new javax.swing.JButton();
        alertMessage = new javax.swing.JLabel();
        date = new org.jdesktop.swingx.JXDatePicker();
        testSuite = new javax.swing.JTextField();

        setTitle("ExcelForm");

        TC_Num_Label.setText("TC No. / Test Suite");

        execution_Label.setText("Execution / Date");

        comment_Label.setText("Comment");

        defect_Label.setText("Defect");

        modification_Label.setText("Script Modification");

        modification.setText("Need Modification");
        modification.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                modificationMouseClicked(evt);
            }
        });
        modification.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                modificationActionPerformed(evt);
            }
        });

        TC_Num.setFont(new java.awt.Font("Arial", 0, 12)); // NOI18N
        TC_Num.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                TC_NumActionPerformed(evt);
            }
        });

        execution.setFont(new java.awt.Font("Arial", 0, 12)); // NOI18N
        execution.setText("QA / Stg");
        execution.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                executionMouseClicked(evt);
            }
        });
        execution.addInputMethodListener(new java.awt.event.InputMethodListener() {
            public void caretPositionChanged(java.awt.event.InputMethodEvent evt) {
            }
            public void inputMethodTextChanged(java.awt.event.InputMethodEvent evt) {
                executionInputMethodTextChanged(evt);
            }
        });
        execution.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                executionActionPerformed(evt);
            }
        });
        execution.addKeyListener(new java.awt.event.KeyAdapter() {
            public void keyPressed(java.awt.event.KeyEvent evt) {
                executionKeyPressed(evt);
            }
        });

        defect.setFont(new java.awt.Font("Arial", 0, 12)); // NOI18N
        defect.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                defectActionPerformed(evt);
            }
        });

        editButton.setText("Edit");
        editButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                editButtonActionPerformed(evt);
            }
        });

        comment.setColumns(20);
        comment.setFont(new java.awt.Font("Arial", 0, 12)); // NOI18N
        comment.setRows(5);
        commentScrollPane.setViewportView(comment);

        previousButton.setText("Previous");
        previousButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                previousButtonActionPerformed(evt);
            }
        });

        nextButton.setText("Next");
        nextButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                nextButtonActionPerformed(evt);
            }
        });

        saveButton.setText("Save");
        saveButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                saveButtonActionPerformed(evt);
            }
        });
        saveButton.addPropertyChangeListener(new java.beans.PropertyChangeListener() {
            public void propertyChange(java.beans.PropertyChangeEvent evt) {
                saveButtonPropertyChange(evt);
            }
        });

        alertMessage.setForeground(new java.awt.Color(153, 0, 0));
        alertMessage.setText("  ");

        date.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                dateActionPerformed(evt);
            }
        });

        testSuite.setFont(new java.awt.Font("Arial", 0, 12)); // NOI18N
        testSuite.setText("Test Suite Name");
        testSuite.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                testSuiteMouseClicked(evt);
            }
        });
        testSuite.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                testSuiteActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(layout.createSequentialGroup()
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                            .addComponent(TC_Num_Label, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                            .addComponent(testSuite, javax.swing.GroupLayout.PREFERRED_SIZE, 1, Short.MAX_VALUE)
                            .addComponent(TC_Num))
                        .addGap(18, 18, Short.MAX_VALUE)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING, false)
                            .addComponent(date, javax.swing.GroupLayout.DEFAULT_SIZE, 110, Short.MAX_VALUE)
                            .addComponent(execution)
                            .addComponent(execution_Label, javax.swing.GroupLayout.Alignment.LEADING))
                        .addGap(25, 25, 25)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(commentScrollPane, javax.swing.GroupLayout.PREFERRED_SIZE, 169, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(comment_Label)))
                    .addGroup(layout.createSequentialGroup()
                        .addComponent(alertMessage, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(editButton)))
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(layout.createSequentialGroup()
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(saveButton)
                        .addGap(28, 28, 28)
                        .addComponent(previousButton, javax.swing.GroupLayout.PREFERRED_SIZE, 77, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(nextButton, javax.swing.GroupLayout.PREFERRED_SIZE, 74, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addGroup(layout.createSequentialGroup()
                        .addGap(15, 15, 15)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(defect, javax.swing.GroupLayout.PREFERRED_SIZE, 81, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(defect_Label))
                        .addGap(18, 18, 18)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(modification_Label)
                            .addComponent(modification))))
                .addGap(24, 24, 24))
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(TC_Num_Label)
                    .addComponent(execution_Label, javax.swing.GroupLayout.PREFERRED_SIZE, 16, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(comment_Label)
                    .addComponent(defect_Label)
                    .addComponent(modification_Label))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(layout.createSequentialGroup()
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(TC_Num, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(execution))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(testSuite, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(date, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addGap(31, 31, 31))
                    .addGroup(layout.createSequentialGroup()
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(commentScrollPane, javax.swing.GroupLayout.PREFERRED_SIZE, 61, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                                .addComponent(modification, javax.swing.GroupLayout.PREFERRED_SIZE, 16, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addComponent(defect, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)))
                        .addGap(18, 18, Short.MAX_VALUE)))
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(alertMessage, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addComponent(editButton)
                    .addComponent(saveButton)
                    .addComponent(previousButton)
                    .addComponent(nextButton))
                .addContainerGap())
        );

        pack();
    }// </editor-fold>//GEN-END:initComponents

    private void modificationActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_modificationActionPerformed
        // TODO add your handling code here:
         if(!editCheckbox){
            if(modification.isSelected()){
                modification.setSelected(false);
            }else{
                modification.setSelected(true);
            }
        }   
    }//GEN-LAST:event_modificationActionPerformed

    private void TC_NumActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_TC_NumActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_TC_NumActionPerformed

    private void executionActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_executionActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_executionActionPerformed

    private void defectActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_defectActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_defectActionPerformed

    private void editButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_editButtonActionPerformed
        // TODO add your handling code here:
        defaultTextFieldEditable(true);
        
        AlertHandling.excelFrameAlerts(PropertiesFile.alertMsg.getProperty("EditBtn_Clicked"));
        // make the checkbox editable
        //editCheckbox=true;
         /*  OR
            modification.setEnabled(true);
        */
        
         // save button should be enabled
        // excelform.saveButton.setEnabled(true);
    }//GEN-LAST:event_editButtonActionPerformed

    private void nextButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_nextButtonActionPerformed
        // TODO add your handling code here:
         /*
         steps to implement the next button click
         1. find the next sheet, if available then search tc num, else show message no next Execution data available
         2. pass the next sheet date or whole sheet name(for whole sheet name need to modify the parameter)
         3. search tc num in that sheet and if found populate data, if not found print nextTC notFound alert msg
        */         
        System.out.println("Next Button clicked ");
        tcSheetEval.search_flag=NEXT;
        String sheetName_value=tcSheetEval.getFormattedSheetName(date_value, execution_value);
        String nextSheetName_value=tcSheetEval.getNextSheetName(sheetName_value);       
        if(! findTC(nextSheetName_value)){
             AlertHandling.excelFrameAlerts(PropertiesFile.alertMsg.getProperty("NextTC_notfound"));   
        }  
    }//GEN-LAST:event_nextButtonActionPerformed

    public static String writeData_AlertFlag;
    public static final String TC_UPDATED=PropertiesFile.alertMsg.getProperty("TC_Updated");
    public static final String TC_UPDATE_FAILED=PropertiesFile.alertMsg.getProperty("TC_Update_failed");
    public static final String TC_ADDED_TO_EXISTING_SHEET=PropertiesFile.alertMsg.getProperty("TC_AddedToExistingSheet");
    public static final String TC_ADDED_TO_EXISTING_SHEET_FAILED=PropertiesFile.alertMsg.getProperty("TC_AddedToExistingSheet_failed");
    public static final String TC_ADDED_TO_NEW_SHEET=PropertiesFile.alertMsg.getProperty("TC_AddedToNewSheet");
    public static final String TC_ADDED_TO_NEW_SHEET_FAILED=PropertiesFile.alertMsg.getProperty("TC_AddedToNewSheet_failed");
    public static final String TC_DATA_WRITE_FAILED= PropertiesFile.alertMsg.getProperty("TC_DataWrite_failed"); 
    public static final String CREATE_NEW_SHEET_FAILED= PropertiesFile.alertMsg.getProperty("Create_NewSheet_failed"); 
    public static final String ADD_BASIC_DATA_FAILED= PropertiesFile.alertMsg.getProperty("Add_BasicData_failed");
    
    private void saveButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_saveButtonActionPerformed
        // TODO add your handling code here:
        
        //Alert handling for save button if clicked yes then write to excel is happeneing
        int showConfirmDialogValue = JOptionPane.showConfirmDialog(null, "Are sure to Save Entered Details ?","Alert !!",JOptionPane.YES_NO_OPTION);
        boolean writeExcel_status;
         if(showConfirmDialogValue == JOptionPane.YES_OPTION){
            System.out.println("Clicked Yes Option");
            excelform.extractExcelFormValues();
            writeExcel_status=tcSheetEval.writeExcelData(values);           //values is array of the objects
            if(writeExcel_status){
                
                if(writeData_AlertFlag.equals(TC_UPDATED)){
                    AlertHandling.excelFrameAlerts(TC_UPDATED);
                }else if(writeData_AlertFlag.equals(TC_ADDED_TO_EXISTING_SHEET)){
                    AlertHandling.excelFrameAlerts(TC_ADDED_TO_EXISTING_SHEET);
                }else if(writeData_AlertFlag.equals(TC_ADDED_TO_NEW_SHEET)){
                    AlertHandling.excelFrameAlerts(TC_ADDED_TO_NEW_SHEET);
                }              
                
                 //after click on the save button save button and all text fields are not editable and changeable
                excelform.defaultTextFieldEditable(false);
            }else{
                
                if(writeData_AlertFlag.equals(TC_UPDATE_FAILED)){
                    AlertHandling.excelFrameAlerts(TC_UPDATE_FAILED);
                }else if(writeData_AlertFlag.equals(TC_ADDED_TO_EXISTING_SHEET_FAILED)){
                    AlertHandling.excelFrameAlerts(TC_ADDED_TO_EXISTING_SHEET_FAILED);
                }else if(writeData_AlertFlag.equals(TC_ADDED_TO_NEW_SHEET_FAILED)){
                    AlertHandling.excelFrameAlerts(TC_ADDED_TO_NEW_SHEET_FAILED);
                }else if(writeData_AlertFlag.equals(CREATE_NEW_SHEET_FAILED)){
                    AlertHandling.excelFrameAlerts(CREATE_NEW_SHEET_FAILED);
                }                     
                //ExcelForm.alertMessage.setText(PropertiesFile.alertMsg.getProperty("WriteExcel_fail"));
                AlertHandling.excelFrameAlerts(PropertiesFile.alertMsg.getProperty("WriteExcel_fail"));
            } 
         }   
         
    }//GEN-LAST:event_saveButtonActionPerformed

    private void modificationMouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_modificationMouseClicked
        //instead of using setEnabled to the check box used below method
        //if- if condition is true then checkbox will be read only, only when if condition is false checkbox will be 
            
    }//GEN-LAST:event_modificationMouseClicked

    private void executionMouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_executionMouseClicked
        /*if(execution.isEditable()){
            execution.setText("");
        }
        */
    }//GEN-LAST:event_executionMouseClicked

    private void dateActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_dateActionPerformed
        // TODO add your handling code heresd:
    }//GEN-LAST:event_dateActionPerformed

    private void saveButtonPropertyChange(java.beans.PropertyChangeEvent evt) {//GEN-FIRST:event_saveButtonPropertyChange
        // TODO add your handling code here:
    }//GEN-LAST:event_saveButtonPropertyChange

    private void testSuiteActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_testSuiteActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_testSuiteActionPerformed

    private void testSuiteMouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_testSuiteMouseClicked
        // TODO add your handling code here:
        /*
         if(testSuite.isEditable()){
            execution.setText("");
        }
        */
    }//GEN-LAST:event_testSuiteMouseClicked

    private Object excelValues[];
    private void previousButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_previousButtonActionPerformed
        // TODO add your handling code here:
        /*
         steps to implement the previous button click
         1. find the previous sheet, if found print TC found alert msg,if not print  PrevTC_notfound alert msg
         2. pass the previous sheet date or whole sheet name(for whole sheet name need to modify the parameter)
         3. search tc num in that sheet and if found populate data, if not found search in more previous sheets
        */
        //2. compare to excel sheet (tcNumMatched)
        //make different class- TestCaseSheetEvaluation with method searchTestCase in excel_handling package to find tc in excel       
        System.out.println("Previous Button clicked ");
        tcSheetEval.search_flag=PREVIOUS;
        String sheetName_value=tcSheetEval.getFormattedSheetName(date_value, execution_value);
        String prevSheetName_value=tcSheetEval.getPreviousSheetName(sheetName_value);
        if(prevSheetName_value.equals("No_Previous_Sheet")){
            AlertHandling.excelFrameAlerts(PropertiesFile.alertMsg.getProperty("PrevTC_notfound")); 
        }else{
             if(! findTC(prevSheetName_value)){
                AlertHandling.excelFrameAlerts(PropertiesFile.alertMsg.getProperty("PrevTC_notfound"));   
            }   
        }                   
    }//GEN-LAST:event_previousButtonActionPerformed

    public boolean findTC (String sheet_name){
        
        System.out.println("Previous sheet name fetched : "+sheet_name);
        boolean tcNumMatched=tcSheetEval.searchTestCase(tc_num_value, sheet_name);
        
        System.out.println("Test case matched : "+tcNumMatched);

        //3. if found in previous execution, return all row values to here
            if(tcNumMatched){
                excelValues=tcSheetEval.readExcelData();   
        //4. set all values to the excel form class, if tc not found then set alert value             
                ExcelForm.setExcelFormValues(excelValues);   
        //5. populate all values to the excel form frame
                try {
                    ExcelForm.populateExcelFormValues(tcNumMatched);
                } catch (ParseException ex) {
                    Logger.getLogger(EventListener.class.getName()).log(Level.SEVERE, null, ex);
                    AlertHandling.excelFrameAlerts(PropertiesFile.alertMsg.getProperty("PopulateSheetData_Error"));    
                }
                //ExcelForm.alertMessage.setText(PropertiesFile.alertMsg.getProperty("TC_found"));
                 AlertHandling.excelFrameAlerts(PropertiesFile.alertMsg.getProperty("TC_found"));    
            }else{                
                 return false;
            }                                
        return true;
    }
    
    private void executionKeyPressed(java.awt.event.KeyEvent evt) {//GEN-FIRST:event_executionKeyPressed
      
    }//GEN-LAST:event_executionKeyPressed

    private void executionInputMethodTextChanged(java.awt.event.InputMethodEvent evt) {//GEN-FIRST:event_executionInputMethodTextChanged
        // TODO add your handling code here:
    }//GEN-LAST:event_executionInputMethodTextChanged
        
    /**
     * @param args the command line arguments
     */
    public static void main(String args[]) {
        /* Set the Nimbus look and feel */
        //<editor-fold defaultstate="collapsed" desc=" Look and feel setting code (optional) ">
        /* If Nimbus (introduced in Java SE 6) is not available, stay with the default look and feel.
         * For details see http://download.oracle.com/javase/tutorial/uiswing/lookandfeel/plaf.html 
         */
        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Nimbus".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (ClassNotFoundException ex) {
            java.util.logging.Logger.getLogger(ExcelForm.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(ExcelForm.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(ExcelForm.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(ExcelForm.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
               ExcelForm excelform= new ExcelForm();
               excelform.setVisible(true);                        
            }
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JTextField TC_Num;
    private javax.swing.JLabel TC_Num_Label;
    public static javax.swing.JLabel alertMessage;
    private javax.swing.JTextArea comment;
    private javax.swing.JScrollPane commentScrollPane;
    private javax.swing.JLabel comment_Label;
    public static org.jdesktop.swingx.JXDatePicker date;
    private javax.swing.JTextField defect;
    private javax.swing.JLabel defect_Label;
    private javax.swing.JTextField execution;
    private javax.swing.JLabel execution_Label;
    private javax.swing.JCheckBox modification;
    private javax.swing.JLabel modification_Label;
    private javax.swing.JButton nextButton;
    private javax.swing.JButton previousButton;
    private static javax.swing.JButton saveButton;
    private javax.swing.JTextField testSuite;
    // End of variables declaration//GEN-END:variables
}

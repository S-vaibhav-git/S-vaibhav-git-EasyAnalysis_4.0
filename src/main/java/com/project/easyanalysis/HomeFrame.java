/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/GUIForms/JFrame.java to edit this template
 */
package com.project.easyanalysis;

import com.project.utilities.PropertiesFile;
import com.project.excel_handling.DataSheetValidation;
import com.project.excel_handling.TestCaseSheetEvaluation;
import com.project.listeners.EventListener;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Properties;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JOptionPane;

/**
  * @author TechAnim
 */  
public class HomeFrame extends javax.swing.JFrame{
 
    static int status=0;
    EventListener event;
    //public static boolean stop_flag=false;   
    public static HomeFrame homeframe=null;
    //public static Object testHF=homeframe;  
    public static TestCaseSheetEvaluation tcSheetEval;
    
    public HomeFrame() throws IOException {
        initComponents();
        this.setAlwaysOnTop(true);
        event=new EventListener(); 
        //setup properties file   
        PropertiesFile.loadPropertiesFile();
        //validate the data sheet excel exist
        DataSheetValidation.validateSheet();
        tcSheetEval=new TestCaseSheetEvaluation();           
    }
    
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        buttonGroup1 = new javax.swing.ButtonGroup();
        ExitButton = new javax.swing.JButton();
        StartBtn = new javax.swing.JRadioButton();
        StopBtn = new javax.swing.JRadioButton();
        alertMessage = new javax.swing.JLabel();

        setDefaultCloseOperation(javax.swing.WindowConstants.DO_NOTHING_ON_CLOSE);
        setTitle("Easy Analysis");

        ExitButton.setText("Exit");
        ExitButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                ExitButtonActionPerformed(evt);
            }
        });

        buttonGroup1.add(StartBtn);
        StartBtn.setText("START");
        StartBtn.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                StartBtnActionPerformed(evt);
            }
        });

        buttonGroup1.add(StopBtn);
        StopBtn.setText("STOP");
        StopBtn.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                StopBtnActionPerformed(evt);
            }
        });

        alertMessage.setForeground(new java.awt.Color(255, 0, 0));
        alertMessage.setText(" ");

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addGap(34, 34, 34)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(layout.createSequentialGroup()
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addGroup(layout.createSequentialGroup()
                                .addComponent(StartBtn)
                                .addGap(18, 18, 18)
                                .addComponent(StopBtn))
                            .addComponent(ExitButton))
                        .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                    .addGroup(layout.createSequentialGroup()
                        .addComponent(alertMessage, javax.swing.GroupLayout.PREFERRED_SIZE, 315, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(0, 58, Short.MAX_VALUE))))
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(StartBtn)
                    .addComponent(StopBtn))
                .addGap(18, 18, 18)
                .addComponent(ExitButton)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(alertMessage)
                .addContainerGap())
        );

        pack();
    }// </editor-fold>//GEN-END:initComponents
     
    public static void setVisibility(boolean visibility){   
        if(visibility){
            EventListener.start_flag=1;
            homeframe.setVisible(visibility);
        }
        else{
            EventListener.start_flag=0;
            homeframe.setVisible(visibility);            
        }       
    }
    
    private void StartBtnActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_StartBtnActionPerformed
         if(event.t1.getState() == Thread.State.NEW){
            event.t1.start();                 
         }                 
         if(EventListener.start_flag==0){
             EventListener.start_flag=1;             
         }         
    }//GEN-LAST:event_StartBtnActionPerformed

    private void StopBtnActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_StopBtnActionPerformed
        EventListener.start_flag=0;  
        System.out.println("STOP button pressed");                 
    }//GEN-LAST:event_StopBtnActionPerformed

    private void ExitButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_ExitButtonActionPerformed
        if(event!=null && event.t1 != null){
            event.t1=null;
              event=null;
        }    
        System.exit(NORMAL);
    }//GEN-LAST:event_ExitButtonActionPerformed

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
            java.util.logging.Logger.getLogger(HomeFrame.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(HomeFrame.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(HomeFrame.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(HomeFrame.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                try {
                    homeframe=new HomeFrame();
                } catch (IOException ex) {
                    Logger.getLogger(HomeFrame.class.getName()).log(Level.SEVERE, null, ex);
                }
                homeframe.setVisible(true); 
                System.out.println(Thread.currentThread().getName());
               // testHF=homeframe;   //for testing purpose
                homeframe.addWindowListener(new WindowAdapter(){
                    @Override
                    public void windowClosing(WindowEvent e) {
                        int optionSelected=JOptionPane.showConfirmDialog(null, "Are sure to close this window ?");
                        if(optionSelected == JOptionPane.YES_OPTION){
                           System.exit(NORMAL);
                        }                        
                    }                    
                });
            }
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton ExitButton;
    private javax.swing.JRadioButton StartBtn;
    private javax.swing.JRadioButton StopBtn;
    public static javax.swing.JLabel alertMessage;
    private javax.swing.ButtonGroup buttonGroup1;
    // End of variables declaration//GEN-END:variables

   
}

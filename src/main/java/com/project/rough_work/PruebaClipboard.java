/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.project.rough_work;

/**
 *
 * @author TechAnim
 */import javax.swing.*;
import java.awt.*;
import java.awt.datatransfer.*;
import java.awt.event.ActionEvent;
import java.io.IOException;

public class PruebaClipboard {
    public static void main(String[] args) {
        JFrame frame = new JFrame();
        frame.setTitle("Copy/Paste");
        frame.getContentPane().setLayout(new BorderLayout());
        JPanel btnPanel = new JPanel();
        JButton btnCopy = new JButton("copy");
        JButton btnPaste = new JButton("paste");
        btnPanel.add(btnCopy);
        btnPanel.add(btnPaste);
        final JTextArea textComp = new JTextArea(7,15);

        Action copyAction = new AbstractAction("copy") {
            public void actionPerformed(ActionEvent e) {
                Clipboard clipboard = Toolkit.getDefaultToolkit().getSystemClipboard();
                StringSelection stringSelection = new StringSelection(textComp.getText());
                clipboard.setContents(stringSelection, stringSelection);
            }
        };

        btnCopy.setAction(copyAction);
        Action pasteAction = new AbstractAction("paste") {
            public void actionPerformed(ActionEvent e) {
                Clipboard clipboard = Toolkit.getDefaultToolkit().getSystemClipboard();
                //odd: the Object param of getContents is not currently used
                Transferable contents = clipboard.getContents(null);
                boolean hasTransferableText = (contents != null) && contents.isDataFlavorSupported(DataFlavor.stringFlavor);
                if (hasTransferableText) {
                    try {
                        String result = "";
                        result = (String) contents.getTransferData(DataFlavor.stringFlavor);
                        textComp.append(result);
                    } catch (UnsupportedFlavorException ex) {
                        //highly unlikely since we are using a standard DataFlavor
                        System.out.println(ex);
                        ex.printStackTrace();
                    } catch (IOException ex) {
                        System.out.println(ex);
                        ex.printStackTrace();
                    }
                }
            }
        };
        btnPaste.setAction(pasteAction);

        frame.getContentPane().add(textComp);
        frame.getContentPane().add(btnPanel, BorderLayout.SOUTH);

        frame.setSize(new Dimension(400, 300));
        frame.setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);
        frame.pack();
        frame.setVisible(true);
    }
}
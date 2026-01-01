package org.example;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class ExamResultConverter {
    private JButton inputFolderButton;
    private JTextField inputTxt;
    private JButton viewOutputFolderButton;
    private JTextField outputTxt;
    private JButton convertButton;
    private JPanel panel1;
    private JFileChooser fileChooser;
    SimpleExcelReader reader = new SimpleExcelReader();
    SimpleExcelExporter exporter = new SimpleExcelExporter();

    public ExamResultConverter() {
        inputFolderButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                fileChooser = new JFileChooser(inputTxt.getText());
                fileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
                fileChooser.setDialogTitle("Select a folder to check for Excel files");

                int returnValue = fileChooser.showOpenDialog(null);

                if (returnValue == JFileChooser.APPROVE_OPTION) {
                    String folderPath = fileChooser.getSelectedFile().toString();
                    inputTxt.setText(folderPath);
                    setOutputTxt(folderPath);
                    boolean hasExcelFiles = ExcelFileChecker.hasExcelFiles(folderPath);

                    String message;
                    if (hasExcelFiles) {

                    } else {
                        message = "The selected folder does NOT contain Excel files";
                        JOptionPane.showMessageDialog(null, message, "Excel Files Check",
                                JOptionPane.INFORMATION_MESSAGE);
                    }
                }


            }
        });
        convertButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                try {
                    convertMultipleExcelFiles(getExcelFileNames(inputTxt.getText()), inputTxt.getText());
                } catch (IOException ex) {
                    throw new RuntimeException(ex);
                }
            }
        });

        viewOutputFolderButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                try {
                    Desktop.getDesktop().open(new File(outputTxt.getText()));
                } catch (IOException ex) {
                    throw new RuntimeException(ex);
                }
            }
        });
    }

    public static void main(String[] args) {
        ExamResultConverter examResultConverter = new ExamResultConverter();
        examResultConverter.createAndShowGUI();
    }

    private void createAndShowGUI() {
        JFrame frame = new JFrame("Exam Result Converter");
        frame.setContentPane(this.panel1);
        frame.setSize(800, 600);
        frame.setDefaultCloseOperation(3);
        frame.setLocationRelativeTo((Component) null);
        frame.setVisible(true);
    }

    private static List<String> getExcelFileNames(String folderPath) {
        List<String> excelFileNames = new ArrayList<>();

        // Check if the provided path is a valid directory
        File folder = new File(folderPath);

        if (!folder.exists() || !folder.isDirectory()) {
            System.err.println("The provided path is not a valid directory: " + folderPath);
            return excelFileNames; // Return empty list
        }

        // Define Excel file extensions (both .xls and .xlsx)
        String[] excelExtensions = {".xls", ".xlsx", ".xlsm", ".xlsb"};

        // List all files in the directory
        File[] files = folder.listFiles();

        if (files == null) {
            System.err.println("Unable to read directory: " + folderPath);
            return excelFileNames; // Return empty list
        }

        // Check each file for Excel extensions
        for (File file : files) {
            if (file.isFile()) {
                String fileName = file.getName();
                String fileNameLower = fileName.toLowerCase();

                for (String extension : excelExtensions) {
                    if (fileNameLower.endsWith(extension)) {
                        excelFileNames.add(fileName); // Add the original file name
                        break; // No need to check other extensions for this file
                    }
                }
            }
        }

        return excelFileNames;
    }

    private void convertMultipleExcelFiles(List<String> excelFiles, String folderPath) throws IOException {
        for(String fileName: excelFiles){
            String fullPath = folderPath + "\\" + fileName;
            List<List<String>> originalData = reader.readColumns(fullPath);

            File directory = new File(outputTxt.getText());
            directory.mkdirs();
            exporter.exportWithHeaders(outputTxt.getText() + "output - " + fileName, originalData,
                    "Last Name", "First Name", "Full Name", "Department", "Grade/100.00");
        }

    }

    private void setOutputTxt(String inputFolder){
        outputTxt.setText(inputFolder + "\\output\\");
    }
}

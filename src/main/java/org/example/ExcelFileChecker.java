package org.example;

import java.io.File;

public class ExcelFileChecker {

    /**
     * Checks if a folder contains Excel files
     * @param folderPath The path to the folder to check
     * @return true if the folder contains Excel files, false otherwise
     */
    public static boolean hasExcelFiles(String folderPath) {
        // Check if the provided path is a valid directory
        File folder = new File(folderPath);

        if (!folder.exists() || !folder.isDirectory()) {
            System.err.println("The provided path is not a valid directory: " + folderPath);
            return false;
        }

        // Define Excel file extensions (both .xls and .xlsx)
        String[] excelExtensions = {".xls", ".xlsx", ".xlsm", ".xlsb"};

        // List all files in the directory
        File[] files = folder.listFiles();

        if (files == null) {
            System.err.println("Unable to read directory: " + folderPath);
            return false;
        }

        // Check each file for Excel extensions
        for (File file : files) {
            if (file.isFile()) {
                String fileName = file.getName().toLowerCase();

                for (String extension : excelExtensions) {
                    if (fileName.endsWith(extension)) {
                        return true; // Found at least one Excel file
                    }
                }
            }
        }

        return false; // No Excel files found
    }

}
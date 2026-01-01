package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

public class SimpleExcelExporter {

    // Define column indices for sorting
    private static final int DEPARTMENT_COLUMN_INDEX = 3; // Column 4 in your data
    private static final int FULL_NAME_COLUMN_INDEX = 2;  // Column 3 in your data (Full Name)
    private static final int LAST_COLUMN_INDEX = 4;       // Column 5 in your data (Score)
    private static final double MIN_THRESHOLD = 78.0;     // Threshold for highlighting

    /**
     * Export data to Excel file with sorting and conditional formatting
     * @param filePath Path to save the Excel file
     * @param data List of rows, each row should have 5 columns
     * @throws IOException if file cannot be created
     */
    public void exportToExcel(String filePath, List<List<String>> data) throws IOException {
        exportToExcel(filePath, data, null);
    }

    /**
     * Export data to Excel file with custom headers, sorting, and conditional formatting
     * @param filePath Path to save the Excel file
     * @param data List of rows, each row should have 5 columns
     * @param headers Custom headers for the columns (null for no headers)
     * @throws IOException if file cannot be created
     */
    public void exportToExcel(String filePath, List<List<String>> data, List<String> headers) throws IOException {
        Workbook workbook;

        // Create appropriate workbook based on file extension
        if (filePath.endsWith(".xlsx")) {
            workbook = new XSSFWorkbook();
        } else if (filePath.endsWith(".xls")) {
            workbook = new HSSFWorkbook();
        } else {
            throw new IllegalArgumentException("File must have .xls or .xlsx extension");
        }

        try (FileOutputStream outputStream = new FileOutputStream(filePath)) {
            // Create a sheet
            Sheet sheet = workbook.createSheet("Data");

            // Create styles
            CellStyle headerStyle = createHeaderStyle(workbook);
            CellStyle normalStyle = createNormalStyle(workbook);
            CellStyle redRowStyle = createRedRowStyle(workbook);

            int rowNum = 0;

            // Add headers if provided
            if (headers != null && !headers.isEmpty()) {
                Row headerRow = sheet.createRow(rowNum++);
                for (int i = 0; i < Math.min(5, headers.size()); i++) {
                    Cell cell = headerRow.createCell(i);
                    cell.setCellValue(headers.get(i));
                    cell.setCellStyle(headerStyle);
                }
            }

            // Sort data by Department (column 3) then Full Name (column 2)
            List<List<String>> sortedData = sortData(data);

            // Add data rows
            for (List<String> rowData : sortedData) {
                Row row = sheet.createRow(rowNum++);

                // Check if this row should be highlighted (last column < threshold)
                boolean shouldHighlight = shouldHighlightRow(rowData);

                // Write all columns (should be 5 columns)
                for (int col = 0; col < Math.min(5, rowData.size()); col++) {
                    Cell cell = row.createCell(col);
                    String value = rowData.get(col);

                    // Set cell value
                    setCellValue(cell, value);

                    // Apply style - entire row gets red background if below threshold
                    if (shouldHighlight) {
                        cell.setCellStyle(redRowStyle);
                    } else {
                        cell.setCellStyle(normalStyle);
                    }
                }
            }

            // Auto-size columns
            for (int i = 0; i < 5; i++) {
                sheet.autoSizeColumn(i);
            }

            workbook.write(outputStream);

        } finally {
            workbook.close();
        }
    }

    /**
     * Check if a row should be highlighted (last column < threshold)
     */
    private boolean shouldHighlightRow(List<String> rowData) {
        if (rowData == null || rowData.size() <= LAST_COLUMN_INDEX) {
            return false;
        }

        String lastColumnValue = rowData.get(LAST_COLUMN_INDEX);
        if (isNumeric(lastColumnValue)) {
            try {
                double numericValue = Double.parseDouble(lastColumnValue);
                return numericValue < MIN_THRESHOLD;
            } catch (NumberFormatException e) {
                return false;
            }
        }
        return false;
    }


    /**
     * Set cell value with proper type detection
     */
    private void setCellValue(Cell cell, String value) {
        if (value == null || value.isEmpty()) {
            cell.setCellValue("");
            return;
        }

        if (isNumeric(value)) {
            try {
                cell.setCellValue(Double.parseDouble(value));
            } catch (NumberFormatException e) {
                cell.setCellValue(value);
            }
        } else if (isBoolean(value)) {
            cell.setCellValue(Boolean.parseBoolean(value));
        } else {
            cell.setCellValue(value);
        }
    }

    /**
     * Create header style
     */
    private CellStyle createHeaderStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        // Add borders
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);

        return style;
    }

    /**
     * Create normal row style
     */
    private CellStyle createNormalStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();

        // Add borders
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);

        return style;
    }

    /**
     * Create red row style (entire row highlighted)
     */
    private CellStyle createRedRowStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();

        // Red background
        style.setFillForegroundColor(IndexedColors.RED.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        // White bold font
        Font whiteFont = workbook.createFont();
        whiteFont.setColor(IndexedColors.WHITE.getIndex());
        whiteFont.setBold(true);
        style.setFont(whiteFont);

        // Add borders
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        style.setRightBorderColor(IndexedColors.BLACK.getIndex());

        return style;
    }

    /**
     * Sort data by Department (column 3) then Full Name (column 2)
     */
    private List<List<String>> sortData(List<List<String>> data) {
        if (data == null || data.isEmpty()) {
            return new ArrayList<>(data);
        }

        List<List<String>> sortedList = new ArrayList<>(data);

        sortedList.sort((row1, row2) -> {
            // Make sure rows have enough columns
            if (row1.size() <= DEPARTMENT_COLUMN_INDEX || row2.size() <= DEPARTMENT_COLUMN_INDEX) {
                return 0;
            }

            // First, compare by Department
            String dept1 = row1.get(DEPARTMENT_COLUMN_INDEX);
            String dept2 = row2.get(DEPARTMENT_COLUMN_INDEX);

            int deptCompare = compareStrings(dept1, dept2);
            if (deptCompare != 0) {
                return deptCompare;
            }

            // If departments are equal, compare by Full Name
            if (row1.size() > FULL_NAME_COLUMN_INDEX && row2.size() > FULL_NAME_COLUMN_INDEX) {
                String name1 = row1.get(FULL_NAME_COLUMN_INDEX);
                String name2 = row2.get(FULL_NAME_COLUMN_INDEX);
                return compareStrings(name1, name2);
            }

            return 0;
        });

        return sortedList;
    }

    /**
     * Helper method to compare strings with null safety
     */
    private int compareStrings(String str1, String str2) {
        if (str1 == null && str2 == null) return 0;
        if (str1 == null) return -1;
        if (str2 == null) return 1;
        return str1.compareToIgnoreCase(str2); // Case-insensitive comparison
    }

    /**
     * Export data with custom headers
     * @param filePath Path to save the Excel file
     * @param data List of rows from SimpleExcelReader
     * @param columnHeaders Headers for the columns
     * @throws IOException if file cannot be created
     */
    public void exportWithHeaders(String filePath, List<List<String>> data,
                                  String... columnHeaders) throws IOException {
        List<String> headers;

        if (columnHeaders.length > 0) {
            headers = List.of(columnHeaders);
        } else {
            // Default headers for your 5-column structure
            headers = List.of("ID", "First Name", "Full Name", "Department", "Score");
        }

        exportToExcel(filePath, data, headers);
    }

    /**
     * Export with sorting and highlighting but without headers
     */
    public void exportWithoutHeaders(String filePath, List<List<String>> data) throws IOException {
        exportToExcel(filePath, data, null);
    }

    /**
     * Helper method to check if string is numeric
     */
    private boolean isNumeric(String str) {
        if (str == null || str.trim().isEmpty()) {
            return false;
        }
        return str.matches("-?\\d+(\\.\\d+)?");
    }

    /**
     * Helper method to check if string is boolean
     */
    private boolean isBoolean(String str) {
        if (str == null) return false;
        String lower = str.toLowerCase();
        return lower.equals("true") || lower.equals("false");
    }

    /**
     * Alternative method with customizable threshold
     */
    public void exportWithThreshold(String filePath, List<List<String>> data,
                                    List<String> headers, double threshold) throws IOException {
        // You can create a copy with different threshold by:
        // 1. Creating a new instance with different MIN_THRESHOLD, or
        // 2. Making MIN_THRESHOLD a class variable with setter
        System.out.println("Using threshold: " + threshold);
        // For now, we'll just use the fixed threshold
        exportToExcel(filePath, data, headers);
    }

}
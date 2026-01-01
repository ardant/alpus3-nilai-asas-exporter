package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class SimpleExcelReader {

    /**
     * MODIFIED: Read columns 0, 1, 4, 10 and add Full Name after column 1
     */
    public List<List<String>> readColumns(String filePath) throws IOException {
        List<List<String>> result = new ArrayList<>();

        try (FileInputStream fis = new FileInputStream(filePath)) {
            Workbook workbook;

            if (filePath.endsWith(".xlsx")) {
                workbook = new XSSFWorkbook(fis);
            } else if (filePath.endsWith(".xls")) {
                workbook = new HSSFWorkbook(fis);
            } else {
                throw new IOException("Unsupported file format. Use .xls or .xlsx");
            }

            try {
                Sheet sheet = workbook.getSheetAt(0);
                boolean isFirstRow = true; // Flag to track first row

                for (Row row : sheet) {
                    // Skip the first row
                    if (isFirstRow) {
                        isFirstRow = false;
                        continue; // Skip to next row
                    }

                    List<String> rowData = new ArrayList<>();

                    // Get values for columns 0, 1, 4, 10
                    String col0 = cellToString(row.getCell(0));
                    String col1 = cellToString(row.getCell(1));
                    String col4 = cellToString(row.getCell(4));
                    String col10 = cellToString(row.getCell(10));

                    // Add to row in order: col0, col1, Full Name, col4, col10
                    rowData.add(col0);                     // Column 0
                    rowData.add(col1);                     // Column 1
                    rowData.add((col1 + " " + col0).trim()); // Full Name (new column)
                    rowData.add(col4);                     // Column 4
                    rowData.add(col10);                    // Column 10

                    result.add(rowData);
                }
            } finally {
                workbook.close();
            }
        }

        return result;
    }

    /**
     * Helper method unchanged
     */
    private String cellToString(Cell cell) {
        if (cell == null) return "";

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue().trim();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                }
                double num = cell.getNumericCellValue();
                if (num == Math.floor(num)) {
                    return String.valueOf((int) num);
                }
                return String.valueOf(num);
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.toString();
            default:
                return "";
        }
    }

}
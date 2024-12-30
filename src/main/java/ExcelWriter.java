/*
Program Created by: Vincent Haney Jr.
Date: 02/25/2024
 */

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ExcelWriter {
    public static void writeToExcel(String fullName, String address, String date, String time, String price, int numBedrooms, int numBathrooms, String filePath) throws IOException {
        Workbook workbook;
        Sheet sheet;
        FileInputStream fileInputStream = null;

        try {
            // Open existing Excel file or create a new one if it doesn't exist
            fileInputStream = new FileInputStream(filePath);
            workbook = WorkbookFactory.create(fileInputStream);
        } catch (IOException e) {
            // If the file doesn't exist, create a new workbook
            workbook = new XSSFWorkbook();
            // Create a new sheet
            sheet = workbook.createSheet("Property Details");
            // Create headers
            Row headerRow = sheet.createRow(0);
            String[] headers = {"Full Name", "Address", "Date", "Time", "Price", "Bedrooms", "Bathrooms"};
            // Create a CellStyle with bold and underline font for headers
            CellStyle headerCellStyle = workbook.createCellStyle();
            Font headerFont = workbook.createFont();
            headerFont.setBold(true);
            headerFont.setUnderline(Font.U_SINGLE);
            headerCellStyle.setFont(headerFont);
            // Apply the CellStyle to the header cells
            for (int i = 0; i < headers.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers[i]);
                cell.setCellStyle(headerCellStyle);
            }
        } finally {
            if (fileInputStream != null) {
                fileInputStream.close();
            }
        }

        // Get the first sheet
        sheet = workbook.getSheetAt(0);

        // Find the next available row
        int lastRowNum = sheet.getLastRowNum();
        Row dataRow = sheet.createRow(lastRowNum + 1);

        // Write data to the new row
        dataRow.createCell(0).setCellValue(fullName);
        dataRow.createCell(1).setCellValue(address);
        dataRow.createCell(2).setCellValue(date);
        dataRow.createCell(3).setCellValue(time);
        dataRow.createCell(4).setCellValue(price);
        dataRow.createCell(5).setCellValue(numBedrooms);
        dataRow.createCell(6).setCellValue(numBathrooms);
        
        
        // Set column widths for each field
        sheet.setColumnWidth(0, 20 * 256); // Column for Full Name
        sheet.setColumnWidth(1, 30 * 256); // Column for Address
        sheet.setColumnWidth(2, 15 * 256); // Column for Date
        sheet.setColumnWidth(3, 15 * 256); // Column for Time
        sheet.setColumnWidth(4, 15 * 256); // Column for Price
        sheet.setColumnWidth(5, 15 * 256); // Column for Bedrooms
        sheet.setColumnWidth(6, 15 * 256); // Column for Bathrooms        

        // Write the updated workbook back to the file
        try (FileOutputStream fileOut = new FileOutputStream(filePath)) {
            workbook.write(fileOut);
        }

        // Close the workbook to prevent resource leaks
        workbook.close();
    }
}
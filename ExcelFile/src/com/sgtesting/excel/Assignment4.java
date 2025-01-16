package com.sgtesting.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class Assignment4 {
    public static void main(String[] args) {
        // Array of 20 country names
        String[] countries = {
                "India", "USA", "Canada", "Australia", "Germany", "France", "Italy", "Spain",
                "Brazil", "China", "Japan", "South Korea", "Russia", "South Africa", "Mexico",
                "Indonesia", "UK", "Argentina", "Saudi Arabia", "Turkey"
        };

        // Create a workbook and sheet
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("FirstSheet");

        // Write country names diagonally
        for (int i = 0; i < countries.length; i++) {
            Row row = sheet.createRow(i);      // Create a row (i-th row)
            Cell cell = row.createCell(i);    // Create a cell in the i-th column
            cell.setCellValue(countries[i]);  // Set cell value to the country name
        }

        // Write the workbook to a file
        try (FileOutputStream fileOut = new FileOutputStream("D:\\Demo\\Test\\countries.xlsx")) {
            workbook.write(fileOut);
            System.out.println("Excel file with country names written diagonally successfully!");
        } catch (IOException e) {
            System.out.println("An error occurred while writing the file: " + e.getMessage());
        } finally {
            try {
                workbook.close();
            } catch (IOException e) {
                System.out.println("Error closing the workbook: " + e.getMessage());
            }
        }
    }
}

package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

public class Main {
    public static void main(String[] args) {
        try {
            FileInputStream file = new FileInputStream(new File("C:\\Users\\Lenovo\\projects\\ExcelLab\\ExcelLab\\src\\TestExcel.xlsx"));

            // Create Workbook instance holding reference to .xlsx file
            XSSFWorkbook workbook = new XSSFWorkbook(file);

            // Get first sheet from the workbook
            XSSFSheet sheet = workbook.getSheetAt(0);

            // Iterate through each row
            Iterator<Row> rowIterator = sheet.iterator();
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();

                // Iterate through each cell
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();

                    // Check the cell type and print accordingly
                    switch (cell.getCellType()) {
                        case STRING:
                            System.out.print(cell.getStringCellValue() + "\t");
                            break;
                        case NUMERIC:
                            System.out.print(cell.getNumericCellValue() + "\t");
                            break;
                        case BOOLEAN:
                            System.out.print(cell.getBooleanCellValue() + "\t");
                            break;
                        case FORMULA:
                            System.out.print(cell.getCellFormula() + "\t");
                            break;
                        default:
                            System.out.print("UNKNOWN\t");
                            break;
                    }
                }
                System.out.println();
            }

            workbook.close();
            file.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}

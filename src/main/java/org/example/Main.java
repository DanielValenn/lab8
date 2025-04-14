package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

public class Main {
    public static void main(String[] args) {
        String inputPath = "C:\\Users\\Lenovo\\projects\\ExcelLab\\ExcelLab\\src\\TestExcel.xlsx";
        String outputPath = "C:\\Users\\Lenovo\\projects\\ExcelLab\\ExcelLab\\src\\Output.xlsx";

        try {
            // Read
            FileInputStream file = new FileInputStream(new File(inputPath));
            XSSFWorkbook workbookIn = new XSSFWorkbook(file);
            XSSFSheet sheetIn = workbookIn.getSheetAt(0);

            // Output Workbook
            XSSFWorkbook workbookOut = new XSSFWorkbook();
            XSSFSheet sheetOut = workbookOut.createSheet("Incremented Ages");

            int rowIndex = 0;

            for (Row row : sheetIn) {
                Row newRow = sheetOut.createRow(rowIndex++);
                int cellIndex = 0;

                for (Cell cell : row) {
                    Cell newCell = newRow.createCell(cellIndex);

                    // If last row & numeric type
                    if (cellIndex == row.getLastCellNum() - 1 && cell.getCellType() == CellType.NUMERIC) {
                        double newAge = cell.getNumericCellValue() + 1;
                        newCell.setCellValue(newAge);
                    } else {
                        switch (cell.getCellType()) {
                            case STRING:
                                newCell.setCellValue(cell.getStringCellValue());
                                break;
                            case NUMERIC:
                                newCell.setCellValue(cell.getNumericCellValue());
                                break;
                            case BOOLEAN:
                                newCell.setCellValue(cell.getBooleanCellValue());
                                break;
                            case FORMULA:
                                newCell.setCellFormula(cell.getCellFormula());
                                break;
                            default:
                                newCell.setCellValue("UNKNOWN");
                                break;
                        }
                    }

                    cellIndex++;
                }
            }

            // Write new
            FileOutputStream outFile = new FileOutputStream(new File(outputPath));
            workbookOut.write(outFile);

            // Close
            outFile.close();
            workbookIn.close();
            workbookOut.close();
            file.close();

            System.out.println("Datele au fost scrise în " + outputPath + " cu vârstele incrementate.");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}

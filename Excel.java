package org.example;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

public class Excel {
    public static void main(String[] args) throws IOException {
        Workbook workbook = new HSSFWorkbook();
        Sheet sheet = workbook.createSheet("Sheet1");
        Row row0 = sheet.createRow(0);
        Cell cell0 = row0.createCell(0);
        cell0.setCellValue("Name");
        Cell cell1 = row0.createCell(1);
        cell1.setCellValue("Age");
        Cell cell2 = row0.createCell(2);
        cell2.setCellValue("email");
        Row row1 = sheet.createRow(1);
        Cell rc10 = row1.createCell(0);
        rc10.setCellValue("John Doe");
        Cell rc11 = row1.createCell(1);
        rc11.setCellValue("30");
        Cell rc12 = row1.createCell(2);
        rc12.setCellValue("john@test.com");
        Row row2 = sheet.createRow(2);
        Cell rc20 = row2.createCell(0);
        rc20.setCellValue("John Doe");
        Cell rc21 = row2.createCell(1);
        rc21.setCellValue("28");
        Cell rc22 = row2.createCell(2);
        rc22.setCellValue("john@test.com");
        Row row3 = sheet.createRow(3);
        Cell rc30 = row3.createCell(0);
        rc30.setCellValue("Bob Smith");
        Cell rc31 = row3.createCell(1);
        rc31.setCellValue("35");
        Cell rc32 = row3.createCell(2);
        rc32.setCellValue("jacky@example.com");
        Row row4 = sheet.createRow(4);
        Cell rc40 = row4.createCell(0);
        rc40.setCellValue("Swapnil");
        Cell rc41 = row4.createCell(1);
        rc41.setCellValue("37");
        Cell rc42 = row4.createCell(2);
        rc42.setCellValue("Swapnil@example.com");
        FileOutputStream fileOutputStream = new FileOutputStream("Employee.xls");
        workbook.write(fileOutputStream);

        Workbook wb = WorkbookFactory.create(new File("Employee.xls"));
        Sheet sh = wb.getSheetAt(0);

        for (Row row : sh) {
            for (Cell cell : row) {
                System.out.print(cell.getStringCellValue() + " ");
            }
            System.out.println();
        }
    }
}

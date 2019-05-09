package com.company;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class Main {

    public static void main(String[] args) throws IOException {
        //TODO: WORKS!!! need to implement some logic (WRITE)
        String[] books = {
                "The Tempest",
                "Gitanjali",
                "Harry Potter"
        };
        String[] authors = {
                "William Shakespeare",
                "Rabindranath Tagore",
                "J. K. Rowling"
        };

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet();
        sheet.setColumnWidth((short) 0, (short)((50 * 8) / ((double) 1 / 20)));
        sheet.setColumnWidth((short) 1, (short)((50 * 8) / ((double) 1 / 20)));
        workbook.setSheetName(0, "XSSFWorkbook example");

        Row headerRow = sheet.createRow(0);
        Cell cell1 = headerRow.createCell(0);
        cell1.setCellValue("Book");
        Cell cell2 = headerRow.createCell(1);
        cell2.setCellValue("Author");

        Row row = null;
        Cell cell = null;
        for (int rownum = 1; rownum <= books.length; rownum++) {
            row = sheet.createRow(rownum);
            cell = row.createCell(0);
            cell.setCellValue(books[rownum - 1]);
            cell = row.createCell(1);
            cell.setCellValue(authors[rownum - 1]);
        }

        final String FILE_NAME = "./books-test.xlsx";
        FileOutputStream outputStream = new FileOutputStream(FILE_NAME);
        workbook.write(outputStream);
        outputStream.close();
        workbook.close();

    }
}

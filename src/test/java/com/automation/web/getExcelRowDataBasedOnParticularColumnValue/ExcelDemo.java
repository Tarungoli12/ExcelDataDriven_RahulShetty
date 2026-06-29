package com.automation.web.getExcelRowDataBasedOnParticularColumnValue;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

public class ExcelDemo {
    public static void main(String[] args) throws IOException {
        FileInputStream fis = new FileInputStream("C://Users//280713//Downloads//ExcelDataDriven.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        int sheetsCount = workbook.getNumberOfSheets();
        for (int i = 0; i < sheetsCount; i++) {
            if (workbook.getSheetName(i).equalsIgnoreCase("studentData")) {
                XSSFSheet sheet = workbook.getSheetAt(i);
                int rowsCount = sheet.getPhysicalNumberOfRows();
                int columnIndex = 0;
                for (int j = 0; j < rowsCount; j++) {
                    Iterator<Cell> cellIterator = sheet.getRow(j).cellIterator();
                    while (cellIterator.hasNext()) {
                        Cell cell = cellIterator.next();
                        if (cell.getCellType() == CellType.STRING && cell.getStringCellValue().trim().equalsIgnoreCase("Testcase")) {
                            columnIndex = cell.getColumnIndex();
                            break;
                        }
                    }
                }

                int targetRowIndex = -1;

                for (int k = 1; k <= sheet.getLastRowNum(); k++) {
                    Row row = sheet.getRow(k);

                    if (row != null) {
                        Cell cell = row.getCell(columnIndex);
                        if (cell != null && String.valueOf((int) cell.getNumericCellValue()).equals("113")) {
                            targetRowIndex = k;
                            break;
                        }
                    }
                }
                DataFormatter formatter = new DataFormatter();
                if (targetRowIndex != -1) {
                    Row targetRow = sheet.getRow(targetRowIndex);

                    for (Cell cell : targetRow) {
                        System.out.print(formatter.formatCellValue(cell) + " ");
                    }
                }
            }
        }
    }
}

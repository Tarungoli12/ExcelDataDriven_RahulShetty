package org;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

public class DataDriven {

    public ArrayList<String> getExcelData(String testCaseName) throws IOException {

        ArrayList<String> list = new ArrayList<>();

        // connect to Excel file
        XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream("C://Users//280713//Downloads//ExcelDataDriven.xlsx"));

        //get no of sheets in Excel file
        int workBookSheets = workbook.getNumberOfSheets();

        for (int i = 0; i < workBookSheets; i++) {

            //check if sheets in Excel file are equal to sheet name
            if (workbook.getSheetName(i).equalsIgnoreCase("DataDriven")) {

                //store the sheet by using the index
                XSSFSheet sheet = workbook.getSheetAt(i);

                //get the no of rows in sheet
                Iterator<Row> rows = sheet.iterator();

                // store the first row
                Row firstRow = rows.next();

                //get the no of cells in first row
                Iterator<Cell> cell = firstRow.cellIterator();

                //check if cell name is equal to testcase and store the matched cell index
                int column = 0;
                int k = 0;
                while (cell.hasNext()) {
                    if (cell.next().getStringCellValue().equalsIgnoreCase("Testcase")) {
                        column = k;
                    }
                    k++;
                }
                System.out.println(column);

                //now by using the matched cell index(Testcase),we will switch to rows,
                // so that we go to check the particular 'purchase' row and print the purchase row cell values by using matched cell index.
                while (rows.hasNext()) {
                    Row r = rows.next();
                    if (r.getCell(column).getStringCellValue().equalsIgnoreCase(testCaseName)) {
                        Iterator<Cell> cells = r.cellIterator();
                        while (cells.hasNext()) {
                            Cell c = cells.next();
                            if(c.getCellType()== CellType.STRING){
                                list.add(c.getStringCellValue());
                            }else{
                                list.add(NumberToTextConverter.toText(c.getNumericCellValue()));
                            }
                        }
                    }
                }
            }
        }
        return list;
    }
}
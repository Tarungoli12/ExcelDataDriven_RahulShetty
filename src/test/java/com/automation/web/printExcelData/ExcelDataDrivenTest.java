package com.automation.web.printExcelData;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;
import java.io.FileInputStream;
import java.io.IOException;


public class ExcelDataDrivenTest {

    @Test(dataProvider = "ExcelData")
    public void DataDrivenTestExcel(String data1,String data2,String data3,String data4){
        System.out.println(data1 + data2 + data3 + data4);
    }


    @DataProvider(name = "ExcelData")
    public Object[][] getData() throws IOException {

        DataFormatter formatter = new DataFormatter();

        //connect to Excel file
        XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream("C://Users//280713//Downloads//ExcelDataDriven.xlsx"));

        //switch to first Excel sheet in Excel file
        XSSFSheet sheet = workbook.getSheetAt(0);

        //get not of rows in sheet
        int rowCount = sheet.getPhysicalNumberOfRows();

        //get the first row
        Row firstRow = sheet.getRow(0);

        //get the no of columns by using the first row with last cell num (index value)
        int columnCount = firstRow.getLastCellNum();

        //create a 2-dimensional array by skipping the first row(which is a header-Testcase)
        Object[][] data = new Object[rowCount-1][columnCount];

        for(int i=0;i<rowCount-1;i++){

            //get the rows based on i but skip the first header row
            XSSFRow row = sheet.getRow(i+1);

            for(int j=0;j<columnCount;j++){

                //get the cells for the row and store into 2-dimensional array
                XSSFCell cell = row.getCell(j);
                data[i][j] = formatter.formatCellValue(cell);

            }
        }
        return data;
    }
}

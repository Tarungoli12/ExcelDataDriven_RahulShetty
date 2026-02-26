package org;

import java.io.IOException;
import java.util.ArrayList;

public class DataDrivenTest {
    public static void main(String[] args) throws IOException {
        DataDriven driven = new DataDriven();
        ArrayList data = driven.getExcelData("Add Profile");
        System.out.println(data.get(0));
        System.out.println(data.get(1));
        System.out.println(data.get(2));
        System.out.println(data.get(3));
    }
}

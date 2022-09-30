package com.syntax;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

public class Main {
    public static void main(String[] args) throws IOException {
        File myFile=new File("C:\\Users\\party\\OneDrive\\Documents\\excel\\DSAI140202.xlsx");
        FileInputStream fis = new FileInputStream(myFile);
        XSSFWorkbook wb = new XSSFWorkbook(fis);
        XSSFSheet firstSheet = wb.getSheetAt(0);
        ArrayList<String> skuNameList = new ArrayList<String>();
        ArrayList<String> picNameList = new ArrayList<String>();
        Iterator<Row> rowIterator = firstSheet.iterator();
        //CarMeth eat = new CarMeth("/Volumes/Krishna/Jar Files/poi-3.16-beta1/TestData.xlsx");


        //makes new excel file
        // (OutputStream fileOut = new FileOutputStream("skuResult.xlsx")) {
            //wb.write(fileOut);
        //}

        Sheet sheet = wb.getSheet("Sheet1");
        int rowIndex = 1; //as you go higher goes down column
        Row row = sheet.getRow(rowIndex - 1);
        Row row1 = sheet.getRow(rowIndex);
        int cellIndex = 0;//0 is skuName       1 is image
        String skuName = String.valueOf(row.getCell(rowIndex - 1));
        String skuName1 = String.valueOf(row.getCell(rowIndex));
        int rowCount = sheet.getLastRowNum()+1;

        while( rowIndex < rowCount) {
            row = sheet.getRow(rowIndex - 1);
            row1 = sheet.getRow(rowIndex);
            skuName = String.valueOf(row.getCell(cellIndex));
            skuName1 = String.valueOf(row1.getCell(cellIndex));

            //System.out.println(skuName + " " + skuName1 + "\n" + skuName.equals(skuName1));

            if(skuName.equals(skuName1)) {
                rowIndex++;
                continue;
            }

            else {
                skuNameList.add(String.valueOf(row.getCell(cellIndex)));
                rowIndex++;
            }
        }


        for(int i = 0; i < skuNameList.size(); i++)
            System.out.println(skuNameList.get(i) + "\n");






    }
}
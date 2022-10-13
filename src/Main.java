import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

public class Main {
    public static void main(String[] args) throws IOException {
        File myFile=new File("/Users/darienlin/Documents/BC work/FGGP7PTR01.xlsx");
        FileInputStream fis = new FileInputStream(myFile);
        XSSFWorkbook wb = new XSSFWorkbook(fis);
        ArrayList<String> skuNameList = new ArrayList<>();
        ArrayList<String> picNameList = new ArrayList<>();

       //creates new workbook
        XSSFWorkbook skuResult = new XSSFWorkbook();
        CreationHelper createHelper = skuResult.getCreationHelper();

        //creates new sheet
        Sheet page1 = skuResult.createSheet("page1");

        //makes new excel file
        // (OutputStream fileOut = new FileOutputStream("skuResult.xlsx")) {
            //wb.write(fileOut);
        //}

        Sheet sheet = wb.getSheet("Sheet1");
        int rowIndex = 1;//as you go higher goes down column
        int row2Index;
        Row row = sheet.getRow(rowIndex - 1);
        Row row1 = sheet.getRow(rowIndex);
        Row imageRow;
        int cellIndex = 0;//0 is skuName       1 is image
        String skuName = String.valueOf(row.getCell(rowIndex - 1));
        String skuName1 = String.valueOf(row.getCell(rowIndex));
        String xImages = String.valueOf(row.getCell(rowIndex));
        String allXImages ="";
        String picNow;
        int rowCount = sheet.getLastRowNum()+1;
        int xImageCounter = 0;
        boolean first = false, ploop = true;

        while( rowIndex < rowCount) {
            row = sheet.getRow(rowIndex - 1);
            row1 = sheet.getRow(rowIndex);
            skuName = String.valueOf(row.getCell(0));
            skuName1 = String.valueOf(row1.getCell(0));

            //System.out.println(skuName + " " + skuName1 + "\n" + skuName.equals(skuName1));

            if(skuName.equals(skuName1)) {
                rowIndex++;
                xImageCounter++;

                if(first){
                    row2Index = rowIndex - 1;
                    first = false;
                    for(int i = 1; i < xImageCounter; i++){
                        imageRow = sheet.getRow(i);
                        xImages = String.valueOf(imageRow.getCell(1));
                        xImages = xImages.substring(0,xImages.indexOf(".jpg") + 4);
                        allXImages = allXImages + "," + xImages;
                    }
                }

                else
                continue;
            }

            else {
                skuNameList.add(String.valueOf(row1.getCell(0)));
                picNow = String.valueOf(row1.getCell(1));
                picNow = picNow.substring(0,picNow.indexOf(".jpg") + 4);
                picNameList.add(picNow + allXImages);
                rowIndex++;

                if(ploop){
                    first = true;
                    ploop = false;
                }
            }
        }

        row1 = sheet.getRow(0);
        skuNameList.add(String.valueOf(row1.getCell(0)));
        picNow = String.valueOf(row1.getCell(1));
        picNow = picNow.substring(0,picNow.indexOf(".jpg") + 4);
        picNameList.add(picNow + allXImages);

        for(int i = 0; i < skuNameList.size(); i++) {
            Row skuRow = page1.createRow(i);
            Cell cell = skuRow.createCell(0);
            cell.setCellValue(skuNameList.get(i));
        }

        for(int i = 0; i < picNameList.size(); i++) {
            Row imgRow = page1.createRow(i);
            Cell cell = imgRow.createCell(1);
            cell.setCellValue(picNameList.get(i));
        }

        //for(int i = 0; i < skuNameList.size();i++)
            //sheet.autoSizeColumn(i);

        FileOutputStream fileOut = new FileOutputStream("poi-generated-file.xlsx");
        skuResult.write(fileOut);
        fileOut.close();
        skuResult.close();








    }
}
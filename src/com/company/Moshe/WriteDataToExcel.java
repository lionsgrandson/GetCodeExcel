package com.company.Moshe;

// Java program to write data in excel sheet using java code

import java.io.File;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

public class WriteDataToExcel {

    // any exceptions need to be caught
//    public static void main(String[] args) throws Exception
//    {
    int NumRowId = 0;

    public WriteDataToExcel() {



    }
    // workbook object
    XSSFWorkbook workbook = new XSSFWorkbook();

    // spreadsheet object
    XSSFSheet spreadsheet
            = workbook.createSheet(" Emails ");

    // creating a row object
    XSSFRow row;

    // This data needs to be written (Object[])
    Map<Integer, Object[]> studentData
            = new TreeMap<Integer, Object[]>();
    public void Excel(String website, String email) throws Exception {


//        studentData.put(
//                1,
//                new Object[]{"Website", "Email"});
        studentData.put(
                NumRowId,
                new Object[]{website, email});


//        System.out.println("row is: " + NumRowId);
        Set<Integer> keyid = studentData.keySet();
        NumRowId++;
        int rowid = NumRowId;

//https://www.maimonides.org/
// https://mosheschwartzberg.com
// writing the data into the sheets...
        //https://byboston.org

        for (Integer key : keyid) {

            row = spreadsheet.createRow(rowid++);
            Object[] objectArr = studentData.get(key);

            int cellid = 0;
            for (Object obj : objectArr) {

                Cell cell = row.createCell(cellid++);
                cell.setCellValue((String) obj);
//                System.out.println((String) obj + " cell "+cellid + " row " + rowid + " key " + keyid);
            }
        }

        // .xlsx is the format for Excel Sheets...
        // writing the workbook into the file...
        String path ="C:/Users/moshe/OneDrive/Documents/WZO/project/Eurpoe-Names/EmailsSheet.xlsx";
        File output = new File(path);
        if(!output.exists()) {
            output.createNewFile();
        }

        FileOutputStream out = new FileOutputStream(output);

        workbook.write(out);


        if(out!=null) {
            out.close();
        }


    }

}

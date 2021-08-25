package com.company.Moshe;

import java.io.File;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

public class WriteDataToExcel {
    int NumRowId = 0;
    //TODO - Fix file path problem, allow it to work on other devices
    //TODO - Add an array for multiple emails at once
    //TODO - change it from writing to excel to write to a text file with the option to add that into exel
    //TODO - or fix the overwriting file problem
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
            = new TreeMap<>();
    public void Excel(String website, String email) throws Exception {

        studentData.put(
                NumRowId,
                new Object[]{website, email});


        Set<Integer> keyid = studentData.keySet();
        NumRowId++;
        int rowid = NumRowId;

//https://www.maimonides.org/
// https://mosheschwartzberg.com
//https://byboston.org
        // writing the data into the sheets...

        for (Integer key : keyid) {

            row = spreadsheet.createRow(rowid++);
            Object[] objectArr = studentData.get(key);

            int cellid = 0;
            for (Object obj : objectArr) {

                Cell cell = row.createCell(cellid++);
                cell.setCellValue((String) obj);
            }
        }

        // .xlsx is the format for Excel Sheets...
        // writing the workbook into the file...
        String path ="C:/Users/Dell/Documents/WZO/project/Eurpoe Names/EmailsSheet.xlsx";
        File output = new File(path);
        output.createNewFile();

        FileOutputStream out = new FileOutputStream(output);

        workbook.write(out);

        out.close();

    }

}

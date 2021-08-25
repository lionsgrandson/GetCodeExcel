package Moshe;

import java.io.BufferedInputStream;
import java.net.URL;
import java.net.URLConnection;
import java.util.Scanner;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Main {

    public static void main(String[] args) {
        int exit = 1;
//        int NumRowId = 1;
        WriteDataToExcel writeDataToExcel = new WriteDataToExcel();


        Scanner myObj = new Scanner(System.in);
        while (exit < 11) {
            System.out.println("Enter Website");
            String urlString = myObj.nextLine();
            try {


                URL url;
                URLConnection uc;
                StringBuilder parsedContentFromUrl = new StringBuilder();
//        String urlString="https://www.shaarei.org/contact.html";
                System.out.println("Getting content for URl : " + urlString);
                url = new URL(urlString);
                uc = url.openConnection();
                uc.connect();
                uc = url.openConnection();
                uc.addRequestProperty("User-Agent",
                        "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)");
                uc.getInputStream();
                BufferedInputStream in = new BufferedInputStream(uc.getInputStream());
                int ch;
                while ((ch = in.read()) != -1) {
                    parsedContentFromUrl.append((char) ch);
                }

                Matcher m = Pattern.compile("[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\\.[a-zA-Z0-9-.]+").matcher(parsedContentFromUrl);
                String emails="";
                while (m.find()) {
                    System.out.println(m.group());
                     emails = m.group();
                }
                writeDataToExcel.Excel(urlString,emails);
            } catch (Exception e) {
                System.out.println(e.getMessage() + " Failed to find more");
                try {
                    writeDataToExcel.Excel(urlString,e.getMessage());
                } catch (Exception exception) {
                    System.out.println(e.getMessage() + "failed to write");
                }
            }

            if (exit == 10) {
                System.out.println("\n" +
                        "*************************\n" +
                        "Do you want to continue?" +
                        "\n*************************");
                String exStr = myObj.nextLine().toLowerCase();
                if (exStr.equals("y".toLowerCase())) {
                    exit = 0;
                } else {
                    break;
                }
            } else {
                exit++;
            }
        }
    }
}
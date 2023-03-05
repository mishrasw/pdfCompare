package com.pdfcomp;

import de.redsix.pdfcompare.PdfComparator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

public class pdfvalidation {

    public static void main(String[] args) throws IOException {

        System.out.println(pdfvalidation.class.getProtectionDomain().getCodeSource().getLocation().toString().substring(1).replace("/","\\").replace("ile:\\",""));
        String recFolder = pdfvalidation.class.getProtectionDomain().getCodeSource().getLocation().toString().substring(1).replace("/","\\").replace("ile:\\","");
        recFolder=pathincaseJar(recFolder);
        String actualFolder = recFolder+"\\actual";
        String baseFolder =recFolder+"\\baseline";
        String configFolder =recFolder+"\\config";
        String diffFolder = recFolder+"\\diff";
        String TestResPath = recFolder+"\\Test_Result.xlsx";
        FileInputStream file = new FileInputStream(new File(TestResPath));
        //FileInputStream file = new FileInputStream(new File("C:\\javaProject\\pdfValidation\\target\\pdfValidation\\resources\\Test_Result.xlsx"));

        XSSFWorkbook workbook = new XSSFWorkbook(file);
        XSSFSheet sheet = workbook.getSheetAt(0);

        Iterator<Row> rowIterator = sheet.iterator();
        String fileName;
        String config;

        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            if (!row.getCell(0).getStringCellValue().trim().equals("")) {
                System.out.println(row.getCell(0).getStringCellValue().trim());
                fileName = row.getCell(0).getStringCellValue().trim();
                config = row.getCell(1).getStringCellValue().trim();
                boolean boolresult;
                if (row.getRowNum() == 0)
                    continue;
                /*
                if (!(config.trim().equals("")))
                    boolresult = new PdfComparator(".resources\\baseline\\" + fileName + ".pdf",
                            ".resources\\actual\\" + fileName + ".pdf")
                            .withIgnore(".resources\\config\\" + config + ".conf")
                            .compare().writeTo(".resources\\diff\\diff-" + fileName + "");

                else
                    boolresult = new PdfComparator(".resources\\baseline\\" + fileName + ".pdf",
                            ".resources\\actual\\" + fileName + ".pdf")
                            .compare().writeTo(".resources\\diff\\diff-" + fileName + "");
*/
                if (!(config.trim().equals("")))
                    boolresult = new PdfComparator(baseFolder+"\\" + fileName + ".pdf",
                            actualFolder+"\\" + fileName + ".pdf")
                            .withIgnore(configFolder+"\\" + config + ".conf")
                            .compare().writeTo(diffFolder+"\\diff-" + fileName);

                else
                    boolresult = new PdfComparator(baseFolder+"\\" + fileName + ".pdf",
                            actualFolder+"\\" + fileName + ".pdf")
                            .compare().writeTo(diffFolder+"\\diff-" + fileName);
                if (boolresult)
                    row.getCell(3).setCellValue("PASS");
                else
                    row.getCell(3).setCellValue("FAIL");
            }
            else{
                break;
            }
        }

            workbook.write(new FileOutputStream(new File(TestResPath)));
        //workbook.write(new FileOutputStream(pdfvalidation.class.getClassLoader().getResource("Test_Result.xlsx").getFile()));
    }



    public static String pathincaseJar(String path)
    {
        String localpath = path;
        if(localpath.contains(".jar")){
            /*int index = localpath.lastIndexOf('\\');
            String lastpath = localpath.substring(index +1);
           String temppath = localpath.substring(0, localpath.lastIndexOf("\\"));
           String firstpath = temppath.substring(0, temppath.lastIndexOf("\\"));
            localpath = firstpath+"\\"+lastpath;*/
            localpath=localpath.replace("pdfValidation\\pdfvalidation.jar","classes\\");
        }
System.out.println(localpath);
        return localpath;
    }
}
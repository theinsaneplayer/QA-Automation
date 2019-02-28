import  java.io.*;
import  org.apache.poi.hssf.usermodel.HSSFSheet;
import  org.apache.poi.hssf.usermodel.HSSFWorkbook;
import  org.apache.poi.hssf.usermodel.HSSFRow;

public class CreateExcelFile{
    public static void main(String[]args) {
        try {
            String filename = "C:\\Users\\p.shatskov\\Documents\\QA Automation/NewExcelFile.xls" ;
            HSSFWorkbook workbook = new HSSFWorkbook();
            HSSFSheet sheet = workbook.createSheet("FirstSheet");

            HSSFRow rowhead = sheet.createRow((short)0);
            rowhead.createCell(0).setCellValue("Имя");
            rowhead.createCell(1).setCellValue("Фамилия");
            rowhead.createCell(2).setCellValue("Отчество");
            rowhead.createCell(3).setCellValue("Возраст");
            rowhead.createCell(4).setCellValue("Пол");
            rowhead.createCell(5).setCellValue("Дата рождения");
            rowhead.createCell(6).setCellValue("ИНН");
            rowhead.createCell(7).setCellValue("Почтовый индекс");
            rowhead.createCell(7).setCellValue("Страна");
            rowhead.createCell(8).setCellValue("Область");
            rowhead.createCell(9).setCellValue("Город");
            rowhead.createCell(10).setCellValue("Улица");
            rowhead.createCell(11).setCellValue("Дом");
            rowhead.createCell(12).setCellValue("Квартира");


            HSSFRow row = sheet.createRow((short)1);
            row.createCell(0).setCellValue("");
            row.createCell(1).setCellValue("");
            row.createCell(2).setCellValue("");
            row.createCell(3).setCellValue("");

            FileOutputStream fileOut = new FileOutputStream(filename);
            workbook.write(fileOut);
            fileOut.close();
            workbook.close();
            System.out.println("Файл Excel успешно создан");

        } catch ( Exception ex ) {
            System.out.println(ex);
        }
    }
}
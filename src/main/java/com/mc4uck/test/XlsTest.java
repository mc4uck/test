package com.mc4uck.test;


import org.apache.poi.hssf.usermodel.*;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

/**
 * Created by mc4uck
 * on 08.10.2019.
 */
public class XlsTest {
    public static void main(String[] args) throws IOException {
        writeIntoExcel("e:\\\\J\\test.xls");
        readFromExcel("e:\\\\J\\test.xls");

    }

    @SuppressWarnings("deprecation")
    public static void writeIntoExcel(String file) throws FileNotFoundException, IOException {
        HSSFWorkbook book =  new HSSFWorkbook();
        HSSFSheet sheet = book.createSheet("Birthdays");

        // Нумерация начинается с нуля
        HSSFRow row = sheet.createRow(0);

        // Мы запишем имя и дату в два столбца
        // имя будет String, а дата рождения --- Date,
        // формата dd.mm.yyyy
        HSSFCell name = row.createCell((short) 0);
        name.setCellValue("John");

        HSSFCell birthdate = row.createCell((short) 1);

        HSSFDataFormat format = book.createDataFormat();
        HSSFCellStyle dateStyle = book.createCellStyle();
        dateStyle.setDataFormat(format.getFormat("dd.mm.yyyy"));
        birthdate.setCellStyle(dateStyle);


        // Нумерация лет начинается с 1900-го
        birthdate.setCellValue(new Date(110, 10, 10));

        // Меняем размер столбца
        sheet.autoSizeColumn((short)1);

        // Записываем всё в файл
        book.write(new FileOutputStream(file));
    }

    public static void readFromExcel(String file) throws IOException{
        HSSFWorkbook myExcelBook = new HSSFWorkbook(new FileInputStream(file));
        HSSFSheet myExcelSheet = myExcelBook.getSheet("Birthdays");
        HSSFRow row = myExcelSheet.getRow(0);

        if(row.getCell(0).getCellType() == HSSFCell.CELL_TYPE_STRING){
            String name = row.getCell(0).getStringCellValue();
            System.out.println("name : " + name);
        }

        if(row.getCell(1).getCellType() == HSSFCell.CELL_TYPE_NUMERIC){
            Date birthdate = row.getCell(1).getDateCellValue();
            System.out.println("birthdate :" + birthdate);
        }

        //myExcelBook.close();

    }

}

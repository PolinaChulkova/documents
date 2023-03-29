package com.example.documents.xlsx;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

@Service
public class ExcelReaderService {

    public void readExcelDocument() {
        try (FileInputStream inputStream = new FileInputStream("excelFile.xlsx")){
            Workbook workbook = new XSSFWorkbook(inputStream);

            Sheet sheet = workbook.getSheetAt(0);

            Iterator<Row> rowIterator = sheet.rowIterator();
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    System.out.print(cell.getStringCellValue() + " ");
                }
                System.out.println(" ");
            }

        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public void createExcelDocument() {
        // получаем файл в формате xls
        try (FileInputStream inputStream = new FileInputStream(new File("excelFile.xlsx"))) {
            // формируем из файла экземпляр HSSFWorkbook
            XSSFWorkbook excelBook = new XSSFWorkbook(inputStream);

            // выбираем первый лист для обработки
            // нумерация начинается с 0
            XSSFSheet sheet = excelBook.getSheetAt(0);
            XSSFRow row = sheet.getRow(0);

            // получаем Iterator по всем строкам в листе
            Iterator<Row> rowIterator = sheet.iterator();
            // получаем Iterator по всем ячейкам в строке
            Iterator<Cell> cellIterator = row.cellIterator();

        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}

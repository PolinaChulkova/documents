package com.example.documents.xlsx;

import com.example.documents.Employee;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

@Service
public class ExcelWriterService {

    public void createExcel(List<Employee> employees) {
        try {
            Workbook workbook = new XSSFWorkbook();

            Sheet sheet = workbook.createSheet("Employee data");
            sheet.setColumnWidth(0, 6000);
            sheet.setColumnWidth(0, 4000);

            Map<String, Object[]> data = new TreeMap<>();
            data.put("0", new Object[]{"ID", "First name", "Last name", "Position"});
            for (Employee employee : employees) {
                data.put(employee.getId().toString(), new Object[]{
                        employee.getId().toString(),
                        employee.getFirstName(),
                        employee.getLastName(),
                        employee.getPosition()});
            }

            Set<String> keyset = data.keySet();
            int rownum = 0;
            for (String key : keyset) {
                Row row = sheet.createRow(rownum++);
                Object[] objArray = data.get(key);
                int cellnum = 0;
                for (Object obj : objArray) {
                    Cell cell = row.createCell(cellnum++);
                    cell.setCellValue((String) obj);
                }
            }

            FileOutputStream outputStream = new FileOutputStream(new File("excelFile.xlsx"));
            workbook.write(outputStream);
            outputStream.close();

        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}

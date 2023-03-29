package com.example.documents;

import com.example.documents.docx.WordReaderService;
import com.example.documents.docx.WordWriterService;
import com.example.documents.xlsx.ExcelReaderService;
import com.example.documents.xlsx.ExcelWriterService;
import lombok.RequiredArgsConstructor;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
@RequestMapping("/")
@RequiredArgsConstructor
public class Controller {

    private final EmployeeService employeeService;
    private final WordWriterService wordWriterService;
    private final WordReaderService wordReaderService;
    private final ExcelReaderService excelReaderService;
    private final ExcelWriterService excelWriterService;

    @PostMapping("generate-docx")
    public void generateWord() {
        wordWriterService.createWord(employeeService.addEmployees());
    }

    @GetMapping("read-docx")
    public void readWord() {
        wordReaderService.readWordDocument();
    }

    @PostMapping("generate-xlsx")
    public void generateExcel() {
        excelWriterService.createExcel(employeeService.addEmployees());
    }

    @GetMapping("read-xlsx")
    public void readExcel() {
        excelReaderService.readExcelDocument();
    }
}

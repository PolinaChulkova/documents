package com.example.documents.controller;

import com.example.documents.service.EmployeeService;
import com.example.documents.service.docx.WordReaderService;
import com.example.documents.service.docx.WordWriterService;
import com.example.documents.service.xlsx.ExcelReaderService;
import com.example.documents.service.xlsx.ExcelWriterService;
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
        wordWriterService.createDocxDocument(employeeService.addEmployees());
    }

    @GetMapping("read-docx")
    public void readWord() {
        wordReaderService.readDocxDocument();
    }

    @PostMapping("generate-xlsx")
    public void generateExcel() {
        excelWriterService.createXlsxDocument(employeeService.addEmployees());
    }

    @GetMapping("read-xlsx")
    public void readExcel() {
        excelReaderService.readXlsxDocument();
    }
}

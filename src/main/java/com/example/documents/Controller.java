package com.example.documents;

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

    @PostMapping("generate-docx")
    public void generateWord() {
        wordWriterService.createWord(employeeService.addEmployees());
    }

    @GetMapping("read-docx")
    public void readWord() {
        wordReaderService.readWordDocument();
    }
}

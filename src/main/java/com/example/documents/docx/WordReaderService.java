package com.example.documents.docx;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.springframework.stereotype.Service;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.List;

@Service
public class WordReaderService {
    public void readWordDocument() {
        try(FileInputStream inputStream = new FileInputStream("testFile2.docx")) {
            XWPFDocument document = new XWPFDocument(OPCPackage.open("testFile2.docx"));
            XWPFHeaderFooterPolicy headerFooterPolicy = new XWPFHeaderFooterPolicy(document);

            XWPFHeader docHeader = headerFooterPolicy.getDefaultHeader();
            System.out.println(docHeader.getText());

            List<XWPFParagraph> paragraphs = document.getParagraphs();
            for (XWPFParagraph paragraph : paragraphs) {
                System.out.println(paragraph.getText());
            }

            XWPFFooter docFooter = headerFooterPolicy.getDefaultFooter();
            System.out.println(docFooter.getText());

        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        } catch (InvalidFormatException e) {
            throw new RuntimeException(e);
        }
    }
}

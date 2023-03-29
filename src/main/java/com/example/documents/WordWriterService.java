package com.example.documents;

import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;
import org.springframework.stereotype.Service;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

@Service
public class WordWriterService {
    public void createWord(List<Employee> employees) {
        try {
            // создаем модель docx документа,
            // к которой будем прикручивать наполнение (колонтитулы, текст)
            XWPFDocument document = new XWPFDocument();
            CTSectPr ctSectPr = document.getDocument().getBody().getSectPr();

            // получаем экземпляр XWPFHeaderFooterPolicy для работы с колонтитулами
            XWPFHeaderFooterPolicy headerFooterPolicy = new XWPFHeaderFooterPolicy(document);
            // создаем верхний колонтитул Word файла
            CTP ctpHeaderModel = createModel("Количество работников: " + employees.size());
            // устанавливаем сформированный верхний
            // колонтитул в модель документа Word
            XWPFParagraph headerParagraph = new XWPFParagraph(ctpHeaderModel, document);
            headerFooterPolicy.createHeader(
                    XWPFHeaderFooterPolicy.DEFAULT,
                    new XWPFParagraph[]{headerParagraph}
            );

            // создаем нижний колонтитул docx файла
            CTP ctpFooterModel = createModel("Просто нижний колонтитул");
            // устанавливаем сформированый нижний
            // колонтитул в модель документа Word
            XWPFParagraph footerParagraph = new XWPFParagraph(ctpFooterModel, document);
            headerFooterPolicy.createFooter(
                    XWPFHeaderFooterPolicy.DEFAULT,
                    new XWPFParagraph[]{footerParagraph}
            );


            for (Employee employee : employees) {
                // создаем обычный параграф, который будет расположен слева,
                // будет синим курсивом со шрифтом 25 размера
                XWPFParagraph bodyParagraph = document.createParagraph();
                bodyParagraph.setAlignment(ParagraphAlignment.RIGHT);
                XWPFRun paragraphConfig = bodyParagraph.createRun();
                paragraphConfig.setItalic(true);
                paragraphConfig.setFontSize(14);
                paragraphConfig.setColor("06357a");
                paragraphConfig.setText(String.format("Id работника: %d \n Имя работника: %s \n Фамилия работника: %s \n" +
                                "Должность работника: %s \n", employee.getId(), employee.getFirstName(),
                        employee.getLastName(), employee.getPosition()));

            }


            // сохраняем модель docx документа в файл
            FileOutputStream fileOutputStream = new FileOutputStream("testFile2.docx");
            document.write(fileOutputStream);
            fileOutputStream.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public CTP createModel(String content) {
        // создаем футер или нижний колонтитул
        CTP ctpHeaderModel = CTP.Factory.newInstance();
        CTR ctrHeaderModel = ctpHeaderModel.addNewR();
        CTText cttHeader = ctrHeaderModel.addNewT();

        cttHeader.setStringValue(content);
        return ctpHeaderModel;
    }
}

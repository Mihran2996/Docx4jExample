package com.example.demo;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.io.File;

@SpringBootApplication
public class DemoApplication {

    public static void main(String[] args) throws Exception {
        SpringApplication.run(DemoApplication.class, args);
        WordprocessingMLPackage wordPackage = WordprocessingMLPackage.createPackage();
        MainDocumentPart mainDocumentPart = wordPackage.getMainDocumentPart();
        File export = new File("file1.docx");

        mainDocumentPart.getContent().add(Util.addImage(wordPackage));
        mainDocumentPart.addStyledParagraphOfText("Title", "Resume ");
        for (int i = 0; i < 5; i++) {
            mainDocumentPart.getContent().add(Util.addFields(i));
        }
        mainDocumentPart.getContent().add(Util.createTable(wordPackage));
        wordPackage.save(export);
    }

}

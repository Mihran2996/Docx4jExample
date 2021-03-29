package com.example.demo;

import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.jaxb.Context;
import org.docx4j.model.table.TblFactory;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.wml.*;

import java.io.File;
import java.nio.file.Files;
import java.util.List;

public class Util {

    public static P addImage(WordprocessingMLPackage wordPackage) throws Exception {
        File image = new File("img.jpg");
        byte[] fileContent = Files.readAllBytes(image.toPath());
        BinaryPartAbstractImage imagePart = BinaryPartAbstractImage.createImagePart(wordPackage, fileContent);
        Inline inline = imagePart.createImageInline("Baeldung Image (filename hint)", "Alt Text", 1, 2, false);
        return addImageToParagraph(inline);
    }

    private static P addImageToParagraph(Inline inline) {
        ObjectFactory factory = new ObjectFactory();
        P p = factory.createP();
        R r = factory.createR();
        p.getContent().add(r);
        Drawing drawing = factory.createDrawing();
        r.getContent().add(drawing);
        drawing.getAnchorOrInline().add(inline);
        return p;
    }

    public static P addFields(int n) {
        ObjectFactory objectFactory = Context.getWmlObjectFactory();
        RPr rPr = objectFactory.createRPr();
        BooleanDefaultTrue b = new BooleanDefaultTrue();
        Color green = objectFactory.createColor();
        green.setVal("green");
        rPr.setB(b);
        rPr.setI(b);
        rPr.setShadow(b);
        rPr.setColor(green);
        for (int i = n; i < 5; i++) {
            Text text = objectFactory.createText();
            if (i == 0) text.setValue("Name `   Poxos");
            if (i == 1) text.setValue("Surname `   Poxosyan");
            if (i == 2) text.setValue("age `   22");
            if (i == 3) text.setValue("Email `   poxos@mail.com");
            if (i == 4) text.setValue("PhoneNumber `  +374 939392");
            P p = objectFactory.createP();
            R r = objectFactory.createR();
            r.setRPr(rPr);
            r.getContent().add(text);
            p.getContent().add(r);
            return p;
        }
        return new P();
    }

    public static Tbl createTable(WordprocessingMLPackage wordPackage) {
        ObjectFactory wmlObjectFactory = Context.getWmlObjectFactory();
        P p = wmlObjectFactory.createP();
        R r = wmlObjectFactory.createR();
        Text t = wmlObjectFactory.createText();
        t.setValue("Empty");
        r.getContent().add(t);
        p.getContent().add(r);
        int writableWidthTwips = wordPackage.getDocumentModel().getSections().get(0).getPageDimensions().getWritableWidthTwips();
        int columnNumber = 3;
        Tbl tbl = TblFactory.createTable(3, 3, writableWidthTwips / columnNumber);
        List<Object> rows = tbl.getContent();
        for (Object row : rows) {
            Tr tr = (Tr) row;
            List<Object> cells = tr.getContent();
            for (Object cell : cells) {
                Tc td = (Tc) cell;
                td.getContent().add(p);
            }
        }
        return tbl;
    }


}

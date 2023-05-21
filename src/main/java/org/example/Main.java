package org.example;

import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.PageSize;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.pdf.PdfWriter;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;


import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class Main {
    public static void main(String[] args) {
        String templateFile = "C:/Users/Juria/IdeaProjects/DocumentGenerator/src/main/java/org/example/template.docx";
        String[] districts = {"Paramaribo Noord-Oost", "Paramaribo Zuid-West"};

        for (String district : districts) {
            String outputFile = "VergunningAanvraag_" + district + ".docx";

            replacePlaceholder(templateFile, outputFile, district, "22 mei 2023","OX88 Midyear Rally", "Marcellino Chen", "+597 856-7199");

            System.out.println("Generated: " + outputFile);
        }
    }

    private static void replacePlaceholder(String inputFile, String outputFile, String district, String datum, String rallynaam, String uitzetternaam, String uitzettertelefoon) {
        try {
            FileInputStream fis = new FileInputStream(inputFile);
            XWPFDocument document = new XWPFDocument(fis);
            for (XWPFParagraph paragraph : document.getParagraphs()) {
                for (XWPFRun run : paragraph.getRuns()) {
                    String text = run.getText(0);
                    if (text != null && text.contains("[district]")) {
                        run.setText(text.replace("[district]", district), 0);
                    }
                    if (text != null && text.contains("[datum]")) {
                        run.setText(text.replace("[datum]", datum), 0);
                    }
                    if (text != null && text.contains("[rallynaam]")) {
                        run.setText(text.replace("[rallynaam]", rallynaam), 0);
                    }
                    if (text != null && text.contains("[uitzetter]")) {
                        run.setText(text.replace("[uitzetter]", uitzetternaam), 0);
                    }
                    if (text != null && text.contains("[telefoon]")) {
                        run.setText(text.replace("[telefoon]", uitzettertelefoon), 0);
                    }
                }
            }

            FileOutputStream fos = new FileOutputStream(outputFile);
            document.write(fos);

            fis.close();
            fos.close();
            document.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void convertToPdf(String inputFile, String outputFile) {
        try {
            FileInputStream fis = new FileInputStream(inputFile);
            XWPFDocument document = new XWPFDocument(fis);
            Document pdfDocument = new Document(PageSize.A4);
            PdfWriter.getInstance(pdfDocument, new FileOutputStream(outputFile));
            pdfDocument.open();
            for (XWPFParagraph paragraph : document.getParagraphs()) {
                String text = paragraph.getText();
                pdfDocument.add(new Paragraph(text));
            }
            pdfDocument.close();
            fis.close();
            document.close();
        } catch (IOException | DocumentException e) {
            e.printStackTrace();
        }
    }
}
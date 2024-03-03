package org.example;

import org.apache.poi.xslf.usermodel.*;

import java.awt.*;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;


public class Main {
    public static void main(String[] args) throws IOException {
        System.out.println("Hello world!");
        Path path = Paths.get("", "src/main/resources/Portfolio.pptx");
        XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(path.toString()));
        XSLFSlide slide =  ppt.getSlides().get(0);
        // Iterate through each shape in the slide
        for (XSLFShape shape : slide.getShapes()) {
            // Check if the shape is a text box
            if (shape instanceof org.apache.poi.xslf.usermodel.XSLFTextShape) {
                // Get the text content of the shape
                String text = ((org.apache.poi.xslf.usermodel.XSLFTextShape) shape).getText();

                // Search for the text "name" (case-insensitive)
                if (text != null && text.toLowerCase().contains("name")) {
                    // Iterate through each text run in the shape
                    for (XSLFTextRun textRun : ((org.apache.poi.xslf.usermodel.XSLFTextShape) shape).getTextParagraphs().get(0).getTextRuns()) {
                        // Search for the text "name" (case-insensitive) within the text run
                        String runText = textRun.getRawText();
                        if (runText != null && runText.toLowerCase().contains("name")) {
                            // Add a hyperlink to the text and set font color to red
                            XSLFHyperlink hyperlink = textRun.createHyperlink();
                            hyperlink.setAddress("http://example.com"); // Set your hyperlink URL
                            textRun.setFontColor(Color.RED);
                            System.out.println("Found required text and added a hyperlink");
                        }
                    }
                }
            }
        }

        ppt.write(new FileOutputStream("output.pptx"));
    }
}
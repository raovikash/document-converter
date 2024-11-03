package com.docviewer.service;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.*;
import org.apache.poi.xwpf.usermodel.*;
import org.springframework.stereotype.Service;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.Base64;
import java.util.regex.Pattern;

@Service
public class DocConverterService {
    
    // Pattern to validate base64 string
    private static final Pattern BASE64_PATTERN = Pattern.compile("^[A-Za-z0-9+/]*={0,2}$");
    
    public String convertDocToDocx(String base64Doc) throws IOException {
        // Validate input
        if (base64Doc == null || base64Doc.trim().isEmpty()) {
            throw new IllegalArgumentException("Base64 input cannot be null or empty");
        }
        
        // Only remove whitespace if it exists to preserve original content when possible
        if (base64Doc.matches(".*\\s+.*")) {
            base64Doc = base64Doc.replaceAll("\\s+", "");
        }
        
        // Validate base64 format
        if (!BASE64_PATTERN.matcher(base64Doc).matches()) {
            throw new IllegalArgumentException("Invalid base64 format. Please ensure the input is properly base64 encoded.");
        }
        
        try {
            // Decode base64 to byte array
            byte[] docBytes = Base64.getDecoder().decode(base64Doc);
            
            try (ByteArrayInputStream docInputStream = new ByteArrayInputStream(docBytes);
                 HWPFDocument doc = new HWPFDocument(docInputStream);
                 ByteArrayOutputStream docxOutputStream = new ByteArrayOutputStream()) {
                
                // Create new DOCX document
                XWPFDocument docx = new XWPFDocument();
                
                // Get the range of text from DOC
                Range range = doc.getRange();
                
                // Iterate through paragraphs
                for (int i = 0; i < range.numParagraphs(); i++) {
                    Paragraph docParagraph = range.getParagraph(i);
                    XWPFParagraph docxParagraph = docx.createParagraph();
                    
                    // Copy paragraph properties
                    docxParagraph.setAlignment(convertAlignment(docParagraph.getJustification()));
                    docxParagraph.setSpacingBefore(docParagraph.getSpacingBefore());
                    docxParagraph.setSpacingAfter(docParagraph.getSpacingAfter());
                    
                    // Iterate through runs (text with same formatting)
                    for (int j = 0; j < docParagraph.numCharacterRuns(); j++) {
                        CharacterRun docRun = docParagraph.getCharacterRun(j);
                        String text = docRun.text();
                        
                        if (text.trim().length() > 0) {
                            XWPFRun docxRun = docxParagraph.createRun();
                            docxRun.setText(text);
                            
                            // Copy run properties
                            docxRun.setBold(docRun.isBold());
                            docxRun.setItalic(docRun.isItalic());
                            docxRun.setUnderline(docRun.getUnderlineCode() != 0 ? 
                                UnderlinePatterns.SINGLE : UnderlinePatterns.NONE);
                            docxRun.setFontSize(docRun.getFontSize());
                            docxRun.setFontFamily(docRun.getFontName());
                            
                            // Handle text color
                            int color = docRun.getColor();
                            if (color != -1) {
                                docxRun.setColor(String.format("%06X", color));
                            }
                        }
                    }
                }
                
                // Write DOCX to output stream
                docx.write(docxOutputStream);
                docx.close();
                
                // Convert to base64
                byte[] docxBytes = docxOutputStream.toByteArray();
                return Base64.getEncoder().encodeToString(docxBytes);
            }
        } catch (IllegalArgumentException e) {
            throw new IllegalArgumentException("Failed to decode base64 content: " + e.getMessage());
        } catch (IOException e) {
            throw new IOException("Failed to process document: " + e.getMessage());
        }
    }
    
    private ParagraphAlignment convertAlignment(int justification) {
        switch (justification) {
            case 1: return ParagraphAlignment.LEFT;
            case 2: return ParagraphAlignment.CENTER;
            case 3: return ParagraphAlignment.RIGHT;
            case 4: return ParagraphAlignment.BOTH;
            default: return ParagraphAlignment.LEFT;
        }
    }
}

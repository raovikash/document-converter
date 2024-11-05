package com.docviewer.service;

import com.artofsolving.jodconverter.DocumentConverter;
import com.artofsolving.jodconverter.openoffice.connection.OpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.connection.SocketOpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.converter.OpenOfficeDocumentConverter;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.*;
import org.apache.poi.xwpf.usermodel.*;
import org.springframework.stereotype.Service;
import java.io.*;
import java.nio.file.Files;
import java.util.Base64;
import java.util.regex.Pattern;

@Service
public class DocConverterService {
    
    private static final Pattern BASE64_PATTERN = Pattern.compile("^[A-Za-z0-9+/]*={0,2}$");
    private static final String OPENOFFICE_HOST = "127.0.0.1";
    private static final int OPENOFFICE_PORT = 8100;

    public String convertDocToDocxUsingJod(String base64Doc) throws IOException {
        // Validate input
        if (base64Doc == null || base64Doc.trim().isEmpty()) {
            throw new IllegalArgumentException("Base64 input cannot be null or empty");
        }
        
        if (base64Doc.matches(".*\\s+.*")) {
            base64Doc = base64Doc.replaceAll("\\s+", "");
        }
        
        if (!BASE64_PATTERN.matcher(base64Doc).matches()) {
            throw new IllegalArgumentException("Invalid base64 format");
        }

        // Create temporary files for input and output
        File inputFile = File.createTempFile("input", ".doc");
        File outputFile = File.createTempFile("output", ".docx");
        
        OpenOfficeConnection connection = null;
        
        try {
            // Decode base64 and write to input file
            byte[] docBytes = Base64.getDecoder().decode(base64Doc);
            Files.write(inputFile.toPath(), docBytes);
            
            // Connect to OpenOffice/LibreOffice
            connection = new SocketOpenOfficeConnection(OPENOFFICE_HOST, OPENOFFICE_PORT);
            connection.connect();
            
            // Convert the document
            DocumentConverter converter = new OpenOfficeDocumentConverter(connection);
            converter.convert(inputFile, outputFile);
            
            // Read the output file and convert to base64
            byte[] docxBytes = Files.readAllBytes(outputFile.toPath());
            return Base64.getEncoder().encodeToString(docxBytes);
            
        } finally {
            // Clean up
            if (connection != null && connection.isConnected()) {
                try {
                    connection.disconnect();
                } catch (Exception e) {
                    System.err.println("Warning: Failed to disconnect: " + e.getMessage());
                }
            }
            
            try {
                Files.deleteIfExists(inputFile.toPath());
                Files.deleteIfExists(outputFile.toPath());
            } catch (IOException e) {
                System.err.println("Warning: Failed to delete temporary files: " + e.getMessage());
            }
        }
    }
    
    // Original Apache POI implementation kept as fallback
    public String convertDocToDocx(String base64Doc) throws IOException {
        // Validate input
        if (base64Doc == null || base64Doc.trim().isEmpty()) {
            throw new IllegalArgumentException("Base64 input cannot be null or empty");
        }
        
        if (base64Doc.matches(".*\\s+.*")) {
            base64Doc = base64Doc.replaceAll("\\s+", "");
        }
        
        if (!BASE64_PATTERN.matcher(base64Doc).matches()) {
            throw new IllegalArgumentException("Invalid base64 format");
        }
        
        try {
            byte[] docBytes = Base64.getDecoder().decode(base64Doc);
            
            try (ByteArrayInputStream docInputStream = new ByteArrayInputStream(docBytes);
                 HWPFDocument doc = new HWPFDocument(docInputStream);
                 ByteArrayOutputStream docxOutputStream = new ByteArrayOutputStream()) {
                
                XWPFDocument docx = new XWPFDocument();
                Range range = doc.getRange();
                
                for (int i = 0; i < range.numParagraphs(); i++) {
                    Paragraph docParagraph = range.getParagraph(i);
                    XWPFParagraph docxParagraph = docx.createParagraph();
                    
                    docxParagraph.setAlignment(convertAlignment(docParagraph.getJustification()));
                    docxParagraph.setSpacingBefore(docParagraph.getSpacingBefore());
                    docxParagraph.setSpacingAfter(docParagraph.getSpacingAfter());
                    
                    for (int j = 0; j < docParagraph.numCharacterRuns(); j++) {
                        CharacterRun docRun = docParagraph.getCharacterRun(j);
                        String text = docRun.text();
                        
                        if (text.trim().length() > 0) {
                            XWPFRun docxRun = docxParagraph.createRun();
                            docxRun.setText(text);
                            
                            docxRun.setBold(docRun.isBold());
                            docxRun.setItalic(docRun.isItalic());
                            docxRun.setUnderline(docRun.getUnderlineCode() != 0 ? 
                                UnderlinePatterns.SINGLE : UnderlinePatterns.NONE);
                            docxRun.setFontSize(docRun.getFontSize() / 2);
                            docxRun.setFontFamily(docRun.getFontName());
                            
                            int color = docRun.getColor();
                            if (color != -1) {
                                docxRun.setColor(String.format("%06X", color));
                            }
                        }
                    }
                }
                
                docx.write(docxOutputStream);
                docx.close();
                
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

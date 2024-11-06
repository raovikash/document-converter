package com.docviewer.service;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.*;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageMar;
import org.springframework.stereotype.Service;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.Base64;
import java.util.regex.Pattern;
import java.math.BigInteger;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

@Service
public class DocConverterService {
    
    private static final Logger logger = LoggerFactory.getLogger(DocConverterService.class);
    private static final Pattern BASE64_PATTERN = Pattern.compile("^[A-Za-z0-9+/]*={0,2}$");
    private static final int DEFAULT_MARGIN = 10; 
    
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
                
                // Set document margins
                setDocumentMargins(docx);
                
                Range range = doc.getRange();
                
                // Process sections for page layout
                for (int sectionIdx = 0; sectionIdx < range.numSections(); sectionIdx++) {
                    Section section = range.getSection(sectionIdx);
                    
                    // Process paragraphs in each section
                    for (int i = 0; i < section.numParagraphs(); i++) {
                        Paragraph docParagraph = section.getParagraph(i);
                        XWPFParagraph docxParagraph = docx.createParagraph();
                        
                        // Set paragraph style
                        setParagraphStyle(docParagraph, docxParagraph);
                        
                        // Handle bullet points and numbering
                        if (docParagraph.isInList()) {
                            handleListFormatting(docParagraph, docxParagraph);
                        }
                        
                        // Process character runs (text with formatting)
                        for (int j = 0; j < docParagraph.numCharacterRuns(); j++) {
                            CharacterRun docRun = docParagraph.getCharacterRun(j);
                            String text = docRun.text();
                            
                            if (text.trim().length() > 0) {
                                XWPFRun docxRun = docxParagraph.createRun();
                                setRunProperties(docRun, docxRun, text);
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
            logger.error("Failed to decode base64 content", e);
            throw new IllegalArgumentException("Failed to decode base64 content: " + e.getMessage());
        } catch (IOException e) {
            logger.error("Failed to process document", e);
            throw new IOException("Failed to process document: " + e.getMessage());
        }
    }
    
    private void setDocumentMargins(XWPFDocument docx) {
        CTSectPr sectPr = docx.getDocument().getBody().isSetSectPr() ? 
            docx.getDocument().getBody().getSectPr() : 
            docx.getDocument().getBody().addNewSectPr();
            
        CTPageMar pageMar = sectPr.isSetPgMar() ? 
            sectPr.getPgMar() : 
            sectPr.addNewPgMar();
        
        pageMar.setLeft(BigInteger.valueOf(DEFAULT_MARGIN));
        // pageMar.setRight(BigInteger.valueOf(DEFAULT_MARGIN));
        // pageMar.setTop(BigInteger.valueOf(DEFAULT_MARGIN));
        // pageMar.setBottom(BigInteger.valueOf(DEFAULT_MARGIN));
    }
    
    private void setParagraphStyle(Paragraph docParagraph, XWPFParagraph docxParagraph) {
        // Basic alignment
        docxParagraph.setAlignment(convertAlignment(docParagraph.getJustification()));
        
        // Spacing
        docxParagraph.setSpacingBefore(docParagraph.getSpacingBefore());
        docxParagraph.setSpacingAfter(docParagraph.getSpacingAfter());
        
        // Line spacing
        LineSpacingDescriptor spacing = docParagraph.getLineSpacing();
        if (spacing != null) {
            docxParagraph.setSpacingLineRule(LineSpacingRule.AUTO);
            docxParagraph.setSpacingBetween(1.15); // Default spacing
        }
        
        // Indentation
        int leftIndent = docParagraph.getIndentFromLeft();
        int rightIndent = docParagraph.getIndentFromRight();
        int firstLineIndent = docParagraph.getFirstLineIndent();
        docxParagraph.setIndentationLeft(leftIndent);
        docxParagraph.setIndentationRight(rightIndent);
        docxParagraph.setFirstLineIndent(firstLineIndent);
    }
    
    private void handleListFormatting(Paragraph docParagraph, XWPFParagraph docxParagraph) {
        int level = docParagraph.getIlvl();
        
        if (level >= 0) {
            BigInteger numId = BigInteger.valueOf(1);
            docxParagraph.setNumID(numId);
            docxParagraph.setNumILvl(BigInteger.valueOf(level));
            
            if (level == 0) {
                docxParagraph.setStyle("ListBullet");
            } else {
                docxParagraph.setStyle("ListNumber");
            }
        }
    }
    
    private void setRunProperties(CharacterRun docRun, XWPFRun docxRun, String text) {
        // Set text
        docxRun.setText(text);
        
        // Basic formatting
        docxRun.setBold(docRun.isBold());
        docxRun.setItalic(docRun.isItalic());
        docxRun.setStrike(docRun.isStrikeThrough());
        
        // Handle subscript/superscript
        int scriptIndex = docRun.getSubSuperScriptIndex();
        if (scriptIndex == 1) {
            docxRun.setVerticalAlignment("subscript");
        } else if (scriptIndex == 2) {
            docxRun.setVerticalAlignment("superscript");
        }
        
        // Font properties
        docxRun.setFontSize(docRun.getFontSize() / 2);
        docxRun.setFontFamily(docRun.getFontName());
        
        // Underline
        if (docRun.getUnderlineCode() != 0) {
            docxRun.setUnderline(convertUnderline(docRun.getUnderlineCode()));
        }
        
        // Text color
        int color = docRun.getColor();
        if (color != -1) {
            docxRun.setColor(String.format("%06X", color));
        }
        
        // Text effects
        docxRun.setEmbossed(docRun.isEmbossed());
        docxRun.setImprinted(docRun.isImprinted());
        docxRun.setShadow(docRun.isShadowed());
        
        // Character spacing
        if (docRun.getKerning() != 0) {
            docxRun.setKerning(docRun.getKerning());
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
    
    private UnderlinePatterns convertUnderline(int code) {
        switch (code) {
            case 1: return UnderlinePatterns.SINGLE;
            case 2: return UnderlinePatterns.DOUBLE;
            case 3: return UnderlinePatterns.DOTTED;
            case 4: return UnderlinePatterns.DASH;
            case 5: return UnderlinePatterns.WORDS;
            case 6: return UnderlinePatterns.THICK;
            case 7: return UnderlinePatterns.WAVE;
            default: return UnderlinePatterns.NONE;
        }
    }
}

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
    private static final int DEFAULT_MARGIN = 100; // 1 inch in twips
    private static final String ARROW_BULLET = "►"; // Arrow bullet point
    
    public String convertDocToDocx(String base64Doc) throws IOException {
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
                setDocumentMargins(docx);
                
                Range range = doc.getRange();
                
                for (int sectionIdx = 0; sectionIdx < range.numSections(); sectionIdx++) {
                    Section section = range.getSection(sectionIdx);
                    
                    for (int i = 0; i < section.numParagraphs(); i++) {
                        Paragraph docParagraph = section.getParagraph(i);
                        XWPFParagraph docxParagraph = docx.createParagraph();
                        
                        // Set default alignment to left for all paragraphs
                        docxParagraph.setAlignment(ParagraphAlignment.LEFT);
                        
                        setParagraphStyle(docParagraph, docxParagraph);
                        
                        // Handle lists and bullet points
                        if (docParagraph.isInList() || startsWithArrow(docParagraph.text())) {
                            handleListFormatting(docParagraph, docxParagraph);
                        }
                        
                        // Process text and formatting
                        processCharacterRuns(docParagraph, docxParagraph);
                    }
                }
                
                docx.write(docxOutputStream);
                docx.close();
                
                byte[] docxBytes = docxOutputStream.toByteArray();
                return Base64.getEncoder().encodeToString(docxBytes);
            }
        } catch (Exception e) {
            logger.error("Failed to convert document", e);
            throw new IOException("Failed to convert document: " + e.getMessage());
        }
    }
    
    private void setDocumentMargins(XWPFDocument docx) {
        CTSectPr sectPr = docx.getDocument().getBody().isSetSectPr() ? 
            docx.getDocument().getBody().getSectPr() : 
            docx.getDocument().getBody().addNewSectPr();
            
        CTPageMar pageMar = sectPr.isSetPgMar() ? 
            sectPr.getPgMar() : 
            sectPr.addNewPgMar();
        
        // Set margins
        pageMar.setRight(BigInteger.valueOf(DEFAULT_MARGIN));
        
        // Set gutter margin to 0
        pageMar.setGutter(BigInteger.ZERO);
    }
    
    private boolean startsWithArrow(String text) {
        return text.trim().startsWith("►") || text.trim().startsWith(">");
    }
    
    private void processCharacterRuns(Paragraph docParagraph, XWPFParagraph docxParagraph) {
        String paragraphText = docParagraph.text().trim();
        boolean isArrowBullet = startsWithArrow(paragraphText);
        
        for (int j = 0; j < docParagraph.numCharacterRuns(); j++) {
            CharacterRun docRun = docParagraph.getCharacterRun(j);
            String text = docRun.text();
            
            if (text.trim().length() > 0) {
                if (isArrowBullet && j == 0) {
                    // Handle arrow bullet point
                    XWPFRun bulletRun = docxParagraph.createRun();
                    bulletRun.setText(ARROW_BULLET + " ");
                    bulletRun.setFontFamily("Symbol");
                    bulletRun.setFontSize(10);
                    
                    // Create another run for the actual text
                    XWPFRun textRun = docxParagraph.createRun();
                    setRunProperties(docRun, textRun, text.replaceFirst("^[►>]\\s*", ""));
                } else if (text.contains("@") && text.contains(".")) {
                    createEmailLink(docxParagraph, text);
                } else {
                    XWPFRun docxRun = docxParagraph.createRun();
                    setRunProperties(docRun, docxRun, text);
                }
            }
        }
    }
    
    private void createEmailLink(XWPFParagraph paragraph, String email) {
        XWPFRun run = paragraph.createRun();
        run.setText(email);
        run.setUnderline(UnderlinePatterns.SINGLE);
        run.setColor("0000FF"); // Blue color
        run.setFontFamily("Times New Roman");
        run.setFontSize(11);
    }
    
    private void setParagraphStyle(Paragraph docParagraph, XWPFParagraph docxParagraph) {
        // Always set left alignment as default
        docxParagraph.setAlignment(ParagraphAlignment.LEFT);
        
        // Spacing
        docxParagraph.setSpacingBefore(docParagraph.getSpacingBefore());
        docxParagraph.setSpacingAfter(docParagraph.getSpacingAfter());
        
        // Line spacing
        docxParagraph.setSpacingLineRule(LineSpacingRule.EXACT);
        docxParagraph.setSpacingBetween(1.15);
        
        // Reset any existing indentation
        docxParagraph.setIndentationLeft(0);
        docxParagraph.setIndentationRight(0);
        docxParagraph.setFirstLineIndent(0);
        
        // Only set indentation if specifically needed
        int leftIndent = docParagraph.getIndentFromLeft();
        int rightIndent = docParagraph.getIndentFromRight();
        int firstLineIndent = docParagraph.getFirstLineIndent();
        
        if (leftIndent > 0) {
            docxParagraph.setIndentationLeft(leftIndent);
        }
        if (rightIndent > 0) {
            docxParagraph.setIndentationRight(rightIndent);
        }
        if (firstLineIndent != 0) {
            docxParagraph.setFirstLineIndent(firstLineIndent);
        }
        
        // Set borders based on the original paragraph's borders
        BorderCode topBorder = docParagraph.getTopBorder();
        BorderCode bottomBorder = docParagraph.getBottomBorder();
        BorderCode leftBorder = docParagraph.getLeftBorder();
        BorderCode rightBorder = docParagraph.getRightBorder();
        
        // Convert border styles
        if (topBorder != null && topBorder.getBorderType() > 0) {
            docxParagraph.setBorderTop(Borders.SINGLE);
        }
        if (bottomBorder != null && bottomBorder.getBorderType() > 0) {
            docxParagraph.setBorderBottom(Borders.SINGLE);
        }
        if (leftBorder != null && leftBorder.getBorderType() > 0) {
            docxParagraph.setBorderLeft(Borders.SINGLE);
        }
        if (rightBorder != null && rightBorder.getBorderType() > 0) {
            docxParagraph.setBorderRight(Borders.SINGLE);
        }
    }
    
    private void handleListFormatting(Paragraph docParagraph, XWPFParagraph docxParagraph) {
        // Check if it's an arrow bullet point
        if (startsWithArrow(docParagraph.text())) {
            docxParagraph.setIndentationLeft(360); // 0.25 inch
            return; // Skip regular list formatting
        }
        
        int level = docParagraph.getIlvl();
        
        if (level >= 0) {
            // Set numbering
            BigInteger numId = BigInteger.valueOf(1);
            docxParagraph.setNumID(numId);
            docxParagraph.setNumILvl(BigInteger.valueOf(level));
            
            // Add proper indentation for lists
            docxParagraph.setIndentationLeft(720); // 0.5 inch
            if (level > 0) {
                docxParagraph.setIndentationLeft(720 * (level + 1));
            }
            
            // Set list style
            if (level == 0) {
                docxParagraph.setStyle("ListNumber");
            } else {
                docxParagraph.setStyle("ListBullet");
            }
        }
    }
    
    private void setRunProperties(CharacterRun docRun, XWPFRun docxRun, String text) {
        docxRun.setText(text);
        
        // Font settings
        docxRun.setFontFamily("Times New Roman");
        docxRun.setFontSize(docRun.getFontSize() / 2);
        
        // Basic formatting
        docxRun.setBold(docRun.isBold());
        docxRun.setItalic(docRun.isItalic());
        docxRun.setStrike(docRun.isStrikeThrough());
        
        // Vertical alignment
        if (docRun.getSubSuperScriptIndex() == 1) {
            docxRun.setVerticalAlignment("subscript");
        } else if (docRun.getSubSuperScriptIndex() == 2) {
            docxRun.setVerticalAlignment("superscript");
        }
        
        // Underline - preserve original underline if it exists
        if (docRun.getUnderlineCode() != 0) {
            docxRun.setUnderline(convertUnderline(docRun.getUnderlineCode()));
        }
        
        // Text color - preserve original color
        int color = docRun.getColor();
        if (color != -1) {
            docxRun.setColor(String.format("%06X", color));
        }
        
        // Character spacing
        if (docRun.getKerning() != 0) {
            docxRun.setKerning(docRun.getKerning());
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

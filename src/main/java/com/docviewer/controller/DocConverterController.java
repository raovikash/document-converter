package com.docviewer.controller;

import com.docviewer.models.DocConvertRequest;
import com.docviewer.service.DocConverterService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

@RestController
@RequestMapping("/api/convert")
public class DocConverterController {

    private final DocConverterService docConverterService;

    @Autowired
    public DocConverterController(DocConverterService docConverterService) {
        this.docConverterService = docConverterService;
    }

    @PostMapping("/doc-to-docx")
    public ResponseEntity<?> convertDocToDocx(@RequestBody DocConvertRequest request) {
        try {
            if (request == null || request.getDocBase64() == null) {
                return ResponseEntity.badRequest()
                        .body(createErrorResponse("Request body or docBase64 content cannot be null"));
            }
            
            String base64Docx = docConverterService.convertDocToDocx(request.getDocBase64());
            
            Map<String, String> response = new HashMap<>();
            response.put("docxBase64", base64Docx);
            return ResponseEntity.ok().body(response);
            
        } catch (IllegalArgumentException e) {
            return ResponseEntity.badRequest()
                    .body(createErrorResponse("Invalid input: " + e.getMessage()));
        } catch (IOException e) {
            return ResponseEntity.badRequest()
                    .body(createErrorResponse("Error processing document: " + e.getMessage()));
        } catch (Exception e) {
            return ResponseEntity.internalServerError()
                    .body(createErrorResponse("Unexpected error: " + e.getMessage()));
        }
    }

    @PostMapping("/doc-to-docx/jod")
    public ResponseEntity<?> convertDocToDocxUsingJod(@RequestBody DocConvertRequest request) {
        try {
            if (request == null || request.getDocBase64() == null) {
                return ResponseEntity.badRequest()
                        .body(createErrorResponse("Request body or docBase64 content cannot be null"));
            }
            
            String base64Docx = docConverterService.convertDocToDocxUsingJod(request.getDocBase64());
            
            Map<String, String> response = new HashMap<>();
            response.put("docxBase64", base64Docx);
            return ResponseEntity.ok().body(response);
            
        } catch (IllegalArgumentException e) {
            return ResponseEntity.badRequest()
                    .body(createErrorResponse("Invalid input: " + e.getMessage()));
        } catch (IOException e) {
            return ResponseEntity.badRequest()
                    .body(createErrorResponse("Error processing document: " + e.getMessage()));
        } catch (Exception e) {
            return ResponseEntity.internalServerError()
                    .body(createErrorResponse("Unexpected error: " + e.getMessage()));
        }
    }
    
    private Map<String, String> createErrorResponse(String message) {
        Map<String, String> response = new HashMap<>();
        response.put("error", message);
        return response;
    }
}

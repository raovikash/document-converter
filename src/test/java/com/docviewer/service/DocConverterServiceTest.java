package com.docviewer.service;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.InjectMocks;
import org.mockito.junit.jupiter.MockitoExtension;
import org.springframework.boot.test.context.SpringBootTest;
import static org.junit.jupiter.api.Assertions.*;

import java.io.IOException;
import java.util.Base64;

@ExtendWith(MockitoExtension.class)
class DocConverterServiceTest {
    @InjectMocks
    private DocConverterService docConverterService;
    
    @BeforeEach
    void setUp() {
        docConverterService = new DocConverterService();
    }

    @Test
    @DisplayName("Should throw IllegalArgumentException for null input")
    void throwExceptionForNullInput() {
        IllegalArgumentException exception = assertThrows(
            IllegalArgumentException.class,
            () -> docConverterService.convertDocToDocx(null)
        );
        
        assertEquals("Base64 input cannot be null or empty", exception.getMessage());
    }

    @Test
    @DisplayName("Should throw IllegalArgumentException for empty input")
    void throwExceptionForEmptyInput() {
        IllegalArgumentException exception = assertThrows(
            IllegalArgumentException.class,
            () -> docConverterService.convertDocToDocx("")
        );
        
        assertEquals("Base64 input cannot be null or empty", exception.getMessage());
    }

    @Test
    @DisplayName("Should throw IllegalArgumentException for invalid base64 input")
    void throwExceptionForInvalidBase64Input() {
        String invalidBase64 = "This is not a valid base64 string!@#$";
        
        IllegalArgumentException exception = assertThrows(
            IllegalArgumentException.class,
            () -> docConverterService.convertDocToDocx(invalidBase64)
        );
        
        assertEquals("Invalid base64 format. Please ensure the input is properly base64 encoded.", 
                    exception.getMessage());
    }

    @Test
    @DisplayName("Should throw IllegalArgumentException for malformed base64 DOC")
    void throwExceptionForMalformedBase64Doc() {
        String malformedBase64 = "AAECAwQFBgc="; // Valid base64 but not a valid DOC file
        
        IllegalArgumentException exception = assertThrows(
            IllegalArgumentException.class,
            () -> docConverterService.convertDocToDocx(malformedBase64)
        );
        
        assertTrue(exception.getMessage().contains("Failed to decode base64 content"));
    }
}

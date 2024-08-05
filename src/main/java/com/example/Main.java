package com.example;
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;

@Controller
public class Main {
    

    private ProcessBuilder processBuilder;

    @GetMapping("/")
    public String uploadFile() {
        return "upload";
    }

    @PostMapping("/")
    public ResponseEntity<String> handleFileUpload(@RequestParam("file") MultipartFile file) throws IOException {
        
        if (file.isEmpty()) {
            return new ResponseEntity<>("No file uploaded", HttpStatus.BAD_REQUEST);
        }

        String fileName = file.getOriginalFilename();
        if (fileName == null || fileName.isEmpty()) {
            return new ResponseEntity<>("No file selected", HttpStatus.BAD_REQUEST);
        }
        try(InputStream inputStream=file.getInputStream()){
            Files.copy(inputStream, Paths.get("uploads/doc.docx"),StandardCopyOption.REPLACE_EXISTING);
        }
        if (isAllowedFile(fileName)) {
            String filePath = "uploads/doc.docx";
            try {
                String documentXmlPath = "uploads/";
                String documentRelsPath = "uploads/";
                String imageFolderPath = "uploads/";
                extractDocxContents(filePath, documentXmlPath, documentRelsPath, imageFolderPath);
                M2 m2=new M2();
                m2.fun();
                
                
            
                    return new ResponseEntity<>("document.xml, document.xml.rels, and images extracted and saved successfully", HttpStatus.OK);
                
            } catch (IOException e) {
                return new ResponseEntity<>("Error occurred during file processing", HttpStatus.INTERNAL_SERVER_ERROR);
            }
        } else {
            return new ResponseEntity<>("Invalid file type", HttpStatus.BAD_REQUEST);
        }
        
    }

    private boolean isAllowedFile(String fileName) {
        String[] parts = fileName.split("\\.");
        if (parts.length > 1) {
            String extension = parts[parts.length - 1].toLowerCase();
            return extension.equals("doc") || extension.equals("docx");
        }
        return false;
    }

    private void extractDocxContents(String docxFilePath, String documentXmlPath, String documentRelsPath, String imageFolderPath) throws IOException {
        try (ZipFile zipFile = new ZipFile(docxFilePath)) {
            ZipEntry documentXmlEntry = zipFile.getEntry("word/document.xml");
            ZipEntry documentRelsEntry = zipFile.getEntry("word/_rels/document.xml.rels");
            Files.createDirectories(Paths.get(imageFolderPath, "media"));

            if (documentXmlEntry != null) {
                try (InputStream inputStream = zipFile.getInputStream(documentXmlEntry);
                     OutputStream outputStream = new FileOutputStream(new File(documentXmlPath, "document.xml"))) {
                    byte[] buffer = new byte[1024];
                    int length;
                    while ((length = inputStream.read(buffer)) > 0) {
                        outputStream.write(buffer, 0, length);
                    }
                }
            }

            if (documentRelsEntry != null) {
                try (InputStream inputStream = zipFile.getInputStream(documentRelsEntry);
                     OutputStream outputStream = new FileOutputStream(new File(documentRelsPath, "document.xml.rels"))) {
                    byte[] buffer = new byte[1024];
                    int length;
                    while ((length = inputStream.read(buffer)) > 0) {
                        outputStream.write(buffer, 0, length);
                    }
                }
            }

            for (ZipEntry entry : zipFile.stream().filter(e -> e.getName().startsWith("word/media/")).toList()) {
                String imageName = Paths.get(entry.getName()).getFileName().toString();
                Path imagePath = Paths.get(imageFolderPath, "media", imageName);
                Files.createDirectories(imagePath.getParent());
                try (InputStream inputStream = zipFile.getInputStream(entry);
                     OutputStream outputStream = new FileOutputStream(imagePath.toFile())) {
                    byte[] buffer = new byte[1024];
                    int length;
                    while ((length = inputStream.read(buffer)) > 0) {
                        outputStream.write(buffer, 0, length);
                    }
                }
            }
        }
    }
}


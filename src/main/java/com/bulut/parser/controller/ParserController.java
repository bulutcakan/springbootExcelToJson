package com.bulut.parser.controller;

import com.bulut.parser.model.ExcelSheet;
import com.bulut.parser.service.ExcelService;
import com.bulut.parser.service.OfficeValidatorService;
import io.github.millij.poi.SpreadsheetReadException;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.InputStreamResource;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;

/**
 * Created by Bulut Cakan (179997) on
 * Hour :19:42
 * Day: Saturday
 * Month:December
 * Year:2020
 */
@RestController
@RequestMapping("/api/v1/excel")
public class ParserController {

    @Autowired
    ExcelService excelService;


    @Autowired
    OfficeValidatorService officeValidatorService;


    @PostMapping("/convert")
    public ResponseEntity excelToPojo(@RequestParam("file") MultipartFile excelDataFile) throws Exception {
        return new ResponseEntity<>(excelService.parse(excelDataFile), HttpStatus.OK);
    }

    @PostMapping("/export")
    public ResponseEntity<Resource> pojoToExport(@RequestBody ExcelSheet sheet) throws Exception {

        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);

        InputStream is = new ByteArrayInputStream(excelService.toExcel(sheet.getContent(),"Testt"));; // get your input stream here
        Resource resource = new InputStreamResource(is);
        return new ResponseEntity<>(resource, headers, HttpStatus.OK);

    }

}


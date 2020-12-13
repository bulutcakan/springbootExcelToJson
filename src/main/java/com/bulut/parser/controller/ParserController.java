package com.bulut.parser.controller;

import com.bulut.parser.service.ExcelService;
import com.bulut.parser.service.OfficeValidatorService;
import io.github.millij.poi.SpreadsheetReadException;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;

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
    public ResponseEntity excelToPojo(@RequestParam("file") MultipartFile excelDataFile) throws IOException, SpreadsheetReadException {
        if (!officeValidatorService.isValid(excelDataFile))
            return new ResponseEntity("Not valid Excel File ", HttpStatus.BAD_REQUEST);
        return new ResponseEntity<>(excelService.parse(excelDataFile), HttpStatus.OK);
    }

}


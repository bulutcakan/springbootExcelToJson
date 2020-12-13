package com.bulut.parser.service.impl;

import com.bulut.parser.service.OfficeValidatorService;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;

@Service
public class OfficeValidatorServiceImpl implements OfficeValidatorService {


    @Override
    public boolean isValid(MultipartFile uploadedImage) {
        if (!isOfficeContentType(uploadedImage.getContentType()))
            return false;

        return isValidOfficeFile(uploadedImage);
    }

    @Override
    public boolean isOfficeContentType(String contentType) {
        return isExcelFile(contentType);
    }


    private boolean isExcelFile(String contentType) {
        return contentType.equalsIgnoreCase("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                || contentType.equalsIgnoreCase("application/vnd.ms-excel");
    }

    boolean isValidOfficeFile(MultipartFile multipartFile) {
        boolean result = true;
        try {

            if (isExcelFile(multipartFile.getContentType())) {
                Workbook workbook = new XSSFWorkbook(multipartFile.getInputStream());
            }

        } catch (IOException e) {
            result = false;
        }
        return result;
    }
}

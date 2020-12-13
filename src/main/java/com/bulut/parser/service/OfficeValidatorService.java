package com.bulut.parser.service;

import org.springframework.web.multipart.MultipartFile;

public interface OfficeValidatorService {

    boolean isValid(MultipartFile uploadedImage);

    boolean isOfficeContentType(String contentType);
}

package com.bulut.parser.service;

import com.bulut.parser.model.ExcelResult;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;
import java.util.List;
import java.util.Map;

/**
 * Created by Bulut Cakan (179997) on
 * Hour :15:36
 * Day: Sunday
 * Month:December
 * Year:2020
 */
public interface ExcelService {

    ExcelResult parse(MultipartFile file) throws IOException;

    ExcelResult parse(InputStream stream) throws IOException;

    byte[] toExcel(List<Map<String, Object>> values, String sheetName) throws Exception;


}

package com.bulut.parser.service.impl;

import com.bulut.parser.model.ExcelResult;
import com.bulut.parser.model.ExcelSheet;
import com.bulut.parser.service.ExcelService;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;

/**
 * Created by Bulut Cakan (179997) on
 * Hour :15:29
 * Day: Sunday
 * Month:December
 * Year:2020
 */
@Service
public class ExcelServiceImpl implements ExcelService {

    @Override
    public ExcelResult parse(MultipartFile file) throws IOException {
        return parse(file.getInputStream());
    }

    @Override
    public ExcelResult parse(InputStream stream) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook(stream);
        ExcelResult excelResult = new ExcelResult();
        excelResult.setSheetSize(workbook.getNumberOfSheets());
        int totalRow = 0;

        for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
            ExcelSheet excelSheet = new ExcelSheet();
            Sheet sheet = workbook.getSheetAt(sheetIndex);
            excelSheet.setName(sheet.getSheetName());
            int rowCount = sheet.getLastRowNum();
            excelSheet.setRowCount(rowCount);
            List<String> headerList = new ArrayList<>();

            //For Header Values
            Row firstRow = sheet.getRow(0);
            firstRow.forEach(
                    cell -> {
                        headerList.add(cell.getRichStringCellValue().getString().trim());
                    }
            );
            //

            for (int i = 1; i <= rowCount; i++) {
                Row row = sheet.getRow(i);
                Map<String, Object> rowMap = new HashMap<>();
                AtomicInteger headerIndex = new AtomicInteger();
                row.forEach(cell -> {
                    rowMap.put(headerList.get(headerIndex.get()), getCellValue(cell));
                    headerIndex.getAndIncrement();
                });
                totalRow++;
                excelSheet.getContent().add(rowMap);
            }
            excelResult.getSheets().add(excelSheet);
        }
        excelResult.setTotalRow(totalRow + excelResult.getSheetSize());
        workbook.close();
        return excelResult;
    }


    private Object getCellValue(Cell cell) {
        if (cell.getCellType().equals(CellType.BOOLEAN))
            return cell.getBooleanCellValue();
        if (cell.getCellType().equals(CellType.STRING))
            return cell.getRichStringCellValue().getString();
        if (cell.getCellType().equals(CellType.NUMERIC)) {
            if (DateUtil.isCellDateFormatted(cell)) {
                return cell.getDateCellValue();
            } else {
                return cell.getNumericCellValue();
            }
        }
        if (cell.getCellType().equals(CellType.FORMULA))
            return cell.getCellFormula();

        return StringUtils.EMPTY;
    }
}
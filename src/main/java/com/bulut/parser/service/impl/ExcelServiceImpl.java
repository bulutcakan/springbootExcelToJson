package com.bulut.parser.service.impl;

import com.bulut.parser.model.ExcelResult;
import com.bulut.parser.model.ExcelSheet;
import com.bulut.parser.service.ExcelService;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.util.CollectionUtils;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
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


    @Override
    public byte[] toExcel(List<Map<String, Object>> values, String sheetName) throws Exception {
        if (CollectionUtils.isEmpty(values))
            throw new Exception("Data map can not be null or empty");

        ByteArrayOutputStream out = new ByteArrayOutputStream();

        Workbook workbook = new XSSFWorkbook();
        CreationHelper createHelper = workbook.getCreationHelper();

        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 18);
        headerFont.setColor(IndexedColors.BLACK.getIndex());

        Sheet sheet = workbook.createSheet(sheetName);

        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFont(headerFont);
        headerFont.setColor(IndexedColors.WHITE.index);
        headerCellStyle.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.index);
        headerCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Row headerRow = sheet.createRow(0);
        Set<String> headerSet = values.get(0).keySet();
        headerCellStyle.setFont(headerFont);

        AtomicInteger i = new AtomicInteger();
        headerSet.forEach(s -> {
            Cell cell = headerRow.createCell(i.get());
            cell.setCellValue(s);
            cell.setCellStyle(headerCellStyle);
            i.getAndIncrement();
        });

        CellStyle dateCellStyle = workbook.createCellStyle();
        AtomicInteger rowNum = new AtomicInteger();
        rowNum.set(1);
        values.stream().forEach(stringObjectMap -> {
            Row row = sheet.createRow(rowNum.getAndIncrement());
            AtomicInteger cellNum = new AtomicInteger(0);
            headerSet.stream().forEach(s -> {
                Object value = stringObjectMap.get(s);
                setCellValue(row, cellNum, value);
            });
            cellNum.set(0);
        });

        for (int x = 0; x < headerSet.size(); x++) {
            sheet.autoSizeColumn(x);
        }


        workbook.write(out);
        workbook.close();
        return out.toByteArray();
    }

    private void setCellValue(Row row, AtomicInteger cellNum, Object value) {
        if (value instanceof Integer)
            row.createCell(cellNum.getAndIncrement())
                    .setCellValue((Integer) value);
        if (value instanceof Double)
            row.createCell(cellNum.getAndIncrement())
                    .setCellValue((Double) value);
        if (value instanceof String)
            row.createCell(cellNum.getAndIncrement())
                    .setCellValue((String) value);
        if (value instanceof Float)
            row.createCell(cellNum.getAndIncrement())
                    .setCellValue((Float) value);
        if (value instanceof Date)
            row.createCell(cellNum.getAndIncrement())
                    .setCellValue((Date) value);
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
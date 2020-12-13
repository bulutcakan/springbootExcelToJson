package com.bulut.parser.model;

import lombok.Getter;
import lombok.Setter;

import java.util.ArrayList;
import java.util.List;

/**
 * Created by Bulut Cakan (179997) on
 * Hour :08:31
 * Day: Sunday
 * Month:December
 * Year:2020
 */
@Getter
@Setter
public class ExcelResult {

    private long totalRow = 0;

    private long sheetSize = 0;

    private List<ExcelSheet> sheets = new ArrayList<>();

}

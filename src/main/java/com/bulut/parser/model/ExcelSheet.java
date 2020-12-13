package com.bulut.parser.model;

import lombok.Getter;
import lombok.Setter;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * Created by Bulut Cakan (179997) on
 * Hour :11:00
 * Day: Sunday
 * Month:December
 * Year:2020
 */
@Getter
@Setter
public class ExcelSheet {
    private int rowCount = 0;

    private String name;

    private List<Map<String, Object>> content = new ArrayList<>();

}

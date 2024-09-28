package com.simple.vo;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.ColumnWidth;
import lombok.Data;

@Data
public class Test {

    @ColumnWidth(210)
    @ExcelProperty(index = 0)
    private String name;

    private int age;

    private String address;
}

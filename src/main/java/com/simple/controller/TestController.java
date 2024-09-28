package com.simple.controller;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.annotation.format.NumberFormat;
import com.simple.utils.EasyExcelUtils;
import com.simple.vo.Test;
import jakarta.servlet.http.HttpServletResponse;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;
import java.util.*;

@RestController
@RequestMapping("/easyexcel")
public class TestController {

    public TestController() {
    }

    @PostMapping("/test")
    public void testExcel(@RequestParam("file") MultipartFile file) {
        System.out.println(file);
    }

    @PostMapping("/download")
    public void testExcel(HttpServletResponse response) throws IOException {
        // 这里注意 有同学反应使用swagger 会导致各种问题，请直接用浏览器或者用postman
//        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
//        response.setCharacterEncoding("utf-8");
//        // 这里URLEncoder.encode可以防止中文乱码 当然和easyexcel没有关系
//        String fileName = URLEncoder.encode("测试", "UTF-8").replaceAll("\\+", "%20");
//        response.setHeader("Content-disposition", "attachment;filename*=utf-8''" + fileName + ".xlsx");
        List<Object> dataList = new ArrayList<>();
        Test test = new Test();
        test.setName("1234");
        test.setAge(14);
        test.setAddress("fff");
        dataList.add(test);
        List<Test> dataList3 = new ArrayList<>();
        Test test3 = new Test();
        test3.setName("555");
        test3.setAge(55);
        test3.setAddress("555");
        dataList3.add(test3);
        List<Object> dataList2 = new ArrayList<>();
        Test test2 = new Test();
        test2.setName("555");
        test2.setAge(55);
        test2.setAddress("555");
        dataList2.add(test2);
        List<List<String>> headers = new ArrayList<>();
        List<String> header = new ArrayList<>();
        header.add("厕所 1");
        header.add("厕所");
        header.add("厕所 3");
        headers.add(header);
        Map<String, List<Object>> listMap1 = new HashMap<>();
        listMap1.put("sheet1a", dataList);
        listMap1.put("sheet2a", dataList2);
        List<String> headers2 = new ArrayList<>();
        headers2.add("titi");
        headers2.add("tee");
        List<String> headers3 = new ArrayList<>();
        headers3.add("titi11");
        headers3.add("tee222");
        Map<String, List<String>> listMap2 = new HashMap<>();
        listMap2.put("sheet1a", headers3);
        listMap2.put("sheet2a", headers2);
        EasyExcelUtils.write(response, "中华.xlsx", "tt", Test.class, dataList3);
//        EasyExcelUtils.writeMultipleSheets(response, "中华.xlsx",  listMap1, listMap2);
//        EasyExcel.write(response.getOutputStream(), Test.class).head(headers).sheet("模板").doWrite(dataList);
    }
}

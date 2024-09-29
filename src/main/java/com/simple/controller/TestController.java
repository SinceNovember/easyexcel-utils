package com.simple.controller;


import com.simple.utils.EasyExcelUtils;
import jakarta.servlet.http.HttpServletResponse;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;
import java.io.IOException;
import java.util.*;

@RestController
@RequestMapping("/easyexcel")
public class TestController {

    @PostMapping("/test")
    public void testExcel(@RequestParam("file") MultipartFile file) {
        try {
            EasyExcelUtils.read(file.getInputStream());
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    @PostMapping("/download")
    public void testExcel(HttpServletResponse response) throws IOException {
        Map<String, List<Object>> map = new HashMap<>();
        List<Object> item = new ArrayList<>();
        item.add("1");
        item.add("2");
        item.add("3");
        map.put("sheet1", item);
        EasyExcelUtils.write("test.xlsx", "sheet1", item);
    }
}

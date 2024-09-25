package com.simple.controller;

import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

@RestController
@RequestMapping("/easyexcel")
public class TestController {

    @PostMapping("/test")
    public void testExcel(@RequestParam("file") MultipartFile file) {
        System.out.println(file);
    }
}

package com.rsramos.report.controller;

import com.google.gson.JsonObject;
import com.rsramos.report.service.ReportService;
import org.apache.commons.lang3.StringUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import java.io.*;
import java.util.Date;

@RestController
@RequestMapping("report")
public class ReportController {

    @Autowired
    private ReportService service;

    @ResponseBody
    @PostMapping(path = "xls")
    public ResponseEntity<byte[]> createXLS(@RequestBody JsonObject json) {
        return ResponseEntity.ok(service.createXLS(json));
    }

    @ResponseBody
    @PostMapping(path = "pdf")
    public ResponseEntity<byte[]> createPDF(@RequestBody JsonObject json) {
        return ResponseEntity.ok(service.createPDF(json));
    }

    @ResponseBody
    @PostMapping(path = "test/xls")
    public void testCreateXLS(@RequestBody JsonObject json) {
        File file = new File(StringUtils.join("C:\\temp\\REPORT_XLS\\", new Date().toString().replaceAll(":", "-"), ".xlsx"));
        try {
            file.createNewFile();
            FileOutputStream outputStream = new FileOutputStream(file);
            outputStream.write(service.createXLS(json));
//            outputStream.flush();
            outputStream.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}

package com.rsramos.report.controller;

import com.rsramos.report.domain.Report;
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
    public ResponseEntity<byte[]> createXLS(@RequestBody Report report) {
        return ResponseEntity.ok(service.createXLS(report));
    }

    @ResponseBody
    @PostMapping(path = "pdf")
    public ResponseEntity<byte[]> createPDF(@RequestBody Report report) {
        return ResponseEntity.ok(service.createPDF(report));
    }

    @ResponseBody
    @PostMapping(path = "test/xls")
    public ResponseEntity<byte[]> testCreateXLS(@RequestBody Report report) {
        byte[] ret = null;
        try {
            ret = service.createXLS(report);
            File file = new File(StringUtils.join("C:\\temp\\REPORT_XLS\\", new Date().toString().replaceAll(":", "-"), ".xlsx"));
            file.createNewFile();
            FileOutputStream outputStream = new FileOutputStream(file);
            outputStream.write(ret);
            outputStream.flush();
            outputStream.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return ResponseEntity.ok(ret);
    }

}

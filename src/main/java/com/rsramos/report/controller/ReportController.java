package com.rsramos.report.controller;

import com.rsramos.report.domain.ReportSheet;
import com.rsramos.report.service.ReportService;
import org.apache.commons.lang3.StringUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;
import java.util.List;

@RestController
@RequestMapping("report")
public class ReportController {

    @Autowired
    private ReportService service;

    @ResponseBody
    @PostMapping(path = "xls")
    public ResponseEntity<byte[]> createXLS(@RequestBody List<ReportSheet> sheets) {
        return ResponseEntity.ok(service.createXLS(sheets));
    }

    @ResponseBody
    @PostMapping(path = "xls/base64")
    public ResponseEntity<byte[]> createXLSBase64(@RequestBody List<ReportSheet> sheets) {
        return ResponseEntity.ok(service.createXLSBase64(sheets));
    }

    @ResponseBody
    @PostMapping(path = "pdf")
    public ResponseEntity<byte[]> createPDF(@RequestBody List<ReportSheet> sheets) {
        return ResponseEntity.ok(service.createPDF(sheets));
    }

    @ResponseBody
    @PostMapping(path = "test/xls")
    public ResponseEntity<byte[]> testCreateXLS(@RequestBody List<ReportSheet> sheets) {
        byte[] ret = null;
        try {
            ret = service.createXLS(sheets);
            File file = new File(StringUtils.join("C:\\temp\\REPORT_XLS\\", String.valueOf(new Date().getTime()), ".xlsx"));
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

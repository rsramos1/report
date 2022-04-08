package com.rsramos.report.controller;

import com.google.gson.JsonObject;
import com.rsramos.report.service.ReportService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

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

}

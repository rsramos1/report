package com.rsramos.report.service;

import com.google.gson.JsonObject;
import com.rsramos.report.report.CreateReport;
import org.springframework.stereotype.Service;

@Service
public class ReportService {

    public byte[] createXLS(JsonObject json) {
        return CreateReport.createXLS(json);
    }

    public byte[] createPDF(JsonObject json) {
        return null; // TODO
    }

}

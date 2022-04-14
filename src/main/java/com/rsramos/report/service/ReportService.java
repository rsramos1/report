package com.rsramos.report.service;

import com.rsramos.report.domain.Report;
import com.rsramos.report.report.CreateReport;
import org.springframework.stereotype.Service;

@Service
public class ReportService {

    public byte[] createXLS(Report report) {
        return CreateReport.createXLS(report);
    }

    public byte[] createPDF(Report report) {
        return null; // TODO
    }

}

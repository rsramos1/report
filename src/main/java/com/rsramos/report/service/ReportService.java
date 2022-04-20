package com.rsramos.report.service;

import com.rsramos.report.domain.Report;
import com.rsramos.report.report.CreateXLS;
import org.springframework.stereotype.Service;

@Service
public class ReportService {

    public byte[] createXLS(Report report) {
        return CreateXLS.createXLS(report);
    }

    public byte[] createPDF(Report report) {
        return null; // TODO
    }

}

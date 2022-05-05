package com.rsramos.report.service;

import com.rsramos.report.domain.ReportSheet;
import com.rsramos.report.report.CreateXLS;
import org.springframework.stereotype.Service;

import java.util.List;

@Service
public class ReportService {

    public byte[] createXLS(List<ReportSheet> sheets) {
        return CreateXLS.createXLS(sheets).build();
    }

    public byte[] createXLSBase64(List<ReportSheet> sheets) {
        return CreateXLS.createXLS(sheets).buildBase64();
    }

    public byte[] createPDF(List<ReportSheet> sheets) {
        return null; // TODO
    }

}

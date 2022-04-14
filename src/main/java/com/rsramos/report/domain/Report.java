package com.rsramos.report.domain;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.List;

public class Report implements Serializable {
    private static final long serialVersionUID = 1L;

    private String fileName;
    private String extension;
    private List<ReportSheet> sheets = new ArrayList<>();

    public String getFileName() {
        return fileName;
    }

    public void setFileName(String fileName) {
        this.fileName = fileName;
    }

    public String getExtension() {
        return extension;
    }

    public void setExtension(String extension) {
        this.extension = extension;
    }

    public List<ReportSheet> getSheets() {
        return sheets;
    }

    public void setSheets(List<ReportSheet> sheets) {
        this.sheets = sheets;
    }
}

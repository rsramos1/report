package com.rsramos.report.domain;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public class ReportSheet implements Serializable {
    private static final long serialVersionUID = 1L;

    private String name;
    private CellConfig header = new CellConfig();
    private CellConfig body = new CellConfig();
    private List<ColumnConfig> columns = new ArrayList<>();
    private List<Map<String, String>> data;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public CellConfig getHeader() {
        return header;
    }

    public void setHeader(CellConfig header) {
        this.header = header;
    }

    public CellConfig getBody() {
        return body;
    }

    public void setBody(CellConfig body) {
        this.body = body;
    }

    public List<ColumnConfig> getColumns() {
        return columns;
    }

    public void setColumns(List<ColumnConfig> columns) {
        this.columns = columns;
    }

    public List<Map<String, String>> getData() {
        return data;
    }

    public void setData(List<Map<String, String>> data) {
        this.data = data;
    }

    public ColumnConfig getColumn(String field) {
        for (ColumnConfig column : getColumns()) {
            if (field.equals(column.getField())) {
                return column;
            }
        }
        return null;
    }
}

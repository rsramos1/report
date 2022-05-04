package com.rsramos.report.domain;

import java.io.Serial;
import java.io.Serializable;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public class ReportSheet implements Serializable {
    @Serial
    private static final long serialVersionUID = 1L;

    private String name;
    private CellConfig headerStyle = new CellConfig();
    private CellConfig bodyStyle = new CellConfig();
    private List<ColumnConfig> columns = new ArrayList<>();
    private Map<String, String>[] data;
    private boolean autoFilter;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public CellConfig getHeaderStyle() {
        return headerStyle;
    }

    public void setHeaderStyle(CellConfig headerStyle) {
        this.headerStyle = headerStyle;
    }

    public CellConfig getBodyStyle() {
        return bodyStyle;
    }

    public void setBodyStyle(CellConfig bodyStyle) {
        this.bodyStyle = bodyStyle;
    }

    public List<ColumnConfig> getColumns() {
        return columns;
    }

    public void setColumns(List<ColumnConfig> columns) {
        this.columns = columns;
    }

    public Map<String, String>[] getData() {
        return data;
    }

    public void setData(Map<String, String>[] data) {
        this.data = data;
    }

    public boolean isAutoFilter() {
        return autoFilter;
    }

    public void setAutoFilter(boolean autoFilter) {
        this.autoFilter = autoFilter;
    }

    public ColumnConfig getColumn(String field) {
        for (ColumnConfig column : getColumns()) {
            if (field.equals(column.getField())) {
                return column;
            }
        }
        return new ColumnConfig();
    }
}

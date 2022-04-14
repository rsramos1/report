package com.rsramos.report.domain;

import org.json.JSONArray;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.List;

public class ReportSheet implements Serializable {
    private static final long serialVersionUID = 1L;

    private String name;
    private CellConfig header;
    private CellConfig body;
    private List<ColumnConfig> columns = new ArrayList<>();
    private JSONArray data;

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

    public JSONArray getData() {
        return data;
    }

    public void setData(JSONArray data) {
        this.data = data;
    }

    public ColumnConfig getColumn(String field) {
        for(ColumnConfig column : getColumns()) {
            if(field.equals(column.getField())){
                return column;
            }
        }
        return null;
    }
}

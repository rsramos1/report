package com.rsramos.report.domain;

import java.io.Serializable;

public class ColumnConfig implements Serializable {
    private static final long serialVersionUID = 1L;

    public static final String[] NUMERIC_TYPE = {"INT", "DOUBLE", "INTEGER", "INT",
            "NUMBER", "FLOAT", "SHORT", "BYTE", "DECIMAL", "BIGDECIMAL", "BIG_DECIMAL"};
    public static final String[] BOOLEAN_TYPE = {"BOOL", "BOOLEAN"};
    public static final String[] LOCAL_DATE_TYPE = {"LOCALDATE", "LOCAL_DATE"};
    public static final String[] LOCAL_DATE_TIME_TYPE = {"LOCALDATETIME", "LOCAL_DATE_TIME"};
    public static final String DATE_TYPE = "DATE";
    public static final String CALENDAR_TYPE = "CALENDAR";

    public static final String WIDTH_AUTO = "AUTO";

    private String field;
    private String label;
    private String width;
    private String type;
    private CellConfig style;

    public String getField() {
        return field;
    }

    public void setField(String field) {
        this.field = field;
    }

    public String getLabel() {
        return label;
    }

    public void setLabel(String label) {
        this.label = label;
    }

    public String getWidth() {
        return width;
    }

    public void setWidth(String width) {
        this.width = width;
    }

    public String getType() {
        return type;
    }

    public void setType(String type) {
        this.type = type;
    }

    public CellConfig getStyle() {
        return style;
    }

    public void setStyle(CellConfig style) {
        this.style = style;
    }
}

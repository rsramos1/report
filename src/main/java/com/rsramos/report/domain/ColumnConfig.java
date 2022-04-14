package com.rsramos.report.domain;

import java.io.Serializable;

public class ColumnConfig implements Serializable {
    private static final long serialVersionUID = 1L;

    private String field;
    private String label;
    private String width;

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
}
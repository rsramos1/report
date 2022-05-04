package com.rsramos.report.domain;

import java.io.Serializable;

public class ColumnConfig implements Serializable {
    private static final long serialVersionUID = 1L;

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

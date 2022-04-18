package com.rsramos.report.domain;

import java.io.Serializable;

public class CellConfig implements Serializable {
    private static final long serialVersionUID = 1L;

    private Float height;
    private String background;
    private String foreground;
    private Short fontSize;
    private String fontFamily;
    private Boolean fontBold;
    private Boolean fontItalic;
    private String borderType;
    private String borderColor;

    public Float getHeight() {
        return height;
    }

    public void setHeight(Float height) {
        this.height = height;
    }

    public String getBackground() {
        return background;
    }

    public void setBackground(String background) {
        this.background = background;
    }

    public String getForeground() {
        return foreground;
    }

    public void setForeground(String foreground) {
        this.foreground = foreground;
    }

    public Short getFontSize() {
        return fontSize;
    }

    public void setFontSize(Short fontSize) {
        this.fontSize = fontSize;
    }

    public String getFontFamily() {
        return fontFamily;
    }

    public void setFontFamily(String fontFamily) {
        this.fontFamily = fontFamily;
    }

    public Boolean getFontBold() {
        return fontBold;
    }

    public void setFontBold(Boolean fontBold) {
        this.fontBold = fontBold;
    }

    public Boolean getFontItalic() {
        return fontItalic;
    }

    public void setFontItalic(Boolean fontItalic) {
        this.fontItalic = fontItalic;
    }

    public String getBorderType() {
        return borderType;
    }

    public void setBorderType(String borderType) {
        this.borderType = borderType;
    }

    public String getBorderColor() {
        return borderColor;
    }

    public void setBorderColor(String borderColor) {
        this.borderColor = borderColor;
    }
}

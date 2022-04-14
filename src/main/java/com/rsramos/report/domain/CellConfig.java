package com.rsramos.report.domain;

import java.io.Serializable;

public class CellConfig implements Serializable {
    private static final long serialVersionUID = 1L;

    private String height;
    private String background;
    private String foreground;
    private String fontSize;
    private String fontFamily;
    private String fontBold;
    private String fontItalic;
    private String borderType;
    private String borderColor;

    public String getHeight() {
        return height;
    }

    public void setHeight(String height) {
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

    public String getFontSize() {
        return fontSize;
    }

    public void setFontSize(String fontSize) {
        this.fontSize = fontSize;
    }

    public String getFontFamily() {
        return fontFamily;
    }

    public void setFontFamily(String fontFamily) {
        this.fontFamily = fontFamily;
    }

    public String getFontBold() {
        return fontBold;
    }

    public void setFontBold(String fontBold) {
        this.fontBold = fontBold;
    }

    public String getFontItalic() {
        return fontItalic;
    }

    public void setFontItalic(String fontItalic) {
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

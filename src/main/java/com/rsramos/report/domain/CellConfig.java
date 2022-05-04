package com.rsramos.report.domain;

import java.io.Serial;
import java.io.Serializable;

public class CellConfig implements Serializable {
    @Serial
    private static final long serialVersionUID = 1L;

    private int startRow;
    private int startColumn;
    private Float height;
    private String background;
    private String foreground;
    private Short fontSize;
    private String fontFamily;
    private boolean fontBold;
    private boolean fontItalic;
    private boolean fontUnderline;
    private boolean fontUnderlineSingle;
    private boolean fontUnderlineDouble;
    private boolean fontUnderlineSingleAccounting;
    private boolean fontUnderlineDoubleAccounting;
    private String borderType;
    private String borderColor;
    private String verticalAlign;
    private String horizontalAlign;

    public int getStartRow() {
        return startRow;
    }

    public void setStartRow(int startRow) {
        this.startRow = startRow;
    }

    public int getStartColumn() {
        return startColumn;
    }

    public void setStartColumn(int startColumn) {
        this.startColumn = startColumn;
    }

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

    public boolean isFontBold() {
        return fontBold;
    }

    public void setFontBold(boolean fontBold) {
        this.fontBold = fontBold;
    }

    public boolean isFontItalic() {
        return fontItalic;
    }

    public void setFontItalic(boolean fontItalic) {
        this.fontItalic = fontItalic;
    }

    public boolean isFontUnderline() {
        return fontUnderline;
    }

    public void setFontUnderline(boolean fontUnderline) {
        this.fontUnderline = fontUnderline;
    }

    public boolean isFontUnderlineSingle() {
        return fontUnderlineSingle;
    }

    public void setFontUnderlineSingle(boolean fontUnderlineSingle) {
        this.fontUnderlineSingle = fontUnderlineSingle;
    }

    public boolean isFontUnderlineDouble() {
        return fontUnderlineDouble;
    }

    public void setFontUnderlineDouble(boolean fontUnderlineDouble) {
        this.fontUnderlineDouble = fontUnderlineDouble;
    }

    public boolean isFontUnderlineSingleAccounting() {
        return fontUnderlineSingleAccounting;
    }

    public void setFontUnderlineSingleAccounting(boolean fontUnderlineSingleAccounting) {
        this.fontUnderlineSingleAccounting = fontUnderlineSingleAccounting;
    }

    public boolean isFontUnderlineDoubleAccounting() {
        return fontUnderlineDoubleAccounting;
    }

    public void setFontUnderlineDoubleAccounting(boolean fontUnderlineDoubleAccounting) {
        this.fontUnderlineDoubleAccounting = fontUnderlineDoubleAccounting;
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

    public String getVerticalAlign() {
        return verticalAlign;
    }

    public void setVerticalAlign(String verticalAlign) {
        this.verticalAlign = verticalAlign;
    }

    public String getHorizontalAlign() {
        return horizontalAlign;
    }

    public void setHorizontalAlign(String horizontalAlign) {
        this.horizontalAlign = horizontalAlign;
    }
}

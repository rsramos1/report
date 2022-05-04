package com.rsramos.report.report;

import com.rsramos.report.domain.CellConfig;
import com.rsramos.report.domain.ColumnConfig;
import com.rsramos.report.domain.ReportSheet;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Stream;

public class CreateXLS extends CreateReport {

    protected Workbook workbook;

    protected CreateXLS(List<ReportSheet> sheets) {
        super(sheets);
    }

    protected void createWorkbook() {
        this.workbook = new SXSSFWorkbook();
        this.sheets.forEach(sheet -> {
            String name = StringUtils.defaultIfBlank(sheet.getName(),
                    StringUtils.join("Plan", this.sheets.indexOf(sheet) + 1));
            createSheet(sheet, name);
        });
    }

    public void createSheet(ReportSheet reportSheet, String name) {
        Set<String> keys = reportSheet.getData()[0].keySet();
        Sheet sheet = this.workbook.createSheet(name);
        List<Integer> autoSizeColumns = new ArrayList<>();
        createHeader(reportSheet, keys, autoSizeColumns, sheet);
        createBody(reportSheet, keys, sheet);
        autoSizeColumns.forEach(sheet::autoSizeColumn);
        applyAutoFilter(reportSheet, sheet);
    }

    protected void applyAutoFilter(ReportSheet reportSheet, Sheet sheet) {
        if (reportSheet.isAutoFilter()) {
            sheet.setAutoFilter(CellRangeAddress.valueOf(StringUtils.join(
                    CellReference.convertNumToColString(reportSheet.getHeader().getStartColumn()),
                    reportSheet.getHeader().getStartRow() + 1,
                    ":",
                    CellReference.convertNumToColString(Integer.sum(
                            reportSheet.getBody().getStartColumn(),
                            reportSheet.getData()[0].size() - 1)),
                    Integer.sum(reportSheet.getBody().getStartRow(),
                            reportSheet.getData().length))));
        }
    }

    protected void createHeader(ReportSheet reportSheet, Set<String> keys, List<Integer> autoSizeColumns, Sheet sheet) {
        CellStyle cellStyle = createCellStyle(reportSheet.getHeader());
        Row row = sheet.createRow(reportSheet.getHeader().getStartRow());
        Optional.ofNullable(reportSheet.getHeader().getHeight()).ifPresent(row::setHeightInPoints);

        AtomicInteger index = new AtomicInteger(reportSheet.getHeader().getStartColumn());
        keys.forEach(key -> {
            ColumnConfig columnConfig = reportSheet.getColumn(key);
            if (StringUtils.isNotBlank(columnConfig.getWidth())) {
                if (StringUtils.isNumeric(columnConfig.getWidth())) {
                    int width = (int) Math.round((Double.parseDouble(
                            columnConfig.getWidth()) * 256.0D / 7.001699924468994D));
                    sheet.setColumnWidth(index.get(), width <= 65280 ? width : 65280);
                } else if (StringUtils.equalsIgnoreCase("AUTO", columnConfig.getWidth().trim())) {
                    autoSizeColumns.add(index.get());
                }
            }

            Cell cell = row.createCell(index.getAndIncrement());
            cell.setCellValue(Objects.nonNull(columnConfig.getLabel()) ? columnConfig.getLabel() : key);
            cell.setCellStyle(cellStyle);
        });
        if (!autoSizeColumns.isEmpty()) {
            ((SXSSFSheet) sheet).trackColumnsForAutoSizing(autoSizeColumns);
        }
    }

    protected void createBody(ReportSheet reportSheet, Set<String> keys, Sheet sheet) {
        CellStyle cellStyle = createCellStyle(reportSheet.getBody());

        AtomicInteger rowIndex = new AtomicInteger(reportSheet.getBody().getStartRow() <= reportSheet.getHeader().getStartRow() ?
                reportSheet.getHeader().getStartRow() + 1 : reportSheet.getBody().getStartRow());
        Stream.of(reportSheet.getData()).forEach(element -> {
            AtomicInteger columnIndex = new AtomicInteger(reportSheet.getBody().getStartColumn());
            Row row = sheet.createRow(rowIndex.getAndIncrement());
            Optional.ofNullable(reportSheet.getBody().getHeight()).ifPresent(row::setHeightInPoints);

            keys.forEach(key -> {
                Cell cell = row.createCell(columnIndex.getAndIncrement());
                setCellValue(cell, element.get(key), reportSheet.getColumn(key));
                cell.setCellStyle(cellStyle);
            });
        });
    }

    protected void setCellValue(Cell cell, String value, ColumnConfig columnConfig) {
        String type = StringUtils.defaultIfBlank(columnConfig.getType(), "").trim();
        if (StringUtils.isNotBlank(value)) {
            if (StringUtils.equalsAnyIgnoreCase(type, "INT", "DOUBLE", "INTEGER", "INT", "NUMBER",
                    "FLOAT", "SHORT", "BYTE", "DECIMAL", "BIGDECIMAL", "BIG_DECIMAL")) {
                cell.setCellValue(Double.parseDouble(value));
            } else if (StringUtils.equalsAnyIgnoreCase(type, "BOOL", "BOOLEAN")) {
                cell.setCellValue(Boolean.parseBoolean(value));
            } else if (StringUtils.equalsAnyIgnoreCase(type, "LOCALDATE", "LOCAL_DATE")) {
                if (StringUtils.contains(value, "-")) {
                    cell.setCellValue(LocalDate.parse(value));
                } else {
                    StringBuilder sb = new StringBuilder();
                    sb.append(StringUtils.substring(value, 0, 4));
                    sb.append("-");
                    sb.append(StringUtils.substring(value, 4, 6));
                    sb.append("-");
                    sb.append(StringUtils.substring(value, 6, 8));
                    cell.setCellValue(LocalDate.parse(sb.toString()));
                }
            } else if (StringUtils.equalsAnyIgnoreCase(type, "LOCALDATETIME", "LOCAL_DATE_TIME")) {
                if (StringUtils.contains(value, "-")) {
                    cell.setCellValue(LocalDateTime.parse(value.toUpperCase()));
                } else {
                    StringBuilder sb = new StringBuilder();
                    sb.append(StringUtils.substring(value, 0, 4));
                    sb.append("-");
                    sb.append(StringUtils.substring(value, 4, 6));
                    sb.append("-");
                    sb.append(StringUtils.substring(value, 6, 8));
                    sb.append("T");
                    sb.append(StringUtils.substring(value, 8, 10));
                    sb.append(":");
                    sb.append(StringUtils.substring(value, 10, 12));
                    sb.append(":");
                    sb.append(StringUtils.substring(value, 12, 14));
                    cell.setCellValue(LocalDateTime.parse(sb.toString()));
                }
            } else if (StringUtils.equalsIgnoreCase(type, "DATE")) {
                cell.setCellValue(new Date(Long.parseLong(value)));
            } else if (StringUtils.equalsIgnoreCase(type, "CALENDAR")) {
                cell.setCellValue(Calendar.getInstance());
            } else {
                cell.setCellValue(value);
            }
        } else if (StringUtils.equalsIgnoreCase(type, "CALENDAR")) {
            cell.setCellValue(Calendar.getInstance());
        }
    }

    protected CellStyle createCellStyle(CellConfig styleConfig) {
        XSSFCellStyle cellStyle = (XSSFCellStyle) this.workbook.createCellStyle();
        XSSFFont font = (XSSFFont) this.workbook.createFont();

        Optional.ofNullable(styleConfig.getForeground()).ifPresent(obj -> {
            if (StringUtils.isNotBlank(obj)) {
                font.setColor(createColor(obj));
            }
        });

        Optional.ofNullable(styleConfig.getBackground()).ifPresent(obj -> {
            if (StringUtils.isNotBlank(obj)) {
                cellStyle.setFillForegroundColor(createColor(obj));
                cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            }
        });

        Optional.ofNullable(styleConfig.getFontSize()).ifPresent(obj -> {
            if (Objects.nonNull(obj)) {
                font.setFontHeightInPoints(obj);
            }
        });

        Optional.ofNullable(styleConfig.getFontFamily()).ifPresent(obj -> {
            if (StringUtils.isNotBlank(obj)) {
                font.setFontName(obj);
            }
        });

        if (styleConfig.isFontBold()) {
            font.setBold(true);
        }

        if (styleConfig.isFontItalic()) {
            font.setItalic(true);
        }

        if (styleConfig.isFontUnderline() || styleConfig.isFontUnderlineSingle()) {
            font.setUnderline(FontUnderline.SINGLE);
        }

        if (styleConfig.isFontUnderlineDouble()) {
            font.setUnderline(FontUnderline.DOUBLE);
        }

        if (styleConfig.isFontUnderlineSingleAccounting()) {
            font.setUnderline(FontUnderline.SINGLE_ACCOUNTING);
        }

        if (styleConfig.isFontUnderlineDoubleAccounting()) {
            font.setUnderline(FontUnderline.DOUBLE_ACCOUNTING);
        }

        Optional.ofNullable(styleConfig.getBorderType()).ifPresent(obj -> {
            if (StringUtils.isNotBlank(obj)) {
                BorderStyle borderStyle = StringUtils.isNumeric(obj) ?
                        BorderStyle.valueOf(Short.parseShort(obj)) :
                        BorderStyle.valueOf(obj.toUpperCase());
                cellStyle.setBorderTop(borderStyle);
                cellStyle.setBorderRight(borderStyle);
                cellStyle.setBorderBottom(borderStyle);
                cellStyle.setBorderLeft(borderStyle);
            }
        });

        Optional.ofNullable(styleConfig.getBorderColor()).ifPresent(obj -> {
            if (StringUtils.isNotBlank(obj)) {
                XSSFColor xssfColor = createColor(obj);
                cellStyle.setTopBorderColor(xssfColor);
                cellStyle.setRightBorderColor(xssfColor);
                cellStyle.setBottomBorderColor(xssfColor);
                cellStyle.setLeftBorderColor(xssfColor);
            }
        });

        Optional.ofNullable(styleConfig.getHorizontalAlign()).ifPresent(obj -> {
            if (StringUtils.isNotBlank(obj)) {
                cellStyle.setAlignment(StringUtils.isNumeric(obj) ?
                        HorizontalAlignment.forInt(Integer.parseInt(obj)) :
                        HorizontalAlignment.valueOf(obj.toUpperCase()));
            }
        });

        Optional.ofNullable(styleConfig.getVerticalAlign()).ifPresent(obj -> {
            if (StringUtils.isNotBlank(obj)) {
                cellStyle.setVerticalAlignment(StringUtils.isNumeric(obj) ?
                        VerticalAlignment.forInt(Integer.parseInt(obj)) :
                        VerticalAlignment.valueOf(obj.toUpperCase()));
            }
        });

        cellStyle.setFont(font);
        return cellStyle;
    }

    protected XSSFColor createColor(String color) {
        byte[] hex = getDecodedColorHex(color);
        if (Objects.nonNull(hex)) {
            return new XSSFColor(hex, null);
        }
        XSSFColor xssfColor = new XSSFColor();
        if (StringUtils.isNumeric(color)) {
            xssfColor.setIndexed(IndexedColors.fromInt(Integer.parseInt(color)).getIndex());
        } else {
            xssfColor.setIndexed(IndexedColors.valueOf(color.toUpperCase()).getIndex());
        }
        return xssfColor;
    }

    @Override
    public ByteArrayOutputStream buildByteArrayOutputStream() throws IOException {
        createWorkbook();
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        workbook.write(outputStream);
        return outputStream;
    }

    public static byte[] createXLS(List<ReportSheet> sheets) {
        return new CreateXLS(sheets).build();
    }
}

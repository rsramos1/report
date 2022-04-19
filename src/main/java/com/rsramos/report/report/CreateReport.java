package com.rsramos.report.report;

import com.rsramos.report.domain.CellConfig;
import com.rsramos.report.domain.ColumnConfig;
import com.rsramos.report.domain.Report;
import com.rsramos.report.domain.ReportSheet;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.bouncycastle.util.encoders.Hex;

import java.io.*;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.*;

public class CreateReport implements Serializable {

    @Serial
    private static final long serialVersionUID = 1L;

    protected XSSFWorkbook workbook;

    private final Report report;

    protected CreateReport(Report report) {
        this.report = report;
    }

    protected void createWorkbook() {
        this.workbook = new XSSFWorkbook();
        for (ReportSheet sheet : this.report.getSheets()) {
            String name = StringUtils.defaultIfBlank(sheet.getName(),
                    StringUtils.join("Plan", report.getSheets().indexOf(sheet) + 1));
            createSheet(sheet, name);
        }
    }

    public void createSheet(ReportSheet reportSheet, String name) {
        Set<String> keys = reportSheet.getData().get(0).keySet();
        XSSFSheet sheet = this.workbook.createSheet(name);
        List<Integer> autoSizeColumns = new ArrayList<>();

        createHeader(reportSheet, keys, autoSizeColumns, sheet);
        createBody(reportSheet, keys, sheet);

        resizeColumns(autoSizeColumns, sheet);
    }

    protected void createHeader(ReportSheet reportSheet, Set<String> keys, List<Integer> autoSizeColumns, XSSFSheet sheet) {
        XSSFCellStyle cellStyle = createCellStyle(reportSheet.getHeader());
        XSSFRow row = sheet.createRow(reportSheet.getHeader().getStartRow());
        Optional.ofNullable(reportSheet.getHeader().getHeight()).ifPresent(row::setHeightInPoints);

        int index = reportSheet.getHeader().getStartColumn();
        for (String key : keys) {
            ColumnConfig columnConfig = reportSheet.getColumn(key);
            if (StringUtils.isNotBlank(columnConfig.getWidth())) {
                if (StringUtils.isNumeric(columnConfig.getWidth())) {
                    sheet.setColumnWidth(index, Integer.parseInt(columnConfig.getWidth()));
                } else if (StringUtils.equalsIgnoreCase("AUTO", columnConfig.getWidth().trim())) {
                    autoSizeColumns.add(index);
                }
            }

            XSSFCell cell = row.createCell(index++);
            cell.setCellValue(Objects.nonNull(columnConfig.getLabel()) ? columnConfig.getLabel() : key);
            Optional.ofNullable(cellStyle).ifPresent(cell::setCellStyle);
        }
    }

    protected void createBody(ReportSheet reportSheet, Set<String> keys, XSSFSheet sheet) {
        XSSFCellStyle cellStyle = createCellStyle(reportSheet.getBody());

        int rowIndex = reportSheet.getHeader().getStartRow() >= reportSheet.getBody().getStartRow() ?
                reportSheet.getHeader().getStartRow() + 1 :
                reportSheet.getBody().getStartRow();
        for (Map<String, String> element : reportSheet.getData()) {
            int columnIndex = reportSheet.getBody().getStartRow();
            XSSFRow row = sheet.createRow(rowIndex++);
            Optional.ofNullable(reportSheet.getBody().getHeight()).ifPresent(row::setHeightInPoints);

            for (String key : keys) {
                XSSFCell cell = row.createCell(columnIndex++);
                setCellValue(cell, element.get(key), reportSheet.getColumn(key));
                Optional.ofNullable(cellStyle).ifPresent(cell::setCellStyle);
            }
        }
    }

    protected void setCellValue(XSSFCell cell, String value, ColumnConfig columnConfig) {
        if (Objects.nonNull(value)) {
            String type = columnConfig.getType().trim();
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
        } else if (StringUtils.equalsIgnoreCase(columnConfig.getType().trim(), "CALENDAR")) {
            cell.setCellValue(Calendar.getInstance());
        }
    }

    protected XSSFCellStyle createCellStyle(CellConfig styleConfig) {
        XSSFCellStyle cellStyle = this.workbook.createCellStyle();
        XSSFFont font = this.workbook.createFont();

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

        Optional.ofNullable(styleConfig.getFontBold()).ifPresent(obj -> {
            if (Objects.nonNull(obj)) {
                font.setBold(obj);
            }
        });

        Optional.ofNullable(styleConfig.getFontItalic()).ifPresent(obj -> {
            if (Objects.nonNull(obj)) {
                font.setItalic(obj);
            }
        });

        Optional.ofNullable(styleConfig.getBorderType()).ifPresent(obj -> {
            if (StringUtils.isNotBlank(obj)) {
                if (StringUtils.isNumeric(obj)) {
                    cellStyle.setBorderTop(BorderStyle.valueOf(Short.parseShort(obj)));
                    cellStyle.setBorderRight(BorderStyle.valueOf(Short.parseShort(obj)));
                    cellStyle.setBorderBottom(BorderStyle.valueOf(Short.parseShort(obj)));
                    cellStyle.setBorderLeft(BorderStyle.valueOf(Short.parseShort(obj)));
                } else {
                    cellStyle.setBorderTop(BorderStyle.valueOf(obj.toUpperCase()));
                    cellStyle.setBorderRight(BorderStyle.valueOf(obj.toUpperCase()));
                    cellStyle.setBorderBottom(BorderStyle.valueOf(obj.toUpperCase()));
                    cellStyle.setBorderLeft(BorderStyle.valueOf(obj.toUpperCase()));
                }
            }
        });

        Optional.ofNullable(styleConfig.getBorderColor()).ifPresent(obj -> {
            if (StringUtils.isNotBlank(obj)) {
                cellStyle.setTopBorderColor(createColor(obj));
                cellStyle.setRightBorderColor(createColor(obj));
                cellStyle.setBottomBorderColor(createColor(obj));
                cellStyle.setLeftBorderColor(createColor(obj));
            }
        });

        Optional.ofNullable(styleConfig.getHorizontalAlign()).ifPresent(obj -> {
            if (StringUtils.isNotBlank(obj)) {
                if (StringUtils.isNumeric(obj)) {
                    cellStyle.setAlignment(HorizontalAlignment.forInt(Integer.parseInt(obj)));
                } else {
                    cellStyle.setAlignment(HorizontalAlignment.valueOf(obj.toUpperCase()));
                }
            }
        });

        Optional.ofNullable(styleConfig.getVerticalAlign()).ifPresent(obj -> {
            if (StringUtils.isNotBlank(obj)) {
                if (StringUtils.isNumeric(obj)) {
                    cellStyle.setVerticalAlignment(VerticalAlignment.forInt(Integer.parseInt(obj)));
                } else {
                    cellStyle.setVerticalAlignment(VerticalAlignment.valueOf(obj.toUpperCase()));
                }
            }
        });

        cellStyle.setFont(font);
        return cellStyle;
    }

    protected XSSFColor createColor(String color) {
        XSSFColor xssfColor = getDecodedColor(color);
        if (Objects.isNull(xssfColor)) {
            xssfColor = new XSSFColor();
            if (StringUtils.isNumeric(color)) {
                xssfColor.setIndexed(IndexedColors.fromInt(Integer.parseInt(color)).getIndex());
            } else {
                xssfColor.setIndexed(IndexedColors.valueOf(color.toUpperCase()).getIndex());
            }
        }
        return xssfColor;
    }

    protected XSSFColor getDecodedColor(String color) {
        if (color.contains("#")) {
            color = color.replace("#", "");
            if (color.length() == 3) {
                color += color;
            }
            return new XSSFColor(Hex.decode(color), null);
        } else if (color.contains(",")) {
            StringBuilder hex = new StringBuilder();
            Arrays.stream(color.toUpperCase()
                    .replace("(", "")
                    .replace(")", "")
                    .replace("R", "")
                    .replace("G", "")
                    .replace("B", "")
                    .split(",")).forEach(rgb -> {
                String aux = Integer.toHexString(Integer.parseInt(rgb));
                if (aux.length() == 1) {
                    hex.append("0");
                }
                hex.append(aux);
            });

            return new XSSFColor(Hex.decode(hex.toString()), null);
        }
        return null;
    }

    protected void resizeColumns(List<Integer> columns, XSSFSheet sheet) {
        columns.forEach(sheet::autoSizeColumn);
    }

    protected File build() {
        createWorkbook();
        File file = null;
        FileOutputStream outputStream = null;
        try {
            String fileName = StringUtils.defaultIfBlank(this.report.getFileName(), "report");
            String extension = StringUtils.defaultIfBlank(this.report.getExtension(), "xls");
            file = File.createTempFile(fileName, StringUtils.join(".", extension.replace(".", "")));
            outputStream = new FileOutputStream(file);
            workbook.write(outputStream);
        } catch (IOException e) {
            e.printStackTrace();
        }
        if (Objects.nonNull(outputStream)) {
            try {
                outputStream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return file;
    }

    public static byte[] createXLS(Report report) {
        byte[] bytes = {};
        try {
            CreateReport createReport = new CreateReport(report);
            File file = createReport.build();
            FileInputStream inputStream = new FileInputStream(file);
            bytes = new byte[(int) file.length()];
            inputStream.read(bytes);
            inputStream.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return bytes;
    }
}

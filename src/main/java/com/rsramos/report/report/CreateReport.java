package com.rsramos.report.report;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.rsramos.report.domain.CellConfig;
import com.rsramos.report.domain.ColumnConfig;
import com.rsramos.report.domain.Report;
import com.rsramos.report.domain.ReportSheet;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFColor;

import java.io.*;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.*;

public class CreateReport implements Serializable {

    @Serial
    private static final long serialVersionUID = 1L;

    protected Workbook workbook;
    private final String fileName;
    private final String extension;
    private final int startRowIndex;
    private final int startColumnIndex;

    private final Report report;

    protected CreateReport(Report report, int startRowIndex, int startColumnIndex) {
        this.fileName = StringUtils.defaultIfBlank(report.getFileName(), "report");
        this.extension = StringUtils.defaultIfBlank(report.getExtension(), "xls");
        this.report = report;
        this.startRowIndex = startRowIndex;
        this.startColumnIndex = startColumnIndex;
    }

    protected CreateReport(Report json) {
        this(json, 0, 0);
    }

    protected Set<String> getKeys(Map<String, String> data) {
        Set<String> keys = new HashSet<>();
        try {
            new ObjectMapper().readTree(data.toString()).fieldNames().forEachRemaining(keys::add);
        } catch (JsonProcessingException e) {
            throw new RuntimeException(e);
        }
        return keys;
    }

    protected void createWorkbook() {
        this.workbook = new SXSSFWorkbook();
        int indice = 1;
        for (ReportSheet sheet : this.report.getSheets()) {
            String name = StringUtils.defaultIfBlank(sheet.getName(), StringUtils.join("Plan", indice++));
            createSheet(sheet, name);
        }
    }

    public void createSheet(ReportSheet reportSheet, String name) {
        Set<String> keys = getKeys(reportSheet.getData().get(0));
        Sheet sheet = this.workbook.createSheet(name);
        List<Integer> autoSizeColumns = new ArrayList<>();

        createHeader(reportSheet, keys, autoSizeColumns, sheet);
        createBody(reportSheet, keys, sheet);

        resizeColumns(autoSizeColumns, sheet);
    }

    protected void createHeader(ReportSheet reportSheet, Set<String> keys, List<Integer> autoSizeColumns, Sheet sheet) {
        Float height = null;
        if (Objects.nonNull(reportSheet.getHeader()) && StringUtils.isNotBlank(reportSheet.getHeader().getHeight())) {
            height = Float.parseFloat(reportSheet.getHeader().getHeight());
        }

        CellStyle cellStyle = createCellStyle(reportSheet.getHeader());
        Row row = sheet.createRow(startRowIndex);

        if (Objects.nonNull(height)) {
            row.setHeightInPoints(height);
        }

        int index = 0;
        for (String key : keys) {
            Cell cell = row.createCell(index);

            ColumnConfig columnConfig = reportSheet.getColumn(key);
            String columnValue = Objects.nonNull(columnConfig) && Objects.nonNull(columnConfig.getLabel()) ?
                    columnConfig.getLabel() : key;

            cell.setCellValue(columnValue);

            if (Objects.nonNull(cellStyle)) {
                cell.setCellStyle(cellStyle);
            }

            final int columnIndex = index++;
            Optional.ofNullable(columnConfig).ifPresent(column ->
                    Optional.ofNullable(column.getWidth()).ifPresentOrElse(width -> {
                        if (StringUtils.isNotBlank(width)) {
                            sheet.setColumnWidth(cell.getColumnIndex(), Integer.parseInt(width) + 4000);
                        } else {
                            autoSizeColumns.add(columnIndex);
                        }
                    }, () -> autoSizeColumns.add(columnIndex)));
        }
    }

    protected void createBody(ReportSheet reportSheet, Set<String> keys, Sheet sheet) {
        int rowIndex = this.startRowIndex + 1;
        CellStyle cellStyle = createCellStyle(reportSheet.getBody());

        for (Map<String, String> element : reportSheet.getData()) {
            int columnIndex = this.startColumnIndex;
            Row row = sheet.createRow(rowIndex++);

            for (String key : keys) {
                Cell cell = row.createCell(columnIndex++);

                setCellValue(cell, element.get(key), reportSheet.getColumn(key));

                if (Objects.nonNull(cellStyle)) {
                    cell.setCellStyle(cellStyle);
                }
            }

            if (Objects.nonNull(reportSheet.getBody()) && StringUtils.isNotBlank(reportSheet.getBody().getHeight())) {
                row.setHeightInPoints(Float.parseFloat(reportSheet.getBody().getHeight()));
            }
        }
    }

    protected void setCellValue(Cell cell, String value, ColumnConfig columnConfig) {
        boolean config = Objects.nonNull(columnConfig) && StringUtils.isNotBlank(columnConfig.getType());
        if (Objects.nonNull(value)) {
            if (config && !StringUtils.equalsIgnoreCase(columnConfig.getType().trim(), "STRING")) {
                String type = columnConfig.getType().trim();
                if (!StringUtils.equalsAnyIgnoreCase(type, "INT", "DOUBLE", "INTEGER", "INT", "NUMBER",
                        "FLOAT", "SHORT", "BYTE", "DECIMAL", "BIGDECIMAL", "BIG_DECIMAL")) {
                    cell.setCellValue(Double.parseDouble(value));
                } else if (!StringUtils.equalsAnyIgnoreCase(type, "BOOL", "BOOLEAN")) {
                    cell.setCellValue(Boolean.parseBoolean(value));
                } else if (!StringUtils.equalsAnyIgnoreCase(type, "LOCALDATE", "LOCAL_DATE")) {
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
                } else if (!StringUtils.equalsAnyIgnoreCase(type, "LOCALDATETIME", "LOCAL_DATE_TIME")) {
                    if (StringUtils.contains(value, "-")) {
                        cell.setCellValue(LocalDateTime.parse(value));
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
                } else if (!StringUtils.equalsIgnoreCase(type, "DATE")) {
                    cell.setCellValue(new Date(Long.parseLong(value)));
                } else if (!StringUtils.equalsIgnoreCase(type, "CALENDAR")) {
                    cell.setCellValue(Calendar.getInstance());
                }
            } else {
                cell.setCellValue(value);
            }
        } else if (config && StringUtils.equalsIgnoreCase(columnConfig.getType().trim(), "CALENDAR")) {
            cell.setCellValue(Calendar.getInstance());
        }
    }

    protected CellStyle createCellStyle(CellConfig styleConfig) {
        if (Objects.isNull(styleConfig)) {
            return null;
        }
        CellStyle cellStyle = this.workbook.createCellStyle();
        Font font = this.workbook.createFont();

        Optional.ofNullable(styleConfig.getBackground()).ifPresent(obj -> {
            if (StringUtils.isNotBlank(obj)) {
                cellStyle.setFillBackgroundColor(createColor(obj).getIndex());
            }
        });

        Optional.ofNullable(styleConfig.getForeground()).ifPresent(obj -> {
            if (StringUtils.isNotBlank(obj)) {
                font.setColor(createColor(obj).getIndex());
            }
        });

        Optional.ofNullable(styleConfig.getFontSize()).ifPresent(obj -> {
            if (StringUtils.isNotBlank(obj)) {
                font.setFontHeightInPoints(Short.parseShort(obj));
            }
        });

        Optional.ofNullable(styleConfig.getFontFamily()).ifPresent(obj -> {
            if (StringUtils.isNotBlank(obj)) {
                font.setFontName(obj);
            }
        });

        Optional.ofNullable(styleConfig.getFontBold()).ifPresent(obj -> {
            if (StringUtils.isNotBlank(obj)) {
                font.setBold(Boolean.parseBoolean(obj));
            }
        });

        Optional.ofNullable(styleConfig.getFontItalic()).ifPresent(obj -> {
            if (StringUtils.isNotBlank(obj)) {
                font.setItalic(Boolean.parseBoolean(obj));
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
                if (StringUtils.isNumeric(obj)) {
                    cellStyle.setTopBorderColor(IndexedColors.fromInt(Integer.parseInt(obj)).getIndex());
                    cellStyle.setRightBorderColor(IndexedColors.fromInt(Integer.parseInt(obj)).getIndex());
                    cellStyle.setBottomBorderColor(IndexedColors.fromInt(Integer.parseInt(obj)).getIndex());
                    cellStyle.setLeftBorderColor(IndexedColors.fromInt(Integer.parseInt(obj)).getIndex());
                } else {
                    cellStyle.setTopBorderColor(IndexedColors.valueOf(obj.toUpperCase()).getIndex());
                    cellStyle.setRightBorderColor(IndexedColors.valueOf(obj.toUpperCase()).getIndex());
                    cellStyle.setBottomBorderColor(IndexedColors.valueOf(obj.toUpperCase()).getIndex());
                    cellStyle.setLeftBorderColor(IndexedColors.valueOf(obj.toUpperCase()).getIndex());
                }
            }
        });

        cellStyle.setFont(font);
        return cellStyle;
    }

    protected XSSFColor createColor(String color) {
        XSSFColor xssfColor = new XSSFColor();
        if (color.contains("#") || color.length() == 3 || color.length() == 6) {
            color = color.replace("#", "");
            if (color.length() == 3) {
                color += color;
            }
            xssfColor.setARGBHex(color);
        } else {
            String[] rgb = color
                    .replace("(", "").replace(")", "")
                    .replace("R", "").replace("r", "")
                    .replace("G", "").replace("g", "")
                    .replace("B", "").replace("b", "")
                    .split(",");

            xssfColor.setRGB(new byte[]{
                    (byte) (Integer.parseInt(rgb[0]) - 127),
                    (byte) (Integer.parseInt(rgb[1]) - 127),
                    (byte) (Integer.parseInt(rgb[2]) - 127)
            });
        }
        return xssfColor;
    }

    protected void resizeColumns(List<Integer> columns, Sheet sheet) {
        columns.forEach(sheet::autoSizeColumn);
    }

    protected File build() {
        createWorkbook();
        File file = null;
        FileOutputStream outputStream = null;
        try {
            file = File.createTempFile(this.fileName, this.extension);
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

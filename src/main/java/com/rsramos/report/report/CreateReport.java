package com.rsramos.report.report;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.google.gson.JsonArray;
import com.google.gson.JsonElement;
import com.google.gson.JsonObject;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFColor;

import java.io.*;
import java.util.*;

public class CreateReport implements Serializable {

    @Serial
    private static final long serialVersionUID = 1L;

    private static final String EXTENSION_ATTRIBUTE = "extension";
    private static final String FILE_NAME_ATTRIBUTE = "fileName";
    private static final String SHEETS_ATTRIBUTE = "sheets";

    private static final String NAME_ATTRIBUTE = "name";
    private static final String CONFIG_ATTRIBUTE = "config";
    private static final String HEADER_ATTRIBUTE = "header";
    private static final String BODY_ATTRIBUTE = "body";
    private static final String COLUMN_ATTRIBUTE = "column";
    private static final String LABEL_ATTRIBUTE = "label";
    private static final String WIDTH_ATTRIBUTE = "width";

    private static final String HEIGHT_ATTRIBUTE = "height";
    private static final String BACKGROUND_ATTRIBUTE = "background";
    private static final String FOREGROUND_ATTRIBUTE = "foreground";
    private static final String FONT_SIZE_ATTRIBUTE = "fontSize";
    private static final String FONT_NAME_ATTRIBUTE = "fontName";
    private static final String FONT_BOLD_ATTRIBUTE = "fontBold";
    private static final String FONT_ITALIC_ATTRIBUTE = "fontItalic";

    private static final String BORDER_ATTRIBUTE = "border";
    private static final String TYPE_ATTRIBUTE = "type";
    private static final String COLOR_ATTRIBUTE = "color";

    private static final String DATA_ATTRIBUTE = "data";

    protected Workbook workbook;
    private String fileName;
    private String extension;
    private final int startRowIndex;
    private final int startColumnIndex;

    private final JsonArray sheets;

    protected CreateReport(JsonObject json, int startRowIndex, int startColumnIndex) {
        Optional.ofNullable(json.get(EXTENSION_ATTRIBUTE)).ifPresentOrElse(
                extension -> this.extension = extension.getAsString(),
                () -> this.extension = "xls");
        Optional.ofNullable(json.get(FILE_NAME_ATTRIBUTE)).ifPresentOrElse(
                fileName -> this.fileName = fileName.getAsString(),
                () -> this.extension = "report");
        this.sheets = json.getAsJsonArray(SHEETS_ATTRIBUTE);
        this.startRowIndex = startRowIndex;
        this.startColumnIndex = startColumnIndex;
    }

    protected CreateReport(JsonObject json) {
        this(json, 0, 0);
    }

    protected Map<String, String> loadColumnNames(JsonObject config, JsonObject data) {
        try {
            Map<String, String> headerFileds = new HashMap<>();
            List<String> keys = new ArrayList<>();
            new ObjectMapper().readTree(data.toString()).fieldNames().forEachRemaining(keys::add);
            JsonElement column = config.get(COLUMN_ATTRIBUTE);
            keys.forEach(key -> {
                String value = key;
                if (Objects.nonNull(column)) {
                    JsonElement columnKey = ((JsonObject) column).get(key);
                    if (Objects.nonNull(columnKey)) {
                        JsonElement columnLabel = ((JsonObject) columnKey).get(LABEL_ATTRIBUTE);
                        if (Objects.nonNull(columnLabel)) {
                            value = columnLabel.getAsString();
                        }
                    }
                }
                headerFileds.put(key, value);
            });
            return headerFileds;
        } catch (JsonProcessingException e) {
            throw new RuntimeException(e);
        }
    }

    protected void createWorkbook() {
        this.workbook = new SXSSFWorkbook();

        int indice = 1;
        for (JsonElement sheetJson : this.sheets) {
            JsonElement nameElement = ((JsonObject) sheetJson).get(NAME_ATTRIBUTE);
            String name = Objects.isNull(nameElement) ?
                    StringUtils.join("Plan", indice++) :
                    nameElement.getAsString();
            createSheet(sheetJson.getAsJsonObject(), name);
        }
    }

    public void createSheet(JsonObject sheetJson, String name) {
        JsonArray data = sheetJson.getAsJsonArray(DATA_ATTRIBUTE);
        JsonObject config = sheetJson.getAsJsonObject(CONFIG_ATTRIBUTE);
        Map<String, String> columnNames = loadColumnNames(config, data.get(0).getAsJsonObject());
        Sheet sheet = this.workbook.createSheet(name);

        createHeader(config, columnNames, sheet);
        createBody(data, config.get(BODY_ATTRIBUTE).getAsJsonObject(), columnNames.keySet(), sheet);
    }

    protected void createHeader(JsonObject config, Map<String, String> columnNames, Sheet sheet) {
        JsonObject column = config.getAsJsonObject(COLUMN_ATTRIBUTE);
        JsonElement headerStyle = config.getAsJsonObject(HEADER_ATTRIBUTE);
        Float height = getHeight(headerStyle.getAsJsonObject());

        CellStyle cellStyle = createCellStyle(headerStyle.getAsJsonObject());
        Row row = sheet.createRow(startRowIndex);

        if (Objects.nonNull(height)) {
            row.setHeightInPoints(height);
        }

        int index = 0;
        Iterator<String> keys = columnNames.keySet().iterator();
        while (keys.hasNext()) {
            String key = keys.next();
            Cell cell = row.createCell(index++);
            cell.setCellValue(columnNames.get(key));

            if (Objects.nonNull(cellStyle)) {
                cell.setCellStyle(cellStyle);
            }

            Optional.ofNullable(column.get(key)).ifPresent(obj -> {
                Optional.ofNullable(obj.getAsJsonObject().get(WIDTH_ATTRIBUTE)).ifPresent(width -> {
                    String value = width.getAsString();
                    if (StringUtils.isNotBlank(value)) {
                        sheet.setColumnWidth(cell.getColumnIndex(), Integer.parseInt(width.getAsString()) + 4000);
                    }
                });
            });
        }
    }

    protected void createBody(JsonArray data, JsonObject bodyConfig, Set<String> keys, Sheet sheet) {
        int rowIndex = this.startRowIndex + 1;
        Iterator<String> keysIterator = keys.iterator();
        CellStyle cellStyle = createCellStyle(bodyConfig);
        Float height = getHeight(bodyConfig);

        for (JsonElement dataElement : data) {
            int columnIndex = this.startColumnIndex;
            Row row = sheet.createRow(rowIndex++);
            if (Objects.nonNull(height)) {
                row.setHeightInPoints(height);
            }

            JsonObject dataObject = dataElement.getAsJsonObject();
            while (keysIterator.hasNext()) {
                String key = keysIterator.next();
                Cell cell = row.createCell(columnIndex++);

                Optional.ofNullable(dataObject.get(key)).ifPresentOrElse(
                        value -> cell.setCellValue(value.getAsString()),
                        () -> cell.setCellValue(""));

                if (Objects.nonNull(cellStyle)) {
                    cell.setCellStyle(cellStyle);
                }
            }
        }
    }

    protected float getHeight(JsonObject configStyleElement) {
        Float height = null;

        JsonElement heightElement = configStyleElement.get(HEIGHT_ATTRIBUTE);
        if (Objects.nonNull(heightElement)) {
            String heightAttribute = heightElement.getAsString();
            if (StringUtils.isNotBlank(heightAttribute) && StringUtils.isNumeric(heightAttribute)) {
                height = Float.parseFloat(heightAttribute);
            }
        }

        return height;
    }

    protected CellStyle createCellStyle(JsonObject styleConfig) {
        if (Objects.isNull(styleConfig)) {
            return null;
        }
        CellStyle cellStyle = this.workbook.createCellStyle();
        Font font = this.workbook.createFont();

        Optional.ofNullable(styleConfig.get(BACKGROUND_ATTRIBUTE)).ifPresent(obj -> {
            String color = obj.getAsString();
            if (StringUtils.isNotBlank(color)) {
                cellStyle.setFillBackgroundColor(createColor(color).getIndex());
            }
        });

        Optional.ofNullable(styleConfig.get(FOREGROUND_ATTRIBUTE)).ifPresent(obj -> {
            String color = obj.getAsString();
            if (StringUtils.isNotBlank(color)) {
                font.setColor(createColor(color).getIndex());
            }
        });

        Optional.ofNullable(styleConfig.get(FONT_SIZE_ATTRIBUTE)).ifPresent(obj -> {
            String size = obj.getAsString();
            if (StringUtils.isNotBlank(size)) {
                font.setFontHeightInPoints(Short.parseShort(size));
            }
        });

        Optional.ofNullable(styleConfig.get(FONT_NAME_ATTRIBUTE)).ifPresent(obj -> {
            String fontName = obj.getAsString();
            if (StringUtils.isNotBlank(fontName)) {
                font.setFontName(fontName);
            }
        });

        Optional.ofNullable(styleConfig.get(FONT_BOLD_ATTRIBUTE)).ifPresent(obj -> {
            String fontBold = obj.getAsString();
            if (StringUtils.isNotBlank(fontBold)) {
                font.setBold(Boolean.parseBoolean(fontBold));
            }
        });

        Optional.ofNullable(styleConfig.get(FONT_ITALIC_ATTRIBUTE)).ifPresent(obj -> {
            String fontItalic = obj.getAsString();
            if (StringUtils.isNotBlank(fontItalic)) {
                font.setItalic(Boolean.parseBoolean(fontItalic));
            }
        });

        Optional.ofNullable(styleConfig.get(BORDER_ATTRIBUTE)).ifPresent(obj -> {
            JsonObject border = obj.getAsJsonObject();
            Optional.ofNullable(border.get(TYPE_ATTRIBUTE)).ifPresent(typeObj -> {
                String type = typeObj.getAsString();
                if (StringUtils.isNotBlank(type)) {
                    if (StringUtils.isNumeric(type)) {
                        cellStyle.setBorderTop(BorderStyle.valueOf(Short.parseShort(type)));
                        cellStyle.setBorderRight(BorderStyle.valueOf(Short.parseShort(type)));
                        cellStyle.setBorderBottom(BorderStyle.valueOf(Short.parseShort(type)));
                        cellStyle.setBorderLeft(BorderStyle.valueOf(Short.parseShort(type)));
                    } else {
                        cellStyle.setBorderTop(BorderStyle.valueOf(type.toUpperCase()));
                        cellStyle.setBorderRight(BorderStyle.valueOf(type.toUpperCase()));
                        cellStyle.setBorderBottom(BorderStyle.valueOf(type.toUpperCase()));
                        cellStyle.setBorderLeft(BorderStyle.valueOf(type.toUpperCase()));
                    }
                }
            });
            Optional.ofNullable(border.get(COLOR_ATTRIBUTE)).ifPresent(colorObj -> {
                String color = colorObj.getAsString();
                if (StringUtils.isNotBlank(color)) {
                    if (StringUtils.isNumeric(color)) {
                        cellStyle.setTopBorderColor(IndexedColors.fromInt(Integer.parseInt(color)).getIndex());
                        cellStyle.setRightBorderColor(IndexedColors.fromInt(Integer.parseInt(color)).getIndex());
                        cellStyle.setBottomBorderColor(IndexedColors.fromInt(Integer.parseInt(color)).getIndex());
                        cellStyle.setLeftBorderColor(IndexedColors.fromInt(Integer.parseInt(color)).getIndex());
                    } else {
                        cellStyle.setTopBorderColor(IndexedColors.valueOf(color.toUpperCase()).getIndex());
                        cellStyle.setRightBorderColor(IndexedColors.valueOf(color.toUpperCase()).getIndex());
                        cellStyle.setBottomBorderColor(IndexedColors.valueOf(color.toUpperCase()).getIndex());
                        cellStyle.setLeftBorderColor(IndexedColors.valueOf(color.toUpperCase()).getIndex());
                    }
                }
            });
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

    public static byte[] createXLS(JsonObject json) {
        byte[] bytes = {};
        try {
            CreateReport report = new CreateReport(json);
            File file = report.build();
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

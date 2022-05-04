package com.rsramos.report.report;

import com.rsramos.report.domain.CellConfig;
import com.rsramos.report.domain.ColumnConfig;
import com.rsramos.report.domain.ReportSheet;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.*;
import java.util.List;

public class CreatePDF {//} extends CreateReport {

//    Document document;
//
//    protected CreatePDF(List<ReportSheet> sheets) throws DocumentException {
//        super(sheets);
//        createDocument();
//    }
//
//    protected void createDocument() throws DocumentException {
//        this.document = new Document();
//        document.open();
//        for (ReportSheet sheet : this.sheets) {
//            String name = StringUtils.defaultIfBlank(sheet.getName(),
//                    StringUtils.join("Plan", this.sheets.indexOf(sheet) + 1));
//            createTable(sheet, name);
//        }
//        document.close();
//    }
//
//    public void createTable(ReportSheet reportSheet, String name) throws DocumentException {
//        if (Objects.nonNull(name)) {
//            Font font = new Font(Font.FontFamily.UNDEFINED, 14, Font.BOLD, new BaseColor(0, 0, 0));
//            Paragraph title = new Paragraph(name, font);
//            title.setAlignment(Element.ALIGN_CENTER);
//            title.setSpacingAfter(20);
//            document.add(title);
//        }
//        Set<String> keys = reportSheet.getData().get(0).keySet();
//        PdfPTable table = new PdfPTable(keys.size());
//        List<Integer> autoSizeColumns = new ArrayList<>();
//        createHeader(reportSheet, keys, table);
//        createBody(reportSheet, keys, table);
//    }
//
//    protected void createHeader(ReportSheet reportSheet, Set<String> keys, PdfPTable table) {
//        keys.forEach(key -> {
//            ColumnConfig columnConfig = reportSheet.getColumn(key);
//            Paragraph paragraph = new Paragraph(columnConfig.getLabel());
//            PdfPCell cell = createPDFCell(reportSheet.getHeader(), columnConfig, paragraph);
//            table.addCell(cell);
//        });
//    }
//
//    protected void createBody(ReportSheet reportSheet, Set<String> keys, PdfPTable table) {
//        reportSheet.getData().forEach(data -> {
//            keys.forEach(key -> {
//                ColumnConfig columnConfig = reportSheet.getColumn(key);
//                Paragraph paragraph = new Paragraph(columnConfig.getLabel());
//                PdfPCell cell = createPDFCell(reportSheet.getHeader(), columnConfig, paragraph);
//                table.addCell(cell);
//            });
//        });
//    }
//
//    public PdfPCell createPDFCell(CellConfig cellConfig, ColumnConfig columnConfig, Paragraph paragraph) {
////        Font headerFont = new Font(Font.FontFamily.UNDEFINED, Font.DEFAULTSIZE, Font.BOLD, BaseColor.WHITE);
//        PdfPCell cell = new PdfPCell(paragraph);
//        Font font = new Font();
//
//        Optional.ofNullable(cellConfig.getForeground()).ifPresent(obj -> {
//            if (StringUtils.isNotBlank(obj)) {
//                font.setColor(createColor(obj));
//            }
//        });
//
//        Optional.ofNullable(cellConfig.getBackground()).ifPresent(obj -> {
//            if (StringUtils.isNotBlank(obj)) {
//                cell.setBackgroundColor(createColor(obj));
//            }
//        });
//
//        Optional.ofNullable(cellConfig.getFontSize()).ifPresent(obj -> {
//            if (Objects.nonNull(obj)) {
//                font.setSize(obj);
//            }
//        });
//
//        Optional.ofNullable(cellConfig.getFontFamily()).ifPresent(obj -> {
//            if (StringUtils.isNotBlank(obj)) {
//                font.setFamily(obj);
//            }
//        });
//
////        TODO
////        if (cellConfig.isFontBold()) {
////            font.setBold(true);
////        }
////
////        if (cellConfig.isFontItalic()) {
////            font.setItalic(true);
////        }
////
////        if (cellConfig.isFontUnderline() || styleConfig.isFontUnderlineSingle()) {
////            font.setUnderline(FontUnderline.SINGLE);
////        }
////
////        if (cellConfig.isFontUnderlineDouble()) {
////            font.setUnderline(FontUnderline.DOUBLE);
////        }
////
////        if (cellConfig.isFontUnderlineSingleAccounting()) {
////            font.setUnderline(FontUnderline.SINGLE_ACCOUNTING);
////        }
////
////        if (cellConfig.isFontUnderlineDoubleAccounting()) {
////            font.setUnderline(FontUnderline.DOUBLE_ACCOUNTING);
////        }
//
////        TODO
//        Optional.ofNullable(cellConfig.getBorderType()).ifPresent(obj -> {
//            if (StringUtils.isNotBlank(obj)) {
//                BorderStyle borderStyle = StringUtils.isNumeric(obj) ?
//                        BorderStyle.valueOf(Short.parseShort(obj)) :
//                        BorderStyle.valueOf(obj.toUpperCase());
//                if (borderStyle != null) {
////                    cell.setBorder(PdfPCell.RIGHT);
//                    cell.setBorderWidth(1);
//                }
//            }
//        });
//
//        Optional.ofNullable(cellConfig.getBorderColor()).ifPresentOrElse(obj -> {
//            if (StringUtils.isNotBlank(obj)) {
//                cell.setBorderColor(createColor(obj));
//            }
//        }, () -> cell.setBorderColor(new BaseColor(0, 0, 0)));
//
////        TODO
//        Optional.ofNullable(cellConfig.getHorizontalAlign()).ifPresent(obj -> {
//            if (StringUtils.isNotBlank(obj)) {
//                cell.setHorizontalAlignment(1); // PdfPCell.ALIGN_CENTER
//            }
//        });
//
////        TODO
//        Optional.ofNullable(cellConfig.getVerticalAlign()).ifPresent(obj -> {
//            if (StringUtils.isNotBlank(obj)) {
//                cell.setVerticalAlignment(1); // PdfPCell.ALIGN_CENTER
//            }
//        });
//
//        Optional.ofNullable(cellConfig.getHeight()).ifPresent(obj -> {
//            if (Objects.nonNull(obj)) {
//                cell.setFixedHeight(obj);
//            }
//        });
//
//        return cell;
//    }
//
//    protected BaseColor createColor(String color) {
//        int[] rgb = getColorRGB(color);
//        return new BaseColor(rgb[0], rgb[1], rgb[2]);
//    }
//
//    @Override
//    public ByteArrayOutputStream buildByteArrayOutputStream() throws IOException {
//        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
//        try {
//            PdfWriter.getInstance(document, outputStream);
//        } catch (DocumentException ex) {
//            throw new RuntimeException(ex);
//        }
//        return outputStream;
//    }
//
//    public static byte[] createPDF(List<ReportSheet> sheets) {
//        try {
//            return new CreatePDF(sheets).build();
//        } catch (DocumentException ex) {
//            throw new RuntimeException(ex);
//        }
//    }
}

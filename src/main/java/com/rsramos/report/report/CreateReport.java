package com.rsramos.report.report;

import com.rsramos.report.domain.ReportSheet;
import org.apache.commons.codec.binary.Base64;
import org.apache.commons.lang3.ArrayUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.text.TextStringBuilder;
import org.bouncycastle.util.encoders.Hex;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.Serial;
import java.io.Serializable;
import java.util.Arrays;
import java.util.List;

public abstract class CreateReport implements Serializable {
    @Serial
    private static final long serialVersionUID = 1L;

    public static final String SHEETS_ATTRIBUTE = "sheets";
    public static final String FILE_ATTRIBUTE = "file";

    protected final List<ReportSheet> sheets;

    protected CreateReport(List<ReportSheet> sheets) {
        this.sheets = sheets;
        validate();
    }

    public void validate() {
        if (this.sheets.isEmpty()) {
            throw new IllegalArgumentException("Sheets cannot be empty");
        } else if (this.sheets.stream().anyMatch(sheet -> ArrayUtils.isEmpty(sheet.getData()))) {
            throw new IllegalArgumentException("Data cannot be empty");
        }
    }

    protected String localDateStringFormatted(String value) {
        TextStringBuilder sb = new TextStringBuilder();
        sb.append(StringUtils.substring(value, 0, 4));
        sb.append("-");
        sb.append(StringUtils.substring(value, 4, 6));
        sb.append("-");
        sb.append(StringUtils.substring(value, 6, 8));
        return sb.toString();
    }

    protected String localDateTimeStringFormatted(String value) {
        TextStringBuilder sb = new TextStringBuilder();
        sb.append(localDateStringFormatted(value));
        sb.append("T");
        sb.append(StringUtils.substring(value, 8, 10));
        sb.append(":");
        sb.append(StringUtils.substring(value, 10, 12));
        sb.append(":");
        sb.append(StringUtils.substring(value, 12, 14));
        return sb.toString();
    }

    public abstract ByteArrayOutputStream buildByteArrayOutputStream() throws IOException;

    public byte[] build() {
        byte[] bytes = {};
        try {
            ByteArrayOutputStream outputStream = buildByteArrayOutputStream();
            bytes = outputStream.toByteArray();
            outputStream.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return bytes;
    }

    public byte[] buildBase64() {
        return Base64.encodeBase64(build());
    }

    protected byte[] getDecodedColorHex(String color) {
        if (color.contains("#")) {
            color = color.replace("#", "");
            if (color.length() == 3) {
                color += color;
            }
            return Hex.decode(color);
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
            return Hex.decode(hex.toString());
        }
        return null;
    }
}

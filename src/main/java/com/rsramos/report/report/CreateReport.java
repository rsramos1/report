package com.rsramos.report.report;

import com.rsramos.report.domain.ReportSheet;
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

    protected final List<ReportSheet> sheets;

    protected CreateReport(List<ReportSheet> sheets) {
        if (sheets.isEmpty()) {
            throw new IllegalArgumentException("Sheets cannot be empty");
        } else if (sheets.stream().anyMatch(sheet -> sheet.getData().isEmpty())) {
            throw new IllegalArgumentException("Data cannot be empty");
        }
        this.sheets = sheets;
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

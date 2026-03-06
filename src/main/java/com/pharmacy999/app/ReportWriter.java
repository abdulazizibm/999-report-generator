package com.pharmacy999.app;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.*;

public class ReportWriter {

    public void write(Path output, ReportModel model) throws Exception {
        try (Workbook wb = new XSSFWorkbook()) {
            writeDataSheet(wb, model);

            try (OutputStream out = Files.newOutputStream(output)) {
                wb.write(out);
            }
        }
    }


    private void writeDataSheet(Workbook wb, ReportModel model) {
        Sheet s = wb.createSheet("Data");

        // union headers across all rows (for PoC)
        LinkedHashSet<String> headers = new LinkedHashSet<>();
        for (RowRecord rr : model.normalizedRows()) headers.addAll(rr.values().keySet());

        List<String> headerList = new ArrayList<>(headers);

        int r = 0;
        Row headerRow = s.createRow(r++);
        headerRow.createCell(0).setCellValue("sourceFile");
        headerRow.createCell(1).setCellValue("rowNumber");
        for (int i = 0; i < headerList.size(); i++) {
            headerRow.createCell(i + 2).setCellValue(headerList.get(i));
        }

        for (RowRecord rr : model.normalizedRows()) {
            Row row = s.createRow(r++);
            row.createCell(0).setCellValue(rr.sourceFile());
            row.createCell(1).setCellValue(rr.rowNumber());
            for (int i = 0; i < headerList.size(); i++) {
                String key = headerList.get(i);
                row.createCell(i + 2).setCellValue(rr.values().getOrDefault(key, ""));
            }
        }

        for (int c = 0; c < Math.min(headerList.size() + 2, 30); c++) {
            s.autoSizeColumn(c);
        }
    }
}
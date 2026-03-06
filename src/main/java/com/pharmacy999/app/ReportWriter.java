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
    DataFormat dataFormat = wb.createDataFormat();

    CellStyle groupedNumberStyle = wb.createCellStyle();
    groupedNumberStyle.setDataFormat(dataFormat.getFormat("#,##0.0"));

    CellStyle percentStyle = wb.createCellStyle();
    percentStyle.setDataFormat(dataFormat.getFormat("0.0%"));

    // union headers across all rows (for PoC)
    LinkedHashSet<String> headers = new LinkedHashSet<>();
    for (RowRecord rr : model.normalizedRows()) {
      headers.addAll(rr.values().keySet());
    }

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
        //row.createCell(i + 2).setCellValue(rr.values().getOrDefault(key, ""));
        String value = rr.values().getOrDefault(key, "");
        Cell cell = row.createCell(i + 2);

        if (isNumericColumn(key) && !value.isBlank()) {
          double numericValue = parseRuNumber(value);
          if (!Double.isNaN(numericValue)) {
            cell.setCellValue(numericValue);

            if ("ДоляПрибыли".equals(key)) {
              cell.setCellStyle(percentStyle);
            } else {
              cell.setCellStyle(groupedNumberStyle);
            }

          } else {
            cell.setCellValue(value);
          }
        } else {
          cell.setCellValue(value);
        }
      }
    }

    for (int c = 0; c < Math.min(headerList.size() + 2, 30); c++) {
      s.autoSizeColumn(c);
    }
  }

  private boolean isNumericColumn(String key) {
    return "Кол-во".equals(key)
        || "Сумма".equals(key)
        || "Прибыль".equals(key)
        || "ДоляПрибыли".equals(key);
  }

  private double parseRuNumber(String s) {
    if (s == null) {
      return Double.NaN;
    }

    String t = s.trim()
        .replace("\u00A0", "")
        .replace(" ", "");

    if (t.isEmpty()) {
      return Double.NaN;
    }

    try {
      if (t.contains(",") && t.contains(".")) {
        t = t.replace(".", "")
            .replace(",", ".");
        return Double.parseDouble(t);
      }

      if (t.contains(",")) {
        t = t.replace(",", ".");
        return Double.parseDouble(t);
      }

      return Double.parseDouble(t);

    } catch (NumberFormatException e) {
      return Double.NaN;
    }
  }

    /*private double parseRuNumber(String s) {
        if (s == null) return Double.NaN;

        String t = s.trim()
            .replace("\u00A0", "")
            .replace(" ", "")
            .replace(".", "")
            .replace(",", ".");

        if (t.isEmpty()) return Double.NaN;

        try {
            return Double.parseDouble(t);
        } catch (NumberFormatException e) {
            return Double.NaN;
        }
    }*/
}
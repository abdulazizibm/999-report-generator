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
    percentStyle.setDataFormat(dataFormat.getFormat("0.00%"));

    // union headers across all rows (for PoC)
    LinkedHashSet<String> headers = new LinkedHashSet<>();
    for (RowRecord rr : model.normalizedRows()) {
      headers.addAll(rr.values().keySet());
    }

    List<String> headerList = new ArrayList<>(headers);

    int r = 0;
    Row headerRow = s.createRow(r++);
    for (int i = 0; i < headerList.size(); i++) {
      headerRow.createCell(i).setCellValue(headerList.get(i));
    }

    for (RowRecord rr : model.normalizedRows()) {
      Row row = s.createRow(r++);
      for (int i = 0; i < headerList.size(); i++) {
        String key = headerList.get(i);
        String value = rr.values().getOrDefault(key, "");
        Cell cell = row.createCell(i);

        if (("ДоляПрибыли".equals(key) || "НакопДоля".equals(key)) && !value.isBlank()) {
          try {
            double numericValue = Double.parseDouble(value);
            cell.setCellValue(numericValue);
            cell.setCellStyle(percentStyle);
          } catch (NumberFormatException e) {
            cell.setCellValue(value);
          }
        } else if (isRuNumericColumn(key) && !value.isBlank()) {
          double numericValue = parseRuNumber(value);
          if (!Double.isNaN(numericValue)) {
            cell.setCellValue(numericValue);
            cell.setCellStyle(groupedNumberStyle);

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

  private boolean isRuNumericColumn(String key) {
    return "Кол-во".equals(key)
        || "Сумма".equals(key)
        || "Сумма приходная".equals(key)
        || "Сумма со скидкой".equals(key)
        || "Прибыль".equals(key)
        || "Мин на 3 дня".equals(key)
        || "Max на 7 дней".equals(key)
        || "Снабжение на 3 дня".equals(key)
        || "Снабжение на 7 дней".equals(key);
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

}
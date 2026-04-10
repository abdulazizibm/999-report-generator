package com.pharmacy999.app;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.*;

public class ReportWriter {

  public void write(Path output, List<RowRecord> rows) throws Exception {
    try (Workbook wb = new XSSFWorkbook()) {
      writeDataSheet(wb, rows);

      try (OutputStream out = Files.newOutputStream(output)) {
        wb.write(out);
      }
    }
  }
  public void write2(Path output, List<RowRecord> rows) throws Exception {
    try (Workbook wb = new XSSFWorkbook()) {
      writeDataSheet2(wb, rows);

      try (OutputStream out = Files.newOutputStream(output)) {
        wb.write(out);
      }
    }
  }


  private void writeDataSheet(Workbook wb, List<RowRecord> normalizedRows) {
    Sheet s = wb.createSheet("Data");
    DataFormat dataFormat = wb.createDataFormat();

    CellStyle groupedNumberStyle = wb.createCellStyle();
    groupedNumberStyle.setDataFormat(dataFormat.getFormat("#,##0.0"));

    CellStyle percentStyle = wb.createCellStyle();
    percentStyle.setDataFormat(dataFormat.getFormat("0.00%"));

    // union headers across all rows (for PoC)
    LinkedHashSet<String> headers = new LinkedHashSet<>();
    for (RowRecord rr : normalizedRows) {
      headers.addAll(rr.values()
          .keySet());
    }

    List<String> headerList = new ArrayList<>(headers);

    int r = 0;
    Row headerRow = s.createRow(r++);
    for (int i = 0; i < headerList.size(); i++) {
      headerRow.createCell(i)
          .setCellValue(headerList.get(i));
    }

    for (RowRecord rr : normalizedRows) {
      Row row = s.createRow(r++);
      for (int i = 0; i < headerList.size(); i++) {
        String key = headerList.get(i);
        String value = rr.values()
            .getOrDefault(key, "");
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
  public void writeDataSheet2(Workbook wb, List<RowRecord> outputRows) {
    String sheetName = WorkbookUtil.createSafeSheetName("Data");
    Sheet sheet = wb.createSheet(sheetName);

    if (outputRows == null || outputRows.isEmpty()) {
      return;
    }

    List<String> headers = new ArrayList<>(outputRows.get(0).values().keySet());

    CellStyle headerStyle = createHeaderStyle(wb);
    CellStyle textStyle = createTextStyle(wb);
    CellStyle qtyStyle = createQtyStyle(wb);
    Map<AaColor, CellStyle> aaStyles = createAaStyles(wb);

    // Header
    Row headerRow = sheet.createRow(0);
    for (int c = 0; c < headers.size(); c++) {
      Cell cell = headerRow.createCell(c);
      cell.setCellValue(headers.get(c));
      cell.setCellStyle(headerStyle);
    }

    // Data
    for (int r = 0; r < outputRows.size(); r++) {
      RowRecord rr = outputRows.get(r);
      Row row = sheet.createRow(r + 1);

      for (int c = 0; c < headers.size(); c++) {
        String columnName = headers.get(c);
        String rawValue = rr.values().getOrDefault(columnName, "");
        String value = rawValue == null ? "" : rawValue.trim();

        Cell cell = row.createCell(c);

        if (isAaColumn(columnName)) {
          cell.setCellValue(value);
          cell.setCellStyle(aaStyles.getOrDefault(resolveAaColor(value), textStyle));
        } else if (isQtyColumn(columnName)) {
          Double numeric = tryParseNumber(value);
          if (numeric != null) {
            cell.setCellValue(numeric);
          } else {
            cell.setCellValue(value);
          }
          cell.setCellStyle(qtyStyle);
        } else {
          cell.setCellValue(value);
          cell.setCellStyle(textStyle);
        }
      }
    }

    // Freeze header
    sheet.createFreezePane(0, 1);

    // Filter
    sheet.setAutoFilter(new org.apache.poi.ss.util.CellRangeAddress(
        0, Math.max(0, outputRows.size()), 0, headers.size() - 1
    ));

    // Widths
    for (int c = 0; c < headers.size(); c++) {
      String header = headers.get(c);

      if ("Наименование".equalsIgnoreCase(header)) {
        sheet.setColumnWidth(c, 40 * 256);
      } else if ("Производитель".equalsIgnoreCase(header)) {
        sheet.setColumnWidth(c, 24 * 256);
      } else if (isAaColumn(header)) {
        sheet.setColumnWidth(c, 16 * 256);
      } else if (isQtyColumn(header)) {
        sheet.setColumnWidth(c, 12 * 256);
      } else {
        sheet.autoSizeColumn(c);
        int current = sheet.getColumnWidth(c);
        sheet.setColumnWidth(c, Math.min(current + 512, 40 * 256));
      }
    }
  }
  public static AaColor resolveAaColor(String aaValue) {
    String v = aaValue == null ? "" : aaValue.trim();
    if (v.isEmpty()) {
      return AaColor.NONE;
    }

    return switch (v) {
      // A block
      case "Ядро A" -> AaColor.DARK_GREEN;
      case "A2/B1" -> AaColor.GREEN;
      case "A1/B2" -> AaColor.LIGHT_GREEN;
      case "A2" -> AaColor.LIGHTER_GREEN;
      case "A1/B1" -> AaColor.VERY_LIGHT_GREEN;

      // B block
      case "Ядро B" -> AaColor.DARK_YELLOW;
      case "B2/C1" -> AaColor.YELLOW;
      case "B1/C2" -> AaColor.LIGHT_YELLOW;
      case "B1", "B2" -> AaColor.LIGHTER_YELLOW;
      case "B1/C1" -> AaColor.VERY_LIGHT_YELLOW;

      // C block
      case "Ядро C" -> AaColor.DARK_RED;
      case "C1", "C2" -> AaColor.LIGHT_RED;

      default -> AaColor.NONE;
    };
  }



  private static boolean isAaColumn(String columnName) {
    return columnName != null && columnName.matches("^AA\\d+$");
  }

  private static boolean isQtyColumn(String columnName) {
    return columnName != null && columnName.matches("^A\\d+$");
  }

  private static Double tryParseNumber(String raw) {
    if (raw == null) {
      return null;
    }

    String s = raw.trim();
    if (s.isEmpty()) {
      return null;
    }

    // Supports both "1234,56" and "1234.56"
    s = s.replace(" ", "").replace(",", ".");

    try {
      return Double.parseDouble(s);
    } catch (NumberFormatException e) {
      return null;
    }
  }

  private static CellStyle createHeaderStyle(Workbook wb) {
    Font font = wb.createFont();
    font.setBold(true);

    CellStyle style = wb.createCellStyle();
    style.setFont(font);
    style.setAlignment(HorizontalAlignment.CENTER);
    style.setVerticalAlignment(VerticalAlignment.CENTER);
    style.setBorderTop(BorderStyle.THIN);
    style.setBorderBottom(BorderStyle.THIN);
    style.setBorderLeft(BorderStyle.THIN);
    style.setBorderRight(BorderStyle.THIN);
    style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
    style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    return style;
  }

  private static CellStyle createTextStyle(Workbook wb) {
    CellStyle style = wb.createCellStyle();
    style.setAlignment(HorizontalAlignment.LEFT);
    style.setVerticalAlignment(VerticalAlignment.CENTER);
    style.setBorderTop(BorderStyle.THIN);
    style.setBorderBottom(BorderStyle.THIN);
    style.setBorderLeft(BorderStyle.THIN);
    style.setBorderRight(BorderStyle.THIN);
    return style;
  }

  private static CellStyle createQtyStyle(Workbook wb) {
    CellStyle style = wb.createCellStyle();
    style.cloneStyleFrom(createTextStyle(wb));
    style.setAlignment(HorizontalAlignment.RIGHT);

    DataFormat df = wb.createDataFormat();
    style.setDataFormat(df.getFormat("0.##"));
    return style;
  }
  private static CellStyle createFilledStyle(Workbook wb, byte[] rgb) {
    XSSFCellStyle style = (XSSFCellStyle) wb.createCellStyle();
    style.cloneStyleFrom(createTextStyle(wb));
    style.setFillForegroundColor(new XSSFColor(rgb, null));
    style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    return style;
  }

  private static Map<AaColor, CellStyle> createAaStyles(Workbook wb) {
    Map<AaColor, CellStyle> styles = new EnumMap<>(AaColor.class);

    styles.put(AaColor.NONE, createTextStyle(wb));

    styles.put(AaColor.DARK_GREEN, createFilledStyle(wb, new byte[]{0x00, 0x61, 0x00}));
    styles.put(AaColor.GREEN, createFilledStyle(wb, new byte[]{0x00, (byte) 0xB0, 0x50}));
    styles.put(AaColor.LIGHT_GREEN, createFilledStyle(wb, new byte[]{(byte) 0x92, (byte) 0xD0, 0x50}));
    styles.put(AaColor.LIGHTER_GREEN, createFilledStyle(wb, new byte[]{(byte) 0xC6, (byte) 0xE0, (byte) 0xB4}));
    styles.put(AaColor.VERY_LIGHT_GREEN, createFilledStyle(wb, new byte[]{(byte) 0xE2, (byte) 0xF0, (byte) 0xD9}));

    styles.put(AaColor.DARK_YELLOW, createFilledStyle(wb, new byte[]{(byte) 0xBF, (byte) 0x8F, 0x00}));
    styles.put(AaColor.YELLOW, createFilledStyle(wb, new byte[]{(byte) 0xFF, (byte) 0xC0, 0x00}));
    styles.put(AaColor.LIGHT_YELLOW, createFilledStyle(wb, new byte[]{(byte) 0xFF, (byte) 0xD9, 0x66}));
    styles.put(AaColor.LIGHTER_YELLOW, createFilledStyle(wb, new byte[]{(byte) 0xFF, (byte) 0xEB, (byte) 0x9C}));
    styles.put(AaColor.VERY_LIGHT_YELLOW, createFilledStyle(wb, new byte[]{(byte) 0xFF, (byte) 0xF2, (byte) 0xCC}));

    styles.put(AaColor.DARK_RED, createFilledStyle(wb, new byte[]{(byte) 0xC0, 0x00, 0x00}));
    styles.put(AaColor.LIGHT_RED, createFilledStyle(wb, new byte[]{(byte) 0xF4, (byte) 0xCC, (byte) 0xCC}));

    return styles;
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

  private boolean isAXcolumn(String key) {
    List<String> list = List.of("A1", "A2", "A3", "A4", "A5", "A6", "A7", "A8", "A9", "A10", "A11",
        "A12", "A13", "A14", "A15", "A16", "A17", "A18", "A19", "A20");
    return list.contains(key);
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
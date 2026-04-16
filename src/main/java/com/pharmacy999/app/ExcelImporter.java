package com.pharmacy999.app;

import java.io.IOException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.util.*;

public class ExcelImporter {

  public interface ProgressCallback {

    void onProgress(long done, long total, String message);
  }

  // Excel rows are 1-based for humans, 0-based in POI:
  private static final int HEADER_ROW_INDEX = 3;     // Excel row 4
  private static final int DATA_START_ROW_INDEX = 4; // Excel row 5

  private final DataFormatter dataFormatter = new DataFormatter(Locale.of("ru", "RU"));

  private static final int BRANCH_COL = 1;         // B - Филиал
  private static final int PRODUCT_COL = 2;        // C - Наименование
  private static final int MANUFACTURER_COL = 3;   // D - Производитель
  private static final int SALES_COL = 4;          // E - Кол-во
  private static final int SUM_COL = 5;            // F - Сумма
  private static final int INCOME_SUM_COL = 6;     // G - Сумма приходн...
  private static final int DISCOUNT_SUM_COL = 7;   // H - Сумма со скидк...
  private static final int PROFIT_COL = 8;         // I - Прибыль

  private static final String BRANCH_HEADER = "Филиал";
  private static final String PRODUCT_HEADER = "Наименование";
  private static final String MANUFACTURER_HEADER = "Производитель";
  private static final String SALES_HEADER = "Кол-во";
  private static final String SUM_HEADER = "Сумма";
  private static final String INCOME_SUM_HEADER = "Сумма приходная";
  private static final String DISCOUNT_SUM_HEADER = "Сумма со скидкой";
  private static final String PROFIT_HEADER = "Прибыль";


  public ImportResult readSalesFile(File f, ProgressCallback cb)
      throws IOException {
    List<RowRecord> out = new ArrayList<>();
    double totalProfit = 0.0;

    cb.onProgress(0, 1, "Reading " + f.getName() + "...");

    try (FileInputStream in = new FileInputStream(f);
        Workbook wb = new XSSFWorkbook(in)) {

      Sheet sheet = wb.getSheetAt(0);
      FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();

      validateHeaderRow(sheet, f.getName());
      int lastRow = sheet.getLastRowNum();
      String currentBranch = "";

      // last row contains totals
      Row totalsRow = sheet.getRow(lastRow);
      if (totalsRow == null) {
        throw new IllegalArgumentException("Totals row not found in " + f.getName());
      }

      totalProfit = parseRuNumber(getCellAsString(totalsRow.getCell(PROFIT_COL), evaluator));

      for (int r = DATA_START_ROW_INDEX; r < lastRow; r++) {
        Row row = sheet.getRow(r);
        if (row == null || isRelevantDataRowEmpty(row)) {
          continue;
        }
        String branch = getCellAsString(row.getCell(BRANCH_COL));

        if (isClosedPharmacy(branch)) {
          continue;
        }

        Map<String, String> values = new LinkedHashMap<>();
        values.put(BRANCH_HEADER, branch);
        values.put(PRODUCT_HEADER, getCellAsString(row.getCell(PRODUCT_COL)));
        values.put(MANUFACTURER_HEADER, getCellAsString(row.getCell(MANUFACTURER_COL)));
        values.put(SALES_HEADER, getCellAsString(row.getCell(SALES_COL)));
        values.put(SUM_HEADER, getCellAsString(row.getCell(SUM_COL)));
        values.put(INCOME_SUM_HEADER, getCellAsString(row.getCell(INCOME_SUM_COL)));
        values.put(DISCOUNT_SUM_HEADER, getCellAsString(row.getCell(DISCOUNT_SUM_COL)));
        values.put(PROFIT_HEADER, getCellAsString(row.getCell(PROFIT_COL)));

        out.add(new RowRecord(values));

      }
    }

    return new ImportResult(out, totalProfit);

  }

  public List<RowRecord> readABCandStockFile(File f, ProgressCallback cb) throws IOException {
    List<RowRecord> result = new ArrayList<>();

    cb.onProgress(0, 1, "Reading " + f.getName() + "...");

    try (FileInputStream in = new FileInputStream(f);
        Workbook wb = new XSSFWorkbook(in)) {

      Sheet sheet = wb.getSheetAt(0);

      Iterator<Row> rowIterator = sheet.iterator();
      if (!rowIterator.hasNext()) {
        return result;
      }

      // --- Read header row --
      int headerRowIndex = detectHeaderRowIndex(sheet);
      Row headerRow = sheet.getRow(headerRowIndex);
      List<String> headers = readHeaders(headerRow);

      int rowIndex = 0;

      // --- Read data rows ---
      for (int r = headerRowIndex + 1; r <= sheet.getLastRowNum(); r++) {
        Row row = sheet.getRow(r);
        if (row == null) {
          continue;
        }

        Map<String, String> values = new LinkedHashMap<>();

        for (int c = 0; c < headers.size(); c++) {
          String header = headers.get(c);
          Cell cell = row.getCell(c, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);

          //String value = readCell(cell);
          String value = getCellAsString(cell);
          values.put(header, value);
        }

        // skip completely empty rows
        if (isEmptyRow(values)) {
          continue;
        }

        result.add(new RowRecord(values));
      }
    }

    return result;
  }


  private List<String> readHeaders(Row headerRow) {
    List<String> headers = new ArrayList<>();

    int lastCell = headerRow.getLastCellNum();
    for (int c = 0; c < lastCell; c++) {
      Cell cell = headerRow.getCell(c, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
      //String header = readCell(cell);
      String header = getCellAsString(cell);

      if (header.isBlank()) {
        header = "COLUMN_" + c; // fallback
      }

      headers.add(header.trim());
    }

    return headers;
  }

  private static int detectHeaderRowIndex(Sheet sheet) {
    Row firstRow = sheet.getRow(0);
    if (firstRow == null) {
      return 0;
    }
    Cell cell = firstRow.getCell(0, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
    if (cell == null) {
      return 0;
    }
    DataFormatter formatter = new DataFormatter();
    String firstRowText = formatter.formatCellValue(cell).toLowerCase(Locale.ROOT);

    if (firstRowText.contains("оборотная ведомость")
        && firstRowText.contains("горизонтальная")) {
      return 1;
    }

    return 0;
  }

  private void validateHeaderRow(Sheet sheet, String fileName) {
    Row headerRow = sheet.getRow(HEADER_ROW_INDEX);
    if (headerRow == null) {
      throw new IllegalArgumentException("Header row (Excel row 4) not found in " + fileName);
    }

    checkHeader(headerRow, BRANCH_COL, "Филиал", fileName);
    checkHeader(headerRow, PRODUCT_COL, "Наименование", fileName);
    checkHeader(headerRow, MANUFACTURER_COL, "Производитель", fileName);
    checkHeader(headerRow, SALES_COL, "Кол-во", fileName);
    checkHeader(headerRow, SUM_COL, "Сумма", fileName);
    checkHeader(headerRow, PROFIT_COL, "Прибыль", fileName);
  }

  private void checkHeader(Row headerRow, int colIndex, String expected, String fileName) {
    String actual = getCellAsString(headerRow.getCell(colIndex));
    if (!actual.equalsIgnoreCase(expected)) {
      throw new IllegalArgumentException(
          "Expected header '" + expected + "' at column " + (colIndex + 1) +
              " in " + fileName + ", but found '" + actual + "'"
      );
    }
  }

  private boolean isRelevantDataRowEmpty(Row row) {
    return getCellAsString(row.getCell(BRANCH_COL)).isEmpty()
        && getCellAsString(row.getCell(PRODUCT_COL)).isEmpty()
        && getCellAsString(row.getCell(MANUFACTURER_COL)).isEmpty()
        && getCellAsString(row.getCell(SALES_COL)).isEmpty()
        && getCellAsString(row.getCell(SUM_COL)).isEmpty()
        && getCellAsString(row.getCell(INCOME_SUM_COL)).isEmpty()
        && getCellAsString(row.getCell(DISCOUNT_SUM_COL)).isEmpty()
        && getCellAsString(row.getCell(PROFIT_COL)).isEmpty();
  }


  private String getCellAsString(Cell cell) {
    if (cell == null) {
      return "";
    }
    return dataFormatter.formatCellValue(cell)
        .trim();
  }

  private static boolean isEmptyRow(Map<String, String> values) {
    for (String v : values.values()) {
      if (v != null && !v.isBlank()) {
        return false;
      }
    }
    return true;
  }

  private String getCellAsString(Cell cell, FormulaEvaluator evaluator) {
    if (cell == null) {
      return "";
    }
    return dataFormatter.formatCellValue(cell, evaluator)
        .trim();
  }

  private double parseRuNumber(String s) {
    if (s == null) {
      return Double.NaN;
    }
    String t = s.trim()
        .replace("\u00A0", "") // NBSP
        .replace(" ", "")
        .replace(".", "")
        .replace(",", ".");
    if (t.isEmpty()) {
      return Double.NaN;
    }
    try {
      return Double.parseDouble(t);
    } catch (NumberFormatException e) {
      return Double.NaN;
    }
  }

  private boolean isClosedPharmacy(String name) {
    String sanitized = name.trim();
    return "Аптека №013".equalsIgnoreCase(sanitized) ||
        "Аптека №005 (Гунча)".equalsIgnoreCase(sanitized) ||
        "Аптека №012 (Ялангач)".equalsIgnoreCase(sanitized) ||
        "Аптека №019 (Ц-1)".equalsIgnoreCase(sanitized) ||
        "Аптека №013 (Шифонур)".equalsIgnoreCase(sanitized);


  }
}
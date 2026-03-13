package com.pharmacy999.app;

import java.io.FileNotFoundException;
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


  public ImportResult readAll(List<File> files, ProgressCallback cb)
      throws IOException {
    List<RowRecord> out = new ArrayList<>();
    double totalProfit = 0.0;
    double totalSales = 0.0;

    for (File f : files) {
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

        for (int r = DATA_START_ROW_INDEX; r <= lastRow; r++) {
          Row row = sheet.getRow(r);
          if (row == null || isRelevantDataRowEmpty(row)) {
            continue;
          }

          String branch = getCellAsString(row.getCell(BRANCH_COL));
          if (!branch.isEmpty()) {
            currentBranch = branch;
          } else {
            // In case branch cells are blank within a grouped block,
            // keep using the last seen branch name.
            branch = currentBranch;
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

          out.add(new RowRecord(f.getName(), r + 1, values));

        }
      }


    }
    return new ImportResult(out, totalProfit, totalSales);


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
  private String getCellAsString(Cell cell, FormulaEvaluator evaluator) {
    if (cell == null) return "";
    return dataFormatter.formatCellValue(cell, evaluator).trim();
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
}
package com.pharmacy999.app;

import java.util.ArrayList;
import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class AbcStockAdder {
  private static final String DRUG_NAME_HEADER = "Наименование";
  private static final String MANUFACTURER_HEADER = "Производитель";
  private static final String END_STOCK_HEADER = "Конечный Остаток";
  private static final Pattern TURNOVER_COLUMN_PATTERN = Pattern.compile("^Оборачиваемость A(\\d+)$");

  /**
   * \\s+ one or more spaces
   * \\s* zero or more spaces
   * 0* any number of leading zeros
   * \\d+ one or more digits
   * (\\d+) capture one or more digits
   */
  private static final Pattern START_STOCK_PHARMACY_PATTERN =
      Pattern.compile("Начальный\\s+остаток\\s+в\\s+аптеке\\s+Аптека\\s*№\\s*0*(\\d+).*", Pattern.CASE_INSENSITIVE | Pattern.UNICODE_CASE);
  private static final Pattern END_STOCK_PHARMACY_PATTERN =
      Pattern.compile("Конечный\\s+остаток\\s+в\\s+аптеке\\s+Аптека\\s*№\\s*0*(\\d+).*", Pattern.CASE_INSENSITIVE | Pattern.UNICODE_CASE);

  /**
   * Enriches ABC output rows with stock columns:
   * AA1, A1, Остаток A1, AA2, A2, Остаток A2, ...
   */
  public static List<RowRecord> addStockColumns(List<RowRecord> abcRows, List<RowRecord> stockRows) {
    List<RowRecord> result = new ArrayList<>();

    if (abcRows == null || abcRows.isEmpty()) {
      return List.of();
    }
    Map<SkuKey, Map<Integer, String>> stockBySkuAndPharmacy = buildDataLookup(stockRows, END_STOCK_HEADER);

    for (RowRecord abcRow : abcRows) {
      Map<String, String> headersToValues = abcRow.headersToValues();

      String drugName = safeTrim(headersToValues.get(DRUG_NAME_HEADER));
      String manufacturer = safeTrim(headersToValues.get(MANUFACTURER_HEADER));
      SkuKey skuKey = new SkuKey(drugName, manufacturer);

      // key -> pharmacy number, value -> stock number of a drug on a current row
      Map<Integer, String> pharmacyToStock =
          stockBySkuAndPharmacy.getOrDefault(skuKey, Collections.emptyMap());

      Map<String, String> out = new LinkedHashMap<>();

      for (Map.Entry<String, String> entry : headersToValues.entrySet()) {
        String columnName = entry.getKey();
        String value = entry.getValue();

        out.put(columnName, value);

        Integer pharmacyNumber = extractPharmacyNumberFromTurnOverColumn(columnName);
        if (pharmacyNumber != null) {
          String stockColumnName = "Остаток A" + pharmacyNumber;
          String stockValue = pharmacyToStock.getOrDefault(pharmacyNumber, "-1");
          out.put(stockColumnName, stockValue);
        }
      }

      result.add(new RowRecord( out));
    }

    return result;
  }

  /**
   * This method takes the file "Оборотная_Ведомость_Горизонтальная" and extracts values from a given column
   * @param stockRows is the input file as a list of RowRecord
   * @param headerName is the column from which values should be extracted, e.g. "Начальный Остаток...",
   * "Количество прихода..." or "Продажа в аптеке..."
   * @return map where key is SkuKey (drug name + manufacturer) and another map as value where
   * key is a pharmacy number and a value is the actual value from @param headerName
   */
  public static Map<SkuKey, Map<Integer, String>> buildDataLookup(List<RowRecord> stockRows, String headerName) {
    Map<SkuKey, Map<Integer, Double>> numericLookup = new LinkedHashMap<>();

    if (stockRows == null || stockRows.isEmpty()) {
      return Map.of();
    }

    for (RowRecord row : stockRows) {
      // e.g. key Наименование -> value 999 КРАФТ ПАКЕТ
      // e.g. key Производитель -> value IBRAT COMPANY
      // e.g. key Кон. Остаток -> value 173
      Map<String, String> headersToValues = row.headersToValues();

      String drugName = safeTrim(headersToValues.get(DRUG_NAME_HEADER));
      String manufacturer = safeTrim(headersToValues.get(MANUFACTURER_HEADER));

      if (drugName.isEmpty() && manufacturer.isEmpty()) {
        continue;
      }
      SkuKey skuKey = new SkuKey(drugName, manufacturer);
      Map<Integer, Double> pharmacyNumberToStock =
          numericLookup.computeIfAbsent(skuKey, k -> new LinkedHashMap<>());

      Pattern headerNamePattern;


      if("конечный остаток".equals(headerName.toLowerCase(Locale.ROOT).trim())){
        headerNamePattern = END_STOCK_PHARMACY_PATTERN;
      }
      else if("начальный остаток".equals(headerName.toLowerCase(Locale.ROOT).trim())){
        headerNamePattern = START_STOCK_PHARMACY_PATTERN;
      }
      else{
        throw new IllegalArgumentException("Extracting data from column " + headerName + " not supported yet!");
      }

      for (Map.Entry<String, String> entry : headersToValues.entrySet()) {
        String columnName = entry.getKey();

        Integer pharmacyNumber = extractPharmacyNumberFromHeader(columnName, headerNamePattern);

        if (pharmacyNumber == null) {
          //current header is not a stock header, keep looking
          continue;
        }

        double stockValue = parseRuNumber(entry.getValue());

        // Sum stocks if duplicate SKU rows exist in stock file
        pharmacyNumberToStock.merge(pharmacyNumber, stockValue, Double::sum);
      }
    }

    Map<SkuKey, Map<Integer, String>> result = new LinkedHashMap<>();
    for (Map.Entry<SkuKey, Map<Integer, Double>> skuEntry : numericLookup.entrySet()) {
      Map<Integer, String> formatted = new LinkedHashMap<>();
      for (Map.Entry<Integer, Double> stockEntry : skuEntry.getValue().entrySet()) {
        formatted.put(stockEntry.getKey(), formatQty(stockEntry.getValue()));
      }
      result.put(skuEntry.getKey(), formatted);
    }

    return result;
  }
  /**
   * Detects A1, A2, A5... and returns the pharmacy number.
   */
  private static Integer extractPharmacyNumberFromTurnOverColumn(String columnName) {
    if (columnName == null) {
      return null;
    }

    Matcher m = TURNOVER_COLUMN_PATTERN.matcher(columnName.trim());
    if (!m.matches()) {
      return null;
    }

    return Integer.parseInt(m.group(1));
  }
  /**
   * Detects stock columns like:
   * "Конечный остаток в аптеке Аптека №001 (Максим Горький) (2)" or
   * "Начальный остаток в аптеке Аптека №001 (Максим Горький) (2)"
   *  and returns 1
   */
  public static Integer extractPharmacyNumberFromHeader(String columnName, Pattern headerNamePattern) {
    if (columnName == null) {
      return null;
    }
    Matcher m = headerNamePattern.matcher(columnName.trim());
    if (!m.matches()) {
      return null;
    }
    // gives the first captured group with \\d+
    String pharmacyNumber = m.group(1);
    return Integer.parseInt(pharmacyNumber);
  }

  private static String safeTrim(String s) {
    return s == null ? "" : s.trim();
  }

  private static double parseRuNumber(String s) {
    if (s == null) {
      return 0.0;
    }
    String t = s.trim()
        .replace("\u00A0", "")
        .replace(" ", "");

    if (t.isEmpty()) {
      return 0.0;
    }

    try {
      if (t.contains(",") && t.contains(".")) {
        t = t.replace(".", "").replace(",", ".");
      } else if (t.contains(",")) {
        t = t.replace(",", ".");
      }
      return Double.parseDouble(t);
    } catch (NumberFormatException e) {
      return Double.NaN;
    }
  }

  private static String formatQty(double value) {
    if (Math.abs(value - Math.rint(value)) < 1e-9) {
      return String.valueOf((long) Math.rint(value));
    }
    return String.valueOf(value).replace(".", ",");
  }
}



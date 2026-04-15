package com.pharmacy999.app;

import java.util.ArrayList;
import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class AbcStockAdder {
  private static final String NAME_HEADER = "Наименование";
  private static final String MANUFACTURER_HEADER = "Производитель";

  private static final Pattern A_COLUMN_PATTERN = Pattern.compile("^A(\\d+)$");
  private static final Pattern STOCK_PHARMACY_PATTERN =
      Pattern.compile("Конечный\\s+остаток\\s+в\\s+аптеке\\s+Аптека\\s*№\\s*0*(\\d+).*", Pattern.CASE_INSENSITIVE | Pattern.UNICODE_CASE);

  /**
   * Enriches 3ABC output rows with stock columns:
   * AA1, A1, Остаток A1, AA2, A2, Остаток A2, ...
   */
  public static List<RowRecord> addStockColumns(List<RowRecord> abcRows, List<RowRecord> stockRows) {
    if (abcRows == null || abcRows.isEmpty()) {
      return List.of();
    }

    Map<SkuKey, Map<Integer, String>> stockBySkuAndPharmacy = buildStockLookup(stockRows);

    List<RowRecord> result = new ArrayList<>();
    int rowIndex = 0;

    for (RowRecord abcRow : abcRows) {
      Map<String, String> input = abcRow.values();

      String name = safeTrim(input.get(NAME_HEADER));
      String manufacturer = safeTrim(input.get(MANUFACTURER_HEADER));
      SkuKey skuKey = new SkuKey(name, manufacturer);

      Map<Integer, String> stockByPharmacy =
          stockBySkuAndPharmacy.getOrDefault(skuKey, Collections.emptyMap());

      Map<String, String> out = new LinkedHashMap<>();

      for (Map.Entry<String, String> entry : input.entrySet()) {
        String columnName = entry.getKey();
        String value = entry.getValue();

        out.put(columnName, value);

        Integer pharmacyNumber = extractPharmacyNumberFromAColumn(columnName);
        if (pharmacyNumber != null) {
          String stockColumnName = "Остаток A" + pharmacyNumber;
          String stockValue = stockByPharmacy.getOrDefault(pharmacyNumber, "0");
          out.put(stockColumnName, stockValue);
        }
      }

      result.add(new RowRecord("sourceFileName",rowIndex++, out));
    }

    return result;
  }
  private static Map<SkuKey, Map<Integer, String>> buildStockLookup(List<RowRecord> stockRows) {
    Map<SkuKey, Map<Integer, Double>> numericLookup = new LinkedHashMap<>();

    if (stockRows == null || stockRows.isEmpty()) {
      return Map.of();
    }

    for (RowRecord row : stockRows) {
      Map<String, String> values = row.values();

      String name = safeTrim(values.get(NAME_HEADER));
      String manufacturer = safeTrim(values.get(MANUFACTURER_HEADER));

      if (name.isEmpty() && manufacturer.isEmpty()) {
        continue;
      }

      SkuKey skuKey = new SkuKey(name, manufacturer);
      Map<Integer, Double> byPharmacy =
          numericLookup.computeIfAbsent(skuKey, k -> new LinkedHashMap<>());

      for (Map.Entry<String, String> entry : values.entrySet()) {
        String columnName = entry.getKey();
        Integer pharmacyNumber = extractPharmacyNumberFromStockHeader(columnName);

        if (pharmacyNumber == null) {
          continue;
        }

        double stockValue = parseRuNumber(entry.getValue());

        // Sum if duplicate SKU rows exist in stock file
        byPharmacy.merge(pharmacyNumber, stockValue, Double::sum);
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
  private static Integer extractPharmacyNumberFromAColumn(String columnName) {
    if (columnName == null) {
      return null;
    }

    Matcher m = A_COLUMN_PATTERN.matcher(columnName.trim());
    if (!m.matches()) {
      return null;
    }

    return Integer.parseInt(m.group(1));
  }
  /**
   * Detects stock columns like:
   * "Конечный остаток в аптеке Аптека №001 (Максим Горький) (2)"
   * and returns 1
   */
  private static Integer extractPharmacyNumberFromStockHeader(String columnName) {
    if (columnName == null) {
      return null;
    }

    Matcher m = STOCK_PHARMACY_PATTERN.matcher(columnName.trim());
    if (!m.matches()) {
      return null;
    }

    return Integer.parseInt(m.group(1));
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
      return 0.0;
    }
  }

  private static String formatQty(double value) {
    if (Math.abs(value - Math.rint(value)) < 1e-9) {
      return String.valueOf((long) Math.rint(value));
    }
    return String.valueOf(value).replace(".", ",");
  }

  private record SkuKey(String name, String manufacturer) {}
}



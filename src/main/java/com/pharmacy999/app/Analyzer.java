package com.pharmacy999.app;

import java.math.RoundingMode;
import java.text.DecimalFormat;
import java.text.DecimalFormatSymbols;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;

public class Analyzer {

    private static final String PROFIT_HEADER = "Прибыль";
    private static final String OUTPUT_COLUMN = "ДоляПрибыли";

    private static final String PHARMACY_HEADER = "Филиал";
    private static final String NAME_HEADER = "Наименование";
    private static final String MANUFACTURER_HEADER = "Производитель";
    private static final String QTY_HEADER = "Кол-во";
    private static final String ABC_HEADER = "ABC";

    public static List<RowRecord> computeTotal(List<RowRecord> rows, double totalProfit) {
        if (rows.isEmpty()) {
            return Collections.emptyList();
        }

        List<RowRecord> outRows = new ArrayList<>(rows.size());


        for (RowRecord rr : rows) {
            Map<String, String> newValues = new LinkedHashMap<>(rr.values());

            double profit = parseRuNumber(rr.values().get(PROFIT_HEADER));
            double ratio = 0.0;

            if (!Double.isNaN(totalProfit) && totalProfit != 0.0) {
                ratio = profit / totalProfit;
            }

            newValues.put(OUTPUT_COLUMN, format(ratio));

            outRows.add(new RowRecord(rr.sourceFile(), rr.rowNumber(), newValues));
        }

        outRows.sort((a, b) -> {
            double va = parseDoubleSafe(a.values().get("ДоляПрибыли"));
            double vb = parseDoubleSafe(b.values().get("ДоляПрибыли"));
            return Double.compare(vb, va); // descending
        });

        double cumulative = 0.0;

        for (int i = 0; i < outRows.size(); i++) {
            RowRecord rr = outRows.get(i);

            Map<String, String> newValues = new LinkedHashMap<>(rr.values());

            double ratio = parseDoubleSafe(newValues.get("ДоляПрибыли"));
            cumulative += ratio;

            if (cumulative > 1.0) {
                cumulative = 1.0;
            }
            newValues.put("НакопДоля", Double.toString(cumulative));

            // ABC classification
            String abc;
            if (cumulative <= 0.80) {
                abc = "A";
            } else if (cumulative <= 0.95) {
                abc = "B";
            } else {
                abc = "C";
            }
            newValues.put("ABC", abc);

            // Min-Max logic
            double sales = parseRuNumber(rr.values().get("Кол-во"));
            double minDouble = ((sales / 30) * 3);
            double maxDouble = ((sales / 30) * 7);

            int min = (int) Math.round(minDouble);
            int max = (int) Math.round(maxDouble);

            newValues.put("Мин на 3 дня", Integer.toString(min));
            newValues.put("Max на 7 дней", Integer.toString(max));


            outRows.set(i, new RowRecord(rr.sourceFile(), rr.rowNumber(), newValues));
        }

        return outRows;
    }
    public static List<RowRecord> computePerPharmacy(List<RowRecord> rows) {
        if (rows.isEmpty()) {
            return Collections.emptyList();
        }
        Map<String, Double> pharmacyTotals = new HashMap<>();
        for (RowRecord rr : rows) {

            String branch = rr.values().get("Филиал");
            double profit = parseRuNumber(rr.values().get("Прибыль"));

            pharmacyTotals.merge(branch, profit, Double::sum);
        }

        List<RowRecord> outRows = new ArrayList<>(rows.size());

        for (RowRecord rr : rows) {
            Map<String, String> newValues = new LinkedHashMap<>(rr.values());
            String branch = rr.values().get("Филиал");

            double profit = parseRuNumber(rr.values().get(PROFIT_HEADER));
            double pharmacyTotal = pharmacyTotals.getOrDefault(branch, 0.0);
            double ratio = 0.0;

            if (pharmacyTotal != 0.0) {
                ratio = profit / pharmacyTotal;
            }

            // store result as string for now (consistent with Map<String,String>)
            newValues.put(OUTPUT_COLUMN, format(ratio));

            outRows.add(new RowRecord(rr.sourceFile(), rr.rowNumber(), newValues));
        }

        Map<String, List<RowRecord>> rowsByBranch = new LinkedHashMap<>();

        for (RowRecord rr : outRows) {
            String branch = rr.values().get("Филиал");
            rowsByBranch.computeIfAbsent(branch, k -> new ArrayList<>()).add(rr);
        }

        List<RowRecord> finalRows = new ArrayList<>(outRows.size());

        for (List<RowRecord> branchRows : rowsByBranch.values()) {

            branchRows.sort((a, b) -> {
                double va = parseDoubleSafe(a.values().get("ДоляПрибыли"));
                double vb = parseDoubleSafe(b.values().get("ДоляПрибыли"));
                return Double.compare(vb, va); // descending inside one pharmacy
            });

            double cumulative = 0.0;
            for (RowRecord rr : branchRows) {
                Map<String, String> newValues = new LinkedHashMap<>(rr.values());

                double ratio = parseDoubleSafe(newValues.get("ДоляПрибыли"));
                cumulative += ratio;

                if (cumulative > 1.0) {
                    cumulative = 1.0;
                }

                newValues.put("НакопДоля", Double.toString(cumulative));

                // ABC classification stays the same, but now per pharmacy
                String abc;
                if (cumulative <= 0.80) {
                    abc = "A";
                } else if (cumulative <= 0.95) {
                    abc = "B";
                } else {
                    abc = "C";
                }
                newValues.put("ABC", abc);

                // Min-Max logic unchanged
                double sales = parseRuNumber(rr.values().get("Кол-во"));
                double minDouble = (sales / 30.0) * 3.0;
                double maxDouble = (sales / 30.0) * 7.0;

                int min = (int) Math.round(minDouble);
                int max = (int) Math.round(maxDouble);

                newValues.put("Мин на 3 дня", Integer.toString(min));
                newValues.put("Max на 7 дней", Integer.toString(max));

                finalRows.add(new RowRecord(rr.sourceFile(), rr.rowNumber(), newValues));
            }

        }
        return finalRows;
    }
    public static List<RowRecord> generateCore(List<List<RowRecord>> files){
        if (files == null || files.isEmpty()) {
            return List.of();
        }

        List<RowRecord> lastFile = files.get(files.size() - 1);

        // Pharmacy order is derived from the last file first, then any missing pharmacies from older files.
        LinkedHashSet<String> pharmacyNamesOrdered = new LinkedHashSet<>();
        collectPharmacies(pharmacyNamesOrdered, lastFile);
        for (List<RowRecord> file : files) {
            collectPharmacies(pharmacyNamesOrdered, file);
        }

        List<String> pharmacies = new ArrayList<>(pharmacyNamesOrdered);

        // Map pharmacy name -> 1-based output index
        Map<String, Integer> pharmacyIndexMap = new LinkedHashMap<>();
        for (int i = 0; i < pharmacies.size(); i++) {
            pharmacyIndexMap.put(pharmacies.get(i), i + 1);
        }

        // SKU -> pharmacy -> aggregate
        Map<SkuKey, Map<String, PharmacySkuAggregate>> data = new LinkedHashMap<>();

        for (List<RowRecord> file : files) {
            for (RowRecord row : file) {
                Map<String, String> values = row.values();

                String pharmacy = safeTrim(values.get(PHARMACY_HEADER));
                String name = safeTrim(values.get(NAME_HEADER));
                String manufacturer = safeTrim(values.get(MANUFACTURER_HEADER));
                String abc = normalizeAbc(values.get(ABC_HEADER));
                String salesString= values.get(QTY_HEADER);

                if (name.isEmpty() && manufacturer.isEmpty()) {
                    continue;
                }
                if (pharmacy.isEmpty()) {
                    continue;
                }

                SkuKey skuKey = new SkuKey(name, manufacturer);

                Map<String, PharmacySkuAggregate> byPharmacy =
                    data.computeIfAbsent(skuKey, k -> new LinkedHashMap<>());

                PharmacySkuAggregate agg =
                    byPharmacy.computeIfAbsent(pharmacy, p -> new PharmacySkuAggregate());

                agg.addAbc(abc);
                agg.addSales(parseRuNumber(salesString));
            }
        }

        // Build output rows
        List<RowRecord> result = new ArrayList<>();
        int outRowIndex = 0;

        for (Map.Entry<SkuKey, Map<String, PharmacySkuAggregate>> entry : data.entrySet()) {
            SkuKey sku = entry.getKey();
            Map<String, PharmacySkuAggregate> byPharmacy = entry.getValue();

            Map<String, String> outValues = new LinkedHashMap<>();
            outValues.put(NAME_HEADER, sku.name());
            outValues.put(MANUFACTURER_HEADER, sku.manufacturer());

            for (String pharmacy : pharmacies) {
                int idx = pharmacyIndexMap.get(pharmacy);
                String aaColumn = "AA" + idx;
                String qtyColumn = "A" + idx;

                PharmacySkuAggregate agg = byPharmacy.get(pharmacy);

                if (agg == null) {
                    outValues.put(aaColumn, "");
                    outValues.put(qtyColumn, "0");
                } else {
                    outValues.put(aaColumn, agg.formatAbcSummary(files.size()));
                    outValues.put(qtyColumn, formatQty(agg.totalSales));
                }
            }

            result.add(new RowRecord(null,outRowIndex++, outValues));
        }

        return result;
    }

    private static void collectPharmacies(Set<String> pharmacyNames, List<RowRecord> rows) {
        for (RowRecord row : rows) {
            String pharmacy = safeTrim(row.values().get(PHARMACY_HEADER));
            if (!pharmacy.isEmpty()) {
                pharmacyNames.add(pharmacy);
            }
        }
    }
    private static String safeTrim(String s) {
        return s == null ? "" : s.trim();
    }
    private static String normalizeAbc(String raw) {
        String s = safeTrim(raw).toUpperCase(Locale.ROOT);
        return switch (s) {
            case "A", "B", "C" -> s;
            default -> "";
        };
    }

    private static String formatQty(double value) {
        /*if (Math.abs(value - Math.rint(value)) < 1e-9) {
            return String.valueOf((long) Math.rint(value));
        }
        return String.valueOf(value).replace(".", ",");*/
        DecimalFormatSymbols symbols = new DecimalFormatSymbols(new Locale("ru", "RU"));
        symbols.setDecimalSeparator(',');
        symbols.setGroupingSeparator(' ');

        DecimalFormat df = new DecimalFormat("0.##", symbols);
        df.setRoundingMode(RoundingMode.HALF_UP);

        return df.format(value);
    }


    private static double parseRuNumber(String s) {
        if (s == null)
            return 0;
        String t = s.trim()
            .replace("\u00A0", "")
            .replace(" ", "")
            .replace(".", "")
            .replace(",", ".");
        if (t.isEmpty())
            return Double.NaN;
        try {
            return Double.parseDouble(t);
        } catch (NumberFormatException e) {
            return Double.NaN;
        }
    }
    private static double parseDoubleSafe(String s) {
        if (s == null || s.isBlank()) return 0.0;
        try {
            return Double.parseDouble(s);
        } catch (NumberFormatException e) {
            return Double.NaN;
        }
    }


    private static String format(double v) {
        if (Double.isNaN(v) || Double.isInfinite(v)) return "";
        return Double.toString(v);
    }
}
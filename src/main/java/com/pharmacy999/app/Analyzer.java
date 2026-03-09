package com.pharmacy999.app;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

public class Analyzer {

    // PoC: choose which column we are dividing
    private static final String PROFIT_HEADER = "Прибыль";
    private static final String OUTPUT_COLUMN = "ДоляПрибыли";

    public static ReportModel compute(List<RowRecord> rows, double totalProfit, double totalSales) {
        if (rows.isEmpty()) {
            return new ReportModel(List.of());
        }

        List<RowRecord> outRows = new ArrayList<>(rows.size());

        for (RowRecord rr : rows) {
            Map<String, String> newValues = new LinkedHashMap<>(rr.values());

            double profit = parseRuNumber(rr.values().get(PROFIT_HEADER));
            double ratio = profit / totalProfit;

            // store result as string for now (consistent with Map<String,String>)
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

            // Supply logic
            int remainingGoods = (int) Math.round(parseRuNumber(rr.values().get("Остаток на конец")));
            int buyForMin = min > remainingGoods ? min - remainingGoods : 0;
            int buyForMax = max > remainingGoods ? max - remainingGoods : 0;

            newValues.put("Снабжение на 3 дня", Integer.toString(buyForMin));
            newValues.put("Снабжение на 7 дней", Integer.toString(buyForMax));

            outRows.set(i, new RowRecord(rr.sourceFile(), rr.rowNumber(), newValues));
        }

        return new ReportModel(outRows);
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
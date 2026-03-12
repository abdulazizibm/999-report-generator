package com.pharmacy999.app;

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
    private static final int HEADER_ROW_INDEX = 2;     // Excel row 3
    private static final int DATA_START_ROW_INDEX = 3; // Excel row 4
    private static final int TOTALS_ROW_INDEX = 1;
    private static final String PROFIT_HEADER = "Прибыль";
    private static final String SALES_HEADER = "Кол-во";// Excel row 2 (Pribil)
    //private final DataFormatter dataFormatter = new DataFormatter(new Locale("ru", "RU"));
    private final DataFormatter dataFormatter = new DataFormatter(Locale.of("ru", "RU"));



    public ImportResult readAll(List<File> files, ProgressCallback cb) throws Exception {
        List<RowRecord> out = new ArrayList<>();
        double totalProfit = Double.NaN;// empty list for all Excel files
        double totalSales = Double.NaN;

        for (File f : files) {
            cb.onProgress(0, 1, "Reading " + f.getName() + "...");

            try (FileInputStream in = new FileInputStream(f); // read the file from the disk
                Workbook wb = new XSSFWorkbook(in)) { //represents te entire Excel file in memory

                Sheet sheet = wb.getSheetAt(0);

                Row headerRow = sheet.getRow(HEADER_ROW_INDEX);
                if (headerRow == null) continue; //skip Excel file if header is missing

                List<String> headers = readHeaders(headerRow);
                int profitCol = headers.indexOf(PROFIT_HEADER);
                if (profitCol < 0) {
                    throw new IllegalArgumentException("Column '" + PROFIT_HEADER + "' not found in " + f.getName());
                }
                int salesCol = headers.indexOf(SALES_HEADER);
                if (salesCol < 0) {
                    throw new IllegalArgumentException("Column '" + SALES_HEADER + "' not found in " + f.getName());
                }

                Row totalsRow = sheet.getRow(TOTALS_ROW_INDEX);
                if (totalsRow == null) {
                    throw new IllegalArgumentException("Totals row (Excel row 2) not found in " + f.getName());
                }

                String totalProfitStr = getCellAsString(totalsRow.getCell(profitCol));
                String salesStr = getCellAsString(totalsRow.getCell(salesCol));
                double fileTotalProfit = parseRuNumber(totalProfitStr);
                totalSales = parseRuNumber(salesStr);

                totalProfit = fileTotalProfit;

                int lastRow = sheet.getLastRowNum(); // index of the last row

                //start iterating through every row
                for (int r = DATA_START_ROW_INDEX; r <= lastRow; r++) {
                    Row row = sheet.getRow(r);
                    if (row == null) continue;

                    Map<String, String> values = new LinkedHashMap<>();
                    for (int c = 0; c < headers.size(); c++) {
                        values.put(headers.get(c), getCellAsString(row.getCell(c)));
                    }

                    // rowNumber store as Excel 1-based
                    out.add(new RowRecord(f.getName(), r + 1, values));
                }
            }
        }

        return new ImportResult(out, totalProfit, totalSales);
    }

    private List<String> readHeaders(Row headerRow) {
        List<String> headers = new ArrayList<>();
        int last = headerRow.getLastCellNum();
        for (int c = 0; c < last; c++) {
            String h = getCellAsString(headerRow.getCell(c)).trim();
            headers.add(h.isEmpty() ? ("COL_" + (c + 1)) : h);
        }
        return headers;
    }
    private String getCellAsString(Cell cell) {
        if (cell == null) return "";
        return dataFormatter.formatCellValue(cell).trim();
    }

    private double parseRuNumber(String s) {
        if (s == null) return Double.NaN;
        String t = s.trim()
            .replace("\u00A0", "") // NBSP
            .replace(" ", "")
            .replace(".", "")
            .replace(",", ".");
        if (t.isEmpty()) return Double.NaN;
        try {
            return Double.parseDouble(t);
        } catch (NumberFormatException e) {
            return Double.NaN;
        }
    }
}
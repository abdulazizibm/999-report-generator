package com.pharmacy999.app;

import java.io.File;
import java.io.IOException;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;


public class ReportService {

    public interface ProgressCallback {
        void onProgress(long done, long total, String message);
    }

    public void generateReport(List<File> inputs, Path output, ProgressCallback cb, CalculationMode mode)
        throws Exception {
        List<RowRecord> outputRows;
        ReportWriter writer = new ReportWriter();

        if(mode.equals(CalculationMode.ABC_TOTAL) || mode.equals(CalculationMode.ABC_PER_PHARMACY)){
            outputRows = generateReport2(inputs, cb, mode);
            cb.onProgress(2, 3, "Создаю отчет…");
            writer.write(output, outputRows);

        }
        else if (mode.equals(CalculationMode.ABC_TOTAL_3_MONTHS)){
            outputRows = generateReport3(inputs, cb);
            cb.onProgress(2, 3, "Создаю отчет…");
            writer.write2(output, outputRows);
        }
        else{
            outputRows = generateReport4(inputs, cb);
            cb.onProgress(2, 3, "Создаю отчет…");
            writer.write2(output, outputRows);

        }


        cb.onProgress(3, 3, "Готово.");
    }

    public List<RowRecord> generateReport2(List<File> inputs, ProgressCallback cb, CalculationMode mode) throws Exception {
        // PoC: read all rows, do simple math, write report
        ExcelImporter importer = new ExcelImporter();
        ImportResult res = importer.readSalesFile(inputs.getFirst(), (done, total, msg) -> cb.onProgress(done, total, msg));

        cb.onProgress(1, 3, "Анализирую…");

        List<RowRecord> outputRows;
        if(mode.equals(CalculationMode.ABC_TOTAL)){
            outputRows = Analyzer.computeTotal(res.rows(), res.totalProfit());
        }
        else {
            outputRows = Analyzer.computePerPharmacy(res.rows());
        }
        return outputRows;

    }

    public List<RowRecord> generateReport3(List<File> inputs, ProgressCallback cb)
        throws IOException {
        ExcelImporter importer = new ExcelImporter();
        List<List<RowRecord>> listOfRows = new ArrayList<>();

        for(File input: inputs){
            ImportResult res = importer.readSalesFile(input, (done, total, msg) -> cb.onProgress(done, total, msg));
            List<RowRecord> rows = Analyzer.computePerPharmacy(res.rows());
            listOfRows.add(rows);
        }
        return Analyzer.generateCore(listOfRows);

    }
    public List<RowRecord> generateReport4(List<File> inputs, ProgressCallback cb)
        throws IOException {
        ExcelImporter importer = new ExcelImporter();
        List<RowRecord> abcFile = new ArrayList<>();
        List<RowRecord> stockFile = new ArrayList<>(); // оборотная ведомость горизонтальная


        for(File input: inputs){
            String fileName = input.getName().toLowerCase(Locale.ROOT);
            if(fileName.contains("оборотная_ведомость_горизонтальная")){
                stockFile = importer.readABCandStockFile(input, (done, total, msg) -> cb.onProgress(done, total, msg));
            }
            else{
                abcFile = importer.readABCandStockFile(input, (done, total, msg) -> cb.onProgress(done, total, msg));

            }

        }
        return AbcStockAdder.addStockColumns(abcFile, stockFile);

    }
}

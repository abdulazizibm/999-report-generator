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
        else if (mode.equals(CalculationMode.ABC_3_MONTHS_CORE)){
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

    public List<RowRecord> generateReport3(List<File> files, ProgressCallback cb)
        throws IOException {
        ExcelImporter importer = new ExcelImporter();
        List<List<RowRecord>> salesFileRows = new ArrayList<>();
        List<RowRecord> stockRows = new ArrayList<>();

        for(File file: files){
            String fileName = file.getName().toLowerCase(Locale.ROOT);

            if(!fileName.contains("оборотная_ведомость_горизонтальная")){
                ImportResult res = importer.readSalesFile(file, (done, total, msg) -> cb.onProgress(done, total, msg));
                List<RowRecord> rows = Analyzer.computePerPharmacy(res.rows());
                salesFileRows.add(rows);
            }
            else{
                stockRows = importer.readABCandStockFile(file, (done, total, msg) -> cb.onProgress(done, total, msg));
            }

        }
        return Analyzer.generateCore(salesFileRows, stockRows);

    }
    public List<RowRecord> generateReport4(List<File> inputs, ProgressCallback cb)
        throws IOException {
        ExcelImporter importer = new ExcelImporter();
        List<RowRecord> abcRows = new ArrayList<>();
        List<RowRecord> stockRows = new ArrayList<>(); // оборотная ведомость горизонтальная


        for(File input: inputs){
            String fileName = input.getName().toLowerCase(Locale.ROOT);
            if(fileName.contains("оборотная_ведомость_горизонтальная")){
                stockRows = importer.readABCandStockFile(input, (done, total, msg) -> cb.onProgress(done, total, msg));
            }
            else{
                abcRows = importer.readABCandStockFile(input, (done, total, msg) -> cb.onProgress(done, total, msg));

            }

        }
        return AbcStockAdder.addStockColumns(abcRows, stockRows);

    }
}

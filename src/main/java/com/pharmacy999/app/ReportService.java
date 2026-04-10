package com.pharmacy999.app;

import java.io.File;
import java.io.IOException;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;


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
        else{
            outputRows = generateReport3(inputs, cb);
            cb.onProgress(2, 3, "Создаю отчет…");
            writer.write2(output, outputRows);
        }


        cb.onProgress(3, 3, "Готово.");
    }

    public List<RowRecord> generateReport2(List<File> inputs, ProgressCallback cb, CalculationMode mode) throws Exception {
        // PoC: read all rows, do simple math, write report
        ExcelImporter importer = new ExcelImporter();
        ImportResult res = importer.readAll(inputs.getFirst(), (done, total, msg) -> cb.onProgress(done, total, msg));

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
            ImportResult res = importer.readAll(input, (done, total, msg) -> cb.onProgress(done, total, msg));
            List<RowRecord> rows = Analyzer.computePerPharmacy(res.rows());
            System.out.println("DEBUG_ROW_RECORD: " + rows.getFirst());
            listOfRows.add(rows);
        }
        return Analyzer.generateCore(listOfRows);

    }
}

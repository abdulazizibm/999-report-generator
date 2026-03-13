package com.pharmacy999.app;

import java.io.File;
import java.nio.file.Path;
import java.util.List;


public class ReportService {

    public interface ProgressCallback {
        void onProgress(long done, long total, String message);
    }

    public void generateReport(List<File> inputs, Path output, ProgressCallback cb, CalculationMode mode) throws Exception {
        // PoC: read all rows, do simple math, write report
        ExcelImporter importer = new ExcelImporter();
        ImportResult res = importer.readAll(inputs, (done, total, msg) -> cb.onProgress(done, total, msg));

        cb.onProgress(1, 3, "Анализирую…");
        ReportModel model;
        if(mode.equals(CalculationMode.ABC_TOTAL)){
            model = Analyzer.computeTotal(res.rows(), res.totalProfit());
        }
        else{
            model = Analyzer.computePerPharmacy(res.rows());
        }

        cb.onProgress(2, 3, "Создаю отчет…");
        ReportWriter writer = new ReportWriter();
        writer.write(output, model);

        cb.onProgress(3, 3, "Готово.");
    }
}

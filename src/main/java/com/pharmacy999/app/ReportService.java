package com.pharmacy999.app;

import java.io.File;
import java.nio.file.Path;
import java.util.List;


public class ReportService {

    public interface ProgressCallback {
        void onProgress(long done, long total, String message);
    }

    public void generateReport(List<File> inputs, Path output, ProgressCallback cb) throws Exception {
        // PoC: read all rows, do simple math, write report
        ExcelImporter importer = new ExcelImporter();
        ImportResult res = importer.readAll(inputs, (done, total, msg) -> cb.onProgress(done, total, msg));

        cb.onProgress(1, 3, "Analyzing…");
        ReportModel model = Analyzer.compute(res.rows(), res.totalProfit());

        cb.onProgress(2, 3, "Writing report…");
        ReportWriter writer = new ReportWriter();
        writer.write(output, model);

        cb.onProgress(3, 3, "Finished.");
    }
}

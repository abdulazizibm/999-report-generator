package com.pharmacy999.app;

import java.util.List;
import java.util.Map;

public record ReportModel(
    List<RowRecord> normalizedRows
) {

}
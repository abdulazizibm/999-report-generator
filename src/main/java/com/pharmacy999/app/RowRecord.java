package com.pharmacy999.app;

import java.util.Map;

public record RowRecord(String sourceFile, int rowNumber, Map<String, String> values) {}
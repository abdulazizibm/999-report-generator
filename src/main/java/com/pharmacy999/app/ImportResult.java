package com.pharmacy999.app;

import java.util.List;

public record ImportResult(List<RowRecord> rows, double totalProfit) {}
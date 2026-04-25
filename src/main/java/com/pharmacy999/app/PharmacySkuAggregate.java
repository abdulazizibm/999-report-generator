package com.pharmacy999.app;

import java.util.ArrayList;
import java.util.List;

public class PharmacySkuAggregate {

  int countA;
  int countB;
  int countC;
  int countF;
  double totalSales = 0.0;

  void addAbc(String abc) {
    switch (abc) {
      case "A" -> countA++;
      case "B" -> countB++;
      case "C" -> countC++;
      case "F" -> countF++;
      default -> {
        // ignore blanks/invalid values
      }
    }
  }
  void addSales(double sales){
    totalSales += sales;
  }

  String formatAbcSummary(int totalFiles) {
    if (countA == totalFiles && totalFiles > 0) {
      return "Ядро A";
    }
    if (countB == totalFiles && totalFiles > 0) {
      return "Ядро B";
    }
    if (countC == totalFiles && totalFiles > 0) {
      return "Ядро C";
    }
    if (countF == totalFiles && totalFiles > 0) {
      return "Ядро F ⚠";
    }

    List<String> parts = new ArrayList<>(3);
    if (countA > 0) {
      parts.add("A" + countA);
    }
    if (countB > 0) {
      parts.add("B" + countB);
    }
    if (countC > 0) {
      parts.add("C" + countC);
    }
    if (countF > 0) {
      parts.add("F" + countF + " ⚠");
    }

    return String.join("/", parts);
  }
}


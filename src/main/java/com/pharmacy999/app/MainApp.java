package com.pharmacy999.app;

import static java.text.MessageFormat.format;

import javafx.application.Application;
import javafx.concurrent.Task;
import javafx.geometry.Insets;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.layout.*;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

public class MainApp extends Application {

  private final List<File> selectedFiles = new ArrayList<>();
  private final ListView<String> fileList = new ListView<>();
  private final ProgressBar progressBar = new ProgressBar(0);
  private final Label statusLabel = new Label("Выберите Excel файлы для начала.");
  private final String VERSION = "v1.1";

  @Override
  public void start(Stage stage) {
    Button pickFilesBtn = new Button("Выберите Excel файлы…");
    Button runBtn = new Button("Запустить");
    runBtn.setDisable(true);

    RadioButton abcPerPharmacyRadio = new RadioButton("ABC по аптекам");
    RadioButton abcTotalRadio = new RadioButton("ABC по сети");
    RadioButton abc3monthsCore = new RadioButton("ABC 3 месяца (Палитра)");
    RadioButton abcAndStock = new RadioButton("ABC + Остаток на конец");


    ToggleGroup modeGroup = new ToggleGroup();
    abcPerPharmacyRadio.setToggleGroup(modeGroup);
    abcTotalRadio.setToggleGroup(modeGroup);
    abc3monthsCore.setToggleGroup(modeGroup);
    abcAndStock.setToggleGroup(modeGroup);

    Label inputHintLabel = new Label();
    inputHintLabel.setWrapText(true);
    inputHintLabel.setStyle("""
    -fx-background-color: #f3f6fa;
    -fx-padding: 8;
    -fx-border-color: #d0d7de;
    -fx-border-radius: 4;
    -fx-background-radius: 4;
    """);

    // default selection
    abcPerPharmacyRadio.setSelected(true);
    inputHintLabel.setText("""
        Загрузите один Excel файл с продажами за период.""");

    modeGroup.selectedToggleProperty().addListener((obs, oldToggle, newToggle) -> {
      if (newToggle == abcPerPharmacyRadio) {
        inputHintLabel.setText("""
            Загрузите один Excel файл с продажами за период.""");
      }
      else if (newToggle == abcTotalRadio) {
        inputHintLabel.setText("""
            Загрузите один Excel файл с продажами за период.""");
      }
      else if (newToggle == abc3monthsCore) {
        inputHintLabel.setText("""
        Загрузите 3 файла с продажами за 3 мес. и опционально Горизонт. Оборот. Ведомость за тот же период.""");
      }
      else if (newToggle == abcAndStock) {
        inputHintLabel.setText("""
        Загрузите файл "ABC 3 месяца (палитра)" и горизонтальную оборотную ведомость.
        """);
      }
    });

    pickFilesBtn.setOnAction(e -> {
      FileChooser chooser = new FileChooser();
      chooser.setTitle("Выберите Excel файлы");
      chooser.getExtensionFilters()
          .add(
              new FileChooser.ExtensionFilter("Excel files (*.xlsx)", "*.xlsx")
          );
      List<File> files = chooser.showOpenMultipleDialog(stage);
      if (files != null && !files.isEmpty()) {

        for (File file : files) {
          if (!selectedFiles.contains(file)) {
            selectedFiles.add(file);
            fileList.getItems().add(file.getAbsolutePath());
          }
        }
        statusLabel.setText("Выбраны " + selectedFiles.size() + " файл(а).");
        runBtn.setDisable(selectedFiles.isEmpty());
      }
    });


    runBtn.setOnAction(e -> {
      FileChooser saveChooser = new FileChooser();
      saveChooser.setTitle("Сохранить отчет в");
      saveChooser.getExtensionFilters()
          .add(
              new FileChooser.ExtensionFilter("Excel files (*.xlsx)", "*.xlsx")
          );

      CalculationMode mode;

      if(abcPerPharmacyRadio.isSelected()){
        mode = CalculationMode.ABC_PER_PHARMACY;
      }
      else if(abcTotalRadio.isSelected()){
        mode = CalculationMode.ABC_TOTAL;
      }
      else if(abc3monthsCore.isSelected()){
        mode = CalculationMode.ABC_3_MONTHS_CORE;
      }
      else{
        mode = CalculationMode.ABC_and_STOCK;
      }

      String fileName = selectedFiles.getFirst().getName();
      if(mode.equals(CalculationMode.ABC_PER_PHARMACY)){
        saveChooser.setInitialFileName("ABC по аптекам " + fileName);
      }
      else if (mode.equals(CalculationMode.ABC_TOTAL)){
        saveChooser.setInitialFileName("ABC по cети " + fileName);
      }
      else if(mode.equals(CalculationMode.ABC_3_MONTHS_CORE)){
        saveChooser.setInitialFileName("ABC 3 месяца палитра");
        //TODO extract dates from input files and set those after the name,
        //TODO e.g., ABC 3 месяца палитра с 01.01.2026 - 31.03.2026
      }
      else{
        saveChooser.setInitialFileName("ABC 3 месяца палитра + остаток");
      }
      File out = saveChooser.showSaveDialog(stage);
      if (out == null) {
        return;
      }

      runBtn.setDisable(true);
      pickFilesBtn.setDisable(true);
      abcTotalRadio.setDisable(true);
      abcPerPharmacyRadio.setDisable(true);
      abc3monthsCore.setDisable(true);
      abcAndStock.setDisable(true);

      progressBar.setProgress(0);
      statusLabel.setText("Работаю…");

      Task<Void> task = new Task<>() {
        @Override
        protected Void call() throws Exception {
          ReportService service = new ReportService();
          service.generateReport(
              selectedFiles,
              out.toPath(),
              (done, total, msg) -> {
                updateProgress(done, total);
                updateMessage(msg);
              }, mode
          );
          return null;
        }
      };

      statusLabel.textProperty()
          .bind(task.messageProperty());
      progressBar.progressProperty()
          .bind(task.progressProperty());

      task.setOnSucceeded(ev -> {
        statusLabel.textProperty()
            .unbind();
        progressBar.progressProperty()
            .unbind();
        statusLabel.setText("Готово. Отчет сохранен в: " + out.getAbsolutePath());
        pickFilesBtn.setDisable(false);
        runBtn.setDisable(false);
        abcTotalRadio.setDisable(false);
        abcPerPharmacyRadio.setDisable(false);
        abc3monthsCore.setDisable(false);
        abcAndStock.setDisable(false);
      });

      task.setOnFailed(ev -> {
        statusLabel.textProperty()
            .unbind();
        progressBar.progressProperty()
            .unbind();
        Throwable ex = task.getException();
        statusLabel.setText("Ошибка: " + (ex != null ? ex.getMessage() : "Неизвестная ошибка"));

        pickFilesBtn.setDisable(false);
        runBtn.setDisable(false);
        abcTotalRadio.setDisable(false);
        abcPerPharmacyRadio.setDisable(false);
        abc3monthsCore.setDisable(false);
        abcAndStock.setDisable(false);
      });

      new Thread(task, "report-task").start();
    });

    VBox root = new VBox(10,
        new HBox(10, pickFilesBtn, runBtn),
        new Label("Режим анализа:"),
        new HBox(16, abcPerPharmacyRadio, abcTotalRadio, abc3monthsCore, abcAndStock),
        inputHintLabel,
        new Label("Выбранные файлы:"),
        fileList,
        new Label("Статус:"),
        progressBar,
        statusLabel
    );
    root.setPadding(new Insets(12));
    fileList.setPrefHeight(200);

    stage.setTitle(format("ABC - Min/Max Report Generator ({0})", VERSION));
    stage.setScene(new Scene(root, 720, 500));
    stage.show();
  }

  public static void main(String[] args) {
    launch(args);
  }
}
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


    ToggleGroup modeGroup = new ToggleGroup();
    abcPerPharmacyRadio.setToggleGroup(modeGroup);
    abcTotalRadio.setToggleGroup(modeGroup);

    // default selection
    abcPerPharmacyRadio.setSelected(true);

    pickFilesBtn.setOnAction(e -> {
      FileChooser chooser = new FileChooser();
      chooser.setTitle("Выберите Excel файлы");
      chooser.getExtensionFilters()
          .add(
              new FileChooser.ExtensionFilter("Excel files (*.xlsx)", "*.xlsx")
          );
      List<File> files = chooser.showOpenMultipleDialog(stage);
      if (files != null && !files.isEmpty()) {
        selectedFiles.clear();
        selectedFiles.addAll(files);
        fileList.getItems()
            .setAll(selectedFiles.stream()
                .map(File::getAbsolutePath)
                .toList());
        statusLabel.setText("Выбраны " + selectedFiles.size() + " файлы.");
        runBtn.setDisable(false);
      }
    });


    runBtn.setOnAction(e -> {
      FileChooser saveChooser = new FileChooser();
      saveChooser.setTitle("Сохранить отчет в");
      saveChooser.getExtensionFilters()
          .add(
              new FileChooser.ExtensionFilter("Excel files (*.xlsx)", "*.xlsx")
          );

      CalculationMode mode = abcPerPharmacyRadio.isSelected()
          ? CalculationMode.ABC_PER_PHARMACY
          : CalculationMode.ABC_TOTAL;

      String fileName = selectedFiles.getFirst().getName();
      if(mode.equals(CalculationMode.ABC_PER_PHARMACY)){
        saveChooser.setInitialFileName("ABC по аптекам " + fileName);
      }
      else{
        saveChooser.setInitialFileName("ABC по cети " + fileName);
      }
      File out = saveChooser.showSaveDialog(stage);
      if (out == null) {
        return;
      }

      runBtn.setDisable(true);
      pickFilesBtn.setDisable(true);
      abcTotalRadio.setDisable(true);
      abcPerPharmacyRadio.setDisable(true);

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
      });

      new Thread(task, "report-task").start();
    });

    VBox root = new VBox(10,
        new HBox(10, pickFilesBtn, runBtn),
        new Label("Режим анализа:"),
        new HBox(16, abcPerPharmacyRadio, abcTotalRadio),
        new Label("Выбранные файлы:"),
        fileList,
        new Label("Статус:"),
        progressBar,
        statusLabel
    );
    root.setPadding(new Insets(12));
    fileList.setPrefHeight(200);

    stage.setTitle(format("ABC - Min/Max Report Generator ({0})", VERSION));
    stage.setScene(new Scene(root, 720, 420));
    stage.show();
  }

  public static void main(String[] args) {
    launch(args);
  }
}
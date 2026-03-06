package com.pharmacy999.app;

import javafx.application.Application;
import javafx.concurrent.Task;
import javafx.geometry.Insets;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.layout.*;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

import java.io.File;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;

public class MainApp extends Application {

    private final List<File> selectedFiles = new ArrayList<>();
    private final ListView<String> fileList = new ListView<>();
    private final ProgressBar progressBar = new ProgressBar(0);
    private final Label statusLabel = new Label("Select Excel files to start.");

    @Override
    public void start(Stage stage) {
        Button pickFilesBtn = new Button("Select Excel files…");
        Button runBtn = new Button("Run");
        runBtn.setDisable(true);

        pickFilesBtn.setOnAction(e -> {
            FileChooser chooser = new FileChooser();
            chooser.setTitle("Select Excel files");
            chooser.getExtensionFilters().add(
                    new FileChooser.ExtensionFilter("Excel files (*.xlsx)", "*.xlsx")
            );
            List<File> files = chooser.showOpenMultipleDialog(stage);
            if (files != null && !files.isEmpty()) {
                selectedFiles.clear();
                selectedFiles.addAll(files);
                fileList.getItems().setAll(selectedFiles.stream().map(File::getAbsolutePath).toList());
                statusLabel.setText("Selected " + selectedFiles.size() + " file(s).");
                runBtn.setDisable(false);
            }
        });

        runBtn.setOnAction(e -> {
            FileChooser saveChooser = new FileChooser();
            saveChooser.setTitle("Save output report");
            saveChooser.getExtensionFilters().add(
                    new FileChooser.ExtensionFilter("Excel files (*.xlsx)", "*.xlsx")
            );
            saveChooser.setInitialFileName("report.xlsx");
            File out = saveChooser.showSaveDialog(stage);
            if (out == null) return;

            runBtn.setDisable(true);
            pickFilesBtn.setDisable(true);
            progressBar.setProgress(0);
            statusLabel.setText("Running…");

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
                            }
                    );
                    return null;
                }
            };

            statusLabel.textProperty().bind(task.messageProperty());
            progressBar.progressProperty().bind(task.progressProperty());

            task.setOnSucceeded(ev -> {
                statusLabel.textProperty().unbind();
                progressBar.progressProperty().unbind();
                statusLabel.setText("Done. Report saved to: " + out.getAbsolutePath());
                pickFilesBtn.setDisable(false);
                runBtn.setDisable(false);
            });

            task.setOnFailed(ev -> {
                statusLabel.textProperty().unbind();
                progressBar.progressProperty().unbind();
                Throwable ex = task.getException();
                statusLabel.setText("Failed: " + (ex != null ? ex.getMessage() : "unknown error"));
                pickFilesBtn.setDisable(false);
                runBtn.setDisable(false);
            });

            new Thread(task, "report-task").start();
        });

        VBox root = new VBox(10,
                new HBox(10, pickFilesBtn, runBtn),
                new Label("Selected files:"),
                fileList,
                new Label("Progress:"),
                progressBar,
                statusLabel
        );
        root.setPadding(new Insets(12));
        fileList.setPrefHeight(200);

        stage.setTitle("Excel Report PoC");
        stage.setScene(new Scene(root, 720, 420));
        stage.show();
    }

    public static void main(String[] args) {
        launch(args);
    }
}
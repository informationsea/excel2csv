/*
 *  excel2csv  xls/xlsx/csv/tsv converter
 *  Copyright (C) 2015 Yasunobu OKAMURA
 *
 *  This program is free software: you can redistribute it and/or modify
 *  it under the terms of the GNU General Public License as published by
 *  the Free Software Foundation, either version 3 of the License, or
 *  (at your option) any later version.
 *
 *  This program is distributed in the hope that it will be useful,
 *  but WITHOUT ANY WARRANTY; without even the implied warranty of
 *  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 *  GNU General Public License for more details.
 *
 *  You should have received a copy of the GNU General Public License
 *  along with this program.  If not, see <http://www.gnu.org/licenses/>.
 */

package info.informationsea.java.excel2csv.gui;

import info.informationsea.java.excel2csv.Converter;
import info.informationsea.java.excel2csv.Utilities;
import info.informationsea.java.excel2csv.VersionResolver;
import javafx.beans.value.ChangeListener;
import javafx.beans.value.ObservableValue;
import javafx.fxml.FXML;
import javafx.fxml.Initializable;
import javafx.scene.control.Alert;
import javafx.scene.control.CheckBox;
import javafx.scene.control.CheckMenuItem;
import javafx.scene.control.MenuBar;
import javafx.scene.input.DragEvent;
import javafx.scene.input.Dragboard;
import javafx.scene.input.TransferMode;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import lombok.Setter;
import lombok.extern.slf4j.Slf4j;

import java.io.File;
import java.net.URL;
import java.util.List;
import java.util.ResourceBundle;
import java.util.prefs.BackingStoreException;
import java.util.prefs.Preferences;

@Slf4j
public class MainWindowController implements Initializable {

    @Setter
    private Stage stage;

    @FXML
    private MenuBar menuBar;

    @FXML
    private CheckBox prettyTable;

    @FXML
    private CheckBox convertType;

    private Preferences preferences;

    @Override
    public void initialize(URL location, ResourceBundle resources) {
        preferences = Application.getApplicationPreferences();

        prettyTable.setSelected(preferences.getBoolean(PRETTY_TABLE, false));
        prettyTable.selectedProperty().addListener(new ChangeListener<Boolean>() {
            @Override
            public void changed(ObservableValue<? extends Boolean> observable, Boolean oldValue, Boolean newValue) {
                preferences.putBoolean(PRETTY_TABLE, newValue);
                try {
                    preferences.sync();
                } catch (BackingStoreException e) {
                    e.printStackTrace();
                }
            }
        });

        convertType.setSelected(preferences.getBoolean(CELL_TYPE_CONVERT, false));
        convertType.selectedProperty().addListener(new ChangeListener<Boolean>() {
            @Override
            public void changed(ObservableValue<? extends Boolean> observable, Boolean oldValue, Boolean newValue) {
                preferences.putBoolean(CELL_TYPE_CONVERT, newValue);
                try {
                    preferences.sync();
                } catch (BackingStoreException e) {
                    e.printStackTrace();
                }
            }
        });

        menuBar.setUseSystemMenuBar(true);
    }

    private static final String OPEN_DIRECTORY_DEFAULT = "OPEN_DIRECTORY_DEFAULT";
    private static final String SAVE_DIRECTORY_DEFAULT = "SAVE_DIRECTORY_DEFAULT";
    private static final String PRETTY_TABLE = "PRETTY_TABLE";
    private static final String CELL_TYPE_CONVERT = "CELL_TYPE_CONVERT";

    @FXML
    public void clickConvert() {
        FileChooser fileChooser = new FileChooser();
        fileChooser.setTitle("Open files to convert");
        fileChooser.getExtensionFilters().addAll(new FileChooser.ExtensionFilter("All Supported Files", "*.csv", "*.txt", "*.xls", "*.xlsx"));
        fileChooser.setInitialDirectory(new File(preferences.get(OPEN_DIRECTORY_DEFAULT, System.getProperty("user.home"))));
        List<File> selected = fileChooser.showOpenMultipleDialog(stage);
        log.info("Selected Files {}", selected);
        if (selected == null)
            return;

        preferences.put(OPEN_DIRECTORY_DEFAULT, selected.get(0).getParentFile().getAbsolutePath());
        try {
            preferences.sync();
        } catch (BackingStoreException e) {
            e.printStackTrace();
        }

        convertFiles(selected);
    }

    private void convertFiles(List<File> files) {
        FileChooser fileChooser = new FileChooser();
        fileChooser.setTitle("Save converted files");
        if (files.size() > 1) {
            fileChooser.getExtensionFilters().addAll(new FileChooser.ExtensionFilter("Excel format", "*.xlsx", "*.xls"));
        } else {
            fileChooser.getExtensionFilters().addAll(
                    new FileChooser.ExtensionFilter("Excel Format", "*.xlsx", "*.xls"),
                    new FileChooser.ExtensionFilter("Classical Excel Format", "*.xls"),
                    new FileChooser.ExtensionFilter("CSV", "*.csv"),
                    new FileChooser.ExtensionFilter("Tab delimited text", "*.txt"),
                    new FileChooser.ExtensionFilter("All Supported Files", "*.csv", "*.txt", "*.xls", "*.xlsx"));
        }

        fileChooser.setInitialDirectory(new File(preferences.get(SAVE_DIRECTORY_DEFAULT, System.getProperty("user.home"))));

        File saveFile = fileChooser.showSaveDialog(stage);
        if (saveFile == null)
            return;

        preferences.put(SAVE_DIRECTORY_DEFAULT, saveFile.getParentFile().getAbsolutePath());
        try {
            preferences.sync();
        } catch (BackingStoreException e) {
            e.printStackTrace();
        }

        convertFilesToFile(files, saveFile);
    }

    private void convertFilesToFile(List<File> inputFiles, File outputFile) {
        try {
            Converter.builder().prettyTable(prettyTable.isSelected()).convertCellTypes(convertType.isSelected()).build().doConvert(inputFiles, outputFile);
        } catch (Exception e) {
            e.printStackTrace();
            Alert alert = new Alert(Alert.AlertType.ERROR);
            alert.setHeaderText("Error");
            alert.setTitle("Excel2CSV Error");
            alert.setContentText(e.getLocalizedMessage());
            alert.showAndWait();
        }

    }

    @FXML
    public void onAbout() {
        Alert alert = new Alert(Alert.AlertType.INFORMATION,
                "Excel2CSV\nVersion: "+ VersionResolver.getVersion() +  "\n" +
                        "Git Commit: " + VersionResolver.getGitCommit() + "\n" +
                        "Build Date: " + VersionResolver.getBuildDate() + "\n\n" +
                        "Webpage: https://github.com/informationsea/excel2csv"
        );
        alert.setTitle("About Excel2CSV");
        alert.setHeaderText("Excel2CSV");
        alert.showAndWait();
    }

    @FXML
    public void onQuit() {
        stage.close();
    }

    @FXML
    public void onDragOver(DragEvent event) {
        Dragboard db = event.getDragboard();
        if (db.hasFiles()) {
            for (File one : db.getFiles()) {
                if (Utilities.suggestFileTypeFromName(one.getName()) != Utilities.FileType.FILETYPE_UNKNOWN) {
                    event.acceptTransferModes(TransferMode.COPY);
                    break;
                }
            }
        }
        event.consume();
    }

    @FXML
    public void onDragDropped(DragEvent event) {
        Dragboard db = event.getDragboard();
        boolean success = false;
        if (db.hasFiles()) {
            //log.info("File dropped {}", db.getFiles());
            success = true;
        }
        event.setDropCompleted(success);
        event.consume();

        if (success) {
            convertFiles(db.getFiles());
        }
    }
}

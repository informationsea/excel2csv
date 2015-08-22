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

import javafx.fxml.FXMLLoader;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.stage.Stage;
import lombok.extern.slf4j.Slf4j;

import java.util.prefs.Preferences;

@Slf4j
public class Application extends javafx.application.Application {
    @Override
    public void start(Stage primaryStage) throws Exception {
        FXMLLoader loader = new FXMLLoader(getClass().getResource("MainWindow.fxml"));
        Parent parent = (Parent)loader.load();
        MainWindowController controller = loader.getController();
        controller.setStage(primaryStage);
        Scene scene = new Scene(parent);
        primaryStage.setScene(scene);
        primaryStage.setTitle("Excel2csv");
        primaryStage.setMinWidth(326);
        primaryStage.setMinHeight(243);
        primaryStage.show();
    }

    private static Preferences applicationProperties = null;

    public static Preferences getApplicationPreferences() {
        if (applicationProperties == null) {
            applicationProperties = Preferences.userNodeForPackage(Application.class);
        }
        log.info("Preferences {}", applicationProperties);
        return applicationProperties;
    }
}

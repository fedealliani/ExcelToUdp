/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package exceltoudp;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.net.DatagramPacket;
import java.net.DatagramSocket;
import java.net.InetAddress;
import java.net.SocketException;
import java.net.UnknownHostException;
import java.util.logging.Level;
import java.util.logging.Logger;
import javafx.application.Application;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.event.ActionEvent;
import javafx.event.EventHandler;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.ProgressBar;
import javafx.scene.control.Spinner;
import javafx.scene.control.SpinnerValueFactory;
import javafx.scene.control.TextField;
import javafx.scene.layout.GridPane;
import javafx.scene.text.Font;
import javafx.scene.text.FontWeight;
import javafx.scene.text.Text;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

/**
 *
 * @author Federico
 */
public class ExcelToUdp extends Application {

    private File file = null;
    private String ip = null;
    private int port = -1;
    private Button btn;
    private Button send;
    private Button cancel;
    private ProgressBar pb;
    private Thread one;
    private Label tiempoRestante;
    private boolean cancelar = false;
    private Spinner<String> spinner;

    @Override
    public void start(Stage primaryStage) throws SocketException {
        DatagramSocket udpSock;
        udpSock = new DatagramSocket(50122);
        GridPane grid = new GridPane();
        grid.setAlignment(Pos.CENTER);
        grid.setHgap(20);
        grid.setVgap(20);
        grid.setPadding(new Insets(25, 25, 25, 25));
        Scene scene = new Scene(grid, 640, 480);
        Text scenetitle = new Text("Excel to UDP");
        scenetitle.setFont(Font.font("Tahoma", FontWeight.NORMAL, 20));
        grid.add(scenetitle, 0, 0, 2, 1);

        Label label = new Label("Select Level:");
        spinner = new Spinner<String>();
        tiempoRestante = new Label("Tiempo Restante");
        ObservableList<String> abecedario = FXCollections.observableArrayList(//
                "Z", "Y", "X", "W", //
                "V", "U", "T", "S", //
                "R", "Q", "P", "O",
                "N", "M", "L", "K", "J",
                "I", "H", "G", "F", "E",
                "D", "C", "B", "A");
        SpinnerValueFactory<String> valueFactory = new SpinnerValueFactory.ListSpinnerValueFactory<String>(abecedario);
        valueFactory.setValue("A");
        spinner.setValueFactory(valueFactory);
        cancel = new Button();
        cancel.setText("Cancelar");
        cancel.setVisible(false);
        cancel.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent event) {
                cancelar = true;
            }
        });
        grid.add(cancel, 1, 6);
        btn = new Button();
        Label fileSelected = new Label("");
        btn.setText("Select xls file");
        btn.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent event) {
                FileChooser fileChooser = new FileChooser();
                FileChooser.ExtensionFilter extFilter = new FileChooser.ExtensionFilter("XLS files (*.xls)", "*.xls");
                fileChooser.getExtensionFilters().add(extFilter);
                fileChooser.setTitle("Open XLS File");
                file = fileChooser.showOpenDialog(primaryStage);
                if (file != null) {
                    String name = file.getName();
                    if (name.contains(".xls")) {
                        String FilePath = file.getAbsolutePath();
                        fileSelected.setText(FilePath);
                    } else {
                        file = null;
                        Alert alert = new Alert(AlertType.INFORMATION);
                        alert.setTitle("Information Dialog");
                        alert.setHeaderText("Error");
                        alert.setContentText("Error file is not .xls");
                        alert.showAndWait();
                    }
                }
            }
        });
        grid.add(btn, 0, 1);

        grid.add(fileSelected, 1, 1);
        Label spinnerLabel = new Label("Select column");
        grid.add(spinnerLabel, 0, 2);
        grid.add(spinner, 1, 2);
        Label ipLabel = new Label("Ip Server");
        grid.add(ipLabel, 0, 3);

        TextField ipTextField = new TextField();
        grid.add(ipTextField, 1, 3);

        Label portLabel = new Label("Port Server");
        grid.add(portLabel, 0, 4);

        TextField portTextField = new TextField();

        grid.add(portTextField, 1, 4);
        pb = new ProgressBar(0);

        grid.add(pb, 1, 7);
        send = new Button();

        Label delayLabel = new Label("Delay packet(Sec)");
        Spinner<Integer> delay = new Spinner<Integer>();

        final int initialValue = 1;

        SpinnerValueFactory<Integer> valueFactoryDelay =   new SpinnerValueFactory.IntegerSpinnerValueFactory(1, 3600, initialValue);

        delay.setValueFactory(valueFactoryDelay);
        grid.add(delayLabel, 0, 5);
        grid.add(delay, 1, 5);
        send.setText("Send");
        send.setOnAction(new EventHandler<ActionEvent>() {

            @Override
            public void handle(ActionEvent event) {

                if (file != null) {
                    String name = file.getName();
                    if (name.contains(".xls")) {
                        FileInputStream fs = null;
                        try {
                            String FilePath = file.getAbsolutePath();
                            fileSelected.setText(FilePath);
                            fs = new FileInputStream(FilePath);
                            Workbook wb = Workbook.getWorkbook(fs);
                            //Accedo a la hoja nro 0
                            Sheet sh = wb.getSheet(0);
                            // To get the number of rows present in sheet
                            int totalNoOfRows = sh.getRows();

                            // To get the number of columns present in sheet
                            int totalNoOfCols = sh.getColumns();
                            int columna = convertLetterToNumber(spinner.getValue());
                            if (totalNoOfCols - 1 < columna) {
                                Alert alert = new Alert(AlertType.INFORMATION);
                                alert.setTitle("Information Dialog");
                                alert.setHeaderText("Error");
                                alert.setContentText("Column empty.");
                                alert.showAndWait();
                                return;
                            }
                            if (ipTextField.getText().contains(".")) {
                                ip = ipTextField.getText();
                            } else {
                                ip = null;
                            }
                            if (portTextField.getText().length() > 0 && portTextField.getText().length() <= 5) {
                                int i;
                                boolean resFor = true;
                                for (i = 0; i < portTextField.getText().length(); i++) {
                                    char charAt = portTextField.getText().charAt(i);
                                    boolean isDigit = Character.isDigit(charAt);
                                    if (!isDigit) {
                                        resFor = false;
                                    }
                                }
                                if (resFor) {
                                    port = Integer.valueOf(portTextField.getText());
                                } else {
                                    port = -1;
                                }
                            } else {
                                port = -1;
                            }
                            int delayTime= delay.getValue();
                            if (ip != null && ip.length() > 2) {
                                if (port != -1 && port > 0 && port < 65535) {
                                    one = new Thread() {
                                        public void run() {
                                            try {
                                                send.setVisible(false);
                                                cancel.setVisible(true);
                                                //tiempoRestante.setVisible(true);
                                                for (int row = 0; row < totalNoOfRows; row++) {
                                                    if (cancelar) {
                                                        row = totalNoOfRows;
                                                        cancelar = false;
                                                    } else {
                                                        //int segundosFaltantes = totalNoOfRows * 2 - row;
                                                        //changeTiempoRestante("Time to finish:" + segundosFaltantes + " Seg");
                                                        if (!sh.getCell(convertLetterToNumber(spinner.getValue()), row).getContents().isEmpty()) {
                                                            System.out.print(sh.getCell(columna, row).getContents() + "\t");
                                                            System.out.println();
                                                            DatagramPacket r = new DatagramPacket(sh.getCell(columna, row).getContents().getBytes(), sh.getCell(columna, row).getContents().getBytes().length);
                                                            InetAddress rmh = InetAddress.getByName(ip);
                                                            r.setAddress(rmh);
                                                            r.setPort(port);
                                                            udpSock.send(r);
                                                        }
                                                        double progress = (double) ((double) (row) / (double) (totalNoOfRows));
                                                        pb.setProgress(progress);
                                                        Thread.sleep(delayTime*1000);
                                                    }
                                                }
                                                //tiempoRestante.setVisible(false);
                                                cancel.setVisible(false);
                                                send.setVisible(true);
                                                pb.setProgress(0);
                                            } catch (InterruptedException v) {
                                                System.out.println(v);
                                            } catch (UnknownHostException ex) {
                                                Logger.getLogger(ExcelToUdp.class.getName()).log(Level.SEVERE, null, ex);
                                            } catch (IOException ex) {
                                                Logger.getLogger(ExcelToUdp.class.getName()).log(Level.SEVERE, null, ex);
                                            }
                                        }

                                    };
                                    one.start();
                                } else {
                                    Alert alert = new Alert(AlertType.INFORMATION);
                                    alert.setTitle("Information Dialog");
                                    alert.setHeaderText("Error");
                                    alert.setContentText("Port invalid.");
                                    alert.showAndWait();
                                }
                            } else {
                                Alert alert = new Alert(AlertType.INFORMATION);
                                alert.setTitle("Information Dialog");
                                alert.setHeaderText("Error");
                                alert.setContentText("IP invalid.");
                                alert.showAndWait();
                            }

                        } catch (FileNotFoundException ex) {
                            Logger.getLogger(ExcelToUdp.class.getName()).log(Level.SEVERE, null, ex);
                        } catch (IOException | BiffException ex) {
                            Logger.getLogger(ExcelToUdp.class.getName()).log(Level.SEVERE, null, ex);
                        } finally {
                            try {
                                fs.close();
                            } catch (IOException ex) {
                                Logger.getLogger(ExcelToUdp.class.getName()).log(Level.SEVERE, null, ex);
                            }
                        }
                    } else {
                        Alert alert = new Alert(AlertType.INFORMATION);
                        alert.setTitle("Information Dialog");
                        alert.setHeaderText("Error");
                        alert.setContentText("Select a vaild file.");
                        alert.showAndWait();
                    }
                } else {
                    Alert alert = new Alert(AlertType.INFORMATION);
                    alert.setTitle("Information Dialog");
                    alert.setHeaderText("Error");
                    alert.setContentText("Select a vaild file.");
                    alert.showAndWait();
                }
            }
        });
        grid.add(send, 1, 6);

        tiempoRestante.setVisible(false);
        grid.add(tiempoRestante, 1, 7);
        primaryStage.setTitle("ExcelToUdp");
        primaryStage.setScene(scene);
        primaryStage.show();
    }

    private void changeTiempoRestante(final String text) {

        tiempoRestante.setText(text);

    }

    private int convertLetterToNumber(String value) {
        switch (value) {
            case "A":
                return 0;
            case "B":
                return 1;
            case "C":
                return 2;
            case "D":
                return 3;
            case "E":
                return 4;
            case "F":
                return 5;
            case "G":
                return 6;
            case "H":
                return 7;
            case "I":
                return 8;
            case "J":
                return 9;
            case "K":
                return 10;
            case "L":
                return 11;
            case "M":
                return 12;
            case "N":
                return 13;
            case "O":
                return 14;
            case "P":
                return 15;
            case "Q":
                return 16;
            case "R":
                return 17;
            case "S":
                return 18;
            case "T":
                return 19;
            case "U":
                return 20;
            case "V":
                return 21;
            case "W":
                return 22;
            case "X":
                return 23;
            case "Y":
                return 24;
            case "Z":
                return 25;
        }
        return 0;
    }

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        launch(args);
    }

}

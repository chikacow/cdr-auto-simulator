package vn.mercury.cdr;

import javafx.application.Application;
import javafx.application.Platform;
import javafx.beans.property.BooleanProperty;
import javafx.beans.property.SimpleBooleanProperty;
import javafx.beans.property.SimpleStringProperty;
import javafx.beans.property.StringProperty;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.concurrent.Task;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Node;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.control.cell.CheckBoxTableCell;
import javafx.scene.layout.*;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Instant;
import java.time.ZoneOffset;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/** Desktop tool for sending the CDR cURLs stored in an Excel file. */
public final class CdrSimulatorApp extends Application {
    private static final String RETEST_MERCURY = "retest mercury";
    private static final DateTimeFormatter ISO_MILLIS = DateTimeFormatter
            .ofPattern("yyyy-MM-dd'T'HH:mm:ss.SSS'Z'")
            .withZone(ZoneOffset.UTC);
    private static final Pattern ENDPOINT = Pattern.compile("/charge/offline/(data|session|event)(?=['\\\"\\s\\\\]|$)", Pattern.CASE_INSENSITIVE);
    private static final Pattern CUSTOMER_SUBSCRIPTION = Pattern.compile("/customers/([^/]+)/subscriptions/([^/]+)/charge/offline/(data|session|event)", Pattern.CASE_INSENSITIVE);
    private static final Pattern ACCESS_KEY = Pattern.compile("(?i)(--header\\s+['\\\"]\\s*accessKey\\s*:\\s*)[^'\\\"]*(['\\\"])");
    private static final Pattern ACCESS_KEY_VALUE = Pattern.compile("(?i)--header\\s+['\\\"]\\s*accessKey\\s*:\\s*([^'\\\"]*)['\\\"]");
    private static final Pattern JSON_STRING = Pattern.compile("(?m)(\\\"%s\\\"\\s*:\\s*\\\")[^\\\"]*(\\\")");
    private static final Pattern RESPONSE_DETAIL_OK = Pattern.compile("\\\"responseDetail\\\"\\s*:\\s*\\\"OK\\\"");
    private static final AtomicInteger SESSION_SEQUENCE = new AtomicInteger();

    private final ObservableList<CdrRow> rows = FXCollections.observableArrayList();
    private final TextField clientId = new TextField();
    private final PasswordField accessKey = new PasswordField();
    private final Map<CdrType, IdFields> ids = new EnumMap<>(CdrType.class);
    private final TextArea logArea = new TextArea();
    private final Label fileLabel = new Label("No file selected");
    private final Label resultLabel = new Label("0 CDR");
    private TableView<CdrRow> table;
    private Stage stage;

    @Override
    public void start(Stage primaryStage) {
        this.stage = primaryStage;
        primaryStage.setTitle("Mercury CDR Simulator");

        BorderPane root = new BorderPane();
        root.setPadding(new Insets(14));
        root.setTop(buildTopBar());
        root.setCenter(buildCenter());
        root.setBottom(buildLog());
        BorderPane.setMargin(root.getCenter(), new Insets(12, 0, 10, 0));

        Scene scene = new Scene(root, 1320, 820);
        primaryStage.setScene(scene);
        primaryStage.show();
    }

    private Node buildTopBar() {
        Button importButton = new Button("Import Excel");
        importButton.setOnAction(event -> importExcel());
        Button templateButton = new Button("Create template");
        templateButton.setOnAction(event -> createTemplate());
        Button prepareButton = new Button("Prepare cURL and SQL");
        prepareButton.setOnAction(event -> prepareRows());
        Button simulateButton = new Button("Simulate selected");
        simulateButton.setOnAction(event -> simulateSelected());
        Button selectAllButton = new Button("Select all");
        selectAllButton.setOnAction(event -> rows.forEach(row -> row.selected.set(row.hasCurl())));
        Button unselectAllButton = new Button("Unselect all");
        unselectAllButton.setOnAction(event -> rows.forEach(row -> row.selected.set(false)));
        Button exportButton = new Button("Export Excel results");
        exportButton.setOnAction(event -> exportResults());

        HBox buttons = new HBox(8, importButton, templateButton, prepareButton, simulateButton, selectAllButton, unselectAllButton, exportButton);
        buttons.setAlignment(Pos.CENTER_LEFT);
        VBox box = new VBox(7,
                new Label("Mercury CDR Simulator"),
                buttons,
                new HBox(12, new Label("File:"), fileLabel, new Region(), resultLabel));
        HBox.setHgrow(box.getChildren().get(2), Priority.ALWAYS);
        return box;
    }

    private Node buildCenter() {
        table = new TableView<>(rows);
        table.setEditable(true);
        table.getColumns().add(selectColumn());
        table.getColumns().add(textColumn("Usage", "usage", 260));
        table.getColumns().add(textColumn("Note", "note", 220));
        table.getColumns().add(useToolConfigColumn());
        table.getColumns().add(textColumn("Type", "type", 90));
        table.getColumns().add(textColumn("Session ID", "sessionId", 240));
        table.getColumns().add(textColumn("Status", "status", 160));
        table.getColumns().add(resultColumn());
        table.getColumns().add(textColumn("Request", "request", 360));
        table.getColumns().add(textColumn("Response", "response", 360));
        table.setPlaceholder(new Label("Import an Excel file with the Usage and Curl columns."));

        VBox configuration = new VBox(10, buildClientConfiguration(), buildIdConfiguration(), buildGuide());
        configuration.setPadding(new Insets(0, 0, 0, 12));
        configuration.setPrefWidth(390);
        SplitPane splitPane = new SplitPane(table, new ScrollPane(configuration));
        splitPane.setDividerPositions(0.68);
        return splitPane;
    }

    private Node buildClientConfiguration() {
        GridPane grid = new GridPane();
        grid.setHgap(8);
        grid.setVgap(8);
        grid.add(new Label("Client ID"), 0, 0);
        grid.add(clientId, 1, 0);
        grid.add(new Label("Access key"), 0, 1);
        grid.add(accessKey, 1, 1);
        clientId.setPromptText("Blank = any non-1130 client");
        accessKey.setPromptText("Blank = keep cURL value");
        GridPane.setHgrow(clientId, Priority.ALWAYS);
        GridPane.setHgrow(accessKey, Priority.ALWAYS);
        return titled("Test client", grid);
    }

    private Node buildIdConfiguration() {
        GridPane grid = new GridPane();
        grid.setHgap(8);
        grid.setVgap(8);
        grid.addRow(0, new Label("Type"), new Label("Customer ID"), new Label("Subscription ID"));
        int row = 1;
        for (CdrType type : CdrType.values()) {
            IdFields fields = new IdFields();
            fields.customerId.setPromptText("Blank = keep cURL value");
            fields.subscriptionId.setPromptText("Blank = keep cURL value");
            ids.put(type, fields);
            grid.add(new Label(type.display), 0, row);
            grid.add(fields.customerId, 1, row);
            grid.add(fields.subscriptionId, 2, row);
            GridPane.setHgrow(fields.customerId, Priority.ALWAYS);
            GridPane.setHgrow(fields.subscriptionId, Priority.ALWAYS);
            row++;
        }
        return titled("IDs by cURL type", grid);
    }

    private Node buildGuide() {
        Label guide = new Label("Input format: the first worksheet must have Usage and Curl as its first two headers; Note is an optional third header. "
                + "You may import the reference file directly: the tool reads only the 'retest mercury' sheet, Usage in column C and Curl in column E.\n\n"
                + "When preparing cURLs, the tool replaces Customer ID, Subscription ID, accessKey, current UTC timestamps and the unique Session ID. "
                + "Leave Customer ID, Subscription ID or Access key blank to keep the value already in each cURL. "
                + "Use tool config is selected by default for every Usage; clear it to keep all IDs and Access key from that Usage's original cURL. "
                + "Rows with a Usage but no cURL are marked Missing cURL and cannot be selected. "
                + "Each Prepare or Simulate action resets the prior run's status, result, request, response, Session ID and SQL checks. "
                + "Leave Client ID blank to check for any non-1130 client and client 1130.");
        guide.setWrapText(true);
        return titled("Instructions", guide);
    }

    private Node buildLog() {
        logArea.setEditable(false);
        logArea.setPrefRowCount(5);
        logArea.setWrapText(true);
        TitledPane pane = new TitledPane("Log", logArea);
        pane.setCollapsible(false);
        return pane;
    }

    private TitledPane titled(String title, Node content) {
        TitledPane pane = new TitledPane(title, content);
        pane.setCollapsible(false);
        return pane;
    }

    private TableColumn<CdrRow, Boolean> selectColumn() {
        TableColumn<CdrRow, Boolean> column = new TableColumn<>("Select");
        column.setCellValueFactory(cell -> cell.getValue().selectedProperty());
        column.setCellFactory(ignored -> new CheckBoxTableCell<>() {
            @Override
            public void updateItem(Boolean item, boolean empty) {
                super.updateItem(item, empty);
                setDisable(!empty && getIndex() >= 0 && getIndex() < getTableView().getItems().size()
                        && !getTableView().getItems().get(getIndex()).hasCurl());
            }
        });
        column.setPrefWidth(60);
        return column;
    }

    private TableColumn<CdrRow, Boolean> useToolConfigColumn() {
        TableColumn<CdrRow, Boolean> column = new TableColumn<>("Use tool config");
        column.setCellValueFactory(cell -> cell.getValue().useToolConfigProperty());
        column.setCellFactory(ignored -> new CheckBoxTableCell<>() {
            @Override
            public void updateItem(Boolean item, boolean empty) {
                super.updateItem(item, empty);
                setDisable(!empty && getIndex() >= 0 && getIndex() < getTableView().getItems().size()
                        && !getTableView().getItems().get(getIndex()).hasCurl());
            }
        });
        column.setPrefWidth(125);
        return column;
    }

    private TableColumn<CdrRow, String> textColumn(String title, String property, int width) {
        TableColumn<CdrRow, String> column = new TableColumn<>(title);
        column.setCellValueFactory(cell -> switch (property) {
            case "usage" -> cell.getValue().usageProperty();
            case "note" -> cell.getValue().noteProperty();
            case "type" -> cell.getValue().typeProperty();
            case "sessionId" -> cell.getValue().sessionIdProperty();
            case "status" -> cell.getValue().statusProperty();
            case "request" -> cell.getValue().requestProperty();
            default -> cell.getValue().responseProperty();
        });
        column.setPrefWidth(width);
        return column;
    }

    private TableColumn<CdrRow, String> resultColumn() {
        TableColumn<CdrRow, String> column = new TableColumn<>("Result");
        column.setCellValueFactory(cell -> cell.getValue().resultProperty());
        column.setCellFactory(ignored -> new TableCell<>() {
            @Override
            protected void updateItem(String item, boolean empty) {
                super.updateItem(item, empty);
                if (empty || item == null) {
                    setText(null);
                    setStyle("");
                } else {
                    setText(item);
                    setStyle(switch (item) {
                        case "Success" -> "-fx-background-color: #C6EFCE; -fx-text-fill: #006100; -fx-font-weight: bold;";
                        case "Failed" -> "-fx-background-color: #FFC7CE; -fx-text-fill: #9C0006; -fx-font-weight: bold;";
                        default -> "";
                    });
                }
            }
        });
        column.setPrefWidth(110);
        return column;
    }

    private void importExcel() {
        FileChooser chooser = xlsxChooser("Select source Excel file");
        File file = chooser.showOpenDialog(stage);
        if (file == null) return;
        try {
            List<CdrRow> imported = readRows(file.toPath());
            rows.setAll(imported);
            fileLabel.setText(file.getName());
            resultLabel.setText(imported.size() + " CDRs imported");
            log("Imported " + imported.size() + " CDRs from " + file.getName());
        } catch (Exception exception) {
            showError("Unable to import Excel", exception.getMessage());
        }
    }

    private List<CdrRow> readRows(Path source) throws IOException {
        try (InputStream input = Files.newInputStream(source); Workbook workbook = WorkbookFactory.create(input)) {
            Sheet mercury = workbook.getSheet(RETEST_MERCURY);
            if (mercury != null) return rowsFromRetestMercury(mercury);
            return rowsFromTwoColumnSheet(workbook.getSheetAt(0));
        }
    }

    private List<CdrRow> rowsFromRetestMercury(Sheet sheet) {
        List<CdrRow> imported = new ArrayList<>();
        DataFormatter formatter = new DataFormatter();
        for (Row row : sheet) {
            String usage = value(formatter, row.getCell(2));
            String curl = value(formatter, row.getCell(4));
            String note = value(formatter, row.getCell(5));
            if (!usage.isBlank()) imported.add(new CdrRow(usage, curl, note));
        }
        return imported;
    }

    private List<CdrRow> rowsFromTwoColumnSheet(Sheet sheet) throws IOException {
        DataFormatter formatter = new DataFormatter();
        Row header = sheet.getRow(sheet.getFirstRowNum());
        if (header == null || !"usage".equalsIgnoreCase(value(formatter, header.getCell(0)).trim())
                || !"curl".equalsIgnoreCase(value(formatter, header.getCell(1)).trim())) {
            throw new IOException("The file must start with the Usage and Curl columns, or contain a 'retest mercury' sheet.");
        }
        String thirdHeader = value(formatter, header.getCell(2)).trim();
        if (!thirdHeader.isBlank() && !"note".equalsIgnoreCase(thirdHeader)) {
            throw new IOException("The optional third column must be named Note.");
        }
        List<CdrRow> imported = new ArrayList<>();
        for (int index = header.getRowNum() + 1; index <= sheet.getLastRowNum(); index++) {
            Row row = sheet.getRow(index);
            if (row == null) continue;
            String usage = value(formatter, row.getCell(0));
            String curl = value(formatter, row.getCell(1));
            String note = value(formatter, row.getCell(2));
            if (!usage.isBlank()) {
                imported.add(new CdrRow(usage, curl, note));
            } else if (!curl.isBlank()) {
                throw new IOException("Row " + (index + 1) + " has a Curl but no Usage.");
            }
        }
        return imported;
    }

    private String value(DataFormatter formatter, org.apache.poi.ss.usermodel.Cell cell) {
        return cell == null ? "" : formatter.formatCellValue(cell).trim();
    }

    private void createTemplate() {
        FileChooser chooser = xlsxChooser("Save input template");
        chooser.setInitialFileName("cdr-input-template.xlsx");
        File file = chooser.showSaveDialog(stage);
        if (file == null) return;
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("CDR Input");
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("Usage");
            header.createCell(1).setCellValue("Curl");
            header.createCell(2).setCellValue("Note");
            sheet.setColumnWidth(0, 42 * 256);
            sheet.setColumnWidth(1, 110 * 256);
            sheet.setColumnWidth(2, 42 * 256);
            try (OutputStream output = Files.newOutputStream(file.toPath())) { workbook.write(output); }
            log("Created template: " + file.getName());
        } catch (Exception exception) {
            showError("Unable to create template", exception.getMessage());
        }
    }

    private void prepareRows() {
        try {
            assertCommonConfiguration();
            resetRunState();
            int prepared = 0;
            for (CdrRow row : rows.stream().filter(CdrRow::hasCurl).toList()) {
                prepare(row);
                prepared++;
            }
            resultLabel.setText(prepared + " CDRs ready");
            log("Applied configuration to " + prepared + " cURLs and generated check SQL.");
        } catch (IllegalArgumentException exception) {
            showError("Missing or invalid configuration", exception.getMessage());
        }
    }

    private void assertCommonConfiguration() {
        if (rows.isEmpty()) throw new IllegalArgumentException("No CDRs are available to process.");
    }

    private void prepare(CdrRow row) {
        CdrType type = CdrType.fromCurl(row.originalCurl);
        if (type == null) {
            row.status.set("Unrecognized data/session/event endpoint");
            throw new IllegalArgumentException("Usage '" + row.usage.get() + "' has no /data, /session or /event endpoint.");
        }
        IdFields groupIds = ids.get(type);
        CurlDefaults defaults = defaultsFrom(row.originalCurl);
        String customerId = row.useToolConfig.get()
                ? configuredOrDefault(groupIds.customerId.getText(), defaults.customerId)
                : defaults.customerId;
        String subscriptionId = row.useToolConfig.get()
                ? configuredOrDefault(groupIds.subscriptionId.getText(), defaults.subscriptionId)
                : defaults.subscriptionId;
        String configuredAccessKey = row.useToolConfig.get()
                ? configuredOrDefault(accessKey.getText(), defaults.accessKey)
                : defaults.accessKey;
        String now = ISO_MILLIS.format(Instant.now());
        String session = uniqueSessionId();
        String curl = rewriteCurl(row.originalCurl, type, customerId, subscriptionId, configuredAccessKey, now, session);
        row.type.set(type.display);
        row.sessionId.set(session);
        row.preparedCurl = curl;
        row.request.set(curl);
        row.sqlCheck = sqlFor(row.usage.get(), session, clientId.getText().trim());
        row.status.set("Ready");
        row.result.set("Not run");
        row.response.set("");
    }

    private CurlDefaults defaultsFrom(String curl) {
        Matcher pathMatcher = CUSTOMER_SUBSCRIPTION.matcher(curl);
        if (!pathMatcher.find()) throw new IllegalArgumentException("The cURL does not contain a customers/subscriptions URL.");
        Matcher keyMatcher = ACCESS_KEY_VALUE.matcher(curl);
        if (!keyMatcher.find() || keyMatcher.group(1).trim().isEmpty()) {
            throw new IllegalArgumentException("The cURL does not contain an accessKey header.");
        }
        return new CurlDefaults(pathMatcher.group(1), pathMatcher.group(2), keyMatcher.group(1).trim());
    }

    private String configuredOrDefault(String configuredValue, String defaultValue) {
        return configuredValue == null || configuredValue.isBlank() ? defaultValue : configuredValue.trim();
    }

    private String rewriteCurl(String original, CdrType type, String customer, String subscription, String key, String now, String session) {
        Matcher pathMatcher = CUSTOMER_SUBSCRIPTION.matcher(original);
        if (!pathMatcher.find()) throw new IllegalArgumentException("The cURL does not contain a customers/subscriptions URL.");
        String curl = pathMatcher.replaceFirst("/customers/" + Matcher.quoteReplacement(customer)
                + "/subscriptions/" + Matcher.quoteReplacement(subscription) + "/charge/offline/" + type.endpoint);
        Matcher keyMatcher = ACCESS_KEY.matcher(curl);
        if (!keyMatcher.find()) throw new IllegalArgumentException("The cURL does not contain an accessKey header.");
        curl = keyMatcher.replaceFirst("$1" + Matcher.quoteReplacement(key) + "$2");
        curl = replaceJsonField(curl, "sessionBeginTimeStamp", now);
        curl = replaceJsonField(curl, "startTime", now);
        curl = replaceJsonField(curl, "sessionId", session);
        return curl;
    }

    private String replaceJsonField(String curl, String key, String value) {
        Pattern pattern = Pattern.compile(String.format(JSON_STRING.pattern(), Pattern.quote(key)));
        Matcher matcher = pattern.matcher(curl);
        if (!matcher.find()) throw new IllegalArgumentException("The cURL does not contain the JSON field '" + key + "'.");
        return matcher.replaceFirst("$1" + Matcher.quoteReplacement(value) + "$2");
    }

    private String uniqueSessionId() {
        return "cdr-" + DateTimeFormatter.ofPattern("yyyyMMddHHmmssSSS").withZone(ZoneOffset.UTC).format(Instant.now())
                + "-" + SESSION_SEQUENCE.incrementAndGet() + "-" + UUID.randomUUID().toString().substring(0, 8);
    }

    private void simulateSelected() {
        try {
            assertCommonConfiguration();
            List<CdrRow> selected = rows.stream().filter(row -> row.hasCurl() && row.selected.get()).toList();
            if (selected.isEmpty()) throw new IllegalArgumentException("Select at least one CDR to simulate.");
            resetRunState();
            for (CdrRow row : selected) prepare(row);
            Alert confirmation = new Alert(Alert.AlertType.CONFIRMATION,
                    "The tool will send " + selected.size() + " cURLs to the environment in the file. Continue?",
                    ButtonType.CANCEL, ButtonType.OK);
            confirmation.setHeaderText("Confirm CDR simulation");
            if (confirmation.showAndWait().orElse(ButtonType.CANCEL) != ButtonType.OK) return;
            runSimulation(selected);
        } catch (IllegalArgumentException exception) {
            showError("Unable to simulate", exception.getMessage());
        }
    }

    private void resetRunState() {
        for (CdrRow row : rows) {
            row.sessionId.set("");
            row.preparedCurl = null;
            row.request.set("");
            row.response.set("");
            row.sqlCheck = "";
            if (row.hasCurl()) {
                row.status.set("Not run");
                row.result.set("Not run");
            } else {
                row.status.set("Missing cURL");
                row.result.set("Not available");
            }
        }
    }

    private void runSimulation(List<CdrRow> selected) {
        Task<Void> task = new Task<>() {
            @Override protected Void call() {
                for (CdrRow row : selected) {
                    Platform.runLater(() -> { row.status.set("Sending"); row.result.set("Checking"); });
                    try {
                        Process process = new ProcessBuilder("/bin/zsh", "-lc", row.preparedCurl).redirectErrorStream(true).start();
                        boolean complete = process.waitFor(45, TimeUnit.SECONDS);
                        String response = readLimited(process.getInputStream(), 4_000);
                        if (!complete) {
                            process.destroyForcibly();
                            Platform.runLater(() -> { row.status.set("Timed out"); row.result.set("Failed"); row.response.set(response); });
                        } else {
                            int exit = process.exitValue();
                            boolean success = exit == 0 && RESPONSE_DETAIL_OK.matcher(response).find();
                            Platform.runLater(() -> {
                                row.status.set(exit == 0 ? "Sent" : "cURL error (" + exit + ")");
                                row.result.set(success ? "Success" : "Failed");
                                row.response.set(response);
                            });
                        }
                    } catch (Exception exception) {
                        Platform.runLater(() -> { row.status.set("Send failed"); row.result.set("Failed"); row.response.set(exception.getMessage()); });
                    }
                }
                return null;
            }
        };
        task.setOnSucceeded(event -> log("Simulation completed for " + selected.size() + " CDRs. Export Excel to get the check SQL."));
        task.setOnFailed(event -> showError("Simulation error", task.getException().getMessage()));
        new Thread(task, "cdr-simulation").start();
    }

    private String readLimited(InputStream input, int limit) throws IOException {
        byte[] bytes = input.readAllBytes();
        String text = new String(bytes, StandardCharsets.UTF_8);
        return text.length() <= limit ? text : text.substring(0, limit) + " ...(truncated)";
    }

    private void exportResults() {
        if (rows.isEmpty()) { showError("No data", "Import Excel before exporting."); return; }
        FileChooser chooser = xlsxChooser("Export Excel results");
        chooser.setInitialFileName("cdr-simulation-result.xlsx");
        File file = chooser.showSaveDialog(stage);
        if (file == null) return;
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("CDR Results");
            String[] headers = {"Usage", "SQL Check", "Request", "Note", "Configuration Source", "Type", "Session ID", "Status", "Result", "Response"};
            Row header = sheet.createRow(0);
            CellStyle style = workbook.createCellStyle();
            Font font = workbook.createFont(); font.setBold(true); style.setFont(font);
            CellStyle successStyle = resultStyle(workbook, IndexedColors.LIGHT_GREEN, IndexedColors.DARK_GREEN);
            CellStyle failureStyle = resultStyle(workbook, IndexedColors.ROSE, IndexedColors.DARK_RED);
            for (int i = 0; i < headers.length; i++) { org.apache.poi.ss.usermodel.Cell cell = header.createCell(i); cell.setCellValue(headers[i]); cell.setCellStyle(style); }
            int index = 1;
            for (CdrRow row : rows) {
                Row target = sheet.createRow(index++);
                target.createCell(0).setCellValue(row.usage.get());
                target.createCell(1).setCellValue(row.sqlCheck);
                target.createCell(2).setCellValue(row.request.get());
                target.createCell(3).setCellValue(row.note.get());
                target.createCell(4).setCellValue(row.useToolConfig.get() ? "Tool configuration" : "Original cURL");
                target.createCell(5).setCellValue(row.type.get());
                target.createCell(6).setCellValue(row.sessionId.get());
                target.createCell(7).setCellValue(row.status.get());
                org.apache.poi.ss.usermodel.Cell resultCell = target.createCell(8);
                resultCell.setCellValue(row.result.get());
                if ("Success".equals(row.result.get())) resultCell.setCellStyle(successStyle);
                if ("Failed".equals(row.result.get())) resultCell.setCellStyle(failureStyle);
                target.createCell(9).setCellValue(row.response.get());
            }
            sheet.createFreezePane(0, 1);
            sheet.setColumnWidth(0, 40 * 256);
            sheet.setColumnWidth(1, 95 * 256);
            sheet.setColumnWidth(2, 120 * 256);
            sheet.setColumnWidth(3, 42 * 256);
            sheet.setColumnWidth(4, 24 * 256);
            sheet.setColumnWidth(5, 14 * 256);
            sheet.setColumnWidth(6, 38 * 256);
            sheet.setColumnWidth(7, 22 * 256);
            sheet.setColumnWidth(8, 16 * 256);
            sheet.setColumnWidth(9, 65 * 256);
            writeBatchCheckSheet(workbook, style);
            try (OutputStream output = Files.newOutputStream(file.toPath())) { workbook.write(output); }
            log("Exported results: " + file.getName());
        } catch (Exception exception) {
            showError("Unable to export Excel", exception.getMessage());
        }
    }

    private void writeBatchCheckSheet(Workbook workbook, CellStyle headerStyle) {
        List<CdrRow> preparedRows = rows.stream()
                .filter(row -> !row.sessionId.get().isBlank())
                .toList();
        Sheet batchSheet = workbook.createSheet("Batch Check");
        Row header = batchSheet.createRow(0);
        org.apache.poi.ss.usermodel.Cell headerCell = header.createCell(0);
        headerCell.setCellValue("Batch SQL Check");
        headerCell.setCellStyle(headerStyle);

        Row queryRow = batchSheet.createRow(1);
        queryRow.createCell(0).setCellValue(batchSqlFor(preparedRows, clientId.getText().trim()));
        queryRow.setHeightInPoints(720);
        CellStyle queryStyle = workbook.createCellStyle();
        queryStyle.setWrapText(true);
        queryStyle.setVerticalAlignment(VerticalAlignment.TOP);
        queryRow.getCell(0).setCellStyle(queryStyle);
        batchSheet.setColumnWidth(0, 150 * 256);
        batchSheet.createFreezePane(0, 1);
    }

    private String sqlFor(String usage, String session, String configuredClient) {
        String expectedClients = configuredClient.isBlank()
                ? "one client other than '1130' and client '1130'"
                : "dc.client = '" + sqlEscape(configuredClient) + "' and '1130'";
        return "-- Expected: 2 rows; " + expectedClients + "\n"
                + "select dc.id,\n       ds.status,\n       dc.ref_id,\n       ccir.session_id,\n       ccir.location_zone,\n       ccir.destination_zone,\n       ccir.location_network,\n       bu.usage,\n       dc.client\n"
                + "from d_cdrs dc\n"
                + "join d_services ds on dc.service = ds.id\n"
                + "join c_cdr_inew_raw ccir on ccir.id = dc.ref_id\n"
                + "left join d_usage du on dc.usage = du.id\n"
                + "left join b_usage bu on bu.id = du.usage\n"
                + "where dc.ref_id is not null\n"
                + "  and ccir.session_id = '" + sqlEscape(session) + "'\n"
                + "  and bu.usage = '" + sqlEscape(usage) + "'\n"
                + "order by dc.ref_id desc, dc.client;";
    }

    private String batchSqlFor(List<CdrRow> preparedRows, String configuredClient) {
        if (preparedRows.isEmpty()) {
            return "-- No CDR session IDs have been generated yet. Select Prepare cURL and SQL or simulate at least one CDR first.";
        }
        boolean hasConfiguredClient = !configuredClient.isBlank();
        String clientPredicate = hasConfiguredClient
                ? "client::text = '" + sqlEscape(configuredClient) + "'"
                : "client::text <> '1130'";
        String configuredClientLabel = hasConfiguredClient ? configuredClient : "ANY_NON_1130_CLIENT";
        String missingClientIssue = hasConfiguredClient ? "MISSING_CONFIGURED_CLIENT" : "MISSING_NON_1130_CLIENT";
        String missingBothIssue = hasConfiguredClient ? "MISSING_CONFIGURED_CLIENT_AND_1130" : "MISSING_NON_1130_CLIENT_AND_1130";
        String values = preparedRows.stream()
                .map(row -> "        ('" + sqlEscape(row.sessionId.get()) + "', '" + sqlEscape(row.usage.get()) + "')")
                .reduce((first, second) -> first + ",\n" + second)
                .orElseThrow();
        return "-- Returns only exceptions: no CDR created, required client missing, or client 1130 missing.\n"
                + "WITH expected_cdr (session_id, usage) AS (\n"
                + "    VALUES\n" + values + "\n"
                + "),\n"
                + "actual_cdr AS (\n"
                + "    SELECT e.session_id, e.usage, actual.id, actual.ref_id, actual.client\n"
                + "    FROM expected_cdr e\n"
                + "    LEFT JOIN c_cdr_inew_raw ccir\n"
                + "      ON ccir.session_id::text = e.session_id::text\n"
                + "    LEFT JOIN (\n"
                + "        SELECT dc.id, dc.ref_id, dc.client, bu.usage\n"
                + "        FROM aax2.d_cdrs dc\n"
                + "        LEFT JOIN d_usage du ON dc.usage = du.id\n"
                + "        LEFT JOIN b_usage bu ON bu.id = du.usage\n"
                + "    ) actual ON actual.ref_id = ccir.id\n"
                + "            AND actual.usage = e.usage\n"
                + ")\n"
                + "SELECT session_id,\n"
                + "       usage,\n"
                + "       '" + sqlEscape(configuredClientLabel) + "' AS configured_client,\n"
                + "       COALESCE(string_agg(DISTINCT ref_id::text, ', '), '') AS ref_ids,\n"
                + "       COUNT(id) AS cdr_count,\n"
                + "       COUNT(id) FILTER (WHERE " + clientPredicate + ") AS configured_client_cdr_count,\n"
                + "       COUNT(id) FILTER (WHERE client::text = '1130') AS client_1130_cdr_count,\n"
                + "       CASE\n"
                + "         WHEN COUNT(id) = 0 THEN 'NO_CDR_CREATED'\n"
                + "         WHEN COUNT(id) FILTER (WHERE " + clientPredicate + ") = 0\n"
                + "          AND COUNT(id) FILTER (WHERE client::text = '1130') = 0 THEN '" + missingBothIssue + "'\n"
                + "         WHEN COUNT(id) FILTER (WHERE " + clientPredicate + ") = 0 THEN '" + missingClientIssue + "'\n"
                + "         WHEN COUNT(id) FILTER (WHERE client::text = '1130') = 0 THEN 'MISSING_1130'\n"
                + "       END AS issue\n"
                + "FROM actual_cdr\n"
                + "GROUP BY session_id, usage\n"
                + "HAVING COUNT(id) = 0\n"
                + "    OR COUNT(id) FILTER (WHERE " + clientPredicate + ") = 0\n"
                + "    OR COUNT(id) FILTER (WHERE client::text = '1130') = 0\n"
                + "ORDER BY session_id, usage;";
    }

    private CellStyle resultStyle(Workbook workbook, IndexedColors fill, IndexedColors fontColor) {
        CellStyle style = workbook.createCellStyle();
        style.setFillForegroundColor(fill.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Font font = workbook.createFont();
        font.setBold(true);
        font.setColor(fontColor.getIndex());
        style.setFont(font);
        return style;
    }

    private String sqlEscape(String value) { return value.replace("'", "''"); }

    private FileChooser xlsxChooser(String title) {
        FileChooser chooser = new FileChooser();
        chooser.setTitle(title);
        chooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("Excel (*.xlsx)", "*.xlsx"));
        return chooser;
    }

    private void log(String message) { logArea.appendText(ISO_MILLIS.format(Instant.now()) + "  " + message + System.lineSeparator()); }

    private void showError(String title, String detail) {
        Alert alert = new Alert(Alert.AlertType.ERROR, detail == null ? "Unknown error" : detail, ButtonType.OK);
        alert.setHeaderText(title);
        alert.showAndWait();
    }

    private enum CdrType {
        DATA("data", "Data"), SESSION("session", "Session"), EVENT("event", "Event");
        private final String endpoint;
        private final String display;
        CdrType(String endpoint, String display) { this.endpoint = endpoint; this.display = display; }
        static CdrType fromCurl(String curl) {
            Matcher matcher = ENDPOINT.matcher(curl);
            if (!matcher.find()) return null;
            return switch (matcher.group(1).toLowerCase(Locale.ROOT)) { case "data" -> DATA; case "session" -> SESSION; default -> EVENT; };
        }
    }

    private static final class IdFields {
        private final TextField customerId = new TextField();
        private final TextField subscriptionId = new TextField();
    }

    private record CurlDefaults(String customerId, String subscriptionId, String accessKey) { }

    private static final class CdrRow {
        private final BooleanProperty selected = new SimpleBooleanProperty(true);
        private final BooleanProperty useToolConfig = new SimpleBooleanProperty(true);
        private final StringProperty usage = new SimpleStringProperty("");
        private final StringProperty note = new SimpleStringProperty("");
        private final StringProperty type = new SimpleStringProperty("");
        private final StringProperty sessionId = new SimpleStringProperty("");
        private final StringProperty status = new SimpleStringProperty("Not configured");
        private final StringProperty result = new SimpleStringProperty("Not run");
        private final StringProperty request = new SimpleStringProperty("");
        private final StringProperty response = new SimpleStringProperty("");
        private final String originalCurl;
        private final boolean hasCurl;
        private String preparedCurl;
        private String sqlCheck = "";
        private CdrRow(String usage, String originalCurl, String note) {
            this.usage.set(usage);
            this.originalCurl = originalCurl;
            this.note.set(note);
            this.hasCurl = !originalCurl.isBlank();
            this.selected.set(hasCurl);
            CdrType detected = CdrType.fromCurl(originalCurl);
            this.type.set(hasCurl && detected != null ? detected.display : "No cURL");
            if (!hasCurl) {
                this.status.set("Missing cURL");
                this.result.set("Not available");
            }
        }
        BooleanProperty selectedProperty() { return selected; }
        BooleanProperty useToolConfigProperty() { return useToolConfig; }
        boolean hasCurl() { return hasCurl; }
        StringProperty usageProperty() { return usage; }
        StringProperty noteProperty() { return note; }
        StringProperty typeProperty() { return type; }
        StringProperty sessionIdProperty() { return sessionId; }
        StringProperty statusProperty() { return status; }
        StringProperty resultProperty() { return result; }
        StringProperty requestProperty() { return request; }
        StringProperty responseProperty() { return response; }
    }

    public static void main(String[] args) { launch(args); }
}

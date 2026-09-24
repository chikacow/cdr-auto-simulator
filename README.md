# Mercury CDR Simulator (JavaFX)

Desktop application for preparing and sending Mercury CDR cURLs from Excel.

## Run the application

Java 17 and Maven are required.

```bash
cd /Users/kiencc/Downloads/cdr-simulator-fx
mvn javafx:run
```

On macOS, you can start `run-cdr-simulator.command` after granting execute permission once:

```bash
chmod +x /Users/kiencc/Downloads/cdr-simulator-fx/run-cdr-simulator.command
/Users/kiencc/Downloads/cdr-simulator-fx/run-cdr-simulator.command
```

Do not open `target/cdr-simulator-fx-1.0.0.jar` directly or run `java -jar ...`. That jar does not package the JavaFX native runtime and will report `JavaFX runtime components are missing`.

## Run and debug in IntelliJ

The project includes the **Run Mercury CDR Simulator** Application configuration in `.run`. After opening the project in IntelliJ and loading Maven dependencies, select that configuration in the top-right run selector and click Run or Debug. It compiles the code, runs `vn.mercury.cdr.CdrSimulatorApp`, and includes the JavaFX runtime for macOS Apple Silicon.

## Input Excel file

- Standard format: the first worksheet starts with `Usage` and `Curl`; `Note` is an optional third column immediately to the right of `Curl`.
- You may import `cdr die.xlsx` directly. The application reads only its `retest mercury` sheet: Usage in column C and Curl in column E. All other worksheets are ignored.

Rows with a Usage but no Curl stay visible as `Missing cURL`, but cannot be selected, prepared, or simulated.

Each cURL must contain `/charge/offline/data`, `/charge/offline/session`, or `/charge/offline/event`; an `accessKey` header; and the JSON fields `sessionBeginTimeStamp`, `startTime`, and `sessionId`.

## Workflow

1. Import the Excel file. Client ID is optional: enter it to check that exact client alongside `1130`; leave it blank to check for any non-`1130` client alongside `1130`.
2. Customer ID, Subscription ID, and Access key are optional. Leave a field blank to keep the value already present in each individual cURL; enter a value to override it for that CDR type. Notes are displayed next to Usage and included in exported results.
3. Every Usage is selected for **Use tool config** by default. Clear that checkbox for a specific Usage to preserve its original Customer ID, Subscription ID, and Access key, even when tool configuration is entered.
4. Select **Prepare cURL and SQL** to create configured cURLs and a unique Session ID for each CDR.
5. Select the required rows and click **Simulate selected**. The application asks for confirmation before sending requests.
6. A response containing exactly `"responseDetail": "OK"` is marked **Success** in green. Every other response is **Failed** in red.
7. Click **Export Excel results**. The `CDR Results` sheet provides one SQL query per CDR. The `Batch Check` sheet provides one query for all generated CDRs. It returns only exceptions: no CDR created, required client missing, or client `1130` missing.

Every Prepare or Simulate action starts a new run. It clears the prior Status, Result, Request, Response, Session ID, and SQL checks for all Usage rows before generating new values for the rows that are processed.

Timestamps use UTC in `yyyy-MM-dd'T'HH:mm:ss.SSSZ` format, for example `2026-09-23T09:21:49.622Z`. The Access key is only kept in memory while the application runs.
# cdr-auto-simulator

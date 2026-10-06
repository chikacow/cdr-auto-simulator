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

On Windows, install Java 17+ and Apache Maven, then double-click `run-cdr-simulator.bat`, or run this from Command Prompt in the project directory:

```bat
run-cdr-simulator.bat
```

On Windows, the application uses Git Bash when available. It writes each cURL to a temporary `.sh` file before running it, preserving the quoted JSON body. Without Git Bash, it parses the cURL and calls `curl.exe` directly.

Do not open `target/cdr-simulator-fx-1.0.0.jar` directly or run `java -jar ...`. That jar does not package the JavaFX native runtime and will report `JavaFX runtime components are missing`.

## Download a packaged desktop app

The **Package desktop applications** GitHub Actions workflow creates self-contained ZIP files for Windows x64, macOS Intel, and macOS Apple Silicon. The ZIP includes Java and JavaFX, so end users only unzip it and open the application; they do not need Java, Maven, Git, or a terminal.

To create packages, open the repository's **Actions** tab, select **Package desktop applications**, and choose **Run workflow**. Download the matching artifact after the workflow finishes. On the first launch of an unsigned macOS package, use Control-click, then choose **Open**.

## Run and debug in IntelliJ

The bundled **Run Mercury CDR Simulator** Application configuration is for macOS Apple Silicon. On Windows, run the Maven goal `javafx:run` from IntelliJ's Maven panel, or use `run-cdr-simulator.bat`.

## Input Excel file

- Standard format: the worksheet to import starts with `Usage` and `Curl`; `Note` is an optional third column immediately to the right of `Curl`.
- After selecting an Excel file, choose the worksheet to import. Other worksheets are not changed or read.

Rows with a Usage but no Curl stay visible as `Missing cURL`, but cannot be selected, prepared, or simulated.

Each cURL must contain `/charge/offline/data`, `/charge/offline/session`, or `/charge/offline/event`; an `accessKey` header; and the JSON fields `sessionBeginTimeStamp`, `startTime`, and `sessionId`.

## Workflow

1. Import the Excel file. Client ID is optional: enter it to check that exact client alongside `1130`; leave it blank to check for any non-`1130` client alongside `1130`.
2. Customer ID, Subscription ID, and Access key are optional. Leave a field blank to keep the value already present in each individual cURL; enter a value to override it for that CDR type. Notes are displayed next to Usage and included in exported results.
3. Every Usage is selected for **Use tool config** by default. Clear that checkbox for a specific Usage to preserve its original Customer ID, Subscription ID, and Access key, even when tool configuration is entered.
4. Choose a Session ID mode: **Auto generate** (default), **Use exact Session ID**, or **Auto generate with prefix**. An exact Session ID may be shared by multiple Usage rows.
5. Select **Prepare cURL and SQL** to create configured cURLs and a Session ID for each CDR.
6. Select the required rows and click **Simulate selected**. The application asks for confirmation before sending requests.
7. A response containing exactly `"responseDetail": "OK"` is marked **Success** in green. Every other response is **Failed** in red.
8. Click **Export Excel results**. The `CDR Results` sheet provides one SQL query per CDR. The `Batch Check` sheet provides one query for all generated CDRs. It returns only exceptions: no CDR created, required client missing, or client `1130` missing.

Every Prepare or Simulate action starts a new run. It clears the prior Status, Result, Request, Response, Session ID, and SQL checks for all Usage rows before generating new values for the rows that are processed.

Timestamps use UTC in `yyyy-MM-dd'T'HH:mm:ss.SSSZ` format, for example `2026-09-23T09:21:49.622Z`. The Access key is only kept in memory while the application runs.
# cdr-auto-simulator

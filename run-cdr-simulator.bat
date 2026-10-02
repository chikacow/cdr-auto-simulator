@echo off
rem Mercury CDR Simulator launcher for Windows
rem Starts through the JavaFX Maven plugin so JavaFX's native runtime is on the module path.
rem Requires Java 17+ and Apache Maven installed and in PATH.

cd /d "%~dp0"
mvn -q -f pom.xml javafx:run

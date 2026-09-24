#!/bin/zsh

# Start through the JavaFX Maven plugin so JavaFX's native macOS runtime is on
# the module path. Do not launch target/cdr-simulator-fx-1.0.0.jar directly.
set -euo pipefail

APP_DIR="${0:A:h}"
exec mvn -q -f "$APP_DIR/pom.xml" javafx:run

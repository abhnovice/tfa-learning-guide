#!/bin/zsh
# Starts TradeFlow Atlas. The Python server loads outputs/.env itself.
set -eu

APP_DIR="$(cd "$(dirname "$0")" && pwd)"
PYTHON="$APP_DIR/../work/.venv/bin/python"

if [[ ! -x "$PYTHON" ]]; then
  print "The TradeFlow virtual environment is missing. Run the setup again."
  exit 1
fi

: "${TRADEFLOW_PORT:=8816}"
export TRADEFLOW_PORT
exec "$PYTHON" -u "$APP_DIR/tradeflow_server.py"

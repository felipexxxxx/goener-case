#!/usr/bin/env sh
set -eu
cd "$(dirname "$0")"
python_bin="${PYTHON:-python3}"
"$python_bin" -c "import sys; assert (3,11) <= sys.version_info[:2] < (3,14), 'Use Python 3.11, 3.12 ou 3.13'"
"$python_bin" -m pip install -r requirements-lock.txt
exec "$python_bin" -m goener --config config.toml "$@"

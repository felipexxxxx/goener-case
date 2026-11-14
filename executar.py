"""Entrada simples para Python e para o executável Windows."""

import multiprocessing
import os
import sys
from pathlib import Path

from goener.pipeline import main

if __name__ == "__main__":
    multiprocessing.freeze_support()
    if getattr(sys, "frozen", False):
        os.chdir(Path(sys.executable).resolve().parent)
    raise SystemExit(main())

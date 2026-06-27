"""Enable ``python -m dividend_contribution``."""
import sys

from .cli import main

if __name__ == "__main__":
    sys.exit(main())

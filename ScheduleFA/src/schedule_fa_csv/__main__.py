"""Enable ``python -m schedule_fa_csv``."""
import sys

from .cli import main

if __name__ == "__main__":
    sys.exit(main())

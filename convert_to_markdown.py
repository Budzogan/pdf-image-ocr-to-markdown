import sys

from scan_to_markdown_docling import main


if __name__ == "__main__":
    try:
        sys.exit(main())
    except FileNotFoundError as exc:
        print(f"ERROR: {exc}")
        sys.exit(1)
    except KeyboardInterrupt:
        print("\nStopped by user (Ctrl+C).")
        sys.exit(130)
    except Exception as exc:
        print(f"ERROR: Unexpected failure: {exc}")
        sys.exit(1)

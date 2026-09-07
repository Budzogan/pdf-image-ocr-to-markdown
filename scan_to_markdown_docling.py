"""Thin entrypoint; full Docling converter lives in impl modules."""
from scan_to_markdown_docling_impl_b import *  # noqa: F403

if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except FileNotFoundError as exc:
        print(f"ERROR: {exc}")
        raise SystemExit(1)
    except KeyboardInterrupt:
        print("\nStopped by user (Ctrl+C).")
        raise SystemExit(130)
    except Exception as exc:
        print(f"ERROR: Unexpected failure: {exc}")
        raise SystemExit(1)

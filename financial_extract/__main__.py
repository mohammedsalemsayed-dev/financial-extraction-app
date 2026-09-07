"""
Entry point:  python -m financial_extract  [--port 8765] [--no-browser]
"""
import argparse

from .server import serve


def main():
    ap = argparse.ArgumentParser(prog="financial_extract",
                                 description="Local PDF -> Excel financial-table extraction app.")
    ap.add_argument("--port", type=int, default=8765)
    ap.add_argument("--no-browser", action="store_true",
                    help="do not open a web browser on start")
    args = ap.parse_args()
    serve(port=args.port, open_browser=not args.no_browser)


if __name__ == "__main__":
    main()

# HWP → TXT Converter

A compact Streamlit application that focuses on a single job: turning Hangul Word Processor (HWP/HWPX) files into clean UTF-8 text.
The converter is optimized for Streamlit Cloud – no external binaries, APIs, or heavyweight Python packages.

## How it works
- **Direct BodyText parsing** for classic HWP files using `olefile`, `struct`, and `zlib`.
- **XML extraction** for HWPX files via Python's built-in `zipfile` and `xml.etree.ElementTree`.
- **Clean UTF-8 output** in the UI with a simple download button—no additional validation steps.

## Running locally
```bash
pip install -e .
streamlit run app.py
```

## Why it's lightweight
- Only two third-party dependencies: `streamlit` and `olefile`.
- No temporary files, subprocesses, or API round-trips.
- The parsing functions stream through BodyText records and keep memory usage small, which keeps Streamlit Cloud happy.

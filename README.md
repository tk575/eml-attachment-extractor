# EML Attachment Extractor

Python script for extracting attachments from .eml files.

The script extracts all attachments of these formats: `docx`, `zip`, `pdf`, `rar`, `tar.gz`, `pptx`. It later gathers the `docx`, `pdf` and `pptx` inside the archives.

### Requirements

Before running the script, you need to have the following installed.

##### Python 3

It can be downloaded from their official website: https://www.python.org/downloads/.

##### Portable Archive Manager (patool)

It can be downloaded from pip3 using `pip3 install patool`. Alternatively, it can be built from source from https://pypi.python.org/pypi/patool.

### Running

For running the script, copy `run.py` into required folder contaning the `.eml` files and use `python3 run.py` in terminal. The output files will be stored in `output` folder.

### Customizing

The extensions to be looked for can be changed by modifying the `exts` list at the start and at the archive part.

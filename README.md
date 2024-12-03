# MetaPlaig

MetaPlaig is a Python application designed to extract files from various archive formats, analyze their metadata, and calculate hash values. It provides a graphical interface to browse and process files, displaying extracted metadata such as author information and hash values in a tree view.

Features

- **Archive Extraction**: Supports `.zip`, `.tar`, `.tar.gz`, `.tar.bz2`, `.rar`, and `.7z` formats.
- **Metadata Extraction**: Retrieves author information from `.docx` and `.pdf` files.
- **Hash Calculation**: Computes file hashes using the SHA-256 algorithm (default).
- **Duplicate Highlighting**: Highlights duplicate authors in the tree view.
- **Graphical User Interface**: Intuitive interface built with Tkinter.

## Requirements

To run MetaPlaig, ensure the following Python libraries are installed:

```
tkinter==builtin
ttk==builtin
zipfile==builtin
tarfile==builtin
os==builtin
hashlib==builtin
rarfile==4.0
py7zr==0.20.4
docx==0.2.4
python-docx==0.8.11
PyPDF2==3.0.1
```

## Installation

1. Clone the repository or copy the script to your local machine.
2. Install the required dependencies using pip:
   ```bash
   pip install rarfile py7zr python-docx PyPDF2
   ```

## Usage

1. Run the script:
   ```bash
   python metaplaig.py
   ```
2. In the GUI:
   - Click **Browse and Extract Files**.
   - Select a folder containing archive files.
   - Choose an extraction folder.
3. The application will extract files, process `.docx` and `.pdf` files for author metadata, calculate file hashes, and display the information in the tree view.

## Tree View Columns

- **Serial**: Index of the file in the processed list.
- **File Name**: Name of the processed file.
- **Author**: Extracted author metadata ("Unknown" if unavailable).
- **Hash**: SHA-256 hash of the file.

## Duplicate Detection

Authors with more than one file in the folder will be highlighted in red in the tree view.

## Known Issues

- Limited error handling for corrupted or unsupported files.
- Author extraction depends on metadata availability.

## License

MetaPlaig is released under the MIT License.

## Contributions

Feel free to fork the repository and submit pull requests for enhancements or bug fixes.

## Acknowledgements

This application leverages the following Python libraries:

- `Tkinter` for the GUI
- `zipfile`, `tarfile`, `rarfile`, and `py7zr` for archive extraction
- `python-docx` and `PyPDF2` for metadata processing
- `hashlib` for hash calculation


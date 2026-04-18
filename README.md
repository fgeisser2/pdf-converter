# PDF-Converter

Spawns a number of instances of libreoffice, converts all files with the selected extensions in a folder to pdf and places the pdf right next to the preexisting file.
All file-formats supported by libreoffice can be converted.


## Install

Download [pdf-converter.py](https://raw.githubusercontent.com/fgeisser2/pdf-converter/refs/heads/main/pdf-converter.py) and, under windows, install [Libreoffice](https://de.libreoffice.org/download/download/) (64bit)

A python-interpreter is required, but normally preinstalled under linux and bundled by libreoffice under windows.

## Usage

### Linux

```bash
python pdf-converter.py [-h] [--path PATH] [--threads (number of libreoffice instances)] [--logfile LOGFILE] [--extensions (comma-separated list of file extensions, default: doc,docx,rtf,txt)]
```


### Windows

```bat
C:\Program Files\LibreOffice\program\python.exe" .\pdf-converter.py [-h] [--path PATH] [--threads (number of libreoffice instances)] [--logfile LOGFILE] [--extensions (comma-separated list of file extensions, default: doc,docx,rtf,txt)]
```


## Development

`main.py` collects the files and launches the instances of libreoffice executing `macro.py`.
`build.py` bundles both into `pdf-converter.py` for easier deployment.


## Uninstall

Just remove the script

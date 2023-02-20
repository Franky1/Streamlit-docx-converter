<!-- markdownlint-disable MD033 -->
# Streamlit Docx Converter

[![Streamlit App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://franky1-streamlit-docx-converter-streamlit-app-n67qvo.streamlit.app/)

Streamlit App to convert **Microsoft Word** or **LibreOffice Writer** documents to PDF.

Supported file formats are: `docx`, `doc`, `odt`, `rtf`

*Disclaimer*: The conversion is done using a headless version of **LibreOffice**.
As we all know, there are some compatibility issues between LibreOffice and Microsoft Office. Therefore, the conversion might not be perfect and may fail in some cases.
Especially if you convert a Microsoft Word document and/or the document is more complex.

## Status

> Demo application is working - Last changed: 2023-02-20

## Description

Trying to evaluate the best way to convert docx to pdf in a headless environment. The goal is to use this in a docker container. It seems that **libreoffice** is currently the best option, but it is not easy to install in a docker container.

## Issues

None.

## ToDo

- [x] Make a streamlit demo app that can convert docx to pdf
- [x] Add session based temporary file storage
- [x] Make a wider page layout
- [x] Add some app styling with css
- [x] Test it on streamlit cloud
- [x] Accumulate multiple files in a temporary folder
- [x] Allow to download multiple files as zip
- [x] Add button for manual cleanup of temporary files
- [x] Allow to upload multiple files at once
- [x] Cleanup temporary files after a certain time
- [x] Cleanup unused functions
- [ ] Use hash of file content to further improve caching and speed of application

## Resources

### Libraries

- libreoffice
  - *does work, can be used in a headless environment, but installation is a bit tricky*
- msoffice2pdf
  - *does work, but only a wrapper for libreoffice call*
  - *requires libreoffice to be installed*
  - <https://github.com/robsonlimadeveloper/msoffice2pdf>
- docx2pdf
  - *cannot be used, it requires a local Word installation*
- aspose-words
  - *commerial*
  - *does not work, requires some additional dotnet installation and probably more*
  - <https://blog.aspose.com/words/convert-word-to-pdf-in-python/>
  - <https://pypi.org/project/aspose-words/>
  - <https://github.com/aspose-words/Aspose.Words-for-Python-via-.NET>

### LibreOffice

Since libreoffice is the only library that works in a headless linux docker container based environment, it is the only option that is tested.<br>
Below commands are working in the shell of the docker container:

```bash
soffice --headless --convert-to pdf sample.docx
soffice --headless --convert-to pdf:writer_pdf_Export sample.docx
soffice --headless --convert-to pdf:writer_pdf_Export --outdir . sample.docx
```

Or using python subprocess call:

```python
import subprocess
def generate_pdf(doc_path, path):
    subprocess.call(['soffice',
        '--headless',
        '--convert-to',
        'pdf',
        '--outdir',
        path,
        doc_path])
    return doc_path
generate_pdf("docx_path.docx", "output_path")
```

#### LibreOffice Headless in Docker

- <https://stackoverflow.com/questions/72434000/running-libreoffice-converter-on-docker>
- <https://stackoverflow.com/questions/30349542/command-libreoffice-headless-convert-to-pdf-test-docx-outdir-pdf-is-not>
- <https://unix.stackexchange.com/questions/700814/debian-minimum-libreoffice-packages-to-convert-docx-files-to-pdf-in-headless-mo>

Required packages to install in the docker container:

```text
default-jre-headless
libreoffice-core-nogui
libreoffice-writer-nogui
libreoffice-java-common
```

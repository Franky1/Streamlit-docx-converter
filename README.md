# Streamlit Template Project

Streamlit template project for new streamlit projects.

[![Streamlit App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://streamlit.io/)

## Status

> Work in progress - Do not use yet - Last changed: 2023-02-16

## Decription

Trying to evaluate the best way to convert docx to pdf in a headless environment. The goal is to use this in a docker container. It seems that **libreoffice** is currently the best option, but it is not easy to install in a docker container.

## ToDo

- [ ] Make a streamlit demo app that can convert docx to pdf
- [ ] Add session based temporary file storage
- [ ] Allow to upload multiple files
- [ ] Allow to download multiple files
- [ ] Test it on streamlit cloud

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

Since libreoffice is the only library that works, it is the only one that is tested.

Below commands are working in the shell of the docker container.

```shell
soffice --headless --convert-to pdf sample.docx
soffice --headless --convert-to pdf:writer_pdf_Export sample.docx
```

> below untested yet, but should work

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

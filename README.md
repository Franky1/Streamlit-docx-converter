# Streamlit Template Project

Streamlit template project for new streamlit projects.

[![Streamlit App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://streamlit.io/)

## Decription

## Resources

### Libraries

> untested

- docx2pdf
  - *cannot be used, it requires a Word installation*
- aspose-words
  - *probably commerial?*
  - <https://blog.aspose.com/words/convert-word-to-pdf-in-python/>
  - <https://pypi.org/project/aspose-words/>
  - <https://github.com/aspose-words/Aspose.Words-for-Python-via-.NET>
- libreoffice
- msoffice2pdf
  - <https://github.com/robsonlimadeveloper/msoffice2pdf>

### LibreOffice

> untested

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

## Status

> Unfinished - Do not use - Last changed: 2023-02-15

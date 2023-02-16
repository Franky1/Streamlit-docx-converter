'''
This is a test of the Aspose Words for Python library.
Unfortunately, it does not work.
The installation of the pip package is not sufficient alone.
It requires additional dotnet libraries to be installed.
Installation is not straightforward and not documented.
I assume that some sort of license is also required.
Probably not worth the effort.
'''

import aspose.words as aw

# Load word document
doc = aw.Document("sample.docx")

# Save as PDF
doc.save("PDF.pdf")

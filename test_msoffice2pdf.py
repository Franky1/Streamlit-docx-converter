'''
Test libreoffice conversion with msoffice2pdf.
msoffice2pdf is just a wrapper for the libreoffice command line call.
It requires libreoffice to be installed.
This works, however, i prefer to call libreoffice directly, because it is more flexible.
'''

from msoffice2pdf import convert

output = convert(source="sample.docx", output_dir=".", soft=1)

print(output)

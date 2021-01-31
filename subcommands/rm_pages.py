import os
from PyPDF2.pdf import *




def removePagesBasedOnTextSub(inputname, outputname, identifiers):
    os.makedirs('output', exist_ok=True)
    os.makedirs('output/removedPages', exist_ok=True)

    with open(inputname, "rb") as inputfile:
        with open(f"output/removedPages/{outputname}", "wb") as outputfile:
            output = PdfFileWriter()
            for page in PdfFileReader(inputfile).pages:
                text = page.extractText()
                if all(x in text for x in identifiers):
                    continue
                output.addPage(page)
            output.write(outputfile)

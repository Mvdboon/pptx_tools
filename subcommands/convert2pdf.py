import os
import comtypes.client


def pptx2pdfSub(inputFileName, hidden=False):
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1
    hidden_setting = 1 if hidden else 0

    os.makedirs('output/converted2PDF', exist_ok=True)
    input = "{}\{}".format(os.path.dirname(
        os.path.realpath(__file__)), inputFileName)
    output = "{}\{}".format(os.path.dirname(os.path.realpath(
        __file__)), inputFileName.replace("pptx", "pdf").replace("input", "output\converted2PDF"))

    deck = powerpoint.Presentations.Open(input)
    deck.ExportAsFixedFormat(output, FixedFormatType=2,
                             PrintHiddenSlides=hidden_setting)

    deck.Close()
    powerpoint.Quit()

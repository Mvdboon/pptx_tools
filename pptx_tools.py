import glob
import json
import multiprocessing
import os
import re
import sys
from datetime import datetime

import click


from subcommands.replace_string import replaceStr
import subcommands.fix_year
from subcommands.convert2pdf import pptx2pdfSub
from  subcommands.rm_pages import removePagesBasedOnTextSub

@click.group()
def main():
    """For explanation on the subcommands use them plus the help flag."""
    print("hoi")

    
@main.command()
@click.argument('find', required=True, nargs=1)
@click.argument('replace', required=True, nargs=1)
@click.option("-f", "--file", default=None, help="If you only want to process one file, select it here.")
@click.option("-r", "--regex", is_flag=True, help="Use regex")
def string_replace(file, find, replace):
    "Replace FIND with REPLACE string in the file. Default behaviour is to process all pptx in the input folder. Output will be placed in output folder."
    if file:
        replaceStr(file, [find], replace, fixyear=True)
    else:
        if not os.path.exists('input'):
            os.mkdir("input")
            click.echo("ERROR: Please place the files in the input folder.")
        for f in glob.glob("input/*.pptx"):
            replaceStr(f, [find], replace, fixyear=True)

@main.command()
@click.option("-f", "--file", default=None, help="If you only want to process one file, select it here.")
@click.option('--year', default=str(datetime.today().year), help=f"Change if you need a different year then {datetime.today().year}.", type=str)
def fix_year(file, year):
    "Replace old years in file(s) with YEAR. Default search string is '<year> Deloitte'"
    replacement = f"{year} Deloitte"
    if file:
        replaceStr(file, ["2016 Deloitte", "2017 Deloitte", "2018 Deloitte",
                          "2019 Deloitte", "2020 Deloitte", "2021 Deloitte"], year, fixyear=True)
    else:
        if not os.path.exists('input'):
            os.mkdir("input")
            click.echo("ERROR: Please place the files in the input folder.")
        for f in glob.glob("input/*.pptx"):
            replaceStr(f, ["2016 Deloitte", "2017 Deloitte", "2018 Deloitte",
                           "2019 Deloitte", "2020 Deloitte", "2021 Deloitte"], year, fixyear=True)

@main.command()
@click.option("-n", "--num-processes", default=4, show_default=True, help="number of parallel jobs.")
@click.option("-h", "--hidden", default=False, help="Include hidden slides in PDF.", type=bool)
@click.option("-f", "--file", default=None, help="If you only want to convert one file, select it here.")
def convert2pdf(num_processes, file, hidden):
    """Convert pptx to pdf. Default behaviour is to convert all pptx in the input folder. Output will be placed in output folder."""
    if not file:
        if not os.path.exists('input'):
            os.mkdir("input")
            click.echo("ERROR: Please place the files in the input folder.")
        convertall(num_processes, hidden)
    else:
        pptx2pdfSub(file, hidden) if "pptx" in file else click.Abort(
            "Wrong file type")

def convertall(num_processes, hidden):
    if not os.path.exists('input'):
        click.Abort("No input folder")
    p = multiprocessing.Pool(num_processes)
    for file in glob.glob("input/*.pptx"):
        p.apply_async(pptx2pdfSub, (file, hidden))
    p.close()
    p.join()

@main.command()
@click.option("-i", "--identifiers", default=["A.", "B.", "C."], show_default=True, help="List of strings that all have to be present for excluding the page.")
def rm_Pages(identifiers):
    """Remove pages based on list of strings present on PDF page."""
    for file in glob.glob("input/*.pdf"):
        removePagesBasedOnTextSub(file, file, identifiers)

if __name__ == "__main__":
    multiprocessing.freeze_support()
    if getattr(sys, 'frozen', False):
        print("hallo")
        main()
        print("hey")
    print("hey3")
    
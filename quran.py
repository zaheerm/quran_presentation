"""
quran.py

Usage:
  quran.py <sura> <outputfile> [--start=<ayat>] [--end=<ayat>]

Options:
  --start=<ayat>        Start ayat
  --end=<ayat>          End ayat
"""

import docopt
import os,sys,getopt,struct
import csv
from cStringIO import StringIO
from odf.opendocument import OpenDocumentPresentation
from odf.style import Style, MasterPage, PageLayout, PageLayoutProperties, \
TextProperties, GraphicProperties, ParagraphProperties, DrawingPageProperties
from odf.text import P
from odf.draw  import Page, Frame, TextBox, Image, LayerSet, Layer
from xml.dom import minidom

# also defined getImageData function that returns content_type, width, height of image.

def load_suras(xmlfilename, translationfilename):
    xmldoc = minidom.parse(xmlfilename)
    suras = {}
    for sura in xmldoc.getElementsByTagName("sura"):
        this_sura = {}
        for ayat in sura.getElementsByTagName("aya"):
            this_sura[int(ayat.getAttribute("index"))] = dict(arabic=ayat.getAttribute("text"))
        suras[int(sura.getAttribute("index"))] = this_sura
    for row in csv.reader(open(translationfilename, "r")):
        sura = int(row[1])
        ayat = int(row[2])
        english = row[3]
        suras[sura][ayat]["english"] = english
    return suras

def create_presentation(sura_number, outputfile, start=None, end=None):
    suras = load_suras("quran-simple.xml", "shakir_table.csv")
    doc = OpenDocumentPresentation()

    # We must describe the dimensions of the page
    pagelayout = PageLayout(name="MyLayout")
    dp = Style(name="dp1", family="drawing-page")
    dp.addElement(DrawingPageProperties(backgroundvisible="true", backgroundobjectsvisible="true"))
    doc.automaticstyles.addElement(pagelayout)
    doc.automaticstyles.addElement(dp)
    pagelayout.addElement(PageLayoutProperties(margin="0pt", pagewidth="800pt",
        pageheight="600pt", printorientation="landscape", backgroundcolor="#000000"))

    ls = LayerSet()
    ls.addElement(Layer(name="layout"))
    ls.addElement(Layer(name="background"))
    ls.addElement(Layer(name="backgroundobjects"))
    ls.addElement(Layer(name="title"))
    doc.masterstyles.addElement(ls)

    # Style for the title frame of the page
    # We set a centered 34pt font with yellowish background
    titlestyle = Style(name="MyMaster-title", family="presentation")
    titlestyle.addElement(ParagraphProperties(textalign="center"))
    titlestyle.addElement(TextProperties(fontsize="60pt", fontsizeasian="96pt", fontsizecomplex="96pt", color="#ffffff", fontfamily="'Al Bayan'", fontfamilyasian="'Al Bayan'", fontfamilycomplex="'Al Bayan'"))
    titlestyle.addElement(GraphicProperties(fillcolor="#000000"))
    doc.styles.addElement(titlestyle)
    masterstyle = Style(name="MyMaster-dp", family="drawing-page")
    masterstyle.addElement(DrawingPageProperties(fill="solid", fillcolor="#000000", backgroundsize="border", fillimagewidth="0cm", fillimageheight="0cm"))
    doc.styles.addElement(masterstyle)
    # Every drawing page must have a master page assigned to it.
    masterpage = MasterPage(name="MyMaster", pagelayoutname=pagelayout, stylename=masterstyle)

    doc.masterstyles.addElement(masterpage)

    for number, ayat in suras[sura_number].iteritems():
        if start is None or (number >= int(start) and number <= int(end)):
            page = Page(stylename=dp, masterpagename=masterpage)
            doc.presentation.addElement(page)
            titleframe = Frame(stylename=titlestyle, width="800pt", height="300pt", x="0pt", y="0pt")
            page.addElement(titleframe)
            textbox = TextBox()
            titleframe.addElement(textbox)
            textbox.addElement(P(stylename=titlestyle, text=ayat["arabic"]))
            secondframe = Frame(stylename=titlestyle, width="800pt", height="300pt", x="0pt", y="300pt")
            page.addElement(secondframe)
            secondbox = TextBox()
            secondframe.addElement(secondbox)
            secondbox.addElement(P(stylename=titlestyle, text=ayat["english"]))
            print "Added ", number, ayat
    doc.save(outputfile)

if __name__ == '__main__':
    arguments = docopt.docopt(__doc__, version='quran.py 0.1')
    create_presentation(int(arguments["<sura>"]), arguments["<outputfile>"], arguments["--start"], arguments["--end"])
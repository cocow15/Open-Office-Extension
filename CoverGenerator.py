import uno

# a UNO struct later needed to create a document
from com.sun.star.text.ControlCharacter import PARAGRAPH_BREAK
from com.sun.star.text.TextContentAnchorType import AS_CHARACTER
from com.sun.star.awt import Size

from com.sun.star.lang import XMain


def insertTextIntoCenterCell( table, cellName, text):
    tableText = table.getCellByName( cellName )
    tableText.setString( text )
    tableText.getByIndex(0)

def createTable():
    ctx = uno.getComponentContext()
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext( "com.sun.star.frame.Desktop",ctx)

    # open a writer document
    doc = desktop.loadComponentFromURL( "private:factory/swriter","_blank", 0, () )

    text = doc.Text
    cursor = text.createTextCursor()

    textFrame = doc.createInstance( "com.sun.star.text.TextFrame" )
    textFrame.setSize( Size(6000,200))
    textFrame.setPropertyValue( "AnchorType" , AS_CHARACTER )
    text.insertTextContent( cursor, textFrame, 0 )
     
    textInTextFrame = textFrame.getText()
    cursorInTextFrame = textInTextFrame.createTextCursor()
    textInTextFrame.insertString( cursorInTextFrame, "Welcome to Cover Generator!!", 0 )

    # create a text table
    table = doc.createInstance( "com.sun.star.text.TextTable" )

    # with 3 rows and 2 columns
    table.initialize( 3, 2)

    text.insertTextContent( cursor, table, 0 )

    table.getCellByName("A1").setString("Nama:   Dimas Wisnu Saputro")
    table.getCellByName("B1").setString("Mata Kuliah:   APPL Praktik")

    table.getCellByName("A2").setString(" NIM:       201524005")
    table.getCellByName("B2").setString(" Tanggal:      18 Maret 2022")
 
    table.getCellByName("A3").setString(" Prodi:    DIV - Teknik Informatika")
    insertTextIntoCenterCell( table, "B3", "                            TUGAS")


g_exportedScripts = createTable,

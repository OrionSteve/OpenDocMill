#!/usr/bin/env python3

"""Adds OpenDocMill template fields to an ODT document"""

import xml.dom.minidom
import zipfile, sys

TEXT = "urn:oasis:names:tc:opendocument:xmlns:text:1.0"
TABLE = "urn:oasis:names:tc:opendocument:xmlns:table:1.0"
OFFICE = "urn:oasis:names:tc:opendocument:xmlns:office:1.0"

class TemplateError(IOError):
    pass

def getOneByTagName(doc, NS, tag):
    nodes = doc.getElementsByTagNameNS(NS, tag)
    if len(nodes) == 0:
        raise TemplateError(f"Cannot find {NS}:{tag}")
    if len(nodes) > 1:
        raise TemplateError(f"Multiple {NS}:{tag} nodes")
    return nodes[0]

def getVariableDecls(doc):
    """XXX for some reason this breaks if called more than once"""
    nodes = doc.getElementsByTagNameNS(TEXT, "variable-decls")
    if len(nodes) > 1:
        raise TemplateError("Multiple text:variable-decls nodes" % (NS, tag))
    elif len(nodes) == 0:
        decls = doc.createElement("text:variable-decls")
        officeText = getOneByTagName(doc, OFFICE, "text")
        officeText.insertBefore(decls, officeText.firstChild)
        return decls
    else:
        return nodes[0]

def create(inTemplateTemplate, outTemplate, fields, tables):
    inZipFile = zipfile.ZipFile(inTemplateTemplate, 'r')
    inContent = inZipFile.read("content.xml")
    doc = xml.dom.minidom.parseString(inContent)

    decls = getVariableDecls(doc)
    appendFields(doc, fields, decls)
    for tableName, tableFields in tables.items():
        appendTable(doc, tableName, tableFields, decls)

    outContent = doc.toxml().encode('UTF-8')
    outZipFile = zipfile.ZipFile(outTemplate, "w")
    for fileInfo in inZipFile.filelist:
        if fileInfo.filename == "content.xml":
            outZipFile.writestr(fileInfo, outContent)
        else:
            outZipFile.writestr(fileInfo, inZipFile.read(fileInfo.filename))
    inZipFile.close()
    outZipFile.close()

def appendFields(doc, fields, decls):
    officeText = getOneByTagName(doc, OFFICE, "text")
    for varName, text in fields.items():
        if isinstance(varName, str): varName = varName.decode('UTF-8')
        if isinstance(text, str): text = text.decode('UTF-8')
        # in minidom, the *NS functions do not work when creating content    
        p = doc.createElement("text:p")
        p.setAttribute("text:style-name", "Standard")
        p.appendChild(doc.createTextNode(text + ": "))
        v = doc.createElement("text:variable-set")
        v.setAttribute("text:name", varName)
        v.setAttribute("office:value-type", "string")
        v.appendChild(doc.createTextNode(varName))
        p.appendChild(v)
        officeText.appendChild(p)
        decl = doc.createElement("text:variable-decl")
        decl.setAttribute("office:value-type", "string")
        decl.setAttribute("text:name", varName)
        decls.appendChild(decl)

def appendTable(doc, tableName, tableFields, decls):
    if isinstance(tableName, str): tableName = tableName.decode('UTF-8')
    officeText = getOneByTagName(doc, OFFICE, "text")

    # in minidom, the *NS functions do not work when creating content    
    nameP = doc.createElement("text:p")
    nameP.appendChild(doc.createTextNode(tableName + u":"))
    officeText.appendChild(nameP)

    table = doc.createElement("table:table")
    table.setAttribute("table:name", tableName)
    cols = doc.createElement("table:table-column")
    cols.setAttribute("table:number-columns-repeated", str(len(tableFields)))
    table.appendChild(cols)
    officeText.appendChild(table)
    officeText.appendChild(doc.createElement("text:p")) # throw para

    nameRow = doc.createElement("table:table-row")
    valRow = doc.createElement("table:table-row")
    table.appendChild(nameRow)
    table.appendChild(valRow)
    for fieldName, fieldText in tableFields:
        fullName = tableName + u"." + fieldName
        if isinstance(fieldName, str): fieldName = fieldName.decode('UTF-8')
        if isinstance(fieldText, str): fieldText = fieldText.decode('UTF-8')
        nameCell = doc.createElement("table:table-cell")
        nameCell.setAttribute("office:value-type", "string")
        nameP = doc.createElement("text:p")
        nameP.appendChild(doc.createTextNode(fieldText))
        nameCell.appendChild(nameP)
        nameRow.appendChild(nameCell)
        valCell = doc.createElement("table:table-cell")
        valCell.setAttribute("office-value-type", "string")
        valP = doc.createElement("text:p")
        valV = doc.createElement("text:variable-set")
        valV.setAttribute("text:name", fullName)
        valV.setAttribute("office:value-type", "string")
        valV.appendChild(doc.createTextNode(fullName))
        valP.appendChild(valV)
        valCell.appendChild(valP)
        valRow.appendChild(valCell)
        decl = doc.createElement("text:variable-decl")
        decl.setAttribute("office:value-type", "string")
        decl.setAttribute("text:name", fullName)
        decls.appendChild(decl)

if __name__ == '__main__':
    try:
        inFile, outFile = sys.argv[1:3]
    except ValueError:
        print >> sys.stderr, "Usage: %s inFile.odt outFile.out"
        sys.exit(1)

    fields = [("field1", "name1"), ("field2", "name2")]
    tables = [
        ("tbl1", [
            ("field1", "tname1"),
            ("field2", "tname2"),
            ("field3", "tname3"),
        ]),
        ("tbl2", [
            ("field1", "tname1"),
            ("field2", "tname2"),
            ("field3", "tname3"),
        ]),
    ]
    create(inFile, outFile, fields, tables)

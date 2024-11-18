#!/usr/bin/env python

import zipfile
import StringIO
import os.path
import re

class DataError(Exception): pass
class TemplateError(Exception): pass

import OpenDocMill.Reader

def getStructure(ob):
    if hasattr(ob, "getStructure"): return ob.getStructure()
    return ob


def oldFormatToBookData(data):
    bd = OpenDocMill.BookData()
    if not isinstance(data, (list, tuple)):
        raise TypeError("Bad old format: should be list of dicts")
    for i, row in enumerate(data):
        if not isinstance(row, dict) or not row.has_key("name"):
            raise TypeError("Bad format for section %d: should be dict(name='', fields={}, tables={}, images={}" % i)
        name = row["name"]
        sd = SectionData(
            fields=row.get("fields", {}),
            tables=row.get("tables", {}),
            images=row.get("images", {}),
        )
        if name == "#header": bd.setHeaderData(sd)
        elif name == "#footer": bd.setFooterData(sd)
        else: bd.addSection(name, sd)
    return bd

class HeadFootData(object):
    def __init__(self):
        self.headerData = SectionData()
        self.footerData = SectionData()

    def setHeaderData(self, sectionData=None, fields={}, tables={}, images={}):
        if sectionData is None:
            sectionData = SectionData(fields, tables, images)
        self.headerData = sectionData

    def setFooterData(self, sectionData=None, fields={}, tables={}, images={}):
        if sectionData is None:
            sectionData = SectionData(fields, tables, images)
        self.footerData = sectionData


class BookData(HeadFootData):
    def __init__(self):
        super(BookData, self).__init__()
        self.sections = []

    def addSection(self, name, sectionData=None, fields={}, tables={}, images={}):
        if sectionData is None:
            sectionData = SectionData(fields, tables, images)
        if not isinstance(sectionData, SectionData):
            raise TypeError("Expected SectionData object, not %r)" % type(sectionData))
        self.sections.append((name, sectionData))


class ReportData(HeadFootData):
    def __init__(self, fields={}, tables={}, images={}):
        super(ReportData, self).__init__()
        self.mainSection = SectionData(fields, tables, images)


class SectionData(object):
    def __init__(self, fields={}, tables={}, images={}):
        fieldErrors = self.findFieldErrors(fields)
        tableErrors = self.findTableErrors(tables)
        imageErrors = self.findImageErrors(images)
        errorString = [] # join at end
        if fieldErrors:
            errorString.append("Errors in fields:\n")
            for name, msg in fieldErrors:
                errorString.append("    %s: %s\n" % (name, msg))
        if tableErrors:
            errorString.append("Errors in tables:\n")
            for name, msg in tableErrors:
                errorString.append("    %s: %s\n" % (name, msg))
        if imageErrors:
            errorString.append("Errors in images:\n")
            for name, msg in imageErrors:
                errorString.append("    %s: %s\n" % (name, msg))
        if errorString:
            raise DataError("".join(errorString))
        self.fields = fields
        self.tables = tables
        self.images = images

    def findFieldErrors(self, fields):
        fieldErrors = []
        for name in fields:
            val = fields[name]
            if not isinstance(val, (str, unicode, int, float)):
                fieldErrors.append((name, "Bad type: %r" % type(val)))
        return fieldErrors

    def findTableErrors(self, tables):
        tableErrors = []
        for name in tables:
            rows = tables[name]
            if not isinstance(rows, (list, tuple)):
                tableErrors.append((name, "Expected list of row dicts: %r" % name))
                continue
            if len(rows) == 0: continue
            badTypeRows = []
            for i, row in enumerate(rows):
                if not isinstance(row, dict): badTypeRows.append(str(i))
            if badTypeRows:
                tableErrors.append((name, "Bad types for rows: " + ",".join(badTypeRows)))
        return tableErrors

    def findImageErrors(self, images):
        imageErrors = []
        for name in images:
            filename = images[name]
            if not isinstance(filename, (str, unicode)):
                imageErrors.append((name, "Filename is not a string: %r" % type(val)))
        return imageErrors


class ODTFileTemplate(object):
    def __init__(self, inZipFilename):
        self.inZipFilename = inZipFilename
        self.contentTemplate = None
        self.stylesTemplate = None
        self.imageList = []
    
    def setContentTemplate(self, contentTemplate): self.contentTemplate = contentTemplate
    def setStylesTemplate(self, stylesTemplate): self.stylesTemplate = stylesTemplate
    def appendImage(self, filename):
        self.imageList.append(filename)

    def getStructure(self):
        return getStructure(self.contentTemplate) + getStructure(self.stylesTemplate)

    def write(self, outZipFilename, data):
        inZipFile = zipfile.ZipFile(self.inZipFilename, "r")
        outZipFile = zipfile.ZipFile(outZipFilename, "w")
        for fileInfo in inZipFile.filelist:
            if fileInfo.filename == "content.xml" and self.contentTemplate is not None:
                s = StringIO.StringIO()
                self.contentTemplate.write(s, data, self.appendImage)
                _make_utf8_evilly(s)
                outZipFile.writestr(fileInfo, s.getvalue())
            elif fileInfo.filename == "styles.xml" and self.stylesTemplate is not None:
                s = StringIO.StringIO()
                self.stylesTemplate.write(s, data, self.appendImage)
                _make_utf8_evilly(s)
                outZipFile.writestr(fileInfo, s.getvalue())
            elif fileInfo.filename == "META-INF/manifest.xml":
                pass # XXX will do this at the end, to add images
            else:
                outZipFile.writestr(fileInfo, inZipFile.read(fileInfo.filename))
        for filename in self.imageList:
            basename = os.path.basename(filename)
            outZipFile.write(filename, arcname="Pictures/%s" % basename)
       
        manifestFileList = [x for x in inZipFile.filelist if x.filename == "META-INF/manifest.xml"]
        if manifestFileList:
            manifestFileInfo = manifestFileList[0]
            manifestStr = inZipFile.read(manifestFileInfo.filename)
            extraFileTags = ["""<manifest:file-entry manifest:media-type="image/png" manifest:full-path="Pictures/%s"/>""" % os.path.basename(f) for f in self.imageList]
            newManifestStr = re.sub(ur"(?=<manifest:file-entry\b)", "".join(extraFileTags), manifestStr)
            outZipFile.writestr(manifestFileInfo, newManifestStr)

        outZipFile.close()

class XMLFileTemplate(object):
    def __init__(self, identifier, appendImage):
        self.identifier = identifier
        self.beforeText = []
        self.afterText = []
        self.appendImage = appendImage

    def addBeforeText(self, text):
        self.beforeText.append(text)

    def addAfterText(self, text):
        self.afterText.append(text)

    def write(self, stream, data, appendImage):
        # Ok, now actually write data
        stream.writelines(self.beforeText)
        self.writeParts(stream, data, appendImage)
        stream.writelines(self.afterText)


class ReportContentTemplate(XMLFileTemplate):
    def __init__(self, *args, **kwargs):
        super(ReportContentTemplate, self).__init__(*args, **kwargs)
        self.mainSection = None

    def addMainSection(self, section):
        if self.mainSection is not None:
            raise TemplateError("addMainSection has already been called")
        self.mainSection = section

    def getStructure(self):
        return [("content", "", x) for x in getStructure(self.mainSection)]

    def writeParts(self, stream, data, appendImage):
        if self.mainSection is None:
            raise TemplateError("need to call addMainSection before write")
        if not isinstance(data, ReportData):
            raise TypeError("Expected ReportData object")
        self.mainSection.write(stream, data.mainSection, appendImage)


class BookContentTemplate(XMLFileTemplate):
    def __init__(self, *args, **kwargs):
        super(BookContentTemplate, self).__init__(*args, **kwargs)
        self.sections = {}

    def addSection(self, name, section):
        if self.sections.has_key(name): raise TemplateError("Duplicate section name: %r" % (name,))
        self.sections[name] = section

    def getStructure(self):
        structure = []
        for name, section in self.sections.items():
            for x in getStructure(section):
                structure.append(("content", name, x))
        return structure

    def writeParts(self, stream, data, appendImage):
        if isinstance(data, BookData):
            dataOb = data
        elif isinstance(data, (list, tuple)):
            dataOb = oldFormatToBookData(data)
        else:
            raise TypeError("Expected BookData object, not %r" % type(data))

        errorStrings = []

        # Check for missing sections
        sectionNames = set(x[0] for x in dataOb.sections)
        missingSections = sectionNames - set(self.sections.keys())
        if missingSections:
            errorStrings.append(("The following sections are in the data but not the template: %r" % tuple(sorted(missingSections))))

        for i, (sectionName, sectionData) in enumerate(dataOb.sections):
            section = self.sections.get(sectionName)
            if section is None: continue # error trapped above
            try:
                section.write(stream, sectionData, appendImage)
            except (DataError, TypeError), ex:
                msg = str(ex)
                errorStrings.append("section %i (%r): %s" % (i, sectionName, msg))
        if errorStrings:
            raise DataError("\n".join(errorStrings))


class StylesTemplate(XMLFileTemplate):
    def __init__(self, *args, **kwargs):
        super(StylesTemplate, self).__init__(*args, **kwargs)
        self.headerSection = None
        self.footerSection = None

    def setHeaderSection(self, section):
        self.headerSection = section

    def setFooterSection(self, section):
        self.footerSection = section

    def getStructure(self):
        structure = []
        for x in getStructure(self.headerSection) or []:
            structure.append(("styles", "header", x))
        for x in getStructure(self.footerSection) or []:
            structure.append(("styles", "footer", x))
        return structure

    def writeParts(self, stream, data, appendImage):
        if isinstance(data, (list, tuple)):
            dataOb = oldFormatToBookData(data)
        elif isinstance(data, HeadFootData):
            dataOb = data
        else:
            raise TypeError("Expected HeadFootData object like ReportData or BookData (or alternatively a list), not %r" % type(data))
        if self.headerSection is not None:
            self.headerSection.write(stream, dataOb.headerData, appendImage)
        if self.footerSection is not None:
            self.footerSection.write(stream, dataOb.footerData, appendImage)


class Section(object):
    def __init__(self, identifier):
        self.identifier = identifier
        self.elements = []

    def addText(self, text): self.elements.append(("TEXT", text))
    def addVariable(self, varName): self.elements.append(("VARIABLE", varName))
    def addTable(self, tableName, table): self.elements.append(("TABLE", (tableName, table)))
    def addImage(self, imageName, defaultArcFilename): self.elements.append(("IMAGE", (imageName, defaultArcFilename)))

    def getStructure(self):
        variables = []
        for eType, eVal in self.elements:
            if eType == "VARIABLE":
                variables.append(eVal)
            elif eType == "TABLE":
                tName, tVal = eVal
                for v in getStructure(tVal):
                    variables.append(tName + "." + v)
        return variables

    def __repr__(self):
        return "Section(" + '\n'.join([repr(x) for x in self.elements]) + ")"

    def write(self, stream, data, appendImage):
        if not isinstance(data, SectionData):
            raise TypeError("Expected SectionData, not %r" % type(data))
        fieldData = data.fields
        tableData = data.tables
        imageData = data.images

        for eType, e in self.elements:
            if eType == "TEXT":
                stream.write(e)
            elif eType == "VARIABLE":
                try:
                    rawString = unicode(fieldData[e])
                except UnicodeError:
                    rawString = u"xxx"
                except KeyError:
                    raise ValueError, "No value for field %r in section %r" % (e, self.identifier)
                stream.write(xmlEscape(rawString))
            elif eType == "IMAGE":
                imageName, defaultArcFilename = e
                filename = imageData.get(imageName)
                if filename is None:
                    rawArcFilename = defaultArcFilename
                else:
                    rawArcFilename = u"Pictures/%s" % os.path.basename(filename)
                    appendImage(filename)
                stream.write(xmlEscapeAttr(rawArcFilename))
            elif eType == "TABLE":
                tableName, table = e
                try:
                    tData = tableData[tableName]
                except KeyError:
                    raise ValueError, "No data for table %r in section %r" % (tableName, self.identifier)
                table.write(stream, tData)
            else:
                raise ValueError, "Unknown type %r in template, section %r " % (eType, self.identifier)


class Table(object):
    def __init__(self, identifier):
        self.identifier = identifier
        self.beforeText = []
        self.row = None
        self.afterText = []

    def addBeforeText(self, text): self.beforeText.append(text)
    def setRow(self, row): self.row = row
    def addAfterText(self, text): self.afterText.append(text)

    def getStructure(self):
        return getStructure(self.row)

    def __repr__(self):
        return '\n'.join(
            ["    TBLB: %r" % x for x in self.beforeText]
            + ["    ROW : %r" % self.row]
            + ["    TBLA: %r" % x for x in self.afterText])

    def write(self, stream, data):
        stream.writelines(self.beforeText)
        for i in xrange(len(data)):
            rowFields = data[i]
            self.row.write(stream, rowFields, rowNo=i)
        stream.writelines(self.afterText)


class Row(object):
    def __init__(self, tableIdentifier):
        self.tableIdentifier = tableIdentifier
        self.elements = []

    def addText(self, text): self.elements.append(("TEXT", text))
    def addVariable(self, varName): self.elements.append(("VARIABLE", varName))

    def getStructure(self):
        parts = []
        for eType, eVal in self.elements:
            if eType == "VARIABLE": parts.append(eVal)
        return parts

    def __repr__(self):
        return '\n'.join(["    %r" % (x,) for x in self.elements])

    def write(self, stream, fields, rowNo):
        for eType, e in self.elements:
            if eType == "TEXT":
                stream.write(e)
            elif eType == "VARIABLE":
                try:
                    v = fields[e]
                    if v is None:
                        rawString = u""
                    if isinstance(v, str):
                        rawString = v.decode('UTF-8')
                    else:
                        rawString = unicode(v)
                except KeyError:
                    raise ValueError, "No value for field %r in table %r[row=%d]" % (e, self.tableIdentifier, rowNo)
                stream.write(xmlEscape(rawString))
            else:
                raise ValueError, "Unknown type %r in template, table %r" % (eType, self.tableIdentifier)


def xmlEscape(s):
    print "s=%r" % (s,)
    return s.replace('&', '&amp;').replace('<', '&lt;')
def xmlEscapeAttr(s):
    return s.replace('&', '&amp;').replace('<', '&lt;').replace('"', '&quot;')
def _make_utf8_evilly(stringIO):
    b = stringIO.buflist
    for i in range(len(b)):
        if isinstance(b[i], unicode): b[i] = b[i].encode("UTF-8")


#!/usr/bin/env python

import xml.dom.minidom
import xml.dom.ext.Printer
import StringIO
import zipfile
import OpenDocMill

TEXT = "urn:oasis:names:tc:opendocument:xmlns:text:1.0"
DRAW = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
TABLE = "urn:oasis:names:tc:opendocument:xmlns:table:1.0"
STYLE = "urn:oasis:names:tc:opendocument:xmlns:style:1.0"
OFFICE = "urn:oasis:names:tc:opendocument:xmlns:office:1.0"
XLINK = "http://www.w3.org/1999/xlink"

DataError = OpenDocMill.DataError
TemplateError = OpenDocMill.TemplateError

class FakeStream(object):
    def write(self, text):
        """Overwrite this method after instantiation, before calling"""
        raise NotImplementedError


class ODTVisitor(xml.dom.ext.Printer.PrintVisitor):
    def __init__(self, template, idTables, idLastRows, nsHints):
        self.template = template
        self.fakeStream = FakeStream()
        self.fakeStream.write = self.template.addBeforeText
        xml.dom.ext.Printer.PrintVisitor.__init__(self, self.fakeStream, "UTF-8", nsHints=nsHints)
        self.writeState = "TEMPLATE"
        self.idTables = idTables
        self.idLastRows = idLastRows
        self.section = None
        self.tableName = None
        self.table = None
        self.row = None
        self.sectionParent = None
        self.imageName = None
        self.defaultFilename = None

    def isSectionNode(self, node): raise NotImplementedError
    def getSectionName(self, node): raise NotImplementedError

    def visitElement(self, node):
        if self.isSectionNode(node):
            self.visitSectionStart(node)
        elif node.namespaceURI == TEXT and node.localName == "variable-set":
            self.visitVariableSet(node)
        elif node.namespaceURI == TABLE and node.localName == "table" and id(node) in self.idTables:
            self.visitTable(node)
        elif node.namespaceURI == TABLE and node.localName == "table-row" and id(node) in self.idLastRows:
            self.visitLastRow(node)
        elif node.namespaceURI == DRAW and node.localName == "image":
            self.visitImage(node)
        else:
            # just print as normal
            xml.dom.ext.Printer.PrintVisitor.visitElement(self, node)

    def visitNodeList(self, nodeList, *args, **kwargs): 
        xml.dom.ext.Printer.PrintVisitor.visitNodeList(self, nodeList, *args, **kwargs)
        if len(nodeList) and nodeList[0].parentNode is self.sectionParent:
            # will only be true AFTER visiting an H1 child
            self.sectionsHaveEnded()

    def addSection(self, node):
        """Sets self.section and adds self.section to the template"""
        raise NotImplementedError

    def visitSectionStart(self, node):
        # change state
        self.addSection(node)
        self.fakeStream.write = self.section.addText
        self.writeState = "SECTION"
        self.sectionParent = node.parentNode
        # produce xml
        xml.dom.ext.Printer.PrintVisitor.visitElement(self, node)

    def sectionsHaveEnded(self):
        # set state
        self.fakeStream.write = self.template.addAfterText
        self.writeState = "TEMPLATE"
        # Don't need to produce xml; done elsewhere

    def visitVariableSet(self, node):
        # change state
        variableName = node.getAttributeNS(TEXT, "name")
        assert variableName
        if self.writeState == "SECTION" or (self.writeState == "TABLE" and "." not in variableName):
            self.section.addVariable(variableName)
        elif self.writeState == "ROW" and variableName.startswith(self.tableName + "."):
            localName = variableName[len(self.tableName + "."):] # remove prefix
            self.row.addVariable(localName)
        else:
            raise TemplateError, "Unexpected variable %r: state=%r, tableName=%r, template=%r" % (variableName, self.writeState, self.tableName, self.template.identifier)
        # Don't need to produce xml; we're ignoring it


    def visitImage(self, node):
        assert self.writeState == "SECTION"
        assert node.parentNode.namespaceURI == DRAW and node.parentNode.localName == "frame"
        # save state
        self.imageName = node.parentNode.getAttributeNS(DRAW, "name")
        self.defaultFilename = node.getAttributeNS(XLINK, "href")
        assert self.imageName is not None
        assert self.defaultFilename is not None
        self.writeState = "IMAGE"
        # produce xml
        xml.dom.ext.Printer.PrintVisitor.visitElement(self, node)
        # restore state
        self.writeState = "SECTION"
        self.imageName = None
        self.defaultFilename = None

    def visitAttr(self, node):
        if self.writeState == "IMAGE" and node.namespaceURI == XLINK and node.localName == "href":
            self.fakeStream.write(u' %s="' % node.name)
            self.section.addImage(self.imageName, self.defaultFilename)
            self.fakeStream.write(u'"')
        else:
            xml.dom.ext.Printer.PrintVisitor.visitAttr(self, node)
        

    def visitTable(self, node):
        assert self.writeState == "SECTION"
        # save state
        oldWrite = self.fakeStream.write 
        # change state
        self.writeState = "TABLE"
        self.tableName = node.getAttributeNS(TABLE, "name")
        self.table = OpenDocMill.Table(self.section.identifier + "/" + self.tableName)
        self.section.addTable(self.tableName, self.table)
        self.fakeStream.write = self.table.addBeforeText
        # produce xml (recurse)
        xml.dom.ext.Printer.PrintVisitor.visitElement(self, node)
        # restore state
        self.fakeStream.write = oldWrite
        self.writeState = "SECTION"
        self.tableName = None
        self.table = None

    def visitLastRow(self, node):
        assert self.writeState == "TABLE"
        # change state
        self.row = OpenDocMill.Row(self.table.identifier)
        self.table.setRow(self.row)
        self.fakeStream.write = self.row.addText
        self.writeState = "ROW"
        # produce xml (recurse)
        xml.dom.ext.Printer.PrintVisitor.visitElement(self, node)
        # restore state
        self.writeState = "TABLE"
        self.fakeStream.write = self.table.addAfterText
        self.row = None

class ODTBookContentVisitor(ODTVisitor):
    def isSectionNode(self, node):
        return node.namespaceURI == TEXT and node.localName == "h" and node.getAttributeNS(TEXT, "outline-level") == "1"

    def addSection(self, node):
        sectionName = u''.join([k.nodeValue for k in node.childNodes if k.nodeType == k.TEXT_NODE]) # just reads the text
        self.section = OpenDocMill.Section(self.template.identifier + "#" + sectionName)
        self.template.addSection(sectionName, self.section)

class ODTReportContentVisitor(ODTVisitor):
    def isSectionNode(self, node):
        return node.namespaceURI == TEXT and node.localName == "variable-decls"

    def addSection(self, node):
        self.section = OpenDocMill.Section(self.template.identifier + "#MAIN")
        self.template.addMainSection(self.section)

class ODTStyleVisitor(ODTVisitor):
    def isSectionNode(self, node):
        parent = node.parentNode
        if parent is None: return False
        if not (parent.namespaceURI == STYLE and parent.localName == "master-page"): return False
        if not parent.getAttributeNS(STYLE, "name") == "Standard": return False
        return True

    def addSection(self, node):
        if node.namespaceURI != STYLE or node.localName not in ['header', 'footer']:
            raise TemplateError("unknown section type %s:%s in template %r" % (node.namespaceURI, node.localName, self.template.identifier))

        self.section = OpenDocMill.Section(self.template.identifier + "#" + node.localName)
        if node.localName == 'header':
            self.template.setHeaderSection(self.section)
        elif node.localName == 'footer':
            self.template.setFooterSection(self.section)
        else:
            pass # unreachable (checked earlier)


def parentTable(node):
    if node is None: return None
    if node.namespaceURI == TABLE and node.localName == "table":
        return node
    return node.parentNode

def getTableAndLastRowIDs(doc):
    """find table nodes which contain a last row with set-variables <tableName>.<something>"""
    idTables = {}
    idLastRows = {}
    for t in doc.getElementsByTagNameNS(TABLE, "table"):
        tableName = t.getAttributeNS(TABLE, "name")
        lastRow = (t.getElementsByTagNameNS(TABLE, "table-row") or [None])[-1]
        if t is not parentTable(lastRow): continue
        v = lastRow.getElementsByTagNameNS(TEXT, "variable-set")
        tableVars = [x for x in v if x.getAttributeNS(TEXT, "name").startswith(tableName + ".")]
        if not tableVars: continue
        # save table and last row ID, for efficient checking on next pass
        idTables[id(t)] = t
        idLastRows[id(lastRow)] = lastRow
    return idTables, idLastRows

def readXML(xmlStream, fileIdentifier, VisitorClass, TemplateClass, appendImage):
    doc = xml.dom.minidom.parse(xmlStream)
    idTables, idLastRows = getTableAndLastRowIDs(doc)
    nss = xml.dom.ext.SeekNss(doc)
    template = TemplateClass(unicode(fileIdentifier), appendImage)
    visitor = VisitorClass(template, idTables, idLastRows, nss)
    xml.dom.ext.Printer.PrintWalker(visitor, doc).run()
    return template

def readBookContentXML(xmlStream, filename, appendImage):
    return readXML(xmlStream, filename + "#content.xml", ODTBookContentVisitor, OpenDocMill.BookContentTemplate, appendImage)

def readReportContentXML(xmlStream, filename, appendImage):
    return readXML(xmlStream, filename + "#content.xml", ODTReportContentVisitor, OpenDocMill.ReportContentTemplate, appendImage)

def readStylesXML(xmlStream, filename, appendImage):
    return readXML(xmlStream, filename + "#styles.xml", ODTStyleVisitor, OpenDocMill.StylesTemplate, appendImage)

def readBookODT(filename):
    template = OpenDocMill.ODTFileTemplate(filename)
    inZipFile = zipfile.ZipFile(filename, 'r')
    content = StringIO.StringIO(inZipFile.read("content.xml"))
    styles = StringIO.StringIO(inZipFile.read("styles.xml"))
    inZipFile.close()
    template.setContentTemplate(readBookContentXML(content, filename, template.appendImage))
    template.setStylesTemplate(readStylesXML(styles, filename, template.appendImage))
    return template

def readReportODT(filename):
    template = OpenDocMill.ODTFileTemplate(filename)
    inZipFile = zipfile.ZipFile(filename, 'r')
    content = StringIO.StringIO(inZipFile.read("content.xml"))
    styles = StringIO.StringIO(inZipFile.read("styles.xml"))
    inZipFile.close()
    template.setContentTemplate(readReportContentXML(content, filename, template.appendImage))
    template.setStylesTemplate(readStylesXML(styles, filename, template.appendImage))
    return template

def readODT(filename):
    return readBookODT(filename)

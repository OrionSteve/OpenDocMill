#!/usr/bin/env python3

import xml.dom.minidom
import io
import zipfile
import OpenDocMill
from xml.etree import ElementTree

TEXT = "urn:oasis:names:tc:opendocument:xmlns:text:1.0"
DRAW = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
TABLE = "urn:oasis:names:tc:opendocument:xmlns:table:1.0"
STYLE = "urn:oasis:names:tc:opendocument:xmlns:style:1.0"
OFFICE = "urn:oasis:names:tc:opendocument:xmlns:office:1.0"
XLINK = "http://www.w3.org/1999/xlink"

DataError = OpenDocMill.DataError
TemplateError = OpenDocMill.TemplateError

class XMLPrinter:
    """Modern replacement for xml.dom.ext.Printer"""
    def __init__(self, stream, encoding="UTF-8", nsHints=None):
        self.stream = stream
        self.encoding = encoding
        self.nsHints = nsHints or {}
        self.indent_level = 0
        
    def write(self, text):
        if isinstance(text, (list, tuple)):
            for item in text:
                self.write(item)
        else:
            self.stream.write(str(text))

    def visit(self, node):
        if node.nodeType == node.ELEMENT_NODE:
            self.visitElement(node)
        elif node.nodeType == node.TEXT_NODE:
            self.visitText(node)
        elif node.nodeType == node.ATTRIBUTE_NODE:
            self.visitAttr(node)

    def visitElement(self, node):
        # Get the original XML string for this node
        if node.prefix:
            tagName = f"{node.prefix}:{node.localName}"
        else:
            tagName = node.localName

        # Start tag
        self.write(f"<{tagName}")
        
        # Write all attributes
        for attr in node.attributes.values():
            self.visitAttr(attr)
            
        if node.childNodes:
            self.write(">")
            # Visit child nodes
            for child in node.childNodes:
                self.visit(child)
            self.write(f"</{tagName}>")
        else:
            self.write("/>")

    def visitText(self, node):
        # Write text content directly, preserving whitespace
        text = node.data
        if text:
            self.write(text)

    def visitAttr(self, attr):
        if attr.prefix:
            name = f"{attr.prefix}:{attr.localName}"
        else:
            name = attr.localName
        value = attr.value
        self.write(f' {name}="{value}"')

class FakeStream(object):
    def write(self, text):
        """Overwrite this method after instantiation, before calling"""
        raise NotImplementedError

class ODTVisitor(XMLPrinter):
    def __init__(self, template, idTables, idLastRows, nsHints):
        self.template = template
        self.fakeStream = FakeStream()
        self.fakeStream.write = self.template.addBeforeText
        super().__init__(self.fakeStream, "UTF-8", nsHints)
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
            XMLPrinter.visitElement(self, node)

    def visitNodeList(self, nodeList, *args, **kwargs): 
        XMLPrinter.visitNodeList(self, nodeList, *args, **kwargs)
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
        XMLPrinter.visitElement(self, node)

    def sectionsHaveEnded(self):
        # set state
        self.fakeStream.write = self.template.addAfterText
        self.writeState = "TEMPLATE"
        # Don't need to produce xml; done elsewhere

    def visitVariableSet(self, node):
        # change state
        variableName = node.getAttributeNS(TEXT, "name")
        assert variableName
        
        # Handle table variables
        if "." in variableName:
            tableName, localName = variableName.split(".", 1)
            if self.writeState == "ROW" and tableName == self.tableName:
                if self.row is None:
                    raise TemplateError(f"Found table variable {variableName} but no row is active")
                self.row.addVariable(localName)
                # Write the variable placeholder
                self.write(f'<text:variable-set text:name="{variableName}">')
                self.write(f"${{{variableName}}}")  # Add placeholder
                self.write('</text:variable-set>')
            elif self.writeState == "TABLE" and tableName == self.tableName:
                if self.row is None:
                    self.row = OpenDocMill.Row(self.table.identifier)
                    self.table.setRow(self.row)
                self.row.addVariable(localName)
                # Write the variable placeholder
                self.write(f'<text:variable-set text:name="{variableName}">')
                self.write(f"${{{variableName}}}")  # Add placeholder
                self.write('</text:variable-set>')
            else:
                raise TemplateError("Unexpected variable %r: state=%r, tableName=%r, template=%r" % 
                                  (variableName, self.writeState, self.tableName, self.template.identifier))
        # Handle non-table variables
        elif self.writeState == "SECTION":
            self.section.addVariable(variableName)
            # Write the variable placeholder
            self.write(f'<text:variable-set text:name="{variableName}">')
            self.write(f"${{{variableName}}}")  # Add placeholder
            self.write('</text:variable-set>')
        elif self.writeState == "TABLE":
            self.section.addVariable(variableName)
            # Write the variable placeholder
            self.write(f'<text:variable-set text:name="{variableName}">')
            self.write(f"${{{variableName}}}")  # Add placeholder
            self.write('</text:variable-set>')
        else:
            raise TemplateError("Unexpected variable %r: state=%r, tableName=%r, template=%r" % 
                              (variableName, self.writeState, self.tableName, self.template.identifier))

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
        XMLPrinter.visitElement(self, node)
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
            XMLPrinter.visitAttr(self, node)
        

    def visitTable(self, node):
        # If we're not in SECTION state but have a valid section, switch to it
        if self.writeState != "SECTION" and self.section is not None:
            self.writeState = "SECTION"
        
        if self.writeState != "SECTION":
            raise TemplateError(f"Found table but not in section state (state={self.writeState})")

        # save state
        oldWrite = self.fakeStream.write 
        # change state
        self.writeState = "TABLE"
        self.tableName = node.getAttributeNS(TABLE, "name")
        self.table = OpenDocMill.Table(self.section.identifier + "/" + self.tableName)
        self.section.addTable(self.tableName, self.table)
        self.fakeStream.write = self.table.addBeforeText
        
        # produce xml (recurse)
        XMLPrinter.visitElement(self, node)
        
        # restore state
        self.fakeStream.write = oldWrite
        self.writeState = "SECTION"
        self.tableName = None
        self.table = None

    def visitLastRow(self, node):
        # Check if we're in a table context, even if not explicitly in TABLE state
        if self.table is None:
            parent = parentTable(node)
            if parent:
                self.tableName = parent.getAttributeNS(TABLE, "name")
                self.table = OpenDocMill.Table(self.section.identifier + "/" + self.tableName)
                self.section.addTable(self.tableName, self.table)
                self.writeState = "TABLE"

        if self.writeState != "TABLE":
            raise TemplateError(f"Found table row but not in table state (state={self.writeState})")

        # change state
        self.row = OpenDocMill.Row(self.table.identifier)
        self.table.setRow(self.row)
        self.fakeStream.write = self.row.addText
        self.writeState = "ROW"
        # produce xml (recurse)
        XMLPrinter.visitElement(self, node)
        # restore state
        self.writeState = "TABLE"
        self.fakeStream.write = self.table.addAfterText
        self.row = None

class ODTBookContentVisitor(ODTVisitor):
    def isSectionNode(self, node):
        return node.namespaceURI == TEXT and node.localName == "h" and node.getAttributeNS(TEXT, "outline-level") == "1"

    def addSection(self, node):
        sectionName = ''.join([k.nodeValue for k in node.childNodes if k.nodeType == k.TEXT_NODE])
        self.section = OpenDocMill.Section(self.template.identifier + "#" + sectionName)
        self.template.addSection(sectionName, self.section)

class ODTReportContentVisitor(ODTVisitor):
    def __init__(self, template, idTables, idLastRows, nsHints):
        super().__init__(template, idTables, idLastRows, nsHints)
        # Initialize main section immediately
        self.section = OpenDocMill.Section(self.template.identifier + "#MAIN")
        self.template.addMainSection(self.section)
        self.writeState = "SECTION"  # Start in SECTION state

    def isSectionNode(self, node):
        # For report templates, we want to start the main section at the beginning
        # of the document
        return False  # We don't need to look for section nodes since we create the section immediately

    def addSection(self, node):
        pass  # Section is already created in __init__

    def visitElement(self, node):
        if node.namespaceURI == TEXT and node.localName == "variable-set":
            self.visitVariableSet(node)
        elif node.namespaceURI == TABLE and node.localName == "table" and id(node) in self.idTables:
            self.visitTable(node)
        elif node.namespaceURI == TABLE and node.localName == "table-row" and id(node) in self.idLastRows:
            self.visitLastRow(node)
        elif node.namespaceURI == DRAW and node.localName == "image":
            self.visitImage(node)
        else:
            # just print as normal
            XMLPrinter.visitElement(self, node)

class ODTStyleVisitor(ODTVisitor):
    def isSectionNode(self, node):
        parent = node.parentNode
        if parent is None: return False
        if not (parent.namespaceURI == STYLE and parent.localName == "master-page"): return False
        if not parent.getAttributeNS(STYLE, "name") == "Standard": return False
        return True

    def addSection(self, node):
        if node.namespaceURI != STYLE or node.localName not in ['header', 'footer']:
            raise TemplateError("unknown section type %s:%s in template %r" % 
                              (node.namespaceURI, node.localName, self.template.identifier))

        self.section = OpenDocMill.Section(self.template.identifier + "#" + node.localName)
        if node.localName == 'header':
            self.template.setHeaderSection(self.section)
        elif node.localName == 'footer':
            self.template.setFooterSection(self.section)

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
    nss = {}  # Replace xml.dom.ext.SeekNss with empty dict for now
    template = TemplateClass(str(fileIdentifier), appendImage)  # Changed unicode to str
    visitor = VisitorClass(template, idTables, idLastRows, nss)
    visitor.visit(doc)  # Use our new XMLPrinter's visit method
    return template

def readBookContentXML(xmlStream, filename, appendImage):
    return readXML(xmlStream, filename + "#content.xml", ODTBookContentVisitor, OpenDocMill.BookContentTemplate, appendImage)

def readReportContentXML(xmlStream, filename, appendImage):
    return readXML(xmlStream, filename + "#content.xml", ODTReportContentVisitor, OpenDocMill.ReportContentTemplate, appendImage)

def readStylesXML(xmlStream, filename, appendImage):
    return readXML(xmlStream, filename + "#styles.xml", ODTStyleVisitor, OpenDocMill.StylesTemplate, appendImage)

def readBookODT(filename):
    template = OpenDocMill.ODTFileTemplate(filename)
    with zipfile.ZipFile(filename, 'r') as inZipFile:
        content = io.BytesIO(inZipFile.read("content.xml"))
        styles = io.BytesIO(inZipFile.read("styles.xml"))
    template.setContentTemplate(readBookContentXML(content, filename, template.appendImage))
    template.setStylesTemplate(readStylesXML(styles, filename, template.appendImage))
    return template

def readReportODT(filename):
    template = OpenDocMill.ODTFileTemplate(filename)
    with zipfile.ZipFile(filename, 'r') as inZipFile:
        content = io.BytesIO(inZipFile.read("content.xml"))
        styles = io.BytesIO(inZipFile.read("styles.xml"))
    template.setContentTemplate(readReportContentXML(content, filename, template.appendImage))
    template.setStylesTemplate(readStylesXML(styles, filename, template.appendImage))
    return template

def readODT(filename):
    return readBookODT(filename)

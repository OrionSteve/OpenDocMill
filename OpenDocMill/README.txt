This is a perl wrapper for OpenDoc Mill (a python ODT generator).
It uses JSON to pass complex data structures from perl to python.

PREREQUISITES

perl: JSON.pm
python
EasyInstall (see http://peak.telecommunity.com/DevCenter/EasyInstall)
Then do: sudo easy_install OpenDocMill*.egg

USAGE

See invoiceTemplate.odt for the ODT template.
See createInvoice.pl for an example data structure.
Running createInvoice.pl creates an output file: out.odt

TEMPLATE CREATION

The template is created in OpenOffice.org as an ODT document (*not* an OOo
template).

There should be a "Heading 1" section at the very start of the document.
Simple templates should contain no other Heading 1 sections.

Plain fields:

- These are single items of text that appear once per section.

- Fields are added by choosing Insert->Fields->Other->Set Variable, and then
setting the Name and Value to the same thing (e.g. to "invoiceNo"), and the
Format to Text.

Table row fields:

- These appear in the last row of the table template.  The row is filled in, 
repeated as many times as the data set requires.

- All OpenOffice.org tables have a name, set by selecting the table and 
changing Table->Table Properties->Table->Name (e.g. to "table1").

- In the last row of the table, add fields preceded by the table name and 
a dot (e.g. set both the field's Name and Value to "table1.vatRate").

OUTPUT

Running createInvoice.pl loads the template and feeds in a data structure to
produce an output document.  Rendering, formatting, word-wrapping, headers,
footers etc. are all defined by the template in the natural way.

OpenOffice.org itself is not used to produce the output document from the
template.  The template file is loaded directly into an XML parser, and 
output documents are produced from this by replacing variable placeholders
by the requisite values, which is fast (less than a second for a several
hundred page document) and requires little memory.

The process produces OpenOffice.org ODT documents.  However PDFs can be
generated programmatically from these: see core/warehouse/makepdfs.py .


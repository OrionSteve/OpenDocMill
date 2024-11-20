#!/usr/bin/env python

import sys
import os
import codecs
scriptdir = os.path.dirname(sys.argv[0])
libdir = os.path.join(scriptdir, "OpenDocMill")
if os.path.isdir(libdir):
    sys.path.append(libdir)
import json
try:
    import OpenDocMill
except ImportError:
    if not os.path.isdir(libdir):
        print >> sys.stderr, "WARNING: Cannot find %r" % libdir
    raise

progName = sys.argv[0]
args = sys.argv[1:]

if len(args) != 2:
    print >> sys.stderr, "Usage: ", progName, "inTemplate.odt outDoc.odt < data.json"
    sys.exit(1)

inTemplate, outDoc = args
# TODO check with  utf-8 encoded data.
input = sys.stdin.read() # read whole multi-line input as string

inputData = json.read(input) # convert to list of (page info data structure)
fields    = inputData.get("fields", {})
tables    = inputData.get("tables", {})
images    = inputData.get("images", {})
# print("FIELDS: %r" % (fields,))
# print("TABLES: %r" % (tables,))
# print( "IMAGES: %r" % (images,))

reportData = OpenDocMill.ReportData(fields=fields, tables=tables, images=images)
reportData.setHeaderData(fields=fields, tables=tables, images=images)
reportData.setFooterData(fields=fields, tables=tables, images=images)

reportTemplate = OpenDocMill.Reader.readReportODT(inTemplate) # load template
reportTemplate.write(outDoc, reportData) # add data; create output

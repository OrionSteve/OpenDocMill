#!/usr/bin/env python3

import sys
import os
#import json as stdlib_json  # Rename to avoid conflict with our custom json module

scriptdir = os.path.dirname(sys.argv[0])
libdir = os.path.join(scriptdir, "OpenDocMill")
if os.path.isdir(libdir):
    sys.path.append(libdir)

try:
    import OpenDocMill
    import json  # Our custom json module
except ImportError:
    if not os.path.isdir(libdir):
        print("WARNING: Cannot find %r" % libdir, file=sys.stderr)
    raise

progName = sys.argv[0]
args = sys.argv[1:]

if len(args) != 2:
    print("Usage: ", progName, "inTemplate.odt outDoc.odt < data.json", file=sys.stderr)
    sys.exit(1)

inTemplate, outDoc = args

input_data = sys.stdin.read()  # read whole multi-line input as string 
raw_data = json.loads(input_data) 

# Convert the raw data into a ReportData object
if isinstance(raw_data, dict):
    # Extract fields, tables, and images from the input data
    fields = raw_data.get('fields', {})
    tables = raw_data.get('tables', {})
    images = raw_data.get('images', {})
    inputData = OpenDocMill.ReportData(fields=fields, tables=tables, images=images)
elif isinstance(raw_data, list):
    # Handle old format
    if len(raw_data) > 0 and isinstance(raw_data[0], dict):
        first_section = raw_data[0]
        fields = first_section.get('fields', {})
        tables = first_section.get('tables', {})
        images = first_section.get('images', {})
        inputData = OpenDocMill.ReportData(fields=fields, tables=tables, images=images)
    else:
        inputData = OpenDocMill.ReportData()
else:
    inputData = OpenDocMill.ReportData()

reportTemplate = OpenDocMill.Reader.readReportODT(inTemplate)  # load template
reportTemplate.write(outDoc, inputData)  # add data; create output

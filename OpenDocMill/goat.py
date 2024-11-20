#!/usr/bin/env python3

import sys
import os
import json as stdlib_json  # Rename to avoid conflict with our custom json module

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

with open('blob.json', 'r') as file:
    input_data = file.read()
inputData  = json.read(input_data)  # Use our custom json.read instead of json.loads

reportTemplate = OpenDocMill.Reader.readODT(inTemplate)  # load template
reportTemplate.write(outDoc, inputData)  # add data; create output

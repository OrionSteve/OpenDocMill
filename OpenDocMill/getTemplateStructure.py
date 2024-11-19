#!/usr/bin/env python3

import sys
import os
import json

scriptdir = os.path.dirname(sys.argv[0])
libdir = os.path.join(scriptdir, "OpenDocMill")
if os.path.isdir(libdir):
    sys.path.append(libdir)

try:
    import OpenDocMill
except ImportError:
    if not os.path.isdir(libdir):
        print("WARNING: Cannot find %r" % libdir, file=sys.stderr)
    raise

progName = sys.argv[0]
args = sys.argv[1:]

if len(args) != 1:
    print("Usage: ", progName, "inTemplate.odt > data.json", file=sys.stderr)
    sys.exit(1)

inTemplate, = args

reportTemplate = OpenDocMill.Reader.readReportODT(inTemplate)  # load template
structure = reportTemplate.getStructure()
print(json.dumps(structure))

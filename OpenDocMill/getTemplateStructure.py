#!/usr/bin/env python

import sys
import os
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

if len(args) != 1:
    print >> sys.stderr, "Usage: ", progName, "inTemplate.odt > data.json"
    sys.exit(1)

inTemplate, = args

reportTemplate = OpenDocMill.Reader.readReportODT(inTemplate) # load template
structure = reportTemplate.getStructure()
print json.write(structure)

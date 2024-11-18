#!/usr/bin/env python

import sys
import os
scriptdir = os.path.dirname(sys.argv[0])
eggname = "OpenDocMill-1.0-py%d.%d.egg" % sys.version_info[:2]
eggpath = os.path.join(scriptdir, eggname)
if os.path.exists(eggpath):
    sys.path.append(eggpath)
import json
try:
    import OpenDocMill.TemplateCreator
except ImportError:
    if not os.path.exists(eggpath):
        print >> sys.stderr, "WARNING: Cannot find %r" % eggpath
    raise


progName = sys.argv[0]
args = sys.argv[1:]

if len(args) != 2:
    print >> sys.stderr, "Usage: ", progName, "inTemplateTemplate.odt outTemplate.odt < data.json"
    sys.exit(1)

inTemplateTemplate, outTemplate = args

input = sys.stdin.read() # read whole multi-line input as string 
inputData = json.read(input) # convert to list of (page info data structure)
fields = inputData["fields"]
tables = inputData["tables"]

OpenDocMill.TemplateCreator.create(inTemplateTemplate, outTemplate, fields, tables)

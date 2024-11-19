#!/usr/bin/env python3

import sys
import os
import json

scriptdir = os.path.dirname(sys.argv[0])
eggname = "OpenDocMill-1.0-py%d.%d.egg" % sys.version_info[:2]
eggpath = os.path.join(scriptdir, eggname)
if os.path.exists(eggpath):
    sys.path.append(eggpath)

try:
    import OpenDocMill.TemplateCreator
except ImportError:
    if not os.path.exists(eggpath):
        print("WARNING: Cannot find %r" % eggpath, file=sys.stderr)
    raise

progName = sys.argv[0]
args = sys.argv[1:]

if len(args) != 2:
    print("Usage: ", progName, "inTemplateTemplate.odt outTemplate.odt < data.json", file=sys.stderr)
    sys.exit(1)

inTemplateTemplate, outTemplate = args

input_data = sys.stdin.read()  # read whole multi-line input as string 
inputData = json.loads(input_data)  # convert to list of (page info data structure)
fields = inputData["fields"]
tables = inputData["tables"]

OpenDocMill.TemplateCreator.create(inTemplateTemplate, outTemplate, fields, tables)

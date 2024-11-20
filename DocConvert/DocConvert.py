#!/usr/bin/env python3
"""Convert between document formats with OOo, from the command line.


AUTHOR

Copyright (c) 2024, Orion Web Technologies (https://www.orionwt.co.uk/)

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.
"""
import sys
import os
import re
import time
import getopt
import subprocess
import shutil

class ConvertError(Exception): pass

def absolute(filePath):
    return os.path.abspath(os.path.join(os.getcwd(), filePath))

def bEscape(s):
    """Returns string with BASIC escaping"""
    return s.replace('"', '""')

def waitFor(f, args, timeout):
    t = time.time()
    while True:
        val = f(*args)
        if val is not None: return val
        if time.time() - t > timeout: return None
        time.sleep(0.1)

def readLog(logFilename, sessionId):
    """Returns text if logFile is finished (=ends with sessionId); else None"""
    if not os.path.exists(logFilename): return None
    with open(logFilename) as logFile:
        log = logFile.read()
    if log.endswith(sessionId):
        return log[:-len(sessionId)]
    return None

FILTER_FOR_EXTENSION = dict(
    pdf="writer_pdf_Export",
    xhtml="XHTML Writer File",
    html="XHTML Writer File",
    odt="writer8",
    doc="MS Word 97",
)

class DocConverter(object):
    def __init__(self, sofficeCmd, timeout):
        self.sofficeCmd = sofficeCmd
        self.timeout=timeout

    def guessFilterCode(self, extension):
        e = extension.lower()[1:]
        if e not in FILTER_FOR_EXTENSION:
            raise ConvertError("Unknown output file extension: %r" % extension)
        return e + ":" + FILTER_FOR_EXTENSION[e]

    def convert(self, inFilename, outFilename):
        inAbs = absolute(inFilename)
        outAbs = absolute(outFilename)

        inStart, inExt = os.path.splitext(inAbs)
        outStart, outExt = os.path.splitext(outAbs)

        # If there is no conversion to do, just copy
        if inExt.lower() == outExt.lower():
            shutil.copy(inAbs, outAbs)
            return

        outDirName = os.path.dirname(outAbs)

        inTmp = outStart + inExt
        if inTmp != inAbs:
            shutil.copy(inAbs, inTmp)

        filterCode = self.guessFilterCode(outExt)
        args = [self.sofficeCmd, "--headless", "--convert-to", filterCode, "--outdir", outDirName, inTmp]
        sys.stderr.write("Running Conversion: %r" % args)
        p = subprocess.Popen(args)
        # wait for p to terminate
        while True:
            time.sleep(0.1)
            result = p.poll()
            if result is not None: break
        # wait for logFilename to appear
        startTime = time.time()
        while not os.path.exists(outAbs):
            if time.time() - startTime > 20:
                raise ConvertError("Conversion did not run or was very slow")
            time.sleep(0.1)
        if inTmp != inAbs:
            os.unlink(inTmp)

if __name__ == '__main__':
    opts, args = getopt.gnu_getopt(sys.argv[1:], "", ["cmd=", "timeout="])
    try:
        inFile, outFile = args
    except ValueError:
        sys.stderr.write("Usage: %s [--cmd=soffice] [--timeout=20] inFile outFile\n" % sys.argv[0])
        sys.exit(1)
    cmd = dict(opts).get("--cmd", "soffice")
    timeout = int(dict(opts).get("--timeout", "20"))
    dc = DocConverter(cmd, timeout)
    dc.convert(inFile, outFile)

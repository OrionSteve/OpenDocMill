#!/usr/bin/env python3
import OpenDocMill
t = OpenDocMill.Reader.readReportODT("report-in.odt")
for x in t.getStructure(): print(x)
print("===")
t = OpenDocMill.Reader.readBookODT("invoiceTemplate.odt")
for x in t.getStructure(): print(x)

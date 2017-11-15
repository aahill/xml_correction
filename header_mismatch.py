import os
import re
import subprocess
"""
iterates through a CSV file containing header mismatch information, and inject the correct 
header xml to a WDCML file
"""
# constants for the location to the header mismatch csv, 
# and SourceDepot location, respectively. edit these for your system

#NOTE: this script is setup to process a csv containing 
#       only header mismatches, with a pipe char ("|") as a delimiter

CSV_LOC = "C:\\Users\\aahi\\projects\\migration\\header_mismatch\\Header_Mismatch_Data_3.csv"
SD_LOC = "C:\\Users\\aahi\\Sd"

file_info =[]
lines = []
errors = []
processed = []
num_processed = 0

#open the csv file containing the header_mismatch information, and append the split lines to a list
with open(CSV_LOC, "r") as f:
    header = True # skip over the header line in the csv
    for line in f:
        if header:
            header = False
        else:
            lines.append(line.split("|"))


for line in lines:
    #variables located in the csv lines
    filename = line[0] # the "name" of the API
    md_headers = line[1] # the header(s) listed in the API's .md file
    wdcml_headers = line[2] # the headers listed in the API's WDCML file
    project = line[5] # the project
    filetype = line[7] # the WDCML topic type
    subtype = line[8] # the WDCML subtopic type
    xml_loc = line[13] # the location of the xml (WDCML) file

    # conditions for processing a file as listed in the CSV
    # NOTE: <ovw> WDCML types cannot have <header> tags, and so must be skipped.
    # NOTE: refpages have dissimilar syntax, and are currently skipped
    if filetype != "ovw" and subtype != "ovw" and filetype != "refpage":

        file_loc = os.path.join(SD_LOC, project, xml_loc.replace(".htm",".xml"))
        content = "" # blank variable to store xml content
        with open(file_loc, "r") as xml_file:
            content = xml_file.read() # read in file
            # use regex to find the header in the xml
            # and get the specific header name
            header = re.search(r"\<header\>\s*<filename>([\w\.]+)\s*</filename>",content)
            if header:
                header = header.group(1)
            # regex to find the include headers
            include_headers= re.findall(r"<include_header>\s*<filename>([\w\.]+)\s*</filename>\s*</include_header", content)
            #list containing all unique header names to remove duplicates from the content
            new_include = []
            for h in include_headers:
                #create list with normalized header names already captured from the file. 
                #this list is compared against the current header name to determine if it's a duplicate. 
                #if not, add it ot the list of unique header names
                if h.lower() not in [x.lower() for x in new_include]:
                    new_include.append(h)
            #add the main header to become a <include_header> if it isn't already; the header specified in the .md file
            #  will be the <header>
            if header and header.lower() not in [x.lower() for x in new_include]:
                new_include.append(header)
            # create new header xml tag to replace the old one
            new_header = "<header><filename>"+ md_headers + "</filename>\n"
            if "</header>" not in content:
                new_header += "</header>"
            # start building a string of <include_header> tags
            new_include_str = ""
            for include in new_include:
                new_include_str += "<include_header><filename>"+include+"</filename></include_header>\n"

            # replace header tag with the new header xml line, if it exists
            if header is not None:
                content = re.sub(r"\<header\>\s*<filename>.+?</filename>", new_header, content)
            # if a <header> tag does not exist, try searching for <info> or </info> to fill out. If not,
            # the script will fail, and it must either be programmed to be created, or manually inserted.
            else:
                try:
                    assert "<info>" in content
                    content = re.sub("<info>", "<info>\n"+new_header+"\n", content)
                except AssertionError:
                    assert "<info/>" in content, filename+" does not have an info tag"
                    content = re.sub("<info/>", "<info>\n"+new_header+"\n</info>", content)
            # replace header_include (or insert if necesary)
            if new_include:
                if "<include_header>" in content:
                    content = re.sub(r"<include_header>.+</include_header>", new_include_str, content)
                else:
                    content = re.sub(r"</header>",new_include_str+"\n"+"</header>", content)

        # automatically checkout the file from source depot, so it can be edited.
        # if the file is checked out or out of sync, an authorization exception will be raised 
        powershell = "C:\\WINDOWS\\system32\\WindowsPowerShell\\v1.0\\powershell.exe "
        command = "sd edit " + file_loc
        # NOTE: set the execution policy to allow the script to be run
        call = [powershell, '-ExecutionPolicy', 'Unrestricted', " & ", command]
        # NOTE change the working directory to the project path before execution
        return_code = subprocess.call(call, cwd=SD_LOC)
        # store the processed file for output later
        processed.append(os.path.join(xml_loc.replace(".htm",".xml")))
        # write the edited topic content to the file
        with open(file_loc, "w") as f_write:
            f_write.write(content)
            #update the count of processed files
            num_processed += 1
            
#print the processed files
print("files (", num_processed, ") processed:")
for p in processed:
    print (p)
#print(num_processed, " files successfully fixed")
#print("\nthe following files could not be written:")
#for f in errors:
#    print(f)

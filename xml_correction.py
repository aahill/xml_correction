#!/usr/bin/env python3
"""
fixes xml files that have a page type mismatch

"""

import xml.etree.ElementTree as ET
import re
import os
import subprocess

# the terms found in xml files to specify API parameters
P_TERMS = ["members", "parameters", "params", "constants"]

# location of SD files and xml stub location
sd_loc = "C:\\Users\\aahi\\Sd" # source depot location
stub_loc = "//ANDKI_BOOK\\Output\\StubFilesOutput"


def preprocess_xml(xml_file):
    """take an xml file, read the contents, and prepare it for processing by
    ElementTree. Most importantly, this function removes namespaces from the xml,
    which hinder functionality"""

    with open(xml_file, 'r', encoding="utf-8") as f:
        xml_text = f.read()
    # strip xml declarations (lines starting with <?xml) from the xml string
    xml_text = re.sub(r'\<\?xml[^\>]+\>\n', "", xml_text)
    # remove schema and namespace information, which can cause issues when the
    # xml is processed
    xml_text = re.sub(r'<(\w+) (?:xsi:schemaLocation|xmlns)[^>]+>', r"<\1>", xml_text)
    # strip all non-unicode characters
    xml_text = re.sub(r'[^\x00-\x7F]+', '', xml_text)
    # remove namespace for easier parsing
    xml_text = re.sub(r'\<(\w+)\sxmlns=.+\>\n', r"<\1>", xml_text)

    return xml_text

def replace(oldtree, newtree):
    """clears an xml subtree and replaces it with the new subtree.
    Note that the root of both subtrees should be the same"""
    oldtree.clear()
    oldtree.text = newtree.text
    oldtree.tail = newtree.tail
    for elem in list(newtree):
        oldtree.append(elem)

def transfer_single_node(stub_tree, orig_tree, tag, default_loc):
    """transfers the content from a single node in orig_tree to stub_tree,
    along with all child nodes. if the node doesn't exist in orig_tree,
    it will be placed in the default location."""
    stub_node = find_node(tag, stub_tree)
    orig_node = find_node(tag, orig_tree)
    if orig_node:
        # if the seealso section exists, replace the stub one, or place it in the xml
        if stub_node:
            replace(stub_node, orig_node)
        else:
            stub_tree.getroot().find(default_loc).append(orig_node)

def find_node(node_name, tree):
    """ returns a specific node from the xml tree"""
    for node in tree.iter():
        if node.tag == node_name:
            return node
    return None

def get_param_term(tree):
    "gets the xml tag for parameters"
    for node in tree.getroot().iter():
        tag = node.tag.lower()
        if tag in P_TERMS:
            return tag
    return None


def transfer_retval(stub_tree, orig_tree, pagetype):
    real_retval = find_node("retval", orig_tree)
    stub_retval = find_node("retval", stub_tree)
    param_val = get_param_term(stub_tree)
    if real_retval:
        if not stub_retval:
            ret = stub_tree.getroot().find("content/syntax/"+param_val)
            ret.append(real_retval)
        else:
            replace(stub_retval, real_retval)

def transfer_metadata(stub_tree, orig_tree):
    """get the metadata info from the elementTree (ET) generated from the original xml file
    and append it into the stub xml ET"""

    stub_metadata = stub_tree.getroot().find("metadata")
    #stub_root.remove(stub_abstract)
    real_metadata = orig_tree.getroot().find("metadata")

    # dictionary containing metadata attributes from the original xml
    metadata_dict = {
        "msdnID": real_metadata.get("msdnID"),
        "beta": "0"
    }
    for tag, value in metadata_dict.items():
        stub_metadata.set(tag, value)

def transfer_metadata_old(stub_tree, original_tree, page_type):
    """get the metadata from the elementTree (ET) generated from the original xml file
    and append it into the stub xml ET"""

    # capture the metadata elements from each xml tree
    for orig_metadata in original_tree.iter('metadata'):
        for stub_metadata in stub_tree.iter('metadata'):
            # replace the type listed in the metadata with the proper type
            stub_metadata.set('type', page_type)
            stub_metadata.set('beta', "0")
            stub_metadata = orig_metadata
            return None

def fill_xml(htm_file, xml_stub, original_xml):
    """starts the file correction process
    takes an htm file and extracts its text to the xml stub file"""

    # preprocess xml files to remove namespaces, amongst other things
    # note that a full ET tree is created after parsing the xml passed in a string
    stub_tree = ET.ElementTree(
        ET.fromstring(
            preprocess_xml(xml_stub)))

    orig_tree = ET.ElementTree(
        ET.fromstring(
            preprocess_xml(
                original_xml)))

    pagetype = stub_tree.getroot().tag # the stub xml file contains the correct pagetype
    transfer_metadata(stub_tree, orig_tree) # place original metadata in stub file

    # extract htm file parameters and associate them with the proper text as a dictionary
    p_term = get_param_term(orig_tree)
    param_dict = extract_params_from_htm(htm_file, p_term)
    if pagetype != "ioctl":
        add_params_to_stub(pagetype, param_dict, stub_tree)
    # add abstract
    add_abstract_to_stub(stub_tree, orig_tree)
    # add info tag
    #add_info_to_stub(stub_tree, orig_tree)

    # add remarks, seealso, and info
    transfer_single_node(stub_tree, orig_tree, "remarks", "content")
    transfer_single_node(stub_tree, orig_tree, "seealso", "content")
    transfer_single_node(stub_tree, orig_tree, "info", "content")

    if pagetype != "ioctl":
        transfer_retval(stub_tree, orig_tree, pagetype)

    # add schema information to make the xml render as wdcml
    schema_info = {
        "xsi:schemaLocation":"http://microsoft.com/wdcml ../../BuildX/Schema/xsd/wdcml.xsd",
        "xmlns":"http://microsoft.com/wdcml",
        "xmlns:xsi":"http://www.w3.org/2001/XMLSchema-instance",
        "xmlns:msxsl": "urn:schemas-microsoft-com:xslt",
        "xmlns:x":"http://www.w3.org/1999/xhtml",
        "xmlns:mssdk":"winsdk",
        "xmlns:script":"urn:script",
        "xmlns:build":"urn:build",
        "xmlns:MSHelp":"http://msdn.microsoft.com/mshelp",
        "xmlns:dx":"http://ddue.schemas.microsoft.com/authoring/2003/5",
        "xmlns:xlink":"http://www.w3.org/1999/xlink"
    }
    for attribute, val in schema_info.items():
        stub_tree.getroot().set(attribute, val)
    return stub_tree

def add_abstract_to_stub(stub_tree, orig_tree):
    real_abstract = orig_tree.getroot().find("content/desc/p/abstract")
    if not real_abstract:
        real_abstract = orig_tree.getroot().find("content/desc")
    stub_abstract_loc = stub_tree.getroot().find("content").find("desc")
    replace(stub_abstract_loc, real_abstract)

def add_info_to_stub(stub_tree, orig_tree):
    """adds original info to the stub xml tree"""
    real_info = orig_tree.getroot().find("content").find("info")
    for orig_metadata in orig_tree.iter('info'):
        for stub_metadata in stub_tree.iter('info'):
            stub_metadata = orig_metadata

def add_params_to_stub(pagetype, param_fields, stub_tree):
    """adds the parameter text scraped from the htm file, and appends it to the stub tree"""

    # maps the api type to the param section
    # (which is most likely named differently depending on api type)
    

    param_location = get_param_term(stub_tree)
    # param terms don't occur in pluralized tag names
    if param_location[-1] == "s":
        param_location = param_location[:-1]
    for param_element in stub_tree.getroot().iter(param_location):
        if param_element.tag == param_location:
            try:
                param_name = param_element.find("name").text
                param_text = param_element.find("desc")
                param_text.text = param_fields[param_name]
            except:
                print(param_name, ": NOT FOUND in FILE")

def alter_html_tags(html_string):
    """converts or removes html tags from text"""

    conversion_dict = {
        "p":""
    }

    converted = html_string
    for html, md in conversion_dict.items():
        converted = converted.replace("<" + html + ">", md) \
            .replace("</" + html + ">", md) # perform replacement for closing tags

    return converted

def extract_params_from_htm(htm_file, p_term):
    """iterates through an html file, looking for an h2 tag that signifies
    where the params are listed. These params and their associated text
    are also extracted, converted to markdown, and returned in a dictionary"""
    text_dict = {}
    # p_term is usually capitalized in xml file, and must be lenthened if "params"
    if p_term == "params":
        p_term = "Parameters"
    else:
        p_term = p_term[0].upper() + p_term[1:]
    # find section of text containing param info
    htm_lines = open(htm_file).read().replace('\n', '')
    re_term = '<h2>' + p_term + '</h2>(.+?)<h2>'
    param_search_str = re.compile(re_term) # used to find parameters section'
    param_html = param_search_str.search(htm_lines).group()

    # get param names and associated info
    ##param_name_list = re.findall(r'<dt>(.+?)</dt>\s*<dd>(.+?)</dd>', param_html)
    param_name_list = re.findall(r'<dt>(?:<\w+?>)(\w+).*?</dt>\s*<dd>(.+?)</dd>', param_html)
    for param in param_name_list:
        param_title = param[0].strip()
        param_text = alter_html_tags(param[1].strip())
        text_dict[param_title] = param_text
    return text_dict

def get_file_mapping(csv_loc):
    """looks through csv "type mismatch" file to generate a
    map of files that need correcting. Returns a dict of projects
    and their files, with relevant info"""
    file_dict = {}
    header_line = True
    with open(csv_loc, "r") as f:
        for line in f:
            # header in csv should be skipped
            if not header_line:
                # use the regex to split commas that are not inside other brackets
                line_l = re.split(r",(?![^\[]*?[^\]]\])", line) #(r",\s*(?![^\[\]]*\))",line)
                title = line_l[0]
                #print(title)
                header = line_l[1]
                project = line_l[5]
                htm_location = line_l[13]
                stub_location = line_l[14]
                owner = line_l[15]
                file_dict[title] = {
                    "title":title,
                    "header":header,
                    "project":project,
                    "htm_location":htm_location,
                    "md_location":stub_location,
                    "owner":owner
                }
            else:
                header_line = False
    return file_dict

def get_filepaths(file_info, stub_loc, sd_loc):
    """gets the three filepaths needed for conversion:
    the original xml, the stub xml, and the htm file"""

    def generate_project_html(project_path):
        """ runs an html build process for a project"""

        powershell = "C:\\WINDOWS\\system32\\WindowsPowerShell\\v1.0\\powershell.exe "
        cbx_command = "CbX msdn hxs_msdn"
        # set the execution policy to allow the script to be run
        # set the execution policy to allow the script to be run
        call = [powershell, '-ExecutionPolicy', 'Unrestricted', " & ", cbx_command]
        # change the working directory to the project path before execution
        return_code = subprocess.call(call, cwd=project_path)
        return return_code
    project = file_info["project"]
    # get the stub location from md_location
    stub_filepath = file_info["md_location"] \
        .replace("SkeletonMD", "SkeletonXML") \
        .replace(".md", ".xml")
    stub_filepath = os.path.join(stub_loc, stub_filepath)

    # get original xml filename from project/filename combo
    orig_filename = file_info["htm_location"] \
        .split("\\")[1] \
        .replace(".htm", "") # .htm not useful
    orig_filepath = os.path.join(sd_loc, project, project, orig_filename) + ".xml"

    # get the htm file, building the project if necessary
    htm_build_path = os.path.join(sd_loc, project, "build", "HxS_MSDN")
    htm_filepath = os.path.join(htm_build_path, file_info["htm_location"])

    if not os.path.exists(htm_build_path):
        ret_code = generate_project_html(os.path.join(sd_loc, project))
        # retcode == 1 signifies a build error
        assert ret_code != 1, ("build error for project: " + htm_build_path)

    return (orig_filepath, stub_filepath, htm_filepath)

def write_tree(tree, filename, output_path):
    """ write ET xml tree to output file """
    #change filename if illegal characters exist
    fname = filename.lower().replace("::", "_")
    #write xml versioning info
    with open(os.path.join(output_path, fname +".xml"), 'w') as f:

        f.writelines('<?xml version="1.0" encoding="utf-8"?>\n')
        f.writelines('<?xml-stylesheet type="text/xsl" href="../../BuildX/Script2/preview.xslt"?>\n')
    #write actual tree content
    with open(os.path.join(output_path, fname +".xml"), 'ab') as f:
        tree.write(f)


def main():
    """ main entry to the migration program """
    cwd = os.getcwd()
    csv_loc = os.path.join(cwd, "type_mismatch_3.csv")

    base_output_dir = "out"
    writers = ["aahi", "andki", "bagold", "domars", "dumacmic", "nabazan", "prwilk", "tedhudek"]
    #get a mapping of projects with files that need to be converted
    file_dict = get_file_mapping(csv_loc)

    #make output folders
    for writer in writers:
        if not os.path.exists(base_output_dir):
            os.makedirs(os.path.join(base_output_dir, writer))
    failed_files = []
    
    #clear file containing filese that couldn't be converted,
    # before writing to it
    open("failed_files.txt", 'w').close()
    # get file info, and start conversion
    for conversion_info in file_dict.items():
        try:
            # prune conversion info to the stuff that helps
            conversion_info = conversion_info[1]

            print("converting: ", conversion_info["title"])
            orig, stub, htm = get_filepaths(conversion_info, stub_loc, sd_loc)
            converted_tree = fill_xml(htm, stub, orig)
            owner = conversion_info["owner"].replace("REDMOND\\", "").replace("\n", "")
            out = os.path.join(base_output_dir, owner, conversion_info["project"])
            #make directory if needed
            if not os.path.exists(out):
                os.makedirs(out)
            write_tree(converted_tree, conversion_info["title"], out)
        # if an error occurs during the correction process, print it to the log file
        except:
            failure_str = "title: " +conversion_info["title"] + "\n" + \
                "project: " + conversion_info["project"] + "\n" + \
                "owner: " + conversion_info["owner"]
            print("!!failure!!")
            failed_files.append(failure_str)
    
    with open("failed_files.txt","a") as f:
        for failed in failed_files:
            f.write(failed+"\n")


def test():
    """test xml correction process on single file"""
    cwd = os.getcwd()
    #filename = "poscxclose"
    filename = "nfccxdevicedeinitialize"
    name_base = os.path.join(cwd, filename)
    test_htm = name_base + ".htm"
    test_xml_stub = name_base + "_stub.xml"
    test_orig_xml = name_base + "_orig.xml"

    filled_tree = fill_xml(test_htm, test_xml_stub, test_orig_xml)

    filled_tree.write(open('test_'+filename+'.xml', 'wb'))

main()

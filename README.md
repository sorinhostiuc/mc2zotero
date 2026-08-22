# MC2Zotero

MC2Zotero converts Mendeley Cite content controls in Microsoft Word `.docx` files into Zotero citation fields. The converted document can then be edited and refreshed with the Zotero Word plugin.

## Before conversion

Import the Mendeley library into Zotero first. In Zotero, use `File > Import` and select the Mendeley option. Better matches are possible when the imported items retain their DOI values and Mendeley identifiers.

## Installation

Download `mc2zotero-1.0.1.xpi` from the latest release. Open `Tools > Plugins` in Zotero, choose `Install Add-on From File`, and select the XPI. The current release supports Zotero 7, 8, and 9.

## Converting a document

Open `Tools > Convert Mendeley Cite Citations...`, select the `.docx` document, and review the detected references. During conversion, MC2Zotero reads the Mendeley content controls embedded in the Word package, identifies each cited record by its Mendeley identifier when that value is available, falls back to the DOI or normalized title and year, and can create a missing Zotero item from the metadata retained in the document.

After conversion, open the output file in Word and select `Refresh` from the Zotero toolbar. Check any unresolved citations before continuing work on the document.

## Building from source

On Windows, run `npm run build:windows`. On Unix-like systems, run `npm run build`. The XPI is written to the repository root.

## License

MC2Zotero is released under the MIT License. See [LICENSE](LICENSE).

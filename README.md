# Office-File-Explorer

The purpose of this tool is to provide potential file specific troubleshooting of Office Open Xml formatted documents for Word, Excel and PowerPoint (.docx, .dotx, .docm, .xlsx, .xlst, .xlsm, .pptx, .pptm).

# Note
Keep in mind if you use this on a production document and choose to use something that changes or removes data, you should be working on a copy of the file, not the original.  

## List of features

### Word
* List function to display (content controls, styles, hyperlinks, List Templates, fonts, footnotes, endnotes, document properties, authors, revisions/tracked changes, comments, field codes, bookmarks, paragraphs, paragraph styles)
* Delete content (headers / footers, orphan list templates, page breaks, comments, hidden text, footnotes, endnotes)
* Convert Macro enabled file (.docm) to non-macro enabled (.docx)
* Fix corrupt documents
* Fix corrupt bookmarks
* Fix corrupt revisions
* Remove PII

### Excel
* List function to display (links, comments, worksheets, hidden rows & columns, shared strings, cell values, connections, defined names)
* Delete content (comments, links)
* Convert Macro enabled file (.xlsm) to non-macro enabled (.xlsx) 

### PowerPoint
* List function to display (hyperlinks, slide titles, slide text, comments)
* Convert Macro enabled file (.pptm) to non-macro enabled (.pptx)
* Reset note page size to default value (if the button doesn't fix the issue, go to File | Settings and enable the notes master checkbox)
* Remove PII

### Shared
* List function to display (Ole Objects, shapes, custom properties, package parts)
* Add custom properties for a file
* Change theme for a file
* Validate underlying xml file details

### Batch File Processing (following features can be used to change many documents at one time)
* Change Theme (Word, Excel, PowerPoint)
* Add Custom Properties (Word, Excel, PowerPoint)
* Reset note page size to default value (PowerPoint)
* Fix corrupt bookmarks (Word)
* Fix corrupt revisions (Word)
* Remove Personally Identifiable Information (Word, Excel, PowerPoint)

# App UI

### Main Window
![image](https://github.com/desjarlais/desjarlais.github.io/blob/master/img/ofcmain.png?raw=true)

### Batch File Processing Window
![image](https://github.com/desjarlais/desjarlais.github.io/blob/master/img/ofcbatch.png?raw=true)

# Help
If you need assistance (find a bug, have a question or any suggestions or feedback), please report them using the [Issues tab](https://github.com/desjarlais/Office-File-Explorer/issues)

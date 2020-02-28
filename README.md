# Office-File-Explorer

The purpose of this tool is to provide potential file specific troubleshooting of Office Open Xml formatted documents for Word, Excel or PowerPoint (.docx, .dotx, .docm, .xlsx, .xlst, .xlsm, .pptx, .pptm).

## List of features

### Word
* list function to display (content controls, styles, hyperlinks, List Templates, fonts, footnotes, endnotes, document properties, authors, revisions/tracked changes, comments, field codes, bookmarks, paragraphs, paragraph styles)
* delete content (headers / footers, list templates, page breaks, comments, hidden text, footnotes, endnotes)
* convert Macro enabled file (.docm) to non-macro enabled (.docx)
* Fix corrupt documents
* Remove PII

### Excel
* list function to display (links, comments, worksheets, hidden rows & columns, shared strings, cell values, connections, defined names)
* delete content (comments, links)
* convert Macro enabled file (.xlsm) to non-macro enabled (.xlsx) 

### PowerPoint
* list function to display (hyperlinks, slide titles, slide text, comments)
* convert Macro enabled file (.pptm) to non-macro enabled (.pptx)
* reset note page size to default value

### Shared
* list function to display (Ole Objects, shapes, custom properties, package parts)
* add custom properties for a file
* change theme for a file
* validate underlying xml file details

### Batch Processing (following features can be used to change many documents at one time)
* Change Theme
* Add Custom Properties
* Fix Note Page Size
* Remove Personally Identifiable Information

# Note
Keep in mind if you use this on a production document and choose to use something that changes or removes data, you should be working on a copy of the file, not the original.  

# App UI

## Main Window
![image](https://github.com/desjarlais/desjarlais.github.io/blob/master/img/ofcmain.png?raw=true)

## Batch Window
![image](https://github.com/desjarlais/desjarlais.github.io/blob/master/img/ofcbatch.png?raw=true)

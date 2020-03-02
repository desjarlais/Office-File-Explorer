namespace Office_File_Explorer.App_Helpers
{
    class StringResources
    {
        // global app strings
        public const string word = "Word";
        public const string excel = "Excel";
        public const string powerpoint = "PowerPoint";
        public const string noCustomDocProps = "** There are no custom file properties in this file **";
        public const string fileDoesNotExist = "** File does not exist **";
        public const string noFootnotes = "** No Footnotes in this document **";
        public const string noEndnotes = "** No Endnotes in this document **";
        public const string themeFileAdded = "Theme File Added.";
        public const string unableToDownloadUpdate = "Unable to download update.";
        public const string noOle = "** This document does not contain OLE objects **";
        public const string noShapes = "** Document does not contain any shapes **";
        public const string txtFallbackStart = "<mc:Fallback>";
        public const string txtFallbackEnd = "</mc:Fallback>";
        public const string invalidTag = "Invalid Tag: ";
        public const string replacedWith = "Replaced With: ";
        public const string errorUnableToFixDocument = "ERROR: Unable to fix document.";
        public const string errorText = "Error: ";
        public const string pptNotesSizeReset = "Notes Page Size Reset.";
        public const string sEnd = "end";
        public const string sBegin = "begin";
        public const string colon = ": ";
        public const string colonBuffer = " : ";
        public const string period = ". ";
        public const string emptyString = "";
        public const string docSecurity = "DocSecurity";
        public const string arrow = " --> ";
        public const string nonEmptyId = "Target Id cannot be empty.";
        public const string duplicateId = "OOXML part Id <1> already exists.";
        public const string helpLocation = "https://github.com/desjarlais/Office-File-Explorer/issues";
        public const string shpChart = "Chart";
        public const string shpOfficeDrawing = ". Office Drawing";
        public const string shpVml = "Vml Shape";
        public const string shpMath = ". Math Shape";
        public const string shpDrawingDgm = ". Drawing Diagram Shape";
        public const string shpChartDraw = ". Chart Drawing Shape";
        public const string shpChartShape = ". Chart Shape";
        public const string shpShape = ". Shape";
        public const string shp3D = ". 3D Shape";
        public const string shpXlDraw = ". Spreadsheet Drawing";


        // schema base urls
        public const string schemaOxml2006 = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
        public const string schemaMsft2007 = "http://schemas.microsoft.com/office/2007/relationships/";
        public const string schemaMsft2006 = "http://schemas.microsoft.com/office/2006/relationships/";

        // Office package relationship ids
        public const string CustomUIPartRelType = schemaMsft2006 + "ui/extensibility";
        public const string CustomUI14PartRelType = schemaMsft2007 + "ui/extensibility";
        public const string QATPartRelType = schemaMsft2006 + "ui/customization";
        public const string ImagePartRelType = schemaOxml2006 + "image";

        // WordprocessingML package relationship ids
        public const string wordMainAttributeNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        public const string AfPartType = schemaOxml2006 + "aFChunk";
        public const string CommentsPartType = schemaOxml2006 + "comments"; // same as Excel, PowerPoint
        public const string DocumentSettingsPartType = schemaOxml2006 + "settings";
        public const string EndnotesPartType = schemaOxml2006 + "endnotes";
        public const string FontsTablePartType = schemaOxml2006 + "fontTable";
        public const string FooterPartType = schemaOxml2006 + "footer";
        public const string FootnotesPartType = schemaOxml2006 + "footnotes";
        public const string GlossaryDocPartType = schemaOxml2006 + "glossaryDocument";
        public const string HeaderPartType = schemaOxml2006 + "header";
        public const string MainDocumentPartType = schemaOxml2006 + "officeDocument"; // same as Excel, PowerPoint
        public const string NumberingDefsPartType = schemaOxml2006 + "numbering";
        public const string StyleDefsPartType = schemaOxml2006 + "styles"; // same as Excel
        public const string WebSettingsPartType = schemaOxml2006 + "webSettings";
        public const string DocumentTemplatePartType = schemaOxml2006 + "attachedTemplate";
        public const string FramesetsPartType = schemaOxml2006 + "frame";
        public const string MasterSubDocumentsPartType = schemaOxml2006 + "subDocument";
        public const string MailMergeDataSourcePartType = schemaOxml2006 + "mailMergeSource";
        public const string MailMergeHeaderSourcePartType = schemaOxml2006 + "mailMergeHeaderSource";
        public const string XslTransformationPartType = schemaOxml2006 + "transform";

        // SpreadsheetML package relationship ids
        public const string CalcChainPartType = schemaOxml2006 + "calcChain";
        public const string ChartSheetPartType = schemaOxml2006 + "chartSheet";
        public const string ConnectionsPartType = schemaOxml2006 + "connections";
        public const string CustomPropertyPartType = schemaOxml2006 + "customProperty";
        public const string CustomXmlMappingsPartType = schemaOxml2006 + "xmlMaps";
        public const string DialogsheetPartType = schemaOxml2006 + "dialogSheet";
        public const string DrawingsPartType = schemaOxml2006 + "drawing";
        public const string ExternalWorkbookRefsPartType = schemaOxml2006 + "externalLink";
        public const string MetadataPartType = schemaOxml2006 + "sheetMetadata";
        public const string PivotTablePartType = schemaOxml2006 + "pivotTable";
        public const string PivotCacheDefPartType = schemaOxml2006 + "pivotCacheDefinition";
        public const string PivotTableCacheRecordsPartType = schemaOxml2006 + "pivotCacheRecords";
        public const string QueryTablePartType = schemaOxml2006 + "queryTable";
        public const string SharedStringsPartType = schemaOxml2006 + "sharedStrings";
        public const string SharedWorkbookRevisionHeadersPartType = schemaOxml2006 + "revisionHeaders";
        public const string SharedWorkbookRevisionLogPartType = schemaOxml2006 + "revisionLog";
        public const string SharedWorkbookUserDataPartType = schemaOxml2006 + "usernames";
        public const string SingleCellTableDefsPartType = schemaOxml2006 + "tableSingleCells";
        public const string TableDefsPartType = schemaOxml2006 + "table";
        public const string VolatileDependenciesPartType = schemaOxml2006 + "volatileDependencies";
        public const string WorksheetPartType = schemaOxml2006 + "worksheet";
        public const string ExternalWorkbooksPartType = schemaOxml2006 + "externalLinkPath";

        // PresentationML package relationship ids
        public const string CommentAuthorsPartType = schemaOxml2006 + "commentAuthors";
        public const string HandoutMasterPartType = schemaOxml2006 + "handoutMaster";
        public const string NotesMasterPartType = schemaOxml2006 + "notesMaster";
        public const string NotesSlidePartType = schemaOxml2006 + "notesSlide";
        public const string PresentationPropertiesPartType = schemaOxml2006 + "presProps";
        public const string SlidePartType = schemaOxml2006 + "slide";
        public const string SlideLayoutPartType = schemaOxml2006 + "slideLayout";
        public const string SlideMasterPartType = schemaOxml2006 + "slideMaster";
        public const string SlideSynchronizationDataPartType = schemaOxml2006 + "slideUpdateInfo";
        public const string UserDefinedTagsPartType = schemaOxml2006 + "tags";
        public const string ViewPropertiesPartType = schemaOxml2006 + "viewProps";
        public const string HtmlPublishLocationPartType = schemaOxml2006 + "htmlPubSaveAs";
        public const string SlideSynchronizationServerLocationPartType = schemaOxml2006 + "slideUpdateUrl";

        // DrawingML package relationship ids
        public const string ChartPartType = schemaOxml2006 + "chart";
        public const string ChartDrawingPartType = schemaOxml2006 + "chartUserShapes";
        public const string DiagramColorsPartType = schemaOxml2006 + "diagramColors";
        public const string DiagramDataPartType = schemaOxml2006 + "diagramData";
        public const string DiagramLayoutPartType = schemaOxml2006 + "diagramLayout";
        public const string DiagramStylePartType = schemaOxml2006 + "diagramQuickStyle";
        public const string ThemePartType = schemaOxml2006 + "theme";
        public const string ThemeOverridePartType = schemaOxml2006 + "themeOverride";
        public const string TableStylesPartType = schemaOxml2006 + "tableStyles";

        // SharedML package relationship ids
        public const string AudioPartType = schemaOxml2006 + "audio";
        public const string EmbeddedControlPartType = schemaOxml2006 + "control";
        public const string EmbeddedObjectPartType = schemaOxml2006 + "oleObject";
        public const string EmbeddedPackagePartType = schemaOxml2006 + "package";
        public const string CoreFilePropertiesPartType = schemaOxml2006 + "metadata/core-properties";
        public const string FontPartType = schemaOxml2006 + "font";
        public const string ImagePartType = schemaOxml2006 + "image";
        public const string PrinterSettingsPartType = schemaOxml2006 + "printerSettings";
        public const string ThumbnailPartType = schemaOxml2006 + "thumbnail";
        public const string VideoPartType = schemaOxml2006 + "video";
        public const string HyperlinkPartType = schemaOxml2006 + "hyperlink";
    }
}
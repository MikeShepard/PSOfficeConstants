#constants for Access based on https://docs.microsoft.com/en-us/office/vba/api/Access
$ac = [Ordered]@{

    #values from access.acaggregatetype
    #****************************
    acAggregateAverage                   =	2		# Average
    acAggregateCount                     =	6		# Count
    acAggregateDistinct                  =	5		# Distinct
    acAggregateMaximum                   =	4		# Maximum
    acAggregateMinimum                   =	3		# Minimum
    acAggregateNone                      =	0		# None
    acAggregateSum                       =	1		# Sum

    #values from access.acaxisrange
    #****************************
    acAxisRangeAuto                      =	0		# The represented range is determined automatically by the lowest and highest values in the set.
    acAxisRangedFixed                    =	1		# The represented range is determined by fixed minimum/maximum values and may be clipped accordingly.

    #values from access.acaxisunits
    #****************************
    acAxisUnitsBillions                  =	9		# Billions
    acAxisUnitsHundredBillions           =	11		# Hundred Billions
    acAxisUnitsHundredMillions           =	8		# Hundred Millions
    acAxisUnitsHundreds                  =	2		# Hundreds
    acAxisUnitsHundredThousands          =	5		# Hundred Thousands
    acAxisUnitsMillions                  =	6		# Millions
    acAxisUnitsNone                      =	0		# None (original values)
    acAxisUnitsPercentage                =	1		# Percentage
    acAxisUnitsTenBillions               =	10		# Ten Billions
    acAxisUnitsTenMillions               =	7		# Ten Millions
    acAxisUnitsTenThousands              =	4		# Ten Thousands
    acAxisUnitsThousands                 =	3		# Thousands
    acAxisUnitsTrillions                 =	12		# Trillions

    #values from access.acbrowsetoobjecttype
    #****************************
    acBrowseToForm                       =	2		# Open a form.
    acBrowseToReport                     =	3		# Open a report.

    #values from access.accharttype
    #****************************
    acChartBarClustered                  =	3		# Clustered Bar
    acChartBarStacked                    =	4		# Stacked Bar
    acChartBarStacked100                 =	5		# 100% Stacked Bar
    acChartColumnClustered               =	0		# Clustered Column
    acChartColumnStacked                 =	1		# Stacked Column
    acChartColumnStacked100              =	2		# 100% Stacked Column
    acChartCombo                         =	10		# Combo
    acChartLine                          =	6		# Line
    acChartLineStacked                   =	7		# Line Stacked
    acChartLineStacked100                =	8		# 100% Stacked Line
    acChartPie                           =	9		# Pie

    #values from access.acclosesave
    #****************************
    acSaveNo                             =	2		# The specified object is not saved.
    acSavePrompt                         =	0		# The user is asked whether or not they want to save the object.NOTE: This value is ignored if you are closing a Visual Basic module. The module will be closed, but changes to the module will not be saved.
    acSaveYes                            =	1		# The specified object is saved.

    #values from access.accolorindex
    #****************************
    acColorIndexAqua                     =	14		# Aqua color.
    acColorIndexBlack                    =	0		# Black color.
    acColorIndexBlue                     =	12		# Blue color.
    acColorIndexBrightGreen              =	10		# Bright green color.
    acColorIndexDarkBlue                 =	4		# Dark blue color.
    acColorIndexFuchsia                  =	13		# Fuchsia color.
    acColorIndexGray                     =	7		# Gray color.
    acColorIndexGreen                    =	2		# Green color.
    acColorIndexMaroon                   =	1		# Maroon color.
    acColorIndexOlive                    =	3		# Olive color.
    acColorIndexRed                      =	9		# Red color.
    acColorIndexSilver                   =	8		# Silver color.
    acColorIndexTeal                     =	6		# Teal color.
    acColorIndexViolet                   =	5		# Violet color.
    acColorIndexWhite                    =	15		# White color.
    acColorIndexYellow                   =	11		# Yellow color.

    #values from access.accommand
    #****************************
    acCmdAboutMicrosoftAccess            =	35

    #values from access.accontroltype
    #****************************
    acAttachment                         =	126		# Attachment control
    acBoundObjectFrame                   =	108		# BoundObjectFrame control
    acCheckBox                           =	106		# CheckBox control
    acComboBox                           =	111		# ComboBox control
    acCommandButton                      =	104		# CommandButton control
    acCustomControl                      =	119		# ActiveX control
    acEmptyCell                          =	127		# EmptyCell control
    acImage                              =	103		# Image control
    acLabel                              =	100		# Label control
    acLine                               =	102		# Line control
    acListBox                            =	110		# ListBox control
    acNavigationButton                   =	130		# NavigationButton control
    acNavigationControl                  =	129		# NavigationControl control
    acObjectFrame                        =	114		# Unbound ObjectFrame control
    acOptionButton                       =	105		# OptionButton control
    acOptionGroup                        =	107		# OptionGroup control
    acPage                               =	124		# Page control
    acPageBreak                          =	118		# PageBreak control
    acRectangle                          =	101		# Rectangle control
    acSubForm                            =	112		# SubForm control
    acTabCtl                             =	123		# Tab control
    acTextBox                            =	109		# TextBox control
    acToggleButton                       =	122		# ToggleButton control
    acWebBrowser                         =	128		# WebBrowserControl control

    #values from access.accurrentview
    #****************************
    acCurViewDatasheet                   =	2		# The object is in Datasheet view.
    acCurViewDesign                      =	0		# The object is in Design view.
    acCurViewFormBrowse                  =	1		# The object is in Form view.
    acCurViewLayout                      =	7		# The object is in Layout view.
    acCurViewPivotChart                  =	4		# The object is in PivotChart view.
    acCurViewPivotTable                  =	3		# The object is in PivotTable view.
    acCurViewPreview                     =	5		# The object is in Print Preview.
    acCurViewReportBrowse                =	6		# The object is in Report view.

    #values from access.accursoronhover
    #****************************
    acCursorOnHoverDefault               =	0		# The default cursor is displayed.
    acCursorOnHoverHyperlinkHand         =	1		# The hyperlink hand cursor is displayed.

    #values from access.acdashtype
    #****************************
    acDashTypeDash                       =	1		# Dash
    acDashTypeDashDot                    =	3		# Dash Dot
    acDashTypeDot                        =	2		# Dot
    acDashTypeSolid                      =	0		# Solid

    #values from access.acdataobjecttype
    #****************************
    acActiveDataObject                   =	-1		# The active object contains the record.
    acDataForm                           =	2		# A form contains the record.
    acDataFunction                       =	10		# A user-defined function contains the record (Microsoft Access project only).
    acDataQuery                          =	1		# A query contains the record.
    acDataReport                         =	3		# A report contains the record.
    acDataServerView                     =	7		# A server view contains the record (Microsoft Access project only).
    acDataStoredProcedure                =	9		# A stored procedure contains the record (Microsoft Access project only).
    acDataTable                          =	0		# A table contains the record.

    #values from access.acdatatransfertype
    #****************************
    acExport                             =	1		# The data is exported.
    acImport                             =	0		# (Default) The data is imported.
    acLink                               =	2		# The database is linked to the specified data source.

    #values from access.acdategrouptype
    #****************************
    acDateGroupDay                       =	4		# The date is grouped by day.
    acDateGroupMonth                     =	3		# The date is grouped by month.
    acDateGroupNone                      =	0		# No grouping is applied.
    acDateGroupQuarter                   =	2		# The date is grouped by quarter.
    acDateGroupYear                      =	1		# The date is grouped by year.

    #values from access.acdefreportview
    #****************************
    acDefViewPreview                     =	0		# The report opens in Print Preview.
    acDefViewReportBrowse                =	1		# The report opens in Report view.

    #values from access.acdefview
    #****************************
    acDefViewContinuous                  =	1		# The form displays multiple records (as many as will fit in the current window), each in its own copy of the form's detail section.
    acDefViewDatasheet                   =	2		# Displays the form fields arranged in rows and columns like a spreadsheet.
    acDefViewPivotChart                  =	4		# Displays the form as a PivotChart.
    acDefViewPivotTable                  =	3		# Displays the form as a PivotTable.
    acDefViewSingle                      =	0		# (Default) Displays one record at a time.
    acDefViewSplitForm                   =	5		# Displays the form in Split Form view.

    #values from access.acdisplayas
    #****************************
    acDisplayAsIcon                      =	1		# The attachment is displayed as the icon for that file type.
    acDisplayAsImageIcon                 =	0		# If the attachment is a supported image format, then the image is displayed. If the attachment is not a supported image format, then the icon for that file type is displayed.
    acDisplayAsPaperclip                 =	2		# A paper clip is displayed.

    #values from access.acdisplayashyperlink
    #****************************
    acDisplayAsHyperlinkAlways           =	1		# Always display the contents of the control as a hyperlink.
    acDisplayAsHyperlinkIfHlink          =	0		# Display the contents of the control as a hyperlink only when its contents are in the form of a Uniform Resource Locator (URL).
    acDisplayAsHyperlinkOnScreenOnly     =	2		# Display the contents of the control as a hyperlink only on the screen.

    #values from access.acexportquality
    #****************************
    acExportQualityPrint                 =	0		# The output is optimized for printing.
    acExportQualityScreen                =	1		# The output is optimized for onscreen display.

    #values from access.acexportxmlencoding
    #****************************
    acUTF16                              =	1		# UTF16 encoding.
    acUTF8                               =	0		# (Default) UTF8 encoding.

    #values from access.acexportxmlobjecttype
    #****************************
    acExportForm                         =	2		# Form
    acExportFunction                     =	10		# User-defined function (Microsoft Access project only)
    acExportQuery                        =	1		# Query
    acExportReport                       =	3		# Report
    acExportServerView                   =	7		# Server view (Microsoft Access project only)
    acExportStoredProcedure              =	9		# Stored procedure (Microsoft Access project only)
    acExportTable                        =	0		# Table

    #values from access.acexportxmlotherflags
    #****************************
    acEmbedSchema                        =	1		# Writes schema information into the document specified by the DataTarget argument; this value takes precedence over the SchemaTarget argument.
    acExcludePrimaryKeyAndIndexes        =	2		# Does not export primary key and index schema properties.
    acExportAllTableAndFieldProperties   =	32		# The exported schema contains properties of the table and its fields.
    acLiveReportSource                   =	8		# Creates a live link to a remote Microsoft SQL Server 2000 database. Valid only when you are exporting reports that are bound to a Microsoft SQL Server 2000 database.
    acPersistReportML                    =	16		# Persists the exported object's ReportML information.
    acRunFromServer                      =	4		# Creates an Active Server Pages (ASP) wrapper; otherwise, default is an HTML wrapper. Applies only when you are exporting reports.

    #values from access.acexportxmlschemaformat
    #****************************
    acSchemaNone                         =	0

    #values from access.acfileformat
    #****************************
    acFileFormatAccess12                 =	12		# Microsoft Access 2007 format
    acFileFormatAccess2                  =	2		# Microsoft Access 2.0 format
    acFileFormatAccess2000               =	9		# Microsoft Access 2000 format
    acFileFormatAccess2002               =	10		# Microsoft Access 2002 format
    acFileFormatAccess95                 =	7		# Microsoft Access 95 format
    acFileFormatAccess97                 =	8		# Microsoft Access 97 format

    #values from access.acfiltertype
    #****************************
    acFilterNormal                       =	0

    #values from access.acfindfield
    #****************************
    acAll                                =	0		# Search in all fields in each record.
    acCurrent                            =	-1		# Confine the search to the current field.

    #values from access.acfindmatch
    #****************************
    acAnywhere                           =	0		# Search for data in any part of the field.
    acEntire                             =	1		# Search for data that fills the entire field.
    acStart                              =	2		# Search for data located at the beginning of the field.

    #values from access.acformatbarlimits
    #****************************
    acAutomatic                          =	0		# For the shortest bar, the lowest value is used. For the highest bar, the highest value is used.
    acNumber                             =	1		# A number is used.
    acPercent                            =	2		# A percentage is used.

    #values from access.acformatconditionoperator
    #****************************
    acBetween                            =	0		# The value must be between the values specified by the Expression1 and Expression2 arguments.
    acEqual                              =	2		# The value must equal to the value specified by the Expression1 argument.
    acGreaterThan                        =	4		# The value must be greater than the value specified by the Expression1 argument.
    acGreaterThanOrEqual                 =	6		# The value must be greater than or equal to the value specified by the Expression1 argument.
    acLessThan                           =	5		# The value must be less than the value specified by the Expression1 argument.
    acLessThanOrEqual                    =	7		# The value must be less than or equal to the value specified by the Expression1 argument.
    acNotBetween                         =	1		# The value must not be between the values specified by the Expression1 and Expression2 arguments.
    acNotEqual                           =	3		# The value must not be equal to the value specified by the Expression1 argument.

    #values from access.acformatconditiontype
    #****************************
    acDataBar                            =	3		# The conditional format is displayed as a data bar.
    acExpression                         =	1		# The conditional format is based on an expression.
    acFieldHasFocus                      =	2		# The conditional format is based on the value of the control that has focus on a form.
    acFieldValue                         =	0		# The conditional format is based on values in the selected control.

    #values from access.acformopendatamode
    #****************************
    acFormAdd                            =	0		# The user can add new records but can't edit existing records.
    acFormEdit                           =	1		# The user can edit existing records and add new records.
    acFormPropertySettings               =	-1		# The user can only change the form's properties.
    acFormReadOnly                       =	2		# The user can only view records.

    #values from access.acformview
    #****************************
    acDesign                             =	1		# The form opens in Design view.
    acFormDS                             =	3		# The form opens in Datasheet view.
    acFormPivotChart                     =	5		# The form opens in PivotChart view.
    acFormPivotTable                     =	4		# The form opens in PivotTable view.
    acLayout                             =	6		# The form opens in Layout view.
    acNormal                             =	0		# (Default) The form opens in Form view.
    acPreview                            =	2		# The form opens in Print Preview.

    #values from access.achorizontalanchor
    #****************************
    acHorizontalAnchorBoth               =	2		# The control is stretched horizontally across its layout.
    acHorizontalAnchorLeft               =	0		# The control is anchored to the left side of its layout.
    acHorizontalAnchorRight              =	1		# The control is anchored to the right side of its layout.

    #values from access.achyperlinkpart
    #****************************
    acAddress                            =	2		# The address part of a Hyperlink field.
    acDisplayedValue                     =	0		# The underlined text displayed in a hyperlink.
    acDisplayText                        =	1		# The displaytext part of a Hyperlink field.
    acFullAddress                        =	5		# The address and subaddress parts of a Hyperlink field, delimited by a &quot;#&quot; character.
    acScreenTip                          =	4		# The tooltip part of a Hyperlink field.
    acSubAddress                         =	3		# The subaddress part of a Hyperlink field.

    #values from access.acimemode
    #****************************
    acImeModeAlpha                       =	8		# Activates the IME in half-width Latin mode.
    acImeModeAlphaFull                   =	7		# Activates the IME in full-width Latin mode.
    acImeModeDisable                     =	3		# Disables the IME.
    acImeModeHangul                      =	10		# Activates the IME in half-width Hangul mode.
    acImeModeHangulFull                  =	9		# Activates the IME in full-width Hangul mode.
    acImeModeHiragana                    =	4		# Activates the IME in full-width hiragana mode.
    acImeModeKatakana                    =	5		# Activates the IME in full-width katakana mode.
    acImeModeKatakanaHalf                =	6		# Activates the IME in half-width katakana mode.
    acImeModeNoControl                   =	0		# Does not change the IME mode.
    acImeModeOff                         =	2		# Disables the IME and activates Latin text entry.
    acImeModeOn                          =	1		# Activates the IME.

    #values from access.acimesentencemode
    #****************************


    #values from access.acimportxmloption
    #****************************
    acAppendData                         =	2		# Imports the data into an existing table.
    acStructureAndData                   =	1		# Imports the data into a new table based on the structure of the specified XML file.
    acStructureOnly                      =	0		# Creates a new table based on the structure of the specified XML file.

    #values from access.aclayouttype
    #****************************
    acLayoutNone                         =	0		# The control is not part of a layout.
    acLayoutStacked                      =	2		# The control is part of a stacked layout.
    acLayoutTabular                      =	1		# The control is part of a tabular layout.

    #values from access.aclegendposition
    #****************************
    acLegendPositionBottom               =	3		# Bottom edge of the chart
    acLegendPositionLeft                 =	0		# Left edge of the chart
    acLegendPositionRight                =	2		# Right edge of the chart
    acLegendPositionTop                  =	1		# Top edge of the chart

    #values from access.acmarkertype
    #****************************
    acMarkerAsterisk                     =	4		# Asterisk
    acMarkerCircle                       =	8		# Circle
    acMarkerDiamond                      =	2		# Diamond
    acMarkerLongDash                     =	7		# Long Dash
    acMarkerNone                         =	0		# None
    acMarkerPlus                         =	9		# Plus
    acMarkerShortDash                    =	6		# Short Dash
    acMarkerSquare                       =	1		# Square
    acMarkerTriangle                     =	3		# Triangle
    acMarkerX                            =	4		# X

    #values from access.acmissingdatapolicy
    #****************************
    acDoNotPlot                          =	1		# Do not plot.
    acPlotAsInterpolated                 =	2		# Plot as interpolated.
    acPlotAsZero                         =	0		# Plot as zero.

    #values from access.acmoduletype
    #****************************
    acClassModule                        =	1		# The specified module is a class module.
    acStandardModule                     =	0		# The specified module is a standard module.

    #values from access.acnavigationspan
    #****************************
    acHorizontal                         =	0		# The navigation buttons are displayed horizontally.
    acVertical                           =	1		# The navigation buttons are displayed vertically.

    #values from access.acnewdatabaseformat
    #****************************
    acNewDatabaseFormatAccess12          =	12		# Create a database in the Access (.accdb) file format.
    acNewDatabaseFormatAccess2000        =	9		# Create a database in the Microsoft Access 2000 (.mdb) file format.
    acNewDatabaseFormatAccess2002        =	10		# Create a database in the Microsoft Access 2002-2003 (.mdb) file format.
    acNewDatabaseFormatUserDefault       =	0		# Create a database in the default file format.

    #values from access.acobjecttype
    #****************************
    acDatabaseProperties                 =	11		# Database property
    acDefault                            =	-1

    #values from access.acopendatamode
    #****************************
    acAdd                                =	0		# The user can add new records but can't view or edit existing records.
    acEdit                               =	1		# The user can view or edit existing records and add new records.
    acReadOnly                           =	2		# The user can only view records.

    #values from access.acoutputobjecttype
    #****************************
    acOutputForm                         =	2		# Form
    acOutputFunction                     =	10		# User-Defined Function
    acOutputModule                       =	5		# Module
    acOutputQuery                        =	1		# Query
    acOutputReport                       =	3		# Report
    acOutputServerView                   =	7		# Server View
    acOutputStoredProcedure              =	9		# Stored Procedure
    acOutputTable                        =	0		# Table

    #values from access.acpicturecaptionarrangement
    #****************************
    acBottom                             =	3		# The caption appears below the picture.
    acGeneral                            =	1		# The caption uses the General alignment setting.
    acLeft                               =	4		# The caption appears to the left of the picture.
    acNoPictureCaption                   =	0		# The caption is not displayed.
    acRight                              =	5		# The caption appears to the right of the picture.
    acTop                                =	2		# The caption appears above the picture.

    #values from access.acprintcolor
    #****************************
    acPRCMColor                          =	2		# The printer should print in color.
    acPRCMMonochrome                     =	1		# The printer should print in monochrome.

    #values from access.acprintduplex
    #****************************
    acPRDPHorizontal                     =	2		# Double-sided printing using a horizontal page turn.
    acPRDPSimplex                        =	1		# Single-sided printing with the current orientation setting.
    acPRDPVertical                       =	3		# Double-sided printing using a vertical page turn.

    #values from access.acprintitemlayout
    #****************************
    acPRHorizontalColumnLayout           =	1953		# Columns are laid across, then down.
    acPRVerticalColumnLayout             =	1954		# Columns are laid down, then across.

    #values from access.acprintobjquality
    #****************************
    acPRPQDraft                          =	-1		# The printer prints in draft quality.
    acPRPQHigh                           =	-4		# The printer prints in high quality.
    acPRPQLow                            =	-2		# The printer prints in low quality.
    acPRPQMedium                         =	-3		# The printer prints in medium quality.

    #values from access.acprintorientation
    #****************************
    acPRORLandscape                      =	2		# Landscape orientation
    acPRORPortrait                       =	1		# Portrait orientation

    #values from access.acprintpaperbin
    #****************************
    acPRBNAuto                           =	7		# (Default) Use paper from the current default bin.
    acPRBNCassette                       =	14		# Use paper from the attached cassette cartridge.
    acPRBNEnvelope                       =	5		# Use envelopes from the envelope feeder.
    acPRBNEnvManual                      =	6		# Use envelopes from the envelope feeder, but wait for manual insertion.
    acPRBNFormSource                     =	15		# Use paper from the forms bin.
    acPRBNLargeCapacity                  =	11		# Use paper from the large capacity feeder.
    acPRBNLargeFmt                       =	10		# Use paper from the large paper bin.
    acPRBNLower                          =	2		# Use paper from the lower bin.
    acPRBNManual                         =	4		# Wait for manual insertion of each sheet of paper.
    acPRBNMiddle                         =	3		# Use paper from the middle bin.
    acPRBNSmallFmt                       =	9		# Use paper from the small paper feeder.
    acPRBNTractor                        =	8		# Use paper from the tractor feeder.
    acPRBNUpper                          =	1		# Use paper from the upper bin.

    #values from access.acprintpapersize
    #****************************
    acPRPS10x14                          =	16		# 10 x 14 in.
    acPRPS11x17                          =	17		# 11 x 17 in.
    acPRPSA3                             =	8		# A3 (297 mm x 420 mm)
    acPRPSA4                             =	9		# A4 (210 mm x 297 mm)
    acPRPSA4Small                        =	10		# A4 Small (210 mm x 297 mm)
    acPRPSA5                             =	11		# A5 (148 mm x 210 mm)
    acPRPSB4                             =	12		# B4 (250 mm x 354 mm)
    acPRPSB5                             =	13		# B5 (148 mm x 210 mm)
    acPRPSCSheet                         =	24		# C size sheet
    acPRPSDSheet                         =	25		# D size sheet
    acPRPSEnv10                          =	20		# Envelope #10 (4-1/8 in. x 9-1/2 in.)
    acPRPSEnv11                          =	21		# Envelope #11 (4-1/2 in. x 10-3/8 in.)
    acPRPSEnv12                          =	22		# Envelope #12 (4-1/2 in. x 11 in.)
    acPRPSEnv14                          =	23		# Envelope #14 (5 in. x 11-1/2 in.)
    acPRPSEnv9                           =	19		# Envelope #9 (3-7/8 in. x 8-7/8 in.)
    acPRPSEnvB4                          =	33		# Envelope B4 (250 mm x 353 mm)
    acPRPSEnvB5                          =	34		# Envelope B5 (176 mm x 250 mm)
    acPRPSEnvB6                          =	35		# Envelope B6 (176 mm x 125 mm)
    acPRPSEnvC3                          =	29		# Envelope C3 (324 mm x 458 mm)
    acPRPSEnvC4                          =	30		# Envelope C4 (229 mm x 324 mm)
    acPRPSEnvC5                          =	28		# Envelope C5 (162 mm x 229 mm)
    acPRPSEnvC6                          =	31		# Envelope C6 (114 mm x 162 mm)
    acPRPSEnvC65                         =	32		# Envelope C65 (114 mm x 229 mm)
    acPRPSEnvDL                          =	27		# Envelope DL (110 mm x 220 mm)
    acPRPSEnvItaly                       =	36		# Italian envelope (110 mm x 230 mm)
    acPRPSEnvMonarch                     =	37		# Monarch envelope (3-7/8 in. x 7-1/2 in.)
    acPRPSEnvPersonal                    =	38		# Envelope (3-5/8 in. x 6-1/2 in.)
    acPRPSESheet                         =	26		# E size sheet
    acPRPSExecutive                      =	7		# Executive (7-1/2 in. x 10-1/2 in.)
    acPRPSFanfoldLglGerman               =	41		# German Legal Fanfold (8-1/2 in. x 13 in.)
    acPRPSFanfoldStdGerman               =	40		# German Standard Fanfold (8-1/2 in. x 12 in.)
    acPRPSFanfoldUS                      =	39		# U.S. Standard Fanfold (14-7/8 in. x 11 in.)
    acPRPSFolio                          =	14		# Folio (8-1/2 in. x 13 in.)
    acPRPSLedger                         =	4		# Ledger (17 in. x 11 in.)
    acPRPSLegal                          =	5		# Legal (8-1/2 in. x 14 in.)
    acPRPSLetter                         =	1		# Letter (8-1/2 in. x 11 in.)
    acPRPSLetterSmall                    =	2		# Letter Small (8-1/2 in. x 11 in.)
    acPRPSNote                           =	18		# Note (8-1/2 in. x 11 in.)
    acPRPSQuarto                         =	15		# Quarto (215 mm x 275 mm)
    acPRPSStatement                      =	6		# Statement (5-1/2 in. x 8-1/2 in.)
    acPRPSTabloid                        =	3		# Tabloid (11 in. x 17 in.)
    acPRPSUser                           =	256		# User-defined

    #values from access.acprintquality
    #****************************
    acDraft                              =	3		# Draft quality
    acHigh                               =	0		# (Default) High quality
    acLow                                =	2		# Low quality
    acMedium                             =	1		# Medium quality

    #values from access.acprintrange
    #****************************
    acPages                              =	2		# A specific range of pages will be printed. Use the PageFrom and PageTo arguments to specify the range of pages to print.
    acPrintAll                           =	0		# Prints all of the object.
    acSelection                          =	1		# Prints the selected part of the object.

    #values from access.acprojecttype
    #****************************
    acADP                                =	1		# The current project is a Microsoft Access project.
    acMDB                                =	2		# The current project is a Microsoft Access database.
    acNull                               =	0

    #values from access.acproperty
    #****************************
    acPropertyBackColor                  =	8		# Set the BackColor property.
    acPropertyCaption                    =	9		# Set the Caption property.
    acPropertyEnabled                    =	0		# Set the Enabled property.
    acPropertyForeColor                  =	7		# Set the ForeColor property.
    acPropertyHeight                     =	6		# Set the Height property.
    acPropertyLeft                       =	3		# Set the Left property.
    acPropertyLocked                     =	2		# Set the Locked property.
    acPropertyTop                        =	4		# Set the Top property.
    acPropertyValue                      =	10		# Set the Value property.
    acPropertyVisible                    =	1		# Set the Visible property.
    acPropertyWidth                      =	5		# Set the Width property.

    #values from access.acquitoption
    #****************************
    acQuitPrompt                         =	0		# Displays a dialog box that asks whether you want to save any database objects that have been changed but not saved.
    acQuitSaveAll                        =	1		# (Default) Saves all objects without displaying a dialog box.
    acQuitSaveNone                       =	2		# Quits Microsoft Access without saving any objects.

    #values from access.acrecord
    #****************************
    acFirst                              =	2		# Make the first record the current record.
    acGoTo                               =	4		# Make the specified record the current record.
    acLast                               =	3		# Make the last record the current record.
    acNewRec                             =	5		# Make a new record the current record.
    acNext                               =	1		# Make the next record the current record.
    acPrevious                           =	0		# Make the previous record the current record.

    #values from access.acresourcetype
    #****************************
    acResourceImage                      =	1		# Image.
    acResourceTheme                      =	0		# Office theme.

    #values from access.acsearchdirection
    #****************************
    acDown                               =	1		# Search all records below the current record.
    acSearchAll                          =	2		# Search all records.
    acUp                                 =	0		# Search all records above the current record.

    #values from access.acsection
    #****************************
    acDetail                             =	0		# (Default) Detail section
    acFooter                             =	2		# Form or report footer
    acGroupLevel1Footer                  =	6		# Group-level 1 footer (reports only)
    acGroupLevel1Header                  =	5		# Group-level 1 header (reports only)
    acGroupLevel2Footer                  =	8		# Group-level 2 footer (reports only)
    acGroupLevel2Header                  =	7		# Group-level 2 header (reports only)
    acHeader                             =	1		# Form or report header
    acPageFooter                         =	4		# Page footer
    acPageHeader                         =	3		# Page header

    #values from access.acsendobjecttype
    #****************************
    acSendForm                           =	2		# Send a Form.
    acSendModule                         =	5		# Send a Module.
    acSendNoObject                       =	-1		# (Default) Don't send a database object.
    acSendQuery                          =	1		# Send a Query.
    acSendReport                         =	3		# Send a Report.
    acSendTable                          =	0		# Send a Table.

    #values from access.acseparatorcharacters
    #****************************
    acSeparatorCharactersComma           =	3		# A comma (,) is used as the separator character.
    acSeparatorCharactersNewLine         =	1		# Each value appears on its own line.
    acSeparatorCharactersSemiColon       =	2		# A semicolon (;) is used as the separator character.
    acSeparatorCharactersSystemSeparator	=	0		# The List separator setting in the Regional and Language Options in the Windows Control Panel is used as the separator character.

    #values from access.acsharepointlisttransfertype
    #****************************
    acImportSharePointList               =	0		# Import the SharePoint list.
    acLinkSharePointList                 =	1		# Link to the SharePoint list.

    #values from access.acshowtoolbar
    #****************************
    acToolbarNo                          =	2		# Hide the toolbar.
    acToolbarWhereApprop                 =	1		# Display the toolbar while in the appropriate view.
    acToolbarYes                         =	0		# Display the toolbar.

    #values from access.acsplitformdatasheet
    #****************************
    acDatasheetAllowEdits                =	0		# The user can edit the contents of the datasheet.
    acDatasheetReadOnly                  =	1		# The user cannot edit the contents of the datasheet.

    #values from access.acsplitformorientation
    #****************************
    acDatasheetOnBottom                  =	1		# The datasheet is displayed below the form.
    acDatasheetOnLeft                    =	2		# The datasheet is displayed to the left of the form.
    acDatasheetOnRight                   =	3		# The datasheet is displayed to the right of the form.
    acDatasheetOnTop                     =	0		# The datasheet is displayed above the form.

    #values from access.acsplitformprinting
    #****************************
    acFormOnly                           =	0		# The contents of the form are printed.
    acGridOnly                           =	1		# The contents of the datasheet are printed.

    #values from access.acspreadsheettype
    #****************************
    acSpreadsheetTypeExcel3              =	0		# Microsoft Excel 3.0 format
    acSpreadsheetTypeExcel4              =	6		# Microsoft Excel 4.0 format
    acSpreadsheetTypeExcel5              =	5		# Microsoft Excel 5.0 format
    acSpreadsheetTypeExcel7              =	5		# Microsoft Excel 95 format
    acSpreadsheetTypeExcel8              =	8		# Microsoft Excel 97 format
    acSpreadsheetTypeExcel9              =	8		# Microsoft Excel 2000 format
    acSpreadsheetTypeExcel12             =	9		# Microsoft Excel 2010 format
    acSpreadsheetTypeExcel12Xml          =	10		# Microsoft Excel 2010/2013/2016 XML format (.xlsx, .xlsm, .xlsb)

    #values from access.acsyscmdaction
    #****************************
    acSysCmdAccessDir                    =	9		# Returns the name of the directory where Msaccess.exe is located.
    acSysCmdAccessVer                    =	7		# Returns the version number of Microsoft Access.
    acSysCmdClearHelpTopic               =	11

    #values from access.actextformat
    #****************************
    acTextFormatHTMLRichText             =	1		# Rich text can be displayed.
    acTextFormatPlain                    =	0		# (Default) Plain text is displayed.

    #values from access.actexttransfertype
    #****************************
    acExportDelim                        =	2		# Export Delimited
    acExportFixed                        =	3		# Export Fixed Width
    acExportHTML                         =	8		# Export HTML
    acExportMerge                        =	4		# Export Microsot Word Merge
    acImportDelim                        =	0		# Import Delimited
    acImportFixed                        =	1		# Import Fixed Width
    acImportHTML                         =	7		# Import HTML
    acLinkDelim                          =	5		# Link Delimited
    acLinkFixed                          =	6		# Link Fixed Width
    acLinkHTML                           =	9		# Link HTML

    #values from access.actransformxmlscriptoption
    #****************************
    acDisableScript                      =	2		# The script is disabled.
    acEnableScript                       =	0		# The script is enabled.
    acPromptScript                       =	1		# The user is prompted to disable or enable the script.

    #values from access.actrendlineoptions
    #****************************
    acTrendlineExponential               =	2		# Exponential
    acTrendlineLinear                    =	1		# Linear
    acTrendlineLogarithmic               =	3		# Logarithmic
    acTrendlineMovingAverage             =	6		# Moving Average
    acTrendlineNone                      =	0		# None
    acTrendlinePolynomial                =	4		# Polynomial
    acTrendlinePower                     =	5		# Power

    #values from access.acvalueaxis
    #****************************
    acPrimaryAxis                        =	0		# Primary axis
    acSecondaryAxis                      =	1		# Secondary axis

    #values from access.acverticalanchor
    #****************************
    acVerticalAnchorBoth                 =	2		# The control is stretched vertically across its layout.
    acVerticalAnchorBottom               =	1		# The control is anchored at the bottom of its layout.
    acVerticalAnchorTop                  =	0		# The control is anchored at the top of its layout.

    #values from access.acview
    #****************************
    acViewDesign                         =	1		# Design view
    acViewLayout                         =	6		# Layout view
    acViewNormal                         =	0		# (Default) Normal view
    acViewPivotChart                     =	4		# PivotChart view
    acViewPivotTable                     =	3		# PivotTable view
    acViewPreview                        =	2		# Print Preview
    acViewReport                         =	5		# Report view

    #values from access.acwebbrowserscrollbars
    #****************************
    acScrollAuto                         =	0		# Scroll bars are displayed if the current page in the control is too large to be displayed in its entirely.
    acScrollNo                           =	2		# Scroll bars are not displayed.
    acScrollYes                          =	1		# Scroll bars are displayed.

    #values from access.acwebbrowserstate
    #****************************
    acComplete                           =	3		# The web browser control has finished loading the new document and all its contents.
    acInteractive                        =	3		# The web browser control has loaded enough of the document to allow limited user interaction, such as choosing hyperlinks that have been displayed.
    acLoaded                             =	2		# The web browser control has loaded and initialized the new document, but has not yet received all the document data.
    acLoading                            =	1		# The web browser control is loading a new document.
    acUninitialized                      =	0		# No document is currently loaded.

    #values from access.acwebuserdisplay
    #****************************
    acWebUserEmail                       =	3		# The current user's email address.
    acWebUserID                          =	0		# The current user's member ID.
    acWebUserLoginName                   =	2		# The current user's login name.
    acWebUserName                        =	1		# The current user's display name.

    #values from access.acwebusergroupsdisplay
    #****************************
    acWebUserGroupID                     =	0		# The identifiers of the groups.
    acWebUserGroupName                   =	1		# The names of the groups.

    #values from access.acwindowmode
    #****************************
    acDialog                             =	3		# The form or report's Modal and PopUp properties are set to Yes.
    acHidden                             =	1		# The form or report is hidden.
    acIcon                               =	2		# The form or report opens minimized in the Windows taskbar.
    acWindowNormal                       =	0		# (Default) The form or report opens in the mode set by its properties.
}
#End Enum

$ac = new-object PSCustomObject -Property $ac


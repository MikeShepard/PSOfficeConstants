#constants for Excel based on https://docs.microsoft.com/en-us/office/vba/api/Excel(enumerations)
$xl = [Ordered]@{

    #values from excel.constants
    #****************************
    xl3DBar                                   =	-4099		# 3D Bar
    xl3DEffects1                              =	13		# 3D Effects1
    xl3DEffects2                              =	14		# 3D Effects2
    xl3DSurface                               =	-4103		# 3D Surface
    xlAbove                                   =	0		# Above
    xlAccounting1                             =	4		# Accounting1
    xlAccounting2                             =	5		# Accounting2
    xlAccounting4                             =	17		# Accounting4
    xlAdd                                     =	2		# Add
    xlAll                                     =	-4104		# All
    xlAccounting3                             =	6		# Accounting3
    xlAllExceptBorders                        =	7		# All Except Borders
    xlAutomatic                               =	-4105		# Automatic
    xlBar                                     =	2		# Automatic
    xlBelow                                   =	1		# Below
    xlBidi                                    =	-5000		# Bidi
    xlBidiCalendar                            =	3		# BidiCalendar
    xlBoth                                    =	1		# Both
    xlBottom                                  =	-4107		# Bottom
    xlCascade                                 =	7		# Cascade
    xlCenter                                  =	-4108		# Center
    xlCenterAcrossSelection                   =	7		# Center Across Selection
    xlChart4                                  =	2		# Chart 4
    xlChartSeries                             =	17		# Chart Series
    xlChartShort                              =	6		# Chart Short
    xlChartTitles                             =	18		# Chart Titles
    xlChecker                                 =	9		# Checker
    xlCircle                                  =	8		# Circle
    xlClassic1                                =	1		# Classic1
    xlClassic2                                =	2		# Classic2
    xlClassic3                                =	3		# Classic3
    xlClosed                                  =	3		# Closed
    xlColor1                                  =	7		# Color1
    xlColor2                                  =	8		# Color2
    xlColor3                                  =	9		# Color3
    xlColumn                                  =	3		# Column
    xlCombination                             =	-4111		# Combination
    xlComplete                                =	4		# Complete
    xlConstants                               =	2		# Constants
    xlContents                                =	2		# Contents
    xlContext                                 =	-5002		# Context
    xlCorner                                  =	2		# Corner
    xlCrissCross                              =	16		# CrissCross
    xlCross                                   =	4		# Cross
    xlCustom                                  =	-4114		# Custom
    xlDebugCodePane                           =	13		# Debug Code Pane
    xlDefaultAutoFormat                       =	-1		# Default Auto Format
    xlDesktop                                 =	9		# Desktop
    xlDiamond                                 =	2		# Diamond
    xlDirect                                  =	1		# Direct
    xlDistributed                             =	-4117		# Distributed
    xlDivide                                  =	5		# Divide
    xlDoubleAccounting                        =	5		# Double Accounting
    xlDoubleClosed                            =	5		# Double Closed
    xlDoubleOpen                              =	4		# Double Open
    xlDoubleQuote                             =	1		# Double Quote
    xlDrawingObject                           =	14		# Drawing Object
    xlEntireChart                             =	20		# Entire Chart
    xlExcelMenus                              =	1		# Excel Menus
    xlExtended                                =	3		# Extended
    xlFill                                    =	5		# Fill
    xlFirst                                   =	0		# First
    xlFixedValue                              =	1		# Fixed Value
    xlFloating                                =	5		# Floating
    xlFormats                                 =	-4122		# Formats
    xlFormula                                 =	5		# Formula
    xlFullScript                              =	1		# Full Script
    xlGeneral                                 =	1		# General
    xlGray16                                  =	17		# Gray16
    xlGray25                                  =	-4124		# Gray25
    xlGray50                                  =	-4125		# Gray50
    xlGray75                                  =	-4126		# Gray75
    xlGray8                                   =	18		# Gray8
    xlGregorian                               =	2		# Gregorian
    xlGrid                                    =	15		# Grid
    xlGridline                                =	22		# Gridline
    xlHigh                                    =	-4127		# High
    xlHindiNumerals                           =	3		# Hindi Numerals
    xlIcons                                   =	1		# Icons
    xlImmediatePane                           =	12		# Immediate Pane
    xlInside                                  =	2		# Inside
    xlInteger                                 =	2		# Integer
    xlJustify                                 =	-4130		# Justify
    xlLast                                    =	1		# Last
    xlLastCell                                =	11		# Last Cell
    xlLatin                                   =	-5001		# Latin
    xlLeft                                    =	-4131		# Left
    xlLeftToRight                             =	2		# Left To Right
    xlLightDown                               =	13		# Light Down
    xlLightHorizontal                         =	11		# Light Horizontal
    xlLightUp                                 =	14		# Light Up
    xlLightVertical                           =	12		# Light Vertical
    xlList1                                   =	10		# List1
    xlList2                                   =	11		# List2
    xlList3                                   =	12		# List3
    xlLocalFormat1                            =	15		# Local Format1
    xlLocalFormat2                            =	16		# Local Format2
    xlLogicalCursor                           =	1		# Logical Cursor
    xlLong                                    =	3		# Long
    xlLotusHelp                               =	2		# Lotus Help
    xlLow                                     =	-4134		# Low
    xlLTR                                     =	-5003		# LTR
    xlMacrosheetCell                          =	7		# MacrosheetCell

    xlMaximum                                 =	2		# Maximum
    xlMinimum                                 =	4		# Minimum
    xlMinusValues                             =	3		# Minus Values
    xlMixed                                   =	2		# Mixed
    xlMixedAuthorizedScript                   =	4		# Mixed Authorized Script
    xlMixedScript                             =	3		# Mixed Script
    xlModule                                  =	-4141		# Module
    xlMultiply                                =	4		# Multiply
    xlNarrow                                  =	1		# Narrow
    xlNextToAxis                              =	4		# Next To Axis
    xlNoDocuments                             =	3		# No Documents
    xlNone                                    =	-4142		# None
    xlNotes                                   =	-4144		# Notes
    xlOff                                     =	-4146		# Off
    xlOn                                      =	1		# On
    xlOpaque                                  =	3		# Opaque
    xlOpen                                    =	2		# Open
    xlOutside                                 =	3		# Outside
    xlPartial                                 =	3		# Partial
    xlPartialScript                           =	2		# Partial Script
    xlPercent                                 =	2		# Percent
    xlPlus                                    =	9		# Plus
    xlPlusValues                              =	2		# Plus Values
    xlReference                               =	4		# Reference
    xlRight                                   =	-4152		# Right
    xlRTL                                     =	-5004		# RTL
    xlScale                                   =	3		# Scale
    xlSemiautomatic                           =	2		# Semiautomatic
    xlSemiGray75                              =	10		# SemiGray75
    xlShort                                   =	1		# Short
    xlShowLabel                               =	4		# Show Label
    xlShowLabelAndPercent                     =	5		# Show Label and Percent
    xlShowPercent                             =	3		# Show Percent
    xlShowValue                               =	2		# Show Value
    xlSimple                                  =	-4154		# Simple
    xlSingle                                  =	2		# Single
    xlSingleAccounting                        =	4		# Single Accounting
    xlSingleQuote                             =	2		# Single Quote
    xlSolid                                   =	1		# Solid
    xlSquare                                  =	1		# Square
    xlStar                                    =	5		# Star
    xlStError                                 =	4		# St Error
    xlStrict                                  =	2		# Strict
    xlSubtract                                =	3		# Subtract
    xlSystem                                  =	1		# System
    xlTextBox                                 =	16		# Text Box
    xlTiled                                   =	1		# Tiled
    xlTitleBar                                =	8		# Title Bar
    xlToolbar                                 =	1		# Toolbar
    xlToolbarButton                           =	2		# Toolbar Button
    xlTop                                     =	-4160		# Top
    xlTopToBottom                             =	1		# Top To Bottom
    xlTransparent                             =	2		# Transparent
    xlTriangle                                =	3		# Triangle
    xlVeryHidden                              =	2		# Very Hidden
    xlVisible                                 =	12		# Visible
    xlVisualCursor                            =	2		# Visual Cursor
    xlWatchPane                               =	11		# Watch Pane
    xlWide                                    =	3		# Wide
    xlWorkbookTab                             =	6		# Workbook Tab
    xlWorksheet4                              =	1		# Worksheet4
    xlWorksheetCell                           =	3		# Worksheet Cell
    xlWorksheetShort                          =	5		# Worksheet Short

    #values from excel.xlabovebelow
    #****************************
    xlAboveAverage                            =	0		# Above average
    xlAboveStdDev                             =	4		# Above standard deviation
    xlBelowAverage                            =	1		# Below average
    xlBelowStdDev                             =	5		# Below standard deviation
    xlEqualAboveAverage                       =	2		# Equal above average
    xlEqualBelowAverage                       =	3		# Equal below average

    #values from excel.xlactiontype
    #****************************
    xlActionTypeDrillthrough                  =	256		# Drill through
    xlActionTypeReport                        =	128		# Report
    xlActionTypeRowset                        =	16		# Rowset
    xlActionTypeUrl                           =	1		# URL

    #values from excel.xlallocation
    #****************************
    xlAutomaticAllocation                     =	2		# Calculate changes automatically after each value is changed.
    xlManualAllocation                        =	1		# Calculate changes manually.

    #values from excel.xlallocationmethod
    #****************************
    xlEqualAllocation                         =	1		# Use equal allocation.
    xlWeightedAllocation                      =	2		# Use weighted allocation.

    #values from excel.xlallocationvalue
    #****************************
    xlAllocateIncrement                       =	2		# Increment based on the old value.
    xlAllocateValue                           =	1		# The value entered divided by the number of allocations.

    #values from excel.xlapplicationinternational
    #****************************
    xl24HourClock                             =	33		# True if you are using 24-hour time; False if you are using 12-hour time.
    xl4DigitYears                             =	43		# True if you are using four-digit years; False if you are using two-digit years.
    xlAlternateArraySeparator                 =	16		# Alternate array item separator to be used if the current array separator is the same as the decimal separator.
    xlColumnSeparator                         =	14		# Character used to separate columns in array literals.
    xlCountryCode                             =	1		# Country/Region version of Microsoft Excel.
    xlCountrySetting                          =	2		# Current country/region setting in the Windows Control Panel.
    xlCurrencyBefore                          =	37		# True if the currency symbol precedes the currency values; False if it follows them.
    xlCurrencyCode                            =	25		# Currency symbol.
    xlCurrencyDigits                          =	27		# Number of decimal digits to be used in currency formats.
    xlCurrencyLeadingZeros                    =	40		# True if leading zeros are displayed for zero currency values.
    xlCurrencyMinusSign                       =	38		# True if you are using a minus sign for negative numbers; False if you are using parentheses.
    xlCurrencyNegative                        =	28		# Currency format for negative currency values:0 = (symbolx) or (xsymbol), 1 = -symbolx or -xsymbol, 2 = symbol-x or x-symbol, or 3 = symbolx- or xsymbol-, where symbol is the currency symbol of the country or region.Note that the position of the currency symbol is determined by xlCurrencyBefore.
    xlCurrencySpaceBefore                     =	36		# True if a space is added before the currency symbol.
    xlCurrencyTrailingZeros                   =	39		# True if trailing zeros are displayed for zero currency values.
    xlDateOrder                               =	32		# Order of date elements: 0 = month-day-year, 1 = day-month-year, 2 = year-month-day
    xlDateSeparator                           =	17		# Date separator (/).
    xlDayCode                                 =	21		# Day symbol (d).
    xlDayLeadingZero                          =	42		# True if a leading zero is displayed in days.
    xlDecimalSeparator                        =	3		# Decimal separator.
    xlGeneralFormatName                       =	26		# Name of the General number format.
    xlHourCode                                =	22		# Hour symbol (h).
    xlLeftBrace                               =	12		# Character used instead of the left brace ({) in array literals.
    xlLeftBracket                             =	10		# Character used instead of the left bracket ([) in R1C1-style relative references.
    xlListSeparator                           =	5		# List separator.
    xlLowerCaseColumnLetter                   =	9		# Lowercase column letter.
    xlLowerCaseRowLetter                      =	8		# Lowercase row letter.
    xlMDY                                     =	44		# True if the date order is month-day-year for dates displayed in the long form; False if the date order is day-month-year.
    xlMetric                                  =	35		# True if you are using the metric system; False if you are using the English measurement system.
    xlMinuteCode                              =	23		# Minute symbol (m).
    xlMonthCode                               =	20		# Month symbol (m).
    xlMonthLeadingZero                        =	41		# True if a leading zero is displayed in months (when months are displayed as numbers).
    xlMonthNameChars                          =	30		# Always returns three characters for backward compatibility. Abbreviated month names are read from Microsoft Windows and can be any length.
    xlNoncurrencyDigits                       =	29		# Number of decimal digits to be used in noncurrency formats.
    xlNonEnglishFunctions                     =	34		# True if you are not displaying functions in English.
    xlRightBrace                              =	13		# Character used instead of the right brace (}) in array literals.
    xlRightBracket                            =	11		# Character used instead of the right bracket (]) in R1C1-style references.
    xlRowSeparator                            =	15		# Character used to separate rows in array literals.
    xlSecondCode                              =	24		# Second symbol (s).
    xlThousandsSeparator                      =	4		# Zero or thousands separator.
    xlTimeLeadingZero                         =	45		# True if a leading zero is displayed in times.
    xlTimeSeparator                           =	18		# Time separator (:).
    xlUpperCaseColumnLetter                   =	7		# Uppercase column letter.
    xlUpperCaseRowLetter                      =	6		# Uppercase row letter (for R1C1-style references).
    xlWeekdayNameChars                        =	31		# Always returns three characters for backward compatibility. Abbreviated weekday names are read from Microsoft Windows and can be any length.
    xlYearCode                                =	19		# Year symbol in number formats (y).

    #values from excel.xlapplynamesorder
    #****************************
    xlColumnThenRow                           =	2		# Columns listed before rows
    xlRowThenColumn                           =	1		# Rows listed before columns

    #values from excel.xlarabicmodes
    #****************************
    xlArabicBothStrict                        =	3		# The spelling checker uses spelling rules regarding both Arabic words ending with the letter yaa and Arabic words beginning with an alef hamza.
    xlArabicNone                              =	0		# The spelling checker ignores spelling rules regarding either Arabic words ending with the letter yaa or Arabic words beginning with an alef hamza.
    xlArabicStrictAlefHamza                   =	1		# The spelling checker uses spelling rules regarding Arabic words beginning with an alef hamza.
    xlArabicStrictFinalYaa                    =	2		# The spelling checker uses spelling rules regarding Arabic words ending with the letter yaa.

    #values from excel.xlarrangestyle
    #****************************
    xlArrangeStyleCascade                     =	7		# Windows are cascaded.
    xlArrangeStyleHorizontal                  =	-4128		# Windows are arranged horizontally.
    xlArrangeStyleTiled                       =	1		# Default. Windows are tiled.
    xlArrangeStyleVertical                    =	-4166		# Windows are arranged vertically.

    #values from excel.xlarrowheadlength
    #****************************
    xlArrowHeadLengthLong                     =	3		# Longest arrowhead
    xlArrowHeadLengthMedium                   =	-4138		# Medium-length arrowhead
    xlArrowHeadLengthShort                    =	1		# Shortest arrowhead

    #values from excel.xlarrowheadstyle
    #****************************
    xlArrowHeadStyleClosed                    =	3		# Small arrowhead with curved edge at connection to line.
    xlArrowHeadStyleDoubleClosed              =	5		# Large diamond-shaped arrowhead.
    xlArrowHeadStyleDoubleOpen                =	4		# Large arrowhead with curved edge at connection to line.
    xlArrowHeadStyleNone                      =	-4142		# No arrowhead.
    xlArrowHeadStyleOpen                      =	2		# Large triangular arrowhead.

    #values from excel.xlarrowheadwidth
    #****************************
    xlArrowHeadWidthMedium                    =	-4138		# Medium-width arrowhead
    xlArrowHeadWidthNarrow                    =	1		# Narrowest arrowhead
    xlArrowHeadWidthWide                      =	3		# Widest arrowhead

    #values from excel.xlautofilltype
    #****************************
    xlFillCopy                                =	1		# Copy the values and formats from the source range to the target range, repeating if necessary.
    xlFillDays                                =	5		# Extend the names of the days of the week in the source range into the target range. Formats are copied from the source range to the target range, repeating if necessary.
    xlFillDefault                             =	0		# Excel determines the values and formats used to fill the target range.
    xlFillFormats                             =	3		# Copy only the formats from the source range to the target range, repeating if necessary.
    xlFillMonths                              =	7		# Extend the names of the months in the source range into the target range. Formats are copied from the source range to the target range, repeating if necessary.
    xlFillSeries                              =	2		# Extend the values in the source range into the target range as a series (for example, '1, 2' is extended as '3, 4, 5'). Formats are copied from the source range to the target range, repeating if necessary.
    xlFillValues                              =	4		# Copy only the values from the source range to the target range, repeating if necessary.
    xlFillWeekdays                            =	6		# Extend the names of the days of the workweek in the source range into the target range. Formats are copied from the source range to the target range, repeating if necessary.
    xlFillYears                               =	8		# Extend the years in the source range into the target range. Formats are copied from the source range to the target range, repeating if necessary.
    xlGrowthTrend                             =	10		# Extend the numeric values from the source range into the target range, assuming that the relationships between the numbers in the source range are multiplicative (for example, '1, 2,' is extended as '4, 8, 16', assuming that each number is a result of multiplying the previous number by some value). Formats are copied from the source range to the target range, repeating if necessary.
    xlLinearTrend                             =	9		# Extend the numeric values from the source range into the target range, assuming that the relationships between the numbers is additive (for example, '1, 2,' is extended as '3, 4, 5', assuming that each number is a result of adding some value to the previous number). Formats are copied from the source range to the target range, repeating if necessary.
    xlFlashFill                               =	11		# Extend the values from the source range into the target range based on the detected pattern of previous user actions, repeating if necessary.

    #values from excel.xlautofilteroperator
    #****************************
    xlAnd                                     =	1		# Logical AND of Criteria1 and Criteria2
    xlBottom10Items                           =	4		# Lowest-valued items displayed (number of items specified in Criteria1)
    xlBottom10Percent                         =	6		# Lowest-valued items displayed (percentage specified in Criteria1)
    xlFilterCellColor                         =	8		# Color of the cell
    xlFilterDynamic                           =	11		# Dynamic filter
    xlFilterFontColor                         =	9		# Color of the font
    xlFilterIcon                              =	10		# Filter icon
    xlFilterValues                            =	7		# Filter values
    xlOr                                      =	2		# Logical OR of Criteria1 or Criteria2
    xlTop10Items                              =	3		# Highest-valued items displayed (number of items specified in Criteria1)
    xlTop10Percent                            =	5		# Highest-valued items displayed (percentage specified in Criteria1)

    #values from excel.xlaxiscrosses
    #****************************
    xlAxisCrossesAutomatic                    =	-4105		# Microsoft Excel sets the axis crossing point.
    xlAxisCrossesCustom                       =	-4114		# The CrossesAt property specifies the axis crossing point.
    xlAxisCrossesMaximum                      =	2		# The axis crosses at the maximum value.
    xlAxisCrossesMinimum                      =	4		# The axis crosses at the minimum value.

    #values from excel.xlaxisgroup
    #****************************
    xlPrimary                                 =	1		# Primary axis group
    xlSecondary                               =	2		# Secondary axis group

    #values from excel.xlaxistype
    #****************************
    xlCategory                                =	1		# Axis displays categories.
    xlSeriesAxis                              =	3		# Axis displays data series.
    xlValue                                   =	2		# Axis displays values.

    #values from excel.xlbackground
    #****************************
    xlBackgroundAutomatic                     =	-4105		# Excel controls the background.
    xlBackgroundOpaque                        =	3		# Opaque background.
    xlBackgroundTransparent                   =	2		# Transparent background.

    #values from excel.xlbarshape
    #****************************
    xlBox                                     =	0		# Box.
    xlConeToMax                               =	5		# Cone, truncated at value.
    xlConeToPoint                             =	4		# Cone, coming to point at value.
    xlCylinder                                =	3		# Cylinder.
    xlPyramidToMax                            =	2		# Pyramid, truncated at value.
    xlPyramidToPoint                          =	1		# Pyramid, coming to point at value.

    #values from excel.xlbinstype
    #****************************
    xlBinsTypeAutomatic                       =	0		# Sets bins type automatically.
    xlBinsTypeCategorical                     =	1		# Sets bins type by category.
    xlBinsTypeManual                          =	2		# Sets bins type manually.
    xlBinsTypeBinSize                         =	3		# Sets bins type by size.
    xlBinsTypeBinCount                        =	4		# Sets bins type by count.

    #values from excel.xlborderweight
    #****************************
    xlHairline                                =	1		# Hairline (thinnest border).
    xlMedium                                  =	-4138		# Medium.
    xlThick                                   =	4		# Thick (widest border).
    xlThin                                    =	2		# Thin.

    #values from excel.xlbordersindex
    #****************************
    xlDiagonalDown                            =	5		# Border running from the upper-left corner to the lower-right of each cell in the range.
    xlDiagonalUp                              =	6		# Border running from the lower-left corner to the upper-right of each cell in the range.
    xlEdgeBottom                              =	9		# Border at the bottom of the range.
    xlEdgeLeft                                =	7		# Border at the left edge of the range.
    xlEdgeRight                               =	10		# Border at the right edge of the range.
    xlEdgeTop                                 =	8		# Border at the top of the range.
    xlInsideHorizontal                        =	12		# Horizontal borders for all cells in the range except borders on the outside of the range.
    xlInsideVertical                          =	11		# Vertical borders for all the cells in the range except borders on the outside of the range.

    #values from excel.xlbuiltindialog
    #****************************
    xlDialogActivate                          =	103		# Activate dialog box
    xlDialogActiveCellFont                    =	476		# Active Cell Font dialog box
    xlDialogAddChartAutoformat                =	390		# Add Chart Autoformat dialog box
    xlDialogAddinManager                      =	321		# Addin Manager dialog box
    xlDialogAlignment                         =	43		# Alignment dialog box
    xlDialogApplyNames                        =	133		# Apply Names dialog box
    xlDialogApplyStyle                        =	212		# Apply Style dialog box
    xlDialogAppMove                           =	170		# AppMove dialog box
    xlDialogAppSize                           =	171		# AppSize dialog box
    xlDialogArrangeAll                        =	12		# Arrange All dialog box
    xlDialogAssignToObject                    =	213		# Assign To Object dialog box
    xlDialogAssignToTool                      =	293		# Assign To Tool dialog box
    xlDialogAttachText                        =	80		# Attach Text dialog box
    xlDialogAttachToolbars                    =	323		# Attach Toolbars dialog box
    xlDialogAutoCorrect                       =	485		# Auto Correct dialog box
    xlDialogAxes                              =	78		# Axes dialog box
    xlDialogBorder                            =	45		# Border dialog box
    xlDialogCalculation                       =	32		# Calculation dialog box
    xlDialogCellProtection                    =	46		# Cell Protection dialog box
    xlDialogChangeLink                        =	166		# Change Link dialog box
    xlDialogChartAddData                      =	392		# Chart Add Data dialog box
    xlDialogChartLocation                     =	527		# Chart Location dialog box
    xlDialogChartOptionsDataLabelMultiple     =	724		# Chart Options DataLabel Multiple dialog box
    xlDialogChartOptionsDataLabels            =	505		# Chart Options DataLabels dialog box
    xlDialogChartOptionsDataTable             =	506		# Chart Options DataTable dialog box
    xlDialogChartSourceData                   =	540		# Chart SourceData dialog box
    xlDialogChartTrend                        =	350		# Chart Trend dialog box
    xlDialogChartType                         =	526		# Chart Type dialog box
    xlDialogChartWizard                       =	288		# ChartWizard dialog box
    xlDialogCheckboxProperties                =	435		# Checkbox Properties dialog box
    xlDialogClear                             =	52		# Clear dialog box
    xlDialogColorPalette                      =	161		# Color Palette dialog box
    xlDialogColumnWidth                       =	47		# Column Width dialog box
    xlDialogCombination                       =	73		# Combination dialog box
    xlDialogConditionalFormatting             =	583		# Conditional Formatting dialog box
    xlDialogConsolidate                       =	191		# Consolidate dialog box
    xlDialogCopyChart                         =	147		# Copy Chart dialog box
    xlDialogCopyPicture                       =	108		# Copy Picture dialog box
    xlDialogCreateList                        =	796		# Create List dialog box
    xlDialogCreateNames                       =	62		# Create Names dialog box
    xlDialogCreatePublisher                   =	217		# Create Publisher dialog box
    xlDialogCreateRelationship                =	1272		# Create Relationship dialog box
    xlDialogCustomizeToolbar                  =	276		# Customize Toolbar dialog box
    xlDialogCustomViews                       =	493		# Custom Views dialog box
    xlDialogDataDelete                        =	36		# Data Delete dialog box
    xlDialogDataLabel                         =	379		# Data Label dialog box
    xlDialogDataLabelMultiple                 =	723		# Data Label Multiple dialog box
    xlDialogDataSeries                        =	40		# Data Series dialog box
    xlDialogDataValidation                    =	525		# Data Validation dialog box
    xlDialogDefineName                        =	61		# Define Name dialog box
    xlDialogDefineStyle                       =	229		# Define Style dialog box
    xlDialogDeleteFormat                      =	111		# Delete Format dialog box
    xlDialogDeleteName                        =	110		# Delete Name dialog box
    xlDialogDemote                            =	203		# Demote dialog box
    xlDialogDisplay                           =	27		# Display dialog box
    xlDialogDocumentInspector                 =	862		# Document Inspector dialog box
    xlDialogEditboxProperties                 =	438		# Editbox Properties dialog box
    xlDialogEditColor                         =	223		# Edit Color dialog box
    xlDialogEditDelete                        =	54		# Edit Delete dialog box
    xlDialogEditionOptions                    =	251		# Edition Options dialog box
    xlDialogEditSeries                        =	228		# Edit Series dialog box
    xlDialogErrorbarX                         =	463		# Errorbar X dialog box
    xlDialogErrorbarY                         =	464		# Errorbar Y dialog box
    xlDialogErrorChecking                     =	732		# Error Checking dialog box
    xlDialogEvaluateFormula                   =	709		# Evaluate Formula dialog box
    xlDialogExternalDataProperties            =	530		# External Data Properties dialog box
    xlDialogExtract                           =	35		# Extract dialog box
    xlDialogFileDelete                        =	6		# File Delete dialog box
    xlDialogFileSharing                       =	481		# File Sharing dialog box
    xlDialogFillGroup                         =	200		# Fill Group dialog box
    xlDialogFillWorkgroup                     =	301		# Fill Workgroup dialog box
    xlDialogFilter                            =	447		# Dialog Filter dialog box
    xlDialogFilterAdvanced                    =	370		# Filter Advanced dialog box
    xlDialogFindFile                          =	475		# Find File dialog box
    xlDialogFont                              =	26		# Font dialog box
    xlDialogFontProperties                    =	381		# Font Properties dialog box
    xlDialogFormatAuto                        =	269		# Format Auto dialog box
    xlDialogFormatChart                       =	465		# Format Chart dialog box
    xlDialogFormatCharttype                   =	423		# Format Charttype dialog box
    xlDialogFormatFont                        =	150		# Format Font dialog box
    xlDialogFormatLegend                      =	88		# Format Legend dialog box
    xlDialogFormatMain                        =	225		# Format Main dialog box
    xlDialogFormatMove                        =	128		# Format Move dialog box
    xlDialogFormatNumber                      =	42		# Format Number dialog box
    xlDialogFormatOverlay                     =	226		# Format Overlay dialog box
    xlDialogFormatSize                        =	129		# Format Size dialog box
    xlDialogFormatText                        =	89		# Format Text dialog box
    xlDialogFormulaFind                       =	64		# Formula Find dialog box
    xlDialogFormulaGoto                       =	63		# Formula Goto dialog box
    xlDialogFormulaReplace                    =	130		# Formula Replace dialog box
    xlDialogFunctionWizard                    =	450		# Function Wizard dialog box
    xlDialogGallery3dArea                     =	193		# Gallery 3D Area dialog box
    xlDialogGallery3dBar                      =	272		# Gallery 3D Bar dialog box
    xlDialogGallery3dColumn                   =	194		# Gallery 3D Column dialog box
    xlDialogGallery3dLine                     =	195		# Gallery 3D Line dialog box
    xlDialogGallery3dPie                      =	196		# Gallery 3D Pie dialog box
    xlDialogGallery3dSurface                  =	273		# Gallery 3D Surface dialog box
    xlDialogGalleryArea                       =	67		# Gallery Area dialog box
    xlDialogGalleryBar                        =	68		# Gallery Bar dialog box
    xlDialogGalleryColumn                     =	69		# Gallery Column dialog box
    xlDialogGalleryCustom                     =	388		# Gallery Custom dialog box
    xlDialogGalleryDoughnut                   =	344		# Gallery Doughnut dialog box
    xlDialogGalleryLine                       =	70		# Gallery Line dialog box
    xlDialogGalleryPie                        =	71		# Gallery Pie dialog box
    xlDialogGalleryRadar                      =	249		# Gallery Radar dialog box
    xlDialogGalleryScatter                    =	72		# Gallery Scatter dialog box
    xlDialogGoalSeek                          =	198		# Goal Seek dialog box
    xlDialogGridlines                         =	76		# Gridlines dialog box
    xlDialogImportTextFile                    =	666		# Import Text File dialog box
    xlDialogInsert                            =	55		# Insert dialog box
    xlDialogInsertHyperlink                   =	596		# Insert Hyperlink dialog box
    xlDialogInsertObject                      =	259		# Insert Object dialog box
    xlDialogInsertPicture                     =	342		# Insert Picture dialog box
    xlDialogInsertTitle                       =	380		# Insert Title dialog box
    xlDialogLabelProperties                   =	436		# Label Properties dialog box
    xlDialogListboxProperties                 =	437		# Listbox Properties dialog box
    xlDialogMacroOptions                      =	382		# Macro Options dialog box
    xlDialogMailEditMailer                    =	470		# Mail Edit Mailer dialog box
    xlDialogMailLogon                         =	339		# Mail Logon dialog box
    xlDialogMailNextLetter                    =	378		# Mail Next Letter dialog box
    xlDialogMainChart                         =	85		# Main Chart dialog box
    xlDialogMainChartType                     =	185		# Main Chart Type dialog box
    xlDialogManageRelationships               =	1271		# Manage Relationships dialog box
    xlDialogMenuEditor                        =	322		# Menu Editor dialog box
    xlDialogMove                              =	262		# Move dialog box
    xlDialogMyPermission                      =	834		# My Permission dialog box
    xlDialogNameManager                       =	977		# NameManager dialog box
    xlDialogNew                               =	119		# New dialog box
    xlDialogNewName                           =	978		# NewName dialog box
    xlDialogNewWebQuery                       =	667		# New Web Query dialog box
    xlDialogNote                              =	154		# Note dialog box
    xlDialogObjectProperties                  =	207		# Object Properties dialog box
    xlDialogObjectProtection                  =	214		# Object Protection dialog box
    xlDialogOpen                              =	1		# Open dialog box
    xlDialogOpenLinks                         =	2		# Open Links dialog box
    xlDialogOpenMail                          =	188		# Open Mail dialog box
    xlDialogOpenText                          =	441		# Open Text dialog box
    xlDialogOptionsCalculation                =	318		# Options Calculation dialog box
    xlDialogOptionsChart                      =	325		# Options Chart dialog box
    xlDialogOptionsEdit                       =	319		# Options Edit dialog box
    xlDialogOptionsGeneral                    =	356		# Options General dialog box
    xlDialogOptionsListsAdd                   =	458		# Options Lists Add dialog box
    xlDialogOptionsME                         =	647		# OptionsME dialog box
    xlDialogOptionsTransition                 =	355		# Options Transition dialog box
    xlDialogOptionsView                       =	320		# Options View dialog box
    xlDialogOutline                           =	142		# Outline dialog box
    xlDialogOverlay                           =	86		# Overlay dialog box
    xlDialogOverlayChartType                  =	186		# Overlay ChartType dialog box
    xlDialogPageSetup                         =	7		# Page Setup dialog box
    xlDialogParse                             =	91		# Parse dialog box
    xlDialogPasteNames                        =	58		# Paste Names dialog box
    xlDialogPasteSpecial                      =	53		# Paste Special dialog box
    xlDialogPatterns                          =	84		# Patterns dialog box
    xlDialogPermission                        =	832		# Permission dialog box
    xlDialogPhonetic                          =	656		# Phonetic dialog box
    xlDialogPivotCalculatedField              =	570		# Pivot Calculated Field dialog box
    xlDialogPivotCalculatedItem               =	572		# Pivot Calculated Item dialog box
    xlDialogPivotClientServerSet              =	689		# Pivot Client Server Set dialog box
    xlDialogPivotFieldGroup                   =	433		# Pivot Field Group dialog box
    xlDialogPivotFieldProperties              =	313		# Pivot Field Properties dialog box
    xlDialogPivotFieldUngroup                 =	434		# Pivot Field Ungroup dialog box
    xlDialogPivotShowPages                    =	421		# Pivot Show Pages dialog box
    xlDialogPivotSolveOrder                   =	568		# Pivot Solve Order dialog box
    xlDialogPivotTableOptions                 =	567		# Pivot Table Options dialog box
    xlDialogPivotTableSlicerConnections       =	1183		# Pivot Table Slicer Connections dialog box
    xlDialogPivotTableWhatIfAnalysisSettings  =	1153		# Pivot Table What If Analysis Settings dialog box
    xlDialogPivotTableWizard                  =	312		# Pivot Table Wizard dialog box
    xlDialogPlacement                         =	300		# Placement dialog box
    xlDialogPrint                             =	8		# Print dialog box
    xlDialogPrinterSetup                      =	9		# Printer Setup dialog box
    xlDialogPrintPreview                      =	222		# Print Preview dialog box
    xlDialogPromote                           =	202		# Promote dialog box
    xlDialogProperties                        =	474		# Properties dialog box
    xlDialogPropertyFields                    =	754		# Property Fields dialog box
    xlDialogProtectDocument                   =	28		# Protect Document dialog box
    xlDialogProtectSharing                    =	620		# Protect Sharing dialog box
    xlDialogPublishAsWebPage                  =	653		# Publish As WebPage dialog box
    xlDialogPushbuttonProperties              =	445		# Pushbutton Properties dialog box
    xlDialogRecommendedPivotTables            =	1258		# Recommended PivotTables dialog box
    xlDialogReplaceFont                       =	134		# Replace Font dialog box
    xlDialogRoutingSlip                       =	336		# This object or member has been deprecated, but it remains part of the object model for backward compatibility. You should not use it in new applications.
    xlDialogRowHeight                         =	127		# Row Height dialog box
    xlDialogRun                               =	17		# Run dialog box
    xlDialogSaveAs                            =	5		# SaveAs dialog box
    xlDialogSaveCopyAs                        =	456		# SaveCopyAs dialog box
    xlDialogSaveNewObject                     =	208		# Save New Object dialog box
    xlDialogSaveWorkbook                      =	145		# Save Workbook dialog box
    xlDialogSaveWorkspace                     =	285		# Save Workspace dialog box
    xlDialogScale                             =	87		# Scale dialog box
    xlDialogScenarioAdd                       =	307		# Scenario Add dialog box
    xlDialogScenarioCells                     =	305		# Scenario Cells dialog box
    xlDialogScenarioEdit                      =	308		# Scenario Edit dialog box
    xlDialogScenarioMerge                     =	473		# Scenario Merge dialog box
    xlDialogScenarioSummary                   =	311		# Scenario Summary dialog box
    xlDialogScrollbarProperties               =	420		# Scrollbar Properties dialog box
    xlDialogSearch                            =	731		# Search dialog box
    xlDialogSelectSpecial                     =	132		# Select Special dialog box
    xlDialogSendMail                          =	189		# Send Mail dialog box
    xlDialogSeriesAxes                        =	460		# Series Axes dialog box
    xlDialogSeriesOptions                     =	557		# Series Options dialog box
    xlDialogSeriesOrder                       =	466		# Series Order dialog box
    xlDialogSeriesShape                       =	504		# Series Shape dialog box
    xlDialogSeriesX                           =	461		# Series X dialog box
    xlDialogSeriesY                           =	462		# Series Y dialog box
    xlDialogSetBackgroundPicture              =	509		# Set Background Picture dialog box
    xlDialogSetManager                        =	1109		# Set Manager dialog box
    xlDialogSetMDXEditor                      =	1208		# Set MDX Editor dialog box
    xlDialogSetPrintTitles                    =	23		# Set Print Titles dialog box
    xlDialogSetTupleEditorOnColumns           =	1108		# Set Tuple Editor On Columns dialog box
    xlDialogSetTupleEditorOnRows              =	1107		# Set Tuple Editor On Rows dialog box
    xlDialogSetUpdateStatus                   =	159		# Set Update Status dialog box
    xlDialogShowDetail                        =	204		# Show Detail dialog box
    xlDialogShowToolbar                       =	220		# Show Toolbar dialog box
    xlDialogSize                              =	261		# Size dialog box
    xlDialogSlicerCreation                    =	1182		# Slicer Creation dialog box
    xlDialogSlicerPivotTableConnections       =	1184		# Slicer Pivot Table Connections dialog box
    xlDialogSlicerSettings                    =	1179		# Slicer Settings dialog box
    xlDialogSort                              =	39		# Sort dialog box
    xlDialogSortSpecial                       =	192		# Sort Special dialog box
    xlDialogSparklineInsertColumn             =	1134		# Sparkline Insert Column dialog box
    xlDialogSparklineInsertLine               =	1133		# Sparkline Insert Line dialog box
    xlDialogSparklineInsertWinLoss            =	1135		# Sparkline Insert Win Loss dialog box
    xlDialogSplit                             =	137		# Split dialog box
    xlDialogStandardFont                      =	190		# Standard Font dialog box
    xlDialogStandardWidth                     =	472		# Standard Width dialog box
    xlDialogStyle                             =	44		# Style dialog box
    xlDialogSubscribeTo                       =	218		# Subscribe To dialog box
    xlDialogSubtotalCreate                    =	398		# Subtotal Create dialog box
    xlDialogSummaryInfo                       =	474		# Summary Info dialog box
    xlDialogTable                             =	41		# Table dialog box
    xlDialogTabOrder                          =	394		# Tab Order dialog box
    xlDialogTextToColumns                     =	422		# Text To Columns dialog box
    xlDialogUnhide                            =	94		# Unhide dialog box
    xlDialogUpdateLink                        =	201		# Update Link dialog box
    xlDialogVbaInsertFile                     =	328		# VBA Insert File dialog box
    xlDialogVbaMakeAddin                      =	478		# VBA Make Addin dialog box
    xlDialogVbaProcedureDefinition            =	330		# VBA Procedure Definition dialog box
    xlDialogView3d                            =	197		# View 3D dialog box
    xlDialogWebOptionsBrowsers                =	773		# Web Options Browsers dialog box
    xlDialogWebOptionsEncoding                =	686		# Web Options Encoding dialog box
    xlDialogWebOptionsFiles                   =	684		# Web Options Files dialog box
    xlDialogWebOptionsFonts                   =	687		# Web Options Fonts dialog box
    xlDialogWebOptionsGeneral                 =	683		# Web Options General dialog box
    xlDialogWebOptionsPictures                =	685		# Web Options Pictures dialog box
    xlDialogWindowMove                        =	14		# Window Move dialog box
    xlDialogWindowSize                        =	13		# Window Size dialog box
    xlDialogWorkbookAdd                       =	281		# Workbook Add dialog box
    xlDialogWorkbookCopy                      =	283		# Workbook Copy dialog box
    xlDialogWorkbookInsert                    =	354		# Workbook Insert dialog box
    xlDialogWorkbookMove                      =	282		# Workbook Move dialog box
    xlDialogWorkbookName                      =	386		# Workbook Name dialog box
    xlDialogWorkbookNew                       =	302		# Workbook New dialog box
    xlDialogWorkbookOptions                   =	284		# Workbook Options dialog box
    xlDialogWorkbookProtect                   =	417		# Workbook Protect dialog box
    xlDialogWorkbookTabSplit                  =	415		# Workbook Tab Split dialog box
    xlDialogWorkbookUnhide                    =	384		# Workbook Unhide dialog box
    xlDialogWorkgroup                         =	199		# Workgroup dialog box
    xlDialogWorkspace                         =	95		# Workspace dialog box
    xlDialogZoom                              =	256		# Zoom dialog box

    #values from excel.xlcverror
    #****************************
    xlErrDiv0                                 =	2007		# Error number: 2007
    xlErrNA                                   =	2042		# Error number: 2042
    xlErrName                                 =	2029		# Error number: 2029
    xlErrNull                                 =	2000		# Error number: 2000
    xlErrNum                                  =	2036		# Error number: 2036
    xlErrRef                                  =	2023		# Error number: 2023
    xlErrValue                                =	2015		# Error number: 2015

    #values from excel.xlcalcfor
    #****************************
    xlAllValues                               =	0		# All values.
    xlColGroups                               =	2		# Column groups.
    xlRowGroups                               =	1		# Row groups.

    #values from excel.xlcalcmemnumberformattype
    #****************************
    xlNumberFormatTypeDefault                 =	0		# Use the default format type of the calculated member for the cell value.
    xlNumberFormatTypeNumber                  =	1		# Calculated member cell format is a number.
    xlNumberFormatTypePercent                 =	2		# Calculated member cell format is a percentage.

    #values from excel.xlcalculatedmembertype
    #****************************
    xlCalculatedMeasure                       =	2		# The member is a Multidimensional Expressions (MDX) expression that defines the measure.
    xlCalculatedMember                        =	0		# The member uses a Multidimensional Expression (MDX) formula.
    xlCalculatedSet                           =	1		# The member contains an MDX formula for a set in a cube field.

    #values from excel.xlcalculation
    #****************************
    xlCalculationAutomatic                    =	-4105		# Excel controls recalculation.
    xlCalculationManual                       =	-4135		# Calculation is done when the user requests it.
    xlCalculationSemiautomatic                =	2		# Excel controls recalculation but ignores changes in tables.

    #values from excel.xlcalculationinterruptkey
    #****************************
    xlAnyKey                                  =	2		# Pressing any key interrupts recalculation.
    xlEscKey                                  =	1		# Pressing the ESC key interrupts recalculation.
    xlNoKey                                   =	0		# No key press can interrupt recalculation.

    #values from excel.xlcalculationstate
    #****************************
    xlCalculating                             =	1		# Calculations in process.
    xlDone                                    =	0		# Calculations complete.
    xlPending                                 =	2		# Changes that trigger calculation have been made, but a recalculation has not yet been performed.

    #values from excel.xlcategorylabellevel
    #****************************
    xlCategoryLabelLevelAll                   =	-1		# Set category labels to all category label levels w/in range on the chart.
    xlCategoryLabelLevelCustom                =	-2		# Indicates literal data in the category labels.
    xlCategoryLabelLevelNone                  =	-3		# Set no category labels in the chart. Defaults to automatic indexed labels.

    #values from excel.xlcategorytype
    #****************************
    xlAutomaticScale                          =	-4105		# Excel controls the axis type.
    xlCategoryScale                           =	2		# Axis groups data by an arbitrary set of categories.
    xlTimeScale                               =	3		# Axis groups data on a time scale.

    #values from excel.xlcellchangedstate
    #****************************
    xlCellChangeApplied                       =	3		# The value in the cell has been edited or recalculated, and that change has been applied to the data source. (Applies only PivotTable reports with OLAP data sources)
    xlCellChanged                             =	2		# The value in the cell has been edited or recalculated.
    xlCellNotChanged                          =	1		# The value in the cell has not been edited or recalculated.

    #values from excel.xlcellinsertionmode
    #****************************
    xlInsertDeleteCells                       =	1		# Partial rows are inserted or deleted to match the exact number of rows required for the new recordset.
    xlInsertEntireRows                        =	2		# Entire rows are inserted, if necessary, to accommodate any overflow. No cells or rows are deleted from the worksheet.
    xlOverwriteCells                          =	0		# No new cells or rows are added to the worksheet. Data in surrounding cells is overwritten to accommodate any overflow.

    #values from excel.xlcelltype
    #****************************
    xlCellTypeAllFormatConditions             =	-4172		# Cells of any format.
    xlCellTypeAllValidation                   =	-4174		# Cells having validation criteria.
    xlCellTypeBlanks                          =	4		# Empty cells.
    xlCellTypeComments                        =	-4144		# Cells containing notes.
    xlCellTypeConstants                       =	2		# Cells containing constants.
    xlCellTypeFormulas                        =	-4123		# Cells containing formulas.
    xlCellTypeLastCell                        =	11		# The last cell in the used range.
    xlCellTypeSameFormatConditions            =	-4173		# Cells having the same format.
    xlCellTypeSameValidation                  =	-4175		# Cells having the same validation criteria.
    xlCellTypeVisible                         =	12		# All visible cells.

    #values from excel.xlchartelementposition
    #****************************
    xlChartElementPositionAutomatic           =	-4105		# Automatically sets the position of the chart element.
    xlChartElementPositionCustom              =	-4114		# Specifies a specific position for the chart element.

    #values from excel.xlchartgallery
    #****************************
    xlAnyGallery                              =	23		# Either of the galleries.
    xlBuiltIn                                 =	21		# The built-in gallery.
    xlUserDefined                             =	22		# The user-defined gallery.

    #values from excel.xlchartitem
    #****************************
    xlAxis                                    =	21		# Axis.
    xlAxisTitle                               =	17		# Axis title.
    xlChartArea                               =	2		# Chart area.
    xlChartTitle                              =	4		# Chart title.
    xlCorners                                 =	6		# Corners.
    xlDataLabel                               =	0		# Data label.
    xlDataTable                               =	7		# Data table.
    xlDisplayUnitLabel                        =	30		# Display unit label.
    xlDownBars                                =	20		# Down bars.
    xlDropLines                               =	26		# Drop lines.
    xlErrorBars                               =	9		# Error bars.
    xlFloor                                   =	23		# Floor.
    xlHiLoLines                               =	25		# HiLo lines.
    xlLeaderLines                             =	29		# Leader lines.
    xlLegend                                  =	24		# Legend.
    xlLegendEntry                             =	12		# Legend entry.
    xlLegendKey                               =	13		# Legend key.
    xlMajorGridlines                          =	15		# Major gridlines.
    xlMinorGridlines                          =	16		# Minor gridlines.
    xlNothing                                 =	28		# Nothing.
    xlPivotChartDropZone                      =	32		# PivotChart drop zone.
    xlPivotChartFieldButton                   =	31		# PivotChart field button.
    xlPlotArea                                =	19		# Plot area.
    xlRadarAxisLabels                         =	27		# Radar axis labels.
    xlSeries                                  =	3		# Series.
    xlSeriesLines                             =	22		# Series lines.
    xlShape                                   =	14		# Shape.
    xlTrendline                               =	8		# Trend line.
    xlUpBars                                  =	18		# Up bars.
    xlWalls                                   =	5		# Walls.
    xlXErrorBars                              =	10		# X error bars.
    xlYErrorBars                              =	11		# Y error bars.

    #values from excel.xlchartlocation
    #****************************
    xlLocationAsNewSheet                      =	1		# Chart is moved to a new sheet.
    xlLocationAsObject                        =	2		# Chart is to be embedded in an existing sheet.
    xlLocationAutomatic                       =	3		# Excel controls chart location.

    #values from excel.xlchartpictureplacement
    #****************************
    xlAllFaces                                =	7		# Display on all faces.
    xlEnd                                     =	2		# Display on end.
    xlEndSides                                =	3		# Display on end and sides.
    xlFront                                   =	4		# Display on front.
    xlFrontEnd                                =	6		# Display on front and end.
    xlFrontSides                              =	5		# Display on front and sides.
    xlSides                                   =	1		# Display on sides.

    #values from excel.xlchartpicturetype
    #****************************
    xlStack                                   =	2		# Picture is sized to repeat a maximum of 15 times in the longest stacked bar.
    xlStackScale                              =	3		# Picture is sized to a specified number of units and repeated the length of the bar.
    xlStretch                                 =	1		# Picture is stretched the full length of the stacked bar.

    #values from excel.xlchartsplittype
    #****************************
    xlSplitByCustomSplit                      =	4		# Arbitrary slides are displayed in the second chart.
    xlSplitByPercentValue                     =	3		# Second chart displays values less than some percentage of the total value. The percentage is specified by the  SplitValue property.
    xlSplitByPosition                         =	1		# Second chart displays the smallest values in the data series. The number of values to display is specified by the  SplitValue property.
    xlSplitByValue                            =	2		# Second chart displays values less than the value specified by the  SplitValue property.

    #values from excel.xlcharttype
    #****************************
    xl3DArea                                  =	-4098		# 3D Area.
    xl3DAreaStacked                           =	78		# 3D Stacked Area.
    xl3DAreaStacked100                        =	79		# 100% Stacked Area.
    xl3DBarClustered                          =	60		# 3D Clustered Bar.
    xl3DBarStacked                            =	61		# 3D Stacked Bar.
    xl3DBarStacked100                         =	62		# 3D 100% Stacked Bar.
    xl3DColumn                                =	-4100		# 3D Column.
    xl3DColumnClustered                       =	54		# 3D Clustered Column.
    xl3DColumnStacked                         =	55		# 3D Stacked Column.
    xl3DColumnStacked100                      =	56		# 3D 100% Stacked Column.
    xl3DLine                                  =	-4101		# 3D Line.
    xl3DPie                                   =	-4102		# 3D Pie.
    xl3DPieExploded                           =	70		# Exploded 3D Pie.
    xlArea                                    =	1		# Area
    xlAreaStacked                             =	76		# Stacked Area.
    xlAreaStacked100                          =	77		# 100% Stacked Area.
    xlBarClustered                            =	57		# Clustered Bar.
    xlBarOfPie                                =	71		# Bar of Pie.
    xlBarStacked                              =	58		# Stacked Bar.
    xlBarStacked100                           =	59		# 100% Stacked Bar.
    xlBubble                                  =	15		# Bubble.
    xlBubble3DEffect                          =	87		# Bubble with 3D effects.
    xlColumnClustered                         =	51		# Clustered Column.
    xlColumnStacked                           =	52		# Stacked Column.
    xlColumnStacked100                        =	53		# 100% Stacked Column.
    xlConeBarClustered                        =	102		# Clustered Cone Bar.
    xlConeBarStacked                          =	103		# Stacked Cone Bar.
    xlConeBarStacked100                       =	104		# 100% Stacked Cone Bar.
    xlConeCol                                 =	105		# 3D Cone Column.
    xlConeColClustered                        =	99		# Clustered Cone Column.
    xlConeColStacked                          =	100		# Stacked Cone Column.
    xlConeColStacked100                       =	101		# 100% Stacked Cone Column.
    xlCylinderBarClustered                    =	95		# Clustered Cylinder Bar.
    xlCylinderBarStacked                      =	96		# Stacked Cylinder Bar.
    xlCylinderBarStacked100                   =	97		# 100% Stacked Cylinder Bar.
    xlCylinderCol                             =	98		# 3D Cylinder Column.
    xlCylinderColClustered                    =	92		# Clustered Cone Column.
    xlCylinderColStacked                      =	93		# Stacked Cone Column.
    xlCylinderColStacked100                   =	94		# 100% Stacked Cylinder Column.
    xlDoughnut                                =	-4120		# Doughnut.
    xlDoughnutExploded                        =	80		# Exploded Doughnut.
    xlLine                                    =	4		# Line.
    xlLineMarkers                             =	65		# Line with Markers.
    xlLineMarkersStacked                      =	66		# Stacked Line with Markers.
    xlLineMarkersStacked100                   =	67		# 100% Stacked Line with Markers.
    xlLineStacked                             =	63		# Stacked Line.
    xlLineStacked100                          =	64		# 100% Stacked Line.
    xlPie                                     =	5		# Pie.
    xlPieExploded                             =	69		# Exploded Pie.
    xlPieOfPie                                =	68		# Pie of Pie.
    xlPyramidBarClustered                     =	109		# Clustered Pyramid Bar.
    xlPyramidBarStacked                       =	110		# Stacked Pyramid Bar.
    xlPyramidBarStacked100                    =	111		# 100% Stacked Pyramid Bar.
    xlPyramidCol                              =	112		# 3D Pyramid Column.
    xlPyramidColClustered                     =	106		# Clustered Pyramid Column.
    xlPyramidColStacked                       =	107		# Stacked Pyramid Column.
    xlPyramidColStacked100                    =	108		# 100% Stacked Pyramid Column.
    xlRadar                                   =	-4151		# Radar.
    xlRadarFilled                             =	82		# Filled Radar.
    xlRadarMarkers                            =	81		# Radar with Data Markers.
    xlStockHLC                                =	88		# High-Low-Close.
    xlStockOHLC                               =	89		# Open-High-Low-Close.
    xlStockVHLC                               =	90		# Volume-High-Low-Close.
    xlStockVOHLC                              =	91		# Volume-Open-High-Low-Close.
    xlSurface                                 =	83		# 3D Surface.
    xlSurfaceTopView                          =	85		# Surface (Top View).
    xlSurfaceTopViewWireframe                 =	86		# Surface (Top View wireframe).
    xlSurfaceWireframe                        =	84		# 3D Surface (wireframe).
    xlXYScatter                               =	-4169		# Scatter.
    xlXYScatterLines                          =	74		# Scatter with Lines.
    xlXYScatterLinesNoMarkers                 =	75		# Scatter with Lines and No Data Markers.
    xlXYScatterSmooth                         =	72		# Scatter with Smoothed Lines.
    xlXYScatterSmoothNoMarkers                =	73		# Scatter with Smoothed Lines and No Data Markers.

    #values from excel.xlcheckinversiontype
    #****************************
    xlCheckInMajorVersion                     =	1		# Check in the major version.
    xlCheckInMinorVersion                     =	0		# Check in the minor version.
    xlCheckInOverwriteVersion                 =	2		# Overwrite current version on the server.

    #values from excel.xlclipboardformat
    #****************************
    xlClipboardFormatBIFF                     =	8		# Binary Interchange file format for Excel version 2.x
    xlClipboardFormatBIFF12                   =	63		# Binary Interchange file format 12
    xlClipboardFormatBIFF2                    =	18		# Binary Interchange file format 2
    xlClipboardFormatBIFF3                    =	20		# Binary Interchange file format 3
    xlClipboardFormatBIFF4                    =	30		# Binary Interchange file format 4
    xlClipboardFormatBinary                   =	15		# Binary format
    xlClipboardFormatBitmap                   =	9		# Bitmap format
    xlClipboardFormatCGM                      =	13		# CGM format
    xlClipboardFormatCSV                      =	5		# CSV format
    xlClipboardFormatDIF                      =	4		# DIF format
    xlClipboardFormatDspText                  =	12		# Dsp Text format
    xlClipboardFormatEmbeddedObject           =	21		# Embedded Object
    xlClipboardFormatEmbedSource              =	22		# Embedded Source
    xlClipboardFormatLink                     =	11		# Link
    xlClipboardFormatLinkSource               =	23		# Link to the source file
    xlClipboardFormatLinkSourceDesc           =	32		# Link to the source description
    xlClipboardFormatMovie                    =	24		# Movie
    xlClipboardFormatNative                   =	14		# Native
    xlClipboardFormatObjectDesc               =	31		# Object description
    xlClipboardFormatObjectLink               =	19		# Object link
    xlClipboardFormatOwnerLink                =	17		# Link to the owner
    xlClipboardFormatPICT                     =	2		# Picture
    xlClipboardFormatPrintPICT                =	3		# Print picture
    xlClipboardFormatRTF                      =	7		# RTF format
    xlClipboardFormatScreenPICT               =	29		# Screen Picture
    xlClipboardFormatStandardFont             =	28		# Standard Font
    xlClipboardFormatStandardScale            =	27		# Standard Scale
    xlClipboardFormatSYLK                     =	6		# SYLK
    xlClipboardFormatTable                    =	16		# Table
    xlClipboardFormatText                     =	0		# Text
    xlClipboardFormatToolFace                 =	25		# Tool Face
    xlClipboardFormatToolFacePICT             =	26		# Tool Face Picture
    xlClipboardFormatVALU                     =	1		# Value
    xlClipboardFormatWK1                      =	10		# Workbook

    #values from excel.xlcmdtype
    #****************************
    xlCmdCube                                 =	1		# Contains a cube name for an OLAP data source.
    xlCmdDAX                                  =	8		# Contains a Data Analysis Expressions (DAX) formula.
    xlCmdDefault                              =	4		# Contains command text that the OLE DB provider understands.
    xlCmdExcel                                =	7		# Contains an Excel formula.
    xlCmdList                                 =	5		# Contains a pointer to list data.
    xlCmdSql                                  =	2		# Contains an SQL statement.
    xlCmdTable                                =	3		# Contains a table name for accessing OLE DB data sources.
    xlCmdTableCollection                      =	6		# Contains the name of a table collection.

    #values from excel.xlcolorindex
    #****************************
    xlColorIndexAutomatic                     =	-4105		# Automatic color.
    xlColorIndexNone                          =	-4142		# No color.

    #values from excel.xlcolumndatatype
    #****************************
    xlDMYFormat                               =	4		# DMY date format.
    xlDYMFormat                               =	7		# DYM date format.
    xlEMDFormat                               =	10		# EMD date format.
    xlGeneralFormat                           =	1		# General.
    xlMDYFormat                               =	3		# MDY date format.
    xlMYDFormat                               =	6		# MYD date format.
    xlSkipColumn                              =	9		# Column is not parsed.
    xlTextFormat                              =	2		# Text.
    xlYDMFormat                               =	8		# YDM date format.
    xlYMDFormat                               =	5		# YMD date format.

    #values from excel.xlcommandunderlines
    #****************************
    xlCommandUnderlinesAutomatic              =	-4105		# Excel controls the display of command underlines.
    xlCommandUnderlinesOff                    =	-4146		# Command underlines are not displayed.
    xlCommandUnderlinesOn                     =	1		# Command underlines are displayed.

    #values from excel.xlcommentdisplaymode
    #****************************
    xlCommentAndIndicator                     =	1		# Display comment and indicator at all times.
    xlCommentIndicatorOnly                    =	-1		# Display comment indicator only. Display comment when mouse pointer is moved over cell.
    xlNoIndicator                             =	0		# Display neither the comment nor the comment indicator at any time.

    #values from excel.xlconditionvaluetypes
    #****************************
    xlConditionValueAutomaticMax              =	7		# The longest data bar is proportional to the maximum value in the range.
    xlConditionValueAutomaticMin              =	6		# The shortest data bar is proportional to the minimum value in the range.
    xlConditionValueFormula                   =	4		# Formula is used.
    xlConditionValueHighestValue              =	2		# Highest value from the list of values.
    xlConditionValueLowestValue               =	1		# Lowest value from the list of values.
    xlConditionValueNone                      =	-1		# No conditional value.
    xlConditionValueNumber                    =	0		# Number is used.
    xlConditionValuePercent                   =	3		# Percentage is used.
    xlConditionValuePercentile                =	5		# Percentile is used.

    #values from excel.xlconnectiontype
    #****************************
    xlConnectionTypeDATAFEED                  =	6		# Data Feed
    xlConnectionTypeMODEL                     =	7		# PowerPivot Model
    xlConnectionTypeNOSOURCE                  =	9		# No source
    xlConnectionTypeODBC                      =	2		# ODBC
    xlConnectionTypeOLEDB                     =	1		# OLEDB
    xlConnectionTypeTEXT                      =	4		# Text
    xlConnectionTypeWEB                       =	5		# Web
    xlConnectionTypeWORKSHEET                 =	8		# Worksheet
    xlConnectionTypeXMLMAP                    =	3		# XML MAP

    #values from excel.xlconsolidationfunction
    #****************************
    xlAverage                                 =	-4106		# Average.
    xlCount                                   =	-4112		# Count.
    xlCountNums                               =	-4113		# Count numerical values only.
    xlDistinctCount                           =	111		# Count using Distinct Count analysis.
    xlMax                                     =	-4136		# Maximum.
    xlMin                                     =	-4139		# Minimum.
    xlProduct                                 =	-4149		# Multiply.
    xlStDev                                   =	-4155		# Standard deviation, based on a sample.
    xlStDevP                                  =	-4156		# Standard deviation, based on the whole population.
    xlSum                                     =	-4157		# Sum.
    xlUnknown                                 =	1000		# No subtotal function specified.
    xlVar                                     =	-4164		# Variation, based on a sample.
    xlVarP                                    =	-4165		# Variation, based on the whole population.

    #values from excel.xlcontainsoperator
    #****************************
    xlBeginsWith                              =	2		# Begins with a specified value.
    xlContains                                =	0		# Contains a specified value.
    xlDoesNotContain                          =	1		# Does not contain the specified value.
    xlEndsWith                                =	3		# Endswith the specified value

    #values from excel.xlcopypictureformat
    #****************************
    xlBitmap                                  =	2		# Bitmap (.bmp, .jpg, .gif).
    xlPicture                                 =	-4147		# Drawn picture (.png, .wmf, .mix).

    #values from excel.xlcorruptload
    #****************************
    xlExtractData                             =	2		# Workbook is opened in extract data mode.
    xlNormalLoad                              =	0		# Workbook is opened normally.
    xlRepairFile                              =	1		# Workbook is opened in repair mode.

    #values from excel.xlcreator
    #****************************
    xlCreatorCode                             =	1480803660		# The Excel for Macintosh creator code.

    #values from excel.xlcredentialsmethod
    #****************************
    CredentialsMethodIntegrated               =	0		# Integrated
    CredentialsMethodNone                     =	1		# No credentials used
    CredentialsMethodStored                   =	2		# Use stored credentials

    #values from excel.xlcubefieldsubtype
    #****************************
    xlCubeAttribute                           =	4		# Attribute
    xlCubeCalculatedMeasure                   =	5		# Calculated Measure
    xlCubeHierarchy                           =	1		# Hierarchy
    xlCubeImplicitMeasure                     =	11		# An implicit measure
    xlCubeKPIGoal                             =	7		# KPI Goal
    xlCubeKPIStatus                           =	8		# KPI Status
    xlCubeKPITrend                            =	9		# KPI Trend
    xlCubeKPIValue                            =	6		# KPI Value
    xlCubeKPIWeight                           =	10		# KPI Weight
    xlCubeMeasure                             =	2		# Measure
    xlCubeSet                                 =	3		# Set

    #values from excel.xlcubefieldtype
    #****************************
    xlHierarchy                               =	1		# OLAP field is a hierarchy.
    xlMeasure                                 =	2		# OLAP field is a measure.
    xlSet                                     =	3		# OLAP field is a set.

    #values from excel.xlcutcopymode
    #****************************
    xlCopy                                    =	1		# In Copy mode
    xlCut                                     =	2		# In Cut mode

    #values from excel.xldvalertstyle
    #****************************
    xlValidAlertInformation                   =	3		# Information icon.
    xlValidAlertStop                          =	1		# Stop icon.
    xlValidAlertWarning                       =	2		# Warning icon.

    #values from excel.xldvtype
    #****************************
    xlValidateCustom                          =	7		# Data is validated using an arbitrary formula.
    xlValidateDate                            =	4		# Date values.
    xlValidateDecimal                         =	2		# Numeric values.
    xlValidateInputOnly                       =	0		# Validate only when user changes the value.
    xlValidateList                            =	3		# Value must be present in a specified list.
    xlValidateTextLength                      =	6		# Length of text.
    xlValidateTime                            =	5		# Time values.
    xlValidateWholeNumber                     =	1		# Whole numeric values.

    #values from excel.xldatabaraxisposition
    #****************************
    xlDataBarAxisAutomatic                    =	0		# Display the axis at a variable position based on the ratio of the minimum negative value to the maximum positive value in the range. Positive values are displayed in a left-to-right direction. Negative values are displayed in a right-to-left direction. When all values are positive or all values are negative, no axis is displayed.
    xlDataBarAxisMidpoint                     =	1		# Display the axis at the midpoint of the cell regardless of the set of values in the range. Positive values are displayed in a left-to-right direction. Negative values are displayed in a right-to-left direction.
    xlDataBarAxisNone                         =	2		# No axis is displayed, and both positive and negative values are displayed in the left-to-right direction.

    #values from excel.xldatabarbordertype
    #****************************
    xlDataBarBorderNone                       =	0		# The data bar has no border.
    xlDataBarBorderSolid                      =	1		# The data bar has a solid border.

    #values from excel.xldatabarfilltype
    #****************************
    xlDataBarFillGradient                     =	1		# The data bar is filled with a color gradient.
    xlDataBarFillSolid                        =	0		# The data bar is filled with solid color.

    #values from excel.xldatabarnegativecolortype
    #****************************
    xlDataBarColor                            =	0		# Use the color specified in the  Negative Value and Axis Setting dialog box or by using the ColorType and BorderColorType properties of the NegativeBarFormat object.
    xlDataBarSameAsPositive                   =	1		# Use the same color as positive data bars.

    #values from excel.xldatalabelposition
    #****************************
    xlLabelPositionAbove                      =	0		# Data label is positioned above the data point.
    xlLabelPositionBelow                      =	1		# Data label is positioned below the data point.
    xlLabelPositionBestFit                    =	5		# Microsoft Office Excel 2007 sets the position of the data label.
    xlLabelPositionCenter                     =	-4108		# Data label is centered on the data point or is inside a bar or pie chart.
    xlLabelPositionCustom                     =	7		# Data label is in a custom position.
    xlLabelPositionInsideBase                 =	4		# Data label is positioned inside the data point at the bottom edge.
    xlLabelPositionInsideEnd                  =	3		# Data label is positioned inside the data point at the top edge.
    xlLabelPositionLeft                       =	-4131		# Data label is positioned to the left of the data point.
    xlLabelPositionMixed                      =	6		# Data labels are in multiple positions.
    xlLabelPositionOutsideEnd                 =	2		# Data label is positioned outside the data point at the top edge.
    xlLabelPositionRight                      =	-4152		# Data label is positioned to the right of the data point.

    #values from excel.xldatalabelseparator
    #****************************
    xlDataLabelSeparatorDefault               =	1		# Excel selects the separator.

    #values from excel.xldatalabelstype
    #****************************
    xlDataLabelsShowBubbleSizes               =	6		# Show the size of the bubble in reference to the absolute value.
    xlDataLabelsShowLabel                     =	4		# Category for the point.
    xlDataLabelsShowLabelAndPercent           =	5		# Percentage of the total, and category for the point. Available only for pie charts and doughnut charts.
    xlDataLabelsShowNone                      =	-4142		# No data labels.
    xlDataLabelsShowPercent                   =	3		# Percentage of the total. Available only for pie charts and doughnut charts.
    xlDataLabelsShowValue                     =	2		# Default value for the point (assumed if this argument is not specified).

    #values from excel.xldataseriesdate
    #****************************
    xlDay                                     =	1		# Day
    xlMonth                                   =	3		# Month
    xlWeekday                                 =	2		# Weekdays
    xlYear                                    =	4		# Year

    #values from excel.xldataseriestype
    #****************************
    xlAutoFill                                =	4		# Fill series according to AutoFill settings.
    xlChronological                           =	3		# Fill with date values.
    xlDataSeriesLinear                        =	-4132		# Extend values, assuming an additive progression (for example, '1, 2' is extended as '3, 4, 5').
    xlGrowth                                  =	2		# Extend values, assuming a multiplicative progression (for example, '1, 2' is extended as '4, 8, 16').

    #values from excel.xldeleteshiftdirection
    #****************************
    xlShiftToLeft                             =	-4159		# Cells are shifted to the left.
    xlShiftUp                                 =	-4162		# Cells are shifted up.

    #values from excel.xldirection
    #****************************
    xlDown                                    =	-4121		# Down.
    xlToLeft                                  =	-4159		# To left.
    xlToRight                                 =	-4161		# To right.
    xlUp                                      =	-4162		# Up.

    #values from excel.xldisplayblanksas
    #****************************
    xlInterpolated                            =	3		# Values are interpolated into the chart.
    xlNotPlotted                              =	1		# Blank cells are not plotted.
    xlZero                                    =	2		# Blanks are plotted as zero.

    #values from excel.xldisplaydrawingobjects
    #****************************
    xlDisplayShapes                           =	-4104		# Show all shapes.
    xlHide                                    =	3		# Hide all shapes.
    xlPlaceholders                            =	2		# Show only placeholders.

    #values from excel.xldisplayunit
    #****************************
    xlHundredMillions                         =	-8		# Hundreds of millions.
    xlHundreds                                =	-2		# Hundreds.
    xlHundredThousands                        =	-5		# Hundreds of thousands.
    xlMillionMillions                         =	-10		# Millions of millions.
    xlMillions                                =	-6		# Millions.
    xlTenMillions                             =	-7		# Tens of millions.
    xlTenThousands                            =	-4		# Tens of thousands.
    xlThousandMillions                        =	-9		# Thousands of millions.
    xlThousands                               =	-3		# Thousands.

    #values from excel.xldupeunique
    #****************************
    xlDuplicate                               =	1		# Display duplicate values.
    xlUnique                                  =	0		# Display unique values.

    #values from excel.xldynamicfiltercriteria
    #****************************
    xlFilterAboveAverage                      =	33		# Filter all above-average values.
    xlFilterAllDatesInPeriodApril             =	24		# Filter all dates in April.
    xlFilterAllDatesInPeriodAugust            =	28		# Filter all dates in August.
    xlFilterAllDatesInPeriodDecember          =	32		# Filter all dates in December.
    xlFilterAllDatesInPeriodFebruary          =	22		# Filter all dates in February.
    xlFilterAllDatesInPeriodJanuary           =	21		# Filter all dates in January.
    xlFilterAllDatesInPeriodJuly              =	27		# Filter all dates in July.
    xlFilterAllDatesInPeriodJune              =	26		# Filter all dates in June.
    xlFilterAllDatesInPeriodMarch             =	23		# Filter all dates in March.
    xlFilterAllDatesInPeriodMay               =	25		# Filter all dates in May.
    xlFilterAllDatesInPeriodNovember          =	31		# Filter all dates in November.
    xlFilterAllDatesInPeriodOctober           =	30		# Filter all dates in October.
    xlFilterAllDatesInPeriodQuarter1          =	17		# Filter all dates in Quarter1.
    xlFilterAllDatesInPeriodQuarter2          =	18		# Filter all dates in Quarter2.
    xlFilterAllDatesInPeriodQuarter3          =	19		# Filter all dates in Quarter3.
    xlFilterAllDatesInPeriodQuarter4          =	20		# Filter all dates in Quarter4.
    xlFilterAllDatesInPeriodSeptember         =	29		# Filter all dates in September.
    xlFilterBelowAverage                      =	34		# Filter all below-average values.
    xlFilterLastMonth                         =	8		# Filter all values related to last month.
    xlFilterLastQuarter                       =	11		# Filter all values related to last quarter.
    xlFilterLastWeek                          =	5		# Filter all values related to last week.
    xlFilterLastYear                          =	14		# Filter all values related to last year.
    xlFilterNextMonth                         =	9		# Filter all values related to next month.
    xlFilterNextQuarter                       =	12		# Filter all values related to next quarter.
    xlFilterNextWeek                          =	6		# Filter all values related to next week.
    xlFilterNextYear                          =	15		# Filter all values related to next year.
    xlFilterThisMonth                         =	7		# Filter all values related to the current month.
    xlFilterThisQuarter                       =	10		# Filter all values related to the current quarter.
    xlFilterThisWeek                          =	4		# Filter all values related to the current week.
    xlFilterThisYear                          =	13		# Filter all values related to the current year.
    xlFilterToday                             =	1		# Filter all values related to the current date.
    xlFilterTomorrow                          =	3		# Filter all values related to tomorrow.
    xlFilterYearToDate                        =	16		# Filter all values from today until a year ago.
    xlFilterYesterday                         =	2		# Filter all values related to yesterday.

    #values from excel.xleditionformat
    #****************************
    xlBIFF                                    =	2		# Binary Interchange file format.
    xlPICT                                    =	1		# Metafile picture structure (.wmf).
    xlRTF                                     =	4		# Rich Text Format (.rtf).
    xlVALU                                    =	8		# VALU.

    #values from excel.xleditionoptionsoption
    #****************************
    xlAutomaticUpdate                         =	4		# Automatic update.
    xlCancel                                  =	1		# Cancel.
    xlChangeAttributes                        =	6		# Change attributes.
    xlManualUpdate                            =	5		# Manual update.
    xlOpenSource                              =	3		# Open source.
    xlSelect                                  =	3		# Select.
    xlSendPublisher                           =	2		# Send to Microsoft Publisher.
    xlUpdateSubscriber                        =	2		# Update subscriber.

    #values from excel.xleditiontype
    #****************************
    xlPublisher                               =	1		# Publisher
    xlSubscriber                              =	2		# Subscriber

    #values from excel.xlenablecancelkey
    #****************************
    xlDisabled                                =	0		# Cancel key trapping is completely disabled.
    xlErrorHandler                            =	2		# The interrupt is sent to the running procedure as an error, trappable by an error handler set up with an On Error GoTo statement. The trappable error code is 18.
    xlInterrupt                               =	1		# The current procedure is interrupted, and the user can debug or end the procedure.

    #values from excel.xlenableselection
    #****************************
    xlNoRestrictions                          =	0		# Anything can be selected.
    xlNoSelection                             =	-4142		# Nothing can be selected.
    xlUnlockedCells                           =	1		# Only unlocked cells can be selected.

    #values from excel.xlendstylecap
    #****************************
    xlCap                                     =	1		# Caps applied.
    xlNoCap                                   =	2		# No caps applied.

    #values from excel.xlerrorbardirection
    #****************************
    xlX                                       =	-4168		# Bars run parallel to the Y axis for X-axis values.
    xlY                                       =	1		# Bars run parallel to the X axis for Y-axis values.

    #values from excel.xlerrorbarinclude
    #****************************
    xlErrorBarIncludeBoth                     =	1		# Both positive and negative error range.
    xlErrorBarIncludeMinusValues              =	3		# Only negative error range.
    xlErrorBarIncludeNone                     =	-4142		# No error bar range.
    xlErrorBarIncludePlusValues               =	2		# Only positive error range.

    #values from excel.xlerrorbartype
    #****************************
    xlErrorBarTypeCustom                      =	-4114		# Range is set by fixed values or cell values.
    xlErrorBarTypeFixedValue                  =	1		# Fixed-length error bars.
    xlErrorBarTypePercent                     =	2		# Percentage of range to be covered by the error bars.
    xlErrorBarTypeStDev                       =	-4155		# Shows range for specified number of standard deviations.
    xlErrorBarTypeStError                     =	4		# Shows standard error range.

    #values from excel.xlerrorchecks
    #****************************
    xlEmptyCellReferences                     =	7		# The cell contains a formula referring to empty cells.
    xlEvaluateToError                         =	1		# The cell evaluates to an error value.
    xlInconsistentFormula                     =	4		# The cell contains an inconsistent formula for a region.
    xlInconsistentListFormula                 =	9		# The cell contains an inconsistent formula for a list.
    xlListDataValidation                      =	8		# Data in the list contains a validation error.
    xlNumberAsText                            =	3		# Number entered as text.
    xlOmittedCells                            =	5		# Cells omitted.
    xlTextDate                                =	2		# Date entered as text.
    xlUnlockedFormulaCells                    =	6		# Formula cells are unlocked.

    #values from excel.xlfileaccess
    #****************************
    xlReadOnly                                =	3		# Read only.
    xlReadWrite                               =	2		# Read/write.

    #values from excel.xlfileformat
    #****************************
    xlAddIn                                   =	18		# Microsoft Excel 97-2003 Add-In

    #values from excel.xlfilevalidationpivotmode
    #****************************
    xlFileValidationPivotDefault              =	0		# Validate the contents of data caches as specified by the  PivotOptions registry setting (default).
    xlFileValidationPivotRun                  =	1		# Validate the contents of all data caches regardless of the registry setting.
    xlFileValidationPivotSkip                 =	2		# Do not validate the contents of data caches.

    #values from excel.xlfillwith
    #****************************
    xlFillWithAll                             =	-4104		# Copy contents and formats.
    xlFillWithContents                        =	2		# Copy contents only.
    xlFillWithFormats                         =	-4122		# Copy formats only.

    #values from excel.xlfilteraction
    #****************************
    xlFilterCopy                              =	2		# Copy filtered data to new location.
    xlFilterInPlace                           =	1		# Leave data in place.

    #values from excel.xlfilteralldatesinperiod
    #****************************
    xlFilterAllDatesInPeriodDay               =	2		# Filter all dates for the specified date.
    xlFilterAllDatesInPeriodHour              =	3		# Filter all dates for the specified hour.
    xlFilterAllDatesInPeriodMinute            =	4		# Filter all dates until the specified minute.
    xlFilterAllDatesInPeriodMonth             =	1		# Filter all dates for the specified month.
    xlFilterAllDatesInPeriodSecond            =	5		# Filter all dates until the specified second.
    xlFilterAllDatesInPeriodYear              =	0		# Filter all dates for the specified year.

    #values from excel.xlfilterstatus
    #****************************
    xlFilterStatusOK                          =	0		# Signifies OK or successful.
    xlFilterStatusDateWrongOrder              =	1		# SetFilterDateRange(?): StartDate &gt; EndDate
    xlFilterStatusDateHasTime                 =	2		# SetFilterDateRange(?): StartDate or EndDate have a time portion.
    xlFilterStatusInvalidDate                 =	3		# SetFilterDateRange(?): StartDate or EndDate are not valid dates.

    #values from excel.xlfindlookin
    #****************************
    xlComments                                =	-4144		# Comments.
    xlFormulas                                =	-4123		# Formulas.
    xlValues                                  =	-4163		# Values.

    #values from excel.xlfixedformatquality
    #****************************
    xlQualityMinimum                          =	1		# Minimum quality
    xlQualityStandard                         =	0		# Standard quality

    #values from excel.xlfixedformattype
    #****************************
    xlTypePDF                                 =	0		# &quot;PDF&quot; ? Portable Document Format file (.pdf).
    xlTypeXPS                                 =	1		# &quot;XPS&quot; ? XPS Document (.xps).

    #values from excel.xlforecastaggregation
    #****************************
    xlForecastAggregationCount                =	2

    #values from excel.xlforecastcharttype
    #****************************
    xlForecastChartTypeLine                   =	0

    #values from excel.xlforecastdatacompletion
    #****************************
    xlForecastDataCompletionZeros             =	0

    #values from excel.xlformcontrol
    #****************************
    xlButtonControl                           =	0		# Button.
    xlCheckBox                                =	1		# Check box.
    xlDropDown                                =	2		# Combo box.
    xlEditBox                                 =	3		# Text box.
    xlGroupBox                                =	4		# Group box.
    xlLabel                                   =	5		# Label.
    xlListBox                                 =	6		# List box.
    xlOptionButton                            =	7		# Option button.
    xlScrollBar                               =	8		# Scroll bar.
    xlSpinner                                 =	9		# Spinner.

    #values from excel.xlformatconditionoperator
    #****************************
    xlBetween                                 =	1		# Between. Can be used only if two formulas are provided.
    xlEqual                                   =	3		# Equal.
    xlGreater                                 =	5		# Greater than.
    xlGreaterEqual                            =	7		# Greater than or equal to.
    xlLess                                    =	6		# Less than.
    xlLessEqual                               =	8		# Less than or equal to.
    xlNotBetween                              =	2		# Not between. Can be used only if two formulas are provided.
    xlNotEqual                                =	4		# Not equal.

    #values from excel.xlformatconditiontype
    #****************************
    xlAboveAverageCondition                   =	12		# Above average condition
    xlBlanksCondition                         =	10		# Blanks condition
    xlCellValue                               =	1		# Cell value
    xlColorScale                              =	3		# Color scale
    xlDatabar                                 =	4		# Databar
    xlErrorsCondition                         =	16		# Errors condition
    xlExpression                              =	2		# Expression
    xlIconSet                                 =	6		# Icon set
    xlNoBlanksCondition                       =	13		# No blanks condition
    xlNoErrorsCondition                       =	17		# No errors condition
    xlTextString                              =	9		# Text string
    xlTimePeriod                              =	11		# Time period
    xlTop10                                   =	5		# Top 10 values
    xlUniqueValues                            =	8		# Unique values

    #values from excel.xlformatfiltertypes
    #****************************
    FilterBottom                              =	0		# Bottom.
    FilterBottomPercent                       =	2		# Bottom Percent.
    FilterTop                                 =	1		# Top.
    FilterTopPercent                          =	3		# Top Percent.

    #values from excel.xlformulalabel
    #****************************
    xlColumnLabels                            =	2		# Column labels only.
    xlMixedLabels                             =	3		# Row and column labels.
    xlNoLabels                                =	-4142		# No labels.
    xlRowLabels                               =	1		# Row labels only.

    #values from excel.xlgeneratetablerefs
    #****************************
    xlA1TableRefs                             =	0		# A1 Table References.
    xlTableNames                              =	1		# Table Names.

    #values from excel.xlgradientfilltype
    #****************************
    GradientFillLinear                        =	0		# Gradient is filled in a straight line.
    GradientFillPath                          =	1		# Gradient is filled in a non-linear or curved path.

    #values from excel.xlhalign
    #****************************
    xlHAlignCenter                            =	-4108		# Center.
    xlHAlignCenterAcrossSelection             =	7		# Center across selection.
    xlHAlignDistributed                       =	-4117		# Distribute.
    xlHAlignFill                              =	5		# Fill.
    xlHAlignGeneral                           =	1		# Align according to data type.
    xlHAlignJustify                           =	-4130		# Justify.
    xlHAlignLeft                              =	-4131		# Left.
    xlHAlignRight                             =	-4152		# Right.

    #values from excel.xlhebrewmodes
    #****************************
    xlHebrewFullScript                        =	0		# The conventional script type as required by the Hebrew Language Academy when writing text without diacritics.
    xlHebrewMixedAuthorizedScript             =	3		# The Hebrew traditional script.
    xlHebrewMixedScript                       =	2		# In this mode the speller accepts any word recognized as Hebrew, whether in Full Script, Partial Script, or any unconventional spelling variation that is known to the speller.
    xlHebrewPartialScript                     =	1		# In this mode the speller accepts words both in Full Script and Partial Script. Some words will be flagged since this spelling is not authorized in either Full script or Partial script.

    #values from excel.xlhighlightchangestime
    #****************************
    xlAllChanges                              =	2		# Show all changes.
    xlNotYetReviewed                          =	3		# Show only changes not yet reviewed.
    xlSinceMyLastSave                         =	1		# Show changes made since last save by last user.

    #values from excel.xlhtmltype
    #****************************
    xlHtmlCalc                                =	1		# Use the Spreadsheet component. Deprecated.
    xlHtmlChart                               =	3		# Use the Chart component. Deprecated.
    xlHtmlList                                =	2		# Use the PivotTable component. Deprecated.
    xlHtmlStatic                              =	0		# Use static (noninteractive) HTML for viewing only.

    #values from excel.xlimemode
    #****************************
    xlIMEModeAlpha                            =	8		# Half-width alphanumeric.
    xlIMEModeAlphaFull                        =	7		# Full-width alphanumeric.
    xlIMEModeDisable                          =	3		# Disable.
    xlIMEModeHangul                           =	10		# Hangul.
    xlIMEModeHangulFull                       =	9		# Full-width Hangul.
    xlIMEModeHiragana                         =	4		# Hiragana.
    xlIMEModeKatakana                         =	5		# Katakana.
    xlIMEModeKatakanaHalf                     =	6		# Half-width Katakana.
    xlIMEModeNoControl                        =	0		# No control.
    xlIMEModeOff                              =	2		# Off (English mode).
    xlIMEModeOn                               =	1		# Mode on.

    #values from excel.xlicon
    #****************************
    xlIcon0Bars                               =	37		# Signal Meter With No Filled Bars
    xlIcon0FilledBoxes                        =	52		# 0 Filled Boxes
    xlIcon1Bar                                =	38		# Signal Meter With One Filled Bar
    xlIcon1FilledBox                          =	51		# 1 Filled Boxes
    xlIcon2Bars                               =	39		# Signal Meter With Two Filled Bars
    xlIcon2FilledBoxes                        =	50		# 2 Filled Boxes
    xlIcon3Bars                               =	40		# Signal Meter With Three Filled Bars
    xlIcon3FilledBoxes                        =	49		# 3 Filled Boxes
    xlIcon4Bars                               =	41		# Signal Meter With Four Filled Bars
    xlIcon4FilledBoxes                        =	48		# 4 Filled Boxes
    xlIconBlackCircle                         =	32		# Black Circle
    xlIconBlackCircleWithBorder               =	13		# Black Circle With Border
    xlIconCircleWithOneWhiteQuarter           =	33		# Circle With One White Quarter
    xlIconCircleWithThreeWhiteQuarters        =	35		# Circle With Three White Quarters
    xlIconCircleWithTwoWhiteQuarters          =	34		# Circle With Two White Quarters
    xlIconGoldStar                            =	42		# Gold Star
    xlIconGrayCircle                          =	31		# Gray Circle
    xlIconGrayDownArrow                       =	6		# Gray Down Arrow
    xlIconGrayDownInclineArrow                =	28		# Gray Down Incline Arrow
    xlIconGraySideArrow                       =	5		# Gray Side Arrow
    xlIconGrayUpArrow                         =	4		# Gray Up Arrow
    xlIconGrayUpInclineArrow                  =	27		# Gray Up Incline Arrow
    xlIconGreenCheck                          =	22		# Green Check
    xlIconGreenCheckSymbol                    =	19		# Green Check Symbol
    xlIconGreenCircle                         =	10		# Green Circle
    xlIconGreenFlag                           =	7		# Green Flag
    xlIconGreenTrafficLight                   =	14		# Green Traffic Light
    xlIconGreenUpArrow                        =	1		# Green Up Arrow
    xlIconGreenUpTriangle                     =	45		# Green Up Triangle
    xlIconHalfGoldStar                        =	43		# Half Gold Star
    xlIconNoCellIcon                          =	-1		# No Cell Icon
    xlIconPinkCircle                          =	30		# Pink Circle
    xlIconRedCircle                           =	29		# Red Circle
    xlIconRedCircleWithBorder                 =	12		# Red Circle With Border
    xlIconRedCross                            =	24		# Red Cross
    xlIconRedCrossSymbol                      =	21		# Red Cross Symbol
    xlIconRedDiamond                          =	18		# Red Diamond
    xlIconRedDownArrow                        =	3		# Red Down Arrow
    xlIconRedDownTriangle                     =	47		# Red Down Triangle
    xlIconRedFlag                             =	9		# Red Flag
    xlIconRedTrafficLight                     =	16		# Red Traffic Light
    xlIconSilverStar                          =	44		# Silver Star
    xlIconWhiteCircleAllWhiteQuarters         =	36		# White Circle (All White Quarters)
    xlIconYellowCircle                        =	11		# Yellow Circle
    xlIconYellowDash                          =	46		# Yellow Dash
    xlIconYellowDownInclineArrow              =	26		# Yellow Down Incline Arrow
    xlIconYellowExclamation                   =	23		# Yellow Exclamation
    xlIconYellowExclamationSymbol             =	20		# Yellow Exclamation Symbol
    xlIconYellowFlag                          =	8		# Yellow Flag
    xlIconYellowSideArrow                     =	2		# Yellow Side Arrow
    xlIconYellowTrafficLight                  =	15		# Yellow Traffic Light
    xlIconYellowTriangle                      =	17		# Yellow Triangle
    xlIconYellowUpInclineArrow                =	25		# Yellow Up Incline Arrow

    #values from excel.xliconset
    #****************************
    xl3Arrows                                 =	1		# 3 Arrows
    xl3ArrowsGray                             =	2		# 3 Arrows Gray
    xl3Flags                                  =	3		# 3 Flags
    xl3Signs                                  =	6		# 3 Signs
    xl3Symbols                                =	7		# 3 Symbols
    xl3TrafficLights1                         =	4		# 3 Traffic Lights 1
    xl3TrafficLights2                         =	5		# 3 Traffic Lights 2
    xl4Arrows                                 =	8		# 4 Arrows
    xl4ArrowsGray                             =	9		# 4 Arrows Gray
    xl4CRV                                    =	11		# 4 CRV
    xl4RedToBlack                             =	10		# 4 Red To Black
    xl4TrafficLights                          =	12		# 4 Traffic Lights
    xl5Arrows                                 =	13		# 5 Arrows
    xl5ArrowsGray                             =	14		# 5 Arrows Gray
    xl5CRV                                    =	15		# 5 CRV
    xl5Quarters                               =	16		# 5 Quarters

    #values from excel.xlimportdataas
    #****************************
    xlPivotTableReport                        =	1		# Returns the data as a PivotTable.
    xlQueryTable                              =	0		# Returns the data as a QueryTable.

    #values from excel.xlinsertformatorigin
    #****************************
    xlFormatFromLeftOrAbove                   =	0		# Copy the format from cells above and/or to the left.
    xlFormatFromRightOrBelow                  =	1		# Copy the format from cells below and/or to the right.

    #values from excel.xlinsertshiftdirection
    #****************************
    xlShiftDown                               =	-4121		# Shift cells down.
    xlShiftToRight                            =	-4161		# Shift cells to the right.

    #values from excel.xllayoutformtype
    #****************************
    xlOutline                                 =	1		# The  LayoutSubtotalLocation property specifies where the subtotal appears in the PivotTable report.
    xlTabular                                 =	0		# Default.

    #values from excel.xllayoutrowtype
    #****************************
    xlCompactRow                              =	0		# Compact Row
    xlOutlineRow                              =	2		# Outline Row
    xlTabularRow                              =	1		# Tabular Row

    #values from excel.xllegendposition
    #****************************
    xlLegendPositionBottom                    =	-4107		# Below the chart.
    xlLegendPositionCorner                    =	2		# In the upper-right corner of the chart border.
    xlLegendPositionCustom                    =	-4161		# A custom position.
    xlLegendPositionLeft                      =	-4131		# Left of the chart.
    xlLegendPositionRight                     =	-4152		# Right of the chart.
    xlLegendPositionTop                       =	-4160		# Above the chart.

    #values from excel.xllinestyle
    #****************************
    xlContinuous                              =	1		# Continuous line.
    xlDash                                    =	-4115		# Dashed line.
    xlDashDot                                 =	4		# Alternating dashes and dots.
    xlDashDotDot                              =	5		# Dash followed by two dots.
    xlDot                                     =	-4118		# Dotted line.
    xlDouble                                  =	-4119		# Double line.
    xlLineStyleNone                           =	-4142		# No line.
    xlSlantDashDot                            =	13		# Slanted dashes.

    #values from excel.xllink
    #****************************
    xlExcelLinks                              =	1		# The link is to an Excel worksheet.
    xlOLELinks                                =	2		# The link is to an OLE source.
    xlPublishers                              =	5		# Macintosh only.
    xlSubscribers                             =	6		# Macintosh only.

    #values from excel.xllinkinfo
    #****************************
    xlEditionDate                             =	2		# Applies only to editions in the Macintosh operating system.
    xlLinkInfoStatus                          =	3		# Returns the link status.
    xlUpdateState                             =	1		# Specifies whether the link updates automatically or manually.

    #values from excel.xllinkinfotype
    #****************************
    xlLinkInfoOLELinks                        =	2		# OLE or DDE server
    xlLinkInfoPublishers                      =	5		# Publisher
    xlLinkInfoSubscribers                     =	6		# Subscriber

    #values from excel.xllinkstatus
    #****************************
    xlLinkStatusCopiedValues                  =	10		# Copied values.
    xlLinkStatusIndeterminate                 =	5		# Unable to determine status.
    xlLinkStatusInvalidName                   =	7		# Invalid name.
    xlLinkStatusMissingFile                   =	1		# File missing.
    xlLinkStatusMissingSheet                  =	2		# Sheet missing.
    xlLinkStatusNotStarted                    =	6		# Not started.
    xlLinkStatusOK                            =	0		# No errors.
    xlLinkStatusOld                           =	3		# Status may be out of date.
    xlLinkStatusSourceNotCalculated           =	4		# Not yet calculated.
    xlLinkStatusSourceNotOpen                 =	8		# Not open.
    xlLinkStatusSourceOpen                    =	9		# Source document is open.

    #values from excel.xllinktype
    #****************************
    xlLinkTypeExcelLinks                      =	1		# A link to a Microsoft Excel source.
    xlLinkTypeOLELinks                        =	2		# A link to an OLE source.

    #values from excel.xllinkeddatatypestate
    #****************************
    xlLinkedDataTypeStateNone                 =	0		# The cell does not contain any Linked data types.
    xlLinkedDataTypeStateValidLinkedData      =	1		# The cell contains a Linked data type.
    xlLinkedDataTypeStateDisambiguationNeeded	=	2		# The cell needs to be disambiguated by the user before a Linked data type can be inserted. For example, if the user types &quot;New York&quot; into a cell and attempts to convert it to a &quot;Geography&quot; data type, they may need to select whether they meant New York State or New York City. Until they do so, the cell will be in this state.
    xlLinkedDataTypeStateBrokenLinkedData     =	3		# There is a valid Linked data type in the cell, but entity no longer exists on the service.
    xlLinkedDataTypeStateFetchingData         =	4		# The Linked data type in the cell is in the middle of refreshing new data from the service.

    #values from excel.xllistconflict
    #****************************
    xlListConflictDialog                      =	0		# Display a dialog box that allows the user to choose how to resolve conflicts.
    xlListConflictDiscardAllConflicts         =	2		# Accept the version of the data stored on the SharePoint site.
    xlListConflictError                       =	3		# Raise an error if a conflict occurs.
    xlListConflictRetryAllConflicts           =	1		# Overwrite the version of the data stored on the SharePoint site.

    #values from excel.xllistdatatype
    #****************************
    xlListDataTypeCheckbox                    =	9		# Check box.
    xlListDataTypeChoice                      =	6		# Single-choice field.
    xlListDataTypeChoiceMulti                 =	7		# Multiple-choice field.
    xlListDataTypeCounter                     =	11		# Counter.
    xlListDataTypeCurrency                    =	4		# Currency.
    xlListDataTypeDateTime                    =	5		# Date/time.
    xlListDataTypeHyperLink                   =	10		# Hyperlink.
    xlListDataTypeListLookup                  =	8		# Lookup list.
    xlListDataTypeMultiLineRichText           =	12		# Rich text format with multiple lines.
    xlListDataTypeMultiLineText               =	2		# Plain text with multiple lines.
    xlListDataTypeNone                        =	0		# Type not specified.
    xlListDataTypeNumber                      =	3		# Numerical.
    xlListDataTypeText                        =	1		# Plain text.

    #values from excel.xllistobjectsourcetype
    #****************************
    xlSrcExternal                             =	0		# External data source (Microsoft SharePoint Foundation site).
    xlSrcModel                                =	4		# PowerPivot Model
    xlSrcQuery                                =	3		# Query
    xlSrcRange                                =	1		# Range
    xlSrcXml                                  =	2		# XML

    #values from excel.xllocationintable
    #****************************
    xlColumnHeader                            =	-4110		# Column header
    xlColumnItem                              =	5		# Column item
    xlDataHeader                              =	3		# Data header
    xlDataItem                                =	7		# Data item
    xlPageHeader                              =	2		# Page header
    xlPageItem                                =	6		# Page item
    xlRowHeader                               =	-4153		# Row header
    xlRowItem                                 =	4		# Row item
    xlTableBody                               =	8		# Table body

    #values from excel.xllookat
    #****************************
    xlPart                                    =	2		# Match against any part of the search text.
    xlWhole                                   =	1		# Match against the whole of the search text.

    #values from excel.xllookfor
    #****************************
    LookForBlanks                             =	0		# Blanks
    LookForErrors                             =	1		# Errors
    LookForFormulas                           =	2		# Formulas

    #values from excel.xlmsapplication
    #****************************
    xlMicrosoftAccess                         =	4		# Microsoft Office Access
    xlMicrosoftFoxPro                         =	5		# Microsoft FoxPro
    xlMicrosoftMail                           =	3		# Microsoft Office Outlook
    xlMicrosoftPowerPoint                     =	2		# Microsoft Office PowerPoint
    xlMicrosoftProject                        =	6		# Microsoft Office Project
    xlMicrosoftSchedulePlus                   =	7		# Microsoft Schedule Plus
    xlMicrosoftWord                           =	1		# Microsoft Office Word

    #values from excel.xlmailsystem
    #****************************
    xlMAPI                                    =	1		# MAPI-complaint system
    xlNoMailSystem                            =	0		# No mail system
    xlPowerTalk                               =	2		# PowerTalk mail system

    #values from excel.xlmarkerstyle
    #****************************
    xlMarkerStyleAutomatic                    =	-4105		# Automatic markers
    xlMarkerStyleCircle                       =	8		# Circular markers
    xlMarkerStyleDash                         =	-4115		# Long bar markers
    xlMarkerStyleDiamond                      =	2		# Diamond-shaped markers
    xlMarkerStyleDot                          =	-4118		# Short bar markers
    xlMarkerStyleNone                         =	-4142		# No markers
    xlMarkerStylePicture                      =	-4147		# Picture markers
    xlMarkerStylePlus                         =	9		# Square markers with a plus sign
    xlMarkerStyleSquare                       =	1		# Square markers
    xlMarkerStyleStar                         =	5		# Square markers with an asterisk
    xlMarkerStyleTriangle                     =	3		# Triangular markers
    xlMarkerStyleX                            =	-4168		# Square markers with an X

    #values from excel.xlmeasurementunits
    #****************************
    xlCentimeters                             =	1		# Centimeters
    xlInches                                  =	0		# Inches
    xlMillimeters                             =	2		# Millimeters

    #values from excel.xlmodelchangesource
    #****************************
    xlChangeByExcel                           =	0		# Excel
    xlChangeByPowerPivotAddIn                 =	1		# PowerPivot add-in

    #values from excel.xlmousebutton
    #****************************
    xlNoButton                                =	0		# No button was pressed.
    xlPrimaryButton                           =	1		# The primary button (normally the left mouse button) was pressed.
    xlSecondaryButton                         =	2		# The secondary button (normally the right mouse button) was pressed.

    #values from excel.xlmousepointer
    #****************************
    xlDefault                                 =	-4143		# The default pointer.
    xlIBeam                                   =	3		# The I-beam pointer.
    xlNorthwestArrow                          =	1		# The northwest-arrow pointer.
    xlWait                                    =	2		# The hourglass pointer.

    #values from excel.xloletype
    #****************************
    xlOLEControl                              =	2		# ActiveX control
    xlOLEEmbed                                =	1		# Embedded OLE object
    xlOLELink                                 =	0		# Linked OLE object

    #values from excel.xloleverb
    #****************************
    xlVerbOpen                                =	2		# Open the object.
    xlVerbPrimary                             =	1		# Perform the primary action for the server.

    #values from excel.xloarthorizontaloverflow
    #****************************
    xlOartHorizontalOverflowClip              =	1		# Hide text that does not fit horizontally in the text frame.
    xlOartHorizontalOverflowOverflow          =	0		# Allow text to overflow the text frame horizontally.

    #values from excel.xloartverticaloverflow
    #****************************
    xlOartVerticalOverflowClip                =	1		# Hide text that does not fit vertically within the text frame.
    xlOartVerticalOverflowEllipsis            =	2		# Hide text that does not fit vertically within the text frame, and add an ellipsis (...) at the end of the visible text.
    xlOartVerticalOverflowOverflow            =	0		# Allow text to overflow the text frame vertically (can be from the top, bottom, or both depending on the text alignment).

    #values from excel.xlobjectsize
    #****************************
    xlFitToPage                               =	2		# Print the chart as large as possible, while retaining the chart's height-to-width ratio as shown on the screen.
    xlFullPage                                =	3		# Print the chart to fit the page, adjusting the height-to-width ratio as necessary.
    xlScreenSize                              =	1		# Print the chart the same size as it appears on the screen.

    #values from excel.xlorder
    #****************************
    xlDownThenOver                            =	1		# Process down the rows before processing across pages or page fields to the right.
    xlOverThenDown                            =	2		# Process across pages or page fields to the right before moving down the rows.

    #values from excel.xlorientation
    #****************************
    xlDownward                                =	-4170		# Text runs downward.
    xlHorizontal                              =	-4128		# Text runs horizontally.
    xlUpward                                  =	-4171		# Text runs upward.
    xlVertical                                =	-4166		# Text runs downward and is centered in the cell.

    #values from excel.xlptselectionmode
    #****************************
    xlBlanks                                  =	4		# Blanks
    xlButton                                  =	15		# Buttons
    xlDataAndLabel                            =	0		# Data and labels
    xlDataOnly                                =	2		# Data
    xlFirstRow                                =	256		# First row
    xlLabelOnly                               =	1		# Label
    xlOrigin                                  =	3		# Origin

    #values from excel.xlpagebreak
    #****************************
    xlPageBreakAutomatic                      =	-4105		# Excel will automatically add page breaks.
    xlPageBreakManual                         =	-4135		# Page breaks are manually inserted.
    xlPageBreakNone                           =	-4142		# Page breaks are not inserted in the worksheet.

    #values from excel.xlpagebreakextent
    #****************************
    xlPageBreakFull                           =	1		# Full screen.
    xlPageBreakPartial                        =	2		# Only within print area.

    #values from excel.xlpageorientation
    #****************************
    xlLandscape                               =	2		# Landscape mode.
    xlPortrait                                =	1		# Portrait mode.

    #values from excel.xlpapersize
    #****************************
    xlPaper10x14                              =	16		# 10 in. x 14 in.
    xlPaper11x17                              =	17		# 11 in. x 17 in.
    xlPaperA3                                 =	8		# A3 (297 mm x 420 mm)
    xlPaperA4                                 =	9		# A4 (210 mm x 297 mm)
    xlPaperA4Small                            =	10		# A4 Small (210 mm x 297 mm)
    xlPaperA5                                 =	11		# A5 (148 mm x 210 mm)
    xlPaperB4                                 =	12		# B4 (250 mm x 354 mm)
    xlPaperB5                                 =	13		# A5 (148 mm x 210 mm)
    xlPaperCsheet                             =	24		# C size sheet
    xlPaperDsheet                             =	25		# D size sheet
    xlPaperEnvelope10                         =	20		# Envelope #10 (4-1/8 in. x 9-1/2 in.)
    xlPaperEnvelope11                         =	21		# Envelope #11 (4-1/2 in. x 10-3/8 in.)
    xlPaperEnvelope12                         =	22		# Envelope #12 (4-1/2 in. x 11 in.)
    xlPaperEnvelope14                         =	23		# Envelope #14 (5 in. x 11-1/2 in.)
    xlPaperEnvelope9                          =	19		# Envelope #9 (3-7/8 in. x 8-7/8 in.)
    xlPaperEnvelopeB4                         =	33		# Envelope B4 (250 mm x 353 mm)
    xlPaperEnvelopeB5                         =	34		# Envelope B5 (176 mm x 250 mm)
    xlPaperEnvelopeB6                         =	35		# Envelope B6 (176 mm x 125 mm)
    xlPaperEnvelopeC3                         =	29		# Envelope C3 (324 mm x 458 mm)
    xlPaperEnvelopeC4                         =	30		# Envelope C4 (229 mm x 324 mm)
    xlPaperEnvelopeC5                         =	28		# Envelope C5 (162 mm x 229 mm)
    xlPaperEnvelopeC6                         =	31		# Envelope C6 (114 mm x 162 mm)
    xlPaperEnvelopeC65                        =	32		# Envelope C65 (114 mm x 229 mm)
    xlPaperEnvelopeDL                         =	27		# Envelope DL (110 mm x 220 mm)
    xlPaperEnvelopeItaly                      =	36		# Envelope (110 mm x 230 mm)
    xlPaperEnvelopeMonarch                    =	37		# Envelope Monarch (3-7/8 in. x 7-1/2 in.)
    xlPaperEnvelopePersonal                   =	38		# Envelope (3-5/8 in. x 6-1/2 in.)
    xlPaperEsheet                             =	26		# E size sheet
    xlPaperExecutive                          =	7		# Executive (7-1/2 in. x 10-1/2 in.)
    xlPaperFanfoldLegalGerman                 =	41		# German Legal Fanfold (8-1/2 in. x 13 in.)
    xlPaperFanfoldStdGerman                   =	40		# German Legal Fanfold (8-1/2 in. x 13 in.)
    xlPaperFanfoldUS                          =	39		# U.S. Standard Fanfold (14-7/8 in. x 11 in.)
    xlPaperFolio                              =	14		# Folio (8-1/2 in. x 13 in.)
    xlPaperLedger                             =	4		# Ledger (17 in. x 11 in.)
    xlPaperLegal                              =	5		# Legal (8-1/2 in. x 14 in.)
    xlPaperLetter                             =	1		# Letter (8-1/2 in. x 11 in.)
    xlPaperLetterSmall                        =	2		# Letter Small (8-1/2 in. x 11 in.)
    xlPaperNote                               =	18		# Note (8-1/2 in. x 11 in.)
    xlPaperQuarto                             =	15		# Quarto (215 mm x 275 mm)
    xlPaperStatement                          =	6		# Statement (5-1/2 in. x 8-1/2 in.)
    xlPaperTabloid                            =	3		# Tabloid (11 in. x 17 in.)
    xlPaperUser                               =	256		# User-defined

    #values from excel.xlparameterdatatype
    #****************************
    xlParamTypeBigInt                         =	-5		# Big integer.
    xlParamTypeBinary                         =	-2		# Binary.
    xlParamTypeBit                            =	-7		# Bit.
    xlParamTypeChar                           =	1		# String.
    xlParamTypeDate                           =	9		# Date.
    xlParamTypeDecimal                        =	3		# Decimal.
    xlParamTypeDouble                         =	8		# Double.
    xlParamTypeFloat                          =	6		# Float.
    xlParamTypeInteger                        =	4		# Integer.
    xlParamTypeLongVarBinary                  =	-4		# Long binary.
    xlParamTypeLongVarChar                    =	-1		# Long string.
    xlParamTypeNumeric                        =	2		# Numeric.
    xlParamTypeReal                           =	7		# Real.
    xlParamTypeSmallInt                       =	5		# Small integer.
    xlParamTypeTime                           =	10		# Time.
    xlParamTypeTimestamp                      =	11		# Time stamp.
    xlParamTypeTinyInt                        =	-6		# Tiny integer.
    xlParamTypeUnknown                        =	0		# Type unknown.
    xlParamTypeVarBinary                      =	-3		# Variable-length binary.
    xlParamTypeVarChar                        =	12		# Variable-length string.
    xlParamTypeWChar                          =	-8		# Unicode character string.

    #values from excel.xlparametertype
    #****************************
    xlConstant                                =	1		# Uses the value specified by the Value argument.
    xlPrompt                                  =	0		# Displays a dialog box that prompts the user for the value. The Value argument specifies the text shown in the dialog box.
    xlRange                                   =	2		# Uses the value of the cell in the upper-left corner of the range. The Value argument specifies a Range object.

    #values from excel.xlparentdatalabeloptions
    #****************************
    xlParentDataLabelOptionsBanner            =	1		# Banner parent data label
    xlParentDataLabelOptionsNone              =	0		# No parent data label
    xlParentDataLabelOptionsOverlapping       =	2		# Overlapping parent data label

    #values from excel.xlpastespecialoperation
    #****************************
    xlPasteSpecialOperationAdd                =	2		# Copied data will be added to the value in the destination cell.
    xlPasteSpecialOperationDivide             =	5		# Copied data will divide the value in the destination cell.
    xlPasteSpecialOperationMultiply           =	4		# Copied data will multiply the value in the destination cell.
    xlPasteSpecialOperationNone               =	-4142		# No calculation will be done in the paste operation.
    xlPasteSpecialOperationSubtract           =	3		# Copied data will be subtracted from the value in the destination cell.

    #values from excel.xlpastetype
    #****************************
    xlPasteAll                                =	-4104		# Everything will be pasted.
    xlPasteAllExceptBorders                   =	7		# Everything except borders will be pasted.
    xlPasteAllMergingConditionalFormats       =	14		# Everything will be pasted and conditional formats will be merged.
    xlPasteAllUsingSourceTheme                =	13		# Everything will be pasted using the source theme.
    xlPasteColumnWidths                       =	8		# Copied column width is pasted.
    xlPasteComments                           =	-4144		# Comments are pasted.
    xlPasteFormats                            =	-4122		# Copied source format is pasted.
    xlPasteFormulas                           =	-4123		# Formulas are pasted.
    xlPasteFormulasAndNumberFormats           =	11		# Formulas and Number formats are pasted.
    xlPasteValidation                         =	6		# Validations are pasted.
    xlPasteValues                             =	-4163		# Values are pasted.
    xlPasteValuesAndNumberFormats             =	12		# Values and Number formats are pasted.

    #values from excel.xlpattern
    #****************************
    xlPatternAutomatic                        =	-4105		# Excel controls the pattern.
    xlPatternChecker                          =	9		# Checkerboard.
    xlPatternCrissCross                       =	16		# Criss-cross lines.
    xlPatternDown                             =	-4121		# Dark diagonal lines running from the upper-left to the lower-right.
    xlPatternGray16                           =	17		# 16% gray.
    xlPatternGray25                           =	-4124		# 25% gray.
    xlPatternGray50                           =	-4125		# 50% gray.
    xlPatternGray75                           =	-4126		# 75% gray.
    xlPatternGray8                            =	18		# 8% gray.
    xlPatternGrid                             =	15		# Grid.
    xlPatternHorizontal                       =	-4128		# Dark horizontal lines.
    xlPatternLightDown                        =	13		# Light diagonal lines running from the upper-left to the lower-right.
    xlPatternLightHorizontal                  =	11		# Light horizontal lines.
    xlPatternLightUp                          =	14		# Light diagonal lines running from the lower-left to the upper-right.
    xlPatternLightVertical                    =	12		# Light vertical bars.
    xlPatternNone                             =	-4142		# No pattern.
    xlPatternSemiGray75                       =	10		# 75% dark gray.
    xlPatternSolid                            =	1		# Solid color.
    xlPatternUp                               =	-4162		# Dark diagonal lines running from the lower-left to the upper-right.
    xlPatternVertical                         =	-4166		# Dark vertical bars.

    #values from excel.xlphoneticalignment
    #****************************
    xlPhoneticAlignCenter                     =	2		# Centered
    xlPhoneticAlignDistributed                =	3		# Distributed
    xlPhoneticAlignLeft                       =	1		# Left aligned
    xlPhoneticAlignNoControl                  =	0		# Excel controls alignment

    #values from excel.xlphoneticcharactertype
    #****************************
    xlHiragana                                =	2		# Hiragana
    xlKatakana                                =	1		# Katakana
    xlKatakanaHalf                            =	0		# Half-size Katakana
    xlNoConversion                            =	3		# No conversion

    #values from excel.xlpictureappearance
    #****************************
    xlPrinter                                 =	2		# The picture is copied as it will look when it is printed.
    xlScreen                                  =	1		# The picture is copied to resemble its display on the screen as closely as possible.

    #values from excel.xlpictureconvertortype
    #****************************
    xlBMP                                     =	1		# Windows version 2.0?compatible bitmap
    xlCGM                                     =	7		# Computer Graphics Metafile
    xlDRW                                     =	4		# DRW
    xlDXF                                     =	5		# DXF
    xlEPS                                     =	8		# Encapsulated Postscript
    xlHGL                                     =	6		# HGL
    xlPCT                                     =	13		# Bitmap Graphic (Apple PICT format)
    xlPCX                                     =	10		# PC Paintbrush Bitmap Graphic
    xlPIC                                     =	11		# PIC
    xlPLT                                     =	12		# PLT
    xlTIF                                     =	9		# Tagged Image Format File
    xlWMF                                     =	2		# Windows Metafile
    xlWPG                                     =	3		# WordPerfect/DrawPerfect Graphic

    #values from excel.xlpiesliceindex
    #****************************
    xlCenterPoint                             =	5		# The center point of a pie slice.
    xlInnerCenterPoint                        =	8		# The innermost center point of a doughnut slice.
    xlInnerClockwisePoint                     =	7		# The innermost point of the most clockwise radius of a doughnut slice.
    xlInnerCounterClockwisePoint              =	9		# The innermost point of the most counterclockwise radius of a doughnut slice.
    xlMidClockwiseRadiusPoint                 =	4		# The midpoint of the most clockwise radius of a slice.
    xlMidCounterClockwiseRadiusPoint          =	6		# The midpoint of the most counterclockwise radius of a slice.
    xlOuterCenterPoint                        =	2		# The outer center point of the circumference of a slice.
    xlOuterClockwisePoint                     =	3		# The outermost clockwise point of the circumference of a slice.
    xlOuterCounterClockwisePoint              =	1		# The outermost counterclockwise point of the circumference of a slice.

    #values from excel.xlpieslicelocation
    #****************************
    xlHorizontalCoordinate                    =	1		# The horizontal coordinate (x)
    xlVerticalCoordinate                      =	2		# The vertical coordinate (y)

    #values from excel.xlpivotcelltype
    #****************************
    xlPivotCellBlankCell                      =	9		# A structural blank cell in the PivotTable.
    xlPivotCellCustomSubtotal                 =	7		# A cell in the row or column area that is a custom subtotal.
    xlPivotCellDataField                      =	4		# A data field label (not the  Data button).
    xlPivotCellDataPivotField                 =	8		# The  Data button.
    xlPivotCellGrandTotal                     =	3		# A cell in a row or column area that is a grand total.
    xlPivotCellPageFieldItem                  =	6		# The cell that shows the selected item of a Page field.
    xlPivotCellPivotField                     =	5		# The button for a field (not the  Data button).
    xlPivotCellPivotItem                      =	1		# A cell in the row or column area that is not a subtotal, grand total, custom subtotal, or blank line.
    xlPivotCellSubtotal                       =	2		# A cell in the row or column area that is a subtotal.
    xlPivotCellValue                          =	0		# Any cell in the data area (except a blank row).

    #values from excel.xlpivotconditionscope
    #****************************
    xlDataFieldScope                          =	2		# Based on the data in the specified fields.
    xlFieldsScope                             =	1		# Based on the specified fields.
    xlSelectionScope                          =	0		# Based on the specified selection criteria.

    #values from excel.xlpivotfieldcalculation
    #****************************
    xlDifferenceFrom                          =	2		# The difference from the value of the Base item in the Base field.
    xlIndex                                   =	9		# Data calculated as ((value in cell) x (Grand Total of Grand Totals)) / ((Grand Row Total) x (Grand Column Total)).
    xlNoAdditionalCalculation                 =	-4143		# No calculation.
    xlPercentDifferenceFrom                   =	4		# Percentage difference from the value of the Base item in the Base field.
    xlPercentOf                               =	3		# Percentage of the value of the Base item in the Base field.
    xlPercentOfColumn                         =	7		# Percentage of the total for the column or series.
    xlPercentOfParent                         =	12		# Percentage of the total of the specified parent Base field.
    xlPercentOfParentColumn                   =	11		# Percentage of the total of the parent column.
    xlPercentOfParentRow                      =	10		# Percentage of the total of the parent row.
    xlPercentOfRow                            =	6		# Percentage of the total for the row or category.
    xlPercentOfTotal                          =	8		# Percentage of the grand total of all the data or data points in the report.
    xlPercentRunningTotal                     =	13		# Percentage of the running total of the specified Base field.
    xlRankAscending                           =	14		# Rank smallest to largest.
    xlRankDecending                           =	15		# Rank largest to smallest.
    xlRunningTotal                            =	5		# Data for successive items in the Base field as a running total.

    #values from excel.xlpivotfielddatatype
    #****************************
    xlDate                                    =	2		# Contains a date.
    xlNumber                                  =	-4145		# Contains a number.
    xlText                                    =	-4158		# Contains text.

    #values from excel.xlpivotfieldorientation
    #****************************
    xlColumnField                             =	2		# Column
    xlDataField                               =	4		# Data
    xlHidden                                  =	0		# Hidden
    xlPageField                               =	3		# Page
    xlRowField                                =	1		# Row

    #values from excel.xlpivotfieldrepeatlabels
    #****************************
    xlDoNotRepeatLabels                       =	1		# Do not repeat item labels.
    xlRepeatLabels                            =	2		# Repeat all item labels.

    #values from excel.xlpivotfiltertype
    #****************************
    xlBefore                                  =	31		# Filters for all dates before a specified date
    xlBeforeOrEqualTo                         =	32		# Filters for all dates on or before a specified date
    xlAfter                                   =	33		# Filters for all dates after a specified date
    xlAfterOrEqualTo                          =	34		# Filters for all dates on or after a specified date
    xlAllDatesInPeriodJanuary                 =	57		# Filters for all dates in January
    xlAllDatesInPeriodFebruary                =	58		# Filters for all dates in February
    xlAllDatesInPeriodMarch                   =	59		# Filters for all dates in March
    xlAllDatesInPeriodApril                   =	60		# Filters for all dates in April
    xlAllDatesInPeriodMay                     =	61		# Filters for all dates in May
    xlAllDatesInPeriodJune                    =	62		# Filters for all dates in June
    xlAllDatesInPeriodJuly                    =	63		# Filters for all dates in July
    xlAllDatesInPeriodAugust                  =	64		# Filters for all dates in August
    xlAllDatesInPeriodSeptember               =	65		# Filters for all dates in September
    xlAllDatesInPeriodOctober                 =	66		# Filters for all dates in October
    xlAllDatesInPeriodNovember                =	67		# Filters for all dates in November
    xlAllDatesInPeriodDecember                =	68		# Filters for all dates in December
    xlAllDatesInPeriodQuarter1                =	53		# Filters for all dates in Quarter1
    xlAllDatesInPeriodQuarter2                =	54		# Filters for all dates in Quarter2
    xlAllDatesInPeriodQuarter3                =	55		# Filters for all dates in Quarter3
    xlAllDatesInPeriodQuarter4                =	56		# Filters for all dates in Quarter 4
    xlBottomCount                             =	2		# Filters for the specified number of values from the bottom of a list
    xlBottomPercent                           =	4		# Filters for the specified percentage of values from the bottom of a list
    xlBottomSum                               =	6		# Sum of the values from the bottom of the list
    xlCaptionBeginsWith                       =	17		# Filters for all captions beginning with the specified string
    xlCaptionContains                         =	21		# Filters for all captions that contain the specified string
    xlCaptionDoesNotBeginWith                 =	18		# Filters for all captions that do not begin with the specified string
    xlCaptionDoesNotContain                   =	22		# Filters for all captions that do not contain the specified string
    xlCaptionDoesNotEndWith                   =	20		# Filters for all captions that do not end with the specified string
    xlCaptionDoesNotEqual                     =	16		# Filters for all captions that do not match the specified string
    xlCaptionEndsWith                         =	19		# Filters for all captions that end with the specified string
    xlCaptionEquals                           =	15		# Filters for all captions that match the specified string
    xlCaptionIsBetween                        =	27		# Filters for all captions that are between a specified range of values
    xlCaptionIsGreaterThan                    =	23		# Filters for all captions that are greater than the specified value
    xlCaptionIsGreaterThanOrEqualTo           =	24		# Filters for all captions that are greater than or match the specified value
    xlCaptionIsLessThan                       =	25		# Filters for all captions that are less than the specified value
    xlCaptionIsLessThanOrEqualTo              =	26		# Filters for all captions that are less than or match the specified value
    xlCaptionIsNotBetween                     =	28		# Filters for all captions that are not between a specified range of values
    xlDateBetween                             =	35		# Filters for all dates that are between a specified range of dates
    xlDateLastMonth                           =	45		# Filters for all dates that apply to the previous month
    xlDateLastQuarter                         =	48		# Filters for all dates that apply to the previous quarter
    xlDateLastWeek                            =	42		# Filters for all dates that apply to the previous week
    xlDateLastYear                            =	51		# Filters for all dates that apply to the previous year
    xlDateNextMonth                           =	43		# Filters for all dates that apply to the next month
    xlDateNextQuarter                         =	46		# Filters for all dates that apply to the next quarter
    xlDateNextWeek                            =	40		# Filters for all dates that apply to the next week
    xlDateNextYear                            =	49		# Filters for all dates that apply to the next year
    xlDateThisMonth                           =	44		# Filters for all dates that apply to the current month
    xlDateThisQuarter                         =	47		# Filters for all dates that apply to the current quarter
    xlDateThisWeek                            =	41		# Filters for all dates that apply to the current week
    xlDateThisYear                            =	50		# Filters for all dates that apply to the current year
    xlDateToday                               =	38		# Filters for all dates that apply to the current date
    xlDateTomorrow                            =	37		# Filters for all dates that apply to the next day
    xlDateYesterday                           =	39		# Filters for all dates that apply to the previous day
    xlNotSpecificDate                         =	30		# Filters for all dates that do not match a specified date
    xlSpecificDate                            =	29		# Filters for all dates that match a specified date
    xlTopCount                                =	1		# Filters for the specified number of values from the top of a list
    xlTopPercent                              =	3		# Filters for the specified percentage of values from a list
    xlTopSum                                  =	5		# Sum of the values from the top of the list
    xlValueDoesNotEqual                       =	8		# Filters for all values that do not match the specified value
    xlValueEquals                             =	7		# Filters for all values that match the specified value
    xlValueIsBetween                          =	13		# Filters for all values that are between a specified range of values
    xlValueIsGreaterThan                      =	9		# Filters for all values that are greater than the specified value
    xlValueIsGreaterThanOrEqualTo             =	10		# Filters for all values that are greater than or match the specified value
    xlValueIsLessThan                         =	11		# Filters for all values that are less than the specified value
    xlValueIsLessThanOrEqualTo                =	12		# Filters for all values that are less than or match the specified value
    xlValueIsNotBetween                       =	14		# Filters for all values that are not between a specified range of values
    xlYearToDate                              =	52		# Filters for all values that are within one year of a specified date

    #values from excel.xlpivotformattype
    #****************************
    xlPTClassic                               =	20		# PivotTable classic format.
    xlPTNone                                  =	21		# Does not apply formatting to the PivotTable report.
    xlReport1                                 =	0		# Use the xlReport1 formatting for the PivotTable.
    xlReport10                                =	9		# Use the xlReport10 formatting for the PivotTable.
    xlReport2                                 =	1		# Use the xlReport2 formatting for the PivotTable.
    xlReport3                                 =	2		# Use the xlReport3 formatting for the PivotTable.
    xlReport4                                 =	3		# Use the xlReport4 formatting for the PivotTable.
    xlReport5                                 =	4		# Use the xlReport5 formatting for the PivotTable.
    xlReport6                                 =	5		# Use the xlReport6 formatting for the PivotTable.
    xlReport7                                 =	6		# Use the xlReport7 formatting for the PivotTable.
    xlReport8                                 =	7		# Use the xlReport8 formatting for the PivotTable.
    xlReport9                                 =	8		# Use the xlReport9 formatting for the PivotTable.
    xlTable1                                  =	10		# Use the xlTable1 formatting for the PivotTable.
    xlTable10                                 =	19		# Use the xlTable10 formatting for the PivotTable.
    xlTable2                                  =	11		# Use the xlTable2 formatting for the PivotTable.
    xlTable3                                  =	12		# Use the xlTable3 formatting for the PivotTable.
    xlTable4                                  =	13		# Use the xlTable4 formatting for the PivotTable.
    xlTable5                                  =	14		# Use the xlTable5 formatting for the PivotTable.
    xlTable6                                  =	15		# Use the xlTable6 formatting for the PivotTable.
    xlTable7                                  =	16		# Use the xlTable7 formatting for the PivotTable.
    xlTable8                                  =	17		# Use the xlTable8 formatting for the PivotTable.
    xlTable9                                  =	18		# Use the xlTable9 formatting for the PivotTable.

    #values from excel.xlpivotlinetype
    #****************************
    xlPivotLineBlank                          =	3		# Blank line after each group.
    xlPivotLineGrandTotal                     =	2		# Grand Total line.
    xlPivotLineRegular                        =	0		# Regular PivotLine with pivot items.
    xlPivotLineSubtotal                       =	1		# Subtotal line.

    #values from excel.xlpivottablemissingitems
    #****************************
    xlMissingItemsDefault                     =	-1		# The default number of unique items per PivotField allowed.
    xlMissingItemsMax                         =	32500		# The maximum number of unique items per PivotField allowed (32,500) for a pre-Excel 2007 PivotTable.
    xlMissingItemsMax2                        =	1048576		# The maximum number of unique items per PivotField allowed (1,048,576) for PivotTables in Excel 2007 and later.
    xlMissingItemsNone                        =	0		# No unique items per PivotField allowed (zero).

    #values from excel.xlpivottablesourcetype
    #****************************
    xlConsolidation                           =	3		# Multiple consolidation ranges.
    xlDatabase                                =	1		# Microsoft Excel list or database.
    xlExternal                                =	2		# Data from another application.
    xlPivotTable                              =	-4148		# Same source as another PivotTable report.
    xlScenario                                =	4		# Data is based on scenarios created using the Scenario Manager.

    #values from excel.xlpivottableversionlist
    #****************************
    xlPivotTableVersion2000                   =	0		# Excel 2000
    xlPivotTableVersion10                     =	1		# Excel 2002
    xlPivotTableVersion11                     =	2		# Excel 2003
    xlPivotTableVersion12                     =	3		# Excel 2007
    xlPivotTableVersion14                     =	4		# Excel 2010
    xlPivotTableVersion15                     =	5		# Excel 2013
    xlPivotTableVersionCurrent                =	-1		# Provided only for backward compatibility

    #values from excel.xlplacement
    #****************************
    xlFreeFloating                            =	3		# Object is free floating.
    xlMove                                    =	2		# Object is moved with the cells.
    xlMoveAndSize                             =	1		# Object is moved and sized with the cells.

    #values from excel.xlplatform
    #****************************
    xlMacintosh                               =	1		# Macintosh
    xlMSDOS                                   =	3		# MS-DOS
    xlWindows                                 =	2		# Microsoft Windows

    #values from excel.xlportuguesereform
    #****************************
    xlPortugueseBoth                          =	3		# The spelling checker recognizes both pre-reform and post-reform spellings.
    xlPortuguesePostReform                    =	2		# The spelling checker recognizes only post-reform spellings.
    xlPortuguesePreReform                     =	1		# The spelling checker recognizes only pre-reform spellings.

    #values from excel.xlprinterrors
    #****************************
    xlPrintErrorsBlank                        =	1		# Print errors are blank.
    xlPrintErrorsDash                         =	2		# Print errors are displayed as dashes.
    xlPrintErrorsDisplayed                    =	0		# All print errors are displayed.
    xlPrintErrorsNA                           =	3		# Print errors are displayed as not available.

    #values from excel.xlprintlocation
    #****************************
    xlPrintInPlace                            =	16		# Comments will be printed where they were inserted in the worksheet.
    xlPrintNoComments                         =	-4142		# Comments will not be printed.
    xlPrintSheetEnd                           =	1		# Comments will be printed as end notes at the end of the worksheet.

    #values from excel.xlpriority
    #****************************
    xlPriorityHigh                            =	-4127		# High
    xlPriorityLow                             =	-4134		# Low
    xlPriorityNormal                          =	-4143		# Normal

    #values from excel.xlpropertydisplayedin
    #****************************
    xlDisplayPropertyInPivotTable             =	1		# Displays member property in the PivotTable only. This is the default value.
    xlDisplayPropertyInPivotTableAndTooltip   =	3		# Displays member property in the tooltip only.
    xlDisplayPropertyInTooltip                =	2		# Displays member property in both the tooltip and the PivotTable.

    #values from excel.xlprotectedviewclosereason
    #****************************
    xlProtectedViewCloseEdit                  =	1		# The window was closed when the user clicked the  Enable Editing button.
    xlProtectedViewCloseForced                =	2		# The window was closed because the application shut it down forcefully or stopped responding.
    xlProtectedViewCloseNormal                =	0		# The window was closed normally.

    #values from excel.xlprotectedviewwindowstate
    #****************************
    xlProtectedViewWindowMaximized            =	2		# Maximized
    xlProtectedViewWindowMinimized            =	1		# Minimized
    xlProtectedViewWindowNormal               =	0		# Normal

    #values from excel.xlquerytype
    #****************************
    xlADORecordset                            =	7		# Based on an ADO recordset query
    xlDAORecordset                            =	2		# Based on a DAO recordset query, for query tables only
    xlODBCQuery                               =	1		# Based on an ODBC data source
    xlOLEDBQuery                              =	5		# Based on an OLE DB query, including OLAP data sources
    xlTextImport                              =	6		# Based on a text file, for query tables only
    xlWebQuery                                =	4		# Based on a web page, for query tables only

    #values from excel.xlquickanalysismode
    #****************************
    xlLensOnly                                =	0		# Show the button but no callout user interface
    xlFormatConditions                        =	1		# Conditional Formatting
    xlRecommendedCharts                       =	2		# Charts
    xlTotals                                  =	3		# Totals
    xlTables                                  =	4		# Tables
    xlSparklines                              =	5		# Sparklines

    #values from excel.xlrangeautoformat
    #****************************
    xlRangeAutoFormat3DEffects1               =	13		# 3D effects 1.
    xlRangeAutoFormat3DEffects2               =	14		# 3D effects 2.
    xlRangeAutoFormatAccounting1              =	4		# Accounting 1.
    xlRangeAutoFormatAccounting2              =	5		# Accounting 2.
    xlRangeAutoFormatAccounting3              =	6		# Accounting 3.
    xlRangeAutoFormatAccounting4              =	17		# Accounting 4.
    xlRangeAutoFormatClassic1                 =	1		# Classic 1.
    xlRangeAutoFormatClassic2                 =	2		# Classic 2.
    xlRangeAutoFormatClassic3                 =	3		# Classic 3.
    xlRangeAutoFormatClassicPivotTable        =	31		# Classic pivot table.
    xlRangeAutoFormatColor1                   =	7		# Color 1.
    xlRangeAutoFormatColor2                   =	8		# Color 2.
    xlRangeAutoFormatColor3                   =	9		# Color 3.
    xlRangeAutoFormatList1                    =	10		# List 1.
    xlRangeAutoFormatList2                    =	11		# List 2.
    xlRangeAutoFormatList3                    =	12		# List 3.
    xlRangeAutoFormatLocalFormat1             =	15		# Local Format 1.
    xlRangeAutoFormatLocalFormat2             =	16		# Local Format 2.
    xlRangeAutoFormatLocalFormat3             =	19		# Local Format 3.
    xlRangeAutoFormatLocalFormat4             =	20		# Local Format 4.
    xlRangeAutoFormatNone                     =	-4142		# No specified format.
    xlRangeAutoFormatPTNone                   =	42		# No specified pivot table format.
    xlRangeAutoFormatReport1                  =	21		# Report 1.
    xlRangeAutoFormatReport10                 =	30		# Report 10.
    xlRangeAutoFormatReport2                  =	22		# Report 2.
    xlRangeAutoFormatReport3                  =	23		# Report 3.
    xlRangeAutoFormatReport4                  =	24		# Report 4.
    xlRangeAutoFormatReport5                  =	25		# Report 5.
    xlRangeAutoFormatReport6                  =	26		# Report 6.
    xlRangeAutoFormatReport7                  =	27		# Report 7.
    xlRangeAutoFormatReport8                  =	28		# Report 8.
    xlRangeAutoFormatReport9                  =	29		# Report 9.
    xlRangeAutoFormatSimple                   =	-4154		# Simple.
    xlRangeAutoFormatTable1                   =	32		# Table 1.
    xlRangeAutoFormatTable10                  =	41		# Table 10.
    xlRangeAutoFormatTable2                   =	33		# Table 2.
    xlRangeAutoFormatTable3                   =	34		# Table 3.
    xlRangeAutoFormatTable4                   =	35		# Table 4.
    xlRangeAutoFormatTable5                   =	36		# Table 5.
    xlRangeAutoFormatTable6                   =	37		# Table 6.
    xlRangeAutoFormatTable7                   =	38		# Table 7.
    xlRangeAutoFormatTable8                   =	39		# Table 8.
    xlRangeAutoFormatTable9                   =	40		# Table 9.

    #values from excel.xlrangevaluedatatype
    #****************************
    xlRangeValueDefault                       =	10		# Default. If the specified  Range object is empty, returns the value Empty (use the IsEmpty function to test for this case). If the Range object contains more than one cell, returns an array of values (use the IsArray function to test for this case).
    xlRangeValueMSPersistXML                  =	12		# Returns the recordset representation of the specified  Range object in an XML format.
    xlRangeValueXMLSpreadsheet                =	11		# Returns the values, formatting, formulas, and names of the specified  Range object in the XML Spreadsheet format.

    #values from excel.xlreferencestyle
    #****************************
    xlA1                                      =	1		# Default. Use  xlA1 to return an A1-style reference.
    xlR1C1                                    =	-4150		# Use  xlR1C1 to return an R1C1-style reference.

    #values from excel.xlreferencetype
    #****************************
    xlAbsolute                                =	1		# Convert to absolute row and column style.
    xlAbsRowRelColumn                         =	2		# Convert to absolute row and relative column style.
    xlRelative                                =	4		# Convert to relative row and column style.
    xlRelRowAbsColumn                         =	3		# Convert to relative row and absolute column style.

    #values from excel.xlremovedocinfotype
    #****************************
    xlRDIAll                                  =	99		# Removes all documentation information.
    xlRDIComments                             =	1		# Removes comments from the document information.
    xlRDIContentType                          =	16		# Removes content type data from the document information.
    xlRDIDefinedNameComments                  =	18		# Removes defined name comments from the documentation information.
    xlRDIDocumentManagementPolicy             =	15		# Removes document management policy data from the document information.
    xlRDIDocumentProperties                   =	8		# Removes document properties from the document information.
    xlRDIDocumentServerProperties             =	14		# Removes server properties from the document information.
    xlRDIDocumentWorkspace                    =	10		# Removes workspace data from the document information.
    xlRDIEmailHeader                          =	5		# Removes email headers from the document information.
    xlRDIExcelDataModel                       =	23		# Removes Data Model data from the document information.
    xlRDIInactiveDataConnections              =	19		# Removes inactive data connection data from the document information.
    xlRDIInkAnnotations                       =	11		# Removes ink annotations from the document information.
    xlRDIInlineWebExtensions                  =	21		# Removes inline Web Extensions from the document information.
    xlRDIPrinterPath                          =	20		# Removes printer paths from the document information.
    xlRDIPublishInfo                          =	13		# Removes the publish information data from the document information.
    xlRDIRemovePersonalInformation            =	4		# Removes personal information from the document information.
    xlRDIRoutingSlip                          =	6		# Removes routing slip information from the document information.
    xlRDIScenarioComments                     =	12		# Removes scenario comments from the document information.
    xlRDISendForReview                        =	7		# Removes the send for review information from the document information.
    xlRDITaskpaneWebExtensions                =	22		# Removes task pane Web Extensions from the document information.

    #values from excel.xlrgbcolor
    #****************************
    rgbAliceBlue                              =	16775408		# Alice Blue
    rgbAntiqueWhite                           =	14150650		# Antique White
    rgbAqua                                   =	16776960		# Aqua
    rgbAquamarine                             =	13959039		# Aquamarine
    rgbAzure                                  =	16777200		# Azure
    rgbBeige                                  =	14480885		# Beige
    rgbBisque                                 =	12903679		# Bisque
    rgbBlack                                  =	0		# Black
    rgbBlanchedAlmond                         =	13495295		# Blanched Almond
    rgbBlue                                   =	16711680		# Blue
    rgbBlueViolet                             =	14822282		# Blue Violet
    rgbBrown                                  =	2763429		# Brown
    rgbBurlyWood                              =	8894686		# Burly Wood
    rgbCadetBlue                              =	10526303		# Cadet Blue
    rgbChartreuse                             =	65407		# Chartreuse
    rgbCoral                                  =	5275647		# Coral
    rgbCornflowerBlue                         =	15570276		# Cornflower Blue
    rgbCornsilk                               =	14481663		# Cornsilk
    rgbCrimson                                =	3937500		# Crimson
    rgbDarkBlue                               =	9109504		# Dark Blue
    rgbDarkCyan                               =	9145088		# Dark Cyan
    rgbDarkGoldenrod                          =	755384		# Dark Goldenrod
    rgbDarkGray                               =	11119017		# Dark Gray
    rgbDarkGreen                              =	25600		# Dark Green
    rgbDarkGrey                               =	11119017		# Dark Grey
    rgbDarkKhaki                              =	7059389		# Dark Khaki
    rgbDarkMagenta                            =	9109643		# Dark Magenta
    rgbDarkOliveGreen                         =	3107669		# Dark Olive Green
    rgbDarkOrange                             =	36095		# Dark Orange
    rgbDarkOrchid                             =	13382297		# Dark Orchid
    rgbDarkRed                                =	139		# Dark Red
    rgbDarkSalmon                             =	8034025		# Dark Salmon
    rgbDarkSeaGreen                           =	9419919		# Dark Sea Green
    rgbDarkSlateBlue                          =	9125192		# Dark Slate Blue
    rgbDarkSlateGray                          =	5197615		# Dark Slate Gray
    rgbDarkSlateGrey                          =	5197615		# Dark Slate Grey
    rgbDarkTurquoise                          =	13749760		# Dark Turquoise
    rgbDarkViolet                             =	13828244		# Dark Violet
    rgbDeepPink                               =	9639167		# Deep Pink
    rgbDeepSkyBlue                            =	16760576		# Deep Sky Blue
    rgbDimGray                                =	6908265		# Dim Gray
    rgbDimGrey                                =	6908265		# Dim Grey
    rgbDodgerBlue                             =	16748574		# Dodger Blue
    rgbFireBrick                              =	2237106		# Fire Brick
    rgbFloralWhite                            =	15792895		# Floral White
    rgbForestGreen                            =	2263842		# Forest Green
    rgbFuchsia                                =	16711935		# Fuchsia
    rgbGainsboro                              =	14474460		# Gainsboro
    rgbGhostWhite                             =	16775416		# Ghost White
    rgbGold                                   =	55295		# Gold
    rgbGoldenrod                              =	2139610		# Goldenrod
    rgbGray                                   =	8421504		# Gray
    rgbGreen                                  =	32768		# Green
    rgbGreenYellow                            =	3145645		# Green Yellow
    rgbGrey                                   =	8421504		# Grey
    rgbHoneydew                               =	15794160		# Honeydew
    rgbHotPink                                =	11823615		# Hot Pink
    rgbIndianRed                              =	6053069		# Indian Red
    rgbIndigo                                 =	8519755		# Indigo
    rgbIvory                                  =	15794175		# Ivory
    rgbKhaki                                  =	9234160		# Khaki
    rgbLavender                               =	16443110		# Lavender
    rgbLavenderBlush                          =	16118015		# Lavender Blush
    rgbLawnGreen                              =	64636		# Lawn Green
    rgbLemonChiffon                           =	13499135		# Lemon Chiffon
    rgbLightBlue                              =	15128749		# Light Blue
    rgbLightCoral                             =	8421616		# Light Coral
    rgbLightCyan                              =	9145088		# Light Cyan
    rgbLightGoldenrodYellow                   =	13826810		# LightGoldenrodYellow
    rgbLightGray                              =	13882323		# Light Gray
    rgbLightGreen                             =	9498256		# Light Green
    rgbLightGrey                              =	13882323		# Light Grey
    rgbLightPink                              =	12695295		# Light Pink
    rgbLightSalmon                            =	8036607		# Light Salmon
    rgbLightSeaGreen                          =	11186720		# Light Sea Green
    rgbLightSkyBlue                           =	16436871		# Light Sky Blue
    rgbLightSlateGray                         =	10061943		# Light Slate Gray
    rgbLightSteelBlue                         =	14599344		# Light Steel Blue
    rgbLightYellow                            =	14745599		# Light Yellow
    rgbLime                                   =	65280		# Lime
    rgbLimeGreen                              =	3329330		# Lime Green
    rgbLinen                                  =	15134970		# Linen
    rgbMaroon                                 =	128		# Maroon
    rgbMediumAquamarine                       =	11206502		# Medium Aquamarine
    rgbMediumBlue                             =	13434880		# Medium Blue
    rgbMediumOrchid                           =	13850042		# Medium Orchid
    rgbMediumPurple                           =	14381203		# Medium Purple
    rgbMediumSeaGreen                         =	7451452		# Medium Sea Green
    rgbMediumSlateBlue                        =	15624315		# Medium Slate Blue
    rgbMediumSpringGreen                      =	10156544		# Medium Spring Green
    rgbMediumTurquoise                        =	13422920		# Medium Turquoise
    rgbMediumVioletRed                        =	8721863		# Medium Violet Red
    rgbMidnightBlue                           =	7346457		# Midnight Blue
    rgbMintCream                              =	16449525		# Mint Cream
    rgbMistyRose                              =	14804223		# Misty Rose
    rgbMoccasin                               =	11920639		# Moccasin
    rgbNavajoWhite                            =	11394815		# Navajo White
    rgbNavy                                   =	8388608		# Navy
    rgbNavyBlue                               =	8388608		# Navy Blue
    rgbOldLace                                =	15136253		# Old Lace
    rgbOlive                                  =	32896		# Olive
    rgbOliveDrab                              =	2330219		# Olive Drab
    rgbOrange                                 =	42495		# Orange
    rgbOrangeRed                              =	17919		# Orange Red
    rgbOrchid                                 =	14053594		# Orchid
    rgbPaleGoldenrod                          =	7071982		# Pale Goldenrod
    rgbPaleGreen                              =	10025880		# Pale Green
    rgbPaleTurquoise                          =	15658671		# Pale Turquoise
    rgbPaleVioletRed                          =	9662683		# Pale Violet Red
    rgbPapayaWhip                             =	14020607		# Papaya Whip
    rgbPeachPuff                              =	12180223		# Peach Puff
    rgbPeru                                   =	4163021		# Peru
    rgbPink                                   =	13353215		# Pink
    rgbPlum                                   =	14524637		# Plum
    rgbPowderBlue                             =	15130800		# Powder Blue
    rgbPurple                                 =	8388736		# Purple
    rgbRed                                    =	255		# Red
    rgbRosyBrown                              =	9408444		# Rosy Brown
    rgbRoyalBlue                              =	14772545		# Royal Blue
    rgbSalmon                                 =	7504122		# Salmon
    rgbSandyBrown                             =	6333684		# Sandy Brown
    rgbSeaGreen                               =	5737262		# Sea Green
    rgbSeashell                               =	15660543		# Seashell
    rgbSienna                                 =	2970272		# Sienna
    rgbSilver                                 =	12632256		# Silver
    rgbSkyBlue                                =	15453831		# Sky Blue
    rgbSlateBlue                              =	13458026		# Slate Blue
    rgbSlateGray                              =	9470064		# Slate Gray
    rgbSnow                                   =	16448255		# Snow
    rgbSpringGreen                            =	8388352		# Spring Green
    rgbSteelBlue                              =	11829830		# Steel Blue
    rgbTan                                    =	9221330		# Tan
    rgbTeal                                   =	8421376		# Teal
    rgbThistle                                =	14204888		# Thistle
    rgbTomato                                 =	4678655		# Tomato
    rgbTurquoise                              =	13688896		# Turquoise
    rgbViolet                                 =	15631086		# Violet
    rgbWheat                                  =	11788021		# Wheat
    rgbWhite                                  =	16777215		# White
    rgbWhiteSmoke                             =	16119285		# White Smoke
    rgbYellow                                 =	65535		# Yellow
    rgbYellowGreen                            =	3329434		# Yellow Green

    #values from excel.xlrobustconnect
    #****************************
    xlAlways                                  =	1		# The PivotTable cache or query table always uses external source information (as defined by the  SourceConnectionFile or SourceDataFile property) to reconnect.
    xlAsRequired                              =	0		# The PivotTable cache or query table uses external source information to reconnect, using the  Connection property.
    xlNever                                   =	2		# The PivotTable cache or query table never uses source information to reconnect.

    #values from excel.xlroutingslipdelivery
    #****************************


    #values from excel.xlroutingslipstatus
    #****************************


    #values from excel.xlrowcol
    #****************************
    xlColumns                                 =	2		# Data series is in a row.
    xlRows                                    =	1		# Data series is in a column.

    #values from excel.xlrunautomacro
    #****************************
    xlAutoActivate                            =	3		# Auto_Activate macros
    xlAutoClose                               =	2		# Auto_Close macros
    xlAutoDeactivate                          =	4		# Auto_Deactivate macros
    xlAutoOpen                                =	1		# Auto_Open macros

    #values from excel.xlsaveaction
    #****************************
    xlDoNotSaveChanges                        =	2		# Changes will not be saved.
    xlSaveChanges                             =	1		# Changes will be saved.

    #values from excel.xlsaveasaccessmode
    #****************************
    xlExclusive                               =	3		# Exclusive mode
    xlNoChange                                =	1		# Default (does not change the access mode)
    xlShared                                  =	2		# Share list

    #values from excel.xlsaveconflictresolution
    #****************************
    xlLocalSessionChanges                     =	2		# The local user's changes are always accepted.
    xlOtherSessionChanges                     =	3		# The local user's changes are always rejected.
    xlUserResolution                          =	1		# A dialog box asks the user to resolve the conflict.

    #values from excel.xlscaletype
    #****************************
    xlScaleLinear                             =	-4132		# Linear
    xlScaleLogarithmic                        =	-4133		# Logarithmic

    #values from excel.xlsearchdirection
    #****************************
    xlNext                                    =	1		# Search for next matching value in range.
    xlPrevious                                =	2		# Search for previous matching value in range.

    #values from excel.xlsearchorder
    #****************************
    xlByColumns                               =	2		# Searches down through a column, then moves to the next column.
    xlByRows                                  =	1		# Searches across a row, then moves to the next row.

    #values from excel.xlsearchwithin
    #****************************
    xlWithinSheet                             =	1		# Limit search to current sheet.
    xlWithinWorkbook                          =	2		# Search whole workbook.

    #values from excel.xlseriesnamelevel
    #****************************
    xlSeriesNameLevelAll                      =	-1		# Set series names to all series name levels w/in range on the chart.
    xlSeriesNameLevelCustom                   =	-2		# Indicates literal data in the series names.
    xlSeriesNameLevelNone                     =	-3		# Set no category labels in the chart. Defaults to automatic indexed labels.

    #values from excel.xlsheettype
    #****************************
    xlChart                                   =	-4109		# Chart
    xlDialogSheet                             =	-4116		# Dialog sheet
    xlExcel4IntlMacroSheet                    =	4		# Excel version 4 international macro sheet
    xlExcel4MacroSheet                        =	3		# Excel version 4 macro sheet
    xlWorksheet                               =	-4167		# Worksheet

    #values from excel.xlsheetvisibility
    #****************************
    xlSheetHidden                             =	0		# Hides the worksheet which the user can unhide via menu.
    xlSheetVeryHidden                         =	2		# Hides the object so that the only way for you to make it visible again is by setting this property to True (the user cannot make the object visible).
    xlSheetVisible                            =	-1		# Displays the sheet.

    #values from excel.xlsizerepresents
    #****************************
    xlSizeIsArea                              =	1		# Area of the bubble.
    xlSizeIsWidth                             =	2		# Width of the bubble.

    #values from excel.xlslicercachetype
    #****************************
    xlSlicer                                  =	1		# Slicer cache represents a Slicer.
    xlTimeline                                =	2		# Slicer cache represents a Timeline.

    #values from excel.xlslicercrossfiltertype
    #****************************
    xlSlicerCrossFilterHideButtonsWithNoData  =	4		# Cross filtering is turned on for this slicer cache, any tile with no data for a filtering selection in other slicers connected to the same data source will be dimmed. Additionally, buttons will be hidden.
    xlSlicerCrossFilterShowItemsWithDataAtTop	=	2		# Cross filtering is turned on for this slicer cache, any tile with no data for a filtering selection in other slicers connected to the same data source will be dimmed. Additionally, tiles with data are moved to the top in the slicer. (Default)
    xlSlicerCrossFilterShowItemsWithNoData    =	3		# Cross filtering is turned on for this slicer cache, any tile with no data for a filtering selection in other slicers connected to the same data source will be dimmed.
    xlSlicerNoCrossFilter                     =	1		# Cross filtering is turned off entirely, so all tiles are displayed and active (not dimmed) regardless of filtering selections in other slicers.

    #values from excel.xlslicersort
    #****************************
    xlSlicerSortAscending                     =	2		# Slicer items are sorted in ascending order by item captions.
    xlSlicerSortDataSourceOrder               =	1		# Slicer items are displayed in the order provided by the data source.
    xlSlicerSortDescending                    =	3		# Slicer items are sorted in descending order by item captions.

    #values from excel.xlsmarttagcontroltype
    #****************************


    #values from excel.xlsmarttagdisplaymode
    #****************************


    #values from excel.xlsortdataoption
    #****************************
    xlSortNormal                              =	0		# default. Sorts numeric and text data separately.
    xlSortTextAsNumbers                       =	1		# Treat text as numeric data for the sort.

    #values from excel.xlsortmethod
    #****************************
    xlPinYin                                  =	1		# Phonetic Chinese sort order for characters. This is the default value.
    xlStroke                                  =	2		# Sort by the quantity of strokes in each character.

    #values from excel.xlsortmethodold
    #****************************
    xlCodePage                                =	2		# Sort by code page.
    xlSyllabary                               =	1		# Sort phonetically.

    #values from excel.xlsorton
    #****************************
    SortOnCellColor                           =	1		# Cell color.
    SortOnFontColor                           =	2		# Font color.
    SortOnIcon                                =	3		# Icon.
    SortOnValues                              =	0		# Values.

    #values from excel.xlsortorder
    #****************************
    xlAscending                               =	1		# Sorts the specified field in ascending order. This is the default value.
    xlDescending                              =	2		# Sorts the specified field in descending order.
    xlManual                                  =	-4135		# Manual sort (you can drag items to rearrange them).

    #values from excel.xlsortorientation
    #****************************
    xlSortColumns                             =	1		# Sorts by column.
    xlSortRows                                =	2		# Sorts by row. This is the default value.

    #values from excel.xlsorttype
    #****************************
    xlSortLabels                              =	2		# Sorts the PivotTable report by labels.
    xlSortValues                              =	1		# Sorts the PivotTable report by values.

    #values from excel.xlsourcetype
    #****************************
    xlSourceAutoFilter                        =	3		# An AutoFilter range
    xlSourceChart                             =	5		# A chart
    xlSourcePivotTable                        =	6		# A PivotTable report
    xlSourcePrintArea                         =	2		# A range of cells selected for printing
    xlSourceQuery                             =	7		# A query table (external data range)
    xlSourceRange                             =	4		# A range of cells
    xlSourceSheet                             =	1		# An entire worksheet
    xlSourceWorkbook                          =	0		# A workbook

    #values from excel.xlspanishmodes
    #****************************
    xlSpanishTuteoAndVoseo                    =	1		# Tuteo and Voseo verb forms.
    xlSpanishTuteoOnly                        =	0		# Tuteo verb forms only.
    xlSpanishVoseoOnly                        =	2		# Voseo verb forms only.

    #values from excel.xlsparkscale
    #****************************
    xlSparkScaleCustom                        =	3		# The minimum or maximum value for the vertical axis of the sparkline has a user-defined value.
    xlSparkScaleGroup                         =	1		# The minimum or maximum value for the vertical axes of all of the sparklines in the group have the same value.
    xlSparkScaleSingle                        =	2		# The minimum or maximum value for the vertical axis of each sparkline in the group is automatically set to its own calculated value.

    #values from excel.xlsparktype
    #****************************
    xlSparkColumn                             =	2		# A column chart sparkline.
    xlSparkColumnStacked100                   =	3		# A win/loss chart sparkline.
    xlSparkLine                               =	1		# A line chart sparkline.

    #values from excel.xlsparklinerowcol
    #****************************
    SparklineColumnsSquare                    =	2		# Plot the data by columns.
    SparklineNonSquare                        =	0		# The sparkline is not bound to data in a square-shaped range.
    SparklineRowsSquare                       =	1		# Plot the data by rows.

    #values from excel.xlspeakdirection
    #****************************
    xlSpeakByColumns                          =	1		# Reads down a column, then moves to the next column.
    xlSpeakByRows                             =	0		# Reads across a row, then moves to the next row.

    #values from excel.xlspecialcellsvalue
    #****************************
    xlErrors                                  =	16		# Cells with errors.
    xlLogical                                 =	4		# Cells with logical values.
    xlNumbers                                 =	1		# Cells with numeric values.
    xlTextValues                              =	2		# Cells with text.

    #values from excel.xlstdcolorscale
    #****************************
    ColorScaleBlackWhite                      =	3		# Black over White.
    ColorScaleGYR                             =	2		# GYR.
    ColorScaleRYG                             =	1		# RYG.
    ColorScaleWhiteBlack                      =	4		# White over Black.

    #values from excel.xlsubscribetoformat
    #****************************
    xlSubscribeToPicture                      =	-4147		# Picture
    xlSubscribeToText                         =	-4158		# Text

    #values from excel.xlsubtototallocationtype
    #****************************
    xlAtBottom                                =	2		# Subtotal will be at the bottom.
    xlAtTop                                   =	1		# Subtotal will be at the top.

    #values from excel.xlsummarycolumn
    #****************************
    xlSummaryOnLeft                           =	-4131		# The summary column will be positioned to the left of the detail columns in the outline.
    xlSummaryOnRight                          =	-4152		# The summary column will be positioned to the right of the detail columns in the outline.

    #values from excel.xlsummaryreporttype
    #****************************
    xlStandardSummary                         =	1		# List scenarios side by side.
    xlSummaryPivotTable                       =	-4148		# Display scenarios in a PivotTable report.

    #values from excel.xlsummaryrow
    #****************************
    xlSummaryAbove                            =	0		# The summary row will be positioned above the detail rows in the outline.
    xlSummaryBelow                            =	1		# The summary row will be positioned below the detail rows in the outline.

    #values from excel.xltabposition
    #****************************
    xlTabPositionFirst                        =	0		# First tab position.
    xlTabPositionLast                         =	1		# Last tab position.

    #values from excel.xltablestyleelementtype
    #****************************
    xlBlankRow                                =	19		# Blank row
    xlColumnStripe1                           =	7		# Column Stripe1
    xlColumnStripe2                           =	8		# Column Stripe2
    xlColumnSubheading1                       =	20		# Column Subheading1
    xlColumnSubheading2                       =	21		# Column Subheading2
    xlColumnSubheading3                       =	22		# Column Subheading3
    xlFirstColumn                             =	3		# First column
    xlFirstHeaderCell                         =	9		# First header cell
    xlFirstTotalCell                          =	11		# First total cell
    xlGrandTotalColumn                        =	4		# Grand total column
    xlGrandTotalRow                           =	2		# Grand total row
    xlHeaderRow                               =	1		# Header row
    xlLastColumn                              =	4		# Last column
    xlLastHeaderCell                          =	10		# Last header cell
    xlLastTotalCell                           =	12		# Last total cell
    xlPageFieldLabels                         =	26		# Page field labels
    xlPageFieldValues                         =	27		# Page field values
    xlRowStripe1                              =	5		# Row Stripe1
    xlRowStripe2                              =	6		# Row Stripe2
    xlRowSubheading1                          =	23		# Row Subheading1
    xlRowSubheading2                          =	24		# Row Subheading2
    xlRowSubheading3                          =	25		# Row Subheading3
    xlSlicerHoveredSelectedItemWithData       =	33		# A selected item, hovered over by the user, that contains data.
    xlSlicerHoveredSelectedItemWithNoData     =	35		# A selected item, hovered over by the user, that does not contain data.
    xlSlicerHoveredUnselectedItemWithData     =	32		# An item, hovered over by the user, that is not selected and that contains data.
    xlSlicerHoveredUnselectedItemWithNoData   =	34		# A selected item, hovered over by the user, that is not selected and that does not contain data.
    xlSlicerSelectedItemWithData              =	30		# A selected item that contains data.
    xlSlicerSelectedItemWithNoData            =	31		# A selected item that does not contain data.
    xlSlicerUnselectedItemWithData            =	28		# An item that is not selected that contains data.
    xlSlicerUnselectedItemWithNoData          =	29		# An item that is not selected that does not contain data.
    xlSubtotalColumn1                         =	13		# Subtotal Column1
    xlSubtotalColumn2                         =	14		# Subtotal Column2
    xlSubtotalColumn3                         =	15		# Subtotal Column3
    xlSubtotalRow1                            =	16		# Subtotal Row1
    xlSubtotalRow2                            =	17		# Subtotal Row2
    xlSubtotalRow3                            =	18		# Subtotal Row3
    xlTimelinePeriodLabels1                   =	38		# Timeline Period Label
    xlTimelinePeriodLabels2                   =	39		# Additional Timeline Period Label
    xlTimelineSelectedTimeBlock               =	40		# Selected Timeline Time Block
    xlTimelineSelectedTimeBlockSpace          =	42		# Selected Timeline Time Block space
    xlTimelineSelectionLabel                  =	36		# Timeline Selection Label
    xlTimelineTimeLevel                       =	37		# Timeline Level
    xlTimelineUnselectedTimeBlock             =	41		# Unselected Timeline Time Block
    xlTotalRow                                =	2		# Total Row
    xlWholeTable                              =	0		# Whole Table

    #values from excel.xltextparsingtype
    #****************************
    xlDelimited                               =	1		# Default. Indicates that the file is delimited by delimiter characters.
    xlFixedWidth                              =	2		# Indicates that the data in the file is arranged in columns of fixed widths.

    #values from excel.xltextqualifier
    #****************************
    xlTextQualifierDoubleQuote                =	1		# Double quotation mark (&quot;).
    xlTextQualifierNone                       =	-4142		# No delimiter.
    xlTextQualifierSingleQuote                =	2		# Single quotation mark (').

    #values from excel.xltextvisuallayouttype
    #****************************
    xlTextVisualLTR                           =	1		# Left-to-right
    xlTextVisualRTL                           =	2		# Right-to-left

    #values from excel.xlthemecolor
    #****************************
    xlThemeColorAccent1                       =	5		# Accent1
    xlThemeColorAccent2                       =	6		# Accent2
    xlThemeColorAccent3                       =	7		# Accent3
    xlThemeColorAccent4                       =	8		# Accent4
    xlThemeColorAccent5                       =	9		# Accent5
    xlThemeColorAccent6                       =	10		# Accent6
    xlThemeColorDark1                         =	1		# Dark1
    xlThemeColorDark2                         =	3		# Dark2
    xlThemeColorFollowedHyperlink             =	12		# Followed hyperlink
    xlThemeColorHyperlink                     =	11		# Hyperlink
    xlThemeColorLight1                        =	2		# Light1
    xlThemeColorLight2                        =	4		# Light2

    #values from excel.xlthemefont
    #****************************
    xlThemeFontMajor                          =	2		# Major.
    xlThemeFontMinor                          =	1		# Minor.
    xlThemeFontNone                           =	0		# Do not use any theme font.

    #values from excel.xlthreadmode
    #****************************
    xlThreadModeAutomatic                     =	0		# Multi-threaded calculation mode is automatic.
    xlThreadModeManual                        =	1		# Multi-threaded calculation mode is manual.

    #values from excel.xlticklabelorientation
    #****************************
    xlTickLabelOrientationAutomatic           =	-4105		# Text orientation set by Excel.
    xlTickLabelOrientationDownward            =	-4170		# Text runs down.
    xlTickLabelOrientationHorizontal          =	-4128		# Characters run horizontally.
    xlTickLabelOrientationUpward              =	-4171		# Text runs up.
    xlTickLabelOrientationVertical            =	-4166		# Characters run vertically.

    #values from excel.xlticklabelposition
    #****************************
    xlTickLabelPositionHigh                   =	-4127		# Top or right side of the chart.
    xlTickLabelPositionLow                    =	-4134		# Bottom or left side of the chart.
    xlTickLabelPositionNextToAxis             =	4		# Next to axis (where axis is not at either side of the chart).
    xlTickLabelPositionNone                   =	-4142		# No tick marks.

    #values from excel.xltickmark
    #****************************
    xlTickMarkCross                           =	4		# Crosses the axis
    xlTickMarkInside                          =	2		# Inside the axis
    xlTickMarkNone                            =	-4142		# No mark
    xlTickMarkOutside                         =	3		# Outside the axis

    #values from excel.xltimeperiods
    #****************************
    xlLast7Days                               =	2		# Last 7 days
    xlLastMonth                               =	5		# Last month
    xlLastWeek                                =	4		# Last week
    xlNextMonth                               =	8		# Next month
    xlNextWeek                                =	7		# Next week
    xlThisMonth                               =	9		# This month
    xlThisWeek                                =	3		# This week
    xlToday                                   =	0		# Today
    xlTomorrow                                =	6		# Tomorrow
    xlYesterday                               =	1		# Yesterday

    #values from excel.xltimeunit
    #****************************
    xlDays                                    =	0		# Days
    xlMonths                                  =	1		# Months
    xlYears                                   =	2		# Years

    #values from excel.xltimelinelevel
    #****************************
    xlTimelineLevelYears                      =	0		# Years level
    xlTimelineLevelQuarters                   =	1		# Quarters level
    xlTimelineLevelMonths                     =	2		# Months level
    xlTimelineLevelDays                       =	3		# Days level

    #values from excel.xltoolbarprotection
    #****************************
    xlNoButtonChanges                         =	1		# No button changes permitted.
    xlNoChanges                               =	4		# No changes of any kind.
    xlNoDockingChanges                        =	3		# No changes to toolbar's docking position.
    xlNoShapeChanges                          =	2		# No changes to toolbar shape.
    xlToolbarProtectionNone                   =	-4143		# All changes permitted.

    #values from excel.xltopbottom
    #****************************
    xlTop10Bottom                             =	0		# Top 10 bottom values
    xlTop10Top                                =	1		# Top 10 values

    #values from excel.xltotalscalculation
    #****************************
    xlTotalsCalculationAverage                =	2		# Average
    xlTotalsCalculationCount                  =	3		# Count of non-empty cells
    xlTotalsCalculationCountNums              =	4		# Count of cells with numeric values
    xlTotalsCalculationCustom                 =	9		# Custom calculation
    xlTotalsCalculationMax                    =	6		# Maximum value in the list
    xlTotalsCalculationMin                    =	5		# Minimum value in the list
    xlTotalsCalculationNone                   =	0		# No calculation
    xlTotalsCalculationStdDev                 =	7		# Standard deviation value
    xlTotalsCalculationSum                    =	1		# Sum of all values in the list column
    xlTotalsCalculationVar                    =	8		# Variable

    #values from excel.xltrendlinetype
    #****************************
    xlExponential                             =	5		# Uses an equation to calculate the least squares fit through points, for example, y=ab^x .
    xlLinear                                  =	-4132		# Uses the linear equation y = mx + b to calculate the least squares fit through points.
    xlLogarithmic                             =	-4133		# Uses the equation y = c ln x + b to calculate the least squares fit through points.
    xlMovingAvg                               =	6		# Uses a sequence of averages computed from parts of the data series. The number of points equals the total number of points in the series less the number specified for the period.
    xlPolynomial                              =	3		# Uses an equation to calculate the least squares fit through points, for example, y = ax^6 + bx^5 + cx^4 + dx^3 + ex^2 + fx + g.
    xlPower                                   =	4		# Uses an equation to calculate the least squares fit through points, for example, y = ax^b.

    #values from excel.xlunderlinestyle
    #****************************
    xlUnderlineStyleDouble                    =	-4119		# Double thick underline.
    xlUnderlineStyleDoubleAccounting          =	5		# Two thin underlines placed close together.
    xlUnderlineStyleNone                      =	-4142		# No underlining.
    xlUnderlineStyleSingle                    =	2		# Single underlining.
    xlUnderlineStyleSingleAccounting          =	4		# Not supported.

    #values from excel.xlupdatelinks
    #****************************
    xlUpdateLinksAlways                       =	3		# Embedded OLE links are always updated for the specified workbook.
    xlUpdateLinksNever                        =	2		# Embedded OLE links are never updated for the specified workbook.
    xlUpdateLinksUserSetting                  =	1		# Embedded OLE links are updated according to the user's settings for the specified workbook.

    #values from excel.xlvalign
    #****************************
    xlVAlignBottom                            =	-4107		# Bottom
    xlVAlignCenter                            =	-4108		# Center
    xlVAlignDistributed                       =	-4117		# Distributed
    xlVAlignJustify                           =	-4130		# Justify
    xlVAlignTop                               =	-4160		# Top

    #values from excel.xlwbatemplate
    #****************************
    xlWBATChart                               =	-4109		# Chart
    xlWBATExcel4IntlMacroSheet                =	4		# Excel version 4 macro
    xlWBATExcel4MacroSheet                    =	3		# Excel version 4 international macro
    xlWBATWorksheet                           =	-4167		# Worksheet

    #values from excel.xlwebformatting
    #****************************
    xlWebFormattingAll                        =	1		# All formatting is imported.
    xlWebFormattingNone                       =	3		# No formatting is imported.
    xlWebFormattingRTF                        =	2		# Rich Text Format?compatible formatting is imported.

    #values from excel.xlwebselectiontype
    #****************************
    xlAllTables                               =	2		# All tables
    xlEntirePage                              =	1		# Entire page
    xlSpecifiedTables                         =	3		# Specified tables

    #values from excel.xlwindowstate
    #****************************
    xlMaximized                               =	-4137		# Maximized
    xlMinimized                               =	-4140		# Minimized
    xlNormal                                  =	-4143		# Normal

    #values from excel.xlwindowtype
    #****************************
    xlChartAsWindow                           =	5		# The chart will open in a new window.
    xlChartInPlace                            =	4		# The chart will be displayed on the current worksheet.
    xlClipboard                               =	3		# The chart is copied to the clipboard.
    xlInfo                                    =	-4129		# This constant has been deprecated.
    xlWorkbook                                =	1		# This constant applies to Macintosh only.

    #values from excel.xlwindowview
    #****************************
    xlNormalView                              =	1		# Normal.
    xlPageBreakPreview                        =	2		# Page break preview.
    xlPageLayoutView                          =	3		# Page layout view.

    #values from excel.xlxlmmacrotype
    #****************************
    xlCommand                                 =	2		# Custom command.
    xlFunction                                =	1		# Custom function.
    xlNotXLM                                  =	3		# Not a macro.

    #values from excel.xlxmlexportresult
    #****************************
    xlXmlExportSuccess                        =	0		# The XML data file was successfully exported.
    xlXmlExportValidationFailed               =	1		# The contents of the XML data file do not match the specified schema map.

    #values from excel.xlxmlimportresult
    #****************************
    xlXmlImportElementsTruncated              =	1		# The contents of the specified XML data file have been truncated because the XML data file is too large for the worksheet.
    xlXmlImportSuccess                        =	0		# The XML data file was successfully imported.
    xlXmlImportValidationFailed               =	2		# The contents of the XML data file do not match the specified schema map.

    #values from excel.xlxmlloadoption
    #****************************
    xlXmlLoadImportToList                     =	2		# Places the contents of the XML data file in an XML table.
    xlXmlLoadMapXml                           =	3		# Displays the schema of the XML data file in the  XML Structure task pane.
    xlXmlLoadOpenXml                          =	1		# Opens the XML data file. The contents of the file will be flattened.
    xlXmlLoadPromptUser                       =	0		# Prompts the user to choose how to open the file.

    #values from excel.xlyesnoguess
    #****************************
    xlGuess                                   =	0		# Excel determines whether there is a header, and where it is, if there is one.
    xlNo                                      =	2		# Default. The entire range should be sorted.
    xlYes                                     =	1		# The entire range should not be sorted.
}
#End Enum

$xl = new-object PSCustomObject -Property $xl


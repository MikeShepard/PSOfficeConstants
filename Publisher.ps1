#constants for Publisher based on https://docs.microsoft.com/en-us/office/vba/api/Publisher(enumerations)
$pb = [Ordered]@{

    #values from publisher.pbbuildingblockgallery
    #****************************
    pbBBGalAccents                                   =	1		# Borders &amp; Accents gallery
    pbBBGalAdvertisements                            =	0		# Advertisements gallery
    pbBBGalBusinessInfo                              =	3		# Business Information gallery
    pbBBGalCalendars                                 =	2		# Calendars gallery
    pbBBGalNone                                      =	-1		# No gallery
    pbBBGalPageParts                                 =	4		# Page Parts gallery

    #values from publisher.pbbuildingblocktype
    #****************************
    pbBBBuiltIn                                      =	1		# Built-in type
    pbBBDownloaded                                   =	2		# Downloaded type
    pbBBNone                                         =	0		# No type
    pbBBUser                                         =	3		# User-defined type
    pbBBWorkgroup                                    =	4		# Workgroup-defined type

    #values from publisher.pbcalendartype
    #****************************
    pbCalendarTypeArabicHijri                        =	1		# Arabic Hijri calendar
    pbCalendarTypeChineseNational                    =	3		# Chinese National calendar
    pbCalendarTypeHebrewLunar                        =	2		# Hebrew Lunar calendar
    pbCalendarTypeJapaneseEmperor                    =	4		# Japanese Emperor calendar
    pbCalendarTypeKoreanDanki                        =	6		# Korean Danki calendar
    pbCalendarTypeSakaEra                            =	7		# Saka Era calendar
    pbCalendarTypeThaiBuddhist                       =	5		# Thai Buddhist calendar
    pbCalendarTypeTranslitEnglish                    =	8		# English calendar
    pbCalendarTypeTranslitFrench                     =	9		# French calendar
    pbCalendarTypeWestern                            =	0		# Western calendar

    #values from publisher.pbcanvasarrangementtype
    #****************************
    pbCanvasArrangementTypeColsCanvas                =	1		# Canvas arranged into columns.
    pbCanvasArrangementTypeOneCanvas                 =	0		# Canvas arranged as a single unit.
    pbCanvasArrangementTypeRowsCanvas                =	2		# Canvas arranged into rows.

    #values from publisher.pbcatalogmergefieldtype
    #****************************
    pbCatalogMergeFieldTypeText                      =	0		# Text field.
    pbCatalogMergeFieldTypePicture                   =	1		# Picture field.

    #values from publisher.pbcelldiagonaltype
    #****************************
    pbTableCellDiagonalDown                          =	2		# Diagonal Down
    pbTableCellDiagonalMixed                         =	-2		# Diagonal Mixed
    pbTableCellDiagonalNone                          =	0		# Not split Diagonally
    pbTableCellDiagonalUp                            =	1		# Diagonal Up

    #values from publisher.pbcollapsedirection
    #****************************
    pbCollapseEnd                                    =	2		# Collapse at the end.
    pbCollapseStart                                  =	1		# Collapse at the start.

    #values from publisher.pbcolormodel
    #****************************
    pbColorModelCMYK                                 =	2		# CMYK
    pbColorModelGreyScale                            =	3		# GreyScale
    pbColorModelRGB                                  =	1		# RGB
    pbColorModelUnknown                              =	4		# Unknown

    #values from publisher.pbcolorscheme
    #****************************
    pbColorSchemeAlpine                              =	-1		# Alpine
    pbColorSchemeAqua                                =	-2		# Aqua
    pbColorSchemeBerry                               =	-3		# Berry
    pbColorSchemeBlackGray                           =	-4		# Black Gray
    pbColorSchemeBlackWhite                          =	-58		# Black and White
    pbColorSchemeBrown                               =	-5		# Brown
    pbColorSchemeBurgundy                            =	-6		# Burgundy
    pbColorSchemeCavern                              =	-7		# Cavern
    pbColorSchemeCelebration                         =	-1004		# Celebration
    pbColorSchemeCherry                              =	-1002		# Cherry
    pbColorSchemeCitrus                              =	-8		# Citrus
    pbColorSchemeClay                                =	-9		# Clay
    pbColorSchemeCranberry                           =	-10		# Cranberry
    pbColorSchemeCrocus                              =	-11		# Crocus
    pbColorSchemeCustom                              =	1		# Custom
    pbColorSchemeDarkBlue                            =	-61		# DarkBlue
    pbColorSchemeDesert                              =	-12		# Desert
    pbColorSchemeField                               =	-13		# Field
    pbColorSchemeFirstUserDefined                    =	2000		# FirstUserDefined
    pbColorSchemeFjord                               =	-14		# Fjord
    pbColorSchemeFloral                              =	-15		# Floral
    pbColorSchemeGarnet                              =	-16		# Garnet
    pbColorSchemeGlacier                             =	-17

    #values from publisher.pbcolortype
    #****************************
    pbColorTypeCMS                                   =	4		# CMS
    pbColorTypeCMYK                                  =	3		# CMYK
    pbColorTypeInk                                   =	5		# Ink
    pbColorTypeMixed                                 =	-2		# Mixed
    pbColorTypeRGB                                   =	1		# RGB
    pbColorTypeScheme                                =	2		# Scheme

    #values from publisher.pbcommandbuttontype
    #****************************
    pbCommandButtonReset                             =	2		# Reset or clear the form.
    pbCommandButtonSubmit                            =	1		# Submit the form.

    #values from publisher.pbdatetimeformat
    #****************************
    pbDateEnglish                                    =	8		# English
    pbDateISO                                        =	4		# ISO
    pbDateLong                                       =	2		# Long
    pbDateLongDay                                    =	1		# Long Day
    pbDateMon_Yr                                     =	10		# Month or Year
    pbDateMonthYr                                    =	9		# Month and Year
    pbDateShort                                      =	0		# Short
    pbDateShortAbb                                   =	7		# Short Abb
    pbDateShortAlt                                   =	3		# Short Alt
    pbDateShortMon                                   =	5		# Short Month
    pbDateShortSlash                                 =	6		# Short Slash
    pbDateTimeEastAsia1                              =	17		# East Asia 1
    pbDateTimeEastAsia2                              =	18		# East Asia 2
    pbDateTimeEastAsia3                              =	19		# East Asia 3
    pbDateTimeEastAsia4                              =	20		# East Asia 4
    pbDateTimeEastAsia5                              =	21		# East Asia 5
    pbTime24                                         =	15		# 24 hours
    pbTimeDatePM                                     =	11		# Date/Time in A.M./P.M. format
    pbTimeDateSecPM                                  =	12		# Date/Time/Second in A.M./P.M. format
    pbTimePM                                         =	13		# Time in A.M./P.M. format
    pbTimeSec24                                      =	16		# Time in 24-hour format
    pbTimeSecPM                                      =	14		# Time including seconds in A.M./P.M. format

    #values from publisher.pbdirectiontype
    #****************************
    pbDirectionLeftToRight                           =	1		# Left to Right
    pbDirectionRightToLeft                           =	2		# Right to Left

    #values from publisher.pbdrivertype
    #****************************
    pbDriverTypeNonPostScript                        =	1		# Non PostScript
    pbDriverTypePostScript1                          =	2		# PostScript 1
    pbDriverTypePostScript2                          =	3		# PostScript 2
    pbDriverTypePostScript3                          =	4		# PostScript 3

    #values from publisher.pbemailmergepriority
    #****************************
    pbPriorityHigh                                   =	1		# High priority
    pbPriorityLow                                    =	2		# Low priority
    pbPriorityNone                                   =	0		# No priority set

    #values from publisher.pbfieldtype
    #****************************
    pbFieldDateTime                                  =	4		# Date and time
    pbFieldHyperlinkAbsolutePage                     =	11		# Absolute page hyperlink
    pbFieldHyperlinkEmail                            =	12		# Email hyperlink
    pbFieldHyperlinkFile                             =	13		# File hyperlink
    pbFieldHyperlinkRelativePage                     =	10		# Relative page hyperlink
    pbFieldHyperlinkURL                              =	9		# URL hyperlink
    pbFieldIHIV                                      =	6		# IHIV
    pbFieldMailMerge                                 =	5		# Mail merge
    pbFieldNone                                      =	0		# None
    pbFieldPageNumber                                =	1		# Page number
    pbFieldPageNumberNext                            =	2		# Next page number
    pbFieldPageNumberPrev                            =	3		# Previous page number
    pbFieldPersonalizedHyperlinkURL                  =	14		# Personalized hyperlink URL
    pbFieldPhoneticGuide                             =	7		# Phonetic guide
    pbFieldWizardSampleText                          =	8		# Wizard sample text

    #values from publisher.pbfileformat
    #****************************
    pbFileHTMLFiltered                               =	7		# The file was saved in HTML Filtered format.
    pbFilePlainText                                  =	8		# The file was saved in plain text format.
    pbFilePublication                                =	1		# The file was saved with the current version of Microsoft Publisher.
    pbFilePublisher2000                              =	3		# The file was saved in a Publisher 2000 file format.
    pbFilePublisher98                                =	2		# The file was saved in a Publisher 98 file format.
    pbFileRTF                                        =	6		# The file was saved in Rich Text Format.
    pbFileUnicodeText                                =	9		# The file was saved in Unicode Text Format.
    pbFileWebArchive                                 =	5		# The file was saved in the MHTML format that allows users to save a Web page and all its related files as a single file.

    #values from publisher.pbfixedformatintent
    #****************************
    pbIntentCommercial                               =	4		# Submit the publication to a commercial press.
    pbIntentMinimum                                  =	1		# Squeeze the publication to the smallest file size. This satisfies the on-screen viewing scenario where the publication is viewed on a computer monitor.
    pbIntentPrinting                                 =	3		# Print the publication on a desktop printer or at a copy store, such as Kinko's.
    pbIntentStandard                                 =	2		# Distribute the publication as an email message or from a Web site. Note that the user does not know how the publication will be viewed: on-screen or printed from a desktop printer. Both the desktop printing scenario and the on-screen viewing scenario must be met by this intent.

    #values from publisher.pbfixedformattype
    #****************************
    pbFixedFormatTypePDF                             =	2		# PDF format
    pbFixedFormatTypeXPS                             =	1		# XPS format

    #values from publisher.pbfontscripttype
    #****************************
    pbFontScriptArabic                               =	7		# Arabic
    pbFontScriptArmenian                             =	5		# Armenian
    pbFontScriptAsciiLatin                           =	1		# Ascii Latin
    pbFontScriptAsciiSym                             =	43		# Ascii Sym
    pbFontScriptBengali                              =	9		# Bengali
    pbFontScriptBopomofo                             =	23		# Bopomofo
    pbFontScriptBraille                              =	41		# Braille
    pbFontScriptCanadianAbor                         =	36		# Canadian Abor
    pbFontScriptCherokee                             =	35		# Cherokee
    pbFontScriptCurrency                             =	42		# Currency
    pbFontScriptCyrillic                             =	4		# Cyrillic
    pbFontScriptDefault                              =	0		# Default
    pbFontScriptDevanagari                           =	8		# Devanagari
    pbFontScriptEthiopic                             =	34		# Ethiopic
    pbFontScriptEUDC                                 =	26		# EUDC
    pbFontScriptGeorgian                             =	20		# Georgian
    pbFontScriptGreek                                =	3		# Greek
    pbFontScriptGujarati                             =	11		# Gujarati
    pbFontScriptGurmukhi                             =	10		# Gurmukhi
    pbFontScriptHalfWidthKana                        =	25		# Half WidthKana
    pbFontScriptHan                                  =	24		# Han
    pbFontScriptHangul                               =	21		# Hangul
    pbFontScriptHanSurrogate                         =	28		# Han Surrogate
    pbFontScriptHebrew                               =	6		# Hebrew
    pbFontScriptKana                                 =	22		# Kana
    pbFontScriptKannada                              =	15		# Kannada
    pbFontScriptKhmer                                =	39		# Khmer
    pbFontScriptLao                                  =	18		# Lao
    pbFontScriptLatin                                =	2		# Latin
    pbFontScriptMalayalam                            =	16		# Malayalam
    pbFontScriptMixed                                =	-2		# Mixed
    pbFontScriptMongolian                            =	40		# Mongolian
    pbFontScriptMyanmar                              =	32		# Myanmar
    pbFontScriptNonHanSurrogate                      =	29		# Non-Han Surrogate
    pbFontScriptOgham                                =	37		# Ogham
    pbFontScriptOriya                                =	12		# Oriya
    pbFontScriptRunic                                =	38		# Runic
    pbFontScriptSinhala                              =	33		# Sinhala
    pbFontScriptSyriac                               =	30		# Syriac
    pbFontScriptTamil                                =	13		# Tamil
    pbFontScriptTelugu                               =	14		# Telugu
    pbFontScriptThaana                               =	31		# Thaana
    pbFontScriptThai                                 =	17		# Thai
    pbFontScriptTibetan                              =	19		# Tibetan
    pbFontScriptYi                                   =	27		# Yi

    #values from publisher.pbhelptype
    #****************************
    pbHelp                                           =	1		# Displays the  Help Topics dialog box.
    pbHelpActiveWindow                               =	2		# Displays Help describing the command associated with the active view or pane.
    pbHelpPSSHelp                                    =	3		# Displays product support information.

    #values from publisher.pbhlinktargettype
    #****************************
    pbHlinkTargetTypeEmail                           =	2		# Email
    pbHlinkTargetTypeFirstPage                       =	3		# First Page
    pbHlinkTargetTypeLastPage                        =	4		# Last Page
    pbHlinkTargetTypeNextPage                        =	5		# Next Page
    pbHlinkTargetTypeNone                            =	0		# None
    pbHlinkTargetTypePageID                          =	7		# Page ID
    pbHlinkTargetTypePersonalized                    =	8		# Personalized
    pbHlinkTargetTypePreviousPage                    =	6		# Previous Page
    pbHlinkTargetTypeURL                             =	1		# URL

    #values from publisher.pbhorizontalpicturelocking
    #****************************
    pbHorizontalLockingLeft                          =	1		# New pictures are inserted along the left edge of the frame.
    pbHorizontalLockingNone                          =	0		# New pictures are inserted in the middle between the left and right edges of the frame.
    pbHorizontalLockingRight                         =	2		# New pictures are inserted along the right edge of the frame.
    pbHorizontalLockingStretch                       =	3		# New pictures are horizontally stretched to the full width of the frame.

    #values from publisher.pbimageformat
    #****************************
    pbImageFormatCMYKJPEG                            =	10		# CMYKJPEG
    pbImageFormatDIB                                 =	7		# DIB
    pbImageFormatEMF                                 =	2		# EMF
    pbImageFormatGIF                                 =	8		# GIF
    pbImageFormatJPEG                                =	5		# JPEG
    pbImageFormatPICT                                =	4		# PICT
    pbImageFormatPNG                                 =	6		# PNG
    pbImageFormatTIFF                                =	9		# TIFF
    pbImageFormatUNKNOWN                             =	1		# Unknown
    pbImageFormatWMF                                 =	3		# WMF

    #values from publisher.pbinlinealignment
    #****************************
    pbInlineAlignmentCharacter                       =	0		# Shape is aligned with the text characters.
    pbInlineAlignmentLeft                            =	1		# Shape is left-aligned.
    pbInlineAlignmentMixed                           =	-2		# Shape is mixed-aligned.
    pbInlineAlignmentRight                           =	2		# Shape is right-aligned.

    #values from publisher.pbligaturepresettype
    #****************************
    pbLigatureAll                                    =	3		# Ligature applied to all characters
    pbLigatureMixed                                  =	-1		# Ligature applied to some characters, but not others
    pbLigatureNone                                   =	4		# No ligature applied
    pbLigatureStandard                               =	0		# Standard ligature applied
    pbLigatureStandardHistorical                     =	2		# Standard historical ligature applied
    pbLigatureStandardOptional                       =	1		# Standard optional ligature applied

    #values from publisher.pblinespacingrule
    #****************************
    pbLineSpacing1pt5                                =	1		# Sets paragraph line spacing to a line and a half.
    pbLineSpacingDouble                              =	2		# Sets paragraph line spacing to two lines.
    pbLineSpacingExactly                             =	4		# Sets paragraph line spacing to exactly the value of the  LineSpacing property, even if a larger font is used within the paragraph.
    pbLineSpacingMixed                               =	-9999999		# A return value for a paragraph that has line spacing of varying values.
    pbLineSpacingMultiple                            =	5		# A  LineSpacing property value must be specified, in number of lines.
    pbLineSpacingSingle                              =	0		# Sets paragraph line spacing to one space.

    #values from publisher.pblinkedfilestatus
    #****************************
    pbLinkedFileMissing                              =	2		# The file can no longer be found at the specified path.
    pbLinkedFileModified                             =	3		# The linked file has been modified since it was linked to the picture.
    pbLinkedFileOK                                   =	1		# The file resides at the specified path, and has not been modified since it was linked to the picture.

    #values from publisher.pblistseparator
    #****************************
    pbListSeparatorColon                             =	327680		# Colon
    pbListSeparatorDoubleHyphen                      =	458752		# Double Hyphen
    pbListSeparatorDoubleParen                       =	65536		# Double Parenthesis
    pbListSeparatorDoubleSquare                      =	393216		# Double Square
    pbListSeparatorParenthesis                       =	0		# Parenthesis
    pbListSeparatorPeriod                            =	131072		# Period
    pbListSeparatorPlain                             =	196608		# Plain
    pbListSeparatorSquare                            =	262144		# Square
    pbListSeparatorWideComma                         =	524288		# WideComma

    #values from publisher.pblisttype
    #****************************
    pbListTypeAiueo                                  =	12		# Aiueo
    pbListTypeArabic                                 =	0		# Arabic
    pbListTypeArabic1                                =	46		# Arabic1
    pbListTypeArabic2                                =	48		# Arabic2
    pbListTypeArabicLeadingZero                      =	22		# Arabic Leading Zero
    pbListTypeBullet                                 =	23		# Bullet
    pbListTypeCardinalText                           =	6		# Cardinal Text
    pbListTypeChnDbNum2                              =	38		# ChnDbNum2
    pbListTypeChnDbNum3                              =	39		# ChnDbNum3
    pbListTypeChosung                                =	25		# Chosung
    pbListTypeCirclenum                              =	18		# Circle num
    pbListTypeDAiueo                                 =	20		# DAiueo
    pbListTypeDbChar                                 =	14		# DbChar
    pbListTypeDbNum1                                 =	10		# DbNum1
    pbListTypeDbNum2                                 =	11		# DbNum2
    pbListTypeDbNum3                                 =	16		# DbNum3
    pbListTypeDbNum4                                 =	17		# DbNum4
    pbListTypeDIroha                                 =	21		# DIroha
    pbListTypeGanada                                 =	24		# Ganada
    pbListTypeHebrew1                                =	45		# Hebrew1
    pbListTypeHebrew2                                =	47		# Hebrew2
    pbListTypeHindi1                                 =	49		# Hindi1
    pbListTypeHindi2                                 =	50		# Hindi2
    pbListTypeHindi3                                 =	51		# Hindi3
    pbListTypeHindi4                                 =	52		# Hindi4
    pbListTypeIroha                                  =	13		# Iroha
    pbListTypeKorDbNum1                              =	41		# KorDbNum1
    pbListTypeKorDbNum2                              =	42		# KorDbNum2
    pbListTypeKorDbNum3                              =	43		# KorDbNum3
    pbListTypeKorDbNum4                              =	44		# KorDbNum4
    pbListTypeLowerCaseLetter                        =	4		# Lowercase Letter
    pbListTypeLowerCaseRoman                         =	2		# Lowercase Roman
    pbListTypeLowerCaseRussian                       =	58		# Lowercase Russian
    pbListTypeNone                                   =	255		# None
    pbListTypeOrdinal                                =	5		# Ordinal
    pbListTypeOrdinalText                            =	7		# OrdinalText
    pbListTypeThai1                                  =	53		# Thai1
    pbListTypeThai2                                  =	54		# Thai2
    pbListTypeThai3                                  =	55		# Thai3
    pbListTypeTpeDbNum2                              =	34		# DbNum2
    pbListTypeTpeDbNum3                              =	35		# DbNum3
    pbListTypeUpperCaseLetter                        =	3		# Uppercase Letter
    pbListTypeUpperCaseRoman                         =	1		# Uppercase Roman
    pbListTypeUpperCaseRussian                       =	59		# Uppercase Russian
    pbListTypeVietnamese1                            =	56		# Vietnamese1
    pbListTypeZodiac1                                =	30		# Zodiac1
    pbListTypeZodiac2                                =	31		# Zodiac2

    #values from publisher.pbmailmergedatafieldtype
    #****************************
    pbMailMergeDataFieldPicture                      =	1		# Contains a picture.
    pbMailMergeDataFieldString                       =	0		# Contains a string.

    #values from publisher.pbmailmergedatasource
    #****************************
    pbMergeInfoFromODSO                              =	5		# From ODSO
    pbMergeInfoSubODSO                               =	6		# Sub ODSO

    #values from publisher.pbmailmergedestination
    #****************************
    pbMergeToExistingPublication                     =	3		# Default Merge to an exisiting presentation.
    pbMergeToNewPublication                          =	2		# Merge to a new presentation.
    pbSendEmail                                      =	4		# Merge and send as an email message.
    pbSendToPrinter                                  =	1		# Merge and send to the printer.

    #values from publisher.pbmergetype
    #****************************
    pbCatalogMerge                                   =	3		# Catalog merge
    pbEmailMerge                                     =	4		# Email merge
    pbMailMerge                                      =	2		# Mail merge
    pbMergeDefault                                   =	0		# Default merge

    #values from publisher.pbnavbarorientation
    #****************************
    pbNavBarOrientHorizontal                         =	1		# Horizontal orientation
    pbNavBarOrientVertical                           =	2		# Vertical orientation

    #values from publisher.pbnumberstylestype
    #****************************
    pbNumberStyleDefault                             =	0		# Default number style
    pbNumberStyleMixed                               =	-1		# Mixed number styles
    pbNumberStyleProportionalLining                  =	1		# Full-height numbers spaced proportionally
    pbNumberStyleProportionalOldstyle                =	3		# Numbers that read well with text
    pbNumberStyleTabularLining                       =	2		# Full-height numbers spaced equally
    pbNumberStyleTabularOldstyle                     =	4		# Numbers that read well and are spaced equally

    #values from publisher.pborientationtype
    #****************************
    pbOrientationLandscape                           =	2		# Landscape orientation
    pbOrientationPortrait                            =	1		# Portrait orientation

    #values from publisher.pbpagenumberformat
    #****************************
    pbPageNumberFormatAiueo                          =	12		# Aiueo
    pbPageNumberFormatArabic                         =	0		# Arabic
    pbPageNumberFormatArabic1                        =	46		# Arabic1
    pbPageNumberFormatArabic2                        =	48		# Arabic2
    pbPageNumberFormatArabicLZ                       =	22		# ArabicLZ
    pbPageNumberFormatCardtext                       =	6		# Cardtext
    pbPageNumberFormatChnDbNum2                      =	38		# ChnDbNum2
    pbPageNumberFormatChnDbNum3                      =	39		# ChnDbNum3
    pbPageNumberFormatChosung                        =	25		# Chosung
    pbPageNumberFormatCirclenum                      =	18		# Circlenum
    pbPageNumberFormatDAiueo                         =	20		# DAiueo
    pbPageNumberFormatDbChar                         =	14		# DbChar
    pbPageNumberFormatDbNum1                         =	10		# DbNum1
    pbPageNumberFormatDbNum2                         =	11		# DbNum2
    pbPageNumberFormatDbNum3                         =	16		# DbNum3
    pbPageNumberFormatDIroha                         =	21		# DIroha
    pbPageNumberFormatGanada                         =	24		# Ganada
    pbPageNumberFormatHebrew1                        =	45		# Hebrew1
    pbPageNumberFormatHebrew2                        =	47		# Hebrew2
    pbPageNumberFormatHindi1                         =	49		# Hindi1
    pbPageNumberFormatHindi2                         =	50		# Hindi2
    pbPageNumberFormatHindi3                         =	51		# Hindi3
    pbPageNumberFormatHindi4                         =	52		# Hindi4
    pbPageNumberFormatIroha                          =	13		# Iroha
    pbPageNumberFormatKorDbNum1                      =	41		# KorDbNum1
    pbPageNumberFormatKorDbNum2                      =	42		# KorDbNum2
    pbPageNumberFormatKorDbNum3                      =	43		# KorDbNum3
    pbPageNumberFormatKorDbNum4                      =	44		# KorDbNum4
    pbPageNumberFormatLCLetter                       =	4		# LC Letter
    pbPageNumberFormatLCRoman                        =	2		# LC Roman
    pbPageNumberFormatLCRus                          =	58		# LC Rus
    pbPageNumberFormatOrdinal                        =	5		# Ordinal
    pbPageNumberFormatOrdtext                        =	7		# Ordtext
    pbPageNumberFormatThai1                          =	53		# Thai1
    pbPageNumberFormatThai2                          =	54		# Thai2
    pbPageNumberFormatThai3                          =	55		# Thai3
    pbPageNumberFormatTpeDbNum2                      =	34		# TpeDbNum2
    pbPageNumberFormatTpeDbNum3                      =	35		# TpeDbNum3
    pbPageNumberFormatUCLetter                       =	3		# UC Letter
    pbPageNumberFormatUCRoman                        =	1		# UC Roman
    pbPageNumberFormatUCRus                          =	59		# UC Rus
    pbPageNumberFormatViet1                          =	56		# Viet 1
    pbPageNumberFormatZodiac1                        =	30		# Zodiac1
    pbPageNumberFormatZodiac2                        =	31		# Zodiac2

    #values from publisher.pbpagenumbertype
    #****************************
    pbPageNumberCurrent                              =	1		# Default.
    pbPageNumberNextInStory                          =	2		# Inserts the page number of the next linked text box.
    pbPageNumberPreviousInStory                      =	3		# Inserts the page number of the previous linked text box.

    #values from publisher.pbpagetype
    #****************************
    pbPageLeftPage                                   =	1		# Left page
    pbPageMasterPage                                 =	4		# Master page
    pbPageRightPage                                  =	2		# Right page
    pbPageScratchPage                                =	3		# Scratch page

    #values from publisher.pbparagraphalignmenttype
    #****************************
    pbParagraphAlignmentCenter                       =	1		# Center alignment
    pbParagraphAlignmentDistribute                   =	4		# Distribute alignment
    pbParagraphAlignmentDistributeAll                =	9		# Distribute all
    pbParagraphAlignmentDistributeCenterLast         =	10		# Distribute center last
    pbParagraphAlignmentDistributeEastAsia           =	5		# Distribute East Asia
    pbParagraphAlignmentInterCluster                 =	8		# Inter Cluster
    pbParagraphAlignmentInterIdeograph               =	7		# Inter Ideograph
    pbParagraphAlignmentInterWord                    =	3		# Inter Word
    pbParagraphAlignmentJustified                    =	6		# Justified
    pbParagraphAlignmentKashida                      =	11		# Kashida
    pbParagraphAlignmentLeft                         =	0		# Left alignment
    pbParagraphAlignmentMixed                        =	-9999999		# Mixed alignment
    pbParagraphAlignmentRight                        =	2		# Right alignment

    #values from publisher.pbpersonalinfoset
    #****************************
    pbPersonalInfoHome                               =	4		# Home
    pbPersonalInfoOtherOrganization                  =	3		# Other organization
    pbPersonalInfoPrimaryBusiness                    =	1		# Primary business
    pbPersonalInfoSecondaryBusiness                  =	2		# Secondary business

    #values from publisher.pbphoneticguidealignmenttype
    #****************************
    pbPhoneticGuideAlignmentCenter                   =	3		# Center aligned
    pbPhoneticGuideAlignmentDefault                  =	0		# Default alignment
    pbPhoneticGuideAlignmentLeft                     =	4		# Left alignment
    pbPhoneticGuideAlignmentOneTwoOne                =	2		# One Two One alignment
    pbPhoneticGuideAlignmentRight                    =	5		# Right alignment
    pbPhoneticGuideAlignmentZeroOneZero              =	1		# Zero One Zero alignment

    #values from publisher.pbpictureinsertas
    #****************************
    pbPictureInsertAsEmbedded                        =	1		# Embed all images.
    pbPictureInsertAsLinked                          =	2		# Images can either be linked externally or internally.
    pbPictureInsertAsOriginalState                   =	3		# Default. Image is inserted in its original state.

    #values from publisher.pbpictureinsertfit
    #****************************


    #values from publisher.pbpictureresolution
    #****************************
    pbPictureResolutionCommercialPrint_300dpi        =	3		# 300 dpi
    pbPictureResolutionDefault                       =	0		# Default
    pbPictureResolutionDesktopPrint_150dpi           =	2		# 150 dpi
    pbPictureResolutionWeb_96dpi                     =	1		# 96 dpi

    #values from publisher.pbplacementtype
    #****************************
    pbPlacementCenter                                =	3		# Center
    pbPlacementLeft                                  =	1		# Left
    pbPlacementRight                                 =	2		# Right

    #values from publisher.pbpresetwordart
    #****************************
    pbPresetWordArt1                                 =	0		# Style 1
    pbPresetWordArt2                                 =	1		# Style 2
    pbPresetWordArt3                                 =	2		# Style 3
    pbPresetWordArt4                                 =	3		# Style 4
    pbPresetWordArt5                                 =	4		# Style 5
    pbPresetWordArt6                                 =	5		# Style 6
    pbPresetWordArt7                                 =	6		# Style 7
    pbPresetWordArt8                                 =	7		# Style 8
    pbPresetWordArt9                                 =	8		# Style 9
    pbPresetWordArt10                                =	9		# Style 10
    pbPresetWordArt11                                =	10		# Style 11
    pbPresetWordArt12                                =	11		# Style 12
    pbPresetWordArt13                                =	12		# Style 13
    pbPresetWordArt14                                =	13		# Style 14
    pbPresetWordArt15                                =	14		# Style 15
    pbPresetWordArt16                                =	15		# Style 16
    pbPresetWordArt17                                =	16		# Style 17
    pbPresetWordArt18                                =	17		# Style 18
    pbPresetWordArt19                                =	18		# Style 19
    pbPresetWordArt20                                =	19		# Style 20
    pbPresetWordArt21                                =	20		# Style 21
    pbPresetWordArt22                                =	21		# Style 22
    pbPresetWordArt23                                =	22		# Style 23
    pbPresetWordArt24                                =	23		# Style 24
    pbPresetWordArt25                                =	24		# Style 25
    pbPresetWordArt26                                =	25		# Style 26
    pbPresetWordArt27                                =	26		# Style 27
    pbPresetWordArt28                                =	27		# Style 28
    pbPresetWordArt29                                =	28		# Style 29
    pbPresetWordArt30                                =	29		# Style 30
    pbPresetWordArt31                                =	30		# Style 31
    pbPresetWordArt32                                =	31		# Style 32
    pbPresetWordArt33                                =	32		# Style 33
    pbPresetWordArt34                                =	33		# Style 34
    pbPresetWordArt35                                =	34		# Style 35
    pbPresetWordArt36                                =	35		# Style 36
    pbPresetWordArt37                                =	36		# Style 37
    pbPresetWordArt38                                =	37		# Style 38
    pbPresetWordArt39                                =	38		# Style 39
    pbPresetWordArt40                                =	39		# Style 40
    pbPresetWordArt41                                =	40		# Style 41
    pbPresetWordArt42                                =	41		# Style 42
    pbPresetWordArt43                                =	42		# Style 43
    pbPresetWordArt44                                =	43		# Style 44
    pbPresetWordArt45                                =	44		# Style 45
    pbPresetWordArt46                                =	45		# Style 46
    pbPresetWordArt47                                =	46		# Style 47
    pbPresetWordArt48                                =	47		# Style 48
    pbPresetWordArt49                                =	48		# Style 49
    pbPresetWordArt50                                =	49		# Style 50
    pbPresetWordArt51                                =	50		# Style 51
    pbPresetWordArt52                                =	51		# Style 52
    pbPresetWordArt53                                =	52		# Style 53
    pbPresetWordArt54                                =	53		# Style 54
    pbPresetWordArt55                                =	54		# Style 55
    pbPresetWordArt56                                =	55		# Style 56
    pbPresetWordArt57                                =	56		# Style 57
    pbPresetWordArt58                                =	57		# Style 58
    pbPresetWordArt59                                =	58		# Style 59
    pbPresetWordArt60                                =	59		# Style 60
    pbPresetWordArtMixed                             =	-2		# A combination of styles

    #values from publisher.pbprintgraphics
    #****************************
    pbPrintHighResolution                            =	1		# Default. Print linked graphics using the full-resolution linked version.
    pbPrintLowResolution                             =	2		# Print linked graphics using the low-resolution placeholder version that is stored in the publication.
    pbPrintNoGraphics                                =	3		# Print a box in place of linked graphics.

    #values from publisher.pbprintmode
    #****************************
    pbPrintModeCompositeCMYK                         =	3		# Print a composite whose colors are defined by the CMYK color model.
    pbPrintModeCompositeGrayscale                    =	4		# Print a composite whose colors are defined as shades of gray.
    pbPrintModeCompositeRGB                          =	1		# Print a composite whose colors are defined by the RGB color model.
    pbPrintModeSeparations                           =	2		# Print a separate plate for each ink used in the publication.

    #values from publisher.pbprintstyle
    #****************************
    pbPrintStyleBookletSideFold                      =	5		# Prints the publication in the booklet style with a side fold.
    pbPrintStyleBookletTopFold                       =	6		# Prints the publication in the booklet style with a top fold.
    pbPrintStyleDefault                              =	0		# Prints the publication in the default style.
    pbPrintStyleEnvelope                             =	11		# Prints the publication in the envelope style.
    pbPrintStyleHalfFoldSide                         =	7		# Prints the publication in the half fold on the side style.
    pbPrintStyleHalfFoldTop                          =	8		# Prints the publication in the half fold on the top style.
    pbPrintStyleMultipleCopiesPerSheet               =	3		# Prints multiple copies of the publication per sheet.
    pbPrintStyleMultiplePagesPerSheet                =	4		# Prints multiple pages of the publication per sheet.
    pbPrintStyleOnePagePerSheet                      =	1		# Prints one page of the publication on one sheet.
    pbPrintStyleQuarterFoldSide                      =	10		# Prints the publication in the quarter fold side style.
    pbPrintStyleQuarterFoldTop                       =	9		# Prints the publication in the quarter fold top style.
    pbPrintStyleTiled                                =	2		# Prints the publication in the tiled style.

    #values from publisher.pbpublicationlayout
    #****************************


    #values from publisher.pbpublicationtype
    #****************************
    pbTypePrint                                      =	1		# Print publication
    pbTypeWeb                                        =	2		# Web publication

    #values from publisher.pbrecipientlistfiletype
    #****************************
    pbAsCsvFile                                      =	1		# Save as comma-delimited CSV file.
    pbAsMdbFile                                      =	0		# Save as Microsoft Office Access MDB file.

    #values from publisher.pbreplacescope
    #****************************
    pbReplaceScopeAll                                =	2		# All need to be replaced.
    pbReplaceScopeNone                               =	0		# None need to be replaced.
    pbReplaceScopeOne                                =	1		# One needs to be replaced.

    #values from publisher.pbreplacetint
    #****************************
    pbReplaceTintKeepTints                           =	1		# Maintain the same tint percentage in the ink represented by the replacement plate as in the deleted plate. For example, replace a 100% tint of yellow with a 100% tint of blue.
    pbReplaceTintMaintainLuminosity                  =	2		# Maintain the same lightness value in the ink represented by the replacement plate as in the deleted plate. For example, replace a 100% tint of yellow with an approximately 10% tint of blue.
    pbReplaceTintUseDefault                          =	0		# Default.

    #values from publisher.pbrulerguidetype
    #****************************
    pbRulerGuideTypeHorizontal                       =	2		# Horizontal ruler
    pbRulerGuideTypeVertical                         =	1		# Vertical ruler

    #values from publisher.pbsaveoptions
    #****************************
    pbDoNotSaveChanges                               =	3		# Close the open publication without saving any changes.
    pbPromptToSaveChanges                            =	1		# Default. Prompt the user whether to save changes in the open publication.
    pbSaveChanges                                    =	2		# Save the open publication before closing it.

    #values from publisher.pbschemecolorindex
    #****************************
    pbSchemeColorAccent1                             =	2		# Sets the color to Accent1 scheme color.
    pbSchemeColorAccent2                             =	3		# Sets the color to Accent2 scheme color.
    pbSchemeColorAccent3                             =	4		# Sets the color to Accent3 scheme color.
    pbSchemeColorAccent4                             =	5		# Sets the color to Accent4 scheme color.
    pbSchemeColorAccent5                             =	8		# Sets the color to Accent5 scheme color.
    pbSchemeColorFollowedHyperlink                   =	7		# Sets the color scheme to followed hyperlink.
    pbSchemeColorHyperlink                           =	6		# Sets the color scheme to color hyperlink.
    pbSchemeColorMain                                =	1		# Sets the color to main.
    pbSchemeColorNone                                =	0		# Sets the color scheme to none.

    #values from publisher.pbselectiontype
    #****************************
    pbSelectionNone                                  =	0		# Selection type none.
    pbSelectionShape                                 =	1		# Selection type shape.
    pbSelectionShapeSubSelection                     =	4		# Selection type subselection.
    pbSelectionTableCells                            =	3		# Selection type table cells.
    pbSelectionText                                  =	2		# Selection type text.

    #values from publisher.pbshapetype
    #****************************
    pbAutoShape                                      =	1		# AutoShape
    pbCallout                                        =	2		# Callout
    pbCatalogMergeArea                               =	111		# Catalog Merge Area
    pbChart                                          =	3		# Chart
    pbComment                                        =	4		# Comment
    pbEmbeddedOLEObject                              =	7		# Embedded OLE Object
    pbFormControl                                    =	8		# Form Control
    pbFreeform                                       =	5		# Freeform
    pbGroup                                          =	6		# Group
    pbGroupWizard                                    =	108		# Group Wizard
    pbLine                                           =	9		# Line
    pbLinkedOLEObject                                =	10		# Linked OLE Object
    pbLinkedPicture                                  =	11		# Picture
    pbMedia                                          =	16		# Media
    pbOLEControlObject                               =	12		# OLE Control Object
    pbPicture                                        =	13		# Picture
    pbPlaceholder                                    =	14		# Placeholder
    pbShapeTypeMixed                                 =	-2		# Shape Type Mixed
    pbTable                                          =	18		# Table
    pbTextEffect                                     =	15		# Text effect
    pbTextFrame                                      =	17		# Text frame
    pbWebCheckBox                                    =	100		# Web check box
    pbWebCommandButton                               =	101		# Web command button
    pbWebHotSpot                                     =	110		# Web hot spot
    pbWebHTMLFragment                                =	107		# Web HTML Fragment
    pbWebListBox                                     =	102		# Web list box
    pbWebMultiLineTextBox                            =	103		# Web multiLine text box
    pbWebNavigationBar                               =	112		# Web navigation bar
    pbWebOptionButton                                =	104		# Web option button
    pbWebSingleLineTextBox                           =	105		# Web single-line textbox
    pbWebWebComponent                                =	106		# Web Web component

    #values from publisher.pbstorytype
    #****************************
    pbStoryContinuedFrom                             =	2		# Story continued from which text frame.
    pbStoryContinuedOn                               =	3		# Story continued on to which text frame.
    pbStoryTable                                     =	0		# Story table.
    pbStoryTextFrame                                 =	1		# Story text frame.

    #values from publisher.pbsubmitdataformattype
    #****************************
    pbSubmitDataFormatCSV                            =	3		# Saves Web form data to a comma-delimited text file.
    pbSubmitDataFormatHTML                           =	1		# Saves Web form data to an HTML file.
    pbSubmitDataFormatRichText                       =	2		# Saves Web form data to a formatted file.
    pbSubmitDataFormatTab                            =	4		# Saves Web form data to a tab-delimited text file.

    #values from publisher.pbsubmitdataretrievalmethodtype
    #****************************
    pbSubmitDataRetrievalEmail                       =	2		# Processes form data by sending an email message to a specified email address.
    pbSubmitDataRetrievalProgram                     =	3		# Processes form data using a script program provided by your Internet service provider.
    pbSubmitDataRetrievalSaveOnServer                =	1		# Saves form data to a file stored on your Web server.

    #values from publisher.pbtabalignmenttype
    #****************************
    pbTabAlignmentCenter                             =	1		# Central tab alignment
    pbTabAlignmentDecimal                            =	3		# Decimal tab alignment
    pbTabAlignmentLeading                            =	0		# Leading tab alignment
    pbTabAlignmentTrailing                           =	2		# Trailing tab alignment

    #values from publisher.pbtableadertype
    #****************************
    pbTabLeaderBullet                                =	5		# Leader bullet tab
    pbTabLeaderDashes                                =	2		# Leader dashes tab
    pbTabLeaderDot                                   =	1		# Leader dot tab
    pbTabLeaderLine                                  =	3		# Tab leader line
    pbTabLeaderNone                                  =	0		# Tab leader none

    #values from publisher.pbtableautoformattype
    #****************************
    pbTableAutoFormatCheckbookRegister               =	0		# Checkbook register
    pbTableAutoFormatCheckerboard                    =	20		# Checkerboard
    pbTableAutoFormatDefault                         =	-3		# Default
    pbTableAutoFormatList1                           =	1		# AutoFormatList1
    pbTableAutoFormatList2                           =	2		# AutoFormatList2
    pbTableAutoFormatList3                           =	3		# AutoFormatList3
    pbTableAutoFormatList4                           =	4		# AutoFormatList4
    pbTableAutoFormatList5                           =	5		# AutoFormatList5
    pbTableAutoFormatList6                           =	6		# AutoFormatList6
    pbTableAutoFormatList7                           =	7		# AutoFormatList7
    pbTableAutoFormatListWithTitle1                  =	8		# Auto Format List with Title1
    pbTableAutoFormatListWithTitle2                  =	9		# Auto Format List with Title2
    pbTableAutoFormatListWithTitle3                  =	10		# Auto Format List with Title3
    pbTableAutoFormatMixed                           =	-1		# Auto Format mixed
    pbTableAutoFormatNone                            =	-2		# Auto Format none
    pbTableAutoFormatNumbers1                        =	11		# Auto Format Numbers1
    pbTableAutoFormatNumbers2                        =	12		# Auto Format Numbers2
    pbTableAutoFormatNumbers3                        =	13		# Auto Format Numbers3
    pbTableAutoFormatNumbers4                        =	14		# Auto Format Numbers4
    pbTableAutoFormatNumbers5                        =	15		# Auto Format Numbers5
    pbTableAutoFormatNumbers6                        =	16		# Auto Format Numbers6
    pbTableAutoFormatTableOfContents1                =	17		# Auto Format Table Of Contents1
    pbTableAutoFormatTableOfContents2                =	18		# Auto Format Table Of Contents2
    pbTableAutoFormatTableOfContents3                =	19		# Auto Format Table Of Contents3

    #values from publisher.pbtabledirectiontype
    #****************************
    pbTableDirectionLeftToRight                      =	1		# Left to Right
    pbTableDirectionRightToLeft                      =	2		# Right to Left

    #values from publisher.pbtextautofittype
    #****************************
    pbTextAutoFitBestFit                             =	2		# Text frame size adjusts to fit text.
    pbTextAutoFitNone                                =	0		# Allows text to overflow the text frame.
    pbTextAutoFitShrinkOnOverflow                    =	1		# Text font reduces so text fits within the text frame.

    #values from publisher.pbtextdirection
    #****************************
    pbTextDirectionLeftToRight                       =	1		# Text flows from left to right.
    pbTextDirectionMixed                             =	-9999999		# Return value indicating a range containing some left-to-right text and some right-to-left text.
    pbTextDirectionRightToLeft                       =	2		# Text flows from right to left.

    #values from publisher.pbtextorientation
    #****************************
    pbTextOrientationHorizontal                      =	1		# Horizontal text orientation
    pbTextOrientationMixed                           =	-2		# Mixed text orientation
    pbTextOrientationRightToLeft                     =	256		# RightToLeft text orientation
    pbTextOrientationVerticalEastAsia                =	2		# VerticalEastAsia text orientation

    #values from publisher.pbtextunit
    #****************************
    pbTextUnitCell                                   =	12		# Expand by a cell
    pbTextUnitCharacter                              =	1		# Expand by a character
    pbTextUnitCharFormat                             =	13		# Expand by a CharFormat
    pbTextUnitCodePoint                              =	17		# Expand by a code point
    pbTextUnitColumn                                 =	9		# Expand by a unit column
    pbTextUnitLine                                   =	5		# Expand by a unit line
    pbTextUnitObject                                 =	16		# Expand by an object
    pbTextUnitParaFormat                             =	14		# Expand by a ParaFormat
    pbTextUnitParagraph                              =	4		# Expand by a paragraph
    pbTextUnitRow                                    =	10		# Expand by a row
    pbTextUnitScreen                                 =	7		# Expand by a screen
    pbTextUnitSection                                =	8		# Expand by a section
    pbTextUnitSentence                               =	3		# Expand by a sentence
    pbTextUnitStory                                  =	6		# Expand by a story
    pbTextUnitTable                                  =	15		# Expand by a table
    pbTextUnitWindow                                 =	11		# Expand by a window
    pbTextUnitWord                                   =	2		# Expand by a word

    #values from publisher.pbtrackingpresettype
    #****************************
    pbTrackingCustom                                 =	-1		# Custom
    pbTrackingLoose                                  =	1		# Loose
    pbTrackingMixed                                  =	-2		# Mixed
    pbTrackingNormal                                 =	2		# Normal
    pbTrackingTight                                  =	3		# Tight
    pbTrackingVeryLoose                              =	0		# Very Loose
    pbTrackingVeryTight                              =	4		# Very Tight

    #values from publisher.pbunderlinetype
    #****************************
    pbUnderlineDash                                  =	6		# Dash
    pbUnderlineDashHeavy                             =	12		# Dash Heavy
    pbUnderlineDashLong                              =	15		# Dash Long
    pbUnderlineDashLongHeavy                         =	16		# Dash Long Heavy
    pbUnderlineDotDash                               =	7		# Dot Dash
    pbUnderlineDotDashHeavy                          =	13		# Dot Dash Heavy
    pbUnderlineDotDotDash                            =	8		# Dot Dot Dash
    pbUnderlineDotDotDashHeavy                       =	14		# Dot Dot Dash Heavy
    pbUnderlineDotHeavy                              =	11		# Dot Heavy
    pbUnderlineDotted                                =	4		# Dotted
    pbUnderlineDouble                                =	3		# Double
    pbUnderlineMixed                                 =	-1		# Mixed
    pbUnderlineNone                                  =	0		# None
    pbUnderlineSingle                                =	1		# Single
    pbUnderlineThick                                 =	5		# Thick
    pbUnderlineWavy                                  =	9		# Wavy
    pbUnderlineWavyDouble                            =	17		# Wavy Double
    pbUnderlineWavyHeavy                             =	10		# Wavy Heavy
    pbUnderlineWordsOnly                             =	2		# Words Only

    #values from publisher.pbunittype
    #****************************
    pbUnitCM                                         =	1		# Sets the unit of measurement to centimeters.
    pbUnitEmu                                        =	4		# Sets the unit of measurement to Emu.
    pbUnitFeet                                       =	6		# Sets the unit of measurement to feet.
    pbUnitHa                                         =	9		# Sets the unit of measurement to Ha.
    pbUnitInch                                       =	0		# Sets the unit of measurement to inches.
    pbUnitKyu                                        =	8		# Sets the unit of measurement to Kyu.
    pbUnitMeter                                      =	7		# Sets the unit of measurement to meters.
    pbUnitPica                                       =	2		# Sets the unit of measurement to picas.
    pbUnitPixel                                      =	10		# Sets the unit of measurement to pixels.
    pbUnitPoint                                      =	3		# Sets the unit of measurement to points.
    pbUnitTwip                                       =	5		# Sets the unit of measurement to twip.

    #values from publisher.pbverticalpicturelocking
    #****************************
    pbVerticalLockingBottom                          =	2		# New pictures are inserted along the bottom edge of the frame.
    pbVerticalLockingNone                            =	0		# New pictures are inserted in the center between the top and bottom edges of the frame.
    pbVerticalLockingStretch                         =	3		# New pictures are vertically stretched to the full height of the frame.
    pbVerticalLockingTop                             =	1		# New pictures are inserted along the top edge of the frame.

    #values from publisher.pbverticaltextalignmenttype
    #****************************
    pbVerticalTextAlignmentBottom                    =	2		# Text is aligned to the bottom.
    pbVerticalTextAlignmentCenter                    =	1		# Text is aligned to the center.
    pbVerticalTextAlignmentTop                       =	0		# Text is aligned to the top.

    #values from publisher.pbwebcontroltype
    #****************************
    pbWebControlCheckBox                             =	100		# Adds a check box.
    pbWebControlCommandButton                        =	101		# Adds a command button.
    pbWebControlHotSpot                              =	110		# Adds a hot spot.
    pbWebControlHTMLFragment                         =	107		# Adds an HTML fragment.
    pbWebControlListBox                              =	102		# Adds a list box.
    pbWebControlMultiLineTextBox                     =	103		# Adds a multiple-line text area.
    pbWebControlOptionButton                         =	104		# Adds an option button.
    pbWebControlSingleLineTextBox                    =	105		# Adds a single-line text box.
    pbWebControlWebComponent                         =	106		# Adds a single-line Web component.

    #values from publisher.pbwindowstate
    #****************************
    pbWindowStateMaximize                            =	0		# Window is maximized.
    pbWindowStateMinimize                            =	1		# Window is minimized.
    pbWindowStateNormal                              =	2		# Window is neither maximized nor minimized.

    #values from publisher.pbwizard
    #****************************
    pbWizardAdvertisements                           =	12		# Creates advertisements
    pbWizardAirplanes                                =	23		# Creates airplanes
    pbWizardBanners                                  =	21		# Creates banners
    pbWizardBrochures                                =	8		# Creates brochures
    pbWizardBusinessCards                            =	3		# Creates business cards
    pbWizardBusinessForms                            =	20		# Creates business forms
    pbWizardCalendars                                =	13		# Creates calendars
    pbWizardCatalogs                                 =	161		# Creates catalogs
    pbWizardCertificates                             =	62		# Creates certificates
    pbWizardEmailActivityEvent                       =	302		# Creates email activity events
    pbWizardEmailAutomatic                           =	305		# Creates email messages automatically
    pbWizardEmailFeaturedProduct                     =	304		# Creates email messages for featured products
    pbWizardEmailLetter                              =	300		# Creates email letters
    pbWizardEmailNewsletter                          =	39		# Creates email newsletters
    pbWizardEmailProductList                         =	303		# Creates email product lists
    pbWizardEmailSpeakerEvent                        =	301		# Creates email speaker event
    pbWizardEnvelopes                                =	7		# Creates envelopes
    pbWizardFlyers                                   =	16		# Creates flyers
    pbWizardGiftCertificates                         =	63		# Creates gift certificates
    pbWizardGreetingCard                             =	40		# Creates greeting cards
    pbWizardInvitation                               =	41		# Creates invitations
    pbWizardJapaneseAdvertisements                   =	165		# Creates Japanese advertisements
    pbWizardJapaneseAirplanes                        =	164		# Creates Japanese airplanes
    pbWizardJapaneseBanners                          =	121		# Creates Japanese banners
    pbWizardJapaneseBrochures                        =	92		# Creates Japanese brochures
    pbWizardJapaneseBusinessCards                    =	91		# Creates Japanese business cards
    pbWizardJapaneseBusinessForms                    =	123		# Creates Japanese business forms
    pbWizardJapaneseCalendars                        =	82		# Creates Japanese calendars
    pbWizardJapaneseCatalogs                         =	177		# Creates Japanese catalogs
    pbWizardJapaneseCertificates                     =	119		# Creates Japanese certificates
    pbWizardJapaneseEnvelopes                        =	93		# Creates Japanese envelopes
    pbWizardJapaneseFlyers                           =	94		# Creates Japanese flyers
    pbWizardJapaneseGiftCertificates                 =	122		# Creates Japanese gift certificates
    pbWizardJapaneseGreetingCards                    =	80		# Creates Japanese greeting cards
    pbWizardJapaneseInvitations                      =	81		# Creates Japanese invitations
    pbWizardJapaneseLabels                           =	118		# Creates Japanese labels
    pbWizardJapaneseLetterheads                      =	95		# Creates Japanese letterheads
    pbWizardJapaneseMenus                            =	116		# Creates Japanese menus
    pbWizardJapaneseNewsletters                      =	117		# Creates Japanese newsletters
    pbWizardJapaneseOrigami                          =	163		# Creates Japanese Origami
    pbWizardJapanesePostcards                        =	78		# Creates Japanese postcards
    pbWizardJapanesePrograms                         =	115		# Creates Japanese programs
    pbWizardJapaneseSigns                            =	149		# Creates Japanese signs
    pbWizardJapaneseWebSites                         =	120		# Creates Japanese Web sites
    pbWizardLabels                                   =	19		# Creates labels
    pbWizardLetterheads                              =	6		# Creates letterheads
    pbWizardMenus                                    =	59		# Creates menus
    pbWizardNewsletters                              =	9		# Creates newsletters
    pbWizardNone                                     =	0		# Default
    pbWizardOrigami                                  =	22		# Creates Origami
    pbWizardPostcards                                =	10		# Creates postcards
    pbWizardPrograms                                 =	76		# Creates programs
    pbWizardQuickPublications                        =	179		# Creates QuickPublications
    pbWizardResumes                                  =	18		# Creates resumes
    pbWizardSigns                                    =	17		# Creates signs
    pbWizardWebSiteBlank                             =	203		# Creates a blank Web site
    pbWizardWebSiteHomePage                          =	5		# Creates a home page for a Web site
    pbWizardWebSiteProductSales                      =	201		# Creates a product sales Web site
    pbWizardWebSiteServices                          =	202		# Creates a services Web site
    pbWizardWebSiteThreePage                         =	200		# Creates a three-page Web site
    pbWizardWithComplimentsCards                     =	73		# Creates with compliments cards
    pbWizardWordDocument                             =	189		# Creates a Microsoft Office Word document

    #values from publisher.pbwizardgroup
    #****************************
    pbWizardGroupAccentBox                           =	151		# Accent Box
    pbWizardGroupAccessoryBar                        =	154		# Accessory Bar
    pbWizardGroupAdvertisements                      =	68		# Advertisements
    pbWizardGroupAttentionGetter                     =	61		# Attention Getter
    pbWizardGroupBarbells                            =	52		# Barbells
    pbWizardGroupBorders                             =	155		# Borders
    pbWizardGroupBoxes                               =	50		# Boxes
    pbWizardGroupCalendars                           =	77		# Calendars
    pbWizardGroupCheckerboards                       =	53		# Checkerboards
    pbWizardGroupCoupon                              =	60		# Coupon
    pbWizardGroupDots                                =	49		# Dots
    pbWizardGroupEastAsiaZipCode                     =	181		# EastAsia ZipCode
    pbWizardGroupJapaneseAccentBox                   =	168		# Japanese Accent Box
    pbWizardGroupJapaneseAccessoryBar                =	171		# Japanese Accessory Bar
    pbWizardGroupJapaneseAttentionGetters            =	97		# Japanese Attention Getters
    pbWizardGroupJapaneseBorders                     =	172		# Japanese Borders
    pbWizardGroupJapaneseCalendar                    =	83		# Japanese Calendar
    pbWizardGroupJapaneseCoupons                     =	99		# Japanese Coupons
    pbWizardGroupJapaneseLinearAccent                =	170		# Japanese Linear Accent
    pbWizardGroupJapaneseMarquees                    =	167		# Japanese Marquees
    pbWizardGroupJapaneseMastheads                   =	141		# Japanese Mastheads
    pbWizardGroupJapanesePullQuotes                  =	144		# Japanese Pull Quotes
    pbWizardGroupJapaneseReplyForms                  =	137		# Japanese Reply Forms
    pbWizardGroupJapaneseSidebars                    =	143		# Japanese Sidebars
    pbWizardGroupJapaneseTableOfContents             =	142		# Japanese Table Of Contents
    pbWizardGroupJapaneseWebButtonEmail              =	182		# Japanese Web Button Email
    pbWizardGroupJapaneseWebButtonHome               =	183		# Japanese Web Button Home
    pbWizardGroupJapaneseWebButtonLink               =	184		# Japanese Web Button Link
    pbWizardGroupJapaneseWebMastheads                =	138		# Japanese Web Mastheads
    pbWizardGroupJapaneseWebNavigationBars           =	148		# Japanese Web Navigation Bars
    pbWizardGroupJapaneseWebPullQuotes               =	139		# Japanese Web Pull Quotes
    pbWizardGroupJapaneseWebSidebars                 =	140		# Japanese Web Sidebars
    pbWizardGroupLinearAccent                        =	153		# Linear Accent
    pbWizardGroupLogo                                =	4		# Logo
    pbWizardGroupMarquee                             =	150		# Marquee
    pbWizardGroupMastheads                           =	105		# Mastheads
    pbWizardGroupPhoneTearoff                        =	66		# Phone Tearoff
    pbWizardGroupPictureCaptions                     =	109		# Picture Captions
    pbWizardGroupPullQuotes                          =	108		# Pull Quotes
    pbWizardGroupPunctuation                         =	152		# Punctuation
    pbWizardGroupReplyForms                          =	79		# Reply Forms
    pbWizardGroupSidebars                            =	107		# Sidebars
    pbWizardGroupTableOfContents                     =	106		# Table of Contents
    pbWizardGroupWebButtonsEmail                     =	133		# Web Buttons Email
    pbWizardGroupWebButtonsHome                      =	134		# Web Buttons Home
    pbWizardGroupWebButtonsLink                      =	136		# Web Buttons Link
    pbWizardGroupWebCalendars                        =	35		# Web Calendars
    pbWizardGroupWebMastheads                        =	102		# Web Mastheads
    pbWizardGroupWebNavigationBars                   =	75		# Web Navigation Bars
    pbWizardGroupWebSidebars                         =	104		# Group Web Sidebars
    pbWizardGroupWellPullQuotes                      =	103		# Group Well Pull Quotes

    #values from publisher.pbwizardnavbaralignment
    #****************************
    pbnbAlignCenter                                  =	2		# Center-aligned
    pbnbAlignLeft                                    =	1		# Left-aligned
    pbnbAlignRight                                   =	3		# Right-aligned

    #values from publisher.pbwizardnavbarbuttonstyle
    #****************************
    pbnbButtonStyleLarge                             =	2		# Large buttons
    pbnbButtonStyleSmall                             =	1		# Small buttons
    pbnbButtonStyleText                              =	3		# Text-only buttons

    #values from publisher.pbwizardnavbardesign
    #****************************
    pbnbDesignAmbient                                =	2		# Ambient
    pbnbDesignBaseline                               =	26		# Baseline
    pbnbDesignBracket                                =	11		# Bracket
    pbnbDesignBulletStaff                            =	20		# BulletStaff
    pbnbDesignCapsule                                =	3		# Capsule
    pbnbDesignCornice                                =	15		# Cornice
    pbnbDesignCounter                                =	13		# Counter
    pbnbDesignDimension                              =	8		# Dimension
    pbnbDesignDottedArrow                            =	9		# Dotted Arrow
    pbnbDesignEdge                                   =	17		# Edge
    pbnbDesignEnclosedArrow                          =	12		# Enclosed Arrow
    pbnbDesignEndCap                                 =	14		# End Cap
    pbnbDesignHollowArrow                            =	10		# Hollow Arrow
    pbnbDesignKeyPunch                               =	22		# Key Punch
    pbnbDesignOffset                                 =	7		# Offset
    pbnbDesignOutline                                =	5		# Outline
    pbnbDesignRadius                                 =	6		# Radius
    pbnbDesignRectangle                              =	1		# Rectangle
    pbnbDesignRoundBullet                            =	23		# RoundBullet
    pbnbDesignSquareBullet                           =	24		# SquareBullet
    pbnbDesignStaff                                  =	16		# Staff
    pbnbDesignTopBar                                 =	21		# TopBar
    pbnbDesignTopDrawer                              =	4		# TopDrawer
    pbnbDesignTopLine                                =	18		# TopLine
    pbnbDesignUnderscore                             =	19		# Underscore
    pbnbDesignWatermark                              =	25		# Watermark

    #values from publisher.pbwizardpagetype
    #****************************
    pbWizardPageTypeCatalogBlank                     =	35		# CatalogBlank
    pbWizardPageTypeCatalogCalendar                  =	22		# Calendar
    pbWizardPageTypeCatalogEightItemsOneColumn       =	33		# CatalogEightItemsOneColumn
    pbWizardPageTypeCatalogEightItemsTwoColumns      =	34		# CatalogEightItemsTwoColumns
    pbWizardPageTypeCatalogFeaturedItem              =	24		# CatalogFeaturedItem
    pbWizardPageTypeCatalogForm                      =	36		# Catalog Form
    pbWizardPageTypeCatalogFourItemsAlignedPictures  =	30		# Catalog Four Items Aligned Pictures
    pbWizardPageTypeCatalogFourItemsOffsetPictures   =	31		# Catalog Four Items Offset Pictures
    pbWizardPageTypeCatalogFourItemsSquaredPictures  =	32		# Catalog Four Items Squared Pictures
    pbWizardPageTypeCatalogOneColumnText             =	18		# Catalog One Column Text
    pbWizardPageTypeCatalogOneColumnTextPicture      =	19		# Catalog One Column Text Picture
    pbWizardPageTypeCatalogTableOfContents           =	23		# Catalog Table of Contents
    pbWizardPageTypeCatalogThreeItemsAlignedPictures	=	27		# Three Items Aligned Pictures
    pbWizardPageTypeCatalogThreeItemsOffsetPictures  =	28		# Three Items Offset Pictures
    pbWizardPageTypeCatalogThreeItemsStackedPictures	=	29		# Three Items Stacked Pictures
    pbWizardPageTypeCatalogTwoColumnsText            =	20		# Catalog Two Columns Text
    pbWizardPageTypeCatalogTwoColumnsTextPicture     =	21		# Catalog Two Columns Text Picture
    pbWizardPageTypeCatalogTwoItemsAlignedPictures   =	25		# Two Items Aligned Pictures
    pbWizardPageTypeCatalogTwoItemsOffsetPictures    =	26		# Two Items Offset Pictures
    pbWizardPageTypeNewsletter3Stories               =	1		# Newsletter3 Stories
    pbWizardPageTypeNewsletterCalendar               =	2		# Newsletter Calendar
    pbWizardPageTypeNewsletterOrderForm              =	15		# Newsletter OrderForm
    pbWizardPageTypeNewsletterResponseForm           =	16		# Newsletter Response Form
    pbWizardPageTypeNewsletterSignupForm             =	17		# Newsletter Signup Form
    pbWizardPageTypeNone                             =	-1		# None
    pbWizardPageTypeWebAboutUs                       =	501		# Web About Us
    pbWizardPageTypeWebArticle                       =	512		# Web Article
    pbWizardPageTypeWebBlank                         =	524		# Web Blank
    pbWizardPageTypeWebCalendarPage                  =	504		# Web Calendar Page
    pbWizardPageTypeWebCalendarWithLinks             =	800		# Web Calendar With Links
    pbWizardPageTypeWebContactUs                     =	505		# Web Contact Us
    pbWizardPageTypeWebEmployee                      =	507		# Web Employee
    pbWizardPageTypeWebEmployeeList                  =	506		# Web Employee List
    pbWizardPageTypeWebEmployeesWithLinks            =	802		# Web Employees With Links
    pbWizardPageTypeWebFAQ                           =	508		# Web FAQ
    pbWizardPageTypeWebHome                          =	509		# Web Home
    pbWizardPageTypeWebInformational                 =	502		# Web Informational
    pbWizardPageTypeWebJobs                          =	510		# Web Jobs
    pbWizardPageTypeWebLegal                         =	511		# Web Legal
    pbWizardPageTypeWebLinks                         =	518		# Web Links
    pbWizardPageTypeWebList                          =	503		# Web List
    pbWizardPageTypeWebOrderForm                     =	525		# Web Order Form
    pbWizardPageTypeWebPhoto                         =	513		# Web Photo
    pbWizardPageTypeWebPhotoGallery                  =	514		# Web Photo Gallery
    pbWizardPageTypeWebPhotosWithLinks               =	805		# Web Photos With Links
    pbWizardPageTypeWebProduct                       =	515		# Web Product
    pbWizardPageTypeWebProductList                   =	516		# Web Product List
    pbWizardPageTypeWebProductsWithLinks             =	801		# Web Products With Links
    pbWizardPageTypeWebProjectList                   =	517		# Web Project List
    pbWizardPageTypeWebProjectsWithLinks             =	804		# Web Projects With Links
    pbWizardPageTypeWebResponseForm                  =	526		# Web Response Form
    pbWizardPageTypeWebSeminar                       =	519		# Web Seminar
    pbWizardPageTypeWebService                       =	521		# Web Service
    pbWizardPageTypeWebServiceList                   =	520		# Web Service List
    pbWizardPageTypeWebServicesWithLinks             =	803		# Web Services With Links
    pbWizardPageTypeWebSignupForm                    =	527		# Web Signup Form
    pbWizardPageTypeWebSpecial                       =	522		# Web Special

    #values from publisher.pbwizardtag
    #****************************
    pbWizardTagAddress                               =	10		# Address
    pbWizardTagAddressGroup                          =	117		# Address Group
    pbWizardTagBriefDescriptionCaption               =	1361		# Description Caption
    pbWizardTagBriefDescriptionGraphic               =	1359		# Description Graphic
    pbWizardTagBriefDescriptionSummary               =	1353		# Description Summary
    pbWizardTagBriefDescriptionSummaryPrimary        =	1365		# Description Summary Primary
    pbWizardTagBriefDescriptionTitle                 =	1364		# Brief Description Title
    pbWizardTagBusinessDescription                   =	685		# Business Description
    pbWizardTagCustomerMailingAddress                =	560		# Customer Mailing Address
    pbWizardTagDate                                  =	1835		# Tag Date
    pbWizardTagEAPostalCodeBox                       =	2151		# EA Postal Code Box
    pbWizardTagEAPostalCodeGroup                     =	2150		# EA Postal Code Group
    pbWizardTagEAPostalCodeLine                      =	2152		# EA Postal Code Line
    pbWizardTagFloatingGraphicCaption                =	1362		# Floating Graphic Caption
    pbWizardTagHourTimeDateInformation               =	684		# Hour Time Date Information
    pbWizardTagJobTitle                              =	115		# Job Title
    pbWizardTagLinkedStoryPrimary                    =	1354		# Linked Primary Story
    pbWizardTagLinkedStorySecondary                  =	1355		# Secondary Story
    pbWizardTagLinkedStoryTertiary                   =	1356		# Linked Story Tertiary
    pbWizardTagList                                  =	1837		# List
    pbWizardTagLocation                              =	488		# Location
    pbWizardTagLogoGroup                             =	5		# Logo Group
    pbWizardTagMainFloatingGraphic                   =	1357		# Main Floating Graphic
    pbWizardTagMainGraphic                           =	1833		# Main Graphic
    pbWizardTagMainTitle                             =	1832		# Title
    pbWizardTagMapPicture                            =	489		# Picture
    pbWizardTagMasthead                              =	1831		# Masthead
    pbWizardTagNewsletterTitle                       =	1344		# Newsletter Title
    pbWizardTagOrganizationName                      =	7		# Organization Name
    pbWizardTagOrganizationNameGroup                 =	118		# Organization Name Group
    pbWizardTagPageNumber                            =	1346		# Page Number
    pbWizardTagPersonalName                          =	8		# Personal Name
    pbWizardTagPersonalNameGroup                     =	116		# Personal Name Group
    pbWizardTagPhoneFaxEmail                         =	113		# Phone/Fax/Email
    pbWizardTagPhoneFaxEmailGroup                    =	120		# Phone/Fax/Email Group
    pbWizardTagPhoneNumber                           =	114		# Phone Number
    pbWizardTagPhotoPlaceholderFrame                 =	1134		# Photo Placeholder Frame
    pbWizardTagPhotoPlaceholderText                  =	1135		# Photo Placeholder Text
    pbWizardTagPublicationDate                       =	1341		# Publication Date
    pbWizardTagQuickPubContent                       =	2143		# Quick Pub Content
    pbWizardTagQuickPubHeading                       =	2140		# Quick Pub Heading
    pbWizardTagQuickPubMessage                       =	2141		# Quick Pub Message
    pbWizardTagQuickPubPicture                       =	2142		# Quick Pub Picture
    pbWizardTagReturnAddressLines                    =	793		# Return Address Lines
    pbWizardTagStampBox                              =	887		# Stamp Box
    pbWizardTagStampBoxOutline                       =	794		# Stamp Box Outline
    pbWizardTagStory                                 =	1349		# Story
    pbWizardTagStoryCaptionPrimary                   =	1351		# Caption Primary
    pbWizardTagStoryCaptionSecondary                 =	1373		# Caption Secondary
    pbWizardTagStoryGraphicPrimary                   =	1350		# Graphic Primary
    pbWizardTagStoryGraphicSecondary                 =	1360		# Graphic Secondary
    pbWizardTagStoryTitle                            =	1348		# Story Title
    pbWizardTagTableOfContents                       =	1343		# Table of Contents
    pbWizardTagTableOfContentsTitle                  =	1342		# Table of Contents Title
    pbWizardTagTagLine                               =	112		# Tag Line
    pbWizardTagTagLineGroup                          =	119		# Tag Line Group
    pbWizardTagTime                                  =	1836		# Tag Time

    #values from publisher.pbwrapsidetype
    #****************************
    pbWrapSideBoth                                   =	0		# Wrap both sides of the shape
    pbWrapSideLarger                                 =	3		# Wrap the larger side of the shape
    pbWrapSideLeft                                   =	1		# Wrap the left side of the shape
    pbWrapSideMixed                                  =	-1		# Wrap the shape in different proportions
    pbWrapSideNeither                                =	4		# Does not wrap the shape on the sides
    pbWrapSideRight                                  =	2		# Wrap the right side of the shape

    #values from publisher.pbwraptype
    #****************************
    pbWrapTypeMixed                                  =	-1		# Mixed
    pbWrapTypeNone                                   =	0		# None
    pbWrapTypeSquare                                 =	1		# Square
    pbWrapTypeThrough                                =	3		# Through
    pbWrapTypeTight                                  =	2		# Tight
    pbWrapTypeTopAndBottom                           =	4		# Top and Bottom

    #values from publisher.pbzoom
    #****************************
    pbZoomFitSelection                               =	-3		# Resizes the page view to the size of the current selection.
    pbZoomPageWidth                                  =	-1		# Resizes the page view to the width of the publication.
    pbZoomWholePage                                  =	-2		# Resizes the page view to the size of a whole page.
}
#End Enum

$pb = new-object PSCustomObject -Property $pb


#constants for Word based on https://docs.microsoft.com/en-us/office/vba/api/Word(enumerations)
$wd=[Ordered]@{

#values from word.wdalertlevel
#****************************
wdAlertsAll	=	-1		# All message boxes and alerts are displayed; errors are returned to the macro.
wdAlertsMessageBox	=	-2		# Only message boxes are displayed; errors are trapped and returned to the macro.
wdAlertsNone	=	0		# No alerts or message boxes are displayed. If a macro encounters a message box, the default value is chosen and the macro continues.

#values from word.wdalignmenttabalignment
#****************************
wdCenter	=	1		# Centered tab.
wdLeft	=	0		# Left-aligned tab.
wdRight	=	2		# Right-aligned tab.

#values from word.wdalignmenttabrelative
#****************************
wdIndent	=	1		# Word calculates tab alignment relative to the paragraph indents.
wdMargin	=	0		# Word calculates tab alignment relative to the margins

#values from word.wdapplyquickstylesets
#****************************
wdSessionStartSet	=	1		# Resets the Quick Style to the style set in use when the document was opened.
wdTemplateSet	=	2		# Resets the Quick Style to the style set from the template, if any.

#values from word.wdarabicnumeral
#****************************
wdNumeralArabic	=	0		# Arabic shape is used for numerals.
wdNumeralContext	=	2		# Numeral shape depends on text surrounding it.
wdNumeralHindi	=	1		# Hindi shape is used for numerals.
wdNumeralSystem	=	3		# Numeral shape is determined by system settings.

#values from word.wdaraspeller
#****************************
wdBoth	=	3		# The spelling checker uses spelling rules regarding both Arabic words ending with the letter yaa and Arabic words beginning with an alef hamza.
wdFinalYaa	=	2		# The spelling checker uses spelling rules regarding Arabic words ending with the letter yaa.
wdInitialAlef	=	1		# The spelling checker uses spelling rules regarding Arabic words beginning with an alef hamza.
wdNone	=	0		# The spelling checker ignores spelling rules regarding either Arabic words ending with the letter yaa or Arabic words beginning with an alef hamza.

#values from word.wdarrangestyle
#****************************
wdIcons	=	1		# Windows are displayed as icons in a single window.
wdTiled	=	0		# Windows are tiled into a single window.

#values from word.wdautofitbehavior
#****************************
wdAutoFitContent	=	1		# The table is automatically sized to fit the content contained in the table.
wdAutoFitFixed	=	0		# The table is set to a fixed size, regardless of the content, and is not automatically sized.
wdAutoFitWindow	=	2		# The table is automatically sized to the width of the active window.

#values from word.wdautomacros
#****************************
wdAutoClose	=	3		# AutoClose macro.
wdAutoExec	=	0		# AutoExec macro.
wdAutoExit	=	4		# AutoExit macro.
wdAutoNew	=	1		# AutoNew macro.
wdAutoOpen	=	2		# AutoOpen macro.
wdAutoSync	=	5		# AutoSync macro.

#values from word.wdautoversions
#****************************
wdAutoVersionOff	=	0		# No document version is saved.
wdAutoVersionOnClose	=	1		# A document version is saved automatically when the document is closed.

#values from word.wdbaselinealignment
#****************************
wdBaselineAlignAuto	=	4		# Microsoft Word automatically adjusts the baseline font alignment.
wdBaselineAlignBaseline	=	2		# Align to a baseline for the paragraph.
wdBaselineAlignCenter	=	1		# Align center points of each font.
wdBaselineAlignFarEast50	=	3		# Align using the baseline for Asian language font standards.
wdBaselineAlignTop	=	0		# Align along top of each font.

#values from word.wdbookmarksortby
#****************************
wdSortByLocation	=	1		# Sorted by location in document.
wdSortByName	=	0		# Sorted by bookmark name.

#values from word.wdborderdistancefrom
#****************************
wdBorderDistanceFromPageEdge	=	1		# From the edge of the page.
wdBorderDistanceFromText	=	0		# From the text it surrounds.

#values from word.wdbordertype
#****************************
wdBorderBottom	=	-3		# A bottom border.
wdBorderDiagonalDown	=	-7		# A diagonal border starting in the upper-left corner.
wdBorderDiagonalUp	=	-8		# A diagonal border starting in the lower-left corner.
wdBorderHorizontal	=	-5		# Horizontal borders.
wdBorderLeft	=	-2		# A left border.
wdBorderRight	=	-4		# A right border.
wdBorderTop	=	-1		# A top border.
wdBorderVertical	=	-6		# Vertical borders.

#values from word.wdbreaktype
#****************************
wdColumnBreak	=	8		# Column break at the insertion point.
wdLineBreak	=	6		# Line break.
wdLineBreakClearLeft	=	9		# Line break.
wdLineBreakClearRight	=	10		# Line break.
wdPageBreak	=	7		# Page break at the insertion point.
wdSectionBreakContinuous	=	3		# New section without a corresponding page break.
wdSectionBreakEvenPage	=	4		# Section break with the next section beginning on the next even-numbered page. If the section break falls on an even-numbered page, Word leaves the next odd-numbered page blank.
wdSectionBreakNextPage	=	2		# Section break on next page.
wdSectionBreakOddPage	=	5		# Section break with the next section beginning on the next odd-numbered page. If the section break falls on an odd-numbered page, Word leaves the next even-numbered page blank.
wdTextWrappingBreak	=	11		# Ends the current line and forces the text to continue below a picture, table, or other item. The text continues on the next blank line that does not contain a table aligned with the left or right margin.

#values from word.wdbrowserlevel
#****************************
wdBrowserLevelMicrosoftInternetExplorer5	=	1		# Microsoft Internet Explorer 5.
wdBrowserLevelMicrosoftInternetExplorer6	=	2		# Microsoft Internet Explorer 6.
wdBrowserLevelV4	=	0		# Microsoft Internet Explorer 4.

#values from word.wdbrowsetarget
#****************************
wdBrowseComment	=	3		# Places insertion point before next or previous comment.
wdBrowseEdit	=	10		# Places insertion point before next or previous edit.
wdBrowseEndnote	=	5		# Places insertion point before next or previous endnote.
wdBrowseField	=	6		# Places insertion point before next or previous browsefield.
wdBrowseFind	=	11		# Places insertion point before next or previous browsefind.
wdBrowseFootnote	=	4		# Places insertion point before next or previous footnote.
wdBrowseGoTo	=	12		# Places insertion point before next or previous GoTo item.
wdBrowseGraphic	=	8		# Places insertion point before next or previous graphic.
wdBrowseHeading	=	9		# Places insertion point before next or previous heading.
wdBrowsePage	=	1		# Places insertion point before next or previous page.
wdBrowseSection	=	2		# Places insertion point before next or previous section.
wdBrowseTable	=	7		# Places insertion point before next or previous table.

#values from word.wdbuildingblocktypes
#****************************
wdTypeAutoText	=	9		# Autotext building block.
wdTypeBibliography	=	34		# Bibliography building block.
wdTypeCoverPage	=	2		# Cover page building block.
wdTypeCustom1	=	29		# Custom building block.
wdTypeCustom2	=	30		# Custom building block.
wdTypeCustom3	=	31		# Custom building block.
wdTypeCustom4	=	32		# Custom building block.
wdTypeCustom5	=	33		# Custom building block.
wdTypeCustomAutoText	=	23		# Custom autotext building block.
wdTypeCustomBibliography	=	35		# Custom bibliography building block.
wdTypeCustomCoverPage	=	16		# Custom cover page building block.
wdTypeCustomEquations	=	17		# Custom equations building block.
wdTypeCustomFooters	=	18		# Custom footers building block.
wdTypeCustomHeaders	=	19		# Custom headers building block.
wdTypeCustomPageNumber	=	20		# Custom page numbering building block.
wdTypeCustomPageNumberBottom	=	26		# Building block for custom page numbering on the bottom of the page.
wdTypeCustomPageNumberPage	=	27		# Custom page numbering building block.
wdTypeCustomPageNumberTop	=	25		# Building block for custom page numbering on the top of the page.
wdTypeCustomQuickParts	=	15		# Custom quick parts building block.
wdTypeCustomTableOfContents	=	28		# Custom table of contents building block.
wdTypeCustomTables	=	21		# Custom table building block.
wdTypeCustomTextBox	=	24		# Custom text box building block.
wdTypeCustomWatermarks	=	22		# Custom watermark building block.
wdTypeEquations	=	3		# Equation building block.
wdTypeFooters	=	4		# Footer building block.
wdTypeHeaders	=	5		# Header building block.
wdTypePageNumber	=	6		# Page numbering building block.
wdTypePageNumberBottom	=	12		# Building block for page numbering on the bottom of the page.
wdTypePageNumberPage	=	13		# Page numbering building block.
wdTypePageNumberTop	=	11		# Building block for page numbering on the top of the page.
wdTypeQuickParts	=	1		# Quick parts building block.
wdTypeTableOfContents	=	14		# Table of contents building block.
wdTypeTables	=	7		# Table building block.
wdTypeTextBox	=	10		# Text box building block.
wdTypeWatermarks	=	8		# Watermark building block.

#values from word.wdbuiltinproperty
#****************************
wdPropertyAppName	=	9		# Name of application.
wdPropertyAuthor	=	3		# Author.
wdPropertyBytes	=	22		# Byte count.
wdPropertyCategory	=	18		# Category.
wdPropertyCharacters	=	16		# Character count.
wdPropertyCharsWSpaces	=	30		# Character count with spaces.
wdPropertyComments	=	5		# Comments.
wdPropertyCompany	=	21		# Company.
wdPropertyFormat	=	19		# Not supported.
wdPropertyHiddenSlides	=	27		# Not supported.
wdPropertyHyperlinkBase	=	29		# Not supported.
wdPropertyKeywords	=	4		# Keywords.
wdPropertyLastAuthor	=	7		# Last author.
wdPropertyLines	=	23		# Line count.
wdPropertyManager	=	20		# Manager.
wdPropertyMMClips	=	28		# Not supported.
wdPropertyNotes	=	26		# Notes.
wdPropertyPages	=	14		# Page count.
wdPropertyParas	=	24		# Paragraph count.
wdPropertyRevision	=	8		# Revision number.
wdPropertySecurity	=	17		# Security setting.
wdPropertySlides	=	25		# Not supported.
wdPropertySubject	=	2		# Subject.
wdPropertyTemplate	=	6		# Template name.
wdPropertyTimeCreated	=	11		# Time created.
wdPropertyTimeLastPrinted	=	10		# Time last printed.
wdPropertyTimeLastSaved	=	12		# Time last saved.
wdPropertyTitle	=	1		# Title.
wdPropertyVBATotalEdit	=	13		# Number of edits to VBA project.
wdPropertyWords	=	15		# Word count.

#values from word.wdbuiltinstyle
#****************************
wdStyleBlockQuotation	=	-85		# Block Text.
wdStyleBodyText	=	-67		# Body Text.
wdStyleBodyText2	=	-81		# Body Text 2.
wdStyleBodyText3	=	-82		# Body Text 3.
wdStyleBodyTextFirstIndent	=	-78		# Body Text First Indent.
wdStyleBodyTextFirstIndent2	=	-79		# Body Text First Indent 2.
wdStyleBodyTextIndent	=	-68		# Body Text Indent.
wdStyleBodyTextIndent2	=	-83		# Body Text Indent 2.
wdStyleBodyTextIndent3	=	-84		# Body Text Indent 3.
wdStyleBookTitle	=	-265		# Book Title.
wdStyleCaption	=	-35		# Caption.
wdStyleClosing	=	-64		# Closing.
wdStyleCommentReference	=	-40		# Comment Reference.
wdStyleCommentText	=	-31		# Comment Text.
wdStyleDate	=	-77		# Date.
wdStyleDefaultParagraphFont	=	-66		# Default Paragraph Font.
wdStyleEmphasis	=	-89		# Emphasis.
wdStyleEndnoteReference	=	-43		# Endnote Reference.
wdStyleEndnoteText	=	-44		# Endnote Text.
wdStyleEnvelopeAddress	=	-37		# Envelope Address.
wdStyleEnvelopeReturn	=	-38		# Envelope Return.
wdStyleFooter	=	-33		# Footer.
wdStyleFootnoteReference	=	-39		# Footnote Reference.
wdStyleFootnoteText	=	-30		# Footnote Text.
wdStyleHeader	=	-32		# Header.
wdStyleHeading1	=	-2		# Heading 1.
wdStyleHeading2	=	-3		# Heading 2.
wdStyleHeading3	=	-4		# Heading 3.
wdStyleHeading4	=	-5		# Heading 4.
wdStyleHeading5	=	-6		# Heading 5.
wdStyleHeading6	=	-7		# Heading 6.
wdStyleHeading7	=	-8		# Heading 7.
wdStyleHeading8	=	-9		# Heading 8.
wdStyleHeading9	=	-10		# Heading 9.
wdStyleHtmlAcronym	=	-96		# HTML Acronym.
wdStyleHtmlAddress	=	-97		# HTML Address.
wdStyleHtmlCite	=	-98		# HTML Cite.
wdStyleHtmlCode	=	-99		# HTML Code.
wdStyleHtmlDfn	=	-100		# HTML Definition.
wdStyleHtmlKbd	=	-101		# HTML Keyboard.
wdStyleHtmlNormal	=	-95		# Normal (Web).
wdStyleHtmlPre	=	-102		# HTML Preformatted.
wdStyleHtmlSamp	=	-103		# HTML Sample.
wdStyleHtmlTt	=	-104		# HTML Typewriter.
wdStyleHtmlVar	=	-105		# HTML Variable.
wdStyleHyperlink	=	-86		# Hyperlink.
wdStyleHyperlinkFollowed	=	-87		# Followed Hyperlink.
wdStyleIndex1	=	-11		# Index 1.
wdStyleIndex2	=	-12		# Index 2.
wdStyleIndex3	=	-13		# Index 3.
wdStyleIndex4	=	-14		# Index 4.
wdStyleIndex5	=	-15		# Index 5.
wdStyleIndex6	=	-16		# Index 6.
wdStyleIndex7	=	-17		# Index 7.
wdStyleIndex8	=	-18		# Index 8.
wdStyleIndex9	=	-19		# Index 9.
wdStyleIndexHeading	=	-34		# Index Heading
wdStyleIntenseEmphasis	=	-262		# Intense Emphasis.
wdStyleIntenseQuote	=	-182		# Intense Quote.
wdStyleIntenseReference	=	-264		# Intense Reference.
wdStyleLineNumber	=	-41		# Line Number.
wdStyleList	=	-48		# List.
wdStyleList2	=	-51		# List 2.
wdStyleList3	=	-52		# List 3.
wdStyleList4	=	-53		# List 4.
wdStyleList5	=	-54		# List 5.
wdStyleListBullet	=	-49		# List Bullet.
wdStyleListBullet2	=	-55		# List Bullet 2.
wdStyleListBullet3	=	-56		# List Bullet 3.
wdStyleListBullet4	=	-57		# List Bullet 4.
wdStyleListBullet5	=	-58		# List Bullet 5.
wdStyleListContinue	=	-69		# List Continue.
wdStyleListContinue2	=	-70		# List Continue 2.
wdStyleListContinue3	=	-71		# List Continue 3.
wdStyleListContinue4	=	-72		# List Continue 4.
wdStyleListContinue5	=	-73		# List Continue 5.
wdStyleListNumber	=	-50		# List Number.
wdStyleListNumber2	=	-59		# List Number 2.
wdStyleListNumber3	=	-60		# List Number 3.
wdStyleListNumber4	=	-61		# List Number 4.
wdStyleListNumber5	=	-62		# List Number 5.
wdStyleListParagraph	=	-180		# List Paragraph.
wdStyleMacroText	=	-46		# Macro Text.
wdStyleMessageHeader	=	-74		# Message Header.
wdStyleNavPane	=	-90		# Document Map.
wdStyleNormal	=	-1		# Normal.
wdStyleNormalIndent	=	-29		# Normal Indent.
wdStyleNormalObject	=	-158		# Normal (applied to an object).
wdStyleNormalTable	=	-106		# Normal (applied within a table).
wdStyleNoteHeading	=	-80		# Note Heading.
wdStylePageNumber	=	-42		# Page Number.
wdStylePlainText	=	-91		# Plain Text.
wdStyleQuote	=	-181		# Quote.
wdStyleSalutation	=	-76		# Salutation.
wdStyleSignature	=	-65		# Signature.
wdStyleStrong	=	-88		# Strong.
wdStyleSubtitle	=	-75		# Subtitle.
wdStyleSubtleEmphasis	=	-261		# Subtle Emphasis.
wdStyleSubtleReference	=	-263		# Subtle Reference.
wdStyleTableColorfulGrid	=	-172		# Colorful Grid.
wdStyleTableColorfulList	=	-171		# Colorful List.
wdStyleTableColorfulShading	=	-170		# Colorful Shading.
wdStyleTableDarkList	=	-169		# Dark List.
wdStyleTableLightGrid	=	-161		# Light Grid.
wdStyleTableLightGridAccent1	=	-175		# Light Grid Accent 1.
wdStyleTableLightList	=	-160		# Light List.
wdStyleTableLightListAccent1	=	-174		# Light List Accent 1.
wdStyleTableLightShading	=	-159		# Light Shading.
wdStyleTableLightShadingAccent1	=	-173		# Light Shading Accent 1.
wdStyleTableMediumGrid1	=	-166		# Medium Grid 1.
wdStyleTableMediumGrid2	=	-167		# Medium Grid 2.
wdStyleTableMediumGrid3	=	-168		# Medium Grid 3.
wdStyleTableMediumList1	=	-164		# Medium List 1.
wdStyleTableMediumList1Accent1	=	-178		# Medium List 1 Accent 1.
wdStyleTableMediumList2	=	-165		# Medium List 2.
wdStyleTableMediumShading1	=	-162		# Medium Shading 1.
wdStyleTableMediumShading1Accent1	=	-176		# Medium Shading 1 Accent 1.
wdStyleTableMediumShading2	=	-163		# Medium Shading 2.
wdStyleTableMediumShading2Accent1	=	-177		# Medium Shading 2 Accent 1.
wdStyleTableOfAuthorities	=	-45		# Table of Authorities.
wdStyleTableOfFigures	=	-36		# Table of Figures.
wdStyleTitle	=	-63		# Title.
wdStyleTOAHeading	=	-47		# TOA Heading.
wdStyleTOC1	=	-20		# TOC 1.
wdStyleTOC2	=	-21		# TOC 2.
wdStyleTOC3	=	-22		# TOC 3.
wdStyleTOC4	=	-23		# TOC 4.
wdStyleTOC5	=	-24		# TOC 5.
wdStyleTOC6	=	-25		# TOC 6.
wdStyleTOC7	=	-26		# TOC 7.
wdStyleTOC8	=	-27		# TOC 8.
wdStyleTOC9	=	-28		# TOC 9.

#values from word.wdcalendartype
#****************************
wdCalendarArabic	=	1		# Arabic Hijri calendar.
wdCalendarHebrew	=	2		# Hebrew Lunar calendar.
wdCalendarJapan	=	4		# Japanese Emperor Era calendar.
wdCalendarKorean	=	6		# Korean Danki calendar.
wdCalendarSakaEra	=	7		# Saka Era calendar.
wdCalendarTaiwan	=	3		# Taiwan calendar.
wdCalendarThai	=	5		# Thai calendar.
wdCalendarTranslitEnglish	=	8		# English transliterated. Gregorian calendar with English month and day names transliterated to the Arabic script. Unsupported.
wdCalendarTranslitFrench	=	9		# French transliterated. Gregorian calendar with French month and day names transliterated to the Arabic script. Unsupported.
wdCalendarUmalqura	=	13		# Um-al-Qura calendar.
wdCalendarWestern	=	0		# Western. Corresponds to the Gregorian calendar.

#values from word.wdcalendartypebi
#****************************
wdCalendarTypeBidi	=	99		# Bi-directional calendar.
wdCalendarTypeGregorian	=	100		# Gregorian calendar.

#values from word.wdcaptionlabelid
#****************************
wdCaptionEquation	=	-3		# Equation.
wdCaptionFigure	=	-1		# Figure.
wdCaptionTable	=	-2		# Table.

#values from word.wdcaptionnumberstyle
#****************************
wdCaptionNumberStyleArabic	=	0		# Arabic style.
wdCaptionNumberStyleArabicFullWidth	=	14		# Full-width Arabic style.
wdCaptionNumberStyleArabicLetter1	=	46		# Arabic letter style 1.
wdCaptionNumberStyleArabicLetter2	=	48		# Arabic letter style 2.
wdCaptionNumberStyleChosung	=	25		# Chosung style.
wdCaptionNumberStyleGanada	=	24		# Ganada style.
wdCaptionNumberStyleHanjaRead	=	41		# Hanja read style.
wdCaptionNumberStyleHanjaReadDigit	=	42		# Hanja read digit style.
wdCaptionNumberStyleHebrewLetter1	=	45		# Hebrew letter style 1.
wdCaptionNumberStyleHebrewLetter2	=	47		# Hebrew letter style 2.
wdCaptionNumberStyleHindiArabic	=	51		# Hindi Arabic style.
wdCaptionNumberStyleHindiCardinalText	=	52		# Hindi cardinal style.
wdCaptionNumberStyleHindiLetter1	=	49		# Hindi letter style 1.
wdCaptionNumberStyleHindiLetter2	=	50		# Hindi letter style 2.
wdCaptionNumberStyleKanji	=	10		# Kanji style.
wdCaptionNumberStyleKanjiDigit	=	11		# Kanji digit style.
wdCaptionNumberStyleKanjiTraditional	=	16		# Kanji traditional style.
wdCaptionNumberStyleLowercaseLetter	=	4		# Lowercase letter style.
wdCaptionNumberStyleLowercaseRoman	=	2		# Lowercase roman style.
wdCaptionNumberStyleNumberInCircle	=	18		# Number in circle style.
wdCaptionNumberStyleSimpChinNum2	=	38		# Simplified Chinese number style 2.
wdCaptionNumberStyleSimpChinNum3	=	39		# Simplified Chinese number style 3.
wdCaptionNumberStyleThaiArabic	=	54		# Thai Arabic style.
wdCaptionNumberStyleThaiCardinalText	=	55		# Thai cardinal text style.
wdCaptionNumberStyleThaiLetter	=	53		# Thai letter style.
wdCaptionNumberStyleTradChinNum2	=	34		# Traditional Chinese number style 2.
wdCaptionNumberStyleTradChinNum3	=	35		# Traditional Chinese number style 3.
wdCaptionNumberStyleUppercaseLetter	=	3		# Uppercase letter style.
wdCaptionNumberStyleUppercaseRoman	=	1		# Uppercase roman style.
wdCaptionNumberStyleVietCardinalText	=	56		# Vietnamese cardinal text style.
wdCaptionNumberStyleZodiac1	=	30		# Zodiac style 1.
wdCaptionNumberStyleZodiac2	=	31		# Zodiac style 2.

#values from word.wdcaptionposition
#****************************
wdCaptionPositionAbove	=	0		# The caption label is added above.
wdCaptionPositionBelow	=	1		# The caption label is added below.

#values from word.wdcellcolor
#****************************
wdCellColorByAuthor	=	-1		# Highlighting color determined by reviewer.
wdCellColorLightBlue	=	2		# Light blue.
wdCellColorLightGray	=	7		# Light gray.
wdCellColorLightGreen	=	6		# Light green.
wdCellColorLightOrange	=	5		# Light orange.
wdCellColorLightPurple	=	4		# Light purple.
wdCellColorLightYellow	=	3		# Light yellow.
wdCellColorNoHighlight	=	0		# No highlighting.
wdCellColorPink	=	1		# Pink.

#values from word.wdcellverticalalignment
#****************************
wdCellAlignVerticalBottom	=	3		# Text is aligned to the bottom border of the cell.
wdCellAlignVerticalCenter	=	1		# Text is aligned to the center of the cell.
wdCellAlignVerticalTop	=	0		# Text is aligned to the top border of the cell.

#values from word.wdcharactercase
#****************************
wdFullWidth	=	7		# Full-width. Used for Japanese characters.
wdHalfWidth	=	6		# Half-width. Used for Japanese characters.
wdHiragana	=	9		# Hiragana characters. Used with Japanese text.
wdKatakana	=	8		# Katakana characters. Used with Japanese text.
wdLowerCase	=	0		# Lowercase.
wdNextCase	=	-1		# Toggles between uppercase, lowercase, and sentence case.
wdTitleSentence	=	4		# Sentence case.
wdTitleWord	=	2		# Title word case.
wdToggleCase	=	5		# Switches uppercase characters to lowercase, and lowercase characters to uppercase.
wdUpperCase	=	1		# Uppercase.

#values from word.wdcharacterwidth
#****************************
wdWidthFullWidth	=	7		# Characters are displayed in full character width.
wdWidthHalfWidth	=	6		# Characters are displayed in half the character width.

#values from word.wdcheckinversiontype
#****************************
wdCheckInMajorVersion	=	1		# Major version.
wdCheckInMinorVersion	=	0		# Minor version.
wdCheckInOverwriteVersion	=	2		# Overwrite current version on the server.

#values from word.wdchevronconvertrule
#****************************
wdAlwaysConvert	=	1		# The converter attempts to convert text enclosed in chevrons (? ?) to mail merge fields.
wdAskToConvert	=	3		# The converter prompts the user to convert or not convert chevrons when a Word for the Macintosh document is opened.
wdAskToNotConvert	=	2		# The converter prompts the user to convert or not convert chevrons when a Word for the Macintosh document is opened.
wdNeverConvert	=	0		# The converter passes the text through without attempting any interpretation.

#values from word.wdcollapsedirection
#****************************
wdCollapseEnd	=	0		# Collapse the range to the ending point.
wdCollapseStart	=	1		# Collapse the range to the starting point.

#values from word.wdcolor
#****************************
wdColorAqua	=	13421619		# Aqua color.
wdColorAutomatic	=	-16777216		# Automatic color. Default; usually black.
wdColorBlack	=	0		# Black color.
wdColorBlue	=	16711680		# Blue color.
wdColorBlueGray	=	10053222		# Blue-gray color.
wdColorBrightGreen	=	65280		# Bright green color.
wdColorBrown	=	13209		# Brown color.
wdColorDarkBlue	=	8388608		# Dark blue color.
wdColorDarkGreen	=	13056		# Dark green color.
wdColorDarkRed	=	128		# Dark red color.
wdColorDarkTeal	=	6697728		# Dark teal color.
wdColorDarkYellow	=	32896		# Dark yellow color.
wdColorGold	=	52479		# Gold color.
wdColorGray05	=	15987699		# Shade 05 of gray color.
wdColorGray10	=	15132390		# Shade 10 of gray color.
wdColorGray125	=	14737632		# Shade 125 of gray color.
wdColorGray15	=	14277081		# Shade 15 of gray color.
wdColorGray20	=	13421772		# Shade 20 of gray color.
wdColorGray25	=	12632256		# Shade 25 of gray color.
wdColorGray30	=	11776947		# Shade 30 of gray color.
wdColorGray35	=	10921638		# Shade 35 of gray color.
wdColorGray375	=	10526880		# Shade 375 of gray color.
wdColorGray40	=	10066329		# Shade 40 of gray color.
wdColorGray45	=	9211020		# Shade 45 of gray color.
wdColorGray50	=	8421504		# Shade 50 of gray color.
wdColorGray55	=	7566195		# Shade 55 of gray color.
wdColorGray60	=	6710886		# Shade 60 of gray color.
wdColorGray625	=	6316128		# Shade 625 of gray color.
wdColorGray65	=	5855577		# Shade 65 of gray color.
wdColorGray70	=	5000268		# Shade 70 of gray color.
wdColorGray75	=	4210752		# Shade 75 of gray color.
wdColorGray80	=	3355443		# Shade 80 of gray color.
wdColorGray85	=	2500134		# Shade 85 of gray color.
wdColorGray875	=	2105376		# Shade 875 of gray color.
wdColorGray90	=	1644825		# Shade 90 of gray color.
wdColorGray95	=	789516		# Shade 95 of gray color.
wdColorGreen	=	32768		# Green color.
wdColorIndigo	=	10040115		# Indigo color.
wdColorLavender	=	16751052		# Lavender color.
wdColorLightBlue	=	16737843		# Light blue color.
wdColorLightGreen	=	13434828		# Light green color.
wdColorLightOrange	=	39423		# Light orange color.
wdColorLightTurquoise	=	16777164		# Light turquoise color.
wdColorLightYellow	=	10092543		# Light yellow color.
wdColorLime	=	52377		# Lime color.
wdColorOliveGreen	=	13107		# Olive green color.
wdColorOrange	=	26367		# Orange color.
wdColorPaleBlue	=	16764057		# Pale blue color.
wdColorPink	=	16711935		# Pink color.
wdColorPlum	=	6697881		# Plum color.
wdColorRed	=	255		# Red color.
wdColorRose	=	13408767		# Rose color.
wdColorSeaGreen	=	6723891		# Sea green color.
wdColorSkyBlue	=	16763904		# Sky blue color.
wdColorTan	=	10079487		# Tan color.
wdColorTeal	=	8421376		# Teal color.
wdColorTurquoise	=	16776960		# Turquoise color.
wdColorViolet	=	8388736		# Violet color.
wdColorWhite	=	16777215		# White color.
wdColorYellow	=	65535		# Yellow color.

#values from word.wdcolorindex
#****************************
wdAuto	=	0		# Automatic color. Default; usually black.
wdBlack	=	1		# Black color.
wdBlue	=	2		# Blue color.
wdBrightGreen	=	4		# Bright green color.
wdByAuthor	=	-1		# Color defined by document author.
wdDarkBlue	=	9		# Dark blue color.
wdDarkRed	=	13		# Dark red color.
wdDarkYellow	=	14		# Dark yellow color.
wdGray25	=	16		# Shade 25 of gray color.
wdGray50	=	15		# Shade 50 of gray color.
wdGreen	=	11		# Green color.
wdNoHighlight	=	0		# Removes highlighting that has been applied.
wdPink	=	5		# Pink color.
wdRed	=	6		# Red color.
wdTeal	=	10		# Teal color.
wdTurquoise	=	3		# Turquoise color.
wdViolet	=	12		# Violet color.
wdWhite	=	8		# White color.
wdYellow	=	7		# Yellow color.

#values from word.wdcolumnwidth
#****************************
wdColumnWidthDefault	=	2		# Default column width.
wdColumnWidthNarrow	=	1		# Narrow column width.
wdColumnWidthWide	=	3		# Wide column width.

#values from word.wdcomparedestination
#****************************
wdCompareDestinationNew	=	2		# Creates a new file and tracks the differences between the original document and the revised document using tracked changes.
wdCompareDestinationOriginal	=	0		# Tracks the differences between the two files using tracked changes in the original document.
wdCompareDestinationRevised	=	1		# Tracks the differences between the two files using tracked changes in the revised document.

#values from word.wdcomparetarget
#****************************
wdCompareTargetCurrent	=	1		# Places comparison differences in the current document. Default.
wdCompareTargetNew	=	2		# Places comparison differences in a new document.
wdCompareTargetSelected	=	0		# Places comparison differences in the target document.

#values from word.wdcompatibility
#****************************
wdAlignTablesRowByRow	=	39		# Align table rows independently.
wdApplyBreakingRules	=	46		# Use line-breaking rules.
wdAutospaceLikeWW7	=	38		# Autospace like Microsoft Word 95.
wdConvMailMergeEsc	=	6		# Treat &quot; as &quot;&quot; in mail merge data sources.
wdDontAdjustLineHeightInTable	=	36		# Adjust line height to grid height in the table.
wdDontBalanceSingleByteDoubleByteWidth	=	16		# Balance SBCS characters and DBCS characters.
wdDontBreakWrappedTables	=	43		# Do not break wrapped tables across pages.
wdDontSnapTextToGridInTableWithObjects	=	44		# Do not snap text to grid inside table with inline objects.
wdDontULTrailSpace	=	15		# Draw underline on trailing spaces.
wdDontUseAsianBreakRulesInGrid	=	48		# Do not use Asian rules for line breaks with character grid.
wdDontUseHTMLParagraphAutoSpacing	=	35		# Do not use HTML paragraph auto spacing.
wdDontWrapTextWithPunctuation	=	47		# Do not allow hanging punctuation with character grid.
wdExactOnTop	=	28		# Do not center &quot;exact line height&quot; lines.
wdExpandShiftReturn	=	14		# Do not expand character spaces on the line ending Shift+Return.
wdFootnoteLayoutLikeWW8	=	34		# Lay out footnotes like Word 6.x/95/97.
wdForgetLastTabAlignment	=	37		# Forget last tab alignment.
wdGrowAutofit	=	50		# Allow tables to extend into margins.
wdLayoutRawTableWidth	=	40		# Lay out tables with raw width.
wdLayoutTableRowsApart	=	41		# Allow table rows to lay out apart.
wdLeaveBackslashAlone	=	13		# Convert backslash characters into yen signs.
wdLineWrapLikeWord6	=	32		# Line wrap like Word 6.0.
wdMWSmallCaps	=	22		# Use larger small caps like Word 5.x for the Macintosh.
wdNoColumnBalance	=	5		# Do not balance columns for continuous section starts.
wdNoExtraLineSpacing	=	23		# Suppress extra line spacing like WordPerfect 5.x.
wdNoLeading	=	20		# Do not add leading (extra space) between rows of text.
wdNoSpaceForUL	=	21		# Add space for underline.
wdNoSpaceRaiseLower	=	2		# Do not add extra space for raised/lowered characters.
wdNoTabHangIndent	=	1		# Do not add automatic tab stop for hanging indent.
wdOrigWordTableRules	=	9		# Combine table borders like Word 5.x for the Macintosh.
wdPrintBodyTextBeforeHeader	=	19		# Print body text before header/footer.
wdPrintColBlack	=	3		# Print colors as black on noncolor printers.
wdSelectFieldWithFirstOrLastCharacter	=	45		# Select entire field with first or last character.
wdShapeLayoutLikeWW8	=	33		# Lay out autoshapes like Word 97.
wdShowBreaksInFrames	=	11		# Show hard page or column breaks in frames.
wdSpacingInWholePoints	=	18		# Expand/condense by whole number of points.
wdSubFontBySize	=	25		# Substitute fonts based on font size.
wdSuppressBottomSpacing	=	29		# Suppress extra line spacing at bottom of page.
wdSuppressSpBfAfterPgBrk	=	7		# Suppress Space Before after a hard page or column break.
wdSuppressTopSpacing	=	8		# Suppress extra line spacing at top of page.
wdSuppressTopSpacingMac5	=	17		# Suppress extra line spacing at top of page like Word 5.x for the Macintosh.
wdSwapBordersFacingPages	=	12		# Swap left and right borders on odd facing pages.
wdTransparentMetafiles	=	10		# Do not blank the area behind metafile pictures.
wdTruncateFontHeight	=	24		# Truncate font height.
wdUsePrinterMetrics	=	26		# Use printer metrics to lay out document.
wdUseWord2002TableStyleRules	=	49		# Use Microsoft Word 2002 table style rules.
wdUseWord2010TableStyleRules	=	69		# Use Microsoft Word 2010 table style rules.
wdUseWord97LineBreakingRules	=	42		# Use Microsoft Word 97 line breaking rules for Asian text.
wdWPJustification	=	31		# Do full justification like WordPerfect 6.x for Windows.
wdWPSpaceWidth	=	30		# Set the width of a space like WordPerfect 5.x.
wdWrapTrailSpaces	=	4		# Wrap trailing spaces to next line.
wdWW6BorderRules	=	27		# Use Word 6.x/95 border rules.
wdAllowSpaceOfSameStyleInTable	=	54		# Allow space between paragraphs of the same style in a table.
wdAutofitLikeWW11	=	57		# Use Microsoft Word 2003 table autofit rules.
wdDontAutofitConstrainedTables	=	56		# Do not autofit tables next to wrapped objects.
wdDontUseIndentAsNumberingTabStop	=	52		# Do not use hanging indent as tab stop for bullets and numbering.
wdFELineBreak11	=	53		# Use Word 2003 hanging-punctuation rules in Asian languages.
wdHangulWidthLikeWW11	=	59		# Do not use proportional width for Korean characters.
wdSplitPgBreakAndParaMark	=	60		# Split apart page break and paragraph mark.
wdUnderlineTabInNumList	=	58		# Underline the tab character between the number and the text in numbered lists.
wdUseNormalStyleForList	=	51		# Use the Normal style instead of the List Paragraph style for bulleted or numbered lists.
wdWW11IndentRules	=	55		# Use Word 2003 indent rules for text next to wrapped objects.

#values from word.wdcompatibilitymode
#****************************
wdCurrent	=	65535		# Compatibility mode equivalent to the latest version of Word.
wdWord2003	=	11		# Word is put into a mode that is most compatible with Word 2003. Features new to Word are disabled in this mode.
wdWord2007	=	12		# Word is put into a mode that is most compatible with Office Word 2007. Features new to Wordare disabled in this mode.
wdWord2010	=	14		# Word is put into a mode that is most compatible with . Features new to Wordare disabled in this mode.
wdWord2013	=	15		# Default. All Word features are enabled.

#values from word.wdconditioncode
#****************************
wdEvenColumnBanding	=	7		# Applies formatting to even-numbered columns.
wdEvenRowBanding	=	3		# Applies formatting to even-numbered rows.
wdFirstColumn	=	4		# Applies formatting to the first column in a table.
wdFirstRow	=	0		# Applies formatting to the first row in a table.
wdLastColumn	=	5		# Applies formatting to the last column in a table.
wdLastRow	=	1		# Applies formatting to the last row in a table.
wdNECell	=	8		# Applies formatting to the last cell in the first row.
wdNWCell	=	9		# Applies formatting to the first cell in the first row.
wdOddColumnBanding	=	6		# Applies formatting to odd-numbered columns.
wdOddRowBanding	=	2		# Applies formatting to odd-numbered rows.
wdSECell	=	10		# Applies formatting to the last cell in the table.
wdSWCell	=	11		# Applies formatting to first cell in the last row of the table.

#values from word.wdconstants
#****************************
wdAutoPosition	=	0		# Represents the Auto value for the specified setting.
wdBackward	=	-1073741823		# Indicates that selection will be extended backward using the  MoveStartUntil or MoveStartWhile method of the Range or Selection object.
wdCreatorCode	=	1297307460		# Represents the creator code for objects created by Microsoft Word.
wdFirst	=	1		# Represents the first item in a collection.
wdForward	=	1073741823		# Indicates that selection will be extended forward using the  MoveStartUntil or MoveStartWhile method of the Range or Selection object.
wdToggle	=	9999998		# Toggles a property's value.
wdUndefined	=	9999999		# Represents an undefined value.

#values from word.wdcontentcontrolappearance
#****************************
wdContentControlBoundingBox	=	0		# Represents a content control shown as a shaded rectangle or bounding box (with optional title).
wdContentControlTags	=	2		# Represents a content control shown as start and end markers.
wdContentControlHidden	=	1		# Represents a content control that is not shown.

#values from word.wdcontentcontroldatestorageformat
#****************************
wdContentControlDateStorageDate	=	1		# Specifies to store or retrieve the date value for a date content control as a date in the standard XML Schema DateTime format.
wdContentControlDateStorageDateTime	=	2		# Specifies to store or retrieve the date value for a date content control as a time in the standard XML Schema DateTime format.
wdContentControlDateStorageText	=	0		# Specifies to store or retrieve the date value for a date content control as text.

#values from word.wdcontentcontrollevel
#****************************
wdContentControlLevelCell	=	3		# Represents a content control that surrounds a table cell.
wdContentControlLevelInline	=	0		# Represents a content control that surrounds content within a single paragraph.
wdContentControlLevelParagraph	=	1		# Represents a content control that surrounds one or more complete paragraphs.
wdContentControlLevelRow	=	2		# Represents a content control that surrounds a table row.

#values from word.wdcontentcontroltype
#****************************
wdContentControlBuildingBlockGallery	=	5		# Specifies a building block gallery content control.
wdContentControlCheckbox	=	8		# Specifies a checkbox content control.
wdContentControlComboBox	=	3		# Specifies a combo box content control.
wdContentControlDate	=	6		# Specifies a date content control.
wdContentControlGroup	=	7		# Specifies a group content control.
wdContentControlDropdownList	=	4		# Specifies a drop-down list content control.
wdContentControlPicture	=	2		# Specifies a picture content control.
wdContentControlRepeatingSection	=	9		# Specifies a repeating section content control.
wdContentControlRichText	=	0		# Specifies a rich-text content control.
wdContentControlText	=	1		# Specifies a text content control

#values from word.wdcontinue
#****************************
wdContinueDisabled	=	0		# Formatting cannot continue from the previous list.
wdContinueList	=	2		# Formatting can continue from the previous list.
wdResetList	=	1		# Numbering can be restarted.

#values from word.wdcountry
#****************************
wdArgentina	=	54		# Argentina
wdBrazil	=	55		# Brazil
wdCanada	=	2		# Canada
wdChile	=	56		# Chile
wdChina	=	86		# China
wdDenmark	=	45		# Denmark
wdFinland	=	358		# Finland
wdFrance	=	33		# France
wdGermany	=	49		# Germany
wdIceland	=	354		# Iceland
wdItaly	=	39		# Italy
wdJapan	=	81		# Japan
wdKorea	=	82		# Korea
wdLatinAmerica	=	3		# Latin America
wdMexico	=	52		# Mexico
wdNetherlands	=	31		# Netherlands
wdNorway	=	47		# Norway
wdPeru	=	51		# Peru
wdSpain	=	34		# Spain
wdSweden	=	46		# Sweden
wdTaiwan	=	886		# Taiwan
wdUK	=	44		# United Kingdom
wdUS	=	1		# United States
wdVenezuela	=	58		# Venezuela

#values from word.wdcursormovement
#****************************
wdCursorMovementLogical	=	0		# Insertion point progresses according to the direction of the language Microsoft Word detects.
wdCursorMovementVisual	=	1		# Insertion point progresses to the next visually adjacent character.

#values from word.wdcursortype
#****************************
wdCursorIBeam	=	1		# I-beam cursor shape.
wdCursorNormal	=	2		# Normal cursor shape. Default; cursor takes shape designated by Microsoft Windows or the application.
wdCursorNorthwestArrow	=	3		# Diagonal cursor shape starting at upper-left corner.
wdCursorWait	=	0		# Hourglass cursor shape.

#values from word.wdcustomlabelpagesize
#****************************
wdCustomLabelA4	=	2		# A4 portrait label dimensions.
wdCustomLabelA4LS	=	3		# A4 landscape label dimensions.
wdCustomLabelA5	=	4		# A5 portrait label dimensions.
wdCustomLabelA5LS	=	5		# A5 landscape label dimensions.
wdCustomLabelB4JIS	=	13		# B4 JIS label dimensions.
wdCustomLabelB5	=	6		# B5 label dimensions.
wdCustomLabelFanfold	=	8		# Fanfold label dimensions.
wdCustomLabelHigaki	=	11		# Higaki portrait label dimensions.
wdCustomLabelHigakiLS	=	12		# Higaki landscape label dimensions.
wdCustomLabelLetter	=	0		# Standard letter portrait label dimensions.
wdCustomLabelLetterLS	=	1		# Standard letter landscape label dimensions.
wdCustomLabelMini	=	7		# Mini label dimensions.
wdCustomLabelVertHalfSheet	=	9		# Half-sheet portrait label dimensions.
wdCustomLabelVertHalfSheetLS	=	10		# Half-sheet landscape label dimensions.

#values from word.wddatelanguage
#****************************
wdDateLanguageBidi	=	10		# Bidirectional date/time format.
wdDateLanguageLatin	=	1033		# Latin date/time format.

#values from word.wddefaultfilepath
#****************************
wdAutoRecoverPath	=	5		# Path for Auto Recover files.
wdBorderArtPath	=	19		# Border art path.
wdCurrentFolderPath	=	14		# Current folder path.
wdDocumentsPath	=	0		# Documents path.
wdGraphicsFiltersPath	=	10		# Graphics filters path.
wdPicturesPath	=	1		# Pictures path.
wdProgramPath	=	9		# Program path.
wdProofingToolsPath	=	12		# Proofing tools path.
wdStartupPath	=	8		# Startup path.
wdStyleGalleryPath	=	15		# Style Gallery path.
wdTempFilePath	=	13		# Temp file path.
wdTextConvertersPath	=	11		# Text converters path.
wdToolsPath	=	6		# Tools path.
wdTutorialPath	=	7		# Tutorial path.
wdUserOptionsPath	=	4		# User Options path.
wdUserTemplatesPath	=	2		# User templates path.
wdWorkgroupTemplatesPath	=	3		# Workgroup templates path.

#values from word.wddefaultlistbehavior
#****************************
wdWord10ListBehavior	=	2		# Use formatting compatible with Microsoft Word 2002.
wdWord8ListBehavior	=	0		# Use formatting compatible with Microsoft Word 97.
wdWord9ListBehavior	=	1		# Use Web-oriented formatting as introduced in Microsoft Word 2000.

#values from word.wddefaulttablebehavior
#****************************
wdWord8TableBehavior	=	0		# Disables AutoFit. Default.
wdWord9TableBehavior	=	1		# Enables AutoFit.

#values from word.wddeletecells
#****************************
wdDeleteCellsEntireColumn	=	3		# Delete the entire column of cells from the table.
wdDeleteCellsEntireRow	=	2		# Delete the entire row of cells from the table.
wdDeleteCellsShiftLeft	=	0		# Shift remaining cells left in the row where the deletion occurred after a cell or range of cells has been deleted.
wdDeleteCellsShiftUp	=	1		# Shift remaining cells up in the column where the deletion occurred after a cell or range of cells has been deleted.

#values from word.wddeletedtextmark
#****************************
wdDeletedTextMarkBold	=	5		# Deleted text is displayed in bold.
wdDeletedTextMarkCaret	=	2		# Deleted text is marked up by using caret characters.
wdDeletedTextMarkColorOnly	=	9		# Deleted text is displayed in a specified color (default is red).
wdDeletedTextMarkDoubleUnderline	=	8		# Deleted text is marked up by using double-underline characters.
wdDeletedTextMarkHidden	=	0		# Deleted text is hidden.
wdDeletedTextMarkItalic	=	6		# Deleted text is displayed in italic.
wdDeletedTextMarkNone	=	4		# Deleted text is not marked up.
wdDeletedTextMarkPound	=	3		# Deleted text is marked up by using pound characters.
wdDeletedTextMarkStrikeThrough	=	1		# Deleted text is marked up by using strikethrough characters.
wdDeletedTextMarkUnderline	=	7		# Deleted text is underlined.
wdDeletedTextMarkDoubleStrikeThrough	=	10		# Deleted text is marked up by using double-strikethrough characters.

#values from word.wddiacriticcolor
#****************************
wdDiacriticColorBidi	=	0		# Bi-directional language (Arabic, Hebrew, and so forth).
wdDiacriticColorLatin	=	1		# Latin style languages.

#values from word.wddictionarytype
#****************************
wdGrammar	=	1		# Grammar.
wdHangulHanjaConversion	=	8		# Dictionary for converting between Hangul and Hanja. Available only if you have enabled support for Korean through Microsoft Office Language Settings.
wdHangulHanjaConversionCustom	=	9		# Custom dictionary for converting between Hangul and Hanja.
wdHyphenation	=	3		# Hyphenation.
wdSpelling	=	0		# Spelling.
wdSpellingComplete	=	4		# Not supported.
wdSpellingCustom	=	5		# Custom spelling dictionary.
wdSpellingLegal	=	6		# Legal dictionary.
wdSpellingMedical	=	7		# Medical dictionary.
wdThesaurus	=	2		# Thesaurus.

#values from word.wddisablefeaturesintroducedafter
#****************************
wd70	=	0		# Specifies Word for Windows 95, versions 7.0 and 7.0a.
wd70FE	=	1		# Specifies Word for Windows 95, versions 7.0 and 7.0a, Asian edition.
wd80	=	2		# Specifies Word 97 for Windows. Default.

#values from word.wddocpartinsertoptions
#****************************
wdInsertContent	=	0		# Inline building block.
wdInsertPage	=	2		# Page-level building block.
wdInsertParagraph	=	1		# Paragraph-level building block.

#values from word.wddocumentdirection
#****************************
wdLeftToRight	=	0		# Left to right.
wdRightToLeft	=	1		# Right to left.

#values from word.wddocumentkind
#****************************
wdDocumentEmail	=	2		# Email format.
wdDocumentLetter	=	1		# Letter format.
wdDocumentNotSpecified	=	0		# No format specified.

#values from word.wddocumentmedium
#****************************
wdDocument	=	1		# Document.
wdEmailMessage	=	0		# Email message.
wdWebPage	=	2		# Web page.

#values from word.wddocumenttype
#****************************
wdTypeDocument	=	0		# Document.
wdTypeFrameset	=	2		# Frameset.
wdTypeTemplate	=	1		# Template.

#values from word.wddocumentviewdirection
#****************************
wdDocumentViewLtr	=	1		# Displays the document with left alignment and left-to-right reading order.
wdDocumentViewRtl	=	0		# Displays the document with right alignment and right-to-left reading order.

#values from word.wddropposition
#****************************
wdDropMargin	=	2		# Dropped capital letter ends at the left margin.
wdDropNone	=	0		# No dropped capital letter.
wdDropNormal	=	1		# Dropped capital letter begins at the left margin.

#values from word.wdeditionoption
#****************************
wdAutomaticUpdate	=	3		# Not supported.
wdCancelPublisher	=	0		# Not supported.
wdChangeAttributes	=	5		# Not supported.
wdManualUpdate	=	4		# Not supported.
wdOpenSource	=	7		# Not supported.
wdSelectPublisher	=	2		# Not supported.
wdSendPublisher	=	1		# Not supported.
wdUpdateSubscriber	=	6		# Not supported.

#values from word.wdeditiontype
#****************************
wdPublisher	=	0		# Not supported.
wdSubscriber	=	1		# Not supported.

#values from word.wdeditortype
#****************************
wdEditorCurrent	=	-6		# Represents the current user of the document.
wdEditorEditors	=	-5		# Represents the Editors group for documents that use Information Rights Management.
wdEditorEveryone	=	-1		# Represents all users who open a document.
wdEditorOwners	=	-4		# Represents the Owners group for documents that use Information Rights Management.

#values from word.wdemailhtmlfidelity
#****************************
wdEmailHTMLFidelityHigh	=	3		# Leaves HTML intact.
wdEmailHTMLFidelityLow	=	1		# Removes all HTML tags that do not affect how a message displays.
wdEmailHTMLFidelityMedium	=	2		# Not supported.

#values from word.wdemphasismark
#****************************
wdEmphasisMarkNone	=	0		# No emphasis mark.
wdEmphasisMarkOverComma	=	2		# A comma.
wdEmphasisMarkOverSolidCircle	=	1		# A solid black circle.
wdEmphasisMarkOverWhiteCircle	=	3		# An empty white circle.
wdEmphasisMarkUnderSolidCircle	=	4		# A solid black circle.

#values from word.wdenablecancelkey
#****************************
wdCancelDisabled	=	0		# Prevents CTRL+BREAK from interrupting a macro.
wdCancelInterrupt	=	1		# Allows a macro to be interrupted by CTRL+BREAK.

#values from word.wdenclosestyle
#****************************
wdEncloseStyleLarge	=	2		# The enclosure is larger.
wdEncloseStyleNone	=	0		# The enclosure assumes the default size.
wdEncloseStyleSmall	=	1		# The enclosure is smaller.

#values from word.wdenclosuretype
#****************************
wdEnclosureCircle	=	0		# A circle.
wdEnclosureDiamond	=	3		# A diamond.
wdEnclosureSquare	=	1		# A square.
wdEnclosureTriangle	=	2		# A triangle.

#values from word.wdendnotelocation
#****************************
wdEndOfDocument	=	1		# At end of active document.
wdEndOfSection	=	0		# At end of current section.

#values from word.wdenvelopeorientation
#****************************
wdCenterClockwise	=	7		# Center clockwise orientation.
wdCenterLandscape	=	4		# Center landscape orientation.
wdCenterPortrait	=	1		# Center portrait orientation.
wdLeftClockwise	=	6		# Left clockwise orientation.
wdLeftLandscape	=	3		# Left landscape orientation.
wdLeftPortrait	=	0		# Left portrait orientation.
wdRightClockwise	=	8		# Right clockwise orientation.
wdRightLandscape	=	5		# Right landscape orientation.
wdRightPortrait	=	2		# Right portrait orientation.

#values from word.wdexportcreatebookmarks
#****************************
wdExportCreateHeadingBookmarks	=	1		# Create a bookmark in the exported document for each Microsoft Word heading, which includes only headings within the main document and text boxes not within headers, footers, endnotes, footnotes, or comments.
wdExportCreateNoBookmarks	=	0		# Do not create bookmarks in the exported document.
wdExportCreateWordBookmarks	=	2		# Create a bookmark in the exported document for each Word bookmark, which includes all bookmarks except those contained within headers and footers.

#values from word.wdexportformat
#****************************
wdExportFormatPDF	=	17		# Export document into PDF format.
wdExportFormatXPS	=	18		# Export document into XML Paper Specification (XPS) format.

#values from word.wdexportitem
#****************************
wdExportDocumentContent	=	0		# Exports the document without markup.
wdExportDocumentWithMarkup	=	7		# Exports the document with markup.

#values from word.wdexportoptimizefor
#****************************
wdExportOptimizeForOnScreen	=	1		# Export for screen, which is a lower quality and results in a smaller file size.
wdExportOptimizeForPrint	=	0		# Export for print, which is higher quality and results in a larger file size.

#values from word.wdexportrange
#****************************
wdExportAllDocument	=	0		# Exports the entire document.
wdExportCurrentPage	=	2		# Exports the current page.
wdExportFromTo	=	3		# Exports the contents of a range using the starting and ending positions.
wdExportSelection	=	1		# Exports the contents of the current selection.

#values from word.wdfareastlinebreaklanguageid
#****************************
wdLineBreakJapanese	=	1041		# Japanese.
wdLineBreakKorean	=	1042		# Korean.
wdLineBreakSimplifiedChinese	=	2052		# Simplified Chinese.
wdLineBreakTraditionalChinese	=	1028		# Traditional Chinese.

#values from word.wdfareastlinebreaklevel
#****************************
wdFarEastLineBreakLevelCustom	=	2		# Custom line break control.
wdFarEastLineBreakLevelNormal	=	0		# Normal line break control.
wdFarEastLineBreakLevelStrict	=	1		# Strict line break control.

#values from word.wdfieldkind
#****************************
wdFieldKindCold	=	3		# A field that does not have a result, for example, an Index Entry (XE), Table of Contents Entry (TC), or Private field.
wdFieldKindHot	=	1		# A field that's automatically updated each time it is displayed or each time the page is reformatted, but which can also be manually updated (for example, INCLUDEPICTURE or FORMDROPDOWN).
wdFieldKindNone	=	0		# An invalid field (for example, a pair of field characters with nothing inside).
wdFieldKindWarm	=	2		# A field that can be updated and has a result. This type includes fields that are automatically updated when the source changes and fields that can be manually updated (for example, DATE or INCLUDETEXT).

#values from word.wdfieldshading
#****************************
wdFieldShadingAlways	=	1		# Always apply.
wdFieldShadingNever	=	0		# Never apply.
wdFieldShadingWhenSelected	=	2		# Apply only when form field is selected.

#values from word.wdfieldtype
#****************************
wdFieldAddin	=	81		# Add-in field. Not available through the  Field dialog box. Used to store data that is hidden from the user interface.
wdFieldAddressBlock	=	93		# AddressBlock field.
wdFieldAdvance	=	84		# Advance field.
wdFieldAsk	=	38		# Ask field.
wdFieldAuthor	=	17		# Author field.
wdFieldAutoNum	=	54		# AutoNum field.
wdFieldAutoNumLegal	=	53		# AutoNumLgl field.
wdFieldAutoNumOutline	=	52		# AutoNumOut field.
wdFieldAutoText	=	79		# AutoText field.
wdFieldAutoTextList	=	89		# AutoTextList field.
wdFieldBarCode	=	63		# BarCode field.
wdFieldBidiOutline	=	92		# BidiOutline field.
wdFieldComments	=	19		# Comments field.
wdFieldCompare	=	80		# Compare field.
wdFieldCreateDate	=	21		# CreateDate field.
wdFieldData	=	40		# Data field.
wdFieldDatabase	=	78		# Database field.
wdFieldDate	=	31		# Date field.
wdFieldDDE	=	45		# DDE field. No longer available through the  Field dialog box, but supported for documents created in earlier versions of Word.
wdFieldDDEAuto	=	46		# DDEAuto field. No longer available through the  Field dialog box, but supported for documents created in earlier versions of Word.
wdFieldDisplayBarcode	=	99		# DisplayBarcode field.
wdFieldDocProperty	=	85		# DocProperty field.
wdFieldDocVariable	=	64		# DocVariable field.
wdFieldEditTime	=	25		# EditTime field.
wdFieldEmbed	=	58		# Embedded field.
wdFieldEmpty	=	-1		# Empty field. Acts as a placeholder for field content that has not yet been added. A field added by pressing Ctrl+F9 in the user interface is an Empty field.
wdFieldExpression	=	34		# = (Formula) field.
wdFieldFileName	=	29		# FileName field.
wdFieldFileSize	=	69		# FileSize field.
wdFieldFillIn	=	39		# Fill-In field.
wdFieldFootnoteRef	=	5		# FootnoteRef field. Not available through the  Field dialog box. Inserted programmatically or interactively.
wdFieldFormCheckBox	=	71		# FormCheckBox field.
wdFieldFormDropDown	=	83		# FormDropDown field.
wdFieldFormTextInput	=	70		# FormText field.
wdFieldFormula	=	49		# EQ (Equation) field.
wdFieldGlossary	=	47		# Glossary field. No longer supported in Word.
wdFieldGoToButton	=	50		# GoToButton field.
wdFieldGreetingLine	=	94		# GreetingLine field.
wdFieldHTMLActiveX	=	91		# HTMLActiveX field. Not currently supported.
wdFieldHyperlink	=	88		# Hyperlink field.
wdFieldIf	=	7		# If field.
wdFieldImport	=	55		# Import field. Cannot be added through the  Field dialog box, but can be added interactively or through code.
wdFieldInclude	=	36		# Include field. Cannot be added through the  Field dialog box, but can be added interactively or through code.
wdFieldIncludePicture	=	67		# IncludePicture field.
wdFieldIncludeText	=	68		# IncludeText field.
wdFieldIndex	=	8		# Index field.
wdFieldIndexEntry	=	4		# XE (Index Entry) field.
wdFieldInfo	=	14		# Info field.
wdFieldKeyWord	=	18		# Keywords field.
wdFieldLastSavedBy	=	20		# LastSavedBy field.
wdFieldLink	=	56		# Link field.
wdFieldListNum	=	90		# ListNum field.
wdFieldMacroButton	=	51		# MacroButton field.
wdFieldMergeBarcode	=	98		# MergeBarcode field.
wdFieldMergeField	=	59		# MergeField field.
wdFieldMergeRec	=	44		# MergeRec field.
wdFieldMergeSeq	=	75		# MergeSeq field.
wdFieldNext	=	41		# Next field.
wdFieldNextIf	=	42		# NextIf field.
wdFieldNoteRef	=	72		# NoteRef field.
wdFieldNumChars	=	28		# NumChars field.
wdFieldNumPages	=	26		# NumPages field.
wdFieldNumWords	=	27		# NumWords field.
wdFieldOCX	=	87		# OCX field. Cannot be added through the  Field dialog box, but can be added through code by using the AddOLEControl method of the Shapes collection or of the InlineShapes collection.
wdFieldPage	=	33		# Page field.
wdFieldPageRef	=	37		# PageRef field.
wdFieldPrint	=	48		# Print field.
wdFieldPrintDate	=	23		# PrintDate field.
wdFieldPrivate	=	77		# Private field.
wdFieldQuote	=	35		# Quote field.
wdFieldRef	=	3		# Ref field.
wdFieldRefDoc	=	11		# RD (Reference Document) field.
wdFieldRevisionNum	=	24		# RevNum field.
wdFieldSaveDate	=	22		# SaveDate field.
wdFieldSection	=	65		# Section field.
wdFieldSectionPages	=	66		# SectionPages field.
wdFieldSequence	=	12		# Seq (Sequence) field.
wdFieldSet	=	6		# Set field.
wdFieldShape	=	95		# Shape field. Automatically created for any drawn picture.
wdFieldSkipIf	=	43		# SkipIf field.
wdFieldStyleRef	=	10		# StyleRef field.
wdFieldSubject	=	16		# Subject field.
wdFieldSubscriber	=	82		# Macintosh only. For information about this constant, consult the language reference Help included with Microsoft Office Macintosh Edition.
wdFieldSymbol	=	57		# Symbol field.
wdFieldTemplate	=	30		# Template field.
wdFieldTime	=	32		# Time field.
wdFieldTitle	=	15		# Title field.
wdFieldTOA	=	73		# TOA (Table of Authorities) field.
wdFieldTOAEntry	=	74		# TOA (Table of Authorities Entry) field.
wdFieldTOC	=	13		# TOC (Table of Contents) field.
wdFieldTOCEntry	=	9		# TOC (Table of Contents Entry) field.
wdFieldUserAddress	=	62		# UserAddress field.
wdFieldUserInitials	=	61		# UserInitials field.
wdFieldUserName	=	60		# UserName field.
wdFieldBibliography	=	97		# Bibliography field.
wdFieldCitation	=	96		# Citation field.

#values from word.wdfindmatch
#****************************
wdMatchAnyCharacter	=	65599		# Not supported.
wdMatchAnyDigit	=	65567		# Not supported.
wdMatchAnyLetter	=	65583		# Not supported.
wdMatchCaretCharacter	=	11		# Not supported.
wdMatchColumnBreak	=	14		# Not supported.
wdMatchCommentMark	=	5		# Not supported.
wdMatchEmDash	=	8212		# Not supported.
wdMatchEnDash	=	8211		# Not supported.
wdMatchEndnoteMark	=	65555		# Not supported.
wdMatchField	=	19		# Not supported.
wdMatchFootnoteMark	=	65554		# Not supported.
wdMatchGraphic	=	1		# Not supported.
wdMatchManualLineBreak	=	65551		# Not supported.
wdMatchManualPageBreak	=	65564		# Not supported.
wdMatchNonbreakingHyphen	=	30		# Not supported.
wdMatchNonbreakingSpace	=	160		# Not supported.
wdMatchOptionalHyphen	=	31		# Not supported.
wdMatchParagraphMark	=	65551		# Not supported.
wdMatchSectionBreak	=	65580		# Not supported.
wdMatchTabCharacter	=	9		# Not supported.
wdMatchWhiteSpace	=	65655		# Not supported.

#values from word.wdfindwrap
#****************************
wdFindAsk	=	2		# After searching the selection or range, Microsoft Word displays a message asking whether to search the remainder of the document.
wdFindContinue	=	1		# The find operation continues if the beginning or end of the search range is reached.
wdFindStop	=	0		# The find operation ends if the beginning or end of the search range is reached.

#values from word.wdflowdirection
#****************************
wdFlowLtr	=	0		# Text in columns flows from left to right.
wdFlowRtl	=	1		# Text in columns flows from right to left.

#values from word.wdfontbias
#****************************
wdFontBiasDefault	=	0		# Default font bias.
wdFontBiasDontCare	=	255		# No font bias specified.
wdFontBiasFareast	=	1		# Font bias for Asian languages.

#values from word.wdfootnotelocation
#****************************
wdBeneathText	=	1		# Beneath current text.
wdBottomOfPage	=	0		# At bottom of current page.

#values from word.wdframeposition
#****************************
wdFrameBottom	=	-999997		# Bottom margin.
wdFrameCenter	=	-999995		# Center of document.
wdFrameInside	=	-999994		# Content inside frame.
wdFrameLeft	=	-999998		# Left margin.
wdFrameOutside	=	-999993		# Content outside frame.
wdFrameRight	=	-999996		# Right margin.
wdFrameTop	=	-999999		# Top margin.

#values from word.wdframesetnewframelocation
#****************************
wdFramesetNewFrameAbove	=	0		# Above existing frame.
wdFramesetNewFrameBelow	=	1		# Below existing frame.
wdFramesetNewFrameLeft	=	3		# To the left of existing frame.
wdFramesetNewFrameRight	=	2		# To the right of existing frame.

#values from word.wdframesetsizetype
#****************************
wdFramesetSizeTypeFixed	=	1		# Microsoft Word interprets the height or width of the specified frame as a fixed value (in points).
wdFramesetSizeTypePercent	=	0		# Word interprets the height or width of the specified frame as a percentage of the screen height or width.
wdFramesetSizeTypeRelative	=	2		# Word interprets the height or width of the specified frame as relative to the height or width of other frames on the frames page.

#values from word.wdframesettype
#****************************
wdFramesetTypeFrame	=	1		# A single frame.
wdFramesetTypeFrameset	=	0		# A frameset.

#values from word.wdframesizerule
#****************************
wdFrameAtLeast	=	1		# Sets the height or width to a value equal to or greater than the value specified by the  Height property or Width property.
wdFrameAuto	=	0		# Sets the height or width according to the height or width of the item in the frame.
wdFrameExact	=	2		# Sets the height or width to an exact value specified by the  Height property or Width property.

#values from word.wdfrenchspeller
#****************************
wdFrenchBoth	=	0		# Use both Post Reform and Pre-Reform French dictionaries when checking French language spelling.
wdFrenchPostReform	=	2		# Use only the Post Reform French dictionary when checking French language spelling.
wdFrenchPreReform	=	1		# Use only the Pre-Reform French dictionary when checking French language spelling.

#values from word.wdgotodirection
#****************************
wdGoToAbsolute	=	1		# An absolute position.
wdGoToFirst	=	1		# The first instance of the specified object.
wdGoToLast	=	-1		# The last instance of the specified object.
wdGoToNext	=	2		# The next instance of the specified object.
wdGoToPrevious	=	3		# The previous instance of the specified object.
wdGoToRelative	=	2		# A position relative to the current position.

#values from word.wdgotoitem
#****************************
wdGoToBookmark	=	-1		# A bookmark.
wdGoToComment	=	6		# A comment.
wdGoToEndnote	=	5		# An endnote.
wdGoToEquation	=	10		# An equation.
wdGoToField	=	7		# A field.
wdGoToFootnote	=	4		# A footnote.
wdGoToGrammaticalError	=	14		# A grammatical error.
wdGoToGraphic	=	8		# A graphic.
wdGoToHeading	=	11		# A heading.
wdGoToLine	=	3		# A line.
wdGoToObject	=	9		# An object.
wdGoToPage	=	1		# A page.
wdGoToPercent	=	12		# A percent.
wdGoToProofreadingError	=	15		# A proofreading error.
wdGoToSection	=	0		# A section.
wdGoToSpellingError	=	13		# A spelling error.
wdGoToTable	=	2		# A table.

#values from word.wdgranularity
#****************************
wdGranularityCharLevel	=	0		# Tracks character-level changes.
wdGranularityWordLevel	=	1		# Tracks word-level changes.

#values from word.wdgutterstyle
#****************************
wdGutterPosLeft	=	0		# On the left side.
wdGutterPosRight	=	2		# On the right side.
wdGutterPosTop	=	1		# At the top.

#values from word.wdgutterstyleold
#****************************
wdGutterStyleBidi	=	2		# Bidirectional gutter should be used to conform to right-to-left text flow.
wdGutterStyleLatin	=	-10		# Latin gutter should be used to conform to left-to-right text flow.

#values from word.wdheaderfooterindex
#****************************
wdHeaderFooterEvenPages	=	3		# Returns all headers or footers on even-numbered pages.
wdHeaderFooterFirstPage	=	2		# Returns the first header or footer in a document or section.
wdHeaderFooterPrimary	=	1		# Returns the header or footer on all pages other than the first page of a document or section.

#values from word.wdheadingseparator
#****************************
wdHeadingSeparatorBlankLine	=	1		# A blank line.
wdHeadingSeparatorLetter	=	2		# A designated letter.
wdHeadingSeparatorLetterFull	=	4		# A designated uppercase letter.
wdHeadingSeparatorLetterLow	=	3		# A designated lowercase letter.
wdHeadingSeparatorNone	=	0		# No separator.

#values from word.wdhebspellstart
#****************************
wdFullScript	=	0		# The spelling checker follows rules for the conventional script required by the Hebrew Language Academy for writing text without diacritics.
wdMixedAuthorizedScript	=	3		# The spelling checker follows rules for full and partial script, but highlights as potential mistakes any spelling variations not permitted within either system and any completely unrecognized words.
wdMixedScript	=	2		# The spelling checker follows rules for full and partial script and allows non-conventional spelling variations. Only completely unrecognized words are highlighted as potential mistakes.
wdPartialScript	=	1		# The spelling checker follows rules for the traditional script used only for text with diacritics.

#values from word.wdhelptype
#****************************
wdHelp	=	0		# Displays the  Help Topics dialog box.
wdHelpAbout	=	1		# Displays the  About Microsoft Word dialog box.
wdHelpActiveWindow	=	2		# Displays Help describing the command associated with the active view or pane.
wdHelpContents	=	3		# Displays the  Help Topics dialog box.
wdHelpExamplesAndDemos	=	4		# Displays examples and demos.
wdHelpHWP	=	13		# Displays Help topics for AreA Hangul users.
wdHelpIchitaro	=	11		# Displays Help topics for Ichitaro users.
wdHelpIndex	=	5		# Displays the  Help Topics dialog box.
wdHelpKeyboard	=	6		# Displays keyboard shortcuts associated with help.
wdHelpPE2	=	12		# Displays Help topics for IBM Personal Editor 2 users.
wdHelpPSSHelp	=	7		# Displays product support information
wdHelpQuickPreview	=	8		# Displays quick previews.
wdHelpSearch	=	9		# Displays the  Help Topics dialog box.
wdHelpUsingHelp	=	10		# Displays a list of Help topics that describe how to use Help.

#values from word.wdhighansitext
#****************************
wdAutoDetectHighAnsiFarEast	=	2		# Microsoft Word interprets high-ANSI text as East Asian characters only if Word automatically detects East Asian language text.
wdHighAnsiIsFarEast	=	0		# Word doesn't interpret any high-ANSI text as East Asian characters.
wdHighAnsiIsHighAnsi	=	1		# Word interprets all high-ANSI text as East Asian characters.

#values from word.wdhorizontalinverticaltype
#****************************
wdHorizontalInVerticalFitInLine	=	1		# The horizontal text is sized to fit in the line of vertical text.
wdHorizontalInVerticalNone	=	0		# No formatting is applied to the horizontal text.
wdHorizontalInVerticalResizeLine	=	2		# The line of vertical text is sized to accommodate the horizontal text.

#values from word.wdhorizontallinealignment
#****************************
wdHorizontalLineAlignCenter	=	1		# Centered.
wdHorizontalLineAlignLeft	=	0		# Aligned to the left.
wdHorizontalLineAlignRight	=	2		# Aligned to the right.

#values from word.wdhorizontallinewidthtype
#****************************
wdHorizontalLineFixedWidth	=	-2		# Microsoft Word interprets the width (length) of the specified horizontal line as a fixed value (in points). This is the default value for horizontal lines added with the  AddHorizontalLine method. Setting the Width property for the InlineShape object associated with a horizontal line sets the WidthType property to this value.
wdHorizontalLinePercentWidth	=	-1		# Word interprets the width (length) of the specified horizontal line as a percentage of the screen width. This is the default value for horizontal lines added with the  AddHorizontalLineStandard method. Setting the PercentWidth property on a horizontal line sets the WidthType property to this value.

#values from word.wdimemode
#****************************
wdIMEModeAlpha	=	8		# Activates the IME in half-width Latin mode.
wdIMEModeAlphaFull	=	7		# Activates the IME in full-width Latin mode.
wdIMEModeHangul	=	10		# Activates the IME in half-width Hangul mode.
wdIMEModeHangulFull	=	9		# Activates the IME in full-width Hangul mode.
wdIMEModeHiragana	=	4		# Activates the IME in full-width hiragana mode.
wdIMEModeKatakana	=	5		# Activates the IME in full-width katakana mode.
wdIMEModeKatakanaHalf	=	6		# Activates the IME in half-width katakana mode.
wdIMEModeNoControl	=	0		# Does not change the IME mode.
wdIMEModeOff	=	2		# Disables the IME and activates Latin text entry.
wdIMEModeOn	=	1		# Activates the IME.

#values from word.wdindexfilter
#****************************
wdIndexFilterAiueo	=	1		# Japanese words use the AIUEO method of alphabetizing.
wdIndexFilterAkasatana	=	2		# Japanese words use Akasatana.
wdIndexFilterChosung	=	3		# Korean words use Chosung.
wdIndexFilterFull	=	6		# Korean words use Chosung.
wdIndexFilterLow	=	4		# Japanese words use Akasatana.
wdIndexFilterMedium	=	5		# Japanese words use the AIUEO method of alphabetizing.
wdIndexFilterNone	=	0		# No special filtering.

#values from word.wdindexformat
#****************************
wdIndexBulleted	=	4		# Bulleted.
wdIndexClassic	=	1		# Classic.
wdIndexFancy	=	2		# Fancy.
wdIndexFormal	=	5		# Formal.
wdIndexModern	=	3		# Modern.
wdIndexSimple	=	6		# Simple.
wdIndexTemplate	=	0		# From template.

#values from word.wdindexsortby
#****************************
wdIndexSortByStroke	=	0		# Sort by the number of strokes in a character.
wdIndexSortBySyllable	=	1		# Sort phonetically.

#values from word.wdindextype
#****************************
wdIndexIndent	=	0		# An indented index.
wdIndexRunin	=	1		# A run-in index.

#values from word.wdinformation
#****************************
wdActiveEndAdjustedPageNumber	=	1		# Returns the number of the page that contains the active end of the specified selection or range. If you set a starting page number or make other manual adjustments, returns the adjusted page number (unlike  wdActiveEndPageNumber).
wdActiveEndPageNumber	=	3		# Returns the number of the page that contains the active end of the specified selection or range, counting from the beginning of the document. Any manual adjustments to page numbering are disregarded (unlike  wdActiveEndAdjustedPageNumber).
wdActiveEndSectionNumber	=	2		# Returns the number of the section that contains the active end of the specified selection or range.
wdAtEndOfRowMarker	=	31		# Returns  True if the specified selection or range is at the end-of-row mark in a table.
wdCapsLock	=	21		# Returns  True if Caps Lock is in effect.
wdEndOfRangeColumnNumber	=	17		# Returns the table column number that contains the end of the specified selection or range.
wdEndOfRangeRowNumber	=	14		# Returns the table row number that contains the end of the specified selection or range.
wdFirstCharacterColumnNumber	=	9		# Returns the character position of the first character in the specified selection or range. If the selection or range is collapsed, the character number immediately to the right of the range or selection is returned (this is the same as the character column number displayed in the status bar after &quot;Col&quot;).
wdFirstCharacterLineNumber	=	10		# Returns the character position of the first character in the specified selection or range. If the selection or range is collapsed, the character number immediately to the right of the range or selection is returned (this is the same as the character line number displayed in the status bar after &quot;Ln&quot;).
wdFrameIsSelected	=	11		# Returns  True if the selection or range is an entire frame or text box.
wdHeaderFooterType	=	33		# Returns a value that indicates the type of header or footer that contains the specified selection or range. See the table in the remarks section for additional information.
wdHorizontalPositionRelativeToPage	=	5		# Returns the horizontal position of the specified selection or range; this is the distance from the left edge of the selection or range to the left edge of the page measured in points (1 point = 20 twips, 72 points = 1 inch). If the selection or range isn't within the screen area, returns ? 1.
wdHorizontalPositionRelativeToTextBoundary	=	7		# Returns the horizontal position of the specified selection or range relative to the left edge of the nearest text boundary enclosing it, in points (1 point = 20 twips, 72 points = 1 inch). If the selection or range isn't within the screen area, returns - 1.
wdInBibliography	=	42		# Returns  True if the specified selection or range is in a bibliography.
wdInCitation	=	43		# Returns  True if the specified selection or range is in a citation.
wdInClipboard	=	38		# For information about this constant, consult the language reference Help included with Microsoft Office Macintosh Edition.
wdInCommentPane	=	26		# Returns  True if the specified selection or range is in a comment pane.
wdInContentControl	=	46		# Returns  True if the specified selection or range is in a content control.
wdInCoverPage	=	41		# Returns  True if the specified selection or range is in a cover page.
wdInEndnote	=	36		# Returns  True if the specified selection or range is in an endnote area in print layout view or in the endnote pane in normal view.
wdInFieldCode	=	44		# Returns  True if the specified selection or range is in a field code.
wdInFieldResult	=	45		# Returns  True if the specified selection or range is in a field result.
wdInFootnote	=	35		# Returns  True if the specified selection or range is in a footnote area in print layout view or in the footnote pane in normal view.
wdInFootnoteEndnotePane	=	25		# Returns  True if the specified selection or range is in the footnote or endnote pane in normal view or in a footnote or endnote area in print layout view. For more information, see the descriptions of wdInFootnote and wdInEndnote in the preceding paragraphs.
wdInHeaderFooter	=	28		# Returns  True if the selection or range is in the header or footer pane or in a header or footer in print layout view.
wdInMasterDocument	=	34		# Returns  True if the selection or range is in a master document (that is, a document that contains at least one subdocument).
wdInWordMail	=	37		# Returns  True if the selection or range is in the header or footer pane or in a header or footer in print layout view.
wdMaximumNumberOfColumns	=	18		# Returns the greatest number of table columns within any row in the selection or range.
wdMaximumNumberOfRows	=	15		# Returns the greatest number of table rows within the table in the specified selection or range.
wdNumberOfPagesInDocument	=	4		# Returns the number of pages in the document associated with the selection or range.
wdNumLock	=	22		# Returns  True if Num Lock is in effect.
wdOverType	=	23		# Returns  True if Overtype mode is in effect. The Overtype property can be used to change the state of the Overtype mode.
wdReferenceOfType	=	32		# Returns a value that indicates where the selection is in relation to a footnote, endnote, or comment reference, as shown in the table in the remarks section.
wdRevisionMarking	=	24		# Returns  True if change tracking is in effect.
wdSelectionMode	=	20		# Returns a value that indicates the current selection mode, as shown in the following table.
wdStartOfRangeColumnNumber	=	16		# Returns the table column number that contains the beginning of the selection or range.
wdStartOfRangeRowNumber	=	13		# Returns the table row number that contains the beginning of the selection or range.
wdVerticalPositionRelativeToPage	=	6		# Returns the vertical position of the selection or range; this is the distance from the top edge of the selection to the top edge of the page measured in points (1 point = 20 twips, 72 points = 1 inch). If the selection isn't visible in the document window, returns ? 1.
wdVerticalPositionRelativeToTextBoundary	=	8		# Returns the vertical position of the selection or range relative to the top edge of the nearest text boundary enclosing it, in points (1 point = 20 twips, 72 points = 1 inch). This is useful for determining the position of the insertion point within a frame or table cell. If the selection isn't visible, returns ? 1.
wdWithInTable	=	12		# Returns  True if the selection is in a table.
wdZoomPercentage	=	19		# Returns the current percentage of magnification as set by the  Percentage property.

#values from word.wdinlineshapetype
#****************************
wdInlineShape3DModel	=	19		# 3D Model.
wdInlineShapeChart	=	12		# Inline chart.
wdInlineShapeDiagram	=	13		# Inline diagram.
wdInlineShapeEmbeddedOLEObject	=	1		# Embedded OLE object.
wdInlineShapeHorizontalLine	=	6		# Horizontal line.
wdInlineShapeLinked3DModel	=	20		# Linked 3D Model.
wdInlineShapeLinkedOLEObject	=	2		# Linked OLE object.
wdInlineShapeLinkedPicture	=	4		# Linked picture.
wdInlineShapeLinkedPictureHorizontalLine	=	8		# Linked picture with horizontal line.
wdInlineShapeLockedCanvas	=	14		# Locked inline shape canvas.
wdInlineShapeOLEControlObject	=	5		# OLE control object.
wdInlineShapeOWSAnchor	=	11		# OWS anchor.
wdInlineShapePicture	=	3		# Picture.
wdInlineShapePictureBullet	=	9		# Picture used as a bullet.
wdInlineShapePictureHorizontalLine	=	7		# Picture with horizontal line.
wdInlineShapeScriptAnchor	=	10		# Script anchor. Refers to anchor location for block of script stored with a document.
wdInlineShapeSmartArt	=	15		# A SmartArt graphic.
wdInlineShapeWebVideo	=	16		# A picture acting as a poster frame for a web video.

#values from word.wdinsertcells
#****************************
wdInsertCellsEntireColumn	=	3		# Inserts an entire column to the left of the column that contains the selection.
wdInsertCellsEntireRow	=	2		# Inserts an entire row above the row that contains the selection.
wdInsertCellsShiftDown	=	1		# Inserts new cells above the selected cells.
wdInsertCellsShiftRight	=	0		# Insert new cells to the left of the selected cells.

#values from word.wdinsertedtextmark
#****************************
wdInsertedTextMarkBold	=	1		# Inserted text is displayed in bold.
wdInsertedTextMarkColorOnly	=	5		# Inserted text is displayed in a specified color.
wdInsertedTextMarkDoubleUnderline	=	4		# Inserted text is marked up by using double-underline characters.
wdInsertedTextMarkItalic	=	2		# Inserted text is displayed in italic.
wdInsertedTextMarkNone	=	0		# Inserted text is not marked up.
wdInsertedTextMarkStrikeThrough	=	6		# Inserted text is marked up by using strikethrough characters.
wdInsertedTextMarkUnderline	=	3		# Inserted text is underlined.
wdInsertedTextMarkDoubleStrikeThrough	=	7		# Inserted text is marked up by using double-strikethrough characters.

#values from word.wdinternationalindex
#****************************
wd24HourClock	=	21		# Returns  True if you are using 24-hour time; returns False if you are using 12-hour time.
wdCurrencyCode	=	20		# Returns the currency symbol ($ in U.S. English).
wdDateSeparator	=	25		# Returns the date separator (/ in U.S. English).
wdDecimalSeparator	=	18		# Returns the decimal separator (. in U.S. English).
wdInternationalAM	=	22		# Returns the string used to indicate morning hours (for example, 10 A.M).
wdInternationalPM	=	23		# Returns the string used to indicate afternoon and evening hours (for example, 2 P.M).
wdListSeparator	=	17		# Returns the list separator (, in U.S. English).
wdProductLanguageID	=	26		# Returns the language version of Word.
wdThousandsSeparator	=	19		# Returns the thousands separator (, in U.S. English).
wdTimeSeparator	=	24		# Returns the time separator (: in U.S. English).

#values from word.wdjustificationmode
#****************************
wdJustificationModeCompress	=	1		# Compress.
wdJustificationModeCompressKana	=	2		# Compress, using rules of the kana syllabaries, Hiragana and Katakana.
wdJustificationModeExpand	=	0		# Expand.

#values from word.wdkana
#****************************
wdKanaHiragana	=	9		# The text is formatted as Hiragana.
wdKanaKatakana	=	8		# The text is formatted as Katakana.

#values from word.wdkey
#****************************
wdKey0	=	48		# The 0 key.
wdKey1	=	49		# The 1 key.
wdKey2	=	50		# The 2 key.
wdKey3	=	51		# The 3 key.
wdKey4	=	52		# The 4 key.
wdKey5	=	53		# The 5 key.
wdKey6	=	54		# The 6 key.
wdKey7	=	55		# The 7 key.
wdKey8	=	56		# The 8 key.
wdKey9	=	57		# The 9 key.
wdKeyA	=	65		# The A key.
wdKeyAlt	=	1024		# The ALT key.
wdKeyB	=	66		# The B key.
wdKeyBackSingleQuote	=	192		# The ` key.
wdKeyBackSlash	=	220		# The \ key.
wdKeyBackspace	=	8		# The BACKSPACE key.
wdKeyC	=	67		# The C key.
wdKeyCloseSquareBrace	=	221		# The ] key.
wdKeyComma	=	188		# The , key.
wdKeyCommand	=	512		# The Windows command key or Macintosh COMMAND key.
wdKeyControl	=	512		# The CTRL key.
wdKeyD	=	68		# The D key.
wdKeyDelete	=	46		# The DELETE key.
wdKeyE	=	69		# The E key.
wdKeyEnd	=	35		# The END key.
wdKeyEquals	=	187		# The = key.
wdKeyEsc	=	27		# The ESC key.
wdKeyF	=	70		# The F key.
wdKeyF1	=	112		# The F1 key.
wdKeyF10	=	121		# The F10 key.
wdKeyF11	=	122		# The F11 key.
wdKeyF12	=	123		# The F12 key.
wdKeyF13	=	124		# The F13 key.
wdKeyF14	=	125		# The F14 key.
wdKeyF15	=	126		# The F15 key.
wdKeyF16	=	127		# The F16 key.
wdKeyF2	=	113		# The F2 key.
wdKeyF3	=	114		# The F3 key.
wdKeyF4	=	115		# The F4 key.
wdKeyF5	=	116		# The F5 key.
wdKeyF6	=	117		# The F6 key.
wdKeyF7	=	118		# The F7 key.
wdKeyF8	=	119		# The F8 key.
wdKeyF9	=	120		# The F9 key.
wdKeyG	=	71		# The G key.
wdKeyH	=	72		# The H key.
wdKeyHome	=	36		# The HOME key.
wdKeyHyphen	=	189		# The - key.
wdKeyI	=	73		# The I key.
wdKeyInsert	=	45		# The INSERT key.
wdKeyJ	=	74		# The J key.
wdKeyK	=	75		# The K key.
wdKeyL	=	76		# The L key.
wdKeyM	=	77		# The M key.
wdKeyN	=	78		# The N key.
wdKeyNumeric0	=	96		# The 0 key.
wdKeyNumeric1	=	97		# The 1 key.
wdKeyNumeric2	=	98		# The 2 key.
wdKeyNumeric3	=	99		# The 3 key.
wdKeyNumeric4	=	100		# The 4 key.
wdKeyNumeric5	=	101		# The 5 key.
wdKeyNumeric5Special	=	12		# .
wdKeyNumeric6	=	102		# The 6 key.
wdKeyNumeric7	=	103		# The 7 key.
wdKeyNumeric8	=	104		# The 8 key.
wdKeyNumeric9	=	105		# The 9 key.
wdKeyNumericAdd	=	107		# The + key on the numeric keypad.
wdKeyNumericDecimal	=	110		# The . key on the numeric keypad.
wdKeyNumericDivide	=	111		# The / key on the numeric keypad.
wdKeyNumericMultiply	=	106		# The * key on the numeric keypad.
wdKeyNumericSubtract	=	109		# The - key on the numeric keypad.
wdKeyO	=	79		# The O key.
wdKeyOpenSquareBrace	=	219		# The [ key.
wdKeyOption	=	1024		# The mouse option key or Macintosh OPTION key.
wdKeyP	=	80		# The P key.
wdKeyPageDown	=	34		# The PAGE DOWN key.
wdKeyPageUp	=	33		# The PAGE UP key.
wdKeyPause	=	19		# The PAUSE key.
wdKeyPeriod	=	190		# The . key.
wdKeyQ	=	81		# The Q key.
wdKeyR	=	82		# The R key.
wdKeyReturn	=	13		# The ENTER or RETURN key.
wdKeyS	=	83		# The S key.
wdKeyScrollLock	=	145		# The SCROLL LOCK key.
wdKeySemiColon	=	186		# The ; key.
wdKeyShift	=	256		# The SHIFT key.
wdKeySingleQuote	=	222		# The ' key.
wdKeySlash	=	191		# The / key.
wdKeySpacebar	=	32		# The SPACEBAR key.
wdKeyT	=	84		# The T key.
wdKeyTab	=	9		# The TAB key.
wdKeyU	=	85		# The U key.
wdKeyV	=	86		# The V key.
wdKeyW	=	87		# The W key.
wdKeyX	=	88		# The X key.
wdKeyY	=	89		# The Y key.
wdKeyZ	=	90		# The Z key.
wdNoKey	=	255		# No key.

#values from word.wdkeycategory
#****************************
wdKeyCategoryAutoText	=	4		# Key is assigned to autotext.
wdKeyCategoryCommand	=	1		# Key is assigned to a command.
wdKeyCategoryDisable	=	0		# Key is disabled.
wdKeyCategoryFont	=	3		# Key is assigned to a font.
wdKeyCategoryMacro	=	2		# Key is assigned to a macro.
wdKeyCategoryNil	=	-1		# Key is not assigned.
wdKeyCategoryPrefix	=	7		# Key is assigned to a prefix.
wdKeyCategoryStyle	=	5		# Key is assigned to a style.
wdKeyCategorySymbol	=	6		# Key is assigned to a symbol.

#values from word.wdlanguageid
#****************************
wdAfrikaans	=	1078		# African language.
wdAlbanian	=	1052		# Albanian language.
wdAmharic	=	1118		# Amharic language.
wdArabic	=	1025		# Arabic language.
wdArabicAlgeria	=	5121		# Arabic Algerian language.
wdArabicBahrain	=	15361		# Arabic Bahraini language.
wdArabicEgypt	=	3073		# Arabic Egyptian language.
wdArabicIraq	=	2049		# Arabic Iraqi language.
wdArabicJordan	=	11265		# Arabic Jordanian language.
wdArabicKuwait	=	13313		# Arabic Kuwaiti language.
wdArabicLebanon	=	12289		# Arabic Lebanese language.
wdArabicLibya	=	4097		# Arabic Libyan language.
wdArabicMorocco	=	6145		# Arabic Moroccan language.
wdArabicOman	=	8193		# Arabic Omani language.
wdArabicQatar	=	16385		# Arabic Qatari language.
wdArabicSyria	=	10241		# Arabic Syrian language.
wdArabicTunisia	=	7169		# Arabic Tunisian language.
wdArabicUAE	=	14337		# Arabic United Arab Emirates language.
wdArabicYemen	=	9217		# Arabic Yemeni language.
wdArmenian	=	1067		# Armenian language.
wdAssamese	=	1101		# Assamese language.
wdAzeriCyrillic	=	2092		# Azeri Cyrillic language.
wdAzeriLatin	=	1068		# Azeri Latin language.
wdBasque	=	1069		# Basque (Basque).
wdBelgianDutch	=	2067		# Belgian Dutch language.
wdBelgianFrench	=	2060		# Belgian French language.
wdBengali	=	1093		# Bengali language.
wdBulgarian	=	1026		# Bulgarian language.
wdBurmese	=	1109		# Burmese language.
wdByelorussian	=	1059		# Belarusian language.
wdCatalan	=	1027		# Catalan language.
wdCherokee	=	1116		# Cherokee language.
wdChineseHongKongSAR	=	3076		# Chinese Hong Kong SAR language.
wdChineseMacaoSAR	=	5124		# Chinese Macao SAR language.
wdChineseSingapore	=	4100		# Chinese Singapore language.
wdCroatian	=	1050		# Croatian language.
wdCzech	=	1029		# Czech language.
wdDanish	=	1030		# Danish language.
wdDivehi	=	1125		# Divehi language.
wdDutch	=	1043		# Dutch language.
wdEdo	=	1126		# Edo language.
wdEnglishAUS	=	3081		# Australian English language.
wdEnglishBelize	=	10249		# Belize English language.
wdEnglishCanadian	=	4105		# Canadian English language.
wdEnglishCaribbean	=	9225		# Caribbean English language.
wdEnglishIndonesia	=	14345		# Indonesian English language.
wdEnglishIreland	=	6153		# Irish English language.
wdEnglishJamaica	=	8201		# Jamaican English language.
wdEnglishNewZealand	=	5129		# New Zealand English language.
wdEnglishPhilippines	=	13321		# Filipino English language.
wdEnglishSouthAfrica	=	7177		# South African English language.
wdEnglishTrinidadTobago	=	11273		# Tobago Trinidad English language.
wdEnglishUK	=	2057		# United Kingdom English language.
wdEnglishUS	=	1033		# United States English language.
wdEnglishZimbabwe	=	12297		# Zimbabwe English language.
wdEstonian	=	1061		# Estonian language.
wdFaeroese	=	1080		# Faeroese language.
wdFilipino	=	1124		# Filipino language.
wdFinnish	=	1035		# Finnish language.
wdFrench	=	1036		# French language.
wdFrenchCameroon	=	11276		# French Cameroon language.
wdFrenchCanadian	=	3084		# French Canadian language.
wdFrenchCongoDRC	=	9228		# French (Congo (DRC)) language.
wdFrenchCotedIvoire	=	12300		# French Cote d'Ivoire language.
wdFrenchHaiti	=	15372		# French Haiti language.
wdFrenchLuxembourg	=	5132		# French Luxembourg language.
wdFrenchMali	=	13324		# French Mali language.
wdFrenchMonaco	=	6156		# French Monaco language.
wdFrenchMorocco	=	14348		# French Morocco language.
wdFrenchReunion	=	8204		# French Reunion language.
wdFrenchSenegal	=	10252		# French Senegal language.
wdFrenchWestIndies	=	7180		# French West Indies language.
wdFrisianNetherlands	=	1122		# Frisian Netherlands language.
wdFulfulde	=	1127		# Fulfulde language.
wdGaelicIreland	=	2108		# Irish (Irish) language.
wdGaelicScotland	=	1084		# Scottish Gaelic language.
wdGalician	=	1110		# Galician language.
wdGeorgian	=	1079		# Georgian language.
wdGerman	=	1031		# German language.
wdGermanAustria	=	3079		# German Austrian language.
wdGermanLiechtenstein	=	5127		# German Liechtenstein language.
wdGermanLuxembourg	=	4103		# German Luxembourg language.
wdGreek	=	1032		# Greek language.
wdGuarani	=	1140		# Guarani language.
wdGujarati	=	1095		# Gujarati language.
wdHausa	=	1128		# Hausa language.
wdHawaiian	=	1141		# Hawaiian language.
wdHebrew	=	1037		# Hebrew language.
wdHindi	=	1081		# Hindi language.
wdHungarian	=	1038		# Hungarian language.
wdIbibio	=	1129		# Ibibio language.
wdIcelandic	=	1039		# Icelandic language.
wdIgbo	=	1136		# Igbo language.
wdIndonesian	=	1057		# Indonesian language.
wdInuktitut	=	1117		# Inuktitut language.
wdItalian	=	1040		# Italian language.
wdJapanese	=	1041		# Japanese language.
wdKannada	=	1099		# Kannada language.
wdKanuri	=	1137		# Kanuri language.
wdKashmiri	=	1120		# Kashmiri language.
wdKazakh	=	1087		# Kazakh language.
wdKhmer	=	1107		# Khmer language.
wdKirghiz	=	1088		# Kirghiz language.
wdKonkani	=	1111		# Konkani language.
wdKorean	=	1042		# Korean language.
wdKyrgyz	=	1088		# Kyrgyz language.
wdLanguageNone	=	0		# No specified language.
wdLao	=	1108		# Lao language.
wdLatin	=	1142		# Latin language.
wdLatvian	=	1062		# Latvian language.
wdLithuanian	=	1063		# Lithuanian language.
wdMacedonianFYROM	=	1071		# Macedonian (FYROM) language.
wdMalayalam	=	1100		# Malayalam language.
wdMalayBruneiDarussalam	=	2110		# Malay Brunei Darussalam language.
wdMalaysian	=	1086		# Malaysian language.
wdMaltese	=	1082		# Maltese language.
wdManipuri	=	1112		# Manipuri language.
wdMarathi	=	1102		# Marathi language.
wdMexicanSpanish	=	2058		# Mexican Spanish language.
wdMongolian	=	1104		# Mongolian language.
wdNepali	=	1121		# Nepali language.
wdNoProofing	=	1024		# Disables proofing if the language ID identifies a language in which an object is grammatically validated using the Microsoft Word proofing tools.
wdNorwegianBokmol	=	1044		# Norwegian Bokmol language.
wdNorwegianNynorsk	=	2068		# Norwegian Nynorsk language.
wdOriya	=	1096		# Oriya language.
wdOromo	=	1138		# Oromo language.
wdPashto	=	1123		# Pashto language.
wdPersian	=	1065		# Persian language.
wdPolish	=	1045		# Polish language.
wdPortuguese	=	2070		# Portuguese language.
wdPortugueseBrazil	=	1046		# Portuguese (Brazil) language.
wdPunjabi	=	1094		# Punjabi language.
wdRhaetoRomanic	=	1047		# Rhaeto Romanic language.
wdRomanian	=	1048		# Romanian language.
wdRomanianMoldova	=	2072		# Romanian Moldova language.
wdRussian	=	1049		# Russian language.
wdRussianMoldova	=	2073		# Russian Moldova language.
wdSamiLappish	=	1083		# Sami Lappish language.
wdSanskrit	=	1103		# Sanskrit language.
wdSerbianCyrillic	=	3098		# Serbian Cyrillic language.
wdSerbianLatin	=	2074		# Serbian Latin language.
wdSesotho	=	1072		# Sesotho language.
wdSimplifiedChinese	=	2052		# Simplified Chinese language.
wdSindhi	=	1113		# Sindhi language.
wdSindhiPakistan	=	2137		# Sindhi (Pakistan) language.
wdSinhalese	=	1115		# Sinhalese language.
wdSlovak	=	1051		# Slovakian language.
wdSlovenian	=	1060		# Slovenian language.
wdSomali	=	1143		# Somali language.
wdSorbian	=	1070		# Sorbian language.
wdSpanish	=	1034		# Spanish language.
wdSpanishArgentina	=	11274		# Spanish Argentina language.
wdSpanishBolivia	=	16394		# Spanish Bolivian language.
wdSpanishChile	=	13322		# Spanish Chilean language.
wdSpanishColombia	=	9226		# Spanish Colombian language.
wdSpanishCostaRica	=	5130		# Spanish Costa Rican language.
wdSpanishDominicanRepublic	=	7178		# Spanish Dominican Republic language.
wdSpanishEcuador	=	12298		# Spanish Ecuadorian language.
wdSpanishElSalvador	=	17418		# Spanish El Salvadorian language.
wdSpanishGuatemala	=	4106		# Spanish Guatemala language.
wdSpanishHonduras	=	18442		# Spanish Honduran language.
wdSpanishModernSort	=	3082		# Spanish Modern Sort language.
wdSpanishNicaragua	=	19466		# Spanish Nicaraguan language.
wdSpanishPanama	=	6154		# Spanish Panamanian language.
wdSpanishParaguay	=	15370		# Spanish Paraguayan language.
wdSpanishPeru	=	10250		# Spanish Peruvian language.
wdSpanishPuertoRico	=	20490		# Spanish Puerto Rican language.
wdSpanishUruguay	=	14346		# Spanish Uruguayan language.
wdSpanishVenezuela	=	8202		# Spanish Venezuelan language.
wdSutu	=	1072		# Sutu language.
wdSwahili	=	1089		# Swahili language.
wdSwedish	=	1053		# Swedish language.
wdSwedishFinland	=	2077		# Swedish Finnish language.
wdSwissFrench	=	4108		# Swiss French language.
wdSwissGerman	=	2055		# Swiss German language.
wdSwissItalian	=	2064		# Swiss Italian language.
wdSyriac	=	1114		# Syriac language.
wdTajik	=	1064		# Tajik language.
wdTamazight	=	1119		# Tamazight language.
wdTamazightLatin	=	2143		# Tamazight Latin language.
wdTamil	=	1097		# Tamil language.
wdTatar	=	1092		# Tatar language.
wdTelugu	=	1098		# Telugu language.
wdThai	=	1054		# Thai language.
wdTibetan	=	1105		# Tibetan language.
wdTigrignaEritrea	=	2163		# Tigrigna Eritrea language.
wdTigrignaEthiopic	=	1139		# Tigrigna Ethiopic language.
wdTraditionalChinese	=	1028		# Traditional Chinese language.
wdTsonga	=	1073		# Tsonga language.
wdTswana	=	1074		# Tswana language.
wdTurkish	=	1055		# Turkish language.
wdTurkmen	=	1090		# Turkmen language.
wdUkrainian	=	1058		# Ukrainian language.
wdUrdu	=	1056		# Urdu language.
wdUzbekCyrillic	=	2115		# Uzbek Cyrillic language.
wdUzbekLatin	=	1091		# Uzbek Latin language.
wdVenda	=	1075		# Venda language.
wdVietnamese	=	1066		# Vietnamese language.
wdWelsh	=	1106		# Welsh language.
wdXhosa	=	1076		# Xhosa language.
wdYi	=	1144		# Yi language.
wdYiddish	=	1085		# Yiddish language.
wdYoruba	=	1130		# Yoruba language.
wdZulu	=	1077		# Zulu language.

#values from word.wdlayoutmode
#****************************
wdLayoutModeDefault	=	0		# No grid is used to lay out text.
wdLayoutModeGenko	=	3		# Text is laid out on a grid; the user specifies the number of lines and the number of characters per line. As the user types, Microsoft Word automatically aligns characters with gridlines.
wdLayoutModeGrid	=	1		# Text is laid out on a grid; the user specifies the number of lines and the number of characters per line. As the user types, Microsoft Word doesn't automatically align characters with gridlines.
wdLayoutModeLineGrid	=	2		# Text is laid out on a grid; the user specifies the number of lines, but not the number of characters per line.

#values from word.wdletterheadlocation
#****************************
wdLetterBottom	=	1		# At the bottom of the letter.
wdLetterLeft	=	2		# To the left of the letter.
wdLetterRight	=	3		# To the right of the letter.
wdLetterTop	=	0		# At the top of the letter.

#values from word.wdletterstyle
#****************************
wdFullBlock	=	0		# Full block.
wdModifiedBlock	=	1		# Modified block.
wdSemiBlock	=	2		# Semi-block.

#values from word.wdligatures
#****************************
wdLigaturesAll	=	15		# Applies all types of ligatures to the font.
wdLigaturesContextual	=	2		# Applies contextual ligatures to the font. Contextual ligatures are often designed to enhance readability, but may also be solely ornamental. Contextual ligatures may also be contextual alternates.
wdLigaturesContextualDiscretional	=	10		# Applies contextual and discretional ligatures to the font.
wdLigaturesContextualHistorical	=	6		# Applies contextual and historical ligatures to the font.
wdLigaturesContextualHistoricalDiscretional	=	14		# Applies contextual, historical, and discretional ligatures to a font.
wdLigaturesDiscretional	=	8		# Applies discretional ligatures to the font. Discretional ligatures are most often designed to be ornamental at the discretion of the type developer.
wdLigaturesHistorical	=	4		# Applies historical ligatures to the font. Historical ligatures are similar to standard ligatures in that they were originally intended to improve the readability of the font, but may look archaic to the modern reader.
wdLigaturesHistoricalDiscretional	=	12		# Applies historical and discretional ligatures to the font.
wdLigaturesNone	=	0		# Does not apply any ligatures to the font.
wdLigaturesStandard	=	1		# Applies standard ligatures to the font. Standard ligatures are designed to enhance readability. Standard ligatures in Latin languages include &quot;fi&quot;, &quot;fl&quot;, and &quot;ff&quot;, for example.
wdLigaturesStandardContextual	=	3		# Applies standard and contextual ligatures to the font.
wdLigaturesStandardContextualDiscretional	=	11		# Applies standard, contextual and discretional ligatures to the font.
wdLigaturesStandardContextualHistorical	=	7		# Applies standard, contextual, and historical ligatures to the font.
wdLigaturesStandardDiscretional	=	9		# Applies standard and discretional ligatures to the font.
wdLigaturesStandardHistorical	=	5		# Applies standard and historical ligatures to the font.
wdLigaturesStandardHistoricalDiscretional	=	13		# Applies standard historical and discretional ligatures to the font.

#values from word.wdlineendingtype
#****************************
wdCRLF	=	0		# Carriage return plus line feed.
wdCROnly	=	1		# Carriage return only.
wdLFCR	=	3		# Line feed plus carriage return.
wdLFOnly	=	2		# Line feed only.
wdLSPS	=	4		# Not supported.

#values from word.wdlinespacing
#****************************
wdLineSpace1pt5	=	1		# Space-and-a-half line spacing. Spacing is equivalent to the current font size plus 6 points.
wdLineSpaceAtLeast	=	3		# Line spacing is always at least a specified amount. The amount is specified separately.
wdLineSpaceDouble	=	2		# Double spaced.
wdLineSpaceExactly	=	4		# Line spacing is only the exact maximum amount of space required. This setting commonly uses less space than single spacing.
wdLineSpaceMultiple	=	5		# Line spacing determined by the number of lines indicated.
wdLineSpaceSingle	=	0		# Single spaced. default

#values from word.wdlinestyle
#****************************
wdLineStyleDashDot	=	5		# A dash followed by a dot.
wdLineStyleDashDotDot	=	6		# A dash followed by two dots.
wdLineStyleDashDotStroked	=	20		# A dash followed by a dot stroke, thus rendering a border similar to a barber pole.
wdLineStyleDashLargeGap	=	4		# A dash followed by a large gap.
wdLineStyleDashSmallGap	=	3		# A dash followed by a small gap.
wdLineStyleDot	=	2		# Dots.
wdLineStyleDouble	=	7		# Double solid lines.
wdLineStyleDoubleWavy	=	19		# Double wavy solid lines.
wdLineStyleEmboss3D	=	21		# The border appears to have a 3D embossed look.
wdLineStyleEngrave3D	=	22		# The border appears to have a 3D engraved look.
wdLineStyleInset	=	24		# The border appears to be inset.
wdLineStyleNone	=	0		# No border.
wdLineStyleOutset	=	23		# The border appears to be outset.
wdLineStyleSingle	=	1		# A single solid line.
wdLineStyleSingleWavy	=	18		# A single wavy solid line.
wdLineStyleThickThinLargeGap	=	16		# An internal single thick solid line surrounded by a single thin solid line with a large gap between them.
wdLineStyleThickThinMedGap	=	13		# An internal single thick solid line surrounded by a single thin solid line with a medium gap between them.
wdLineStyleThickThinSmallGap	=	10		# An internal single thick solid line surrounded by a single thin solid line with a small gap between them.
wdLineStyleThinThickLargeGap	=	15		# An internal single thin solid line surrounded by a single thick solid line with a large gap between them.
wdLineStyleThinThickMedGap	=	12		# An internal single thin solid line surrounded by a single thick solid line with a medium gap between them.
wdLineStyleThinThickSmallGap	=	9		# An internal single thin solid line surrounded by a single thick solid line with a small gap between them.
wdLineStyleThinThickThinLargeGap	=	17		# An internal single thin solid line surrounded by a single thick solid line surrounded by a single thin solid line with a large gap between all lines.
wdLineStyleThinThickThinMedGap	=	14		# An internal single thin solid line surrounded by a single thick solid line surrounded by a single thin solid line with a medium gap between all lines.
wdLineStyleThinThickThinSmallGap	=	11		# An internal single thin solid line surrounded by a single thick solid line surrounded by a single thin solid line with a small gap between all lines.
wdLineStyleTriple	=	8		# Three solid thin lines.

#values from word.wdlinetype
#****************************
wdTableRow	=	1		# A table row.
wdTextLine	=	0		# A line of text in the body of the document.

#values from word.wdlinewidth
#****************************
wdLineWidth025pt	=	2		# 0.25 point.
wdLineWidth050pt	=	4		# 0.50 point.
wdLineWidth075pt	=	6		# 0.75 point.
wdLineWidth100pt	=	8		# 1.00 point. default.
wdLineWidth150pt	=	12		# 1.50 points.
wdLineWidth225pt	=	18		# 2.25 points.
wdLineWidth300pt	=	24		# 3.00 points.
wdLineWidth450pt	=	36		# 4.50 points.
wdLineWidth600pt	=	48		# 6.00 points.

#values from word.wdlinktype
#****************************
wdLinkTypeChart	=	8		# Microsoft Excel chart.
wdLinkTypeDDE	=	6		# Dynamic Data Exchange.
wdLinkTypeDDEAuto	=	7		# DDE automatic.
wdLinkTypeImport	=	5		# Import file.
wdLinkTypeInclude	=	4		# Include file.
wdLinkTypeOLE	=	0		# OLE object.
wdLinkTypePicture	=	1		# Picture.
wdLinkTypeReference	=	3		# Reference library.
wdLinkTypeText	=	2		# Text.

#values from word.wdlistapplyto
#****************************
wdListApplyToSelection	=	2		# Selection.
wdListApplyToThisPointForward	=	1		# From cursor insertion point to end of list.
wdListApplyToWholeList	=	0		# Entire list.

#values from word.wdlistgallerytype
#****************************
wdBulletGallery	=	1		# Bulleted list.
wdNumberGallery	=	2		# Numbered list.
wdOutlineNumberGallery	=	3		# Outline numbered list.

#values from word.wdlistlevelalignment
#****************************
wdListLevelAlignCenter	=	1		# Center-aligned.
wdListLevelAlignLeft	=	0		# Left-aligned.
wdListLevelAlignRight	=	2		# Right-aligned.

#values from word.wdlistnumberstyle
#****************************
wdListNumberStyleAiueo	=	20		# Aiueo numeric style.
wdListNumberStyleAiueoHalfWidth	=	12		# Aiueo half-width numeric style.
wdListNumberStyleArabic	=	0		# Arabic numeric style.
wdListNumberStyleArabic1	=	46		# Arabic 1 numeric style.
wdListNumberStyleArabic2	=	48		# Arabic 2 numeric style.
wdListNumberStyleArabicFullWidth	=	14		# Arabic full-width numeric style.
wdListNumberStyleArabicLZ	=	22		# Arabic LZ numeric style.
wdListNumberStyleArabicLZ2	=	62		# Arabic LZ2 numeric style.
wdListNumberStyleArabicLZ3	=	63		# Arabic LZ3 numeric style.
wdListNumberStyleArabicLZ4	=	64		# Arabic LZ4 numeric style.
wdListNumberStyleBullet	=	23		# Bullet style.
wdListNumberStyleCardinalText	=	6		# Cardinal text style.
wdListNumberStyleChosung	=	25		# Chosung style.
wdListNumberStyleGanada	=	24		# Ganada style.
wdListNumberStyleGBNum1	=	26		# GB numeric 1 style.
wdListNumberStyleGBNum2	=	27		# GB numeric 2 style.
wdListNumberStyleGBNum3	=	28		# GB numeric 3 style.
wdListNumberStyleGBNum4	=	29		# GB numeric 4 style.
wdListNumberStyleHangul	=	43		# Hanqul style.
wdListNumberStyleHanja	=	44		# Hanja style.
wdListNumberStyleHanjaRead	=	41		# Hanja Read style.
wdListNumberStyleHanjaReadDigit	=	42		# Hanja Read Digit style.
wdListNumberStyleHebrew1	=	45		# Hebrew 1 style.
wdListNumberStyleHebrew2	=	47		# Hebrew 2 style.
wdListNumberStyleHindiArabic	=	51		# Hindi Arabic style.
wdListNumberStyleHindiCardinalText	=	52		# Hindi Cardinal text style.
wdListNumberStyleHindiLetter1	=	49		# Hindi letter 1 style.
wdListNumberStyleHindiLetter2	=	50		# Hindi letter 2 style.
wdListNumberStyleIroha	=	21		# Iroha style.
wdListNumberStyleIrohaHalfWidth	=	13		# Iroha half width style.
wdListNumberStyleKanji	=	10		# Kanji style.
wdListNumberStyleKanjiDigit	=	11		# Kanji Digit style.
wdListNumberStyleKanjiTraditional	=	16		# Kanji traditional style.
wdListNumberStyleKanjiTraditional2	=	17		# Kanji traditional 2 style.
wdListNumberStyleLegal	=	253		# Legal style.
wdListNumberStyleLegalLZ	=	254		# Legal LZ style.
wdListNumberStyleLowercaseBulgarian	=	67		# Lowercase Bulgarian style.
wdListNumberStyleLowercaseGreek	=	60		# Lowercase Greek style.
wdListNumberStyleLowercaseLetter	=	4		# Lowercase letter style.
wdListNumberStyleLowercaseRoman	=	2		# Lowercase Roman style.
wdListNumberStyleLowercaseRussian	=	58		# Lowercase Russian style.
wdListNumberStyleLowercaseTurkish	=	65		# Lowercase Turkish style.
wdListNumberStyleNone	=	255		# No style applied.
wdListNumberStyleNumberInCircle	=	18		# Number in circle style.
wdListNumberStyleOrdinal	=	5		# Ordinal style.
wdListNumberStyleOrdinalText	=	7		# Ordinal text style.
wdListNumberStylePictureBullet	=	249		# Picture bullet style.
wdListNumberStyleSimpChinNum1	=	37		# Simplified Chinese numeric 1 style.
wdListNumberStyleSimpChinNum2	=	38		# Simplified Chinese numeric 2 style.
wdListNumberStyleSimpChinNum3	=	39		# Simplified Chinese numeric 3 style.
wdListNumberStyleSimpChinNum4	=	40		# Simplified Chinese numeric 4 style.
wdListNumberStyleThaiArabic	=	54		# Thai Arabic style.
wdListNumberStyleThaiCardinalText	=	55		# Thai Cardinal text style.
wdListNumberStyleThaiLetter	=	53		# Thai letter style.
wdListNumberStyleTradChinNum1	=	33		# Traditional Chinese numeric 1 style.
wdListNumberStyleTradChinNum2	=	34		# Traditional Chinese numeric 2 style.
wdListNumberStyleTradChinNum3	=	35		# Traditional Chinese numeric 3 style.
wdListNumberStyleTradChinNum4	=	36		# Traditional Chinese numeric 4 style.
wdListNumberStyleUppercaseBulgarian	=	68		# Uppercase Bulgarian style.
wdListNumberStyleUppercaseGreek	=	61		# Uppercase Greek style.
wdListNumberStyleUppercaseLetter	=	3		# Uppercase letter style.
wdListNumberStyleUppercaseRoman	=	1		# Uppercase Roman style.
wdListNumberStyleUppercaseRussian	=	59		# Uppercase Russian style.
wdListNumberStyleUppercaseTurkish	=	66		# Uppercase Turkish style.
wdListNumberStyleVietCardinalText	=	56		# Vietnamese Cardinal text style.
wdListNumberStyleZodiac1	=	30		# Zodiac 1 style.
wdListNumberStyleZodiac2	=	31		# Zodiac 2 style.
wdListNumberStyleZodiac3	=	32		# Zodiac 3 style.

#values from word.wdlisttype
#****************************
wdListBullet	=	2		# Bulleted list.
wdListListNumOnly	=	1		# ListNum fields that can be used in the body of a paragraph.
wdListMixedNumbering	=	5		# Mixed numeric list.
wdListNoNumbering	=	0		# List with no bullets, numbering, or outlining.
wdListOutlineNumbering	=	4		# Outlined list.
wdListPictureBullet	=	6		# Picture bulleted list.
wdListSimpleNumbering	=	3		# Simple numeric list.

#values from word.wdlocktype
#****************************
wdLockChanged	=	3		# Specifies a placeholder lock. A placeholder lock indicates that another user has removed their lock from the range, but the current user has not updated their view of the document by saving.
wdLockEphemeral	=	2		# Specifies an ephemeral lock. Word implicitly places an ephemeral lock on a range when a user begins editing a range in a document with coauthoring enabled.
wdLockNone	=	0		# Reserved for future use.
wdLockReservation	=	1		# Specifies a reservation lock. A reservation lock is explicitly created by a user through the  Block Authors button on the Review tab in Word.

#values from word.wdmailerpriority
#****************************
wdPriorityHigh	=	3		# Not supported.
wdPriorityLow	=	2		# Not supported.
wdPriorityNormal	=	1		# Not supported.

#values from word.wdmailmergeactiverecord
#****************************
wdFirstDataSourceRecord	=	-6		# The first record in the data source.
wdFirstRecord	=	-4		# The first record in the result set.
wdLastDataSourceRecord	=	-7		# The last record in the data source.
wdLastRecord	=	-5		# The last record in the result set.
wdNextDataSourceRecord	=	-8		# The next record in the data source.
wdNextRecord	=	-2		# The next record in the result set.
wdNoActiveRecord	=	-1		# No active record.
wdPreviousDataSourceRecord	=	-9		# The previous record in the data source.
wdPreviousRecord	=	-3		# The previous record in the result set.

#values from word.wdmailmergecomparison
#****************************
wdMergeIfEqual	=	0		# A value is output if the mail merge field is equal to a value.
wdMergeIfGreaterThan	=	3		# A value is output if the mail merge field is greater than a value.
wdMergeIfGreaterThanOrEqual	=	5		# A value is output if the mail merge field is greater than or equal to a value.
wdMergeIfIsBlank	=	6		# A value is output if the mail merge field is blank.
wdMergeIfIsNotBlank	=	7		# A value is output if the mail merge field is not blank.
wdMergeIfLessThan	=	2		# A value is output if the mail merge field is less than a value.
wdMergeIfLessThanOrEqual	=	4		# A value is output if the mail merge field is less than or equal to a value.
wdMergeIfNotEqual	=	1		# A value is output if the mail merge field is not equal to a value.

#values from word.wdmailmergedatasource
#****************************
wdMergeInfoFromAccessDDE	=	1		# From Microsoft Access using Dynamic Data Exchange (DDE).
wdMergeInfoFromExcelDDE	=	2		# From Microsoft Excel using DDE.
wdMergeInfoFromMSQueryDDE	=	3		# From MSQuery using DDE.
wdMergeInfoFromODBC	=	4		# From an Open Database Connectivity (ODBC) connection.
wdMergeInfoFromODSO	=	5		# From an Office Data Source Object (ODSO).
wdMergeInfoFromWord	=	0		# From Microsoft Word.
wdNoMergeInfo	=	-1		# No merge information provided.

#values from word.wdmailmergedefaultrecord
#****************************
wdDefaultFirstRecord	=	1		# Use the first record in the result set as the default record.
wdDefaultLastRecord	=	-16		# Use the last record in the result set as the default record.

#values from word.wdmailmergedestination
#****************************
wdSendToEmail	=	2		# Send results to email recipient.
wdSendToFax	=	3		# Send results to fax recipient.
wdSendToNewDocument	=	0		# Send results to a new Word document.
wdSendToPrinter	=	1		# Send results to a printer.

#values from word.wdmailmergemailformat
#****************************
wdMailFormatHTML	=	1		# Sends mail merge email documents using HTML format.
wdMailFormatPlainText	=	0		# Sends mail merge email documents using plain text.

#values from word.wdmailmergemaindoctype
#****************************
wdCatalog	=	3		# Catalog.
wdDirectory	=	3		# Directory.
wdEMail	=	4		# Email message.
wdEnvelopes	=	2		# Envelope.
wdFax	=	5		# Fax.
wdFormLetters	=	0		# Form letter.
wdMailingLabels	=	1		# Mailing label.
wdNotAMergeDocument	=	-1		# Not a merge document.

#values from word.wdmailmergestate
#****************************
wdDataSource	=	5		# A data source with no main document.
wdMainAndDataSource	=	2		# A main document with an attached data source.
wdMainAndHeader	=	3		# A main document with an attached header source.
wdMainAndSourceAndHeader	=	4		# A main document with attached data source and header source.
wdMainDocumentOnly	=	1		# A main document with no data attached.
wdNormalDocument	=	0		# Document is not involved in a mail merge operation.

#values from word.wdmailsystem
#****************************
wdMAPI	=	1		# Standard Messaging Application Programming Interface (MAPI) mail system.
wdMAPIandPowerTalk	=	3		# Both a standard Messaging Application Programming Interface (MAPI) mail system and a PowerTalk mail system.
wdNoMailSystem	=	0		# No mail system.
wdPowerTalk	=	2		# PowerTalk mail system.

#values from word.wdmappeddatafields
#****************************
wdAddress1	=	10		# Address 1 field.
wdAddress2	=	11		# Address 2 field.
wdAddress3	=	29		# Address 3 field.
wdBusinessFax	=	17		# Business fax field.
wdBusinessPhone	=	16		# Business phone field.
wdCity	=	12		# City field.
wdCompany	=	9		# Company field.
wdCountryRegion	=	15		# Country/region field.
wdCourtesyTitle	=	2		# Courtesy title field.
wdDepartment	=	30		# Department field.
wdEmailAddress	=	20		# Email address field.
wdFirstName	=	3		# First name field.
wdHomeFax	=	19		# Home fax field.
wdHomePhone	=	18		# Home phone field.
wdJobTitle	=	8		# Job title field.
wdLastName	=	5		# Last name field.
wdMiddleName	=	4		# Middle name field.
wdNickname	=	7		# Nickname field.
wdPostalCode	=	14		# Postal code field.
wdRubyFirstName	=	27		# Ruby first name field.
wdRubyLastName	=	28		# Ruby last name field.
wdSpouseCourtesyTitle	=	22		# Spouse/partner courtesy title field.
wdSpouseFirstName	=	23		# Spouse/partner first name field.
wdSpouseLastName	=	25		# Spouse/partner last name field.
wdSpouseMiddleName	=	24		# Spouse/partner middle name field.
wdSpouseNickname	=	26		# Spouse/partner nickname field.
wdState	=	13		# State field.
wdSuffix	=	6		# Suffix field.
wdUniqueIdentifier	=	1		# Unique identifier field.
wdWebPageURL	=	21		# Web page uniform resource locator (URL) field.

#values from word.wdmeasurementunits
#****************************
wdCentimeters	=	1		# Centimeters.
wdInches	=	0		# Inches.
wdMillimeters	=	2		# Millimeters.
wdPicas	=	4		# Picas (commonly used in traditional typewriter font spacing).
wdPoints	=	3		# Points.

#values from word.wdmergeformatfrom
#****************************
wdMergeFormatFromOriginal	=	0		# Retains formatting from the original document.
wdMergeFormatFromPrompt	=	2		# Prompt the user for the document to use for formatting.
wdMergeFormatFromRevised	=	1		# Retains formatting from the revised document.

#values from word.wdmergesubtype
#****************************
wdMergeSubTypeAccess	=	1		# Microsoft Access.
wdMergeSubTypeOAL	=	2		# Office Address List.
wdMergeSubTypeOLEDBText	=	5		# OLE database.
wdMergeSubTypeOLEDBWord	=	3		# OLE database.
wdMergeSubTypeOther	=	0		# Other type of data source.
wdMergeSubTypeOutlook	=	6		# Microsoft Outlook.
wdMergeSubTypeWord	=	7		# Microsoft Word.
wdMergeSubTypeWord2000	=	8		# Microsoft Word 2000.
wdMergeSubTypeWorks	=	4		# Microsoft Works.

#values from word.wdmergetarget
#****************************
wdMergeTargetCurrent	=	1		# Merge into current document.
wdMergeTargetNew	=	2		# Merge into new document.
wdMergeTargetSelected	=	0		# Merge into selected document.

#values from word.wdmonthnames
#****************************
wdMonthNamesArabic	=	0		# Arabic format.
wdMonthNamesEnglish	=	1		# English format.
wdMonthNamesFrench	=	2		# French format.

#values from word.wdmovefromtextmark
#****************************
wdMoveFromTextMarkBold	=	6		# Marks moved text with bold formatting.
wdMoveFromTextMarkCaret	=	3		# Marks moved text with a caret.
wdMoveFromTextMarkColorOnly	=	10		# Marks moved text with color only. Use the  MoveFromTextColor property to set the color of moved text.
wdMoveFromTextMarkDoubleStrikeThrough	=	1		# Marks moved text with a double strikethrough.
wdMoveFromTextMarkDoubleUnderline	=	9		# Marks moved text with a double underline.
wdMoveFromTextMarkHidden	=	0		# Hides moved text.
wdMoveFromTextMarkItalic	=	7		# Marks moved text with italic formatting.
wdMoveFromTextMarkNone	=	5		# No special formatting for moved text.
wdMoveFromTextMarkPound	=	4		# Marks moved text with a pound (number) sign.
wdMoveFromTextMarkStrikeThrough	=	2		# Marks moved text with a strikethrough.
wdMoveFromTextMarkUnderline	=	8		# Underlines moved text.

#values from word.wdmovementtype
#****************************
wdExtend	=	1		# The end of the selection is extended to the end of the specified unit.
wdMove	=	0		# The selection is collapsed to an insertion point and moved to the end of the specified unit. Default.

#values from word.wdmovetotextmark
#****************************
wdMoveToTextMarkBold	=	1		# Marks moved text with bold formatting.
wdMoveToTextMarkColorOnly	=	5		# Marks moved text with color only. Use the  MoveToTextColor property to set the color of moved text.
wdMoveToTextMarkDoubleStrikeThrough	=	7		# Moved text is marked with a double strikethrough.
wdMoveToTextMarkDoubleUnderline	=	4		# Moved text is marked with a double underline.
wdMoveToTextMarkItalic	=	2		# Marks moved text with italic formatting.
wdMoveToTextMarkNone	=	0		# No special formatting for moved text.
wdMoveToTextMarkStrikeThrough	=	6		# Moved text is marked with a strikethrough.
wdMoveToTextMarkUnderline	=	3		# Underlines moved text.

#values from word.wdmultiplewordconversionsmode
#****************************
wdHangulToHanja	=	0		# Hangul to Hanja.
wdHanjaToHangul	=	1		# Hanja to Hangul.

#values from word.wdnewdocumenttype
#****************************
wdNewBlankDocument	=	0		# Blank document.
wdNewEmailMessage	=	2		# Email message.
wdNewFrameset	=	3		# Frameset.
wdNewWebPage	=	1		# Web page.
wdNewXMLDocument	=	4		# XML document.

#values from word.wdnotenumberstyle
#****************************
wdNoteNumberStyleArabic	=	0		# Arabic number style.
wdNoteNumberStyleArabicFullWidth	=	14		# Arabic full-width number style.
wdNoteNumberStyleArabicLetter1	=	46		# Arabic letter style 1.
wdNoteNumberStyleArabicLetter2	=	48		# Arabic letter style 2.
wdNoteNumberStyleHanjaRead	=	41		# Hanja read number style.
wdNoteNumberStyleHanjaReadDigit	=	42		# Hanja read digit number style.
wdNoteNumberStyleHebrewLetter1	=	45		# Hebrew letter style 1.
wdNoteNumberStyleHebrewLetter2	=	47		# Hebrew letter style 2.
wdNoteNumberStyleHindiArabic	=	51		# Hindi Arabic number style.
wdNoteNumberStyleHindiCardinalText	=	52		# Hindi Cardinal text style.
wdNoteNumberStyleHindiLetter1	=	49		# Hindi letter style 1.
wdNoteNumberStyleHindiLetter2	=	50		# Hindi letter style 2.
wdNoteNumberStyleKanji	=	10		# Kanji number style.
wdNoteNumberStyleKanjiDigit	=	11		# Kanji digit number style.
wdNoteNumberStyleKanjiTraditional	=	16		# Kanji traditional number style.
wdNoteNumberStyleLowercaseLetter	=	4		# Lowercase letter style.
wdNoteNumberStyleLowercaseRoman	=	2		# Lowercase Roman number style.
wdNoteNumberStyleNumberInCircle	=	18		# Number in circle number style.
wdNoteNumberStyleSimpChinNum1	=	37		# Simplified Chinese number style 1.
wdNoteNumberStyleSimpChinNum2	=	38		# Simplified Chinese number style 2.
wdNoteNumberStyleSymbol	=	9		# Symbol number style.
wdNoteNumberStyleThaiArabic	=	54		# Thai Arabic number style.
wdNoteNumberStyleThaiCardinalText	=	55		# Thai Cardinal text style.
wdNoteNumberStyleThaiLetter	=	53		# Thai letter style.
wdNoteNumberStyleTradChinNum1	=	33		# Traditional Chinese number style 1.
wdNoteNumberStyleTradChinNum2	=	34		# Traditional Chinese number style 2.
wdNoteNumberStyleUppercaseLetter	=	3		# Uppercase letter style.
wdNoteNumberStyleUppercaseRoman	=	1		# Uppercase Roman number style.
wdNoteNumberStyleVietCardinalText	=	56		# Vietnamese Cardinal text style.

#values from word.wdnumberform
#****************************
wdNumberFormDefault	=	0		# Applies the default number form for the font.
wdNumberFormLining	=	1		# Applies the lining number form to the font.
wdNumberFormOldstyle	=	2		# Applies the &quot;old-style&quot; number form to the font.

#values from word.wdnumberingrule
#****************************
wdRestartContinuous	=	0		# Numbers are assigned continuously.
wdRestartPage	=	2		# Numbers are reset for each page.
wdRestartSection	=	1		# Numbers are reset for each section.

#values from word.wdnumberspacing
#****************************
wdNumberSpacingDefault	=	0		# Applies the default number spacing for the font.
wdNumberSpacingProportional	=	1		# Applies proportional number spacing to the font.
wdNumberSpacingTabular	=	2		# Applies tabular number spacing to the font.

#values from word.wdnumberstylewordbasicbidi
#****************************
wdCaptionNumberStyleBidiLetter1	=	49		# Not supported.
wdCaptionNumberStyleBidiLetter2	=	50		# Not supported.
wdListNumberStyleBidi1	=	49		# Not supported.
wdListNumberStyleBidi2	=	50		# Not supported.
wdNoteNumberStyleBidiLetter1	=	49		# Not supported.
wdNoteNumberStyleBidiLetter2	=	50		# Not supported.
wdPageNumberStyleBidiLetter1	=	49		# Not supported.
wdPageNumberStyleBidiLetter2	=	50		# Not supported.

#values from word.wdnumbertype
#****************************
wdNumberAllNumbers	=	3		# Default value for all other cases.
wdNumberListNum	=	2		# Default value for LISTNUM fields.
wdNumberParagraph	=	1		# Preset numbers you can add to paragraphs by selecting a template in the  Bullets and Numbering dialog box.

#values from word.wdoleplacement
#****************************
wdFloatOverText	=	1		# Float over text.
wdInLine	=	0		# In line with text.

#values from word.wdoletype
#****************************
wdOLEControl	=	2		# OLE control.
wdOLEEmbed	=	1		# Embedded OLE object.
wdOLELink	=	0		# Linked OLE object.

#values from word.wdoleverb
#****************************
wdOLEVerbDiscardUndoState	=	-6		# Forces the object to discard any undo state that it might be maintaining; note that the object remains active, however.
wdOLEVerbHide	=	-3		# Removes the object's user interface from view.
wdOLEVerbInPlaceActivate	=	-5		# Runs the object and installs its window, but doesn't install any user-interface tools.
wdOLEVerbOpen	=	-2		# Opens the object in a separate window.
wdOLEVerbPrimary	=	0		# Performs the verb that is invoked when the user double-clicks the object.
wdOLEVerbShow	=	-1		# Shows the object to the user for editing or viewing. Use it to show a newly inserted object for initial editing.
wdOLEVerbUIActivate	=	-4		# Activates the object in place and displays any user-interface tools that the object needs, such as menus or toolbars.

#values from word.wdomathbreakbin
#****************************
wdOMathBreakBinAfter	=	1		# Places the operator before a line break, at the end of the line.
wdOMathBreakBinBefore	=	0		# Places the operator after a line break, at the beginning of the following line.
wdOMathBreakBinRepeat	=	2		# Repeats the operator before a line break at the end of the line and after a line break at the beginning of the following line.

#values from word.wdomathbreaksub
#****************************
wdOMathBreakSubMinusMinus	=	0		# Repeats a minus sign that ends before a line break at the beginning of the next line. Default.
wdOMathBreakSubMinusPlus	=	2		# Inserts a minus sign at the end of the first line, before the line break, and a plus sign at the beginning of the following line, before the number.
wdOMathBreakSubPlusMinus	=	1		# Inserts a plus sign at the end of the first line, before the line break, and a minus sign at the beginning of the following line, before the number.

#values from word.wdomathfractype
#****************************
wdOMathFracBar	=	0		# Normal fraction bar.
wdOMathFracLin	=	3		# Show fraction inline.
wdOMathFracNoBar	=	1		# No fraction bar.
wdOMathFracSkw	=	2		# Skewed fraction bar.

#values from word.wdomathfunctiontype
#****************************
wdOMathFunctionAcc	=	1		# Equation accent mark.
wdOMathFunctionBar	=	2		# Equation fraction bar.
wdOMathFunctionBorderBox	=	4		# Border box.
wdOMathFunctionBox	=	3		# Box.
wdOMathFunctionDelim	=	5		# Equation delimiters.
wdOMathFunctionEqArray	=	6		# Equation array.
wdOMathFunctionFrac	=	7		# Equation fraction.
wdOMathFunctionFunc	=	8		# Equation function.
wdOMathFunctionGroupChar	=	9		# Group character.
wdOMathFunctionLimLow	=	10		# Equation lower limit.
wdOMathFunctionLimUpp	=	11		# Equation upper limit.
wdOMathFunctionMat	=	12		# Equation matrix.
wdOMathFunctionNary	=	13		# Equation N-ary operator.
wdOMathFunctionNormalText	=	21		# Equation normal text.
wdOMathFunctionPhantom	=	14		# Equation phantom.
wdOMathFunctionRad	=	16		# Equation base expression.
wdOMathFunctionScrPre	=	15		# Scr pre.
wdOMathFunctionScrSub	=	17		# Scr. sub.
wdOMathFunctionScrSubSup	=	18		# Scr. sub sup.
wdOMathFunctionScrSup	=	19		# Scr sup.
wdOMathFunctionText	=	20		# Equation text.

#values from word.wdomathhorizaligntype
#****************************
wdOMathHorizAlignCenter	=	0		# Centered.
wdOMathHorizAlignLeft	=	1		# Left alignment.
wdOMathHorizAlignRight	=	2		# Right alignment.

#values from word.wdomathjc
#****************************
wdOMathJcCenter	=	2		# Center.
wdOMathJcCenterGroup	=	1		# Center as a group.
wdOMathJcInline	=	7		# Inline.
wdOMathJcLeft	=	3		# Left.
wdOMathJcRight	=	4		# Right.

#values from word.wdomathshapetype
#****************************
wdOMathShapeCentered	=	0		# Vertically centers delimiters around the entire height of the equation causing delimiters grow equally above and below their midpoint.
wdOMathShapeMatch	=	1		# Matches the shape of the delimiters to the size of their contents.

#values from word.wdomathspacingrule
#****************************
wdOMathSpacing1pt5	=	1		# One and half spaces for each line.
wdOMathSpacingDouble	=	2		# Double spacing.
wdOMathSpacingExactly	=	3		# Exact spacing measurement.
wdOMathSpacingMultiple	=	4		# Custom spacing measurement.
wdOMathSpacingSingle	=	0		# Single spacing.

#values from word.wdomathtype
#****************************
wdOMathDisplay	=	0		# Professional format.
wdOMathInline	=	1		# Inline.

#values from word.wdomathvertaligntype
#****************************
wdOMathVertAlignBottom	=	2		# Aligns the equation on the bottom of the shape canvas or line.
wdOMathVertAlignCenter	=	0		# Vertically centers the equation in the shape canvas or line.
wdOMathVertAlignTop	=	1		# Aligns the equation on the top of the shape canvas or line.

#values from word.wdopenformat
#****************************
wdOpenFormatAllWord	=	6		# A Microsoft Word format that is backward compatible with earlier versions of Word.
wdOpenFormatAuto	=	0		# The existing format.
wdOpenFormatDocument	=	1		# Word format.
wdOpenFormatEncodedText	=	5		# Encoded text format.
wdOpenFormatRTF	=	3		# Rich text format (RTF).
wdOpenFormatTemplate	=	2		# As a Word template.
wdOpenFormatText	=	4		# Unencoded text format.
wdOpenFormatOpenDocumentText	=	18		# OpenDocument Text format.
wdOpenFormatUnicodeText	=	5		# Unicode text format.
wdOpenFormatWebPages	=	7		# HTML format.
wdOpenFormatXML	=	8		# XML format.
wdOpenFormatAllWordTemplates	=	13		# Word template format.
wdOpenFormatDocument97	=	1		# Microsoft Word 97 document format.
wdOpenFormatTemplate97	=	2		# Word 97 template format.
wdOpenFormatXMLDocument	=	9		# XML document format.
wdOpenFormatXMLDocumentSerialized	=	14		# Open XML file format saved as a single XML file.
wdOpenFormatXMLDocumentMacroEnabled	=	10		# XML document format with macros enabled.
wdOpenFormatXMLDocumentMacroEnabledSerialized	=	15		# Open XML file format with macros enabled saved as a single XML file.
wdOpenFormatXMLTemplate	=	11		# XML template format.
wdOpenFormatXMLTemplateSerialized	=	16		# Open XML template format saved as a XML single file.
wdOpenFormatXMLTemplateMacroEnabled	=	12		# XML template format with macros enabled.
wdOpenFormatXMLTemplateMacroEnabledSerialized	=	17		# Open XML template format with macros enabled saved as a single XML file.

#values from word.wdorganizerobject
#****************************
wdOrganizerObjectAutoText	=	1		# An AutoText item.
wdOrganizerObjectCommandBars	=	2		# A command bar item.
wdOrganizerObjectProjectItems	=	3		# A project item.
wdOrganizerObjectStyles	=	0		# A style item.

#values from word.wdorientation
#****************************
wdOrientLandscape	=	1		# Landscape orientation.
wdOrientPortrait	=	0		# Portrait orientation.

#values from word.wdoriginalformat
#****************************
wdOriginalDocumentFormat	=	1		# Original document format.
wdPromptUser	=	2		# Prompt user to select a document format.
wdWordDocument	=	0		# Microsoft Word document format.

#values from word.wdoutlinelevel
#****************************
wdOutlineLevel1	=	1		# Outline level 1.
wdOutlineLevel2	=	2		# Outline level 2.
wdOutlineLevel3	=	3		# Outline level 3.
wdOutlineLevel4	=	4		# Outline level 4.
wdOutlineLevel5	=	5		# Outline level 5.
wdOutlineLevel6	=	6		# Outline level 6.
wdOutlineLevel7	=	7		# Outline level 7.
wdOutlineLevel8	=	8		# Outline level 8.
wdOutlineLevel9	=	9		# Outline level 9.
wdOutlineLevelBodyText	=	10		# No outline level.

#values from word.wdpageborderart
#****************************
wdArtApples	=	1		# An apple border.
wdArtArchedScallops	=	97		# An arched scalloped border.
wdArtBabyPacifier	=	70		# A baby pacifier border.
wdArtBabyRattle	=	71		# A baby rattle border.
wdArtBalloons3Colors	=	11		# Balloons in three colors as the border.
wdArtBalloonsHotAir	=	12		# A hot air balloon border.
wdArtBasicBlackDashes	=	155		# A basic black-dashed border.
wdArtBasicBlackDots	=	156		# A basic black-dotted border.
wdArtBasicBlackSquares	=	154		# A basic black squares border.
wdArtBasicThinLines	=	151		# A basic thin-lines border.
wdArtBasicWhiteDashes	=	152		# A basic white-dashed border.
wdArtBasicWhiteDots	=	147		# A basic white-dotted border.
wdArtBasicWhiteSquares	=	153		# A basic white squares border.
wdArtBasicWideInline	=	150		# A basic wide inline border.
wdArtBasicWideMidline	=	148		# A basic wide midline border.
wdArtBasicWideOutline	=	149		# A basic wide outline border.
wdArtBats	=	37		# A bats border.
wdArtBirds	=	102		# A birds border.
wdArtBirdsFlight	=	35		# A birds-in-flight border.
wdArtCabins	=	72		# A cabins border.
wdArtCakeSlice	=	3		# A cake slice border.
wdArtCandyCorn	=	4		# A candy corn border.
wdArtCelticKnotwork	=	99		# A Celtic knotwork border.
wdArtCertificateBanner	=	158		# A certificate banner border.
wdArtChainLink	=	128		# A chain-link border.
wdArtChampagneBottle	=	6		# A champagne bottle border.
wdArtCheckedBarBlack	=	145		# A checked-bar black border.
wdArtCheckedBarColor	=	61		# A checked-bar colored border.
wdArtCheckered	=	144		# A checkered border.
wdArtChristmasTree	=	8		# A Christmas tree border.
wdArtCirclesLines	=	91		# A circles-and-lines border.
wdArtCirclesRectangles	=	140		# A circles-and-rectangles border.
wdArtClassicalWave	=	56		# A classical wave border.
wdArtClocks	=	27		# A clocks border.
wdArtCompass	=	54		# A compass border.
wdArtConfetti	=	31		# A confetti border.
wdArtConfettiGrays	=	115		# A confetti border using shades of gray.
wdArtConfettiOutline	=	116		# A confetti outline border.
wdArtConfettiStreamers	=	14		# A confetti streamers border.
wdArtConfettiWhite	=	117		# A confetti white border.
wdArtCornerTriangles	=	141		# A triangles border.
wdArtCouponCutoutDashes	=	163		# A coupon-cut-out dashes border.
wdArtCouponCutoutDots	=	164		# A coupon-cut-out dots border.
wdArtCrazyMaze	=	100		# A crazy maze border.
wdArtCreaturesButterfly	=	32		# A butterfly border.
wdArtCreaturesFish	=	34		# A fish border.
wdArtCreaturesInsects	=	142		# An insect border.
wdArtCreaturesLadyBug	=	33		# A ladybug border.
wdArtCrossStitch	=	138		# A cross-stitch border.
wdArtCup	=	67		# A cup border.
wdArtDecoArch	=	89		# A deco arch border.
wdArtDecoArchColor	=	50		# A deco arch colored border.
wdArtDecoBlocks	=	90		# A deco blocks border.
wdArtDiamondsGray	=	88		# A diamond border using shades of gray.
wdArtDoubleD	=	55		# A double-D border.
wdArtDoubleDiamonds	=	127		# A double-diamonds border.
wdArtEarth1	=	22		# An earth number 1 border.
wdArtEarth2	=	21		# An earth number 2 border.
wdArtEclipsingSquares1	=	101		# An eclipsing squares number 1 border.
wdArtEclipsingSquares2	=	86		# An eclipsing squares number 2 border.
wdArtEggsBlack	=	66		# A black eggs border.
wdArtFans	=	51		# A fans border.
wdArtFilm	=	52		# A film border.
wdArtFirecrackers	=	28		# A fire crackers border.
wdArtFlowersBlockPrint	=	49		# A block flowers print border.
wdArtFlowersDaisies	=	48		# A daisies border.
wdArtFlowersModern1	=	45		# A modern flowers number 1 border.
wdArtFlowersModern2	=	44		# A modern flowers number 2 border.
wdArtFlowersPansy	=	43		# A pansy border.
wdArtFlowersRedRose	=	39		# A red rose border.
wdArtFlowersRoses	=	38		# A rose border.
wdArtFlowersTeacup	=	103		# A teacup border.
wdArtFlowersTiny	=	42		# A tiny flower border.
wdArtGems	=	139		# A gems border.
wdArtGingerbreadMan	=	69		# A gingerbread man border.
wdArtGradient	=	122		# A gradient border.
wdArtHandmade1	=	159		# A handmade number 1 border.
wdArtHandmade2	=	160		# A handmade number 2 border.
wdArtHeartBalloon	=	16		# A heart-balloon border.
wdArtHeartGray	=	68		# A heart border in shades of gray.
wdArtHearts	=	15		# A hearts border.
wdArtHeebieJeebies	=	120		# A heebie-jeebies border.
wdArtHolly	=	41		# A holly border.
wdArtHouseFunky	=	73		# A funky house border.
wdArtHypnotic	=	87		# An hypnotic border.
wdArtIceCreamCones	=	5		# An ice cream cones border.
wdArtLightBulb	=	121		# A light bulb border.
wdArtLightning1	=	53		# A lightning number 1 border.
wdArtLightning2	=	119		# A lightning number 2 border.
wdArtMapleLeaf	=	81		# A maple leaf border.
wdArtMapleMuffins	=	2		# A maple muffins border.
wdArtMapPins	=	30		# A map pins border.
wdArtMarquee	=	146		# A marquee border.
wdArtMarqueeToothed	=	131		# A marquee toothed border.
wdArtMoons	=	125		# A moons border.
wdArtMosaic	=	118		# A mosaic border.
wdArtMusicNotes	=	79		# A music notes border.
wdArtNorthwest	=	104		# A northwest border.
wdArtOvals	=	126		# An ovals border.
wdArtPackages	=	26		# A packages border.
wdArtPalmsBlack	=	80		# A black palms border.
wdArtPalmsColor	=	10		# A colored palms border.
wdArtPaperClips	=	82		# A paper clips border.
wdArtPapyrus	=	92		# A papyrus border.
wdArtPartyFavor	=	13		# A party favor border.
wdArtPartyGlass	=	7		# A party glass border.
wdArtPencils	=	25		# A pencils border.
wdArtPeople	=	84		# A people border.
wdArtPeopleHats	=	23		# A people-wearing-hats border.
wdArtPeopleWaving	=	85		# A people-waving border.
wdArtPoinsettias	=	40		# A poinsettias border.
wdArtPostageStamp	=	135		# A postage stamp border.
wdArtPumpkin1	=	65		# A pumpkin number 1 border.
wdArtPushPinNote1	=	63		# A pushpin note number 1 border.
wdArtPushPinNote2	=	64		# A pushpin note number 2 border.
wdArtPyramids	=	113		# A pyramids border.
wdArtPyramidsAbove	=	114		# An external pyramids border.
wdArtQuadrants	=	60		# A quadrants border.
wdArtRings	=	29		# A rings border.
wdArtSafari	=	98		# A safari border.
wdArtSawtooth	=	133		# A saw-tooth border.
wdArtSawtoothGray	=	134		# A saw-tooth border using shades of gray.
wdArtScaredCat	=	36		# A scared cat border.
wdArtSeattle	=	78		# A Seattle border.
wdArtShadowedSquares	=	57		# A shadowed squared border.
wdArtSharksTeeth	=	132		# A shark-tooth border.
wdArtShorebirdTracks	=	83		# A shorebird tracks border.
wdArtSkyrocket	=	77		# A sky rocket border.
wdArtSnowflakeFancy	=	76		# A fancy snowflake border.
wdArtSnowflakes	=	75		# A snowflake border.
wdArtSombrero	=	24		# A sombrero border.
wdArtSouthwest	=	105		# A southwest border.
wdArtStars	=	19		# A stars border.
wdArtStars3D	=	17		# A 3D stars border.
wdArtStarsBlack	=	74		# A black stars border.
wdArtStarsShadowed	=	18		# A shadowed stars border.
wdArtStarsTop	=	157		# A stars-on-top border.
wdArtSun	=	20		# A sun border.
wdArtSwirligig	=	62		# A swirling border.
wdArtTornPaper	=	161		# A torn-paper border.
wdArtTornPaperBlack	=	162		# A black torn-paper border.
wdArtTrees	=	9		# A trees border.
wdArtTriangleParty	=	123		# A triangle party border.
wdArtTriangles	=	129		# A triangles border.
wdArtTribal1	=	130		# A tribal number 1 border.
wdArtTribal2	=	109		# A tribal number 2 border.
wdArtTribal3	=	108		# A tribal number 3 border.
wdArtTribal4	=	107		# A tribal number 4 border.
wdArtTribal5	=	110		# A tribal number 5 border.
wdArtTribal6	=	106		# A tribal number 6 border.
wdArtTwistedLines1	=	58		# A twisted lines number 1 border.
wdArtTwistedLines2	=	124		# A twisted lines number 2 border.
wdArtVine	=	47		# A vine border.
wdArtWaveline	=	59		# A wave-line border.
wdArtWeavingAngles	=	96		# A weaving angle border.
wdArtWeavingBraid	=	94		# A weaving braid border.
wdArtWeavingRibbon	=	95		# A weaving ribbon border.
wdArtWeavingStrips	=	136		# A weaving strips border.
wdArtWhiteFlowers	=	46		# A white flower border.
wdArtWoodwork	=	93		# A woodwork border.
wdArtXIllusions	=	111		# An X illusion border.
wdArtZanyTriangles	=	112		# A zany triangle border.
wdArtZigZag	=	137		# A zigzag border.
wdArtZigZagStitch	=	143		# A zigzag stitch border.

#values from word.wdpagecolor
#****************************
wdPageColorInverse	=	2		# Inverse page color. Renders the document content in a manner that resembles high-contrast black, although not necessarily exactly so. Some figures are rendered in full color on a black background.
wdPageColorNone	=	0		# No page color, the default. The page background is rendered in white. Any assigned page background colors are ignored.
wdPageColorSepia	=	1		# Sepia page color, RGB (112, 66, 20) at 80% transparency. Makes no changes to the contents of the document.

#values from word.wdpagefit
#****************************
wdPageFitBestFit	=	2		# Best fit the page to the active window.
wdPageFitFullPage	=	1		# View the full page.
wdPageFitNone	=	0		# Do not adjust the view settings for the page.
wdPageFitTextFit	=	3		# Best fit the text of the page to the active window.

#values from word.wdpagemovementtype
#****************************
wdVertical	=	1		# Document page movement type vertical.
wdSideToSide	=	2		# Document page movement type side-to-side.

#values from word.wdpagenumberalignment
#****************************
wdAlignPageNumberCenter	=	1		# Centered.
wdAlignPageNumberInside	=	3		# Left-aligned just inside the footer.
wdAlignPageNumberLeft	=	0		# Left-aligned.
wdAlignPageNumberOutside	=	4		# Right-aligned just outside the footer.
wdAlignPageNumberRight	=	2		# Right-aligned.

#values from word.wdpagenumberstyle
#****************************
wdPageNumberStyleArabic	=	0		# Arabic style.
wdPageNumberStyleArabicFullWidth	=	14		# Arabic full width style.
wdPageNumberStyleArabicLetter1	=	46		# Arabic letter 1 style.
wdPageNumberStyleArabicLetter2	=	48		# Arabic letter 2 style.
wdPageNumberStyleHanjaRead	=	41		# Hanja Read style.
wdPageNumberStyleHanjaReadDigit	=	42		# Hanja Read Digit style.
wdPageNumberStyleHebrewLetter1	=	45		# Hebrew letter 1 style.
wdPageNumberStyleHebrewLetter2	=	47		# Hebrew letter 2 style.
wdPageNumberStyleHindiArabic	=	51		# Hindi Arabic style.
wdPageNumberStyleHindiCardinalText	=	52		# Hindi Cardinal text style.
wdPageNumberStyleHindiLetter1	=	49		# Hindi letter 1 style.
wdPageNumberStyleHindiLetter2	=	50		# Hindi letter 2 style.
wdPageNumberStyleKanji	=	10		# Kanji style.
wdPageNumberStyleKanjiDigit	=	11		# Kanji Digit style.
wdPageNumberStyleKanjiTraditional	=	16		# Kanji traditional style.
wdPageNumberStyleLowercaseLetter	=	4		# Lowercase letter style.
wdPageNumberStyleLowercaseRoman	=	2		# Lowercase Roman style.
wdPageNumberStyleNumberInCircle	=	18		# Number in circle style.
wdPageNumberStyleNumberInDash	=	57		# Number in dash style.
wdPageNumberStyleSimpChinNum1	=	37		# Simplified Chinese number 1 style.
wdPageNumberStyleSimpChinNum2	=	38		# Simplified Chinese number 2 style.
wdPageNumberStyleThaiArabic	=	54		# Thai Arabic style.
wdPageNumberStyleThaiCardinalText	=	55		# Thai Cardinal Text style.
wdPageNumberStyleThaiLetter	=	53		# Thai letter style.
wdPageNumberStyleTradChinNum1	=	33		# Traditional Chinese number 1 style.
wdPageNumberStyleTradChinNum2	=	34		# Traditional Chinese number 2 style.
wdPageNumberStyleUppercaseLetter	=	3		# Uppercase letter style.
wdPageNumberStyleUppercaseRoman	=	1		# Uppercase Roman style.
wdPageNumberStyleVietCardinalText	=	56		# Vietnamese Cardinal text style.

#values from word.wdpapersize
#****************************
wdPaper10x14	=	0		# 10 inches wide, 14 inches long.
wdPaper11x17	=	1		# Legal 11 inches wide, 17 inches long.
wdPaperA3	=	6		# A3 dimensions.
wdPaperA4	=	7		# A4 dimensions.
wdPaperA4Small	=	8		# Small A4 dimensions.
wdPaperA5	=	9		# A5 dimensions.
wdPaperB4	=	10		# B4 dimensions.
wdPaperB5	=	11		# B5 dimensions.
wdPaperCSheet	=	12		# C sheet dimensions.
wdPaperCustom	=	41		# Custom paper size.
wdPaperDSheet	=	13		# D sheet dimensions.
wdPaperEnvelope10	=	25		# Legal envelope, size 10.
wdPaperEnvelope11	=	26		# Envelope, size 11.
wdPaperEnvelope12	=	27		# Envelope, size 12.
wdPaperEnvelope14	=	28		# Envelope, size 14.
wdPaperEnvelope9	=	24		# Envelope, size 9.
wdPaperEnvelopeB4	=	29		# B4 envelope.
wdPaperEnvelopeB5	=	30		# B5 envelope.
wdPaperEnvelopeB6	=	31		# B6 envelope.
wdPaperEnvelopeC3	=	32		# C3 envelope.
wdPaperEnvelopeC4	=	33		# C4 envelope.
wdPaperEnvelopeC5	=	34		# C5 envelope.
wdPaperEnvelopeC6	=	35		# C6 envelope.
wdPaperEnvelopeC65	=	36		# C65 envelope.
wdPaperEnvelopeDL	=	37		# DL envelope.
wdPaperEnvelopeItaly	=	38		# Italian envelope.
wdPaperEnvelopeMonarch	=	39		# Monarch envelope.
wdPaperEnvelopePersonal	=	40		# Personal envelope.
wdPaperESheet	=	14		# E sheet dimensions.
wdPaperExecutive	=	5		# Executive dimensions.
wdPaperFanfoldLegalGerman	=	15		# German legal fanfold dimensions.
wdPaperFanfoldStdGerman	=	16		# German standard fanfold dimensions.
wdPaperFanfoldUS	=	17		# United States fanfold dimensions.
wdPaperFolio	=	18		# Folio dimensions.
wdPaperLedger	=	19		# Ledger dimensions.
wdPaperLegal	=	4		# Legal dimensions.
wdPaperLetter	=	2		# Letter dimensions.
wdPaperLetterSmall	=	3		# Small letter dimensions.
wdPaperNote	=	20		# Note dimensions.
wdPaperQuarto	=	21		# Quarto dimensions.
wdPaperStatement	=	22		# Statement dimensions.
wdPaperTabloid	=	23		# Tabloid dimensions.

#values from word.wdpapertray
#****************************
wdPrinterAutomaticSheetFeed	=	7		# Automatic sheet feed.
wdPrinterDefaultBin	=	0		# Default bin.
wdPrinterEnvelopeFeed	=	5		# Envelope feed.
wdPrinterFormSource	=	15		# Form source.
wdPrinterLargeCapacityBin	=	11		# Large-capacity bin.
wdPrinterLargeFormatBin	=	10		# Large-format bin.
wdPrinterLowerBin	=	2		# Lower bin.
wdPrinterManualEnvelopeFeed	=	6		# Manual envelope feed.
wdPrinterManualFeed	=	4		# Manual feed.
wdPrinterMiddleBin	=	3		# Middle bin.
wdPrinterOnlyBin	=	1		# Printer's only bin.
wdPrinterPaperCassette	=	14		# Paper cassette.
wdPrinterSmallFormatBin	=	9		# Small-format bin.
wdPrinterTractorFeed	=	8		# Tractor feed.
wdPrinterUpperBin	=	1		# Upper bin.

#values from word.wdparagraphalignment
#****************************
wdAlignParagraphCenter	=	1		# Center-aligned.
wdAlignParagraphDistribute	=	4		# Paragraph characters are distributed to fill the entire width of the paragraph.
wdAlignParagraphJustify	=	3		# Fully justified.
wdAlignParagraphJustifyHi	=	7		# Justified with a high character compression ratio.
wdAlignParagraphJustifyLow	=	8		# Justified with a low character compression ratio.
wdAlignParagraphJustifyMed	=	5		# Justified with a medium character compression ratio.
wdAlignParagraphLeft	=	0		# Left-aligned.
wdAlignParagraphRight	=	2		# Right-aligned.
wdAlignParagraphThaiJustify	=	9		# Justified according to Thai formatting layout.

#values from word.wdpartofspeech
#****************************
wdAdjective	=	0		# An adjective.
wdAdverb	=	2		# An adverb.
wdConjunction	=	5		# A conjunction.
wdIdiom	=	8		# An idiom.
wdInterjection	=	7		# An interjection.
wdNoun	=	1		# A noun.
wdOther	=	9		# Some other part of speech.
wdPreposition	=	6		# A preposition.
wdPronoun	=	4		# A pronoun.
wdVerb	=	3		# A verb.

#values from word.wdpastedatatype
#****************************
wdPasteBitmap	=	4		# Bitmap.
wdPasteDeviceIndependentBitmap	=	5		# Device-independent bitmap.
wdPasteEnhancedMetafile	=	9		# Enhanced metafile.
wdPasteHTML	=	10		# HTML.
wdPasteHyperlink	=	7		# Hyperlink.
wdPasteMetafilePicture	=	3		# Metafile picture.
wdPasteOLEObject	=	0		# OLE object.
wdPasteRTF	=	1		# Rich Text Format (RTF).
wdPasteShape	=	8		# Shape.
wdPasteText	=	2		# Text.

#values from word.wdpasteoptions
#****************************
wdKeepSourceFormatting	=	0		# Keeps formatting from the source document.
wdKeepTextOnly	=	2		# Keeps text only, without formatting.
wdMatchDestinationFormatting	=	1		# Matches formatting to the destination document.
wdUseDestinationStyles	=	3		# Matches formatting to the destination document using styles for formatting.

#values from word.wdphoneticguidealignmenttype
#****************************
wdPhoneticGuideAlignmentCenter	=	0		# Microsoft Word centers phonetic text over the specified range. This is the default value.
wdPhoneticGuideAlignmentLeft	=	3		# Word left-aligns phonetic text with the specified range.
wdPhoneticGuideAlignmentOneTwoOne	=	2		# Word adjusts the inside and outside spacing of the phonetic text in a 1:2:1 ratio.
wdPhoneticGuideAlignmentRight	=	4		# Word right-aligns phonetic text with the specified range.
wdPhoneticGuideAlignmentRightVertical	=	5		# Word aligns the phonetic text on the right side of vertical text.
wdPhoneticGuideAlignmentZeroOneZero	=	1		# Word adjusts the inside and outside spacing of the phonetic text in a 0:1:0 ratio.

#values from word.wdpicturelinktype
#****************************
wdLinkDataInDoc	=	1		# Embed the picture in the document.
wdLinkDataOnDisk	=	2		# Link the picture to the document.
wdLinkNone	=	0		# Do not link to or embed the picture in the document.

#values from word.wdportuguesereform
#****************************
wdPortugueseBoth	=	3		# Use both the pre-reform and post-reform spelling rules.
wdPortuguesePostReform	=	2		# Use the post-reform spelling rules.
wdPortuguesePreReform	=	1		# Use the pre-reform spelling rules.

#values from word.wdpreferredwidthtype
#****************************
wdPreferredWidthAuto	=	1		# Automatically select the unit of measure to use based on the current selection.
wdPreferredWidthPercent	=	2		# Measure the current item width using a specified percentage.
wdPreferredWidthPoints	=	3		# Measure the current item width using a specified number of points.

#values from word.wdprintoutitem
#****************************
wdPrintAutoTextEntries	=	4		# Autotext entries in the current document.
wdPrintComments	=	2		# Comments in the current document.
wdPrintDocumentContent	=	0		# Current document content.
wdPrintDocumentWithMarkup	=	7		# Current document content including markup.
wdPrintEnvelope	=	6		# An envelope.
wdPrintKeyAssignments	=	5		# Key assignments in the current document.
wdPrintMarkup	=	2		# Markup in the current document.
wdPrintProperties	=	1		# Properties in the current document.
wdPrintStyles	=	3		# Styles in the current document.

#values from word.wdprintoutpages
#****************************
wdPrintAllPages	=	0		# All pages.
wdPrintEvenPagesOnly	=	2		# Even-numbered pages only.
wdPrintOddPagesOnly	=	1		# Odd-numbered pages only.

#values from word.wdprintoutrange
#****************************
wdPrintAllDocument	=	0		# The entire document.
wdPrintCurrentPage	=	2		# The current page.
wdPrintFromTo	=	3		# A specified range.
wdPrintRangeOfPages	=	4		# A specified range of pages.
wdPrintSelection	=	1		# The current selection.

#values from word.wdproofreadingerrortype
#****************************
wdGrammaticalError	=	1		# Grammatical error.
wdSpellingError	=	0		# Spelling error.

#values from word.wdprotectedviewclosereason
#****************************
wdProtectedViewCloseEdit	=	1		# The window was closed when the user clicked the  Enable Editing or Edit Anyway button while in Protected View.
wdProtectedViewCloseForced	=	2		# The window was closed because the application shut it down forcefully or it stopped responding.
wdProtectedViewCloseNormal	=	0		# The window was closed normally.

#values from word.wdprotectiontype
#****************************
wdAllowOnlyComments	=	1		# Allow only comments to be added to the document.
wdAllowOnlyFormFields	=	2		# Allow content to be added to the document only through form fields.
wdAllowOnlyReading	=	3		# Allow read-only access to the document.
wdAllowOnlyRevisions	=	0		# Allow only revisions to be made to existing content.
wdNoProtection	=	-1		# Do not apply protection to the document.

#values from word.wdreadinglayoutmargin
#****************************
wdAutomaticMargin	=	0		# Shows the pages without margins.
wdFullMargin	=	2		# Shows the pages with margins.
wdSuppressMargin	=	1		# Microsoft Word determines automatically whether to show or hide the margins based on the available space.

#values from word.wdreadingorder
#****************************
wdReadingOrderLtr	=	1		# Left-to-right reading order.
wdReadingOrderRtl	=	0		# Right-to-left reading order.

#values from word.wdrecoverytype
#****************************
wdChart	=	14		# Pastes a Microsoft Office Excel chart as an embedded OLE object.
wdChartLinked	=	15		# Pastes an Excel chart and links it to the original Excel spreadsheet.
wdChartPicture	=	13		# Pastes an Excel chart as a picture.
wdFormatOriginalFormatting	=	16		# Preserves original formatting of the pasted material.
wdFormatPlainText	=	22		# Pastes as plain, unformatted text.
wdFormatSurroundingFormattingWithEmphasis	=	20		# Matches the formatting of the pasted text to the formatting of surrounding text.
wdListCombineWithExistingList	=	24		# Merges a pasted list with neighboring lists.
wdListContinueNumbering	=	7		# Continues numbering of a pasted list from the list in the document.
wdListDontMerge	=	25		# Not supported.
wdListRestartNumbering	=	8		# Restarts numbering of a pasted list.
wdPasteDefault	=	0		# Not supported.
wdSingleCellTable	=	6		# Pastes a single cell table as a separate table.
wdSingleCellText	=	5		# Pastes a single cell as text.
wdTableAppendTable	=	10		# Merges pasted cells into an existing table by inserting the pasted rows between the selected rows.
wdTableInsertAsRows	=	11		# Inserts a pasted table as rows between two rows in the target table.
wdTableOriginalFormatting	=	12		# Pastes an appended table without merging table styles.
wdTableOverwriteCells	=	23		# Pastes table cells and overwrites existing table cells.
wdUseDestinationStylesRecovery	=	19		# Uses the styles that are in use in the destination document.

#values from word.wdrectangletype
#****************************
wdLineBetweenColumnRectangle	=	5		# Represents a region corresponding to a line that separates columns.
wdMarkupRectangle	=	2		# Represents a space occupied by a comment balloon.
wdMarkupRectangleButton	=	3		# Represents a space occupied by the more (...) indicator that appears in a comment balloon when there is additional text for the comment.
wdPageBorderRectangle	=	4		# Represents a space occupied by a page border.
wdSelection	=	6		# Represents a space occupied by a selection tool, for example the table selection tool in the upper-left corner of a table or the anchor for an image.
wdShapeRectangle	=	1		# Represents a space occupied by a shape.
wdSystem	=	7		# Not applicable.
wdTextRectangle	=	0		# Represents a space occupied by text.
wdDocumentControlRectangle	=	13		# Represents space occupied by a content control, equation, or document building block in-document control.
wdMailNavArea	=	12		# Represents space occupied by the email message navigation buttons when reading email in Microsoft Office Outlook.
wdMarkupRectangleArea	=	8		# Represents space occupied for the presentation of revision balloons on the page. This space is only printed if you print using &quot;Document Showing Markup&quot; in the  Print dialog box.
wdMarkupRectangleMoveMatch	=	10		# Represents space occupied by the  Go button used to find matching pairs of tracked moves in a document.
wdReadingModeNavigation	=	9		# Represents space occupied by the page navigation buttons when reading a document in full page reading view.
wdReadingModePanningArea	=	11		# Represents space occupied for page turning when reading a document in full page reading view.

#values from word.wdreferencekind
#****************************
wdContentText	=	-1		# Insert text value of the specified item. For example, insert text of the specified heading.
wdEndnoteNumber	=	6		# Insert endnote reference mark.
wdEndnoteNumberFormatted	=	17		# Insert formatted endnote reference mark.
wdEntireCaption	=	2		# Insert label, number, and any additional caption of specified equation, figure, or table.
wdFootnoteNumber	=	5		# Insert footnote reference mark.
wdFootnoteNumberFormatted	=	16		# Insert formatted footnote reference mark.
wdNumberFullContext	=	-4		# Insert complete heading or paragraph number.
wdNumberNoContext	=	-3		# Insert heading or paragraph without its relative position in the outline numbered list.
wdNumberRelativeContext	=	-2		# Insert heading or paragraph with as much of its relative position in the outline numbered list as necessary to identify the item.
wdOnlyCaptionText	=	4		# Insert only the caption text of the specified equation, figure, or table.
wdOnlyLabelAndNumber	=	3		# Insert only the label and number of the specified equation, figure, or table.
wdPageNumber	=	7		# Insert page number of specified item.
wdPosition	=	15		# Insert the word &quot;Above&quot; or the word &quot;Below&quot; as appropriate.

#values from word.wdreferencetype
#****************************
wdRefTypeBookmark	=	2		# Bookmark.
wdRefTypeEndnote	=	4		# Endnote.
wdRefTypeFootnote	=	3		# Footnote.
wdRefTypeHeading	=	1		# Heading.
wdRefTypeNumberedItem	=	0		# Numbered item.

#values from word.wdrelativehorizontalposition
#****************************
wdRelativeHorizontalPositionCharacter	=	3		# Relative to character.
wdRelativeHorizontalPositionColumn	=	2		# Relative to column.
wdRelativeHorizontalPositionMargin	=	0		# Relative to margin.
wdRelativeHorizontalPositionPage	=	1		# Relative to page.
wdRelativeHorizontalPositionInnerMarginArea	=	6		# Relative to inner margin area.
wdRelativeHorizontalPositionLeftMarginArea	=	4		# Relative to left margin.
wdRelativeHorizontalPositionOuterMarginArea	=	7		# Relative to outer margin area.
wdRelativeHorizontalPositionRightMarginArea	=	5		# Relative to right margin.

#values from word.wdrelativehorizontalsize
#****************************
wdRelativeHorizontalSizeInnerMarginArea	=	4		# Width is relative to the size of the inside margin?to the size of the left margin for odd pages, and to the size of the right margin for even pages.
wdRelativeHorizontalSizeLeftMarginArea	=	2		# Width is relative to the size of the left margin.
wdRelativeHorizontalSizeMargin	=	0		# Width is relative to the space between the left margin and the right margin.
wdRelativeHorizontalSizeOuterMarginArea	=	5		# Width is relative to the size of the outside margin?to the size of the right margin for odd pages, and to the size of the left margin for even pages.
wdRelativeHorizontalSizePage	=	1		# Width is relative to the width of the page.
wdRelativeHorizontalSizeRightMarginArea	=	3		# Width is relative to the width of the right margin.

#values from word.wdrelativeverticalposition
#****************************
wdRelativeVerticalPositionLine	=	3		# Relative to line.
wdRelativeVerticalPositionMargin	=	0		# Relative to margin.
wdRelativeVerticalPositionPage	=	1		# Relative to page.
wdRelativeVerticalPositionParagraph	=	2		# Relative to paragraph.
wdRelativeVerticalPositionBottomMarginArea	=	5		# Relative to bottom margin.
wdRelativeVerticalPositionInnerMarginArea	=	6		# Relative to inner margin area.
wdRelativeVerticalPositionOuterMarginArea	=	7		# Relative to outer margin area.
wdRelativeVerticalPositionTopMarginArea	=	4		# Relative to top margin.

#values from word.wdrelativeverticalsize
#****************************
wdRelativeVerticalSizeBottomMarginArea	=	3		# Height is relative to the size of the bottom margin.
wdRelativeVerticalSizeInnerMarginArea	=	4		# Height is relative to the size of the inside margin?to the size of the top margin for odd pages, and to the size of the bottom margin for even pages.
wdRelativeVerticalSizeMargin	=	0		# Height is relative to the space between the left margin and the right margin.
wdRelativeVerticalSizeOuterMarginArea	=	5		# Height is relative to the size of the outside margin?to the size of the bottom margin for odd pages, and to the size of the top margin for even pages.
wdRelativeVerticalSizePage	=	1		# Height is relative to the height of the page.
wdRelativeVerticalSizeTopMarginArea	=	2		# Height is relative to the size of the top margin.

#values from word.wdrelocate
#****************************
wdRelocateDown	=	1		# Below the next visible paragraph.
wdRelocateUp	=	0		# Above the previous visible paragraph.

#values from word.wdremovedocinfotype
#****************************
wdRDIAll	=	99		# Removes all document information.
wdRDIComments	=	1		# Removes document comments.
wdRDIContentType	=	16		# Removes content type information.
wdRDIDocumentManagementPolicy	=	15		# Removes document management policy information.
wdRDIDocumentProperties	=	8		# Removes document properties.
wdRDIDocumentServerProperties	=	14		# Removes document server properties.
wdRDIDocumentWorkspace	=	10		# Removes document workspace information.
wdRDIEmailHeader	=	5		# Removes email header information.
wdRDIInkAnnotTations	=	11		# Removes ink annotations.
wdRDIRemovePersonalInformation	=	4		# Removes personal information.
wdRDIRevisions	=	2		# Removes revision marks.
wdRDIRoutingSlip	=	6		# Removes routing slip information.
wdRDISendForReview	=	7		# Removes information stored when sending a document for review.
wdRDITemplate	=	9		# Removes template information.
wdRDITaskpaneWebExtensions	=	17		# Removes taskpane web extensions information.
wdRDIVersions	=	3		# Removes document version information.

#values from word.wdreplace
#****************************
wdReplaceAll	=	2		# Replace all occurrences.
wdReplaceNone	=	0		# Replace no occurrences.
wdReplaceOne	=	1		# Replace the first occurrence encountered.

#values from word.wdrevisedlinesmark
#****************************
wdRevisedLinesMarkLeftBorder	=	1		# In the left border.
wdRevisedLinesMarkNone	=	0		# Not displayed.
wdRevisedLinesMarkOutsideBorder	=	3		# Outside the border.
wdRevisedLinesMarkRightBorder	=	2		# In the right border.

#values from word.wdrevisedpropertiesmark
#****************************
wdRevisedPropertiesMarkBold	=	1		# In bold.
wdRevisedPropertiesMarkColorOnly	=	5		# In the designated color.
wdRevisedPropertiesMarkDoubleStrikeThrough	=	7		# Using double-strikethrough characters.
wdRevisedPropertiesMarkDoubleUnderline	=	4		# With double-underline characters.
wdRevisedPropertiesMarkItalic	=	2		# In italic.
wdRevisedPropertiesMarkNone	=	0		# Using a special character.
wdRevisedPropertiesMarkStrikeThrough	=	6		# Using strikethrough characters.
wdRevisedPropertiesMarkUnderline	=	3		# In underline.

#values from word.wdrevisionsballoonmargin
#****************************
wdLeftMargin	=	0		# Left margin.
wdRightMargin	=	1		# Right margin. default.

#values from word.wdrevisionsballoonprintorientation
#****************************
wdBalloonPrintOrientationAuto	=	0		# Microsoft Word automatically selects the orientation that keeps the zoom factor closest to 100%.
wdBalloonPrintOrientationForceLandscape	=	2		# Word forces all sections to be printed in Landscape mode, regardless of original orientation, and prints the revision and comment balloons on the side opposite to the document text.
wdBalloonPrintOrientationPreserve	=	1		# Word preserves the orientation of the original, uncommented document.

#values from word.wdrevisionsballoonwidthtype
#****************************
wdBalloonWidthPercent	=	0		# Measured as a percentage of the width of the document.
wdBalloonWidthPoints	=	1		# Measured in points.

#values from word.wdrevisionsmarkup
#****************************
wdRevisionsMarkupAll	=	2		# Displays the final document with all markup visible.
wdRevisionsMarkupNone	=	0		# Displays the final document with no markup visible.
wdRevisionsMarkupSimple	=	1		# Displays the final document in simple markup: with revisions incorporated, but with no markup visible.

#values from word.wdrevisionsmode
#****************************
wdBalloonRevisions	=	0		# Displays revisions in balloons in the left or right margin.
wdInLineRevisions	=	1		# Displays revisions within the text using strikethrough for deletions and underlining for insertions. This is the default setting for prior versions of Word.
wdMixedRevisions	=	2		# Not supported.

#values from word.wdrevisionsview
#****************************
wdRevisionsViewFinal	=	0		# Displays the document with formatting and content changes applied.
wdRevisionsViewOriginal	=	1		# Displays the document before changes were made.

#values from word.wdrevisionswrap
#****************************
wdWrapAlways	=	1		# Revisions are wrapped.
wdWrapAsk	=	2		# Ask the user if revisions should be wrapped.
wdWrapNever	=	0		# Never wrap revisions.

#values from word.wdrevisiontype
#****************************
wdNoRevision	=	0		# No revision.
wdRevisionCellDeletion	=	17		# Table cell deleted.
wdRevisionCellInsertion	=	16		# Table cell inserted.
wdRevisionCellMerge	=	18		# Table cells merged.
wdRevisionCellSplit	=	19		# This object, member, or enumeration is deprecated and is not intended to be used in your code.
wdRevisionConflict	=	7		# Revision marked as a conflict.
wdRevisionConflictDelete	=	21		# Deletion revision conflict in a coauthored document.
wdRevisionConflictInsert	=	20		# Insertion revision conflict in a coauthored document
wdRevisionDelete	=	2		# Deletion.
wdRevisionDisplayField	=	5		# Field display changed.
wdRevisionInsert	=	1		# Insertion.
wdRevisionMovedFrom	=	14		# Content moved from.
wdRevisionMovedTo	=	15		# Content moved to.
wdRevisionParagraphNumber	=	4		# Paragraph number changed.
wdRevisionParagraphProperty	=	10		# Paragraph property changed.
wdRevisionProperty	=	3		# Property changed.
wdRevisionReconcile	=	6		# Revision marked as reconciled conflict.
wdRevisionReplace	=	9		# Replaced.
wdRevisionSectionProperty	=	12		# Section property changed.
wdRevisionStyle	=	8		# Style changed.
wdRevisionStyleDefinition	=	13		# Style definition changed.
wdRevisionTableProperty	=	11		# Table property changed.

#values from word.wdrowalignment
#****************************
wdAlignRowCenter	=	1		# Centered.
wdAlignRowLeft	=	0		# Left-aligned. Default.
wdAlignRowRight	=	2		# Right-aligned.

#values from word.wdrowheightrule
#****************************
wdRowHeightAtLeast	=	1		# The row height is at least a minimum specified value.
wdRowHeightAuto	=	0		# The row height is adjusted to accommodate the tallest value in the row.
wdRowHeightExactly	=	2		# The row height is an exact value.

#values from word.wdrulerstyle
#****************************
wdAdjustFirstColumn	=	2		# Adjusts the left edge of the first column only, preserving the positions of the other columns and the right edge of the table.
wdAdjustNone	=	0		# Adjusts the left edge of row or rows, preserving the width of all columns by shifting them to the left or right. This is the default value.
wdAdjustProportional	=	1		# Adjusts the left edge of the first column, preserving the position of the right edge of the table by proportionally adjusting the widths of all the cells in the specified row or rows.
wdAdjustSameWidth	=	3		# Adjusts the left edge of the first column, preserving the position of the right edge of the table by setting the widths of all the cells in the specified row or rows to the same value.

#values from word.wdsalutationgender
#****************************
wdGenderFemale	=	0		# Female gender.
wdGenderMale	=	1		# Male gender.
wdGenderNeutral	=	2		# Neutral gender.
wdGenderUnknown	=	3		# Unknown gender.

#values from word.wdsalutationtype
#****************************
wdSalutationBusiness	=	2		# Business salutation
wdSalutationFormal	=	1		# Format salutation.
wdSalutationInformal	=	0		# Informal salutation.
wdSalutationOther	=	3		# Custom salutation.

#values from word.wdsaveformat
#****************************
wdFormatDocument	=	0		# Microsoft Office Word 97 - 2003 binary file format.
wdFormatDOSText	=	4		# Microsoft DOS text format.
wdFormatDOSTextLineBreaks	=	5		# Microsoft DOS text with line breaks preserved.
wdFormatEncodedText	=	7		# Encoded text format.
wdFormatFilteredHTML	=	10		# Filtered HTML format.
wdFormatFlatXML	=	19		# Open XML file format saved as a single XML file.
wdFormatFlatXMLMacroEnabled	=	20		# Open XML file format with macros enabled saved as a single XML file.
wdFormatFlatXMLTemplate	=	21		# Open XML template format saved as a XML single file.
wdFormatFlatXMLTemplateMacroEnabled	=	22		# Open XML template format with macros enabled saved as a single XML file.
wdFormatOpenDocumentText	=	23		# OpenDocument Text format.
wdFormatHTML	=	8		# Standard HTML format.
wdFormatRTF	=	6		# Rich text format (RTF).
wdFormatStrictOpenXMLDocument	=	24		# Strict Open XML document format.
wdFormatTemplate	=	1		# Word template format.
wdFormatText	=	2		# Microsoft Windows text format.
wdFormatTextLineBreaks	=	3		# Windows text format with line breaks preserved.
wdFormatUnicodeText	=	7		# Unicode text format.
wdFormatWebArchive	=	9		# Web archive format.
wdFormatXML	=	11		# Extensible Markup Language (XML) format.
wdFormatDocument97	=	0		# Microsoft Word 97 document format.
wdFormatDocumentDefault	=	16		# Word default document file format. For Word, this is the DOCX format.
wdFormatPDF	=	17		# PDF format.
wdFormatTemplate97	=	1		# Word 97 template format.
wdFormatXMLDocument	=	12		# XML document format.
wdFormatXMLDocumentMacroEnabled	=	13		# XML document format with macros enabled.
wdFormatXMLTemplate	=	14		# XML template format.
wdFormatXMLTemplateMacroEnabled	=	15		# XML template format with macros enabled.
wdFormatXPS	=	18		# XPS format.

#values from word.wdsaveoptions
#****************************
wdDoNotSaveChanges	=	0		# Do not save pending changes.
wdPromptToSaveChanges	=	-2		# Prompt the user to save pending changes.
wdSaveChanges	=	-1		# Save pending changes automatically without prompting the user.

#values from word.wdscrollbartype
#****************************
wdScrollbarTypeAuto	=	0		# Scroll bars are available for the specified frame only if the contents are too large to fit in the allotted space.
wdScrollbarTypeNo	=	2		# Scroll bars are never available for the specified frame.
wdScrollbarTypeYes	=	1		# Scroll bars are always available for the specified frame.

#values from word.wdsectiondirection
#****************************
wdSectionDirectionLtr	=	1		# Displays the section with left alignment and left-to-right reading order.
wdSectionDirectionRtl	=	0		# Displays the section with right alignment and right-to-left reading order.

#values from word.wdsectionstart
#****************************
wdSectionContinuous	=	0		# Continuous section break.
wdSectionEvenPage	=	3		# Even pages section break.
wdSectionNewColumn	=	1		# New column section break.
wdSectionNewPage	=	2		# New page section break.
wdSectionOddPage	=	4		# Odd pages section break.

#values from word.wdseekview
#****************************
wdSeekCurrentPageFooter	=	10		# The current page footer.
wdSeekCurrentPageHeader	=	9		# The current page header.
wdSeekEndnotes	=	8		# Endnotes.
wdSeekEvenPagesFooter	=	6		# The even pages footer.
wdSeekEvenPagesHeader	=	3		# The even pages header.
wdSeekFirstPageFooter	=	5		# The first page footer.
wdSeekFirstPageHeader	=	2		# The first page header.
wdSeekFootnotes	=	7		# Footnotes.
wdSeekMainDocument	=	0		# The main document.
wdSeekPrimaryFooter	=	4		# The primary footer.
wdSeekPrimaryHeader	=	1		# The primary header.

#values from word.wdselectionflags
#****************************
wdSelActive	=	8		# The selection is the active selection.
wdSelAtEOL	=	2		# The selection is at the end of the letter.
wdSelOvertype	=	4		# The selection was overtyped.
wdSelReplace	=	16		# The selection was replaced.
wdSelStartActive	=	1		# The selection is at the start of the active document.

#values from word.wdselectiontype
#****************************
wdNoSelection	=	0		# No selection.
wdSelectionBlock	=	6		# A block selection.
wdSelectionColumn	=	4		# A column selection.
wdSelectionFrame	=	3		# A frame selection.
wdSelectionInlineShape	=	7		# An inline shape selection.
wdSelectionIP	=	1		# An inline paragraph selection.
wdSelectionNormal	=	2		# A normal or user-defined selection.
wdSelectionRow	=	5		# A row selection.
wdSelectionShape	=	8		# A shape selection.

#values from word.wdseparatortype
#****************************
wdSeparatorColon	=	2		# A colon.
wdSeparatorEmDash	=	3		# An emphasized dash.
wdSeparatorEnDash	=	4		# A standard dash.
wdSeparatorHyphen	=	0		# A hyphen.
wdSeparatorPeriod	=	1		# A period.

#values from word.wdshapeposition
#****************************
wdShapeBottom	=	-999997		# At the bottom.
wdShapeCenter	=	-999995		# In the center.
wdShapeInside	=	-999994		# Inside the selected range.
wdShapeLeft	=	-999998		# On the left.
wdShapeOutside	=	-999993		# Outside the selected range.
wdShapeRight	=	-999996		# On the right.
wdShapeTop	=	-999999		# At the top.

#values from word.wdshapepositionrelative
#****************************
wdShapePositionRelativeNone	=	-999999		# Specifies that the  LeftRelative or TopRelative property is not currently valid, so the shape is positioned according to the value specified in the Left or Top property, respectively.

#values from word.wdshapesizerelative
#****************************
wdShapeSizeRelativeNone	=	-999999		# Specifies that the  WidthRelative or HeightRelative property is not currently valid, so the shape is positioned according to the value specified in the Left or Top property, respectively.

#values from word.wdshowfilter
#****************************
wdShowFilterFormattingAvailable	=	4		# All formatting available.
wdShowFilterFormattingInUse	=	3		# All formatting in use.
wdShowFilterStylesAll	=	2		# All styles.
wdShowFilterStylesAvailable	=	0		# All styles available.
wdShowFilterStylesInUse	=	1		# All styles in use.
wdShowFilterFormattingRecommended	=	5		# Only recommended styles.

#values from word.wdshowsourcedocuments
#****************************
wdShowSourceDocumentsBoth	=	3		# Shows both original and revised documents.
wdShowSourceDocumentsNone	=	0		# Shows neither the original nor the revised documents for the source document used in a Compare function.
wdShowSourceDocumentsOriginal	=	1		# Shows the original document only.
wdShowSourceDocumentsRevised	=	2		# Shows the revised document only.

#values from word.wdsmarttagcontroltype
#****************************
wdControlActiveX	=	13		# ActiveX control.
wdControlButton	=	6		# Button.
wdControlCheckbox	=	9		# Check box.
wdControlCombo	=	12		# Combo box.
wdControlDocumentFragment	=	14		# Document fragment.
wdControlDocumentFragmentURL	=	15		# Document fragment URL.
wdControlHelp	=	3		# Help.
wdControlHelpURL	=	4		# Help URL.
wdControlImage	=	8		# Image.
wdControlLabel	=	7		# Label.
wdControlLink	=	2		# Link.
wdControlListbox	=	11		# List box.
wdControlRadioGroup	=	16		# Radio group.
wdControlSeparator	=	5		# Separator.
wdControlSmartTag	=	1		# Smart tag.
wdControlTextbox	=	10		# Text box.

#values from word.wdsortfieldtype
#****************************
wdSortFieldAlphanumeric	=	0		# Alphanumeric order.
wdSortFieldDate	=	2		# Date order.
wdSortFieldJapanJIS	=	4		# Japanese JIS order.
wdSortFieldKoreaKS	=	6		# Korean KS order.
wdSortFieldNumeric	=	1		# Numeric order.
wdSortFieldStroke	=	5		# Stroke order.
wdSortFieldSyllable	=	3		# Syllable order.

#values from word.wdsortorder
#****************************
wdSortOrderAscending	=	0		# Ascending order. Default.
wdSortOrderDescending	=	1		# Descending order.

#values from word.wdsortseparator
#****************************
wdSortSeparateByCommas	=	1		# Comma.
wdSortSeparateByDefaultTableSeparator	=	2		# Default table separator.
wdSortSeparateByTabs	=	0		# Tab.

#values from word.wdspanishspeller
#****************************
wdSpanishTuteoAndVoseo	=	1		# The Spanish spelling checker recognizes both tuteo and voseo verb forms.
wdSpanishTuteoOnly	=	0		# The Spanish spelling checker recognizes only tuteo verb forms.
wdSpanishVoseoOnly	=	2		# The Spanish spelling checker recognizes only voseo verb forms.

#values from word.wdspecialpane
#****************************
wdPaneComments	=	15		# Selected comments.
wdPaneCurrentPageFooter	=	17		# The page footer.
wdPaneCurrentPageHeader	=	16		# The page header.
wdPaneEndnoteContinuationNotice	=	12		# The endnote continuation notice.
wdPaneEndnoteContinuationSeparator	=	13		# The endnote continuation separator.
wdPaneEndnotes	=	8		# Endnotes.
wdPaneEndnoteSeparator	=	14		# The endnote separator.
wdPaneEvenPagesFooter	=	6		# The even pages footer.
wdPaneEvenPagesHeader	=	3		# The even pages header.
wdPaneFirstPageFooter	=	5		# The first page footer.
wdPaneFirstPageHeader	=	2		# The first page header.
wdPaneFootnoteContinuationNotice	=	9		# The footnote continuation notice.
wdPaneFootnoteContinuationSeparator	=	10		# The footnote continuation separator.
wdPaneFootnotes	=	7		# Footnotes.
wdPaneFootnoteSeparator	=	11		# The footnote separator.
wdPaneNone	=	0		# No display.
wdPanePrimaryFooter	=	4		# The primary footer pane.
wdPanePrimaryHeader	=	1		# The primary header pane.
wdPaneRevisions	=	18		# The revisions pane.
wdPaneRevisionsHoriz	=	19		# The revisions pane displays along the bottom of the document window.
wdPaneRevisionsVert	=	20		# The revisions pane displays along the left side of the document window.

#values from word.wdspellingerrortype
#****************************
wdSpellingCapitalization	=	2		# Capitalization error.
wdSpellingCorrect	=	0		# Spelling is correct.
wdSpellingNotInDictionary	=	1		# The word is not in the specified dictionary.

#values from word.wdspellingwordtype
#****************************
wdAnagram	=	2		# Anagram searching.
wdSpellword	=	0		# Spellword searching.
wdWildcard	=	1		# Wildcard searching.

#values from word.wdstatistic
#****************************
wdStatisticCharacters	=	3		# Count of characters.
wdStatisticCharactersWithSpaces	=	5		# Count of characters including spaces.
wdStatisticFarEastCharacters	=	6		# Count of characters for Asian languages.
wdStatisticLines	=	1		# Count of lines.
wdStatisticPages	=	2		# Count of pages.
wdStatisticParagraphs	=	4		# Count of paragraphs.
wdStatisticWords	=	0		# Count of words.

#values from word.wdstorytype
#****************************
wdCommentsStory	=	4		# Comments story.
wdEndnoteContinuationNoticeStory	=	17		# Endnote continuation notice story.
wdEndnoteContinuationSeparatorStory	=	16		# Endnote continuation separator story.
wdEndnoteSeparatorStory	=	15		# Endnote separator story.
wdEndnotesStory	=	3		# Endnotes story.
wdEvenPagesFooterStory	=	8		# Even pages footer story.
wdEvenPagesHeaderStory	=	6		# Even pages header story.
wdFirstPageFooterStory	=	11		# First page footer story.
wdFirstPageHeaderStory	=	10		# First page header story.
wdFootnoteContinuationNoticeStory	=	14		# Footnote continuation notice story.
wdFootnoteContinuationSeparatorStory	=	13		# Footnote continuation separator story.
wdFootnoteSeparatorStory	=	12		# Footnote separator story.
wdFootnotesStory	=	2		# Footnotes story.
wdMainTextStory	=	1		# Main text story.
wdPrimaryFooterStory	=	9		# Primary footer story.
wdPrimaryHeaderStory	=	7		# Primary header story.
wdTextFrameStory	=	5		# Text frame story.

#values from word.wdstylesheetlinktype
#****************************
wdStyleSheetLinkTypeImported	=	1		# Imported internal style sheet.
wdStyleSheetLinkTypeLinked	=	0		# Linked external style sheet.

#values from word.wdstylesheetprecedence
#****************************
wdStyleSheetPrecedenceHigher	=	-1		# Raise precedence.
wdStyleSheetPrecedenceHighest	=	1		# Highest precedence.
wdStyleSheetPrecedenceLower	=	-2		# Lower precedence.
wdStyleSheetPrecedenceLowest	=	0		# Lowest precedence.

#values from word.wdstylesort
#****************************
wdStyleSortByBasedOn	=	3		# Sorts styles based on the item indicated in the  Sort Styles Based On option.
wdStyleSortByFont	=	2		# Sorts styles based on the name of the font used.
wdStyleSortByName	=	0		# Sorts styles alphabetically based on the name of the style.
wdStyleSortByType	=	4		# Sorts styles based on whether the style is a paragraph style or character style.
wdStyleSortRecommended	=	1		# Sorts styles based on whether they are recommended for use.

#values from word.wdstyletype
#****************************
wdStyleTypeCharacter	=	2		# Body character style.
wdStyleTypeList	=	4		# List style.
wdStyleTypeParagraph	=	1		# Paragraph style.
wdStyleTypeTable	=	3		# Table style.

#values from word.wdstylisticset
#****************************
wdStylisticSet01	=	1		# First stylistic set for the specified font.
wdStylisticSet02	=	2		# Second stylistic set for the specified font.
wdStylisticSet03	=	4		# Third stylistic set for the specified font.
wdStylisticSet04	=	8		# Fourth stylistic set for the specified font.
wdStylisticSet05	=	16		# Fifth stylistic set for the specified font.
wdStylisticSet06	=	32		# Sixth stylistic set for the specified font.
wdStylisticSet07	=	64		# Seventh stylistic set for the specified font.
wdStylisticSet08	=	128		# Eighth stylistic set for the specified font.
wdStylisticSet09	=	256		# Ninth stylistic set for the specified font.
wdStylisticSet10	=	512		# Tenth stylistic set for the specified font.
wdStylisticSet11	=	1024		# Eleventh stylistic set for the specified font.
wdStylisticSet12	=	2048		# Twelfth stylistic set for the specified font.
wdStylisticSet13	=	4096		# Thirteenth stylistic set for the specified font.
wdStylisticSet14	=	8192		# Fourtheenth stylistic set for the specified font.
wdStylisticSet15	=	16384		# Fifthteenth stylistic set for the specified font.
wdStylisticSet16	=	32768		# Sixteenth stylistic set for the specified font.
wdStylisticSet17	=	65536		# Seventeenth stylistic set for the specified font.
wdStylisticSet18	=	131072		# Eighteenth stylistic set for the specified font.
wdStylisticSet19	=	262144		# Nineteenth stylistic set for the specified font.
wdStylisticSet20	=	524288		# Twentieth stylistic set for the specified font.
wdStylisticSetDefault	=	0		# Default stylistic set for the specified font.

#values from word.wdsubscriberformats
#****************************
wdSubscriberBestFormat	=	0		# Not supported.
wdSubscriberPict	=	4		# Not supported.
wdSubscriberRTF	=	1		# Not supported.
wdSubscriberText	=	2		# Not supported.

#values from word.wdtabalignment
#****************************
wdAlignTabBar	=	4		# Bar-aligned.
wdAlignTabCenter	=	1		# Center-aligned.
wdAlignTabDecimal	=	3		# Decimal-aligned.
wdAlignTabLeft	=	0		# Left-aligned.
wdAlignTabList	=	6		# List-aligned.
wdAlignTabRight	=	2		# Right-aligned.

#values from word.wdtableader
#****************************
wdTabLeaderDashes	=	2		# Dashes.
wdTabLeaderDots	=	1		# Dots.
wdTabLeaderHeavy	=	4		# A heavy line.
wdTabLeaderLines	=	3		# Double lines.
wdTabLeaderMiddleDot	=	5		# A middle dot.
wdTabLeaderSpaces	=	0		# Spaces. Default.

#values from word.wdtabledirection
#****************************
wdTableDirectionLtr	=	1		# The selected rows are arranged with the first column in the leftmost position.
wdTableDirectionRtl	=	0		# The selected rows are arranged with the first column in the rightmost position.

#values from word.wdtablefieldseparator
#****************************
wdSeparateByCommas	=	2		# A comma.
wdSeparateByDefaultListSeparator	=	3		# The default list separator.
wdSeparateByParagraphs	=	0		# Paragraph markers.
wdSeparateByTabs	=	1		# A tab.

#values from word.wdtableformat
#****************************
wdTableFormat3DEffects1	=	32		# 3D effects format number 1.
wdTableFormat3DEffects2	=	33		# 3D effects format number 2.
wdTableFormat3DEffects3	=	34		# 3D effects format number 3.
wdTableFormatClassic1	=	4		# Classic format number 1.
wdTableFormatClassic2	=	5		# Classic format number 2.
wdTableFormatClassic3	=	6		# Classic format number 3.
wdTableFormatClassic4	=	7		# Classic format number 4.
wdTableFormatColorful1	=	8		# Colorful format number 1.
wdTableFormatColorful2	=	9		# Colorful format number 2.
wdTableFormatColorful3	=	10		# Colorful format number 3.
wdTableFormatColumns1	=	11		# Columns format number 1.
wdTableFormatColumns2	=	12		# Columns format number 2.
wdTableFormatColumns3	=	13		# Columns format number 3.
wdTableFormatColumns4	=	14		# Columns format number 4.
wdTableFormatColumns5	=	15		# Columns format number 5.
wdTableFormatContemporary	=	35		# Contemporary format.
wdTableFormatElegant	=	36		# Elegant format.
wdTableFormatGrid1	=	16		# Grid format number 1.
wdTableFormatGrid2	=	17		# Grid format number 2.
wdTableFormatGrid3	=	18		# Grid format number 3.
wdTableFormatGrid4	=	19		# Grid format number 4.
wdTableFormatGrid5	=	20		# Grid format number 5.
wdTableFormatGrid6	=	21		# Grid format number 6.
wdTableFormatGrid7	=	22		# Grid format number 7.
wdTableFormatGrid8	=	23		# Grid format number 8.
wdTableFormatList1	=	24		# List format number 1.
wdTableFormatList2	=	25		# List format number 2.
wdTableFormatList3	=	26		# List format number 3.
wdTableFormatList4	=	27		# List format number 4.
wdTableFormatList5	=	28		# List format number 5.
wdTableFormatList6	=	29		# List format number 6.
wdTableFormatList7	=	30		# List format number 7.
wdTableFormatList8	=	31		# List format number 8.
wdTableFormatNone	=	0		# No formatting.
wdTableFormatProfessional	=	37		# Professional format.
wdTableFormatSimple1	=	1		# Simple format number 1.
wdTableFormatSimple2	=	2		# Simple format number 2.
wdTableFormatSimple3	=	3		# Simple format number 3.
wdTableFormatSubtle1	=	38		# Subtle format number 1.
wdTableFormatSubtle2	=	39		# Subtle format number 2.
wdTableFormatWeb1	=	40		# Web format number 1.
wdTableFormatWeb2	=	41		# Web format number 2.
wdTableFormatWeb3	=	42		# Web format number 3.

#values from word.wdtableformatapply
#****************************
wdTableFormatApplyAutoFit	=	16		# AutoFit.
wdTableFormatApplyBorders	=	1		# Borders.
wdTableFormatApplyColor	=	8		# Color.
wdTableFormatApplyFirstColumn	=	128		# Apply AutoFormat to first column.
wdTableFormatApplyFont	=	4		# Font.
wdTableFormatApplyHeadingRows	=	32		# Apply AutoFormat to heading rows.
wdTableFormatApplyLastColumn	=	256		# Apply AutoFormat to last column.
wdTableFormatApplyLastRow	=	64		# Apply AutoFormat to last row.
wdTableFormatApplyShading	=	2		# Shading.

#values from word.wdtableposition
#****************************
wdTableBottom	=	-999997		# At the bottom of the document.
wdTableCenter	=	-999995		# Centered.
wdTableInside	=	-999994		# Placed inside a range.
wdTableLeft	=	-999998		# Aligned to the left side of the document.
wdTableOutside	=	-999993		# Placed outside a range.
wdTableRight	=	-999996		# Aligned to the right side of the document.
wdTableTop	=	-999999		# At the top of the document.

#values from word.wdtaskpanes
#****************************
wdTaskPaneApplyStyles	=	17		# Apply styles pane.
wdTaskPaneDocumentActions	=	7		# Document actions pane.
wdTaskPaneDocumentProtection	=	6		# Document protection pane.
wdTaskPaneFaxService	=	11		# Fax service pane.
wdTaskPaneFormatting	=	0		# Formatting pane.
wdTaskPaneHelp	=	9		# Help pane.
wdTaskPaneMailMerge	=	2		# Mail merge pane.
wdTaskPaneProofing	=	20		# Proofing pane.
wdTaskPaneResearch	=	10		# Research pane.
wdTaskPaneRevealFormatting	=	1		# Reveal formatting codes pane.
wdTaskPaneRevPaneFlex	=	22		# Revisions pane flex pane.
wdTaskPaneSearch	=	4		# Search pane.
wdTaskPaneSignature	=	14		# Signature pane.
wdTaskPaneStyleInspector	=	15		# Style inspector pane.
wdTaskPaneThesaurus	=	23		# Thesaurus pane.
wdTaskPaneTranslate	=	3		# Translate pane.
wdTaskPaneXMLDocument	=	12		# XML document pane.
wdTaskPaneXMLMapping	=	21		# XML mapping pane.
wdTaskPaneXMLStructure	=	5		# XML structure pane.

#values from word.wdtcscconverterdirection
#****************************
wdTCSCConverterDirectionAuto	=	2		# Convert in the appropriate direction based on the detected language of the specified range.
wdTCSCConverterDirectionSCTC	=	0		# Convert from Simplified Chinese to Traditional Chinese.
wdTCSCConverterDirectionTCSC	=	1		# Convert from Traditional Chinese to Simplified Chinese.

#values from word.wdtemplatetype
#****************************
wdAttachedTemplate	=	2		# An attached template.
wdGlobalTemplate	=	1		# A global template.
wdNormalTemplate	=	0		# The normal default template.

#values from word.wdtextboxtightwrap
#****************************
wdTightAll	=	1		# Wraps text around the text box tightly to the contents of the text box on all lines.
wdTightFirstAndLastLines	=	2		# Wraps text tightly only on first and last lines.
wdTightFirstLineOnly	=	3		# Wraps text tightly only on the first line.
wdTightLastLineOnly	=	4		# Wraps text tightly only on the last line.
wdTightNone	=	0		# Does not wrap text tightly around the contents of a text box.

#values from word.wdtextformfieldtype
#****************************
wdCalculationText	=	5		# Calculation text field.
wdCurrentDateText	=	3		# Current date text field.
wdCurrentTimeText	=	4		# Current time text field.
wdDateText	=	2		# Date text field.
wdNumberText	=	1		# Number text field.
wdRegularText	=	0		# Regular text field.

#values from word.wdtextorientation
#****************************
wdTextOrientationDownward	=	3		# Text flows downward on a slope.
wdTextOrientationHorizontal	=	0		# Text flows horizontally. Default.
wdTextOrientationHorizontalRotatedFarEast	=	4		# Text flows horizontally but from right to left to accommodate right-to-left languages.
wdTextOrientationUpward	=	2		# Text flows upward on a slope.
wdTextOrientationVerticalFarEast	=	1		# Text flows vertically and reads downward from the top, right to left.
wdTextOrientationVertical	=	5		# Text flows vertically and reads downward from the top, left to right.

#values from word.wdtextureindex
#****************************
wdTexture10Percent	=	100		# 10 percent shading.
wdTexture12Pt5Percent	=	125		# 12.5 percent shading.
wdTexture15Percent	=	150		# 15 percent shading.
wdTexture17Pt5Percent	=	175		# 17.5 percent shading.
wdTexture20Percent	=	200		# 20 percent shading.
wdTexture22Pt5Percent	=	225		# 22.5 percent shading.
wdTexture25Percent	=	250		# 25 percent shading.
wdTexture27Pt5Percent	=	275		# 27.5 percent shading.
wdTexture2Pt5Percent	=	25		# 2.5 percent shading.
wdTexture30Percent	=	300		# 30 percent shading.
wdTexture32Pt5Percent	=	325		# 32.5 percent shading.
wdTexture35Percent	=	350		# 35 percent shading.
wdTexture37Pt5Percent	=	375		# 37.5 percent shading.
wdTexture40Percent	=	400		# 40 percent shading.
wdTexture42Pt5Percent	=	425		# 42.5 percent shading.
wdTexture45Percent	=	450		# 45 percent shading.
wdTexture47Pt5Percent	=	475		# 47.5 percent shading.
wdTexture50Percent	=	500		# 50 percent shading.
wdTexture52Pt5Percent	=	525		# 52.5 percent shading.
wdTexture55Percent	=	550		# 55 percent shading.
wdTexture57Pt5Percent	=	575		# 57.5 percent shading.
wdTexture5Percent	=	50		# 5 percent shading.
wdTexture60Percent	=	600		# 60 percent shading.
wdTexture62Pt5Percent	=	625		# 62.5 percent shading.
wdTexture65Percent	=	650		# 65 percent shading.
wdTexture67Pt5Percent	=	675		# 67.5 percent shading.
wdTexture70Percent	=	700		# 70 percent shading.
wdTexture72Pt5Percent	=	725		# 72.5 percent shading.
wdTexture75Percent	=	750		# 75 percent shading.
wdTexture77Pt5Percent	=	775		# 77.5 percent shading.
wdTexture7Pt5Percent	=	75		# 7.5 percent shading.
wdTexture80Percent	=	800		# 80 percent shading.
wdTexture82Pt5Percent	=	825		# 82.5 percent shading.
wdTexture85Percent	=	850		# 85 percent shading.
wdTexture87Pt5Percent	=	875		# 87.5 percent shading.
wdTexture90Percent	=	900		# 90 percent shading.
wdTexture92Pt5Percent	=	925		# 92.5 percent shading.
wdTexture95Percent	=	950		# 95 percent shading.
wdTexture97Pt5Percent	=	975		# 97.5 percent shading.
wdTextureCross	=	-11		# Horizontal cross shading.
wdTextureDarkCross	=	-5		# Dark horizontal cross shading.
wdTextureDarkDiagonalCross	=	-6		# Dark diagonal cross shading.
wdTextureDarkDiagonalDown	=	-3		# Dark diagonal down shading.
wdTextureDarkDiagonalUp	=	-4		# Dark diagonal up shading.
wdTextureDarkHorizontal	=	-1		# Dark horizontal shading.
wdTextureDarkVertical	=	-2		# Dark vertical shading.
wdTextureDiagonalCross	=	-12		# Diagonal cross shading.
wdTextureDiagonalDown	=	-9		# Diagonal down shading.
wdTextureDiagonalUp	=	-10		# Diagonal up shading.
wdTextureHorizontal	=	-7		# Horizontal shading.
wdTextureNone	=	0		# No shading.
wdTextureSolid	=	1000		# Solid shading.
wdTextureVertical	=	-8		# Vertical shading.

#values from word.wdthemecolorindex
#****************************
wdNotThemeColor	=	-1		# No color.
wdThemeColorAccent1	=	4		# Accent color 1.
wdThemeColorAccent2	=	5		# Accent color 2.
wdThemeColorAccent3	=	6		# Accent color 3.
wdThemeColorAccent4	=	7		# Accent color 4.
wdThemeColorAccent5	=	8		# Accent color 5.
wdThemeColorAccent6	=	9		# Accent color 6.
wdThemeColorBackground1	=	12		# Background color 1.
wdThemeColorBackground2	=	14		# Background color 2.
wdThemeColorHyperlink	=	10		# Hyperlink color.
wdThemeColorHyperlinkFollowed	=	11		# Followed hyperlink color.
wdThemeColorMainDark1	=	0		# Dark main color 1.
wdThemeColorMainDark2	=	2		# Dark main color 2.
wdThemeColorMainLight1	=	1		# Light main color 1.
wdThemeColorMainLight2	=	3		# Light main color 2.
wdThemeColorText1	=	13		# Text color 1.
wdThemeColorText2	=	15		# Text color 2.

#values from word.wdtoaformat
#****************************
wdTOAClassic	=	1		# Classic formatting.
wdTOADistinctive	=	2		# Distinctive formatting.
wdTOAFormal	=	3		# Formal formatting.
wdTOASimple	=	4		# Simple formatting.
wdTOATemplate	=	0		# Template formatting.

#values from word.wdtocformat
#****************************
wdTOCClassic	=	1		# Classic formatting.
wdTOCDistinctive	=	2		# Distinctive formatting.
wdTOCFancy	=	3		# Fancy formatting.
wdTOCFormal	=	5		# Formal formatting.
wdTOCModern	=	4		# Modern formatting.
wdTOCSimple	=	6		# Simple formatting.
wdTOCTemplate	=	0		# Template formatting.

#values from word.wdtofformat
#****************************
wdTOFCentered	=	3		# Centered formatting.
wdTOFClassic	=	1		# Classic formatting.
wdTOFDistinctive	=	2		# Distinctive formatting.
wdTOFFormal	=	4		# Formal formatting.
wdTOFSimple	=	5		# Simple formatting.
wdTOFTemplate	=	0		# Template formatting.

#values from word.wdtrailingcharacter
#****************************
wdTrailingNone	=	2		# No character is inserted.
wdTrailingSpace	=	1		# A space is inserted. Default.
wdTrailingTab	=	0		# A tab is inserted.

#values from word.wdtwolinesinonetype
#****************************
wdTwoLinesInOneAngleBrackets	=	4		# Enclose the lines using angle brackets.
wdTwoLinesInOneCurlyBrackets	=	5		# Enclose the lines using curly brackets.
wdTwoLinesInOneNoBrackets	=	1		# Use no enclosing character.
wdTwoLinesInOneNone	=	0		# Restore the two lines of text written into one to two separate lines.
wdTwoLinesInOneParentheses	=	2		# Enclose the lines using parentheses.
wdTwoLinesInOneSquareBrackets	=	3		# Enclose the lines using square brackets.

#values from word.wdunderline
#****************************
wdUnderlineDash	=	7		# Dashes.
wdUnderlineDashHeavy	=	23		# Heavy dashes.
wdUnderlineDashLong	=	39		# Long dashes.
wdUnderlineDashLongHeavy	=	55		# Long heavy dashes.
wdUnderlineDotDash	=	9		# Alternating dots and dashes.
wdUnderlineDotDashHeavy	=	25		# Alternating heavy dots and heavy dashes.
wdUnderlineDotDotDash	=	10		# An alternating dot-dot-dash pattern.
wdUnderlineDotDotDashHeavy	=	26		# An alternating heavy dot-dot-dash pattern.
wdUnderlineDotted	=	4		# Dots.
wdUnderlineDottedHeavy	=	20		# Heavy dots.
wdUnderlineDouble	=	3		# A double line.
wdUnderlineNone	=	0		# No underline.
wdUnderlineSingle	=	1		# A single line. default.
wdUnderlineThick	=	6		# A single thick line.
wdUnderlineWavy	=	11		# A single wavy line.
wdUnderlineWavyDouble	=	43		# A double wavy line.
wdUnderlineWavyHeavy	=	27		# A heavy wavy line.
wdUnderlineWords	=	2		# Underline individual words only.

#values from word.wdunits
#****************************
wdCell	=	12		# A cell.
wdCharacter	=	1		# A character.
wdCharacterFormatting	=	13		# Character formatting.
wdColumn	=	9		# A column.
wdItem	=	16		# The selected item.
wdLine	=	5		# A line.
wdParagraph	=	4		# A paragraph.
wdParagraphFormatting	=	14		# Paragraph formatting.
wdRow	=	10		# A row.
wdScreen	=	7		# The screen dimensions.
wdSection	=	8		# A section.
wdSentence	=	3		# A sentence.
wdStory	=	6		# A story.
wdTable	=	15		# A table.
wdWindow	=	11		# A window.
wdWord	=	2		# A word.

#values from word.wdupdatestylelistbehavior
#****************************
wdListBehaviorAddBulletsNumbering	=	1		# Adds the numbering or bullets pattern of the selection to all paragraphs in the document that use the same style.
wdListBehaviorKeepPreviousPattern	=	0		# Keeps the existing numbering or bullets pattern for all other paragraphs that use the same style and does not apply the numbering or bullets pattern of the selection.

#values from word.wduseformattingfrom
#****************************
wdFormattingFromCurrent	=	0		# Copy source formatting from the current item.
wdFormattingFromPrompt	=	2		# Prompt the user for formatting to use.
wdFormattingFromSelected	=	1		# Copy source formatting from the current selection.

#values from word.wdverticalalignment
#****************************
wdAlignVerticalBottom	=	3		# Bottom vertical alignment.
wdAlignVerticalCenter	=	1		# Center vertical alignment.
wdAlignVerticalJustify	=	2		# Justified vertical alignment.
wdAlignVerticalTop	=	0		# Top vertical alignment.

#values from word.wdviewtype
#****************************
wdMasterView	=	5		# A master view.
wdNormalView	=	1		# A normal view.
wdOutlineView	=	2		# An outline view.
wdPrintPreview	=	4		# A print preview view.
wdPrintView	=	3		# A print view.
wdReadingView	=	7		# A reading view.
wdWebView	=	6		# A Web view.

#values from word.wdvisualselection
#****************************
wdVisualSelectionBlock	=	0		# All selected lines are the same width.
wdVisualSelectionContinuous	=	1		# The selection wraps from line to line.

#values from word.wdwindowstate
#****************************
wdWindowStateMaximize	=	1		# Maximized.
wdWindowStateMinimize	=	2		# Minimized.
wdWindowStateNormal	=	0		# Normal.

#values from word.wdwindowtype
#****************************
wdWindowDocument	=	0		# A document window.
wdWindowTemplate	=	1		# A template window.

#values from word.wdworddialog
#****************************
wdDialogBuildingBlockOrganizer	=	2067		# (none)
wdDialogConnect	=	420		# Drive, Path, Password
wdDialogConsistencyChecker	=	1121		# (none)
wdDialogContentControlProperties	=	2394		# (none)
wdDialogControlRun	=	235		# Application
wdDialogConvertObject	=	392		# IconNumber, ActivateAs, IconFileName, Caption, Class, DisplayIcon, Floating
wdDialogCopyFile	=	300		# FileName, Directory
wdDialogCreateAutoText	=	872		# (none)
wdDialogCreateSource	=	1922		# (none)
wdDialogCSSLinks	=	1261		# LinkStyles
wdDialogDocumentInspector	=	1482		# (none)
wdDialogDocumentStatistics	=	78		# FileName, Directory, Template, Title, Created, LastSaved, LastSavedBy, Revision, Time, Printed, Pages, Words, Characters, Paragraphs, Lines, FileSize
wdDialogDrawAlign	=	634		# Horizontal, Vertical, RelativeTo
wdDialogDrawSnapToGrid	=	633		# SnapToGrid, XGrid, YGrid, XOrigin, YOrigin, SnapToShapes, XGridDisplay, YGridDisplay, FollowMargins, ViewGridLines, DefineLineBasedOnGrid
wdDialogEditAutoText	=	985		# Name, Context, InsertAs, Insert, Add, Define, InsertAsText, Delete, CompleteAT
wdDialogEditCreatePublisher	=	732		# Macintosh-only. For information about this constant, consult the language reference Help included with Microsoft Office Macintosh Edition.
wdDialogEditFind	=	112		# Find, Replace, Direction, MatchCase, WholeWord, PatternMatch, SoundsLike, FindNext, ReplaceOne, ReplaceAll, Format, Wrap, FindAllWordForms, MatchByte, FuzzyFind, Destination, CorrectEnd, MatchKashida, MatchDiacritics, MatchAlefHamza, MatchControl
wdDialogEditFrame	=	458		# Wrap, WidthRule, FixedWidth, HeightRule, FixedHeight, PositionHorz, PositionHorzRel, DistFromText, PositionVert, PositionVertRel, DistVertFromText, MoveWithText, LockAnchor, RemoveFrame
wdDialogEditGoTo	=	896		# Find, Replace, Direction, MatchCase, WholeWord, PatternMatch, SoundsLike, FindNext, ReplaceOne, ReplaceAll, Format, Wrap, FindAllWordForms, MatchByte, FuzzyFind, Destination, CorrectEnd, MatchKashida, MatchDiacritics, MatchAlefHamza, MatchControl
wdDialogEditGoToOld	=	811		# (none)
wdDialogEditLinks	=	124		# UpdateMode, Locked, SavePictureInDoc, UpdateNow, OpenSource, KillLink, Link, Application, Item, FileName, PreserveFormatLinkUpdate
wdDialogEditObject	=	125		# Verb
wdDialogEditPasteSpecial	=	111		# IconNumber, Link, DisplayIcon, Class, DataType, IconFileName, Caption, Floating
wdDialogEditPublishOptions	=	735		# Macintosh-only. For information about this constant, consult the language reference Help included with Microsoft Office Macintosh Edition.
wdDialogEditReplace	=	117		# Find, Replace, Direction, MatchCase, WholeWord, PatternMatch, SoundsLike, FindNext, ReplaceOne, ReplaceAll, Format, Wrap, FindAllWordForms, MatchByte, FuzzyFind, Destination, CorrectEnd, MatchKashida, MatchDiacritics, MatchAlefHamza, MatchControl
wdDialogEditStyle	=	120		# (none)
wdDialogEditSubscribeOptions	=	736		# Macintosh-only. For information about this constant, consult the language reference Help included with Microsoft Office Macintosh Edition.
wdDialogEditSubscribeTo	=	733		# Macintosh-only. For information about this constant, consult the language reference Help included with Microsoft Office Macintosh Edition.
wdDialogEditTOACategory	=	625		# Category, CategoryName
wdDialogEmailOptions	=	863		# (none)
wdDialogFileDocumentLayout	=	178		# Tab, PaperSize, TopMargin, BottomMargin, LeftMargin, RightMargin, Gutter, PageWidth, PageHeight, Orientation, FirstPage, OtherPages, VertAlign, ApplyPropsTo, Default, FacingPages, HeaderDistance, FooterDistance, SectionStart, OddAndEvenPages, DifferentFirstPage, Endnotes, LineNum, StartingNum, FromText, CountBy, NumMode, TwoOnOne, GutterPosition, LayoutMode, CharsLine, LinesPage, CharPitch, LinePitch, DocFontName, DocFontSize, PageColumns, TextFlow, FirstPageOnLeft, SectionType, RTLAlignment
wdDialogFileFind	=	99		# SearchName, SearchPath, Name, SubDir, Title, Author, Keywords, Subject, Options, MatchCase, Text, PatternMatch, DateSavedFrom, DateSavedTo, SavedBy, DateCreatedFrom, DateCreatedTo, View, SortBy, ListBy, SelectedFile, Add, Delete, ShowFolders, MatchByte
wdDialogFileMacCustomPageSetupGX	=	737		# Macintosh-only. For information about this constant, consult the language reference Help included with Microsoft Office Macintosh Edition.
wdDialogFileMacPageSetup	=	685		# Macintosh-only. For information about this constant, consult the language reference Help included with Microsoft Office Macintosh Edition.
wdDialogFileMacPageSetupGX	=	444		# Macintosh-only. For information about this constant, consult the language reference Help included with Microsoft Office Macintosh Edition.
wdDialogFileNew	=	79		# Template, NewTemplate, DocumentType, Visible
wdDialogFileOpen	=	80		# Name, ConfirmConversions, ReadOnly, LinkToSource, AddToMru, PasswordDoc, PasswordDot, Revert, WritePasswordDoc, WritePasswordDot, Connection, SQLStatement, SQLStatement1, Format, Encoding, Visible, OpenExclusive, OpenAndRepair, SubType, DocumentDirection, NoEncodingDialog, XMLTransform
wdDialogFilePageSetup	=	178		# Tab, PaperSize, TopMargin, BottomMargin, LeftMargin, RightMargin, Gutter, PageWidth, PageHeight, Orientation, FirstPage, OtherPages, VertAlign, ApplyPropsTo, Default, FacingPages, HeaderDistance, FooterDistance, SectionStart, OddAndEvenPages, DifferentFirstPage, Endnotes, LineNum, StartingNum, FromText, CountBy, NumMode, TwoOnOne, GutterPosition, LayoutMode, CharsLine, LinesPage, CharPitch, LinePitch, DocFontName, DocFontSize, PageColumns, TextFlow, FirstPageOnLeft, SectionType, RTLAlignment, FolioPrint
wdDialogFilePrint	=	88		# Background, AppendPrFile, Range, PrToFileName, From, To, Type, NumCopies, Pages, Order, PrintToFile, Collate, FileName, Printer, OutputPrinter, DuplexPrint, PrintZoomColumn, PrintZoomRow, PrintZoomPaperWidth, PrintZoomPaperHeight, ZoomPaper
wdDialogFilePrintOneCopy	=	445		# Macintosh-only. For information about this constant, consult the language reference Help included with Microsoft Office Macintosh Edition.
wdDialogFilePrintSetup	=	97		# Printer, Options, Network, DoNotSetAsSysDefault
wdDialogFileRoutingSlip	=	624		# Subject, Message, AllAtOnce, ReturnWhenDone, TrackStatus, Protect, AddSlip, RouteDocument, AddRecipient, OldRecipient, ResetSlip, ClearSlip, ClearRecipients, Address
wdDialogFileSaveAs	=	84		# Name, Format, LockAnnot, Password, AddToMru, WritePassword, RecommendReadOnly, EmbedFonts, NativePictureFormat, FormsData, SaveAsAOCELetter, WriteVersion, VersionDesc, InsertLineBreaks, AllowSubstitutions, LineEnding, AddBiDiMarks
wdDialogFileSaveVersion	=	1007		# (none)
wdDialogFileSummaryInfo	=	86		# Title, Subject, Author, Keywords, Comments, FileName, Directory, Template, CreateDate, LastSavedDate, LastSavedBy, RevisionNumber, EditTime, LastPrintedDate, NumPages, NumWords, NumChars, NumParas, NumLines, Update, FileSize
wdDialogFileVersions	=	945		# AutoVersion, VersionDesc
wdDialogFitText	=	983		# FitTextWidth
wdDialogFontSubstitution	=	581		# UnavailableFont, SubstituteFont
wdDialogFormatAddrFonts	=	103		# Points, Underline, Color, StrikeThrough, Superscript, Subscript, Hidden, SmallCaps, AllCaps, Spacing, Position, Kerning, KerningMin, Default, Tab, Font, Bold, Italic, DoubleStrikeThrough, Shadow, Outline, Emboss, Engrave, Scale, Animations, CharAccent, FontMajor, FontLowAnsi, FontHighAnsi, CharacterWidthGrid, ColorRGB, UnderlineColor, PointsBi, ColorBi, FontNameBi, BoldBi, ItalicBi, DiacColor
wdDialogFormatBordersAndShading	=	189		# ApplyTo, Shadow, TopBorder, LeftBorder, BottomBorder, RightBorder, HorizBorder, VertBorder, TopColor, LeftColor, BottomColor, RightColor, HorizColor, VertColor, FromText, Shading, Foreground, Background, Tab, FineShading, TopStyle, LeftStyle, BottomStyle, RightStyle, HorizStyle, VertStyle, TopWeight, LeftWeight, BottomWeight, RightWeight, HorizWeight, VertWeight, BorderObjectType, BorderArtWeight, BorderArt, FromTextTop, FromTextBottom, FromTextLeft, FromTextRight, OffsetFrom, InFront, SurroundHeader, SurroundFooter, JoinBorder, LineColor, WhichPages, TL2BRBorder, TR2BLBorder, TL2BRColor, TR2BLColor, TL2BRStyle, TR2BLStyle, TL2BRWeight, TR2BLWeight, ForegroundRGB, BackgroundRGB, TopColorRGB, LeftColorRGB, BottomColorRGB, RightColorRGB, HorizColorRGB, VertColorRGB, TL2BRColorRGB, TR2BLColorRGB, LineColorRGB
wdDialogFormatBulletsAndNumbering	=	824		# (none)
wdDialogFormatCallout	=	610		# Type, Gap, Angle, Drop, Length, Border, AutoAttach, Accent
wdDialogFormatChangeCase	=	322		# Type
wdDialogFormatColumns	=	177		# Columns, ColumnNo, ColumnWidth, ColumnSpacing, EvenlySpaced, ApplyColsTo, ColLine, StartNewCol, FlowColumnsRtl
wdDialogFormatDefineStyleBorders	=	185		# ApplyTo, Shadow, TopBorder, LeftBorder, BottomBorder, RightBorder, HorizBorder, VertBorder, TopColor, LeftColor, BottomColor, RightColor, HorizColor, VertColor, FromText, Shading, Foreground, Background, Tab, FineShading, TopStyle, LeftStyle, BottomStyle, RightStyle, HorizStyle, VertStyle, TopWeight, LeftWeight, BottomWeight, RightWeight, HorizWeight, VertWeight, BorderObjectType, BorderArtWeight, BorderArt, FromTextTop, FromTextBottom, FromTextLeft, FromTextRight, OffsetFrom, InFront, SurroundHeader, SurroundFooter, JoinBorder, LineColor, WhichPages, TL2BRBorder, TR2BLBorder, TL2BRColor, TR2BLColor, TL2BRStyle, TR2BLStyle, TL2BRWeight, TR2BLWeight, ForegroundRGB, BackgroundRGB, TopColorRGB, LeftColorRGB, BottomColorRGB, RightColorRGB, HorizColorRGB, VertColorRGB, TL2BRColorRGB, TR2BLColorRGB, LineColorRGB
wdDialogFormatDefineStyleFont	=	181		# Points, Underline, Color, StrikeThrough, Superscript, Subscript, Hidden, SmallCaps, AllCaps, Spacing, Position, Kerning, KerningMin, Default, Tab, Font, Bold, Italic, DoubleStrikeThrough, Shadow, Outline, Emboss, Engrave, Scale, Animations, CharAccent, FontMajor, FontLowAnsi, FontHighAnsi, CharacterWidthGrid, ColorRGB, UnderlineColor, PointsBi, ColorBi, FontNameBi, BoldBi, ItalicBi, DiacColor
wdDialogFormatDefineStyleFrame	=	184		# Wrap, WidthRule, FixedWidth, HeightRule, FixedHeight, PositionHorz, PositionHorzRel, DistFromText, PositionVert, PositionVertRel, DistVertFromText, MoveWithText, LockAnchor, RemoveFrame
wdDialogFormatDefineStyleLang	=	186		# Language, CheckLanguage, Default, NoProof
wdDialogFormatDefineStylePara	=	182		# LeftIndent, RightIndent, Before, After, LineSpacingRule, LineSpacing, Alignment, WidowControl, KeepWithNext, KeepTogether, PageBreak, NoLineNum, DontHyphen, Tab, FirstIndent, OutlineLevel, Kinsoku, WordWrap, OverflowPunct, TopLinePunct, AutoSpaceDE, LineHeightGrid, AutoSpaceDN, CharAlign, CharacterUnitLeftIndent, AdjustRight, CharacterUnitFirstIndent, CharacterUnitRightIndent, LineUnitBefore, LineUnitAfter, NoSpaceBetweenParagraphsOfSameStyle, OrientationBi
wdDialogFormatDefineStyleTabs	=	183		# Position, DefTabs, Align, Leader, Set, Clear, ClearAll
wdDialogFormatDrawingObject	=	960		# Left, PositionHorzRel, Top, PositionVertRel, LockAnchor, FloatOverText, LayoutInCell, WrapSide, TopDistanceFromText, BottomDistanceFromText, LeftDistanceFromText, RightDistanceFromText, Wrap, WordWrap, AutoSize, HRWidthType, HRHeight, HRNoshade, HRAlign, Text, AllowOverlap, HorizRule
wdDialogFormatDropCap	=	488		# Position, Font, DropHeight, DistFromText
wdDialogFormatEncloseCharacters	=	1162		# Style, Text, Enclosure
wdDialogFormatFont	=	174		# Points, Underline, Color, StrikeThrough, Superscript, Subscript, Hidden, SmallCaps, AllCaps, Spacing, Position, Kerning, KerningMin, Default, Tab, Font, Bold, Italic, DoubleStrikeThrough, Shadow, Outline, Emboss, Engrave, Scale, Animations, CharAccent, FontMajor, FontLowAnsi, FontHighAnsi, CharacterWidthGrid, ColorRGB, UnderlineColor, PointsBi, ColorBi, FontNameBi, BoldBi, ItalicBi, DiacColor
wdDialogFormatFrame	=	190		# Wrap, WidthRule, FixedWidth, HeightRule, FixedHeight, PositionHorz, PositionHorzRel, DistFromText, PositionVert, PositionVertRel, DistVertFromText, MoveWithText, LockAnchor, RemoveFrame
wdDialogFormatPageNumber	=	298		# ChapterNumber, NumRestart, NumFormat, StartingNum, Level, Separator, DoubleQuote, PgNumberingStyle
wdDialogFormatParagraph	=	175		# LeftIndent, RightIndent, Before, After, LineSpacingRule, LineSpacing, Alignment, WidowControl, KeepWithNext, KeepTogether, PageBreak, NoLineNum, DontHyphen, Tab, FirstIndent, OutlineLevel, Kinsoku, WordWrap, OverflowPunct, TopLinePunct, AutoSpaceDE, LineHeightGrid, AutoSpaceDN, CharAlign, CharacterUnitLeftIndent, AdjustRight, CharacterUnitFirstIndent, CharacterUnitRightIndent, LineUnitBefore, LineUnitAfter, NoSpaceBetweenParagraphsOfSameStyle, OrientationBi
wdDialogFormatPicture	=	187		# SetSize, CropLeft, CropRight, CropTop, CropBottom, ScaleX, ScaleY, SizeX, SizeY
wdDialogFormatRetAddrFonts	=	221		# Points, Underline, Color, StrikeThrough, Superscript, Subscript, Hidden, SmallCaps, AllCaps, Spacing, Position, Kerning, KerningMin, Default, Tab, Font, Bold, Italic, DoubleStrikeThrough, Shadow, Outline, Emboss, Engrave, Scale, Animations, CharAccent, FontMajor, FontLowAnsi, FontHighAnsi, CharacterWidthGrid, ColorRGB, UnderlineColor, PointsBi, ColorBi, FontNameBi, BoldBi, ItalicBi, DiacColor
wdDialogFormatSectionLayout	=	176		# SectionStart, VertAlign, Endnotes, LineNum, StartingNum, FromText, CountBy, NumMode, SectionType
wdDialogFormatStyle	=	180		# Name, Delete, Merge, NewName, BasedOn, NextStyle, Type, FileName, Source, AddToTemplate, Define, Rename, Apply, New, Link
wdDialogFormatStyleGallery	=	505		# Template, Preview
wdDialogFormatStylesCustom	=	1248		# (none)
wdDialogFormatTabs	=	179		# Position, DefTabs, Align, Leader, Set, Clear, ClearAll
wdDialogFormatTheme	=	855		# (none)
wdDialogFormattingRestrictions	=	1427		# (none)
wdDialogFormFieldHelp	=	361		# (none)
wdDialogFormFieldOptions	=	353		# Entry, Exit, Name, Enable, TextType, TextWidth, TextDefault, TextFormat, CheckSize, CheckWidth, CheckDefault, Type, OwnHelp, HelpText, OwnStat, StatText, Calculate
wdDialogFrameSetProperties	=	1074		# (none)
wdDialogHelpAbout	=	9		# APPNAME, APPCOPYRIGHT, APPUSERNAME, APPORGANIZATION, APPSERIALNUMBER
wdDialogHelpWordPerfectHelp	=	10		# WPCommand, HelpText, DemoGuidance
wdDialogHelpWordPerfectHelpOptions	=	511		# CommandKeyHelp, DocNavKeys, MouseSimulation, DemoGuidance, DemoSpeed, HelpType
wdDialogHorizontalInVertical	=	1160		# (none)
wdDialogIMESetDefault	=	1094		# (none)
wdDialogInsertAddCaption	=	402		# Name
wdDialogInsertAutoCaption	=	359		# Clear, ClearAll, Object, Label, Position
wdDialogInsertBookmark	=	168		# Name, SortBy, Add, Delete, Goto, Hidden
wdDialogInsertBreak	=	159		# Type
wdDialogInsertCaption	=	357		# Label, TitleAutoText, Title, Delete, Position, AutoCaption, ExcludeLabel
wdDialogInsertCaptionNumbering	=	358		# Label, FormatNumber, ChapterNumber, Level, Separator, CapNumberingStyle
wdDialogInsertCrossReference	=	367		# ReferenceType, ReferenceKind, ReferenceItem, InsertAsHyperLink, InsertPosition, SeparateNumbers, SeparatorCharacters
wdDialogInsertDatabase	=	341		# Format, Style, LinkToSource, Connection, SQLStatement, SQLStatement1, PasswordDoc, PasswordDot, DataSource, From, To, IncludeFields, WritePasswordDoc, WritePasswordDot
wdDialogInsertDateTime	=	165		# DateTimePic, InsertAsField, DbCharField, DateLanguage, CalendarType
wdDialogInsertField	=	166		# Field
wdDialogInsertFile	=	164		# Name, Range, ConfirmConversions, Link, Attachment
wdDialogInsertFootnote	=	370		# Reference, NoteType, Symbol, FootNumberAs, EndNumberAs, FootnotesAt, EndnotesAt, FootNumberingStyle, EndNumberingStyle, FootStartingNum, FootRestartNum, EndStartingNum, EndRestartNum, ApplyPropsTo
wdDialogInsertFormField	=	483		# Entry, Exit, Name, Enable, TextType, TextWidth, TextDefault, TextFormat, CheckSize, CheckWidth, CheckDefault, Type, OwnHelp, HelpText, OwnStat, StatText, Calculate
wdDialogInsertHyperlink	=	925		# (none)
wdDialogInsertIndex	=	170		# Outline, Fields, From, To, TableId, AddedStyles, Caption, HeadingSeparator, Replace, MarkEntry, AutoMark, MarkCitation, Type, RightAlignPageNumbers, Passim, KeepFormatting, Columns, Category, Label, ShowPageNumbers, AccentedLetters, Filter, SortBy, Leader, TOCUseHyperlinks, TOCHidePageNumInWeb, IndexLanguage, UseOutlineLevel
wdDialogInsertIndexAndTables	=	473		# Outline, Fields, From, To, TableId, AddedStyles, Caption, HeadingSeparator, Replace, MarkEntry, AutoMark, MarkCitation, Type, RightAlignPageNumbers, Passim, KeepFormatting, Columns, Category, Label, ShowPageNumbers, AccentedLetters, Filter, SortBy, Leader, TOCUseHyperlinks, TOCHidePageNumInWeb, IndexLanguage, UseOutlineLevel
wdDialogInsertMergeField	=	167		# MergeField, WordField
wdDialogInsertNumber	=	812		# NumPic
wdDialogInsertObject	=	172		# IconNumber, FileName, Link, DisplayIcon, Tab, Class, IconFileName, Caption, Floating
wdDialogInsertPageNumbers	=	294		# Type, Position, FirstPage
wdDialogInsertPicture	=	163		# Name, LinkToFile, New, FloatOverText
wdDialogInsertPlaceholder	=	2348		# (none)
wdDialogInsertSource	=	2120		# (none)
wdDialogInsertSubdocument	=	583		# Name, ConfirmConversions, ReadOnly, LinkToSource, AddToMru, PasswordDoc, PasswordDot, Revert, WritePasswordDoc, WritePasswordDot, Connection, SQLStatement, SQLStatement1, Format, Encoding, Visible, OpenExclusive, OpenAndRepair, SubType, DocumentDirection, NoEncodingDialog, XMLTransform
wdDialogInsertSymbol	=	162		# Font, Tab, CharNum, CharNumLow, Unicode, Hint
wdDialogInsertTableOfAuthorities	=	471		# Outline, Fields, From, To, TableId, AddedStyles, Caption, HeadingSeparator, Replace, MarkEntry, AutoMark, MarkCitation, Type, RightAlignPageNumbers, Passim, KeepFormatting, Columns, Category, Label, ShowPageNumbers, AccentedLetters, Filter, SortBy, Leader, TOCUseHyperlinks, TOCHidePageNumInWeb, IndexLanguage, UseOutlineLevel
wdDialogInsertTableOfContents	=	171		# Outline, Fields, From, To, TableId, AddedStyles, Caption, HeadingSeparator, Replace, MarkEntry, AutoMark, MarkCitation, Type, RightAlignPageNumbers, Passim, KeepFormatting, Columns, Category, Label, ShowPageNumbers, AccentedLetters, Filter, SortBy, Leader, TOCUseHyperlinks, TOCHidePageNumInWeb, IndexLanguage, UseOutlineLevel
wdDialogInsertTableOfFigures	=	472		# Outline, Fields, From, To, TableId, AddedStyles, Caption, HeadingSeparator, Replace, MarkEntry, AutoMark, MarkCitation, Type, RightAlignPageNumbers, Passim, KeepFormatting, Columns, Category, Label, ShowPageNumbers, AccentedLetters, Filter, SortBy, Leader, TOCUseHyperlinks, TOCHidePageNumInWeb, IndexLanguage, UseOutlineLevel
wdDialogInsertWebComponent	=	1324		# IconNumber, FileName, Link, DisplayIcon, Tab, Class, IconFileName, Caption, Floating
wdDialogLabelOptions	=	1367		# (none)
wdDialogLetterWizard	=	821		# SenderCity, DateFormat, IncludeHeaderFooter, LetterStyle, Letterhead, LetterheadLocation, LetterheadSize, RecipientName, RecipientAddress, Salutation, SalutationType, RecipientGender, RecipientReference, MailingInstructions, AttentionLine, LetterSubject, CCList, SenderName, ReturnAddress, Closing, SenderJobTitle, SenderCompany, SenderInitials, EnclosureNumber, PageDesign, InfoBlock, SenderGender, ReturnAddressSF, RecipientCode, SenderCode, SenderReference
wdDialogListCommands	=	723		# ListType
wdDialogMailMerge	=	676		# CheckErrors, Destination, MergeRecords, From, To, Suppression, MailMerge, QueryOptions, MailSubject, MailAsAttachment, MailAddress
wdDialogMailMergeCheck	=	677		# CheckErrors
wdDialogMailMergeCreateDataSource	=	642		# FileName, PasswordDoc, PasswordDot, HeaderRecord, MSQuery, SQLStatement, SQLStatement1, Connection, LinkToSource, WritePasswordDoc
wdDialogMailMergeCreateHeaderSource	=	643		# FileName, PasswordDoc, PasswordDot, HeaderRecord, MSQuery, SQLStatement, SQLStatement1, Connection, LinkToSource, WritePasswordDoc
wdDialogMailMergeFieldMapping	=	1304		# (none)
wdDialogMailMergeFindRecipient	=	1326		# (none)
wdDialogMailMergeFindRecord	=	569		# (none)
wdDialogMailMergeHelper	=	680		# (none)
wdDialogMailMergeInsertAddressBlock	=	1305		# (none)
wdDialogMailMergeInsertAsk	=	4047		# (none)
wdDialogMailMergeInsertFields	=	1307		# (none)
wdDialogMailMergeInsertFillIn	=	4048		# (none)
wdDialogMailMergeInsertGreetingLine	=	1306		# (none)
wdDialogMailMergeInsertIf	=	4049		# (none)
wdDialogMailMergeInsertNextIf	=	4053		# (none)
wdDialogMailMergeInsertSet	=	4054		# (none)
wdDialogMailMergeInsertSkipIf	=	4055		# (none)
wdDialogMailMergeOpenDataSource	=	81		# (none)
wdDialogMailMergeOpenHeaderSource	=	82		# (none)
wdDialogMailMergeQueryOptions	=	681		# (none)
wdDialogMailMergeRecipients	=	1308		# (none)
wdDialogMailMergeSetDocumentType	=	1339		# (none)
wdDialogMailMergeUseAddressBook	=	779		# (none)
wdDialogMarkCitation	=	463		# (none)
wdDialogMarkIndexEntry	=	169		# (none)
wdDialogMarkTableOfContentsEntry	=	442		# (none)
wdDialogMyPermission	=	1437		# (none)
wdDialogNewToolbar	=	586		# (none)
wdDialogNoteOptions	=	373		# (none)
wdDialogOMathRecognizedFunctions	=	2165		# (none)
wdDialogOrganizer	=	222		# (none)
wdDialogPermission	=	1469		# (none)
wdDialogPhoneticGuide	=	986		# (none)
wdDialogReviewAfmtRevisions	=	570		# (none)
wdDialogSchemaLibrary	=	1417		# (none)
wdDialogSearch	=	1363		# (none)
wdDialogShowRepairs	=	1381		# (none)
wdDialogSourceManager	=	1920		# (none)
wdDialogStyleManagement	=	1948		# (none)
wdDialogTableAutoFormat	=	563		# (none)
wdDialogTableCellOptions	=	1081		# (none)
wdDialogTableColumnWidth	=	143		# (none)
wdDialogTableDeleteCells	=	133		# (none)
wdDialogTableFormatCell	=	612		# (none)
wdDialogTableFormula	=	348		# (none)
wdDialogTableInsertCells	=	130		# (none)
wdDialogTableInsertRow	=	131		# (none)
wdDialogTableInsertTable	=	129		# (none)
wdDialogTableOfCaptionsOptions	=	551		# (none)
wdDialogTableOfContentsOptions	=	470		# (none)
wdDialogTableProperties	=	861		# (none)
wdDialogTableRowHeight	=	142		# (none)
wdDialogTableSort	=	199		# (none)
wdDialogTableSplitCells	=	137		# (none)
wdDialogTableTableOptions	=	1080		# (none)
wdDialogTableToText	=	128		# (none)
wdDialogTableWrapping	=	854		# (none)
wdDialogTCSCTranslator	=	1156		# (none)
wdDialogTextToTable	=	127		# (none)
wdDialogToolsAcceptRejectChanges	=	506		# (none)
wdDialogToolsAdvancedSettings	=	206		# (none)
wdDialogToolsAutoCorrect	=	378		# (none)
wdDialogToolsAutoCorrectExceptions	=	762		# (none)
wdDialogToolsAutoManager	=	915		# (none)
wdDialogToolsAutoSummarize	=	874		# (none)
wdDialogToolsBulletsNumbers	=	196		# (none)
wdDialogToolsCompareDocuments	=	198		# (none)
wdDialogToolsCreateDirectory	=	833		# (none)
wdDialogToolsCreateEnvelope	=	173		# (none)
wdDialogToolsCreateLabels	=	489		# (none)
wdDialogToolsCustomize	=	152		# (none)
wdDialogToolsCustomizeKeyboard	=	432		# (none)
wdDialogToolsCustomizeMenuBar	=	615		# (none)
wdDialogToolsCustomizeMenus	=	433		# (none)
wdDialogToolsDictionary	=	989		# (none)
wdDialogToolsEnvelopesAndLabels	=	607		# (none)
wdDialogToolsGrammarSettings	=	885		# (none)
wdDialogToolsHangulHanjaConversion	=	784		# (none)
wdDialogToolsHighlightChanges	=	197		# (none)
wdDialogToolsHyphenation	=	195		# (none)
wdDialogToolsLanguage	=	188		# (none)
wdDialogToolsMacro	=	215		# (none)
wdDialogToolsMacroRecord	=	214		# (none)
wdDialogToolsManageFields	=	631		# (none)
wdDialogToolsMergeDocuments	=	435		# (none)
wdDialogToolsOptions	=	974		# (none)
wdDialogToolsOptionsAutoFormat	=	959		# (none)
wdDialogToolsOptionsAutoFormatAsYouType	=	778		# (none)
wdDialogToolsOptionsBidi	=	1029		# (none)
wdDialogToolsOptionsCompatibility	=	525		# (none)
wdDialogToolsOptionsEdit	=	224		# (none)
wdDialogToolsOptionsEditCopyPaste	=	1356		# (none)
wdDialogToolsOptionsFileLocations	=	225		# (none)
wdDialogToolsOptionsFuzzy	=	790		# (none)
wdDialogToolsOptionsGeneral	=	203		# (none)
wdDialogToolsOptionsPrint	=	208		# (none)
wdDialogToolsOptionsSave	=	209		# (none)
wdDialogToolsOptionsSecurity	=	1361		# (none)
wdDialogToolsOptionsSmartTag	=	1395		# (none)
wdDialogToolsOptionsSpellingAndGrammar	=	211		# (none)
wdDialogToolsOptionsTrackChanges	=	386		# (none)
wdDialogToolsOptionsTypography	=	739		# (none)
wdDialogToolsOptionsUserInfo	=	213		# (none)
wdDialogToolsOptionsView	=	204		# (none)
wdDialogToolsProtectDocument	=	503		# (none)
wdDialogToolsProtectSection	=	578		# (none)
wdDialogToolsRevisions	=	197		# (none)
wdDialogToolsSpellingAndGrammar	=	828		# (none)
wdDialogToolsTemplates	=	87		# (none)
wdDialogToolsThesaurus	=	194		# (none)
wdDialogToolsUnprotectDocument	=	521		# (none)
wdDialogToolsWordCount	=	228		# (none)
wdDialogTwoLinesInOne	=	1161		# (none)
wdDialogUpdateTOC	=	331		# (none)
wdDialogViewZoom	=	577		# (none)
wdDialogWebOptions	=	898		# (none)
wdDialogWindowActivate	=	220		# (none)
wdDialogXMLElementAttributes	=	1460		# (none)
wdDialogXMLOptions	=	1425		# (none)

#values from word.wdworddialogtab
#****************************
wdDialogEmailOptionsTabQuoting	=	1900002		# General tab of the Email Options dialog box.
wdDialogEmailOptionsTabSignature	=	1900000		# Email Signature tab of the Email Options dialog box.
wdDialogEmailOptionsTabStationary	=	1900001		# Personal Stationary tab of the Email Options dialog box.
wdDialogFilePageSetupTabCharsLines	=	150004		# Margins tab of the Page Setup dialog box, with Apply To drop-down list active.
wdDialogFilePageSetupTabLayout	=	150003		# Layout tab of the Page Setup dialog box.
wdDialogFilePageSetupTabMargins	=	150000		# Margins tab of the Page Setup dialog box.
wdDialogFilePageSetupTabPaper	=	150001		# Paper tab of the Page Setup dialog box.
wdDialogFormatBordersAndShadingTabBorders	=	700000		# Borders tab of the Borders dialog box.
wdDialogFormatBordersAndShadingTabPageBorder	=	700001		# Page Border tab of the Borders dialog box.
wdDialogFormatBordersAndShadingTabShading	=	700002		# Shading tab of the Borders dialog box.
wdDialogFormatBulletsAndNumberingTabBulleted	=	1500000		# Bulleted tab of the Bullets and Numbering dialog box.
wdDialogFormatBulletsAndNumberingTabNumbered	=	1500001		# Numbered tab of the Bullets and Numbering dialog box.
wdDialogFormatBulletsAndNumberingTabOutlineNumbered	=	1500002		# Outline Numbered tab of the Bullets and Numbering dialog box.
wdDialogFormatDrawingObjectTabColorsAndLines	=	1200000		# Colors and Lines tab of the Format Drawing Object dialog box.
wdDialogFormatDrawingObjectTabHR	=	1200007		# Colors and Lines tab of the Format Drawing Object dialog box.
wdDialogFormatDrawingObjectTabPicture	=	1200004		# Picture tab of the Format Drawing Object dialog box.
wdDialogFormatDrawingObjectTabPosition	=	1200002		# Position tab of the Format Drawing Object dialog box.
wdDialogFormatDrawingObjectTabSize	=	1200001		# Size tab of the Format Drawing Object dialog box.
wdDialogFormatDrawingObjectTabTextbox	=	1200005		# Textbox tab of the Format Drawing Object dialog box.
wdDialogFormatDrawingObjectTabWeb	=	1200006		# Web tab of the Format Drawing Object dialog box.
wdDialogFormatDrawingObjectTabWrapping	=	1200003		# Wrapping tab of the Format Drawing Object dialog box.
wdDialogFormatFontTabAnimation	=	600002		# Animation tab of the Font dialog box.
wdDialogFormatFontTabCharacterSpacing	=	600001		# Character Spacing tab of the Font dialog box.
wdDialogFormatFontTabFont	=	600000		# Font tab of the Font dialog box.
wdDialogFormatParagraphTabIndentsAndSpacing	=	1000000		# Indents and Spacing tab of the Paragraph dialog box.
wdDialogFormatParagraphTabTeisai	=	1000002		# Line and Page Breaks tab of the Paragraph dialog box, with choices appropriate for Asian text.
wdDialogFormatParagraphTabTextFlow	=	1000001		# Line and Page Breaks tab of the Paragraph dialog box.
wdDialogInsertIndexAndTablesTabIndex	=	400000		# Index tab of the Index and Tables dialog box.
wdDialogInsertIndexAndTablesTabTableOfAuthorities	=	400003		# Table of Authorities tab of the Index and Tables dialog box.
wdDialogInsertIndexAndTablesTabTableOfContents	=	400001		# Table of Contents tab of the Index and Tables dialog box.
wdDialogInsertIndexAndTablesTabTableOfFigures	=	400002		# Table of Figures tab of the Index and Tables dialog box.
wdDialogInsertSymbolTabSpecialCharacters	=	200001		# Special Characters tab of the Symbol dialog box.
wdDialogInsertSymbolTabSymbols	=	200000		# Symbols tab of the Symbol dialog box.
wdDialogLetterWizardTabLetterFormat	=	1600000		# Letter Format tab of the Letter Wizard dialog box.
wdDialogLetterWizardTabOtherElements	=	1600002		# Other Elements tab of the Letter Wizard dialog box.
wdDialogLetterWizardTabRecipientInfo	=	1600001		# Recipient Info tab of the Letter Wizard dialog box.
wdDialogLetterWizardTabSenderInfo	=	1600003		# Sender Info tab of the Letter Wizard dialog box.
wdDialogNoteOptionsTabAllEndnotes	=	300001		# All Endnotes tab of the Note Options dialog box.
wdDialogNoteOptionsTabAllFootnotes	=	300000		# All Footnotes tab of the Note Options dialog box.
wdDialogOrganizerTabAutoText	=	500001		# AutoText tab of the Organizer dialog box.
wdDialogOrganizerTabCommandBars	=	500002		# Command Bars tab of the Organizer dialog box.
wdDialogOrganizerTabMacros	=	500003		# Macros tab of the Organizer dialog box.
wdDialogOrganizerTabStyles	=	500000		# Styles tab of the Organizer dialog box.
wdDialogTablePropertiesTabCell	=	1800003		# Cell tab of the Table Properties dialog box.
wdDialogTablePropertiesTabColumn	=	1800002		# Column tab of the Table Properties dialog box.
wdDialogTablePropertiesTabRow	=	1800001		# Row tab of the Table Properties dialog box.
wdDialogTablePropertiesTabTable	=	1800000		# Table tab of the Table Properties dialog box.
wdDialogTemplates	=	2100000		# Templates tab of the Templates and Add-ins dialog box.
wdDialogTemplatesLinkedCSS	=	2100003		# Linked CSS tab of the Templates and Add-ins dialog box.
wdDialogTemplatesXMLExpansionPacks	=	2100002		# XML Expansion Packs tab of the Templates and Add-ins dialog box.
wdDialogTemplatesXMLSchema	=	2100001		# XML Schema tab of the Templates and Add-ins dialog box.
wdDialogToolsAutoCorrectExceptionsTabFirstLetter	=	1400000		# First Letter tab of the AutoCorrect Exceptions dialog box.
wdDialogToolsAutoCorrectExceptionsTabHangulAndAlphabet	=	1400002		# Hangul and Alphabet tab of the AutoCorrect Exceptions dialog box. Available only in multi-language versions.
wdDialogToolsAutoCorrectExceptionsTabIac	=	1400003		# Other Corrections tab of the AutoCorrect Exceptions dialog box.
wdDialogToolsAutoCorrectExceptionsTabInitialCaps	=	1400001		# Initial Caps tab of the AutoCorrect Exceptions dialog box.
wdDialogToolsAutoManagerTabAutoCorrect	=	1700000		# AutoCorrect tab of the AutoCorrect dialog box.
wdDialogToolsAutoManagerTabAutoFormat	=	1700003		# AutoFormat tab of the AutoCorrect dialog box.
wdDialogToolsAutoManagerTabAutoFormatAsYouType	=	1700001		# Format As You Type tab of the AutoCorrect dialog box.
wdDialogToolsAutoManagerTabAutoText	=	1700002		# AutoText tab of the AutoCorrect dialog box.
wdDialogToolsAutoManagerTabSmartTags	=	1700004		# Smart Tags tab of the AutoCorrect dialog box.
wdDialogToolsEnvelopesAndLabelsTabEnvelopes	=	800000		# Envelopes tab of the Envelopes and Labels dialog box.
wdDialogToolsEnvelopesAndLabelsTabLabels	=	800001		# Labels tab of the Envelopes and Labels dialog box.
wdDialogToolsOptionsTabAcetate	=	1266		# Not supported.
wdDialogToolsOptionsTabBidi	=	1029		# Complex Scripts tab of the Options dialog box.
wdDialogToolsOptionsTabCompatibility	=	525		# Compatibility tab of the Options dialog box.
wdDialogToolsOptionsTabEdit	=	224		# Edit tab of the Options dialog box.
wdDialogToolsOptionsTabFileLocations	=	225		# File Locations tab of the Options dialog box.
wdDialogToolsOptionsTabFuzzy	=	790		# Not supported.
wdDialogToolsOptionsTabGeneral	=	203		# General tab of the Options dialog box.
wdDialogToolsOptionsTabHangulHanjaConversion	=	786		# Hangul Hanja Conversion tab of the Options dialog box.
wdDialogToolsOptionsTabPrint	=	208		# Print tab of the Options dialog box.
wdDialogToolsOptionsTabProofread	=	211		# Spelling and Grammar tab of the Options dialog box.
wdDialogToolsOptionsTabSave	=	209		# Save tab of the Options dialog box.
wdDialogToolsOptionsTabSecurity	=	1361		# Security tab of the Options dialog box.
wdDialogToolsOptionsTabTrackChanges	=	386		# Track Changes tab of the Options dialog box.
wdDialogToolsOptionsTabTypography	=	739		# Asian Typography tab of the Options dialog box.
wdDialogToolsOptionsTabUserInfo	=	213		# User Information tab of the Options dialog box.
wdDialogToolsOptionsTabView	=	204		# View tab of the Options dialog box.
wdDialogWebOptionsBrowsers	=	2000000		# Browsers tab of the Web Options dialog box.
wdDialogWebOptionsEncoding	=	2000003		# Encoding tab of the Web Options dialog box.
wdDialogWebOptionsFiles	=	2000001		# Files tab of the Web Options dialog box.
wdDialogWebOptionsFonts	=	2000004		# Fonts tab of the Web Options dialog box.
wdDialogWebOptionsGeneral	=	2000000		# General tab of the Web Options dialog box.
wdDialogWebOptionsPictures	=	2000002		# Pictures tab of the Web Options dialog box.
wdDialogStyleManagementTabEdit	=	2200000		# Edit tab of the Style Management dialog box.
wdDialogStyleManagementTabRecommend	=	2200001		# Recommend tab of the Style Management dialog box.
wdDialogStyleManagementTabRestrict	=	2200002		# Restrict tab of the Style Management dialog box.

#values from word.wdwrapsidetype
#****************************
wdWrapBoth	=	0		# Both sides of the specified shape.
wdWrapLargest	=	3		# Side of the shape that is farthest from the page margin.
wdWrapLeft	=	1		# Left side of shape only.
wdWrapRight	=	2		# Right side of shape only.

#values from word.wdwraptype
#****************************
wdWrapInline	=	7		# Places shapes in line with text.
wdWrapNone	=	3		# Places shape in front of text. See also  wdWrapFront.
wdWrapSquare	=	0		# Wraps text around the shape. Line continuation is on the opposite side of the shape.
wdWrapThrough	=	2		# Wraps text around the shape.
wdWrapTight	=	1		# Wraps text close to the shape.
wdWrapTopBottom	=	4		# Places text above and below the shape.
wdWrapBehind	=	5		# Places shape behind text.
wdWrapFront	=	6		# Places shape in front of text.

#values from word.wdwraptypemerged
#****************************
wdWrapMergeBehind	=	3		# Behind text.
wdWrapMergeFront	=	4		# In front of text.
wdWrapMergeInline	=	0		# In line with text.
wdWrapMergeSquare	=	1		# Square.
wdWrapMergeThrough	=	5		# Through.
wdWrapMergeTight	=	2		# Tight.
wdWrapMergeTopBottom	=	6		# Top and bottom.

#values from word.xlaxiscrosses
#****************************
xlAxisCrossesAutomatic	=	-4105		# Word sets the axis crossing point.
xlAxisCrossesCustom	=	-4114		# The  CrossesAt property specifies the axis crossing point.
xlAxisCrossesMaximum	=	2		# The axis crosses at the maximum value.
xlAxisCrossesMinimum	=	4		# The axis crosses at the minimum value.

#values from word.xlaxisgroup
#****************************
xlPrimary	=	1		# The primary axis group.
xlSecondary	=	2		# The secondary axis group.

#values from word.xlaxistype
#****************************
xlCategory	=	1		# Axis displays categories.
xlSeriesAxis	=	3		# Axis displays data series.
xlValue	=	2		# Axis displays values.

#values from word.xlbackground
#****************************
xlBackgroundAutomatic	=	-4105		# Word controls the background.
xlBackgroundOpaque	=	3		# An opaque background.
xlBackgroundTransparent	=	2		# A transparent background.

#values from word.xlbarshape
#****************************
xlBox	=	0		# A box.
xlConeToMax	=	5		# A cone, truncated at the specified value.
xlConeToPoint	=	4		# A cone, coming to a point at the specified value.
xlCylinder	=	3		# A cylinder.
xlPyramidToMax	=	2		# A pyramid, truncated at the specified value.
xlPyramidToPoint	=	1		# A pyramid, coming to a point at the specified value.

#values from word.xlbinstype
#****************************
xlBinsTypeAutomatic	=	0		# Sets bins type automatically.
xlBinsTypeCategorical	=	1		# Sets bins type by category.
xlBinsTypeManual	=	2		# Sets bins type manually.
xlBinsTypeBinSize	=	3		# Sets bins type by size.
xlBinsTypeBinCount	=	4		# Sets bins type by count.

#values from word.xlborderweight
#****************************
xlHairline	=	1		# A hairline border (thinnest border).
xlMedium	=	-4138		# A medium border.
xlThick	=	4		# A thick border (widest border).
xlThin	=	2		# A thin border.

#values from word.xlcategorylabellevel
#****************************
xlCategoryLabelLevelAll	=	-1		# Use all category label levels within range on the chart. The default.
xlCategoryLabelLevelCustom	=	-2		# Indicates literal data in the category labels.
xlCategoryLabelLevelNone	=	-3		# Use no category labels in the chart. Defaults to automatic indexed labels.

#values from word.xlcategorytype
#****************************
xlAutomaticScale	=	-4105		# Word controls the axis type.
xlCategoryScale	=	2		# Axis groups data by an arbitrary set of categories.
xlTimeScale	=	3		# Axis groups data on a time scale.

#values from word.xlchartelementposition
#****************************
xlChartElementPositionAutomatic	=	-4105		# Automatically sets the position of the chart element.
xlChartElementPositionCustom	=	-4114		# Specifies a specific position for the chart element.

#values from word.xlchartgallery
#****************************
xlAnyGallery	=	23		# Either of the galleries.
xlBuiltIn	=	21		# The built-in gallery.
xlUserDefined	=	22		# The user-defined gallery.

#values from word.xlchartitem
#****************************
xlAxis	=	21		# Axis.
xlAxisTitle	=	17		# Axis title.
xlChartArea	=	2		# Chart area.
xlChartTitle	=	4		# Chart title.
xlCorners	=	6		# Corners.
xlDataLabel	=	0		# Data label.
xlDataTable	=	7		# Data table.
xlDisplayUnitLabel	=	30		# Display unit label.
xlDownBars	=	20		# Down bars.
xlDropLines	=	26		# Drop lines.
xlErrorBars	=	9		# Error bars.
xlFloor	=	23		# Floor.
xlHiLoLines	=	25		# HiLo lines.
xlLeaderLines	=	29		# Leader lines.
xlLegend	=	24		# Legend.
xlLegendEntry	=	12		# Legend entry.
xlLegendKey	=	13		# Legend key.
xlMajorGridlines	=	15		# Major gridlines.
xlMinorGridlines	=	16		# Minor gridlines.
xlNothing	=	28		# Nothing.
xlPivotChartDropZone	=	32		# PivotChart drop zone.
xlPivotChartFieldButton	=	31		# PivotChart field button.
xlPlotArea	=	19		# Plot area.
xlRadarAxisLabels	=	27		# Radar axis labels.
xlSeries	=	3		# Series.
xlSeriesLines	=	22		# Series lines.
xlShape	=	14		# Shape.
xlTrendline	=	8		# Trend line.
xlUpBars	=	18		# Up bars.
xlWalls	=	5		# Walls.
xlXErrorBars	=	10		# X error bars.
xlYErrorBars	=	11		# Y error bars.

#values from word.xlchartpictureplacement
#****************************
xlAllFaces	=	7		# Display on all faces.
xlEnd	=	2		# Display on end.
xlEndSides	=	3		# Display on end and sides.
xlFront	=	4		# Display on front.
xlFrontEnd	=	6		# Display on front and end.
xlFrontSides	=	5		# Display on front and sides.
xlSides	=	1		# Display on sides.

#values from word.xlchartpicturetype
#****************************
xlStack	=	2		# The picture is sized to repeat a maximum of 15 times in the longest stacked bar.
xlStackScale	=	3		# The picture is sized to a specified number of units and repeated the length of the bar.
xlStretch	=	1		# The picture is stretched the full length of the stacked bar.

#values from word.xlchartsplittype
#****************************
xlSplitByCustomSplit	=	4		# The second chart displays arbitrary slides.
xlSplitByPercentValue	=	3		# The second chart displays values less than a percentage of the total value. The percentage is specified by the  SplitValue property.
xlSplitByPosition	=	1		# The second chart displays the smallest values in the data series. The number of values to display is specified by the  SplitValue property.
xlSplitByValue	=	2		# The second chart displays values less than the value specified by the  SplitValue property.

#values from word.xlcolorindex
#****************************
xlColorIndexAutomatic	=	-4105		# Automatic color.
xlColorIndexNone	=	-4142		# No color.

#values from word.xlconstants
#****************************
xl3DBar	=	-4099		# Three-dimensional bar chart group or series.
xl3DSurface	=	-4103		# Three-dimensional surface chart group or series.
xlAbove	=	0		# The summary row is displayed above the specified range.
xlAutomatic	=	-4105		# Word applies automatic settings, such as a color or page number, to the specified object.
xlBar	=	2		# Two-dimensional bar chart group or series.
xlBelow	=	1		# The summary row is displayed below the specified range.
xlBoth	=	1		# Display positive and negative error bars in the specified chart group or series.
xlBottom	=	-4107		# Bottom.
xlCenter	=	-4108		# Center.
xlChecker	=	9		# Checker pattern.
xlCircle	=	8		# Circle.
xlColumn	=	3		# Columnar chart group or series.
xlCombination	=	-4111		# Combination.
xlCorner	=	2		# Corner.
xlCrissCross	=	16		# Criss-cross pattern.
xlCross	=	4		# Cross pattern.
xlCustom	=	-4114		# Word applies custom settings, such as a color or error amount, to the specified object.
xlDefaultAutoFormat	=	-1		# Word applies default or automatic formatting.
xlDiamond	=	2		# Diamond pattern.
xlDistributed	=	-4117		# Distributed.
xlFill	=	5		# Fill.
xlFixedValue	=	1		# Display error amounts as a fixed value.
xlGeneral	=	1		# General.
xlGray16	=	17		# 16% gray pattern.
xlGray25	=	-4124		# 25% gray pattern.
xlGray50	=	-4125		# 50% gray pattern.
xlGray75	=	-4126		# 75% gray pattern.
xlGray8	=	18		# 8% gray pattern.
xlGrid	=	15		# Grid pattern.
xlHigh	=	-4127		# High.
xlInside	=	2		# Inside.
xlJustify	=	-4130		# Justify.
xlLeft	=	-4131		# Left.
xlLightDown	=	13		# Light down line pattern.
xlLightHorizontal	=	11		# Light horizontal line pattern.
xlLightUp	=	14		# Light up line pattern.
xlLightVertical	=	12		# Light vertical line pattern.
xlLow	=	-4134		# Low.
xlMaximum	=	2		# Maximum.
xlMinimum	=	4		# Minimum.
xlMinusValues	=	3		# Minus values.
xlNextToAxis	=	4		# Next to axis.
xlNone	=	-4142		# Do not display error bars in the specified chart group or series.
xlOpaque	=	3		# Opaque fill.
xlOutside	=	3		# Outside.
xlPercent	=	2		# Display error amounts as a percentage.
xlPlus	=	9		# Display positive error bars in the specified chart group or series.
xlPlusValues	=	2		# Plus values.
xlRight	=	-4152		# Right.
xlScale	=	3		# Scale.
xlSemiGray75	=	10		# 75% semi-gray pattern.
xlShowLabel	=	4		# Show label.
xlShowLabelAndPercent	=	5		# Show label and percent.
xlShowPercent	=	3		# Show percent.
xlShowValue	=	2		# Show value.
xlSingle	=	2		# Single line.
xlSolid	=	1		# Solid pattern.
xlSquare	=	1		# Square.
xlStar	=	5		# Star.
xlStError	=	4		# Display error amounts as a standard error.
xlTop	=	-4160		# Top.
xlTransparent	=	2		# Transparent fill.
xlTriangle	=	3		# Triangle.

#values from word.xlcopypictureformat
#****************************
xlBitmap	=	2		# A bitmap (.bmp, .jpg, .gif).
xlPicture	=	-4147		# A drawn picture (.png, .wmf, .mix).

#values from word.xldatalabelposition
#****************************
xlLabelPositionAbove	=	0		# The data label is positioned above the data point.
xlLabelPositionBelow	=	1		# The data label is positioned below the data point.
xlLabelPositionBestFit	=	5		# Word sets the position of the data label.
xlLabelPositionCenter	=	-4108		# The data label is centered on the data point or is inside a bar or pie chart.
xlLabelPositionCustom	=	7		# The data label is in a custom position.
xlLabelPositionInsideBase	=	4		# The data label is positioned inside the data point at the bottom edge.
xlLabelPositionInsideEnd	=	3		# The data label is positioned inside the data point at the top edge.
xlLabelPositionLeft	=	-4131		# The data label is positioned to the left of the data point.
xlLabelPositionMixed	=	6		# Data labels are in multiple positions.
xlLabelPositionOutsideEnd	=	2		# The data label is positioned outside the data point at the top edge.
xlLabelPositionRight	=	-4152		# The data label is positioned to the right of the data point.

#values from word.xldatalabelseparator
#****************************
xlDataLabelSeparatorDefault	=	1		# Word selects the separator.

#values from word.xldatalabelstype
#****************************
xlDataLabelsShowBubbleSizes	=	6		# Show the size of the bubble in reference to the absolute value.
xlDataLabelsShowLabel	=	4		# The category for the point.
xlDataLabelsShowLabelAndPercent	=	5		# The percentage of the total, and the category for the point. Available only for pie charts and doughnut charts.
xlDataLabelsShowNone	=	-4142		# No data labels.
xlDataLabelsShowPercent	=	3		# The percentage of the total. Available only for pie charts and doughnut charts.
xlDataLabelsShowValue	=	2		# The default value for the point (assumed if this argument is not specified).

#values from word.xldisplayblanksas
#****************************
xlInterpolated	=	3		# Values are interpolated into the chart.
xlNotPlotted	=	1		# Blank cells are not plotted.
xlZero	=	2		# Blanks are plotted as zero.

#values from word.xldisplayunit
#****************************
xlHundredMillions	=	-8		# Hundreds of millions.
xlHundreds	=	-2		# Hundreds.
xlHundredThousands	=	-5		# Hundreds of thousands.
xlMillionMillions	=	-10		# Millions of millions.
xlMillions	=	-6		# Millions.
xlTenMillions	=	-7		# Tens of millions.
xlTenThousands	=	-4		# Tens of thousands.
xlThousandMillions	=	-9		# Thousands of millions.
xlThousands	=	-3		# Thousands.

#values from word.xlendstylecap
#****************************
xlCap	=	1		# Caps are applied.
xlNoCap	=	2		# No caps are applied.

#values from word.xlerrorbardirection
#****************************
xlChartX	=	-4168		# Bars run parallel to the y-axis for x-axis values.
xlChartY	=	1		# Bars run parallel to the x-axis for y-axis values.

#values from word.xlerrorbarinclude
#****************************
xlErrorBarIncludeBoth	=	1		# Both the positive and negative error range.
xlErrorBarIncludeMinusValues	=	3		# Only the negative error range.
xlErrorBarIncludeNone	=	-4142		# No error bar range.
xlErrorBarIncludePlusValues	=	2		# Only the positive error range.

#values from word.xlerrorbartype
#****************************
xlErrorBarTypeCustom	=	-4114		# The range is set by fixed values or cell values.
xlErrorBarTypeFixedValue	=	1		# Fixed-length error bars.
xlErrorBarTypePercent	=	2		# The percentage of the range to be covered by the error bars.
xlErrorBarTypeStDev	=	-4155		# Shows the range for a specified number of standard deviations.
xlErrorBarTypeStError	=	4		# Shows the standard error range.

#values from word.xlhalign
#****************************
xlHAlignCenter	=	-4108		# Center.
xlHAlignCenterAcrossSelection	=	7		# Center across selection.
xlHAlignDistributed	=	-4117		# Distribute.
xlHAlignFill	=	5		# Fill.
xlHAlignGeneral	=	1		# Align according to data type.
xlHAlignJustify	=	-4130		# Justify.
xlHAlignLeft	=	-4131		# Left.
xlHAlignRight	=	-4152		# Right.

#values from word.xllegendposition
#****************************
xlLegendPositionBottom	=	-4107		# Below the chart.
xlLegendPositionCorner	=	2		# In the upper-right corner of the chart border.
xlLegendPositionCustom	=	-4161		# A custom position.
xlLegendPositionLeft	=	-4131		# Left of the chart.
xlLegendPositionRight	=	-4152		# Right of the chart.
xlLegendPositionTop	=	-4160		# Above the chart.

#values from word.xllinestyle
#****************************
xlContinuous	=	1		# A continuous line.
xlDash	=	-4115		# A dashed line.
xlDashDot	=	4		# Alternating dashes and dots.
xlDashDotDot	=	5		# A dash followed by two dots.
xlDot	=	-4118		# A dotted line.
xlDouble	=	-4119		# A double line.
xlLineStyleNone	=	-4142		# No line.
xlSlantDashDot	=	13		# Slanted dashes.

#values from word.xlmarkerstyle
#****************************
xlMarkerStyleAutomatic	=	-4105		# Automatic markers.
xlMarkerStyleCircle	=	8		# Circular markers.
xlMarkerStyleDash	=	-4115		# Long bar markers.
xlMarkerStyleDiamond	=	2		# Diamond-shaped markers.
xlMarkerStyleDot	=	-4118		# Short bar markers.
xlMarkerStyleNone	=	-4142		# No markers.
xlMarkerStylePicture	=	-4147		# Picture markers.
xlMarkerStylePlus	=	9		# Square markers with a plus sign.
xlMarkerStyleSquare	=	1		# Square markers.
xlMarkerStyleStar	=	5		# Square markers with an asterisk.
xlMarkerStyleTriangle	=	3		# Triangular markers.
xlMarkerStyleX	=	-4168		# Square markers with an X.

#values from word.xlorientation
#****************************
xlDownward	=	-4170		# Text runs downward.
xlHorizontal	=	-4128		# Text runs horizontally.
xlUpward	=	-4171		# Text runs upward.
xlVertical	=	-4166		# Text runs downward and is centered in the cell.

#values from word.xlparentdatalabeloptions
#****************************
xlParentDataLabelOptionsNone	=	0		# No parent labels are shown.
xlParentDataLabelOptionsBanner	=	1		# The parent label layout is a banner above the category.
xlParentDataLabelOptionsOverlapping	=	2		# The parent label is laid out within the category.

#values from word.xlpattern
#****************************
xlPatternAutomatic	=	-4105		# Word controls the pattern.
xlPatternChecker	=	9		# A checkerboard.
xlPatternCrissCross	=	16		# Criss-crossed lines.
xlPatternDown	=	-4121		# Dark diagonal lines running from the upper-left to the lower-right.
xlPatternGray16	=	17		# 16% gray.
xlPatternGray25	=	-4124		# 25% gray.
xlPatternGray50	=	-4125		# 50% gray.
xlPatternGray75	=	-4126		# 75% gray.
xlPatternGray8	=	18		# 8% gray.
xlPatternGrid	=	15		# A grid.
xlPatternHorizontal	=	-4128		# Dark horizontal lines.
xlPatternLightDown	=	13		# Light diagonal lines running from the upper-left to the lower-right.
xlPatternLightHorizontal	=	11		# Light horizontal lines.
xlPatternLightUp	=	14		# Light diagonal lines running from the lower-left to the upper-right.
xlPatternLightVertical	=	12		# Light vertical bars.
xlPatternLinearGradient	=	4000		# A linear gradient.
xlPatternNone	=	-4142		# No pattern.
xlPatternRectangularGradient	=	4001		# A rectangular gradient.
xlPatternSemiGray75	=	10		# 75% dark moiré.
xlPatternSolid	=	1		# A solid color.
xlPatternUp	=	-4162		# Dark diagonal lines running from the lower-left to the upper-right.
xlPatternVertical	=	-4166		# Dark vertical bars.

#values from word.xlpictureappearance
#****************************
xlPrinter	=	2		# The picture is copied as it will look when it is printed.
xlScreen	=	1		# The picture is copied to resemble its display on the screen as closely as possible.

#values from word.xlpiesliceindex
#****************************
xlCenterPoint	=	5		# The center point of a pie slice.
xlInnerCenterPoint	=	8		# The innermost center point of a doughnut slice.
xlInnerClockwisePoint	=	7		# The innermost point of the most clockwise radius of a doughnut slice.
xlInnerCounterClockwisePoint	=	9		# The innermost point of the most counterclockwise radius of a doughnut slice.
xlMidClockwiseRadiusPoint	=	4		# The midpoint of the most clockwise radius of a slice.
xlMidCounterClockwiseRadiusPoint	=	6		# The midpoint of the most counterclockwise radius of a slice.
xlOuterCenterPoint	=	2		# The outer center point of the circumference of a slice.
xlOuterClockwisePoint	=	3		# The outermost clockwise point of the circumference of a slice.
xlOuterCounterClockwisePoint	=	1		# The outermost counterclockwise point of the circumference of a slice.
}
#End Enum

$wd=new-object PSCustomObject -Property $wd


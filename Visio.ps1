#constants for Visio based on https://docs.microsoft.com/en-us/office/vba/api/Visio(enumerations)
$vis=[Ordered]@{

#values from visio.visarcsweepflags
#****************************
visArcSweepFlagConcave	=	0		# Concave arc
visArcSweepFlagConvex	=	1		# Convex arc

#values from visio.visautoconnectdir
#****************************
visAutoConnectDirDown	=	2		# Place shape below.
visAutoConnectDirLeft	=	3		# Place shape to the left.
visAutoConnectDirNone	=	0		# Do not place shape.
visAutoConnectDirRight	=	4		# Place shape to the right.
visAutoConnectDirUp	=	1		# Place shape above.

#values from visio.visautolinkbehaviors
#****************************
visAutoLinkDontReplaceExistingLinks	=	16		# Do not replace existing links.
visAutoLinkGenericProgressBar	=	2		# Show generic progress bar instead of more detailed one.
visAutoLinkIncludeHiddenProps	=	64		# Include hidden properties.
visAutoLinkNoApplyDataGraphic	=	4		# Do not apply the default data graphic to linked shapes.
visAutolinkNullMatchesNoFormula	=	32		# Allow null database values to map to &quot;No Formula&quot; in the Microsoft Office Visio ShapeSheet spreadsheet.
visAutoLinkReplaceExistingLinks	=	8		# Replace existing links.
visAutoLinkSelectedShapesOnly	=	1		# Link selected shapes only, not sub-shapes of selected shapes.

#values from visio.visautolinkfieldtypes
#****************************
visAutoLinkCustPropsLabel	=	2		# Field is label for shape data (custom property); field name is label and is required.
visAutoLinkGeometryAngle	=	4		# Field is angle in shape geometry.
visAutoLinkGeometryHeight	=	6		# Field is height of shape.
visAutoLinkGeometryWidth	=	5		# Field is width of shape.
visAutoLinkMasterName	=	8		# Field is local name of the master for the shape.
visAutoLinkMasterNameU	=	16		# Field is universal name of the master for the shape.
visAutoLinkObjectData1	=	11		# Field is  Data1 property of Shape object.
visAutoLinkObjectData2	=	12		# Field is  Data2 property of Shape object.
visAutoLinkObjectData3	=	13		# Field is  Data3 property of Shape object.
visAutoLinkObjectID	=	7		# Field is ID of the shape .
visAutoLinkObjectName	=	9		# Field is local name of the shape.
visAutoLinkObjectNameU	=	17		# Field is universal name of the shape.
visAutoLinkObjectType	=	10		# Field is type of shape object.
visAutoLinkPropRowNameU	=	14		# Field is universal property-row name; field name is  Cell.RowNameU and is required.
visAutoLinkShapeText	=	1		# Field is shape text.
visAutoLinkUserRowName	=	3		# Field is user-defined cell local row name; field name is  Cell.RowName and is required.
visAutoLinkUserRowNameU	=	15		# Field is universal user-defined cell row name; field name is  Cell.RowNameU and is required.

#values from visio.visboundingboxargs
#****************************
visBBoxDrawingCoords	=	0x2000		# Return numbers in the drawing coordinate system of the page or master whose shapes are being considered. By default, the returned numbers are drawing units in the local coordinate system of the parent of the considered shapes.
visBBoxExtents	=	0x4		# Return a rectangle that is the smallest rectangle parallel to the local coordinate system of the shape's parent that encloses the paths stroked by the shape's geometry.This rectangle may be larger or smaller than the shape's upright width-height box. The extents box determined for a shape of type  visTypeForeignObject equals that shape's upright width-height box.
visBBoxIgnoreVisible	=	0x20		# Ignore visible geometry.
visBBoxIncludeDataGraphics	=	0x10000		# Include data-graphic callout shapes (and their sub-shapes) that are applied to the shape, or the shapes in a master, page, or selection. Off by default.
visBBoxIncludeGuides	=	0x1000		# Include extents for shapes of type  visTypeguide. By default, the extents of shapes of type visTypeGuide are ignored.If you request guide extents, only the positions of vertical guides and the positions of horizontal guides contribute to the rectangle that is returned. If any vertical guides are reported on, an infinite extent is returned. If any horizontal guides are reported on, an infinite extent is returned. If any rotated guides are reported on, infinite and extents are returned.
visBBoxIncludeHidden	=	0x10		# Include hidden geometry.
visBBoxNoNonPrint	=	0x4000		# Ignore the extents of shapes that are non-printing. A shape is non-printing if the value of its NonPrinting cell is non-zero or it belongs only to non-printing layers.
visBBoxUprightText	=	0x2		# Return a rectangle that is the smallest rectangle parallel to the local coordinate system of the shape's parent that encloses the shape's text.
visBBoxUprightWH	=	0x1		# Return a rectangle that is the smallest rectangle parallel to the local coordinate system of the shape's parent that encloses the shape's width-height box.If the shape is not rotated, its upright width-height box and its width-height box are the same. Paths in the shape's geometry need not and often do not lie entirely within the shape's width-height box.

#values from visio.visbuiltinstenciltypes
#****************************
visBuiltInStencilBackgrounds	=	0		# The hidden stencil that contains the shapes displayed in the  Backgrounds gallery (Design tab).
visBuiltInStencilBorders	=	1		# The hidden stencil that contains the shapes displayed in the  Borders and Titles gallery (Design tab).
visBuiltInStencilContainers	=	2		# The hidden stencil that contains the shapes displayed in the  Container gallery (Insert tab).
visBuiltInStencilCallouts	=	3		# The hidden stencil that contains the shapes displayed in the  Callout gallery (Insert tab).
visBuiltInStencilLegends	=	4		# The hidden stencil that contains the shapes displayed in the  Insert Legend gallery (Data tab).

#values from visio.viscellerror
#****************************
visErrorDivideByZero	=	39		# Division by zero.
visErrorName	=	61		# Invalid name in cell formula.
visErrorNotAvailable	=	74		# Unknown error.
visErrorNumber	=	68		# Invalid number in cell formula.
visErrorReference	=	55		# Invalid reference in cell formula.
visErrorSuccess	=	0		# No error.
visErrorValue	=	47		# Invalid value in cell formula.

#values from visio.viscellindices
#****************************
vis1DBeginX	=	0		# BeginX Cell (1D Endpoints Section)
vis1DBeginY	=	1		# BeginY Cell (1D Endpoints Section)
vis1DEndX	=	2		# EndX Cell (1D Endpoints Section)
vis1DEndY	=	3		# EndY Cell (1D Endpoints Section)
visActionAction	=	3		# Action Cell (Actions Section)
visActionBeginGroup	=	8		# BeginGroup Cell (Actions Section)
visActionButtonFace	=	15		# ButtonFace Cell (Actions Section)
visActionChecked	=	4		# Checked Cell (Actions Section)
visActionDisabled	=	5		# Disabled Cell (Actions Section)
visActionInvisible	=	7		# Invisible Cell (Actions Section)
visActionMenu	=	0		# Menu Cell (Actions Section)
visActionReadOnly	=	6		# ReadOnly Cell (Actions Section)
visActionSortKey	=	16		# SortKey Cell (Actions Section)
visActionTagName	=	14		# TagName Cell (Actions Section)
visAlignBottom	=	5		# AlignBottom Cell (Alignment Section)
visAlignCenter	=	1		# AlignCenter Cell (Alignment Section)
visAlignLeft	=	0		# AlignLeft Cell (Alignment Section)
visAlignMiddle	=	4		# AlignMiddle Cell (Alignment Section)
visAlignRight	=	2		# AlignRight Cell (Alignment Section)
visAlignTop	=	3		# AlignTop Cell (Alignment Section)
visAnnotationComment	=	5		# Comment Cell (Annotation Section)
visAnnotationDate	=	4		# Date Cell (Annotation Section)
visAnnotationLangID	=	6		# LangID Cell (Annotation Section)
visAnnotationMarkerIndex	=	3		# In previous versions of Visio, represented the MarkerIndex Cell (Annotation Section), which is now deprecated.
visAnnotationReviewerID	=	2		# ReviewerID Cell (Annotation Section)
visAnnotationX	=	0		# X Cell (Annotation Section)
visAnnotationY	=	1		# Y Cell (Annotation Section)
visAspectRatio	=	5		# D Cell (Geometry Section)
visBegTrigger	=	11		# BegTrigger Cell (Glue Info Section)
visBevelBottomHeight	=	5		# BevelBottomHeight Cell (Bevel Properties Section)
visBevelBottomType	=	3		# BevelBottomType Cell (Bevel Properties Section)
visBevelBottomWidth	=	4		# BevelBottomWidth Cell (Bevel Properties Section)
visBevelContourColor	=	8		# BevelContourColor Cell (Bevel Properties Section)
visBevelContourSize	=	9		# BevelContourSize Cell (Bevel Properties Section)
visBevelDepthColor	=	6		# BevelDepthColor Cell (Bevel Properties Section)
visBevelDepthSize	=	7		# BevelDepthSize Cell (Bevel Properties Section)
visBevelLightingAngle	=	12		# BevelLightingAngle Cell (Bevel Properties Section)
visBevelLightingType	=	11		# BevelLightingType Cell (Bevel Properties Section)
visBevelMaterialType	=	10		# BevelMaterialType Cell (Bevel Properties Section)
visBevelTopHeight	=	2		# BevelTopHeight Cell (Bevel Properties Section)
visBevelTopType	=	0		# BevelTopType Cell (Bevel Properties Section)
visBevelTopWidth	=	1		# BevelTopWidth Cell (Bevel Properties Section)
visBow	=	2		# A Cell (Geometry Section)
visBulletFontSize	=	11		# BulletSize Cell (Paragraph Section)
visBulletFont	=	9		# BulletFont Cell (Paragraph Section)
visBulletIndex	=	7		# Bullet Cell (Paragraph Section)
visBulletString	=	8		# BulletString Cell (Paragraph Section)
visCellFirst	=	0		# Cell logically at or before every other cell in a row.
visCellInval	=	255		# An index no cell will ever have.
visCellNone	=	255		# Unspecified cell.
visCharacterAsianFont	=	51		# AsianFont Cell (Character Section)
visCharacterCase	=	3		# Case Cell (Character Section)
visCharacterColorTrans	=	17		# Transparency Cell (Character Section)
visCharacterColor	=	1		# Color Cell (Character Section)
visCharacterComplexScriptFont	=	52		# ComplexScriptFont Cell (Character Section)
visCharacterComplexScriptSize	=	54		# ComplexScriptSize Cell (Character Section)
visCharacterDblUnderline	=	8		# DoubleULine Cell (Character Section)
visCharacterDoubleStrikethrough	=	13		# DoubleStrikethrough Cell (Character Section)
visCharacterFont	=	0		# Font Cell (Character Section)
visCharacterFontScale	=	5		# Scale Cell (Character Section)
visCharacterLangID	=	57		# LangID Cell (Character Section)
visCharacterLetterspace	=	16		# Spacing Cell (Character Section)
visCharacterOverline	=	9		# Overline Cell (Character Section)
visCharacterPos	=	4		# Pos Cell (Character Section)
visCharacterSize	=	7		# Size Cell (Character Section)
visCharacterStrikethru	=	10		# Strikethru Cell (Character Section)
visCharacterStyle	=	2		# Style Cell (Character Section)
visColorSchemeIndex	=	0		# ColorSchemeIndex Cell (Theme Properties Section)
visCompoundType	=	10		# Compound Type Cell (Line Format Section)
visConnectorSchemeIndex	=	2		# ConnectorSchemeIndex Cell (Theme Properties Section)
visCnnctAutoGen	=	5		# AutoGen Cell (Connection Points Section)
visCnnctA	=	2		# DirX / A Cell (Connection Points Section)
visCnnctB	=	3		# DirY / B Cell (Connection Points Section)
visCnnctC	=	4		# Type / C Cell (Connection Points Section)
visCnnctDirX	=	2		# DirX / A Cell (Connection Points Section)
visCnnctDirY	=	3		# DirY / B Cell (Connection Points Section)
visCnnctD	=	5		# D Cell (Connection Points Section)
visCnnctType	=	4		# Type / C Cell (Connection Points Section)
visCnnctX	=	0		# X Cell (Connection Points Section)
visCnnctY	=	1		# Y Cell (Connection Points Section)
visComment	=	16		# Comment Cell (Miscellaneous Section)
visCompNoFill	=	0		# NoFill Cell (Geometry Section)
visCompNoLine	=	1		# NoLine Cell (Geometry Section)
visCompNoQuickDrag	=	5		# NoQuickDrag Cell (Geometry Section)
visCompNoShow	=	2		# NoShow Cell (Geometry Section)
visCompNoSnap	=	3		# NoSnap Cell (Geometry Section)
visCompPath	=	4		# Path Cell (Geometry Section)
visControl1X	=	2		# A Cell (Geometry Section)
visControl1Y	=	3		# B Cell (Geometry Section)
visControl2X	=	4		# C Cell (Geometry Section)
visControl2Y	=	5		# D Cell (Geometry Section)
visControlX	=	2		# A Cell (Geometry Section)
visControlY	=	3		# B Cell (Geometry Section)
visCopyright	=	1		# Copyright Cell (Miscellaneous Section)
visCtlGlue	=	6		# Can Glue Cell (Controls Section)
visCtlTip	=	8		# Tip Cell (Controls Section)
visCtlXCon	=	4		# X Behavior Cell (Controls Section)
visCtlXDyn	=	2		# X Dynamics Cell (Controls Section)
visCtlX	=	0		# X Cell (Controls Section)
visCtlYCon	=	5		# Y Behavior Cell (Controls Section)
visCtlYDyn	=	3		# Y Dynamics Cell (Controls Section)
visCtlY	=	1		# Y Cell (Controls Section)
visCustPropsAsk	=	7		# Ask Cell (Shape Data Section)
visCustPropsCalendar	=	15		# Calendar Cell (Shape Data Section)
visCustPropsDataLinked	=	8		# DataLinked Cell (Shape Data Section)
visCustPropsFormat	=	3		# Format Cell (Shape Data Section)
visCustPropsInvis	=	6		# Invisible Cell (Shape Data Section)
visCustPropsLabel	=	2		# Label Cell (Shape Data Section)
visCustPropsLangID	=	14		# LangID Cell (Shape Data Section)
visCustPropsPrompt	=	1		# Prompt Cell (Shape Data Section)
visCustPropsSortKey	=	4		# SortKey Cell (Shape Data Section)
visCustPropsType	=	5		# Type Cell (Shape Data Section)
visCustPropsValue	=	0		# Value Cell (Shape Data Section)
visDocAddMarkup	=	3		# AddMarkup Cell (Document Properties Section)
visDocLangID	=	19		# DocLangID Cell (Document Properties Section)
visDocLockDuplicate	=	7		# LockDuplicate Cell (Document Properties Section)
visDocLockPreview	=	1		# LockPreview Cell (Document Properties Section)
visDocLockReplace	=	5		# LockReplace Cell (Document Properties Section)
visDocNoCoauth	=	6		# NoCoauth Cell (Document Properties Section)
visDocOutputFormat	=	0		# OutputFormat Cell (Document Properties Section)
visDocPreviewQuality	=	9		# PreviewQuality Cell (Document Properties Section)
visDocPreviewScope	=	10		# PreviewScope Cell (Document Properties Section)
visDocViewMarkup	=	4		# ViewMarkup Cell (Document Properties Section)
visDropSource	=	17		# IsDropSource Cell (Miscellaneous Section)
visDynFeedback	=	8		# DynFeedback Cell (Miscellaneous Section)
visEccentricityAngle	=	4		# C Cell (Geometry Section)
visEffectSchemeIndex	=	1		# EffectSchemeIndex Cell (Theme Properties Section)
visEllipseCenterX	=	0		# X Cell (Geometry Section)
visEllipseCenterY	=	1		# Y Cell (Geometry Section)
visEllipseMajorX	=	2		# A Cell (Geometry Section)
visEllipseMajorY	=	3		# B Cell (Geometry Section)
visEllipseMinorX	=	4		# C Cell (Geometry Section)
visEllipseMinorY	=	5		# D Cell (Geometry Section)
visEmbellishmentIndex	=	7		# EmbellishmentIndex Cell (Theme Properties Section)
visEndTrigger	=	12		# EndTrigger Cell (Glue Info Section)
visEvtCellDblClick	=	2		# EventDblClick Cell (Events Section)
visEvtCellDrop	=	4		# EventDrop Cell (Events Section)
visEvtCellMultiDrop	=	22		# EventMultiDrop Cell (Events Section)
visEvtCellTheText	=	1		# TheText Cell (Events Section)
visEvtCellXFMod	=	3		# EventXFMod Cell (Events Section)
visFieldCalendar	=	7		# Calendar Cell (Text Fields Section)
visFieldCell	=	0		# Value Cell (Text Fields Section)
visFieldFormat	=	2		# Format Cell (Text Fields Section)
visFieldObjectKind	=	10		# ObjectKind Cell (Text Fields Section)
visFieldType	=	3		# Type Cell (Text Fields Section)
visFieldUICategory	=	4		# UICategory Cell (Text Fields Section)
visFieldUICode	=	5		# UICode Cell (Text Fields Section)
visFieldUIFormat	=	6		# UIFormat Cell (Text Fields Section)
visFillBkgndTrans	=	7		# FillBkgndTrans Cell (Fill Format Section)
visFillBkgnd	=	1		# FillBkgnd Cell (Fill Format Section)
visFillForegndTrans	=	6		# FillForegndTrans Cell (Fill Format Section)
visFillForegnd	=	0		# FillForegnd Cell (Fill Format Section)
visFillGradientAngle	=	3		# FillGradientAngle Cell (Gradient Properties Section)
visFillGradientDir	=	2		# FillGradientDir Cell (Gradient Properties Section)
visFillGradientEnabled	=	5		# FillGradientEnabled Cell (Gradient Properties Section)
visFillPattern	=	2		# FillPattern Cell (Fill Format Section)
visFillShdwBkgndTrans	=	9		# In previous versions of Visio, represented the ShdwBkgndTrans Cell (Fill Format Section), which is now deprecated.
visFillShdwBkgnd	=	4		# In previous versions of Visio, represented the ShdwBkgnd Cell (Fill Format Section), which is now deprecated.
visFillShdwBlur	=	15		# ShapeShdwBlur Cell (Fill Format Section)
visFillShdwForegndTrans	=	8		# ShdwForegndTrans Cell (Fill Format Section)
visFillShdwForegnd	=	3		# ShdwForegnd Cell (Fill Format Section)
visFillShdwObliqueAngle	=	13		# ShapeShdwObliqueAngle Cell (Fill Format Section)
visFillShdwOffsetX	=	11		# ShapeShdwOffsetX Cell (Fill Format Section)
visFillShdwOffsetY	=	12		# ShapeShdwOffsetY Cell (Fill Format Section)
visFillShdwPattern	=	5		# ShdwPattern Cell (Fill Format Section)
visFillShdwScaleFactor	=	14		# ShapeShdwScaleFactor Cell (Fill Format Section)
visFillShdwShow	=	16		# ShapeShdwShow Cell (Fill Format Section)
visFillShdwType	=	10		# ShdwType Cell (Page Properties Section)
visFlags	=	13		# Flags Cell (Paragraph Section)
visFontSchemeIndex	=	3		# FontSchemeIndex Cell (Theme Properties Section)
visFrgnImgClippingPath	=	4		# ClippingPath Cell (Foreign Image Info Section)
visFrgnImgHeight	=	3		# ImgHeight Cell (Foreign Image Info Section)
visFrgnImgOffsetX	=	0		# ImgOffsetX Cell (Foreign Image Info Section)
visFrgnImgOffsetY	=	1		# ImgOffsetY Cell (Foreign Image Info Section)
visFrgnImgWidth	=	2		# ImgWidth Cell (Foreign Image Info Section)
visGlowColor	=	4		# GlowColor Cell (Additional Effect Properties Section)
visGlowColorTrans	=	5		# GlowColorTrans Cell (Additional Effect Properties Section)
visGlowSize	=	6		# GlowSize Cell (Additional Effect Properties Section)
visGlueType	=	9		# GlueType Cell (Glue Info Section)
visGradientStopColor	=	0		# Color Cell (Fill Gradient Stops Section)
visGradientStopColorTrans	=	1		# Trans Cell (Fill Gradient Stops Section)
visGradientStopPosition	=	2		# Position (Fill Gradient Stops Section)
visGroupDisplayMode	=	1		# DisplayMode Cell (Group Properties Section)
visGroupDontMoveChildren	=	5		# DontMoveChildren Cell (Group Properties Section)
visGroupIsDropTarget	=	2		# IsDropTarget Cell (Group Properties Section)
visGroupIsSnapTarget	=	3		# IsSnapTarget Cell (Group Properties Section)
visGroupIsTextEditTarget	=	4		# IsTextEditTarget Cell (Group Properties Section)
visGroupSelectMode	=	0		# SelectMode Cell (Group Properties Section)
visHideText	=	5		# HideText Cell (Miscellaneous Section)
visHLinkAddress	=	1		# Address Cell (Hyperlinks Section)
visHLinkDefault	=	7		# Default Cell (Hyperlinks Section)
visHLinkDescription	=	0		# Description Cell (Hyperlinks Section)
visHLinkExtraInfo	=	3		# ExtraInfo Cell (Hyperlinks Section)
visHLinkFrame	=	4		# Frame Cell (Hyperlinks Section)
visHLinkInvisible	=	8		# Invisible Cell (Hyperlinks Section)
visHLinkNewWin	=	5		# NewWindow Cell (Hyperlinks Section)
visHLinkSortKey	=	15		# SortKey Cell (Hyperlinks Section)
visHLinkSubAddress	=	2		# SubAddress Cell (Hyperlinks Section)
visHorzAlign	=	6		# HAlign Cell (Paragraph Section)
visImageBlur	=	4		# Blur Cell (Image Properties Section)
visImageBrightness	=	2		# Brightness Cell (Image Properties Section)
visImageContrast	=	1		# Contrast Cell (Image Properties Section)
visImageDenoise	=	5		# Denoise Cell (Image Properties Section)
visImageGamma	=	0		# Gamma Cell (Image Properties Section)
visImageSharpen	=	3		# Sharpen Cell (Image Properties Section)
visImageTransparency	=	6		# Transparency Cell (Image Properties Section)
visIndentFirst	=	0		# IndFirst Cell (Paragraph Section)
visIndentLeft	=	1		# IndLeft Cell (Paragraph Section)
visIndentRight	=	2		# IndRight Cell (Paragraph Section)
visInfiniteLineX1	=	0		# X Cell (Geometry Section)
visInfiniteLineX2	=	2		# A Cell (Geometry Section)
visInfiniteLineY1	=	1		# Y Cell (Geometry Section)
visInfiniteLineY2	=	3		# B Cell (Geometry Section)
visLayerActive	=	6		# Active Cell (Layers Section)
visLayerColorTrans	=	11		# Transparency Cell (Layers Section)
visLayerColor	=	2		# Color Cell (Layers Section)
visLayerGlue	=	9		# Glue Cell (Layers Section)
visLayerLock	=	7		# Lock Cell (Layers Section)
visLayerMember	=	0		# Layer Membership Cell (Layer Membership Section)
visLayerNameUniv	=	10		# Universal Name Cell (Layers Section)
visLayerName	=	0		# Name Cell (Layers Section)
visLayerPrint	=	5		# Print Cell (Layers Section)
visLayerSnap	=	8		# Snap Cell (Layers Section)
visLayerStatus	=	3		# Status Cell (Layers Section)
visLayerVisible	=	4		# Visible Cell (Layers Section)
visLineBeginArrowSize	=	8		# BeginArrowSize Cell (Line Format Section)
visLineBeginArrow	=	5		# BeginArrow Cell (Line Format Section)
visLineColorTrans	=	9		# LineColorTrans Cell (Line Format Section)
visLineColor	=	1		# LineColor Cell (Line Format Section)
visLineEndArrowSize	=	4		# EndArrowSize Cell (Line Format Section)
visLineEndArrow	=	6		# EndArrow Cell (Line Format Section)
visLineEndCap	=	7		# LineCap Cell (Line Format Section)
visLineGradientAngle	=	1		# LineGradientAngle Cell (Gradient Properties Section)
visLineGradientDir	=	0		# LineGradientDir Cell (Gradient Properties Section)
visLineGradientEnabled	=	4		# LineGradientEnabled Cell (Gradient Properties Section)
visLinePattern	=	2		# LinePattern Cell (Line Format Section)
visLineRounding	=	3		# Rounding Cell (Line Format Section)
visLineWeight	=	0		# LineWeight Cell (Line Format Section)
visLockAspect	=	4		# LockAspect Cell (Protection Section)
visLockBegin	=	6		# LockBegin Cell (Protection Section)
visLockCalcWH	=	14		# LockCalcWH Cell (Protection Section)
visLockCrop	=	9		# LockCrop Cell (Protection Section)
visLockCustProp	=	16		# LockCustProp Cell (Protection Section)
visLockDelete	=	5		# LockDelete Cell (Protection Section)
visLockEnd	=	7		# LockEnd Cell (Protection Section)
visLockFormat	=	12		# LockFormat Cell (Protection Section)
visLockFromGroupFormat	=	17		# LockFromGroupFormat Cell (Protection Section)
visLockGroup	=	13		# LockGroup Cell (Protection Section)
visLockHeight	=	1		# LockHeight Cell (Protection Section)
visLockMoveX	=	2		# LockMoveX Cell (Protection Section)
visLockMoveY	=	3		# LockMoveY Cell (Protection Section)
visLockReplace	=	23		# LockReplace Cell (Protection Cell)
visLockRotate	=	8		# LockRotate Cell (Protection Section)
visLockSelect	=	15		# LockSelect Cell (Protection Section)
visLockTextEdit	=	11		# LockTextEdit Cell (Protection Section)
visLockThemeColors	=	18		# LockThemeColors Cell (Protection Section)
visLockThemeConnectors	=	20		# LockThemeConnectors Cell (Protection Section)
visLockThemeEffects	=	19		# LockThemeEffects Cell (Protection Section)
visLockThemeFonts	=	21		# LockThemeFonts Cell (Protection Section)
visLockThemeIndex	=	22		# LockThemeIndex Cell (Protection Section)
visLockVariation	=	24		# LockVariation Cell (Protection Section)
visLockVtxEdit	=	10		# LockVtxEdit Cell (Protection Section)
visLockWidth	=	0		# LockWidth Cell (Protection Section)
visLOFlags	=	13		# ObjType Cell (Miscellaneous Section)
visNoAlignBox	=	3		# NoAlignBox Cell (Miscellaneous Section)
visNoCtlHandles	=	2		# NoCtlHandles Cell (Miscellaneous Section)
visNoLiveDynamics	=	18		# NoLiveDynamics Cell (Miscellaneous Section)
visNonPrinting	=	1		# NonPrinting Cell (Miscellaneous Section)
visNoObjHandles	=	0		# NoObjHandles Cell (Miscellaneous Section)
visNURBSData	=	6		# E Cell (Geometry Section)
visNURBSKnotPrev	=	4		# C Cell (Geometry Section)
visNURBSKnot	=	2		# A Cell (Geometry Section)
visNURBSWeightPrev	=	5		# D Cell (Geometry Section)
visNURBSWeight	=	3		# B Cell (Geometry Section)
visObjCalendar	=	25		# Calendar Cell (Miscellaneous Section)
visObjDropOnPageScale	=	28		# DropOnPageScale Cell (Miscellaneous Section)
visObjHelp	=	0		# HelpTopic Cell (Miscellaneous Section)
visObjKeywords	=	27		# ShapeKeywords cell (Miscellaneous Section)
visObjLangID	=	26		# LangID Cell (Miscellaneous Section)
visObjLocalizeMerge	=	19		# LocalizeMerge Cell (Miscellaneous Section)
visObjNoProofing	=	20		# NoProofing Cell (Miscellaneous Section)
visObjTheme	=	29		# Reserved for internal use only.
visObjThemeModern	=	30		# Reserved for internal use only.
visPageDrawingScale	=	5		# DrawingScale Cell (Page Properties Section)
visPageDrawScaleType	=	7		# DrawingScaleType Cell (Page Properties Section)
visPageDrawSizeType	=	6		# DrawingSizeType Cell (Page Properties Section)
visPageHeight	=	1		# PageHeight Cell (Page Properties Section)
visPageInhibitSnap	=	26		# InhibitSnap Cell (Page Properties Section)
visPageLockDuplicate	=	30		# PageLockDuplicate Cell (Page Properties Section)
visPageLockReplace	=	28		# PageLockReplace Cell (Page Properties Section)
visPageScale	=	4		# InhibitSnap Cell (Page Properties Section)
visPageShdwObliqueAngle	=	36		# ShdwObliqueAngle Cell (Page Properties Section)
visPageShdwOffsetX	=	2		# ShdwOffsetX Cell (Page Properties Section)
visPageShdwOffsetY	=	3		# ShdwOffsetY Cell (Page Properties Section)
visPageShdwScaleFactor	=	37		# ShdwScaleFactor Cell (Page Properties Section)
visPageShdwType	=	35		# ShdwType Cell (Page Properties Section)
visPageUIVisibility	=	34		# UIVisibility Cell (Page Properties Section)
visPageWidth	=	0		# PageWidth Cell (Page Properties Section)
visPageZOrderChanged	=	39		# Reserved for internal use only.
visPLOAvenueSizeX	=	20		# AvenueSizeX Cell (Page Layout Section)
visPLOAvenueSizeY	=	21		# AvenueSizeY Cell (Page Layout Section)
visPLOAvoidPageBreaks	=	4		# AvoidPageBreaks Cell (Page Layout Section)
visPLOBlockSizeX	=	18		# BlockSizeX Cell (Page Layout Section)
visPLOBlockSizeY	=	19		# BlockSizeY Cell (Page Layout Section)
visPLOCtrlAsInput	=	3		# CtrlAsInput Cell (Page Layout Section)
visPLODynamicsOff	=	2		# DynamicsOff Cell (Page Layout Section)
visPLOEnableGrid	=	1		# EnableGrid Cell (Page Layout Section)
visPLOJumpCode	=	12		# LineJumpCode Cell (Page Layout Section)
visPLOJumpDirX	=	14		# LineJumpFactorX Cell (Page Layout Section)
visPLOJumpDirY	=	15		# LineJumpFactorY Cell (Page Layout Section)
visPLOJumpFactorX	=	24		# LineJumpFactorX Cell (Page Layout Section)
visPLOJumpFactorY	=	25		# LineJumpFactorXYCell (Page Layout Section)
visPLOJumpStyle	=	13		# LineJumpStyle Cell (Page Layout Section)
visPLOLineAdjustFrom	=	26		# LineAdjustFrom Cell (Page Layout Section)
visPLOLineAdjustTo	=	27		# LineAdjustTo Cell (Page Layout Section)
visPLOLineRouteExt	=	29		# LineRouteExt Cell (Page Layout Section)
visPLOLineToLineX	=	22		# LineToLineX Cell (Page Layout Section)
visPLOLineToLineY	=	23		# LineToLineY Cell (Page Layout Section)
visPLOLineToNodeX	=	16		# LineToNodeX Cell (Page Layout Section)
visPLOLineToNodeY	=	17		# LineToNodeY Cell (Page Layout Section)
visPLOPlaceDepth	=	10		# PlaceDepth Cell (Page Layout Section)
visPLOPlaceFlip	=	28		# PlaceFlip Cell (Page Layout Section)
visPLOPlaceStyle	=	8		# PlaceStyle Cell (Page Layout Section)
visPLOPlowCode	=	11		# PlowCode Cell (Page Layout Section)
visPLOResizePage	=	0		# ResizePage Cell (Page Layout Section)
visPLORouteStyle	=	9		# RouteStyle Cell (Page Layout Section)
visPLOSplit	=	30		# PageShapeSplit Cell (Page Layout Section)
visPolylineData	=	2		# A Cell (Geometry Section)
visPrintPropertiesBottomMargin	=	3		# PageBottomMargin Cell (Print Properties Section)
visPrintPropertiesCenterX	=	8		# CenterX Cell (Print Properties Section)
visPrintPropertiesCenterY	=	9		# CenterY Cell (Print Properties Section)
visPrintPropertiesLeftMargin	=	0		# PageLeftMargin Cell (Print Properties Section)
visPrintPropertiesOnPage	=	10		# OnPage Cell (Print Properties Section)
visPrintPropertiesPageOrientation	=	16		# PrintPageOrientation Cell (Print Properties Section)
visPrintPropertiesPagesX	=	6		# PagesX Cell (Print Properties Section)
visPrintPropertiesPagesY	=	7		# PagesY Cell (Print Properties Section)
visPrintPropertiesPaperKind	=	17		# PaperKind Cell (Print Properties Section)
visPrintPropertiesPaperSource	=	18		# PaperSource Cell (Print Properties Section)
visPrintPropertiesPrintGrid	=	11		# PrintGrid Cell (Print Properties Section)
visPrintPropertiesRightMargin	=	1		# PageRightMargin Cell (Print Properties Section)
visPrintPropertiesScaleX	=	4		# ScaleX Cell (Print Properties Section)
visPrintPropertiesScaleY	=	5		# ScaleY Cell (Print Properties Section)
visPrintPropertiesTopMargin	=	2		# PageTopMargin Cell (Print Properties Section)
visQuickStyleEffectsMatrix	=	6		# QuickStyleEffectsMatrix Cell (Quick Style Section)
visQuickStyleFillColor	=	1		# QuickStyleFillColor Cell (Quick Style Section)
visQuickStyleFillMatrix	=	5		# QuickStyleFillMatrix Cell (Quick Style Section)
visQuickStyleFontColor	=	3		# QuickStyleFontColor Cell (Quick Style Section)
visQuickStyleFontMatrix	=	7		# QuickStyleFontMatrix Cell (Quick Style Section)
visQuickStyleLineColor	=	0		# QuickStyleLineColor Cell (Quick Style Section)
visQuickStyleLineMatrix	=	4		# QuickStyleLineMatrix Cell (Quick Style Section)
visQuickStyleShadowColor	=	2		# QuickStyleShadowColor Cell (Quick Style Section)
visQuickStyleType	=	8		# QuickStyleType Cell (Quick Style Section)
visQuickStyleVariation	=	9		# QuickStyleVariation Cell (Quick Style Section)
visReflectionBlur	=	3		# ReflectionBlur Cell (Additional Effects Properties Section)
visReflectionDist	=	2		# ReflectionDist Cell (Additional Effects Properties Section)
visReflectionSize	=	1		# ReflectionSize Cell (Additional Effects Properties Section)
visReflectionTrans	=	0		# ReflectionTrans Cell (Additional Effects Properties Section)
visReplaceCopyCells	=	3		# ReplaceCopyCells Cell (Change Shape Behavior Section)
visReplaceLockFormat	=	2		# ReplaceLockFormat Cell (Change Shape Behavior Section)
visReplaceLockShapeData	=	1		# ReplaceLockShapeData Cell (Change Shape Behavior Section)
visReplaceLockText	=	0		# ReplaceLockText Cell (Change Shape Behavior Section)
visReviewerColor	=	2		# Color Cell (Reviewer Section)
visReviewerCurrentIndex	=	4		# CurrentIndex Cell (Reviewer Section)
visReviewerInitials	=	1		# Initials Cell (Reviewer Section)
visReviewerName	=	0		# Name Cell (Reviewer Section)
visReviewerReviewerID	=	3		# ReviewerID Cell (Reviewer Section)
visRotateGradientWithShape	=	6		# RotateGradientWithShape Cell (Gradient Properties Section)
visRotationType	=	3		# RotationType Cell (3D Rotation Properties Section)
visRotationXAngle	=	0		# RotationXAngle Cell (3D Rotation Properties Section)
visRotationYAngle	=	1		# RotationYAngle Cell (3D Rotation Properties Section)
visRotationZAngle	=	2		# RotationZAngle Cell (3D Rotation Properties Section)
visKeepTextFlat	=	6		# KeepTextFlat Cell (3D Rotation Properties Section)
visDistanceFromGround	=	5		# DistanceFromGround Cell (3D Rotation Properties Section)
visPerspective	=	4		# Perspective Cell (3D Rotation Properties Section)
visScratchA	=	2		# A Cell (Scratch Section)
visScratchB	=	3		# B Cell (Scratch Section)
visScratchC	=	4		# C Cell (Scratch Section)
visScratchD	=	5		# D Cell (Scratch Section)
visScratchX	=	0		# X Cell (Scratch Section)
visScratchY	=	1		# Y Cell (Scratch Section)
visSketchAmount	=	10		# SketchAmount Cell (Additional Effect Properties Section)
visSketchEnabled	=	9		# SketchEnabled Cell (Additional Effect Properties Section)
visSketchFillChange	=	13		# SketchFillChange Cell (Additional Effect Properties Section)
visSketchLineChange	=	12		# SketchLineChange Cell (Additional Effect Properties Section)
visSketchLineWeight	=	11		# SketchLineWeight Cell (Additional Effect Properties Section)
visSketchSeed	=	8		# SketchSeed Cell (Additional Effect Properties Section)
visSLOCategoryChanged	=	24		# Reserved for internal use only.
visSLOConFixedCode	=	12		# ConFixedCode Cell (Shape Layout Section)
visSLOFixedCode	=	8		# ShapeFixedCode Cell (Shape Layout Section)
visSLOJumpCode	=	13		# ConLineJumpCode Cell (Shape Layout Section)
visSLOJumpDirX	=	16		# ConLineJumpDirX Cell (Shape Layout Section)
visSLOJumpDirY	=	17		# ConLineJumpDirY Cell (Shape Layout Section)
visSLOJumpStyle	=	14		# ConLineJumpStyle Cell (Shape Layout Section)
visSLOLineRouteExt	=	19		# ConLineRouteExt Cell (Shape Layout Section)
visSLOPermeablePlace	=	2		# ShapePermeablePlace Cell (Shape Layout Section)
visSLOPermX	=	0		# ShapePermeableX Cell (Shape Layout Section)
visSLOPermY	=	1		# ShapePermeableY Cell (Shape Layout Section)
visSLOPlaceFlip	=	18		# ShapePlaceFlip Cell (Shape Layout Section)
visSLOPlaceStyle	=	11		# ShapePlaceStyle Cell (Shape Layout Section)
visSLOPlowCode	=	9		# ShapePlowCode Cell (Shape Layout Section)
visSLORelationships	=	3		# Relationships Cell (Shape Layout Section)
visSLORelChanged	=	23		# Reserved for internal use only.
visSLORouteStyle	=	10		# ShapeRouteStyle Cell (Shape Layout Section)
visSLOSplittable	=	21		# ShapeSplittable Cell (Shape Layout Section)
visSLOSplit	=	20		# ShapeSplit Cell (Shape Layout Section)
visSLOZOrderChanged	=	25		# ZOrderChanged Cell (Shape Layout Section)
visSmartTagButtonFace	=	6		# ButtonFace Cell (Smart Tags Section)
visSmartTagDescription	=	15		# Description Cell (Smart Tags Section)
visSmartTagDisabled	=	7		# Disabled Cell (Smart Tags Section)
visSmartTagDisplayMode	=	5		# DisplayMode Cell (Smart Tags Section)
visSmartTagName	=	2		# TagName Cell (Smart Tags Section)
visSmartTagXJustify	=	3		# X Justify Cell (Smart Tags Section)
visSmartTagX	=	0		# X Cell (Smart Tags Section)
visSmartTagYJustify	=	4		# Y Justify Cell (Smart Tags Section)
visSmartTagY	=	1		# Y Cell (Smart Tags Section)
visSoftEdgeSize	=	7		# SoftEdgesSize Cell (Additional Effect Properties Section)
visSpaceAfter	=	5		# SpAfter Cell (Paragraph Section)
visSpaceBefore	=	4		# SpBefore Cell (Paragraph Section)
visSpaceLine	=	3		# SpLine Cell (Paragraph Section)
visSplineDegree	=	5		# D Cell (Geometry Section)
visSplineKnot2	=	3		# B Cell (Geometry Section)
visSplineKnot3	=	4		# C Cell (Geometry Section)
visSplineKnot	=	2		# A Cell (Geometry Section)
visStyleHidden	=	3		# HideForApply Cell (Style Properties Section)
visStyleIncludesFill	=	1		# EnableFillProps Cell (Style Properties Section)
visStyleIncludesLine	=	0		# EnableLineProps Cell (Style Properties Section)
visStyleIncludesText	=	2		# EnableTextProps Cell (Style Properties Section)
visTabAlign	=	2		# Alignment Cell (Tabs Section)
visTabPos	=	1		# Position Cell (Tabs Section)
visTabStopCount	=	0		# X1 Cell (Tabs Section)
visTextPosAfterBullet	=	12		# TextPosAfterBullet Cell (Paragraph Section)
visThemeIndex	=	4		# ThemeIndex Cell (Theme Properties Section)
visTxtBlkBkgndTrans	=	11		# TextBkgndTrans Cell (Text Block Format Section)
visTxtBlkBkgnd	=	5		# TextBkgnd Cell (Text Block Format Section)
visTxtBlkBottomMargin	=	3		# BottomMargin Cell (Text Block Format Section)
visTxtBlkDefaultTabStop	=	6		# DefaultTabstop Cell (Text Block Format Section)
visTxtBlkDirection	=	10		# TextDirection Cell (Text Block Format Section)
visTxtBlkLeftMargin	=	0		# LeftMargin Cell (Text Block Format Section)
visTxtBlkRightMargin	=	1		# RightMargin Cell (Text Block Format Section)
visTxtBlkTopMargin	=	2		# TopMargin Cell (Text Block Format Section)
visTxtBlkVerticalAlign	=	4		# VerticalAlign Cell (Text Block Format Section)
visUpdateAlignBox	=	4		# UpdateAlignBox Cell (Miscellaneous Section)
visUseGroupGradient	=	7		# UseGroupGradient Cell (Gradient Properties Section)
visUserPrompt	=	1		# Prompt Cell (User-Defined Cells Section)
visUserValue	=	0		# Value Cell (User-Defined Cells Section)
visVariationColorIndex	=	5		# VariationColorIndex Cell (Theme Properties Section)
visVariationStyleIndex	=	6		# VariationStyleIndex Cell (Theme Properties Section)
visWalkPref	=	10		# WalkPreference Cell (Glue Info Section)
visXFormAngle	=	6		# Angle Cell (Shape Transform Section)
visXFormFlipX	=	7		# FlipX Cell (Shape Transform Section)
visXFormFlipY	=	8		# FlipY Cell (Shape Transform Section)
visXFormHeight	=	3		# Height Cell (Shape Transform Section)
visXFormLocPinX	=	4		# LocPinX Cell (Shape Transform Section)
visXFormLocPinY	=	5		# LocPinY Cell (Shape Transform Section)
visXFormPinX	=	0		# PinX Cell (Shape Transform Section)
visXFormPinY	=	1		# PinY Cell (Shape Transform Section)
visXFormResizeMode	=	9		# ResizeMode Cell (Shape Transform Section)
visXFormWidth	=	2		# Width Cell (Shape Transform Section)
visXGridDensity	=	6		# XGridDensity Cell (Ruler &amp; Grid Section)
visXGridOrigin	=	10		# XGridOrigin Cell (Ruler &amp; Grid Section)
visXGridSpacing	=	8		# XGridSpacing Cell (Ruler &amp; Grid Section)
visXRulerDensity	=	0		# XRulerDensity Cell (Ruler &amp; Grid Section)
visXRulerOrigin	=	4		# XRulerOrigin Cell (Ruler &amp; Grid Section)
visX	=	0		# X Cell (Geometry Section)
visYGridDensity	=	7		# YGridDensity Cell (Ruler &amp; Grid Section)
visYGridOrigin	=	11		# YGridOrigin Cell (Ruler &amp; Grid Section)
visYGridSpacing	=	9		# YGridSpacing Cell (Ruler &amp; Grid Section)
visYRulerDensity	=	1		# YRulerDensity Cell (Ruler &amp; Grid Section)
visYRulerOrigin	=	5		# YRulerOrigin Cell (Ruler &amp; Grid Section)
visY	=	1		# Y Cell (Geometry Section)

#values from visio.viscellvals
#****************************
visArchitectural	=	1
visArrowSizeColossal	=	6
visArrowSizeJumbo	=	5
visArrowSizeLarge	=	3
visArrowSizeMedium	=	2
visArrowSizeSmall	=	1
visArrowSizeVeryLarge	=	4
visArrowSizeVerySmall	=	0
visBackDotsMini	=	8
visBackDotsWide	=	18
visBold	=	1
visCalArabicHijri	=	1
visCalChineseTaiwan	=	3
visCalHebrewLunar	=	2
visCalJapaneseEmperor	=	4
visCalKoreanDanki	=	6
visCalSakaEra	=	7
visCaltableHeaderaiBuddhist	=	5
visCalTranslitEnglish	=	8
visCalTranslitFrench	=	9
visCalWestern	=	0
visCaseAllCaps	=	1
visCaseInitialCaps	=	2
visCaseNormal	=	0
visCnnctTypeInwardOutward	=	2
visCnnctTypeInward	=	0
visCnnctTypeOutward	=	1
visComplexBold	=	16
visComplexItalic	=	32
visCtlLockedHidden	=	6
visCtlLocked	=	1
visCtlOffsetMaxHidden	=	9
visCtlOffsetMax	=	4
visCtlOffsetMidHidden	=	8
visCtlOffsetMid	=	3
visCtlOffsetMinHidden	=	7
visCtlOffsetMin	=	2
visCtlProportionalHidden	=	5
visCtlProportional	=	0
visCustom	=	3
visDocPreviewQualityDetailed	=	1
visDocPreviewQualityDraft	=	0
visDocPreviewScope1stPage	=	0
visDocPreviewScopeNone	=	1
visDSArch	=	7
visDSEngr	=	6
visDSMetric	=	5
visDynFBDefault	=	0
visDynFBUCon3Leg	=	1
visDynFBUCon5Leg	=	2
visEngineering	=	2
visForeDotsMini	=	10
visForeDotsNarrow	=	11
visForeDotsWide	=	12
visFSTOblique	=	2
visFSTPageDefault	=	0
visFSTSimple	=	1
visGlueTypeDefault	=	0
visGlueTypeNoWalkingTo	=	8
visGlueTypeNoWalking	=	4
visGlueTypeWalking	=	2
visGridCoarse	=	2
visGridFine	=	8
visGridFixed	=	0
visGridNormal	=	4
visGrpDispModeBack	=	1
visGrpDispModeFront	=	2
visGrpDispModeNone	=	0
visGrpSelModeGroup1st	=	1
visGrpSelModeGroupOnly	=	0
visGrpSelModeMembers1st	=	2
visHalfAndHalf	=	9
visHorzCenter	=	1
visHorzDistribute	=	4
visHorzJustifyHigh	=	8
visHorzJustifyLow	=	6
visHorzJustifyMedium	=	7
visHorzJustify	=	3
visHorzLeft	=	0
visHorzRight	=	2
visItalic	=	2
visLayerAvailable	=	2
visLayerDeleted	=	1
visLayerValid	=	0
visLocFontAlways	=	1
visLocFontIfArialOrSym	=	0
visLocFontNever	=	2
visLOFlagsDont	=	4
visLOFlagsPlacable	=	1
visLOFlagsPNRGroup	=	8
visLOFlagsRoutable	=	2
visLOFlagsVisDecides	=	0
visLOFlipDefault	=	0
visLOFlipNone	=	8
visLOFlipRotate	=	4
visLOFlipX	=	1
visLOFlipY	=	2
visLogical	=	4
visLOJumpDirXDefault	=	0
visLOJumpDirXDown	=	2
visLOJumpDirXUp	=	1
visLOJumpDirYDefault	=	0
visLOJumpDirYLeft	=	1
visLOJumpDirYRight	=	2
visLOJumpStyle2Point	=	5
visLOJumpStyle3Point	=	6
visLOJumpStyle4Point	=	7
visLOJumpStyle5Point	=	8
visLOJumpStyle6Point	=	9
visLOJumpStyleArc	=	1
visLOJumpStyleDefault	=	0
visLOJumpStyleGap	=	2
visLOJumpStyleSquare	=	3
visLOJumpStyleTriangle	=	4
visLOPlaceBottomToTop	=	4
visLOPlaceCircular	=	6
visLOPlaceCompactDownLeft	=	14
visLOPlaceCompactDownRight	=	7
visLOPlaceCompactLeftDown	=	13
visLOPlaceCompactLeftUp	=	12
visLOPlaceCompactRightDown	=	8
visLOPlaceCompactRightUp	=	9
visLOPlaceCompactUpLeft	=	11
visLOPlaceCompactUpRight	=	10
visLOPlaceDefault	=	0
visLOPlaceHierarchyBottomToTopCenter	=	20
visLOPlaceHierarchyBottomToTopLeft	=	19
visLOPlaceHierarchyBottomToTopRight	=	21
visLOPlaceHierarchyLeftToRightBottom	=	24
visLOPlaceHierarchyLeftToRightMiddle	=	23
visLOPlaceHierarchyLeftToRightTop	=	22
visLOPlaceHierarchyRightToLeftBottom	=	27
visLOPlaceHierarchyRightToLeftMiddle	=	26
visLOPlaceHierarchyRightToLeftTop	=	25
visLOPlaceHierarchyTopToBottomCenter	=	17
visLOPlaceHierarchyTopToBottomLeft	=	16
visLOPlaceHierarchyTopToBottomRight	=	18
visLOPlaceLeftToRight	=	2
visLOPlaceParentDefault	=	15
visLOPlaceRadial	=	3
visLOPlaceRightToLeft	=	5
visLOPlaceTopToBottom	=	1
visLORouteCenterToCenter	=	16
visLORouteDefault	=	0
visLORouteExtDefault	=	0
visLORouteExtNURBS	=	2
visLORouteExtStraight	=	1
visLORouteFlowchartEW	=	13
visLORouteFlowchartNS	=	5
visLORouteFlowchartSN	=	12
visLORouteFlowchartWE	=	6
visLORouteNetwork	=	9
visLORouteOrgChartEW	=	11
visLORouteOrgChartNS	=	3
visLORouteOrgChartSN	=	10
visLORouteOrgChartWE	=	4
visLORouteRightAngle	=	1
visLORouteSimpleEW	=	20
visLORouteSimpleHV	=	21
visLORouteSimpleNS	=	17
visLORouteSimpleSN	=	19
visLORouteSimpleVH	=	22
visLORouteSimpleWE	=	18
visLORouteStraight	=	2
visLORouteTreeEW	=	15
visLORouteTreeNS	=	7
visLORouteTreeSN	=	14
visLORouteTreeWE	=	8
visNoFill	=	0
visNoLayerColor	=	255
visNoScale	=	0
visPLOJumpDisplayOrder	=	4
visPLOJumpHorizontal	=	1
visPLOJumpLastRouted	=	3
visPLOJumpNone	=	0
visPLOJumpProhibitAll	=	6
visPLOJumpReverseDisplayOrder	=	5
visPLOJumpVertical	=	2
visPLOLineAdjustFromAll	=	1
visPLOLineAdjustFromNone	=	2
visPLOLineAdjustFromNotRelated	=	0
visPLOLineAdjustFromRoutingDefault	=	3
visPLOLineAdjustToAll	=	1
visPLOLineAdjustToDefault	=	0
visPLOLineAdjustToNone	=	2
visPLOLineAdjustToRelated	=	3
visPLOPlaceBottomToTop	=	4
visPLOPlaceCircular	=	6
visPLOPlaceCompactDownLeft	=	14
visPLOPlaceCompactDownRight	=	7
visPLOPlaceCompactLeftDown	=	13
visPLOPlaceCompactLeftUp	=	12
visPLOPlaceCompactRightDown	=	8
visPLOPlaceCompactRightUp	=	9
visPLOPlaceCompactUpLeft	=	11
visPLOPlaceCompactUpRight	=	10
visPLOPlaceDefault	=	0
visPLOPlaceDeptableHeaderDeep	=	2
visPLOPlaceDeptableHeaderDefault	=	0
visPLOPlaceDeptableHeaderMedium	=	1
visPLOPlaceDeptableHeaderShallow	=	3
visPLOPlaceLeftToRight	=	2
visPLOPlaceRadial	=	3
visPLOPlaceRightToLeft	=	5
visPLOPlaceTopToBottom	=	1
visPLOPlowAll	=	1
visPLOPlowNone	=	0
visPLOSplitAllow	=	1
visPLOSplitNone	=	0
visPosNormal	=	0
visPosSub	=	2
visPosSuper	=	1
visPPFlagsRTLText	=	1
visPPOLandscape	=	2
visPPOPortrait	=	1
visPPOSameAsPrinter	=	0
visPrintSetup	=	0
visPropTypeBool	=	3
visPropTypeCurrency	=	7
visPropTypeDate	=	5
visPropTypeDuration	=	6
visPropTypeListFix	=	1
visPropTypeListVar	=	4
visPropTypeNumber	=	2
visPropTypeString	=	0
visRulerCoarse	=	8
visRulerFine	=	32
visRulerFixed	=	0
visRulerNormal	=	16
visScaleCustom	=	3
visScaleMechanical	=	5
visScaleMetric	=	4
visSLOConFixedRerouteAsNeeded	=	1
visSLOConFixedRerouteFreely	=	0
visSLOConFixedRerouteNever	=	2
visSLOConFixedRerouteOnCrossover	=	3
visSLOFixedConnPtsIgnore	=	32
visSLOFixedConnPtsOnly	=	64
visSLOFixedNoFoldToShape	=	128
visSLOFixedPermeablePlow	=	4
visSLOFixedPlacement	=	1
visSLOFixedPlow	=	2
visSLOJumpAlways	=	2
visSLOJumpDefault	=	0
visSLOJumpNeitableHeaderer	=	4
visSLOJumpNever	=	1
visSLOJumpOtableHeaderer	=	3
visSLOPlowAlways	=	2
visSLOPlowDefault	=	0
visSLOPlowNever	=	1
visSLOSplitAllow	=	1
visSLOSplitNone	=	0
visSLOSplittableAllow	=	1
visSLOSplittableNone	=	0
visSmallCaps	=	8
visSmartTagDispModeAlways	=	2
visSmartTagDispModeMouseOver	=	0
visSmartTagDispModeShapeSelected	=	1
visSmartTagXJustifyCenter	=	1
visSmartTagXJustifyLeft	=	0
visSmartTagXJustifyRight	=	2
visSmartTagYJustifyBottom	=	2
visSmartTagYJustifyMiddle	=	1
visSmartTagYJustifyTop	=	0
visSolid	=	1
visStandard	=	2
visTabStopCenter	=	1
visTabStopComma	=	4
visTabStopDecimal	=	3
visTabStopLeft	=	0
visTabStopRight	=	2
visTFOKHorizontalInVertical	=	1
visTFOKStandard	=	0
vistableHeaderickDiagonalCross	=	17
vistableHeaderickDownDiagonal	=	15
vistableHeaderickHorz	=	13
vistableHeaderickUpDiagonal	=	16
vistableHeaderickVertical	=	14
vistableHeaderinCross	=	23
vistableHeaderinDiagonalCross	=	24
vistableHeaderinDownDiagonal	=	21
vistableHeaderinHorz	=	19
vistableHeaderinUpDiagonal	=	22
vistableHeaderinVert	=	20
visTight	=	1
visTxtBlkFlagsHideHiddens	=	1
visTxtBlkLeftToRight	=	0
visTxtBlkOpaque	=	255
visTxtBlkTopToBottom	=	1
visUIVHidden	=	1
visUIVNormal	=	0
visUnderLine	=	4
visVertBottom	=	2
visVertMiddle	=	1
visVertTop	=	0
visWalkPrefBegNS	=	1
visWalkPrefEndNS	=	2
visWideCross	=	3
visWideDiagonalCross	=	4
visWideDownDiagonal	=	5
visWideHorz	=	6
visWideUpDiagonal	=	2
visWideVert	=	7
visXFormResizeDontCare	=	0
visXFormResizeScale	=	2
visXFormResizeSpread	=	1

#values from visio.viscenterviewflags
#****************************
visCenterViewDefault	=	0		# Display the page that contains the specified shape, and center the view on the shape.
visCenterViewIfOffScreen	=	1		# Center the view only if the shape is currently off screen.
visCenterViewSelectShape	=	2		# Display the page that contains the specified shape, center the view on the shape, and select the shape.

#values from visio.vischarsbias
#****************************
visBiasLeft	=	1		# Specifies the ShapeSheet row that covers character formatting for the character to the left of the insertion point.
visBiasLetVisioChoose	=	0		# Specifies that Microsoft Visio decide the kind of character formatting to apply based on certain rules. See the Characters.CharPropsRow property topic for more information.
visBiasRight	=	2		# Specifies the ShapeSheet row that covers character formatting for the character to the right of the insertion point.

#values from visio.viscoloringmethod
#****************************
visColorDiscrete	=	0
visColorInvalid	=	2
visColorRange	=	1

#values from visio.visconnectedshapesflags
#****************************
visConnectedShapesAllNodes	=	0		# The shapes that are connected by using either incoming or outgoing connections.
visConnectedShapesIncomingNodes	=	1		# The shapes that are connected by using incoming connections.
visConnectedShapesOutgoingNodes	=	2		# The shapes that are connected by using outgoing connections.

#values from visio.visconnectorends
#****************************
visConnectorBeginpoint	=	0		# The begin point of the connector.
visConnectorEndPoint	=	1		# The end point of the connector.
visConnectorBothEnds	=	2		# Both the begin point and the end point of the connector.

#values from visio.viscontainerautoresize
#****************************
visContainerAutoResizeNone	=	0		# Do not automatically resize container.
visContainerAutoResizeExpand	=	1		# Automatically expand the container size, but do not contract.
visContainerAutoResizeExpandContract	=	2		# Automatically expand and contract the container size.

#values from visio.viscontainerflags
#****************************
visContainerFlagsDefault	=	0		# Returns all shape types and includes items in nested containers.
visContainerFlagsExcludeContainers	=	1		# Excludes member shapes that are containers.
visContainerFlagsExcludeConnectors	=	2		# Excludes member shapes that are connectors.
visContainerFlagsExcludeCallouts	=	4		# Excludes member shapes that are callouts.
visContainerFlagsExcludeElements	=	8		# Excludes member shapes that are not containers, connectors, or callouts.
visContainerFlagsExcludeNested	=	16		# Excludes any member shapes that are members of containers nested within the container.
visContainerFlagsExcludeListMembers	=	32		# Excludes members of a list container that are explicitly members of the list. Does not exclude other shapes in the list container.

#values from visio.viscontainerformattype
#****************************
visContainerFormatLockMembership	=	0		# Change whether container membership is locked.
visContainerFormatContainerAutoResize	=	1		# Change how the container resizes automatically.
visContainerFormatFitToContents	=	2		# Force the container to resize so as to tightly include all member shapes, including any applicable margins between the container and the shapes.

#values from visio.viscontainermemberstate
#****************************
visContainerMemberNotAMember	=	0		# The shape is not a member of the container.
visContainerMemberInterior	=	1		# The member shape is within the bounds of the container.
visContainerMemberOnBoundary	=	2		# The member shape is on the boundary of the container.
visContainerMemberOutside	=	3		# The member shape is outside the bounds of the container.
visContainerMemberInList	=	4		# The member shape is a list member.

#values from visio.viscontainernested
#****************************
visContainerIncludeNested	=	0		# Include shapes that are in nested containers.
visContainerExcludeNested	=	1		# Exclude shapes that are in nested containers.

#values from visio.viscontainertypes
#****************************
visContainerTypeNormal	=	0		# Member shapes are not arranged in a list.
visContainerTypeList	=	1		# Member shapes are arranged in a list.

#values from visio.viscutcopypastecodes
#****************************
visCopyPasteNormal	=	0x0		# Follow default copying behavior.
visCopyPasteNoTranslate	=	0x1		# Copy shapes to their original coordinate locations.
visCopyPasteCenter	=	0x2		# Copy shapes to the center of the page.
visCopyPasteNoHealConnectors	=	0x4		# Do not clean up connectors attached to cut shapes.
visCopyPasteNoContainerMembers	=	0x8		# Do not cut and copy unselected members of containers or lists.
visCopyPasteNoAssociatedCallouts	=	0x16		# Do not cut and copy unselected callouts associated with shapes.
visCopyPasteDontAddToContainers	=	0x32		# Do not add pasted shapes to any underlying containers.
visCopyPasteNoCascade	=	0x64		# Do not offset shapes on copy.

#values from visio.visdatacolumnproperties
#****************************
visDataColumnPropertyCalendar	=	3		# Calendar property of the data column.
visDataColumnPropertyCurrency	=	5		# Currency property of the data column.
visDataColumnPropertyDisplayName	=	6		# Display name of the data column in the UI.
visDataColumnPropertyHyperlink	=	8		# Whether the data-column value becomes a hyperlink in the Microsoft Office Visio user interface when it is linked to a shape.
visDataColumnPropertyLangID	=	2		# Language ID property of the data column.
visDataColumnPropertyType	=	1		# Type property of the data column .
visDataColumnPropertyUnits	=	4		# Units property of the data column.
visDataColumnPropertyVisible	=	7		# Whether the data-column property is visible in the UI, and therefore if the data column participates in data linking. For shapes linked to data rows in a data recordset, only visible columns populate shape data items in the shape.

#values from visio.visdatarecordsetaddoptions
#****************************
visDataRecordsetNoExternalDataUI	=	1		# Prevents data in the new data recordset from being displayed in the  External Data window.
visDataRecordsetNoRefreshUI	=	2		# Prevents the data recordset from being displayed in the  Refresh Data dialog box.
visDataRecordsetNoAdvConfig	=	4		# Limits the control users have of how the data recordset is refreshed in the  Configure Refresh dialog box for the data recordset. In particular, users cannot change the primary key or specify when shape data should be overwritten; however, users can set the refresh interval and can change the data source.
visDataRecordsetDelayQuery	=	8		# Adds a data recordset but does not execute the command-string query until the next time you call the  Refresh method.
visDataRecordsetDontCopyLinks	=	16		# Adds a data recordset, but shape-data links are not copied to the Clipboard when shapes are copied or cut.

#values from visio.visdefaultcolors
#****************************
visBlack	=	0		# Black
visBlue	=	4		# Blue
visCyan	=	7		# Cyan
visDarkBlue	=	10		# Dark blue
visDarkCyan	=	13		# Dark cyan
visDarkGray	=	19		# Dark gray
visDarkGreen	=	9		# Dark green
visDarkRed	=	8		# Dark red
visDarkYellow	=	11		# Dark yellow
visGray10	=	15		# 10% gray
visGray20	=	16		# 20% gray
visGray30	=	17		# 30% gray
visGray40	=	18		# 40% gray
visGray50	=	19		# 50% gray
visGray60	=	20		# 60% gray
visGray70	=	21		# 70% gray
visGray80	=	22		# 80% gray
visGray90	=	23		# 90% gray
visGray	=	14		# Gray
visGreen	=	3		# Green
visMagenta	=	6		# Magenta
visPurple	=	12		# Purple
visRed	=	2		# Red
visTransparent	=	0		# Transparent
visWhite	=	1		# White
visYellow	=	5		# Yellow

#values from visio.visdefaultsaveformats
#****************************
visDefaultSaveCurrent	=	0		# File format for current version of Microsoft Visio.
visDefaultSaveCurrentMacroEnabled	=	3		# Macro-enabled file format for current version of Visio.
visDefaultSavePreviousBinary	=	1		# Binary format for previous version of Visio (11.0).

#values from visio.visdeleteflags
#****************************
visDeleteNormal	=	0		# Match the deletion behavior in the user interface.
visDeleteHealConnectors	=	1		# Delete connectors attached to deleted shapes.
visDeleteNoHealConnectors	=	2		# Do not delete connectors attached to deleted shapes.
visDeleteNoContainerMembers	=	4		# Do not delete unselected members of containers or lists.
visDeleteNoAssociatedCallouts	=	8		# Do not delete unselected callouts associated with shapes.

#values from visio.visdiagramservices
#****************************
visServiceNone	=	0		# No diagram services.
visServiceAll	=	-1		# All diagram services.
visServiceAnimations	=	8		# Smooth transition behaviors to match user interface.
visServiceAutoSizePage	=	1		# AutoSize (automatic page-sizing) behaviors.
visServiceStructureBasic	=	2		# Structured-diagram behaviors that maintain existing relationships but do not create new relationships.
visServiceStructureFull	=	4		# Structured-diagram behaviors that match all those in the user interface (UI).
visServiceVersion140	=	7		# All diagram services that exist in Visio.
visServiceVersion150	=	8		# All diagram services that exist in Visio.

#values from visio.visdistributetypes
#****************************
visDistHorzCenter	=	2		# Distributes shapes hoirzontally so that their bottom edges are uniformly spaced.
visDistHorzLeft	=	1		# Distributes shapes horizontally so that their left edges are uniformly spaced.
visDistHorzRight	=	3		# Distributes shapes horizontally so that their right edges are uniformly spaced.
visDistHorzSpace	=	0		# Distributes shapes horizontally so that there is a uniform space between shapes.
visDistVertBottom	=	7		# Distributes shapes vertically so that their bottom edges are uniformly spaced.
visDistVertMiddle	=	6		# Distributes shapes vertically so that their centers are uniformly spaced.
visDistVertSpace	=	4		# Distributes shapes vertically so that there is a uniform space between shapes.
visDistVertTop	=	5		# Distributes shapes vertically so that their top edges are uniformly spaced.

#values from visio.visdoccleanactions
#****************************
visDocCleanActAll	=	0x3FFF		# Perform all actions.
visDocCleanActBadDisplayLists	=	0x100		# Detect invalid display list linkages.
visDocCleanActBadFieldFormulas	=	0x800		# Detect fields that have missing or nonstandard formulas.
visDocCleanActBadFieldMarks	=	0x1000		# Detect fields that have out-of-sync count and marker values. Change the position of escape characters to match character counts.
visDocCleanActBadReferences	=	0x2000		# Detect formulas that have #Ref() errors.
visDocCleanActConstantFormulas	=	0x20		# Detect formulas that can be generated from the result.
visDocCleanActDefault	=	0x1FD8		# Default conditions to detect.
visDocCleanActDeletedFields	=	0x400		# Detect deleted fields.
visDocCleanActDuplicateSubs	=	0x80		# Detect duplicate subscriptions (cell dependencies).
visDocCleanActEmptyRowsAndSects	=	0x2		# Detect empty local rows and sections.
visDocCleanActLocalFormulas	=	0x1		# Detect unnecessary local overrides.
visDocCleanActMissingSubs	=	0x10		# Detect missing subscriptions (cell dependencies).
visDocCleanActNearZero	=	0x40		# Detect results that are almost zero and change them to zero.
visDocCleanActNonDefaultFonts	=	0x4		# Detect non-default font settings.
visDocCleanActStaleResults	=	0x8		# Detect results that don't match formulas.
visDocCleanAlertDefault	=	0x0		# Default conditions to report.
visDocCleanFixDefault	=	0x3D8		# Default conditions to fix.

#values from visio.visdoccleantargets
#****************************
visDocCleanTargAll	=	0xFF		# Examine all objects.
visDocCleanTargBPages	=	0x2		# Examine background pages.
visDocCleanTargDoc	=	0x10		# Examine document sheet.
visDocCleanTargFPages	=	0x1		# Examine foreground pages.
visDocCleanTargMasters	=	0x4		# Examine masters.
visDocCleanTargPageSheet	=	0x100		# Examine page sheet(s).
visDocCleanTargStyles	=	0x8		# Examine styles.

#values from visio.visdocexintent
#****************************
visDocExIntentPrint	=	1		# Intent is to publish online and to print.
visDocExIntentScreen	=	0		# Intent is to publish online.

#values from visio.visdocmodeargs
#****************************
visDocModeDesign	=	1		# The document is in design mode.
visDocModeRun	=	0		# The document is in run mode.
visInvalDocID	=	-1		# The document ID is invalid. Document.ID will never return this value for an open document.

#values from visio.visdocumenttypes
#****************************
visDocTypeInval	=	0		# Not a Microsoft Visio document.
visTypeDrawing	=	1		# Document type is a drawing.
visTypeStencil	=	2		# Document type is a stencil.
visTypeTemplate	=	3		# Document type is a template.

#values from visio.visdocversions
#****************************
visVersion100	=	393216		# Visio version 2000 or 2002 document.
visVersion10	=	65571		# Visio 1.0 document.
visVersion110	=	720896		# Visio 2003, Visio 2007, or Visio 2010 document.
visVersion120	=	720896		# Visio 2003, Visio 2007, or Visio 2010 document.
visVersion140	=	720896		# Visio 2003, Visio 2007, or Visio 2010 document.
visVersion150	=	983040		# Visio document.
visVersion20	=	131072		# Visio 2.0 document.
visVersion30	=	196611		# Visio 3.0 document.
visVersion40	=	262144		# Visio 4.x document.
visVersion50	=	327680		# Visio 5.0 document.
visVersion60	=	393216		# Visio version 2000 or 2002 document.
visVersionUnsaved	=	0		# Visio version number of an unsaved document.

#values from visio.visdrawregionflags
#****************************
visDrawRegionDeleteInput	=	0x4		# Delete items in selection.
visDrawRegionIgnoreVisible	=	0x20		# Exclude visible geometry.
visDrawRegionIncludeDataGraphics	=	0x40		# Include data graphic callout shapes and their sub-shapes.
visDrawRegionIncludeHidden	=	0x10		# Include hidden geometry.

#values from visio.visdrawsplineflags
#****************************
visPolyarcs	=	256		# Draw a sequence of arcs rather than a sequence of line segments.
visPolyline1D	=	8		# Draw a shape that has one-dimensional (1D) behavior.
visSpline1D	=	8		# Draw a shape that has 1D behavior.
visSplineAbrupt	=	4		# Break the spline whenever an abrupt change of direction or curvature in the point's trail is detected.
visSplineDoCircles	=	2		# Recognize circular segments in the given array of points and generate circular arcs for those segments.
visSplinePeriodic	=	1		# Draw a periodic spline.

#values from visio.visedition
#****************************
visEditionStandard	=	0		# Standard edition.
visEditionProfessional	=	1		# Professional edition.

#values from visio.viseventcodes
#****************************
visActCodeAdvise	=	2		# AddAdvise action code
visActCodeRunAddon	=	1		# RunAddon action code
visEvtAdd	=	32768		# Event code for adding an  Event object, passed to the Add and AddAdvise methods. Used in conjunction with object codes for particular objects.
visEvtAfterModal	=	64		# AfterModal event
visEvtApp	=	4096		# Application object
visEvtAppActivate	=	1		# AppActivated event
visEvtAppDeactivate	=	2		# AppDeactivated event
visEvtBeforeModal	=	32		# BeforeModal event
visEvtBeforeQuit	=	16		# BeforeQuit event
visEvtCell	=	2048		# Cell object
visEvtCodeAfterCoauthMerge	=	14		# AfterCoauthMerge event code
visEvtCodeAfterForcedFlush	=	201		# AfterForcedFlush event code
visEvtCodeAfterResume	=	209		# AfterResume event code
visEvtCodeAfterResumeEvents	=	213		# AfterResumeEvents event code
visEvtCodeBefDocSaveAs	=	8		# BeforeDocumentSaveAs event code
visEvtCodeBefDocSave	=	7		# BeforeDocumentSave event code
visEvtCodeBefForcedFlush	=	200		# BeforeForcedFlush event code
visEvtCodeBeforeReplaceShapes	=	913		# BeforeReplaceShapes event code
visEvtCodeBeforeSuspend	=	208		# BeforeSuspend event code
visEvtCodeBeforeSuspendEvents	=	212		# BeforeSuspendEvents event code
visEvtCodeBefSelDel	=	901		# BeforeSelectionDelete event code
visEvtCodeBefWinPageTurn	=	703		# BeforeWindowPageTurn event code
visEvtCodeBefWinSelDel	=	702		# BeforeWindowSelDelete event code
visEvtCodeCalloutRelationshipAdded	=	504		# CalloutRelationshipAdded event code
visEvtCodeCalloutRelationshipDeleted	=	505		# CalloutRelationshipDeleted event code
visEvtCodeCancelConvertToGroup	=	908		# ConvertToGroupCanceled event code
visEvtCodeCancelDocClose	=	10		# DocumentCloseCanceled event code
visEvtCodeCancelMasterDel	=	401		# MasterDeleteCanceled event code
visEvtCodeCancelPageDel	=	501		# PageDeleteCanceled event code
visEvtCodeCancelQuit	=	205		# QuitCanceled event code
visEvtCodeCancelReplaceShapes	=	912		# ReplaceShapesCancelled event code
visEvtCodeCancelSelDel	=	904		# SelectionDeleteCanceled event code
visEvtCodeCancelSelGroup	=	910		# GroupCanceled event code
visEvtCodeCancelStyleDel	=	301		# StyleDeleteCanceled event code
visEvtCodeCancelSuspend	=	207		# SuspendCanceled event code
visEvtCodeCancelSuspendEvents	=	211		# SuspendEventsCanceled event code
visEvtCodeCancelUngroup	=	906		# UngroupCanceled event code
visEvtCodeCancelWinClose	=	707		# WindowCloseCanceled event code
visEvtCodeContainerRelationshipAdded	=	502		# ContainerRelationshipAdded event code
visEvtCodeContainerRelationshipDeleted	=	503		# ContainerRelationshipDeleted event code
visEvtCodeDocCreate	=	1		# DocumentCreated event code
visEvtCodeDocDesign	=	6		# DocumentDesignModeEntered event code
visEvtCodeDocOpen	=	2		# DocumentOpened event code
visEvtCodeDocRunning	=	5		# DocumentRunModeEntered event code
visEvtCodeDocSaveAs	=	4		# DocumentSavedAs event code
visEvtCodeDocSave	=	3		# DocumentSaved event code
visEvtCodeEnterScope	=	202		# EnterScope event code
visEvtCodeExitScope	=	203		# ExitScope event code
visEvtCodeInval	=	0		# An event code no event can have.
visEvtCodeKeyDown	=	712		# KeyDown event code
visEvtCodeKeyPress	=	713		# KeyPress event code
visEvtCodeKeyUp	=	714		# KeyUp event code
visEvtCodeMouseDown	=	709		# MouseDown event code
visEvtCodeMouseMove	=	710		# MousePress event code
visEvtCodeMouseUp	=	711		# MouseUp event code
visEvtCodeQueryCancelConvertToGroup	=	907		# QueryCancelConvertToGroup event code
visEvtCodeQueryCancelDocClose	=	9		# QueryCancelDocumentClose event code
visEvtCodeQueryCancelMasterDel	=	400		# QueryCancelMasterDelete event code
visEvtCodeQueryCancelPageDel	=	500		# QueryCancelPageDelete event code
visEvtCodeQueryCancelReplaceShapes	=	911		# QueryCancelReplaceShapes event code
visEvtCodeQueryCancelQuit	=	204		# QueryCancelQuit event code
visEvtCodeQueryCancelSelDel	=	903		# QueryCancelSelectionDeleted event code
visEvtCodeQueryCancelSelGroup	=	909		# QueryCancelGroup event code
visEvtCodeQueryCancelStyleDel	=	300		# QueryCancelStyleDeleted event code
visEvtCodeQueryCancelSuspend	=	206		# QueryCancelSuspend event code
visEvtCodeQueryCancelSuspendEvents	=	210		# QueryCancelSuspendEvents event code
visEvtCodeQueryCancelUngroup	=	905		# QueryCancelUngroup event code
visEvtCodeQueryCancelWinClose	=	706		# QueryCancelWindowClose event code
visEvtCodeRuleSetValidated	=	13		# RuleSetValidated event code
visEvtCodeSelAdded	=	902		# SelectionAdded event code
visEvtCodeSelectionMovedToSubprocess	=	12		# SelectionMovedToSubprocess event code
visEvtCodeShapeBeforeTextEdit	=	803		# BeforeShapeTextEdit event code
visEvtCodeShapeDelete	=	801		# ShapesDeleted event code
visEvtCodeShapeExitTextEdit	=	804		# ShapeExitedTextEdit event code
visEvtCodeShapeParentChange	=	802		# ShapeParentChanged event code
visEvtCodeShapesReplaced	=	914		# ShapesReplaced event code
visEvtCodeViewChanged	=	705		# ViewChanged event code
visEvtCodeWinOnAddonKeyMSG	=	708		# OnKeystrokeMessageForAddon event code
visEvtCodeWinPageTurn	=	704		# WindowTurnToPage event code
visEvtCodeWinSelChange	=	701		# SelectionChanged event code
visEvtConnect	=	256		# Connect object
visEvtDel	=	16384		# Event code for deleting an  Event object, passed to the Delete and AddAdvise methods. Used in conjunction with object codes for particular objects.
visEvtDataRecordset	=	32		# DataRecordset object
visEvtDoc	=	2		# Document object
visEvtFormula	=	4096		# FormulaChanged event
visEvtIDInval	=	-1		# An ID no event can have.
visEvtIdle	=	1024		# VisioIsIdle event code
visEvtIdMostRecent	=	0		# The ID of the most recent event to fire.
visEvtMarker	=	256		# MarkerEvent event
visEvtMaster	=	8		# Master object
visEvtMod	=	8192		# Used in conjunction with object codes for particular objects to create events that report a change to an object. For example,  visEvtMod + visEvtCell consitutes the CellChanged event.
visEvtNonePending	=	512		# NoEventsPending event
visEvtObjActivate	=	4		# AppObjActivated event
visEvtObjDeactivate	=	8		# AppObjDeactivated
visEvtPage	=	16		# Page object
visEvtRemoveHiddenInformation	=	11		# AfterRemoveHiddenInformation event
visEvtRow	=	1024		# Row object
visEvtSection	=	512		# Section object
visEvtShape	=	64		# Shape object
visEvtShapeDataGraphicChanged	=	807		# ShapeDataGraphicChanged event
visEvtShapeLinkAdded	=	805		# ShapeLinkAdded event
visEvtShapeLinkDeleted	=	806		# ShapeLinkDeleted event
visEvtStyle	=	4		# Style object
visEvtText	=	128		# Text object
visEvtWinActivate	=	128		# WindowActivated event
visEvtWindow	=	1		# Window object
visScopeIDInval	=	-1		# An ID no event scope can have.

#values from visio.visexistsflags
#****************************
visExistsAnywhere	=	0		# The ShapeSheet section either exists locally in the shape or is inherited.
visExistsLocally	=	1		# The ShapeSheet section exists locally in the shape.

#values from visio.visfieldcategories
#****************************
visFCatCustom	=	0		# Custom field.
visFCatDateTime	=	1		# Date/time field.
visFCatDocument	=	2		# Document information field.
visFCatGeometry	=	3		# Geometery field.
visFCatObject	=	4		# Object information.
visFCatPage	=	5		# Page information field.

#values from visio.visfieldcodes
#****************************
visFCodeAngle	=	2		# Angle.
visFCodeBackgroundName	=	0		# Background name.
visFCodeCategory	=	9		# Document category.
visFCodeCompany	=	8		# Document company.
visFCodeCreateDate	=	0		# Document creation date.
visFCodeCreateTime	=	1		# Document creation time.
visFCodeCreator	=	0		# Creator.
visFCodeCurrentDate	=	2		# Current date.
visFCodeCurrentTime	=	3		# Current time.
visFCodeData1	=	0		# Data1 value.
visFCodeData2	=	1		# Data2 value.
visFCodeData3	=	2		# Data3 value.
visFCodeDescription	=	1		# Document description.
visFCodeDirectory	=	2		# Directory (folder).
visFCodeEditDate	=	4		# Edit date.
visFCodeEditTime	=	5		# Edit time.
visFCodeFileName	=	3		# Document file name.
visFCodeHeight	=	1		# Height
visFCodeHyperlinkBase	=	10		# Hyperlink base.
visFCodeKeyWords	=	4		# Keywords.
visFCodeManager	=	7		# Manager name.
visFCodeMasterName	=	4		# Master name.
visFCodeNumberOfPages	=	2		# Number of pages.
visFCodeObjectID	=	3		# Object ID.
visFCodeObjectName	=	5		# Object name.
visFCodeObjectType	=	6		# Object type.
visFCodePageName	=	1		# Page name.
visFCodePageNumber	=	3		# Page number.
visFCodePrintDate	=	6		# Date printed.
visFCodePrintTime	=	7		# Time printed.
visFCodeSubject	=	5		# Document subject.
visFCodeTitle	=	6		# Document title.
visFCodeWidth	=	0		# Width.

#values from visio.visfieldformats
#****************************
visFmt0PlDefUnits	=	3
visFmt0PlNoUnits	=	2
visFmt1PlDefUnits	=	5
visFmt1PlNoUnits	=	4
visFmt2PlDefUnits	=	7
visFmt2PlNoUnits	=	6
visFmt3PlDefUnits	=	9
visFmt3PlNoUnits	=	8
visFmtDateDDMMYY	=	27
visFmtDateDMMMMYYYY	=	29
visFmtDateDMMMYYYY	=	28
visFmtDateDMYY	=	26
visFmtDategeMMMMddd_K	=	62
visFmtDategeMMMMddddww_K	=	60
visFmtDategggemd_J	=	56
visFmtDategggemdww_J	=	54
visFmtDateLong	=	21
visFmtDateMDYY	=	22
visFmtDateMMDDYY	=	23
visFmtDateMmmDYYYY	=	24
visFmtDateMmmmDYYYY	=	25
visFmtDateTWNfyyyymmdd_C	=	53
visFmtDateTWNfYYYYMMDDD_C	=	50
visFmtDateTWNfyyyymmddww_C	=	52
visFmtDateTWNsYYYYMMDDD_C	=	51
visFmtDateShort	=	20
visFmtDatewwyyyymd_S	=	79
visFmtDatewwyyyymmdd_S	=	78
visFmtDateyy_mm_dd	=	65
visFmtDateyymmdd	=	45
visFmtDateyyyy_m_d	=	64
visFmtDateyyyymd_J	=	57
visFmtDateyyyymd_K	=	63
visFmtDateyyyymd_S	=	76
visFmtDateyyyymdww_J	=	55
visFmtDateyyyymdww_K	=	61
visFmtDateyyyymd	=	44
visFmtDateyyyymmdd_S	=	77
visFmtDateYYYYMMMDDD_C	=	59
visFmtDateYYYYMMMDDDWWW_C	=	58
visFmtDegrees	=	12
visFmtFeetAndInches1Pl	=	13
visFmtFeetAndInches2Pl	=	14
visFmtFeetAndInches	=	10
visFmtFraction1PlDefUnits	=	16
visFmtFraction1PlNoUnits	=	15
visFmtFraction2PlDefUnits	=	18
visFmtFraction2PlNoUnits	=	17
visFmtMsoDateEnglish	=	208
visFmtMsoDateISO	=	204
visFmtMsoDateLongDay	=	201
visFmtMsoDateLong	=	202
visFmtMsoDateMon_Yr	=	210
visFmtMsoDateMonthYr	=	209
visFmtMsoDateShortAbb	=	207
visFmtMsoDateShortAlt	=	203
visFmtMsoDateShortMon	=	205
visFmtMsoDateShortSlash	=	206
visFmtMsoDateShort	=	200
visFmtMsoFEExtra1	=	217
visFmtMsoFEExtra2	=	218
visFmtMsoFEExtra3	=	219
visFmtMsoFEExtra4	=	220
visFmtMsoFEExtra5	=	221
visFmtMsoTime24	=	215
visFmtMsoTimeDatePM	=	211
visFmtMsoTimeDateSecPM	=	212
visFmtMsoTimePM	=	213
visFmtMsoTimeSec24	=	216
visFmtMsoTimeSecPM	=	214
visFmtNumGenDefUnits	=	1
visFmtNumGenNoUnits	=	0
visFmtRadians	=	11
visFmtStrLower	=	38
visFmtStrNormal	=	37
visFmtStrUpper	=	39
visFmtTimeAMPM_hmm_C	=	70
visFmtTimeAMPM_hmm_J	=	68
visFmtTimeAMPM_hmm_K	=	72
visFmtTimeAMPMhhmm_S	=	81
visFmtTimeAMPMhmm_C	=	66
visFmtTimeAMPMhmm_J	=	46
visFmtTimeAMPMhmm_K	=	67
visFmtTimeAMPMhmm_S	=	80
visFmtTimeGen	=	30
visFmtTimeHHMM24	=	34
visFmtTimeHHMMAMPM_E	=	75
visFmtTimeHHMMAMPM	=	36
visFmtTimeHHMM	=	32
visFmtTimehmm_C	=	71
visFmtTimehmm_J	=	69
visFmtTimehmm_K	=	73
visFmtTimeHMM24	=	33
visFmtTimeHMMAMPM_E	=	74
visFmtTimeHMMAMPM	=	35
visFmtTimeHMM	=	31

#values from visio.visfilteractions
#****************************
visFilterMouseMoveDragBegin	=	1		# Filter the  DragBegin extension of the MouseMove event.
visFilterMouseMoveDragDrop	=	5		# Filter the  DragDrop extension of the MouseMove event.
visFilterMouseMoveDragEnter	=	2		# Filter the  DragEnter extension of the MouseMove event.
visFilterMouseMoveDragLeave	=	4		# Filter the  DragLeave extension of the MouseMove event.
visFilterMouseMoveDragOver	=	3		# Filter the  DragOver extension of the MouseMove event.
visFilterMouseMoveNoDrag	=	0		# Do not filter any extensions of the  MouseMove event.

#values from visio.visfixedformattypes
#****************************
visFixedFormatPDF	=	1		# PDF format.
visFixedFormatXPS	=	2		# XPS format.

#values from visio.visflipdirection
#****************************
visFlipHorizontal	=	1		# Flip the selection horizontally.
visFlipVertical	=	2		# Flip the selection vertically.

#values from visio.visfliptypes
#****************************
visFlipSelectionWithPin	=	1		# Flip the selection around a pin.
visFlipSelection	=	0		# Flip the selection around its center.
visFlipShapes	=	2		# Flip the selected shapes around their pins.

#values from visio.visfontattributes
#****************************
visFont0Alias	=	128		# Used instead of font 0 (the default font). The font 0 alias is used in some localized versions of Microsoft Visio and is controlled by means of entries in the registry.

#values from visio.visfromparts
#****************************
visBeginX	=	7		# Connection is from the begin point x of a 1D shape.
visBeginY	=	8		# Connection is from the begin point y of a 1D shape.
visBegin	=	9		# Connection is from the begin point of a 1D shape.
visBottomEdge	=	4		# Connection is from bottom edge of shape.
visCenterEdge	=	2		# Connection is from the center (x) of a 1D shape.
visConnectFromError	=	-1		# Connection from an unknown part.
visControlPoint	=	100		# Connection is from the control point plus the row index (see Note).
visEndX	=	10		# Connection is from the endpoint (x) of a 1D shape.
visEndY	=	11		# Connection is from the endpoint (y) of a 1D shape.
visEnd	=	12		# Connection is from the end of a 1D shape.
visFromAngle	=	13		# Connection is from the direction of a connection point.
visFromNone	=	0		# Connection is from nothing.
visFromPin	=	14		# Connection is from the pin of a shape.
visLeftEdge	=	1		# Connection is from the left edge of a shape.
visMiddleEdge	=	5		# Connection is from the middle (y) of a shape.
visRightEdge	=	3		# Connection is from the right edge of a shape.
visTopEdge	=	6		# Connection is from the top edge of a shape.

#values from visio.visgeomflags
#****************************
visGeomExcludeLastPoint	=	1		# The last point (the X and Y cells in the row) is not included in the data.
visGeomWHPct	=	16		# The X and Y values are percentages of width and height.
visGeomXYLocal	=	32		# The X and Y values are local, internal units in the drawing.

#values from visio.visgetsetargs
#****************************
visGetFloats	=	0		# Results returned as doubles (VT_R8).
visGetFormulasU	=	5		# Formulas returned in universal syntax (VT_BSTR).
visGetFormulas	=	4		# Formulas returned as strings (VT_BSTR).
visGetRoundedInts	=	2		# Results returned as rounded long integers (VT_I4).
visGetStrings	=	3		# Results returned as strings (VT_BSTR).
visGetTruncatedInts	=	1		# Results returned as truncated long integers (VT_I4).
visSetBlastGuards	=	2		# Override present cell values even if they're guarded.
visSetFormulas	=	1		# Treat strings in results as formulas.
visSetTestCircular	=	4		# Test for establishment of circular cell references.
visSetUniversalSyntax	=	8		# Formulas are in universal syntax.

#values from visio.visgluedshapesflags
#****************************
visGluedShapesAll1D	=	0		# Return all 1D shapes that are glued to this shape.
visGluedShapesIncoming1D	=	1		# Return the 1D shapes whose end points are glued to this shape.
visGluedShapesOutgoing1D	=	2		# Return the 1D shapes whose begin points are glued to this shape.
visGluedShapesAll2D	=	3		# Return the 2D shapes that are glued to this shape and the 2D shapes to which this shape is glued.
visGluedShapesIncoming2D	=	4		# If the source object is a 1D shape, return the 2D shape to which the begin point is glued. If the source object is a 2D shape, return the 2D shapes that are glued to this shape.
visGluedShapesOutgoing2D	=	5		# If the source object is a 1D shape, return the 2D shape to which the end point is glued. If the source object is a 2D shape, return the 2D shapes to which this shape is glued.

#values from visio.visgluesettings
#****************************
visGlueToConnectionPoints	=	0x8		# Glue to connection points.
visGlueToDisabled	=	0x8000		# Disable glue.
visGlueToGeometry	=	0x20		# Glue to shape geometry.
visGlueToGuides	=	0x1		# Glue to guides.
visGlueToHandles	=	0x2		# Glue to shape handles.
visGlueToNone	=	0x0		# Glue is enabled but no other glue settings are on.
visGlueToVertices	=	0x4		# Glue to shape vertices.

#values from visio.visgraphicfield
#****************************
visGraphicExpression	=	2		# The ShapeSheet formula of a shape data item.
visGraphicPropertyLabel	=	1		# The label of a shape data item.

#values from visio.visgraphicitemtypes
#****************************
visTypeIconSet	=	2		# Represents an  Icon Set graphic item.
visTypeTextCallout	=	3		# Represents a  Text graphic item.
visTypeDataBar	=	4		# Represents a  Data Bar graphic item.
visTypeColorByValue	=	5		# Represents a  Color by Value graphic item.
visTypeHeading	=	6		# Represents a  Text graphic item that has a Callout type of Heading x .

#values from visio.visgraphicpositionhorizontal
#****************************
visGraphicFarLeft	=	0		# The right edge of the graphic item's alignment box is aligned with the left edge of the shape or container's alignment box.
visGraphicLeftEdge	=	1		# The vertical centerline of the graphic item's alignment box is aligned with the left edge of the shape or container's alignment box.
visGraphicLeft	=	2		# The left edge of the graphic item's alignment box is aligned with the left edge of the shape or container's alignment box.
visGraphicCenter	=	3		# The vertical centerline of the graphic item's alignment box is aligned with the vertical centerline of the shape or container's alignment box.
visGraphicRight	=	4		# The right edge of the graphic item's alignment box is aligned with the right edge of the shape or container's alignment box.
visGraphicRightEdge	=	5		# The vertical centerline of the graphic item's alignment box is aligned with the right edge of the shape or container's alignment box.
visGraphicFarRight	=	6		# The left edge of the graphic item's alignment box is aligned with the right edge of the shape or container's alignment box.

#values from visio.visgraphicpositionvertical
#****************************
visGraphicBelow	=	0		# The top edge of the graphic item's alignment box is aligned with the bottom edge of the shape or container's alignment box.
visGraphicBottomEdge	=	1		# The horizontal centerline of the graphic item's alignment box is aligned with the bottom edge of the shape or container's alignment box.
visGraphicBottom	=	2		# The bottom edge of the graphic item's alignment box is aligned with the bottom edge of the shape or container's alignment box.
visGraphicMiddle	=	3		# The horizontal centerline of the graphic item's alignment box is aligned with the horizontal centerline of the shape or container's alignment box.
visGraphicTop	=	4		# The top edge of the graphic item's alignment box is aligned with the top edge of the shape or container's alignment box.
visGraphicTopEdge	=	5		# The horizontal centerline of the graphic item's alignment box is aligned with the top edge of the shape or container's alignment box.
visGraphicAbove	=	6		# The bottom edge of the graphic item's alignment box is aligned with the top edge of the shape or container's alignment box.

#values from visio.visguidetypes
#****************************
visHorz	=	2		# Vertical guide
visPoint	=	1		# Guide point
visVert	=	3		# Horizontal guide

#values from visio.vishittestresults
#****************************
visHitInside	=	2		# Hit position is inside shape boundary.
visHitOnBoundary	=	1		# Hit position is on shape boundary
visHitOutside	=	0		# Hit position is outside shape boundary

#values from visio.vishorizontalaligntypes
#****************************
visHorzAlignCenter	=	2		# Align to the center of the primary selected shape.
visHorzAlignLeft	=	1		# Align to the left of the primary selected shape.
visHorzAlignNone	=	0		# No horizontal alignment.
visHorzAlignRight	=	3		# Align to the right of the primary selected shape.

#values from visio.visinsertobjargs
#****************************
visInsertAsControl	=	8192		# None.
visInsertAsEmbed	=	16384		# None.
visInsertDontShow	=	4096		# Don't execute the new object's show verb.
visInsertIcon	=	16		# Display the new object as an icon
visInsertLink	=	8		# If set, the new shape represents an OLE link to the named file. Otherwise, the InsertFromFile method produces an OLE object from the contents of the named file and embeds it in the document that contains the page, master, or group.
visInsertNoDesignModeTransition	=	256		# If set, when an ActiveX control is inserted, prevents Microsoft Visio from transitioning to design mode.

#values from visio.viskeybuttonflags
#****************************
visKeyControl	=	8		# CTRL key.
visKeyShift	=	4		# SHIFT key.
visMouseLeft	=	1		# Left mouse button.
visMouseMiddle	=	16		# Mouse wheel.
visMouseRight	=	2		# Right mouse button.

#values from visio.vislangflags
#****************************
visLangLocal	=	0		# The page name is a local name.
visLangUniversal	=	1		# The page name is a universal name.

#values from visio.vislayoutdirection
#****************************
visLayoutDirRotateRight	=	0		# Rotates the diagram or selection 90 degrees clockwise.
visLayoutDirRotateLeft	=	1		# Rotates the diagram or selection 90 degrees counter-clockwise.
visLayoutDirFlipVert	=	2		# Flips the diagram or selection vertically.
visLayoutDirFlipHorz	=	3		# Flips the diagram or selection horizontally.

#values from visio.vislayouthorzaligntype
#****************************
visLayoutHorzAlignNone	=	0		# Do not align horizontally.
visLayoutHorzAlignDefault	=	1		# Let Visio choose how to align horizontally.
visLayoutHorzAlignLeft	=	2		# Align the left edges of the shapes.
visLayoutHorzAlignCenter	=	3		# Align the centers of the shapes.
visLayoutHorzAlignRight	=	4		# Align the right edges of the shapes.

#values from visio.vislayoutincrementaltype
#****************************
visLayoutIncrAlign	=	1		# Align shapes.
visLayoutIncrSpace	=	2		# Space shapes evenly.

#values from visio.vislayoutvertaligntype
#****************************
visLayoutVertAlignNone	=	0		# Do not align vertically.
visLayoutVertAlignDefault	=	1		# Let Visio choose how to align vertically.
visLayoutVertAlignTop	=	2		# Align the top edges of the shapes.
visLayoutVertAlignMiddle	=	3		# Align the middles of the shapes.
visLayoutVertAlignBottom	=	4		# Align the bottom edges of the shapes.

#values from visio.vislegendflags
#****************************
visLegendPopulate	=	0		# Drop the legend and populate it.
visLegendNoContents	=	1		# Drop the legend and do not populate it.

#values from visio.vislinkreplacebehavior
#****************************
visLinkReplaceAlways	=	1		# Always replace links when linking to a shape that has existing links.
visLinkReplaceNever	=	0		# Never replace links when linking to a shape that has existing links.
visLinkReplacePrompt	=	2		# Prompt the user before replacing links when the user attempts to create links in the Microsoft Office Visio user interface.

#values from visio.vislistalignment
#****************************
visListAlignLeftOrTop	=	0		# Left-align or top-align shapes.
visListAlignCenterOrMiddle	=	1		# Center-align or middle-align shapes.
visListAlignRightOrBottom	=	2		# Right-align or bottom-align shapes.

#values from visio.vislistdirection
#****************************
visListDirLeftToRight	=	0		# Shapes are arranged horizontally, from left to right.
visListDirRightToLeft	=	1		# Shapes are arranged horizontally, from right to left.
visListDirTopToBottom	=	2		# Shapes are arranged vertically, from top to bottom.
visListDirBottomToTop	=	3		# Shapes are arranged vertically, from bottom to top.

#values from visio.vismasterproperties
#****************************
visAutomatic	=	1		# Generate icon automatically from shape data.
visCenter	=	2		# Master name center-aligned.
visDouble	=	4		# Icon size is double (64 x 64 pixels).
visIconFormatBMP	=	2		# Use bitmap icon format.
visIconFormatVisio	=	0		# Use Visio icon format.
visLeft	=	1		# Master name left-aligned.
visManual	=	0		# Generate icon manually.
visMasFPCenter	=	4096		# Fill pattern is centered.
visMasFPScale	=	16384		# Fill pattern is scaled.
visMasFPStretch	=	8192		# Fill pattern is stretched to fit available space.
visMasFPTile	=	0		# Fill pattern is tiled (repeated).
visMasIsFillPat	=	4		# Master is a fill pattern.
visMasIsLineEnd	=	2		# Master is a line end.
visMasIsLinePat	=	1		# Master is a line pattern.
visMasLEDefault	=	0		# Use default line-end properties.
visMasLEScale	=	1024		# Line end is scaled.
visMasLEUpright	=	256		# Line end always remains upright (does not rotate with the line).
visMasLPAnnotate	=	48		# Line pattern is annotated.
visMasLPScale	=	64		# Line pattern is scaled.
visMasLPStretch	=	32		# Line pattern is stretched to fit available space.
visMasLPTileDeform	=	0		# Line pattern is tiled and deformed when line curves.
visMasLPTile	=	16		# Line pattern is tiled (repeated).
visNormal	=	1		# Icon size is normal (32 x 32 pixels).
visRight	=	3		# Master name right-aligned.
visTall	=	2		# Icon size is tall (32 x 64 pixels).
visWide	=	3		# Icon size is wide (64 x 32 pixels)

#values from visio.vismastertypes
#****************************
visTypeDataGraphic	=	5		# Data graphic master.
visTypeFillPattern	=	2		# Fill pattern master.
visTypeLineEnd	=	4		# Line end master.
visTypeLinePattern	=	3		# Line pattern master.
visTypeMaster	=	1		# Generic master, no overload.
visTypeThemeColors	=	6		# Theme colors master.
visTypeThemeEffects	=	7		# Theme effects master.

#values from visio.vismeasurementsystem
#****************************
visMSDefault	=	0		# Choose metric or U.S., depending on regional options set in Control Panel.
visMSMetric	=	1		# Metric measurement system.
visMSUS	=	2		# U.S. (English) units measurement system.

#values from visio.vismemberaddoptions
#****************************
visMemberAddUseResizeSetting	=	0		# Defer to the setting of the  ContainerProperties.ResizeAsNeeded property.
visMemberAddExpandContainer	=	1		# Expand the container to fit the incoming shape(s).
visMemberAddDoNotExpand	=	2		# Do not expand the container to fit the incoming shape(s).

#values from visio.vismousemovedragstates
#****************************
visMouseMoveDragStatesBegin	=	1		# User is beginning to drag an object with the mouse.
visMouseMoveDragStatesDrop	=	5		# User dropped the dragged object in the drop-target window.
visMouseMoveDragStatesEnter	=	2		# User is dragging an object into the drop-target window with the mouse.
visMouseMoveDragStatesLeave	=	4		# User is moving the mouse out of the drop-target window.
visMouseMoveDragStatesNone	=	0		# Either not a mouse movement or a mouse movement that is not a drag operation.
visMouseMoveDragStatesOver	=	3		# User is moving the dragged object within the drop-target window with the mouse.

#values from visio.visobjecttypes
#****************************
visObjTypeAddon	=	31		# Addon object
visObjTypeAddons	=	32		# Addons collection
visObjTypeApplicationSettings	=	51		# ApplicationSettings object
visObjTypeApp	=	3		# Application object
visObjTypeCell	=	4		# Cell object
visObjTypeChars	=	5		# Characters object
visObjTypeColor	=	29		# Color object
visObjTypeColors	=	30		# Colors collection
visObjTypeComment	=	74		# Comment object
visObjTypeComments	=	74		# Comments collection
visObjTypeConnect	=	8		# Connect object
visObjTypeConnects	=	9		# Connects collection
visObjTypeContainerProperties	=	60		# ContainerProperties collection
visObjTypeCurve	=	42		# Curve object
visObjTypeDataColumn	=	56		# DataColumn object
visObjTypeDataColumns	=	55		# DataColumns collection
visObjTypeDataConnection	=	54		# DataConnection object
visObjTypeDataRecordset	=	53		# DataRecordset object
visObjTypeDataRecordsetChangedEvent	=	57		# DataRecordsetChangedEvent object
visObjTypeDataRecordsets	=	52		# DataRecordsets collection
visObjTypeDoc	=	10		# Document object
visObjTypeDocs	=	11		# Documents collection
visObjTypeEventList	=	34		# EventList collection
visObjTypeEvent	=	33		# Event object
visObjTypeFont	=	27		# Font object
visObjTypeFonts	=	28		# Fonts collection
visObjTypeGlobal	=	36		# Global object
visObjTypeGraphicItem	=	59		# GraphicItem object
visObjTypeGraphicItems	=	58		# GraphicItems collection
visObjTypeHyperlink	=	37		# Hyperlink object
visObjTypeHyperlinks	=	43		# Hyperlinks collection
visObjTypeKeyboardEvent	=	50		# KeyboardEvent object
visObjTypeLayer	=	25		# Layer object
visObjTypeLayers	=	26		# Layers collection
visObjTypeMasterShortcut	=	47		# MasterShortcut object
visObjTypeMasterShortcuts	=	46		# MasterShortcuts collection
visObjTypeMaster	=	12		# Master object
visObjTypeMasters	=	13		# Masters collection
visObjTypeMouseEvent	=	49		# MouseEvent object
visObjTypeMovedSelectionEvent	=	62		# MovedSelectionEvent object
visObjTypeMSGWrap	=	48		# MSGWrap object
visObjTypeOLEObject	=	39		# OLEObject object
visObjTypeOLEObjects	=	38		# OLEObjects collection
visObjTypePage	=	14		# Page object
visObjTypePages	=	15		# Pages collection
visObjTypePath	=	41		# Path object
visObjTypePaths	=	40		# Paths collection
visObjTypeRelatedShapePairEvent	=	61		# RelatedShapePairEvent object
visObjTypeReplaceShapesEvent	=	71		# ReplaceShapesEvent object
visObjTypeRow	=	45		# Row object
visObjTypeSection	=	44		# Section object
visObjTypeSelection	=	16		# Selection object
visObjTypeServerPublishOptions	=	63		# ServerPublishOptions object
visObjTypeShape	=	17		# Shape object
visObjTypeShapes	=	18		# Shapes collection
visObjTypeStyle	=	19		# Style object
visObjTypeStyles	=	20		# Styles object
visObjTypeUnknown	=	1		# Unknown object
visObjTypeValidation	=	64		# Validation object
visObjTypeValidationIssue	=	70		# ValidationIssue object
visObjTypeValidationIssues	=	69		# ValidationIssues collection
visObjTypeValidationRule	=	68		# ValidationRule object
visObjTypeValidationRules	=	67		# ValidationRules collection
visObjTypeValidationRuleSet	=	66		# ValidationRuleSet object
visObjTypeValidationRuleSets	=	65		# ValidationRuleSets collection
visObjTypeWindow	=	21		# Window object
visObjTypeWindows	=	22		# Windows collection

#values from visio.visoncomponententercodes
#****************************
visComponentStateModal	=	1		# The state being identified is a modal state.
visModalDeferEvents	=	0x10000		# Causes Microsoft Visio to attempt to defer firing events while modal. By default, Visio defers firing events when displaying its own dialog boxes, but does not defer firing events when client code has caused a dialog box to appear.In the case of several nested modal scopes, if any scope is deferring events, all scopes within the outermost scope that is deferring events defer events.This flag only has an effect when Visio is entering a modal scope. When exiting a modal scope, Visio behaves as it did when entering the scope.
visModalDisableVisiosFrame	=	0x80000		# Causes Visio to disable its frame window while modal. By default, Visio disables its frame window when showing its own dialog boxes or when showing dialog boxes implemented by Microsoft Visual Basic for Applications (VBA), but not when client code in another process shows a dialog box.If code in another process wants to show a dialog box and have the Visio frame window behave as if it is the Visio process showing the dialog box, it can set this flag.This flag only has an effect when entering a modal scope. When exiting a modal scope, Visio behaves as it did when entering the scope.
visModalDontBlockMessages	=	0x40000		# Prevents Visio from rejecting calls from outside its main thread while modal. By default, Visio does reject calls from outside its thread while modal.In the case of several nested modal scopes, if any scope is deferring events, all scopes within the outermost scope that is deferring events defer events.This flag only has an effect when entering a modal scope. When exiting a modal scope, Visio behaves as it did when entering the scope.
visModalNoBeforeAfter	=	0x20000		# Prevents Visio from firing a  BeforeModal event when entering a modal scope or an AfterModal event when leaving a modal scope.By default, Visio fires these events when displaying its own dialog boxes or displaying dialog boxes implemented by VBA, but does not fire these events when client code displays a dialog box.Calling the  OnComponentEnterState method causes these events to fire unless visModalNoBeforeAfter is specified.

#values from visio.visopensaveargs
#****************************
visAddDeclineAutoRefresh	=	1024		# Adds a document without displaying the  Configure Refresh dialog box.
visAddDocked	=	4		# Adds a document in a docked window.
visAddHidden	=	64		# Adds a document in a hidden window.
visAddMacrosDisabled	=	128		# Adds a document with macros disabled.
visAddMinimized	=	16		# Adds a document in a minimized window
visAddNoWorkspace	=	256		# Adds a document with no workspace information.
visAddStencil	=	512		# Adds a new stencil file.
visOpenCopy	=	1		# Opens a copy of the document.
visOpenCopyOfNaming	=	2048		# Opens a copy of the document, using a copy of the naming.
visOpenDeclineAutoRefresh	=	1024		# Opens a document without displaying the  Configure Refresh dialog box.
visOpenDocked	=	4		# Opens a stencil in a docked window.
visOpenDontList	=	8		# Opens the document without adding it to the Most Recently Used (MRU) list.
visOpenHidden	=	64		# Opens the document in a hidden window.
visOpenMacrosDisabled	=	128		# Opens the document with macros disabled.
visOpenMinimized	=	16		# Opens the document in a minimized window.
visOpenNoWorkspace	=	256		# Opens the document with no workspace information.
visOpenRO	=	2		# Opens the document as read-only.
visOpenRW	=	32		# Opens the document for both reading and writing.
visSaveAsCheckCompatibility	=	8		# Displays the  Compatibility Checker dialog box on save.
visSaveAsListInMRU	=	4		# Saves the document and puts it in the MRU list.
visSaveAsRO	=	1		# Saves the document as read-only.
visSaveAsWS	=	2		# Saves the workspace and the file.

#values from visio.vispageandmasterids
#****************************
visInvalMasterID	=	-1		# An ID no master will ever have.
visInvalPageID	=	-1		# An ID no master will ever have.

#values from visio.vispagesizingbehaviors
#****************************
visNeverResizePages	=	0		# Do not automatically resize pages under any circumstances.
visResizePages	=	1		# Automatically resize all pages when the Microsoft Visio Drawing Control is resized or when a new document is loaded into it. Leave shapes unchanged.

#values from visio.vispagetypes
#****************************
visPageTypeInval	=	0		# Not a Microsoft Visio page.
visTypeBackground	=	2		# A background page.
visTypeForeground	=	1		# A foreground page.
visTypeMarkup	=	3		# An annotation page.

#values from visio.vispapersizes
#****************************
visPaperSizeA3	=	8		# A3 297 x 420 mm
visPaperSizeA4	=	9		# A4 210 x 297 mm
visPaperSizeA5	=	11		# A5 148 x 210 mm
visPaperSizeB4	=	12		# B4 (JIS) 250 x 354 mm
visPaperSizeB5	=	13		# B5 (JIS) 182 x 257 mm
visPaperSizeFolio	=	14		# Folio 8 1/2 x 13 in.
visPaperSizeLegal	=	5		# Legal 8 1/2 x 14 in.
visPaperSizeLetter	=	1		# Letter 8 1/2 x 11 in.
visPaperSizeNote	=	18		# Note 8 1/2 x 11 in.
visPaperSizeSizeC	=	24		# C size sheet 17 x 22 in.
visPaperSizeSizeD	=	25		# D size sheet 22 x 34 in.
visPaperSizeSizeE	=	26		# E size sheet 34 x 44 in.
visPaperSizeUnknown	=	0		# Unknown

#values from visio.vispastespecialcodes
#****************************
visPasteBitmap	=	2		# Paste bitmap.
visPasteDIB	=	8		# Paste device-independent bitmap.
visPasteEMF	=	14		# Paste enhanced metafile.
visPasteHyperlink	=	65538		# Paste hyperlink.
visPasteInk	=	65544		# Paste Ink data.
visPasteMetafile	=	3		# Paste metafile.
visPasteOEMText	=	7		# Paste OEM text.
visPasteOLEObject	=	65536		# Paste OLE object.
visPasteRichText	=	65537		# Paste rich text.
visPasteText	=	1		# Paste ANSI text.
visPasteURL	=	65539		# Paste Uniform Resource Locator (URL).
visPasteVisioIcon	=	65543		# Paste Microsoft Visio icon.
visPasteVisioMastersXML	=	65546		# Paste Visio masters XML.
visPasteVisioMasters	=	65541		# Paste Visio masters.
visPasteVisioShapesXML	=	65545		# Paste Visio shapes XML.
visPasteVisioShapesWithoutDataLinks	=	65548		# Paste Visio drawing data without internal data links.
visPasteVisioShapes	=	65540		# Paste Visio shapes.
visPasteVisioText	=	65542		# Paste Visio text.

#values from visio.visprimarykeysettings
#****************************
visKeyComposite	=	3		# Use multiple columns as primary key columns.
visKeyRowOrder	=	1		# Use row order as the primary key.
visKeySingle	=	2		# Use a single column as the primary key column.

#values from visio.visprintoutrange
#****************************
visPrintAll	=	0		# Print all foreground pages.
visPrintCurrentPage	=	2		# Print current page.
visPrintCurrentView	=	4		# Print current view area.
visPrintFromTo	=	1		# Print pages between from index and to index.
visPrintSelection	=	3		# Print selection.

#values from visio.visprotection
#****************************
visProtectBackgrounds	=	0x8		# Protect document backgrounds from user customization.
visProtectMasters	=	0x4		# Protect document masters from user customization.
visProtectNone	=	0x0		# Document is unprotected.
visProtectPreviews	=	0x10		# Protect document previews from user customization.
visProtectShapes	=	0x2		# Protect document shapes from user customization.
visProtectStyles	=	0x1		# Protect document styles from user customization.

#values from visio.vispublishdatarecordsets
#****************************
visPublishDataRecordsetAll	=	0		# Publish all data recordsets in the document.
visPublishDataRecordsetNone	=	1		# Publish none of the data recordsets in the document.
visPublishDataRecordsetSelect	=	2		# Publish selected data recordsets.

#values from visio.vispublishpages
#****************************
visPublishPageAll	=	0		# Publish all pages.
visPublishPageSelect	=	1		# Publish selected pages.

#values from visio.visquickstylecolors
#****************************


#values from visio.visquickstylematrixindices
#****************************


#values from visio.visrasterexportcolorformat
#****************************
visRasterBiLevel	=	0		# Bi-level color format

#values from visio.visrasterexportcolorreduction
#****************************
visRasterAdaptive	=	0		# Adaptive color reduction; the default for GIF files.
visRasterDiffusion	=	1		# Diffusion color reduction.
visRasterHalftone	=	2		# Halftone color reduction.

#values from visio.visrasterexportdatacompression
#****************************
visRasterNone	=	0		# No compression; the default for BMP.

#values from visio.visrasterexportdataformat
#****************************
visRasterInterlace	=	0		# Interlace format; the default.
visRasterNonInterlace	=	1		# Non-interlace format.

#values from visio.visrasterexportflip
#****************************
visRasterNoFlip	=	0		# No flip, the default.
visRasterFlipHorizontal	=	1		# Flip horizontally.
visRasterFlipVertical	=	2		# Flip vertically.

#values from visio.visrasterexportoperation
#****************************
visRasterBaseline	=	0		# Baseline operation; the default.
visRasterProgressive	=	1		# Progressive operation.

#values from visio.visrasterexportresolution
#****************************
visRasterUseScreenResolution	=	0		# Use screen resolution.
visRasterUsePrinterResolution	=	1		# Use printer resolution.
visRasterUseSourceResolution	=	2		# Use source resolution.
visRasterUseCustomResolution	=	3		# Use custom resolution.

#values from visio.visrasterexportresolutionunits
#****************************
visRasterPixelsPerInch	=	0		# Pixels per inch.
visRasterPixelsPerCm	=	1		# Pixels per centimeter.

#values from visio.visrasterexportrotation
#****************************
visRasterNoRotation	=	0		# No rotation; the default.
visRasterRotateLeft	=	1		# Rotate left.
visRasterRotateRight	=	2		# Rotate right.

#values from visio.visrasterexportsize
#****************************
visRasterFitToScreenSize	=	0		# Use screen size.
visRasterFitToPrinterSize	=	1		# Use printer size.
visRasterFitToSourceSize	=	2		# Use source size.
visRasterFitToCustomSize	=	3		# Use custom size.

#values from visio.visrasterexportsizeunits
#****************************
visRasterPixel	=	0		# Pixels
visRasterCm	=	1		# Centimeters
visRasterInch	=	2		# Inches

#values from visio.visrecordsetfieldstatus
#****************************
visFieldMappedAllCallouts	=	3
visFieldMappedNoCallouts	=	1
visFieldMappedSomeCallouts	=	2
visFieldNotMapped	=	0

#values from visio.visrefreshsettings
#****************************
visRefreshNoReconcilationUI	=	2		# Disables display of the  Refresh Conflicts task pane in the Microsoft Visio user interface after data is refreshed.
visRefreshOverwriteAll	=	1		# When data is refreshed, overwrites all user changes made since the previous refresh operation.

#values from visio.visregionaluioptions
#****************************
visRegionalUIOptionsHide	=	0		# Always hide regional UI.
visRegionalUIOptionsShow	=	1		# Always show regional UI.
visRegionalUIOptionsUseSystemSettings	=	-1		# Not used.

#values from visio.visremovehiddeninfoitems
#****************************
visRHIDataRecordsets	=	16		# Data recordsets.
visRHIMasters	=	4		# Unused masters
visRHINone	=	0		# No information.
visRHIPersonalInfo	=	1		# Personal information.
visRHIPreview	=	2		# Preview thumbnail.
visRHIStyles	=	8		# Unused styles and display formats.
visRHIValidationRules	=	32		# Validation rules.

#values from visio.visreplaceflags
#****************************


#values from visio.visresizedirection
#****************************
visResizeDirE	=	0		# Middle right.
visResizeDirNE	=	1		# Top right.
visResizeDirN	=	2		# Top center.
visResizeDirNW	=	3		# Top left.
visResizeDirW	=	4		# Middle left.
visResizeDirSW	=	5		# Bottom left.
visResizeDirS	=	6		# Bottom center.
visResizeDirSE	=	7		# Bottom right.

#values from visio.visribbonxmodes
#****************************
visRXModeNone	=	0		# Display the custom user interface (UI) when no document is active.
visRXModeDrawing	=	1		# Display the custom UI in Drawing mode.
visRXModeStencil	=	2		# Display the custom UI in Stencil mode.
visRXModePrintPreview	=	4		# Display the custom UI in Print Preview mode.

#values from visio.visroleselectiontypes
#****************************
visRoleSelConnector	=	1		# A selection that contains all connector shapes.
visRoleSelContainer	=	2		# A selection that contains all container shapes.
visRoleSelCallout	=	4		# A selection that contains all callout shapes.

#values from visio.visrotationtypes
#****************************
visRotateSelectionWithPin	=	1		# Rotate the selection around a pin.
visRotateSelection	=	0		# Rotate the selection relative to the center of the selection.
visRotateShapes	=	2		# Rotate the selected shapes around their pins relative to their current angle.

#values from visio.visroundflags
#****************************
visRound	=	1		# Round the result.
visTruncate	=	0		# Truncate the result.

#values from visio.visrowindices
#****************************
visRow1stHyperlink	=	0		# Index of the first row in  visSectionHyperlink.
visRow3DRotationProperties	=	30		# Index of the row in  visSectionObject that defines the 3D rotation properties of the shape.
visRowAction	=	0		# Index of the first row in  visSectionAction.
visRowAlign	=	14		# Index of the row in  visSectionObject that defines the shape's alignment.
visRowAnnotation	=	0		# Index of the first row in  visSectionAnnotation.
visRowBevelProperties	=	29		# Index of the row in  visSectionObject that defines the bevel properties of the shape.
visRowCharacter	=	0		# Index of the first row in  visSectionCharacter.
visRowComponent	=	0		# Index of the component properties row in a Geometry section (visSectionFirstComponent +).
visRowConnectionPts	=	0		# Index of the first row in  visSectionConnectionPts.
visRowControl	=	0		# Index of the first row in  visSectionControl.
visRowDoc	=	20		# Index of the row in  visSectionObject that contains document properties.
visRowEvent	=	5		# Index of the row in  visSectionObject that contains event information.
visRowField	=	0		# Index of the first row in  visSectionTextField.
visRowFill	=	3		# Index of the row in  visSectionObject that defines fill properties.
visRowFirst	=	0		# Row logically before every row in a section.
visRowForeign	=	9		# Index of the row in  visSectionObject that defines foreign properties (shape of type visTypeForeignObject).
visRowGradientProperties	=	26		# Index of the row in  visSectionObject that defines the gradient properties of the shape.
visRowGradientStop	=	0		# Index of first row in  visSectionLineGradientStops and visSectionFillGradientStops.
visRowGroup	=	22		# Index of the row in  visSectionObject that defines foreign properties (shape of type visTypeGroup).
visRowHelpCopyright	=	16		# Index of the row in  visSectionObject that defines Help and copyright properties.
visRowImage	=	21		# Index of the row in  visSectionObject that defines image properties (shape whose property is visTypeBitMap).
visRowInval	=	-1		# An invalid row index.
visRowLast	=	-2		# Row logically after every row in a section.
visRowLayerMem	=	6		# Index of the row in  visSectionObject that defines what layers the shape belongs to.
visRowLayer	=	0		# Index of the first row in  visSectionLayer.
visRowLine	=	2		# Index of the row in  visSectionObject that defines line properties.
visRowLock	=	15		# Index of the row in  visSectionObject that defines its lock properties.
visRowMisc	=	17		# Index of the row in  visSectionObject that defines miscellaneous behaviors.
visRowNone	=	-1		# Unspecified row.
visRowPageLayout	=	24		# Index of the row in  visSectionObject of a page or master that defines placement and routing.
visRowOtherEffectProperties	=	28		# Index of the row in  visSectionObject that defines other effect properties.
visRowPage	=	10		# Index of the row in  visSectionObject that defines page or master properties (shape of type visTypePage).
visRowParagraph	=	0		# Index of the first row in  visSectionParagraph.
visRowPrintProperties	=	25		# Index of the row in  visSectionObject of a document that defines printing properties. (Print Properties section in the ShapeSheet window.)
visRowProp	=	0		# Index of the first row in  visSectionProp.
visRowQuickStyleProperties	=	27		# Index of the row in  visSectionObject that defines QuickStyle properties.
visRowReplaceBehaviors	=	32		# Index of the row in  visSectionObject that defines replace-shape behaviors.
visRowReviewer	=	0		# Index of the first row in  visSectionReview.
visRowRulerGrid	=	18		# Index of the row in  visSectionObject of a page or master that defines the ruler and grid settings.
visRowScratch	=	0		# Index of the first row in  visSectionScratch.
visRowShapeLayout	=	23		# Index of the row in  visSectionObject of shape that defines placement and routing.
visRowSmartTag	=	0		# Index of the first row in  visSectionSmartTag.
visRowStyle	=	8		# Index of the row in  visSectionObject that defines style properties.
visRowTab	=	0		# Index of the first row in visSectionTab.
visRowTextXForm	=	12		# Index of the row in  visSectionObject that defines a shape's text transform properties.
visRowText	=	11		# Index of the row in  visSectionObject that defines a shape or style's text block properties.
visRowThemeProperties	=	31		# Index of the row in  visSectionObject that defines theme properties for a shape.
visRowUser	=	0		# Index of the first row in  visSectionUser.
visRowVertex	=	1		# Index of the first vertex row in a Geometry section.
visRowXForm1D	=	4		# Index of the row in 1D shape's  visSectionObject that defines its endpoints.
visRowXFormOut	=	1		# Index of the row in  visSectionObject that defines the shape's transform properties.

#values from visio.visrowtags
#****************************
visTagArcTo	=	140		# The row type of an ArcTo row in a Geometry section.
visTagCnnctNamed	=	185		# The row type of a row in a  visSectionConnectionPts section that has named rows.
visTagCnnctNamedABCD	=	187		# The row type of an extended row in a  visSectionConnectionPts section that has named rows. Seldom used.
visTagCnnctPt	=	153		# The row type of a row in a  visSectionConnectionPts section that has unnamed rows.
visTagCnnctPtABCD	=	186		# The row type of an extended row in a  visSectionConnectionPts section that has unnamed rows. Seldom used.
visTagComponent	=	137		# The row type of the component properties row in a Geometry section.
visTagCtlPt	=	162		# The row type of a row in  visSectionControls that doesn't supply a ToolTip.
visTagCtlPtTip	=	170		# The row type of a row in  visSectionControls that supplies a ToolTip.
visTagDefault	=	0		# Connotes row of default type to  AddRow, AddNamedRows, or AddRows methods.
visTagEllipse	=	143		# The row type of an Ellipse row in a Geometry section.
visTagEllipticalArcTo	=	144		# The row type of an EllipticalArcTo row in a Geometry section.
visTagInfiniteLine	=	141		# The row type of an InfiniteLine row in a Geometry section.
visTagLineTo	=	139		# The row type of a LineTo row in a Geometry section.
visTagMoveTo	=	138		# The row type of a MoveTo row in a Geometry section.
visTagNURBSTo	=	195		# The row type of a NURBSTo row in a Geometry section.
visTagPolylineTo	=	193		# The row type of a PolylineTo row in a Geometry section.
visTagRelCubBezTo	=	236		# The row type of a RelCubBezTo row in a Geometry section.
visTagRelEllipticalArcTo	=	240		# The row type of a RelEllipticalArcTo row in a Geometry section.
visTagRelLineTo	=	239		# The row type of a RelLineTo row in a Geometry section.
visTagRelMoveTo	=	238		# The row type of a RelMoveTo row in a Geometry section.
visTagRelQuadBezTo	=	23777		# The row type of a RelQuadBezTo row in a Geometry section.
visTagSplineBeg	=	165		# The row type of a SplineStart row in a Geometry section.
visTagSplineSpan	=	166		# The row type of a SplineKnot row in a Geometry section.
visTagTab0	=	136		# The row type of a row in a  visSectionTab section that defines 0 (zero) tab stops.
visTagTab10	=	151		# The row type of a row in a  visSectionTab section that defines up to 10 tab stops.
visTagTab2	=	150		# The row type of a row in a  visSectionTab section that defines up to 2 tab stops.
visTagTab60	=	181		# The row type of a row in a  visSectionTab section that defines up to 60 tab stops.

#values from visio.visrulesetflags
#****************************
visRuleSetDefault	=	0		# The default rule-set property. The rule set appears in the  Rules to Check list (click the Check Diagram arrow on the Process tab).
visRuleSetHidden	=	1		# The rule set does not appear in the  Rules to Check list.

#values from visio.visruletargets
#****************************
visRuleTargetDocument	=	2		# The rule applies to the document itself.
visRuleTargetPage	=	1		# The rule applies to pages in the document.
visRuleTargetShape	=	0		# The rule applies to shapes in the document.

#values from visio.visruntypes
#****************************
visCharPropRow	=	1		# Reports runs of characters that have common character properties. Corresponds to a set of characters covered by one row in a shape's Character section.
visFieldRun	=	20		# Reports runs whose boundaries are between characters that are and aren't the result of the expansion of a text field, or between characters that are the result of the expansion of distinct text fields.
visParaPropRow	=	2		# Reports runs of characters that have common paragraph properties. Corresponds to a set of characters covered by one row in the shape's Paragraph section.
visParaRun	=	11		# Reports runs whose boundaries are between successive paragraphs in the shape's text. Mimics triple-clicking to select text.
visTabPropRow	=	3		# Reports runs of characters that have common tab properties. Corresponds to a set of characters that are covered by one row in shape's Tabs section.
visWordRun	=	10		# Reports runs whose boundaries are between successive words in a shape's text. Mimics double-clicking to select text.

#values from visio.vissavepreviewmode
#****************************
visSavePreviewDraft1st	=	1		# The first page; includes only Microsoft Visio shapes. Does not include embedded objects, text, or gradient fills.
visSavePreviewDraftAll	=	4		# All file pages; includes only Visio shapes. Does not include embedded objects, text, or gradient fills.
visSavePreviewNone	=	0		# No preview picture.

#values from visio.visscrollbarstates
#****************************
visScrollBarBoth	=	0x5		# Show both scrollbars.
visScrollBarHoriz	=	0x1		# Show the horizontal scrollbar.
visScrollBarNeither	=	0x0		# Show neither scrollbar.
visScrollBarVert	=	0x4		# Show the vertical scrollbar.

#values from visio.vissectionindices
#****************************
visSectionAction	=	240		# Stores the actions that appear on the shortcut menu.
visSectionAnnotation	=	246		# Index of a section whose rows represent annotations.
visSectionCharacter	=	3		# Stores character properties; for example, font.
visSectionConnectionPts	=	7		# Stores an object's connection points.
visSectionControls	=	9		# Stores an object's control handles.
visSectionFillGradientStops	=	249		# Index of a section whose rows represent fill gradient stops.
visSectionFirstComponent	=	10		# An object's first Geometry section. Additional sections have indices (visSectionFirstComponent + ).
visSectionFirst	=	0		# Index whose value is less than any other section index.
visSectionHyperlink	=	244		# Stores hyperlinks.
visSectionInval	=	255		# An invalid index that no section will ever have.
visSectionLastComponent	=	239		# An object's last Geometry section.
visSectionLast	=	252		# Index whose value is greater than any other section index.
visSectionLayer	=	241		# Stores a page or master's layer properties.
visSectionLineGradientStops	=	248		# Index of a section whose rows represent line gradient stops.
visSectionNone	=	255		# Unspecified section.
visSectionObject	=	1		# Stores general non-repeating properties of an object.
visSectionParagraph	=	4		# Stores paragraph properties; for example, indentation.
visSectionProp	=	243		# Stores shape data (formerly custom properties).
visSectionReviewer	=	245		# Index of section whose rows represent reviewers.
visSectionScratch	=	6		# Holds scratch cells.
visSectionSmartTag	=	247		# Index of section whose rows represent SmartTags.
visSectionTab	=	5		# Stores position and alignment of tab stops.
visSectionTextField	=	8		# Stores an object's text fields.
visSectionUser	=	242		# Stores cells created and used by an external solution.

#values from visio.visselectargs
#****************************
visDeselect	=	1		# Deselects a shape but leaves the rest of the selection unchanged.
visSelect	=	2		# Selects a shape but leaves the rest of the selection unchanged.
visSubSelect	=	3		# Selects a shape whose parent is already selected.
visSelectAll	=	4		# Selects a shape and all its peers.
visDeselectAll	=	256		# Deselects a shape and all its peers.

#values from visio.visselectiontypes
#****************************
visSelTypeAll	=	1		# A selection that initially contains all shapes.
visSelTypeByDataGraphic	=	6		# A selection that initially contains all shapes that have a given type of data graphic appled.
visSelTypeByLayer	=	3		# A selection that initially contains all the shapes of a given layer.
visSelTypeByMaster	=	5		# A selection that initially contains all the instantiated shapes of a given master.
visSelTypeByRole	=	7		# A selection that initially contains all the shapes that have a given role.
visSelTypeByType	=	4		# A selection that initially contains all the shapes of a given type.
visSelTypeEmpty	=	0		# A selection that initially contains no shapes.
visSelTypeSingle	=	2		# A selection that initially contains one shape.

#values from visio.visselectitemstatus
#****************************
visSelIsPrimaryItem	=	0x1		# The item is the primary item.
visSelIsSubItem	=	0x2		# The item is a subselected item.
visSelIsSuperItem	=	0x4		# The item is a superselected item.

#values from visio.visselectmode
#****************************
visSelModeOnlySub	=	0x0800		# Selection reports only subselected shapes.
visSelModeOnlySuper	=	0x0200		# Selection reports only superselected shapes.
visSelModeSkipSub	=	0x0400		# Selection does not report subselected shapes.
visSelModeSkipSuper	=	0x0100		# Selection does not report superselected shapes.

#values from visio.visshapeids
#****************************
visInvalShapeID	=	-1		# An ID no shape will ever have; used for comparison against valid ID values that are returned by the  Shape.ID and Shapes.ItemFromID properties.
visPageSheetID	=	0		# The ID of a page's or master's page sheet; the value of the  Shape.ID property when the shape is a page sheet.

#values from visio.visshapetypes
#****************************
visTypeBitmap	=	32		# Returned by  Shape.ForeignType if the shape is a bitmap.
visTypeDoc	=	6		# The document's  DocumentSheet.
visTypeForeignObject	=	4		# An imported shape.
visTypeGroup	=	2		# A shape that contains other shapes.
visTypeGuide	=	5		# A shape that is a guide.
visTypeInk	=	64		# Returned by  Shape.ForeignType if the shape is ink.
visTypeInval	=	0		# The type of no shape. Means all types when used as filter code.
visTypeIsControl	=	1024		# Returned by  Shape.ForeignType if the shape is a control.
visTypeIsEmbedded	=	512		# Returned by  Shape.ForeignType if the shape is embedded.
visTypeIsLinked	=	256		# Returned by  Shape.ForeignType if the shape is linked.
visTypeIsOLE2	=	32768		# Returned by  Shape.ForeignType if the shape is linked, embedded, or a control.
visTypeMetafile	=	16		# Returned by  Shape.ForeignType if the shape is a metafile.
visTypePage	=	1		# Page's or master's  PageSheet property.
visTypeShape	=	3		# Native Microsoft Visio shape.

#values from visio.vissnapextensions
#****************************
visSnapExtAlignmentBoxExtension	=	0x1		# Show alignment box extensions.
visSnapExtCenterAxes	=	0x2		# Show center alignment axes.
visSnapExtCurveExtension	=	0x40		# Show curved extensions
visSnapExtCurveTangent	=	0x4		# Show curve interior tangents.
visSnapExtEllipseCenter	=	0x800		# Show ellipse center points.
visSnapExtEndpointHorizontal	=	0x200		# Show horizontal lines at endpoint.
visSnapExtEndpointPerpendicular	=	0x80		# Show endpoint perpendicular lines.
visSnapExtEndpointVertical	=	0x400		# Show vertical lines at endpoint.
visSnapExtEndpoint	=	0x8		# Show segment endpoints.
visSnapExtIsometricAngles	=	0x1000		# Show isometric angle lines.
visSnapExtLinearExtension	=	0x20		# Show linear extensions.
visSnapExtMidpointPerpendicular	=	0x100		# Show midpoint perpendicular lines.
visSnapExtMidpoint	=	0x10		# Show segment midpoints.
visSnapExtNone	=	0x0		# Show no extentions.

#values from visio.vissnapsettings
#****************************
visSnapToAlignmentBox	=	0x200		# Snap to the alignment box.
visSnapToConnectionPoints	=	0x20		# Snap to connection points.
visSnapToDisabled	=	0x8000		# Disable snap.
visSnapToExtensions	=	0x400		# Snap to shape extensions options.
visSnapToGeometry	=	0x100		# Snap to the visible edges of shapes.
visSnapToGrid	=	0x2		# Snap to the grid.
visSnapToGuides	=	0x4		# Snap to guides.
visSnapToHandles	=	0x8		# Snap to selection handles.
visSnapToIntersections	=	0x10000		# Snap to intersections.
visSnapToNone	=	0x0		# Snap to nothing.
visSnapToRulerSubdivisions	=	0x1		# Snap to tick marks on the ruler.
visSnapToVertices	=	0x10		# Snap to vertices.

#values from visio.visspatialrelationcodes
#****************************
visSpatialContainedIn	=	0x4		# A shape can be contained within another shape. Shape B is contained within shape A if shape A encloses every region and path of shape B.
visSpatialContain	=	0x2		# A shape can contain another shape. Shape A contains shape B if shape A encloses every region and path of shape B.
visSpatialOverlap	=	0x1		# Two shapes can overlap. Shapes overlap if their interior regions have at least one point in common. You will also get this result if you compare a shape to itself or if either shape is a sub-shape of the other.
visSpatialTouching	=	0x8		# A shape can be touching another shape. Shape A touches shape B if neither one contains or overlaps the other and they have one or more common points whose distance is within the specified tolerance.

#values from visio.visspatialrelationflags
#****************************
visSpatialBackToFront	=	0x8		# Order items back to front.
visSpatialFrontToBack	=	0x4		# Order items front to back.
visSpatialIgnoreVisible	=	0x20		# Do not consider visible Geometry sections. By default, visible Geometry sections influence the result.
visSpatialIncludeContainerShapes	=	0x80		# Include containers. By default, containers are not included.
visSpatialIncludeDataGraphics	=	0x40		# Include data graphic callout shapes and their sub-shapes. By default, data graphic callout shapes and their subshapes are not included. If the parent shape is itself a data graphic callout, searches are made between the parent shape's geometry and non-callout shapes, unless this flag is set.
visSpatialIncludeGuides	=	0x2		# Consider a guide's Geometry section. By default, guides do not influence the result.
visSpatialIncludeHidden	=	0x10		# Consider hidden Geometry sections. By default, hidden Geometry sections do not influence the result.

#values from visio.visstatcodes
#****************************
visStatAppHasShutdown	=	1		# The application has stopped.
visStatClosed	=	8		# Object is closed.
visStatDeleted	=	2		# Object is deleted.
visStatNormal	=	0		# Object status is normal.
visStatSuspended	=	16		# Object status is suspended.

#values from visio.vissvgexportformat
#****************************


#values from visio.visthemecolors
#****************************
visThemeColorsNone	=	0		# No theme colors
visThemeColorsMonochrome	=	1		# Monochrome
visThemeColorsOffice	=	2		# Microsoft Office colors
visThemeColorsMedian	=	3		# Median
visThemeColorsConcourse	=	4		# Concourse
visThemeColorsSolstice	=	5		# Solstice
visThemeColorsTechnic	=	6		# Technic
visThemeColorsPaper	=	7		# Paper
visThemeColorsFoundry	=	8		# Foundry
visThemeColorsApex	=	9		# Apex
visThemeColorsTrek	=	10		# Trek
visThemeColorsModule	=	11		# Module
visThemeColorsOriel	=	12		# Oriel
visThemeColorsAspect	=	13		# Aspect
visThemeColorsEquity	=	14		# Equity
visThemeColorsCivic	=	15		# Civic
visThemeColorsOpulent	=	16		# Opulent
visThemeColorsVerve	=	17		# Verve
visThemeColorsOrigin	=	18		# Origin
visThemeColorsUrban	=	19		# Urban
visThemeColorsFlow	=	20		# Flow
visThemeColorsMetro	=	21		# Metro
visThemeColorsOfficeLight	=	22		# Microsoft Office colors light
visThemeColorsOfficeDark	=	23		# Microsoft Office colors dark
visThemeColorsMedianLight	=	24		# Median light
visThemeColorsMedianDark	=	25		# Median dark
visThemeColorsConcourseLight	=	26		# Concourse light
visThemeColorsConcourseDark	=	27		# Concourse dark
visThemeColorsPaperLight	=	28		# Paper light
visThemeColorsPaperDark	=	29		# Paper dark
visThemeColorsFoundryLight	=	30		# Foundry light
visThemeColorsFoundryDark	=	31		# Foundry dark
visThemeColorsEquityLight	=	32		# Equity light
visThemeColorsEquityDark	=	33		# Equity dark
visThemeColorsVerveLight	=	34		# Verve light
visThemeColorsVerveDark	=	35		# Verve dark
visThemeColorsBasic	=	36		# Basic
visThemeColorsAdjacency	=	37		# Adjacency
visThemeColorsAngles	=	38		# Angles
visThemeColorsApothecary	=	39		# Apothecary
visThemeColorsAustin	=	40		# Austin
visThemeColorsEssential	=	41		# Essential
visThemeColorsBlackTie	=	42		# Black tie
visThemeColorsComposite	=	43		# Composite
visThemeColorsClarity	=	44		# Clarity
visThemeColorsElemental	=	45		# Elemental
visThemeColorsExecutive	=	46		# Executive
visThemeColorsGrid	=	47		# Grid
visThemeColorsHardcover	=	48		# Hardcover
visThemeColorsHorizon	=	49		# Horizon
visThemeColorsNewsprint	=	50		# Newsprint
visThemeColorsCouture	=	51		# Couture
visThemeColorsPerspective	=	52		# Perspective
visThemeColorsPushpin	=	53		# Pushpin
visThemeColorsSlipstream	=	54		# Slipstream
visThemeColorsThatch	=	55		# Thatch
visThemeColorsWaveform	=	56		# Waveform

#values from visio.visthemeeffects
#****************************
visThemeEffectsNone	=	0		# No theme effect.
visThemeEffectsSubdued	=	1		# Subdued.
visThemeEffectsSimpleShadow	=	2		# Simple shadow.
visThemeEffectsButton	=	3		# Button.
visThemeEffectsSquare	=	4		# Square.
visThemeEffectsPillow	=	5		# Pillow
visThemeEffectsBevelIllusion	=	6		# Bevel illusion.
visThemeEffectsBevelHighlight	=	7		# Bevel highlight.
visThemeEffectsOutline	=	8		# Outline.
visThemeEffectsDecal	=	9		# Decal.
visThemeEffectsRaisedSurface	=	10		# Raised surface.
visThemeEffectsMesh	=	11		# Mesh.
visThemeEffectsPinstripe	=	12		# Pinstripe.
visThemeEffectsStripes	=	13		# Stripes.
visThemeEffectsOblique	=	14		# Oblique.
visThemeEffectsToy	=	15		# Toy.
visThemeEffectsBasicShadow	=	16		# Basic shadow.

#values from visio.visthemetypes
#****************************
visThemeTypeColor	=	1		# Theme colors.
visThemeTypeConnector	=	3		# Theme connectors.
visThemeTypeEffect	=	2		# Theme effects.
visThemeTypeFont	=	4		# Theme fonts.
visThemeTypeIndex	=	0		# Theme indices.

#values from visio.vistoparts
#****************************
visConnectionPoint	=	100		# Connect to specified connection point on target shape.
visConnectToError	=	-1		# Error connecting to shape.
visGuideIntersect	=	4		# Connect to intersection of guides on target shape.
visGuideX	=	1		# Connect to vertical guide on target shape.
visGuideY	=	2		# Connect to horizontal guide on target shape.
visToAngle	=	7		# Connect to angle on target shape.
visToNone	=	0		# Do not connect.
visWholeShape	=	3		# Connect to entire target shape, using dynamic glue.

#values from visio.vistraceflags
#****************************
visTraceAddonInvokes	=	0x4		# Add-on invocations.
visTraceAdvises	=	0x2		# Outgoing advise calls.
visTraceCallsToVBA	=	0x8		# Microsoft Visual Basic for Applications (VBA) invocations.
visTraceEvents	=	0x1		# Event occurrences.

#values from visio.vistypeselectiontypes
#****************************
visTypeSelBitmap	=	16		# A shape that is a bitmap.
visTypeSelGroup	=	1		# A shape that contains other shapes.
visTypeSelGuide	=	4		# A shape that is a guide.
visTypeSelInk	=	32		# A shape that is ink.
visTypeSelMetafile	=	8		# A shape that is a metafile.
visTypeSelOLE	=	64		# A shape that is linked, embedded, or a control.
visTypeSelShape	=	2		# A native Visio shape.

#values from visio.visuibarposition
#****************************
visBarBottom	=	3		# Display docked at bottom of drawing window.
visBarFloating	=	4		# Float in drawing window.
visBarLeft	=	0		# Display docked at left of drawing window.
visBarMenu	=	6		# Display in the menu bar.
visBarPopup	=	5		# Float in drawing window.
visBarRight	=	2		# Display docked at right of drawing window.
visBarTop	=	1		# Display docked at top of drawing window.

#values from visio.visuibarprotection
#****************************
visBarNoChangeDock	=	16		# Can't be docked or floating.
visBarNoCustomize	=	1		# Can't be customized.
visBarNoHorizontalDock	=	64		# Can't be docked horizontally.
visBarNoMove	=	4		# Can't be moved.
visBarNoProtection	=	0		# No protection.
visBarNoResize	=	2		# Can't be resized.
visBarNoVerticalDock	=	32		# Can't be docked vertically.

#values from visio.visuibarrow
#****************************
visBarRowFirst	=	0		# First row.
visBarRowLast	=	-1		# Last row.

#values from visio.visuibuttonstate
#****************************
visButtonDown	=	-1		# Button is down.
visButtonUp	=	0		# Button is up.

#values from visio.visuibuttonstyle
#****************************
visButtonAutomatic	=	0		# Default style.
visButtonCaption	=	2		# Text only (always).
visButtonIconandCaption	=	3		# Image and text.
visButtonIcon	=	1		# Text only (in menus).

#values from visio.visuicmds
#****************************
visCmdABarAutoHeight	=	1684
visCmdABarAutohide	=	1652
visCmdABarHide	=	1650
visCmdABarToggleFloat	=	1651
visCmdAcquireImages	=	1868
visCmdAddConnectPt	=	1263
visCmdAddControlPt	=	1266
visCmdAddDataRecordset	=	1998
visCmdAddMemberToContainer	=	2173
visCmdAddTextShape	=	1181
visCmdAddToAllContainers	=	2251
visCmdAddToNewContainer	=	2268
visCmdAlignBox	=	1768
visCmdAlignObjectBottom	=	1201
visCmdAlignObjectCenter	=	1197
visCmdAlignObjectLeft	=	1196
visCmdAlignObjectMiddle	=	1200
visCmdAlignObjectRight	=	1198
visCmdAlignObjectTop	=	1199
visCmdAllowThemes	=	2056
visCmdApplyDataGraphic	=	2017
visCmdApplyDataGraphicAfterLink	=	2092
visCmdApplyMainTheme	=	2285
visCmdApplyMainThemeToDocument	=	2296
visCmdApplyMainThemeToPage	=	2289
visCmdApplyThemeColors	=	2188
visCmdApplyThemeEffects	=	2189
visCmdApplyThemeToDoc	=	2048
visCmdApplyThemeToNewShapesToggle	=	2297
visCmdApplyThemeToPage	=	2047
visCmdAppMaximize	=	1864
visCmdAppMinimize	=	1865
visCmdAppRestore	=	1904
visCmdAssociateCallout	=	2287
visCmdAutoAlign	=	2223
visCmdAutoAlignAndSpace	=	2222
visCmdAutoConnectToggle	=	2091
visCmdAutoGenerateDataGraphics	=	2105
visCmdAutoSpace	=	2224
visCmdBreakOLELink	=	1900
visCmdBrowseSampleDrawings	=	1645
visCmdBullets	=	1633
visCmdCancelInPlaceEditing	=	1602
visCmdCenterDrawing	=	1202
visCmdCheckCompatibility	=	2182
visCmdCloseInkToolsRibbonTab	=	2213
visCmdCloseWindow	=	1361
visCmdCoauthMerging	=	2385
visCmdCollapseShapesWindow	=	2269
visCmdConnectorEffectCurved	=	1945
visCmdConnectorEffectRightAngle	=	1943
visCmdConnectorEffectStraight	=	1944
visCmdConnPoints	=	1774
visCmdContainerNoHeadingToggle	=	2315
visCmdContainerAutoResizeExpandContract	=	2349
visCmdContainerAutoResizeExpandOnly	=	2348
visCmdContainerAutoResizeOff	=	2347
visCmdCreateEditMaster	=	1899
visCmdCreateNewDrawing	=	1812
visCmdCreateShortcut	=	1791
visCmdCropObject	=	1192
visCmdCropTool	=	1449
visCmdCustomPropertySets	=	1675
visCmdCustProp	=	1658
visCmdCustPropDefine	=	1695
visCmdDataAutoConnect	=	2098
visCmdDataAutoLink	=	2046
visCmdDataAutoLinkWiz	=	2045
visCmdDataColumnSettingsDlg	=	2043
visCmdDataExplorerWindow	=	2044
visCmdDataRecordsetProperties	=	2072
visCmdDataRecordsetSetCommand	=	2037
visCmdDataRecordsetSetPrimaryKey	=	2038
visCmdDataRefresh	=	2021
visCmdDataRefreshAddConflict	=	2094
visCmdDataRefreshConfigDlg	=	2022
visCmdDataRefreshDeleteConflict	=	2095
visCmdDataRefreshDlg	=	2019
visCmdDataRefreshResolveConflict	=	2103
visCmdDataSelectorDlg	=	2011
visCmdDataUnlinkRow	=	2058
visCmdDataUnlinkShape	=	2057
visCmdDecreaseIndent	=	1093
visCmdDecreaseParaSpacing	=	1095
visCmdDelConnectPt	=	1265
visCmdDeleteBackWord	=	1905
visCmdDeleteComment	=	1920
visCmdDeleteConnectors	=	2199
visCmdDeleteDataGraphic	=	2067
visCmdDeleteDataRecordset	=	1999
visCmdDeleteForwardWord	=	2114
visCmdDeleteTheme	=	2052
visCmdDeselectAll	=	1213
visCmdDesignMode	=	1388
visCmdDetectAndRepair	=	1890
visCmdDiagramFlipHorizontal	=	2229
visCmdDiagramFlipVertical	=	2228
visCmdDiagramGallery	=	1982
visCmdDiagramRotateLeft	=	2227
visCmdDiagramRotateRight	=	2226
visCmdDisbandContainer	=	2204
visCmdDistributeBottom	=	1238
visCmdDistributeCenter	=	1233
visCmdDistributeHSpace	=	1231
visCmdDistributeLeft	=	1232
visCmdDistributeMiddle	=	1237
visCmdDistributeRight	=	1234
visCmdDistributeTop	=	1236
visCmdDistributeVSpace	=	1235
visCmdDlgCustomFit	=	1536
visCmdDragDuplicate	=	1184
visCmdDrawFillStyle	=	1123
visCmdDrawGlue	=	1125
visCmdDrawAddGuide	=	1180
visCmdDrawingExplorer	=	1721
visCmdDrawingTools	=	1946
visCmdDrawLineStyle	=	1122
visCmdDrawOval	=	1183
visCmdDrawSnap	=	1124
visCmdDrawRect	=	1182
visCmdDrawRegion	=	1742
visCmdDrawTextStyle	=	1121
visCmdDrawZoom	=	1126
visCmdDRConnectionTool	=	1226
visCmdDRConnectorTool	=	1225
visCmdDRLineTool	=	1221
visCmdDropAndContain	=	2172
visCmdDropAndInsertIntoList	=	2196
visCmdDropCallout	=	2286
visCmdDropManyOnPage	=	1869
visCmdDropManyLinked	=	2108
visCmdDropOnPage	=	1246
visCmdDropOnStencil	=	1244
visCmdDropOnText	=	1243
visCmdDROvalTool	=	1224
visCmdDRPencilTool	=	1220
visCmdDRPointerTool	=	1219
visCmdDRQtrArcTool	=	1222
visCmdDRRectTool	=	1223
visCmdDRRotateTool	=	1228
visCmdDRSplineTool	=	1311
visCmdDRTextTool	=	1227
visCmdDuplicateDataGraphic	=	2106
visCmdDuplicatePage	=	2383
visCmdDuplicateTheme	=	2050
visCmdEditUndoMultiple	=	1682
visCmdDynamicGrid	=	1765
visCmdDynConnReroute	=	1829
visCmdEditConvertObject	=	1439
visCmdEditFind	=	1043
visCmdEditInsertField	=	1032
visCmdEditInsertObject	=	1031
visCmdEditLinks	=	1030
visCmdEditOpenObject	=	1029
visCmdEditPasteLink	=	1028
visCmdEditPasteSpecial	=	1027
visCmdEditRedo	=	1018
visCmdEditRedoMultiple	=	1683
visCmdEditRedoOrRepeat	=	2295
visCmdEditRepeat	=	1019
visCmdEditReplace	=	1179
visCmdEditSelectSpecial	=	1026
visCmdEditTheme	=	2049
visCmdEditThemeColors	=	2190
visCmdEditThemeEffects	=	2191
visCmdEditUndo	=	1017
visCmdEmailRouting	=	1588
visCmdExportDatabaseAddon	=	1891
visCmdFileCheckin	=	1787
visCmdFileCheckout	=	1788
visCmdFileChooseTemplates	=	1583
visCmdFileClose	=	1003
visCmdFileExit	=	1016
visCmdFileImport	=	1007
visCmdFileLastFile1	=	1012
visCmdFileLastFile10	=	2127
visCmdFileLastFile11	=	2128
visCmdFileLastFile12	=	2129
visCmdFileLastFile13	=	2130
visCmdFileLastFile14	=	2131
visCmdFileLastFile15	=	2132
visCmdFileLastFile16	=	2133
visCmdFileLastFile17	=	2134
visCmdFileLastFile18	=	2136
visCmdFileLastFile19	=	1012
visCmdFileLastFile2	=	1013
visCmdFileLastFile20	=	2137
visCmdFileLastFile3	=	1014
visCmdFileLastFile4	=	1015
visCmdFileLastFile5	=	1561
visCmdFileLastFile6	=	1569
visCmdFileLastFile7	=	1570
visCmdFileLastFile8	=	1571
visCmdFileLastFile9	=	1572
visCmdFileNew	=	1001
visCmdFileNewBlankDrawing	=	1579
visCmdFileNewBlankDrawingMetric	=	1671
visCmdFileNewBlankDrawingUS	=	1672
visCmdFileNewBlankStencil	=	1582
visCmdFileNewBlankStencilMetric	=	1673
visCmdFileNewBlankStencilUS	=	1674
visCmdFileNewStencilDlg	=	1580
visCmdFileOpen	=	1002
visCmdFileOpenStencil	=	1442
visCmdFilePrint	=	1010
visCmdFileSave	=	1004
visCmdFileSaveAs	=	1005
visCmdFileSaveAsDrawing	=	2306
visCmdFileSaveAsDrawingPreviousFileFormat	=	2298
visCmdFileSaveAsDWG	=	2305
visCmdFileSaveAsEMF	=	2302
visCmdFileSaveAsJPG	=	2301
visCmdFileSaveAsMacroDrawing	=	2311
visCmdFileSaveAsPNG	=	2300
visCmdFileSaveAsSVG	=	2303
visCmdFileSaveAsTemplate	=	2299
visCmdFileSaveAsWebPage	=	1785
visCmdFileSaveWorkspace	=	1006
visCmdFileSummaryInfoDlg	=	1009
visCmdFileUndoCheckout	=	2109
visCmdFirstTile	=	1515
visCmdFitContainerToContents	=	2195
visCmdFitCurve	=	1538
visCmdFormatAllTextProps	=	1642
visCmdFormatBehavior	=	1071
visCmdFormatBlock	=	1070
visCmdFormatCorners	=	1334
visCmdFormatCustPropDef	=	1687
visCmdFormatCustPropEdit	=	1312
visCmdFormatDefineStyles	=	1064
visCmdFormatDoubleClick	=	1118
visCmdFormatFill	=	1066
visCmdFormatInkDlg	=	1955
visCmdFormatLine	=	1065
visCmdFormatPainter	=	1271
visCmdFormatParagraph	=	1068
visCmdFormatPictureAutobalance	=	2205
visCmdFormatPictureCompressionDlg	=	2212
visCmdFormatProtection	=	1072
visCmdFormatShadow	=	1333
visCmdFormatSpecial	=	1073
visCmdFormatStyle	=	1063
visCmdFormatTabs	=	1069
visCmdFormatText	=	1067
visCmdFullScreenMode	=	1492
visCmdGoToPageToolbar	=	1635
visCmdGrid	=	1767
visCmdGuides	=	1771
visCmdHeaderFooter	=	1720
visCmdHelpAboutVisio	=	1100
visCmdHelpContents	=	1092
visCmdHelpMode	=	1386
visCmdHelpSearch	=	1809
visCmdHelpShapeBasics	=	1822
visCmdHelpTemplates	=	1586
visCmdHideAllToolbars	=	1726
visCmdHideDocumentStencil	=	1689
visCmdHideMoreShapes	=	2291
visCmdHyperlinkHier	=	1611
visCmdHyperlinkList	=	1719
visCmdIconBucketTool	=	1543
visCmdIconLassoTool	=	1544
visCmdIconLeftColor	=	1143
visCmdIconPencilTool	=	1145
visCmdIconRightTool	=	1144
visCmdIconSelectNet	=	1545
visCmdIgnoreValidationIssue	=	2254
visCmdIgnoreValidationRule	=	2256
visCmdImageProperties	=	1887
visCmdImagePropertiesDlg	=	1883
visCmdIncreaseIndent	=	1094
visCmdIncreaseParaSpacing	=	1096
visCmdINETAddToFavorites	=	1506
visCmdINETCopyHyperlink	=	1610
visCmdINETDeleteHlink	=	1609
visCmdINETDiagrammingResources	=	1606
visCmdINETEditHyperlink	=	1619
visCmdINETGoBack	=	1599
visCmdINETGoForward	=	1598
visCmdINETKnowledgeBase	=	1605
visCmdINETOpenHLink	=	1607
visCmdINETOpenHLinkNewWnd	=	1608
visCmdINETPasteAsHyperlink	=	1620
visCmdINETUserSearchPage	=	1595
visCmdINETVisioHomePage	=	1596
visCmdINETVisioOnTheWeb	=	1831
visCmdINETVisioSolutionsLibrary	=	1604
visCmdInkEraser	=	1970
visCmdInkReviewPen	=	1971
visCmdInkStockPen0	=	1973
visCmdInkStockPen1	=	1974
visCmdInkTool	=	1661
visCmdInsertAutoCADAddOn	=	1521
visCmdInsertCheckBoxControl	=	2150
visCmdInsertClipArt	=	1497
visCmdInsertClipArtDlg	=	2345
visCmdInsertComboBoxControl	=	2152
visCmdInsertComment	=	1501
visCmdInsertControlDlg	=	1522
visCmdInsertDataMap	=	1282
visCmdInserTextBoxControl	=	2145
visCmdInsertHyperLink	=	1585
visCmdInsertImageControl	=	2148
visCmdInsertLabelControl	=	2144
visCmdInsertLegendHorizontal	=	2331
visCmdInsertLegendVertical1	=	2335
visCmdInsertListBoxControl	=	2153
visCmdInsertMemberIntoList	=	2174
visCmdInsertMicrosoftGraph	=	1499
visCmdInsertNewBackgroundPage	=	2165
visCmdInsertPageTab	=	2202
visCmdInsertPushButtonControl	=	2147
visCmdInsertRadioButtonControl	=	2151
visCmdInsertScrollBarControl	=	2149
visCmdInsertSpinControl	=	2146
visCmdInsertTextBox	=	2006
visCmdInsertToggleButtonControl	=	2154
visCmdInsertVertTextBox	=	2007
visCmdInsertWordArt	=	1498
visCmdIntersect	=	1453
visCmdJoin	=	1533
visCmdLanguagePreferencesDlg	=	2363
visCmdLast	=	65535
visCmdLastTile	=	1516
visCmdLayerDlg	=	1446
visCmdLayerSetupDlg	=	1448
visCmdLayoutDynamic	=	1493
visCmdLinkRowToShape	=	1997
visCmdListInsertAfter	=	2271
visCmdListInsertBefore	=	2270
visCmdLockContainer	=	2220
visCmdMasterExplorer	=	1916
visCmdMasterSetup	=	1343
visCmdMDIMaximize	=	1901
visCmdMDIMinimize	=	1902
visCmdMDIRestore	=	1903
visCmdMinimizeRibbonToggle	=	2232
visCmdModConnectPt	=	1264
visCmdModControlPt	=	1267
visCmdMovConnectPt	=	1269
visCmdMove1D	=	1186
visCmdMove2D	=	1187
visCmdMoveComment	=	1502
visCmdMoveObject	=	1185
visCmdMsoAutoCorrect	=	1872
visCmdMsoAutoCorrectDlg	=	1866
visCmdMsoAutoFormat	=	1873
visCmdMsoCustomItem	=	1896
visCmdMSOInsertEquation	=	1646
visCmdMSOInsertSymbol	=	1504
visCmdMSOInsertSymbolDlg	=	1505
visCmdMsoMediaGallery	=	1885
visCmdMultipleFileImport	=	2201
visCmdNewDefDocBlankDrawing	=	1906
visCmdNewForegroundPage	=	2361
visCmdNewFromExisting	=	1897
visCmdNewThemeColors	=	2065
visCmdNewThemeEffects	=	2064
visCmdNextCommentMarkup	=	2180
visCmdNextMarkup	=	1914
visCmdNextTile	=	1514
visCmdNextWindow	=	1886
visCmdObjectAddToGroup	=	1053
visCmdObjectAlignObjects	=	1049
visCmdObjectBringForward	=	1045
visCmdObjectBringToFront	=	1046
visCmdObjectCombine	=	1061
visCmdObjectConnectObjects	=	1050
visCmdObjectConvertToGroup	=	1055
visCmdObjectDistributeDlg	=	1230
visCmdObjectFlipHorizontal	=	1058
visCmdObjectFlipVertical	=	1057
visCmdObjectFragment	=	1062
visCmdObjectGroup	=	1051
visCmdObjectHelp	=	1428
visCmdObjectInfoDlg	=	1425
visCmdObjectRemoveFromGroup	=	1054
visCmdObjectReverse	=	1059
visCmdObjectRotate90	=	1056
visCmdObjectSendBackward	=	1047
visCmdObjectSendToBack	=	1048
visCmdObjectSwapEnds	=	1870
visCmdObjectUngroup	=	1052
visCmdObjectUnion	=	1060
visCmdOfficeCenterOptions	=	2141
visCmdOffsetDlg	=	1387
visCmdOfficeDiagnostics	=	1890
visCmdOpenActiveObject	=	1601
visCmdOpenCommentForEdit	=	1503
visCmdOpenInVisio	=	1491
visCmdOptionsColorPaletteDlg	=	1082
visCmdOptionsDeletePages	=	1079
visCmdOptionsEditBackground	=	1075
visCmdOptionsEditDrawing	=	1074
visCmdOptionsGoToDrawing	=	1077
visCmdOptionsNewPage	=	1078
visCmdOptionsPageSetup	=	1076
visCmdOptionsPreferences	=	1081
visCmdOptionsProtectDocument	=	1083
visCmdOptionsReorderPages	=	1080
visCmdOptionsSnapGlueSetup	=	1084
visCmdPageAutoSizeToggle	=	2333
visCmdPageMeasureUnitsDlg	=	1274
visCmdPageSizeDlg	=	2176
visCmdPageSizeToFitDrawing	=	2332
visCmdPagesList	=	1654
visCmdPanObject	=	1193
visCmdPanZoom	=	1653
visCmdPasteShortcut	=	1790
visCmdPasteToLocation	=	2221
visCmdPauseRecordingMacro	=	1778
visCmdPreviousCommentMarkup	=	2179
visCmdPreviousMarkup	=	1915
visCmdPreviousTile	=	1513
visCmdPrintPage	=	1443
visCmdPrintPreview	=	1490
visCmdProgRefHelp	=	1584
visCmdPublishToProcessRepository	=	2294
visCmdPublishToVisioServices	=	2293
visCmdRecalcObjectWH	=	1146
visCmdRecordNewMacro	=	1775
visCmdReinstateValidationIssue	=	2255
visCmdRelayoutShapes	=	2068
visCmdRemoveDataGraphicFromSel	=	2107
visCmdRemoveFromAllContainers	=	2252
visCmdRemoveMemberFromContainer	=	2175
visCmdRemoveMemberFromList	=	2346
visCmdRemoveThemeFromSel	=	2119
visCmdRemoveVBAFromActiveDoc	=	1590
visCmdReorderList	=	2197
visCmdReplaceShape	=	2051
visCmdRHI	=	2009
visCmdRHIDlg	=	2010
visCmdReOrderPage	=	1795
visCmdResearchLookUp	=	1967
visCmdResearchThesaurus	=	2178
visCmdResearchTranslate	=	1968
visCmdResumeRecordingMacro	=	1779
visCmdReviewerVisibilityAll	=	1836
visCmdReviewerVisibilityNone	=	1919
visCmdRightDragCancel	=	1881
visCmdRightDragCopy	=	1879
visCmdRightDragLink	=	1880
visCmdRightDragMove	=	1878
visCmdRotate90Clockwise	=	1494
visCmdRotateObject	=	1190
visCmdRulerGridDlg	=	1318
visCmdRulSub	=	1766
visCmdRunAddOnDlg	=	1484
visCmdRunAddonMenu	=	1090
visCmdSaveAsFixedFormatDlg	=	2117
visCmdSaveForAutoRecover	=	1857
visCmdSelectContainerMembers	=	2219
visCmdSelectionModeExtend	=	1909
visCmdSelectionModeLasso	=	1908
visCmdSelectionModeRect	=	1907
visCmdSendAsMail	=	1292
visCmdSetAddMarkup	=	1744
visCmdSetCharColor	=	1404
visCmdSetCharSizeDown	=	1406
visCmdSetCharSizeUp	=	1405
visCmdSetContainerProperties	=	2181
visCmdSetDynConnAppearanceCurved	=	1895
visCmdSetDynConnAppearanceDefault	=	1893
visCmdSetDynConnAppearanceStraight	=	1894
visCmdSetDynConnLineJumpStyle_2pt	=	1713
visCmdSetDynConnLineJumpStyle_3pt	=	1714
visCmdSetDynConnLineJumpStyle_4pt	=	1715
visCmdSetDynConnLineJumpStyle_5pt	=	1716
visCmdSetDynConnLineJumpStyle_6pt	=	1717
visCmdSetDynConnLineJumpStyle_Arc	=	1709
visCmdSetDynConnLineJumpStyle_Gap	=	1710
visCmdSetDynConnLineJumpStyle_Page	=	1708
visCmdSetDynConnLineJumpStyle_Square	=	1711
visCmdSetDynConnLineJumpStyle_Triangle	=	1712
visCmdSetDynConnRerouteAsNeeded	=	1697
visCmdSetDynConnRerouteFreely	=	1696
visCmdSetDynConnRerouteNever	=	1698
visCmdSetDynConnRerouteOnCrossover	=	1837
visCmdSetDynConnRoutingStyle	=	1700
visCmdSetFillColor	=	1385
visCmdSetFillPattern	=	1399
visCmdSetFillShadow	=	1379
visCmdSetHeaderFooter	=	1858
visCmdSetIndexInStencil	=	1871
visCmdSetLanguageDlg	=	1888
visCmdSetLineColor	=	1359
visCmdSetLineCornerStyle	=	1358
visCmdSetLineEnds	=	1357
visCmdSetLinePattern	=	1356
visCmdSetLineWeight	=	1355
visCmdSetPageLineJumpCode_D	=	1703
visCmdSetPageLineJumpCode_Horz	=	1705
visCmdSetPageLineJumpCode_Last	=	1707
visCmdSetPageLineJumpCode_None	=	1704
visCmdSetPageLineJumpCode_Vert	=	1706
visCmdSetPageOrientation	=	2170
visCmdSetPagePlow	=	1699
visCmdSetPageSize	=	2171
visCmdSetThemeBehavior	=	2382
visCmdShapeActions	=	1309
visCmdShapeComment	=	1686
visCmdShapeCommentDelete	=	1688
visCmdShapeCommentDlg	=	1685
visCmdShapeExplorer	=	1389
visCmdShapeGalleryAddOn	=	1867
visCmdShapeGeo	=	1769
visCmdShapeHand	=	1772
visCmdShapeIntersect	=	1830
visCmdShapeLayerToolbar	=	1634
visCmdShapeSearchWindowToggle	=	2344
visCmdShapeStudioAddOn	=	1985
visCmdShapesWindow	=	1669
visCmdShapeTransparency	=	1875
visCmdShapeTransparencyDlg	=	1874
visCmdShapeVert	=	1773
visCmdShowIgnoredIssuesToggle	=	2258
visCmdShowLineJumpsToggle	=	2231
visCmdShowShapeSheetDocument	=	2169
visCmdShowShapeSheetPage	=	2168
visCmdShowShapeSheetShape	=	2167
visCmdSize1D	=	1188
visCmdSize2D	=	1189
visCmdSizeObjects	=	1925
visCmdSizePos	=	1670
visCmdSizeTextBlock	=	1194
visCmdSmartAlign	=	2223
visCmdSmartAlignAndSpace	=	2222
visCmdSmartSpace	=	2224
visCmdSpaceShapesAvoidPageBreaksToggle	=	2340
visCmdSpellingChange	=	1889
visCmdSpellingOptionsDlg	=	2042
visCmdSSWindowAddSection	=	1384
visCmdSSWindowChangeRowType	=	1383
visCmdSSWindowCollapse	=	1250
visCmdSSWindowDeselect	=	1253
visCmdSSWindowExpand	=	1251
visCmdSSWindowPasteFunction	=	1382
visCmdSSWindowPasteName	=	1381
visCmdSSWindowSelect	=	1252
visCmdSSWindowShowSection	=	1380
visCmdSSWindowShowTraceWindow	=	1781
visCmdStampTool	=	1424
visCmdStartRecordingMacro	=	1776
visCmdStenActivate	=	1458
visCmdStenAutoArrange	=	1483
visCmdStenCleanup	=	1106
visCmdStenClose	=	1452
visCmdStenDrawingExplorer	=	1796
visCmdStenEditDrawing	=	1102
visCmdStenEditIcon	=	1101
visCmdStenEditOff	=	1681
visCmdStenEditOn	=	1680
visCmdStenEditToggle	=	1679
visCmdStenIconAndDetail	=	1892
visCmdStenIconAndName	=	1480
visCmdStenIconOnly	=	1481
visCmdStenImageMaster	=	1105
visCmdStenNameMaster	=	1103
visCmdStenNameOnly	=	1482
visCmdStenNamesUnderIcons	=	2005
visCmdStenNewMaster	=	1104
visCmdStenProperties	=	1678
visCmdStenSave	=	1676
visCmdStenSaveAs	=	1677
visCmdStopIgnoringValidationRule	=	2257
visCmdStopRecordingMacro	=	1777
visCmdSubtract	=	1454
visCmdSWAccept	=	1140
visCmdSWAddSectionDlg	=	1116
visCmdSWCancel	=	1139
visCmdSWChangeRowTypeDlg	=	1114
visCmdSWDeleteRow	=	1115
visCmdSWDeleteSection	=	1117
visCmdSWExpandRow	=	1718
visCmdSWFormula	=	1141
visCmdSWInsertRow	=	1112
visCmdSWInsertRowAfter	=	1113
visCmdSWPasteFunctionDlg	=	1111
visCmdSWPasteNameDlg	=	1110
visCmdSWShapeActionDlg	=	1444
visCmdSWShowFormulas	=	1108
visCmdSWShowSectionsDlg	=	1109
visCmdSWShowToggle	=	1142
visCmdSWShowValues	=	1107
visCmdTaskPane	=	1896
visCmdTaskPaneDataGraphic	=	2024
visCmdTaskPaneDocumentManagement	=	1972
visCmdTaskPaneResearch	=	1969
visCmdTaskPaneReviewer	=	1939
visCmdTaskTogglePreviewSize	=	2053
visCmdTextAllCaps	=	1862
visCmdTextBlockTool	=	1451
visCmdTextBold	=	1131
visCmdTextDoubleStrikeThrough	=	1951
visCmdTextDoubleULine	=	1863
visCmdTextEditRuler	=	1810
visCmdTextEditState	=	1214
visCmdTextFont	=	1129
visCmdTextHAlignCenter	=	1408
visCmdTextHAlignDistribute	=	1952
visCmdTextHAlignJustify	=	1412
visCmdTextHAlignLeft	=	1407
visCmdTextHAlignRight	=	1409
visCmdTextItalic	=	1132
visCmdTextRotate90	=	1098
visCmdTextSize	=	1130
visCmdTextSmallCaps	=	1133
visCmdTextStrikeThrough	=	1741
visCmdTextStyle	=	1128
visCmdTextSubscript	=	1135
visCmdTextSuperScript	=	1134
visCmdTextULine	=	1136
visCmdTextVAlignBottom	=	1422
visCmdTextVAlignMiddle	=	1414
visCmdTextVAlignTop	=	1413
visCmdToggleDocumentStencil	=	1690
visCmdToolbarsDlg	=	1500
visCmdToolsArrayShapesAddOn	=	1354
visCmdToolsInventory	=	1335
visCmdToolsLayoutShapesDlg	=	1574
visCmdToolsMacroDlg	=	1577
visCmdToolSnapLines	=	1807
visCmdToolsRunVBE	=	1576
visCmdToolsSpelling	=	1270
visCmdTranslateOptions	=	2352
visCmdTrim	=	1534
visCmdTrustCenterDlg	=	2104
visCmdTurnToNextPage	=	1148
visCmdTurnToPrevPage	=	1147
visCmdUFEditClear	=	1023
visCmdUFEditCopy	=	1021
visCmdUFEditCut	=	1020
visCmdUFEditDuplicate	=	1024
visCmdUFEditPaste	=	1022
visCmdUFEditSelectAll	=	1025
visCmdUpgradeThemeModel	=	2384
visCmdUpdateColumnsInLinkedShapes	=	2061
visCmdUpdateContentCache	=	1241
visCmdValidateDiagram	=	2253
visCmdValidationIssueNavigateToShape	=	2236
visCmdValidationIssuesArrangeByCategory	=	2279
visCmdValidationIssuesArrangeByIgnored	=	2281
visCmdValidationIssuesArrangeByPage	=	2280
visCmdValidationIssuesArrangeByRule	=	2278
visCmdValidationIssuesArrangeOriginalOrder	=	2282
visCmdValidationIssuesWindowToggle	=	2263
visCmdView100	=	1035
visCmdView150	=	1036
visCmdView200	=	1037
visCmdView400	=	1280
visCmdView50	=	1279
visCmdView75	=	1034
visCmdViewConnections	=	1042
visCmdViewCustom	=	1038
visCmdViewDirectionToggle	=	2012
visCmdViewFitInWindow	=	1033
visCmdViewGrid	=	1040
visCmdViewGuides	=	1041
visCmdViewLeftToRight	=	2013
visCmdViewPageBreaks	=	1509
visCmdViewRightToLeft	=	2014
visCmdViewRulers	=	1039
visCmdViewStatusBar	=	1044
visCmdWindowCascadeAll	=	1086
visCmdWindowNewWindow	=	1085
visCmdWindowShowDrawPage	=	1091
visCmdWindowShowMasterObjects	=	1089
visCmdWindowShowShapeSheet	=	1088
visCmdWindowTileAll	=	1087
visCmdZoomArea	=	1218
visCmdZoomIn	=	1216
visCmdZoomInIgnoreSel	=	1917
visCmdZoomLast	=	1495
visCmdZoomOut	=	1217
visCmdZoomOutIgnoreSel	=	1918
visCmdZoomPageWidth	=	1496
visCmdZoomPt	=	1215
visCmdZoomSingleTile	=	1512

#values from visio.visuictrltypes
#****************************
visCtrlTypeBUTTON_OWNERDRAW	=	33		# Owner-draw push button.
visCtrlTypeBUTTON	=	2		# Push button.
visCtrlTypeCOMBOBOX	=	128		# Combo box.
visCtrlTypeDROPDOWN	=	272		# Drop-down combo box.
visCtrlTypeEDITBOX	=	64		# Text box.
visCtrlTypeLABEL	=	2048		# Label.
visCtrlTypeSPLITBUTTON_MRU_COLOR	=	16		# Split button, with MRU color behavior.
visCtrlTypeSPLITBUTTON_MRU_COMMAND	=	18		# Split button, with MRU command behavior.
visCtrlTypeSPLITBUTTON	=	17		# Split button.

#values from visio.visuiiconids
#****************************
visIconIXACCEPT	=	85
visIconIXADDIN	=	149
visIconIXALIGNBOTTOM	=	69
visIconIXALIGNBOX	=	224
visIconIXALIGNCENTER	=	65
visIconIXALIGNLEFT	=	64
visIconIXALIGNMIDDLE	=	68
visIconIXALIGNRIGHT	=	66
visIconIXALIGNTOP	=	67
visIconIXALIGN	=	63
visIconIXARRANGE	=	83
visIconIXBOLD	=	50
visIconIXBRING_FORWARD	=	245
visIconIXBRINGFRONT	=	90
visIconIXBULLETS	=	113
visIconIXCANCEL	=	84
visIconIXCANTFIND	=	129
visIconIXCASCADE	=	94
visIconIXCHART	=	134
visIconIXCHECKMARK	=	128
visIconIXCLEAR	=	9
visIconIXCLIPART	=	130
visIconIXCLOSE	=	143
visIconIXCONNECTIONPOINTS	=	36
visIconIXCONNECTIONPTTOOL	=	30
visIconIXCONNECTORTOOL	=	96
visIconIXCONNECTSHAPES	=	75
visIconIXCONNPOINTS	=	229
visIconIXCOPY	=	7
visIconIXCORNERSTYLE	=	40
visIconIXCROP	=	29
visIconIXCUSTOM_BALLOON	=	162
visIconIXCUSTOM_BANK	=	153
visIconIXCUSTOM_BELL	=	159
visIconIXCUSTOM_BOX	=	173
visIconIXCUSTOM_CALC	=	164
visIconIXCUSTOM_CAMCORD	=	163
visIconIXCUSTOM_CARDS	=	169
visIconIXCUSTOM_CLUB	=	168
visIconIXCUSTOM_DIAMOND	=	166
visIconIXCUSTOM_DOWN	=	178
visIconIXCUSTOM_EIGHTBALL	=	191
visIconIXCUSTOM_EYE	=	190
visIconIXCUSTOM_FEET	=	174
visIconIXCUSTOM_FISH	=	182
visIconIXCUSTOM_FROWN	=	152
visIconIXCUSTOM_GEARS	=	184
visIconIXCUSTOM_HEART	=	165
visIconIXCUSTOM_HOURGLASS	=	186
visIconIXCUSTOM_KEYBOARD	=	180
visIconIXCUSTOM_KEY	=	183
visIconIXCUSTOM_LEFT	=	175
visIconIXCUSTOM_LOAD	=	155
visIconIXCUSTOM_MAN	=	187
visIconIXCUSTOM_MIC	=	157
visIconIXCUSTOM_MUG	=	170
visIconIXCUSTOM_NOTE	=	160
visIconIXCUSTOM_PAGES	=	181
visIconIXCUSTOM_PASTE	=	154
visIconIXCUSTOM_PENCIL	=	172
visIconIXCUSTOM_PHONE	=	161
visIconIXCUSTOM_QUESTION	=	192
visIconIXCUSTOM_RIGHT	=	176
visIconIXCUSTOM_RUN	=	189
visIconIXCUSTOM_SAVE	=	156
visIconIXCUSTOM_SCALES	=	185
visIconIXCUSTOM_SMILE	=	151
visIconIXCUSTOM_SPADE	=	167
visIconIXCUSTOM_SPEAKER	=	158
visIconIXCUSTOM_TACK	=	179
visIconIXCUSTOM_TRASH	=	171
visIconIXCUSTOM_UP	=	177
visIconIXCUSTOM_WOMAN	=	188
visIconIXCUSTOMPROP_WINDOW	=	215
visIconIXCUSTPROP	=	111
visIconIXCUT	=	6
visIconIXDCREROUTE_ASNEEDED	=	214
visIconIXDCREROUTE_FREELY	=	213
visIconIXDCREROUTE_NEVER	=	235
visIconIXDCREROUTE	=	236
visIconIXDECRINDENT	=	114
visIconIXDECRPARA	=	116
visIconIXDELETECOMMENT	=	195
visIconIXDELETE	=	196
visIconIXDESIGNMODE	=	119
visIconIXDHORZ_CENTER	=	72
visIconIXDHORZ_EQSPACE	=	71
visIconIXDISTRIBUTE	=	70
visIconIXDOUBLE_UNDERLINE	=	253
visIconIXDRAWINGEXPLORER	=	219
visIconIXDVERT_EQSPACE	=	73
visIconIXDVERT_MIDDLE	=	74
visIconIXDYNGRID	=	221
visIconIXEDITCOMMENT	=	194
visIconIXEDITSTEN	=	197
visIconIXEXCHANGEFOLDER	=	137
visIconIXFILLCOLOR	=	43
visIconIXFILLPATTERN	=	47
visIconIXFIND	=	138
visIconIXFIRSTPAGE	=	76
visIconIXFLIPHORIZONTAL	=	18
visIconIXFLIPVERTICAL	=	19
visIconIXFOLDER	=	144
visIconIXFORMATPAINTER	=	101
visIconIXFULLSCREEN	=	124
visIconIXGLUE	=	32
visIconIXGOBACK	=	107
visIconIXGOFORWARD	=	108
visIconIXGRID	=	34
visIconIXGROUP	=	92
visIconIXGUIDES	=	226
visIconIXGUIDE	=	35
visIconIXHELPASSISTANT	=	133
visIconIXHELPBOOK	=	125
visIconIXHELPMODE	=	102
visIconIXICONBUCKET	=	87
visIconIXICONLASSO	=	88
visIconIXICONNAME	=	80
visIconIXICONONLY	=	81
visIconIXICONPENCIL	=	86
visIconIXICONSELNET	=	89
visIconIXIMAGE	=	131
visIconIXINCRINDENT	=	115
visIconIXINCRPARA	=	117
visIconIXINSERT_EQUATION	=	250
visIconIXINSERT_OBJECT	=	248
visIconIXINSERTCOMMENT	=	193
visIconIXINSERTCONTROL	=	118
visIconIXINSERTHYPERLINK	=	105
visIconIXITALIC	=	51
visIconIXLARGE_PADLOCK	=	249
visIconIXLASTPAGE	=	77
visIconIXLAYERPROPERTIES	=	103
visIconIXLAYOUTSHAPES	=	104
visIconIXLINECOLOR	=	44
visIconIXLINEEND	=	41
visIconIXLINEJUMPSTYLE_2PT	=	208
visIconIXLINEJUMPSTYLE_3PT	=	209
visIconIXLINEJUMPSTYLE_4PT	=	210
visIconIXLINEJUMPSTYLE_5PT	=	211
visIconIXLINEJUMPSTYLE_6PT	=	212
visIconIXLINEJUMPSTYLE_ARC	=	204
visIconIXLINEJUMPSTYLE_GAP	=	205
visIconIXLINEJUMPSTYLE_PAGE	=	218
visIconIXLINEJUMPSTYLE_SQUARE	=	206
visIconIXLINEJUMPSTYLE_TRIANGLE	=	207
visIconIXLINEPATTERN	=	46
visIconIXLINETOOL	=	22
visIconIXLINEWEIGHT	=	45
visIconIXMACROS	=	121
visIconIXMAILRECPT	=	135
visIconIXMAXIMIZE	=	142
visIconIXMINIMIZE	=	141
visIconIXNAMEONLY	=	82
visIconIXNEWSTEN	=	198
visIconIXNEWWINDOW	=	39
visIconIXNEW	=	0
visIconIXNEXTPAGE	=	14
visIconIXOPENSTENCIL	=	2
visIconIXOPEN	=	1
visIconIXOVALTOOL	=	25
visIconIXPAGEBREAKS	=	78
visIconIXPAGELINEJUMPCODE_DISP	=	217
visIconIXPAGELINEJUMPCODE_HORZ	=	201
visIconIXPAGELINEJUMPCODE_LASTROUTED	=	203
visIconIXPAGELINEJUMPCODE_NONE	=	200
visIconIXPAGELINEJUMPCODE_RDISP	=	231
visIconIXPAGELINEJUMPCODE_VERT	=	202
visIconIXPAGEPLOW	=	216
visIconIXPANZOOM	=	139
visIconIXPASTE	=	8
visIconIXPENCILTOOL	=	21
visIconIXPOINTERTOOL	=	20
visIconIXPOINTSIZEDOWN	=	48
visIconIXPOINTSIZEUP	=	49
visIconIXPREVIOUSPAGE	=	13
visIconIXPRINTPREVIEW	=	5
visIconIXPRINT	=	4
visIconIXQTRARCTOOL	=	23
visIconIXRECTANGLETOOL	=	24
visIconIXREDO	=	11
visIconIXREPEAT	=	12
visIconIXREPLACE	=	252
visIconIXRESTORE	=	140
visIconIXROTATECLOCKWISE	=	37
visIconIXROTATECOUNTERCLOCKWISE	=	38
visIconIXROTATETEXT	=	112
visIconIXROTATETOOL	=	28
visIconIXROUTINGRECPT	=	136
visIconIXRULER	=	33
visIconIXRULSUB	=	222
visIconIXSAVE	=	3
visIconIXSEARCHTHEWEB	=	106
visIconIXSEND_BACKWARD	=	246
visIconIXSENDBACK	=	91
visIconIXSHADOWSTYLE	=	42
visIconIXSHAPE_INTERSECT	=	220
visIconIXSHAPEEXPLORER	=	126
visIconIXSHAPEEXPL	=	110
visIconIXSHAPEEXT	=	230
visIconIXSHAPEGEO	=	225
visIconIXSHAPEHAND	=	227
visIconIXSHAPESHEET	=	120
visIconIXSHAPEVERT	=	228
visIconIXSHOWDOCSTEN	=	199
visIconIXSINGLETILE	=	99
visIconIXSIZEPOS	=	150
visIconIXSMALL_PADLOCK	=	247
visIconIXSMALLCAPS	=	234
visIconIXSNAP_LINES	=	232
visIconIXSNAPTOGRID	=	223
visIconIXSNAP	=	31
visIconIXSPELLING	=	100
visIconIXSPLINETOOL	=	79
visIconIXSTAMPTOOL	=	26
visIconIXSTRIKETHROUGH	=	233
visIconIXSTYLE	=	251
visIconIXSUBSCRIPT	=	54
visIconIXSUPERSCRIPT	=	53
visIconIXTEXTALIGNBOTTOM	=	62
visIconIXTEXTALIGNCENTER	=	57
visIconIXTEXTALIGNJUSTIFY	=	59
visIconIXTEXTALIGNLEFT	=	56
visIconIXTEXTALIGNMIDDLE	=	61
visIconIXTEXTALIGNRIGHT	=	58
visIconIXTEXTALIGNTOP	=	60
visIconIXTEXTBLOCKTOOL	=	97
visIconIXTEXTCOLOR	=	55
visIconIXTEXTOOL	=	27
visIconIXTILE	=	95
visIconIXUNDERLINE	=	52
visIconIXUNDO	=	10
visIconIXUNGROUP	=	93
visIconIXVBAMACRO	=	148
visIconIXVBEDITOR	=	122
visIconIXVERTICALTEXT	=	123
visIconIXVSD	=	145
visIconIXVSS	=	146
visIconIXVST	=	147
visIconIXWEBPAGE	=	127
visIconIXWEBTOOLBAR	=	109
visIconIXWHOLEPAGE	=	98
visIconIXWORDART	=	132
visIconIXZOOM100	=	17
visIconIXZOOMIN	=	16
visIconIXZOOMOUT	=	15

#values from visio.visuimenuanimation
#****************************
visMenuAnimationNone	=	0		# Animations appear immediately.
visMenuAnimationRandom	=	1		# Animations unfold or slide randomly.
visMenuAnimationSlide	=	3		# Animations appear to slide into view from above.
visMenuAnimationUnfold	=	2		# Animations appear to expand from a point in the upper-left corner of the animation.

#values from visio.visuiobjsets
#****************************
visUIObjSetActiveXDoc	=	18		# Visio is running as an ActiveX document.

#values from visio.visuniqueidargs
#****************************
visDeleteGUID	=	2		# Clear the unique ID of a shape and return a zero-length string (&quot;&quot;).
visDeleteGUIDWithUndo	=	4		# Clear the unique ID of a shape and return a zero-length string (&quot;&quot;). Undoable.
visGetGUID	=	0		# Return the unique ID string only if the shape already has a unique ID.
visGetOrMakeGUID	=	1		# Return the unique ID string of the shape. If the shape does not already have a unique ID, assign one to the shape and return the new ID.
visGetOrMakeGUIDWithUndo	=	3		# Return the unique ID string of the shape. If the shape does not already have a unique ID, assign one to the shape and return the new ID. Undoable.

#values from visio.visunitcodes
#****************************
visAcre	=	36		# Acre
visAngleUnits	=	80		# Angle units
visCentimeters	=	69		# Centimeters
visCicerosAndDidots	=	52		# Ciceros and didots
visCiceros	=	54		# Ciceros
visCurrency	=	111		# Currency
visDate	=	40		# Date
visDegreeMinSec	=	82		# Degrees, minutes, and seconds
visDegrees	=	81		# Degrees
visDidots	=	53		# Didots
visDrawingUnits	=	64		# Drawing units
visDurationUnits	=	42		# Duration
visElapsedDay	=	44		# Elapsed days
visElapsedHour	=	45		# Elapsed hours
visElapsedMin	=	46		# Elapsed minutes
visElapsedSec	=	47		# Elapsed seconds
visElapsedWeek	=	43		# Elapsed weeks
visFeetAndInches	=	67		# Feet and inches
visFeet	=	66		# Feet
visHectare	=	37		# Hectares
visInches	=	65		# Inches
visInchFrac	=	73		# Fractions of inches
visKilometers	=	72		# Kilometers
visMeters	=	71		# Meters
visMileFrac	=	74		# Fractions of miles
visMiles	=	68		# Miles
visMillimeters	=	70		# Millimeters
visMin	=	84		# Minutes
visNautMiles	=	76		# Nautical miles
visNoCast	=	252		# No unit conversion
visNumber	=	32		# Number
visPageUnits	=	63		# Page units
visPercent	=	33		# Percent
visPicasAndPoints	=	49		# Picas and points
visPicas	=	51		# Picas
visPoints	=	50		# Points
visRadians	=	83		# Radians
visSec	=	85		# Seconds
visTypeUnits	=	48		# Type units
visUnitsColor	=	251		# Color units
visUnitsGUID	=	95		# GUID units
visUnitsInval	=	255		# Invalid units
visUnitsNURBS	=	138		# NURBS units
visUnitsPoint	=	225		# Point units
visUnitsPolyline	=	139		# Polyline units
visUnitsString	=	231		# String units
visYards	=	75		# Yards

#values from visio.visvalidationflags
#****************************
visValidationDefault	=	0		# Validate document, and if validation issues are found, open the  Issues window.
visValidationNoOpenWindow	=	1		# Validate document, but do not open the  Issues window.

#values from visio.visverticalaligntypes
#****************************
visVertAlignBottom	=	3		# Align to bottom of primary selected shape.
visVertAlignMiddle	=	2		# Align to middle of primary selected shape.
visVertAlignNone	=	0		# No vertical alignment.
visVertAlignTop	=	1		# Align to top of primary selected shape.

#values from visio.viswindowarrange
#****************************
visArrangeCascade	=	3		# Cascade the windows.
visArrangeTileHorizontal	=	2		# Tile the windows horizontally.
visArrangeTileVertical	=	1		# Tile the windows vertically.

#values from visio.viswindowfit
#****************************
visFitNone	=	0		# No auto-fit.
visFitPage	=	1		# Fit whole page.
visFitWidth	=	2		# Fit to page width.

#values from visio.viswindowscrollx
#****************************
visScrollLeftPage	=	2		# Scroll horizontally so that the left edge of the drawing page is centered in the window.
visScrollLeft	=	0		# Scroll horizontally to the left the same distance as clicking the left scroll button.
visScrollNoneX	=	9		# Do not scroll horizontally.
visScrollRightPage	=	3		# Scroll horizontally so that the right edge of the drawing page is centered in the window.
visScrollRight	=	1		# Scroll horizontally to the right the same distance as clicking the right scroll button.
visScrollToLeft	=	6		# Scroll so that the upper-left corner of the drawing page is centered in the window.
visScrollToRight	=	7		# Scroll so that the lower-right corner of the drawing page is centered in the window.

#values from visio.viswindowscrolly
#****************************
visScrollDownPage	=	3		# Scroll vertically so that the upper edge of the drawing page is centered in the window.
visScrollDown	=	1		# Scroll vertically down the same distance as clicking the down scroll button.
visScrollNoneY	=	9		# Do not scroll vertically.
visScrollToBottom	=	7		# Scroll so that the lower-right corner of the drawing page is centered in the window.
visScrollToTop	=	6		# Scroll so that the upper-left corner of the drawing page is centered in the window.
visScrollUpPage	=	2		# Scroll vertically so that the lower edge of the drawing page is centered in the window.
visScrollUp	=	0		# Scroll vertically up the same distance as clicking the up scroll button.

#values from visio.viswindowstates
#****************************
visWSActive	=	0x4000000		# Active window.
visWSAnchorAutoHide	=	0x200		# Anchor window with AutoHide on.
visWSAnchorBottom	=	0x100		# Window is anchored at the bottom.
visWSAnchorLeft	=	0x20		# Window is anchored at the left.
visWSAnchorMerged	=	0x400		# Window is merged.
visWSAnchorRight	=	0x80		# Window is anchored at the right.
visWSAnchorTop	=	0x40		# Window is anchored at the top.
visWSDockedBottom	=	0x8		# Window is docked at the bottom. Not used for the Shapes window in Visio
visWSDockedLeft	=	0x1		# Window is docked at the left.
visWSDockedRight	=	0x4		# Window is docked at the right.
visWSDockedTop	=	0x2		# Window is docked at the top. Not used for the Shapes window in Visio.
visWSFloating	=	0x10		# Window is floating.
visWSMaximized	=	0x40000000		# Window is maximized.
visWSMinimized	=	0x20000000		# Window is minimized.
visWSNone	=	0x0		# No window state.
visWSRestored	=	0x10000000		# Window is restored.
visWSVisible	=	0x8000000		# Window is visible.

#values from visio.viswintypes
#****************************
visAnchorBarAddon	=	10		# Window created by an add-on that has tabs at the bottom when merged (floating, anchored, or docked window)
visAnchorBarBuiltIn	=	6		# Visio built-in window that has tabs at the bottom when merged?presently, the  Custom Properties, Size &amp; Position, Drawing Explorer, Master Explorer, Pan &amp; Zoom and Validation Issues windows (floating, anchored, or docked windows).
visApplication	=	5		# Microsoft Visio application window.
visDockedStencilAddon	=	11		# An add-on window that has docked stencil behavior.
visDockedStencilBuiltIn	=	7		# Stencil window docked in a drawing window.
visDrawing	=	1		# Drawing window (MDI frame window).
visDrawingAddon	=	8		# Drawing window created by an add-on (MDI frame window).
visIcon	=	4		# Icon editing window (MDI frame window).
visInvalWinID	=	-1		# Window has no ID.
visMasterGroupWin	=	96		# A group editing window of a group in a master.
visMasterWin	=	64		# A master drawing page window.
visPageGroupWin	=	160		# A group editing window of a group on a page.
visPageWin	=	128		# A drawing window showing a page.
visSheet	=	3		# ShapeSheet window (MDI frame window).

#values from visio.viszoombehavior
#****************************
visZoomInPlaceContainer	=	1		# The container performs the zoom.
visZoomNone	=	0		# Undefined zoom behavior; use the zoom behavior of the document or application.
visZoomVisio	=	2		# Microsoft Visio performs the zoom. The default.
visZoomVisioExact	=	4		# Visio zooms when open in place; Visio does not adjust the zoom level.
}
#End Enum

$vis=new-object PSCustomObject -Property $vis


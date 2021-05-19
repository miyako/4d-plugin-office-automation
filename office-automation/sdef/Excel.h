/*
 * Excel.h
 */

#import <AppKit/AppKit.h>
#import <ScriptingBridge/ScriptingBridge.h>


@class ExcelApplication, ExcelDocument, ExcelWindow, ExcelCommandBarControl, ExcelCommandBarButton, ExcelCommandBarCombobox, ExcelCommandBarPopup, ExcelCommandBar, ExcelDocumentProperty, ExcelCustomDocumentProperty, ExcelWebPageFont, ExcelExcelComment, ExcelODBCError, ExcelProtection, ExcelAboveAverageFormatCondition, ExcelAddIn, ExcelApplication, ExcelAutofilter, ExcelBorder, ExcelButton, ExcelCalculatedMember, ExcelCheckbox, ExcelColorScaleCriteria, ExcelColorScaleCriterion, ExcelColorScaleFormatCondition, ExcelColorstop, ExcelColorstops, ExcelConditionValue, ExcelCubeField, ExcelCustomView, ExcelDatabarBorder, ExcelDatabarFormatCondition, ExcelDefaultWebOptions, ExcelDialog, ExcelDisplayFormat, ExcelDropdown, ExcelFilter, ExcelFormatColor, ExcelFormatConditionIconObject, ExcelFormatConditionIconSet, ExcelFormatConditionIconSets, ExcelFormatCondition, ExcelGraphic, ExcelGroupbox, ExcelHorizontalPageBreak, ExcelHyperlink, ExcelIconCriteria, ExcelIconCriterion, ExcelIconSetFormatCondition, ExcelLabel, ExcelLinearGradient, ExcelListColumn, ExcelListObject, ExcelListRow, ExcelListbox, ExcelNamedItem, ExcelNegativeBarFormat, ExcelOptionButton, ExcelOutline, ExcelPageSetup, ExcelPane, ExcelPhonetic, ExcelPivotAxis, ExcelPivotCache, ExcelPivotCell, ExcelPivotField, ExcelCalculatedField, ExcelColumnField, ExcelDataField, ExcelHiddenField, ExcelPageField, ExcelPivotFilter, ExcelActiveFilter, ExcelPivotFormula, ExcelPivotItem, ExcelCalculatedItem, ExcelChildItem, ExcelColumnItem, ExcelHiddenItem, ExcelParentItem, ExcelPivotLine, ExcelPivotTable, ExcelQueryTable, ExcelRecentFile, ExcelRectangularGradient, ExcelRowField, ExcelRowItem, ExcelScenario, ExcelScrollbar, ExcelSheet, ExcelInternationalMacroSheet, ExcelMacroSheet, ExcelSlicer, ExcelSort, ExcelSortfield, ExcelSortfieldset, ExcelSpinner, ExcelTab, ExcelTableStyleElement, ExcelTableStyle, ExcelTextbox, ExcelTop10FormatCondition, ExcelTreeviewControl, ExcelUniqueValuesFormatCondition, ExcelValidation, ExcelValueChange, ExcelVerticalPageBreak, ExcelWebOptions, ExcelWindow, ExcelWorkbookConnection, ExcelWorkbook, ExcelDocument, ExcelWorksheet, ExcelXlspellingOptions, ExcelAdjustment, ExcelArc, ExcelBulletFormat, ExcelCalloutFormat, ExcelConnectorFormat, ExcelFillFormat, ExcelGlowFormat, ExcelGradientStop, ExcelLineFormat, ExcelLine, ExcelOfficeTheme, ExcelOval, ExcelParagraphFormat, ExcelPictureFormat, ExcelRectangle, ExcelReflectionFormat, ExcelRulerLevel, ExcelRuler, ExcelShadowFormat, ExcelShapeFont, ExcelShapeTextFrame, ExcelShape, ExcelCallout, ExcelPicture, ExcelShapeConnector, ExcelShapeLine, ExcelShapeTextbox, ExcelSoftEdgeFormat, ExcelTabStop, ExcelTextColumn, ExcelTextFrame, ExcelThemeColorScheme, ExcelThemeColor, ExcelThemeEffectScheme, ExcelThemeFontScheme, ExcelThemeFont, ExcelMajorThemeFont, ExcelMinorThemeFont, ExcelThreeDFormat, ExcelWordArtFormat, ExcelCharacter, ExcelFont, ExcelStyle, ExcelTextRange, ExcelParagraph, ExcelSentence, ExcelTextFlow, ExcelTextRangeCharacter, ExcelTextRangeLine, ExcelWord, ExcelRange, ExcelCell, ExcelColumn, ExcelRow, ExcelAutocorrect, ExcelAxisTitle, ExcelAxis, ExcelChartArea, ExcelChartFillFormat, ExcelChartFormat, ExcelChartGroup, ExcelAreaGroup, ExcelBarGroup, ExcelChartObject, ExcelChartTitle, ExcelChart, ExcelChartSheet, ExcelColumnGroup, ExcelCorners, ExcelDataLabel, ExcelDataTable, ExcelDisplayUnitLabel, ExcelDoughnutGroup, ExcelDownBars, ExcelDropLines, ExcelErrorBars, ExcelFloor, ExcelGridlines, ExcelHiloLines, ExcelInterior, ExcelLeaderLines, ExcelLegendEntry, ExcelLegendKey, ExcelLegend, ExcelLineGroup, ExcelPieGroup, ExcelPlotArea, ExcelRadarGroup, ExcelSeriesLines, ExcelSeriesPoint, ExcelSeries, ExcelTickLabels, ExcelTrendline, ExcelUpBars, ExcelWalls, ExcelXyGroup;

enum ExcelSaveOptions {
	ExcelSaveOptionsYes = 'yes ' /* Save the file. */,
	ExcelSaveOptionsNo = 'no  ' /* Do not save the file. */,
	ExcelSaveOptionsAsk = 'ask ' /* Ask the user whether or not to save the file. */
};
typedef enum ExcelSaveOptions ExcelSaveOptions;

enum ExcelPrintingErrorHandling {
	ExcelPrintingErrorHandlingStandard = 'lwst' /* Standard PostScript error handling */,
	ExcelPrintingErrorHandlingDetailed = 'lwdt' /* print a detailed report of PostScript errors */
};
typedef enum ExcelPrintingErrorHandling ExcelPrintingErrorHandling;

enum ExcelMsoLineDashStyle {
	ExcelMsoLineDashStyleLineDashStyleUnset = '\000\222\377\376',
	ExcelMsoLineDashStyleLineDashStyleSolid = '\000\223\000\001',
	ExcelMsoLineDashStyleLineDashStyleSquareDot = '\000\223\000\002',
	ExcelMsoLineDashStyleLineDashStyleRoundDot = '\000\223\000\003',
	ExcelMsoLineDashStyleLineDashStyleDash = '\000\223\000\004',
	ExcelMsoLineDashStyleLineDashStyleDashDot = '\000\223\000\005',
	ExcelMsoLineDashStyleLineDashStyleDashDotDot = '\000\223\000\006',
	ExcelMsoLineDashStyleLineDashStyleLongDash = '\000\223\000\007',
	ExcelMsoLineDashStyleLineDashStyleLongDashDot = '\000\223\000\010',
	ExcelMsoLineDashStyleLineDashStyleLongDashDotDot = '\000\223\000\011',
	ExcelMsoLineDashStyleLineDashStyleSystemDash = '\000\223\000\012',
	ExcelMsoLineDashStyleLineDashStyleSystemDot = '\000\223\000\013',
	ExcelMsoLineDashStyleLineDashStyleSystemDashDot = '\000\223\000\014'
};
typedef enum ExcelMsoLineDashStyle ExcelMsoLineDashStyle;

enum ExcelMsoLineStyle {
	ExcelMsoLineStyleLineStyleUnset = '\000\224\377\376',
	ExcelMsoLineStyleSingleLine = '\000\225\000\001',
	ExcelMsoLineStyleThinThinLine = '\000\225\000\002',
	ExcelMsoLineStyleThinThickLine = '\000\225\000\003',
	ExcelMsoLineStyleThickThinLine = '\000\225\000\004',
	ExcelMsoLineStyleThickBetweenThinLine = '\000\225\000\005'
};
typedef enum ExcelMsoLineStyle ExcelMsoLineStyle;

enum ExcelMsoArrowheadStyle {
	ExcelMsoArrowheadStyleArrowheadStyleUnset = '\000\221\377\376',
	ExcelMsoArrowheadStyleNoArrowhead = '\000\222\000\001',
	ExcelMsoArrowheadStyleTriangleArrowhead = '\000\222\000\002',
	ExcelMsoArrowheadStyleOpen_arrowhead = '\000\222\000\003',
	ExcelMsoArrowheadStyleStealthArrowhead = '\000\222\000\004',
	ExcelMsoArrowheadStyleDiamondArrowhead = '\000\222\000\005',
	ExcelMsoArrowheadStyleOvalArrowhead = '\000\222\000\006'
};
typedef enum ExcelMsoArrowheadStyle ExcelMsoArrowheadStyle;

enum ExcelMsoArrowheadWidth {
	ExcelMsoArrowheadWidthArrowheadWidthUnset = '\000\220\377\376',
	ExcelMsoArrowheadWidthNarrowWidthArrowhead = '\000\221\000\001',
	ExcelMsoArrowheadWidthMediumWidthArrowhead = '\000\221\000\002',
	ExcelMsoArrowheadWidthWideArrowhead = '\000\221\000\003'
};
typedef enum ExcelMsoArrowheadWidth ExcelMsoArrowheadWidth;

enum ExcelMsoArrowheadLength {
	ExcelMsoArrowheadLengthArrowheadLengthUnset = '\000\223\377\376',
	ExcelMsoArrowheadLengthShortArrowhead = '\000\224\000\001',
	ExcelMsoArrowheadLengthMediumArrowhead = '\000\224\000\002',
	ExcelMsoArrowheadLengthLongArrowhead = '\000\224\000\003'
};
typedef enum ExcelMsoArrowheadLength ExcelMsoArrowheadLength;

enum ExcelMsoFillType {
	ExcelMsoFillTypeFillUnset = '\000c\377\376',
	ExcelMsoFillTypeFillSolid = '\000d\000\001',
	ExcelMsoFillTypeFillPatterned = '\000d\000\002',
	ExcelMsoFillTypeFillGradient = '\000d\000\003',
	ExcelMsoFillTypeFillTextured = '\000d\000\004',
	ExcelMsoFillTypeFillBackground = '\000d\000\005',
	ExcelMsoFillTypeFillPicture = '\000d\000\006'
};
typedef enum ExcelMsoFillType ExcelMsoFillType;

enum ExcelMsoGradientStyle {
	ExcelMsoGradientStyleGradientUnset = '\000d\377\376',
	ExcelMsoGradientStyleHorizontalGradient = '\000e\000\001',
	ExcelMsoGradientStyleVerticalGradient = '\000e\000\002',
	ExcelMsoGradientStyleDiagonalUpGradient = '\000e\000\003',
	ExcelMsoGradientStyleDiagonalDownGradient = '\000e\000\004',
	ExcelMsoGradientStyleFromCornerGradient = '\000e\000\005',
	ExcelMsoGradientStyleFromTitleGradient = '\000e\000\006',
	ExcelMsoGradientStyleFromCenterGradient = '\000e\000\007'
};
typedef enum ExcelMsoGradientStyle ExcelMsoGradientStyle;

enum ExcelMsoGradientColorType {
	ExcelMsoGradientColorTypeGradientTypeUnset = '\003\357\377\376',
	ExcelMsoGradientColorTypeSingleShadeGradientType = '\003\360\000\001',
	ExcelMsoGradientColorTypeTwoColorsGradientType = '\003\360\000\002',
	ExcelMsoGradientColorTypePresetColorsGradientType = '\003\360\000\003',
	ExcelMsoGradientColorTypeMultiColorsGradientType = '\003\360\000\004'
};
typedef enum ExcelMsoGradientColorType ExcelMsoGradientColorType;

enum ExcelMsoTextureType {
	ExcelMsoTextureTypeTextureTypeTextureTypeUnset = '\003\360\377\376',
	ExcelMsoTextureTypeTextureTypePresetTexture = '\003\361\000\001',
	ExcelMsoTextureTypeTextureTypeUserDefinedTexture = '\003\361\000\002'
};
typedef enum ExcelMsoTextureType ExcelMsoTextureType;

enum ExcelMsoPresetTexture {
	ExcelMsoPresetTexturePresetTextureUnset = '\000e\377\376',
	ExcelMsoPresetTextureTexturePapyrus = '\000f\000\001',
	ExcelMsoPresetTextureTextureCanvas = '\000f\000\002',
	ExcelMsoPresetTextureTextureDenim = '\000f\000\003',
	ExcelMsoPresetTextureTextureWovenMat = '\000f\000\004',
	ExcelMsoPresetTextureTextureWaterDroplets = '\000f\000\005',
	ExcelMsoPresetTextureTexturePaperBag = '\000f\000\006',
	ExcelMsoPresetTextureTextureFishFossil = '\000f\000\007',
	ExcelMsoPresetTextureTextureSand = '\000f\000\010',
	ExcelMsoPresetTextureTextureGreenMarble = '\000f\000\011',
	ExcelMsoPresetTextureTextureWhiteMarble = '\000f\000\012',
	ExcelMsoPresetTextureTextureBrownMarble = '\000f\000\013',
	ExcelMsoPresetTextureTextureGranite = '\000f\000\014',
	ExcelMsoPresetTextureTextureNewsprint = '\000f\000\015',
	ExcelMsoPresetTextureTextureRecycledPaper = '\000f\000\016',
	ExcelMsoPresetTextureTextureParchment = '\000f\000\017',
	ExcelMsoPresetTextureTextureStationery = '\000f\000\020',
	ExcelMsoPresetTextureTextureBlueTissuePaper = '\000f\000\021',
	ExcelMsoPresetTextureTexturePinkTissuePaper = '\000f\000\022',
	ExcelMsoPresetTextureTexturePurpleMesh = '\000f\000\023',
	ExcelMsoPresetTextureTextureBouquet = '\000f\000\024',
	ExcelMsoPresetTextureTextureCork = '\000f\000\025',
	ExcelMsoPresetTextureTextureWalnut = '\000f\000\026',
	ExcelMsoPresetTextureTextureOak = '\000f\000\027',
	ExcelMsoPresetTextureTextureMediumWood = '\000f\000\030'
};
typedef enum ExcelMsoPresetTexture ExcelMsoPresetTexture;

enum ExcelMsoPatternType {
	ExcelMsoPatternTypePatternUnset = '\000f\377\376',
	ExcelMsoPatternTypeFivePercentPattern = '\000g\000\001',
	ExcelMsoPatternTypeTenPercentPattern = '\000g\000\002',
	ExcelMsoPatternTypeTwentyPercentPattern = '\000g\000\003',
	ExcelMsoPatternTypeTwentyFivePercentPattern = '\000g\000\004',
	ExcelMsoPatternTypeThirtyPercentPattern = '\000g\000\005',
	ExcelMsoPatternTypeFortyPercentPattern = '\000g\000\006',
	ExcelMsoPatternTypeFiftyPercentPattern = '\000g\000\007',
	ExcelMsoPatternTypeSixtyPercentPattern = '\000g\000\010',
	ExcelMsoPatternTypeSeventyPercentPattern = '\000g\000\011',
	ExcelMsoPatternTypeSeventyFivePercentPattern = '\000g\000\012',
	ExcelMsoPatternTypeEightyPercentPattern = '\000g\000\013',
	ExcelMsoPatternTypeNinetyPercentPattern = '\000g\000\014',
	ExcelMsoPatternTypeDarkHorizontalPattern = '\000g\000\015',
	ExcelMsoPatternTypeDarkVerticalPattern = '\000g\000\016',
	ExcelMsoPatternTypeDarkDownwardDiagonalPattern = '\000g\000\017',
	ExcelMsoPatternTypeDarkUpwardDiagonalPattern = '\000g\000\020',
	ExcelMsoPatternTypeSmallCheckerBoardPattern = '\000g\000\021',
	ExcelMsoPatternTypeTrellisPattern = '\000g\000\022',
	ExcelMsoPatternTypeLightHorizontalPattern = '\000g\000\023',
	ExcelMsoPatternTypeLightVerticalPattern = '\000g\000\024',
	ExcelMsoPatternTypeLightDownwardDiagonalPattern = '\000g\000\025',
	ExcelMsoPatternTypeLightUpwardDiagonalPattern = '\000g\000\026',
	ExcelMsoPatternTypeSmallGridPattern = '\000g\000\027',
	ExcelMsoPatternTypeDottedDiamondPattern = '\000g\000\030',
	ExcelMsoPatternTypeWideDownwardDiagonal = '\000g\000\031',
	ExcelMsoPatternTypeWideUpwardDiagonalPattern = '\000g\000\032',
	ExcelMsoPatternTypeDashedUpwardDiagonalPattern = '\000g\000\033',
	ExcelMsoPatternTypeDashedDownwardDiagonalPattern = '\000g\000\034',
	ExcelMsoPatternTypeNarrowVerticalPattern = '\000g\000\035',
	ExcelMsoPatternTypeNarrowHorizontalPattern = '\000g\000\036',
	ExcelMsoPatternTypeDashedVerticalPattern = '\000g\000\037',
	ExcelMsoPatternTypeDashedHorizontalPattern = '\000g\000 ',
	ExcelMsoPatternTypeLargeConfettiPattern = '\000g\000!',
	ExcelMsoPatternTypeLargeGridPattern = '\000g\000\"',
	ExcelMsoPatternTypeHorizontalBrickPattern = '\000g\000#',
	ExcelMsoPatternTypeLargeCheckerBoardPattern = '\000g\000$',
	ExcelMsoPatternTypeSmallConfettiPattern = '\000g\000%',
	ExcelMsoPatternTypeZigZagPattern = '\000g\000&',
	ExcelMsoPatternTypeSolidDiamondPattern = '\000g\000\'',
	ExcelMsoPatternTypeDiagonalBrickPattern = '\000g\000(',
	ExcelMsoPatternTypeOutlinedDiamondPattern = '\000g\000)',
	ExcelMsoPatternTypePlaidPattern = '\000g\000*',
	ExcelMsoPatternTypeSpherePattern = '\000g\000+',
	ExcelMsoPatternTypeWeavePattern = '\000g\000,',
	ExcelMsoPatternTypeDottedGridPattern = '\000g\000-',
	ExcelMsoPatternTypeDivotPattern = '\000g\000.',
	ExcelMsoPatternTypeShinglePattern = '\000g\000/',
	ExcelMsoPatternTypeWavePattern = '\000g\0000',
	ExcelMsoPatternTypeHorizontalPattern = '\000g\0001',
	ExcelMsoPatternTypeVerticalPattern = '\000g\0002',
	ExcelMsoPatternTypeCrossPattern = '\000g\0003',
	ExcelMsoPatternTypeDownwardDiagonalPattern = '\000g\0004',
	ExcelMsoPatternTypeUpwardDiagonalPattern = '\000g\0005',
	ExcelMsoPatternTypeDiagonalCrossPattern = '\000g\0005'
};
typedef enum ExcelMsoPatternType ExcelMsoPatternType;

enum ExcelMsoPresetGradientType {
	ExcelMsoPresetGradientTypePresetGradientUnset = '\000g\377\376',
	ExcelMsoPresetGradientTypeGradientEarlySunset = '\000h\000\001',
	ExcelMsoPresetGradientTypeGradientLateSunset = '\000h\000\002',
	ExcelMsoPresetGradientTypeGradientNightfall = '\000h\000\003',
	ExcelMsoPresetGradientTypeGradientDaybreak = '\000h\000\004',
	ExcelMsoPresetGradientTypeGradientHorizon = '\000h\000\005',
	ExcelMsoPresetGradientTypeGradientDesert = '\000h\000\006',
	ExcelMsoPresetGradientTypeGradientOcean = '\000h\000\007',
	ExcelMsoPresetGradientTypeGradientCalmWater = '\000h\000\010',
	ExcelMsoPresetGradientTypeGradientFire = '\000h\000\011',
	ExcelMsoPresetGradientTypeGradientFog = '\000h\000\012',
	ExcelMsoPresetGradientTypeGradientMoss = '\000h\000\013',
	ExcelMsoPresetGradientTypeGradientPeacock = '\000h\000\014',
	ExcelMsoPresetGradientTypeGradientWheat = '\000h\000\015',
	ExcelMsoPresetGradientTypeGradientParchment = '\000h\000\016',
	ExcelMsoPresetGradientTypeGradientMahogany = '\000h\000\017',
	ExcelMsoPresetGradientTypeGradientRainbow = '\000h\000\020',
	ExcelMsoPresetGradientTypeGradientRainbow2 = '\000h\000\021',
	ExcelMsoPresetGradientTypeGradientGold = '\000h\000\022',
	ExcelMsoPresetGradientTypeGradientGold2 = '\000h\000\023',
	ExcelMsoPresetGradientTypeGradientBrass = '\000h\000\024',
	ExcelMsoPresetGradientTypeGradientChrome = '\000h\000\025',
	ExcelMsoPresetGradientTypeGradientChrome2 = '\000h\000\026',
	ExcelMsoPresetGradientTypeGradientSilver = '\000h\000\027',
	ExcelMsoPresetGradientTypeGradientSapphire = '\000h\000\030'
};
typedef enum ExcelMsoPresetGradientType ExcelMsoPresetGradientType;

enum ExcelMsoShadowType {
	ExcelMsoShadowTypeShadowUnset = '\003_\377\376',
	ExcelMsoShadowTypeShadow1 = '\003`\000\001',
	ExcelMsoShadowTypeShadow2 = '\003`\000\002',
	ExcelMsoShadowTypeShadow3 = '\003`\000\003',
	ExcelMsoShadowTypeShadow4 = '\003`\000\004',
	ExcelMsoShadowTypeShadow5 = '\003`\000\005',
	ExcelMsoShadowTypeShadow6 = '\003`\000\006',
	ExcelMsoShadowTypeShadow7 = '\003`\000\007',
	ExcelMsoShadowTypeShadow8 = '\003`\000\010',
	ExcelMsoShadowTypeShadow9 = '\003`\000\011',
	ExcelMsoShadowTypeShadow10 = '\003`\000\012',
	ExcelMsoShadowTypeShadow11 = '\003`\000\013',
	ExcelMsoShadowTypeShadow12 = '\003`\000\014',
	ExcelMsoShadowTypeShadow13 = '\003`\000\015',
	ExcelMsoShadowTypeShadow14 = '\003`\000\016',
	ExcelMsoShadowTypeShadow15 = '\003`\000\017',
	ExcelMsoShadowTypeShadow16 = '\003`\000\020',
	ExcelMsoShadowTypeShadow17 = '\003`\000\021',
	ExcelMsoShadowTypeShadow18 = '\003`\000\022',
	ExcelMsoShadowTypeShadow19 = '\003`\000\023',
	ExcelMsoShadowTypeShadow20 = '\003`\000\024',
	ExcelMsoShadowTypeShadow21 = '\003`\000\025',
	ExcelMsoShadowTypeShadow22 = '\003`\000\026',
	ExcelMsoShadowTypeShadow23 = '\003`\000\027',
	ExcelMsoShadowTypeShadow24 = '\003`\000\030',
	ExcelMsoShadowTypeShadow25 = '\003`\000\031',
	ExcelMsoShadowTypeShadow26 = '\003`\000\032',
	ExcelMsoShadowTypeShadow27 = '\003`\000\033',
	ExcelMsoShadowTypeShadow28 = '\003`\000\034',
	ExcelMsoShadowTypeShadow29 = '\003`\000\035',
	ExcelMsoShadowTypeShadow30 = '\003`\000\036',
	ExcelMsoShadowTypeShadow31 = '\003`\000\037',
	ExcelMsoShadowTypeShadow32 = '\003`\000 ',
	ExcelMsoShadowTypeShadow33 = '\003`\000!',
	ExcelMsoShadowTypeShadow34 = '\003`\000\"',
	ExcelMsoShadowTypeShadow35 = '\003`\000#',
	ExcelMsoShadowTypeShadow36 = '\003`\000$',
	ExcelMsoShadowTypeShadow37 = '\003`\000%',
	ExcelMsoShadowTypeShadow38 = '\003`\000&',
	ExcelMsoShadowTypeShadow39 = '\003`\000\'',
	ExcelMsoShadowTypeShadow40 = '\003`\000(',
	ExcelMsoShadowTypeShadow41 = '\003`\000)',
	ExcelMsoShadowTypeShadow42 = '\003`\000*',
	ExcelMsoShadowTypeShadow43 = '\003`\000+'
};
typedef enum ExcelMsoShadowType ExcelMsoShadowType;

enum ExcelMsoPresetTextEffect {
	ExcelMsoPresetTextEffectWordartFormatUnset = '\003\361\377\376',
	ExcelMsoPresetTextEffectWordartFormat1 = '\003\362\000\000',
	ExcelMsoPresetTextEffectWordartFormat2 = '\003\362\000\001',
	ExcelMsoPresetTextEffectWordartFormat3 = '\003\362\000\002',
	ExcelMsoPresetTextEffectWordartFormat4 = '\003\362\000\003',
	ExcelMsoPresetTextEffectWordartFormat5 = '\003\362\000\004',
	ExcelMsoPresetTextEffectWordartFormat6 = '\003\362\000\005',
	ExcelMsoPresetTextEffectWordartFormat7 = '\003\362\000\006',
	ExcelMsoPresetTextEffectWordartFormat8 = '\003\362\000\007',
	ExcelMsoPresetTextEffectWordartFormat9 = '\003\362\000\010',
	ExcelMsoPresetTextEffectWordartFormat10 = '\003\362\000\011',
	ExcelMsoPresetTextEffectWordartFormat11 = '\003\362\000\012',
	ExcelMsoPresetTextEffectWordartFormat12 = '\003\362\000\013',
	ExcelMsoPresetTextEffectWordartFormat13 = '\003\362\000\014',
	ExcelMsoPresetTextEffectWordartFormat14 = '\003\362\000\015',
	ExcelMsoPresetTextEffectWordartFormat15 = '\003\362\000\016',
	ExcelMsoPresetTextEffectWordartFormat16 = '\003\362\000\017',
	ExcelMsoPresetTextEffectWordartFormat17 = '\003\362\000\020',
	ExcelMsoPresetTextEffectWordartFormat18 = '\003\362\000\021',
	ExcelMsoPresetTextEffectWordartFormat19 = '\003\362\000\022',
	ExcelMsoPresetTextEffectWordartFormat20 = '\003\362\000\023',
	ExcelMsoPresetTextEffectWordartFormat21 = '\003\362\000\024',
	ExcelMsoPresetTextEffectWordartFormat22 = '\003\362\000\025',
	ExcelMsoPresetTextEffectWordartFormat23 = '\003\362\000\026',
	ExcelMsoPresetTextEffectWordartFormat24 = '\003\362\000\027',
	ExcelMsoPresetTextEffectWordartFormat25 = '\003\362\000\030',
	ExcelMsoPresetTextEffectWordartFormat26 = '\003\362\000\031',
	ExcelMsoPresetTextEffectWordartFormat27 = '\003\362\000\032',
	ExcelMsoPresetTextEffectWordartFormat28 = '\003\362\000\033',
	ExcelMsoPresetTextEffectWordartFormat29 = '\003\362\000\034',
	ExcelMsoPresetTextEffectWordartFormat30 = '\003\362\000\035',
	ExcelMsoPresetTextEffectWordartFormat31 = '\003\362\000\036',
	ExcelMsoPresetTextEffectWordartFormat32 = '\003\362\000\037',
	ExcelMsoPresetTextEffectWordartFormat33 = '\003\362\000 ',
	ExcelMsoPresetTextEffectWordartFormat34 = '\003\362\000!',
	ExcelMsoPresetTextEffectWordartFormat35 = '\003\362\000\"',
	ExcelMsoPresetTextEffectWordartFormat36 = '\003\362\000#',
	ExcelMsoPresetTextEffectWordartFormat37 = '\003\362\000$',
	ExcelMsoPresetTextEffectWordartFormat38 = '\003\362\000%',
	ExcelMsoPresetTextEffectWordartFormat39 = '\003\362\000&',
	ExcelMsoPresetTextEffectWordartFormat40 = '\003\362\000\'',
	ExcelMsoPresetTextEffectWordartFormat41 = '\003\362\000(',
	ExcelMsoPresetTextEffectWordartFormat42 = '\003\362\000)',
	ExcelMsoPresetTextEffectWordartFormat43 = '\003\362\000*',
	ExcelMsoPresetTextEffectWordartFormat44 = '\003\362\000+',
	ExcelMsoPresetTextEffectWordartFormat45 = '\003\362\000,',
	ExcelMsoPresetTextEffectWordartFormat46 = '\003\362\000-',
	ExcelMsoPresetTextEffectWordartFormat47 = '\003\362\000.',
	ExcelMsoPresetTextEffectWordartFormat48 = '\003\362\000/',
	ExcelMsoPresetTextEffectWordartFormat49 = '\003\362\0000',
	ExcelMsoPresetTextEffectWordartFormat50 = '\003\362\0001'
};
typedef enum ExcelMsoPresetTextEffect ExcelMsoPresetTextEffect;

enum ExcelMsoPresetTextEffectShape {
	ExcelMsoPresetTextEffectShapeTextEffectShapeUnset = '\000\227\377\376',
	ExcelMsoPresetTextEffectShapePlainText = '\000\230\000\001',
	ExcelMsoPresetTextEffectShapeStop = '\000\230\000\002',
	ExcelMsoPresetTextEffectShapeTriangleUp = '\000\230\000\003',
	ExcelMsoPresetTextEffectShapeTriangleDown = '\000\230\000\004',
	ExcelMsoPresetTextEffectShapeChevronUp = '\000\230\000\005',
	ExcelMsoPresetTextEffectShapeChevronDown = '\000\230\000\006',
	ExcelMsoPresetTextEffectShapeRingInside = '\000\230\000\007',
	ExcelMsoPresetTextEffectShapeRingOutside = '\000\230\000\010',
	ExcelMsoPresetTextEffectShapeArchUpCurve = '\000\230\000\011',
	ExcelMsoPresetTextEffectShapeArchDownCurve = '\000\230\000\012',
	ExcelMsoPresetTextEffectShapeCircleCurve = '\000\230\000\013',
	ExcelMsoPresetTextEffectShapeButtonCurve = '\000\230\000\014',
	ExcelMsoPresetTextEffectShapeArchUpPour = '\000\230\000\015',
	ExcelMsoPresetTextEffectShapeArchDownPour = '\000\230\000\016',
	ExcelMsoPresetTextEffectShapeCirclePour = '\000\230\000\017',
	ExcelMsoPresetTextEffectShapeButtonPour = '\000\230\000\020',
	ExcelMsoPresetTextEffectShapeCurveUp = '\000\230\000\021',
	ExcelMsoPresetTextEffectShapeCurveDown = '\000\230\000\022',
	ExcelMsoPresetTextEffectShapeCanUp = '\000\230\000\023',
	ExcelMsoPresetTextEffectShapeCanDown = '\000\230\000\024',
	ExcelMsoPresetTextEffectShapeWave1 = '\000\230\000\025',
	ExcelMsoPresetTextEffectShapeWave2 = '\000\230\000\026',
	ExcelMsoPresetTextEffectShapeDoubleWave1 = '\000\230\000\027',
	ExcelMsoPresetTextEffectShapeDoubleWave2 = '\000\230\000\030',
	ExcelMsoPresetTextEffectShapeInflate = '\000\230\000\031',
	ExcelMsoPresetTextEffectShapeDeflate = '\000\230\000\032',
	ExcelMsoPresetTextEffectShapeInflateBottom = '\000\230\000\033',
	ExcelMsoPresetTextEffectShapeDeflateBottom = '\000\230\000\034',
	ExcelMsoPresetTextEffectShapeInflateTop = '\000\230\000\035',
	ExcelMsoPresetTextEffectShapeDeflateTop = '\000\230\000\036',
	ExcelMsoPresetTextEffectShapeDeflateInflate = '\000\230\000\037',
	ExcelMsoPresetTextEffectShapeDeflateInflateDeflate = '\000\230\000 ',
	ExcelMsoPresetTextEffectShapeFadeRight = '\000\230\000!',
	ExcelMsoPresetTextEffectShapeFadeLeft = '\000\230\000\"',
	ExcelMsoPresetTextEffectShapeFadeUp = '\000\230\000#',
	ExcelMsoPresetTextEffectShapeFadeDown = '\000\230\000$',
	ExcelMsoPresetTextEffectShapeSlantUp = '\000\230\000%',
	ExcelMsoPresetTextEffectShapeSlantDown = '\000\230\000&',
	ExcelMsoPresetTextEffectShapeCascadeUp = '\000\230\000\'',
	ExcelMsoPresetTextEffectShapeCascadeDown = '\000\230\000('
};
typedef enum ExcelMsoPresetTextEffectShape ExcelMsoPresetTextEffectShape;

enum ExcelMsoTextEffectAlignment {
	ExcelMsoTextEffectAlignmentTextEffectAlignmentUnset = '\000\226\377\376',
	ExcelMsoTextEffectAlignmentLeftTextEffectAlignment = '\000\227\000\001',
	ExcelMsoTextEffectAlignmentCenteredTextEffectAlignment = '\000\227\000\002',
	ExcelMsoTextEffectAlignmentRightTextEffectAlignment = '\000\227\000\003',
	ExcelMsoTextEffectAlignmentJustifyTextEffectAlignment = '\000\227\000\004',
	ExcelMsoTextEffectAlignmentWordJustifyTextEffectAlignment = '\000\227\000\005',
	ExcelMsoTextEffectAlignmentStretchJustifyTextEffectAlignment = '\000\227\000\006'
};
typedef enum ExcelMsoTextEffectAlignment ExcelMsoTextEffectAlignment;

enum ExcelMsoPresetLightingDirection {
	ExcelMsoPresetLightingDirectionPresetLightingDirectionUnset = '\000\233\377\376',
	ExcelMsoPresetLightingDirectionLightFromTopLeft = '\000\234\000\001',
	ExcelMsoPresetLightingDirectionLightFromTop = '\000\234\000\002',
	ExcelMsoPresetLightingDirectionLightFromTopRight = '\000\234\000\003',
	ExcelMsoPresetLightingDirectionLightFromLeft = '\000\234\000\004',
	ExcelMsoPresetLightingDirectionLightFromNone = '\000\234\000\005',
	ExcelMsoPresetLightingDirectionLightFromRight = '\000\234\000\006',
	ExcelMsoPresetLightingDirectionLightFromBottomLeft = '\000\234\000\007',
	ExcelMsoPresetLightingDirectionLightFromBottom = '\000\234\000\010',
	ExcelMsoPresetLightingDirectionLightFromBottomRight = '\000\234\000\011'
};
typedef enum ExcelMsoPresetLightingDirection ExcelMsoPresetLightingDirection;

enum ExcelMsoPresetLightingSoftness {
	ExcelMsoPresetLightingSoftnessLightingSoftnessUnset = '\000\234\377\376',
	ExcelMsoPresetLightingSoftnessLightingDim = '\000\235\000\001',
	ExcelMsoPresetLightingSoftnessLightingNormal = '\000\235\000\002',
	ExcelMsoPresetLightingSoftnessLightingBright = '\000\235\000\003'
};
typedef enum ExcelMsoPresetLightingSoftness ExcelMsoPresetLightingSoftness;

enum ExcelMsoPresetMaterial {
	ExcelMsoPresetMaterialPresetMaterialUnset = '\000\235\377\376',
	ExcelMsoPresetMaterialMatte = '\000\236\000\001',
	ExcelMsoPresetMaterialPlastic = '\000\236\000\002',
	ExcelMsoPresetMaterialMetal = '\000\236\000\003',
	ExcelMsoPresetMaterialWireframe = '\000\236\000\004',
	ExcelMsoPresetMaterialMatte2 = '\000\236\000\005',
	ExcelMsoPresetMaterialPlastic2 = '\000\236\000\006',
	ExcelMsoPresetMaterialMetal2 = '\000\236\000\007',
	ExcelMsoPresetMaterialWarmMatte = '\000\236\000\010',
	ExcelMsoPresetMaterialTranslucentPowder = '\000\236\000\011',
	ExcelMsoPresetMaterialPowder = '\000\236\000\012',
	ExcelMsoPresetMaterialDarkEdge = '\000\236\000\013',
	ExcelMsoPresetMaterialSoftEdge = '\000\236\000\014',
	ExcelMsoPresetMaterialMaterialClear = '\000\236\000\015',
	ExcelMsoPresetMaterialFlat = '\000\236\000\016',
	ExcelMsoPresetMaterialSoftMetal = '\000\236\000\017'
};
typedef enum ExcelMsoPresetMaterial ExcelMsoPresetMaterial;

enum ExcelMsoPresetExtrusionDirection {
	ExcelMsoPresetExtrusionDirectionPresetExtrusionDirectionUnset = '\000\231\377\376',
	ExcelMsoPresetExtrusionDirectionExtrudeBottomRight = '\000\232\000\001',
	ExcelMsoPresetExtrusionDirectionExtrudeBottom = '\000\232\000\002',
	ExcelMsoPresetExtrusionDirectionExtrudeBottomLeft = '\000\232\000\003',
	ExcelMsoPresetExtrusionDirectionExtrudeRight = '\000\232\000\004',
	ExcelMsoPresetExtrusionDirectionExtrudeNone = '\000\232\000\005',
	ExcelMsoPresetExtrusionDirectionExtrudeLeft = '\000\232\000\006',
	ExcelMsoPresetExtrusionDirectionExtrudeTopRight = '\000\232\000\007',
	ExcelMsoPresetExtrusionDirectionExtrudeTop = '\000\232\000\010',
	ExcelMsoPresetExtrusionDirectionExtrudeTopLeft = '\000\232\000\011'
};
typedef enum ExcelMsoPresetExtrusionDirection ExcelMsoPresetExtrusionDirection;

enum ExcelMsoPresetThreeDFormat {
	ExcelMsoPresetThreeDFormatPresetThreeDFormatUnset = '\000\230\377\376',
	ExcelMsoPresetThreeDFormatFormat1 = '\000\231\000\001',
	ExcelMsoPresetThreeDFormatFormat2 = '\000\231\000\002',
	ExcelMsoPresetThreeDFormatFormat3 = '\000\231\000\003',
	ExcelMsoPresetThreeDFormatFormat4 = '\000\231\000\004',
	ExcelMsoPresetThreeDFormatFormat5 = '\000\231\000\005',
	ExcelMsoPresetThreeDFormatFormat6 = '\000\231\000\006',
	ExcelMsoPresetThreeDFormatFormat7 = '\000\231\000\007',
	ExcelMsoPresetThreeDFormatFormat8 = '\000\231\000\010',
	ExcelMsoPresetThreeDFormatFormat9 = '\000\231\000\011',
	ExcelMsoPresetThreeDFormatFormat10 = '\000\231\000\012',
	ExcelMsoPresetThreeDFormatFormat11 = '\000\231\000\013',
	ExcelMsoPresetThreeDFormatFormat12 = '\000\231\000\014',
	ExcelMsoPresetThreeDFormatFormat13 = '\000\231\000\015',
	ExcelMsoPresetThreeDFormatFormat14 = '\000\231\000\016',
	ExcelMsoPresetThreeDFormatFormat15 = '\000\231\000\017',
	ExcelMsoPresetThreeDFormatFormat16 = '\000\231\000\020',
	ExcelMsoPresetThreeDFormatFormat17 = '\000\231\000\021',
	ExcelMsoPresetThreeDFormatFormat18 = '\000\231\000\022',
	ExcelMsoPresetThreeDFormatFormat19 = '\000\231\000\023',
	ExcelMsoPresetThreeDFormatFormat20 = '\000\231\000\024'
};
typedef enum ExcelMsoPresetThreeDFormat ExcelMsoPresetThreeDFormat;

enum ExcelMsoExtrusionColorType {
	ExcelMsoExtrusionColorTypeExtrusionColorUnset = '\000\232\377\376',
	ExcelMsoExtrusionColorTypeExtrusionColorAutomatic = '\000\233\000\001',
	ExcelMsoExtrusionColorTypeExtrusionColorCustom = '\000\233\000\002'
};
typedef enum ExcelMsoExtrusionColorType ExcelMsoExtrusionColorType;

enum ExcelMsoConnectorType {
	ExcelMsoConnectorTypeConnectorTypeUnset = '\000h\377\376',
	ExcelMsoConnectorTypeStraight = '\000i\000\001',
	ExcelMsoConnectorTypeElbow = '\000i\000\002',
	ExcelMsoConnectorTypeCurve = '\000i\000\003'
};
typedef enum ExcelMsoConnectorType ExcelMsoConnectorType;

enum ExcelMsoHorizontalAnchor {
	ExcelMsoHorizontalAnchorHorizontalAnchorUnset = '\000\236\377\376',
	ExcelMsoHorizontalAnchorHorizontalAnchorNone = '\000\237\000\001',
	ExcelMsoHorizontalAnchorHorizontalAnchorCenter = '\000\237\000\002'
};
typedef enum ExcelMsoHorizontalAnchor ExcelMsoHorizontalAnchor;

enum ExcelMsoVerticalAnchor {
	ExcelMsoVerticalAnchorVerticalAnchorUnset = '\000\237\377\376',
	ExcelMsoVerticalAnchorAnchorTop = '\000\240\000\001',
	ExcelMsoVerticalAnchorAnchorTopBaseline = '\000\240\000\002',
	ExcelMsoVerticalAnchorAnchorMiddle = '\000\240\000\003',
	ExcelMsoVerticalAnchorAnchorBottom = '\000\240\000\004',
	ExcelMsoVerticalAnchorAnchorBottomBaseline = '\000\240\000\005'
};
typedef enum ExcelMsoVerticalAnchor ExcelMsoVerticalAnchor;

enum ExcelMsoAutoShapeType {
	ExcelMsoAutoShapeTypeAutoshapeShapeTypeUnset = '\000i\377\376',
	ExcelMsoAutoShapeTypeAutoshapeRectangle = '\000j\000\001',
	ExcelMsoAutoShapeTypeAutoshapeParallelogram = '\000j\000\002',
	ExcelMsoAutoShapeTypeAutoshapeTrapezoid = '\000j\000\003',
	ExcelMsoAutoShapeTypeAutoshapeDiamond = '\000j\000\004',
	ExcelMsoAutoShapeTypeAutoshapeRoundedRectangle = '\000j\000\005',
	ExcelMsoAutoShapeTypeAutoshapeOctagon = '\000j\000\006',
	ExcelMsoAutoShapeTypeAutoshapeIsoscelesTriangle = '\000j\000\007',
	ExcelMsoAutoShapeTypeAutoshapeRightTriangle = '\000j\000\010',
	ExcelMsoAutoShapeTypeAutoshapeOval = '\000j\000\011',
	ExcelMsoAutoShapeTypeAutoshapeHexagon = '\000j\000\012',
	ExcelMsoAutoShapeTypeAutoshapeCross = '\000j\000\013',
	ExcelMsoAutoShapeTypeAutoshapeRegularPentagon = '\000j\000\014',
	ExcelMsoAutoShapeTypeAutoshapeCan = '\000j\000\015',
	ExcelMsoAutoShapeTypeAutoshapeCube = '\000j\000\016',
	ExcelMsoAutoShapeTypeAutoshapeBevel = '\000j\000\017',
	ExcelMsoAutoShapeTypeAutoshapeFoldedCorner = '\000j\000\020',
	ExcelMsoAutoShapeTypeAutoshapeSmileyFace = '\000j\000\021',
	ExcelMsoAutoShapeTypeAutoshapeDonut = '\000j\000\022',
	ExcelMsoAutoShapeTypeAutoshapeNoSymbol = '\000j\000\023',
	ExcelMsoAutoShapeTypeAutoshapeBlockArc = '\000j\000\024',
	ExcelMsoAutoShapeTypeAutoshapeHeart = '\000j\000\025',
	ExcelMsoAutoShapeTypeAutoshapeLightningBolt = '\000j\000\026',
	ExcelMsoAutoShapeTypeAutoshapeSun = '\000j\000\027',
	ExcelMsoAutoShapeTypeAutoshapeMoon = '\000j\000\030',
	ExcelMsoAutoShapeTypeAutoshapeArc = '\000j\000\031',
	ExcelMsoAutoShapeTypeAutoshapeDoubleBracket = '\000j\000\032',
	ExcelMsoAutoShapeTypeAutoshapeDoubleBrace = '\000j\000\033',
	ExcelMsoAutoShapeTypeAutoshapePlaque = '\000j\000\034',
	ExcelMsoAutoShapeTypeAutoshapeLeftBracket = '\000j\000\035',
	ExcelMsoAutoShapeTypeAutoshapeRightBracket = '\000j\000\036',
	ExcelMsoAutoShapeTypeAutoshapeLeftBrace = '\000j\000\037',
	ExcelMsoAutoShapeTypeAutoshapeRightBrace = '\000j\000 ',
	ExcelMsoAutoShapeTypeAutoshapeRightArrow = '\000j\000!',
	ExcelMsoAutoShapeTypeAutoshapeLeftArrow = '\000j\000\"',
	ExcelMsoAutoShapeTypeAutoshapeUpArrow = '\000j\000#',
	ExcelMsoAutoShapeTypeAutoshapeDownArrow = '\000j\000$',
	ExcelMsoAutoShapeTypeAutoshapeLeftRightArrow = '\000j\000%',
	ExcelMsoAutoShapeTypeAutoshapeUpDownArrow = '\000j\000&',
	ExcelMsoAutoShapeTypeAutoshapeQuadArrow = '\000j\000\'',
	ExcelMsoAutoShapeTypeAutoshapeLeftRightUpArrow = '\000j\000(',
	ExcelMsoAutoShapeTypeAutoshapeBentArrow = '\000j\000)',
	ExcelMsoAutoShapeTypeAutoshapeUTurnArrow = '\000j\000*',
	ExcelMsoAutoShapeTypeAutoshapeLeftUpArrow = '\000j\000+',
	ExcelMsoAutoShapeTypeAutoshapeBentUpArrow = '\000j\000,',
	ExcelMsoAutoShapeTypeAutoshapeCurvedRightArrow = '\000j\000-',
	ExcelMsoAutoShapeTypeAutoshapeCurvedLeftArrow = '\000j\000.',
	ExcelMsoAutoShapeTypeAutoshapeCurvedUpArrow = '\000j\000/',
	ExcelMsoAutoShapeTypeAutoshapeCurvedDownArrow = '\000j\0000',
	ExcelMsoAutoShapeTypeAutoshapeStripedRightArrow = '\000j\0001',
	ExcelMsoAutoShapeTypeAutoshapeNotchedRightArrow = '\000j\0002',
	ExcelMsoAutoShapeTypeAutoshapePentagon = '\000j\0003',
	ExcelMsoAutoShapeTypeAutoshapeChevron = '\000j\0004',
	ExcelMsoAutoShapeTypeAutoshapeRightArrowCallout = '\000j\0005',
	ExcelMsoAutoShapeTypeAutoshapeLeftArrowCallout = '\000j\0006',
	ExcelMsoAutoShapeTypeAutoshapeUpArrowCallout = '\000j\0007',
	ExcelMsoAutoShapeTypeAutoshapeDownArrowCallout = '\000j\0008',
	ExcelMsoAutoShapeTypeAutoshapeLeftRightArrowCallout = '\000j\0009',
	ExcelMsoAutoShapeTypeAutoshapeUpDownArrowCallout = '\000j\000:',
	ExcelMsoAutoShapeTypeAutoshapeQuadArrowCallout = '\000j\000;',
	ExcelMsoAutoShapeTypeAutoshapeCircularArrow = '\000j\000<',
	ExcelMsoAutoShapeTypeAutoshapeFlowchartProcess = '\000j\000=',
	ExcelMsoAutoShapeTypeAutoshapeFlowchartAlternateProcess = '\000j\000>',
	ExcelMsoAutoShapeTypeAutoshapeFlowchartDecision = '\000j\000\?',
	ExcelMsoAutoShapeTypeAutoshapeFlowchartData = '\000j\000@',
	ExcelMsoAutoShapeTypeAutoshapeFlowchartPredefinedProcess = '\000j\000A',
	ExcelMsoAutoShapeTypeAutoshapeFlowchartInternalStorage = '\000j\000B',
	ExcelMsoAutoShapeTypeAutoshapeFlowchartDocument = '\000j\000C',
	ExcelMsoAutoShapeTypeAutoshapeFlowchartMultiDocument = '\000j\000D',
	ExcelMsoAutoShapeTypeAutoshapeFlowchartTerminator = '\000j\000E',
	ExcelMsoAutoShapeTypeAutoshapeFlowchartPreparation = '\000j\000F',
	ExcelMsoAutoShapeTypeAutoshapeFlowchartManualInput = '\000j\000G',
	ExcelMsoAutoShapeTypeAutoshapeFlowchartManualOperation = '\000j\000H',
	ExcelMsoAutoShapeTypeAutoshapeFlowchartConnector = '\000j\000I',
	ExcelMsoAutoShapeTypeAutoshapeFlowchartOffpageConnector = '\000j\000J',
	ExcelMsoAutoShapeTypeAutoshapeFlowchartCard = '\000j\000K',
	ExcelMsoAutoShapeTypeAutoshapeFlowchartPunchedTape = '\000j\000L',
	ExcelMsoAutoShapeTypeAutoshapeFlowchartSummingJunction = '\000j\000M',
	ExcelMsoAutoShapeTypeAutoshapeFlowchartOr = '\000j\000N',
	ExcelMsoAutoShapeTypeAutoshapeFlowchartCollate = '\000j\000O',
	ExcelMsoAutoShapeTypeAutoshapeFlowchartSort = '\000j\000P',
	ExcelMsoAutoShapeTypeAutoshapeFlowchartExtract = '\000j\000Q',
	ExcelMsoAutoShapeTypeAutoshapeFlowchartMerge = '\000j\000R',
	ExcelMsoAutoShapeTypeAutoshapeFlowchartStoredData = '\000j\000S',
	ExcelMsoAutoShapeTypeAutoshapeFlowchartDelay = '\000j\000T',
	ExcelMsoAutoShapeTypeAutoshapeFlowchartSequentialAccessStorage = '\000j\000U',
	ExcelMsoAutoShapeTypeAutoshapeFlowchartMagneticDisk = '\000j\000V',
	ExcelMsoAutoShapeTypeAutoshapeFlowchartDirectAccessStorage = '\000j\000W',
	ExcelMsoAutoShapeTypeAutoshapeFlowchartDisplay = '\000j\000X',
	ExcelMsoAutoShapeTypeAutoshapeExplosionOne = '\000j\000Y',
	ExcelMsoAutoShapeTypeAutoshapeExplosionTwo = '\000j\000Z',
	ExcelMsoAutoShapeTypeAutoshapeFourPointStar = '\000j\000[',
	ExcelMsoAutoShapeTypeAutoshapeFivePointStar = '\000j\000\\',
	ExcelMsoAutoShapeTypeAutoshapeEightPointStar = '\000j\000]',
	ExcelMsoAutoShapeTypeAutoshapeSixteenPointStar = '\000j\000^',
	ExcelMsoAutoShapeTypeAutoshapeTwentyFourPointStar = '\000j\000_',
	ExcelMsoAutoShapeTypeAutoshapeThirtyTwoPointStar = '\000j\000`',
	ExcelMsoAutoShapeTypeAutoshapeUpRibbon = '\000j\000a',
	ExcelMsoAutoShapeTypeAutoshapeDownRibbon = '\000j\000b',
	ExcelMsoAutoShapeTypeAutoshapeCurvedUpRibbon = '\000j\000c',
	ExcelMsoAutoShapeTypeAutoshapeCurvedDownRibbon = '\000j\000d',
	ExcelMsoAutoShapeTypeAutoshapeVerticalScroll = '\000j\000e',
	ExcelMsoAutoShapeTypeAutoshapeHorizontalScroll = '\000j\000f',
	ExcelMsoAutoShapeTypeAutoshapeWave = '\000j\000g',
	ExcelMsoAutoShapeTypeAutoshapeDoubleWave = '\000j\000h',
	ExcelMsoAutoShapeTypeAutoshapeRectangularCallout = '\000j\000i',
	ExcelMsoAutoShapeTypeAutoshapeRoundedRectangularCallout = '\000j\000j',
	ExcelMsoAutoShapeTypeAutoshapeOvalCallout = '\000j\000k',
	ExcelMsoAutoShapeTypeAutoshapeCloudCallout = '\000j\000l',
	ExcelMsoAutoShapeTypeAutoshapeLineCalloutOne = '\000j\000m',
	ExcelMsoAutoShapeTypeAutoshapeLineCalloutTwo = '\000j\000n',
	ExcelMsoAutoShapeTypeAutoshapeLineCalloutThree = '\000j\000o',
	ExcelMsoAutoShapeTypeAutoshapeLineCalloutFour = '\000j\000p',
	ExcelMsoAutoShapeTypeAutoshapeLineCalloutOneAccentBar = '\000j\000q',
	ExcelMsoAutoShapeTypeAutoshapeLineCalloutTwoAccentBar = '\000j\000r',
	ExcelMsoAutoShapeTypeAutoshapeLineCalloutThreeAccentBar = '\000j\000s',
	ExcelMsoAutoShapeTypeAutoshapeLineCalloutFourAccentBar = '\000j\000t',
	ExcelMsoAutoShapeTypeAutoshapeLineCalloutOneNoBorder = '\000j\000u',
	ExcelMsoAutoShapeTypeAutoshapeLineCalloutTwoNoBorder = '\000j\000v',
	ExcelMsoAutoShapeTypeAutoshapeLineCalloutThreeNoBorder = '\000j\000w',
	ExcelMsoAutoShapeTypeAutoshapeLineCalloutFourNoBorder = '\000j\000x',
	ExcelMsoAutoShapeTypeAutoshapeCalloutOneBorderAndAccentBar = '\000j\000y',
	ExcelMsoAutoShapeTypeAutoshapeCalloutTwoBorderAndAccentBar = '\000j\000z',
	ExcelMsoAutoShapeTypeAutoshapeCalloutThreeBorderAndAccentBar = '\000j\000{',
	ExcelMsoAutoShapeTypeAutoshapeCalloutFourBorderAndAccentBar = '\000j\000|',
	ExcelMsoAutoShapeTypeAutoshapeActionButtonCustom = '\000j\000}',
	ExcelMsoAutoShapeTypeAutoshapeActionButtonHome = '\000j\000~',
	ExcelMsoAutoShapeTypeAutoshapeActionButtonHelp = '\000j\000\177',
	ExcelMsoAutoShapeTypeAutoshapeActionButtonInformation = '\000j\000\200',
	ExcelMsoAutoShapeTypeAutoshapeActionButtonBackOrPrevious = '\000j\000\201',
	ExcelMsoAutoShapeTypeAutoshapeActionButtonForwardOrNext = '\000j\000\202',
	ExcelMsoAutoShapeTypeAutoshapeActionButtonBeginning = '\000j\000\203',
	ExcelMsoAutoShapeTypeAutoshapeActionButtonEnd = '\000j\000\204',
	ExcelMsoAutoShapeTypeAutoshapeActionButtonReturn = '\000j\000\205',
	ExcelMsoAutoShapeTypeAutoshapeActionButtonDocument = '\000j\000\206',
	ExcelMsoAutoShapeTypeAutoshapeActionButtonSound = '\000j\000\207',
	ExcelMsoAutoShapeTypeAutoshapeActionButtonMovie = '\000j\000\210',
	ExcelMsoAutoShapeTypeAutoshapeBalloon = '\000j\000\211',
	ExcelMsoAutoShapeTypeAutoshapeNotPrimitive = '\000j\000\212',
	ExcelMsoAutoShapeTypeAutoshapeFlowchartOfflineStorage = '\000j\000\213',
	ExcelMsoAutoShapeTypeAutoshapeLeftRightRibbon = '\000j\000\214',
	ExcelMsoAutoShapeTypeAutoshapeDiagonalStripe = '\000j\000\215',
	ExcelMsoAutoShapeTypeAutoshapePie = '\000j\000\216',
	ExcelMsoAutoShapeTypeAutoshapeNonIsoscelesTrapezoid = '\000j\000\217',
	ExcelMsoAutoShapeTypeAutoshapeDecagon = '\000j\000\220',
	ExcelMsoAutoShapeTypeAutoshapeHeptagon = '\000j\000\221',
	ExcelMsoAutoShapeTypeAutoshapeDodecagon = '\000j\000\222',
	ExcelMsoAutoShapeTypeAutoshapeSixPointsStar = '\000j\000\223',
	ExcelMsoAutoShapeTypeAutoshapeSevenPointsStar = '\000j\000\224',
	ExcelMsoAutoShapeTypeAutoshapeTenPointsStar = '\000j\000\225',
	ExcelMsoAutoShapeTypeAutoshapeTwelvePointsStar = '\000j\000\226',
	ExcelMsoAutoShapeTypeAutoshapeRoundOneRectangle = '\000j\000\227',
	ExcelMsoAutoShapeTypeAutoshapeRoundTwoSameRectangle = '\000j\000\230',
	ExcelMsoAutoShapeTypeAutoshapeRoundTwoDiagonalRectangle = '\000j\000\231',
	ExcelMsoAutoShapeTypeAutoshapeSnipRoundRectangle = '\000j\000\232',
	ExcelMsoAutoShapeTypeAutoshapeSnipOneRectangle = '\000j\000\233',
	ExcelMsoAutoShapeTypeAutoshapeSnipTwoSameRectangle = '\000j\000\234',
	ExcelMsoAutoShapeTypeAutoshapeSnipTwoDiagonalRectangle = '\000j\000\235',
	ExcelMsoAutoShapeTypeAutoshapeFrame = '\000j\000\236',
	ExcelMsoAutoShapeTypeAutoshapeHalfFrame = '\000j\000\237',
	ExcelMsoAutoShapeTypeAutoshapeTear = '\000j\000\240',
	ExcelMsoAutoShapeTypeAutoshapeChord = '\000j\000\241',
	ExcelMsoAutoShapeTypeAutoshapeCorner = '\000j\000\242',
	ExcelMsoAutoShapeTypeAutoshapeMathPlus = '\000j\000\243',
	ExcelMsoAutoShapeTypeAutoshapeMathMinus = '\000j\000\244',
	ExcelMsoAutoShapeTypeAutoshapeMathMultiply = '\000j\000\245',
	ExcelMsoAutoShapeTypeAutoshapeMathDivide = '\000j\000\246',
	ExcelMsoAutoShapeTypeAutoshapeMathEqual = '\000j\000\247',
	ExcelMsoAutoShapeTypeAutoshapeMathNotEqual = '\000j\000\250',
	ExcelMsoAutoShapeTypeAutoshapeCornerTabs = '\000j\000\251',
	ExcelMsoAutoShapeTypeAutoshapeSquareTabs = '\000j\000\252',
	ExcelMsoAutoShapeTypeAutoshapePlaqueTabs = '\000j\000\253',
	ExcelMsoAutoShapeTypeAutoshapeGearSix = '\000j\000\254',
	ExcelMsoAutoShapeTypeAutoshapeGearNine = '\000j\000\255',
	ExcelMsoAutoShapeTypeAutoshapeFunnel = '\000j\000\256',
	ExcelMsoAutoShapeTypeAutoshapePieWedge = '\000j\000\257',
	ExcelMsoAutoShapeTypeAutoshapeLeftCircularArrow = '\000j\000\260',
	ExcelMsoAutoShapeTypeAutoshapeLeftRightCircularArrow = '\000j\000\261',
	ExcelMsoAutoShapeTypeAutoshapeSwooshArrow = '\000j\000\262',
	ExcelMsoAutoShapeTypeAutoshapeCloud = '\000j\000\263',
	ExcelMsoAutoShapeTypeAutoshapeChartX = '\000j\000\264',
	ExcelMsoAutoShapeTypeAutoshapeChartStar = '\000j\000\265',
	ExcelMsoAutoShapeTypeAutoshapeChartPlus = '\000j\000\266',
	ExcelMsoAutoShapeTypeAutoshapeLineInverse = '\000j\000\267'
};
typedef enum ExcelMsoAutoShapeType ExcelMsoAutoShapeType;

enum ExcelMsoShapeType {
	ExcelMsoShapeTypeShapeTypeUnset = '\000\213\377\376',
	ExcelMsoShapeTypeShapeTypeAuto = '\000\214\000\001',
	ExcelMsoShapeTypeShapeTypeCallout = '\000\214\000\002',
	ExcelMsoShapeTypeShapeTypeChart = '\000\214\000\003',
	ExcelMsoShapeTypeShapeTypeComment = '\000\214\000\004',
	ExcelMsoShapeTypeShapeTypeFreeForm = '\000\214\000\005',
	ExcelMsoShapeTypeShapeTypeGroup = '\000\214\000\006',
	ExcelMsoShapeTypeShapeTypeEmbeddedOLEControl = '\000\214\000\007',
	ExcelMsoShapeTypeShapeTypeFormControl = '\000\214\000\010',
	ExcelMsoShapeTypeShapeTypeLine = '\000\214\000\011',
	ExcelMsoShapeTypeShapeTypeLinkedOLEObject = '\000\214\000\012',
	ExcelMsoShapeTypeShapeTypeLinkedPicture = '\000\214\000\013',
	ExcelMsoShapeTypeShapeTypeOLEControl = '\000\214\000\014',
	ExcelMsoShapeTypeShapeTypePicture = '\000\214\000\015',
	ExcelMsoShapeTypeShapeTypePlaceHolder = '\000\214\000\016',
	ExcelMsoShapeTypeShapeTypeWordArt = '\000\214\000\017',
	ExcelMsoShapeTypeShapeTypeMedia = '\000\214\000\020',
	ExcelMsoShapeTypeShapeTypeTextBox = '\000\214\000\021',
	ExcelMsoShapeTypeShapeTypeScriptAnchor = '\000\214\000\022',
	ExcelMsoShapeTypeShapeTypeTable = '\000\214\000\023',
	ExcelMsoShapeTypeShapeTypeCanvas = '\000\214\000\024',
	ExcelMsoShapeTypeShapeTypeDiagram = '\000\214\000\025',
	ExcelMsoShapeTypeShapeTypeInk = '\000\214\000\026',
	ExcelMsoShapeTypeShapeTypeInkComment = '\000\214\000\027',
	ExcelMsoShapeTypeShapeTypeSmartartGraphic = '\000\214\000\030',
	ExcelMsoShapeTypeShapeTypeSlicer = '\000\214\000\031',
	ExcelMsoShapeTypeShapeTypeWebVideo = '\000\214\000\032',
	ExcelMsoShapeTypeShapeTypeContentApplication = '\000\214\000\033',
	ExcelMsoShapeTypeShapeTypeGraphic = '\000\214\000\034',
	ExcelMsoShapeTypeShapeTypeLinkedGraphic = '\000\214\000\035',
	ExcelMsoShapeTypeShapeType3dModel = '\000\214\000\036',
	ExcelMsoShapeTypeShapeTypeLinked3dModel = '\000\214\000\037'
};
typedef enum ExcelMsoShapeType ExcelMsoShapeType;

enum ExcelMsoColorType {
	ExcelMsoColorTypeColorTypeUnset = '\000j\377\376',
	ExcelMsoColorTypeRGB = '\000k\000\001',
	ExcelMsoColorTypeScheme = '\000k\000\002'
};
typedef enum ExcelMsoColorType ExcelMsoColorType;

enum ExcelMsoPictureColorType {
	ExcelMsoPictureColorTypePictureColorTypeUnset = '\000\265\377\376',
	ExcelMsoPictureColorTypePictureColorAutomatic = '\000\266\000\001',
	ExcelMsoPictureColorTypePictureColorGrayScale = '\000\266\000\002',
	ExcelMsoPictureColorTypePictureColorBlackAndWhite = '\000\266\000\003',
	ExcelMsoPictureColorTypePictureColorWatermark = '\000\266\000\004'
};
typedef enum ExcelMsoPictureColorType ExcelMsoPictureColorType;

enum ExcelMsoCalloutAngleType {
	ExcelMsoCalloutAngleTypeAngleUnset = '\000k\377\376',
	ExcelMsoCalloutAngleTypeAngleAutomatic = '\000l\000\001',
	ExcelMsoCalloutAngleTypeAngle30 = '\000l\000\002',
	ExcelMsoCalloutAngleTypeAngle45 = '\000l\000\003',
	ExcelMsoCalloutAngleTypeAngle60 = '\000l\000\004',
	ExcelMsoCalloutAngleTypeAngle90 = '\000l\000\005'
};
typedef enum ExcelMsoCalloutAngleType ExcelMsoCalloutAngleType;

enum ExcelMsoCalloutDropType {
	ExcelMsoCalloutDropTypeDropUnset = '\000l\377\376',
	ExcelMsoCalloutDropTypeDropCustom = '\000m\000\001',
	ExcelMsoCalloutDropTypeDropTop = '\000m\000\002',
	ExcelMsoCalloutDropTypeDropCenter = '\000m\000\003',
	ExcelMsoCalloutDropTypeDropBottom = '\000m\000\004'
};
typedef enum ExcelMsoCalloutDropType ExcelMsoCalloutDropType;

enum ExcelMsoCalloutType {
	ExcelMsoCalloutTypeCalloutUnset = '\000m\377\376',
	ExcelMsoCalloutTypeCalloutOne = '\000n\000\001',
	ExcelMsoCalloutTypeCalloutTwo = '\000n\000\002',
	ExcelMsoCalloutTypeCalloutThree = '\000n\000\003',
	ExcelMsoCalloutTypeCalloutFour = '\000n\000\004'
};
typedef enum ExcelMsoCalloutType ExcelMsoCalloutType;

enum ExcelMsoTextOrientation {
	ExcelMsoTextOrientationTextOrientationUnset = '\000\215\377\376',
	ExcelMsoTextOrientationHorizontal = '\000\216\000\001',
	ExcelMsoTextOrientationUpward = '\000\216\000\002',
	ExcelMsoTextOrientationDownward = '\000\216\000\003',
	ExcelMsoTextOrientationVerticalEastAsian = '\000\216\000\004',
	ExcelMsoTextOrientationVertical = '\000\216\000\005',
	ExcelMsoTextOrientationHorizontalRotatedEastAsian = '\000\216\000\006'
};
typedef enum ExcelMsoTextOrientation ExcelMsoTextOrientation;

enum ExcelMsoScaleFrom {
	ExcelMsoScaleFromScaleFromTopLeft = '\000o\000\000',
	ExcelMsoScaleFromScaleFromMiddle = '\000o\000\001',
	ExcelMsoScaleFromScaleFromBottomRight = '\000o\000\002'
};
typedef enum ExcelMsoScaleFrom ExcelMsoScaleFrom;

enum ExcelMsoPresetCamera {
	ExcelMsoPresetCameraPresetCameraUnset = '\000\256\377\376',
	ExcelMsoPresetCameraCameraLegacyObliqueFromTopLeft = '\000\257\000\001',
	ExcelMsoPresetCameraCameraLegacyObliqueFromTop = '\000\257\000\002',
	ExcelMsoPresetCameraCameraLegacyObliqueFromTopright = '\000\257\000\003',
	ExcelMsoPresetCameraCameraLegacyObliqueFromLeft = '\000\257\000\004',
	ExcelMsoPresetCameraCameraLegacyObliqueFromFront = '\000\257\000\005',
	ExcelMsoPresetCameraCameraLegacyObliqueFromRight = '\000\257\000\006',
	ExcelMsoPresetCameraCameraLegacyObliqueFromBottomLeft = '\000\257\000\007',
	ExcelMsoPresetCameraCameraLegacyObliqueFromBottom = '\000\257\000\010',
	ExcelMsoPresetCameraCameraLegacyObliqueFromBottomRight = '\000\257\000\011',
	ExcelMsoPresetCameraCameraLegacyPerspectiveFromTopLeft = '\000\257\000\012',
	ExcelMsoPresetCameraCameraLegacyPerspectiveFromTop = '\000\257\000\013',
	ExcelMsoPresetCameraCameraLegacyPerspectiveFromTopRight = '\000\257\000\014',
	ExcelMsoPresetCameraCameraLegacyPerspectiveFromLeft = '\000\257\000\015',
	ExcelMsoPresetCameraCameraLegacyPerspectiveFromFront = '\000\257\000\016',
	ExcelMsoPresetCameraCameraLegacyPerspectiveFromRight = '\000\257\000\017',
	ExcelMsoPresetCameraCameraLegacyPerspectiveFromBottomLeft = '\000\257\000\020',
	ExcelMsoPresetCameraCameraLegacyPerspectiveFromBottom = '\000\257\000\021',
	ExcelMsoPresetCameraCameraLegacyPerspectiveFromBottomRight = '\000\257\000\022',
	ExcelMsoPresetCameraCameraOrthographic = '\000\257\000\023',
	ExcelMsoPresetCameraCameraIsometricFromTopUp = '\000\257\000\024',
	ExcelMsoPresetCameraCameraIsometricFromTopDown = '\000\257\000\025',
	ExcelMsoPresetCameraCameraIsometricFromBottomUp = '\000\257\000\026',
	ExcelMsoPresetCameraCameraIsometricFromBottomDown = '\000\257\000\027',
	ExcelMsoPresetCameraCameraIsometricFromLeftUp = '\000\257\000\030',
	ExcelMsoPresetCameraCameraIsometricFromLeftDown = '\000\257\000\031',
	ExcelMsoPresetCameraCameraIsometricFromRightUp = '\000\257\000\032',
	ExcelMsoPresetCameraCameraIsometricFromRightDown = '\000\257\000\033',
	ExcelMsoPresetCameraCameraIsometricOffAxis1FromLeft = '\000\257\000\034',
	ExcelMsoPresetCameraCameraIsometricOffAxis1FromRight = '\000\257\000\035',
	ExcelMsoPresetCameraCameraIsometricOffAxis1FromTop = '\000\257\000\036',
	ExcelMsoPresetCameraCameraIsometricOffAxis2FromLeft = '\000\257\000\037',
	ExcelMsoPresetCameraCameraIsometricOffAxis2FromRight = '\000\257\000 ',
	ExcelMsoPresetCameraCameraIsometricOffAxis2FromTop = '\000\257\000!',
	ExcelMsoPresetCameraCameraIsometricOffAxis3FromLeft = '\000\257\000\"',
	ExcelMsoPresetCameraCameraIsometricOffAxis3FromRight = '\000\257\000#',
	ExcelMsoPresetCameraCameraIsometricOffAxis3FromBottom = '\000\257\000$',
	ExcelMsoPresetCameraCameraIsometricOffAxis4FromLeft = '\000\257\000%',
	ExcelMsoPresetCameraCameraIsometricOffAxis4FromRight = '\000\257\000&',
	ExcelMsoPresetCameraCameraIsometricOffAxis4FromBottom = '\000\257\000\'',
	ExcelMsoPresetCameraCameraObliqueFromTopLeft = '\000\257\000(',
	ExcelMsoPresetCameraCameraObliqueFromTop = '\000\257\000)',
	ExcelMsoPresetCameraCameraObliqueFromTopRight = '\000\257\000*',
	ExcelMsoPresetCameraCameraObliqueFromLeft = '\000\257\000+',
	ExcelMsoPresetCameraCameraObliqueFromRight = '\000\257\000,',
	ExcelMsoPresetCameraCameraObliqueFromBottomLeft = '\000\257\000-',
	ExcelMsoPresetCameraCameraObliqueFromBottom = '\000\257\000.',
	ExcelMsoPresetCameraCameraObliqueFromBottomRight = '\000\257\000/',
	ExcelMsoPresetCameraCameraPerspectiveFromFront = '\000\257\0000',
	ExcelMsoPresetCameraCameraPerspectiveFromLeft = '\000\257\0001',
	ExcelMsoPresetCameraCameraPerspectiveFromRight = '\000\257\0002',
	ExcelMsoPresetCameraCameraPerspectiveFromAbove = '\000\257\0003',
	ExcelMsoPresetCameraCameraPerspectiveFromBelow = '\000\257\0004',
	ExcelMsoPresetCameraCameraPerspectiveFromAboveFacingLeft = '\000\257\0005',
	ExcelMsoPresetCameraCameraPerspectiveFromAboveFacingRight = '\000\257\0006',
	ExcelMsoPresetCameraCameraPerspectiveContrastingFacingLeft = '\000\257\0007',
	ExcelMsoPresetCameraCameraPerspectiveContrastingFacingRight = '\000\257\0008',
	ExcelMsoPresetCameraCameraPerspectiveHeroicFacingLeft = '\000\257\0009',
	ExcelMsoPresetCameraCameraPerspectiveHeroicFacingRight = '\000\257\000:',
	ExcelMsoPresetCameraCameraPerspectiveHeroicExtremeFacingLeft = '\000\257\000;',
	ExcelMsoPresetCameraCameraPerspectiveHeroicExtremeFacingRight = '\000\257\000<',
	ExcelMsoPresetCameraCameraPerspectiveRelaxed = '\000\257\000=',
	ExcelMsoPresetCameraCameraPerspectiveRelaxedModerately = '\000\257\000>'
};
typedef enum ExcelMsoPresetCamera ExcelMsoPresetCamera;

enum ExcelMsoLightRigType {
	ExcelMsoLightRigTypeLightRigUnset = '\000\257\377\376',
	ExcelMsoLightRigTypeLightRigFlat1 = '\000\260\000\001',
	ExcelMsoLightRigTypeLightRigFlat2 = '\000\260\000\002',
	ExcelMsoLightRigTypeLightRigFlat3 = '\000\260\000\003',
	ExcelMsoLightRigTypeLightRigFlat4 = '\000\260\000\004',
	ExcelMsoLightRigTypeLightRigNormal1 = '\000\260\000\005',
	ExcelMsoLightRigTypeLightRigNormal2 = '\000\260\000\006',
	ExcelMsoLightRigTypeLightRigNormal3 = '\000\260\000\007',
	ExcelMsoLightRigTypeLightRigNormal4 = '\000\260\000\010',
	ExcelMsoLightRigTypeLightRigHarsh1 = '\000\260\000\011',
	ExcelMsoLightRigTypeLightRigHarsh2 = '\000\260\000\012',
	ExcelMsoLightRigTypeLightRigHarsh3 = '\000\260\000\013',
	ExcelMsoLightRigTypeLightRigHarsh4 = '\000\260\000\014',
	ExcelMsoLightRigTypeLightRigThreePoint = '\000\260\000\015',
	ExcelMsoLightRigTypeLightRigBalanced = '\000\260\000\016',
	ExcelMsoLightRigTypeLightRigSoft = '\000\260\000\017',
	ExcelMsoLightRigTypeLightRigHarsh = '\000\260\000\020',
	ExcelMsoLightRigTypeLightRigFlood = '\000\260\000\021',
	ExcelMsoLightRigTypeLightRigContrasting = '\000\260\000\022',
	ExcelMsoLightRigTypeLightRigMorning = '\000\260\000\023',
	ExcelMsoLightRigTypeLightRigSunrise = '\000\260\000\024',
	ExcelMsoLightRigTypeLightRigSunset = '\000\260\000\025',
	ExcelMsoLightRigTypeLightRigChilly = '\000\260\000\026',
	ExcelMsoLightRigTypeLightRigFreezing = '\000\260\000\027',
	ExcelMsoLightRigTypeLightRigFlat = '\000\260\000\030',
	ExcelMsoLightRigTypeLightRigTwoPoint = '\000\260\000\031',
	ExcelMsoLightRigTypeLightRigGlow = '\000\260\000\032',
	ExcelMsoLightRigTypeLightRigBrightRoom = '\000\260\000\033'
};
typedef enum ExcelMsoLightRigType ExcelMsoLightRigType;

enum ExcelMsoBevelType {
	ExcelMsoBevelTypeBevelTypeUnset = '\000\260\377\376',
	ExcelMsoBevelTypeBevelNone = '\000\261\000\001',
	ExcelMsoBevelTypeBevelRelaxedInset = '\000\261\000\002',
	ExcelMsoBevelTypeBevelCircle = '\000\261\000\003',
	ExcelMsoBevelTypeBevelSlope = '\000\261\000\004',
	ExcelMsoBevelTypeBevelCross = '\000\261\000\005',
	ExcelMsoBevelTypeBevelAngle = '\000\261\000\006',
	ExcelMsoBevelTypeBevelSoftRound = '\000\261\000\007',
	ExcelMsoBevelTypeBevelConvex = '\000\261\000\010',
	ExcelMsoBevelTypeBevelCoolSlant = '\000\261\000\011',
	ExcelMsoBevelTypeBevelDivot = '\000\261\000\012',
	ExcelMsoBevelTypeBevelRiblet = '\000\261\000\013',
	ExcelMsoBevelTypeBevelHardEdge = '\000\261\000\014',
	ExcelMsoBevelTypeBevelArtDeco = '\000\261\000\015'
};
typedef enum ExcelMsoBevelType ExcelMsoBevelType;

enum ExcelMsoShadowStyle {
	ExcelMsoShadowStyleShadowStyleUnset = '\000\261\377\376',
	ExcelMsoShadowStyleShadowStyleInner = '\000\262\000\001',
	ExcelMsoShadowStyleShadowStyleOuter = '\000\262\000\002'
};
typedef enum ExcelMsoShadowStyle ExcelMsoShadowStyle;

enum ExcelMsoParagraphAlignment {
	ExcelMsoParagraphAlignmentParagraphAlignmentUnset = '\000\346\377\376',
	ExcelMsoParagraphAlignmentParagraphAlignLeft = '\000\347\000\001',
	ExcelMsoParagraphAlignmentParagraphAlignCenter = '\000\347\000\002',
	ExcelMsoParagraphAlignmentParagraphAlignRight = '\000\347\000\003',
	ExcelMsoParagraphAlignmentParagraphAlignJustify = '\000\347\000\004',
	ExcelMsoParagraphAlignmentParagraphAlignDistribute = '\000\347\000\005',
	ExcelMsoParagraphAlignmentParagraphAlignThai = '\000\347\000\006',
	ExcelMsoParagraphAlignmentParagraphAlignJustifyLow = '\000\347\000\007'
};
typedef enum ExcelMsoParagraphAlignment ExcelMsoParagraphAlignment;

enum ExcelMsoTextStrike {
	ExcelMsoTextStrikeStrikeUnset = '\000\263\377\376',
	ExcelMsoTextStrikeNoStrike = '\000\264\000\000',
	ExcelMsoTextStrikeSingleStrike = '\000\264\000\001',
	ExcelMsoTextStrikeDoubleStrike = '\000\264\000\002'
};
typedef enum ExcelMsoTextStrike ExcelMsoTextStrike;

enum ExcelMsoTextCaps {
	ExcelMsoTextCapsCapsUnset = '\000\264\377\376',
	ExcelMsoTextCapsNoCaps = '\000\265\000\000',
	ExcelMsoTextCapsSmallCaps = '\000\265\000\001',
	ExcelMsoTextCapsAllCaps = '\000\265\000\002'
};
typedef enum ExcelMsoTextCaps ExcelMsoTextCaps;

enum ExcelMsoTextUnderlineType {
	ExcelMsoTextUnderlineTypeUnderlineUnset = '\003\356\377\376',
	ExcelMsoTextUnderlineTypeNoUnderline = '\003\357\000\000',
	ExcelMsoTextUnderlineTypeUnderlineWordsOnly = '\003\357\000\001',
	ExcelMsoTextUnderlineTypeUnderlineSingleLine = '\003\357\000\002',
	ExcelMsoTextUnderlineTypeUnderlineDoubleLine = '\003\357\000\003',
	ExcelMsoTextUnderlineTypeUnderlineHeavyLine = '\003\357\000\004',
	ExcelMsoTextUnderlineTypeUnderlineDottedLine = '\003\357\000\005',
	ExcelMsoTextUnderlineTypeUnderlineHeavyDottedLine = '\003\357\000\006',
	ExcelMsoTextUnderlineTypeUnderlineDashLine = '\003\357\000\007',
	ExcelMsoTextUnderlineTypeUnderlineHeavyDashLine = '\003\357\000\010',
	ExcelMsoTextUnderlineTypeUnderlineLongDashLine = '\003\357\000\011',
	ExcelMsoTextUnderlineTypeUnderlineHeavyLongDashLine = '\003\357\000\012',
	ExcelMsoTextUnderlineTypeUnderlineDotDashLine = '\003\357\000\013',
	ExcelMsoTextUnderlineTypeUnderlineHeavyDotDashLine = '\003\357\000\014',
	ExcelMsoTextUnderlineTypeUnderlineDotDotDashLine = '\003\357\000\015',
	ExcelMsoTextUnderlineTypeUnderlineHeavyDotDotDashLine = '\003\357\000\016',
	ExcelMsoTextUnderlineTypeUnderlineWavyLine = '\003\357\000\017',
	ExcelMsoTextUnderlineTypeUnderlineHeavyWavyLine = '\003\357\000\020',
	ExcelMsoTextUnderlineTypeUnderlineWavyDoubleLine = '\003\357\000\021'
};
typedef enum ExcelMsoTextUnderlineType ExcelMsoTextUnderlineType;

enum ExcelMsoTextTabAlign {
	ExcelMsoTextTabAlignTabUnset = '\000\266\377\376',
	ExcelMsoTextTabAlignLeftTab = '\000\267\000\000',
	ExcelMsoTextTabAlignCenterTab = '\000\267\000\001',
	ExcelMsoTextTabAlignRightTab = '\000\267\000\002',
	ExcelMsoTextTabAlignDecimalTab = '\000\267\000\003'
};
typedef enum ExcelMsoTextTabAlign ExcelMsoTextTabAlign;

enum ExcelMsoTextCharWrap {
	ExcelMsoTextCharWrapCharacterWrapUnset = '\000\267\377\376',
	ExcelMsoTextCharWrapNoCharacterWrap = '\000\270\000\000',
	ExcelMsoTextCharWrapStandardCharacterWrap = '\000\270\000\001',
	ExcelMsoTextCharWrapStrictCharacterWrap = '\000\270\000\002',
	ExcelMsoTextCharWrapCustomCharacterWrap = '\000\270\000\003'
};
typedef enum ExcelMsoTextCharWrap ExcelMsoTextCharWrap;

enum ExcelMsoTextFontAlign {
	ExcelMsoTextFontAlignFontAlignmentUnset = '\000\270\377\376',
	ExcelMsoTextFontAlignAutomaticAlignment = '\000\271\000\000',
	ExcelMsoTextFontAlignTopAlignment = '\000\271\000\001',
	ExcelMsoTextFontAlignCenterAlignment = '\000\271\000\002',
	ExcelMsoTextFontAlignBaselineAlignment = '\000\271\000\003',
	ExcelMsoTextFontAlignBottomAlignment = '\000\271\000\004'
};
typedef enum ExcelMsoTextFontAlign ExcelMsoTextFontAlign;

enum ExcelMsoAutoSize {
	ExcelMsoAutoSizeAutoSizeUnset = '\000\344\377\376',
	ExcelMsoAutoSizeAutoSizeNone = '\000\345\000\000',
	ExcelMsoAutoSizeShapeToFitText = '\000\345\000\001',
	ExcelMsoAutoSizeTextToFitShape = '\000\345\000\002'
};
typedef enum ExcelMsoAutoSize ExcelMsoAutoSize;

enum ExcelMsoPathFormat {
	ExcelMsoPathFormatPathTypeUnset = '\000\272\377\376',
	ExcelMsoPathFormatNoPathType = '\000\273\000\000',
	ExcelMsoPathFormatPathType1 = '\000\273\000\001',
	ExcelMsoPathFormatPathType2 = '\000\273\000\002',
	ExcelMsoPathFormatPathType3 = '\000\273\000\003',
	ExcelMsoPathFormatPathType4 = '\000\273\000\004'
};
typedef enum ExcelMsoPathFormat ExcelMsoPathFormat;

enum ExcelMsoWarpFormat {
	ExcelMsoWarpFormatWarpFormatUnset = '\000\273\377\376',
	ExcelMsoWarpFormatWarpFormat1 = '\000\274\000\000',
	ExcelMsoWarpFormatWarpFormat2 = '\000\274\000\001',
	ExcelMsoWarpFormatWarpFormat3 = '\000\274\000\002',
	ExcelMsoWarpFormatWarpFormat4 = '\000\274\000\003',
	ExcelMsoWarpFormatWarpFormat5 = '\000\274\000\004',
	ExcelMsoWarpFormatWarpFormat6 = '\000\274\000\005',
	ExcelMsoWarpFormatWarpFormat7 = '\000\274\000\006',
	ExcelMsoWarpFormatWarpFormat8 = '\000\274\000\007',
	ExcelMsoWarpFormatWarpFormat9 = '\000\274\000\010',
	ExcelMsoWarpFormatWarpFormat10 = '\000\274\000\011',
	ExcelMsoWarpFormatWarpFormat11 = '\000\274\000\012',
	ExcelMsoWarpFormatWarpFormat12 = '\000\274\000\013',
	ExcelMsoWarpFormatWarpFormat13 = '\000\274\000\014',
	ExcelMsoWarpFormatWarpFormat14 = '\000\274\000\015',
	ExcelMsoWarpFormatWarpFormat15 = '\000\274\000\016',
	ExcelMsoWarpFormatWarpFormat16 = '\000\274\000\017',
	ExcelMsoWarpFormatWarpFormat17 = '\000\274\000\020',
	ExcelMsoWarpFormatWarpFormat18 = '\000\274\000\021',
	ExcelMsoWarpFormatWarpFormat19 = '\000\274\000\022',
	ExcelMsoWarpFormatWarpFormat20 = '\000\274\000\023',
	ExcelMsoWarpFormatWarpFormat21 = '\000\274\000\024',
	ExcelMsoWarpFormatWarpFormat22 = '\000\274\000\025',
	ExcelMsoWarpFormatWarpFormat23 = '\000\274\000\026',
	ExcelMsoWarpFormatWarpFormat24 = '\000\274\000\027',
	ExcelMsoWarpFormatWarpFormat25 = '\000\274\000\030',
	ExcelMsoWarpFormatWarpFormat26 = '\000\274\000\031',
	ExcelMsoWarpFormatWarpFormat27 = '\000\274\000\032',
	ExcelMsoWarpFormatWarpFormat28 = '\000\274\000\033',
	ExcelMsoWarpFormatWarpFormat29 = '\000\274\000\034',
	ExcelMsoWarpFormatWarpFormat30 = '\000\274\000\035',
	ExcelMsoWarpFormatWarpFormat31 = '\000\274\000\036',
	ExcelMsoWarpFormatWarpFormat32 = '\000\274\000\037',
	ExcelMsoWarpFormatWarpFormat33 = '\000\274\000 ',
	ExcelMsoWarpFormatWarpFormat34 = '\000\274\000!',
	ExcelMsoWarpFormatWarpFormat35 = '\000\274\000\"',
	ExcelMsoWarpFormatWarpFormat36 = '\000\274\000#',
	ExcelMsoWarpFormatWarpFormat37 = '\000\274\000$'
};
typedef enum ExcelMsoWarpFormat ExcelMsoWarpFormat;

enum ExcelMsoTextChangeCase {
	ExcelMsoTextChangeCaseCaseSentence = '\000\344\000\001',
	ExcelMsoTextChangeCaseCaseLower = '\000\344\000\002',
	ExcelMsoTextChangeCaseCaseUpper = '\000\344\000\003',
	ExcelMsoTextChangeCaseCaseTitle = '\000\344\000\004',
	ExcelMsoTextChangeCaseCaseToggle = '\000\344\000\005'
};
typedef enum ExcelMsoTextChangeCase ExcelMsoTextChangeCase;

enum ExcelMsoDateTimeFormat {
	ExcelMsoDateTimeFormatDateTimeFormatUnset = '\000\275\377\376',
	ExcelMsoDateTimeFormatDateTimeFormatMdyy = '\000\276\000\001',
	ExcelMsoDateTimeFormatDateTimeFormatDdddMMMMddyyyy = '\000\276\000\002',
	ExcelMsoDateTimeFormatDateTimeFormatDMMMMyyyy = '\000\276\000\003',
	ExcelMsoDateTimeFormatDateTimeFormatMMMMdyyyy = '\000\276\000\004',
	ExcelMsoDateTimeFormatDateTimeFormatDMMMyy = '\000\276\000\005',
	ExcelMsoDateTimeFormatDateTimeFormatMMMMyy = '\000\276\000\006',
	ExcelMsoDateTimeFormatDateTimeFormatMMyy = '\000\276\000\007',
	ExcelMsoDateTimeFormatDateTimeFormatMMddyyHmm = '\000\276\000\010',
	ExcelMsoDateTimeFormatDateTimeFormatMMddyyhmmAMPM = '\000\276\000\011',
	ExcelMsoDateTimeFormatDateTimeFormatHmm = '\000\276\000\012',
	ExcelMsoDateTimeFormatDateTimeFormatHmmss = '\000\276\000\013',
	ExcelMsoDateTimeFormatDateTimeFormatHmmAMPM = '\000\276\000\014',
	ExcelMsoDateTimeFormatDateTimeFormatHmmssAMPM = '\000\276\000\015',
	ExcelMsoDateTimeFormatDateTimeFormatFigureOut = '\000\276\000\016'
};
typedef enum ExcelMsoDateTimeFormat ExcelMsoDateTimeFormat;

enum ExcelMsoSoftEdgeType {
	ExcelMsoSoftEdgeTypeSoftEdgeUnset = '\000\277\377\376',
	ExcelMsoSoftEdgeTypeNoSoftEdge = '\000\300\000\000',
	ExcelMsoSoftEdgeTypeSoftEdgeType1 = '\000\300\000\001',
	ExcelMsoSoftEdgeTypeSoftEdgeType2 = '\000\300\000\002',
	ExcelMsoSoftEdgeTypeSoftEdgeType3 = '\000\300\000\003',
	ExcelMsoSoftEdgeTypeSoftEdgeType4 = '\000\300\000\004',
	ExcelMsoSoftEdgeTypeSoftEdgeType5 = '\000\300\000\005',
	ExcelMsoSoftEdgeTypeSoftEdgeType6 = '\000\300\000\006'
};
typedef enum ExcelMsoSoftEdgeType ExcelMsoSoftEdgeType;

enum ExcelMsoThemeColorSchemeIndex {
	ExcelMsoThemeColorSchemeIndexFirstDarkSchemeColor = '\000\301\000\001',
	ExcelMsoThemeColorSchemeIndexFirstLightSchemeColor = '\000\301\000\002',
	ExcelMsoThemeColorSchemeIndexSecondDarkSchemeColor = '\000\301\000\003',
	ExcelMsoThemeColorSchemeIndexSecondLightSchemeColor = '\000\301\000\004',
	ExcelMsoThemeColorSchemeIndexFirstAccentSchemeColor = '\000\301\000\005',
	ExcelMsoThemeColorSchemeIndexSecondAccentSchemeColor = '\000\301\000\006',
	ExcelMsoThemeColorSchemeIndexThirdAccentSchemeColor = '\000\301\000\007',
	ExcelMsoThemeColorSchemeIndexFourthAccentSchemeColor = '\000\301\000\010',
	ExcelMsoThemeColorSchemeIndexFifthAccentSchemeColor = '\000\301\000\011',
	ExcelMsoThemeColorSchemeIndexSixthAccentSchemeColor = '\000\301\000\012',
	ExcelMsoThemeColorSchemeIndexHyperlinkSchemeColor = '\000\301\000\013',
	ExcelMsoThemeColorSchemeIndexFollowedHyperlinkSchemeColor = '\000\301\000\014'
};
typedef enum ExcelMsoThemeColorSchemeIndex ExcelMsoThemeColorSchemeIndex;

enum ExcelMsoThemeColorIndex {
	ExcelMsoThemeColorIndexThemeColorUnset = '\000\301\377\376',
	ExcelMsoThemeColorIndexNoThemeColor = '\000\302\000\000',
	ExcelMsoThemeColorIndexFirstDarkThemeColor = '\000\302\000\001',
	ExcelMsoThemeColorIndexFirstLightThemeColor = '\000\302\000\002',
	ExcelMsoThemeColorIndexSecondDarkThemeColor = '\000\302\000\003',
	ExcelMsoThemeColorIndexSecondLightThemeColor = '\000\302\000\004',
	ExcelMsoThemeColorIndexFirstAccentThemeColor = '\000\302\000\005',
	ExcelMsoThemeColorIndexSecondAccentThemeColor = '\000\302\000\006',
	ExcelMsoThemeColorIndexThirdAccentThemeColor = '\000\302\000\007',
	ExcelMsoThemeColorIndexFourthAccentThemeColor = '\000\302\000\010',
	ExcelMsoThemeColorIndexFifthAccentThemeColor = '\000\302\000\011',
	ExcelMsoThemeColorIndexSixthAccentThemeColor = '\000\302\000\012',
	ExcelMsoThemeColorIndexHyperlinkThemeColor = '\000\302\000\013',
	ExcelMsoThemeColorIndexFollowedHyperlinkThemeColor = '\000\302\000\014',
	ExcelMsoThemeColorIndexFirstTextThemeColor = '\000\302\000\015',
	ExcelMsoThemeColorIndexFirstBackgroundThemeColor = '\000\302\000\016',
	ExcelMsoThemeColorIndexSecondTextThemeColor = '\000\302\000\017',
	ExcelMsoThemeColorIndexSecondBackgroundThemeColor = '\000\302\000\020'
};
typedef enum ExcelMsoThemeColorIndex ExcelMsoThemeColorIndex;

enum ExcelMsoFontLanguageIndex {
	ExcelMsoFontLanguageIndexThemeFontLatin = '\000\303\000\001',
	ExcelMsoFontLanguageIndexThemeFontComplexScript = '\000\303\000\002',
	ExcelMsoFontLanguageIndexThemeFontHighAnsi = '\000\303\000\003',
	ExcelMsoFontLanguageIndexThemeFontEastAsian = '\000\303\000\004'
};
typedef enum ExcelMsoFontLanguageIndex ExcelMsoFontLanguageIndex;

enum ExcelMsoShapeStyleIndex {
	ExcelMsoShapeStyleIndexShapeStyleUnset = '\000\303\377\376',
	ExcelMsoShapeStyleIndexShapeNotAPreset = '\000\304\000\000',
	ExcelMsoShapeStyleIndexShapePreset1 = '\000\304\000\001',
	ExcelMsoShapeStyleIndexShapePreset2 = '\000\304\000\002',
	ExcelMsoShapeStyleIndexShapePreset3 = '\000\304\000\003',
	ExcelMsoShapeStyleIndexShapePreset4 = '\000\304\000\004',
	ExcelMsoShapeStyleIndexShapePreset5 = '\000\304\000\005',
	ExcelMsoShapeStyleIndexShapePreset6 = '\000\304\000\006',
	ExcelMsoShapeStyleIndexShapePreset7 = '\000\304\000\007',
	ExcelMsoShapeStyleIndexShapePreset8 = '\000\304\000\010',
	ExcelMsoShapeStyleIndexShapePreset9 = '\000\304\000\011',
	ExcelMsoShapeStyleIndexShapePreset10 = '\000\304\000\012',
	ExcelMsoShapeStyleIndexShapePreset11 = '\000\304\000\013',
	ExcelMsoShapeStyleIndexShapePreset12 = '\000\304\000\014',
	ExcelMsoShapeStyleIndexShapePreset13 = '\000\304\000\015',
	ExcelMsoShapeStyleIndexShapePreset14 = '\000\304\000\016',
	ExcelMsoShapeStyleIndexShapePreset15 = '\000\304\000\017',
	ExcelMsoShapeStyleIndexShapePreset16 = '\000\304\000\020',
	ExcelMsoShapeStyleIndexShapePreset17 = '\000\304\000\021',
	ExcelMsoShapeStyleIndexShapePreset18 = '\000\304\000\022',
	ExcelMsoShapeStyleIndexShapePreset19 = '\000\304\000\023',
	ExcelMsoShapeStyleIndexShapePreset20 = '\000\304\000\024',
	ExcelMsoShapeStyleIndexShapePreset21 = '\000\304\000\025',
	ExcelMsoShapeStyleIndexShapePreset22 = '\000\304\000\026',
	ExcelMsoShapeStyleIndexShapePreset23 = '\000\304\000\027',
	ExcelMsoShapeStyleIndexShapePreset24 = '\000\304\000\030',
	ExcelMsoShapeStyleIndexShapePreset25 = '\000\304\000\031',
	ExcelMsoShapeStyleIndexShapePreset26 = '\000\304\000\032',
	ExcelMsoShapeStyleIndexShapePreset27 = '\000\304\000\033',
	ExcelMsoShapeStyleIndexShapePreset28 = '\000\304\000\034',
	ExcelMsoShapeStyleIndexShapePreset29 = '\000\304\000\035',
	ExcelMsoShapeStyleIndexShapePreset30 = '\000\304\000\036',
	ExcelMsoShapeStyleIndexShapePreset31 = '\000\304\000\037',
	ExcelMsoShapeStyleIndexShapePreset32 = '\000\304\000 ',
	ExcelMsoShapeStyleIndexShapePreset33 = '\000\304\000!',
	ExcelMsoShapeStyleIndexShapePreset34 = '\000\304\000\"',
	ExcelMsoShapeStyleIndexShapePreset35 = '\000\304\000#',
	ExcelMsoShapeStyleIndexShapePreset36 = '\000\304\000$',
	ExcelMsoShapeStyleIndexShapePreset37 = '\000\304\000%',
	ExcelMsoShapeStyleIndexShapePreset38 = '\000\304\000&',
	ExcelMsoShapeStyleIndexShapePreset39 = '\000\304\000\'',
	ExcelMsoShapeStyleIndexShapePreset40 = '\000\304\000(',
	ExcelMsoShapeStyleIndexShapePreset41 = '\000\304\000)',
	ExcelMsoShapeStyleIndexShapePreset42 = '\000\304\000*',
	ExcelMsoShapeStyleIndexLinePreset1 = '\000\304\'\021',
	ExcelMsoShapeStyleIndexLinePreset2 = '\000\304\'\022',
	ExcelMsoShapeStyleIndexLinePreset3 = '\000\304\'\023',
	ExcelMsoShapeStyleIndexLinePreset4 = '\000\304\'\024',
	ExcelMsoShapeStyleIndexLinePreset5 = '\000\304\'\025',
	ExcelMsoShapeStyleIndexLinePreset6 = '\000\304\'\026',
	ExcelMsoShapeStyleIndexLinePreset7 = '\000\304\'\027',
	ExcelMsoShapeStyleIndexLinePreset8 = '\000\304\'\030',
	ExcelMsoShapeStyleIndexLinePreset9 = '\000\304\'\031',
	ExcelMsoShapeStyleIndexLinePreset10 = '\000\304\'\032',
	ExcelMsoShapeStyleIndexLinePreset11 = '\000\304\'\033',
	ExcelMsoShapeStyleIndexLinePreset12 = '\000\304\'\034',
	ExcelMsoShapeStyleIndexLinePreset13 = '\000\304\'\035',
	ExcelMsoShapeStyleIndexLinePreset14 = '\000\304\'\036',
	ExcelMsoShapeStyleIndexLinePreset15 = '\000\304\'\037',
	ExcelMsoShapeStyleIndexLinePreset16 = '\000\304\' ',
	ExcelMsoShapeStyleIndexLinePreset17 = '\000\304\'!',
	ExcelMsoShapeStyleIndexLinePreset18 = '\000\304\'\"',
	ExcelMsoShapeStyleIndexLinePreset19 = '\000\304\'#',
	ExcelMsoShapeStyleIndexLinePreset20 = '\000\304\'$',
	ExcelMsoShapeStyleIndexLinePreset21 = '\000\304\'%'
};
typedef enum ExcelMsoShapeStyleIndex ExcelMsoShapeStyleIndex;

enum ExcelMsoBackgroundStyleIndex {
	ExcelMsoBackgroundStyleIndexBackgroundUnset = '\000\304\377\376',
	ExcelMsoBackgroundStyleIndexBackgroundNotAPreset = '\000\305\000\000',
	ExcelMsoBackgroundStyleIndexBackgroundPreset1 = '\000\305\000\001',
	ExcelMsoBackgroundStyleIndexBackgroundPreset2 = '\000\305\000\002',
	ExcelMsoBackgroundStyleIndexBackgroundPreset3 = '\000\305\000\003',
	ExcelMsoBackgroundStyleIndexBackgroundPreset4 = '\000\305\000\004',
	ExcelMsoBackgroundStyleIndexBackgroundPreset5 = '\000\305\000\005',
	ExcelMsoBackgroundStyleIndexBackgroundPreset6 = '\000\305\000\006',
	ExcelMsoBackgroundStyleIndexBackgroundPreset7 = '\000\305\000\007',
	ExcelMsoBackgroundStyleIndexBackgroundPreset8 = '\000\305\000\010',
	ExcelMsoBackgroundStyleIndexBackgroundPreset9 = '\000\305\000\011',
	ExcelMsoBackgroundStyleIndexBackgroundPreset10 = '\000\305\000\012',
	ExcelMsoBackgroundStyleIndexBackgroundPreset11 = '\000\305\000\013',
	ExcelMsoBackgroundStyleIndexBackgroundPreset12 = '\000\305\000\014'
};
typedef enum ExcelMsoBackgroundStyleIndex ExcelMsoBackgroundStyleIndex;

enum ExcelMsoTextDirection {
	ExcelMsoTextDirectionTextDirectionUnset = '\000\352\377\376',
	ExcelMsoTextDirectionLeftToRight = '\000\353\000\001',
	ExcelMsoTextDirectionRightToLeft = '\000\353\000\002'
};
typedef enum ExcelMsoTextDirection ExcelMsoTextDirection;

enum ExcelMsoBulletType {
	ExcelMsoBulletTypeBulletTypeUnset = '\000\347\377\376',
	ExcelMsoBulletTypeBulletTypeNone = '\000\350\000\000',
	ExcelMsoBulletTypeBulletTypeUnnumbered = '\000\350\000\001',
	ExcelMsoBulletTypeBulletTypeNumbered = '\000\350\000\002',
	ExcelMsoBulletTypePictureBulletType = '\000\350\000\003'
};
typedef enum ExcelMsoBulletType ExcelMsoBulletType;

enum ExcelMsoNumberedBulletStyle {
	ExcelMsoNumberedBulletStyleNumberedBulletStyleUnset = '\000\350\377\376',
	ExcelMsoNumberedBulletStyleNumberedBulletStyleAlphaLowercasePeriod = '\000\351\000\000',
	ExcelMsoNumberedBulletStyleNumberedBulletStyleAlphaUppercasePeriod = '\000\351\000\001',
	ExcelMsoNumberedBulletStyleNumberedBulletStyleArabicRightParen = '\000\351\000\002',
	ExcelMsoNumberedBulletStyleNumberedBulletStyleArabicPeriod = '\000\351\000\003',
	ExcelMsoNumberedBulletStyleNumberedBulletStyleRomanLowercaseParenBoth = '\000\351\000\004',
	ExcelMsoNumberedBulletStyleNumberedBulletStyleRomanLowercaseParenRight = '\000\351\000\005',
	ExcelMsoNumberedBulletStyleNumberedBulletStyleRomanLowercasePeriod = '\000\351\000\006',
	ExcelMsoNumberedBulletStyleNumberedBulletStyleRomanUppercasePeriod = '\000\351\000\007',
	ExcelMsoNumberedBulletStyleNumberedBulletStyleAlphaLowercaseParenBoth = '\000\351\000\010',
	ExcelMsoNumberedBulletStyleNumberedBulletStyleAlphaLowercaseParenRight = '\000\351\000\011',
	ExcelMsoNumberedBulletStyleNumberedBulletStyleAlphaUppercaseParenBoth = '\000\351\000\012',
	ExcelMsoNumberedBulletStyleNumberedBulletStyleAlphaUppercaseParenRight = '\000\351\000\013',
	ExcelMsoNumberedBulletStyleNumberedBulletStyleArabicParenBoth = '\000\351\000\014',
	ExcelMsoNumberedBulletStyleNumberedBulletStyleArabicPlain = '\000\351\000\015',
	ExcelMsoNumberedBulletStyleNumberedBulletStyleRomanUppercaseParenBoth = '\000\351\000\016',
	ExcelMsoNumberedBulletStyleNumberedBulletStyleRomanUppercaseParenRight = '\000\351\000\017',
	ExcelMsoNumberedBulletStyleNumberedBulletStyleSimplifiedChinesePlain = '\000\351\000\020',
	ExcelMsoNumberedBulletStyleNumberedBulletStyleSimplifiedChinesePeriod = '\000\351\000\021',
	ExcelMsoNumberedBulletStyleNumberedBulletStyleCircleNumberPlain = '\000\351\000\022',
	ExcelMsoNumberedBulletStyleNumberedBulletStyleCircleNumberWhitePlain = '\000\351\000\023',
	ExcelMsoNumberedBulletStyleNumberedBulletStyleCircleNumberBlackPlain = '\000\351\000\024',
	ExcelMsoNumberedBulletStyleNumberedBulletStyleTraditionalChinesePlain = '\000\351\000\025',
	ExcelMsoNumberedBulletStyleNumberedBulletStyleTraditionalChinesePeriod = '\000\351\000\026',
	ExcelMsoNumberedBulletStyleNumberedBulletStyleArabicAlphaDash = '\000\351\000\027',
	ExcelMsoNumberedBulletStyleNumberedBulletStyleArabicAbjadDash = '\000\351\000\030',
	ExcelMsoNumberedBulletStyleNumberedBulletStyleHebrewAlphaDash = '\000\351\000\031',
	ExcelMsoNumberedBulletStyleNumberedBulletStyleKanjiKoreanPlain = '\000\351\000\032',
	ExcelMsoNumberedBulletStyleNumberedBulletStyleKanjiKoreanPeriod = '\000\351\000\033',
	ExcelMsoNumberedBulletStyleNumberedBulletStyleArabicDBPlain = '\000\351\000\034',
	ExcelMsoNumberedBulletStyleNumberedBulletStyleArabicDBPeriod = '\000\351\000\035',
	ExcelMsoNumberedBulletStyleNumberedBulletStyleThaiAlphaPeriod = '\000\351\000\036',
	ExcelMsoNumberedBulletStyleNumberedBulletStyleThaiAlphaParenRight = '\000\351\000\037',
	ExcelMsoNumberedBulletStyleNumberedBulletStyleThaiAlphaParenBoth = '\000\351\000 ',
	ExcelMsoNumberedBulletStyleNumberedBulletStyleThaiNumberPeriod = '\000\351\000!',
	ExcelMsoNumberedBulletStyleNumberedBulletStyleThaiNumberParenRight = '\000\351\000\"',
	ExcelMsoNumberedBulletStyleNumberedBulletStyleThaiParenBoth = '\000\351\000#',
	ExcelMsoNumberedBulletStyleNumberedBulletStyleHindiAlphaPeriod = '\000\351\000$',
	ExcelMsoNumberedBulletStyleNumberedBulletStyleHindiNumberPeriod = '\000\351\000%',
	ExcelMsoNumberedBulletStyleNumberedBulletStyleKanjiSimpifiedChineseDBPeriod = '\000\351\000&',
	ExcelMsoNumberedBulletStyleNumberedBulletStyleHindiNumberParenRight = '\000\351\000\'',
	ExcelMsoNumberedBulletStyleNumberedBulletStyleHindiAlpha1Period = '\000\351\000('
};
typedef enum ExcelMsoNumberedBulletStyle ExcelMsoNumberedBulletStyle;

enum ExcelMsoTabStopType {
	ExcelMsoTabStopTypeTabstopUnset = '\000\364\377\376',
	ExcelMsoTabStopTypeTabstopLeft = '\000\365\000\001',
	ExcelMsoTabStopTypeTabstopCenter = '\000\365\000\002',
	ExcelMsoTabStopTypeTabstopRight = '\000\365\000\003',
	ExcelMsoTabStopTypeTabstopDecimal = '\000\365\000\004'
};
typedef enum ExcelMsoTabStopType ExcelMsoTabStopType;

enum ExcelMsoReflectionType {
	ExcelMsoReflectionTypeReflectionUnset = '\003\351\377\376',
	ExcelMsoReflectionTypeReflectionTypeNone = '\003\352\000\000',
	ExcelMsoReflectionTypeReflectionType1 = '\003\352\000\001',
	ExcelMsoReflectionTypeReflectionType2 = '\003\352\000\002',
	ExcelMsoReflectionTypeReflectionType3 = '\003\352\000\003',
	ExcelMsoReflectionTypeReflectionType4 = '\003\352\000\004',
	ExcelMsoReflectionTypeReflectionType5 = '\003\352\000\005',
	ExcelMsoReflectionTypeReflectionType6 = '\003\352\000\006',
	ExcelMsoReflectionTypeReflectionType7 = '\003\352\000\007',
	ExcelMsoReflectionTypeReflectionType8 = '\003\352\000\010',
	ExcelMsoReflectionTypeReflectionType9 = '\003\352\000\011'
};
typedef enum ExcelMsoReflectionType ExcelMsoReflectionType;

enum ExcelMsoTextureAlignment {
	ExcelMsoTextureAlignmentTextureUnset = '\003\352\377\376',
	ExcelMsoTextureAlignmentTextureTopLeft = '\003\353\000\000',
	ExcelMsoTextureAlignmentTextureTop = '\003\353\000\001',
	ExcelMsoTextureAlignmentTextureTopRight = '\003\353\000\002',
	ExcelMsoTextureAlignmentTextureLeft = '\003\353\000\003',
	ExcelMsoTextureAlignmentTextureCenter = '\003\353\000\004',
	ExcelMsoTextureAlignmentTextureRight = '\003\353\000\005',
	ExcelMsoTextureAlignmentTextureBottomLeft = '\003\353\000\006',
	ExcelMsoTextureAlignmentTextureBotton = '\003\353\000\007',
	ExcelMsoTextureAlignmentTextureBottomRight = '\003\353\000\010'
};
typedef enum ExcelMsoTextureAlignment ExcelMsoTextureAlignment;

enum ExcelMsoBaselineAlignment {
	ExcelMsoBaselineAlignmentTextBaselineAlignmentUnset = '\003\353\377\376',
	ExcelMsoBaselineAlignmentTextBaselineAlignBaseline = '\003\354\000\001',
	ExcelMsoBaselineAlignmentTextBaselineAlignTop = '\003\354\000\002',
	ExcelMsoBaselineAlignmentTextBaselineAlignCenter = '\003\354\000\003',
	ExcelMsoBaselineAlignmentTextBaselineAlignEastAsian50 = '\003\354\000\004',
	ExcelMsoBaselineAlignmentTextBaselineAlignAutomatic = '\003\354\000\005'
};
typedef enum ExcelMsoBaselineAlignment ExcelMsoBaselineAlignment;

enum ExcelMsoClipboardFormat {
	ExcelMsoClipboardFormatClipboardFormatUnset = '\003\354\377\376',
	ExcelMsoClipboardFormatNativeClipboardFormat = '\003\355\000\001',
	ExcelMsoClipboardFormatHTMlClipboardFormat = '\003\355\000\002',
	ExcelMsoClipboardFormatRTFClipboardFormat = '\003\355\000\003',
	ExcelMsoClipboardFormatPlainTextClipboardFormat = '\003\355\000\004'
};
typedef enum ExcelMsoClipboardFormat ExcelMsoClipboardFormat;

enum ExcelMsoTextRangeInsertPosition {
	ExcelMsoTextRangeInsertPositionInsertBefore = '\003\356\000\000',
	ExcelMsoTextRangeInsertPositionInsertAfter = '\003\356\000\001'
};
typedef enum ExcelMsoTextRangeInsertPosition ExcelMsoTextRangeInsertPosition;

enum ExcelMsoPictureType {
	ExcelMsoPictureTypeSaveAsDefault = '\003\362\377\376',
	ExcelMsoPictureTypeSaveAsPNGFile = '\003\363\000\000',
	ExcelMsoPictureTypeSaveAsBMPFile = '\003\363\000\001',
	ExcelMsoPictureTypeSaveAsGIFFile = '\003\363\000\002',
	ExcelMsoPictureTypeSaveAsJPGFile = '\003\363\000\003',
	ExcelMsoPictureTypeSaveAsPDFFile = '\003\363\000\004'
};
typedef enum ExcelMsoPictureType ExcelMsoPictureType;

enum ExcelMsoPictureEffectType {
	ExcelMsoPictureEffectTypeNoEffect = '\003\364\000\000',
	ExcelMsoPictureEffectTypeEffectBackgroundRemoval = '\003\364\000\001',
	ExcelMsoPictureEffectTypeEffectBlur = '\003\364\000\002',
	ExcelMsoPictureEffectTypeEffectBrightnessContrast = '\003\364\000\003',
	ExcelMsoPictureEffectTypeEffectCement = '\003\364\000\004',
	ExcelMsoPictureEffectTypeEffectCrisscrossEtching = '\003\364\000\005',
	ExcelMsoPictureEffectTypeEffectChalkSketch = '\003\364\000\006',
	ExcelMsoPictureEffectTypeEffectColorTemperature = '\003\364\000\007',
	ExcelMsoPictureEffectTypeEffectCutout = '\003\364\000\010',
	ExcelMsoPictureEffectTypeEffectFilmGrain = '\003\364\000\011',
	ExcelMsoPictureEffectTypeEffectGlass = '\003\364\000\012',
	ExcelMsoPictureEffectTypeEffectGlowDiffused = '\003\364\000\013',
	ExcelMsoPictureEffectTypeEffectGlowEdges = '\003\364\000\014',
	ExcelMsoPictureEffectTypeEffectLightScreen = '\003\364\000\015',
	ExcelMsoPictureEffectTypeEffectLineDrawing = '\003\364\000\016',
	ExcelMsoPictureEffectTypeEffectMarker = '\003\364\000\017',
	ExcelMsoPictureEffectTypeEffectMosiaicBubbles = '\003\364\000\020',
	ExcelMsoPictureEffectTypeEffectPaintBrush = '\003\364\000\021',
	ExcelMsoPictureEffectTypeEffectPaintStrokes = '\003\364\000\022',
	ExcelMsoPictureEffectTypeEffectPastelsSmooth = '\003\364\000\023',
	ExcelMsoPictureEffectTypeEffectPencilGrayscale = '\003\364\000\024',
	ExcelMsoPictureEffectTypeEffectPencilSketch = '\003\364\000\025',
	ExcelMsoPictureEffectTypeEffectPhotocopy = '\003\364\000\026',
	ExcelMsoPictureEffectTypeEffectPlasticWrap = '\003\364\000\027',
	ExcelMsoPictureEffectTypeEffectSaturation = '\003\364\000\030',
	ExcelMsoPictureEffectTypeEffectSharpenSoften = '\003\364\000\031',
	ExcelMsoPictureEffectTypeEffectTexturizer = '\003\364\000\032',
	ExcelMsoPictureEffectTypeEffectWatercolorSponge = '\003\364\000\033'
};
typedef enum ExcelMsoPictureEffectType ExcelMsoPictureEffectType;

enum ExcelMsoSegmentType {
	ExcelMsoSegmentTypeLine = '\000\217\000\000',
	ExcelMsoSegmentTypeCurve = '\000\217\000\001'
};
typedef enum ExcelMsoSegmentType ExcelMsoSegmentType;

enum ExcelMsoEditingType {
	ExcelMsoEditingTypeAuto = '\000\220\000\000',
	ExcelMsoEditingTypeCorner = '\000\220\000\001',
	ExcelMsoEditingTypeSmooth = '\000\220\000\002',
	ExcelMsoEditingTypeSymmetric = '\000\220\000\003'
};
typedef enum ExcelMsoEditingType ExcelMsoEditingType;

enum ExcelMsoSmartArtNodePosition {
	ExcelMsoSmartArtNodePositionDefaultNodePosition = '\003\365\000\001',
	ExcelMsoSmartArtNodePositionAfterNode = '\003\365\000\002',
	ExcelMsoSmartArtNodePositionBeforeNode = '\003\365\000\003',
	ExcelMsoSmartArtNodePositionAboveNode = '\003\365\000\004',
	ExcelMsoSmartArtNodePositionBelowNode = '\003\365\000\005'
};
typedef enum ExcelMsoSmartArtNodePosition ExcelMsoSmartArtNodePosition;

enum ExcelMsoSmartArtNodeType {
	ExcelMsoSmartArtNodeTypeDefaultNode = '\003\366\000\001',
	ExcelMsoSmartArtNodeTypeAssistantNode = '\003\366\000\002'
};
typedef enum ExcelMsoSmartArtNodeType ExcelMsoSmartArtNodeType;

enum ExcelMsoOrgChartLayoutType {
	ExcelMsoOrgChartLayoutTypeOrgChartLayoutUnset = '\003\366\377\376',
	ExcelMsoOrgChartLayoutTypeOrgChartLayoutStandard = '\003\367\000\001',
	ExcelMsoOrgChartLayoutTypeOrgChartLayoutBothHanging = '\003\367\000\002',
	ExcelMsoOrgChartLayoutTypeOrgChartLayoutLeftHanging = '\003\367\000\003',
	ExcelMsoOrgChartLayoutTypeOrgChartLayoutRightHanging = '\003\367\000\004',
	ExcelMsoOrgChartLayoutTypeOrgChartLayoutDefault = '\003\367\000\005'
};
typedef enum ExcelMsoOrgChartLayoutType ExcelMsoOrgChartLayoutType;

enum ExcelMsoAlignCmd {
	ExcelMsoAlignCmdAlignLefts = '\000\000\000\000',
	ExcelMsoAlignCmdAlignCenters = '\000\000\000\001',
	ExcelMsoAlignCmdAlignRights = '\000\000\000\002',
	ExcelMsoAlignCmdAlignTops = '\000\000\000\003',
	ExcelMsoAlignCmdAlignMiddles = '\000\000\000\004',
	ExcelMsoAlignCmdAlignBottoms = '\000\000\000\005'
};
typedef enum ExcelMsoAlignCmd ExcelMsoAlignCmd;

enum ExcelMsoDistributeCmd {
	ExcelMsoDistributeCmdDistributeHorizontally = '\000\000\000\000',
	ExcelMsoDistributeCmdDistributeVertically = '\000\000\000\001'
};
typedef enum ExcelMsoDistributeCmd ExcelMsoDistributeCmd;

enum ExcelMsoOrientation {
	ExcelMsoOrientationOrientationUnset = '\000\214\377\376',
	ExcelMsoOrientationHorizontalOrientation = '\000\215\000\001',
	ExcelMsoOrientationVerticalOrientation = '\000\215\000\002'
};
typedef enum ExcelMsoOrientation ExcelMsoOrientation;

enum ExcelMsoZOrderCmd {
	ExcelMsoZOrderCmdBringShapeToFront = '\000p\000\000',
	ExcelMsoZOrderCmdSendShapeToBack = '\000p\000\001',
	ExcelMsoZOrderCmdBringShapeForward = '\000p\000\002',
	ExcelMsoZOrderCmdSendShapeBackward = '\000p\000\003',
	ExcelMsoZOrderCmdBringShapeInFrontOfText = '\000p\000\004',
	ExcelMsoZOrderCmdSendShapeBehindText = '\000p\000\005'
};
typedef enum ExcelMsoZOrderCmd ExcelMsoZOrderCmd;

enum ExcelMsoScriptLanguage {
	ExcelMsoScriptLanguageBringShapeToFront = '\000o\000\001',
	ExcelMsoScriptLanguageSendShapeToBack = '\000o\000\002',
	ExcelMsoScriptLanguageBringShapeForward = '\000o\000\003',
	ExcelMsoScriptLanguageSendShapeBackward = '\000o\000\004'
};
typedef enum ExcelMsoScriptLanguage ExcelMsoScriptLanguage;

enum ExcelMsoFlipCmd {
	ExcelMsoFlipCmdFlipHorizontal = '\000q\000\000',
	ExcelMsoFlipCmdFlipVertical = '\000q\000\001'
};
typedef enum ExcelMsoFlipCmd ExcelMsoFlipCmd;

enum ExcelMsoTriState {
	ExcelMsoTriStateTrue = '\000\240\377\377',
	ExcelMsoTriStateFalse = '\000\241\000\000',
	ExcelMsoTriStateCTrue = '\000\241\000\001',
	ExcelMsoTriStateToggle = '\000\240\377\375',
	ExcelMsoTriStateTriStateUnset = '\000\240\377\376'
};
typedef enum ExcelMsoTriState ExcelMsoTriState;

enum ExcelMsoBlackWhiteMode {
	ExcelMsoBlackWhiteModeBlackAndWhiteUnset = '\000\254\377\376',
	ExcelMsoBlackWhiteModeBlackAndWhiteModeAutomatic = '\000\255\000\001',
	ExcelMsoBlackWhiteModeBlackAndWhiteModeGrayScale = '\000\255\000\002',
	ExcelMsoBlackWhiteModeBlackAndWhiteModeLightGrayScale = '\000\255\000\003',
	ExcelMsoBlackWhiteModeBlackAndWhiteModeInverseGrayScale = '\000\255\000\004',
	ExcelMsoBlackWhiteModeBlackAndWhiteModeGrayOutline = '\000\255\000\005',
	ExcelMsoBlackWhiteModeBlackAndWhiteModeBlackTextAndLine = '\000\255\000\006',
	ExcelMsoBlackWhiteModeBlackAndWhiteModeHighContrast = '\000\255\000\007',
	ExcelMsoBlackWhiteModeBlackAndWhiteModeBlack = '\000\255\000\010',
	ExcelMsoBlackWhiteModeBlackAndWhiteModeWhite = '\000\255\000\011',
	ExcelMsoBlackWhiteModeBlackAndWhiteModeDontShow = '\000\255\000\012'
};
typedef enum ExcelMsoBlackWhiteMode ExcelMsoBlackWhiteMode;

enum ExcelMsoBarPosition {
	ExcelMsoBarPositionBarLeft = '\000r\000\000',
	ExcelMsoBarPositionBarTop = '\000r\000\001',
	ExcelMsoBarPositionBarRight = '\000r\000\002',
	ExcelMsoBarPositionBarBottom = '\000r\000\003',
	ExcelMsoBarPositionBarFloating = '\000r\000\004',
	ExcelMsoBarPositionBarPopUp = '\000r\000\005',
	ExcelMsoBarPositionBarMenu = '\000r\000\006'
};
typedef enum ExcelMsoBarPosition ExcelMsoBarPosition;

enum ExcelMsoBarProtection {
	ExcelMsoBarProtectionNoProtection = '\000s\000\000',
	ExcelMsoBarProtectionNoCustomize = '\000s\000\001',
	ExcelMsoBarProtectionNoResize = '\000s\000\002',
	ExcelMsoBarProtectionNoMove = '\000s\000\004',
	ExcelMsoBarProtectionNoChangeVisible = '\000s\000\010',
	ExcelMsoBarProtectionNoChangeDock = '\000s\000\020',
	ExcelMsoBarProtectionNoVerticalDock = '\000s\000 ',
	ExcelMsoBarProtectionNoHorizontalDock = '\000s\000@'
};
typedef enum ExcelMsoBarProtection ExcelMsoBarProtection;

enum ExcelMsoBarType {
	ExcelMsoBarTypeNormalCommandBar = '\000t\000\000',
	ExcelMsoBarTypeMenubarCommandBar = '\000t\000\001',
	ExcelMsoBarTypePopupCommandBar = '\000t\000\002'
};
typedef enum ExcelMsoBarType ExcelMsoBarType;

enum ExcelMsoControlType {
	ExcelMsoControlTypeControlCustom = '\000u\000\000',
	ExcelMsoControlTypeControlButton = '\000u\000\001',
	ExcelMsoControlTypeControlEdit = '\000u\000\002',
	ExcelMsoControlTypeControlDropDown = '\000u\000\003',
	ExcelMsoControlTypeControlCombobox = '\000u\000\004',
	ExcelMsoControlTypeButtonDropDown = '\000u\000\005',
	ExcelMsoControlTypeSplitDropDown = '\000u\000\006',
	ExcelMsoControlTypeOCXDropDown = '\000u\000\007',
	ExcelMsoControlTypeGenericDropDown = '\000u\000\010',
	ExcelMsoControlTypeGraphicDropDown = '\000u\000\011',
	ExcelMsoControlTypeControlPopup = '\000u\000\012',
	ExcelMsoControlTypeGraphicPopup = '\000u\000\013',
	ExcelMsoControlTypeButtonPopup = '\000u\000\014',
	ExcelMsoControlTypeSplitButtonPopup = '\000u\000\015',
	ExcelMsoControlTypeSplitButtonMRUPopup = '\000u\000\016',
	ExcelMsoControlTypeControlLabel = '\000u\000\017',
	ExcelMsoControlTypeExpandingGrid = '\000u\000\020',
	ExcelMsoControlTypeSplitExpandingGrid = '\000u\000\021',
	ExcelMsoControlTypeControlGrid = '\000u\000\022',
	ExcelMsoControlTypeControlGauge = '\000u\000\023',
	ExcelMsoControlTypeGraphicCombobox = '\000u\000\024',
	ExcelMsoControlTypeControlPane = '\000u\000\025',
	ExcelMsoControlTypeActiveX = '\000u\000\026',
	ExcelMsoControlTypeControlGroup = '\000u\000\027',
	ExcelMsoControlTypeControlTab = '\000u\000\030',
	ExcelMsoControlTypeControlSpinner = '\000u\000\031'
};
typedef enum ExcelMsoControlType ExcelMsoControlType;

enum ExcelMsoButtonState {
	ExcelMsoButtonStateButtonStateUp = '\000v\000\000',
	ExcelMsoButtonStateButtonStateDown = '\000u\377\377',
	ExcelMsoButtonStateButtonStateUnset = '\000v\000\002'
};
typedef enum ExcelMsoButtonState ExcelMsoButtonState;

enum ExcelMsoControlOLEUsage {
	ExcelMsoControlOLEUsageNeither = '\000w\000\000',
	ExcelMsoControlOLEUsageServer = '\000w\000\001',
	ExcelMsoControlOLEUsageClient = '\000w\000\002',
	ExcelMsoControlOLEUsageBoth = '\000w\000\003'
};
typedef enum ExcelMsoControlOLEUsage ExcelMsoControlOLEUsage;

enum ExcelMsoButtonStyle {
	ExcelMsoButtonStyleButtonAutomatic = '\000x\000\000',
	ExcelMsoButtonStyleButtonIcon = '\000x\000\001',
	ExcelMsoButtonStyleButtonCaption = '\000x\000\002',
	ExcelMsoButtonStyleButtonIconAndCaption = '\000x\000\003'
};
typedef enum ExcelMsoButtonStyle ExcelMsoButtonStyle;

enum ExcelMsoComboStyle {
	ExcelMsoComboStyleComboboxStyleNormal = '\000y\000\000',
	ExcelMsoComboStyleComboboxStyleLabel = '\000y\000\001'
};
typedef enum ExcelMsoComboStyle ExcelMsoComboStyle;

enum ExcelMsoMenuAnimation {
	ExcelMsoMenuAnimationNone = '\000{\000\000',
	ExcelMsoMenuAnimationRandom = '\000{\000\001',
	ExcelMsoMenuAnimationUnfold = '\000{\000\002',
	ExcelMsoMenuAnimationSlide = '\000{\000\003'
};
typedef enum ExcelMsoMenuAnimation ExcelMsoMenuAnimation;

enum ExcelMsoHyperlinkType {
	ExcelMsoHyperlinkTypeHyperlinkTypeTextRange = '\000\226\000\000',
	ExcelMsoHyperlinkTypeHyperlinkTypeShape = '\000\226\000\001',
	ExcelMsoHyperlinkTypeHyperlinkTypeInlineShape = '\000\226\000\002'
};
typedef enum ExcelMsoHyperlinkType ExcelMsoHyperlinkType;

enum ExcelMsoExtraInfoMethod {
	ExcelMsoExtraInfoMethodAppendString = '\000\256\000\000',
	ExcelMsoExtraInfoMethodPostString = '\000\256\000\001'
};
typedef enum ExcelMsoExtraInfoMethod ExcelMsoExtraInfoMethod;

enum ExcelMsoAnimationType {
	ExcelMsoAnimationTypeIdle = '\000|\000\001',
	ExcelMsoAnimationTypeGreeting = '\000|\000\002',
	ExcelMsoAnimationTypeGoodbye = '\000|\000\003',
	ExcelMsoAnimationTypeBeginSpeaking = '\000|\000\004',
	ExcelMsoAnimationTypeCharacterSuccessMajor = '\000|\000\006',
	ExcelMsoAnimationTypeGetAttentionMajor = '\000|\000\013',
	ExcelMsoAnimationTypeGetAttentionMinor = '\000|\000\014',
	ExcelMsoAnimationTypeSearching = '\000|\000\015',
	ExcelMsoAnimationTypePrinting = '\000|\000\022',
	ExcelMsoAnimationTypeGestureRight = '\000|\000\023',
	ExcelMsoAnimationTypeWritingNotingSomething = '\000|\000\026',
	ExcelMsoAnimationTypeWorkingAtSomething = '\000|\000\027',
	ExcelMsoAnimationTypeThinking = '\000|\000\030',
	ExcelMsoAnimationTypeSendingMail = '\000|\000\031',
	ExcelMsoAnimationTypeListensToComputer = '\000|\000\032',
	ExcelMsoAnimationTypeDisappear = '\000|\000\037',
	ExcelMsoAnimationTypeAppear = '\000|\000 ',
	ExcelMsoAnimationTypeGetArtsy = '\000|\000d',
	ExcelMsoAnimationTypeGetTechy = '\000|\000e',
	ExcelMsoAnimationTypeGetWizardy = '\000|\000f',
	ExcelMsoAnimationTypeCheckingSomething = '\000|\000g',
	ExcelMsoAnimationTypeLookDown = '\000|\000h',
	ExcelMsoAnimationTypeLookDownLeft = '\000|\000i',
	ExcelMsoAnimationTypeLookDownRight = '\000|\000j',
	ExcelMsoAnimationTypeLookLeft = '\000|\000k',
	ExcelMsoAnimationTypeLookRight = '\000|\000l',
	ExcelMsoAnimationTypeLookUp = '\000|\000m',
	ExcelMsoAnimationTypeLookUpLeft = '\000|\000n',
	ExcelMsoAnimationTypeLookUpRight = '\000|\000o',
	ExcelMsoAnimationTypeSaving = '\000|\000p',
	ExcelMsoAnimationTypeGestureDown = '\000|\000q',
	ExcelMsoAnimationTypeGestureLeft = '\000|\000r',
	ExcelMsoAnimationTypeGestureUp = '\000|\000s',
	ExcelMsoAnimationTypeEmptyTrash = '\000|\000t'
};
typedef enum ExcelMsoAnimationType ExcelMsoAnimationType;

enum ExcelMsoButtonSetType {
	ExcelMsoButtonSetTypeButtonNone = '\000}\000\000',
	ExcelMsoButtonSetTypeButtonOk = '\000}\000\001',
	ExcelMsoButtonSetTypeButtonCancel = '\000}\000\002',
	ExcelMsoButtonSetTypeButtonsOkCancel = '\000}\000\003',
	ExcelMsoButtonSetTypeButtonsYesNo = '\000}\000\004',
	ExcelMsoButtonSetTypeButtonsYesNoCancel = '\000}\000\005',
	ExcelMsoButtonSetTypeButtonsBackClose = '\000}\000\006',
	ExcelMsoButtonSetTypeButtonsNextClose = '\000}\000\007',
	ExcelMsoButtonSetTypeButtonsBackNextClose = '\000}\000\010',
	ExcelMsoButtonSetTypeButtonsRetryCancel = '\000}\000\011',
	ExcelMsoButtonSetTypeButtonsAbortRetryIgnore = '\000}\000\012',
	ExcelMsoButtonSetTypeButtonsSearchClose = '\000}\000\013',
	ExcelMsoButtonSetTypeButtonsBackNextSnooze = '\000}\000\014',
	ExcelMsoButtonSetTypeButtonsTipsOptionsClose = '\000}\000\015',
	ExcelMsoButtonSetTypeButtonsYesAllNoCancel = '\000}\000\016'
};
typedef enum ExcelMsoButtonSetType ExcelMsoButtonSetType;

enum ExcelMsoIconType {
	ExcelMsoIconTypeIconNone = '\000~\000\000',
	ExcelMsoIconTypeIconApplication = '\000~\000\001',
	ExcelMsoIconTypeIconAlert = '\000~\000\002',
	ExcelMsoIconTypeIconTip = '\000~\000\003',
	ExcelMsoIconTypeIconAlertCritical = '\000~\000e',
	ExcelMsoIconTypeIconAlertWarning = '\000~\000g',
	ExcelMsoIconTypeIconAlertInfo = '\000~\000h'
};
typedef enum ExcelMsoIconType ExcelMsoIconType;

enum ExcelMsoWizardActType {
	ExcelMsoWizardActTypeInactive = '\000\202\000\000',
	ExcelMsoWizardActTypeActive = '\000\202\000\001',
	ExcelMsoWizardActTypeSuspend = '\000\202\000\002',
	ExcelMsoWizardActTypeResume = '\000\202\000\003'
};
typedef enum ExcelMsoWizardActType ExcelMsoWizardActType;

enum ExcelMsoDocProperties {
	ExcelMsoDocPropertiesPropertyTypeNumber = '\000\242\000\001',
	ExcelMsoDocPropertiesPropertyTypeBoolean = '\000\242\000\002',
	ExcelMsoDocPropertiesPropertyTypeDate = '\000\242\000\003',
	ExcelMsoDocPropertiesPropertyTypeString = '\000\242\000\004',
	ExcelMsoDocPropertiesPropertyTypeFloat = '\000\242\000\005'
};
typedef enum ExcelMsoDocProperties ExcelMsoDocProperties;

enum ExcelMsoAutomationSecurity {
	ExcelMsoAutomationSecurityMsoAutomationSecurityLow = '\000\243\000\001',
	ExcelMsoAutomationSecurityMsoAutomationSecurityByUI = '\000\243\000\002',
	ExcelMsoAutomationSecurityMsoAutomationSecurityForceDisable = '\000\243\000\003'
};
typedef enum ExcelMsoAutomationSecurity ExcelMsoAutomationSecurity;

enum ExcelMsoScreenSize {
	ExcelMsoScreenSizeResolution544x376 = '\000\204\000\000',
	ExcelMsoScreenSizeResolution640x480 = '\000\204\000\001',
	ExcelMsoScreenSizeResolution720x512 = '\000\204\000\002',
	ExcelMsoScreenSizeResolution800x600 = '\000\204\000\003',
	ExcelMsoScreenSizeResolution1024x768 = '\000\204\000\004',
	ExcelMsoScreenSizeResolution1152x882 = '\000\204\000\005',
	ExcelMsoScreenSizeResolution1152x900 = '\000\204\000\006',
	ExcelMsoScreenSizeResolution1280x1024 = '\000\204\000\007',
	ExcelMsoScreenSizeResolution1600x1200 = '\000\204\000\010',
	ExcelMsoScreenSizeResolution1800x1440 = '\000\204\000\011',
	ExcelMsoScreenSizeResolution1920x1200 = '\000\204\000\012'
};
typedef enum ExcelMsoScreenSize ExcelMsoScreenSize;

enum ExcelMsoCharacterSet {
	ExcelMsoCharacterSetArabicCharacterSet = '\000\205\000\001',
	ExcelMsoCharacterSetCyrillicCharacterSet = '\000\205\000\002',
	ExcelMsoCharacterSetEnglishCharacterSet = '\000\205\000\003',
	ExcelMsoCharacterSetGreekCharacterSet = '\000\205\000\004',
	ExcelMsoCharacterSetHebrewCharacterSet = '\000\205\000\005',
	ExcelMsoCharacterSetJapaneseCharacterSet = '\000\205\000\006',
	ExcelMsoCharacterSetKoreanCharacterSet = '\000\205\000\007',
	ExcelMsoCharacterSetMultilingualUnicodeCharacterSet = '\000\205\000\010',
	ExcelMsoCharacterSetSimplifiedChineseCharacterSet = '\000\205\000\011',
	ExcelMsoCharacterSetThaiCharacterSet = '\000\205\000\012',
	ExcelMsoCharacterSetTraditionalChineseCharacterSet = '\000\205\000\013',
	ExcelMsoCharacterSetVietnameseCharacterSet = '\000\205\000\014'
};
typedef enum ExcelMsoCharacterSet ExcelMsoCharacterSet;

enum ExcelMsoEncoding {
	ExcelMsoEncodingEncodingThai = '\000\213\003j',
	ExcelMsoEncodingEncodingJapaneseShiftJIS = '\000\213\003\244',
	ExcelMsoEncodingEncodingSimplifiedChinese = '\000\213\003\250',
	ExcelMsoEncodingEncodingKorean = '\000\213\003\265',
	ExcelMsoEncodingEncodingBig5TraditionalChinese = '\000\213\003\266',
	ExcelMsoEncodingEncodingLittleEndian = '\000\213\004\260',
	ExcelMsoEncodingEncodingBigEndian = '\000\213\004\261',
	ExcelMsoEncodingEncodingCentralEuropean = '\000\213\004\342',
	ExcelMsoEncodingEncodingCyrillic = '\000\213\004\343',
	ExcelMsoEncodingEncodingWestern = '\000\213\004\344',
	ExcelMsoEncodingEncodingGreek = '\000\213\004\345',
	ExcelMsoEncodingEncodingTurkish = '\000\213\004\346',
	ExcelMsoEncodingEncodingHebrew = '\000\213\004\347',
	ExcelMsoEncodingEncodingArabic = '\000\213\004\350',
	ExcelMsoEncodingEncodingBaltic = '\000\213\004\351',
	ExcelMsoEncodingEncodingVietnamese = '\000\213\004\352',
	ExcelMsoEncodingEncodingISO88591Latin1 = '\000\213o\257',
	ExcelMsoEncodingEncodingISO88592CentralEurope = '\000\213o\260',
	ExcelMsoEncodingEncodingISO88593Latin3 = '\000\213o\261',
	ExcelMsoEncodingEncodingISO88594Baltic = '\000\213o\262',
	ExcelMsoEncodingEncodingISO88595Cyrillic = '\000\213o\263',
	ExcelMsoEncodingEncodingISO88596Arabic = '\000\213o\264',
	ExcelMsoEncodingEncodingISO88597Greek = '\000\213o\265',
	ExcelMsoEncodingEncodingISO88598Hebrew = '\000\213o\266',
	ExcelMsoEncodingEncodingISO88599Turkish = '\000\213o\267',
	ExcelMsoEncodingEncodingISO885915Latin9 = '\000\213o\275',
	ExcelMsoEncodingEncodingISO2022JapaneseNoHalfWidthKatakana = '\000\213\304,',
	ExcelMsoEncodingEncodingISO2022JapaneseJISX02021984 = '\000\213\304-',
	ExcelMsoEncodingEncodingISO2022JapaneseJISX02011989 = '\000\213\304.',
	ExcelMsoEncodingEncodingISO2022KR = '\000\213\3041',
	ExcelMsoEncodingEncodingISO2022CNTraditionalChinese = '\000\213\3043',
	ExcelMsoEncodingEncodingISO2022CNSimplifiedChinese = '\000\213\3045',
	ExcelMsoEncodingEncodingMacRoman = '\000\213\'\020',
	ExcelMsoEncodingEncodingMacJapanese = '\000\213\'\021',
	ExcelMsoEncodingEncodingMacTraditionalChinese = '\000\213\'\022',
	ExcelMsoEncodingEncodingMacKorean = '\000\213\'\023',
	ExcelMsoEncodingEncodingMacArabic = '\000\213\'\024',
	ExcelMsoEncodingEncodingMacHebrew = '\000\213\'\025',
	ExcelMsoEncodingEncodingMacGreek1 = '\000\213\'\026',
	ExcelMsoEncodingEncodingMacCyrillic = '\000\213\'\027',
	ExcelMsoEncodingEncodingMacSimplifiedChineseGB2312 = '\000\213\'\030',
	ExcelMsoEncodingEncodingMacRomania = '\000\213\'\032',
	ExcelMsoEncodingEncodingMacUkraine = '\000\213\'!',
	ExcelMsoEncodingEncodingMacLatin2 = '\000\213\'-',
	ExcelMsoEncodingEncodingMacIcelandic = '\000\213\'_',
	ExcelMsoEncodingEncodingMacTurkish = '\000\213\'a',
	ExcelMsoEncodingEncodingMacCroatia = '\000\213\'b',
	ExcelMsoEncodingEncodingEBCDICUSCanada = '\000\213\000%',
	ExcelMsoEncodingEncodingEBCDICInternational = '\000\213\001\364',
	ExcelMsoEncodingEncodingEBCDICMultilingualROECELatin2 = '\000\213\003f',
	ExcelMsoEncodingEncodingEBCDICGreekModern = '\000\213\003k',
	ExcelMsoEncodingEncodingEBCDICTurkishLatin5 = '\000\213\004\002',
	ExcelMsoEncodingEncodingEBCDICGermany = '\000\213O1',
	ExcelMsoEncodingEncodingEBCDICDenmarkNorway = '\000\213O5',
	ExcelMsoEncodingEncodingEBCDICFinlandSweden = '\000\213O6',
	ExcelMsoEncodingEncodingEBCDICItaly = '\000\213O8',
	ExcelMsoEncodingEncodingEBCDICLatinAmericaSpain = '\000\213O<',
	ExcelMsoEncodingEncodingEBCDICUnitedKingdom = '\000\213O=',
	ExcelMsoEncodingEncodingEBCDICJapaneseKatakanaExtended = '\000\213OB',
	ExcelMsoEncodingEncodingEBCDICFrance = '\000\213OI',
	ExcelMsoEncodingEncodingEBCDICArabic = '\000\213O\304',
	ExcelMsoEncodingEncodingEBCDICGreek = '\000\213O\307',
	ExcelMsoEncodingEncodingEBCDICHebrew = '\000\213O\310',
	ExcelMsoEncodingEncodingEBCDICKoreanExtended = '\000\213Qa',
	ExcelMsoEncodingEncodingEBCDICThai = '\000\213Qf',
	ExcelMsoEncodingEncodingEBCDICIcelandic = '\000\213Q\207',
	ExcelMsoEncodingEncodingEBCDICTurkish = '\000\213Q\251',
	ExcelMsoEncodingEncodingEBCDICRussian = '\000\213Q\220',
	ExcelMsoEncodingEncodingEBCDICSerbianBulgarian = '\000\213R!',
	ExcelMsoEncodingEncodingEBCDICJapaneseKatakanaExtendedAndJapanese = '\000\213\306\362',
	ExcelMsoEncodingEncodingEBCDICUSCanadaAndJapanese = '\000\213\306\363',
	ExcelMsoEncodingEncodingEBCDICExtendedAndKorean = '\000\213\306\365',
	ExcelMsoEncodingEncodingEBCDICSimplifiedChineseExtendedAndSimplifiedChinese = '\000\213\306\367',
	ExcelMsoEncodingEncodingEBCDICUSCanadaAndTraditionalChinese = '\000\213\306\371',
	ExcelMsoEncodingEncodingEBCDICJapaneseLatinExtendedAndJapanese = '\000\213\306\373',
	ExcelMsoEncodingEncodingOEMUnitedStates = '\000\213\001\265',
	ExcelMsoEncodingEncodingOEMGreek = '\000\213\002\341',
	ExcelMsoEncodingEncodingOEMBaltic = '\000\213\003\007',
	ExcelMsoEncodingEncodingOEMMultilingualLatinI = '\000\213\003R',
	ExcelMsoEncodingEncodingOEMMultilingualLatinII = '\000\213\003T',
	ExcelMsoEncodingEncodingOEMCyrillic = '\000\213\003W',
	ExcelMsoEncodingEncodingOEMTurkish = '\000\213\003Y',
	ExcelMsoEncodingEncodingOEMPortuguese = '\000\213\003\\',
	ExcelMsoEncodingEncodingOEMIcelandic = '\000\213\003]',
	ExcelMsoEncodingEncodingOEMHebrew = '\000\213\003^',
	ExcelMsoEncodingEncodingOEMCanadianFrench = '\000\213\003_',
	ExcelMsoEncodingEncodingOEMArabic = '\000\213\003`',
	ExcelMsoEncodingEncodingOEMNordic = '\000\213\003a',
	ExcelMsoEncodingEncodingOEMCyrillicII = '\000\213\003b',
	ExcelMsoEncodingEncodingOEMModernGreek = '\000\213\003e',
	ExcelMsoEncodingEncodingEUCJapanese = '\000\213\312\334',
	ExcelMsoEncodingEncodingEUCChineseSimplifiedChinese = '\000\213\312\340',
	ExcelMsoEncodingEncodingEUCKorean = '\000\213\312\355',
	ExcelMsoEncodingEncodingEUCTaiwaneseTraditionalChinese = '\000\213\312\356',
	ExcelMsoEncodingEncodingDevanagari = '\000\213\336\252',
	ExcelMsoEncodingEncodingBengali = '\000\213\336\253',
	ExcelMsoEncodingEncodingTamil = '\000\213\336\254',
	ExcelMsoEncodingEncodingTelugu = '\000\213\336\255',
	ExcelMsoEncodingEncodingAssamese = '\000\213\336\256',
	ExcelMsoEncodingEncodingOriya = '\000\213\336\257',
	ExcelMsoEncodingEncodingKannada = '\000\213\336\260',
	ExcelMsoEncodingEncodingMalayalam = '\000\213\336\261',
	ExcelMsoEncodingEncodingGujarati = '\000\213\336\262',
	ExcelMsoEncodingEncodingPunjabi = '\000\213\336\263',
	ExcelMsoEncodingEncodingArabicASMO = '\000\213\002\304',
	ExcelMsoEncodingEncodingArabicTransparentASMO = '\000\213\002\320',
	ExcelMsoEncodingEncodingKoreanJohab = '\000\213\005Q',
	ExcelMsoEncodingEncodingTaiwanCNS = '\000\213N ',
	ExcelMsoEncodingEncodingTaiwanTCA = '\000\213N!',
	ExcelMsoEncodingEncodingTaiwanEten = '\000\213N\"',
	ExcelMsoEncodingEncodingTaiwanIBM5550 = '\000\213N#',
	ExcelMsoEncodingEncodingTaiwanTeletext = '\000\213N$',
	ExcelMsoEncodingEncodingTaiwanWang = '\000\213N%',
	ExcelMsoEncodingEncodingIA5IRV = '\000\213N\211',
	ExcelMsoEncodingEncodingIA5German = '\000\213N\212',
	ExcelMsoEncodingEncodingIA5Swedish = '\000\213N\213',
	ExcelMsoEncodingEncodingIA5Norwegian = '\000\213N\214',
	ExcelMsoEncodingEncodingUSASCII = '\000\213N\237',
	ExcelMsoEncodingEncodingT61 = '\000\213O%',
	ExcelMsoEncodingEncodingISO6937NonspacingAccent = '\000\213O-',
	ExcelMsoEncodingEncodingKOI8R = '\000\213Q\202',
	ExcelMsoEncodingEncodingExtAlphaLowercase = '\000\213R#',
	ExcelMsoEncodingEncodingKOI8U = '\000\213Uj',
	ExcelMsoEncodingEncodingEuropa3 = '\000\213qI',
	ExcelMsoEncodingEncodingHZGBSimplifiedChinese = '\000\213\316\310',
	ExcelMsoEncodingEncodingUTF7 = '\000\213\375\350',
	ExcelMsoEncodingEncodingUTF8 = '\000\213\375\351'
};
typedef enum ExcelMsoEncoding ExcelMsoEncoding;

enum ExcelReset {
	ExcelResetCommandBar = 'msCB',
	ExcelResetCommandBarControl = 'mCBC'
};
typedef enum ExcelReset ExcelReset;

enum ExcelXlChartGallery {
	ExcelXlChartGalleryBuiltInChart = '\001\366\000\025',
	ExcelXlChartGalleryUserDefined = '\001\366\000\026',
	ExcelXlChartGalleryAnyGallery = '\001\366\000\027'
};
typedef enum ExcelXlChartGallery ExcelXlChartGallery;

enum ExcelXlColorIndex {
	ExcelXlColorIndexColorIndexAutomatic = '\001\366\357\367',
	ExcelXlColorIndexColorIndexNone = '\001\366\357\322',
	ExcelXlColorIndexAColorIndexInteger = '\001\367\000\000'
};
typedef enum ExcelXlColorIndex ExcelXlColorIndex;

enum ExcelXlEndStyleCap {
	ExcelXlEndStyleCapCap = '\001\370\000\001',
	ExcelXlEndStyleCapNoCap = '\001\370\000\002'
};
typedef enum ExcelXlEndStyleCap ExcelXlEndStyleCap;

enum ExcelXlRowCol {
	ExcelXlRowColByColumns = '\001\371\000\002',
	ExcelXlRowColByRows = '\001\371\000\001'
};
typedef enum ExcelXlRowCol ExcelXlRowCol;

enum ExcelXlScaleType {
	ExcelXlScaleTypeScaleLinear = '\001\371\357\334',
	ExcelXlScaleTypeScaleLogarithmic = '\001\371\357\333'
};
typedef enum ExcelXlScaleType ExcelXlScaleType;

enum ExcelXlDataSeriesType {
	ExcelXlDataSeriesTypeAutofillSeries = '\001\373\000\004',
	ExcelXlDataSeriesTypeChronologicalSeries = '\001\373\000\003',
	ExcelXlDataSeriesTypeGrowthSeries = '\001\373\000\002',
	ExcelXlDataSeriesTypeDataSeriesLinear = '\001\372\357\334'
};
typedef enum ExcelXlDataSeriesType ExcelXlDataSeriesType;

enum ExcelXlAxisCrosses {
	ExcelXlAxisCrossesAxisCrossesAutomatic = '\001\373\357\367',
	ExcelXlAxisCrossesAxisCrossesCustom = '\001\373\357\356',
	ExcelXlAxisCrossesAxisCrossesMaximum = '\001\374\000\002',
	ExcelXlAxisCrossesAxisCrossesMinimum = '\001\374\000\004'
};
typedef enum ExcelXlAxisCrosses ExcelXlAxisCrosses;

enum ExcelXlAxisGroup {
	ExcelXlAxisGroupPrimaryAxis = '\001\375\000\001',
	ExcelXlAxisGroupSecondaryAxis = '\001\375\000\002'
};
typedef enum ExcelXlAxisGroup ExcelXlAxisGroup;

enum ExcelXlBackground {
	ExcelXlBackgroundBackgroundAutomatic = '\001\375\357\367',
	ExcelXlBackgroundBackgroundOpaque = '\001\376\000\003',
	ExcelXlBackgroundBackgroundTransparent = '\001\376\000\002'
};
typedef enum ExcelXlBackground ExcelXlBackground;

enum ExcelXlWindowState {
	ExcelXlWindowStateWindowStateMaximized = '\001\376\357\327',
	ExcelXlWindowStateWindowStateMinimized = '\001\376\357\324',
	ExcelXlWindowStateWindowStateNormal = '\001\376\357\321'
};
typedef enum ExcelXlWindowState ExcelXlWindowState;

enum ExcelXlAxisType {
	ExcelXlAxisTypeCategoryAxis = '\002\000\000\001',
	ExcelXlAxisTypeSeriesAxis = '\002\000\000\003',
	ExcelXlAxisTypeValueAxis = '\002\000\000\002'
};
typedef enum ExcelXlAxisType ExcelXlAxisType;

enum ExcelXlArrowHeadLength {
	ExcelXlArrowHeadLengthArrowheadLengthLong = '\002\001\000\003',
	ExcelXlArrowHeadLengthArrowheadLengthMedium = '\002\000\357\326',
	ExcelXlArrowHeadLengthArrowheadLengthShort = '\002\001\000\001'
};
typedef enum ExcelXlArrowHeadLength ExcelXlArrowHeadLength;

enum ExcelXlVAlign {
	ExcelXlVAlignValignBottom = '\002\001\357\365',
	ExcelXlVAlignValignCenter = '\002\001\357\364',
	ExcelXlVAlignValignDistributed = '\002\001\357\353',
	ExcelXlVAlignValignJustify = '\002\001\357\336',
	ExcelXlVAlignValignTop = '\002\001\357\300'
};
typedef enum ExcelXlVAlign ExcelXlVAlign;

enum ExcelXlTickMark {
	ExcelXlTickMarkTickMarkCross = '\002\003\000\004',
	ExcelXlTickMarkTickMarkInside = '\002\003\000\002',
	ExcelXlTickMarkTickMarkNone = '\002\002\357\322',
	ExcelXlTickMarkTickMarkOutside = '\002\003\000\003'
};
typedef enum ExcelXlTickMark ExcelXlTickMark;

enum ExcelXlErrorBarDirection {
	ExcelXlErrorBarDirectionErrorBarDirectionX = '\002\003\357\270',
	ExcelXlErrorBarDirectionErrorBarDirectionY = '\002\004\000\001'
};
typedef enum ExcelXlErrorBarDirection ExcelXlErrorBarDirection;

enum ExcelXlErrorBarInclude {
	ExcelXlErrorBarIncludeErrorBarIncludeBoth = '\002\005\000\001',
	ExcelXlErrorBarIncludeErrorBarIncludeMinusValues = '\002\005\000\003',
	ExcelXlErrorBarIncludeErrorBarIncludeNone = '\002\004\357\322',
	ExcelXlErrorBarIncludeErrorBarIncludePlusValues = '\002\005\000\002'
};
typedef enum ExcelXlErrorBarInclude ExcelXlErrorBarInclude;

enum ExcelXlDisplayBlanksAs {
	ExcelXlDisplayBlanksAsInterpolated = '\002\006\000\003',
	ExcelXlDisplayBlanksAsNotPlotted = '\002\006\000\001',
	ExcelXlDisplayBlanksAsZero = '\002\006\000\002'
};
typedef enum ExcelXlDisplayBlanksAs ExcelXlDisplayBlanksAs;

enum ExcelXlArrowHeadStyle {
	ExcelXlArrowHeadStyleArrowheadStyleClosed = '\002\007\000\003',
	ExcelXlArrowHeadStyleArrowheadStyleDoubleClosed = '\002\007\000\005',
	ExcelXlArrowHeadStyleArrowheadStyleDoubleOpen = '\002\007\000\004',
	ExcelXlArrowHeadStyleArrowheadStyleNone = '\002\006\357\322',
	ExcelXlArrowHeadStyleArrowheadStyleOpen = '\002\007\000\002'
};
typedef enum ExcelXlArrowHeadStyle ExcelXlArrowHeadStyle;

enum ExcelXlArrowHeadWidth {
	ExcelXlArrowHeadWidthArrowheadWidthMedium = '\002\007\357\326',
	ExcelXlArrowHeadWidthArrowheadWidthNarrow = '\002\010\000\001',
	ExcelXlArrowHeadWidthArrowheadWidthWide = '\002\010\000\003'
};
typedef enum ExcelXlArrowHeadWidth ExcelXlArrowHeadWidth;

enum ExcelXlHAlign {
	ExcelXlHAlignHorizontalAlignCenter = '\002\010\357\364',
	ExcelXlHAlignHorizontalAlignCenterAcrossSelection = '\002\011\000\007',
	ExcelXlHAlignHorizontalAlignDistributed = '\002\010\357\353',
	ExcelXlHAlignHorizontalAlignFill = '\002\011\000\005',
	ExcelXlHAlignHorizontalAlignGeneral = '\002\011\000\001',
	ExcelXlHAlignHorizontalAlignJustify = '\002\010\357\336',
	ExcelXlHAlignHorizontalAlignLeft = '\002\010\357\335',
	ExcelXlHAlignHorizontalAlignRight = '\002\010\357\310'
};
typedef enum ExcelXlHAlign ExcelXlHAlign;

enum ExcelXlTickLabelPosition {
	ExcelXlTickLabelPositionTickLabelPositionHigh = '\002\011\357\341',
	ExcelXlTickLabelPositionTickLabelPositionLow = '\002\011\357\332',
	ExcelXlTickLabelPositionTickLabelPositionNextToAxis = '\002\012\000\004',
	ExcelXlTickLabelPositionTickLabelPositionNone = '\002\011\357\322'
};
typedef enum ExcelXlTickLabelPosition ExcelXlTickLabelPosition;

enum ExcelXlLegendPosition {
	ExcelXlLegendPositionLegendPositionBottom = '\002\012\357\365',
	ExcelXlLegendPositionLegendPositionCorner = '\002\013\000\002',
	ExcelXlLegendPositionLegendPositionLeft = '\002\012\357\335',
	ExcelXlLegendPositionLegendPositionRight = '\002\012\357\310',
	ExcelXlLegendPositionLegendPositionTop = '\002\012\357\300'
};
typedef enum ExcelXlLegendPosition ExcelXlLegendPosition;

enum ExcelXlChartPictureType {
	ExcelXlChartPictureTypeChartPictureTypeStackScale = '\002\014\000\003',
	ExcelXlChartPictureTypeChartPictureTypeStack = '\002\014\000\002',
	ExcelXlChartPictureTypeChartPictureTypeStretch = '\002\014\000\001'
};
typedef enum ExcelXlChartPictureType ExcelXlChartPictureType;

enum ExcelXlChartPicturePlacement {
	ExcelXlChartPicturePlacementSides = '\002\015\000\001',
	ExcelXlChartPicturePlacementEnd = '\002\015\000\002',
	ExcelXlChartPicturePlacementEndSides = '\002\015\000\003',
	ExcelXlChartPicturePlacementFront = '\002\015\000\004',
	ExcelXlChartPicturePlacementFrontSides = '\002\015\000\005',
	ExcelXlChartPicturePlacementFrontEnd = '\002\015\000\006',
	ExcelXlChartPicturePlacementAllFaces = '\002\015\000\007'
};
typedef enum ExcelXlChartPicturePlacement ExcelXlChartPicturePlacement;

enum ExcelXlOrientation {
	ExcelXlOrientationOrientationDownward = '\002\015\357\266',
	ExcelXlOrientationOrientationHorizontal = '\002\015\357\340',
	ExcelXlOrientationOrientationUpward = '\002\015\357\265',
	ExcelXlOrientationOrientationVertical = '\002\015\357\272'
};
typedef enum ExcelXlOrientation ExcelXlOrientation;

enum ExcelXlTickLabelOrientation {
	ExcelXlTickLabelOrientationTickLabelOrientationAutomatic = '\002\016\357\367',
	ExcelXlTickLabelOrientationTickLabelOrientationDownward = '\002\016\357\266',
	ExcelXlTickLabelOrientationTickLabelOrientationHorizontal = '\002\016\357\340',
	ExcelXlTickLabelOrientationTickLabelOrientationUpward = '\002\016\357\265',
	ExcelXlTickLabelOrientationTickLabelOrientationVertical = '\002\016\357\272'
};
typedef enum ExcelXlTickLabelOrientation ExcelXlTickLabelOrientation;

enum ExcelXlBorderWeight {
	ExcelXlBorderWeightBorderWeightHairline = '\002\020\000\001',
	ExcelXlBorderWeightBorderWeightMedium = '\002\017\357\326',
	ExcelXlBorderWeightBorderWeightThick = '\002\020\000\004',
	ExcelXlBorderWeightBorderWeightThin = '\002\020\000\002'
};
typedef enum ExcelXlBorderWeight ExcelXlBorderWeight;

enum ExcelXlDataSeriesDate {
	ExcelXlDataSeriesDateSeriesDateDay = '\002\021\000\001',
	ExcelXlDataSeriesDateSeriesDateMonth = '\002\021\000\003',
	ExcelXlDataSeriesDateSeriesDateWeekday = '\002\021\000\002',
	ExcelXlDataSeriesDateSeriesDateYear = '\002\021\000\004'
};
typedef enum ExcelXlDataSeriesDate ExcelXlDataSeriesDate;

enum ExcelXlUnderlineStyle {
	ExcelXlUnderlineStyleUnderlineStyleDouble = '\002\021\357\351',
	ExcelXlUnderlineStyleUnderlineStyleDoubleAccounting = '\002\022\000\005',
	ExcelXlUnderlineStyleUnderlineStyleNone = '\002\021\357\322',
	ExcelXlUnderlineStyleUnderlineStyleSingle = '\002\022\000\002',
	ExcelXlUnderlineStyleUnderlineStyleSingleAccounting = '\002\022\000\004'
};
typedef enum ExcelXlUnderlineStyle ExcelXlUnderlineStyle;

enum ExcelXlErrorBarType {
	ExcelXlErrorBarTypeErrorBarTypeCustom = '\002\022\357\356',
	ExcelXlErrorBarTypeErrorBarTypeFixedValue = '\002\023\000\001',
	ExcelXlErrorBarTypeErrorBarTypePercent = '\002\023\000\002',
	ExcelXlErrorBarTypeErrorBarTypeStandardDeviation = '\002\022\357\305',
	ExcelXlErrorBarTypeErrorBarTypeStandardError = '\002\023\000\004'
};
typedef enum ExcelXlErrorBarType ExcelXlErrorBarType;

enum ExcelXlTrendlineType {
	ExcelXlTrendlineTypeExponential = '\002\024\000\005',
	ExcelXlTrendlineTypeLinear = '\002\023\357\334',
	ExcelXlTrendlineTypeLogarithmic = '\002\023\357\333',
	ExcelXlTrendlineTypeMovingAverage = '\002\024\000\006',
	ExcelXlTrendlineTypePolynomial = '\002\024\000\003',
	ExcelXlTrendlineTypePower = '\002\024\000\004'
};
typedef enum ExcelXlTrendlineType ExcelXlTrendlineType;

enum ExcelXlLineStyle {
	ExcelXlLineStyleContinuous = '\002\025\000\001',
	ExcelXlLineStyleDash = '\002\024\357\355',
	ExcelXlLineStyleDashDot = '\002\025\000\004',
	ExcelXlLineStyleDashDotDot = '\002\025\000\005',
	ExcelXlLineStyleDot = '\002\024\357\352',
	ExcelXlLineStyleDouble = '\002\024\357\351',
	ExcelXlLineStyleSlantDashDot = '\002\025\000\015',
	ExcelXlLineStyleLineStyleNone = '\002\024\357\322'
};
typedef enum ExcelXlLineStyle ExcelXlLineStyle;

enum ExcelXlDataLabelsType {
	ExcelXlDataLabelsTypeDataLabelsShowNone = '\002\025\357\322',
	ExcelXlDataLabelsTypeDataLabelsShowValue = '\002\026\000\002',
	ExcelXlDataLabelsTypeDataLabelsShowPercent = '\002\026\000\003',
	ExcelXlDataLabelsTypeDataLabelsShowLabel = '\002\026\000\004',
	ExcelXlDataLabelsTypeDataLabelsShowLabelAndPercent = '\002\026\000\005',
	ExcelXlDataLabelsTypeDataLabelsShowBubbleSizes = '\002\026\000\006'
};
typedef enum ExcelXlDataLabelsType ExcelXlDataLabelsType;

enum ExcelXlMarkerStyle {
	ExcelXlMarkerStyleMarkerStyleAutomatic = '\002\026\357\367',
	ExcelXlMarkerStyleMarkerStyleCircle = '\002\027\000\010',
	ExcelXlMarkerStyleMarkerStyleDash = '\002\026\357\355',
	ExcelXlMarkerStyleMarkerStyleDiamond = '\002\027\000\002',
	ExcelXlMarkerStyleMarkerStyleDot = '\002\026\357\352',
	ExcelXlMarkerStyleMarkerStyleNone = '\002\026\357\322',
	ExcelXlMarkerStyleMarkerStylePicture = '\002\026\357\315',
	ExcelXlMarkerStyleMarkerStylePlus = '\002\027\000\011',
	ExcelXlMarkerStyleMarkerStyleSquare = '\002\027\000\001',
	ExcelXlMarkerStyleMarkerStyleStar = '\002\027\000\005',
	ExcelXlMarkerStyleMarkerStyleTriangle = '\002\027\000\003',
	ExcelXlMarkerStyleMarkerStyleX = '\002\026\357\270'
};
typedef enum ExcelXlMarkerStyle ExcelXlMarkerStyle;

enum ExcelXlPattern {
	ExcelXlPatternPatternAutomatic = '\002\030\357\367',
	ExcelXlPatternPatternChecker = '\002\031\000\011',
	ExcelXlPatternPatternCrissCross = '\002\031\000\020',
	ExcelXlPatternPatternDown = '\002\030\357\347',
	ExcelXlPatternPatternGray16 = '\002\031\000\021',
	ExcelXlPatternPatternGray25 = '\002\030\357\344',
	ExcelXlPatternPatternGray50 = '\002\030\357\343',
	ExcelXlPatternPatternGray75 = '\002\030\357\342',
	ExcelXlPatternPatternGray8 = '\002\031\000\022',
	ExcelXlPatternPatternGrid = '\002\031\000\017',
	ExcelXlPatternPatternHorizontal = '\002\030\357\340',
	ExcelXlPatternPatternLightDown = '\002\031\000\015',
	ExcelXlPatternPatternLightHorizontal = '\002\031\000\013',
	ExcelXlPatternPatternLightUp = '\002\031\000\016',
	ExcelXlPatternPatternLightVertical = '\002\031\000\014',
	ExcelXlPatternPatternNone = '\002\030\357\322',
	ExcelXlPatternPatternSemiGray75 = '\002\031\000\012',
	ExcelXlPatternPatternSolid = '\002\031\000\001',
	ExcelXlPatternPatternUp = '\002\030\357\276',
	ExcelXlPatternPatternVertical = '\002\030\357\272',
	ExcelXlPatternPatternLinearGradient = '\002\031\017\240',
	ExcelXlPatternPatternRectangularGradient = '\002\031\017\241'
};
typedef enum ExcelXlPattern ExcelXlPattern;

enum ExcelXlChartSplitType {
	ExcelXlChartSplitTypeSplitByPosition = '\002\032\000\001',
	ExcelXlChartSplitTypeSplitByPercentValue = '\002\032\000\003',
	ExcelXlChartSplitTypeSplitByCustomSplit = '\002\032\000\004',
	ExcelXlChartSplitTypeSplitByValue = '\002\032\000\002'
};
typedef enum ExcelXlChartSplitType ExcelXlChartSplitType;

enum ExcelXlDisplayUnit {
	ExcelXlDisplayUnitHundreds = '\002\032\377\376',
	ExcelXlDisplayUnitThousands = '\002\032\377\375',
	ExcelXlDisplayUnitTenThousands = '\002\032\377\374',
	ExcelXlDisplayUnitHundredThousands = '\002\032\377\373',
	ExcelXlDisplayUnitMillions = '\002\032\377\372',
	ExcelXlDisplayUnitTenMillions = '\002\032\377\371',
	ExcelXlDisplayUnitHundredMillions = '\002\032\377\370',
	ExcelXlDisplayUnitThousandMillions = '\002\032\377\367',
	ExcelXlDisplayUnitMillionMillions = '\002\032\377\366',
	ExcelXlDisplayUnitCustomDisplayUnit = '\002\032\357\356'
};
typedef enum ExcelXlDisplayUnit ExcelXlDisplayUnit;

enum ExcelXlDataLabelPosition {
	ExcelXlDataLabelPositionLabelPositionCenter = '\002\033\357\364',
	ExcelXlDataLabelPositionLabelPositionAbove = '\002\034\000\000',
	ExcelXlDataLabelPositionLabelPositionBelow = '\002\034\000\001',
	ExcelXlDataLabelPositionLabelPositionLeft = '\002\033\357\335',
	ExcelXlDataLabelPositionLabelPositionRight = '\002\033\357\310',
	ExcelXlDataLabelPositionLabelPositionOutsideEnd = '\002\034\000\002',
	ExcelXlDataLabelPositionLabelPositionInsideEnd = '\002\034\000\003',
	ExcelXlDataLabelPositionLabelPositionInsideBase = '\002\034\000\004',
	ExcelXlDataLabelPositionLabelPositionBestFit = '\002\034\000\005',
	ExcelXlDataLabelPositionLabelPositionMixed = '\002\034\000\006',
	ExcelXlDataLabelPositionLabelPositionCustom = '\002\034\000\007'
};
typedef enum ExcelXlDataLabelPosition ExcelXlDataLabelPosition;

enum ExcelXlTimeUnit {
	ExcelXlTimeUnitDays = '\002\035\000\000',
	ExcelXlTimeUnitMonths = '\002\035\000\001',
	ExcelXlTimeUnitYears = '\002\035\000\002'
};
typedef enum ExcelXlTimeUnit ExcelXlTimeUnit;

enum ExcelXlCategoryType {
	ExcelXlCategoryTypeCategoryScale = '\002\036\000\002',
	ExcelXlCategoryTypeTimeScale = '\002\036\000\003',
	ExcelXlCategoryTypeAutomaticScale = '\002\035\357\367'
};
typedef enum ExcelXlCategoryType ExcelXlCategoryType;

enum ExcelXlBarShape {
	ExcelXlBarShapeBox = '\002\037\000\000',
	ExcelXlBarShapePyramidToPoint = '\002\037\000\001',
	ExcelXlBarShapePyramidToMax = '\002\037\000\002',
	ExcelXlBarShapeCylinder = '\002\037\000\003',
	ExcelXlBarShapeConeToPoint = '\002\037\000\004',
	ExcelXlBarShapeConeToMax = '\002\037\000\005'
};
typedef enum ExcelXlBarShape ExcelXlBarShape;

enum ExcelXlChartType {
	ExcelXlChartTypeColumnClustered = '\002 \0003',
	ExcelXlChartTypeColumnStacked = '\002 \0004',
	ExcelXlChartTypeColumnStacked100 = '\002 \0005',
	ExcelXlChartTypeThreeDColumnClustered = '\002 \0006',
	ExcelXlChartTypeThreeDColumnStacked = '\002 \0007',
	ExcelXlChartTypeThreeDColumnStacked100 = '\002 \0008',
	ExcelXlChartTypeBarClustered = '\002 \0009',
	ExcelXlChartTypeBarStacked = '\002 \000:',
	ExcelXlChartTypeBarStacked100 = '\002 \000;',
	ExcelXlChartTypeThreeDBarClustered = '\002 \000<',
	ExcelXlChartTypeThreeDBarStacked = '\002 \000=',
	ExcelXlChartTypeThreeDBarStacked100 = '\002 \000>',
	ExcelXlChartTypeLineStacked = '\002 \000\?',
	ExcelXlChartTypeLineStacked100 = '\002 \000@',
	ExcelXlChartTypeLineMarkers = '\002 \000A',
	ExcelXlChartTypeLineMarkersStacked = '\002 \000B',
	ExcelXlChartTypeLineMarkersStacked100 = '\002 \000C',
	ExcelXlChartTypePieOfPie = '\002 \000D',
	ExcelXlChartTypePieExploded = '\002 \000E',
	ExcelXlChartTypeThreeDPieExploded = '\002 \000F',
	ExcelXlChartTypeBarOfPie = '\002 \000G',
	ExcelXlChartTypeXyScatterSmooth = '\002 \000H',
	ExcelXlChartTypeXyScatterSmoothNoMarkers = '\002 \000I',
	ExcelXlChartTypeXyScatterLines = '\002 \000J',
	ExcelXlChartTypeXyScatterLinesNoMarkers = '\002 \000K',
	ExcelXlChartTypeAreaStacked = '\002 \000L',
	ExcelXlChartTypeAreaStacked100 = '\002 \000M',
	ExcelXlChartTypeThreeDAreaStacked = '\002 \000N',
	ExcelXlChartTypeThreeDAreaStacked100 = '\002 \000O',
	ExcelXlChartTypeDoughnutExploded = '\002 \000P',
	ExcelXlChartTypeRadarMarkers = '\002 \000Q',
	ExcelXlChartTypeRadarFilled = '\002 \000R',
	ExcelXlChartTypeSurface = '\002 \000S',
	ExcelXlChartTypeSurfaceWireframe = '\002 \000T',
	ExcelXlChartTypeSurfaceTopView = '\002 \000U',
	ExcelXlChartTypeSurfaceTopViewWireframe = '\002 \000V',
	ExcelXlChartTypeBubble = '\002 \000\017',
	ExcelXlChartTypeBubbleThreeDEffect = '\002 \000W',
	ExcelXlChartTypeStockHLC = '\002 \000X',
	ExcelXlChartTypeStockOHLC = '\002 \000Y',
	ExcelXlChartTypeStockVHLC = '\002 \000Z',
	ExcelXlChartTypeStockVOHLC = '\002 \000[',
	ExcelXlChartTypeCylinderColumnClustered = '\002 \000\\',
	ExcelXlChartTypeCylinderColumnStacked = '\002 \000]',
	ExcelXlChartTypeCylinderColumnStacked100 = '\002 \000^',
	ExcelXlChartTypeCylinderBarClustered = '\002 \000_',
	ExcelXlChartTypeCylinderBarStacked = '\002 \000`',
	ExcelXlChartTypeCylinderBarStacked100 = '\002 \000a',
	ExcelXlChartTypeCylinderColumn = '\002 \000b',
	ExcelXlChartTypeConeColumnClustered = '\002 \000c',
	ExcelXlChartTypeConeColumnStacked = '\002 \000d',
	ExcelXlChartTypeConeColumnStacked100 = '\002 \000e',
	ExcelXlChartTypeConeBarClustered = '\002 \000f',
	ExcelXlChartTypeConeBarStacked = '\002 \000g',
	ExcelXlChartTypeConeBarStacked100 = '\002 \000h',
	ExcelXlChartTypeConeCol = '\002 \000i',
	ExcelXlChartTypePyramidColumnClustered = '\002 \000j',
	ExcelXlChartTypePyramidColumnStacked = '\002 \000k',
	ExcelXlChartTypePyramidColumnStacked100 = '\002 \000l',
	ExcelXlChartTypePyramidBarClustered = '\002 \000m',
	ExcelXlChartTypePyramidBarStacked = '\002 \000n',
	ExcelXlChartTypePyramidBarStacked100 = '\002 \000o',
	ExcelXlChartTypePyramidColumn = '\002 \000p',
	ExcelXlChartTypeThreeDColumn = '\002\037\357\374',
	ExcelXlChartTypeLineChart = '\002 \000\004',
	ExcelXlChartTypeThreeDLine = '\002\037\357\373',
	ExcelXlChartTypeThreeDPie = '\002\037\357\372',
	ExcelXlChartTypePieChart = '\002 \000\005',
	ExcelXlChartTypeXyscatter = '\002\037\357\267',
	ExcelXlChartTypeThreeDArea = '\002\037\357\376',
	ExcelXlChartTypeAreaChart = '\002 \000\001',
	ExcelXlChartTypeDoughnut = '\002\037\357\350',
	ExcelXlChartTypeRadar = '\002\037\357\311',
	ExcelXlChartTypeCombinationChart = '\002\037\357\361',
	ExcelXlChartTypeTreemap = '\002 \000u',
	ExcelXlChartTypeHistogram = '\002 \000v',
	ExcelXlChartTypeWaterfall = '\002 \000w',
	ExcelXlChartTypeSunburst = '\002 \000x',
	ExcelXlChartTypeBoxwhisker = '\002 \000y',
	ExcelXlChartTypePareto = '\002 \000z',
	ExcelXlChartTypeFunnel = '\002 \000{',
	ExcelXlChartTypeMapChart = '\002 \000\214'
};
typedef enum ExcelXlChartType ExcelXlChartType;

enum ExcelXlChartItem {
	ExcelXlChartItemDataLabel = '\002!\000\000',
	ExcelXlChartItemAChartArea = '\002!\000\002',
	ExcelXlChartItemASeries = '\002!\000\003',
	ExcelXlChartItemAChartTitle = '\002!\000\004',
	ExcelXlChartItemWalls = '\002!\000\005',
	ExcelXlChartItemACornersObject = '\002!\000\006',
	ExcelXlChartItemDataTable = '\002!\000\007',
	ExcelXlChartItemTrendline = '\002!\000\010',
	ExcelXlChartItemErrorBarsObject = '\002!\000\011',
	ExcelXlChartItemXerrorBars = '\002!\000\012',
	ExcelXlChartItemYerrorBars = '\002!\000\013',
	ExcelXlChartItemLegendEntry = '\002!\000\014',
	ExcelXlChartItemLegendKey = '\002!\000\015',
	ExcelXlChartItemShape = '\002!\000\016',
	ExcelXlChartItemMajorGridlines = '\002!\000\017',
	ExcelXlChartItemMinorGridlines = '\002!\000\020',
	ExcelXlChartItemAxisTitle = '\002!\000\021',
	ExcelXlChartItemUpBars = '\002!\000\022',
	ExcelXlChartItemPlotArea = '\002!\000\023',
	ExcelXlChartItemDownBars = '\002!\000\024',
	ExcelXlChartItemAxis = '\002!\000\025',
	ExcelXlChartItemSeriesLines = '\002!\000\026',
	ExcelXlChartItemFloor = '\002!\000\027',
	ExcelXlChartItemLegend = '\002!\000\030',
	ExcelXlChartItemHiLoLines = '\002!\000\031',
	ExcelXlChartItemDropLines = '\002!\000\032',
	ExcelXlChartItemRadarAxisLabels = '\002!\000\033',
	ExcelXlChartItemNothing = '\002!\000\034',
	ExcelXlChartItemLeaderLines = '\002!\000\035',
	ExcelXlChartItemDisplayUnitLabel = '\002!\000\036'
};
typedef enum ExcelXlChartItem ExcelXlChartItem;

enum ExcelXlSizeRepresents {
	ExcelXlSizeRepresentsSizeIsWidth = '\002\"\000\002',
	ExcelXlSizeRepresentsSizeIsArea = '\002\"\000\001'
};
typedef enum ExcelXlSizeRepresents ExcelXlSizeRepresents;

enum ExcelXlInsertShiftDirection {
	ExcelXlInsertShiftDirectionShiftDown = '\002\"\357\347',
	ExcelXlInsertShiftDirectionShiftToRight = '\002\"\357\277'
};
typedef enum ExcelXlInsertShiftDirection ExcelXlInsertShiftDirection;

enum ExcelXlDeleteShiftDirection {
	ExcelXlDeleteShiftDirectionShiftToLeft = '\002#\357\301',
	ExcelXlDeleteShiftDirectionShiftUp = '\002#\357\276'
};
typedef enum ExcelXlDeleteShiftDirection ExcelXlDeleteShiftDirection;

enum ExcelXlDirection {
	ExcelXlDirectionTowardTheBottom = '\002$\357\347',
	ExcelXlDirectionTowardTheLeft = '\002$\357\301',
	ExcelXlDirectionTowardTheRight = '\002$\357\277',
	ExcelXlDirectionTowardTheTop = '\002$\357\276'
};
typedef enum ExcelXlDirection ExcelXlDirection;

enum ExcelXlConsolidationFunction {
	ExcelXlConsolidationFunctionDoAverage = '\002%\357\366',
	ExcelXlConsolidationFunctionDoCount = '\002%\357\360',
	ExcelXlConsolidationFunctionDoCountNumbers = '\002%\357\357',
	ExcelXlConsolidationFunctionDoMaximum = '\002%\357\330',
	ExcelXlConsolidationFunctionDoMinimum = '\002%\357\325',
	ExcelXlConsolidationFunctionDoProduct = '\002%\357\313',
	ExcelXlConsolidationFunctionDoStandardDeviation = '\002%\357\305',
	ExcelXlConsolidationFunctionDoStandardDeviationP = '\002%\357\304',
	ExcelXlConsolidationFunctionDoSum = '\002%\357\303',
	ExcelXlConsolidationFunctionDoVar = '\002%\357\274',
	ExcelXlConsolidationFunctionDoVarP = '\002%\357\273'
};
typedef enum ExcelXlConsolidationFunction ExcelXlConsolidationFunction;

enum ExcelXlSheetType {
	ExcelXlSheetTypeSheetTypeChart = '\002&\357\363',
	ExcelXlSheetTypeSheetTypeDialogSheet = '\002&\357\354',
	ExcelXlSheetTypeSheetTypeExcel4IntlMacroSheet = '\002\'\000\004',
	ExcelXlSheetTypeSheetTypeExcel4MacroSheet = '\002\'\000\003',
	ExcelXlSheetTypeSheetTypeWorksheet = '\002&\357\271'
};
typedef enum ExcelXlSheetType ExcelXlSheetType;

enum ExcelXlLocationInTable {
	ExcelXlLocationInTableColumnHeader = '\002\'\357\362',
	ExcelXlLocationInTableColumnItem = '\002(\000\005',
	ExcelXlLocationInTableDataHeader = '\002(\000\003',
	ExcelXlLocationInTableDataItem = '\002(\000\007',
	ExcelXlLocationInTablePageHeader = '\002(\000\002',
	ExcelXlLocationInTablePageItem = '\002(\000\006',
	ExcelXlLocationInTableRowHeader = '\002\'\357\307',
	ExcelXlLocationInTableRowItem = '\002(\000\004',
	ExcelXlLocationInTableTableBody = '\002(\000\010'
};
typedef enum ExcelXlLocationInTable ExcelXlLocationInTable;

enum ExcelXlFindLookIn {
	ExcelXlFindLookInFormulas = '\002(\357\345',
	ExcelXlFindLookInComments = '\002(\357\320',
	ExcelXlFindLookInValues = '\002(\357\275'
};
typedef enum ExcelXlFindLookIn ExcelXlFindLookIn;

enum ExcelXlWindowType {
	ExcelXlWindowTypeWindowTypeChartAsWindow = '\002*\000\005',
	ExcelXlWindowTypeWindowTypeChartInPlace = '\002*\000\004',
	ExcelXlWindowTypeWindowTypeClipboard = '\002*\000\003',
	ExcelXlWindowTypeWindowTypeInfo = '\002)\357\337',
	ExcelXlWindowTypeWindowTypeWorkbook = '\002*\000\001'
};
typedef enum ExcelXlWindowType ExcelXlWindowType;

enum ExcelXlPivotFieldDataType {
	ExcelXlPivotFieldDataTypePivotFieldTypeDate = '\002+\000\002',
	ExcelXlPivotFieldDataTypePivotFieldTypeNumber = '\002*\357\317',
	ExcelXlPivotFieldDataTypePivotFieldTypeText = '\002*\357\302'
};
typedef enum ExcelXlPivotFieldDataType ExcelXlPivotFieldDataType;

enum ExcelXlCopyPictureFormat {
	ExcelXlCopyPictureFormatBitmap = '\002,\000\002',
	ExcelXlCopyPictureFormatPicture = '\002+\357\315'
};
typedef enum ExcelXlCopyPictureFormat ExcelXlCopyPictureFormat;

enum ExcelXlPivotTableSourceType {
	ExcelXlPivotTableSourceTypeConsolidation = '\002-\000\003',
	ExcelXlPivotTableSourceTypeDatabase = '\002-\000\001',
	ExcelXlPivotTableSourceTypeExternal = '\002-\000\002',
	ExcelXlPivotTableSourceTypePivotTable = '\002,\357\314'
};
typedef enum ExcelXlPivotTableSourceType ExcelXlPivotTableSourceType;

enum ExcelXlReferenceStyle {
	ExcelXlReferenceStyleA1 = '\002.\000\001',
	ExcelXlReferenceStyleR1C1 = '\002-\357\312'
};
typedef enum ExcelXlReferenceStyle ExcelXlReferenceStyle;

enum ExcelXlMSApplication {
	ExcelXlMSApplicationMicrosoftAccess = '\002/\000\004',
	ExcelXlMSApplicationMicrosoftFoxPro = '\002/\000\005',
	ExcelXlMSApplicationMicrosoftMail = '\002/\000\003',
	ExcelXlMSApplicationMicrosoftPowerPoint = '\002/\000\002',
	ExcelXlMSApplicationMicrosoftProject = '\002/\000\006',
	ExcelXlMSApplicationMicrosoftSchedulePlus = '\002/\000\007',
	ExcelXlMSApplicationMicrosoftWord = '\002/\000\001'
};
typedef enum ExcelXlMSApplication ExcelXlMSApplication;

enum ExcelXlMouseButton {
	ExcelXlMouseButtonNoButton = '\0020\000\000',
	ExcelXlMouseButtonPrimaryButton = '\0020\000\001',
	ExcelXlMouseButtonSecondaryButton = '\0020\000\002'
};
typedef enum ExcelXlMouseButton ExcelXlMouseButton;

enum ExcelXlCutCopyMode {
	ExcelXlCutCopyModeCopyMode = '\0021\000\001',
	ExcelXlCutCopyModeCutMode = '\0021\000\002'
};
typedef enum ExcelXlCutCopyMode ExcelXlCutCopyMode;

enum ExcelXlFilterAction {
	ExcelXlFilterActionFilterCopy = '\0023\000\002',
	ExcelXlFilterActionFilterInPlace = '\0023\000\001'
};
typedef enum ExcelXlFilterAction ExcelXlFilterAction;

enum ExcelXlFormulaVersion {
	ExcelXlFormulaVersionUsingFormula = '\002\327\000\000',
	ExcelXlFormulaVersionUsingFormula2 = '\002\327\000\001'
};
typedef enum ExcelXlFormulaVersion ExcelXlFormulaVersion;

enum ExcelXlOrder {
	ExcelXlOrderDownThenOver = '\0024\000\001',
	ExcelXlOrderOverThenDown = '\0024\000\002'
};
typedef enum ExcelXlOrder ExcelXlOrder;

enum ExcelXlLinkType {
	ExcelXlLinkTypeLinkTypeExcelLinks = '\0025\000\001',
	ExcelXlLinkTypeLinkTypeOLELinks = '\0025\000\002'
};
typedef enum ExcelXlLinkType ExcelXlLinkType;

enum ExcelXlApplyNamesOrder {
	ExcelXlApplyNamesOrderColumnThenRow = '\0026\000\002',
	ExcelXlApplyNamesOrderRowThenColumn = '\0026\000\001'
};
typedef enum ExcelXlApplyNamesOrder ExcelXlApplyNamesOrder;

enum ExcelXlEnableCancelKey {
	ExcelXlEnableCancelKeyCancelKeyDisabled = '\0027\000\000',
	ExcelXlEnableCancelKeyErrorHandler = '\0027\000\002',
	ExcelXlEnableCancelKeyInterrupt = '\0027\000\001'
};
typedef enum ExcelXlEnableCancelKey ExcelXlEnableCancelKey;

enum ExcelXlPageBreak {
	ExcelXlPageBreakPageBreakAutomatic = '\0027\357\367',
	ExcelXlPageBreakPageBreakManual = '\0027\357\331',
	ExcelXlPageBreakPageBreakNone = '\0028\000\000'
};
typedef enum ExcelXlPageBreak ExcelXlPageBreak;

enum ExcelXlPageOrientation {
	ExcelXlPageOrientationLandscape = '\002:\000\002',
	ExcelXlPageOrientationPortrait = '\002:\000\001'
};
typedef enum ExcelXlPageOrientation ExcelXlPageOrientation;

enum ExcelXlLinkInfo {
	ExcelXlLinkInfoEditionDate = '\002;\000\002',
	ExcelXlLinkInfoUpdateState = '\002;\000\001'
};
typedef enum ExcelXlLinkInfo ExcelXlLinkInfo;

enum ExcelXlCommandUnderlines {
	ExcelXlCommandUnderlinesCommandUnderlinesAutomatic = '\002;\357\367',
	ExcelXlCommandUnderlinesCommandUnderlinesOff = '\002;\357\316',
	ExcelXlCommandUnderlinesCommandUnderlinesOn = '\002<\000\001'
};
typedef enum ExcelXlCommandUnderlines ExcelXlCommandUnderlines;

enum ExcelXlOLEVerb {
	ExcelXlOLEVerbVerbOpen = '\002=\000\002',
	ExcelXlOLEVerbVerbPrimary = '\002=\000\001'
};
typedef enum ExcelXlOLEVerb ExcelXlOLEVerb;

enum ExcelXlCalculation {
	ExcelXlCalculationCalculationAutomatic = '\002=\357\367',
	ExcelXlCalculationCalculationManual = '\002=\357\331',
	ExcelXlCalculationCalculationSemiautomatic = '\002>\000\002'
};
typedef enum ExcelXlCalculation ExcelXlCalculation;

enum ExcelXlFileAccess {
	ExcelXlFileAccessWorkbookReadOnly = '\002\?\000\003',
	ExcelXlFileAccessWorkbookReadWrite = '\002\?\000\002'
};
typedef enum ExcelXlFileAccess ExcelXlFileAccess;

enum ExcelXlObjectSize {
	ExcelXlObjectSizeFitToPage = '\002@\000\002',
	ExcelXlObjectSizeFullPage = '\002@\000\003',
	ExcelXlObjectSizeFullScreen = '\002@\000\001'
};
typedef enum ExcelXlObjectSize ExcelXlObjectSize;

enum ExcelXlLookAt {
	ExcelXlLookAtPart = '\002A\000\002',
	ExcelXlLookAtWhole = '\002A\000\001'
};
typedef enum ExcelXlLookAt ExcelXlLookAt;

enum ExcelXlMailSystem {
	ExcelXlMailSystemMAPI = '\002B\000\001',
	ExcelXlMailSystemNoMailSystem = '\002B\000\000',
	ExcelXlMailSystemPowerTalk = '\002B\000\002'
};
typedef enum ExcelXlMailSystem ExcelXlMailSystem;

enum ExcelXlLinkInfoType {
	ExcelXlLinkInfoTypeLinkInfoOlelinks = '\002C\000\002',
	ExcelXlLinkInfoTypeLinkInfoPublishers = '\002C\000\005',
	ExcelXlLinkInfoTypeLinkInfoSubscribers = '\002C\000\006'
};
typedef enum ExcelXlLinkInfoType ExcelXlLinkInfoType;

enum ExcelXlCellType {
	ExcelXlCellTypeCellTypeBlanks = '\002F\000\004',
	ExcelXlCellTypeCellTypeConstants = '\002F\000\002',
	ExcelXlCellTypeCellTypeFormulas = '\002E\357\345',
	ExcelXlCellTypeCellTypeLastCell = '\002F\000\013',
	ExcelXlCellTypeCellTypeComments = '\002E\357\320',
	ExcelXlCellTypeCellTypeVisible = '\002F\000\014',
	ExcelXlCellTypeCellTypeAllFormatConditions = '\002E\357\264',
	ExcelXlCellTypeCellTypeSameFormatConditions = '\002E\357\263',
	ExcelXlCellTypeCellTypeAllValidation = '\002E\357\262',
	ExcelXlCellTypeCellTypeSameValidation = '\002E\357\261'
};
typedef enum ExcelXlCellType ExcelXlCellType;

enum ExcelXlArrangeStyle {
	ExcelXlArrangeStyleArrangeStyleCascade = '\002G\000\007',
	ExcelXlArrangeStyleArrangeStyleHorizontal = '\002F\357\340',
	ExcelXlArrangeStyleArrangeStyleTiled = '\002G\000\001',
	ExcelXlArrangeStyleArrangeStyleVertical = '\002F\357\272'
};
typedef enum ExcelXlArrangeStyle ExcelXlArrangeStyle;

enum ExcelXlMousePointer {
	ExcelXlMousePointerIBeamCursor = '\002H\000\003',
	ExcelXlMousePointerDefaultCursor = '\002G\357\321',
	ExcelXlMousePointerNorthwestArrowCursor = '\002H\000\001',
	ExcelXlMousePointerWaitCursor = '\002H\000\002'
};
typedef enum ExcelXlMousePointer ExcelXlMousePointer;

enum ExcelXlAutoFillType {
	ExcelXlAutoFillTypeFillDefault = '\002I\000\000',
	ExcelXlAutoFillTypeFillCopy = '\002I\000\001',
	ExcelXlAutoFillTypeFillSeries = '\002I\000\002',
	ExcelXlAutoFillTypeFillFormats = '\002I\000\003',
	ExcelXlAutoFillTypeFillValues = '\002I\000\004',
	ExcelXlAutoFillTypeFillDays = '\002I\000\005',
	ExcelXlAutoFillTypeFillWeekdays = '\002I\000\006',
	ExcelXlAutoFillTypeFillMonths = '\002I\000\007',
	ExcelXlAutoFillTypeFillYears = '\002I\000\010',
	ExcelXlAutoFillTypeLinearTrend = '\002I\000\011',
	ExcelXlAutoFillTypeGrowthTrend = '\002I\000\012',
	ExcelXlAutoFillTypeFlashfill = '\002I\000\013'
};
typedef enum ExcelXlAutoFillType ExcelXlAutoFillType;

enum ExcelXlAutoFilterOperator {
	ExcelXlAutoFilterOperatorAutofilterAnd = '\002J\000\001',
	ExcelXlAutoFilterOperatorBottom10Items = '\002J\000\004',
	ExcelXlAutoFilterOperatorBottom10Percent = '\002J\000\006',
	ExcelXlAutoFilterOperatorAutofilterOr = '\002J\000\002',
	ExcelXlAutoFilterOperatorTop10Items = '\002J\000\003',
	ExcelXlAutoFilterOperatorTop10Percent = '\002J\000\005',
	ExcelXlAutoFilterOperatorFilterByValue = '\002J\000\007',
	ExcelXlAutoFilterOperatorFilterByCellColor = '\002J\000\010',
	ExcelXlAutoFilterOperatorFilterByFontColor = '\002J\000\011',
	ExcelXlAutoFilterOperatorFilterByIcon = '\002J\000\012',
	ExcelXlAutoFilterOperatorFilterDynamic = '\002J\000\013',
	ExcelXlAutoFilterOperatorFilterNoFill = '\002J\000\014',
	ExcelXlAutoFilterOperatorFilterByAutomaticFontColor = '\002J\000\015',
	ExcelXlAutoFilterOperatorFilterByNoIcon = '\002J\000\016'
};
typedef enum ExcelXlAutoFilterOperator ExcelXlAutoFilterOperator;

enum ExcelXlClipboardFormat {
	ExcelXlClipboardFormatClipboardFormatBiff = '\002K\000\010',
	ExcelXlClipboardFormatClipboardFormatBiff2 = '\002K\000\022',
	ExcelXlClipboardFormatClipboardFormatBiff3 = '\002K\000\024',
	ExcelXlClipboardFormatClipboardFormatBiff4 = '\002K\000\036',
	ExcelXlClipboardFormatClipboardFormatBinary = '\002K\000\017',
	ExcelXlClipboardFormatClipboardFormatBitmap = '\002K\000\011',
	ExcelXlClipboardFormatClipboardFormatCgm = '\002K\000\015',
	ExcelXlClipboardFormatClipboardFormatCsv = '\002K\000\005',
	ExcelXlClipboardFormatClipboardFormatDif = '\002K\000\004',
	ExcelXlClipboardFormatClipboardFormatDspText = '\002K\000\014',
	ExcelXlClipboardFormatClipboardFormatEmbeddedObject = '\002K\000\025',
	ExcelXlClipboardFormatClipboardFormatEmbedSource = '\002K\000\026',
	ExcelXlClipboardFormatClipboardFormatLink = '\002K\000\013',
	ExcelXlClipboardFormatClipboardFormatLinkSource = '\002K\000\027',
	ExcelXlClipboardFormatClipboardFormatLinkSourceDesc = '\002K\000 ',
	ExcelXlClipboardFormatClipboardFormatMovie = '\002K\000\030',
	ExcelXlClipboardFormatClipboardFormatNative = '\002K\000\016',
	ExcelXlClipboardFormatClipboardFormatObjectDesc = '\002K\000\037',
	ExcelXlClipboardFormatClipboardFormatObjectLink = '\002K\000\023',
	ExcelXlClipboardFormatClipboardFormatOwnerLink = '\002K\000\021',
	ExcelXlClipboardFormatClipboardFormatPict = '\002K\000\002',
	ExcelXlClipboardFormatClipboardFormatPrintPict = '\002K\000\003',
	ExcelXlClipboardFormatClipboardFormatRtf = '\002K\000\007',
	ExcelXlClipboardFormatClipboardFormatScreenPict = '\002K\000\035',
	ExcelXlClipboardFormatClipboardFormatStandardFont = '\002K\000\034',
	ExcelXlClipboardFormatClipboardFormatStandardScale = '\002K\000\033',
	ExcelXlClipboardFormatClipboardFormatSylk = '\002K\000\006',
	ExcelXlClipboardFormatClipboardFormatTable = '\002K\000\020',
	ExcelXlClipboardFormatClipboardFormatText = '\002K\000\000',
	ExcelXlClipboardFormatClipboardFormatToolFace = '\002K\000\031',
	ExcelXlClipboardFormatClipboardFormatToolFacePict = '\002K\000\032',
	ExcelXlClipboardFormatClipboardFormatValu = '\002K\000\001',
	ExcelXlClipboardFormatClipboardFormatWk1 = '\002K\000\012',
	ExcelXlClipboardFormatClipboardFormatUnicodeText = '\002K\000.',
	ExcelXlClipboardFormatClipboardFormatStyleText = '\002K\0005',
	ExcelXlClipboardFormatClipboardFormatUnicodeStyleText = '\002K\0007',
	ExcelXlClipboardFormatClipboardFormatBiff5 = '\002K\000!',
	ExcelXlClipboardFormatClipboardFormatPictureBuild = '\002K\000\"',
	ExcelXlClipboardFormatClipboardFormatOdbcConn = '\002K\000#',
	ExcelXlClipboardFormatClipboardFormatOdbcSql = '\002K\000$',
	ExcelXlClipboardFormatClipboardFormat3dPicture = '\002K\000%',
	ExcelXlClipboardFormatClipboardFormatUnexpected38 = '\002K\000&',
	ExcelXlClipboardFormatClipboardFormatDrawingDragDrop = '\002K\000\'',
	ExcelXlClipboardFormatClipboardFormatDrawing = '\002K\000(',
	ExcelXlClipboardFormatClipboardFormatUnexpected41 = '\002K\000)',
	ExcelXlClipboardFormatClipboardFormatUnexpected42 = '\002K\000*',
	ExcelXlClipboardFormatClipboardFormatUnexpected43 = '\002K\000+',
	ExcelXlClipboardFormatClipboardFormatHyperlink = '\002K\000,',
	ExcelXlClipboardFormatClipboardFormatUnexpected45 = '\002K\000-',
	ExcelXlClipboardFormatClipboardFormatWindowsBitmap = '\002K\000/',
	ExcelXlClipboardFormatClipboardFormatUniformResourceLocator = '\002K\0000',
	ExcelXlClipboardFormatClipboardFormatFileName = '\002K\0001',
	ExcelXlClipboardFormatClipboardFormatUnexpected50 = '\002K\0002',
	ExcelXlClipboardFormatClipboardFormatUnexpected51 = '\002K\0003',
	ExcelXlClipboardFormatClipboardFormatHypertextMarkupLanguage = '\002K\0004',
	ExcelXlClipboardFormatClipboardFormatOfficeScrapbookInfo = '\002K\0006',
	ExcelXlClipboardFormatClipboardFormatPortableDocumentFormat = '\002K\0008',
	ExcelXlClipboardFormatClipboardFormatExcelInternalShape = '\002K\0009',
	ExcelXlClipboardFormatClipboardFormatOfficeArtShape = '\002K\000:'
};
typedef enum ExcelXlClipboardFormat ExcelXlClipboardFormat;

enum ExcelXlFileFormat {
	ExcelXlFileFormatCSVFileFormat = '\002\274\000\006',
	ExcelXlFileFormatCSVMacFileFormat = '\002\274\000\026',
	ExcelXlFileFormatCSVMSDosFileFormat = '\002\274\000\030',
	ExcelXlFileFormatCSVWindowsFileFormat = '\002\274\000\027',
	ExcelXlFileFormatDBF3FileFormat = '\002\274\000\010',
	ExcelXlFileFormatDBF4FileFormat = '\002\274\000\013',
	ExcelXlFileFormatDIFFileFormat = '\002\274\000\011',
	ExcelXlFileFormatExcel2FileFormat = '\002\274\000\020',
	ExcelXlFileFormatExcel2EastAsianFileFormat = '\002\274\000\033',
	ExcelXlFileFormatExcel3FileFormat = '\002\274\000\035',
	ExcelXlFileFormatExcel4FileFormat = '\002\274\000!',
	ExcelXlFileFormatExcel5FileFormat = '\002\274\000\'',
	ExcelXlFileFormatExcel7FileFormat = '\002\274\000\'',
	ExcelXlFileFormatExcel4WorkbookFileFormat = '\002\274\000#',
	ExcelXlFileFormatInternationalAddInFileFormat = '\002\274\000\032',
	ExcelXlFileFormatInternationalMacroFileFormat = '\002\274\000\031',
	ExcelXlFileFormatWorkbookNormalFileFormat = '\002\274\0003',
	ExcelXlFileFormatSYLKFileFormat = '\002\274\000\002',
	ExcelXlFileFormatCurrentPlatformTextFileFormat = '\002\273\357\302',
	ExcelXlFileFormatTextMacFileFormat = '\002\274\000\023',
	ExcelXlFileFormatTextMSDosFileFormat = '\002\274\000\025',
	ExcelXlFileFormatTextPrinterFileFormat = '\002\274\000$',
	ExcelXlFileFormatTextWindowsFileFormat = '\002\274\000\024',
	ExcelXlFileFormatHTMLFileFormat = '\002\274\000-',
	ExcelXlFileFormatXMLSpreadsheetFileFormat = '\002\274\000.',
	ExcelXlFileFormatPDFFileFormat = '\002\274\0009',
	ExcelXlFileFormatExcelBinaryFileFormat = '\002\274\0002',
	ExcelXlFileFormatExcelXMLFileFormat = '\002\274\0003',
	ExcelXlFileFormatMacroEnabledXMLFileFormat = '\002\274\0004',
	ExcelXlFileFormatMacroEnabledTemplateFileFormat = '\002\274\0005',
	ExcelXlFileFormatTemplateFileFormat = '\002\274\0006',
	ExcelXlFileFormatAddInFileFormat = '\002\274\0007',
	ExcelXlFileFormatExcel98to2004FileFormat = '\002\274\0008',
	ExcelXlFileFormatExcel98to2004TemplateFileFormat = '\002\274\000\021',
	ExcelXlFileFormatExcel98to2004AddInFileFormat = '\002\274\000\022'
};
typedef enum ExcelXlFileFormat ExcelXlFileFormat;

enum ExcelXlApplicationInternational {
	ExcelXlApplicationInternationalTwenty_four_hourClock = '\002M\000!',
	ExcelXlApplicationInternationalFourDigitYears = '\002M\000+',
	ExcelXlApplicationInternationalAlternateArraySeparator = '\002M\000\020',
	ExcelXlApplicationInternationalColumnSeparator = '\002M\000\016',
	ExcelXlApplicationInternationalCountry_code = '\002M\000\001',
	ExcelXlApplicationInternationalCountry_setting = '\002M\000\002',
	ExcelXlApplicationInternationalCurrency_before = '\002M\000%',
	ExcelXlApplicationInternationalCurrency_code = '\002M\000\031',
	ExcelXlApplicationInternationalCurrency_digits = '\002M\000\033',
	ExcelXlApplicationInternationalCurrency_leading_zeros = '\002M\000(',
	ExcelXlApplicationInternationalCurrency_minus_sign = '\002M\000&',
	ExcelXlApplicationInternationalCurrency_negative = '\002M\000\034',
	ExcelXlApplicationInternationalCurrency_space_before = '\002M\000$',
	ExcelXlApplicationInternationalCurrency_trailing_zeros = '\002M\000\'',
	ExcelXlApplicationInternationalDate_order = '\002M\000 ',
	ExcelXlApplicationInternationalDate_separator = '\002M\000\021',
	ExcelXlApplicationInternationalDayCode = '\002M\000\025',
	ExcelXlApplicationInternationalDayLeadingZero = '\002M\000*',
	ExcelXlApplicationInternationalDecimalSeparator = '\002M\000\003',
	ExcelXlApplicationInternationalGeneralFormatName = '\002M\000\032',
	ExcelXlApplicationInternationalHourCode = '\002M\000\026',
	ExcelXlApplicationInternationalLeftBrace = '\002M\000\014',
	ExcelXlApplicationInternationalLeftBracket = '\002M\000\012',
	ExcelXlApplicationInternationalListSeparator = '\002M\000\005',
	ExcelXlApplicationInternationalLowerCaseColumnLetter = '\002M\000\011',
	ExcelXlApplicationInternationalLowerCaseRowLetter = '\002M\000\010',
	ExcelXlApplicationInternationalMdy = '\002M\000,',
	ExcelXlApplicationInternationalMetric = '\002M\000#',
	ExcelXlApplicationInternationalMinute_code = '\002M\000\027',
	ExcelXlApplicationInternationalMonth_code = '\002M\000\024',
	ExcelXlApplicationInternationalMonth_leading_zero = '\002M\000)',
	ExcelXlApplicationInternationalMonth_name_chars = '\002M\000\036',
	ExcelXlApplicationInternationalNoncurrency_digits = '\002M\000\035',
	ExcelXlApplicationInternationalNonEnglishFunctions = '\002M\000\"',
	ExcelXlApplicationInternationalRightBrace = '\002M\000\015',
	ExcelXlApplicationInternationalRightBracket = '\002M\000\013',
	ExcelXlApplicationInternationalRowSeparator = '\002M\000\017',
	ExcelXlApplicationInternationalSecondCode = '\002M\000\030',
	ExcelXlApplicationInternationalThousandsSeparator = '\002M\000\004',
	ExcelXlApplicationInternationalTimeLeadingZero = '\002M\000-',
	ExcelXlApplicationInternationalTimeSeparator = '\002M\000\022',
	ExcelXlApplicationInternationalUpperCaseColumnLetter = '\002M\000\007',
	ExcelXlApplicationInternationalUpperCaseRowLetter = '\002M\000\006',
	ExcelXlApplicationInternationalWeekday_name_chars = '\002M\000\037',
	ExcelXlApplicationInternationalYearCode = '\002M\000\023'
};
typedef enum ExcelXlApplicationInternational ExcelXlApplicationInternational;

enum ExcelXlPageBreakExtent {
	ExcelXlPageBreakExtentPageBreakFull = '\002N\000\001',
	ExcelXlPageBreakExtentPageBreakPartial = '\002N\000\002'
};
typedef enum ExcelXlPageBreakExtent ExcelXlPageBreakExtent;

enum ExcelXlCellInsertionMode {
	ExcelXlCellInsertionModeOverwriteCells = '\002O\000\000',
	ExcelXlCellInsertionModeInsertDeleteCells = '\002O\000\001',
	ExcelXlCellInsertionModeInsertEntireRows = '\002O\000\002'
};
typedef enum ExcelXlCellInsertionMode ExcelXlCellInsertionMode;

enum ExcelXlFormulaLabel {
	ExcelXlFormulaLabelNoLabels = '\002O\357\322',
	ExcelXlFormulaLabelRowLabels = '\002P\000\001',
	ExcelXlFormulaLabelColumnLabels = '\002P\000\002',
	ExcelXlFormulaLabelMixedLabels = '\002P\000\003'
};
typedef enum ExcelXlFormulaLabel ExcelXlFormulaLabel;

enum ExcelXlHighlightChangesTime {
	ExcelXlHighlightChangesTimeSinceMyLastSave = '\002Q\000\001',
	ExcelXlHighlightChangesTimeAllChanges = '\002Q\000\002',
	ExcelXlHighlightChangesTimeNotYetReviewed = '\002Q\000\003'
};
typedef enum ExcelXlHighlightChangesTime ExcelXlHighlightChangesTime;

enum ExcelXlCommentDisplayMode {
	ExcelXlCommentDisplayModeNoIndicator = '\002R\000\000',
	ExcelXlCommentDisplayModeCommentIndicatorOnly = '\002Q\377\377',
	ExcelXlCommentDisplayModeCommentAndIndicator = '\002R\000\001'
};
typedef enum ExcelXlCommentDisplayMode ExcelXlCommentDisplayMode;

enum ExcelXlFormatConditionType {
	ExcelXlFormatConditionTypeCellValue = '\002S\000\001',
	ExcelXlFormatConditionTypeExpression = '\002S\000\002',
	ExcelXlFormatConditionTypeColorScale = '\002S\000\003',
	ExcelXlFormatConditionTypeDatabar = '\002S\000\004',
	ExcelXlFormatConditionTypeTop10 = '\002S\000\005',
	ExcelXlFormatConditionTypeIconSets = '\002S\000\006',
	ExcelXlFormatConditionTypeUniqueValues = '\002S\000\007',
	ExcelXlFormatConditionTypeTextString = '\002S\000\011',
	ExcelXlFormatConditionTypeBlanksCondition = '\002S\000\012',
	ExcelXlFormatConditionTypeTimePeriod = '\002S\000\013',
	ExcelXlFormatConditionTypeAboveAverageCondition = '\002S\000\014',
	ExcelXlFormatConditionTypeNoBlanksCondition = '\002S\000\015',
	ExcelXlFormatConditionTypeErrorsCondition = '\002S\000\020',
	ExcelXlFormatConditionTypeNoErrorsCondition = '\002S\000\021'
};
typedef enum ExcelXlFormatConditionType ExcelXlFormatConditionType;

enum ExcelXlFormatConditionOperator {
	ExcelXlFormatConditionOperatorOperatorBetween = '\002T\000\001',
	ExcelXlFormatConditionOperatorOperatorNotBetween = '\002T\000\002',
	ExcelXlFormatConditionOperatorOperatorEqual = '\002T\000\003',
	ExcelXlFormatConditionOperatorOperatorNotEqual = '\002T\000\004',
	ExcelXlFormatConditionOperatorOperatorGreater = '\002T\000\005',
	ExcelXlFormatConditionOperatorOperatorLess = '\002T\000\006',
	ExcelXlFormatConditionOperatorOperatorGreaterEqual = '\002T\000\007',
	ExcelXlFormatConditionOperatorOperatorLessEqual = '\002T\000\010'
};
typedef enum ExcelXlFormatConditionOperator ExcelXlFormatConditionOperator;

enum ExcelXlEnableSelection {
	ExcelXlEnableSelectionNoRestrictions = '\002U\000\000',
	ExcelXlEnableSelectionUnlockedCells = '\002U\000\001',
	ExcelXlEnableSelectionNoSelection = '\002T\357\322'
};
typedef enum ExcelXlEnableSelection ExcelXlEnableSelection;

enum ExcelXlDVType {
	ExcelXlDVTypeValidateInputOnly = '\002V\000\000',
	ExcelXlDVTypeValidateWholeNumber = '\002V\000\001',
	ExcelXlDVTypeValidateDecimal = '\002V\000\002',
	ExcelXlDVTypeValidateList = '\002V\000\003',
	ExcelXlDVTypeValidatedDate = '\002V\000\004',
	ExcelXlDVTypeValidateTime = '\002V\000\005',
	ExcelXlDVTypeValidateTextLength = '\002V\000\006',
	ExcelXlDVTypeValidateCustom = '\002V\000\007'
};
typedef enum ExcelXlDVType ExcelXlDVType;

enum ExcelXlIMEMode {
	ExcelXlIMEModeIMEModeNoControl = '\002W\000\000',
	ExcelXlIMEModeIMEModeOn = '\002W\000\001',
	ExcelXlIMEModeIMEModeOff = '\002W\000\002',
	ExcelXlIMEModeIMEModeDisable = '\002W\000\003',
	ExcelXlIMEModeIMEModeHiragana = '\002W\000\004',
	ExcelXlIMEModeIMEModeKatakana = '\002W\000\005',
	ExcelXlIMEModeIMEModeKatakanaHalf = '\002W\000\006',
	ExcelXlIMEModeIMEModeAlphaFull = '\002W\000\007',
	ExcelXlIMEModeIMEModeAlpha = '\002W\000\010',
	ExcelXlIMEModeIMEModeHangulFull = '\002W\000\011',
	ExcelXlIMEModeIMEModeHangul = '\002W\000\012'
};
typedef enum ExcelXlIMEMode ExcelXlIMEMode;

enum ExcelXlDVAlertStyle {
	ExcelXlDVAlertStyleValidAlertNone = '\002W\377\377',
	ExcelXlDVAlertStyleValidAlertStop = '\002X\000\001',
	ExcelXlDVAlertStyleValidAlertWarning = '\002X\000\002',
	ExcelXlDVAlertStyleValidAlertInformation = '\002X\000\003'
};
typedef enum ExcelXlDVAlertStyle ExcelXlDVAlertStyle;

enum ExcelXlChartLocation {
	ExcelXlChartLocationLocationAsNewSheet = '\002Y\000\001',
	ExcelXlChartLocationLocationAsObject = '\002Y\000\002',
	ExcelXlChartLocationLocationAutomatic = '\002Y\000\003'
};
typedef enum ExcelXlChartLocation ExcelXlChartLocation;

enum ExcelXlChartElementPosition {
	ExcelXlChartElementPositionAutomatic = '\002Y\357\367',
	ExcelXlChartElementPositionCustom = '\002Y\357\356'
};
typedef enum ExcelXlChartElementPosition ExcelXlChartElementPosition;

enum ExcelXlPivotTableVersionList {
	ExcelXlPivotTableVersionListPivotTableVersion2000 = '\003\204\000\000',
	ExcelXlPivotTableVersionListPivotTableVersion10 = '\003\204\000\001',
	ExcelXlPivotTableVersionListPivotTableVersion11 = '\003\204\000\002',
	ExcelXlPivotTableVersionListPivotTableVersion12 = '\003\204\000\003',
	ExcelXlPivotTableVersionListPivotTableVersion14 = '\003\204\000\004',
	ExcelXlPivotTableVersionListPivotTableVersionCurrent = '\003\203\377\377'
};
typedef enum ExcelXlPivotTableVersionList ExcelXlPivotTableVersionList;

enum ExcelXlLayoutRowType {
	ExcelXlLayoutRowTypeCompactRow = '\003\205\000\000',
	ExcelXlLayoutRowTypeTabularRow = '\003\205\000\001',
	ExcelXlLayoutRowTypeOutlineRow = '\003\205\000\002'
};
typedef enum ExcelXlLayoutRowType ExcelXlLayoutRowType;

enum ExcelXlSubtototalLocationType {
	ExcelXlSubtototalLocationTypeAtTop = '\003\206\000\001',
	ExcelXlSubtototalLocationTypeAtBottom = '\003\206\000\002'
};
typedef enum ExcelXlSubtototalLocationType ExcelXlSubtototalLocationType;

enum ExcelXlAllocation {
	ExcelXlAllocationManualAllocation = '\003\207\000\001',
	ExcelXlAllocationAutomaticAllocation = '\003\207\000\002'
};
typedef enum ExcelXlAllocation ExcelXlAllocation;

enum ExcelXlAllocationValue {
	ExcelXlAllocationValueAllocateValue = '\003\210\000\001',
	ExcelXlAllocationValueAllocateIncrement = '\003\210\000\002'
};
typedef enum ExcelXlAllocationValue ExcelXlAllocationValue;

enum ExcelXlAllocationMethod {
	ExcelXlAllocationMethodEqualAllocation = '\003\211\000\001',
	ExcelXlAllocationMethodWeightAllocation = '\003\211\000\002'
};
typedef enum ExcelXlAllocationMethod ExcelXlAllocationMethod;

enum ExcelXlPivotFieldRepeatLabels {
	ExcelXlPivotFieldRepeatLabelsDoNotRepeatLabels = '\003\212\000\001',
	ExcelXlPivotFieldRepeatLabelsRepeatLabels = '\003\212\000\002'
};
typedef enum ExcelXlPivotFieldRepeatLabels ExcelXlPivotFieldRepeatLabels;

enum ExcelXlPivotTableMissingItems {
	ExcelXlPivotTableMissingItemsMissingItemsDefault = '\003\212\377\377',
	ExcelXlPivotTableMissingItemsMissingItemsNone = '\003\213\000\000',
	ExcelXlPivotTableMissingItemsMissingItemsMax = '\003\213~\364',
	ExcelXlPivotTableMissingItemsMissingItemsMax2 = '\003\233\000\000'
};
typedef enum ExcelXlPivotTableMissingItems ExcelXlPivotTableMissingItems;

enum ExcelXlPivotCellType {
	ExcelXlPivotCellTypePivotCellValue = '\003\214\000\000',
	ExcelXlPivotCellTypePivotCellPivotItem = '\003\214\000\001',
	ExcelXlPivotCellTypePivotCellSubtotal = '\003\214\000\002',
	ExcelXlPivotCellTypePivotCellGrandTotal = '\003\214\000\003',
	ExcelXlPivotCellTypePivotCellDataField = '\003\214\000\004',
	ExcelXlPivotCellTypePivotCellPivotField = '\003\214\000\005',
	ExcelXlPivotCellTypePivotCellPageFieldItem = '\003\214\000\006',
	ExcelXlPivotCellTypePivotCellCustomSubtotal = '\003\214\000\007',
	ExcelXlPivotCellTypePivotCellDataPivotField = '\003\214\000\010',
	ExcelXlPivotCellTypePivotCellBlankCell = '\003\214\000\011'
};
typedef enum ExcelXlPivotCellType ExcelXlPivotCellType;

enum ExcelXlCellChangedState {
	ExcelXlCellChangedStateCellNotChanged = '\003\215\000\001',
	ExcelXlCellChangedStateCellChanged = '\003\215\000\002',
	ExcelXlCellChangedStateCellChangeApplied = '\003\215\000\003'
};
typedef enum ExcelXlCellChangedState ExcelXlCellChangedState;

enum ExcelXlLayoutFormType {
	ExcelXlLayoutFormTypeTabular = '\003\216\000\000',
	ExcelXlLayoutFormTypeOutline = '\003\216\000\001'
};
typedef enum ExcelXlLayoutFormType ExcelXlLayoutFormType;

enum ExcelXlPivotFilterType {
	ExcelXlPivotFilterTypePivotTopCount = '\003\217\000\001',
	ExcelXlPivotFilterTypePivotBottomCount = '\003\217\000\002',
	ExcelXlPivotFilterTypePivotTopPercent = '\003\217\000\003',
	ExcelXlPivotFilterTypePivotBottomPercent = '\003\217\000\004',
	ExcelXlPivotFilterTypePivotTopSum = '\003\217\000\005',
	ExcelXlPivotFilterTypePivotBottomSum = '\003\217\000\006',
	ExcelXlPivotFilterTypePivotValueEquals = '\003\217\000\007',
	ExcelXlPivotFilterTypePivotValueIsNotEqual = '\003\217\000\010',
	ExcelXlPivotFilterTypePivotValueIsGreaterThan = '\003\217\000\011',
	ExcelXlPivotFilterTypePivotValueIsGreaterThanOrEqualTo = '\003\217\000\012',
	ExcelXlPivotFilterTypePivotValueIsLessThan = '\003\217\000\013',
	ExcelXlPivotFilterTypePivotValueIsLessThanOrEqualTo = '\003\217\000\014',
	ExcelXlPivotFilterTypePivotValueIsBetween = '\003\217\000\015',
	ExcelXlPivotFilterTypePivotValueIsNotBetween = '\003\217\000\016',
	ExcelXlPivotFilterTypePivotCaptionEquals = '\003\217\000\017',
	ExcelXlPivotFilterTypePivotCaptionDoesNotEqual = '\003\217\000\020',
	ExcelXlPivotFilterTypePivotCaptionBeginsWith = '\003\217\000\021',
	ExcelXlPivotFilterTypePivotCaptionDoesNotBeginWith = '\003\217\000\022',
	ExcelXlPivotFilterTypePivotCaptionEndsWith = '\003\217\000\023',
	ExcelXlPivotFilterTypePivotCaptionDoesNotEndWith = '\003\217\000\024',
	ExcelXlPivotFilterTypePivotCaptionContains = '\003\217\000\025',
	ExcelXlPivotFilterTypePivotCaptionDoesNotContain = '\003\217\000\026',
	ExcelXlPivotFilterTypePivotCaptionIsGreaterThan = '\003\217\000\027',
	ExcelXlPivotFilterTypePivotCaptionIsGreaterThanOrEqualTo = '\003\217\000\030',
	ExcelXlPivotFilterTypePivotCaptionIsLessThan = '\003\217\000\031',
	ExcelXlPivotFilterTypePivotCaptionIsLessThanOrEqualTo = '\003\217\000\032',
	ExcelXlPivotFilterTypePivotCaptionIsBetween = '\003\217\000\033',
	ExcelXlPivotFilterTypePivotCaptionIsNowBetween = '\003\217\000\034',
	ExcelXlPivotFilterTypePivotSpecificDate = '\003\217\000\035',
	ExcelXlPivotFilterTypePivotNotSpecificDate = '\003\217\000\036',
	ExcelXlPivotFilterTypePivotBefore = '\003\217\000\037',
	ExcelXlPivotFilterTypePivotBeforeOrEqualTo = '\003\217\000 ',
	ExcelXlPivotFilterTypePivotAfter = '\003\217\000!',
	ExcelXlPivotFilterTypePivotAfterOrEqualTo = '\003\217\000\"',
	ExcelXlPivotFilterTypePivotBetween = '\003\217\000#',
	ExcelXlPivotFilterTypePivotNotBetween = '\003\217\000$',
	ExcelXlPivotFilterTypePivotTomorrow = '\003\217\000%',
	ExcelXlPivotFilterTypePivotToday = '\003\217\000&',
	ExcelXlPivotFilterTypePivotYesterday = '\003\217\000\'',
	ExcelXlPivotFilterTypePivotNextWeek = '\003\217\000(',
	ExcelXlPivotFilterTypePivotThisWeek = '\003\217\000)',
	ExcelXlPivotFilterTypePivotLastWeek = '\003\217\000*',
	ExcelXlPivotFilterTypePivotNextMonth = '\003\217\000+',
	ExcelXlPivotFilterTypePivotThisMonth = '\003\217\000,',
	ExcelXlPivotFilterTypePivotLastMonth = '\003\217\000-',
	ExcelXlPivotFilterTypePivotNextQuarter = '\003\217\000.',
	ExcelXlPivotFilterTypePivotThisQuarter = '\003\217\000/',
	ExcelXlPivotFilterTypePivotLastQuarter = '\003\217\0000',
	ExcelXlPivotFilterTypePivotNextYear = '\003\217\0001',
	ExcelXlPivotFilterTypePivotThisYear = '\003\217\0002',
	ExcelXlPivotFilterTypePivotLastYear = '\003\217\0003',
	ExcelXlPivotFilterTypePivotYearToDate = '\003\217\0004',
	ExcelXlPivotFilterTypePivotAllDatesInPeriodQuarter1 = '\003\217\0005',
	ExcelXlPivotFilterTypePivotAllDatesInPeriodQuarter2 = '\003\217\0006',
	ExcelXlPivotFilterTypePivotAllDatesInPeriodQuarter3 = '\003\217\0007',
	ExcelXlPivotFilterTypePivotAllDatesInPeriodQuarter4 = '\003\217\0008',
	ExcelXlPivotFilterTypePivotAllDatesInPeriodJanuary = '\003\217\0009',
	ExcelXlPivotFilterTypePivotAllDatesInPeriodFeberary = '\003\217\000:',
	ExcelXlPivotFilterTypePivotAllDatesInPeriodMarch = '\003\217\000;',
	ExcelXlPivotFilterTypePivotAllDatesInPeriodApril = '\003\217\000<',
	ExcelXlPivotFilterTypePivotAllDatesInPeriodMay = '\003\217\000=',
	ExcelXlPivotFilterTypePivotAllDatesInPeriodJune = '\003\217\000>',
	ExcelXlPivotFilterTypePivotAllDatesInPeriodJuly = '\003\217\000\?',
	ExcelXlPivotFilterTypePivotAllDatesInPeriodAugust = '\003\217\000@',
	ExcelXlPivotFilterTypePivotAllDatesInPeriodSeptember = '\003\217\000A',
	ExcelXlPivotFilterTypePivotAllDatesInPeriodOctober = '\003\217\000B',
	ExcelXlPivotFilterTypePivotAllDatesInPeriodNovember = '\003\217\000C',
	ExcelXlPivotFilterTypePivotAllDatesInPeriodDecember = '\003\217\000D'
};
typedef enum ExcelXlPivotFilterType ExcelXlPivotFilterType;

enum ExcelXlPivotLineType {
	ExcelXlPivotLineTypePivotLineRegular = '\003\220\000\000',
	ExcelXlPivotLineTypePivotLineSubtotal = '\003\220\000\001',
	ExcelXlPivotLineTypePivotLineGrandtotal = '\003\220\000\002',
	ExcelXlPivotLineTypePivotLineBlank = '\003\220\000\003'
};
typedef enum ExcelXlPivotLineType ExcelXlPivotLineType;

enum ExcelXlCubeFieldType {
	ExcelXlCubeFieldTypeHierarchy = '\003\221\000\001',
	ExcelXlCubeFieldTypeMeasure = '\003\221\000\002',
	ExcelXlCubeFieldTypeSet = '\003\221\000\003'
};
typedef enum ExcelXlCubeFieldType ExcelXlCubeFieldType;

enum ExcelXlCubeFieldSubType {
	ExcelXlCubeFieldSubTypeCubeHierarchy = '\003\222\000\001',
	ExcelXlCubeFieldSubTypeCubeMeasure = '\003\222\000\002',
	ExcelXlCubeFieldSubTypeCubeSet = '\003\222\000\003',
	ExcelXlCubeFieldSubTypeCubeAttribute = '\003\222\000\004',
	ExcelXlCubeFieldSubTypeCubeCalculatedMeasure = '\003\222\000\005',
	ExcelXlCubeFieldSubTypeCubeKPIValue = '\003\222\000\006',
	ExcelXlCubeFieldSubTypeCubeKPIGoal = '\003\222\000\007',
	ExcelXlCubeFieldSubTypeCubeKPIStatus = '\003\222\000\010',
	ExcelXlCubeFieldSubTypeCubeKPITrend = '\003\222\000\011',
	ExcelXlCubeFieldSubTypeCubeKPIWeight = '\003\222\000\012'
};
typedef enum ExcelXlCubeFieldSubType ExcelXlCubeFieldSubType;

enum ExcelXlPropertyDisplayedIn {
	ExcelXlPropertyDisplayedInDisplayPropertyInPivotTable = '\003\223\000\001',
	ExcelXlPropertyDisplayedInDisplayPropertyInTooltip = '\003\223\000\002',
	ExcelXlPropertyDisplayedInDisplayPropertyInPivotTableAndTooltip = '\003\223\000\003'
};
typedef enum ExcelXlPropertyDisplayedIn ExcelXlPropertyDisplayedIn;

enum ExcelXlCalculatedMemberType {
	ExcelXlCalculatedMemberTypeCalculatedMember = '\003\224\000\000',
	ExcelXlCalculatedMemberTypeCalculatedSet = '\003\224\000\001'
};
typedef enum ExcelXlCalculatedMemberType ExcelXlCalculatedMemberType;

enum ExcelXlConnectionType {
	ExcelXlConnectionTypeConnectionTypeOLEDB = '\003\225\000\001',
	ExcelXlConnectionTypeConnectionTypeODBC = '\003\225\000\002',
	ExcelXlConnectionTypeConnectionTypeXMLMAP = '\003\225\000\003',
	ExcelXlConnectionTypeConnectionTypeTEXT = '\003\225\000\004',
	ExcelXlConnectionTypeConnectionTypeWEB = '\003\225\000\005'
};
typedef enum ExcelXlConnectionType ExcelXlConnectionType;

enum ExcelXlPasteSpecialOperation {
	ExcelXlPasteSpecialOperationPasteSpecialOperationAdd = '\002[\000\002',
	ExcelXlPasteSpecialOperationPasteSpecialOperationDivide = '\002[\000\005',
	ExcelXlPasteSpecialOperationPasteSpecialOperationMultiply = '\002[\000\004',
	ExcelXlPasteSpecialOperationPasteSpecialOperationNone = '\002Z\357\322',
	ExcelXlPasteSpecialOperationPasteSpecialOperationSubtract = '\002[\000\003'
};
typedef enum ExcelXlPasteSpecialOperation ExcelXlPasteSpecialOperation;

enum ExcelXlPasteType {
	ExcelXlPasteTypePasteAll = '\002[\357\370',
	ExcelXlPasteTypePasteAllUsingSourceTheme = '\002\\\000\015',
	ExcelXlPasteTypePasteAllExceptBorders = '\002\\\000\007',
	ExcelXlPasteTypePasteFormats = '\002[\357\346',
	ExcelXlPasteTypePasteFormulas = '\002[\357\345',
	ExcelXlPasteTypePasteComments = '\002[\357\320',
	ExcelXlPasteTypePasteValues = '\002[\357\275',
	ExcelXlPasteTypePasteColumnWidths = '\002\\\000\010',
	ExcelXlPasteTypePasteValidation = '\002\\\000\006',
	ExcelXlPasteTypePasteFormulasAndNumberFormats = '\002\\\000\013',
	ExcelXlPasteTypePasteValuesAndNumberFormats = '\002\\\000\014'
};
typedef enum ExcelXlPasteType ExcelXlPasteType;

enum ExcelXlPhoneticCharacterType {
	ExcelXlPhoneticCharacterTypePhoneticCharacterHalfWidthKatakana = '\002]\000\000',
	ExcelXlPhoneticCharacterTypePhoneticCharacterFullWidthKatakana = '\002]\000\001',
	ExcelXlPhoneticCharacterTypePhoneticCharacterHiragana = '\002]\000\002',
	ExcelXlPhoneticCharacterTypeNoPhoneticCharacterConversion = '\002]\000\003'
};
typedef enum ExcelXlPhoneticCharacterType ExcelXlPhoneticCharacterType;

enum ExcelXlPhoneticAlignment {
	ExcelXlPhoneticAlignmentPhoneticAlignNoControl = '\002^\000\000',
	ExcelXlPhoneticAlignmentPhoneticAlignLeft = '\002^\000\001',
	ExcelXlPhoneticAlignmentPhoneticAlignCenter = '\002^\000\002',
	ExcelXlPhoneticAlignmentPhoneticAlignDistributed = '\002^\000\003'
};
typedef enum ExcelXlPhoneticAlignment ExcelXlPhoneticAlignment;

enum ExcelXlPictureAppearance {
	ExcelXlPictureAppearancePrinter = '\002_\000\002',
	ExcelXlPictureAppearanceScreen = '\002_\000\001'
};
typedef enum ExcelXlPictureAppearance ExcelXlPictureAppearance;

enum ExcelXlPivotFieldOrientation {
	ExcelXlPivotFieldOrientationOrientAsColumnField = '\002`\000\002',
	ExcelXlPivotFieldOrientationOrientAsDataField = '\002`\000\004',
	ExcelXlPivotFieldOrientationOrientAsHidden = '\002`\000\000',
	ExcelXlPivotFieldOrientationOrientAsPageField = '\002`\000\003',
	ExcelXlPivotFieldOrientationOrientAsRowField = '\002`\000\001'
};
typedef enum ExcelXlPivotFieldOrientation ExcelXlPivotFieldOrientation;

enum ExcelXlPivotFieldCalculation {
	ExcelXlPivotFieldCalculationPivotFieldCalculationDifferenceFrom = '\002a\000\002',
	ExcelXlPivotFieldCalculationPivotFieldCalculationIndex = '\002a\000\011',
	ExcelXlPivotFieldCalculationPivotFieldCalculationNoAdditionalCalculation = '\002`\357\321',
	ExcelXlPivotFieldCalculationPivotFieldCalculationPercentDifferenceFrom = '\002a\000\004',
	ExcelXlPivotFieldCalculationPivotFieldCalculationPercentOf = '\002a\000\003',
	ExcelXlPivotFieldCalculationPivotFieldCalculationPercentOfColumn = '\002a\000\007',
	ExcelXlPivotFieldCalculationPivotFieldCalculationPercentOfRow = '\002a\000\006',
	ExcelXlPivotFieldCalculationPivotFieldCalculationPercentOfTotal = '\002a\000\010',
	ExcelXlPivotFieldCalculationPivotFieldCalculationRunningTotal = '\002a\000\005'
};
typedef enum ExcelXlPivotFieldCalculation ExcelXlPivotFieldCalculation;

enum ExcelXlPlacement {
	ExcelXlPlacementPlacementFreeFloating = '\002b\000\003',
	ExcelXlPlacementPlacementMove = '\002b\000\002',
	ExcelXlPlacementPlacementMoveAndSize = '\002b\000\001'
};
typedef enum ExcelXlPlacement ExcelXlPlacement;

enum ExcelXlPlatform {
	ExcelXlPlatformMacintosh = '\002c\000\001',
	ExcelXlPlatformMSDos = '\002c\000\003',
	ExcelXlPlatformMSWindows = '\002c\000\002'
};
typedef enum ExcelXlPlatform ExcelXlPlatform;

enum ExcelXlPrintLocation {
	ExcelXlPrintLocationPrintSheetEnd = '\002d\000\001',
	ExcelXlPrintLocationPrintInPlace = '\002d\000\020',
	ExcelXlPrintLocationPrintNoComments = '\002c\357\322'
};
typedef enum ExcelXlPrintLocation ExcelXlPrintLocation;

enum ExcelXlPriority {
	ExcelXlPriorityPriorityHigh = '\002d\357\341',
	ExcelXlPriorityPriorityLow = '\002d\357\332',
	ExcelXlPriorityPriorityNormal = '\002d\357\321'
};
typedef enum ExcelXlPriority ExcelXlPriority;

enum ExcelXlPTSelectionMode {
	ExcelXlPTSelectionModeSelectionModeLabelOnly = '\002f\000\001',
	ExcelXlPTSelectionModeSelectionModeDataAndLabel = '\002f\000\000',
	ExcelXlPTSelectionModeSelectionModeDataOnly = '\002f\000\002',
	ExcelXlPTSelectionModeSelectionModeOrigin = '\002f\000\003',
	ExcelXlPTSelectionModeSelectionModeButton = '\002f\000\017',
	ExcelXlPTSelectionModeSelectionModeBlanks = '\002f\000\004'
};
typedef enum ExcelXlPTSelectionMode ExcelXlPTSelectionMode;

enum ExcelXlRangeAutoFormat {
	ExcelXlRangeAutoFormatRangeAutoformatThreeDEffects1 = '\002g\000\015',
	ExcelXlRangeAutoFormatRangeAutoformatThreeDEffects2 = '\002g\000\016',
	ExcelXlRangeAutoFormatRangeAutoformatAccounting1 = '\002g\000\004',
	ExcelXlRangeAutoFormatRangeAutoformatAccounting2 = '\002g\000\005',
	ExcelXlRangeAutoFormatRangeAutoformatAccounting3 = '\002g\000\006',
	ExcelXlRangeAutoFormatRangeAutoformatAccounting4 = '\002g\000\021',
	ExcelXlRangeAutoFormatRangeAutoformatClassic1 = '\002g\000\001',
	ExcelXlRangeAutoFormatRangeAutoformatClassic2 = '\002g\000\002',
	ExcelXlRangeAutoFormatRangeAutoformatClassic3 = '\002g\000\003',
	ExcelXlRangeAutoFormatRangeAutoformatColor1 = '\002g\000\007',
	ExcelXlRangeAutoFormatRangeAutoformatColor2 = '\002g\000\010',
	ExcelXlRangeAutoFormatRangeAutoformatColor3 = '\002g\000\011',
	ExcelXlRangeAutoFormatRangeAutoformatList1 = '\002g\000\012',
	ExcelXlRangeAutoFormatRangeAutoformatList2 = '\002g\000\013',
	ExcelXlRangeAutoFormatRangeAutoformatList3 = '\002g\000\014',
	ExcelXlRangeAutoFormatRangeAutoformatLocalFormat1 = '\002g\000\017',
	ExcelXlRangeAutoFormatRangeAutoformatLocalFormat2 = '\002g\000\020',
	ExcelXlRangeAutoFormatRangeAutoformatLocalFormat3 = '\002g\000\023',
	ExcelXlRangeAutoFormatRangeAutoformatLocalFormat4 = '\002g\000\024',
	ExcelXlRangeAutoFormatRangeAutoformatNone = '\002f\357\322',
	ExcelXlRangeAutoFormatRangeAutoformatSimple = '\002f\357\306'
};
typedef enum ExcelXlRangeAutoFormat ExcelXlRangeAutoFormat;

enum ExcelXlRoutingSlipDelivery {
	ExcelXlRoutingSlipDeliveryAllAtOnce = '\002i\000\002',
	ExcelXlRoutingSlipDeliveryOneAfterAnother = '\002i\000\001'
};
typedef enum ExcelXlRoutingSlipDelivery ExcelXlRoutingSlipDelivery;

enum ExcelXlRoutingSlipStatus {
	ExcelXlRoutingSlipStatusNotYetRouted = '\002j\000\000',
	ExcelXlRoutingSlipStatusRoutingComplete = '\002j\000\002',
	ExcelXlRoutingSlipStatusRoutingInProgress = '\002j\000\001'
};
typedef enum ExcelXlRoutingSlipStatus ExcelXlRoutingSlipStatus;

enum ExcelXlRunAutoMacro {
	ExcelXlRunAutoMacroAutoActivate = '\002k\000\003',
	ExcelXlRunAutoMacroAutoClose = '\002k\000\002',
	ExcelXlRunAutoMacroAutoDeactivate = '\002k\000\004',
	ExcelXlRunAutoMacroAutoOpen = '\002k\000\001'
};
typedef enum ExcelXlRunAutoMacro ExcelXlRunAutoMacro;

enum ExcelXlSaveAsAccessMode {
	ExcelXlSaveAsAccessModeExclusive = '\002m\000\003',
	ExcelXlSaveAsAccessModeNoChange = '\002m\000\001',
	ExcelXlSaveAsAccessModeShared = '\002m\000\002'
};
typedef enum ExcelXlSaveAsAccessMode ExcelXlSaveAsAccessMode;

enum ExcelXlSaveConflictResolution {
	ExcelXlSaveConflictResolutionLocalSessionChanges = '\002n\000\002',
	ExcelXlSaveConflictResolutionOtherSessionChanges = '\002n\000\003',
	ExcelXlSaveConflictResolutionUserResolution = '\002n\000\001'
};
typedef enum ExcelXlSaveConflictResolution ExcelXlSaveConflictResolution;

enum ExcelXlSearchDirection {
	ExcelXlSearchDirectionSearchNext = '\002o\000\001',
	ExcelXlSearchDirectionSearchPrevious = '\002o\000\002'
};
typedef enum ExcelXlSearchDirection ExcelXlSearchDirection;

enum ExcelXlSearchOrder {
	ExcelXlSearchOrderByColumns = '\002p\000\002',
	ExcelXlSearchOrderByRows = '\002p\000\001'
};
typedef enum ExcelXlSearchOrder ExcelXlSearchOrder;

enum ExcelXlSheetVisibility {
	ExcelXlSheetVisibilitySheetVisible = '\002p\377\377',
	ExcelXlSheetVisibilitySheetHidden = '\002q\000\000',
	ExcelXlSheetVisibilitySheetVeryHidden = '\002q\000\002'
};
typedef enum ExcelXlSheetVisibility ExcelXlSheetVisibility;

enum ExcelXlSortMethod {
	ExcelXlSortMethodPinYin = '\002r\000\001' /* Phonetic Chinese/Japanese sort order for characters. This is the default value. */,
	ExcelXlSortMethodStroke = '\002r\000\002' /* Sort by the quantity of strokes in each character. */
};
typedef enum ExcelXlSortMethod ExcelXlSortMethod;

enum ExcelXlSortOrder {
	ExcelXlSortOrderSortAscending = '\002t\000\001' /* Sorts the specified field in ascending order. This is the default value. */,
	ExcelXlSortOrderSortDescending = '\002t\000\002' /* Sorts the specified field in descending order. */,
	ExcelXlSortOrderSortManual = '\002s\357\331' /* It is not supported. */
};
typedef enum ExcelXlSortOrder ExcelXlSortOrder;

enum ExcelXlSortOrientation {
	ExcelXlSortOrientationSortRows = '\002u\000\002' /* Sorts by row. this is the default value. */,
	ExcelXlSortOrientationSortColumns = '\002u\000\001' /* Sorts by column. */
};
typedef enum ExcelXlSortOrientation ExcelXlSortOrientation;

enum ExcelXlSortType {
	ExcelXlSortTypeSortLabels = '\002v\000\002' /* Sorts the PivotTable report by labels. */,
	ExcelXlSortTypeSortValues = '\002v\000\001' /* Sorts the PivotTable report by values. */
};
typedef enum ExcelXlSortType ExcelXlSortType;

enum ExcelXlSpecialCellsValue {
	ExcelXlSpecialCellsValueErrors = '\002w\000\020',
	ExcelXlSpecialCellsValueLogical = '\002w\000\004',
	ExcelXlSpecialCellsValueNumbers = '\002w\000\001',
	ExcelXlSpecialCellsValueTextValues = '\002w\000\002'
};
typedef enum ExcelXlSpecialCellsValue ExcelXlSpecialCellsValue;

enum ExcelXlSummaryRow {
	ExcelXlSummaryRowSummaryAbove = '\002x\000\000',
	ExcelXlSummaryRowSummaryBelow = '\002x\000\001'
};
typedef enum ExcelXlSummaryRow ExcelXlSummaryRow;

enum ExcelXlSummaryColumn {
	ExcelXlSummaryColumnSummaryOnLeft = '\002x\357\335',
	ExcelXlSummaryColumnSummaryOnRight = '\002x\357\310'
};
typedef enum ExcelXlSummaryColumn ExcelXlSummaryColumn;

enum ExcelXlSummaryReportType {
	ExcelXlSummaryReportTypeSummaryPivotTable = '\002y\357\314',
	ExcelXlSummaryReportTypeStandardSummary = '\002z\000\001'
};
typedef enum ExcelXlSummaryReportType ExcelXlSummaryReportType;

enum ExcelXlTextParsingType {
	ExcelXlTextParsingTypeDelimited = '\002|\000\001',
	ExcelXlTextParsingTypeFixedWidth = '\002|\000\002'
};
typedef enum ExcelXlTextParsingType ExcelXlTextParsingType;

enum ExcelXlTextQualifier {
	ExcelXlTextQualifierTextQualifierDoubleQuote = '\002}\000\001',
	ExcelXlTextQualifierTextQualifierNone = '\002|\357\322',
	ExcelXlTextQualifierTextQualifierSingleQuote = '\002}\000\002'
};
typedef enum ExcelXlTextQualifier ExcelXlTextQualifier;

enum ExcelXlWBATemplate {
	ExcelXlWBATemplateChart = '\002}\357\363',
	ExcelXlWBATemplateExcel4IntlMacroSheet = '\002~\000\004',
	ExcelXlWBATemplateExcel4MacroSheet = '\002~\000\003',
	ExcelXlWBATemplateWorksheet = '\002}\357\271'
};
typedef enum ExcelXlWBATemplate ExcelXlWBATemplate;

enum ExcelXlWindowView {
	ExcelXlWindowViewNormalView = '\002\177\000\001',
	ExcelXlWindowViewPageLayoutView = '\002\177\000\003'
};
typedef enum ExcelXlWindowView ExcelXlWindowView;

enum ExcelXlXLMMacroType {
	ExcelXlXLMMacroTypeMacroTypeCommand = '\002\200\000\002',
	ExcelXlXLMMacroTypeMacroTypeFunction = '\002\200\000\001',
	ExcelXlXLMMacroTypeMacroTypeNotXLM = '\002\200\000\003'
};
typedef enum ExcelXlXLMMacroType ExcelXlXLMMacroType;

enum ExcelXlYesNoGuess {
	ExcelXlYesNoGuessHeaderGuess = '\002\201\000\000' /* Default value. Excel determines whether there is a header, and where it is, if there is one. */,
	ExcelXlYesNoGuessHeaderNo = '\002\201\000\002' /* The entire range should be sorted. */,
	ExcelXlYesNoGuessHeaderYes = '\002\201\000\001' /* The entire range should not be sorted. */
};
typedef enum ExcelXlYesNoGuess ExcelXlYesNoGuess;

enum ExcelXlDisplayDrawingObjects {
	ExcelXlDisplayDrawingObjectsDisplayShapes = '\002\201\357\370',
	ExcelXlDisplayDrawingObjectsHide = '\002\202\000\003',
	ExcelXlDisplayDrawingObjectsPlaceholders = '\002\202\000\002'
};
typedef enum ExcelXlDisplayDrawingObjects ExcelXlDisplayDrawingObjects;

enum ExcelXlBordersIndex {
	ExcelXlBordersIndexInsideHorizontal = '\002\203\000\014',
	ExcelXlBordersIndexInsideVertical = '\002\203\000\013',
	ExcelXlBordersIndexDiagonalDown = '\002\203\000\005',
	ExcelXlBordersIndexDiagonalUp = '\002\203\000\006',
	ExcelXlBordersIndexEdgeBottom = '\002\203\000\011',
	ExcelXlBordersIndexEdgeLeft = '\002\203\000\007',
	ExcelXlBordersIndexEdgeRight = '\002\203\000\012',
	ExcelXlBordersIndexEdgeTop = '\002\203\000\010',
	ExcelXlBordersIndexBorderBottom = '\002\202\357\365',
	ExcelXlBordersIndexBorderLeft = '\002\202\357\335',
	ExcelXlBordersIndexBorderRight = '\002\202\357\310',
	ExcelXlBordersIndexBorderTop = '\002\202\357\300'
};
typedef enum ExcelXlBordersIndex ExcelXlBordersIndex;

enum ExcelXlToolbarProtection {
	ExcelXlToolbarProtectionNoButtonChanges = '\002\204\000\001',
	ExcelXlToolbarProtectionNoChanges = '\002\204\000\004',
	ExcelXlToolbarProtectionNoDockingChanges = '\002\204\000\003',
	ExcelXlToolbarProtectionToolbarProtectionNone = '\002\203\357\321',
	ExcelXlToolbarProtectionNoShapeChanges = '\002\204\000\002'
};
typedef enum ExcelXlToolbarProtection ExcelXlToolbarProtection;

enum ExcelXlBuiltInDialog {
	ExcelXlBuiltInDialogDialogOpen = '\002\205\000\001',
	ExcelXlBuiltInDialogDialogOpenLinks = '\002\205\000\002',
	ExcelXlBuiltInDialogDialogSaveAs = '\002\205\000\005',
	ExcelXlBuiltInDialogDialogFileDelete = '\002\205\000\006',
	ExcelXlBuiltInDialogDialogPageSetup = '\002\205\000\007',
	ExcelXlBuiltInDialogDialogPrint = '\002\205\000\010',
	ExcelXlBuiltInDialogDialogPrinterSetup = '\002\205\000\011',
	ExcelXlBuiltInDialogDialogArrangeAll = '\002\205\000\014',
	ExcelXlBuiltInDialogDialogWindowSize = '\002\205\000\015',
	ExcelXlBuiltInDialogDialogWindowMove = '\002\205\000\016',
	ExcelXlBuiltInDialogDialogRun = '\002\205\000\021',
	ExcelXlBuiltInDialogDialogSetPrintTitles = '\002\205\000\027',
	ExcelXlBuiltInDialogDialogFont = '\002\205\000\032',
	ExcelXlBuiltInDialogDialogDisplay = '\002\205\000\033',
	ExcelXlBuiltInDialogDialogProtectDocument = '\002\205\000\034',
	ExcelXlBuiltInDialogDialogCalculation = '\002\205\000 ',
	ExcelXlBuiltInDialogDialogExtract = '\002\205\000#',
	ExcelXlBuiltInDialogDialogDataDelete = '\002\205\000$',
	ExcelXlBuiltInDialogDialogSort = '\002\205\000\'',
	ExcelXlBuiltInDialogDialogDataSeries = '\002\205\000(',
	ExcelXlBuiltInDialogDialogTable = '\002\205\000)',
	ExcelXlBuiltInDialogDialogFormatNumber = '\002\205\000*',
	ExcelXlBuiltInDialogDialogAlignment = '\002\205\000+',
	ExcelXlBuiltInDialogDialogStyle = '\002\205\000,',
	ExcelXlBuiltInDialogDialogBorder = '\002\205\000-',
	ExcelXlBuiltInDialogDialogCellProtection = '\002\205\000.',
	ExcelXlBuiltInDialogDialogColumnWidth = '\002\205\000/',
	ExcelXlBuiltInDialogDialogClear = '\002\205\0004',
	ExcelXlBuiltInDialogDialogPasteSpecial = '\002\205\0005',
	ExcelXlBuiltInDialogDialogEditDelete = '\002\205\0006',
	ExcelXlBuiltInDialogDialogInsert = '\002\205\0007',
	ExcelXlBuiltInDialogDialogPasteNames = '\002\205\000:',
	ExcelXlBuiltInDialogDialogDefineName = '\002\205\000=',
	ExcelXlBuiltInDialogDialogCreateNames = '\002\205\000>',
	ExcelXlBuiltInDialogDialogFormulaGoto = '\002\205\000\?',
	ExcelXlBuiltInDialogDialogFormulaFind = '\002\205\000@',
	ExcelXlBuiltInDialogDialogGalleryArea = '\002\205\000C',
	ExcelXlBuiltInDialogDialogGalleryBar = '\002\205\000D',
	ExcelXlBuiltInDialogDialogGalleryColumn = '\002\205\000E',
	ExcelXlBuiltInDialogDialogGalleryLine = '\002\205\000F',
	ExcelXlBuiltInDialogDialogGalleryPie = '\002\205\000G',
	ExcelXlBuiltInDialogDialogGalleryScatter = '\002\205\000H',
	ExcelXlBuiltInDialogDialogCombination = '\002\205\000I',
	ExcelXlBuiltInDialogDialogGridlines = '\002\205\000L',
	ExcelXlBuiltInDialogDialogAxes = '\002\205\000N',
	ExcelXlBuiltInDialogDialogAttachText = '\002\205\000P',
	ExcelXlBuiltInDialogDialogPatterns = '\002\205\000T',
	ExcelXlBuiltInDialogDialogMainChart = '\002\205\000U',
	ExcelXlBuiltInDialogDialogOverlay = '\002\205\000V',
	ExcelXlBuiltInDialogDialogScale = '\002\205\000W',
	ExcelXlBuiltInDialogDialogFormatLegend = '\002\205\000X',
	ExcelXlBuiltInDialogDialogFormatText = '\002\205\000Y',
	ExcelXlBuiltInDialogDialogParse = '\002\205\000[',
	ExcelXlBuiltInDialogDialogUnhide = '\002\205\000^',
	ExcelXlBuiltInDialogDialogWorkspace = '\002\205\000_',
	ExcelXlBuiltInDialogDialogActivate = '\002\205\000g',
	ExcelXlBuiltInDialogDialogCopyPicture = '\002\205\000l',
	ExcelXlBuiltInDialogDialogDeleteName = '\002\205\000n',
	ExcelXlBuiltInDialogDialogDeleteFormat = '\002\205\000o',
	ExcelXlBuiltInDialogDialogNew = '\002\205\000w',
	ExcelXlBuiltInDialogDialogRowHeight = '\002\205\000\177',
	ExcelXlBuiltInDialogDialogFormatMove = '\002\205\000\200',
	ExcelXlBuiltInDialogDialogFormatSize = '\002\205\000\201',
	ExcelXlBuiltInDialogDialogFormulaReplace = '\002\205\000\202',
	ExcelXlBuiltInDialogDialogSelectSpecial = '\002\205\000\204',
	ExcelXlBuiltInDialogDialogApplyNames = '\002\205\000\205',
	ExcelXlBuiltInDialogDialogReplaceFont = '\002\205\000\206',
	ExcelXlBuiltInDialogDialogSplit = '\002\205\000\211',
	ExcelXlBuiltInDialogDialogOutline = '\002\205\000\216',
	ExcelXlBuiltInDialogDialogSaveWorkbook = '\002\205\000\221',
	ExcelXlBuiltInDialogDialogCopyChart = '\002\205\000\223',
	ExcelXlBuiltInDialogDialogFormatFont = '\002\205\000\226',
	ExcelXlBuiltInDialogDialogNote = '\002\205\000\232',
	ExcelXlBuiltInDialogDialogSetUpdateStatus = '\002\205\000\237',
	ExcelXlBuiltInDialogDialogColorPalette = '\002\205\000\241',
	ExcelXlBuiltInDialogDialogChangeLink = '\002\205\000\246',
	ExcelXlBuiltInDialogDialogAppMove = '\002\205\000\252',
	ExcelXlBuiltInDialogDialogAppSize = '\002\205\000\253',
	ExcelXlBuiltInDialogDialogMainChartType = '\002\205\000\271',
	ExcelXlBuiltInDialogDialogOverlayChartType = '\002\205\000\272',
	ExcelXlBuiltInDialogDialogOpenMail = '\002\205\000\274',
	ExcelXlBuiltInDialogDialogSendMail = '\002\205\000\275',
	ExcelXlBuiltInDialogDialogStandardFont = '\002\205\000\276',
	ExcelXlBuiltInDialogDialogConsolidate = '\002\205\000\277',
	ExcelXlBuiltInDialogDialogSortSpecial = '\002\205\000\300',
	ExcelXlBuiltInDialogDialogGalleryThreeDArea = '\002\205\000\301',
	ExcelXlBuiltInDialogDialogGalleryThreeDColumn = '\002\205\000\302',
	ExcelXlBuiltInDialogDialogGalleryThreeDLine = '\002\205\000\303',
	ExcelXlBuiltInDialogDialogGalleryThreeDPie = '\002\205\000\304',
	ExcelXlBuiltInDialogDialogViewThreeD = '\002\205\000\305',
	ExcelXlBuiltInDialogDialogGoalSeek = '\002\205\000\306',
	ExcelXlBuiltInDialogDialogWorkgroup = '\002\205\000\307',
	ExcelXlBuiltInDialogDialogFillGroup = '\002\205\000\310',
	ExcelXlBuiltInDialogDialogUpdateLink = '\002\205\000\311',
	ExcelXlBuiltInDialogDialogPromote = '\002\205\000\312',
	ExcelXlBuiltInDialogDialogDemote = '\002\205\000\313',
	ExcelXlBuiltInDialogDialogShowDetail = '\002\205\000\314',
	ExcelXlBuiltInDialogDialogObjectProperties = '\002\205\000\317',
	ExcelXlBuiltInDialogDialogSaveNewObject = '\002\205\000\320',
	ExcelXlBuiltInDialogDialogApplyStyle = '\002\205\000\324',
	ExcelXlBuiltInDialogDialogAssignToObject = '\002\205\000\325',
	ExcelXlBuiltInDialogDialogObjectProtection = '\002\205\000\326',
	ExcelXlBuiltInDialogDialogShowToolbar = '\002\205\000\334',
	ExcelXlBuiltInDialogDialogPrintPreview = '\002\205\000\336',
	ExcelXlBuiltInDialogDialogEditColor = '\002\205\000\337',
	ExcelXlBuiltInDialogDialogFormatMain = '\002\205\000\341',
	ExcelXlBuiltInDialogDialogFormatOverlay = '\002\205\000\342',
	ExcelXlBuiltInDialogDialogEditSeries = '\002\205\000\344',
	ExcelXlBuiltInDialogDialogDefineStyle = '\002\205\000\345',
	ExcelXlBuiltInDialogDialogGalleryRadar = '\002\205\000\371',
	ExcelXlBuiltInDialogDialogZoom = '\002\205\001\000',
	ExcelXlBuiltInDialogDialogInsertObject = '\002\205\001\003',
	ExcelXlBuiltInDialogDialogSize = '\002\205\001\005',
	ExcelXlBuiltInDialogDialogMove = '\002\205\001\006',
	ExcelXlBuiltInDialogDialogFormatAuto = '\002\205\001\015',
	ExcelXlBuiltInDialogDialogGalleryThreeDBar = '\002\205\001\020',
	ExcelXlBuiltInDialogDialogGalleryThreeDSurface = '\002\205\001\021',
	ExcelXlBuiltInDialogDialogCustomizeToolbar = '\002\205\001\024',
	ExcelXlBuiltInDialogDialogWorkbookAdd = '\002\205\001\031',
	ExcelXlBuiltInDialogDialogWorkbookMove = '\002\205\001\032',
	ExcelXlBuiltInDialogDialogWorkbookCopy = '\002\205\001\033',
	ExcelXlBuiltInDialogDialogWorkbookOptions = '\002\205\001\034',
	ExcelXlBuiltInDialogDialogSaveWorkspace = '\002\205\001\035',
	ExcelXlBuiltInDialogDialogChartWizard = '\002\205\001 ',
	ExcelXlBuiltInDialogDialogAssignToTool = '\002\205\001%',
	ExcelXlBuiltInDialogDialogPlacement = '\002\205\001,',
	ExcelXlBuiltInDialogDialogFillWorkgroup = '\002\205\001-',
	ExcelXlBuiltInDialogDialogWorkbookNew = '\002\205\001.',
	ExcelXlBuiltInDialogDialogScenarioCells = '\002\205\0011',
	ExcelXlBuiltInDialogDialogScenarioAdd = '\002\205\0013',
	ExcelXlBuiltInDialogDialogScenarioEdit = '\002\205\0014',
	ExcelXlBuiltInDialogDialogScenarioSummary = '\002\205\0017',
	ExcelXlBuiltInDialogDialogPivotTableWizard = '\002\205\0018',
	ExcelXlBuiltInDialogDialogPivotFieldProperties = '\002\205\0019',
	ExcelXlBuiltInDialogDialogOptionsCalculation = '\002\205\001>',
	ExcelXlBuiltInDialogDialogOptionsEdit = '\002\205\001\?',
	ExcelXlBuiltInDialogDialogOptionsView = '\002\205\001@',
	ExcelXlBuiltInDialogDialogAddInManager = '\002\205\001A',
	ExcelXlBuiltInDialogDialogMenuEditor = '\002\205\001B',
	ExcelXlBuiltInDialogDialogAttachToolbars = '\002\205\001C',
	ExcelXlBuiltInDialogDialogOptionsChart = '\002\205\001E',
	ExcelXlBuiltInDialogDialogVbaInsertFile = '\002\205\001H',
	ExcelXlBuiltInDialogDialogVbaProcedureDefinition = '\002\205\001J',
	ExcelXlBuiltInDialogDialogRoutingSlip = '\002\205\001P',
	ExcelXlBuiltInDialogDialogMailLogon = '\002\205\001S',
	ExcelXlBuiltInDialogDialogInsertPicture = '\002\205\001V',
	ExcelXlBuiltInDialogDialogGalleryDoughnut = '\002\205\001X',
	ExcelXlBuiltInDialogDialogChartTrend = '\002\205\001^',
	ExcelXlBuiltInDialogDialogWorkbookInsert = '\002\205\001b',
	ExcelXlBuiltInDialogDialogOptionsTransition = '\002\205\001c',
	ExcelXlBuiltInDialogDialogOptionsGeneral = '\002\205\001d',
	ExcelXlBuiltInDialogDialogFilterAdvanced = '\002\205\001r',
	ExcelXlBuiltInDialogDialogMailNextLetter = '\002\205\001z',
	ExcelXlBuiltInDialogDialogDataLabel = '\002\205\001{',
	ExcelXlBuiltInDialogDialogInsertTitle = '\002\205\001|',
	ExcelXlBuiltInDialogDialogFontProperties = '\002\205\001}',
	ExcelXlBuiltInDialogDialogMacroOptions = '\002\205\001~',
	ExcelXlBuiltInDialogDialogWorkbookUnhide = '\002\205\001\200',
	ExcelXlBuiltInDialogDialogWorkbookName = '\002\205\001\202',
	ExcelXlBuiltInDialogDialogGalleryCustom = '\002\205\001\204',
	ExcelXlBuiltInDialogDialogAddChartAutoformat = '\002\205\001\206',
	ExcelXlBuiltInDialogDialogChartAddData = '\002\205\001\210',
	ExcelXlBuiltInDialogDialogTabOrder = '\002\205\001\212',
	ExcelXlBuiltInDialogDialogSubtotalCreate = '\002\205\001\216',
	ExcelXlBuiltInDialogDialogWorkbookTabSplit = '\002\205\001\237',
	ExcelXlBuiltInDialogDialogWorkbookProtect = '\002\205\001\241',
	ExcelXlBuiltInDialogDialogScrollbarProperties = '\002\205\001\244',
	ExcelXlBuiltInDialogDialogPivotShowPages = '\002\205\001\245',
	ExcelXlBuiltInDialogDialogTextToColumns = '\002\205\001\246',
	ExcelXlBuiltInDialogDialogFormatCharttype = '\002\205\001\247',
	ExcelXlBuiltInDialogDialogPivotFieldGroup = '\002\205\001\261',
	ExcelXlBuiltInDialogDialogPivotFieldUngroup = '\002\205\001\262',
	ExcelXlBuiltInDialogDialogCheckboxProperties = '\002\205\001\263',
	ExcelXlBuiltInDialogDialogLabelProperties = '\002\205\001\264',
	ExcelXlBuiltInDialogDialogListboxProperties = '\002\205\001\265',
	ExcelXlBuiltInDialogDialogEditboxProperties = '\002\205\001\266',
	ExcelXlBuiltInDialogDialogOpenText = '\002\205\001\271',
	ExcelXlBuiltInDialogDialogPushbuttonProperties = '\002\205\001\275',
	ExcelXlBuiltInDialogDialogFilter = '\002\205\001\277',
	ExcelXlBuiltInDialogDialogFunctionWizard = '\002\205\001\302',
	ExcelXlBuiltInDialogDialogSaveCopyAs = '\002\205\001\310',
	ExcelXlBuiltInDialogDialogOptionsListsAdd = '\002\205\001\312',
	ExcelXlBuiltInDialogDialogSeriesAxes = '\002\205\001\314',
	ExcelXlBuiltInDialogDialogSeriesX = '\002\205\001\315',
	ExcelXlBuiltInDialogDialogSeriesY = '\002\205\001\316',
	ExcelXlBuiltInDialogDialogErrorbarX = '\002\205\001\317',
	ExcelXlBuiltInDialogDialogErrorbarY = '\002\205\001\320',
	ExcelXlBuiltInDialogDialogFormatChart = '\002\205\001\321',
	ExcelXlBuiltInDialogDialogSeriesOrder = '\002\205\001\322',
	ExcelXlBuiltInDialogDialogMailEditMailer = '\002\205\001\326',
	ExcelXlBuiltInDialogDialogStandardWidth = '\002\205\001\330',
	ExcelXlBuiltInDialogDialogScenarioMerge = '\002\205\001\331',
	ExcelXlBuiltInDialogDialogProperties = '\002\205\001\332',
	ExcelXlBuiltInDialogDialogSummaryInfo = '\002\205\001\332',
	ExcelXlBuiltInDialogDialogFindFile = '\002\205\001\333',
	ExcelXlBuiltInDialogDialogActiveCellFont = '\002\205\001\334',
	ExcelXlBuiltInDialogDialogVbaMakeAddIn = '\002\205\001\336',
	ExcelXlBuiltInDialogDialogFileSharing = '\002\205\001\341',
	ExcelXlBuiltInDialogDialogAutocorrect = '\002\205\001\345',
	ExcelXlBuiltInDialogDialogCustomViews = '\002\205\001\355',
	ExcelXlBuiltInDialogDialogInsertNameLabel = '\002\205\001\360',
	ExcelXlBuiltInDialogDialogSeriesShape = '\002\205\001\370',
	ExcelXlBuiltInDialogDialogChartOptionsDataLabels = '\002\205\001\371',
	ExcelXlBuiltInDialogDialogChartOptionsDataTable = '\002\205\001\372',
	ExcelXlBuiltInDialogDialogSetBackgroundPicture = '\002\205\001\375',
	ExcelXlBuiltInDialogDialogDataValidation = '\002\205\002\015',
	ExcelXlBuiltInDialogDialogChartType = '\002\205\002\016',
	ExcelXlBuiltInDialogDialogChartLocation = '\002\205\002\017',
	ExcelXlBuiltInDialogDialogChartSourceData = '\002\205\002\035',
	ExcelXlBuiltInDialogDialogSeriesOptions = '\002\205\002-',
	ExcelXlBuiltInDialogDialogPivotTableOptions = '\002\205\0027',
	ExcelXlBuiltInDialogDialogPivotSolveOrder = '\002\205\0028',
	ExcelXlBuiltInDialogDialogPivotCalculatedField = '\002\205\002:',
	ExcelXlBuiltInDialogDialogPivotCalculatedItem = '\002\205\002<',
	ExcelXlBuiltInDialogDialogConditionalFormatting = '\002\205\002G',
	ExcelXlBuiltInDialogDialogInsertHyperlink = '\002\205\002T',
	ExcelXlBuiltInDialogDialogProtectSharing = '\002\205\002l',
	ExcelXlBuiltInDialogDialogPhonetic = '\002\205\002\213',
	ExcelXlBuiltInDialogDialogImportTextFile = '\002\205\002\232',
	ExcelXlBuiltInDialogDialogWebOptionsGeneral = '\002\205\002\264',
	ExcelXlBuiltInDialogDialogWebOptionsPictures = '\002\205\002\266',
	ExcelXlBuiltInDialogDialogWebOptionsFiles = '\002\205\002\265',
	ExcelXlBuiltInDialogDialogWebOptionsFonts = '\002\205\002\270',
	ExcelXlBuiltInDialogDialogWebOptionsEncoding = '\002\205\002\267'
};
typedef enum ExcelXlBuiltInDialog ExcelXlBuiltInDialog;

enum ExcelXlParameterType {
	ExcelXlParameterTypePrompt = '\002\206\000\000',
	ExcelXlParameterTypeConstant = '\002\206\000\001',
	ExcelXlParameterTypeRange = '\002\206\000\002'
};
typedef enum ExcelXlParameterType ExcelXlParameterType;

enum ExcelXlParameterDataType {
	ExcelXlParameterDataTypeParamTypeUnknown = '\002\207\000\000',
	ExcelXlParameterDataTypeParamTypeChar = '\002\207\000\001',
	ExcelXlParameterDataTypeParamTypeNumeric = '\002\207\000\002',
	ExcelXlParameterDataTypeParamTypeDecimal = '\002\207\000\003',
	ExcelXlParameterDataTypeParamTypeNumber = '\002\207\000\004',
	ExcelXlParameterDataTypeParamTypeSmallInt = '\002\207\000\005',
	ExcelXlParameterDataTypeParamTypeFloat = '\002\207\000\006',
	ExcelXlParameterDataTypeParamTypeReal = '\002\207\000\007',
	ExcelXlParameterDataTypeParamTypeDouble = '\002\207\000\010',
	ExcelXlParameterDataTypeParamTypeVarChar = '\002\207\000\014',
	ExcelXlParameterDataTypeParamTypeDate = '\002\207\000\011',
	ExcelXlParameterDataTypeParamTypeTime = '\002\207\000\012',
	ExcelXlParameterDataTypeParamTypeTimestamp = '\002\207\000\013',
	ExcelXlParameterDataTypeParamTypeLongVarChar = '\002\206\377\377',
	ExcelXlParameterDataTypeParamTypeBinary = '\002\206\377\376',
	ExcelXlParameterDataTypeParamTypeVarBinary = '\002\206\377\375',
	ExcelXlParameterDataTypeParamTypeLongVarBinary = '\002\206\377\374',
	ExcelXlParameterDataTypeParamTypeBigInt = '\002\206\377\373',
	ExcelXlParameterDataTypeParamTypeTinyInt = '\002\206\377\372',
	ExcelXlParameterDataTypeParamTypeBit = '\002\206\377\371'
};
typedef enum ExcelXlParameterDataType ExcelXlParameterDataType;

enum ExcelXlFormControl {
	ExcelXlFormControlButtonControl = '\002\210\000\000',
	ExcelXlFormControlCheckBox = '\002\210\000\001',
	ExcelXlFormControlDropDown = '\002\210\000\002',
	ExcelXlFormControlEditBox = '\002\210\000\003',
	ExcelXlFormControlGroupBox = '\002\210\000\004',
	ExcelXlFormControlLabel = '\002\210\000\005',
	ExcelXlFormControlListBox = '\002\210\000\006',
	ExcelXlFormControlOptionButton = '\002\210\000\007',
	ExcelXlFormControlScrollBar = '\002\210\000\010',
	ExcelXlFormControlSpinner = '\002\210\000\011'
};
typedef enum ExcelXlFormControl ExcelXlFormControl;

enum ExcelXlColumnDataType {
	ExcelXlColumnDataTypeGeneralFormat = '\002\211\000\001',
	ExcelXlColumnDataTypeTextFormat = '\002\211\000\002',
	ExcelXlColumnDataTypeMDYFormat = '\002\211\000\003',
	ExcelXlColumnDataTypeDMYFormat = '\002\211\000\004',
	ExcelXlColumnDataTypeYMDFormat = '\002\211\000\005',
	ExcelXlColumnDataTypeMYDFormat = '\002\211\000\006',
	ExcelXlColumnDataTypeDYMFormat = '\002\211\000\007',
	ExcelXlColumnDataTypeYDMFormat = '\002\211\000\010',
	ExcelXlColumnDataTypeSkipColumn = '\002\211\000\011'
};
typedef enum ExcelXlColumnDataType ExcelXlColumnDataType;

enum ExcelXlQueryType {
	ExcelXlQueryTypeODBCQuery = '\002\212\000\001',
	ExcelXlQueryTypeDAORecordSet = '\002\212\000\002',
	ExcelXlQueryTypeWebQuery = '\002\212\000\004',
	ExcelXlQueryTypeOLEDBQuery = '\002\212\000\005',
	ExcelXlQueryTypeTextImport = '\002\212\000\006',
	ExcelXlQueryTypeADORecordset = '\002\212\000\007',
	ExcelXlQueryTypeFileMakerQuery = '\002\212\000\010'
};
typedef enum ExcelXlQueryType ExcelXlQueryType;

enum ExcelXlCmdType {
	ExcelXlCmdTypeCmdCube = '\002\213\000\001',
	ExcelXlCmdTypeCmdSql = '\002\213\000\002',
	ExcelXlCmdTypeCmdTable = '\002\213\000\003',
	ExcelXlCmdTypeCmdDefault = '\002\213\000\004',
	ExcelXlCmdTypeCmdList = '\002\213\000\005'
};
typedef enum ExcelXlCmdType ExcelXlCmdType;

enum ExcelXlListObjectSourceType {
	ExcelXlListObjectSourceTypeSrcExternal = '\002\215\000\000',
	ExcelXlListObjectSourceTypeSrcRange = '\002\215\000\001',
	ExcelXlListObjectSourceTypeSrcXml = '\002\215\000\002',
	ExcelXlListObjectSourceTypeSrcQuery = '\002\215\000\003',
	ExcelXlListObjectSourceTypeSrcModel = '\002\215\000\004'
};
typedef enum ExcelXlListObjectSourceType ExcelXlListObjectSourceType;

enum ExcelXlFMCriteriaOperator {
	ExcelXlFMCriteriaOperatorCriteriaEquals = '\002\221\000\000',
	ExcelXlFMCriteriaOperatorCriteriaLessThanOrEqualTo = '\002\221\000\001',
	ExcelXlFMCriteriaOperatorCriteriaGreaterThanOrEqualTo = '\002\221\000\002',
	ExcelXlFMCriteriaOperatorCriteriaLessThan = '\002\221\000\003',
	ExcelXlFMCriteriaOperatorCriteriaGreaterThan = '\002\221\000\004',
	ExcelXlFMCriteriaOperatorCriteriaBeginsWith = '\002\221\000\005',
	ExcelXlFMCriteriaOperatorCriteriaEndsWith = '\002\221\000\006',
	ExcelXlFMCriteriaOperatorCriteriaContains = '\002\221\000\007'
};
typedef enum ExcelXlFMCriteriaOperator ExcelXlFMCriteriaOperator;

enum ExcelXlFMCriteriaConditional {
	ExcelXlFMCriteriaConditionalNoCondition = '\002\222\000\000',
	ExcelXlFMCriteriaConditionalAndCondition = '\002\222\000\001',
	ExcelXlFMCriteriaConditionalOrCondition = '\002\222\000\002'
};
typedef enum ExcelXlFMCriteriaConditional ExcelXlFMCriteriaConditional;

enum ExcelXlRangeValueDataType {
	ExcelXlRangeValueDataTypeRangeValueDefault = '\002\223\000\012',
	ExcelXlRangeValueDataTypeRangeValueXMLSpreadsheet = '\002\223\000\013',
	ExcelXlRangeValueDataTypeRangeValueMSPersistXML = '\002\223\000\014'
};
typedef enum ExcelXlRangeValueDataType ExcelXlRangeValueDataType;

enum ExcelXLSubTotalsIndex {
	ExcelXLSubTotalsIndexSubtotalAutomatic = '\002\225\000\001',
	ExcelXLSubTotalsIndexSubtotalSum = '\002\225\000\002',
	ExcelXLSubTotalsIndexSubtotalCount = '\002\225\000\003',
	ExcelXLSubTotalsIndexSubtotalAverage = '\002\225\000\004',
	ExcelXLSubTotalsIndexSubtotalMax = '\002\225\000\005',
	ExcelXLSubTotalsIndexSubtotalMin = '\002\225\000\006',
	ExcelXLSubTotalsIndexSubtotalProduct = '\002\225\000\007',
	ExcelXLSubTotalsIndexSubtotalCountNumbers = '\002\225\000\010',
	ExcelXLSubTotalsIndexSubtotalStandardDeviation = '\002\225\000\011',
	ExcelXLSubTotalsIndexSubtotalStandardDeviationP = '\002\225\000\012',
	ExcelXLSubTotalsIndexSubtotalVariable = '\002\225\000\013',
	ExcelXLSubTotalsIndexSubtotalVariableP = '\002\225\000\014'
};
typedef enum ExcelXLSubTotalsIndex ExcelXLSubTotalsIndex;

enum ExcelXLDataEntryMode {
	ExcelXLDataEntryModeDataEntryOn = '\002\226\000\001',
	ExcelXLDataEntryModeDataEntryStrict = '\002\226\000\002',
	ExcelXLDataEntryModeDataEntryOff = '\002\225\357\316'
};
typedef enum ExcelXLDataEntryMode ExcelXLDataEntryMode;

enum ExcelXLStatusBarState {
	ExcelXLStatusBarStateStatusText = '\002\226\377\377',
	ExcelXLStatusBarStateABoolean = '\002\227\000\000'
};
typedef enum ExcelXLStatusBarState ExcelXLStatusBarState;

enum ExcelXLTransitionMenuKeyAction {
	ExcelXLTransitionMenuKeyActionExcelMenus = '\002\230\000\001'
};
typedef enum ExcelXLTransitionMenuKeyAction ExcelXLTransitionMenuKeyAction;

enum ExcelXLDefaultSheetDir {
	ExcelXLDefaultSheetDirLeftToRight = '\002\230\354u',
	ExcelXLDefaultSheetDirRightToLeft = '\002\230\354t',
	ExcelXLDefaultSheetDirContext = '\002\230\354v'
};
typedef enum ExcelXLDefaultSheetDir ExcelXLDefaultSheetDir;

enum ExcelXLCusorMovement {
	ExcelXLCusorMovementNormalCursor = '\002\232\000\000',
	ExcelXLCusorMovementLogicalCursor = '\002\232\000\001',
	ExcelXLCusorMovementVisualCursor = '\002\232\000\002'
};
typedef enum ExcelXLCusorMovement ExcelXLCusorMovement;

enum ExcelXLRangeReference {
	ExcelXLRangeReferenceRangeObject = '\002\233\000\001' /* range object */,
	ExcelXLRangeReferenceA1StyleRangeReference = '\002\233\000\002' /* range R1C1 reference */,
	ExcelXLRangeReferenceNamedRange = '\002\233\000\003' /* range R1C1 reference */
};
typedef enum ExcelXLRangeReference ExcelXLRangeReference;

enum ExcelXLSubTotalType {
	ExcelXLSubTotalTypeAutomaticSubtotal = '\002\234\000\001',
	ExcelXLSubTotalTypeSumSubtotal = '\002\234\000\002',
	ExcelXLSubTotalTypeCountSubtotal = '\002\234\000\003',
	ExcelXLSubTotalTypeAverageSubtotal = '\002\234\000\004',
	ExcelXLSubTotalTypeMaximumValue = '\002\234\000\005',
	ExcelXLSubTotalTypeMinimumValue = '\002\234\000\006',
	ExcelXLSubTotalTypeProductSubtotal = '\002\234\000\007',
	ExcelXLSubTotalTypeCountNumbersSubtotal = '\002\234\000\010',
	ExcelXLSubTotalTypeStandardDeviation = '\002\234\000\011',
	ExcelXLSubTotalTypeStandardDeviationP = '\002\234\000\012',
	ExcelXLSubTotalTypeVarianceSubtotal = '\002\234\000\013',
	ExcelXLSubTotalTypeVariancePSubtotal = '\002\234\000\014'
};
typedef enum ExcelXLSubTotalType ExcelXLSubTotalType;

enum ExcelXLAutoShowType {
	ExcelXLAutoShowTypeType_automatic = '\002\234\357\367',
	ExcelXLAutoShowTypeType_manual = '\002\234\357\331'
};
typedef enum ExcelXLAutoShowType ExcelXLAutoShowType;

enum ExcelXLAutoShowPosition {
	ExcelXLAutoShowPositionPositionTop = '\002\235\357\300',
	ExcelXLAutoShowPositionPositionBottom = '\002\235\357\365'
};
typedef enum ExcelXLAutoShowPosition ExcelXLAutoShowPosition;

enum ExcelXLScrollTabPosition {
	ExcelXLScrollTabPositionScrollTabPositionFirst = '\002\237\000\000',
	ExcelXLScrollTabPositionScrollTabPositionLast = '\002\237\000\001'
};
typedef enum ExcelXLScrollTabPosition ExcelXLScrollTabPosition;

enum ExcelXLPivotTableWizardSourceData {
	ExcelXLPivotTableWizardSourceDataRange = '\002\240\000\000',
	ExcelXLPivotTableWizardSourceDataAListOfRanges = '\002\240\000\001',
	ExcelXLPivotTableWizardSourceDataReportName = '\002\240\000\002',
	ExcelXLPivotTableWizardSourceDataAListOfStringThatIsASQLQuery = '\002\240\000\003'
};
typedef enum ExcelXLPivotTableWizardSourceData ExcelXLPivotTableWizardSourceData;

enum ExcelXLDefaultChartTemplate {
	ExcelXLDefaultChartTemplateBuiltInChartTemplate = '\002\241\000\001',
	ExcelXLDefaultChartTemplateFormatName = '\002\241\000\002'
};
typedef enum ExcelXLDefaultChartTemplate ExcelXLDefaultChartTemplate;

enum ExcelXLDefaultChartType {
	ExcelXLDefaultChartTypeBuiltInChartType = '\002\242\000\025',
	ExcelXLDefaultChartTypeCustomChart = '\002\241\357\356'
};
typedef enum ExcelXLDefaultChartType ExcelXLDefaultChartType;

enum ExcelXLCustomListType {
	ExcelXLCustomListTypeRangeObject = '\002\243\000\001' /* range object */,
	ExcelXLCustomListTypeA1StyleRangeReference = '\002\243\000\002' /* range R1C1 reference */,
	ExcelXLCustomListTypeNamedRange = '\002\243\000\003' /* range R1C1 reference */,
	ExcelXLCustomListTypeListOfStrings = '\002\243\000\004'
};
typedef enum ExcelXLCustomListType ExcelXLCustomListType;

enum ExcelXLInputDefault {
	ExcelXLInputDefaultRangeObject = '\002\244\000\001' /* range object */,
	ExcelXLInputDefaultA1StyleRangeReference = '\002\244\000\002' /* range R1C1 reference */,
	ExcelXLInputDefaultNamedRange = '\002\244\000\003' /* range R1C1 reference */,
	ExcelXLInputDefaultInputDefaultAsString = '\002\244\000\004'
};
typedef enum ExcelXLInputDefault ExcelXLInputDefault;

enum ExcelXLInputType {
	ExcelXLInputTypeANumber = '\002\245\000\001' /* range object */,
	ExcelXLInputTypeInputTypeAsString = '\002\245\000\002' /* range R1C1 reference */,
	ExcelXLInputTypeANumberOrAString = '\002\245\000\003' /* range R1C1 reference */,
	ExcelXLInputTypeABool = '\002\245\000\004',
	ExcelXLInputTypeRangeObject = '\002\245\000\010',
	ExcelXLInputTypeListOfNumbers = '\002\245\000A',
	ExcelXLInputTypeListOfStrings = '\002\245\000B',
	ExcelXLInputTypeListOfNumberOrString = '\002\245\000C',
	ExcelXLInputTypeListOfBools = '\002\245\000D',
	ExcelXLInputTypeListOfRangeObjects = '\002\245\000H'
};
typedef enum ExcelXLInputType ExcelXLInputType;

enum ExcelXLZoomType {
	ExcelXLZoomTypeANumber = '\002\246\000\001',
	ExcelXLZoomTypeABool = '\002\246\000\004'
};
typedef enum ExcelXLZoomType ExcelXLZoomType;

enum ExcelXLSourceDataLocation {
	ExcelXLSourceDataLocationRangeObject = '\002\247\000\001' /* range object */,
	ExcelXLSourceDataLocationA1StyleRangeReference = '\002\247\000\002' /* range R1C1 reference */,
	ExcelXLSourceDataLocationNamedRange = '\002\247\000\003' /* range R1C1 reference */,
	ExcelXLSourceDataLocationListOfStrings = '\002\247\000\004' /* A list of SQL query strings */
};
typedef enum ExcelXLSourceDataLocation ExcelXLSourceDataLocation;

enum ExcelXLSourceData {
	ExcelXLSourceDataPercentable = '\002\250\000\001' /* A percentage between 10 and 400 */,
	ExcelXLSourceDataABool = '\002\250\000\004'
};
typedef enum ExcelXLSourceData ExcelXLSourceData;

enum ExcelXLOnDataType {
	ExcelXLOnDataTypeScript = '\002\251\000\001' /* A script object */,
	ExcelXLOnDataTypeScriptText = '\002\251\000\002'
};
typedef enum ExcelXLOnDataType ExcelXLOnDataType;

enum ExcelXlRangeTarget {
	ExcelXlRangeTargetApplication = '\002\252\000\001',
	ExcelXlRangeTargetWorksheet = '\002\252\000\002',
	ExcelXlRangeTargetA1StyleRangeReference = '\002\252\000\003'
};
typedef enum ExcelXlRangeTarget ExcelXlRangeTarget;

enum ExcelXlHorizAlignmentTarget {
	ExcelXlHorizAlignmentTargetHorizontalAligmentBottom = '\002\252\357\365',
	ExcelXlHorizAlignmentTargetHorizontalAligmentLeft = '\002\252\357\335',
	ExcelXlHorizAlignmentTargetHorizontalAligmentRight = '\002\252\357\310',
	ExcelXlHorizAlignmentTargetHorizontalAligmentTop = '\002\252\357\300'
};
typedef enum ExcelXlHorizAlignmentTarget ExcelXlHorizAlignmentTarget;

enum ExcelXlVerticalAlignmentTarget {
	ExcelXlVerticalAlignmentTargetVerticalAlignmentTop = '\002\253\357\300',
	ExcelXlVerticalAlignmentTargetVerticalAlignmentCenter = '\002\253\357\364',
	ExcelXlVerticalAlignmentTargetVerticalAlignmentBottom = '\002\253\357\365',
	ExcelXlVerticalAlignmentTargetVerticalAlignmentJustify = '\002\253\357\336',
	ExcelXlVerticalAlignmentTargetVerticalAlignmentDistributed = '\002\253\357\353'
};
typedef enum ExcelXlVerticalAlignmentTarget ExcelXlVerticalAlignmentTarget;

enum ExcelXlCheckBoxState {
	ExcelXlCheckBoxStateCheckboxOff = '\002\254\357\316',
	ExcelXlCheckBoxStateCheckboxOn = '\002\255\000\001',
	ExcelXlCheckBoxStateCheckboxMixed = '\002\255\000\002'
};
typedef enum ExcelXlCheckBoxState ExcelXlCheckBoxState;

enum ExcelXlEditBoxItem {
	ExcelXlEditBoxItemText = '\002\255\357\302',
	ExcelXlEditBoxItemANumber = '\002\256\000\002',
	ExcelXlEditBoxItemXlNumber = '\002\255\357\317',
	ExcelXlEditBoxItemReference = '\002\256\000\004',
	ExcelXlEditBoxItemFormula = '\002\256\000\005'
};
typedef enum ExcelXlEditBoxItem ExcelXlEditBoxItem;

enum ExcelXlMultiSelect {
	ExcelXlMultiSelectSelectNone = '\002\261\357\322',
	ExcelXlMultiSelectSelectSimple = '\002\261\357\306',
	ExcelXlMultiSelectSelectExtended = '\002\262\000\003'
};
typedef enum ExcelXlMultiSelect ExcelXlMultiSelect;

enum ExcelXlReplacements {
	ExcelXlReplacementsTextToReplace = '\002\263\000\001',
	ExcelXlReplacementsReplacementText = '\002\263\000\002'
};
typedef enum ExcelXlReplacements ExcelXlReplacements;

enum ExcelXlCategoryNames {
	ExcelXlCategoryNamesRangeObject = '\002\264\000\001' /* range object */,
	ExcelXlCategoryNamesA1StyleRangeReference = '\002\264\000\002' /* range R1C1 reference */,
	ExcelXlCategoryNamesNamedRange = '\002\264\000\003' /* range R1C1 reference */,
	ExcelXlCategoryNamesListOfCategoryNames = '\002\264\000\004' /* A list category names */
};
typedef enum ExcelXlCategoryNames ExcelXlCategoryNames;

enum ExcelMyUDateLinks {
	ExcelMyUDateLinksDoNotUpdateLinks = '\002\266\000\000',
	ExcelMyUDateLinksUpdateExternalLinksOnly = '\002\266\000\001',
	ExcelMyUDateLinksUpdateRemoteLinksOnly = '\002\266\000\002',
	ExcelMyUDateLinksUpdateRemoteAndExternalLinks = '\002\266\000\003'
};
typedef enum ExcelMyUDateLinks ExcelMyUDateLinks;

enum ExcelMyODelimiter {
	ExcelMyODelimiterTabDelimiter = '\002\267\000\001',
	ExcelMyODelimiterCommasDelimiter = '\002\267\000\002',
	ExcelMyODelimiterSpacesDelimiter = '\002\267\000\003',
	ExcelMyODelimiterSemicolonDelimiter = '\002\267\000\004',
	ExcelMyODelimiterNoDelimiter = '\002\267\000\005',
	ExcelMyODelimiterCustomCharacterDelimiter = '\002\267\000\006'
};
typedef enum ExcelMyODelimiter ExcelMyODelimiter;

enum ExcelXlColorVariance {
	ExcelXlColorVarianceVaryByColor = '\002\270\000\001',
	ExcelXlColorVarianceVaryByShade = '\002\270\000\002',
	ExcelXlColorVarianceVaryByGrayscale = '\002\270\000\003',
	ExcelXlColorVarianceVaryBySameColor = '\002\270\000\004'
};
typedef enum ExcelXlColorVariance ExcelXlColorVariance;

enum ExcelXlTickHAlign {
	ExcelXlTickHAlignAlignTickLabelCenter = '\002\272\357\364',
	ExcelXlTickHAlignAlignTickLabelLeft = '\002\272\357\335',
	ExcelXlTickHAlignAlignTickLabelRight = '\002\272\357\310'
};
typedef enum ExcelXlTickHAlign ExcelXlTickHAlign;

enum ExcelXlLanguage {
	ExcelXlLanguageBasque = '\002\274\004-',
	ExcelXlLanguageCatalan = '\002\274\004\003',
	ExcelXlLanguageChinese = '\002\274\010\004',
	ExcelXlLanguageChineseTaiwan = '\002\274\004\004',
	ExcelXlLanguageCzech = '\002\274\004\005',
	ExcelXlLanguageDanish = '\002\274\004\006',
	ExcelXlLanguageDutch = '\002\274\004\023',
	ExcelXlLanguageEnglishUS = '\002\274\004\011',
	ExcelXlLanguageEnglishAUS = '\002\274\014\011',
	ExcelXlLanguageEnglishBritish = '\002\274\010\011',
	ExcelXlLanguageEnglishCAN = '\002\274\020\011',
	ExcelXlLanguageFinnish = '\002\274\004\013',
	ExcelXlLanguageFrench = '\002\274\004\014',
	ExcelXlLanguageFrenchCanadian = '\002\274\014\014',
	ExcelXlLanguageGerman = '\002\274\004\007',
	ExcelXlLanguageGermanAustria = '\002\274\014\007',
	ExcelXlLanguageSwissGerman = '\002\274\010\007',
	ExcelXlLanguageGreek = '\002\274\004\010',
	ExcelXlLanguageHungarian = '\002\274\004\016',
	ExcelXlLanguageItalian = '\002\274\004\020',
	ExcelXlLanguageJapanese = '\002\274\004\021',
	ExcelXlLanguageKorean = '\002\274\004\022',
	ExcelXlLanguageMalaysian = '\002\274\004>',
	ExcelXlLanguageNorwegianBokmal = '\002\274\004\024',
	ExcelXlLanguageNorwegian = '\002\274\004,',
	ExcelXlLanguagePolish = '\002\274\004\025',
	ExcelXlLanguagePortugueseBrazil = '\002\274\004\026',
	ExcelXlLanguagePortugueseIberian = '\002\274\010\026',
	ExcelXlLanguageRussian = '\002\274\004\031',
	ExcelXlLanguageSlovak = '\002\274\004\033',
	ExcelXlLanguageSlovenian = '\002\274\004$',
	ExcelXlLanguageSpanish = '\002\274\004\012',
	ExcelXlLanguageSwedish = '\002\274\004\035',
	ExcelXlLanguageTurkish = '\002\274\004\037'
};
typedef enum ExcelXlLanguage ExcelXlLanguage;

enum ExcelXlSortOn {
	ExcelXlSortOnSortOnCellValue = '\002\275\000\000' /* Values. */,
	ExcelXlSortOnSortOnCellColor = '\002\275\000\001' /* Cell color. */,
	ExcelXlSortOnSortOnFontColor = '\002\275\000\002' /* Font color. */,
	ExcelXlSortOnSortOnIcon = '\002\275\000\003' /* Icon. */
};
typedef enum ExcelXlSortOn ExcelXlSortOn;

enum ExcelXlSortDataOption {
	ExcelXlSortDataOptionSortNormal = '\002\276\000\000' /* Default. Sorts numeric and text data separately. */,
	ExcelXlSortDataOptionSortTextAsNumbers = '\002\276\000\001' /* Treat text as numeric data for the sort. */
};
typedef enum ExcelXlSortDataOption ExcelXlSortDataOption;

enum ExcelXlTotalsCalculation {
	ExcelXlTotalsCalculationNoneTotalsCalc = '\002\277\000\000',
	ExcelXlTotalsCalculationSumTotalsCalc = '\002\277\000\001',
	ExcelXlTotalsCalculationAverageTotalsCalc = '\002\277\000\002',
	ExcelXlTotalsCalculationCountTotalsCalc = '\002\277\000\003',
	ExcelXlTotalsCalculationCountNumberTotalsCalc = '\002\277\000\004',
	ExcelXlTotalsCalculationMinTotalsCalc = '\002\277\000\005',
	ExcelXlTotalsCalculationMaxTotalsCalc = '\002\277\000\006',
	ExcelXlTotalsCalculationDeviationTotalsCalc = '\002\277\000\007',
	ExcelXlTotalsCalculationVarTotalsCalc = '\002\277\000\010',
	ExcelXlTotalsCalculationCustomTotalsCalc = '\002\277\000\011'
};
typedef enum ExcelXlTotalsCalculation ExcelXlTotalsCalculation;

enum ExcelMsoChartElementType {
	ExcelMsoChartElementTypeNoChartTitle = '\003\206\000\000',
	ExcelMsoChartElementTypeChartTitleCenteredOverlay = '\003\206\000\001',
	ExcelMsoChartElementTypeChartTitleAboveChart = '\003\206\000\002',
	ExcelMsoChartElementTypeNoLegend = '\003\206\000d',
	ExcelMsoChartElementTypeLegendRight = '\003\206\000e',
	ExcelMsoChartElementTypeLegendTop = '\003\206\000f',
	ExcelMsoChartElementTypeLegendLeft = '\003\206\000g',
	ExcelMsoChartElementTypeLegendBottom = '\003\206\000h',
	ExcelMsoChartElementTypeLegendRightOverlay = '\003\206\000i',
	ExcelMsoChartElementTypeLegendLeftOverlay = '\003\206\000j',
	ExcelMsoChartElementTypeNoDataLabel = '\003\206\000\310',
	ExcelMsoChartElementTypeShowDataLabel = '\003\206\000\311',
	ExcelMsoChartElementTypeDataLabelCenter = '\003\206\000\312',
	ExcelMsoChartElementTypeDataLabelInsideEnd = '\003\206\000\313',
	ExcelMsoChartElementTypeDataLabelInsideBase = '\003\206\000\314',
	ExcelMsoChartElementTypeDataLabelOutsideEnd = '\003\206\000\315',
	ExcelMsoChartElementTypeDataLabelLeft = '\003\206\000\316',
	ExcelMsoChartElementTypeDataLabelRight = '\003\206\000\317',
	ExcelMsoChartElementTypeDataLabelTop = '\003\206\000\320',
	ExcelMsoChartElementTypeDataLabelBottom = '\003\206\000\321',
	ExcelMsoChartElementTypeDataLabelBestFit = '\003\206\000\322',
	ExcelMsoChartElementTypeNoPrimaryCategoryAxisTitle = '\003\206\001,',
	ExcelMsoChartElementTypePrimaryCategoryAxisTitleAdjacentToAxis = '\003\206\001-',
	ExcelMsoChartElementTypePrimaryCategoryAxisTitleBelowAxis = '\003\206\001.',
	ExcelMsoChartElementTypePrimaryCategoryAxisTitleRotated = '\003\206\001/',
	ExcelMsoChartElementTypePrimaryCategoryAxisTitleVertical = '\003\206\0010',
	ExcelMsoChartElementTypePrimaryCategoryAxisTitleHorizontal = '\003\206\0011',
	ExcelMsoChartElementTypeNoPrimaryValueAxisTitle = '\003\206\0012',
	ExcelMsoChartElementTypePrimaryValueAxisTitleAdjacentToAxis = '\003\206\0013',
	ExcelMsoChartElementTypePrimaryValueAxisTitleBelowAxis = '\003\206\0014',
	ExcelMsoChartElementTypePrimaryValueAxisTitleRotated = '\003\206\0015',
	ExcelMsoChartElementTypePrimaryValueAxisTitleVertical = '\003\206\0016',
	ExcelMsoChartElementTypePrimaryValueAxisTitleHorizontal = '\003\206\0017',
	ExcelMsoChartElementTypeNoSecondaryCategoryAxisTitle = '\003\206\0018',
	ExcelMsoChartElementTypeSecondaryCategoryAxisTitleAdjacentToAxis = '\003\206\0019',
	ExcelMsoChartElementTypeSecondaryCategoryAxisTitleBelowAxis = '\003\206\001:',
	ExcelMsoChartElementTypeSecondaryCategoryAxisTitleRotated = '\003\206\001;',
	ExcelMsoChartElementTypeSecondaryCategoryAxisTitleVertical = '\003\206\001<',
	ExcelMsoChartElementTypeSecondaryCategoryAxisTitleHorizontal = '\003\206\001=',
	ExcelMsoChartElementTypeNoSecondaryValueAxisTitle = '\003\206\001>',
	ExcelMsoChartElementTypeSecondaryValueAxisTitleAdjacentToAxis = '\003\206\001\?',
	ExcelMsoChartElementTypeSecondaryValueAxisTitleBelowAxis = '\003\206\001@',
	ExcelMsoChartElementTypeSecondaryValueAxisTitleRotated = '\003\206\001A',
	ExcelMsoChartElementTypeSecondaryValueAxisTitleVertical = '\003\206\001B',
	ExcelMsoChartElementTypeSecondaryValueAxisTitleHorizontal = '\003\206\001C',
	ExcelMsoChartElementTypeNoSeriesAxisTitle = '\003\206\001D',
	ExcelMsoChartElementTypeSeriesAxisTitleRotated = '\003\206\001E',
	ExcelMsoChartElementTypeSeriesAxisTitleVertical = '\003\206\001F',
	ExcelMsoChartElementTypeSeriesAxisTitleHorizontal = '\003\206\001G',
	ExcelMsoChartElementTypeNoPrimaryValueGridLines = '\003\206\001H',
	ExcelMsoChartElementTypePrimaryValueGridLinesMinor = '\003\206\001I',
	ExcelMsoChartElementTypePrimaryValueGridLinesMajor = '\003\206\001J',
	ExcelMsoChartElementTypePrimaryValueGridLinesMinorMajor = '\003\206\001K',
	ExcelMsoChartElementTypeNoPrimaryCategoryGridLines = '\003\206\001L',
	ExcelMsoChartElementTypePrimaryCategoryGridLinesMinor = '\003\206\001M',
	ExcelMsoChartElementTypePrimaryCategoryGridLinesMajor = '\003\206\001N',
	ExcelMsoChartElementTypePrimaryCategoryGridLinesMinorMajor = '\003\206\001O',
	ExcelMsoChartElementTypeNoSecondaryValueGridLines = '\003\206\001P',
	ExcelMsoChartElementTypeSecondaryValueGridLinesMinor = '\003\206\001Q',
	ExcelMsoChartElementTypeSecondaryValueGridLinesMajor = '\003\206\001R',
	ExcelMsoChartElementTypeSecondaryValueGridLinesMinorMajor = '\003\206\001S',
	ExcelMsoChartElementTypeNoSecondaryCategoryGridLines = '\003\206\001T',
	ExcelMsoChartElementTypeSecondaryCategoryGridLinesMinor = '\003\206\001U',
	ExcelMsoChartElementTypeSecondaryCategoryGridLinesMajor = '\003\206\001V',
	ExcelMsoChartElementTypeSecondaryCategoryGridLinesMinorMajor = '\003\206\001W',
	ExcelMsoChartElementTypeNoSeriesAxisGridLines = '\003\206\001X',
	ExcelMsoChartElementTypeSeriesAxisGridLinesMinor = '\003\206\001Y',
	ExcelMsoChartElementTypeSeriesAxisGridLinesMajor = '\003\206\001Z',
	ExcelMsoChartElementTypeSeriesAxisGridLinesMinorMajor = '\003\206\001[',
	ExcelMsoChartElementTypeNoPrimaryCategoryAxis = '\003\206\001\\',
	ExcelMsoChartElementTypePrimaryCategoryAxisShow = '\003\206\001]',
	ExcelMsoChartElementTypePrimaryCategoryAxisWithoutLabels = '\003\206\001^',
	ExcelMsoChartElementTypePrimaryCategoryAxisReverse = '\003\206\001_',
	ExcelMsoChartElementTypeNoPrimaryValueAxis = '\003\206\001`',
	ExcelMsoChartElementTypeShowPrimaryValueAxis = '\003\206\001a',
	ExcelMsoChartElementTypePrimaryValueAxisThousands = '\003\206\001b',
	ExcelMsoChartElementTypePrimaryValueAxisMillions = '\003\206\001c',
	ExcelMsoChartElementTypePrimaryValueAxisBillions = '\003\206\001d',
	ExcelMsoChartElementTypePrimaryValueAxisLogScale = '\003\206\001e',
	ExcelMsoChartElementTypeNoSecondaryCategoryAxis = '\003\206\001f',
	ExcelMsoChartElementTypeShowSecondaryCategoryAxis = '\003\206\001g',
	ExcelMsoChartElementTypeSecondaryCategoryAxisWithoutLabels = '\003\206\001h',
	ExcelMsoChartElementTypeSecondaryCategoryAxisReverse = '\003\206\001i',
	ExcelMsoChartElementTypeNoSecondaryValueAxis = '\003\206\001j',
	ExcelMsoChartElementTypeShowSecondaryValueAxis = '\003\206\001k',
	ExcelMsoChartElementTypeSecondaryValueAxisThousands = '\003\206\001l',
	ExcelMsoChartElementTypeSecondaryValueAxisMillions = '\003\206\001m',
	ExcelMsoChartElementTypeSecondaryValueAxisBillions = '\003\206\001n',
	ExcelMsoChartElementTypeSecondaryValueAxisLogScale = '\003\206\001o',
	ExcelMsoChartElementTypeNoSeriesAxis = '\003\206\001p',
	ExcelMsoChartElementTypeShowSeriesAxis = '\003\206\001q',
	ExcelMsoChartElementTypeSeriesAxisWithoutLabeling = '\003\206\001r',
	ExcelMsoChartElementTypeSeriesAxisReverse = '\003\206\001s',
	ExcelMsoChartElementTypePrimaryCategoryAxisThousands = '\003\206\001t',
	ExcelMsoChartElementTypePrimaryCategoryAxisMillions = '\003\206\001u',
	ExcelMsoChartElementTypePrimaryCategoryAxisBillions = '\003\206\001v',
	ExcelMsoChartElementTypePrimaryCategoryAxisLogScale = '\003\206\001w',
	ExcelMsoChartElementTypeSecondaryCategoryAxisThousands = '\003\206\001x',
	ExcelMsoChartElementTypeSecondaryCategoryAxisMillions = '\003\206\001y',
	ExcelMsoChartElementTypeSecondaryCategoryAxisBillions = '\003\206\001z',
	ExcelMsoChartElementTypeSecondaryCategoryAxisLogScale = '\003\206\001{',
	ExcelMsoChartElementTypeNoDataTable = '\003\206\001\364',
	ExcelMsoChartElementTypeShowDataTable = '\003\206\001\365',
	ExcelMsoChartElementTypeDataTableWithLegendKeys = '\003\206\001\366',
	ExcelMsoChartElementTypeNoTrendline = '\003\206\002X',
	ExcelMsoChartElementTypeTrendlineAddLinear = '\003\206\002Y',
	ExcelMsoChartElementTypeTrendlineAddExponential = '\003\206\002Z',
	ExcelMsoChartElementTypeTrendlineAddLinearForecast = '\003\206\002[',
	ExcelMsoChartElementTypeTrendlineAddTwoPeriodMovingAverage = '\003\206\002\\',
	ExcelMsoChartElementTypeNoErrorBar = '\003\206\002\274',
	ExcelMsoChartElementTypeErrorBarStandardError = '\003\206\002\275',
	ExcelMsoChartElementTypeErrorBarPercentage = '\003\206\002\276',
	ExcelMsoChartElementTypeErrorBarStandardDeviation = '\003\206\002\277',
	ExcelMsoChartElementTypeNoLine = '\003\206\003 ',
	ExcelMsoChartElementTypeLineDropLine = '\003\206\003!',
	ExcelMsoChartElementTypeLineHiLoLine = '\003\206\003\"',
	ExcelMsoChartElementTypeLineSeriesLine = '\003\206\003#',
	ExcelMsoChartElementTypeLineDropHiloLine = '\003\206\003$',
	ExcelMsoChartElementTypeNoUpDownBars = '\003\206\003\204',
	ExcelMsoChartElementTypeShowUpDownBars = '\003\206\003\205',
	ExcelMsoChartElementTypeNoPlotArea = '\003\206\003\350',
	ExcelMsoChartElementTypeShowPlotArea = '\003\206\003\351',
	ExcelMsoChartElementTypeNoChartWall = '\003\206\004L',
	ExcelMsoChartElementTypeShowChartWall = '\003\206\004M',
	ExcelMsoChartElementTypeNoChartFloor = '\003\206\004\260',
	ExcelMsoChartElementTypeShowChartFloor = '\003\206\004\261'
};
typedef enum ExcelMsoChartElementType ExcelMsoChartElementType;

enum ExcelXlDynamicFilterCriteria {
	ExcelXlDynamicFilterCriteriaFilterAboveAverage = '\003\207\000!' /* Filter all above-average values. */,
	ExcelXlDynamicFilterCriteriaFilterAllDatesInApril = '\003\207\000\030' /* Filter all dates in April. */,
	ExcelXlDynamicFilterCriteriaFilterAllDatesInAugust = '\003\207\000\034' /* Filter all dates in August. */,
	ExcelXlDynamicFilterCriteriaFilterAllDatesInDecember = '\003\207\000 ' /* Filter all dates in December. */,
	ExcelXlDynamicFilterCriteriaFilterAllDatesInFebruary = '\003\207\000\026' /* Filter all dates in February */,
	ExcelXlDynamicFilterCriteriaFilterAllDatesInJanuary = '\003\207\000\025' /* Filter all dates in January. */,
	ExcelXlDynamicFilterCriteriaFilterAllDatesInJuly = '\003\207\000\033' /* Filter all dates in July. */,
	ExcelXlDynamicFilterCriteriaFilterAllDatesInJune = '\003\207\000\032' /* Filter all dates in June. */,
	ExcelXlDynamicFilterCriteriaFilterAllDatesInMarch = '\003\207\000\027' /* Filter all dates in March. */,
	ExcelXlDynamicFilterCriteriaFilterAllDatesInMay = '\003\207\000\031' /* Filter all dates in May. */,
	ExcelXlDynamicFilterCriteriaFilterAllDatesInNovember = '\003\207\000\037' /* Filter all dates in November. */,
	ExcelXlDynamicFilterCriteriaFilterAllDatesInOctober = '\003\207\000\036' /* Filter all dates in October. */,
	ExcelXlDynamicFilterCriteriaFilterAllDatesInQuarter1 = '\003\207\000\021' /* Filter all dates in Quarter1. */,
	ExcelXlDynamicFilterCriteriaFilterAllDatesInQuarter2 = '\003\207\000\022' /* Filter all dates in Quarter2. */,
	ExcelXlDynamicFilterCriteriaFilterAllDatesInQuarter3 = '\003\207\000\023' /* Filter all dates in Quarter3. */,
	ExcelXlDynamicFilterCriteriaFilterAllDatesInQuarter4 = '\003\207\000\024' /* Filter all dates in Quarter4. */,
	ExcelXlDynamicFilterCriteriaFilterAllDatesInSeptember = '\003\207\000\035' /* Filter all dates in September. */,
	ExcelXlDynamicFilterCriteriaFilterBelowAverage = '\003\207\000\"' /* Filter all below-average values. */,
	ExcelXlDynamicFilterCriteriaFilterLastMonth = '\003\207\000\010' /* Filter all values related to last month. */,
	ExcelXlDynamicFilterCriteriaFilterLastQuarter = '\003\207\000\013' /* Filter all values related to last quarter. */,
	ExcelXlDynamicFilterCriteriaFilterLastWeek = '\003\207\000\005' /* Filter all values related to last week. */,
	ExcelXlDynamicFilterCriteriaFilterLastYear = '\003\207\000\016' /* Filter all values related to last year. */,
	ExcelXlDynamicFilterCriteriaFilterNextMonth = '\003\207\000\011' /* Filter all values related to next month. */,
	ExcelXlDynamicFilterCriteriaFilterNextQuarter = '\003\207\000\014' /* Filter all values related to next quarter. */,
	ExcelXlDynamicFilterCriteriaFilterNextWeek = '\003\207\000\006' /* Filter all values related to next week. */,
	ExcelXlDynamicFilterCriteriaFilterNextYear = '\003\207\000\017' /* Filter all values related to next year. */,
	ExcelXlDynamicFilterCriteriaFilterThisMonth = '\003\207\000\007' /* Filter all values related to the current month. */,
	ExcelXlDynamicFilterCriteriaFilterThisQuarter = '\003\207\000\012' /* Filter all values related to the current quarter. */,
	ExcelXlDynamicFilterCriteriaFilterThisWeek = '\003\207\000\004' /* Filter all values related to the current week. */,
	ExcelXlDynamicFilterCriteriaFilterThisYear = '\003\207\000\015' /* Filter all values related to the current year. */,
	ExcelXlDynamicFilterCriteriaFilterToday = '\003\207\000\001' /* Filter all values related to the current date. */,
	ExcelXlDynamicFilterCriteriaFilterTomorrow = '\003\207\000\003' /* Filter all values related to tomorrow. */,
	ExcelXlDynamicFilterCriteriaFilterYearToDate = '\003\207\000\020' /* Filter all values from today until a year ago. */,
	ExcelXlDynamicFilterCriteriaFilterYesterday = '\003\207\000\002' /* Filter all values related to yesterday. */
};
typedef enum ExcelXlDynamicFilterCriteria ExcelXlDynamicFilterCriteria;

enum ExcelXlThemeFont {
	ExcelXlThemeFontThemeFontIndexNone = '\002\300\000\000',
	ExcelXlThemeFontThemeFontIndexMajor = '\002\300\000\001',
	ExcelXlThemeFontThemeFontIndexMinor = '\002\300\000\002'
};
typedef enum ExcelXlThemeFont ExcelXlThemeFont;

enum ExcelXlThemeColor {
	ExcelXlThemeColorColorIndexNone = '\002\300\357\322',
	ExcelXlThemeColorFirstDarkThemeColor = '\002\301\000\001',
	ExcelXlThemeColorFirstLightThemeColor = '\002\301\000\002',
	ExcelXlThemeColorSecondDarkThemeColor = '\002\301\000\003',
	ExcelXlThemeColorSecondLightThemeColor = '\002\301\000\004',
	ExcelXlThemeColorFirstAccentThemeColor = '\002\301\000\005',
	ExcelXlThemeColorSecondAccentThemeColor = '\002\301\000\006',
	ExcelXlThemeColorThirdAccentThemeColor = '\002\301\000\007',
	ExcelXlThemeColorFourthAccentThemeColor = '\002\301\000\010',
	ExcelXlThemeColorFifthAccentThemeColor = '\002\301\000\011',
	ExcelXlThemeColorSixthAccentThemeColor = '\002\301\000\012',
	ExcelXlThemeColorHyperlinkThemeColor = '\002\301\000\013',
	ExcelXlThemeColorFollowedHyperlinkThemeColor = '\002\301\000\014'
};
typedef enum ExcelXlThemeColor ExcelXlThemeColor;

enum ExcelXlCheckInVersionType {
	ExcelXlCheckInVersionTypeMinorVersion = '\002\304\000\000',
	ExcelXlCheckInVersionTypeMajorVersion = '\002\304\000\001',
	ExcelXlCheckInVersionTypeOverwriteCurrentVersion = '\002\304\000\002'
};
typedef enum ExcelXlCheckInVersionType ExcelXlCheckInVersionType;

enum ExcelXlWebSelectionType {
	ExcelXlWebSelectionTypeEntirePage = '\002\305\000\000',
	ExcelXlWebSelectionTypeAllTables = '\002\305\000\001',
	ExcelXlWebSelectionTypeSpecifiedTables = '\002\305\000\002'
};
typedef enum ExcelXlWebSelectionType ExcelXlWebSelectionType;

enum ExcelXlWebFormatting {
	ExcelXlWebFormattingWebFormattingAll = '\002\306\000\000',
	ExcelXlWebFormattingWebFormattingRtf = '\002\306\000\001',
	ExcelXlWebFormattingWebFormattingNone = '\002\306\000\002'
};
typedef enum ExcelXlWebFormatting ExcelXlWebFormatting;

enum ExcelXlRobustConnect {
	ExcelXlRobustConnectAsRequired = '\002\307\000\000',
	ExcelXlRobustConnectAlways = '\002\307\000\001',
	ExcelXlRobustConnectNever = '\002\307\000\002'
};
typedef enum ExcelXlRobustConnect ExcelXlRobustConnect;

enum ExcelXlConditionValueTypes {
	ExcelXlConditionValueTypesConditionValueNone = '\002\307\377\377',
	ExcelXlConditionValueTypesConditionValueNumber = '\002\310\000\000',
	ExcelXlConditionValueTypesConditionValueLowestValue = '\002\310\000\001',
	ExcelXlConditionValueTypesConditionValueHighestValue = '\002\310\000\002',
	ExcelXlConditionValueTypesConditionValuePercent = '\002\310\000\003',
	ExcelXlConditionValueTypesConditionValueFormula = '\002\310\000\004',
	ExcelXlConditionValueTypesConditionValuePercentile = '\002\310\000\005',
	ExcelXlConditionValueTypesConditionValueAutomaticMinimum = '\002\310\000\006',
	ExcelXlConditionValueTypesConditionValueAutomaticMaximum = '\002\310\000\007'
};
typedef enum ExcelXlConditionValueTypes ExcelXlConditionValueTypes;

enum ExcelXlPivotConditionScope {
	ExcelXlPivotConditionScopePivotConditionSelectionScope = '\002\311\000\000',
	ExcelXlPivotConditionScopePivotConditionFieldsScope = '\002\311\000\001',
	ExcelXlPivotConditionScopePivotConditionDataFieldScope = '\002\311\000\002'
};
typedef enum ExcelXlPivotConditionScope ExcelXlPivotConditionScope;

enum ExcelXlDataBarFillType {
	ExcelXlDataBarFillTypeDatabarFillSolid = '\002\312\000\000',
	ExcelXlDataBarFillTypeDatabarFillGradient = '\002\312\000\001'
};
typedef enum ExcelXlDataBarFillType ExcelXlDataBarFillType;

enum ExcelXlDataBarBorderType {
	ExcelXlDataBarBorderTypeDatabarBorderNone = '\002\313\000\000',
	ExcelXlDataBarBorderTypeDatabarBorderSolid = '\002\313\000\001'
};
typedef enum ExcelXlDataBarBorderType ExcelXlDataBarBorderType;

enum ExcelXlDataBarAxisPosition {
	ExcelXlDataBarAxisPositionDatabarAxisAutomatic = '\002\314\000\000',
	ExcelXlDataBarAxisPositionDatabarAxisMidpoint = '\002\314\000\001',
	ExcelXlDataBarAxisPositionDatabarAxisNone = '\002\314\000\002'
};
typedef enum ExcelXlDataBarAxisPosition ExcelXlDataBarAxisPosition;

enum ExcelXlDataBarNegativeFormatType {
	ExcelXlDataBarNegativeFormatTypeDatabarAutomatic = '\002\315\000\000',
	ExcelXlDataBarNegativeFormatTypeDatabarPositiveFormat = '\002\315\000\001',
	ExcelXlDataBarNegativeFormatTypeDatabarCustomFormat = '\002\315\000\002'
};
typedef enum ExcelXlDataBarNegativeFormatType ExcelXlDataBarNegativeFormatType;

enum ExcelXlIcon {
	ExcelXlIconFormatConditionIconNoCellIcon = '\002\315\377\377',
	ExcelXlIconFormatConditionIconGreenUpArrow = '\002\316\000\001',
	ExcelXlIconFormatConditionIconYellowSideArrow = '\002\316\000\002',
	ExcelXlIconFormatConditionIconRedDownArrow = '\002\316\000\003',
	ExcelXlIconFormatConditionIconGrayUpArrow = '\002\316\000\004',
	ExcelXlIconFormatConditionIconGraySideArrow = '\002\316\000\005',
	ExcelXlIconFormatConditionIconGrayDownArrow = '\002\316\000\006',
	ExcelXlIconFormatConditionIconGreenFlag = '\002\316\000\007',
	ExcelXlIconFormatConditionIconYellowFlag = '\002\316\000\010',
	ExcelXlIconFormatConditionIconRedFlag = '\002\316\000\011',
	ExcelXlIconFormatConditionIconGreenCircle = '\002\316\000\012',
	ExcelXlIconFormatConditionIconYellowCircle = '\002\316\000\013',
	ExcelXlIconFormatConditionIconRedCircleWithBorder = '\002\316\000\014',
	ExcelXlIconFormatConditionIconBlackCircleWithBorder = '\002\316\000\015',
	ExcelXlIconFormatConditionIconGreenTrafficLight = '\002\316\000\016',
	ExcelXlIconFormatConditionIconYellowTrafficLight = '\002\316\000\017',
	ExcelXlIconFormatConditionIconRedTrafficLight = '\002\316\000\020',
	ExcelXlIconFormatConditionIconYellowTriangle = '\002\316\000\021',
	ExcelXlIconFormatConditionIconRedDiamond = '\002\316\000\022',
	ExcelXlIconFormatConditionIconGreenCheckSymbol = '\002\316\000\023',
	ExcelXlIconFormatConditionIconYellowExclamationSymbol = '\002\316\000\024',
	ExcelXlIconFormatConditionIconRedCrossSymbol = '\002\316\000\025',
	ExcelXlIconFormatConditionIconGreenCheck = '\002\316\000\026',
	ExcelXlIconFormatConditionIconYellowExclamation = '\002\316\000\027',
	ExcelXlIconFormatConditionIconRedCross = '\002\316\000\030',
	ExcelXlIconFormatConditionIconYellowUpInclineArrow = '\002\316\000\031',
	ExcelXlIconFormatConditionIconYellowDownInclineArrow = '\002\316\000\032',
	ExcelXlIconFormatConditionIconGrayUpInclineArrow = '\002\316\000\033',
	ExcelXlIconFormatConditionIconGrayDownInclineArrow = '\002\316\000\034',
	ExcelXlIconFormatConditionIconRedCircle = '\002\316\000\035',
	ExcelXlIconFormatConditionIconPinkCircle = '\002\316\000\036',
	ExcelXlIconFormatConditionIconGrayCircle = '\002\316\000\037',
	ExcelXlIconFormatConditionIconBlackCircle = '\002\316\000 ',
	ExcelXlIconFormatConditionIconCircleWithOneWhiteQuarter = '\002\316\000!',
	ExcelXlIconFormatConditionIconCircleWithTwoWhiteQuarters = '\002\316\000\"',
	ExcelXlIconFormatConditionIconCircleWithThreeWhiteQuarters = '\002\316\000#',
	ExcelXlIconFormatConditionIconWhiteCircleAllWhiteQuarters = '\002\316\000$',
	ExcelXlIconFormatConditionIcon0Bars = '\002\316\000%',
	ExcelXlIconFormatConditionIcon1Bar = '\002\316\000&',
	ExcelXlIconFormatConditionIcon2Bars = '\002\316\000\'',
	ExcelXlIconFormatConditionIcon3Bars = '\002\316\000(',
	ExcelXlIconFormatConditionIcon4Bars = '\002\316\000)',
	ExcelXlIconFormatConditionIconGoldStar = '\002\316\000*',
	ExcelXlIconFormatConditionIconHalfGoldStar = '\002\316\000+',
	ExcelXlIconFormatConditionIconSilverStar = '\002\316\000,',
	ExcelXlIconFormatConditionIconGreenUpTriangle = '\002\316\000-',
	ExcelXlIconFormatConditionIconYellowDash = '\002\316\000.',
	ExcelXlIconFormatConditionIconRedDownTriangle = '\002\316\000/',
	ExcelXlIconFormatConditionIcon4FilledBoxes = '\002\316\0000',
	ExcelXlIconFormatConditionIcon3FilledBoxes = '\002\316\0001',
	ExcelXlIconFormatConditionIcon2FilledBoxes = '\002\316\0002',
	ExcelXlIconFormatConditionIcon1FilledBox = '\002\316\0003',
	ExcelXlIconFormatConditionIcon0FilledBoxes = '\002\316\0004'
};
typedef enum ExcelXlIcon ExcelXlIcon;

enum ExcelXlIconSet {
	ExcelXlIconSetIconSetCustom = '\002\316\377\377',
	ExcelXlIconSetIconSet3Arrows = '\002\317\000\001',
	ExcelXlIconSetIconSet3ArrowsGray = '\002\317\000\002',
	ExcelXlIconSetIconSet3Flags = '\002\317\000\003',
	ExcelXlIconSetIconSet3TrafficLights1 = '\002\317\000\004',
	ExcelXlIconSetIconSet3TrafficLights2 = '\002\317\000\005',
	ExcelXlIconSetIconSet3Signs = '\002\317\000\006',
	ExcelXlIconSetIconSet3Symbols = '\002\317\000\007',
	ExcelXlIconSetIconSet3Symbols2 = '\002\317\000\010',
	ExcelXlIconSetIconSet4Arrows = '\002\317\000\011',
	ExcelXlIconSetIconSet4ArrowsGray = '\002\317\000\012',
	ExcelXlIconSetIconSet4RedToBlack = '\002\317\000\013',
	ExcelXlIconSetIconSet4CRV = '\002\317\000\014',
	ExcelXlIconSetIconSet4TrafficLights = '\002\317\000\015',
	ExcelXlIconSetIconSet5Arrows = '\002\317\000\016',
	ExcelXlIconSetIconSet5ArrowsGray = '\002\317\000\017',
	ExcelXlIconSetIconSet5CRV = '\002\317\000\020',
	ExcelXlIconSetIconSet5Quarters = '\002\317\000\021',
	ExcelXlIconSetIconSet3Stars = '\002\317\000\022',
	ExcelXlIconSetIconSet3Triangles = '\002\317\000\023',
	ExcelXlIconSetIconSet5Boxes = '\002\317\000\024'
};
typedef enum ExcelXlIconSet ExcelXlIconSet;

enum ExcelXlTopBottom {
	ExcelXlTopBottomTop10Top = '\002\320\000\001',
	ExcelXlTopBottomTop10Bottom = '\002\320\000\000'
};
typedef enum ExcelXlTopBottom ExcelXlTopBottom;

enum ExcelXlCalcFor {
	ExcelXlCalcForCalcForAllValues = '\002\321\000\000',
	ExcelXlCalcForCalcForRowGroups = '\002\321\000\001',
	ExcelXlCalcForCalcForColGroups = '\002\321\000\002'
};
typedef enum ExcelXlCalcFor ExcelXlCalcFor;

enum ExcelXlAboveBelow {
	ExcelXlAboveBelowFormatAboveAverage = '\002\322\000\000',
	ExcelXlAboveBelowFormatBelowAverage = '\002\322\000\001',
	ExcelXlAboveBelowFormatEqualAboveAverage = '\002\322\000\002',
	ExcelXlAboveBelowFormatEqualBelowAverage = '\002\322\000\003',
	ExcelXlAboveBelowFormatAboveStandardDeviation = '\002\322\000\004',
	ExcelXlAboveBelowFormatBelowStandardDeviation = '\002\322\000\005'
};
typedef enum ExcelXlAboveBelow ExcelXlAboveBelow;

enum ExcelXlDupeUnique {
	ExcelXlDupeUniqueFormatUniqueValues = '\002\323\000\000',
	ExcelXlDupeUniqueFormatDuplicateValues = '\002\323\000\001'
};
typedef enum ExcelXlDupeUnique ExcelXlDupeUnique;

enum ExcelXlContainsOperator {
	ExcelXlContainsOperatorTextContains = '\002\324\000\000',
	ExcelXlContainsOperatorTextDoesNotContain = '\002\324\000\001',
	ExcelXlContainsOperatorTextBeginsWith = '\002\324\000\002',
	ExcelXlContainsOperatorTextEndsWith = '\002\324\000\003'
};
typedef enum ExcelXlContainsOperator ExcelXlContainsOperator;

enum ExcelXlTimePeriods {
	ExcelXlTimePeriodsDateIsToday = '\002\325\000\000',
	ExcelXlTimePeriodsDateIsYesterday = '\002\325\000\001',
	ExcelXlTimePeriodsDateIsWithinTheLastSevenDays = '\002\325\000\002',
	ExcelXlTimePeriodsDateIsThisWeek = '\002\325\000\003',
	ExcelXlTimePeriodsDateIsLastWeek = '\002\325\000\004',
	ExcelXlTimePeriodsDateIsLastMonth = '\002\325\000\005',
	ExcelXlTimePeriodsDateIsTomorrow = '\002\325\000\006',
	ExcelXlTimePeriodsDateIsNextWeek = '\002\325\000\007',
	ExcelXlTimePeriodsDateIsNextMonth = '\002\325\000\010',
	ExcelXlTimePeriodsDateIsThisMonth = '\002\325\000\011'
};
typedef enum ExcelXlTimePeriods ExcelXlTimePeriods;

enum ExcelXlDataBarNegativeColorType {
	ExcelXlDataBarNegativeColorTypeDatabarColorTypeColor = '\002\326\000\000',
	ExcelXlDataBarNegativeColorTypeDatabarColorTypeSameAsPositive = '\002\326\000\001'
};
typedef enum ExcelXlDataBarNegativeColorType ExcelXlDataBarNegativeColorType;

enum ExcelLargeScroll {
	ExcelLargeScrollWindow = 'cwin',
	ExcelLargeScrollPane = 'X189'
};
typedef enum ExcelLargeScroll ExcelLargeScroll;

enum ExcelPrintOut {
	ExcelPrintOutWindow = 'cwin',
	ExcelPrintOutSheet = 'X128',
	ExcelPrintOutWorkbook = 'X141'
};
typedef enum ExcelPrintOut ExcelPrintOut;

enum ExcelPrintPreview {
	ExcelPrintPreviewWindow = 'cwin',
	ExcelPrintPreviewSheet = 'X128',
	ExcelPrintPreviewWorkbook = 'X141'
};
typedef enum ExcelPrintPreview ExcelPrintPreview;

enum ExcelSmallScroll {
	ExcelSmallScrollWindow = 'cwin',
	ExcelSmallScrollPane = 'X189'
};
typedef enum ExcelSmallScroll ExcelSmallScroll;

enum ExcelCalculate {
	ExcelCalculateApplication = 'capp',
	ExcelCalculateSheet = 'X128'
};
typedef enum ExcelCalculate ExcelCalculate;

enum ExcelCheckSpelling {
	ExcelCheckSpellingSheet = 'X128',
	ExcelCheckSpellingButton = 'Xbtn',
	ExcelCheckSpellingCheckbox = 'Xckb',
	ExcelCheckSpellingOptionButton = 'XObn',
	ExcelCheckSpellingGroupbox = 'XGBc',
	ExcelCheckSpellingLabel = 'Xlbl',
	ExcelCheckSpellingTextbox = 'XTbx'
};
typedef enum ExcelCheckSpelling ExcelCheckSpelling;

enum ExcelCopyPicture {
	ExcelCopyPictureButton = 'Xbtn',
	ExcelCopyPictureCheckbox = 'Xckb',
	ExcelCopyPictureOptionButton = 'XObn',
	ExcelCopyPictureScrollbar = 'XSrl',
	ExcelCopyPictureListbox = 'XLbx',
	ExcelCopyPictureGroupbox = 'XGBc',
	ExcelCopyPictureDropdown = 'XdpD',
	ExcelCopyPictureSpinner = 'XSpn',
	ExcelCopyPictureLabel = 'Xlbl',
	ExcelCopyPictureTextbox = 'XTbx'
};
typedef enum ExcelCopyPicture ExcelCopyPicture;

enum ExcelCut {
	ExcelCutButton = 'Xbtn',
	ExcelCutCheckbox = 'Xckb',
	ExcelCutOptionButton = 'XObn',
	ExcelCutScrollbar = 'XSrl',
	ExcelCutListbox = 'XLbx',
	ExcelCutGroupbox = 'XGBc',
	ExcelCutDropdown = 'XdpD',
	ExcelCutSpinner = 'XSpn',
	ExcelCutLabel = 'Xlbl',
	ExcelCutTextbox = 'XTbx'
};
typedef enum ExcelCut ExcelCut;

enum ExcelShow {
	ExcelShowDialog = 'X165',
	ExcelShowScenario = 'X191'
};
typedef enum ExcelShow ExcelShow;

enum ExcelUnprotect {
	ExcelUnprotectSheet = 'X128',
	ExcelUnprotectWorkbook = 'X141'
};
typedef enum ExcelUnprotect ExcelUnprotect;

enum ExcelBringToFront {
	ExcelBringToFrontButton = 'Xbtn',
	ExcelBringToFrontCheckbox = 'Xckb',
	ExcelBringToFrontOptionButton = 'XObn',
	ExcelBringToFrontScrollbar = 'XSrl',
	ExcelBringToFrontListbox = 'XLbx',
	ExcelBringToFrontGroupbox = 'XGBc',
	ExcelBringToFrontDropdown = 'XdpD',
	ExcelBringToFrontSpinner = 'XSpn',
	ExcelBringToFrontLabel = 'Xlbl',
	ExcelBringToFrontTextbox = 'XTbx'
};
typedef enum ExcelBringToFront ExcelBringToFront;

enum ExcelSendToBack {
	ExcelSendToBackButton = 'Xbtn',
	ExcelSendToBackCheckbox = 'Xckb',
	ExcelSendToBackOptionButton = 'XObn',
	ExcelSendToBackScrollbar = 'XSrl',
	ExcelSendToBackListbox = 'XLbx',
	ExcelSendToBackGroupbox = 'XGBc',
	ExcelSendToBackDropdown = 'XdpD',
	ExcelSendToBackSpinner = 'XSpn',
	ExcelSendToBackLabel = 'Xlbl',
	ExcelSendToBackTextbox = 'XTbx'
};
typedef enum ExcelSendToBack ExcelSendToBack;

enum ExcelSetFirstPriority {
	ExcelSetFirstPriorityFormatCondition = 'X227',
	ExcelSetFirstPriorityColorScaleFormatCondition = 'X325',
	ExcelSetFirstPriorityDatabarFormatCondition = 'X312',
	ExcelSetFirstPriorityIconSetFormatCondition = 'X315',
	ExcelSetFirstPriorityTop10FormatCondition = 'X321',
	ExcelSetFirstPriorityAboveAverageFormatCondition = 'X322',
	ExcelSetFirstPriorityUniqueValuesFormatCondition = 'X323'
};
typedef enum ExcelSetFirstPriority ExcelSetFirstPriority;

enum ExcelSetLastPriority {
	ExcelSetLastPriorityFormatCondition = 'X227',
	ExcelSetLastPriorityColorScaleFormatCondition = 'X325',
	ExcelSetLastPriorityDatabarFormatCondition = 'X312',
	ExcelSetLastPriorityIconSetFormatCondition = 'X315',
	ExcelSetLastPriorityTop10FormatCondition = 'X321',
	ExcelSetLastPriorityAboveAverageFormatCondition = 'X322',
	ExcelSetLastPriorityUniqueValuesFormatCondition = 'X323'
};
typedef enum ExcelSetLastPriority ExcelSetLastPriority;

enum ExcelModifyAppliesToRange {
	ExcelModifyAppliesToRangeFormatCondition = 'X227',
	ExcelModifyAppliesToRangeColorScaleFormatCondition = 'X325',
	ExcelModifyAppliesToRangeDatabarFormatCondition = 'X312',
	ExcelModifyAppliesToRangeIconSetFormatCondition = 'X315',
	ExcelModifyAppliesToRangeTop10FormatCondition = 'X321',
	ExcelModifyAppliesToRangeAboveAverageFormatCondition = 'X322',
	ExcelModifyAppliesToRangeUniqueValuesFormatCondition = 'X323'
};
typedef enum ExcelModifyAppliesToRange ExcelModifyAppliesToRange;

enum ExcelClearAllFilters {
	ExcelClearAllFiltersPivotTable = 'X155',
	ExcelClearAllFiltersPivotField = 'X157'
};
typedef enum ExcelClearAllFilters ExcelClearAllFilters;

enum ExcelDelete {
	ExcelDeleteCubeField = 'X900',
	ExcelDeleteCalculatedMember = 'X901',
	ExcelDeletePivotFilter = 'X903',
	ExcelDeleteValueChange = 'X905'
};
typedef enum ExcelDelete ExcelDelete;

enum ExcelDrillTo {
	ExcelDrillToPivotField = 'X157',
	ExcelDrillToPivotItem = 'X160'
};
typedef enum ExcelDrillTo ExcelDrillTo;

enum ExcelClearManualFilter {
	ExcelClearManualFilterCubeField = 'X900',
	ExcelClearManualFilterPivotField = 'X157'
};
typedef enum ExcelClearManualFilter ExcelClearManualFilter;

enum ExcelResetTimer {
	ExcelResetTimerPivotCache = 'X151',
	ExcelResetTimerQueryTable = 'X231'
};
typedef enum ExcelResetTimer ExcelResetTimer;

enum ExcelAddItemToList {
	ExcelAddItemToListListbox = 'XLbx',
	ExcelAddItemToListDropdown = 'XdpD'
};
typedef enum ExcelAddItemToList ExcelAddItemToList;

enum ExcelGetBorder {
	ExcelGetBorderFormatCondition = 'X227',
	ExcelGetBorderDisplayFormat = 'X306',
	ExcelGetBorderTop10FormatCondition = 'X321',
	ExcelGetBorderAboveAverageFormatCondition = 'X322',
	ExcelGetBorderUniqueValuesFormatCondition = 'X323'
};
typedef enum ExcelGetBorder ExcelGetBorder;

enum ExcelSetListItem {
	ExcelSetListItemListbox = 'XLbx',
	ExcelSetListItemDropdown = 'XdpD'
};
typedef enum ExcelSetListItem ExcelSetListItem;

enum ExcelCopyObject {
	ExcelCopyObjectButton = 'Xbtn',
	ExcelCopyObjectCheckbox = 'Xckb',
	ExcelCopyObjectOptionButton = 'XObn',
	ExcelCopyObjectScrollbar = 'XSrl',
	ExcelCopyObjectListbox = 'XLbx',
	ExcelCopyObjectGroupbox = 'XGBc',
	ExcelCopyObjectDropdown = 'XdpD',
	ExcelCopyObjectSpinner = 'XSpn',
	ExcelCopyObjectLabel = 'Xlbl',
	ExcelCopyObjectTextbox = 'XTbx'
};
typedef enum ExcelCopyObject ExcelCopyObject;

enum ExcelGetListItem {
	ExcelGetListItemListbox = 'XLbx',
	ExcelGetListItemDropdown = 'XdpD'
};
typedef enum ExcelGetListItem ExcelGetListItem;

enum ExcelRemoveAllItems {
	ExcelRemoveAllItemsListbox = 'XLbx',
	ExcelRemoveAllItemsDropdown = 'XdpD'
};
typedef enum ExcelRemoveAllItems ExcelRemoveAllItems;

enum ExcelRemoveItem {
	ExcelRemoveItemListbox = 'XLbx',
	ExcelRemoveItemDropdown = 'XdpD'
};
typedef enum ExcelRemoveItem ExcelRemoveItem;

enum ExcelActivateObject {
	ExcelActivateObjectWindow = 'cwin',
	ExcelActivateObjectSheet = 'X128',
	ExcelActivateObjectWorkbook = 'X141',
	ExcelActivateObjectPane = 'X189'
};
typedef enum ExcelActivateObject ExcelActivateObject;

enum ExcelXlOartVerticalOverflow {
	ExcelXlOartVerticalOverflowOverflow = '\002\302\000\000',
	ExcelXlOartVerticalOverflowClip = '\002\302\000\001',
	ExcelXlOartVerticalOverflowEllipsis = '\002\302\000\002'
};
typedef enum ExcelXlOartVerticalOverflow ExcelXlOartVerticalOverflow;

enum ExcelXlOartHorizontalOverflow {
	ExcelXlOartHorizontalOverflowOverflow = '\002\303\000\000',
	ExcelXlOartHorizontalOverflowClip = '\002\303\000\001'
};
typedef enum ExcelXlOartHorizontalOverflow ExcelXlOartHorizontalOverflow;

enum ExcelCustomDrop {
	ExcelCustomDropCalloutFormat = 'X101',
	ExcelCustomDropCallout = 'cD00'
};
typedef enum ExcelCustomDrop ExcelCustomDrop;

enum ExcelCustomLength {
	ExcelCustomLengthCalloutFormat = 'X101',
	ExcelCustomLengthCallout = 'cD00'
};
typedef enum ExcelCustomLength ExcelCustomLength;

enum ExcelPresetDrop {
	ExcelPresetDropCalloutFormat = 'X101',
	ExcelPresetDropCallout = 'cD00'
};
typedef enum ExcelPresetDrop ExcelPresetDrop;

enum ExcelChartPatterned {
	ExcelChartPatternedChartFillFormat = 'X253',
	ExcelChartPatternedChartTitle = 'X256',
	ExcelChartPatternedAxisTitle = 'X257',
	ExcelChartPatternedSeriesPoint = 'X262',
	ExcelChartPatternedSeries = 'X263',
	ExcelChartPatternedDataLabel = 'X265',
	ExcelChartPatternedLegendKey = 'X269',
	ExcelChartPatternedDownBars = 'X279',
	ExcelChartPatternedFloor = 'X280',
	ExcelChartPatternedWalls = 'X281',
	ExcelChartPatternedPlotArea = 'X283',
	ExcelChartPatternedChartArea = 'X284',
	ExcelChartPatternedLegend = 'X285',
	ExcelChartPatternedDisplayUnitLabel = 'X299'
};
typedef enum ExcelChartPatterned ExcelChartPatterned;

enum ExcelClear {
	ExcelClearChartArea = 'X284',
	ExcelClearLegend = 'X285'
};
typedef enum ExcelClear ExcelClear;

enum ExcelClearFormats {
	ExcelClearFormatsSeriesPoint = 'X262',
	ExcelClearFormatsSeries = 'X263',
	ExcelClearFormatsLegendKey = 'X269',
	ExcelClearFormatsTrendline = 'X271',
	ExcelClearFormatsFloor = 'X280',
	ExcelClearFormatsWalls = 'X281',
	ExcelClearFormatsPlotArea = 'X283',
	ExcelClearFormatsChartArea = 'X284',
	ExcelClearFormatsErrorBars = 'X286'
};
typedef enum ExcelClearFormats ExcelClearFormats;

enum ExcelApplyDataLabels {
	ExcelApplyDataLabelsChart = 'X119',
	ExcelApplyDataLabelsSeriesPoint = 'X262',
	ExcelApplyDataLabelsSeries = 'X263'
};
typedef enum ExcelApplyDataLabels ExcelApplyDataLabels;

enum ExcelPaste {
	ExcelPasteFloor = 'X280',
	ExcelPasteWalls = 'X281'
};
typedef enum ExcelPaste ExcelPaste;

enum ExcelChartOneColorGradient {
	ExcelChartOneColorGradientChartFillFormat = 'X253',
	ExcelChartOneColorGradientChartTitle = 'X256',
	ExcelChartOneColorGradientAxisTitle = 'X257',
	ExcelChartOneColorGradientSeriesPoint = 'X262',
	ExcelChartOneColorGradientSeries = 'X263',
	ExcelChartOneColorGradientDataLabel = 'X265',
	ExcelChartOneColorGradientLegendKey = 'X269',
	ExcelChartOneColorGradientDownBars = 'X279',
	ExcelChartOneColorGradientFloor = 'X280',
	ExcelChartOneColorGradientWalls = 'X281',
	ExcelChartOneColorGradientPlotArea = 'X283',
	ExcelChartOneColorGradientChartArea = 'X284',
	ExcelChartOneColorGradientLegend = 'X285',
	ExcelChartOneColorGradientDisplayUnitLabel = 'X299'
};
typedef enum ExcelChartOneColorGradient ExcelChartOneColorGradient;

enum ExcelChartTwoColorGradient {
	ExcelChartTwoColorGradientChartFillFormat = 'X253',
	ExcelChartTwoColorGradientChartTitle = 'X256',
	ExcelChartTwoColorGradientAxisTitle = 'X257',
	ExcelChartTwoColorGradientSeriesPoint = 'X262',
	ExcelChartTwoColorGradientSeries = 'X263',
	ExcelChartTwoColorGradientDataLabel = 'X265',
	ExcelChartTwoColorGradientLegendKey = 'X269',
	ExcelChartTwoColorGradientDownBars = 'X279',
	ExcelChartTwoColorGradientFloor = 'X280',
	ExcelChartTwoColorGradientWalls = 'X281',
	ExcelChartTwoColorGradientPlotArea = 'X283',
	ExcelChartTwoColorGradientChartArea = 'X284',
	ExcelChartTwoColorGradientLegend = 'X285',
	ExcelChartTwoColorGradientDisplayUnitLabel = 'X299'
};
typedef enum ExcelChartTwoColorGradient ExcelChartTwoColorGradient;

enum ExcelChartSolid {
	ExcelChartSolidChartFillFormat = 'X253',
	ExcelChartSolidChartTitle = 'X256',
	ExcelChartSolidAxisTitle = 'X257',
	ExcelChartSolidSeriesPoint = 'X262',
	ExcelChartSolidSeries = 'X263',
	ExcelChartSolidDataLabel = 'X265',
	ExcelChartSolidLegendKey = 'X269',
	ExcelChartSolidDownBars = 'X279',
	ExcelChartSolidFloor = 'X280',
	ExcelChartSolidWalls = 'X281',
	ExcelChartSolidPlotArea = 'X283',
	ExcelChartSolidChartArea = 'X284',
	ExcelChartSolidLegend = 'X285',
	ExcelChartSolidDisplayUnitLabel = 'X299'
};
typedef enum ExcelChartSolid ExcelChartSolid;

enum ExcelChartUserPicture {
	ExcelChartUserPictureChartFillFormat = 'X253',
	ExcelChartUserPictureChartTitle = 'X256',
	ExcelChartUserPictureAxisTitle = 'X257',
	ExcelChartUserPictureSeriesPoint = 'X262',
	ExcelChartUserPictureSeries = 'X263',
	ExcelChartUserPictureDataLabel = 'X265',
	ExcelChartUserPictureLegendKey = 'X269',
	ExcelChartUserPictureDownBars = 'X279',
	ExcelChartUserPictureFloor = 'X280',
	ExcelChartUserPictureWalls = 'X281',
	ExcelChartUserPicturePlotArea = 'X283',
	ExcelChartUserPictureChartArea = 'X284',
	ExcelChartUserPictureLegend = 'X285',
	ExcelChartUserPictureDisplayUnitLabel = 'X299'
};
typedef enum ExcelChartUserPicture ExcelChartUserPicture;

enum ExcelDeleteObject {
	ExcelDeleteObjectChartObject = 'X221',
	ExcelDeleteObjectAxis = 'X255'
};
typedef enum ExcelDeleteObject ExcelDeleteObject;

enum ExcelPresetChartGradient {
	ExcelPresetChartGradientChartFillFormat = 'X253',
	ExcelPresetChartGradientChartTitle = 'X256',
	ExcelPresetChartGradientAxisTitle = 'X257',
	ExcelPresetChartGradientSeriesPoint = 'X262',
	ExcelPresetChartGradientSeries = 'X263',
	ExcelPresetChartGradientDataLabel = 'X265',
	ExcelPresetChartGradientLegendKey = 'X269',
	ExcelPresetChartGradientDownBars = 'X279',
	ExcelPresetChartGradientFloor = 'X280',
	ExcelPresetChartGradientWalls = 'X281',
	ExcelPresetChartGradientPlotArea = 'X283',
	ExcelPresetChartGradientChartArea = 'X284',
	ExcelPresetChartGradientLegend = 'X285',
	ExcelPresetChartGradientDisplayUnitLabel = 'X299'
};
typedef enum ExcelPresetChartGradient ExcelPresetChartGradient;

enum ExcelPresetChartTextured {
	ExcelPresetChartTexturedChartFillFormat = 'X253',
	ExcelPresetChartTexturedChartTitle = 'X256',
	ExcelPresetChartTexturedAxisTitle = 'X257',
	ExcelPresetChartTexturedSeriesPoint = 'X262',
	ExcelPresetChartTexturedSeries = 'X263',
	ExcelPresetChartTexturedDataLabel = 'X265',
	ExcelPresetChartTexturedLegendKey = 'X269',
	ExcelPresetChartTexturedDownBars = 'X279',
	ExcelPresetChartTexturedFloor = 'X280',
	ExcelPresetChartTexturedWalls = 'X281',
	ExcelPresetChartTexturedPlotArea = 'X283',
	ExcelPresetChartTexturedChartArea = 'X284',
	ExcelPresetChartTexturedLegend = 'X285',
	ExcelPresetChartTexturedDisplayUnitLabel = 'X299'
};
typedef enum ExcelPresetChartTextured ExcelPresetChartTextured;

enum ExcelChartUserTextured {
	ExcelChartUserTexturedChartFillFormat = 'X253',
	ExcelChartUserTexturedChartTitle = 'X256',
	ExcelChartUserTexturedAxisTitle = 'X257',
	ExcelChartUserTexturedSeriesPoint = 'X262',
	ExcelChartUserTexturedSeries = 'X263',
	ExcelChartUserTexturedDataLabel = 'X265',
	ExcelChartUserTexturedLegendKey = 'X269',
	ExcelChartUserTexturedDownBars = 'X279',
	ExcelChartUserTexturedFloor = 'X280',
	ExcelChartUserTexturedWalls = 'X281',
	ExcelChartUserTexturedPlotArea = 'X283',
	ExcelChartUserTexturedChartArea = 'X284',
	ExcelChartUserTexturedLegend = 'X285',
	ExcelChartUserTexturedDisplayUnitLabel = 'X299'
};
typedef enum ExcelChartUserTextured ExcelChartUserTextured;

@protocol ExcelGenericMethods

- (void) closeSaving:(ExcelSaveOptions)saving savingIn:(NSURL *)savingIn;  // Close a document.
- (void) saveIn:(NSURL *)in_ as:(id)as;  // Save a document.
- (void) printWithProperties:(NSDictionary *)withProperties printDialog:(BOOL)printDialog;  // Print a document.
- (void) delete;  // Delete an object.
- (void) duplicateTo:(SBObject *)to withProperties:(NSDictionary *)withProperties;  // Copy an object.
- (void) moveTo:(SBObject *)to;  // Move an object to a new location.
- (BOOL) RunInProcTestsCommand:(NSString *)command;  // Run InProc Tests.
- (void) applyFilter;  // Apply the defined filter
- (void) applySort;  // Sorts the range based on the currently applied sort states.
- (BOOL) canCheckOutFileName:(NSString *)fileName;  // Returns True if Excel can check out a specified workbook from a server.
- (void) checkOutFileName:(NSString *)fileName;  // Copies a specified workbook from a server to a local computer for editing. Returns a String that represents the local path and file name of the workbook checked out.
- (void) clearColorstops;  // Clears the represented object.
- (void) clearSortfieldset;  // Clears all the sortfield objects.
- (void) deleteSortfield;  // Removes the specified sortfield object from the sortfieldset collection.
- (void) openDataBaseFilename:(NSString *)filename commandText:(NSString *)commandText rcommandType:(NSInteger)rcommandType backGroundQuery:(id)backGroundQuery importDataAs:(NSInteger)importDataAs;  // Open a data base
- (void) openXmlFilename:(NSString *)filename styleSheets:(NSInteger)styleSheets loadOption:(NSInteger)loadOption;  // Open an XML file
- (void) showAll;  // Show all the hidden data, but do not exist the filter mode
- (void) removeDuplicates;  // Removes duplicate values from a range of values.

@end



/*
 * Standard Suite
 */

// The application's top-level scripting object.
@interface ExcelApplication : SBApplication

- (SBElementArray<ExcelDocument *> *) documents;
- (SBElementArray<ExcelWindow *> *) windows;

@property (copy, readonly) NSString *name;  // The name of the application.
@property (readonly) BOOL frontmost;  // Is this the active application?
@property (copy, readonly) NSString *version;  // The version number of the application.

- (id) open:(id)x;  // Open a document.
- (void) print:(id)x withProperties:(NSDictionary *)withProperties printDialog:(BOOL)printDialog;  // Print a document.
- (void) quitSaving:(ExcelSaveOptions)saving;  // Quit the application.
- (BOOL) exists:(id)x;  // Verify that an object exists.
- (void) reset:(id)x;  // Resets a built-in command bar or command bar control to its default configuration.
- (void) ExcelRepeat;  // Repeats the last user-interface action.
- (void) activateObject:(id)x;  // Make the object active
- (void) addChartAutoformatChart:(ExcelChart *)chart name:(NSString *)name objectDescription:(NSString *)objectDescription;  // Adds a custom chart autoformat to the list of available chart autoformats.
- (void) addCustomListListArray:(ExcelXLCustomListType)listArray byRow:(BOOL)byRow;  // Adds a custom list for custom autofill and/or custom sort.
- (void) addItemToList:(id)x itemText:(NSString *)itemText entry_index:(NSInteger)entry_index;  // Adds a string to the list control
- (void) arrange_windowsArrangeStyle:(ExcelXlArrangeStyle)arrangeStyle activeWorkbook:(BOOL)activeWorkbook syncHorizontal:(BOOL)syncHorizontal syncVertical:(BOOL)syncVertical;  // Arranges the windows on the screen
- (void) bringToFront:(id)x;  // Bring the object to the front of the z-order of objects.
- (void) calculate:(id)x;  // Calculates all open workbooks, a specific worksheet in a workbook, or a specified range of cells on a worksheet..
- (void) calculateFull;  // Forces a full calculation of the data in all open workbooks.
- (void) calculateFullRebuild;  // For all open workbooks, forces a full calculation of the data and rebuilds the dependencies
- (double) centimetersToPointsCentimeters:(double)centimeters;  // Converts a measurement from centimeters to points, one point equals 0.035 centimeters.
- (void) checkSpelling:(id)x customDictionary:(NSString *)customDictionary ignoreUppercase:(BOOL)ignoreUppercase alwaysSuggest:(BOOL)alwaysSuggest;  // Checks the spelling of an object.
- (BOOL) checkSpellingForTextToCheck:(NSString *)textToCheck customDictionary:(NSString *)customDictionary ignoreUppercase:(BOOL)ignoreUppercase;  // Checks the spelling of a single word. Returns True if the word is found in one of the dictionaries.
- (void) clearAllFilters:(id)x;  // The ClearAllFilters method deletes all filters currently applied to the PivotTable.
- (void) clearManualFilter:(id)x;  // The ClearManualFilter method provides an easy way to set the Visible property to True for all items of a PivotField in PivotTables, and to empty the HiddenItemsList or VisibleItemsList collections in OLAP PivotTables.
- (NSString *) convertFormulaFormulaToConvert:(NSString *)formulaToConvert fromReferenceStyle:(ExcelXlReferenceStyle)fromReferenceStyle toReferenceStyle:(ExcelXlReferenceStyle)toReferenceStyle toAbsolute:(ExcelXlReferenceStyle)toAbsolute relativeTo:(ExcelXLRangeReference)relativeTo;  // Converts cell references in a formula between the A1 and R1C1 reference styles, between relative and absolute references, or both
- (void) copyObject:(id)x NS_RETURNS_NOT_RETAINED;  // Copies the object to the clipboard.
- (void) copyPicture:(id)x appearance:(ExcelXlPictureAppearance)appearance format:(ExcelXlCopyPictureFormat)format NS_RETURNS_NOT_RETAINED;  // Copies the selected object to the clipboard as a picture.
- (void) cut:(id)x;  // Cuts the object to the clipboard.
- (void) delete:(id)x;  // Deletes the object.
- (void) deleteChartAutoformatName:(NSString *)name;  // Removes a custom chart autoformat from the list of available chart autoformats.
- (void) deleteCustomListListNum:(NSInteger)listNum;  // Deletes a custom list.
- (void) doubleClick;  // Equivalent to double-clicking the active cell.
- (void) drillTo:(id)x field:(NSString *)field;  // The DrillTo method supports drilling to a specified PivotField from another PivotField.
- (id) evaluateName:(id)name;  // Converts a Microsoft Excel name to an object or value.
- (ExcelBorder *) getBorder:(id)x whichBorder:(ExcelXlBordersIndex)whichBorder;  // Returns the specified border object.
- (NSArray<NSAppleEventDescriptor *> *) getClipboardFormats;  // Returns a list of the  formats that are currently on the clipboard.
- (NSArray<id> *) getCustomListContentsListNum:(NSInteger)listNum;  // Returns a custom list of strings.
- (NSInteger) getCustomListNumListArray:(NSArray<NSString *> *)listArray;  // Returns the custom list number for an array of strings. You can use this method to match both built-in lists and custom-defined lists.
- (ExcelDialog *) getDialog:(ExcelXlBuiltInDialog)x;  // Returns a the specified dialog object.
- (NSArray<NSString *> *) getFileConverters;  // Returns a list of all of the file converter objects.
- (id) getInternationalDataType:(ExcelXlApplicationInternational)dataType;  // Returns information about a specific international setting.
- (NSString *) getListItem:(id)x entry_index:(NSInteger)entry_index;  // Returns a string from the list
- (NSString *) getOpenFilenameFileFilter:(NSString *)fileFilter buttonText:(NSString *)buttonText multiSelect:(BOOL)multiSelect;  // Displays the standard Open dialog box and gets a file name from the user without actually opening any files.
- (NSArray<SBObject *> *) getPreviousSelections;  // Returns a list of the last four ranges or names selected. Each element in the list is a range object.
- (NSArray<NSString *> *) getRegisteredFunctions;  // Returns information about functions in code resources that were registered with the REGISTER or REGISTER.ID macro functions.
- (NSString *) getSaveAsFilenameInitialFilename:(NSString *)initialFilename fileFilter:(NSString *)fileFilter filterIndex:(NSInteger)filterIndex buttonText:(NSString *)buttonText;  // Displays the standard save as dialog box and gets a file name from the user without actually saving any files.
- (void) gotoReference:(ExcelXLRangeReference)reference scroll:(BOOL)scroll;  // Selects any range in any workbook, and activates that workbook if it's not already active.
- (void) helpHelpFile:(NSString *)helpFile helpContextId:(NSInteger)helpContextId;  // Displays the Help window.
- (double) inchesToPointsInches:(double)inches;  // Converts a measurement from inches to points.
- (SBObject *) inputBoxPrompt:(NSString *)prompt title:(NSString *)title default:(ExcelXLInputDefault)default_ leftPosition:(NSInteger)leftPosition top:(NSInteger)top type:(ExcelXLInputType)type;  // Displays a dialog box for user input. Returns the information entered in the dialog box.
- (ExcelRange *) intersectRange1:(ExcelRange *)range1 range2:(ExcelRange *)range2 range3:(ExcelRange *)range3 range4:(ExcelRange *)range4 range5:(ExcelRange *)range5 range6:(ExcelRange *)range6 range7:(ExcelRange *)range7 range8:(ExcelRange *)range8 range9:(ExcelRange *)range9 range10:(ExcelRange *)range10 range11:(ExcelRange *)range11 range12:(ExcelRange *)range12 range13:(ExcelRange *)range13 range14:(ExcelRange *)range14 range15:(ExcelRange *)range15 range16:(ExcelRange *)range16 range17:(ExcelRange *)range17 range18:(ExcelRange *)range18 range19:(ExcelRange *)range19 range20:(ExcelRange *)range20 range21:(ExcelRange *)range21 range22:(ExcelRange *)range22 range23:(ExcelRange *)range23 range24:(ExcelRange *)range24 range25:(ExcelRange *)range25 range26:(ExcelRange *)range26 range27:(ExcelRange *)range27 range28:(ExcelRange *)range28 range29:(ExcelRange *)range29 range30:(ExcelRange *)range30;  // Returns a range object that represents the rectangular intersection of two or more ranges.
- (void) largeScroll:(id)x down:(NSInteger)down up:(NSInteger)up toRight:(NSInteger)toRight toLeft:(NSInteger)toLeft;  // Scrolls the contents of the window by pages.
- (void) modifyAppliesToRange:(id)x range:(ExcelRange *)range;  // Changes the range that this format condition applies to.
- (void) onKeyKey:(NSString *)key commandKeyPressed:(BOOL)commandKeyPressed shiftKeyPressed:(BOOL)shiftKeyPressed optionKeyPressed:(BOOL)optionKeyPressed controlKeyPressed:(BOOL)controlKeyPressed procedure:(ExcelXLOnDataType)procedure;  // Runs a specified procedure when a particular key or key combination is pressed.
- (void) openTextFileFilename:(NSString *)filename origin:(ExcelXlPlatform)origin startRow:(NSInteger)startRow dataType:(ExcelXlTextParsingType)dataType textQualifier:(ExcelXlTextQualifier)textQualifier consecutiveDelimiter:(BOOL)consecutiveDelimiter tab:(BOOL)tab semicolon:(BOOL)semicolon comma:(BOOL)comma space:(BOOL)space useOther:(BOOL)useOther otherChar:(NSString *)otherChar fieldInfo:(NSArray<NSAppleEventDescriptor *> *)fieldInfo decimalSeparator:(NSString *)decimalSeparator thousandsSeparator:(NSString *)thousandsSeparator;  // Loads and parses a text file as a new workbook with a single sheet that contains the parsed text-file data.
- (ExcelWorkbook *) openWorkbookWorkbookFileName:(NSString *)workbookFileName updateLinks:(ExcelMyUDateLinks)updateLinks readOnly:(BOOL)readOnly format:(ExcelMyODelimiter)format password:(NSString *)password writeReservedPassword:(NSString *)writeReservedPassword ignoreReadOnlyRecommended:(BOOL)ignoreReadOnlyRecommended origin:(ExcelXlPlatform)origin delimiter:(NSString *)delimiter editable:(BOOL)editable notify:(BOOL)notify converter:(NSInteger)converter addToMru:(BOOL)addToMru;  // Opens a workbook.
- (void) printOut:(id)x from:(NSInteger)from to:(NSInteger)to copies:(NSInteger)copies preview:(BOOL)preview activePrinter:(NSString *)activePrinter printToFile:(BOOL)printToFile collate:(BOOL)collate;  // Prints the object
- (void) printPreview:(id)x enableChanges:(BOOL)enableChanges;  // Shows a preview of the object as it would look when printed. This function has been deprecated.
- (BOOL) registerXllFilename:(NSString *)filename;  // Loads an XLL code resource and automatically registers the functions and commands contained in the resource.
- (void) removeAllItems:(id)x;  // Removes all of the strings from the list
- (void) removeItem:(id)x entry_index:(NSInteger)entry_index count:(NSInteger)count;  // Removes a specified string from the list
- (void) resetTimer:(id)x;  // Resets the refresh timer for the specified PivotTable report to the last interval you set using the RefreshPeriod property.
- (NSString *) runVBMacro:(id)x arg1:(id)arg1 arg2:(id)arg2 arg3:(id)arg3 arg4:(id)arg4 arg5:(id)arg5 arg6:(id)arg6 arg7:(id)arg7 arg8:(id)arg8 arg9:(id)arg9 arg10:(id)arg10 arg11:(id)arg11 arg12:(id)arg12 arg13:(id)arg13 arg14:(id)arg14 arg15:(id)arg15 arg16:(id)arg16 arg17:(id)arg17 arg18:(id)arg18 arg19:(id)arg19 arg20:(id)arg20 arg21:(id)arg21 arg22:(id)arg22 arg23:(id)arg23 arg24:(id)arg24 arg25:(id)arg25 arg26:(id)arg26 arg27:(id)arg27 arg28:(id)arg28 arg29:(id)arg29 arg30:(id)arg30;  // Runs a macro or calls a function. This can be used to run a macro written in the Microsoft Excel 4.0 macro language, or to run a function in a DLL or XLL.
- (NSString *) runXLMMacro:(id)x arg1:(id)arg1 arg2:(id)arg2 arg3:(id)arg3 arg4:(id)arg4 arg5:(id)arg5 arg6:(id)arg6 arg7:(id)arg7 arg8:(id)arg8 arg9:(id)arg9 arg10:(id)arg10 arg11:(id)arg11 arg12:(id)arg12 arg13:(id)arg13 arg14:(id)arg14 arg15:(id)arg15 arg16:(id)arg16 arg17:(id)arg17 arg18:(id)arg18 arg19:(id)arg19 arg20:(id)arg20 arg21:(id)arg21 arg22:(id)arg22 arg23:(id)arg23 arg24:(id)arg24 arg25:(id)arg25 arg26:(id)arg26 arg27:(id)arg27 arg28:(id)arg28 arg29:(id)arg29 arg30:(id)arg30;  // Runs a macro or calls a function. This can be used to run a macro written in the Microsoft Excel 4.0 macro language, or to run a function in a DLL or XLL.
- (void) saveWorkspaceWorkspaceFileName:(NSString *)workspaceFileName;  // Saves the current workspace.
- (void) sendToBack:(id)x;  // Sends the object to the back of the z-order.
- (void) setFirstPriority:(id)x;  // Sets this condition to the highest priority.
- (void) setLastPriority:(id)x;  // Sets this condition to the lowest priority.
- (void) setListItem:(id)x entry_index:(NSInteger)entry_index itemText:(NSString *)itemText;  // Set a string in the list
- (void) show:(id)x;  // Displays the built-in dialog box and waits for the user to input data.
- (void) smallScroll:(id)x down:(NSInteger)down up:(NSInteger)up toRight:(NSInteger)toRight toLeft:(NSInteger)toLeft;  // Scrolls the contents of the window by rows or columns.
- (void) undo;  // Cancels the last user-interface action.
- (ExcelRange *) unionRange1:(ExcelRange *)range1 range2:(ExcelRange *)range2 range3:(ExcelRange *)range3 range4:(ExcelRange *)range4 range5:(ExcelRange *)range5 range6:(ExcelRange *)range6 range7:(ExcelRange *)range7 range8:(ExcelRange *)range8 range9:(ExcelRange *)range9 range10:(ExcelRange *)range10 range11:(ExcelRange *)range11 range12:(ExcelRange *)range12 range13:(ExcelRange *)range13 range14:(ExcelRange *)range14 range15:(ExcelRange *)range15 range16:(ExcelRange *)range16 range17:(ExcelRange *)range17 range18:(ExcelRange *)range18 range19:(ExcelRange *)range19 range20:(ExcelRange *)range20 range21:(ExcelRange *)range21 range22:(ExcelRange *)range22 range23:(ExcelRange *)range23 range24:(ExcelRange *)range24 range25:(ExcelRange *)range25 range26:(ExcelRange *)range26 range27:(ExcelRange *)range27 range28:(ExcelRange *)range28 range29:(ExcelRange *)range29 range30:(ExcelRange *)range30;  // Returns the union of two or more ranges.
- (void) unprotect:(id)x password:(NSString *)password;  // Removes protection from a sheet or workbook. This method has no effect if the sheet or workbook isn't protected.
- (void) bringToFront:(id)x;  // Bring the object to the front of the z-order of objects.
- (void) checkSpelling:(id)x customDictionary:(NSString *)customDictionary ignoreUppercase:(BOOL)ignoreUppercase alwaysSuggest:(BOOL)alwaysSuggest;  // Checks the spelling of an object.
- (void) copyObject:(id)x NS_RETURNS_NOT_RETAINED;  // Copies the object to the clipboard.
- (void) copyPicture:(id)x appearance:(ExcelXlPictureAppearance)appearance format:(ExcelXlCopyPictureFormat)format NS_RETURNS_NOT_RETAINED;  // Copies the selected object to the clipboard as a picture.
- (void) customDrop:(id)x drop:(double)drop;  // Sets the vertical distance in points from the edge of the text bounding box to the place where the callout line attaches to the text box.
- (void) customLength:(id)x length:(double)length;  // Specifies that the first segment of the callout line, the segment attached to the text callout box, retain a fixed length whenever the callout is moved.
- (void) cut:(id)x;  // Cuts the object to the clipboard.
- (void) presetDrop:(id)x dropType:(ExcelMsoCalloutDropType)dropType;  // Specifies whether the callout line attaches to the top, bottom, or center of the callout text box or whether it attaches at a point that's a specified distance from the top or bottom of the text box.
- (void) sendToBack:(id)x;  // Sends the object to the back of the z-order.
- (void) clearHyperlinks;  // clear all the hyperlinks in a range
- (void) dirty;  // Designates a range to be recalculated when the next recalculation occurs.
- (void) activateObject:(id)x;  // Activates the object.
- (void) applyDataLabels:(id)x type:(ExcelXlDataLabelsType)type legendKey:(BOOL)legendKey autoText:(BOOL)autoText hasLeaderLines:(BOOL)hasLeaderLines showSeriesName:(BOOL)showSeriesName showCategoryName:(BOOL)showCategoryName showValue:(BOOL)showValue showPercentage:(BOOL)showPercentage showBubbleSize:(BOOL)showBubbleSize separator:(NSString *)separator;  // Applies data labels to all the series in a chart, a single series or a series point.
- (void) chartOneColorGradient:(id)x gradientStyle:(ExcelMsoGradientStyle)gradientStyle variant:(NSInteger)variant degree:(double)degree;  // Sets the specified fill to a one-color gradient.
- (void) chartPatterned:(id)x pattern:(ExcelMsoPatternType)pattern;  // Sets the specified fill to a pattern.
- (void) chartSolid:(id)x;  // Sets the specified fill to a uniform color. Use this method to convert a gradient, textured, patterned, or background fill back to a solid fill.
- (void) chartTwoColorGradient:(id)x gradientStyle:(ExcelMsoGradientStyle)gradientStyle variant:(NSInteger)variant;  // Sets the specified fill to a two-color gradient.
- (void) chartUserPicture:(id)x pictureFile:(NSString *)pictureFile pictureFormat:(ExcelXlChartPictureType)pictureFormat pictureStackUnit:(NSInteger)pictureStackUnit picturePlacement:(ExcelXlChartPicturePlacement)picturePlacement;  // Fills the specified shape with an image.
- (void) chartUserTextured:(id)x textureFile:(NSString *)textureFile;  // Fills the specified shape with small tiles of an image.
- (void) clear:(id)x;  // Clear the object.
- (void) clearFormats:(id)x;  // Clears the formatting of the object.
- (void) copyObject:(id)x NS_RETURNS_NOT_RETAINED;  // Copies the object to the clipboard.
- (void) deleteObject:(id)x;  // Deletes the object.
- (void) paste:(id)x;
- (void) presetChartGradient:(id)x gradientStyle:(ExcelMsoGradientStyle)gradientStyle variant:(NSInteger)variant presetGradientType:(ExcelMsoPresetGradientType)presetGradientType;  // Sets the specified fill to a preset gradient.
- (void) presetChartTextured:(id)x presetTextureForChart:(ExcelMsoPresetTexture)presetTextureForChart;  // Sets the specified fill format to a preset texture.

@end

// A document.
@interface ExcelDocument : SBObject <ExcelGenericMethods>

@property (copy, readonly) NSString *name;  // Its name.
@property (readonly) BOOL modified;  // Has it been modified since the last save?
@property (copy, readonly) NSURL *file;  // Its location on disk, if it has one.


@end

// A window.
@interface ExcelWindow : SBObject <ExcelGenericMethods>

@property (copy, readonly) NSString *name;  // The title of the window.
- (NSInteger) id;  // The unique identifier of the window.
@property NSInteger index;  // The index of the window, ordered front to back.
@property (copy) ExcelRectangle *bounds;  // The bounding rectangle of the window.
@property (readonly) BOOL closeable;  // Does the window have a close button?
@property (readonly) BOOL miniaturizable;  // Does the window have a minimize button?
@property BOOL miniaturized;  // Is the window minimized right now?
@property (readonly) BOOL resizable;  // Can the window be resized?
@property BOOL visible;  // Is the window visible right now?
@property (readonly) BOOL zoomable;  // Does the window have a zoom button?
@property BOOL zoomed;  // Is the window zoomed right now?
@property (copy, readonly) ExcelDocument *document;  // The document whose contents are displayed in the window.
@property (readonly) NSInteger entryIndex;  // the number of the window
@property NSPoint position;  // upper left coordinates of the window
@property (readonly) BOOL titled;  // Does the window have a title bar?
@property (readonly) BOOL floating;  // Does the window float?
@property (readonly) BOOL modal;  // Is the window modal?
@property (readonly) BOOL collapsable;  // Is the window collapasable?
@property BOOL collapsed;  // Is the window collapsed?
@property (readonly) BOOL sheet;  // Is this window a sheet window?

- (void) activateNext;  // Activates the current window, sends it to the back of the window z-order, then activates the next window according to the z-order.
- (void) activatePrevious;  // Activates the specified window and then activates the window at the back of the window z-order.
- (void) scrollWorkbookTabsSheets:(NSInteger)sheets position:(ExcelXLScrollTabPosition)position;  // Scrolls through the workbook tabs at the bottom of the window. Doesn't affect the active sheet in the workbook.

@end



/*
 * Microsoft Office Suite
 */

// A control within a command bar.
@interface ExcelCommandBarControl : SBObject <ExcelGenericMethods>

@property BOOL beginGroup;  // Returns or sets if the command bar control appears at the beginning of a group of controls on the command bar.
@property (readonly) BOOL builtIn;  // Returns true if the command bar control is a built-in command bar control.
@property (readonly) ExcelMsoControlType controlType;  // Returns the type of the command bar control.
@property (copy) NSString *descriptionText;  // Returns or sets the description for a command bar control.  The description is not displayed to the user, but it can be useful for documenting the behavior of a control.
@property BOOL enabled;  // Returns or sets if the command bar control is enabled.
@property (readonly) NSInteger entry_index;  // Returns the index number for this command bar control.
@property NSInteger height;  // Returns or sets the height of a command bar control.
@property NSInteger helpContextID;  // Returns or sets the help context ID number for the Help topic attached to the command bar control.
@property (copy) NSString *helpFile;  // Returns or sets the file name for the help topic attached to the command bar.  To use this property, you must also set the help context ID property.
- (NSInteger) id;  // Returns the id for a built-in command bar control.
@property (readonly) NSInteger leftPosition;  // Returns the left position of the command bar control.
@property (copy) NSString *name;  // Returns or sets the caption text for a command bar control.
@property (copy) NSString *parameter;  // Returns or sets a string that is used to execute a command.
@property NSInteger priority;  // Returns or sets the priority of a command bar control. A controls priority determines whether the control can be dropped from a docked command bar if the command bar controls can not fit in a single row.  Valid priority number are 0 through 7.
@property (copy) NSString *tag;  // Returns or sets information about the command bar control, such as data that can be used as an argument in procedures, or information that identifies the control.
@property (copy) NSString *tooltipText;  // Returns or sets the text displayed in a command bar controls tooltip.
@property (readonly) NSInteger top;  // Returns the top position of a command bar control.
@property BOOL visible;  // Returns or sets if the command bar control is visible.
@property NSInteger width;  // Returns or sets the width in pixels of the command bar control.

- (void) execute;  // Runs the procedure or built-in command assigned to the specified command bar control.

@end

// A button control within a command bar.
@interface ExcelCommandBarButton : ExcelCommandBarControl

@property (readonly) BOOL buttonFaceIsDefault;  // Returns if the face of a command bar button control is the original built-in face.
@property ExcelMsoButtonState buttonState;  // Returns or set the appearance of a command bar button control.  The property is read-only for built-in command bar buttons.
@property ExcelMsoButtonStyle buttonStyle;  // Returns or sets the way a command button control is displayed.
@property NSInteger faceId;  // Returns or sets the Id number for the face of the command bar button control.


@end

// A combobox menu control within a command bar.
@interface ExcelCommandBarCombobox : ExcelCommandBarControl

@property ExcelMsoComboStyle comboboxStyle;  // Returns or sets the way a command bar combobox control is displayed.
@property (copy) NSString *comboboxText;  // Returns or sets the text in the display or edit portion of the command bar combobox control.
@property NSInteger dropDownLines;  // Returns or sets the number of lines in a command bar control combobox control.  The combobox control must be a custom control.
@property NSInteger dropDownWidth;  // Returns or sets the width in pixels of the list for the specified command bar combobox control.  An error occurs if you attempt to set this property for a built-in combobox control.
@property NSInteger listIndex;

- (void) addItemToComboboxComboboxItem:(NSString *)comboboxItem entry_index:(NSInteger)entry_index;  // Add a new string to a custom combobox control.
- (void) clearCombobox;  // Clear all of the strings form a custom combobox.
- (NSString *) getComboboxItemEntry_index:(NSInteger)entry_index;  // Return the string at the given index within a combobox.
- (NSInteger) getCountOfComboboxItems;  // Return the number of strings within a combobox.
- (void) removeAnItemFromComboboxEntry_index:(NSInteger)entry_index;  // Remove a string from a custom combobox.
- (void) setComboboxItemEntry_index:(NSInteger)entry_index comboboxItem:(NSString *)comboboxItem;  // Set the string an a given index for a custom combobox.

@end

// A popup menu control within a command bar.
@interface ExcelCommandBarPopup : ExcelCommandBarControl

- (SBElementArray<ExcelCommandBarControl *> *) commandBarControls;


@end

// Toolbars used in all of the Office applications.
@interface ExcelCommandBar : SBObject <ExcelGenericMethods>

- (SBElementArray<ExcelCommandBarControl *> *) commandBarControls;

@property ExcelMsoBarPosition barPosition;  // Returns or sets the position of the command bar.
@property (readonly) ExcelMsoBarType barType;  // Returns the type of this command bar.
@property (readonly) BOOL builtIn;  // True if the command bar is built-in.
@property (copy, readonly) NSString *context;  // Returns or sets a string that determines where a command bar will be saved.
@property BOOL enabled;  // Returns or set if the command bar is enabled.
@property (readonly) NSInteger entry_index;  // The index of the command bar.
@property NSInteger height;  // Returns or sets the height of the command bar.
@property NSInteger leftPosition;  // Returns or sets the left position of the command bar.
@property (copy) NSString *localName;  // Returns or sets the name of the command bar in the localized language of the application.  An error is returned when trying to set the local name of a built-in command bar.
@property (copy) NSString *name;  // Returns or sets the name of the command bar.
@property (copy) NSArray<NSAppleEventDescriptor *> *protection;  // Returns or sets the way a command bar is protected from user customization.  It accepts a list of the following items: no protection, no customize, no resize, no move, no change visible, no change dock, no vertical dock, no horizontal dock.
@property NSInteger rowIndex;  // Returns or sets the docking order of a command bar in relation to other command bars in the same docking area.  Can be an integer greater than zero.
@property NSInteger top;  // Returns or sets the top position of a command bar.
@property BOOL visible;  // Returns or sets if the command bar is visible.
@property NSInteger width;  // Returns or sets the width in pixels of the command bar.


@end

@interface ExcelDocumentProperty : SBObject <ExcelGenericMethods>

@property (copy) NSNumber *documentPropertyType;  // Returns or sets the document property type.
@property (copy) NSString *linkSource;  // Returns or sets the source of a lined custom document property.
@property BOOL linkToContent;  // True if the value of the document property is lined to the content of the container document.  False if the value is static.  This only applies to custom document properties.  For built-in properties this is always false.
@property (copy) NSString *name;  // Returns or sets the name of the document property.
@property (copy) NSString *value;  // Returns or sets the value of the document property.


@end

@interface ExcelCustomDocumentProperty : ExcelDocumentProperty


@end

@interface ExcelWebPageFont : SBObject <ExcelGenericMethods>

@property (copy) NSString *fixedWidthFont;  // Returns or sets the fixed-width font setting.
@property double fixedWidthFontSize;  // Returns or sets the fixed-width font size.  You can enter half-point sizes; if you enter other fractional point sizes, they are rounded up or down to the nearest half-point.
@property (copy) NSString *proportionalFont;  // Returns or sets the proportional font setting.
@property double proportionalFontSize;  // Returns or sets the proportional font size.  You can enter half-point sizes; if you enter other fractional point sizes, they are rounded up or down to the nearest half-point.


@end



/*
 * Microsoft Excel Suite
 */

// Represents a cell comment.
@interface ExcelExcelComment : SBObject <ExcelGenericMethods>

@property (copy, readonly) NSString *author;  // Returns the author of the comment.
@property (copy, readonly) ExcelShape *shapeObject;  // Returns a shape object that represents the shape attached to the specified comment.
@property BOOL visible;  // Returns or sets if the specified object is visible.

- (NSString *) ExcelCommentTextText:(NSString *)text start:(NSInteger)start overWrite:(BOOL)overWrite;  // Returns or sets the text of the comment
- (ExcelExcelComment *) nextExcelComment;  // Returns the next Excel comment object
- (ExcelExcelComment *) previousExcelComment;  // Returns the previous Excel comment object

@end

@interface ExcelODBCError : SBObject <ExcelGenericMethods>

@property (copy, readonly) NSString *errorString;  // Returns the ODBC error string.
@property (copy, readonly) NSString *sqlState;  // Returns the SQL state error.


@end

// Represents the various types of protection options available for a worksheet
@interface ExcelProtection : SBObject <ExcelGenericMethods>

@property (readonly) BOOL allowDeletingColumns;  // Returns True if the deleting of columns is allowed on a protected worksheet. Read-only Boolean.
@property (readonly) BOOL allowDeletingRows;  // Returns True if the deleting of rows is allowed on a protected worksheet. Read-only Boolean.
@property (readonly) BOOL allowFiltering;  // Returns True if the filtering is allowed on a protected worksheet. Read-only Boolean.
@property (readonly) BOOL allowFormattingCells;  // Returns True if the formatting of cells is allowed on a protected worksheet. Read-only Boolean.
@property (readonly) BOOL allowFormattingColumns;  // Returns True if the formatting of columns is allowed on a protected worksheet. Read-only Boolean.
@property (readonly) BOOL allowFormattingRows;  // Returns True if the formatting of rows is allowed on a protected worksheet. Read-only Boolean.
@property (readonly) BOOL allowInsertingColumns;  // Returns True if the inserting of columns is allowed on a protected worksheet. Read-only Boolean.
@property (readonly) BOOL allowInsertingHyperlinks;  // Returns True if the inserting of hyperlinks is allowed on a protected worksheet. Read-only Boolean.
@property (readonly) BOOL allowInsertingRows;  // Returns True if the inserting of rows is allowed on a protected worksheet. Read-only Boolean.
@property (readonly) BOOL allowSorting;  // Returns True if the sorting is allowed on a protected worksheet. Read-only Boolean.
@property (readonly) BOOL allowUsingPivotTable;  // Returns True if the using of pivot table is allowed on a protected worksheet. Read-only Boolean.


@end

@interface ExcelAboveAverageFormatCondition : SBObject <ExcelGenericMethods>

@property ExcelXlAboveBelow aboveOrBelow;
@property (copy, readonly) ExcelRange *appliesTo;  // Returns the range this format condition applies to. Read Only.
@property ExcelXlCalcFor calcFor;
@property (copy, readonly) ExcelFont *fontObject;  // Returns a font object that represents the font of the specified object.
@property NSInteger formatConditionPriority;  // Specifies the priority for the format condition. Read/Write.
@property (readonly) ExcelXlFormatConditionType formatConditionType;  // Return the conditional format type.
@property (copy, readonly) ExcelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property (copy) NSString *numberFormat;  // Returns or sets the format code for the object.
@property NSInteger numberOfStandardDeviations;
@property ExcelXlPivotConditionScope pivotConditionScopeType;  // Returns or sets the part of the pivot table that the pivot table format condition is scoped to
@property (readonly) BOOL pivotTableCondition;  // Tells whether this format condition is a pivot table condition.
@property (readonly) BOOL stopIfTrue;  // Tells whether the processing of format conditions stops if this condition is true. Read Only.


@end

// Represents a single add-in, either installed or not installed.
@interface ExcelAddIn : SBObject <ExcelGenericMethods>

@property (copy, readonly) NSString *fullName;  // Returns the name of the add in, including its path on disk, as a string.
@property BOOL installed;  // Returns or sets if the add in is installed.
@property (copy, readonly) NSString *name;  // Returns the name of the object.
@property (copy, readonly) NSString *path;  // Returns the complete path of the object, excluding the final separator and name of the add in.


@end

// A representation of the Microsoft Excel application.
@interface ExcelApplication (MicrosoftExcelSuite)

- (SBElementArray<ExcelAddIn *> *) addIns;
- (SBElementArray<ExcelChartSheet *> *) chartSheets;
- (SBElementArray<ExcelCommandBar *> *) commandBars;
- (SBElementArray<ExcelNamedItem *> *) namedItems;
- (SBElementArray<ExcelRange *> *) ranges;
- (SBElementArray<ExcelCell *> *) cells;
- (SBElementArray<ExcelRow *> *) rows;
- (SBElementArray<ExcelColumn *> *) columns;
- (SBElementArray<ExcelWindow *> *) windows;
- (SBElementArray<ExcelWorkbook *> *) workbooks;
- (SBElementArray<ExcelSheet *> *) sheets;
- (SBElementArray<ExcelWorksheet *> *) worksheets;
- (SBElementArray<ExcelInternationalMacroSheet *> *) internationalMacroSheets;
- (SBElementArray<ExcelMacroSheet *> *) macroSheets;
- (SBElementArray<ExcelRecentFile *> *) recentFiles;
- (SBElementArray<ExcelODBCError *> *) ODBCErrors;

@property BOOL AutoFormatAsYouTypeReplaceHyperlinks;  // True if Microsoft Excel automatically formats hyperlinks as you type. False if Excel does not automatically format hyperlinks as you type.
@property ExcelXlMousePointer ExcelCursor;  // Returns or sets the appearance of the mouse pointer in Microsoft Excel.
@property NSInteger ODBCTimeout;  // Returns or sets the ODBC query time limit, in seconds. The default value is 45 seconds. 
@property (copy, readonly) ExcelCell *activeCell;  // Returns a range object that represents the active cell in the active window, the window on top, or in the specified window. If the window isn't displaying a worksheet, this property fails.
@property (copy, readonly) ExcelChart *activeChart;  // Returns a chart object that represents the active chart, either an embedded chart or a chart sheet. An embedded chart is considered active when it's either selected or activated.
@property (copy) NSString *activePrinter;  // Returns or sets the name of the active printer.
@property (copy, readonly) ExcelSheet *activeSheet;  // Returns an object that represents the active sheet, the sheet on top, in the active workbook or in the specified window or workbook.
@property (copy, readonly) ExcelWindow *activeWindow;  // Returns a window object that represents the active window, the window on top.
@property (copy, readonly) ExcelWorkbook *activeWorkbook;  // Returns a workbook object that represents the workbook in the active window, the window on top.
@property BOOL alertBeforeOverwriting;  // Returns or sets if Microsoft Excel displays a message before overwriting nonblank cells during a drag-and-drop editing operation.
@property (copy) NSString *altStartupPath;  // Returns or sets the name of the alternate startup folder.
@property (readonly) BOOL arbitraryXMLSupportAvailable;  // Returns a Boolean value that indicates whether the XML features in Microsoft Excel are available
@property BOOL askToUpdateLinks;  // Returns or sets if Microsoft Excel asks the user to update links when opening files with links.
@property (copy, readonly) ExcelAutocorrect *autocorrectObject;  // Returns an autocorrect object that represents the Microsoft Excel AutoCorrect attributes.
@property ExcelMsoAutomationSecurity automationSecurity;
@property (readonly) NSInteger build;  // Returns the Microsoft Excel build number.
@property BOOL calculateBeforeSave;  // Returns or sets if workbooks are calculated before they're saved to disk.
@property ExcelXlCalculation calculation;  // Returns or sets the calculation mode.
@property (readonly) NSInteger calculationVersion;  // Returns a number whose rightmost four digits are the minor calculation engine version number, and whose other digits, on the left, are the major version of Microsoft Excel.
@property (copy, readonly) NSString *caption;  // Returns the name of the application.
@property BOOL cellDragAndDrop;  // Returns or sets if dragging and dropping cells is enabled.
@property ExcelXlCommandUnderlines commandUnderlines;  // Returns or sets the state of the command underlines in Microsoft Excel.
@property BOOL copyObjectsWithCells;  // Returns or sets if objects are cut, copied, extracted, and sorted with cells.
@property (readonly) NSInteger customListCount;  // Returns the number of defined custom lists, including built-in lists.
@property ExcelXlCutCopyMode cutCopyMode;  // Returns or sets the status of cut or copy mode.
@property ExcelXLDataEntryMode dataEntryMode;  // Returns or sets data entry mode, as shown in the following table. When in data entry mode, you can enter data only in the unlocked cells in the currently selected range. 
@property (copy) NSString *defaultFilePath;  // Returns or sets the default path that Microsoft Excel uses when it opens files. 
@property ExcelXlFileFormat defaultSaveFormat;  // Returns or sets the default format for saving files.
@property (copy, readonly) ExcelDefaultWebOptions *defaultWebOptionsObject;  // Returns the default web object.
@property BOOL displayAlerts;  // Returns or sets if Microsoft Excel displays certain alerts and messages while handling events from AppleScript.
@property ExcelXlCommentDisplayMode displayCommentIndicator;  // Returns or sets the way cells display comments and indicators.
@property BOOL displayExcel4Menus;  // Returns or sets if Microsoft Excel displays version 4.0 menu bars.
@property BOOL displayFormulaBar;  // Returns or sets  if the formula bar is displayed.
@property BOOL displayFullScreen;  // Returns or sets if Microsoft Excel is in full-screen mode.
@property BOOL displayFunctionTooltips;  // Returns or sets if function tool tips can be displayed.
@property BOOL displayInsertOptions;  // Returns or sets if the insert options button should be displayed. 
@property BOOL displayNoteIndicator;  // Returns or sets if cells containing notes display cell tips and contain note indicators, small dots in their upper-right corners.
@property BOOL displayRecentFiles;  // Returns or sets if the list of recently used files is displayed on the file menu.
@property BOOL displayScrollBars;  // Returns or sets if scroll bars are visible for all workbooks.
@property BOOL displayStatusBar;  // Returns or sets if the status bar is displayed.
@property BOOL editDirectlyInCell;  // Returns or sets if Microsoft Excel allows editing in cells.
@property BOOL enableAnimations;  // Returns or sets if animated insertion and deletion is enabled.
@property BOOL enableAutocomplete;  // Returns or sets if the autocomplete feature is enabled.
@property ExcelXlEnableCancelKey enableCancelKey;  // Controls how Microsoft Excel handles CTRL-BREAK, ESC, or cmd-period user interruptions to the running procedure.
@property BOOL enableEvents;  // Returns or sets if events are enabled for the application object.
@property BOOL enableSound;  // Returns or sets if sound is enabled for Microsoft Office.
@property BOOL extendList;  // Returns or sets if Microsoft Excel automatically extends formatting and formulas to new data that is added to a list.
@property BOOL fixedDecimal;  // All data entered after this property is set to true will be formatted with the number of fixed decimal places set by the fixed decimal places property.
@property NSInteger fixedDecimalPlaces;  // Returns or sets the number of fixed decimal places used when the fixed decimal property is set to true.
@property (copy, readonly) id frontmost;  // Returns if the application is the frontmost window.
@property BOOL includeEmptyCellsInLists;  // Returns or sets if empty cells are included in range lists.
@property BOOL iteration;  // Returns or sets  if Microsoft Excel will use iteration to resolve circular references.
@property (copy, readonly) NSString *libraryPath;  // Returns the path to the Library folder.
@property (readonly) BOOL mathCoprocessorAvailable;  // Returns true if a math coprocessor is available.
@property double maxChange;  // Returns or sets the maximum amount of change between each iteration as Microsoft Excel resolves circular references. 
@property NSInteger maxIterations;  // Returns or sets the maximum number of iterations that Microsoft Excel can use to resolve a circular reference.
@property NSInteger measurementUnit;  // Returns or set the current unit of measure.
@property (readonly) NSInteger memoryFree;  // Returns the amount of memory that's still available for Microsoft Excel to use, in bytes.
@property (readonly) NSInteger memoryTotal;  // Returns the total amount of memory, in bytes, that's available to Microsoft Excel, including memory already in use.
@property (readonly) NSInteger memoryUsed;  // Returns the amount of memory that Microsoft Excel is currently using, in bytes.
@property BOOL moveAfterReturn;  // Returns or sets if the active cell will be moved as soon as the ENTER/RETURN key is pressed.
@property ExcelXlDirection moveAfterReturnDirection;  // Returns or sets the direction in which the active cell is moved when the user presses ENTER.
@property (copy, readonly) NSString *name;  // Returns the name of the application.
@property (copy, readonly) NSString *networkTemplatesPath;  // Returns the network path where templates are stored. If the network path doesn't exist, this property returns an empty string. 
@property (copy, readonly) NSString *operatingSystem;  // Returns the name and version number of the current operating system.
@property (copy, readonly) NSString *organizationName;  // Returns the registered organization name.
@property (copy, readonly) NSString *path;  // Returns the complete path of the application, excluding the final separator and name of the application.
@property (copy, readonly) NSString *pathSeparator;  // Returns the path separator character.
@property BOOL pivotTableSelection;  // Returns or sets if pivot tables use structured selection.
@property BOOL promptForSummaryInfo;  // Returns or sets if Microsoft Excel asks for summary information when files are first saved.
@property ExcelXlReferenceStyle referenceStyle;  // Returns or sets how Microsoft Excel displays cell references and row and column headings in either A1 or R1C1 reference style.
@property BOOL rollZoom;  // Returns or sets if the IntelliMouse zooms instead of scrolling.
@property BOOL screenUpdating;  // Returns or sets if screen updating is turned on. Turn screen updating off to speed up your AppleScript code. You won't be able to see what the code is doing, but it will run faster.
@property (copy, readonly) ExcelRange *selection;  // Returns the selected object in the active window.
@property NSInteger sheetsInNewWorkbook;  // Returns or sets the number of sheets that Microsoft Excel automatically inserts into new workbooks.
@property BOOL showChartTipNames;  // Returns or sets if charts show chart tip names.
@property BOOL showChartTipValues;  // Returns or sets if charts show chart tip values.
@property BOOL showToolTips;  // Returns or sets if tool tips are turned on.
@property (copy, readonly) ExcelXlspellingOptions *spellingOptions;  // Returns the default spelling options object.
@property (copy) NSString *standardFont;  // Returns or sets the name of the standard font.
@property double standardFontSize;  // Returns or sets the standard font size, in points.
@property BOOL startupDialog;  // Returns or sets if the startup dialog should be shown when starting up the application.
@property (copy, readonly) NSString *startupPath;  // Returns the complete path of the startup folder, excluding the final separator.
@property (copy) NSString *statusBar;  // Returns or sets the text in the status bar.
@property (copy, readonly) NSString *templatesPath;  // Returns the local path where templates are stored.
@property (copy, readonly) id thisCell;  // Returns the cell in which the user-defined function is being called from as a Range object.
@property (copy) NSString *transitionMenuKey;  // Returns or sets the alternate menu or help key.
@property ExcelXLTransitionMenuKeyAction transitionMenuKeyAction;  // Returns or sets the action taken when the alternate menu key is pressed.
@property (readonly) double usableHeight;  // Returns the maximum height of the space that a window can occupy in points.
@property (readonly) double usableWidth;  // Returns the maximum width of the space that a window can occupy in points.
@property (copy) NSString *userName;  // Returns or sets the name of the current user.

@end

// Represents autofiltering for the specified worksheet.
@interface ExcelAutofilter : SBObject <ExcelGenericMethods>

- (SBElementArray<ExcelFilter *> *) filters;

@property (readonly) BOOL autofiltermode;  // Returns True if filters have been defined else False.
@property (copy, readonly) ExcelRange *rangeObject;  // Returns a range object that represents the range to which the specified autofilter applies.
@property (copy, readonly) ExcelSort *sortObject;  // Returns the sort object in the auto filter object.


@end

// Represents the border of an object.
@interface ExcelBorder : SBObject <ExcelGenericMethods>

@property (copy) NSArray<NSNumber *> *color;  // Returns or sets the primary color of the object. 
@property ExcelXlColorIndex colorIndex;  // Returns or sets the color of the border. The color is specified as an index value into the current color palette.
@property ExcelXlLineStyle lineStyle;  // Returns or sets the line style for the border.
@property ExcelXlThemeColor themeColor;  // Returns or sets the theme color in the applied color scheme that is associated with the specified object.
@property double tintAndShade;  // Returns or sets a Single that lightens or darkens a color.
@property ExcelXlBorderWeight weight;  // Returns or sets the weight of the border.


@end

// Represents a button control.
@interface ExcelButton : SBObject <ExcelGenericMethods>

- (SBElementArray<ExcelCharacter *> *) characters;

@property (copy) NSString *accelerator;  // Returns or sets the accelerator character for this control.
@property BOOL addIndent;  // Returns or sets if text is automatically indented when the text alignment in a cell is set to equal distribution either horizontally or vertically.
@property BOOL autoScaleFont;  // Returns or sets  if the text in the object changes font size when the object size changes.
@property BOOL autoSize;  // Returns or sets if the size of the specified object is changed automatically to fit text within its boundaries.
@property (copy, readonly) ExcelRange *bottomRightCell;  // Returns the bottom right cell of the range the control is occupying.
@property BOOL cancelButton;  // Returns or sets if this button is the cancel button.
@property (copy) NSString *caption;  // Returns or sets the caption for this object.
@property (copy) NSString *controlText;  // Returns or sets the default text for the control. 
@property BOOL defaultButton;  // Returns or sets if this button is the default button.
@property BOOL dismissButton;  // Returns or sets if this button is the dismiss button.
@property BOOL enabled;  // Returns or sets if the object is enabled.
@property (readonly) NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property (copy, readonly) ExcelFont *fontObject;  // Returns a font object that represents the font of the specified object.
@property (copy) NSString *formula;  // Returns or sets the object's formula, in A1-style notation and in the language of the macro.
@property double height;  // Returns or set the height of the object.
@property BOOL helpButton;  // Returns or sets if this button is the help button.
@property ExcelXlHAlign horizontalAlignment;  // Returns or sets the horizontal alignment for the object.
@property double leftPosition;  // Returns or sets the position of the specified object, in points.
@property BOOL locked;  // Returns or sets if the object is locked, if false the object can be modified when the sheet is protected.
@property BOOL lockedText;  // Returns or sets whether the control's text is locked for editing.
@property (copy) NSString *name;  // Returns or sets the name of the object.
@property (copy) NSString *onAction;  // Returns or sets the name of a string of AppleScript commands that will be executed when the object is clicked on. Please note that if AppleScript commands are set they will not be saved with the document.
@property ExcelXlOrientation orientation;  // May also be a number value from -90 to 90 degrees.
@property (copy) NSString *phoneticAccelerator;  // Returns or sets the link to a range.
@property ExcelXlPlacement placement;  // Returns or sets how the object is placed on the worksheet.
@property BOOL printObject;  // Returns or sets if this object is printed.
@property ExcelXLDefaultSheetDir readingOrder;  // Returns or sets the reading order for the specified object.
@property double top;  // Returns the top position of the specified object, in points.
@property (copy, readonly) ExcelRange *topLeftCell;  // Returns a range object that represents the cell that lies under the upper-left corner of the specified object. 
@property ExcelXlVerticalAlignmentTarget verticalAlignment;  // Returns or sets the vertical alignment of the object.
@property BOOL visible;  // Returns or sets if the object is visible.
@property double width;  // Returns or sets  the width of the object.
@property (readonly) NSInteger zOrderPosition;  // Returns the z-order position of the object.


@end

// Represents the calculated fields and calculated items for PivotTables with Online Analytical Processing data sources.
@interface ExcelCalculatedMember : SBObject <ExcelGenericMethods>

@property (copy, readonly) NSString *_default;
@property (copy, readonly) NSString *displayFolder;  // An ST_Xstring attribute that specifies the display folder of this PivotTable named set.
@property (readonly) BOOL dynamic;
@property (readonly) BOOL flattenHierarchies;
@property (copy, readonly) NSString *formula;  // Returns the member's formula in multidimensional expressions syntax.
@property (readonly) BOOL hierarchizeDistinct;
@property (readonly) BOOL isValid;  // Returns a Boolean that indicates whether the specified calculated member has been successfully instantiated with the OLAP provider during the current session.
@property (copy, readonly) NSString *name;  // Calculated Member Name.
@property (readonly) NSInteger solveOrder;  // Calculated Members Solve Order.
@property (copy, readonly) NSString *sourceName;  // Returns the specified object's name as it appears in the original source data for the specified PivotTable report.
@property (readonly) ExcelXlCalculatedMemberType type;


@end

@interface ExcelCheckbox : SBObject <ExcelGenericMethods>

- (SBElementArray<ExcelCharacter *> *) characters;

@property (copy) NSString *accelerator;  // Returns or sets the accelerator character for this control.
@property (copy, readonly) ExcelBorder *border;  // Returns a border object that represents the border of the object. 
@property (copy, readonly) ExcelRange *bottomRightCell;  // Returns the bottom right cell of the range the control is occupying.
@property (copy) NSString *caption;  // Returns or sets the caption for this object.
@property (copy) NSString *controlText;  // Returns or sets the default text for the control. 
@property BOOL displayThreeDShading;
@property BOOL enabled;  // Returns or sets if the control is enabled
@property (readonly) NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property double height;  // Returns or sets the height of the control.
@property (copy, readonly) ExcelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property double leftPosition;  // Returns or sets the left position of the control
@property (copy) NSString *linkedCell;
@property BOOL locked;  // Returns or sets whether the control is locked for editing.
@property BOOL lockedText;  // Returns or sets whether the control's text is locked for editing.
@property (copy) NSString *name;  // Returns or sets the name of this control
@property (copy) NSString *onAction;  // Returns or sets the name of a string of AppleScript commands that will be executed when the object is clicked on. Please note that if AppleScript commands are set they will not be saved with the document.
@property (copy) NSString *phoneticAccelerator;  // Returns or sets the link to a range.
@property ExcelXlPlacement placement;  // Returns or sets how the object is placed on the worksheet.
@property BOOL printObject;  // Returns or sets if this object is printed.
@property double top;  // Returns or sets the position of the top of the control.
@property (copy, readonly) ExcelRange *topLeftCell;  // Returns the top left cell of the range within which the control is positioned.
@property ExcelXlCheckBoxState value;
@property BOOL visible;  // Returns or sets the current value of the control
@property double width;  // Returns or sets  the width of the object.
@property (readonly) NSInteger zOrderPosition;  // Returns the most recently set z order position, bring shape to front/send shape to back/bring shape forward/send shape backward/bring shape in front of text/send shape behind text.


@end

@interface ExcelColorScaleCriteria : SBObject <ExcelGenericMethods>


@end

@interface ExcelColorScaleCriterion : SBObject <ExcelGenericMethods>

@property (readonly) NSInteger colorScaleCriterionIndex;  // The index of the color scale criterion. Read only.
@property ExcelXlConditionValueTypes colorScaleCriterionType;  // Returns or sets the condition value type of the color scale criterion. Read/Write.
@property (copy) id colorScaleCriterionValue;
@property (copy, readonly) ExcelFormatColor *formatColor;  // Returns the FormatColor for the ColorScaleCriterion object.


@end

@interface ExcelColorScaleFormatCondition : SBObject <ExcelGenericMethods>

@property (copy, readonly) ExcelRange *appliesTo;  // Returns the range this format condition applies to. Read Only.
@property (copy, readonly) ExcelColorScaleCriteria *colorScaleCriteria;  // Returns the ColorScaleCriteria for the ColorScale object.
@property NSInteger colorScaleType;
@property NSInteger formatConditionPriority;  // Specifies the priority for the format condition. Read/Write.
@property (readonly) ExcelXlFormatConditionType formatConditionType;  // Return the conditional format type.
@property (copy) NSString *formula;  // Returns or sets the value or expression associated with the conditional format or data validation.
@property ExcelXlPivotConditionScope pivotConditionScopeType;  // Returns or sets the part of the pivot table that the pivot table format condition is scoped to
@property (readonly) BOOL pivotTableCondition;  // Tells whether this format condition is a pivot table condition.
@property (readonly) BOOL stopIfTrue;  // Tells whether the processing of format conditions stops if this condition is true. Read Only.


@end

// Represents the color stop point for a gradient fill in an range or selection.
@interface ExcelColorstop : SBObject <ExcelGenericMethods>

@property (copy) NSArray<NSNumber *> *color;  // Returns or sets the primary color of the object.
@property double colorstopPosition;  // Returns or sets the position of the ColorStop.
@property ExcelXlThemeColor themeColor;  // Returns or sets the theme color in the applied color scheme that is associated with the specified object.
@property double tintAndShade;  // Returns or sets a Single that lightens or darkens a color.

- (void) deleteColorstop;  // Clears the represented object.

@end

// A collection of all the ColorStop objects for the specified series.
@interface ExcelColorstops : SBObject <ExcelGenericMethods>

- (ExcelColorstop *) addColorstopPosition:(double)position;  // Adds ColorStop object to the ColorStops object.

@end

@interface ExcelConditionValue : SBObject <ExcelGenericMethods>

@property (readonly) ExcelXlConditionValueTypes conditionValueType;
@property (copy, readonly) id conditionValueValue;

- (void) modifyConditionValueType:(ExcelXlConditionValueTypes)type conditionValue:(id)conditionValue;  // Modifies an existing condition value.

@end

// Represents a hierarchy or measure field from an OLAP cube
@interface ExcelCubeField : SBObject <ExcelGenericMethods>

- (SBElementArray<ExcelPivotField *> *) pivotFields;

@property (readonly) BOOL allItemsVisible;  // The AllItemsVisible property checks whether manual filtering is applied to a PivotField or CubeField.
@property (copy) NSString *caption;  // The label text for the cube field.
@property (readonly) ExcelXlCubeFieldSubType cubeFieldSubType;  // Specifies the type of a CubeField.
@property (readonly) ExcelXlCubeFieldType cubeFieldType;  // Indicates whether the OLAP cube field is a hierarchy field or a measure field.
@property (copy) NSString *currentPageName;  // Returns or sets the page name for a CubeField.
@property BOOL dragToColumn;  // True if the specified field can be dragged to the column position.
@property BOOL dragToData;  // True if the specified field can be dragged to the data position.
@property BOOL dragToHide;  // True if the specified field can be dragged to the column position.
@property BOOL dragToPage;  // True if the field can be dragged to the page position.
@property BOOL dragToRow;  // True if the field can be dragged to the row position.
@property BOOL enableMultiplePageItems;  // True to allow multiple items in the page field area for OLAP PivotTables to be selected.
@property BOOL flattenHierarchies;
@property (readonly) BOOL hasMemberProperties;  // Returns True when there are member properties specified to be displayed for the cube field.
@property BOOL hierarchizeDistinct;
@property BOOL includeNewItemsInFilter;  // The IncludeNewItemsInFilter property is used to track included or excluded items in OLAP PivotTables.
@property (readonly) BOOL isDate;  // Returns True if the CubeField is a date.
@property (readonly) ExcelXlLayoutFormType layoutForm;  // Returns or sets the way the specified PivotTable items appear -- in table format or in outline format.
@property ExcelXlSubtototalLocationType layoutSubtotalLocation;  // Returns or sets the position of the PivotTable field subtotals in relation to, either above or below, the specified field.
@property (copy, readonly) NSString *name;  // Returns the name of the object.
@property ExcelXlPivotFieldOrientation orientation;  // The location of the field in the specified PivotTable report.
@property NSInteger position;  // Position of the item in its field if the item is currently showing.
@property BOOL showInFieldList;  // When set to True, a CubeField object will be shown in the field list.
@property (copy, readonly) ExcelTreeviewControl *treeviewControl;
@property (copy, readonly) NSString *value;  // Returns the name of the specified field.

- (void) addMemberPropertyFieldProperty:(NSString *)property propertyOrder:(NSString *)propertyOrder propertyDisplayedIn:(ExcelXlPropertyDisplayedIn)propertyDisplayedIn;  // Adds a member property field to the display for the cube field.
- (void) createPivotFields;  // The CreatePivotFields method enables users to create all PivotFields of a CubeField.

@end

// Represents a custom workbook view.
@interface ExcelCustomView : SBObject <ExcelGenericMethods>

@property (readonly) BOOL customViewPrintSettings;  // True if print settings are included in the custom view.
@property (copy, readonly) NSString *name;  // Returns the name of this object.
@property (readonly) BOOL rowColSettings;  // Returns true if the custom view includes settings for hidden rows and columns, including filter information.

- (void) showCustomView;  // Displays the custom view

@end

@interface ExcelDatabarBorder : SBObject <ExcelGenericMethods>

@property (copy, readonly) ExcelFormatColor *databarBorderColor;
@property ExcelXlDataBarBorderType databarBorderType;  // Returns or sets the type of border of the databar


@end

@interface ExcelDatabarFormatCondition : SBObject <ExcelGenericMethods>

@property (copy, readonly) ExcelRange *appliesTo;  // Returns the range this format condition applies to. Read Only.
@property (copy, readonly) ExcelFormatColor *databarAxisColor;
@property ExcelXlDataBarAxisPosition databarAxisPosition;  // Returns or sets the position of the databar axis
@property (copy, readonly) ExcelFormatColor *databarBarColor;
@property (copy, readonly) ExcelDatabarBorder *databarBorder;  // Returns the DataBarBorder for the Databar object.
@property NSInteger databarDirection;  // Specifies the direction of the databar. Read/Write.
@property ExcelXlDataBarFillType databarFillType;  // Returns or sets the type of fill of the databar
@property NSInteger formatConditionPriority;  // Specifies the priority for the format condition. Read/Write.
@property BOOL formatConditionShowValue;  // Specifies whether to show the cell value along with the databar. Read/Write.
@property (readonly) ExcelXlFormatConditionType formatConditionType;  // Return the conditional format type.
@property (copy) NSString *formula;  // Returns or sets the value or expression associated with the conditional format or data validation.
@property (copy, readonly) ExcelConditionValue *maxPointConditionValue;  // Returns the ConditionValue for the maximum point of the DataBar object.
@property NSInteger maxPointPercent;  // Specifies the percentage of the data bar to draw at the maximum point. Read/Write.
@property (copy, readonly) ExcelConditionValue *minPointConditionValue;  // Returns the ConditionValue for the minimum point of the DataBar object.
@property NSInteger minPointPercent;  // Specifies the percentage of the data bar to draw at the minimum point. Read/Write.
@property (copy, readonly) ExcelNegativeBarFormat *negativeBarFormat;  // Returns the NegativeBarFormat for the Databar object.
@property ExcelXlPivotConditionScope pivotConditionScopeType;  // Returns or sets the part of the pivot table that the pivot table format condition is scoped to
@property (readonly) BOOL pivotTableCondition;  // Tells whether this format condition is a pivot table condition.
@property (readonly) BOOL stopIfTrue;  // Tells whether the processing of format conditions stops if this condition is true. Read Only.


@end

// Contains global application-level attributes used by Microsoft Excel when you save a document as a Web page or open a Web page.
@interface ExcelDefaultWebOptions : SBObject <ExcelGenericMethods>

@property BOOL allowPng;  // Returns or sets if PNG, Portable Network Graphics, is allowed as an image format when you save documents as a Web page.
@property BOOL alwaysSaveInDefaultEncoding;  // Returns or sets if the default encoding is used when you save a Web page or plain text document, independent of the file's original encoding when opened.
@property ExcelMsoEncoding encoding;  // Returns or sets the document encoding, code page or character set, to be used by the Web browser when you view the saved document.
@property (copy) NSString *locationOfComponents;  // Returns or sets the central URL, on the intranet or Web, or path, local or network, to the location from which authorized users can download Microsoft Office Web components when viewing your saved document.
@property NSInteger pixelsPerInch;  // Returns or sets the density, pixels per inch, of graphics images and table cells on a Web page. The range of settings is usually from 19 to 480, and common settings for popular screen sizes are 72, 96, and 120.
@property ExcelMsoScreenSize screenSize;  // Returns or sets the ideal minimum screen size, width by height, in pixels, that you should use when viewing the saved document in a Web browser.
@property BOOL updateLinksOnSave;  // Returns or sets if hyperlinks and paths to all supporting files are automatically updated before you save the document as a Web page, ensuring that the links are up-to-date at the time the document is saved. If false, the links are not updated.


@end

// Represents a built-in Microsoft Excel dialog box.
@interface ExcelDialog : SBObject <ExcelGenericMethods>


@end

// Represents the formatting shown to the user.
@interface ExcelDisplayFormat : SBObject <ExcelGenericMethods>

@property (readonly) BOOL addIndent;  // Returns or sets if text is automatically indented when the text alignment in a cell is set to equal distribution either horizontally or vertically.
@property (copy, readonly) ExcelFont *fontObject;  // Returns a font object that represents the font of the specified object.
@property (readonly) BOOL formulaHidden;  // Returns or sets if the formula will be hidden when the workbook or worksheet is protected.
@property (readonly) ExcelXlHAlign horizontalAlignment;  // Returns or sets the horizontal alignment for the range.
@property (readonly) NSInteger indentLevel;  // Returns or sets the indent level for the range or style. Can be an integer from 0 to 15.
@property (copy, readonly) ExcelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property (readonly) BOOL locked;  // Returns or sets  if the range is locked.
@property (readonly) BOOL mergeCells;  // Returns or sets if the range contains merged cells. 
@property (copy, readonly) NSString *numberFormat;  // Returns or sets the format code for the object.
@property (readonly) ExcelXLDefaultSheetDir readingOrder;  // Returns or sets the reading order for the specified object.
@property (readonly) BOOL shrinkToFit;  // Returns or sets if text automatically shrinks to fit in the available column width.
@property (copy, readonly) ExcelStyle *styleObject;  // Returns or sets a style object that represents the style of the specified range.
@property (readonly) ExcelXlOrientation textOrientation;  // The text orientation. Can be a number value from -90 to 90 degrees.
@property (readonly) ExcelXlVerticalAlignmentTarget verticalAlignment;  // Returns or sets the vertical alignment of the range.
@property (readonly) BOOL wrapText;  // Returns or sets if Microsoft Excel wraps the text in the object.


@end

@interface ExcelDropdown : SBObject <ExcelGenericMethods>

- (SBElementArray<ExcelCharacter *> *) characters;

@property (copy, readonly) ExcelRange *bottomRightCell;  // Returns the bottom right cell of the range the control is occupying. 
@property (copy) NSString *caption;  // Returns or sets the caption for the control.
@property (copy) NSString *controlText;  // Returns or sets the default text for the control.
@property BOOL displayThreeDShading;  // Returns or sets whether a 3-d effect will be employed when displaying the control
@property NSInteger dropDownLines;  // Returns or sets the number of dropdown items.
@property BOOL enabled;  // Returns or sets if the control is enabled
@property (readonly) NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property double height;  // Returns or sets the height of the control.
@property double leftPosition;  // Returns or sets the left position of the control
@property (copy) NSString *linkedCell;  // Returns or sets reference to a call which contains the value of the control.
@property (copy) NSString *listFillRange;  // Returns or sets the items which are contained in the drop down.
@property NSInteger listIndex;  // Returns or sets the selected item.
@property BOOL locked;  // Returns or sets whether the control is locked for editing.
@property (copy) NSString *name;  // Returns or sets the name of this control.
@property (readonly) NSInteger numberOfItemsInList;  // Returns the number of total items in the list.
@property (copy) NSString *onAction;  // Returns or sets the name of a string of AppleScript commands that will be executed when the object is clicked on. Please note that if AppleScript commands are set they will not be saved with the document.
@property ExcelXlPlacement placement;  // Returns or sets how the object is placed on the worksheet.
@property BOOL printObject;  // Returns or sets if this object is printed.
@property double top;  // Returns or sets the position of the top of the control.
@property (copy, readonly) ExcelRange *topLeftCell;  // Returns the top left cell of the range within which the control is positioned.
@property NSInteger value;  // Returns or sets the current value of the control
@property BOOL visible;  // Returns or sets if the control is visible.
@property double width;  // Returns or sets  the width of the object.
@property (readonly) NSInteger zOrderPosition;  // Returns the most recently set z order position, bring shape to front/send shape to back/bring shape forward/send shape backward/bring shape in front of text/send shape behind text.


@end

// Represents a filter for a single column. 
@interface ExcelFilter : SBObject <ExcelGenericMethods>

@property (copy, readonly) id criteria1;  // Returns the first filtered value for the specified column in a filtered range.
@property (copy, readonly) id criteria2;  // Returns the second filtered value for the specified column in a filtered range.
@property (readonly) BOOL filterOn;  // Returns true if the specified filter is on.
- (ExcelXlAutoFilterOperator) operator;  // set or return the operator that associates the two criteria applied by the specified filter.
- (void) setOperator: (ExcelXlAutoFilterOperator) _operator;


@end

// Represents a color that a format condition can be colored with.
@interface ExcelFormatColor : SBObject <ExcelGenericMethods>

@property (copy) NSArray<NSNumber *> *color;  // Returns or sets the primary color of the object.
@property ExcelXlColorIndex colorIndex;
@property ExcelXlThemeColor themeColor;  // Returns or sets the theme color in the applied color scheme that is associated with the specified object.
@property double tintAndShade;  // Returns or sets a Single that lightens or darkens a color.


@end

@interface ExcelFormatConditionIconObject : SBObject <ExcelGenericMethods>

@property (readonly) NSInteger formatConditionIconIndex;  // The index of the icon. Read only.


@end

@interface ExcelFormatConditionIconSet : SBObject <ExcelGenericMethods>

@property (readonly) ExcelXlIconSet iconSetId;


@end

@interface ExcelFormatConditionIconSets : SBObject <ExcelGenericMethods>


@end

// Represents a conditional format.
@interface ExcelFormatCondition : SBObject <ExcelGenericMethods>

@property (copy, readonly) ExcelRange *appliesTo;  // Returns the range this format condition applies to. Read Only.
@property (readonly) ExcelXlFormatConditionOperator conditionOperator;  // Returns the operator for the conditional format.
@property (copy, readonly) ExcelFont *fontObject;  // Returns a font object that represents the font of the specified object.
@property ExcelXlTimePeriods formatConditionDateOperator;
@property NSInteger formatConditionPriority;  // Specifies the priority for the format condition. Read/Write.
@property (copy) NSString *formatConditionText;  // Returns or sets the text for the specified object.
@property ExcelXlContainsOperator formatConditionTextOperator;
@property (readonly) ExcelXlFormatConditionType formatConditionType;  // Return the conditional format type.
@property (copy, readonly) NSString *formula1;  // Returns the value or expression associated with the conditional format or data validation.
@property (copy, readonly) NSString *formula2;  // Returns the value or expression associated with the second part of a conditional format or data validation. Used only when the data validation conditional format Operator property is operator between or operator not between.
@property (copy, readonly) ExcelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property (copy) NSString *numberFormat;  // Returns or sets the format code for the object.
@property ExcelXlPivotConditionScope pivotConditionScopeType;  // Returns or sets the part of the pivot table that the pivot table format condition is scoped to
@property (readonly) BOOL pivotTableCondition;  // Tells whether this format condition is a pivot table condition.
@property BOOL stopIfTrue;  // Tells whether the processing of format conditions stops if this condition is true. Read/Write.

- (void) modifyConditionType:(ExcelXlFormatConditionType)type operator:(ExcelXlFormatConditionOperator)operator_ formula1:(NSString *)formula1 formula2:(NSString *)formula2 string:(NSString *)string operator2:(ExcelXlFormatConditionOperator)operator2;  // Modifies an existing conditional format.

@end

// Contains properties that apply to header and footer picture objects.
@interface ExcelGraphic : SBObject <ExcelGenericMethods>

@property double brightness;  // Returns or sets the brightness of the specified picture. The value for this property must be a number from 0.0, dimmest, to 1.0, brightest.
@property ExcelMsoPictureColorType colorType;  // Returns or sets the type of color transformation applied to the specified picture.
@property double contrast;  // Returns or sets the contrast for the specified picture. The value for this property must be a number from 0.0, the least contrast, to 1.0, the greatest contrast.
@property double cropBottom;  // Returns or sets the number of points that are cropped off the bottom of the specified picture.
@property double cropLeft;  // Returns or sets the number of points that are cropped off the left side of the specified picture.
@property double cropRight;  // Returns or sets the number of points that are cropped off the right side of the specified picture.
@property double cropTop;  // Returns or sets the number of points that are cropped off the top of the specified picture.
@property (copy) NSString *fileName;  // Returns or sets the URL, on the intranet or the Web, or path, local or network, to the location where the specified source object was saved.
@property double height;  // Returns or sets the height of the object.
@property BOOL lockAspectRatio;  // Returns or sets if the specified shape retains its original proportions when you resize it. False if you can change the height and width of the shape independently of one another when you resize it.
@property double width;  // Returns or sets the width of the object.


@end

@interface ExcelGroupbox : SBObject <ExcelGenericMethods>

- (SBElementArray<ExcelCharacter *> *) characters;

@property (copy) NSString *accelerator;  // Returns or sets the accelerator character for this control.
@property (copy, readonly) ExcelRange *bottomRightCell;  // Returns the bottom right cell of the range the control is occupying.
@property (copy) NSString *caption;  // Returns or sets the caption for this object.
@property (copy) NSString *controlText;  // Returns or sets the default text for the control. 
@property BOOL displayThreeDShading;
@property BOOL enabled;  // Returns or sets if the object is enabled.
@property (readonly) NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property double height;  // Returns or set the height of the object.
@property double leftPosition;  // Returns or sets the position of the specified object, in points.
@property BOOL locked;  // Returns or sets if the object is locked, if false the object can be modified when the sheet is protected.
@property BOOL lockedText;  // Returns or sets whether the control's text is locked for editing.
@property (copy) NSString *name;  // Returns or sets the name of the object.
@property (copy) NSString *onAction;  // Returns or sets the name of a string of AppleScript commands that will be executed when the object is clicked on. Please note that if AppleScript commands are set they will not be saved with the document.
@property (copy) NSString *phoneticAccelerator;  // Returns or sets the link to a range.
@property ExcelXlPlacement placement;  // Returns or sets how the object is placed on the worksheet.
@property BOOL printObject;  // Returns or sets if this object is printed.
@property double top;  // Returns or sets the position of the top of the control.
@property (copy, readonly) ExcelRange *topLeftCell;  // Returns the top left cell of the range within which the control is positioned.
@property BOOL visible;  // Returns or sets if the control is visible
@property double width;  // Returns or sets  the width of the object.
@property (readonly) NSInteger zOrderPosition;  // Returns the most recently set z order position, bring shape to front/send shape to back/bring shape forward/send shape backward/bring shape in front of text/send shape behind text.


@end

// Represents a horizontal page break.
@interface ExcelHorizontalPageBreak : SBObject <ExcelGenericMethods>

@property (readonly) ExcelXlPageBreakExtent extent;  // Returns the type of the specified page break: full-screen or only within a print area.
@property ExcelXlPageBreak horizontalPageBreakType;  // Returns or sets the type of horizontal page break.
@property (copy) ExcelRange *location;  // Returns or sets where this horizontal page break is located.


@end

// Represents a hyperlink.
@interface ExcelHyperlink : SBObject <ExcelGenericMethods>

@property (copy) NSString *address;  // Returns or sets the address of the target document.
@property (copy) NSString *emailSubject;  // Returns or sets the text string of the specified hyperlink's e-mail subject line. The subject line is appended to the hyperlink's address.
@property (readonly) ExcelMsoHyperlinkType hyperlinkType;  // Returns the hyperlink type, what the hyperlink is associated with.
@property (copy, readonly) NSString *name;  // Returns the name of the object.
@property (copy, readonly) ExcelRange *rangeObject;  // Returns a range object that represents the range the specified hyperlink is attached to.
@property (copy) NSString *screenTip;  // Returns or sets the screen tip text for the specified hyperlink.
@property (copy, readonly) ExcelShape *shapeObject;  // Returns a shape object that represents the shape attached to the specified hyperlink.
@property (copy) NSString *subAddress;  // Returns or sets the location within the document associated with the hyperlink.
@property (copy) NSString *textToDisplay;  // Returns or sets the text to be displayed for the specified hyperlink. The default value is the address of the hyperlink.

- (void) createNewDocumentFileName:(NSString *)fileName editNow:(BOOL)editNow overwrite:(BOOL)overwrite;  // Creates a new document linked to the specified hyperlink.
- (void) followNewWindow:(BOOL)newWindow extraInfo:(NSString *)extraInfo method:(ExcelMsoExtraInfoMethod)method headerInfo:(NSString *)headerInfo;  // Displays a cached document, if it's already been downloaded. Otherwise, this method resolves the hyperlink, downloads the target document, and displays the document in the appropriate application.

@end

@interface ExcelIconCriteria : SBObject <ExcelGenericMethods>


@end

@interface ExcelIconCriterion : SBObject <ExcelGenericMethods>

@property ExcelXlFormatConditionOperator conditionOperator;  // Returns the operator for the conditional format.
@property ExcelXlIcon iconCriterionIcon;
@property (readonly) NSInteger iconCriterionIndex;  // The index of the icon criterion. Read only.
@property ExcelXlConditionValueTypes iconCriterionType;  // Returns or sets the condition value type of the icon criterion. Read/Write.
@property (copy) id iconCriterionValue;


@end

@interface ExcelIconSetFormatCondition : SBObject <ExcelGenericMethods>

@property (copy, readonly) ExcelRange *appliesTo;  // Returns the range this format condition applies to. Read Only.
@property (copy) ExcelFormatConditionIconSet *formatConditionIconSet;
@property NSInteger formatConditionPriority;  // Specifies the priority for the format condition. Read/Write.
@property (readonly) ExcelXlFormatConditionType formatConditionType;  // Return the conditional format type.
@property (copy) NSString *formula;  // Returns or sets the value or expression associated with the conditional format or data validation.
@property (copy, readonly) ExcelIconCriteria *iconCriteria;  // Returns the IconCriteria for the IconSetCondition object.
@property BOOL percentileValues;
@property ExcelXlPivotConditionScope pivotConditionScopeType;  // Returns or sets the part of the pivot table that the pivot table format condition is scoped to
@property (readonly) BOOL pivotTableCondition;  // Tells whether this format condition is a pivot table condition.
@property BOOL reverseIconSetOrder;
@property BOOL showIconOnly;
@property (readonly) BOOL stopIfTrue;  // Tells whether the processing of format conditions stops if this condition is true. Read Only.


@end

@interface ExcelLabel : SBObject <ExcelGenericMethods>

- (SBElementArray<ExcelCharacter *> *) characters;

@property (copy) NSString *accelerator;  // Returns or sets the accelerator character for this control.
@property (copy, readonly) ExcelRange *bottomRightCell;  // Returns the bottom right cell of the range the control is occupying.
@property (copy) NSString *caption;  // Returns or sets the caption for the control.
@property (copy) NSString *controlText;  // Returns or sets the default text for the control. 
@property BOOL enabled;  // Returns or sets if the control is enabled
@property (readonly) NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property double height;  // Returns or sets the height of the control.
@property double leftPosition;  // Returns or sets the left position of the control
@property BOOL locked;  // Returns or sets whether the control is locked for editing.
@property BOOL lockedText;  // Returns or sets whether the control's text is locked for editing.
@property (copy) NSString *name;  // Returns or sets the name of this control
@property (copy) NSString *onAction;  // Returns or sets the name of a string of AppleScript commands that will be executed when the object is clicked on. Please note that if AppleScript commands are set they will not be saved with the document.
@property (copy) NSString *phoneticAccelerator;  // Returns or sets the link to a range.
@property ExcelXlPlacement placement;  // Returns or sets how the object is placed on the worksheet.
@property BOOL printObject;  // Returns or sets if this object is printed.
@property double top;  // Returns or sets the position of the top of the control.
@property (copy, readonly) ExcelRange *topLeftCell;  // Returns the top left cell of the range within which the control is positioned.
@property BOOL visible;  // Returns or sets if the control is visible
@property double width;  // Returns or sets  the width of the object.
@property (readonly) NSInteger zOrderPosition;  // Returns the most recently set z order position, bring shape to front/send shape to back/bring shape forward/send shape backward/bring shape in front of text/send shape behind text.


@end

// The LinearGradient object transitions through a series of colors in a linear manner along a specific angle.
@interface ExcelLinearGradient : SBObject <ExcelGenericMethods>

@property (copy, readonly) ExcelColorstops *colorstops;  // Returns the ColorStops for the LinearGradient object.
@property double linearGradientDegree;  // The angle of the linear gradient fill within a selection.


@end

// Represents a column in a list object.
@interface ExcelListColumn : SBObject <ExcelGenericMethods>

@property (copy, readonly) ExcelCell *cellTable;  // Returns the cell table from a specified list column. 
@property (readonly) NSInteger index;
@property (copy) NSString *name;  // Returns or sets the name of the object.
@property (copy, readonly) ExcelRange *rangeObject;  // Returns a range object that represents the range to which the specified list column.
@property (copy, readonly) ExcelRange *totalRow;  // Returns the totals row, if any, from a specified list column.
@property ExcelXlTotalsCalculation totalsCalculation;


@end

// Represents a list object on a worksheet.
@interface ExcelListObject : SBObject <ExcelGenericMethods>

- (SBElementArray<ExcelListColumn *> *) listColumns;
- (SBElementArray<ExcelListRow *> *) listRows;

@property (readonly) BOOL active;
@property (copy, readonly) ExcelAutofilter *autofilterObject;  // Returns the autofilter object associated with this list object.
@property (copy, readonly) ExcelCell *cellTable;  // Returns the cell table from a specified list object. 
@property (copy) NSString *comment;  // Returns or sets the comment of the object.
- (NSString *) default;
@property (copy) NSString *displayName;  // Returns or sets the display name of the object.
@property (readonly) BOOL displayRightToLeft;
@property (copy, readonly) ExcelRange *headerRow;  // Returns a range object that represents the used range on the specified worksheet. 
@property (copy, readonly) ExcelRange *insertRow;
@property (copy) NSString *name;  // Returns or sets the name of the object.
@property (copy, readonly) ExcelQueryTable *queryTable;  // Returns a query table object that represents the query table that intersects the specified range.
@property (copy, readonly) ExcelRange *rangeObject;  // Returns or sets a range object that represents the range to which the specified list column, object, or row applies.
@property BOOL showAutofilter;  // Returns or sets if the autofilter is implemented in a list object.
@property BOOL showHeaders;  // Returns or sets if headers is implemented in a list object.
@property BOOL showTableStyleColumnStripes;  // The ShowTableStyleColumnStripes property displays banded columns in which even columns are formatted differently from odd columns.
@property BOOL showTableStyleFirstColumn;  // Returns or sets if the first column is displayed for the specified ListObject object.
@property BOOL showTableStyleLastColumn;  // Returns or sets if the last column is displayed for the specified ListObject object.
@property BOOL showTableStyleRowStripes;  // The ShowTableStyleRowStripes property displays banded rows in which even rows are formatted differently from odd rows.
@property (copy, readonly) ExcelSort *sortObject;  // Returns the sort object associated with this list object.
@property (readonly) ExcelXlListObjectSourceType sourceType;
@property (copy) id tableStyle;  // Returns or sets the style used in the table body.
@property BOOL total;  // Returns or sets if a totals row be implemented in a list object.
@property (copy, readonly) ExcelRange *totalRow;  // Returns the totals row, if any, from a specified list object.

- (void) clearContents;  // Clears the formulas from the range. Clears the data from a chart but leaves the formatting. Clears all the data, formatting, and formulas from a list object.
- (ExcelRange *) convertToRange;  // Converts a list object to a normal Excel range.
- (void) resizeRange:(ExcelRange *)range;

@end

// Represents a row in a list object.
@interface ExcelListRow : SBObject <ExcelGenericMethods>

@property (readonly) NSInteger index;
@property (copy, readonly) ExcelRange *rangeObject;  // Returns a range object that represents the range to which the specified list row applies.


@end

@interface ExcelListbox : SBObject <ExcelGenericMethods>

@property (copy, readonly) ExcelRange *bottomRightCell;  // Returns the bottom right cell of the range the control is occupying.
@property BOOL displayThreeDShading;  // Returns or sets whether a 3-d effect will be employed when displaying the control
@property BOOL enabled;  // Returns or sets if the control is enabled
@property (readonly) NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property double height;  // Returns or sets the height of the control.
@property double leftPosition;  // Returns or sets the left position of the control
@property (copy) NSString *linkedCell;  // Returns or sets reference to a call which contains the value of the control.
@property (copy) NSString *listFillRange;  // Returns or sets the items which are contained in the control.
@property NSInteger listIndex;  // Returns or sets the selected item.
@property BOOL locked;  // Returns or sets whether the control is locked for editing.
@property ExcelXlMultiSelect multiSelect;  // Returns or sets the multiple selection setting.
@property (copy) NSString *name;  // Returns or sets the name of this control
@property (readonly) NSInteger numberOfItemsInList;
@property (copy) NSString *onAction;  // Returns or sets the name of a string of AppleScript commands that will be executed when the object is clicked on. Please note that if AppleScript commands are set they will not be saved with the document.
@property ExcelXlPlacement placement;  // Returns or sets how the object is placed on the worksheet.
@property BOOL printObject;  // Returns or sets if this object is printed.
@property double top;  // Returns or sets the position of the top of the control.
@property (copy, readonly) ExcelRange *topLeftCell;  // Returns the top left cell of the range within which the control is positioned. 
@property NSInteger value;  // Returns or sets the current value of the control.
@property BOOL visible;  // Returns or sets if the control is visible
@property double width;  // Returns or sets  the width of the object.
@property (readonly) NSInteger zOrderPosition;  // Returns the most recently set z order position, bring shape to front/send shape to back/bring shape forward/send shape backward/bring shape in front of text/send shape behind text.


@end

// Represents a defined name for a range of cells. Named items can be either built-in names, such as Database, Print_Area, and Auto_Open, or custom names.
@interface ExcelNamedItem : SBObject <ExcelGenericMethods>

@property (copy) NSString *category;  // Returns or sets the category for the specified name in the language of the macro.
@property (copy) NSString *categoryLocal;  // Returns or sets the category for the specified name, in the language of the user, if the name refers to a custom function or command.
@property (readonly) NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property ExcelXlXLMMacroType macroType;  // Returns or sets what the name refers to.
@property (copy) NSString *name;  // Returns or sets the name of the object.
@property (copy) NSString *nameLocal;  // Returns or sets the name of the object, in the language of the user.
@property (copy) NSString *referenceLocal;  // Returns or sets the formula that the name refers to. The formula is in the language of the user, and it's in A1-style notation, beginning with an equal sign.
@property (copy) NSString *referenceLocalR1c1;  // Returns or sets the formula that the name refers to. This formula is in the language of the user, and it's in R1C1-style notation, beginning with an equal sign.
@property (copy) NSString *referenceR1c1;  // Returns or sets the formula that the name refers to. The formula is in the language of the macro, and it's in R1C1-style notation, beginning with an equal sign.
@property (copy, readonly) ExcelRange *referenceRange;  // Returns the range object referred to by this object.
@property (copy) NSString *references;  // Returns or sets the formula that the name is defined to refer to, in the language of the macro and in A1-style notation, beginning with an equal sign.
@property (copy) NSString *shortcutKey;  // Returns or sets the shortcut key for a name defined as a custom Microsoft Excel 4.0 macro command.
@property (copy) NSString *value;  // Returns or sets a string containing the formula that the name is defined to refer to. The string is in A1-style notation in the language of the macro, and it begins with an equal sign.
@property BOOL visible;  // Returns or sets if the object is visible.


@end

@interface ExcelNegativeBarFormat : SBObject <ExcelGenericMethods>

@property (copy, readonly) ExcelFormatColor *databarBarColor;
@property (copy, readonly) ExcelFormatColor *databarBorderColor;  // Returns the DataBarBorder for the Databar object.
@property ExcelXlDataBarNegativeColorType negativeBarBorderColorType;  // Returns or sets the type of border of the databar
@property ExcelXlDataBarNegativeColorType negativeBarColorType;  // Returns or sets the type of fill of the databar


@end

@interface ExcelOptionButton : SBObject <ExcelGenericMethods>

- (SBElementArray<ExcelCharacter *> *) characters;

@property (copy) NSString *accelerator;  // Returns or sets the accelerator character for this control.
@property (copy, readonly) ExcelBorder *border;  // Returns a border object that represents the border of the object. 
@property (copy, readonly) ExcelRange *bottomRightCell;  // Returns the bottom right cell of the range the control is occupying.
@property (copy) NSString *caption;  // Returns or sets the caption for this object.
@property (copy) NSString *controlText;  // Returns or sets the default text for the control. 
@property BOOL displayThreeDShading;
@property BOOL enabled;  // Returns or sets if the object is enabled.
@property (readonly) NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property (copy, readonly) ExcelGroupbox *groupBox;
@property double height;  // Returns or set the height of the object.
@property (copy, readonly) ExcelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property double leftPosition;  // Returns or sets the position of the specified object, in points.
@property (copy) NSString *linkedCell;
@property BOOL locked;  // Returns or sets if the object is locked, if false the object can be modified when the sheet is protected.
@property BOOL lockedText;  // Returns or sets whether the control's text is locked for editing.
@property (copy) NSString *name;  // Returns or sets the name of the object.
@property (copy) NSString *onAction;  // Returns or sets the name of a string of AppleScript commands that will be executed when the object is clicked on. Please note that if AppleScript commands are set they will not be saved with the document.
@property (copy) NSString *phoneticAccelerator;  // Returns or sets the link to a range.
@property ExcelXlPlacement placement;  // Returns or sets how the object is placed on the worksheet.
@property BOOL printObject;  // Returns or sets if this object is printed.
@property double top;  // Returns the top position of the specified object, in points.
@property (copy, readonly) ExcelRange *topLeftCell;  // Returns a range object that represents the cell that lies under the upper-left corner of the specified object. 
@property ExcelXlCheckBoxState value;
@property BOOL visible;  // Returns or sets if the object is visible.
@property double width;  // Returns or sets  the width of the object.
@property (readonly) NSInteger zOrderPosition;  // Returns the z-order position of the object.


@end

// Represents an outline on a worksheet.
@interface ExcelOutline : SBObject <ExcelGenericMethods>

@property BOOL automaticStyles;  // Returns or sets if the outline uses automatic styles.
@property ExcelXlSummaryColumn summaryColumn;  // Returns or sets the location of the summary columns in the outline, as shown in the following table.
@property ExcelXlSummaryRow summaryRow;  // Returns or sets the location of the summary rows in the outline, as shown in the following table.

- (void) showLevelsRowLevels:(NSInteger)rowLevels columnLevels:(NSInteger)columnLevels;  // Displays the specified number of row and/or column levels of an outline.

@end

// Represents the page setup description. The page setup object contains all page setup attributes, left margin, bottom margin, paper size, and so on, as properties.
@interface ExcelPageSetup : SBObject <ExcelGenericMethods>

@property BOOL blackAndWhite;  // Returns or sets if elements of the document will be printed in black and white.
@property double bottomMargin;  // Returns or sets the size of the bottom margin, in points.
@property (copy) NSString *centerFooter;  // Returns or sets the center part of the footer.
@property (copy, readonly) ExcelGraphic *centerFooterPicture;  // Returns or sets a graphic object that represents the picture for the center section of the footer. Used to set attributes about the picture.
@property (copy) NSString *centerHeader;  // Returns or sets the center part of the header.
@property (copy, readonly) ExcelGraphic *centerHeaderPicture;  // Returns or sets a graphic object that represents the picture for the center section of the footer. Used to set attributes about the picture.
@property BOOL centerHorizontally;  // Returns or sets  if the sheet is centered horizontally on the page when it's printed.
@property BOOL centerVertically;  // Returns or sets if the sheet is centered vertically on the page when it's printed.
@property ExcelXlObjectSize chartSize;  // Returns or sets the way a chart is scaled to fit on a page.
@property BOOL draft;  // Returns or sets if the sheet will be printed without graphics.
@property NSInteger firstPageNumber;  // Returns or sets the first page number that will be used when this sheet is printed.
@property NSInteger fitToPagesTall;  // Returns or sets the number of pages tall the worksheet will be scaled to when it's printed. Applies only to worksheets.
@property NSInteger fitToPagesWide;  // Returns or sets the number of pages wide the worksheet will be scaled to when it's printed. Applies only to worksheets.
@property double footerMargin;  // Returns or sets the distance from the bottom of the page to the footer, in points.
@property double headerMargin;  // Returns or sets the distance from the top of the page to the header, in points.
@property (copy) NSString *leftFooter;  // Returns or sets the left part of the footer.
@property (copy, readonly) ExcelGraphic *leftFooterPicture;  // Returns or sets a graphic object that represents the picture for the left section of the footer. Used to set attributes about the picture.
@property (copy) NSString *leftHeader;  // Returns or sets the left part of the header.
@property (copy, readonly) ExcelGraphic *leftHeaderPicture;  // Returns or sets a graphic object that represents the picture for the left section of the header. Used to set attributes about the picture.
@property double leftMargin;  // Returns or sets the size of the left margin, in points.
@property ExcelXlOrder order;  // Returns or sets the order that Microsoft Excel uses to number pages when printing a large worksheet.
@property ExcelXlPageOrientation pageOrientation;  // Returns or set if the printing mode will be portrait or landscape.
@property ExcelXlPrintLocation printExcelComments;  // Returns or sets the way comments are printed with the sheet.
@property (copy) NSString *printArea;  // Returns or sets the range to be printed, as a string using A1-style references in the language of the macro.
@property BOOL printGridlines;  // Returns or sets if cell gridlines are printed on the page. Applies only to worksheets.
@property BOOL printHeadings;  // Returns or sets if row and column headings are printed with this page. Applies only to worksheets.
@property BOOL printNotes;  // Returns or sets if cell notes are printed as end notes with the sheet. Applies only to worksheets.
@property (copy) NSArray<id> *printQuality;  // Set/Gets a two element list where 1 is the horizontal print quality and 2 is the vertical print quality
@property (copy) NSString *printTitleColumns;  // Returns or sets the columns that contain the cells to be repeated on the left side of each page, as a string in A1-style notation in the language of the macro.
@property (copy) NSString *printTitleRows;  // Returns or sets the rows that contain the cells to be repeated at the top of each page, as a string in A1-style notation in the language of the macro.
@property (copy) NSString *rightFooter;  // Returns or sets the right part of the footer.
@property (copy, readonly) ExcelGraphic *rightFooterPicture;  // Returns or sets a graphic object that represents the picture for the right section of the footer. Used to set attributes about the picture.
@property (copy) NSString *rightHeader;  // Returns or sets the right part of the header.
@property (copy, readonly) ExcelGraphic *rightHeaderPicture;  // Returns or sets a graphic object that represents the picture for the right section of the header. Used to set attributes about the picture.
@property double rightMargin;  // Returns or sets the size of the right margin, in points.
@property double topMargin;  // Returns or sets the size of the top margin, in points.
@property ExcelXLSourceData zoom;  // Returns or sets a percentage, between 10 and 400 percent, by which Microsoft Excel will scale the worksheet for printing. Applies only to worksheets.


@end

// Represents a pane of a window. Pane objects exist only for worksheets and Microsoft Excel 4.0 macro sheets.
@interface ExcelPane : SBObject <ExcelGenericMethods>

@property (readonly) NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property NSInteger scrollColumn;  // Returns or sets the number of the leftmost column in the pane
@property NSInteger scrollRow;  // Returns or sets the number of the row that appears at the top of the pane.
@property (copy, readonly) ExcelRange *visibleRange;  // Returns a Range object that represents the range of cells that are visible in the pane. If a column or row is partially visible, it's included in the range.


@end

// Contains information about a specific phonetic text string in a cell.
@interface ExcelPhonetic : SBObject <ExcelGenericMethods>

@property ExcelXlPhoneticCharacterType characterType;  // Returns or sets the type of phonetic text in the specified cell.
@property (copy, readonly) ExcelFont *fontObject;  // Returns a font object that represents the font of the specified object.
@property ExcelXlPhoneticAlignment phoneticAlignment;  // Returns or sets the alignment for the specified phonetic text.
@property (copy) NSString *phoneticText;  // Returns or sets the text for the specified object.
@property BOOL visible;  // Returns or sets if the object is visible


@end

// Used as the base class for the PivotDataAxis, PivotFilterAxis, and PivotGroupAxis objects. There are no properties or methods which with you can use the PivotAxis object directly.
@interface ExcelPivotAxis : SBObject <ExcelGenericMethods>

- (SBElementArray<ExcelPivotLine *> *) pivotLines;


@end

// Represents the memory cache for a PivotTable report.
@interface ExcelPivotCache : SBObject <ExcelGenericMethods>

@property (copy, readonly) id ADOConnection;  // Returns an ADO connection object if the PivotTable cache is connected to an OLE DB data source.
@property (readonly) BOOL OLAP;  // Returns True if the PivotTable cache is connected to an Online Analytical Processing server.
@property (copy) NSString *SQLQuery;  // Returns or sets the SQL query string used with the specified ODBC data source.
@property BOOL backgroundQuery;  // Returns or sets if queries for the pivot table report or query table are performed asynchronously, in the background.
@property (copy) NSString *commandText;  // Returns or sets the command string for the specified data source.
@property ExcelXlCmdType commandType;  // Returns or sets a constant describing the value type of the CommandText property.
@property (copy) NSString *connection;  // Returns or sets a string that contains one of the following. ODBC settings that enable Microsoft Excel to connect to an ODBC data source, a URL that enables Microsoft Excel to connect to a Web data source, or a file that specifies a database or Web query.
@property BOOL enableRefresh;  // Returns or sets if the pivot table cache can be refreshed by the user.
@property (readonly) NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property (readonly) BOOL isConnected;  // Returns True if the MaintainConnection property is True and the PivotTable cache is currently connected to its source.
@property (copy) NSString *localConnection;  // Returns or sets the connection string to an offline cube file.
@property BOOL maintainConnection;  // True if the connection to the specified data source is maintained after the refresh and until the workbook is closed.
@property (readonly) NSInteger memoryUsed;  // Returns the amount of memory currently being used by the object, in bytes.
@property ExcelXlPivotTableMissingItems missingItemsLimit;  // Returns or sets the maximum quantity of unique items per PivotTable field that are retained even when they have no supporting data in the cache records.
@property BOOL optimizeCache;  // Returns or set if the pivot table cache is optimized when it's constructed.
@property (readonly) ExcelXlQueryType queryType;  // Indicates the type of query used by Microsoft Excel to populate the PivotTable cache.
@property (readonly) NSInteger recordCount;  // Returns the number of records in the pivot table cache or the number of cache records that contain the specified item.
@property (copy, readonly) NSDate *refreshDate;  // Returns the date on which the pivot cache was last refreshed.
@property (copy, readonly) NSString *refreshName;  // Returns the name of the person who last refreshed the pivot cache. 
@property BOOL refreshOnFileOpen;  // Returns or sets if the pivot table cache or query table is automatically updated each time the workbook is opened.
@property NSInteger refreshPeriod;  // Returns or sets the number of minutes between refreshes.
@property ExcelXlRobustConnect robustConnect;  // Returns or sets how the PivotTable cache connects to its data source.
@property BOOL savePassword;  // Returns or set if password information in an ODBC connection string is saved with the specified query. if false, the password is removed.
@property (copy) NSString *sourceConnectionFile;  // Returns or sets a String indicating the Microsoft Office Data Connection file or similar file that was used to create the PivotTable.
@property ExcelXLSourceDataLocation sourceData;  // Returns or sets the data source for the pivot table.
@property (readonly) ExcelXlPivotTableSourceType sourceType;  // Returns a value that identifies the type of item being published.
@property BOOL upgradeOnRefresh;  // Contains information on whether to upgrade the PivotCache and all connected PivotTables on the next refresh.
@property BOOL useLocalConnection;  // True if the LocalConnection property is used to specify the string that enables Microsoft Excel to connect to a data source.
@property (readonly) ExcelXlPivotTableVersionList version;  // Returns the version of Microsoft Excel in which the PivotCache was created.
@property (copy, readonly) id workbookConnection;  // Establishes a connection between the current workbook and the PivotCache object.

- (void) createPivotTableTableDestination:(NSString *)tableDestination tableName:(NSString *)tableName readData:(NSString *)readData defaultVersion:(NSString *)defaultVersion;  // Creates a PivotTable report based on a PivotCache object.
- (void) refresh;  // Updates the pivot table cache or query table.

@end

// Represents a cell in a PivotTable report.
@interface ExcelPivotCell : SBObject <ExcelGenericMethods>

@property (copy, readonly) NSString *MDX;
@property (readonly) ExcelXlCellChangedState cellChanged;
@property (readonly) ExcelXlConsolidationFunction customSubtotalFunction;  // Returns the custom subtotal function field setting of a PivotCell object.
@property (copy, readonly) ExcelPivotField *dataField;  // Returns a PivotField object that corresponds to the selected data field.
@property (copy, readonly) id dataSourceValue;
@property (readonly) ExcelXlPivotCellType pivotCellType;  // Returns a constant that identifies the PivotTable entity the cell corresponds to.
@property (copy, readonly) ExcelPivotLine *pivotColumnLine;  // Returns the PivotLine on a column for a specific PivotCell object.
@property (copy, readonly) ExcelPivotField *pivotField;  // Returns a PivotField object that represents the PivotTable field containing the upper-left corner of the specified range.
@property (copy, readonly) ExcelPivotItem *pivotItem;  // Returns a PivotItem object that represents the PivotTable item containing the upper-left corner of the specified range.
@property (copy, readonly) ExcelPivotLine *pivotRowLine;  // Returns the PivotLine on a row for a specific PivotCell object.
@property (copy, readonly) ExcelPivotTable *pivotTable;  // Returns a PivotTable object that represents the PivotTable report associated with the PivotCell.
@property (copy, readonly) ExcelRange *range;  // Returns a Range object that represents the range the specified PivotCell applies to.
@property (copy) id rowItems;  // Returns a PivotItemList collection that corresponds to the items on the category axis that represent the selected cell.

- (void) allocateChange;
- (void) discardChange;

@end

// Represents a field in a pivot table.
@interface ExcelPivotField : SBObject <ExcelGenericMethods>

- (SBElementArray<ExcelChildItem *> *) childItems;
- (SBElementArray<ExcelHiddenItem *> *) hiddenItems;
- (SBElementArray<ExcelParentItem *> *) parentItems;
- (SBElementArray<ExcelPivotItem *> *) pivotItems;
- (SBElementArray<ExcelCalculatedItem *> *) calculatedItems;
- (SBElementArray<ExcelPivotFilter *> *) pivotFilters;

@property (readonly) BOOL allItemsVisible;  // Used to retrieve a Boolean value that indicates whether or not any manual filtering is applied to the PivotField.
@property (readonly) NSInteger autoShowCount;  // Returns the number of top or bottom items that are automatically shown in the pivot field.
@property (copy, readonly) NSString *autoShowField;  // Returns the name of the data field used to determine the top or bottom items that are automatically shown in the pivot field.
@property (readonly) ExcelXLAutoShowPosition autoShowRange;  // Returns position top if the top items are shown automatically in the pivot field, returns position bottom if the bottom items are shown.
@property (readonly) ExcelXLAutoShowType autoShowType;  // Returns type_automatic if auto show is enabled for the pivot field, or  returns type_manual if auto show is disabled.
@property (readonly) NSInteger autoSortCustomSubtotal;  // Returns the name of the custom subtotal used to sort the specified PivotTable field automatically.
@property (copy, readonly) NSString *autoSortField;  // Returns the name of the data field used to sort the pivot field automatically.
@property (readonly) ExcelXlSortOrder autoSortOrder;  // Returns the order used to sort the pivot field automatically. 
@property (copy, readonly) ExcelPivotLine *autoSortPivotLine;  // Returns the name of the PivotLine used to sort the specified PivotTable field automatically.
@property (copy) NSString *baseField;  // Returns or sets the base field for a custom calculation.
@property (copy) NSString *baseItem;  // Returns or sets the item in the base field for a custom calculation.
@property ExcelXlPivotFieldCalculation calculation;  // Returns or sets the type of calculation done by the specified pivot field.
@property (copy) NSString *caption;  // The label text for the pivot field.
@property (copy, readonly) ExcelPivotField *childField;  // Returns a pivot field object that represents the child pivot field for the specified field, if the field is grouped and has a child field.
@property (copy, readonly) ExcelCubeField *cubeField;  // Returns the CubeField object from which the specified PivotTable field is descended.
@property (copy) id currentPage;  // Returns or sets the current page showing for the page field, only valid for page fields.
@property (copy) id currentPageList;  // Returns or sets an array of strings corresponding to the list of items included in a multiple-item page field of a PivotTable report.
@property (copy) NSString *currentPageName;  // Returns or sets the currently displayed page of the specified PivotTable report.
@property (copy, readonly) ExcelRange *dataRange;  // Returns a range object.  For a data field the range is the data contained in the field, for a row, column or page field it is the items in the field.
@property BOOL databaseSort;  // When set to True, manual repositioning of items in a PivotTable field is allowed.
@property (readonly) BOOL displayAsCaption;  // This property is used to display member properties of PivotFields as captions.
@property BOOL displayAsTooltip;  // This property is used to specify whether or not a specific member property PivotField is displayed in tooltips.
@property BOOL displayInReport;  // This property is used to specify whether the specified member property PivotField is displayed in the PivotTable or not.
@property BOOL dragToColumn;  // Returns or sets if the pivot field can be dragged to the column position.
@property BOOL dragToData;  // Returns or sets if the field can be dragged to the data position.
@property BOOL dragToHide;  // Returns or sets if the field can be hidden by being dragged off the pivot table.
@property BOOL dragToPage;  // Returns or sets if the field can be dragged to the page position.
@property BOOL dragToRow;  // Returns or sets if the field can be dragged to the row position.
@property BOOL drilledDown;  // True if the flag for the specified PivotTable field or PivotTable item is set to drilled.
@property BOOL enableItemSelection;  // When set to False, disables the ability to use the field dropdown in the user interface.
@property BOOL enableMultiplePageItems;  // Used for specifying whether or not check boxes are present in the filter drop-down list for fields in the page area.
@property (copy) NSString *formula;  // Returns or sets the object's formula, in A1-style notation and in the language of the macro.
@property ExcelXlConsolidationFunction function;  // Returns or sets the function used to summarize the pivot field.
@property (readonly) NSInteger groupLevel;  // Returns the placement of the specified field within a group of fields, if the field is a member of a grouped set of fields.
@property BOOL hidden;  // This property is used to hide the individual levels of an OLAP hierarchy.
@property (copy) id hiddenItemsList;  // Returns or sets an Object specifying an array of strings that are hidden items for a PivotTable field.
@property BOOL includeNewItemsInFilter;  // Allows developers to specify whether excluded or included items should be tracked when manual filtering is applied to the PivotField.
@property (readonly) BOOL isCalculated;  // Returns true if the pivot field or item is a calculated field or item.
@property (readonly) BOOL isMemberProperty;  // Returns True when the PivotField contains member properties.
@property (copy, readonly) ExcelRange *labelRange;  // Returns a range object that represents the cell, or cells, that contain the field label.
@property BOOL layoutBlankLine;  // True if a blank row is inserted after the specified row field in a PivotTable report.
@property BOOL layoutCompactRow;  // Specifies whether or not a PivotField is compacted , items of multiple PivotFields are displayed in a single column, when rows are selected.
@property ExcelXlLayoutFormType layoutForm;  // Returns or sets the way the specified PivotTable items appear.
@property BOOL layoutPagebreak;  // True if a page break is inserted after each field.
@property ExcelXlSubtototalLocationType layoutSubtotalLocation;  // Returns or sets the position of the PivotTable field subtotals in relation to either above or below the specified field.
@property (copy) NSString *memberPropertyCaption;  // Setting the MemberPropertyCaption property controls which member property is used as caption for a given level.
@property (readonly) NSInteger memoryUsed;  // Returns the amount of memory currently being used by the object, in bytes.
@property (copy) NSString *name;  // Returns or sets the name of the pivot field.
@property (copy) NSString *numberFormat;  // Returns or sets the format code for the pivot field. 
@property (copy, readonly) ExcelPivotField *parentField;  // Returns a pivot field object that represents the pivot field that's the group parent of the object. The field must be grouped and have a parent field.
@property (readonly) ExcelXlPivotFieldDataType pivotFieldDataType;  // Returns an enumeration describing the type of data in the pivot field.
@property ExcelXlPivotFieldOrientation pivotFieldOrientation;  // The location of the field in the pivot table.
@property NSInteger position;  // Position of the field, first, second, third, and so on, among all the fields in its orientation, rows, columns, pages, data.
@property NSInteger propertyOrder;  // Valid only for PivotTable fields that are member property fields.
@property (copy, readonly) ExcelPivotField *propertyParentField;  // Returns a PivotField object representing the field to which the properties in this field pertain.
@property BOOL repeatLabels;
@property BOOL serverBased;  // Returns or sets if the pivot table's data source is external and only the items matching the page field selection are retrieved.
@property BOOL showAllItems;  // Returns or sets if all items in the pivot table are displayed, even if they don't contain summary data.
@property BOOL showDetail;  // Gets or sets whether the specified PivotField is showing detail.
@property (readonly) BOOL showingInAxis;  // Indicates if the PivotField is currently visible in the PivotTable or not.
@property (copy, readonly) NSString *sourceCaption;  // The SourceCaption property is applicable only for OLAP PivotTables, and returns the original caption from the OLAP server for a PivotField.
@property (copy, readonly) NSString *sourceName;  // Return the specified object's name, as it appears in the original source data for the pivot table. This might be different from the current item name if the user renamed the item after creating the pivot table.
@property (copy) NSString *standardFormula;  // Returns or sets a String specifying formulas with standard English formatting.
@property (copy) NSString *subtotalName;  // Returns or sets the text string label displayed in the subtotal column or row heading in the specified PivotTable report.
@property (readonly) NSInteger totalLevels;  // Returns the total number of fields in the current field group. If the field isn't grouped, or if the data source is OLAP-based, TotalLevels returns the value 1.
@property BOOL useMemberPropertyAsCaption;  // This property is used to control whether member property captions are used for PivotItem captions of the PivotField.
@property (copy) NSString *value;  // Returns or sets the name of the specified field in the pivot table.
@property (copy, readonly) NSArray<SBObject *> *visibleItems;  // Returns a list of all the visible pivot items in the specified field.
@property (copy) id visibleItemsList;  // Returns or sets an array of strings corresponding to the list of items included in a multiple-item page field of a PivotTable report.

- (void) addPageItemItem:(NSString *)item clearList:(NSString *)clearList;  // Adds an additional item to a multiple item page field.
- (void) autoShowType:(ExcelXLAutoShowType)type range:(ExcelXLAutoShowPosition)range count:(NSInteger)count field:(NSString *)field;  // Displays the number of top or bottom items for a row, page, or column field in the specified PivotTable report
- (void) autoSortSortOrder:(ExcelXlSortOrder)sortOrder sortField:(NSString *)sortField;  // Establishes automatic pivot table field-sorting rules.
- (void) clearLabelFilters;  // This method deletes all label filters or all date filters in the PivotFilters collection of the PivotField.
- (void) clearValueFilters;  // Calling this method deletes all value filters in the PivotFilters collection of the PivotField.
- (BOOL) getSubtotalsSubtotalIndex:(ExcelXLSubTotalsIndex)subtotalIndex;  // Returns subtotals displayed with the specified field. Valid only for nondata fields.
- (void) setSubtotalsSubtotalIndex:(ExcelXLSubTotalType)subtotalIndex value:(BOOL)value;  // Sets subtotals displayed with the specified field. Valid only for nondata fields.

@end

@interface ExcelCalculatedField : ExcelPivotField


@end

@interface ExcelColumnField : ExcelPivotField


@end

@interface ExcelDataField : ExcelPivotField


@end

@interface ExcelHiddenField : ExcelPivotField


@end

@interface ExcelPageField : ExcelPivotField


@end

// PivotTable Advanced Filter.
@interface ExcelPivotFilter : SBObject <ExcelGenericMethods>

@property (readonly) BOOL active;  // Returns whether the specified PivotFilter is active.
@property (copy, readonly) ExcelCubeField *dataCubeField;  // This property is applicable only to non-OLAP PivotTables and provides the Value field ,PivotField in the Values area, being filtered by for a value filter.
@property (copy, readonly) ExcelPivotField *dataField;  // This property is applicable only to non-OLAP PivotTables and provides the Value field ,PivotField in the Values area, being filtered by for a value filter.
@property (copy, readonly) NSString *objectDescription;  // Provides an optional description for the PivotFilter object.
@property (readonly) ExcelXlPivotFilterType filterType;  // Specifies the type of filter to be applied.
@property (readonly) BOOL isMemberPropertyFilter;  // Specifies whether the label filter is based on the PivotItem captions of a member property of the field or on the PivotItem captions of the PivotField itself.
@property (copy, readonly) ExcelPivotField *memberPropertyField;  // This property specifies the member property PivotField on which the label filter is based.
@property (copy, readonly) NSString *name;  // This property provides the option of naming filters for reference.
@property NSInteger order;  // Specifies the evaluation order of the filter among all Value filters applied to the entire PivotTable.
@property (copy, readonly) ExcelPivotField *pivotField;  // Specifies the PivotField to which the filter is applied.
@property (copy, readonly) id value1;  // This property is a user-supplied parameter to define a filter for a PivotField.
@property (copy, readonly) id value2;  // This property is a user-supplied parameter to define a filter for a PivotField.


@end

@interface ExcelActiveFilter : ExcelPivotFilter


@end

// Represents a formula used to calculate results in a pivot table report.
@interface ExcelPivotFormula : SBObject <ExcelGenericMethods>

@property NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property (copy) NSString *formula;  // Returns or sets the object's formula, in A1-style notation and in the language of the macro.
@property (copy) NSString *standardFormula;  // Returns or sets a String specifying formulas with standard English formatting.
@property (copy) NSString *value;  // Returns or sets the name of the specified formula in the pivot table formula.


@end

// Represents an item in a pivot field. The items are the individual data entries in a field category.
@interface ExcelPivotItem : SBObject <ExcelGenericMethods>

- (SBElementArray<ExcelChildItem *> *) childItems;

@property (copy) NSString *caption;  // The label text for the pivot item.
@property (copy, readonly) ExcelRange *dataRange;  // Returns a range object specifying the data qualified by the item.
@property BOOL drilledDown;  // True if the flag for the specified PivotTable field or PivotTable item is set to 
@property (copy) NSString *formula;  // Returns or sets the pivot item's formula, in A1-style notation and in the language of the macro.
@property (readonly) BOOL isCalculated;  // Returns true if the pivot  item is a calculated item.
@property (copy, readonly) ExcelRange *labelRange;  // Returns a range object that represents all the pivot table cells that contain the item.
@property (copy) NSString *name;  // Returns or sets the name of the pivot item.
@property (copy, readonly) ExcelParentItem *parentItem;  // Returns a pivot item object that represents the parent pivot item in the parent pivot field object, the field must be grouped so that it has a parent.
@property (readonly) BOOL parentShowDetail;  // Return true if the specified item is showing because one of its parents is showing detail. False if the specified item isn't showing because one of its parents is hiding detail. This property is available only if the item is grouped.
@property NSInteger position;  // Returns or sets the position of the item in its field, if the item is currently showing.
@property (readonly) NSInteger recordCount;  // Returns the number of records in the pivot table cache or the number of cache records that contain the specified item.
@property BOOL showDetail;  // Returns  or sets if the pivot item is showing detail.
@property (copy, readonly) NSString *sourceName;  // Return the specified object's name, as it appears in the original source data for the pivot table. This might be different from the current item name if the user renamed the item after creating the pivot table.
@property (copy, readonly) NSString *sourceNameStandard;  // Returns a String that represents the PivotTable item's source name in standard English format settings.
@property (copy) NSString *standardFormula;  // Returns or sets a String specifying formulas with standard English formatting.
@property (copy) NSString *value;  // Returns or sets set the name of the specified item in the pivot table field.
@property BOOL visible;  // Returns or sets if the pivot item is visible.


@end

@interface ExcelCalculatedItem : ExcelPivotItem


@end

@interface ExcelChildItem : ExcelPivotItem


@end

@interface ExcelColumnItem : ExcelPivotItem


@end

@interface ExcelHiddenItem : ExcelPivotItem


@end

@interface ExcelParentItem : ExcelPivotItem


@end

// The PivotLines object is a collection of lines in a PivotTable, containing all lines on rows or columns of the pivot.
@interface ExcelPivotLine : SBObject <ExcelGenericMethods>

@property (readonly) ExcelXlPivotLineType lineType;  // Returns an XlPivotLineType constant that indicates the type of PivotLine.
@property (copy, readonly) id pivotLineCells;  // Returns a collection of PivotCell objects in a PivotLine.
@property (readonly) NSInteger position;  // Returns or sets the position of the PivotLine object.


@end

// Represents a pivot table on a worksheet.
@interface ExcelPivotTable : SBObject <ExcelGenericMethods>

- (SBElementArray<ExcelColumnField *> *) columnFields;
- (SBElementArray<ExcelDataField *> *) dataFields;
- (SBElementArray<ExcelHiddenField *> *) hiddenFields;
- (SBElementArray<ExcelPageField *> *) pageFields;
- (SBElementArray<ExcelPivotField *> *) pivotFields;
- (SBElementArray<ExcelRowField *> *) rowFields;
- (SBElementArray<ExcelCalculatedField *> *) calculatedFields;
- (SBElementArray<ExcelCalculatedMember *> *) calculatedMembers;
- (SBElementArray<ExcelActiveFilter *> *) activeFilters;
- (SBElementArray<ExcelCubeField *> *) cubeFields;
- (SBElementArray<ExcelSlicer *> *) slicers;

@property NSInteger CompactRowIndent;  // Returns or sets the indent increment for PivotItems when compact row layout form is turned on.
@property ExcelXlAllocation allocation;
@property ExcelXlAllocationMethod allocationMethod;
@property ExcelXlAllocationValue allocationValue;
@property (copy) NSString *allocationWeightExpression;
@property BOOL allowMultipleFilters;  // Sets or retrieves a value that indicates whether a PivotField can have multiple filters applied to it at the same time.
@property (copy) NSString *alternativeText;  // Returns or sets the descriptive alternative text string for a ShapeRange object when the object is saved to a Web page.
@property NSInteger cacheIndex;  // Returns or sets the index number of the pivot table cache.
@property BOOL calculatedMembersInFilters;
@property (copy, readonly) id changeList;
@property BOOL columnGrand;  // Returns or sets if the pivot table shows grand totals for columns.
@property (copy, readonly) ExcelRange *columnRange;  // Returns a range object that represents the range that contains the pivot table column area.
@property (copy) NSString *compactLayoutColumnHeader;  // Specifies the caption that is displayed in the column header of a PivotTable when in compact row layout form.
@property (copy) NSString *compactLayoutRowHeader;  // Specifies the caption that is displayed in the row header of a PivotTable when in compact row layout form.
@property (copy, readonly) ExcelRange *dataBodyRange;  // Returns a range object that represents the range that contains the pivot table data area.
@property (copy, readonly) ExcelRange *dataLabelRange;  // Returns a range object that represents the range that contains the labels for the pivot table data fields. 
@property (copy, readonly) ExcelPivotField *dataPivotField;  // Returns a PivotField object that represents all the data fields in a PivotTable.
@property BOOL displayContextTooltips;  // Controls whether or not tooltips are displayed for PivotTable cells.
@property BOOL displayEmptyColumn;  // Returns True when the nonempty MDX keyword is included in the query to the OLAP provider for the value axis.
@property BOOL displayEmptyRow;  // Returns True when the nonempty MDX keyword is included in the query to the OLAP provider for the category axis.
@property BOOL displayErrorString;  // Returns or sets if the pivot table displays a custom error string in cells that contain errors.
@property BOOL displayFieldCaptions;  // Controls whether or not filter buttons and PivotField captions for rows and columns are displayed in the grid.
@property BOOL displayImmediateItems;  // Returns or sets a Boolean that indicates whether items in the row and column areas are visible when the data area of the PivotTable is empty.
@property BOOL displayMemberPropertyTooltips;  // Controls whether or not to display member properties in tooltips.
@property BOOL displayNullString;  // Returns or sets if the pivot table displays a custom string in cells that contain null values.
@property BOOL enableDataValueEditing;  // True to disable the alert for when the user overwrites values in the data area of the PivotTable.
@property BOOL enableDrilldown;  // Returns or sets if drilldown is enabled.
@property BOOL enableFieldDialog;  // Returns or sets if the pivot table field dialog box is available when the user double-clicks the pivot table field.
@property BOOL enableFieldList;  // False to disable the ability to display the field well for the PivotTable.
@property BOOL enableWizard;  // Returns or sets if the pivot table wizard is available.
@property BOOL enableWriteback;
@property (copy) NSString *errorString;  //  Returns or sets the string displayed in cells that contain errors when the display error string property is true. The default value is an empty string.
@property BOOL fileListSortAscending;  // Controls the sort order of fields in the PivotTable Field List.
@property (copy) NSString *grandTotalName;  // Returns or sets the text string label that is displayed in the grand total column or row heading in the specified PivotTable report
@property BOOL hasAutoformat;  // Returns or sets if the pivot table is automatically formatted when it's refreshed or when fields are moved.
@property BOOL inGridDropZones;  // This property is used to toggle in-grid drop zones for a PivotTable object.
@property (copy) NSString *innerDetail;  // Returns or sets the name of the field that will be shown as detail when the show detail property is true for the innermost row or column field.
@property ExcelXlLayoutRowType layoutRowDefault;  // This property specifies the layout settings for PivotFields when they are added to the PivotTable for the first time.
@property (copy) NSString *location;  // Gets or sets a String that represents the top-left cell in the body of the specified PivotTable.
@property BOOL manualUpdate;  // Returns or sets if the pivot table is recalculated only at the user's request.
@property (copy, readonly) NSString *mdx;  // Returns a String indicating the MDX, Multidimensional-Expression, that would be sent to the provider to populate the current PivotTable view.
@property BOOL mergeLabels;  // Returns or sets if pivot table outer-row item, column item, subtotal, and grand total labels use merged cells.
@property (copy) NSString *name;  // Returns or sets the name of the object.
@property (copy) NSString *nullString;  // Returns or sets the string displayed in cells that contain null values when the display null string property is true. The default value is an empty string.
@property ExcelXlOrder pageFieldOrder;  // Returns or sets the order in which page fields are added to the pivot table layout.
@property (copy) NSString *pageFieldStyle;  // Returns or sets the style used in the bound page field area.
@property NSInteger pageFieldWrapCount;  // Returns or sets the number of pivot table page fields in each column or row.
@property (copy, readonly) ExcelRange *pageRange;  // Returns a range object that represents the range that contains the pivot table page area.
@property (copy, readonly) ExcelRange *pageRangeCells;  // Returns a range object that represents the cells in the pivot table containing only the page fields and item drop-down lists.
@property (copy, readonly) ExcelPivotCache *pivotCache;  // Returns a pivot cache object that represents the cache for the specified pivot table
@property (copy, readonly) ExcelPivotAxis *pivotColumnAxis;  // Returns a PivotAxis object representing the entire column axis 
@property (copy, readonly) ExcelPivotAxis *pivotRowAxis;  // Returns a PivotAxis object representing the entire row axis 
@property (copy) NSString *pivotSelection;  // Returns or sets the pivot table selection, in standard pivot table selection format.
@property (copy) NSString *pivotSelectionStandard;  // Returns or sets a String indicating the PivotTable selection in standard PivotTable report format using English settings.
@property BOOL preserveFormatting;  // Returns or sets if pivot table formatting is preserved when the pivot table is refreshed or recalculated by operations such as pivoting, sorting, or changing page field items.
@property BOOL printDrillIndicators;  // Specifies whether or not drill indicators are printed with the PivotTable.
@property BOOL printTitles;  // set/get whether the print titles for the worksheet are set based on the PivotTable report.
@property (copy, readonly) NSDate *refreshDate;  // Returns the date on which the pivot table was last refreshed.
@property (copy, readonly) NSString *refreshName;  // Returns the name of the person who last refreshed the pivot table data.
@property BOOL repeatItemsOnEachPrintedPage;  // True if row, column, and item labels appear on the first row of each page when the specified PivotTable report is printed. False if labels are printed only on the first page.
@property BOOL rowGrand;  // Returns or sets if the pivot table shows grand totals for rows.
@property (copy, readonly) ExcelRange *rowRange;  // Returns a range object that represents the range including the pivot table row area.
@property BOOL saveData;  // Returns or sets if data for the pivot table is saved with the workbook. If false only the pivot table definition is saved.
@property ExcelXlPTSelectionMode selectionMode;  // Returns or sets the pivot table structured selection mode.
@property BOOL showDrillIndicators;  // When the ShowDrillIndicators property is set to False, the PrintDrillIndicators property has no effect.
@property BOOL showPageMultipleLabel;  // When set to True, Multiple Items will appear in the PivotTable cell on the worksheet whenever items are hidden and an aggregate of nonhidden items is shown in the PivotTable view.
@property BOOL showTableStyleColumnHeaders;  // The ShowTableStyleColumnHeaders property is set to True if the coulmn headers should be displayed in the PivotTable.
@property BOOL showTableStyleColumnStripes;  // The ShowTableStyleColumnStripes property displays banded columns in which even columns are formatted differently from odd columns.
@property BOOL showTableStyleLastColumn;  // Returns or sets if the last column is displayed for the specified ListObject object.
@property BOOL showTableStyleRowHeaders;  // The ShowTableStyleRowHeaders property is set to True if the row headers should be displayed in the PivotTable.
@property BOOL showTableStyleRowStripes;  // The ShowTableStyleRowStripes property displays banded rows in which even rows are formatted differently from odd rows.
@property BOOL showValuesRow;
@property BOOL smallGrid;  // Returns or sets if Microsoft Excel uses a grid that's two cells wide and two cells deep for a newly created pivot table report. False if Excel uses a blank stencil outline.
@property BOOL sortUsingCustomLists;  // The SortUsingCustomLists property controls whether custom lists are used for sorting items of fields.
@property ExcelXLSourceDataLocation sourceData;  // Either a range reference as a string or a list of SQL commands
@property BOOL subtotalHiddenPageItems;  // Returns or sets if hidden page field items in the pivot table are included in row and column subtotals, block totals, and grand totals.
@property (copy) NSString *summary;
@property (copy, readonly) ExcelRange *tableRange1;  // Returns a range object that represents the range containing the entire pivot table, but doesn't include page fields.
@property (copy, readonly) ExcelRange *tableRange2;  // Returns a range object that represents the range containing the entire pivot table, including page fields.
@property (copy) NSString *tableStyle;  // Returns or sets the style used in the pivot table body.
@property (copy) id tableStyle2;  // The TableStyle2 property specifies the PivotTable style currently applied to the PivotTable.
@property (copy) NSString *tag;  // Returns or sets a string saved with the pivot table.
@property BOOL totalsAnnotation;  // True if an asterisk is displayed next to each subtotal and grand total value in the specified PivotTable report if the report is based on an OLAP data source.
@property (copy) NSString *vacatedStyle;  // Returns or sets the style applied to cells vacated when the pivot table is refreshed.
@property (copy) NSString *value;  // Returns or set the name of the pivot table.
@property (readonly) ExcelXlPivotTableVersionList version;
@property BOOL viewCalculatedMembers;  // When set to True, calculated members for Online Analytical Processing, OLAP, PivotTables can be viewed.
@property BOOL visualTotals;  // True to enable Online Analytical Processing, OLAP, PivotTables to retotal after an item has been hidden from view.
@property BOOL visualTotalsForSets;

- (ExcelPivotField *) addDataFieldField:(id)field caption:(NSString *)caption function:(NSString *)function;  // Adds a data field to a PivotTable report.
- (void) addFieldsToPivotTableRowFields:(NSArray<NSString *> *)rowFields columnFields:(NSArray<NSString *> *)columnFields pageFields:(NSArray<NSString *> *)pageFields addToTable:(BOOL)addToTable;  // Adds row, column, and page fields to a pivot table.
- (void) allocateChanges;
- (void) changeConnectionConnection:(ExcelWorkbookConnection *)connection;  // Changes the connection of the specified PivotTable.
- (void) changePivotCachePivotCache:(NSString *)pivotCache;  // Changes the PivotCache of the specified PivotTable.
- (void) clearTable;  // The ClearTable method is used for clearing a PivotTable.
- (void) commitChanges;
- (void) convertToFormulasConvertFilters:(BOOL)convertFilters;  // The ConvertToFormulas method is new in 1st_Excel12 and is used for converting a PivotTable to cube formulas.
- (NSString *) createCubeFileFile:(NSString *)file measures:(NSString *)measures levels:(NSString *)levels members:(NSString *)members properties:(NSString *)properties;  // Creates a cube file from a PivotTable report connected to an Online Analytical Processing data source.
- (void) discardChanges;
- (ExcelRange *) getPivotDataDataField:(NSString *)dataField field1:(NSString *)field1 item1:(NSString *)item1 field2:(NSString *)field2 item2:(NSString *)item2 field3:(NSString *)field3 item3:(NSString *)item3 field4:(NSString *)field4 item4:(NSString *)item4 field5:(NSString *)field5 item5:(NSString *)item5 field6:(NSString *)field6 item6:(NSString *)item6 field7:(NSString *)field7 item7:(NSString *)item7 field8:(NSString *)field8 item8:(NSString *)item8 field9:(NSString *)field9 item9:(NSString *)item9 field10:(NSString *)field10 item10:(NSString *)item10 field11:(NSString *)field11 item11:(NSString *)item11 field12:(NSString *)field12 item12:(NSString *)item12 field13:(NSString *)field13 item13:(NSString *)item13 field14:(NSString *)field14 item14:(NSString *)item14;  // Returns a Range object with information about a data item in a PivotTable report.
- (double) getPivotTableDataName:(NSString *)name;  // Retrieve data from a pivot table
- (NSArray<SBObject *> *) getVisibleFields;  // Returns a list of all the visible pivot fields. Visible pivot fields are shown as row, column, page or data fields.
- (void) listFormulas;  // Creates a list of calculated pivot table items and fields on a separate worksheet.
- (void) makeConnection;  // Establishes a connection for the specified PivotTable cache.
- (void) pivotSelectName:(NSString *)name mode:(ExcelXlPTSelectionMode)mode;  // Selects part of a pivot table.
- (void) refreshDataSourceValues;
- (BOOL) refreshTable;  // Refreshes the pivot table from the source data. Returns true if it's successful.
- (void) repeatAllLabelsRepeat:(ExcelXlPivotFieldRepeatLabels)repeat;
- (void) rowAxisLayoutLayout:(ExcelXlLayoutRowType)layout;  // This method is used for simultaneously setting layout options for all existing PivotFields.
- (void) saveAsODCODCFileName:(NSString *)ODCFileName objectDescription:(NSString *)objectDescription keywords:(NSString *)keywords;  // Saves the PivotTable cache source as a Microsoft Office Data Connection file.
- (void) showPagesPageField:(NSString *)pageField;  // Creates a new pivot table for each item in the page field. Each new pivot table is created on a new worksheet.
- (void) subtotalLocationLocation:(ExcelXlSubtototalLocationType)location;  // This method changes the subtotal location for all existing PivotFields.
- (void) update;  // Updates the link or pivot table.

@end

// Represents a worksheet table built from data returned from an external data source, such as an SQL server.
@interface ExcelQueryTable : SBObject <ExcelGenericMethods>

@property BOOL adjustColumnWidth;  // Returns or sets if the column widths are automatically adjusted for the best fit each time you refresh the specified query table.
@property BOOL backgroundQuery;  // Returns or sets if queries for the query table are performed asynchronously.
@property ExcelXlCmdType commandType;  // Returns or sets one of the XlCmdType constants listed in the following table in the remarks section.
@property (copy) NSString *connection;  // Returns or sets a string that contains one of the following: ODBC settings that enable Microsoft Excel to connect to an ODBC data source; a URL that enables Microsoft Excel to connect to a Web data source; or a file that specifies a database or Web query.
@property (copy, readonly) ExcelRange *destination;  // Returns the cell in the upper-left corner of the query table destination range, the range where the resulting query table will be placed. The destination range must be on the worksheet that contains the query table object.
@property BOOL enableEditing;  // Returns or sets if the user can edit the specified query table. False if the user can only refresh the query table.
@property BOOL enableRefresh;  // Returns or sets if the query table can be refreshed by the user.
@property (readonly) BOOL fetchedRowOverflow;  // Returns true if the number of rows returned by the last use of the refresh method is greater than the number of rows available on the worksheet.
@property BOOL fieldNames;  // Returns or sets if field names from the data source appear as column headings for the returned data.
@property BOOL fillAdjacentFormulas;  // Returns or sets if formulas to the right of the specified query table are automatically updated whenever the query table is refreshed.
@property BOOL hasAutoformat;  // Returns or sets if the query table is automatically formatted when it's refreshed or when fields are moved.
@property (copy) NSString *name;  // Returns or sets the name of the object.
@property (copy) NSString *postText;  // Returns or sets the string used with the post method of inputting data into a Web server to return data from a Web query.
@property (readonly) ExcelXlQueryType queryType;  // Returns the type of query used by Microsoft Excel to populate the query table.
@property BOOL refreshOnFileOpen;  // Returns or sets if the query table is automatically updated each time the workbook is opened.
@property ExcelXlCellInsertionMode refreshStyle;  // Returns or sets the way rows on the specified worksheet are added or deleted to accommodate the number of rows in a recordset returned by a query.
@property (readonly) BOOL refreshing;  // Returns true if there's a background query in progress for the specified query table.
@property (copy, readonly) ExcelRange *resultRange;  // Returns a range object that represents the area of the worksheet occupied by the specified query table.
@property BOOL rowNumbers;  // Returns or sets if row numbers are added as the first column of the specified query table.
@property BOOL saveData;  // Returns or sets if data for the query table is saved with the workbook.
@property BOOL savePassword;  // Returns or sets if password information in an ODBC connection string is saved with the specified query.
@property (copy) NSString *sql;  // Returns or sets the SQL query string used with the specified ODBC data source.
@property BOOL tablesOnlyFromHtml;  // Returns or sets if only the HTML tables in the document are read when a query table is refreshed. False if the entire HTML document is read when a query table is refreshed. This property has an effect only when the query table is using a URL connection.
@property (copy) NSArray<NSAppleEventDescriptor *> *textFileColumnDataTypes;  // A list of enums: general format, text format, MDY format, DMY format, YMD format, MYD format, DYM format, YDM format, skip column
@property BOOL textFileCommaDelimiter;  // Returns or sets if the comma character is the delimiter when you import a text file into a query table.
@property BOOL textFileConsecutiveDelimiter;  // Returns or sets if consecutive delimiters are treated as a single delimiter when you import a text file into a query table.
@property (copy) NSString *textFileDecimalSeparator;  // Returns or sets the decimal separator character that Microsoft Excel uses when you import a text file into a query table. The default is the system decimal separator character.
@property (copy) NSArray<NSNumber *> *textFileFixedColumnWidths;  // Returns or sets a list of numbers that correspond to the widths of the columns in the text file that you are importing into a query table. Valid widths are from 1 through 32767 characters.
@property (copy) NSString *textFileOtherDelimiter;  // Returns or sets the character used as the delimiter when you import a text file into a query table.
@property ExcelXlTextParsingType textFileParseType;  // Returns or sets the column format for the data in the text file that you're importing into a query table.
@property NSInteger textFilePlatform;  // Returns or sets the origin of the text file you're importing into the query table. This property determines which code page is used during the data import.
@property BOOL textFilePromptOnRefresh;  // Returns or sets if you want to specify the name of the imported text file each time the query table is refreshed.
@property BOOL textFileSemicolonDelimiter;  // Returns or sets if the semicolon character is the delimiter when you import a text file into a query table.
@property BOOL textFileSpaceDelimiter;  // Returns or sets if the space character is the delimiter when you import a text file into a query table.
@property NSInteger textFileStartRow;  // Returns or sets the row number at which text parsing will begin when you import a text file into a query table. Valid values are integers from 1 through 32767. The default value is 1.
@property BOOL textFileTabDelimiter;  // Returns or sets if the tab character is the delimiter when you import a text file into a query table.
@property ExcelXlTextQualifier textFileTextQualifier;  // Returns or sets the text qualifier when you import a text file into a query table. The text qualifier specifies that the enclosed data is in text format.
@property (copy) NSString *textFileThousandsSeparator;  // Returns or sets the thousands separator character thatMicrosoft Excel uses when you import a text file into a query table. The default is the system thousands separator character.

- (void) cancelRefresh;  // Cancels all background queries for the specified query table. Use the refreshing property to determine whether a background query is currently in progress.
- (BOOL) refreshQueryTableBackgroundQuery:(BOOL)backgroundQuery;  // Updates the PivotTable cache or query table.

@end

// Represents a file in the list of recently used files.
@interface ExcelRecentFile : SBObject <ExcelGenericMethods>

@property (readonly) NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property (copy, readonly) NSString *name;  // Returns the name of the object.
@property (copy, readonly) NSString *path;  // Returns the complete path to the file, excluding the final separator and name of the file.


@end

// The RectangularGradient object transitions through a series of colors from the center to one or more sides.
@interface ExcelRectangularGradient : SBObject <ExcelGenericMethods>

@property (copy, readonly) ExcelColorstops *colorstops;  // Returns the ColorStops for the LinearGradient object.
@property double rectangularGradientBottom;  // Represents the point or vector that the gradient fill converges to.
@property double rectangularGradientLeft;  // Represents the point or vector that the gradient fill converges to.
@property double rectangularGradientRight;  // Represents the point or vector that the gradient fill converges to.
@property double rectangularGradientTop;  // Represents the point or vector that the gradient fill converges to.


@end

@interface ExcelRowField : ExcelPivotField


@end

@interface ExcelRowItem : ExcelPivotItem


@end

// Represents a scenario on a worksheet. Represents a scenario on a worksheet. A scenario is a group of input values, called changing cells, that's named and saved.
@interface ExcelScenario : SBObject <ExcelGenericMethods>

@property (copy) NSString *ExcelComment;  // Returns or sets the comment associated with the scenario. The comment text cannot exceed 255 characters.
@property (copy, readonly) ExcelRange *changingCells;  // Returns a range object that represents the changing cells for a scenario.
@property (readonly) NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property BOOL hidden;  //  Returns or sets if the scenario is hidden.
@property BOOL locked;  // Returns or sets if the object is locked.
@property (copy) NSString *name;  // Returns or sets the name of the object.

- (void) changeScenarioChangingCells:(ExcelXLRangeReference)changingCells values:(NSArray<id> *)values;  // Changes the scenario to have a new set of changing cells and, optionally, scenario values.
- (NSArray<id> *) getValues;  // Returns or sets a list that contains the current values of the changing cells for the scenario.

@end

@interface ExcelScrollbar : SBObject <ExcelGenericMethods>

@property (copy, readonly) ExcelRange *bottomRightCell;  // Returns the bottom right cell of the range the control is occupying.
@property BOOL displayThreeDShading;
@property BOOL enabled;  // Returns or sets if the object is enabled.
@property (readonly) NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property double height;  // Returns or set the height of the object.
@property NSInteger largeChange;
@property double leftPosition;  // Returns or sets the position of the specified object, in points.
@property (copy) NSString *linkedCell;
@property BOOL locked;  // Returns or sets if the object is locked, if false the object can be modified when the sheet is protected.
@property NSInteger maximumValue;
@property NSInteger minimumValue;
@property (copy) NSString *name;  // Returns or sets the name of the object.
@property (copy) NSString *onAction;  // Returns or sets the name of a string of AppleScript commands that will be executed when the object is clicked on. Please note that if AppleScript commands are set they will not be saved with the document.
@property ExcelXlPlacement placement;  // Returns or sets how the object is placed on the worksheet.
@property BOOL printObject;  // Returns or sets if this object is printed.
@property NSInteger smallChange;
@property double top;  // Returns the top position of the specified object, in points.
@property (copy, readonly) ExcelRange *topLeftCell;  // Returns a range object that represents the cell that lies under the upper-left corner of the specified object. 
@property NSInteger value;
@property BOOL visible;  // Returns or sets if the object is visible.
@property double width;  // Returns or sets  the width of the object.
@property (readonly) NSInteger zOrderPosition;  // Returns the z-order position of the object.


@end

// Represents a worksheet.
@interface ExcelSheet : SBObject <ExcelGenericMethods>

- (SBElementArray<ExcelShape *> *) shapes;
- (SBElementArray<ExcelArc *> *) arcs;
- (SBElementArray<ExcelButton *> *) buttons;
- (SBElementArray<ExcelChartObject *> *) chartObjects;
- (SBElementArray<ExcelCheckbox *> *) checkboxes;
- (SBElementArray<ExcelDropdown *> *) dropdowns;
- (SBElementArray<ExcelGroupbox *> *) groupboxes;
- (SBElementArray<ExcelLabel *> *) labels;
- (SBElementArray<ExcelLine *> *) lines;
- (SBElementArray<ExcelListbox *> *) listboxes;
- (SBElementArray<ExcelNamedItem *> *) namedItems;
- (SBElementArray<ExcelOptionButton *> *) optionButtons;
- (SBElementArray<ExcelOval *> *) ovals;
- (SBElementArray<ExcelPivotTable *> *) pivotTables;
- (SBElementArray<ExcelRange *> *) ranges;
- (SBElementArray<ExcelCell *> *) cells;
- (SBElementArray<ExcelRow *> *) rows;
- (SBElementArray<ExcelColumn *> *) columns;
- (SBElementArray<ExcelRectangle *> *) rectangles;
- (SBElementArray<ExcelScenario *> *) scenarios;
- (SBElementArray<ExcelScrollbar *> *) scrollbars;
- (SBElementArray<ExcelSpinner *> *) spinners;
- (SBElementArray<ExcelTextbox *> *) textboxes;
- (SBElementArray<ExcelHorizontalPageBreak *> *) horizontalPageBreaks;
- (SBElementArray<ExcelVerticalPageBreak *> *) verticalPageBreaks;
- (SBElementArray<ExcelQueryTable *> *) queryTables;
- (SBElementArray<ExcelExcelComment *> *) ExcelComments;
- (SBElementArray<ExcelHyperlink *> *) hyperlinks;
- (SBElementArray<ExcelListObject *> *) listObjects;

@property BOOL autofilterMode;  // Returns or sets if the autofilter drop-down arrows are currently displayed on the sheet. This property is independent of the filter mode property. 
@property (copy, readonly) ExcelAutofilter *autofilterObject;  // Returns the autofilter object associated with this sheet.
@property (copy, readonly) ExcelRange *circularReference;  // Returns a range object that represents the range containing the first circular reference on the sheet, or returns missing value if there's no circular reference on the sheet.
@property (readonly) ExcelXlConsolidationFunction consolidationFunction;  // Returns the function code used for the current consolidation.
@property (copy, readonly) NSArray<NSNumber *> *consolidationOptions;  // Returns a three-element list of boolean values. If the element is true, that option is set. Element 1 is use labels in top row, element 2 is use labels in left column, and element 3 is create links to source data.
@property (copy, readonly) NSArray<NSString *> *consolidationSources;  // Returns an list of string values that name the source sheets for the worksheet's current consolidation.
@property BOOL displayPageBreaks;  // Returns or sets if page breaks, both automatic and manual, on the specified worksheet are displayed.
@property BOOL enableAutofilter;  // Returns or sets if autofilter arrows are enabled when user-interface-only protection is turned on.
@property BOOL enableCalculation;  // Returns or sets if Microsoft Excel automatically recalculates the worksheet when necessary. If false, the user cannot request a recalculation, Microsoft Excel never recalculates the sheet automatically.
@property BOOL enableOutlining;  // Returns or sets if outlining symbols are enabled when user-interface-only protection is turned on.
@property BOOL enablePivotTable;  // Returns or sets if pivot table controls and actions are enabled when user-interface-only protection is turned on.
@property (readonly) NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property (readonly) BOOL filterMode;  // Returns true if the worksheet is in filter mode.
@property (copy) NSString *name;  // Returns or sets the name of the object.
@property (copy, readonly) ExcelSheet *next;  // Returns a worksheet object that represents the next sheet.
@property (copy, readonly) ExcelOutline *outlineObject;  // Returns an outline object that represents the outline for the specified worksheet.
@property (copy, readonly) ExcelPageSetup *pageSetupObject;  // Returns the page setup object associated with this chart.
@property (copy, readonly) ExcelSheet *previous;  // Returns a worksheet object that represents the previous sheet.
@property (readonly) BOOL protectContents;  // Returns true if the contents of the sheet are protected.
@property (readonly) BOOL protectDrawingObjects;  // Returns true if shapes are protected.
@property (readonly) BOOL protectionMode;  // Returns true if user-interface-only protection is turned on. To turn on user interface protection, use the protect method with the user interface only argument set to true.
@property (copy, readonly) ExcelProtection *protectionObject;  // Returns a Protection object that represents the protection options of the worksheet.
@property (copy) NSString *scrollArea;  // Returns or sets the range where scrolling is allowed, as an A1-style range reference. Cells outside the scroll area cannot be selected. 
@property (copy, readonly) ExcelTab *sheetTab;  // Returns the sheet tab of the work sheet
@property (copy, readonly) ExcelSort *sortObject;  // Returns the sort object in the sheet.
@property (readonly) double standardHeight;  // Returns the standard default height of all the rows in the worksheet, in points. 
@property double standardWidth;  // Returns or sets the standard default width of all the columns in the worksheet.
@property BOOL transitionExpressionEvaluation;  // Returns or sets if Microsoft Excel uses Lotus 1-2-3 expression evaluation rules for the worksheet.
@property (copy, readonly) ExcelRange *usedRange;  // Returns a range object that represents the used range on the specified worksheet.
@property ExcelXlSheetVisibility visible;  // Returns or sets if the worksheet is visible.
@property (readonly) ExcelXlSheetType worksheetType;  // Returns the type of this worksheet.

- (void) circleInvalid;  // Circles invalid entries on the worksheet.
- (void) clearArrows;  // Clears the tracer arrows from the worksheet. Tracer arrows are added by using the auditing feature.
- (void) clearCircles;  // Clears circles from invalid entries on the worksheet.
- (void) copyWorksheetBefore:(ExcelSheet *)before after:(ExcelSheet *)after NS_RETURNS_NOT_RETAINED;  // Copies the sheet to another location in the workbook.
- (void) pasteSpecialOnWorksheetFormat:(NSString *)format link:(BOOL)link displayAsIcon:(BOOL)displayAsIcon iconFileName:(NSString *)iconFileName iconIndex:(NSInteger)iconIndex iconLabel:(NSString *)iconLabel noHTMLFormatting:(BOOL)noHTMLFormatting;  // Pastes the contents of the clipboard onto the sheet, using a specified format. Use this method to paste data from other applications or to paste data in a specific format.
- (void) pasteWorksheetDestination:(ExcelXLRangeReference)destination link:(BOOL)link;  // Pastes the contents of the Clipboard onto the sheet.
- (void) protectWorksheetPassword:(NSString *)password drawingObjects:(BOOL)drawingObjects worksheetContents:(BOOL)worksheetContents scenarios:(BOOL)scenarios userInterfaceOnly:(BOOL)userInterfaceOnly allowFormattingCells:(BOOL)allowFormattingCells allowFormattingColumns:(BOOL)allowFormattingColumns allowFormattingRows:(BOOL)allowFormattingRows allowInsertingColumns:(BOOL)allowInsertingColumns allowInsertingRows:(BOOL)allowInsertingRows allowInsertingHyperlinks:(BOOL)allowInsertingHyperlinks allowDeletingColumns:(BOOL)allowDeletingColumns allowDeletingRows:(BOOL)allowDeletingRows allowSorting:(BOOL)allowSorting allowFiltering:(BOOL)allowFiltering allowUsingPivotTable:(BOOL)allowUsingPivotTable;  // Protects a worksheet so that it cannot be modified.
- (void) resetAllPageBreaks;  // Resets all page breaks on the specified worksheet.
- (void) saveAsFilename:(NSString *)filename fileFormat:(ExcelXlFileFormat)fileFormat password:(NSString *)password writeReservationPassword:(NSString *)writeReservationPassword readOnlyRecommended:(BOOL)readOnlyRecommended createBackup:(BOOL)createBackup addToMostRecentlyUsedList:(BOOL)addToMostRecentlyUsedList saveAsLocalLanguage:(BOOL)saveAsLocalLanguage;  // Saves changes into a different file.
- (void) setBackgroundPicturePictureFileName:(NSString *)pictureFileName;  // Sets the background graphic for a worksheet.
- (void) showAllData;  // Makes all rows of the currently filtered list visible. If autofilter is in use, this method changes the arrows to All.
- (void) showDataForm;  // Displays the data form associated with the worksheet.

@end

@interface ExcelInternationalMacroSheet : ExcelSheet

@property ExcelXlEnableSelection enableSelection;  // Returns or sets what can be selected on the sheet.


@end

@interface ExcelMacroSheet : ExcelSheet

@property ExcelXlEnableSelection enableSelection;  // Returns or sets what can be selected on the sheet.


@end

// A slicer is a mechanism for filtering data in PivotTable views and cube functions.
@interface ExcelSlicer : SBObject <ExcelGenericMethods>


@end

// Represents a sort of a range of data.
@interface ExcelSort : SBObject <ExcelGenericMethods>

@property BOOL matchCase;  // Set to true to perform a case-sensitive sort or set to false to perform non-case sensitive sort. Read/Write.
@property ExcelXlYesNoGuess sortHeader;  // Specifies whether the first row contains header information. Read/Write.
@property ExcelXlSortMethod sortMethod;  // Specifies the sort method for Chinese/Japanese languages. Read/Write.
@property ExcelXlSortOrientation sortOrientation;  // Specifies the orientation for the sort. Read/Write.
@property (copy, readonly) id sortfieldset;  // Stores the sort state for workbooks, lists, and autofilters. Read-Only.
@property (copy, readonly) ExcelRange *sortrange;  // Returns a range object that represents the range to which the specified sort applies. Read Only.

- (void) setSortRangeRng:(ExcelRange *)rng;  // Sets the starting and ending character positions for Sort object.

@end

// The sortfield object contains all the sort information for the worksheet, lists, and autofilter object.
@interface ExcelSortfield : SBObject <ExcelGenericMethods>

@property (copy) id customOrder;  // Specifies a custom order to sort the fields. Read/Write.
@property ExcelXlSortDataOption sortDataOption;  // Specifies how to sort text in the range specified in sortfield object. Read/Write.
@property (copy, readonly) ExcelRange *sortKey;  // Specifies the range that is currently being sorted on. Read only.
@property ExcelXlSortOn sortOn;  // Returns or sets what attribute of the cell to sort on. Read/Write.
@property (copy, readonly) id sortOnValues;  // Return the value on which the sort is performed for the specified sortfield object. Read Only.
@property ExcelXlSortOrder sortOrder;  // Determines the sort order for the values specified in the key. Read/Write.
@property NSInteger sortPriority;  // Specifies the priority for the sortfield. Read/Write.

- (void) modifySortKeyRng:(ExcelRange *)rng;  // Modify the key value by which values are sorted in the field.

@end

// The sortfieldset collection is a collection of sortfield objects. It allows developers to store a sort state on worksheets, lists, and autofilters.
@interface ExcelSortfieldset : SBObject <ExcelGenericMethods>

- (void) addSortfieldKey:(ExcelRange *)key sorton:(ExcelXlSortOn)sorton order:(ExcelXlSortOrder)order customorder:(id)customorder dataoption:(ExcelXlSortDataOption)dataoption;  // Creates a new sortfield and returns a sortfieldset object.

@end

@interface ExcelSpinner : SBObject <ExcelGenericMethods>

@property (copy, readonly) ExcelRange *bottomRightCell;  // Returns the bottom right cell of the range the control is occupying.
@property BOOL displayThreeDShading;  // Returns or sets whether a 3-d effect will be employed when displaying the control
@property BOOL enabled;  // Returns or sets if the control is enabled
@property (readonly) NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property double height;  // Returns or sets the height of the control.
@property double leftPosition;  // Returns or sets the left position of the control
@property (copy) NSString *linkedCell;  // Returns or sets reference to a call which contains the value of the control.
@property BOOL locked;  // Returns or sets whether the control is locked for editing.
@property NSInteger maximumValue;  // Returns or sets the maximum value allowed
@property NSInteger minimumValue;  // Returns or sets the minimum value allowed
@property (copy) NSString *name;  // Returns or sets the name of this control
@property (copy) NSString *onAction;  // Returns or sets the name of a string of AppleScript commands that will be executed when the object is clicked on. Please note that if AppleScript commands are set they will not be saved with the document.
@property ExcelXlPlacement placement;  // Returns or sets how the object is placed on the worksheet.
@property BOOL printObject;  // Returns or sets if this object is printed.
@property NSInteger smallChange;  // Returns or sets the change in value of one click on the spinner control.
@property double top;  // Returns or sets the position of the top of the control.
@property (copy, readonly) ExcelRange *topLeftCell;  // Returns the top left cell of the range within which the control is positioned.
@property NSInteger value;  // Returns or sets the current value of the control
@property BOOL visible;  // Returns or sets if the control is visible
@property double width;  // Returns or sets  the width of the object.
@property (readonly) NSInteger zOrderPosition;  // Returns the most recently set z order position, bring shape to front/send shape to back/bring shape forward/send shape backward/bring shape in front of text/send shape behind text.


@end

// Represents the sheet tab of a work sheet or chart sheet.
@interface ExcelTab : SBObject <ExcelGenericMethods>

@property (copy) NSArray<NSNumber *> *color;
@property ExcelXlColorIndex colorIndex;
@property ExcelXlThemeColor themeColor;
@property double tintAndShade;


@end

// Represents a table style element
@interface ExcelTableStyleElement : SBObject <ExcelGenericMethods>

@property (copy, readonly) ExcelFont *fontObject;  // Returns a font object that represents the font of the specified object.
@property (readonly) BOOL hasFormat;  // Returns or sets if table style element has format.
@property (copy, readonly) ExcelInterior *interiorObject;  // Returns a interior object that represents the interior for the specified table style element. 


@end

// Represents a table style
@interface ExcelTableStyle : SBObject <ExcelGenericMethods>

- (SBElementArray<ExcelTableStyleElement *> *) tableStyleElements;

@property (readonly) BOOL builtIn;  // if this is a built in table style or not
- (NSString *) default;
@property (copy, readonly) NSString *name;  // Returns or sets the name of the object.
@property (copy, readonly) NSString *namelocal;  // Returns or sets the local name of the object.
@property BOOL showAsAvailablePivotTableStyle;  // whether to show as avaliable pivot table style
@property BOOL showAsAvailableTableStyle;  // whether to show as avaliable table style


@end

@interface ExcelTextbox : SBObject <ExcelGenericMethods>

- (SBElementArray<ExcelCharacter *> *) characters;

@property BOOL addIndent;  // Returns or sets if text is automatically indented when the text alignment in a cell is set to equal distribution either horizontally or vertically.
@property BOOL autoScaleFont;  // Returns or sets  if the text in the object changes font size when the object size changes.
@property BOOL autoSize;  // Returns or sets if the size of the specified object is changed automatically to fit text within its boundaries.
@property (copy, readonly) ExcelBorder *border;  // Returns a border object that represents the border of the object. 
@property (copy, readonly) ExcelRange *bottomRightCell;  // Returns the bottom right cell of the range the control is occupying.
@property (copy) NSString *caption;  // Returns or sets the caption for this object.
@property (readonly) NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property (copy, readonly) ExcelFont *fontObject;  // Returns a font object that represents the font of the specified object.
@property (copy) NSString *formula;  // Returns or sets the object's formula, in A1-style notation and in the language of the macro.
@property double height;  // Returns or set the height of the object.
@property ExcelXlHAlign horizontalAlignment;  // Returns or sets the horizontal alignment for the object.
@property (copy, readonly) ExcelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property double leftPosition;  // Returns or sets the position of the specified object, in points.
@property BOOL locked;  // Returns or sets if the object is locked, if false the object can be modified when the sheet is protected.
@property BOOL lockedText;  // Returns or sets whether the control's text is locked for editing.
@property (copy) NSString *name;  // Returns or sets the name of the object.
@property (copy) NSString *onAction;  // Returns or sets the name of a string of AppleScript commands that will be executed when the object is clicked on. Please note that if AppleScript commands are set they will not be saved with the document.
@property ExcelXlOrientation orientation;  // May also be a number value from -90 to 90 degrees.
@property ExcelXlPlacement placement;  // Returns or sets how the object is placed on the worksheet.
@property BOOL printObject;  // Returns or sets if this object is printed.
@property ExcelXLDefaultSheetDir readingOrder;  // Returns or sets the reading order for the specified object.
@property BOOL roundedCorners;
@property BOOL shadow;
@property (copy) NSString *stringValue;  // Returns or sets the text of the specified object.
@property double top;  // Returns the top position of the specified object, in points.
@property (copy, readonly) ExcelRange *topLeftCell;  // Returns a range object that represents the cell that lies under the upper-left corner of the specified object. 
@property ExcelXlVerticalAlignmentTarget verticalAlignment;  // Returns or sets the vertical alignment of the object.
@property BOOL visible;  // Returns or sets if the object is visible.
@property double width;  // Returns or sets  the width of the object.
@property (readonly) NSInteger zOrderPosition;  // Returns the z-order position of the object.


@end

@interface ExcelTop10FormatCondition : SBObject <ExcelGenericMethods>

@property (copy, readonly) ExcelRange *appliesTo;  // Returns the range this format condition applies to. Read Only.
@property ExcelXlCalcFor calcFor;
@property (copy, readonly) ExcelFont *fontObject;  // Returns a font object that represents the font of the specified object.
@property NSInteger formatConditionPriority;  // Specifies the priority for the format condition. Read/Write.
@property (readonly) ExcelXlFormatConditionType formatConditionType;  // Return the conditional format type.
@property (copy, readonly) ExcelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property (copy) NSString *numberFormat;  // Returns or sets the format code for the object.
@property ExcelXlPivotConditionScope pivotConditionScopeType;  // Returns or sets the part of the pivot table that the pivot table format condition is scoped to
@property (readonly) BOOL pivotTableCondition;  // Tells whether this format condition is a pivot table condition.
@property (readonly) BOOL stopIfTrue;  // Tells whether the processing of format conditions stops if this condition is true. Read Only.
@property BOOL top10Percentage;
@property NSInteger top10Rank;
@property ExcelXlTopBottom topOrBottom;


@end

// Represents the hierarchical member-selection control of a cube field.
@interface ExcelTreeviewControl : SBObject <ExcelGenericMethods>

@property (copy) id drilled;
@property (copy) id hidden;


@end

@interface ExcelUniqueValuesFormatCondition : SBObject <ExcelGenericMethods>

@property (copy, readonly) ExcelRange *appliesTo;  // Returns the range this format condition applies to. Read Only.
@property ExcelXlDupeUnique duplicateOrUnique;
@property (copy, readonly) ExcelFont *fontObject;  // Returns a font object that represents the font of the specified object.
@property NSInteger formatConditionPriority;  // Specifies the priority for the format condition. Read/Write.
@property (readonly) ExcelXlFormatConditionType formatConditionType;  // Return the conditional format type.
@property (copy, readonly) ExcelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property (copy) NSString *numberFormat;  // Returns or sets the format code for the object.
@property ExcelXlPivotConditionScope pivotConditionScopeType;  // Returns or sets the part of the pivot table that the pivot table format condition is scoped to
@property (readonly) BOOL pivotTableCondition;  // Tells whether this format condition is a pivot table condition.
@property (readonly) BOOL stopIfTrue;  // Tells whether the processing of format conditions stops if this condition is true. Read Only.


@end

// Represents data validation for a worksheet range.
@interface ExcelValidation : SBObject <ExcelGenericMethods>

@property ExcelXlIMEMode IMEMode;  // Returns or sets the description of the Japanese input rules.
@property (readonly) ExcelXlDVAlertStyle alertStyle;  // Returns the validation alert style.
@property (copy) NSString *errorMessage;  // Returns or sets the data validation error message.
@property (copy) NSString *errorTitle;  // Returns or sets the title of the data-validation error dialog box.
@property (copy, readonly) NSString *formula1;  // Returns the value or expression associated with the conditional format or data validation.
@property (copy, readonly) NSString *formula2;  // Returns the value or expression associated with the second part of a conditional format or data validation. Used only when the data validation conditional format Operator property is operator between or operator not between.
@property BOOL ignoreBlank;  // Returns or sets if blank values are permitted by the range data validation.
@property BOOL inCellDropdown;  // Returns or sets if data validation displays a drop-down list that contains acceptable values.
@property (copy) NSString *inputMessage;  // Returns or sets the data validation input message.
@property (copy) NSString *inputTitle;  // Returns or sets the title of the data-validation input dialog box.
@property BOOL showError;  // Returns or sets if the data validation error message will be displayed whenever the user enters invalid data.
@property BOOL showInput;  // Returns or sets if the data validation input message will be displayed whenever the user selects a cell in the data validation range.
@property (readonly) ExcelXlFormatConditionOperator validationOperator;  // Returns the operator for the conditional format or data validation.
@property (readonly) ExcelXlDVType validationType;  // Returns the data validation type.
@property (readonly) BOOL value;  // Returns true if all the validation criteria are met, that is, if the range contains valid data.

- (void) addDataValidationType:(ExcelXlDVType)type alertStyle:(ExcelXlDVAlertStyle)alertStyle operator:(ExcelXlFormatConditionOperator)operator_ formula1:(NSString *)formula1 formula2:(NSString *)formula2;  // Adds data validation to the specified range.
- (void) modifyType:(ExcelXlDVType)type alertStyle:(ExcelXlDVAlertStyle)alertStyle operator:(ExcelXlFormatConditionOperator)operator_ formula1:(NSString *)formula1 formula2:(NSString *)formula2;  // Modifies data validation for a range.

@end

@interface ExcelValueChange : SBObject <ExcelGenericMethods>

@property (readonly) ExcelXlAllocationMethod allocationMethod;
@property (readonly) ExcelXlAllocationValue allocationValue;
@property (copy, readonly) NSString *allocationWeightExpression;
@property (readonly) NSInteger order;
@property (copy, readonly) ExcelPivotCell *pivotCell;
@property (copy, readonly) NSString *tuple;
@property (readonly) double value;
@property (readonly) BOOL visibleInPivotTable;


@end

// Represents a vertical page break. 
@interface ExcelVerticalPageBreak : SBObject <ExcelGenericMethods>

@property (readonly) ExcelXlPageBreakExtent extent;  // Returns the type of the specified page break: full-screen or only within a print area.
@property (copy) ExcelRange *location;  // Returns or sets where this vertical page break is located.
@property ExcelXlPageBreak verticalPageBreakType;  // Returns or sets the type of vertical page break.


@end

// Contains workbook-level attributes used by Microsoft Excel when you save a document as a Web page or open a Web page.
@interface ExcelWebOptions : SBObject <ExcelGenericMethods>

@property BOOL allowPng;  // Returns or sets if PNG, Portable Network Graphics, is allowed as an image format when you save documents as a Web page.
@property ExcelMsoEncoding encoding;  // Returns or sets the document encoding, code page or character set, to be used by the Web browser when you view the saved document.
@property (copy) NSString *locationOfComponents;  // Returns or sets the central URL, on the intranet or Web, or path, local or network, to the location from which authorized users can download Microsoft Office Web components when viewing your saved document.
@property NSInteger pixelsPerInch;  // Returns or sets the density, pixels per inch, of graphics images and table cells on a Web page. The range of settings is usually from 19 to 480, and common settings for popular screen sizes are 72, 96, and 120.
@property (readonly) ExcelMsoScreenSize screenSize;  // Returns the ideal minimum screen size, width by height, in pixels, that you should use when viewing the saved document in a Web browser.

- (void) useDefaultFolderSuffix;  // Sets the folder suffix for the specified document to the default suffix for the language support you have selected or installed.

@end

// Represents a window. Many worksheet characteristics, such as scroll bars and gridlines, are actually properties of the window.
@interface ExcelWindow (MicrosoftExcelSuite)

- (SBElementArray<ExcelPane *> *) panes;

@property (copy, readonly) ExcelCell *activeCell;  // Returns a range object that represents the active cell in the active window, the window on top,  or in the specified window. If the window isn't displaying a worksheet, this property fails.
@property (copy, readonly) ExcelPane *activePane;  // Returns a pane object that represents the active pane in the active window, the window on top,  or in the specified window. If the window isn't displaying a worksheet, this property fails.
@property (copy, readonly) ExcelSheet *activeSheet;  // Returns an object that represents the active sheet, the sheet on top, in the active workbook or in the specified window or workbook.
@property (copy) NSString *caption;  // Returns or sets the name that appears in the title bar of the document window. 
@property BOOL dateGrouping;  // Returns or sets the date grouping flag in the window.
@property BOOL displayFormulas;  // Returns or set if the window is displaying formulas.  If false, the window is displaying values.
@property BOOL displayGridlines;  // Returns or sets if gridlines are displayed.
@property BOOL displayHeadings;  // Returns or sets if both row and column headings are displayed.
@property BOOL displayHorizontalScrollBar;  // Returns or sets if the horizontal scroll bar is displayed.
@property BOOL displayOutline;  // Returns or sets if outline symbols are displayed.
@property BOOL displayVerticalScrollBar;  // Returns or sets if the vertical scroll bar is displayed.
@property BOOL displayWorkbookTabs;  // Returns or sets if the workbook tabs are displayed.
@property BOOL displayZeros;  // Returns or sets if zero values are displayed.
@property BOOL enableResize;  // Returns or sets if the window can be resized. 
@property (readonly) NSInteger entry_index;  // Returns the index of this item in the element list of windows.
@property BOOL freezePanes;  // Returns or sets if split panes are frozen. It's possible for freeze panes to be true and split to be false, or vice versa. This property applies only to worksheets and macro sheets.
@property (copy) NSArray<NSNumber *> *gridlineColor;  // Returns or sets the gridline color as an RGB value. 
@property ExcelXlColorIndex gridlineColorIndex;  // Returns or sets the gridline color as an index into the current color palette.
@property double height;  // Returns or sets the height of the window. Use the usable height property to determine the maximum size for the window. You cannot set this property if the window is maximized or minimized. Use the window state property to determine the window state.
@property double leftPosition;  // Returns or sets the distance from the left edge of the client area to the left edge of the window.
@property (copy, readonly) ExcelRange *rangeSelection;  // Returns a range object that represents the selected cells on the worksheet in the specified window even if a graphic object is active or selected on the worksheet.
@property NSInteger scrollColumn;  // Returns or sets the number of the leftmost column in the window. 
@property NSInteger scrollRow;  // Returns or sets the number of the row that appears at the top of the window.
@property (copy, readonly) NSArray<NSString *> *selectedSheets;  // Returns a list of sheet objects that represents all the selected sheets in the specified window.
@property (copy, readonly) ExcelRange *selection;  // Returns the selected range object in the specified window.
@property BOOL split;  // Returns or sets if the window is split. 
@property NSInteger splitColumn;  // Returns or sets the column number where the window is split into panes, the number of columns to the left of the split line.
@property double splitHorizontal;  // Returns or sets the location of the horizontal window split, in points.
@property NSInteger splitRow;  // Returns or sets the row number where the window is split into panes, the number of rows above the split.
@property double splitVertical;  // Returns or sets the location of the vertical window split, in points.
@property double tabRatio;  // Returns or sets the ratio of the width of the workbook's tab area to the width of the window's horizontal scroll bar, as a number between 0 and 1, the default value is 0.75. 
@property double top;  //  The distance from the top edge of the window to the top edge of the usable area, below the menus, any toolbars docked at the top, and the formula bar. You cannot set this property for a maximized window.
@property (readonly) double usableHeight;  // Returns the maximum height of the space that a window can occupy in points.
@property (readonly) double usableWidth;  // Returns the maximum width of the space that a window can occupy in the application window area, in points.
@property ExcelXlWindowView view;  // Returns or sets the view showing in the window.
@property BOOL visible;  // Returns or sets if the window is visible.
@property (copy, readonly) ExcelRange *visibleRange;  // Returns a range object that represents the range of cells that are visible in the window or pane. If a column or row is partially visible, it's included in the range. 
@property double width;  // Returns or sets the width of the window.
@property (readonly) NSInteger windowNumber;  // Returns the window number. For example, a window named Book1.xls:2 has 2 as its window number. Most windows have the window number 1.
@property ExcelXlWindowState windowState;  // Returns or sets the state of the window.
@property (readonly) ExcelXlWindowType windowType;  // Returns the window type.
@property ExcelXLZoomType zoom;  // Returns or sets the display size of the window, as a percentage,100 equals normal size, 200 equals double size, and so on. 

@end

// A connection is a set of information needed to obtain data from an external data source other than an 1st_Excel12 workbook.
@interface ExcelWorkbookConnection : SBObject <ExcelGenericMethods>

@property (copy, readonly) NSString *_default;
@property (copy, readonly) NSString *objectDescription;  // Returns or sets a brief description for a WorkbookConnection object.
@property (copy, readonly) NSString *name;  // Returns or sets the name of the workbook connection object.
@property (copy, readonly) id ranges;  // Returns the range of object for the specified WorkbookConnection object.
@property (readonly) ExcelXlConnectionType type;


@end

// Represents a Microsoft Excel workbook.
@interface ExcelWorkbook : SBObject <ExcelGenericMethods>

- (SBElementArray<ExcelDocumentProperty *> *) documentProperties;
- (SBElementArray<ExcelChartSheet *> *) chartSheets;
- (SBElementArray<ExcelCommandBar *> *) commandBars;
- (SBElementArray<ExcelCustomDocumentProperty *> *) customDocumentProperties;
- (SBElementArray<ExcelNamedItem *> *) namedItems;
- (SBElementArray<ExcelPivotCache *> *) pivotCaches;
- (SBElementArray<ExcelSheet *> *) sheets;
- (SBElementArray<ExcelStyle *> *) styles;
- (SBElementArray<ExcelCustomView *> *) customViews;
- (SBElementArray<ExcelWindow *> *) windows;
- (SBElementArray<ExcelWorksheet *> *) worksheets;
- (SBElementArray<ExcelInternationalMacroSheet *> *) internationalMacroSheets;
- (SBElementArray<ExcelMacroSheet *> *) macroSheets;
- (SBElementArray<ExcelTableStyle *> *) tableStyles;

@property BOOL acceptLabelsInFormulas;  // Returns or sets if labels can be used in worksheet formulas.
@property NSInteger accuracyVersion;  // Returns or sets the accuracy version for this workbook.
@property (copy, readonly) ExcelSheet *activeSheet;  // Returns an object that represents the active sheet, the sheet on top, in the specified window.
@property NSInteger autoUpdateFrequency;  // Returns or sets the number of minutes between automatic updates to the shared workbook. If this property is set to zero  updates occur only when the workbook is saved.
@property BOOL autoUpdateSaveChanges;  // Returns or sets if current changes to the shared workbook are posted to other users whenever the workbook is automatically updated. if false changes aren't posted, this workbook is still synchronized with changes made by other users.
@property (readonly) NSInteger calculationVersion;  // Returns a number whose rightmost four digits are the minor calculation engine version number, and whose other digits, on the left, are the major version of Microsoft Excel.
@property NSInteger changeHistoryDuration;  // Returns or sets the number of days shown in the shared workbook's change history.
@property ExcelXlSaveConflictResolution conflictResolution;  // Returns or sets the way conflicts are to be resolved whenever a shared workbook is updated.
@property (readonly) BOOL createBackup;  // Returns true if a backup file is created when this file is saved.
@property BOOL date1904;  // Returns or sets if the workbook uses the 1904 date system.
@property (copy) id defaultPivottableStyle;  // Set or get the default pivot table style for the current workbook
@property (copy) id defaultTableStyle;  // Set or get the default table style for the current workbook
@property ExcelXlDisplayDrawingObjects displayDrawingObjects;  // Returns or sets how shapes are displayed.
@property BOOL enableAutoRecover;  // Gets or sets a value that indicates whether Microsoft Office Excel saves changed files, of all formats, on a timed interval.
@property (readonly) BOOL excel8CompatibilityMode;  // Gets a value that indicates whether the workbook is in compatibility mode.
@property (readonly) ExcelXlFileFormat fileFormat;  // Returns the file format of the workbook.
@property (copy, readonly) NSString *fullName;  // Returns the name of the workbook, including its path on disk, as a string.
@property (copy, readonly) NSString *fullNameUrlencoded;  // Returns a String indicating the name of the object, including its path on disk, as a string.
@property (readonly) BOOL hasPassword;  // Returns true if the workbook has a protection password.
@property (readonly) BOOL hasVbProject;  // Gets a value that indicates whether a workbook has an attached Microsoft Visual Basic for Applications <VBA> project.
@property BOOL highlightChangesOnScreen;  // Returns or sets if changes to the shared workbook are highlighted on-screen.
@property BOOL inactiveListBorderVisible;  // Gets or sets a value that indicates whether list borders are visible when a list is not active.
@property BOOL isAddIn;  // Returns or sets if the workbook is running as an add in.
@property BOOL keepChangeHistory;  // Returns or sets if change tracking is enabled for the shared workbook.
@property BOOL listChangesOnNewSheet;  // Returns or sets if changes to the shared workbook are shown on a separate worksheet.
@property (readonly) BOOL multiUserEditing;  // Returns true if the workbook is open as a shared list.
@property (copy, readonly) NSString *name;  // Returns or sets the name of the workbook.
@property (copy) NSString *password;  // Returns or sets the password that must be supplied to open the specified workbook.
@property (copy, readonly) NSString *path;  // Returns the complete path of the object, excluding the final separator and name of the object.
@property BOOL personalViewListSettings;  // Returns or sets if filter and sort settings for lists are included in the user's personal view of the shared workbook.
@property BOOL personalViewPrintSettings;  // Returns or sets if print settings are included in the user's personal view of the shared workbook.
@property BOOL precisionAsDisplayed;  // Returns or sets if calculations in this workbook will be done using only the precision of the numbers as they're displayed.
@property (readonly) BOOL protectStructure;  // Returns true if the order of the sheets in the workbook is protected.
@property (readonly) BOOL protectWindows;  // Returns true if the windows of the workbook are protected.
@property (readonly) BOOL readOnly;  // Returns true if the workbook has been opened as read-only.
@property BOOL readOnlyRecommended;  // Returns or sets if the workbook was saved as read-only recommended.
@property BOOL removePersonalInformation;  // Returns or sets if personal information can be removed from the specified workbook.
@property (readonly) NSInteger revisionNumber;  // Returns the number of times the workbook has been saved while open as a shared list. If the workbook is open in exclusive mode, this property returns zero.
@property BOOL saveLinkValues;  // Returns or sets if Microsoft Excel saves external link values with the workbook.
@property BOOL saved;  // Returns or sets if no changes have been made to the specified workbook since it was last saved.
@property BOOL showConflictHistory;  // Returns or sets if the conflict history worksheet is visible in the workbook that's open as a shared list.
@property BOOL templateRemoveExternalData;  // Returns or sets if external data references are removed when the workbook is saved as a template.
@property (copy, readonly) ExcelOfficeTheme *theme;  // Represents a Microsoft Office theme.
@property BOOL updateRemoteReferences;  // Returns or sets if Microsoft Excel updates remote references in for the workbook.
@property (copy, readonly) NSArray<id> *userStatus;  // Returns a list of lists that provides information about each user who has the workbook open as a shared list. The list is name of user, date and time user opened the workbook, and a number 1 for exclusive, 2 for shared access.
@property (copy, readonly) ExcelWebOptions *webOptions;  // Returns the web options object, which contains workbook-level attributes used by Microsoft Excel when you save a document as a Web page or open a Web page.
@property (copy) NSString *workbookComments;  // Returns or sets the comment string for this workbook.
@property (copy) NSString *writePassword;  // Returns or sets a string for the write password of a workbook.
@property (readonly) BOOL writeReserved;  // Return true if the workbook is write-reserved
@property (copy, readonly) NSString *writeReservedBy;  // Returns the name of the user who currently has write permission for the workbook.

- (void) acceptAllChangesWhen:(NSString *)when who:(NSString *)who where:(NSString *)where;  // Accepts all changes in the specified shared workbook.
- (void) applyThemeFileName:(NSString *)fileName;  // Applies a theme or design template to the specified workbook
- (void) breakLinkName:(NSString *)name type:(ExcelXlLinkType)type;  // Converts formulas linked to other Microsoft Excel sources to values.
- (BOOL) canCheckIn;  // Returns True if Excel can check in a specified workbook to a server.
- (void) changeFileAccessMode:(ExcelXlFileAccess)mode writePassword:(NSString *)writePassword notify:(BOOL)notify;  // Changes the access permissions for the workbook. This may require an updated version to be loaded from the disk.
- (void) changeLinkName:(NSString *)name newName:(NSString *)newName type:(ExcelXlLinkType)type;  // Changes a link from one document to another.
- (void) checkInSaveChanges:(BOOL)saveChanges comments:(NSString *)comments makePublic:(BOOL)makePublic;  // Returns a workbook from a local computer to a server, and sets the local file to read-only so that it cannot be edited locally.
- (void) checkInWithVersionSaveChanges:(BOOL)saveChanges comments:(NSString *)comments makePublic:(BOOL)makePublic versionType:(ExcelXlCheckInVersionType)versionType;  // Returns a workbook from a local computer to a server, and sets the local file to read-only so that it cannot be edited locally.
- (void) deleteNumberFormatNumberFormat:(NSString *)numberFormat;  // Deletes a custom number format from the workbook.
- (BOOL) exclusiveAccess;  // Assigns the current user exclusive access to the workbook that's open as a shared list.
- (void) followHyperlinkAddress:(NSString *)address subAddress:(NSString *)subAddress newWindow:(BOOL)newWindow extraInfo:(NSString *)extraInfo method:(ExcelMsoAnimationType)method headerInfo:(NSString *)headerInfo;  // Displays a cached document, if it's already been downloaded. Otherwise, this method resolves the hyperlink, downloads the target document, and displays the document in the appropriate application.
- (void) highlightChangesOptionsWhen:(ExcelXlHighlightChangesTime)when who:(NSString *)who where:(ExcelXLRangeReference)where;  // Controls how changes are shown in a shared workbook.
- (SBObject *) linkInfoName:(NSString *)name linkInfo:(ExcelXlLinkInfo)linkInfo type:(ExcelXlLinkInfoType)type;  // Returns the link date and update status.
- (NSArray<NSString *> *) linkSourcesType:(ExcelXlLinkType)type;  // Returns a list of links in the workbook. The names in the array are the names of the linked documents. Returns empty if there are no links.
- (void) mergeWorkbookFileName:(NSString *)fileName;  // Merges changes from one workbook into an open workbook.
- (ExcelWindow *) newWindowOnWorkbook NS_RETURNS_NOT_RETAINED;  // Creates a new window or a copy of the specified workbook window.
- (void) openLinksName:(NSString *)name readOnly:(BOOL)readOnly type:(ExcelXlLinkType)type;  // Opens the supporting documents for a link or links.
- (void) protectSharingFileName:(NSString *)fileName password:(NSString *)password writeReservationPassword:(NSString *)writeReservationPassword readOnlyRecommended:(BOOL)readOnlyRecommended createBackup:(BOOL)createBackup sharingPassword:(NSString *)sharingPassword fileFormat:(ExcelXlFileFormat)fileFormat;  // Saves the workbook and protects it for sharing.
- (void) protectWorkbookPassword:(NSString *)password structure:(BOOL)structure windows:(BOOL)windows;  // Protect workbook structure from changes.
- (void) purgeChangeHistoryNowDays:(NSInteger)days sharingPassword:(NSString *)sharingPassword;  // Removes entries from the change log for the specified workbook.
- (void) refreshAll;  // Refreshes all external data ranges and PivotTables in the workbook.
- (void) rejectAllChangesWhen:(NSString *)when who:(NSString *)who where:(NSString *)where;  // Rejects all changes in the specified shared workbook.
- (void) removeUserEntry_index:(NSInteger)entry_index;  // Disconnects the specified user from the shared workbook.
- (void) resetColors;  // Resets the color palette to the default colors.
- (void) saveWorkbookAsFilename:(NSString *)filename fileFormat:(ExcelXlFileFormat)fileFormat password:(NSString *)password writeReservationPassword:(NSString *)writeReservationPassword readOnlyRecommended:(BOOL)readOnlyRecommended createBackup:(BOOL)createBackup accessMode:(ExcelXlSaveAsAccessMode)accessMode conflictResolution:(ExcelXlSaveConflictResolution)conflictResolution addToMostRecentlyUsedList:(BOOL)addToMostRecentlyUsedList;  // Saves changes into a different file.
- (void) sendMail;  // Opens a message window in your registered mail program for sending the specified document as an attachment.
- (void) toggleFormsDesign;  // Toggles Microsoft Office Excel into and out of design mode.
- (void) unprotectSharingSharingPassword:(NSString *)sharingPassword;  // Turns off protection for sharing and saves the workbook.
- (void) updateFromFile;  // Updates a read-only workbook from the saved disk version of the workbook if the disk version is more recent than the copy of the workbook that is loaded in memory.
- (void) updateLinkName:(NSString *)name type:(ExcelXlLinkType)type;  // Updates a Microsoft Excel  link.
- (void) webPagePreview;  // Displays a preview of the specified workbook as it would look if saved as a Web page.

@end

@interface ExcelDocument (MicrosoftExcelSuite)

@end

@interface ExcelWorksheet : ExcelSheet

@property ExcelXlEnableSelection enableSelection;  // Returns or sets what can be selected on the sheet.
@property (readonly) BOOL protectScenarios;  // Returns true if the worksheet scenarios are protected.

- (void) createSummaryForScenariosReportType:(ExcelXlSummaryReportType)reportType resultCells:(ExcelXLRangeReference)resultCells;  // Creates a new worksheet that contains a summary report for the scenarios on the specified worksheet.
- (void) mergeScenariosMergeSource:(id)mergeSource;  // Merges the scenarios from the merge source worksheet into this worksheet

@end

// Contains global application-level spelling options
@interface ExcelXlspellingOptions : SBObject <ExcelGenericMethods>

@property ExcelXlLanguage dictionaryLang;  // Returns or sets the LCID used by the proofing tools


@end



/*
 * Drawing Suite
 */

@interface ExcelAdjustment : SBObject <ExcelGenericMethods>

@property double adjustment_value;


@end

// Represents an arc graphic.
@interface ExcelArc : SBObject <ExcelGenericMethods>

- (SBElementArray<ExcelCharacter *> *) characters;

@property NSInteger addIndent;  // Returns or sets if text is automatically indented when the text alignment in a cell is set to equal distribution either horizontally or vertically.
@property BOOL autoScaleFont;  // Returns or sets  if the text in the object changes font size when the object size changes.
@property NSInteger autoSize;  // Returns or sets if the size of the specified object is changed automatically to fit text within its boundaries.
@property (copy, readonly) ExcelBorder *border;  // Returns a border object that represents the border of the object. 
@property (readonly) NSInteger bottomRightCell;  // Returns the bottom right cell of the range the control is occupying.
@property NSInteger caption;  // Returns or sets the caption for this object.
@property (readonly) NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property (readonly) NSInteger fontObject;  // Returns a font object that represents the font of the specified object.
@property NSInteger formula;  // Returns or sets the object's formula, in A1-style notation and in the language of the macro.
@property NSInteger height;  // Returns or set the height of the object.
@property NSInteger horizontalAlignment;  // Returns or sets the horizontal alignment for the object.
@property (readonly) NSInteger interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property NSInteger leftPosition;  // Returns or sets the position of the specified object, in points.
@property NSInteger locked;  // Returns or sets if the object is locked, if false the object can be modified when the sheet is protected.
@property NSInteger lockedText;  // Returns or sets whether the control's text is locked for editing.
@property NSInteger name;  // Returns or sets the name of the object.
@property (copy) NSString *onAction;  // Returns or sets the name of a string of AppleScript commands that will be executed when the object is clicked on. Please note that if AppleScript commands are set they will not be saved with the document.
@property NSInteger orientation;  // May also be a number value from -90 to 90 degrees.
@property NSInteger placement;  // Returns or sets how the object is placed on the worksheet.
@property NSInteger printObject;  // Returns or sets if this object is printed.
@property ExcelXLDefaultSheetDir readingOrder;  // Returns or sets the reading order for the specified object.
@property NSInteger stringValue;  // Returns or sets the text of the specified object.
@property NSInteger top;  // Returns the top position of the specified object, in points.
@property (readonly) NSInteger topLeftCell;  // Returns a range object that represents the cell that lies under the upper-left corner of the specified object. 
@property NSInteger verticalAlignment;  // Returns or sets the vertical alignment of the object.
@property NSInteger visible;  // Returns or sets if the object is visible.
@property NSInteger width;  // Returns or sets  the width of the object.
@property (readonly) NSInteger zOrderPosition;  // Returns the z-order position of the object.


@end

@interface ExcelBulletFormat : SBObject <ExcelGenericMethods>

@property (copy) NSString *bulletCharacter;  // Returns or sets the Unicode character value that is used for bullets in the specified text.
@property (copy, readonly) ExcelShapeFont *bulletFont;  // Returns a font object that represents character formatting for a bullet format object.
@property (readonly) NSInteger bulletNumber;  // Returns the bullet number of a paragraph.
@property NSInteger bulletStartValue;  // Gets or sets the beginning value of a bulleted list.
@property ExcelMsoNumberedBulletStyle bulletStyle;  // Returns or sets a constant that represents the style of a bullet.
@property ExcelMsoBulletType bulletType;  // Returns or sets a constant that represents the type of bullet.
@property double relativeSize;  // Returns or sets the bullet size relative to the size of the first text character in the paragraph.
@property BOOL useTextColor;  // Determines whether the specified bullets are set to the color of the first text character in the paragraph.
@property BOOL useTextFont;  // Determines whether the specified bullets are set to the font of the first text character in the paragraph.
@property BOOL visible;  // Returns or sets a value that specifies whether the bullet is visible.

- (void) setBulletPictureFileName:(NSString *)FileName;  // Sets the graphics file to be used for bullets in a bulleted list.

@end

// Contains properties and methods that apply to line callouts.
@interface ExcelCalloutFormat : SBObject <ExcelGenericMethods>

@property BOOL accent;  // Returns or sets if a vertical accent bar separates the callout text from the callout line.
@property ExcelMsoCalloutAngleType angle;  // Returns or sets the angle of the callout line. If the callout line contains more than one line segment, this property returns or sets the angle of the segment that is farthest from the callout text box.
@property BOOL autoAttach;  // Returns or sets if the place where the callout line attaches to the callout text box changes depending on whether the origin of the callout line, where the callout points to, is to the left or right of the callout text box.
@property (readonly) BOOL autoLength;  // Returns if the length of the callout line is automatically set. Use the automatic length method to set this property to true, and use the custom length method to set this property to false.
@property BOOL border;  // Returns or sets whether the text in the specified callout is surrounded by a border.
@property (readonly) double calloutFormatLength;  // When the auto length property of the specified callout is set to false, the length property returns the length in points of the first segment of the callout line, the segment attached to the text callout box.
@property ExcelMsoCalloutType calloutFormatType;  // Returns or sets the callout type.
@property (readonly) double drop;  // For callouts with an explicitly set drop value, this property returns the vertical distance in points from the edge of the text bounding box to the place where the callout line attaches to the text box.
@property (readonly) ExcelMsoCalloutDropType dropType;  // Returns a value that indicates where the callout line attaches to the callout text box.
@property double gap;  // Returns or sets the horizontal distance in points between the end of the callout line and the text bounding box.

- (void) automaticLength;  // Specifies that the first segment of the callout line, the segment attached to the text callout box, be scaled automatically when the callout is moved.

@end

// Contains properties and methods that apply to connectors. A connector is a line that attaches two other shapes at points called connection sites.
@interface ExcelConnectorFormat : SBObject <ExcelGenericMethods>

@property (readonly) BOOL beginConnected;  // Returns true if the beginning of the specified connector is connected to a shape.
@property (copy, readonly) ExcelShape *beginConnectedShape;  // Returns a shape object that represents the shape that the beginning of the specified connector is attached to.
@property (readonly) NSInteger beginConnectionSite;  // Returns an integer that specifies the connection site that the beginning of a connector is connected to.
@property ExcelMsoConnectorType connectorFormatType;  // Returns or sets the connector type.
@property (readonly) BOOL endConnected;  // Returns true if the end of the specified connector is connected to a shape.
@property (copy, readonly) ExcelShape *endConnectedShape;  // Returns a shape object that represents the shape that the end of the specified connector is attached to.
@property (readonly) NSInteger endConnectionSite;  // Returns an integer that specifies the connection site that the end of a connector is connected to.

- (void) beginConnectConnectedShape:(ExcelShape *)connectedShape connectionSite:(NSInteger)connectionSite;  // Attaches the beginning of the specified connector to a specified shape. If there's already a connection between the beginning of the connector and another shape, that connection is broken.
- (void) beginDisconnect;  // Detaches the beginning of the specified connector from the shape it's attached to. This method doesn't alter the size or position of the connector: the beginning of the connector remains positioned at a connection site but is no longer connected.
- (void) endConnectConnectedShape:(ExcelShape *)connectedShape connectionSite:(NSInteger)connectionSite;  // Attaches the end of the specified connector to a specified shape. If there's already a connection between the end of the connector and another shape, that connection is broken.
- (void) endDisconnect;  // Detaches the end of the specified connector from the shape it's attached to. This method doesn't alter the size or position of the connector: the end of the connector remains positioned at a connection site but is no longer connected.

@end

// Represents fill formatting for a shape. A shape can have a solid, gradient, texture, pattern, picture, or semi-transparent fill.
@interface ExcelFillFormat : SBObject <ExcelGenericMethods>

@property (copy) NSArray<NSNumber *> *backColor;  // Returns or sets a RGB color that represents the background color for the specified fill or patterned line.
@property ExcelMsoThemeColorIndex backColorThemeIndex;  // Returns or sets the specified fill background color.
@property (readonly) ExcelMsoFillType fillFormatType;  // Returns the shape fill format type.
@property (copy) NSArray<NSNumber *> *foreColor;  // Returns or sets a RGB color that represents the foreground color for the fill, line, or shadow.
@property ExcelMsoThemeColorIndex foreColorThemeIndex;  // Returns or sets the specified foreground fill or solid color.
@property (readonly) ExcelMsoGradientColorType gradientColorType;  // Returns the gradient color type for the specified fill.
@property (readonly) double gradientDegree;  // Returns a value that indicates how dark or light a one-color gradient fill is. A value of zero means that black is mixed in with the shape's foreground color to form the gradient; a value of 1 means that white is mixed in. Values between 1 and zero blend.
@property (readonly) ExcelMsoGradientStyle gradientStyle;  // Returns the gradient style for the specified fill.
@property (readonly) NSInteger gradientVariant;  // Returns the gradient variant for the specified fill as an integer value from 1 to 4 for most gradient fills. If the gradient style is from center gradient, this property returns either 1 or 2.
@property (readonly) ExcelMsoPatternType pattern;  // Returns a value that represents the pattern applied to the specified fill or line.
@property (readonly) ExcelMsoPresetGradientType presetGradientType;  // Returns the preset gradient type for the specified fill.
@property (readonly) ExcelMsoPresetTexture presetTexture;  // Returns the preset texture for the specified fill.
@property BOOL rotateWithObject;  // Returns or sets whether the fill rotates with the specified shape.
@property ExcelMsoTextureAlignment textureAlignment;  // Returns or sets the texture alignment for the specified object.
@property double textureHorizontalScale;  // Returns or sets the texture alignment for the specified object.
@property (copy, readonly) NSString *textureName;  // Returns the name of the custom texture file for the specified fill.
@property double textureOffsetX;  // Returns or sets the texture alignment for the specified object.
@property double textureOffsetY;  // Returns or sets the texture alignment for the specified object.
@property BOOL textureTile;  // Returns the texture tile style for the specified fill.
@property (readonly) ExcelMsoTextureType textureType;  // Returns the texture type for the specified fill.
@property double textureVerticalScale;  // Returns or sets the texture alignment for the specified object.
@property double transparency;  // Returns or sets the degree of transparency of the specified fill, shadow, or line as a value between 0.0, opaque, and 1.0, clear.
@property BOOL visible;  // Returns or sets if the specified object, or the formatting applied to it, is visible.

- (void) deleteGradientStopStopIndex:(NSInteger)stopIndex;  // Removes a gradient stop.
- (void) insertGradientStopCustomColor:(NSArray<NSNumber *> *)customColor position:(double)position transparency:(double)transparency stopIndex:(NSInteger)stopIndex;  // Adds a stop to a gradient.
- (void) oneColorGradientGradientStyle:(ExcelMsoGradientStyle)gradientStyle variant:(NSInteger)variant degree:(double)degree;  // Sets the specified fill to a one-color gradient.
- (void) patternedPattern:(ExcelMsoPatternType)pattern;  // Sets the specified fill to a pattern.
- (void) presetGradientGradientStyle:(ExcelMsoGradientStyle)gradientStyle variant:(NSInteger)variant presetGradientType:(ExcelMsoPresetGradientType)presetGradientType;  // Sets the specified fill to a preset gradient.
- (void) presetTexturedPresetTexture:(ExcelMsoPresetTexture)presetTexture;  // Sets the specified fill to a preset texture.
- (void) solid;  // Sets the specified fill to a uniform color. Use this method to convert a gradient, textured, patterned, or background fill back to a solid fill.
- (void) twoColorGradientGradientStyle:(ExcelMsoGradientStyle)gradientStyle variant:(NSInteger)variant;  // Sets the specified fill to a two-color gradient.
- (void) userPicturePictureFile:(NSString *)pictureFile;  // Fills the specified shape with one large image.
- (void) userTexturedTextureFile:(NSString *)textureFile;  // Fills the specified shape with small tiles of an image.

@end

// Represents the glow formatting for a shape or range of shapes
@interface ExcelGlowFormat : SBObject <ExcelGenericMethods>

@property (copy) NSArray<NSNumber *> *color;  // Returns or sets the color for the specified glow format.
@property ExcelMsoThemeColorIndex colorThemeIndex;  // Returns or sets the color for the specified glow format.
@property double radius;  // Returns or sets the length of the radius for the specified glow format.


@end

// Represents one gradient stop.
@interface ExcelGradientStop : SBObject <ExcelGenericMethods>

@property (copy) NSArray<NSNumber *> *color;  // Returns or sets the color for the specified the gradient stop.
@property ExcelMsoThemeColorIndex colorThemeIndex;  // Returns or sets the color for the specified gradient stop.
@property double position;  // Returns or sets the position for the specified gradient stop expressed as a percent.
@property double transparency;  // Returns or sets a value representing the transparency of the gradient fill expressed as a percent.


@end

// Represents line and arrowhead formatting. For a line, the line format object contains formatting information for the line itself; for a shape with a border, this object contains formatting information for the shape's border.
@interface ExcelLineFormat : SBObject <ExcelGenericMethods>

@property (copy) NSArray<NSNumber *> *backColor;  // Returns or sets a RGB color that represents the background color for the specified fill or patterned line.
@property ExcelMsoThemeColorIndex backColorThemeIndex;  // Returns or sets the background color for a patterned line.
@property ExcelMsoArrowheadLength beginArrowheadLength;  // Returns or sets the length of the arrowhead at the beginning of the specified line.
@property ExcelMsoArrowheadStyle beginArrowheadStyle;  // Returns or sets the style of the arrowhead at the beginning of the specified line.
@property ExcelMsoArrowheadWidth beginArrowheadWidth;  // Returns or sets the width of the arrowhead at the beginning of the specified line.
@property ExcelMsoLineDashStyle dashStyle;  // Returns or sets the dash style for the specified line.
@property ExcelMsoArrowheadLength endArrowheadLength;  // Returns or sets the length of the arrowhead at the end of the specified line.
@property ExcelMsoArrowheadStyle endArrowheadStyle;  // Returns or sets the style of the arrowhead at the end of the specified line.
@property ExcelMsoArrowheadWidth endArrowheadWidth;  // Returns or sets the width of the arrowhead at the end of the specified line.
@property (copy) NSArray<NSNumber *> *foreColor;  // Returns or sets a RGB color that represents the foreground color for the fill, line, or shadow.
@property ExcelMsoThemeColorIndex foreColorThemeIndex;  // Returns or sets the foreground color for the line.
@property ExcelMsoLineStyle lineStyle;  // Returns or sets the line format style.
@property ExcelMsoPatternType pattern;  // Returns or sets a value that represents the pattern applied to the specified fill or line.
@property double transparency;  // Returns or sets the degree of transparency of the specified fill, shadow, or line as a value between 0.0, opaque and 1.0, clear.
@property BOOL visible;  // Returns or sets if the specified object, or the formatting applied to it, is visible.
@property double weight;  // Returns or sets the thickness of the specified line in points.


@end

// Represents a line graphic object.
@interface ExcelLine : SBObject <ExcelGenericMethods>

@property ExcelXlArrowHeadLength arrowheadLength;  // Returns or sets the length of an arrowhead
@property ExcelXlArrowHeadStyle arrowheadStyle;  // Returns or sets the style of an arrowhead.
@property ExcelXlArrowHeadWidth arrowheadWidth;  // Returns or sets the width of an arrowhead.
@property (copy, readonly) ExcelBorder *border;  // Returns a border object that represents the border of the object. 
@property (copy, readonly) ExcelRange *bottomRightCell;  // Returns the bottom right cell of the range the control is occupying.
@property (readonly) NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property double height;  // Returns or set the height of the object.
@property double leftPosition;  // Returns or sets the position of the specified object, in points.
@property BOOL locked;  // Returns or sets if the object is locked, if false the object can be modified when the sheet is protected.
@property (copy) NSString *name;  // Returns or sets the name of the object.
@property (copy) NSString *onAction;  // Returns or sets the name of a string of AppleScript commands that will be executed when the object is clicked on. Please note that if AppleScript commands are set they will not be saved with the document.
@property ExcelXlPlacement placement;  // Returns or sets how the object is placed on the worksheet.
@property BOOL printObject;  // Returns or sets if this object is printed.
@property double top;  // Returns the top position of the specified object, in points.
@property (copy, readonly) ExcelRange *topLeftCell;  // Returns a range object that represents the cell that lies under the upper-left corner of the specified object. 
@property BOOL visible;  // Returns or sets if the object is visible.
@property double width;  // Returns or sets  the width of the object.
@property (readonly) NSInteger zOrderPosition;  // Returns the z-order position of the object.


@end

// Represents a Microsoft Office theme.
@interface ExcelOfficeTheme : SBObject <ExcelGenericMethods>

@property (copy, readonly) ExcelThemeColorScheme *themeColorScheme;  // Returns the color scheme of a Microsoft Office theme.
@property (copy, readonly) ExcelThemeEffectScheme *themeEffectScheme;  // Returns the effects scheme of a Microsoft Office theme.
@property (copy, readonly) ExcelThemeFontScheme *themeFontScheme;  // Returns the font scheme of a Microsoft Office theme.


@end

// Represents an oval graphic.
@interface ExcelOval : SBObject <ExcelGenericMethods>

- (SBElementArray<ExcelCharacter *> *) characters;

@property NSInteger addIndent;  // Returns or sets if text is automatically indented when the text alignment in a cell is set to equal distribution either horizontally or vertically.
@property BOOL autoScaleFont;  // Returns or sets  if the text in the object changes font size when the object size changes.
@property NSInteger autoSize;  // Returns or sets if the size of the specified object is changed automatically to fit text within its boundaries.
@property (copy, readonly) ExcelBorder *border;  // Returns a border object that represents the border of the object. 
@property (readonly) NSInteger bottomRightCell;  // Returns the bottom right cell of the range the control is occupying.
@property NSInteger caption;  // Returns or sets the caption for this object.
@property (readonly) NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property (readonly) NSInteger fontObject;  // Returns a font object that represents the font of the specified object.
@property NSInteger formula;  // Returns or sets the object's formula, in A1-style notation and in the language of the macro.
@property NSInteger height;  // Returns or set the height of the object.
@property NSInteger horizontalAlignment;  // Returns or sets the horizontal alignment for the object.
@property (readonly) NSInteger interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property NSInteger leftPosition;  // Returns or sets the position of the specified object, in points.
@property NSInteger locked;  // Returns or sets if the object is locked, if false the object can be modified when the sheet is protected.
@property NSInteger lockedText;  // Returns or sets whether the control's text is locked for editing.
@property NSInteger name;  // Returns or sets the name of the object.
@property (copy) NSString *onAction;  // Returns or sets the name of a string of AppleScript commands that will be executed when the object is clicked on. Please note that if AppleScript commands are set they will not be saved with the document.
@property NSInteger orientation;  // May also be a number value from -90 to 90 degrees.
@property NSInteger placement;  // Returns or sets how the object is placed on the worksheet.
@property NSInteger printObject;  // Returns or sets if this object is printed.
@property ExcelXLDefaultSheetDir readingOrder;  // Returns or sets the reading order for the specified object.
@property NSInteger shadow;  // Returns or sets if the object has a shadow.
@property NSInteger stringValue;  // Returns or sets the text of the specified object.
@property NSInteger top;  // Returns the top position of the specified object, in points.
@property (readonly) NSInteger topLeftCell;  // Returns a range object that represents the cell that lies under the upper-left corner of the specified object. 
@property NSInteger verticalAlignment;  // Returns or sets the vertical alignment of the object.
@property NSInteger visible;  // Returns or sets if the object is visible.
@property NSInteger width;  // Returns or sets  the width of the object.
@property (readonly) NSInteger zOrderPosition;  // Returns the z-order position of the object.


@end

@interface ExcelParagraphFormat : SBObject <ExcelGenericMethods>

@property ExcelMsoParagraphAlignment alignment;  // Returns or sets a value specifying the alignment of the paragraph.
@property ExcelMsoBaselineAlignment baselineAlignment;  // Returns or sets a constant that represents the vertical position of fonts in a paragraph.
@property (copy, readonly) ExcelBulletFormat *bullet;  // Returns a bullet format object for the paragraph.
@property BOOL eastAsianLineBreakLevel;  // Returns or sets the East Asian line break control level for the specified paragraph.
@property double firstLineIndent;  // Returns or sets the value, in points, for a first line or hanging indent.
@property BOOL hangingPunctuation;  // Determines whether hanging punctuation is enabled for the specified paragraphs.
@property NSInteger indentLevel;  // Returns or sets a value representing the indent level assigned to text in the selected paragraph.
@property double leftIndent;  // Returns or sets a value that represents the left indent value, in points, for the specified paragraphs.
@property BOOL lineRuleAfter;  // Determines whether line spacing after the last line in each paragraph is set to a specific number of points or lines.
@property BOOL lineRuleBefore;  // Determines whether line spacing before the first line in each paragraph is set to a specific number of points or lines.
@property BOOL lineRuleWithin;  // Determines whether line spacing between base lines is set to a specific number of points or lines.
@property double rightIndent;  // Returns or sets the right indent, in points, for the specified paragraphs.
@property double spaceAfter;  // Returns or sets the amount of spacing, in points, after the specified paragraph.
@property double spaceBefore;  // Returns or sets the spacing, in points, before the specified paragraphs.
@property double spaceWithin;  // Returns or sets the amount of space between base lines in the specified paragraph, in points or lines.
@property (copy) id textDirection;  // Returns or sets the text direction for the specified paragraph.
@property BOOL wordWrap;  // Determines whether the application wraps the Latin text in the middle of a word in the specified paragraphs.


@end

// Contains properties and methods that apply to pictures.
@interface ExcelPictureFormat : SBObject <ExcelGenericMethods>

@property double brightness;  // Returns or sets the brightness of the specified picture . The value for this property must be a number from 0.0, dimmest to 1.0, brightest.
@property ExcelMsoPictureColorType colorType;  // Returns or sets the type of color transformation applied to the specified picture.
@property double contrast;  // Returns or sets the contrast for the specified picture. The value for this property must be a number from 0.0, the least contrast to 1.0, the greatest contrast.
@property double cropBottom;  // Returns or sets the number of points that are cropped off the bottom of the specified picture.
@property double cropLeft;  // Returns or sets the number of points that are cropped off the left side of the specified picture.
@property double cropRight;  // Returns or sets the number of points that are cropped off the right side of the specified picture.
@property double cropTop;  // Returns or sets the number of points that are cropped off the top of the specified picture.
@property (copy) id transparencyColor;  // Returns or sets the transparent color for the specified picture as aRGB color. For this property to take effect, the transparent background property must be set to true.
@property BOOL transparentBackground;  // Returns or sets if the parts of the picture that are defined with a transparent color actually appear transparent.


@end

// Represents a rectangle graphic object.
@interface ExcelRectangle : SBObject <ExcelGenericMethods>

- (SBElementArray<ExcelCharacter *> *) characters;

@property NSInteger addIndent;  // Returns or sets if text is automatically indented when the text alignment in a cell is set to equal distribution either horizontally or vertically.
@property BOOL autoScaleFont;  // Returns or sets  if the text in the object changes font size when the object size changes.
@property NSInteger autoSize;  // Returns or sets if the size of the specified object is changed automatically to fit text within its boundaries.
@property (copy, readonly) ExcelBorder *border;  // Returns a border object that represents the border of the object. 
@property (readonly) NSInteger bottomRightCell;  // Returns the bottom right cell of the range the control is occupying.
@property NSInteger caption;  // Returns or sets the caption for this object.
@property (readonly) NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property (readonly) NSInteger fontObject;  // Returns a font object that represents the font of the specified object.
@property NSInteger formula;  // Returns or sets the object's formula, in A1-style notation and in the language of the macro.
@property NSInteger height;  // Returns or set the height of the object.
@property NSInteger horizontalAlignment;  // Returns or sets the horizontal alignment for the object.
@property (readonly) NSInteger interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property NSInteger leftPosition;  // Returns or sets the position of the specified object, in points.
@property NSInteger locked;  // Returns or sets if the object is locked, if false the object can be modified when the sheet is protected.
@property NSInteger lockedText;  // Returns or sets whether the control's text is locked for editing.
@property NSInteger name;  // Returns or sets the name of the object.
@property (copy) NSString *onAction;  // Returns or sets the name of a string of AppleScript commands that will be executed when the object is clicked on. Please note that if AppleScript commands are set they will not be saved with the document.
@property NSInteger orientation;  // May also be a number value from -90 to 90 degrees.
@property NSInteger placement;  // Returns or sets how the object is placed on the worksheet.
@property NSInteger printObject;  // Returns or sets if this object is printed.
@property ExcelXLDefaultSheetDir readingOrder;  // Returns or sets the reading order for the specified object.
@property NSInteger roundedCorners;  // Returns or sets if the rectangle has rounded corners
@property NSInteger shadow;  // Returns or sets if the object has a shadow.
@property NSInteger stringValue;  // Returns or sets the text of the specified object.
@property NSInteger top;  // Returns the top position of the specified object, in points.
@property (readonly) NSInteger topLeftCell;  // Returns a range object that represents the cell that lies under the upper-left corner of the specified object. 
@property NSInteger verticalAlignment;  // Returns or sets the vertical alignment of the object.
@property NSInteger visible;  // Returns or sets if the object is visible.
@property NSInteger width;  // Returns or sets  the width of the object.
@property (readonly) NSInteger zOrderPosition;  // Returns the z-order position of the object.


@end

// Represents the reflection effect in Office graphics.
@interface ExcelReflectionFormat : SBObject <ExcelGenericMethods>

@property ExcelMsoReflectionType reflectionType;  // Returns or sets the type of the reflection format object.


@end

@interface ExcelRulerLevel : SBObject <ExcelGenericMethods>

@property double firstMargin;  // Returns or sets the first-line indent for the specified outline level, in points.
@property double leftMargin;  // Returns or sets the left indent for the specified outline level, in points.


@end

// Represents the ruler for the text in the specified shape or for all text in the specified text style. Contains tab stops and the indentation settings for text outline levels.
@interface ExcelRuler : SBObject <ExcelGenericMethods>

- (SBElementArray<ExcelRulerLevel *> *) rulerLevels;


@end

// Represents shadow formatting for a shape.
@interface ExcelShadowFormat : SBObject <ExcelGenericMethods>

@property double blur;  // Returns or sets the blur, in points, of the specified shadow.
@property (copy) NSArray<NSNumber *> *foreColor;  // Returns or sets a RGB color that represents the foreground color for the fill, line, or shadow.
@property ExcelMsoThemeColorIndex foreColorThemeIndex;  // Returns or sets the foreground color for the shadow format.
@property BOOL obscured;  // Returns or sets if the shadow of the specified shape appears filled in and is obscured by the shape, even if the shape has no fill. If false the shadow has no fill and the outline of the shadow is visible through the shape if the shape has no fill.
@property double offsetX;  // Returns or sets the horizontal offset in points of the shadow from the specified shape. A positive value offsets the shadow to the right of the shape; a negative value offsets it to the left.
@property double offsetY;  // Returns or sets the vertical offset in points of the shadow from the specified shape. A positive value offsets the shadow below the shape; a negative value offsets it above the shape.
@property BOOL rotateWithShape;  // Returns or sets whether to rotate the shadow when rotating the shape.
@property ExcelMsoShadowStyle shadowStyle;  // Returns or sets the style of shadow formatting to apply to a shape.
@property ExcelMsoShadowType shadowType;  // Returns or sets the shape shadow type.
@property double size;  // Returns or sets the width of the shadow.
@property double transparency;  // Returns or sets the degree of transparency of the specified fill, shadow, or line as a value between 0.0, opaque and 1.0, clear.
@property BOOL visible;  // Returns or sets if the specified object, or the formatting applied to it, is visible.


@end

// Contains font attributes such as font name, size, and color, for an object.
@interface ExcelShapeFont : SBObject <ExcelGenericMethods>

@property (copy) NSString *ASCIIName;  // Returns or sets the font used for Latin text; characters with character codes from 0 through 127.
@property double baseLineOffset;  // Returns or sets a value specifying the horizontaol offset of the selected font.
@property BOOL bold;  // Returns or sets a value specifying whether the font should be bold.
@property ExcelMsoTextCaps capsType;  // Returns or sets a value specifying how the text should be capitalized.
@property (copy) NSString *eastAsianName;  // Returns or sets the font name used for Asian text.
@property (readonly) BOOL embedable;  // Returns a value indicating whether the font can be embedded in a page.
@property (readonly) BOOL embedded;  // Returns a value specifying whether the font is embedded in a page.
@property BOOL equalizeCharacterHeight;  // Returns or sets a value specifying whether the text should have the same horizontal height.
@property (copy, readonly) ExcelFillFormat *fillFormat;  // Returns a fill format object that contains fill formatting properties for the specified font.
@property (copy) NSArray<NSNumber *> *fontColor;
@property ExcelMsoThemeColorIndex fontColorThemeIndex;  // Returns or sets the color for the specified font.
@property (copy) NSString *fontName;  // Returns or sets a value specifying the font to use for a selection.
@property (copy) NSString *fontNameOther;  // Returns or sets the font used for characters whose character set numbers are greater than 127.
@property double fontSize;  // Returns or sets the font size.
@property (copy, readonly) ExcelGlowFormat *glowFormat;  // Returns the formatting properties for a glow effect.
@property (copy) NSArray<NSNumber *> *highlightColor;  // Returns or sets the text highlight color for object.
@property ExcelMsoThemeColorIndex highlightColorThemeIndex;  // Returns or sets the highlight color for the specified text.
@property BOOL italic;  // Returns or sets a value specifying whether the text for a selection is italic.
@property double kerning;  // Returns or sets a value specifying the amount of spacing between text characters.
@property (copy, readonly) ExcelLineFormat *lineFormat;  // Returns a value specifiying the format of a line.
@property (copy, readonly) ExcelReflectionFormat *reflectionFormat;  // Returns the formatting properties for a reflection effect.
@property (copy, readonly) ExcelShadowFormat *shadowFormat;  // Returns the value specifying the type of shadow effect for the selection of text.
@property ExcelMsoSoftEdgeType softEdgeType;  // Returns or sets the type soft edge format object.
@property double spacing;  // Returns or sets a value specifying the spacing between characters in a selection of text.
@property ExcelMsoTextStrike strikeType;  // Returns or sets a value specifying the strike format used for a selection of text.
@property BOOL subscript;  // Returns or sets a value specifying that the selected text should be displayed a subscript.
@property BOOL superscript;  // Returns or sets a value specifying that the selected text should be displayed a superscript.
@property (copy) NSArray<NSNumber *> *underlineColor;  // Returns a value specifying the color of the underline for the selected text.
@property ExcelMsoThemeColorIndex underlineColorThemeIndex;  // Returns a value specifying the color of the underline for the selected text.
@property ExcelMsoTextUnderlineType underlineStyle;  // Returns or sets a value specifying the text effect for the selected text.
@property ExcelMsoPresetTextEffect wordArtStylesFormat;  // Returns or sets a value specifying the text effect for the selected text.


@end

// Represents the shape text frame in a shape object. Contains the text in the text frame as well as the properties and methods that control the alignment and anchoring of the text frame.
@interface ExcelShapeTextFrame : SBObject <ExcelGenericMethods>

@property (readonly) BOOL hasText;  // Returns whether the specified text frame has text.
@property ExcelMsoHorizontalAnchor horizontalAnchor;
@property double marginBottom;  // Returns or sets the distance, in points, between the bottom of the text frame and the bottom of the inscribed rectangle of the shape that contains the text.
@property double marginLeft;  // Returns or sets the distance, in points, between the left edge of the text frame and the left edge of the inscribed rectangle of the shape that contains the text.
@property double marginRight;  // Returns or sets the distance, in points, between the right edge of the text frame and the right edge of the inscribed rectangle of the shape that contains the text.
@property double marginTop;  // Returns or sets the distance, in points, between the top of the text frame and the top of the inscribed rectangle of the shape that contains the text.
@property ExcelMsoTextOrientation orientation;  // Returns or sets the text orientation.
@property ExcelMsoPathFormat pathFormat;  // Returns or sets the path type for the specified text frame.
@property (copy, readonly) ExcelRuler *ruler;
@property (copy, readonly) ExcelTextColumn *textColumn;  // Returns the text column object that represents the columns within the text frame.
@property (copy, readonly) ExcelTextRange *textRange;
@property (copy, readonly) ExcelThreeDFormat *threeDFormat;  // Returns the 3-D-effect formatting properties for the specified text.
@property ExcelMsoVerticalAnchor verticalAnchor;
@property ExcelMsoWarpFormat warpFormat;  // Returns or sets the warp type for the specified text frame.
@property BOOL wordWrap;  // Returns or sets text break lines within or past the boundaries of the shape.
@property ExcelMsoAutoSize wordartAutoSize;  // The size of the specified object that changes automatically to fit text within its boundaries.

- (void) ungroup;

@end

// Represents an object in the drawing layer.
@interface ExcelShape : SBObject <ExcelGenericMethods>

- (SBElementArray<ExcelAdjustment *> *) adjustments;

@property (copy) NSString *alternativeText;  // Returns or sets the descriptive alternative text string for a Shape object when the object is saved to a Web page.
@property ExcelMsoAutoShapeType autoShapeType;  // Returns or sets the shape type for the specified shape object, which must represent an auto-shape.
@property ExcelMsoBackgroundStyleIndex backgroundStyle;  // Returns or sets the background style.
@property ExcelMsoBlackWhiteMode blackWhiteMode;  // Returns or sets a value that indicates how the specified shape appears when the presentation is viewed in black-and-white mode.
@property (copy, readonly) ExcelRange *bottomRightCell;  // Returns a range object that represents the cell that lies under the lower-right corner of the object.
@property (copy, readonly) ExcelChart *chart;  // Returns a chart object that represents the chart contained in the shape.
@property (readonly) BOOL child;  // True if the shape is a child shape.
@property (readonly) NSInteger connectionSiteCount;  // Returns the number of connection sites on the specified shape.
@property (readonly) BOOL connector;  // Returns true if the specified shape is a connector.
@property (copy, readonly) ExcelConnectorFormat *connectorFormat;  // Returns a connector format object that contains connector formatting properties if this shape is a connector.
@property (readonly) ExcelMsoConnectorType connectorType;  // Returns the connector type if this shape is a connector.
@property (copy, readonly) ExcelFillFormat *fillFormat;  // Returns a fill format object that contains fill formatting properties for the specified shape.
@property (copy, readonly) ExcelGlowFormat *glowFormat;  // Returns the formatting properties for a glow effect.
@property (readonly) BOOL hasChart;  // True if the specified shape has a chart.
@property double height;  // Returns or sets the height of the object.
@property (readonly) BOOL horizontalFlip;  // Returns true if the specified shape is flipped around the horizontal axis.
@property (copy, readonly) ExcelHyperlink *hyperlink;  // Returns a hyperlink object that represents the hyperlink for the shape.
@property double leftPosition;  // Returns or sets the left position of the specified object, in points.
@property (copy, readonly) ExcelLineFormat *lineFormat;  // Returns a line format object for this shape.
@property BOOL lockAspectRatio;  // Returns or sets if the specified shape retains its original proportions when you resize it. If false, you can change the height and width of the shape independently of one another when you resize it.
@property BOOL locked;  // Returns or sets if the object is locked. If false, the object can be modified when the sheet is protected.
@property (copy) NSString *name;  // Returns or sets the name of the object.
@property (copy, readonly) ExcelShape *parentgroup;  // Returns a Shape object that represents the common parent shape of a child shape or a range of child shapes.
@property ExcelXlPlacement placement;  // Returns or sets how the object is placed on the worksheet.
@property (copy, readonly) ExcelReflectionFormat *reflectionFormat;  // Returns the formatting properties for a reflection effect.
@property double rotation;  // Returns or sets the rotation of the shape, in degrees.
@property (copy, readonly) ExcelShadowFormat *shadowFormat;  // Returns a shadow format object for this shape.
@property (copy) NSString *shapeOnAction;  // Returns or sets the name of a string of AppleScript commands that will be executed when the object is clicked on. Please note that if AppleScript commands are set they will not be saved with the document.
@property ExcelMsoShapeStyleIndex shapeStyle;  // Returns or sets the shape style corresponding to the Quick Styles.
@property (copy, readonly) ExcelShapeTextFrame *shapeTextFrame;  // Returns a shape text frame object that contains the alignment and anchoring properties for the specified shape.
@property (readonly) ExcelMsoShapeType shapeType;  // Returns the shape type.
@property (copy, readonly) ExcelSoftEdgeFormat *softEdgeFormat;  // Returns the formatting properties for a soft edge effect.
@property (copy, readonly) ExcelTextFrame *textFrame;  // Returns a text frame object that contains the alignment and anchoring properties for the specified shape.
@property (copy, readonly) ExcelThreeDFormat *threeDFormat;  // Returns a threeD format object that contains 3-D-effect formatting properties for the specified shape.
@property double top;  // Returns or sets the top position of the specified object, in points.
@property (copy, readonly) ExcelRange *topLeftCell;  // Returns a range object that represents the cell that lies under the upper-left corner of the specified object.
@property (readonly) BOOL verticalFlip;  // Returns true if the specified shape is flipped around the vertical axis.
@property BOOL visible;  // Returns or sets if the object is visible
@property double width;  // Returns or sets  the width of the object.
@property (copy, readonly) ExcelWordArtFormat *wordArtFormat;  // Returns the formatting properties for a word art effect.
@property (readonly) NSInteger zOrderPosition;  // Returns the position of the specified shape in the z-order. To set the shape's position in the z-order, use the z order method.

- (void) apply;  // Applies to the specified shape formatting that's been copied by using the pick up method.
- (void) flipFlipCmd:(ExcelMsoFlipCmd)flipCmd;  // Flips the specified shape around its horizontal or vertical axis.
- (void) pickUp;  // Copies the formatting of the specified shape. Use the apply method to apply the copied formatting to another shape.
- (void) rerouteConnections;  // Reroutes connectors so that they take the shortest possible path between the shapes they connect. To do this, the reroute connections method may detach the ends of a connector and reattach them to different connecting sites on the connected shapes.
- (void) saveAsPicturePictureType:(ExcelMsoPictureType)pictureType fileName:(NSString *)fileName;  // Saves the shape in the requested file using the stated graphic format
- (void) setShapesDefaultProperties;  // Applies the formatting for the specified shape to the default shape. Shapes created after this method has been used will have this formatting applied by default.
- (void) zOrderZOrderCommand:(ExcelMsoZOrderCmd)zOrderCommand;  // Moves the specified shape in front of or behind other shapes in the collection, that is, changes the shape's position in the z-order.

@end

@interface ExcelCallout : ExcelShape

@property (copy, readonly) ExcelCalloutFormat *calloutFormat;  // Returns a connector format object that contains connector formatting properties.
@property (readonly) ExcelMsoCalloutType calloutType;  // Returns the type of callout.


@end

@interface ExcelPicture : ExcelShape

@property (copy, readonly) NSString *fileName;  // Returns he name of the file that has the picture.
@property (readonly) BOOL linkToFile;  // Returns if the picture is lined to the file.
@property (copy, readonly) ExcelPictureFormat *pictureFormat;  // Returns a picture format object for this picture.
@property (readonly) BOOL saveWithDocument;  // Returns if the picture should be saved with the document.

- (void) scaleHeightFactor:(double)factor relativeToOriginalSize:(BOOL)relativeToOriginalSize scale:(ExcelMsoScaleFrom)scale;  // Scales the height of the shape by a specified factor.
- (void) scaleWidthFactor:(double)factor relativeToOriginalSize:(BOOL)relativeToOriginalSize scale:(ExcelMsoScaleFrom)scale;  // Scales the width of the shape by a specified factor. For pictures, you can indicate whether you want to scale the shape relative to the original size or relative to the current size.

@end

@interface ExcelShapeConnector : ExcelShape

@property (copy, readonly) ExcelConnectorFormat *connectorFormat;  // Returns a connector format object that contains connector formatting properties.
@property (readonly) ExcelMsoConnectorType connectorType;  // Returns the connector type.


@end

// The line shape uses begin line X, begin line Y, end line X, and end line Y when created
@interface ExcelShapeLine : ExcelShape

@property double beginLineX;  // Returns or sets the beginning X position of the line.
@property double beginLineY;  // Returns or sets the beginning Y position of the line.
@property double endLineX;  // Returns or sets the ending X position of the line.
@property double endLineY;  // Returns or sets the ending Y position of the line.


@end

@interface ExcelShapeTextbox : ExcelShape

@property (readonly) ExcelMsoTextOrientation textOrientation;  // Returns the text orientation of the object.


@end

// Represents the soft edge formatting for a shape or range of shapes
@interface ExcelSoftEdgeFormat : SBObject <ExcelGenericMethods>

@property ExcelMsoSoftEdgeType softEdgeType;  // Returns or sets the type soft edge format object.


@end

// Represents a single tab stop.
@interface ExcelTabStop : SBObject <ExcelGenericMethods>

@property double tabPosition;  // Returns or sets the position of a tab stop relative to the left margin.
@property ExcelMsoTabStopType tabStopType;  // Returns or sets the type of the tab stop object.


@end

// Represents a single text column.
@interface ExcelTextColumn : SBObject <ExcelGenericMethods>

@property NSInteger columnNumber;  // Returns or sets the index of the text column object.
@property double spacing;  // Returns or sets the spacing between text columns in a text column object.
@property (copy) id textDirection;  // Returns or sets the direction of text in the text column object.


@end

// Represents the text frame in a shape object. Contains the text in the text frame as well as the properties and methods that control the alignment and anchoring of the text frame.
@interface ExcelTextFrame : SBObject <ExcelGenericMethods>

- (SBElementArray<ExcelCharacter *> *) characters;

@property double marginBottom;  // Returns or sets the distance, in points, between the bottom of the text frame and the bottom of the inscribed rectangle of the shape that contains the text.
@property double marginLeft;  // Returns or sets the distance, in points, between the left edge of the text frame and the left edge of the inscribed rectangle of the shape that contains the text.
@property double marginRight;  // Returns or sets the distance, in points, between the right edge of the text frame and the right edge of the inscribed rectangle of the shape that contains the text.
@property double marginTop;  // Returns or sets the distance, in points, between the top of the text frame and the top of the inscribed rectangle of the shape that contains the text.
@property ExcelMsoTextOrientation orientation;  // Returns or sets the text orientation.
@property ExcelXLDefaultSheetDir readingOrder;  // Returns or sets the reading order for the specified object.


@end

// Represents the color scheme of an Office Theme
@interface ExcelThemeColorScheme : SBObject <ExcelGenericMethods>

- (SBElementArray<ExcelThemeColor *> *) themeColors;

- (NSArray<NSNumber *> *) getCustomColorName:(NSString *)name;  // Returns the custom color for the specified Microsoft Office theme.
- (void) loadThemeColorSchemeFileName:(NSString *)fileName;  // Loads the color scheme of a Microsoft Office theme from a file
- (void) saveThemeColorSchemeFileName:(NSString *)fileName;  // Saves the color scheme of a Microsoft Office theme to a file.

@end

// Represents a color in the color scheme of a Microsoft Office 2007 theme.
@interface ExcelThemeColor : SBObject <ExcelGenericMethods>

@property (copy) NSArray<NSNumber *> *RGB;  // Returns or sets a value of a color in the color scheme of a Microsoft Office theme.
@property (readonly) ExcelMsoThemeColorSchemeIndex themeColorSchemeIndex;  // Returns the index value a color scheme of a Microsoft Office theme.


@end

// Represents the effect scheme of a Microsoft Office theme.
@interface ExcelThemeEffectScheme : SBObject <ExcelGenericMethods>

- (void) loadThemeEffectSchemeFileName:(NSString *)fileName;  // Loads the effects scheme of a Microsoft Office theme from a file

@end

// Represents the font scheme of a Microsoft Office theme.
@interface ExcelThemeFontScheme : SBObject <ExcelGenericMethods>

- (SBElementArray<ExcelMinorThemeFont *> *) minorThemeFonts;
- (SBElementArray<ExcelMajorThemeFont *> *) majorThemeFonts;

- (void) loadThemeFontSchemeFileName:(NSString *)fileName;  // Loads the font scheme of a Microsoft Office theme from a file.
- (void) saveThemeFontSchemeFileName:(NSString *)fileName;  // Saves the font scheme of a Microsoft Office theme to a file.

@end

// Represents a container for the font schemes of a Microsoft Office theme.
@interface ExcelThemeFont : SBObject <ExcelGenericMethods>

@property (copy) NSString *name;  // Returns or sets a value specifying the font to use for a selection.


@end

// Represents a container for the font schemes of a Microsoft Office theme.
@interface ExcelMajorThemeFont : ExcelThemeFont


@end

// Represents a container for the font schemes of a Microsoft Office theme.
@interface ExcelMinorThemeFont : ExcelThemeFont


@end

// Represents a shape's three-dimensional formatting.
@interface ExcelThreeDFormat : SBObject <ExcelGenericMethods>

@property double ZDistance;  // Returns or sets a Single that represents the distance from the center of an object or text.
@property double bevelBottomDepth;  // Returns or sets the depth/height of the bottom bevel.
@property double bevelBottomInset;  // Returns or sets the inset size/width for the bottom bevel.
@property ExcelMsoBevelType bevelBottomType;  // Returns or sets the bevel type for the bottom bevel.
@property double bevelTopDepth;  // Returns or sets the depth/height of the top bevel.
@property double bevelTopInset;  // Returns or sets the inset size/width for the top bevel.
@property ExcelMsoBevelType bevelTopType;  // Returns or sets the bevel type for the top bevel.
@property (copy, readonly) NSArray<NSNumber *> *contourColor;  // Returns or sets the color of the contour of an object or text.
@property ExcelMsoThemeColorIndex contourColorThemeIndex;  // Returns or sets the color for the specified contour.
@property double contourWidth;  // Returns or sets the width of the contour of an object or text.
@property double depth;  // Returns or sets the depth of the shape's extrusion.
@property (copy, readonly) NSArray<NSNumber *> *extrusionColor;  // Returns or sets a RGB color that represents the color of the shape's extrusion.
@property ExcelMsoThemeColorIndex extrusionColorThemeIndex;  // Returns or sets the color for the specified extrusion.
@property ExcelMsoExtrusionColorType extrusionColorType;  // Returns or sets a value that indicates what will determine the extrusion color.
@property double fieldOfView;  // Returns or sets the amount of perspective for an object or text.
@property double lightAngle;  // Returns or sets the angle of the lighting.
@property BOOL perspective;  // Returns or sets if the extrusion appears in perspective that is, if the walls of the extrusion narrow toward a vanishing point. If false, the extrusion is a parallel, or orthographic, projection that is, if the walls don't narrow toward a vanishing point.
@property (readonly) ExcelMsoPresetCamera presetCamera;  // Returns a constant that represents the camera preset.
@property (readonly) ExcelMsoPresetExtrusionDirection presetExtrusionDirection;  // Returns or sets the direction taken by the extrusion's sweep path leading away from the extruded shape, the front face of the extrusion.
@property ExcelMsoPresetLightingDirection presetLightingDirection;  // Returns or sets the position of the light source relative to the extrusion.
@property ExcelMsoLightRigType presetLightingRig;  // Returns a constant that represents the lighting preset.
@property ExcelMsoPresetLightingSoftness presetLightingSoftness;  // Returns or sets the intensity of the extrusion lighting.
@property ExcelMsoPresetMaterial presetMaterial;  // Returns or sets the extrusion surface material.
@property (readonly) ExcelMsoPresetThreeDFormat presetThreeDFormat;
@property BOOL projectText;  // Returns or sets whether text on a shape rotates with shape.
@property double rotationX;  // Returns or sets the rotation of the extruded shape around the x-axis in degrees. A positive value indicates upward rotation; a negative value indicates downward rotation.
@property double rotationY;  // Returns or sets the rotation of the extruded shape around the y-axis, in degrees. A positive value indicates rotation to the left; a negative value indicates rotation to the right.
@property double rotationZ;  // Returns or sets the rotation of the extruded shape around the z-axis, in degrees. A positive value indicates clockwise rotation; a negative value indicates anti-clockwise rotation.
@property BOOL visible;  // Returns or sets if the specified object, or the formatting applied to it, is visible.

- (void) resetRotation;  // Resets the extrusion rotation around the x-axis and the y-axis to zero so that the front of the extrusion faces forward. This method doesn't reset the rotation around the z-axis.

@end

// Contains properties and methods that apply to WordArt objects.
@interface ExcelWordArtFormat : SBObject <ExcelGenericMethods>

@property ExcelMsoTextEffectAlignment alignment;  // Returns or sets a constant that represents the alignment for the specified text effect.
@property BOOL bold;  // Returns if the text of the word art shape is formatted as bold.
@property (copy) NSString *fontName;  // Returns or sets the font name of the font used by this word art shape.
@property double fontSize;  // Returns or sets the font size of the font used by this word art shape.
@property BOOL italic;  // Returns if the text of the word art shape is formatted as italic.
@property BOOL kernedPairs;  // Returns or sets if character pairs in a WordArt object have been kerned. 
@property BOOL normalizedHeight;  // Returns or sets if all characters, both uppercase and lowercase, in the specified WordArt are the same height.
@property ExcelMsoPresetTextEffectShape presetShape;  // Returns or sets the shape of the specified WordArt.
@property ExcelMsoPresetTextEffect presetWordArtEffect;  // Returns or sets the style of the specified WordArt.
@property BOOL rotatedChars;  // Returns or sets if characters in the specified WordArt are rotated 90 degrees relative to the WordArt's bounding shape. If false, characters in the specified WordArt retain their original orientation relative to the bounding shape.
@property double tracking;  // Returns or sets the ratio of the horizontal space allotted to each character in the specified WordArt in relation to the width of the character. Can be a value from zero through 5.
@property (copy) NSString *wordArtText;  // Returns or sets the text associated with this word art shape.

- (void) toggleVerticalText;

@end



/*
 * Text Suite
 */

// Represents characters in an object that contains text.
@interface ExcelCharacter : SBObject <ExcelGenericMethods>

@property (copy) NSString *content;  // Returns or sets the text for the specified object.
@property (copy, readonly) ExcelFont *fontObject;  // Returns a font object that represents the font of the specified object.
@property (copy) NSString *phoneticCharacters;  // Returns or sets the phonetic text in the specified characters object.

- (void) insertIntoString:(NSString *)string;  // Inserts a string preceding the selected characters.

@end

// Contains font attributes, such as font name, size, and color, for an object.
@interface ExcelFont : SBObject <ExcelGenericMethods>

@property BOOL bold;  // Returns or sets if the font is formatted as bold.
@property (copy) NSArray<NSNumber *> *color;  // Returns or sets the color for the font.
@property ExcelXlBackground fontBackground;  // Returns or sets the text background type.
@property ExcelXlColorIndex fontColorIndex;  // Returns or sets the color index of the font.
@property NSInteger fontSize;  // Returns or sets the font size.
@property (copy) NSString *fontStyle;  // Returns or sets the font style.
@property BOOL italic;  // Returns or sets if the font is formatted as italic.
@property (copy) NSString *name;  // Returns or sets the font name associated with this font object.
@property BOOL outlineFont;  // Returns or sets if the specified font is formatted as outline.
@property BOOL shadow;  // Returns or sets if the specified font is formatted as shadowed.
@property BOOL strikethrough;  // Returns or sets if the font is formatted as strikethrough text.
@property BOOL subscript;  // Returns or sets if the font is formatted as subscript.
@property BOOL superscript;  // Returns or sets if the font is formatted as superscript.
@property ExcelXlThemeColor themeColor;  // Returns or sets the theme color in the applied color scheme that is associated with the specified object.
@property ExcelXlThemeFont themeFontIndex;  // Returns or sets the theme font in the applied font scheme that is associated with the specified object.
@property double tintAndShade;  // Returns or sets a Single that lightens or darkens a color.
@property ExcelXlUnderlineStyle underline;  // Returns or sets the type of underline applied to the font.


@end

// Represents a style description for a range.
@interface ExcelStyle : SBObject <ExcelGenericMethods>

@property BOOL addIndent;  // Returns or sets if text is automatically indented when the text alignment in a cell is set to equal distribution either horizontally or vertically. 
@property (readonly) BOOL builtIn;  // Returns true if the style is a built-in style.
@property (copy, readonly) ExcelFont *fontObject;  // Returns the font object associated with this style.
@property BOOL formulaHidden;  // Returns or sets if the formula will be hidden when the worksheet is protected.
@property ExcelXlHAlign horizontalAlignment;  // Returns or sets the horizontal alignment for the object.
@property BOOL includeAlignment;  // Returns or sets if the style includes the add indent, horizontal alignment, vertical alignment, wrap text, and orientation properties.
@property BOOL includeBorder;  // Returns or sets if the style includes the color, color index, line style, and weight border properties.
@property BOOL includeFont;  // Returns or sets if the style includes the background, bold, color, color index, font style, italic, name, outline font, shadow, size, strikethrough, subscript, superscript, and underline font properties.
@property BOOL includeNumber;  // Returns or sets if the style includes the number format property. 
@property BOOL includePatterns;  // Returns or sets if the style includes the color, color index, invert if negative, pattern, pattern color, and pattern color index interior properties.
@property BOOL includeProtection;  // Returns or sets if the style includes the formula hidden and locked protection properties.
@property NSInteger indentLevel;  // Returns or sets the indent level for the style. Can be an integer from 0 to 15.
@property (copy, readonly) ExcelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified style.
@property BOOL locked;  // Returns or sets if the range using this style is locked.
@property BOOL mergedCells;  // Returns true if the style contains merged cells.
@property (copy, readonly) NSString *name;  // Returns the name of the style.
@property (copy, readonly) NSString *nameLocal;  // Returns the name of the style, in the language of the user.
@property (copy) NSString *numberFormat;  // Returns or sets the format code for the object.
@property (copy) NSString *numberFormatLocal;  // Returns or sets the format code for the object as a string in the language of the user.
@property ExcelXlOrientation orientation;  // May also be a number value from -90 to 90 degrees.
@property ExcelXLDefaultSheetDir readingOrder;  // Returns or sets the reading order for the specified object.
@property BOOL shrinkToFit;  // Returns or sets if text automatically shrinks to fit in the available column width.
@property (copy, readonly) NSString *value;  // Return the name of the specified style.
@property ExcelXlVAlign verticalAlignment;  // Returns or sets the vertical alignment of the style.
@property BOOL wrapText;  // Returns or sets if Microsoft Excel wraps the text in the object.

- (ExcelBorder *) getBorderWhichBorder:(ExcelXlBordersIndex)whichBorder;  // Returns the specified border object.

@end

// Represents the text frame in a shape object.
@interface ExcelTextRange : SBObject <ExcelGenericMethods>

- (SBElementArray<ExcelTextRangeCharacter *> *) textRangeCharacters;
- (SBElementArray<ExcelWord *> *) words;
- (SBElementArray<ExcelTextRangeLine *> *) textRangeLines;
- (SBElementArray<ExcelSentence *> *) sentences;
- (SBElementArray<ExcelParagraph *> *) paragraphs;
- (SBElementArray<ExcelTextFlow *> *) textFlows;

@property (readonly) double boundHeight;  // Returns the height, in points, of the text bounding box for the specified text.
@property (readonly) double boundLeft;  // Returns the left coordinate, in points, of the text bounding box for the specified text.
@property (readonly) double boundTop;  // Returns the top coordinate, in points, of the text bounding box for the specified text.
@property (readonly) double boundWidth;  // Returns the width, in points, of the text bounding box for the specified text.
@property (copy) NSString *content;  // Returns or sets the text in a text range.
@property (copy, readonly) ExcelShapeFont *font;  // Returns the character formatting for the object.
@property (copy, readonly) ExcelParagraphFormat *paragraphFormat;  // Returns the paragraph formatting for the specified text.
@property (readonly) NSInteger textLength;  // Returns the length of a text range.
@property (readonly) NSInteger textStart;  // Returns the starting point of the specified text range.

- (void) addPeriodsTo;  // Adds period punctuation to the end of the text contained in text range object.
- (void) changeCaseTo:(ExcelMsoTextChangeCase)to;  // Changes the case of a text range object.
- (void) copyTextRange NS_RETURNS_NOT_RETAINED;  // Copies a text range object.
- (void) cutTextRange;  // Removes a portion or all of the text from a range of text.
- (NSArray<id> *) getRotatedTextBounds;  // Returns back a list containing the four point of the text bounds as follows  {x1, y1}, {x2, y2}, {x3, y3}, {x4, y4} }
- (void) insertTextTextRangeInsertWhere:(ExcelMsoTextRangeInsertPosition)insertWhere newText:(NSString *)newText;  // Adds new text to the text range object.
- (void) pasteTextRange;  // Pastes the contents of the Clipboard into the text range object.
- (void) removePeriodsFrom;  // Removes all period punctuation from the text in the text range object.

@end

@interface ExcelParagraph : ExcelTextRange


@end

@interface ExcelSentence : ExcelTextRange


@end

@interface ExcelTextFlow : ExcelTextRange


@end

@interface ExcelTextRangeCharacter : ExcelTextRange


@end

@interface ExcelTextRangeLine : ExcelTextRange


@end

@interface ExcelWord : ExcelTextRange


@end



/*
 * Table Suite
 */

// Represents a cell, a row, a column, a selection of cells containing one or more contiguous blocks of cells, or a 3-D range.
@interface ExcelRange : SBObject <ExcelGenericMethods>

- (SBElementArray<ExcelRange *> *) ranges;
- (SBElementArray<ExcelCell *> *) cells;
- (SBElementArray<ExcelRow *> *) rows;
- (SBElementArray<ExcelColumn *> *) columns;
- (SBElementArray<ExcelCharacter *> *) characters;
- (SBElementArray<ExcelFormatCondition *> *) formatConditions;
- (SBElementArray<ExcelHyperlink *> *) hyperlinks;

@property (copy, readonly) ExcelExcelComment *ExcelComment;  //  Returns a comment object that represents the comment associated with the cell in the upper-left corner of the range.
@property BOOL addIndent;  // Returns or sets if text is automatically indented when the text alignment in a cell is set to equal distribution either horizontally or vertically.
@property (copy, readonly) NSArray<SBObject *> *areas;  // Returns a list of  range objects  that represents all the ranges in a multiple-area selection.
@property double columnWidth;  // Returns or sets the width of all columns in the specified range.
@property (copy, readonly) id countLarge;  // Counts the largest value in a given range of values. Read-only Variant.
@property (copy, readonly) ExcelRange *currentArray;  // If the specified cell is part of an array, returns a range object that represents the entire array.
@property (copy, readonly) ExcelRange *currentRegion;  // Returns a range object that represents the current region. The current region is a range bounded by any combination of blank rows and blank columns.
@property (copy, readonly) ExcelRange *dependents;  // Returns a range object that represents the range containing all the dependents of a cell. This can be a multiple selection, a union of Range objects, if there's more than one dependent.
@property (copy, readonly) ExcelRange *directDependents;  // Returns a range object that represents the range containing all the direct dependents of a cell. This can be a multiple selection, a union of Range objects, if there's more than one dependent.
@property (copy, readonly) ExcelRange *directPrecedents;  // Returns a range object that represents the range containing all the direct precedents of a cell. This can be a multiple selection, a union of Range objects,  if there's more than one precedent.
@property (copy, readonly) ExcelDisplayFormat *displayFormat;  // Returns the display format associated with this range.
@property (copy, readonly) ExcelRange *entireColumn;  // Returns a range object that represents the entire column, or columns,  that contains the specified range.
@property (copy, readonly) ExcelRange *entireRow;  // Returns a range object that represents the entire row, or rows,  that contains the specified range.
@property (readonly) NSInteger firstColumnIndex;  // Returns the number of the first column in the first area in the specified range.
@property (readonly) NSInteger firstRowIndex;  // Returns the number of the first row of the first area in the range.
@property (copy, readonly) ExcelFont *fontObject;  // Returns the font object associated with this range.
@property (copy) id formula;  // Returns or sets the object's formula, in A1-style notation.
@property (copy) NSString *formulaArray;  // Returns or sets the array formula of a range.
@property BOOL formulaHidden;  // Returns or sets if the formula will be hidden when the workbook or worksheet is protected.
@property ExcelXlFormulaLabel formulaLabel;  // Returns or sets the formula label type for the specified range. 
@property (copy) NSString *formulaLocal;  // Returns or sets the formula for the object, using A1-style references in the language of the user.
@property (copy) NSString *formulaR1c1;  // Returns or sets the formula for the object, using R1C1-style notation.
@property (copy) NSString *formulaR1c1Local;  // Returns or sets the formula for the object, using R1C1-style notation in the language of the user.
@property (copy) id formula2;  // Returns or sets the object's formula, in A1-style notation.
@property (copy) NSString *formula2Local;  // Returns or sets the formula for the object, using A1-style references in the language of the user.
@property (copy) NSString *formula2R1c1;  // Returns or sets the formula for the object, using R1C1-style notation.
@property (copy) NSString *formula2R1c1Local;  // Returns or sets the formula for the object, using R1C1-style notation in the language of the user.
@property (readonly) BOOL hasArray;  // Returns true if the specified cell is part of an array formula.
@property (readonly) BOOL hasFormula;  // Returns true if all cells in the range contain formulas.
@property (readonly) double height;  // Returns the height of the range.
@property BOOL hidden;  // Returns or sets if the rows or columns are hidden. The specified range must span an entire column or row.
@property ExcelXlHAlign horizontalAlignment;  // Returns or sets the horizontal alignment for the range.
@property NSInteger indentLevel;  // Returns or sets the indent level for the range or style. Can be an integer from 0 to 15.
@property (copy, readonly) ExcelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property (readonly) double leftPosition;  // The distance from the left edge of column A to the left edge of the range. If the range is discontinuous, the first area is used. If the range is more than one column wide, the leftmost column in the range is used.
@property (readonly) NSInteger listHeaderRows;  // Returns the number of header rows for the specified range.
@property (copy, readonly) ExcelListObject *listObject;  // Returns the list object associated with this range.
@property (readonly) ExcelXlLocationInTable locationInTable;  // Returns an enumeration that describes the part of the pivot table that contains the upper-left corner of the specified range.
@property BOOL locked;  // Returns or sets  if the range is locked.
@property (copy, readonly) ExcelRange *mergeArea;  // Returns a range object that represents the merged range containing the specified cell. If the specified cell isn't in a merged range, this property returns the specified cell.
@property BOOL mergeCells;  // Returns or sets if the range contains merged cells. 
@property (copy) NSString *name;  // Returns or sets the name of the range.
@property (copy, readonly) NSString *namedItem;  // Returns the named item associated with this range.
@property (copy) NSString *numberFormat;  // Returns or sets the format code for the object.
@property (copy) NSString *numberFormatLocal;  // Returns or sets the format code for the object as a string in the language of the user.
@property NSInteger outlineLevel;  // Returns or sets the current outline level of the specified row or column
@property ExcelXlPageBreak pageBreak;  // Returns or sets the location of a page break.
@property (copy, readonly) ExcelPhonetic *phoneticObject;  // Returns the phonetic object, which contains information about a specific phonetic text string in a specified range.
@property (copy, readonly) ExcelPivotField *pivotField;  // Returns a pivot field object that represents the pivot field containing the upper-left corner of the specified range.
@property (copy, readonly) ExcelPivotItem *pivotItem;  // Returns a pivot item object that represents the pivot item containing the upper-left corner of the specified range.
@property (copy, readonly) ExcelPivotTable *pivotTable;  // Returns a pivot table object that represents the pivot table containing the upper-left corner of the specified range.
@property (copy, readonly) ExcelRange *precedents;  // Returns a range object that represents all the precedents of a cell. This can be a multiple selection, a union of Range objects, if there's more than one precedent.
@property (copy, readonly) NSString *prefixCharacter;  // Returns the prefix character for the cell.
@property (copy, readonly) ExcelQueryTable *queryTable;  // Returns a query table object that represents the query table that intersects the specified range.
@property ExcelXLDefaultSheetDir readingOrder;  // Returns or sets the reading order for the specified object.
@property double rowHeight;  // Returns  or sets the height of all the rows in the range specified, measured in points.
@property BOOL showDetail;  // Returns or sets if the outline is expanded for the specified range, so that the detail of the column or row is visible. The specified range must be a single summary column or row in an outline.
@property BOOL shrinkToFit;  // Returns or sets if text automatically shrinks to fit in the available column width.
@property (copy, readonly) id stringValue;  // Returns the text for the range.
@property (copy) ExcelStyle *styleObject;  // Returns or sets a style object that represents the style of the specified range.
@property (readonly) BOOL summary;  // Returns true if the range is an outlining summary row or column. The range should be a row or a column.
@property ExcelXlOrientation textOrientation;  // The text orientation. Can be a number value from -90 to 90 degrees.
@property (readonly) double top;  // The distance from the top edge of row 1 to the top edge of the range. If the range is discontinuous, the first area is used. If the range is more than one row high, the top, lowest numbered, row in the range is used.
@property BOOL useStandardHeight;  // Returns or sets if the row height of the range object equals the standard height of the sheet.
@property BOOL useStandardWidth;  // Returns or sets if the column width of the range object equals the standard width of the sheet.
@property (copy, readonly) ExcelValidation *validation;  // Returns the validation object that represents data validation for the specified range.
@property (copy) id value;  // Returns or sets the value of the range.
@property (copy) id value2;  // Returns or sets the value of the range. The difference between this property and Value is that Value2 uses the numerical representation of cells that are formatted as currency and date.
@property ExcelXlVerticalAlignmentTarget verticalAlignment;  // Returns or sets the vertical alignment of the range.
@property (readonly) double width;  // Returns the width of the range.
@property (copy, readonly) ExcelSheet *worksheetObject;  // Returns a worksheet object that represents the worksheet containing the specified range.
@property BOOL wrapText;  // Returns or sets if Microsoft Excel wraps the text in the object.

- (void) activateObject;  // Activates the object.
- (ExcelExcelComment *) addCommentCommentText:(NSString *)commentText;  // Adds a comment to the range.
- (void) advancedFilterAction:(ExcelXlFilterAction)action criteriaRange:(ExcelXlCategoryNames)criteriaRange copyToRange:(ExcelXlCategoryNames)copyToRange unique:(BOOL)unique;  // Filters or copies data from a list based on a criteria range. If the initial selection is a single cell, that cell's current region is used.
- (void) applyNamesNames:(NSArray<NSString *> *)names ignoreRelativeAbsolute:(BOOL)ignoreRelativeAbsolute useRowColumnNames:(BOOL)useRowColumnNames omitColumn:(BOOL)omitColumn omitRow:(BOOL)omitRow order:(ExcelXlApplyNamesOrder)order appendLast:(BOOL)appendLast;  // Applies names to the cells in the specified range.
- (void) applyOutlineStyles;  // Applies outlining styles to the specified range.
- (void) autoOutline;  // Automatically creates an outline for the specified range. If the range is a single cell, Microsoft Excel creates an outline for the entire sheet. The new outline replaces any existing outline.
- (NSString *) autocompleteString:(NSString *)string;  // Returns an AutoComplete match from the list. If there's no AutoComplete match or if more than one entry in the list matches the string to complete, this method returns an empty string.
- (void) autofillDestination:(ExcelXlCategoryNames)destination type:(ExcelXlAutoFillType)type;  // Performs an autofill on the cells in the specified range.
- (void) autofilterRangeField:(NSInteger)field criteria1:(id)criteria1 operator:(ExcelXlAutoFilterOperator)operator_ criteria2:(id)criteria2 visibleDropDown:(BOOL)visibleDropDown;  // Filters a list using the AutoFilter.
- (void) autofit;  // Changes the width of the specified column range or the height of the specified row range to achieve the best fit.
- (void) autoformatFormat:(ExcelXlRangeAutoFormat)format includeNumber:(BOOL)includeNumber font:(BOOL)font alignment:(BOOL)alignment border:(BOOL)border pattern:(BOOL)pattern width:(BOOL)width;  // Automatically formats a range of cells, using a predefined format.
- (void) borderAroundLineStyle:(ExcelXlLineStyle)lineStyle weight:(ExcelXlBorderWeight)weight colorIndex:(ExcelXlColorIndex)colorIndex color:(NSArray<NSNumber *> *)color;  // Adds a border to a range and sets the color, line style, and weight properties for the new border.
- (void) calculate;  // Calculates all open workbooks, a specific worksheet in a workbook, or a specified range of cells on a worksheet.
- (void) calculateRowMajorOrder;  // Calculates a specfied range of cells.
- (void) checkSpellingCustomDictionary:(NSString *)customDictionary ignoreUppercase:(BOOL)ignoreUppercase alwaysSuggest:(BOOL)alwaysSuggest;  // Checks the spelling of an object.
- (void) clearExcelComments;  // Clears all cell comments from the specified range.
- (void) clearContents;  // Clears all cell comments from the specified range.
- (void) clearOutline;  // Clears the outline for the specified range.
- (void) clearRange;  // Clears the entire object.
- (void) clearRangeFormats;  // Clears the formatting of the object.
- (ExcelRange *) columnDifferencesComparison:(ExcelXlCategoryNames)comparison;  // Returns a range object that represents all the cells whose contents are different from the comparison cell in each column.
- (void) consolidateSources:(NSArray<NSString *> *)sources consolidationFunction:(ExcelXlConsolidationFunction)consolidationFunction topRow:(BOOL)topRow leftColumn:(BOOL)leftColumn createLinks:(BOOL)createLinks;  // Consolidates data from multiple ranges on multiple worksheets into a single range on a single worksheet.
- (void) copyPictureAppearance:(ExcelXlPictureAppearance)appearance format:(ExcelXlCopyPictureFormat)format NS_RETURNS_NOT_RETAINED;  // Copies the selected object to the clipboard as a picture.
- (void) copyRangeDestination:(ExcelRange *)destination NS_RETURNS_NOT_RETAINED;  //  Copies the range to the specified range.
- (void) createNamesTop:(BOOL)top leftPosition:(BOOL)leftPosition bottom:(BOOL)bottom right:(BOOL)right;  // Creates names in the specified range, based on text labels in the sheet.
- (void) cutRangeDestinationOfCut:(ExcelXlCategoryNames)destinationOfCut;  // Cuts the range to the clipboard.
- (void) dataSeriesRowcol:(ExcelXlRowCol)rowcol dataSeriesType:(ExcelXlDataSeriesType)dataSeriesType date:(ExcelXlDataSeriesDate)date increment:(NSInteger)increment stop:(NSInteger)stop trend:(BOOL)trend;  // Creates a data series in the specified range.
- (void) dataTableRowInput:(ExcelCell *)rowInput columnInput:(ExcelCell *)columnInput;  // Creates a data table based on input values and formulas that you define on a worksheet.
- (void) deleteRangeShift:(ExcelXlDeleteShiftDirection)shift;  // Deletes the range
- (void) fillDown;  // Fills down from the top cell or cells in the specified range to the bottom of the range. The contents and formatting of the cell or cells in the top row of a range are copied into the rest of the rows in the range.
- (void) fillLeft;  // Fills left from the rightmost cell or cells in the specified range. The contents and formatting of the cell or cells in the rightmost column of a range are copied into the rest of the columns in the range.
- (void) fillRight;  // Fills right from the leftmost cell or cells in the specified range. The contents and formatting of the cell or cells in the leftmost column of a range are copied into the rest of the columns in the range.
- (void) fillUp;  // Fills up from the bottom cell or cells in the specified range to the top of the range. The contents and formatting of the cell or cells in the bottom row of a range are copied into the rest of the rows in the range.
- (ExcelRange *) findWhat:(NSString *)what after:(ExcelXlCategoryNames)after lookIn:(ExcelXlFindLookIn)lookIn lookAt:(ExcelXlLookAt)lookAt searchOrder:(ExcelXlSearchOrder)searchOrder searchDirection:(ExcelXlSearchDirection)searchDirection matchCase:(BOOL)matchCase matchByte:(BOOL)matchByte;  // Finds specific information in a range, and returns a range object that represents the first cell where that information is found. Returns nothing if no match is found. Doesn't affect the selection or the active cell.
- (ExcelRange *) findNextAfter:(ExcelXlCategoryNames)after;  // Continues a search that was begun with the find method. Finds the next cell that matches those same conditions and returns a range object that represents that cell. Doesn't affect the selection or the active cell.
- (ExcelRange *) findPreviousAfter:(ExcelXlCategoryNames)after;  // Continues a search that was begun with the find method. Finds the previous cell that matches those same conditions and returns a range object that represents that cell. Doesn't affect the selection or the active cell.
- (void) functionWizard;  // Starts the Function Wizard for the upper-left cell of the range.
- (id) getXMLValue;  // Returns the value of the specified range as XML.
- (NSString *) getAddressRowAbsolute:(BOOL)rowAbsolute columnAbsolute:(BOOL)columnAbsolute referenceStyle:(ExcelXlReferenceStyle)referenceStyle external:(BOOL)external relativeTo:(ExcelXlCategoryNames)relativeTo;  // Returns the range reference in the language of the macro.
- (NSString *) getAddressLocalRowAbsolute:(BOOL)rowAbsolute columnAbsolute:(BOOL)columnAbsolute referenceStyle:(ExcelXlReferenceStyle)referenceStyle external:(BOOL)external relativeTo:(ExcelXlCategoryNames)relativeTo;  // Returns the range reference in the language of the user.
- (ExcelBorder *) getBorderWhichBorder:(ExcelXlBordersIndex)whichBorder;  // Returns the specified border object.
- (ExcelRange *) getEndDirection:(ExcelXlDirection)direction;  // Returns a range object that represents the cell at the end of the region that contains the source range.
- (ExcelRange *) getOffsetRowOffset:(NSInteger)rowOffset columnOffset:(NSInteger)columnOffset;  // Returns a range object that represents a range that's offset from the specified range.
- (ExcelRange *) getResizeRowSize:(NSInteger)rowSize columnSize:(NSInteger)columnSize;  // Resizes the specified range. Returns a Range object that represents the resized range.
- (BOOL) goalSeekGoal:(double)goal changingCell:(ExcelRange *)changingCell;  // Calculates the values necessary to achieve a specific goal. If the goal is an amount returned by a formula, this calculates a value that, when supplied to your formula, causes the formula to return the number you want. Returns True if successful.
- (BOOL) groupStart:(NSInteger)start end:(NSInteger)end by:(NSInteger)by periods:(NSArray<NSNumber *> *)periods;  // When the Range object represents a single cell in a pivot table field's data range, the group method performs numeric or date-based grouping in that field.
- (void) insertIndentInsertAmount:(NSInteger)insertAmount;  // Adds an indent to the specified range.
- (void) insertIntoRangeShift:(ExcelXlInsertShiftDirection)shift;  // Inserts a cell or a range of cells into the worksheet or macro sheet and shifts other cells away to make space.
- (void) justify;  // Rearranges the text in a range so that it fills the range evenly.
- (void) listNames;  // Pastes a list of all non-hidden names onto the worksheet, beginning with the first cell in the range.
- (void) mergeAcross:(BOOL)across;  // Creates a merged cell from the specified Range object.
- (void) navigateArrowTowardPrecedent:(BOOL)towardPrecedent arrowNumber:(NSInteger)arrowNumber linkNumber:(NSInteger)linkNumber;  // Navigates a tracer arrow for the specified range to the precedent, dependent, or error-causing cell or cells. Selects the precedent, dependent, or error cells and returns a range object that represents the new selection.
- (void) parseParseLine:(NSString *)parseLine destination:(ExcelXlCategoryNames)destination;  // Parses a range of data and breaks it into multiple cells. Distributes the contents of the range to fill several adjacent columns; the range can be no more than one column wide.
- (void) pasteSpecialWhat:(ExcelXlPasteType)what operation:(ExcelXlPasteSpecialOperation)operation skipBlanks:(BOOL)skipBlanks transpose:(BOOL)transpose;  // Pastes the contents of the Clipboard onto the sheet, using a specified format. Use this method to paste data from other applications or to paste data in a specific format.
- (void) printOutFrom:(NSInteger)from to:(NSInteger)to copies:(NSInteger)copies preview:(BOOL)preview activePrinter:(NSString *)activePrinter printToFile:(BOOL)printToFile collate:(BOOL)collate;  // Prints the object
- (void) printPreviewEnableChanges:(BOOL)enableChanges;  // Shows a preview of the object as it would look when printed. This function has been deprecated.
- (void) removeSubtotal;  // Removes subtotals from a list.
- (BOOL) replaceWhat:(NSString *)what replacement:(NSString *)replacement lookAt:(ExcelXlLookAt)lookAt searchOrder:(ExcelXlSearchOrder)searchOrder matchCase:(BOOL)matchCase matchByte:(BOOL)matchByte matchControlCharacters:(BOOL)matchControlCharacters matchDiacritics:(BOOL)matchDiacritics formulaVersion:(ExcelXlFormulaVersion)formulaVersion;  // Finds and replaces characters in cells within a range. Doesn't change the selection or the active cell.
- (ExcelRange *) rowDifferencesComparison:(ExcelCell *)comparison;  // Returns a range object that represents all the cells whose contents are different from those of the comparison cell in each row.
- (NSString *) runVBMacroArg1:(id)arg1 arg2:(id)arg2 arg3:(id)arg3 arg4:(id)arg4 arg5:(id)arg5 arg6:(id)arg6 arg7:(id)arg7 arg8:(id)arg8 arg9:(id)arg9 arg10:(id)arg10 arg11:(id)arg11 arg12:(id)arg12 arg13:(id)arg13 arg14:(id)arg14 arg15:(id)arg15 arg16:(id)arg16 arg17:(id)arg17 arg18:(id)arg18 arg19:(id)arg19 arg20:(id)arg20 arg21:(id)arg21 arg22:(id)arg22 arg23:(id)arg23 arg24:(id)arg24 arg25:(id)arg25 arg26:(id)arg26 arg27:(id)arg27 arg28:(id)arg28 arg29:(id)arg29 arg30:(id)arg30;  //  Runs a macro or calls a function. This can be used to run a macro written in the Microsoft Excel 4.0 macro language, or to run a function in a DLL or XLL.
- (NSString *) runXLMMacroArg1:(id)arg1 arg2:(id)arg2 arg3:(id)arg3 arg4:(id)arg4 arg5:(id)arg5 arg6:(id)arg6 arg7:(id)arg7 arg8:(id)arg8 arg9:(id)arg9 arg10:(id)arg10 arg11:(id)arg11 arg12:(id)arg12 arg13:(id)arg13 arg14:(id)arg14 arg15:(id)arg15 arg16:(id)arg16 arg17:(id)arg17 arg18:(id)arg18 arg19:(id)arg19 arg20:(id)arg20 arg21:(id)arg21 arg22:(id)arg22 arg23:(id)arg23 arg24:(id)arg24 arg25:(id)arg25 arg26:(id)arg26 arg27:(id)arg27 arg28:(id)arg28 arg29:(id)arg29 arg30:(id)arg30;  //  Runs a macro or calls a function. This can be used to run a macro written in the Microsoft Excel 4.0 macro language, or to run a function in a DLL or XLL.
- (void) setXMLValueRangeValue:(id)rangeValue;  // Set the value of the specified range using XML.
- (void) show;  // Scrolls through the contents of the active window to move the range into view. The range must consist of a single cell in the active document.
- (void) showDependentsRemove:(BOOL)remove;  // Draws tracer arrows to the direct dependents of the range.
- (void) showErrors;  // Draws tracer arrows through the precedents tree to the cell that's the source of the error, and returns the range that contains that cell.
- (void) showPrecedentsRemove:(BOOL)remove;  // Draws tracer arrows to the direct precedents of the range.
- (void) sortKey1:(ExcelRange *)key1 order1:(ExcelXlSortOrder)order1 key2:(ExcelRange *)key2 sortType:(ExcelXlSortType)sortType order2:(ExcelXlSortOrder)order2 key3:(ExcelRange *)key3 order3:(ExcelXlSortOrder)order3 header:(ExcelXlYesNoGuess)header orderCustom:(NSInteger)orderCustom matchCase:(BOOL)matchCase orientation:(ExcelXlSortOrientation)orientation sortMethod:(ExcelXlSortMethod)sortMethod ignoreControlCharacters:(BOOL)ignoreControlCharacters ignoreDiacritics:(BOOL)ignoreDiacritics dataoption1:(ExcelXlSortDataOption)dataoption1 dataoption2:(ExcelXlSortDataOption)dataoption2 dataoption3:(ExcelXlSortDataOption)dataoption3;  // Sorts a pivot table, a range, or the current region, if the specified range contains only one cell.
- (void) sortSpecialSortMethod:(ExcelXlSortMethod)sortMethod key1:(ExcelRange *)key1 order1:(ExcelXlSortOrder)order1 type:(id)type key2:(ExcelRange *)key2 order2:(ExcelXlSortOrder)order2 key3:(ExcelRange *)key3 order3:(ExcelXlSortOrder)order3 header:(ExcelXlYesNoGuess)header orderCustom:(NSInteger)orderCustom matchCase:(BOOL)matchCase orientation:(ExcelXlSortOrientation)orientation dataoption1:(ExcelXlSortDataOption)dataoption1 dataoption2:(ExcelXlSortDataOption)dataoption2 dataoption3:(ExcelXlSortDataOption)dataoption3;  // Uses East Asian sorting methods to sort the range, or uses the current region if the range contains only one cell.
- (ExcelRange *) specialCellsType:(ExcelXlCellType)type value:(ExcelXlSpecialCellsValue)value;  // Returns a Range object that represents all the cells that match the specified type and value.
- (void) subtotalGroupBy:(NSInteger)groupBy function:(ExcelXlConsolidationFunction)function totalList:(NSArray<NSNumber *> *)totalList replace:(BOOL)replace pageBreaks:(BOOL)pageBreaks summaryBelowData:(ExcelXlSummaryRow)summaryBelowData;  // Creates subtotals for the range, or the current region, if the range is a single cell.
- (void) textToColumnsDestination:(ExcelXlCategoryNames)destination dataType:(ExcelXlTextParsingType)dataType textQualifier:(ExcelXlTextQualifier)textQualifier consecutiveDelimiter:(BOOL)consecutiveDelimiter tab:(BOOL)tab semicolon:(BOOL)semicolon comma:(BOOL)comma space:(BOOL)space useOther:(BOOL)useOther otherChar:(NSString *)otherChar fieldInfo:(NSArray<id> *)fieldInfo decimalSeparator:(NSString *)decimalSeparator thousandsSeparator:(NSString *)thousandsSeparator;  // Parses a column of cells that contain text into several columns.
- (void) ungroup;  // Promotes a range in an outline, that is, decreases its outline level. The specified range must be a row or column, or a range of rows or columns. If the range is in a pivot table, this method ungroups the items contained in the range.
- (void) unmerge;  // Separates a merged area into individual cells.

@end

@interface ExcelCell : ExcelRange


@end

@interface ExcelColumn : ExcelRange


@end

@interface ExcelRow : ExcelRange


@end



/*
 * Proofing Suite
 */

// Contains Microsoft Excel autocorrect attributes, capitalization of names of days, correction of two initial capital letters, automatic correction list, and so on.
@interface ExcelAutocorrect : SBObject <ExcelGenericMethods>

@property BOOL AutomaticallyExpandTablesAsIType;  // Returns or sets if resizes the table to include data entered into a neighboring column or row.
@property BOOL AutomaticallyFillFormulas;  // Returns or sets if we automatically copies the formula to all the cells in the column when a formula is entered into a cell.
@property BOOL correctCapsLock;  // Returns or sets if Microsoft Excel automatically corrects accidental use of the CAPS LOCK key.
@property BOOL correctDays;  // Returns or sets if the first letter of day names is capitalized automatically.
@property BOOL correctInitialCaps;  // Returns or sets if words that begin with two capital letters are corrected automatically.
@property BOOL correctSentenceCaps;  // Returns or sets if Microsoft Excel automatically corrects sentence, first word, capitalization.
@property BOOL replaceText;  // Returns or sets if text in the list of autocorrect replacements is replaced automatically.

- (void) addReplacementTextToReplace:(NSString *)textToReplace replacementText:(NSString *)replacementText;  // Adds an entry to the array of autocorrect replacements.
- (void) deleteReplacementWhat:(NSString *)what;  // Deletes an entry from the array of autocorrect replacements. 
- (NSArray<id> *) getReplacementList;  // Returns a list of autocorrect replacements.

@end



/*
 * Chart Suite
 */

// Represents a chart axis title.
@interface ExcelAxisTitle : SBObject <ExcelGenericMethods>

- (SBElementArray<ExcelCharacter *> *) characters;

@property BOOL autoScaleFont;  // Returns or sets if the text in the object changes font size when the object size changes.
@property (copy) NSString *axisTitleText;  // Returns or sets the text associated with this object.
@property (copy, readonly) ExcelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy) NSString *caption;  // Returns or sets the caption for this object.
@property (copy, readonly) ExcelChartFillFormat *chartFillFormatObject;  // Returns a chart fill format object for this object.
@property (copy, readonly) ExcelFont *fontObject;  // Returns a font object that represents the font of the specified object.
@property ExcelXlHAlign horizontalAlignment;  // Returns or sets the horizontal alignment for the object.
@property BOOL includeInLayout;  // Returns or sets if a axis title will occupy the chart layout space when a chart layout is being determined.
@property (copy, readonly) ExcelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property double leftPosition;  // Returns or sets the left position of the specified object, in points.
@property (copy, readonly) NSString *name;  // Returns the name of the object.
@property NSInteger orientation;  // The text orientation.  Can be a number value from -90 to 90 degrees.
@property ExcelXlChartElementPosition position;  // Returns or sets the position of the axis title on the chart.
@property ExcelXLDefaultSheetDir readingOrder;  // Returns or sets the reading order for the specified object.
@property BOOL shadow;  // Returns or sets if the object has a shadow.
@property double top;  // Returns or sets the top position of the specified object, in points.
@property ExcelXlVerticalAlignmentTarget verticalAlignment;  // Returns or sets the vertical alignment of the object.


@end

// Represents a single axis in a chart.
@interface ExcelAxis : SBObject <ExcelGenericMethods>

@property BOOL axisBetweenCategories;  // Returns or sets if the value axis crosses the category axis between categories.
@property (readonly) ExcelXlAxisGroup axisGroup;  // Returns the group for the specified axis.
@property (copy, readonly) ExcelAxisTitle *axisTitle;  // Returns an axis title object that represents the title of the specified axis.
@property ExcelXlAxisType axisType;  // Returns or sets the axis type
@property ExcelXlTimeUnit baseUnit;  // Returns or sets the base unit for the specified category axis.
@property BOOL baseUnitIsAuto;  // Returns or sets if Microsoft Excel chooses appropriate base units for the specified category axis.
@property (copy, readonly) ExcelBorder *border;  // Returns a border object that represents the border of the object.
@property ExcelXlCategoryNames categoryNames;  // Returns or sets all the category names for the specified axis, as a text array. When you set this property, you can set it to either an array or a range object that contains the category names.
@property ExcelXlCategoryType categoryType;  // Returns or sets the category axis type.
@property (copy, readonly) ExcelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the axis.
@property ExcelXlAxisCrosses crosses;  // Returns or sets the point on the specified axis where the other axis crosses.
@property double crossesAt;  // Returns or sets the point on the value axis where the category axis crosses it. Applies only to the value axis.
@property ExcelXlDisplayUnit displayUnit;  // Returns or sets the unit label for the specified axis. 
@property double displayUnitCustom;  // If the value of the display unit property is custom display unit, the display unit custom property returns or sets the value of the displayed units.
@property (copy, readonly) ExcelDisplayUnitLabel *displayUnitLabel;  // Returns the display unit label object for the specified axis
@property BOOL hasDisplayUnitLabel;  // Returns or sets if the label specified by the display unit or display unit custom property is displayed on the specified axis. The default value is true.
@property BOOL hasMajorGridlines;  // Returns or sets if the axis has major gridlines. Only axes in the primary axis group can have gridlines.
@property BOOL hasMinorGridlines;  // Returns or sets if the axis has minor gridlines. Only axes in the primary axis group can have gridlines.
@property BOOL hasTitle;  // Returns or sets if the axis or chart has a visible title.
@property (readonly) double height;  // Returns the height of the object.
@property (readonly) double leftPosition;  // Returns the left position of the specified object, in points.
@property double logBase;  // Returns or sets the base of the logarithm when you are using log scales.
@property (copy, readonly) ExcelGridlines *majorGridlines;  // Returns a gridlines object that represents the major gridlines for the specified axis. Only axes in the primary axis group can have gridlines.
@property ExcelXlTickMark majorTickMark;  // Returns or sets the type of major tick mark for the specified axis. 
@property double majorUnit;  // Returns or sets the major units for the axis.
@property BOOL majorUnitIsAuto;  // Returns or sets if Microsoft Excel calculates the major units for the axis.
@property ExcelXlTimeUnit majorUnitScale;  // Returns or sets the major unit scale value for the category axis when the category type property is set to time scale.
@property double maximumScale;  // Returns or sets the maximum value on the axis.
@property BOOL maximumScaleIsAuto;  // Returns or sets if Microsoft Excel calculates the maximum value for the axis.
@property double minimumScale;  // Returns or sets the minimum value on the axis.
@property BOOL minimumScaleIsAuto;  // Returns or sets if Microsoft Excel calculates the minimum value for the axis.
@property (copy, readonly) ExcelGridlines *minorGridlines;  // Returns a gridlines object that represents the minor gridlines for the specified axis. Only axes in the primary axis group can have gridlines.
@property ExcelXlTickMark minorTickMark;  // Returns or sets the type of minor tick mark for the specified axis.
@property double minorUnit;  // Returns or sets the minor units on the axis.
@property BOOL minorUnitIsAuto;  // Returns or sets if Microsoft Excel calculates minor units for the axis.
@property ExcelXlTimeUnit minorUnitScale;  // Returns or sets the minor unit scale value for the category axis when the category type property is set to time scale. 
@property BOOL reversePlotOrder;  // Returns or sets if Microsoft Excel plots data points from last to first.
@property ExcelXlScaleType scaleType;  // Returns or sets the value axis scale type.
@property ExcelXlTickLabelPosition tickLabelPosition;  // Describes the position of tick-mark labels on the specified axis.
@property NSInteger tickLabelSpacing;  // Returns or sets the number of categories or series between tick-mark labels. Applies only to category and series axes.
@property BOOL tickLabelSpacingIsAuto;  // Returns or sets whether or not the tick label spacing is automatic.
@property (copy, readonly) ExcelTickLabels *tickLabels;  // Returns a tick labels object that represents the tick-mark labels for the specified axis.
@property NSInteger tickMarkSpacing;  // Returns or sets the number of categories or series between tick marks. Applies only to category and series axes.
@property (readonly) double top;  // Returns the top position of the specified object, in points.
@property (readonly) double width;  // Returns the width of the object.


@end

// Represents the chart area of a chart. The chart area on a 2-D chart contains the axes, the chart title, the axis titles, and the legend. The chart area on a 3-D chart contains the chart title and the legend; it doesn't include the plot area .
@interface ExcelChartArea : SBObject <ExcelGenericMethods>

@property BOOL autoScaleFont;  // Returns or sets if the text in the object changes font size when the object size changes.
@property (copy, readonly) ExcelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy, readonly) ExcelChartFillFormat *chartFillFormatObject;  // Returns a chart fill format object for this object.
@property (copy, readonly) ExcelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the chart area.
@property (copy, readonly) ExcelFont *fontObject;  // Returns a font object that represents the font of the specified object.
@property double height;  // Returns or sets the height of the object.
@property (copy, readonly) ExcelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property double leftPosition;  // Returns or sets the left position of the specified object, in points.
@property (copy, readonly) NSString *name;  // Returns the name of the object.
@property (readonly) BOOL roundedCorners;  // Returns or sets if the chart area has rounded corners.
@property BOOL shadow;  // Returns or sets if the object has a shadow.
@property double top;  // Returns or sets the top position of the specified object, in points.
@property double width;  // Returns or sets  the width of the object.

- (void) clearContents;  // Clears the data from a chart but leaves the formatting. 

@end

// Used only with charts. Represents fill formatting for chart elements.
@interface ExcelChartFillFormat : SBObject <ExcelGenericMethods>

@property (copy, readonly) NSArray<NSNumber *> *backColor;  // Returns the background color.
@property NSInteger backgroundSchemeColor;  // Returns or sets the background color as an index in the current color scheme.
@property (readonly) ExcelMsoFillType chartFillFormatType;  // The fill type. 
@property (copy, readonly) NSArray<NSNumber *> *foreColor;  // Returns the foreground color.
@property NSInteger foregroundSchemeColor;  // Returns or sets the foreground color as an index in the current color scheme.
@property (readonly) ExcelMsoGradientColorType gradientColorType;  // Returns the gradient color type for the specified fill.
@property (readonly) double gradientDegree;  // Returns the gradient degree of the specified one-color shaded fill as a floating-point value from 0.0 dark through 1.0 light.
@property (readonly) ExcelMsoGradientStyle gradientStyle;  // Returns the gradient style for the specified fill.
@property (readonly) NSInteger gradientVariant;  // Returns the shade variant for the specified fill as an integer value from 1 through 4.
@property (readonly) ExcelMsoPatternType pattern;  // Returns the fill pattern.
@property (readonly) ExcelMsoPresetGradientType presetGradientType;  // Returns the preset gradient type for the specified fill.
@property (readonly) ExcelMsoPresetTexture presetTexture;  // Returns the preset texture for the specified fill.
@property (copy, readonly) NSString *textureName;  // Returns the name of the custom texture file for the specified fill.
@property (readonly) ExcelMsoTextureType textureType;  // Returns the texture type for the specified fill.
@property BOOL visible;  // Returns or sets if the specified object is visible.


@end

// Provides access to the Office Art formatting for chart elements.
@interface ExcelChartFormat : SBObject <ExcelGenericMethods>

@property (copy, readonly) ExcelFillFormat *fillFormat;  // Returns a fill format object for the parent chart element that contains fill formatting properties for the chart element.
@property (copy, readonly) ExcelGlowFormat *glowFormat;  // Returns a glow format object for a specified chart that contains glow formatting properties for the chart.
@property (copy, readonly) ExcelLineFormat *lineFormat;  // Returns a line format object that contains line formatting properties for the specified chart element.
@property (copy, readonly) ExcelShadowFormat *shadowFormat;  // Returns a shadow format object that contains shadow formatting properties for the chart element.
@property (copy, readonly) ExcelShapeTextFrame *shapeTextFrame;  // Returns a shape text frame object that contains the alignment and anchoring properties for the specified chart.
@property (copy, readonly) ExcelSoftEdgeFormat *softEdgeFormat;  // Returns the formatting properties for a soft edge effect.
@property (copy, readonly) ExcelThreeDFormat *threeDFormat;  // Returns a threeD format object that contains 3-D-effect formatting properties for the specified chart.


@end

// Represents one or more series plotted in a chart with the same format. A chart contains one or more chart groups, each chart group contains one or more series, and each series contains one or more points.
@interface ExcelChartGroup : SBObject <ExcelGenericMethods>

- (SBElementArray<ExcelSeries *> *) seriesCollection;

@property ExcelXlAxisGroup axisGroup;
@property NSInteger bubbleScale;  // Returns or sets the scale factor for bubbles in the specified chart group. Can be an integer value from zero to 300, corresponding to a percentage of the default size. Applies only to bubble charts.
@property NSInteger doughnutHoleSize;  // Returns or sets the size of the hole in a doughnut chart group. The hole size is expressed as a percentage of the chart size, between 10 and 90 percent.
@property (copy, readonly) ExcelDownBars *downBarsObject;  // Returns a down bars object that represents the down bars on a line chart. Applies only to line charts.
@property (copy, readonly) ExcelDropLines *dropLinesObject;  // Returns a drop lines object that represents the drop lines for a series on a line chart or area chart. Applies only to line charts or area charts.
@property (readonly) NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property NSInteger firstSliceAngle;  // Returns or sets the angle of the first pie-chart or doughnut-chart slice, in degrees, clockwise from vertical. Applies only to pie, 3-D pie, and doughnut charts.
@property NSInteger gapWidth;  // Returns or sets: For bar and column charts, the space between bar or column clusters, as a percentage of the bar or column width. For pie of pie and bar of pie charts, the space between the primary and secondary sections of the chart.
@property BOOL hasDropLines;  // Returns or sets if the line chart or area chart has drop lines. Applies only to line and area charts.
@property BOOL hasHiLoLines;  // Returns or sets if the line chart has high-low lines. Applies only to line charts.
@property BOOL hasRadarAxisLabels;  // Returns or sets if a radar chart has axis labels. Applies only to radar charts.
@property BOOL hasSeriesLines;  // Returns or sets if a stacked column chart or bar chart has series lines or if a pie of pie chart or bar of pie chart has connector lines between the two sections. Applies only to stacked column charts, bar charts, pie of pie charts, or bar of pie charts.
@property BOOL hasThreeDShading;  // Returns or sets if the chart group has three-dimensional shading.
@property BOOL hasUpDownBars;  // Returns or sets if a line chart has up and down bars. Applies only to line charts.
@property (copy, readonly) ExcelHiloLines *hiloLinesObject;  // Returns a hilo lines object that represents the high-low lines for a series on a line chart.
@property NSInteger overlap;  // Returns or sets how bars and columns are positioned. Can be a value between -100 and 100. Applies only to 2-D bar and 2-D column charts.
@property (copy, readonly) ExcelTickLabels *radarAxisLabels;  // Returns a tick labels object that represents the radar axis labels for the specified chart group.
@property NSInteger secondPlotSize;  // Returns or sets the size of the secondary section of either a pie of pie chart or a bar of pie chart, as a percentage of the size of the primary pie. Can be a value from 5 to 200. 
@property (copy, readonly) ExcelSeriesLines *seriesLinesObject;  // Returns a series lines object that represents the series lines for a stacked bar chart or a stacked column chart. Applies only to stacked bar and stacked column charts.
@property BOOL showNegativeBubbles;  // Returns or sets if negative bubbles are shown for the chart group. Valid only for bubble charts.
@property ExcelXlSizeRepresents sizeRepresents;  // Returns or sets what the bubble size represents on a bubble chart.
@property ExcelXlChartSplitType splitType;  // Returns or sets the way the two sections of either a pie of pie chart or a bar of pie chart are split.
@property NSInteger splitValue;  // Returns or sets the threshold value separating the two sections of either a pie of pie chart or a bar of pie chart.
@property (copy, readonly) ExcelUpBars *upBarsObject;  // Returns an up bars object that represents the up bars on a line chart. Applies only to line charts.
@property BOOL varyByCategories;  // Returns or sets if Microsoft Excel assigns a different color or pattern to each data marker. The chart must contain only one series.


@end

@interface ExcelAreaGroup : ExcelChartGroup


@end

@interface ExcelBarGroup : ExcelChartGroup


@end

// Represents an embedded chart on a worksheet. The chart object object acts as a container for a chart object. Properties and methods for the chart object object control the appearance and size of the embedded chart on the worksheet.
@interface ExcelChartObject : SBObject <ExcelGenericMethods>

@property (copy, readonly) ExcelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy, readonly) ExcelRange *bottomRightCell;  // Returns a range object that represents the cell that lies under the lower-right corner of the object.
@property (copy, readonly) ExcelChart *chart;  // Returns a chart object that represents the chart contained in the object.
@property BOOL enabled;  // Returns or sets if the object is enabled.
@property (readonly) NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property double height;  // Returns or set the height of the object.
@property (copy, readonly) ExcelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property double leftPosition;  // Returns or sets the position of the specified object, in points.
@property BOOL locked;  // Returns or sets if the object is locked, if false the object can be modified when the sheet is protected.
@property (copy) NSString *name;  // Returns or sets the name of the object.
@property (copy) NSString *onAction;  // Returns or sets the name of a string of AppleScript commands that will be executed when the object is clicked on. Please note that if AppleScript commands are set they will not be saved with the document.
@property ExcelXlPlacement placement;  // Returns or sets the way the object is attached to the cells below it.
@property BOOL printObject;  // Returns or sets if this object is printed.
@property BOOL protectChartObject;  // Returns or sets if the embedded chart frame cannot be moved, resized, or deleted.
@property BOOL roundedCorners;  // Returns or sets if the embedded chart has rounded corners.
@property BOOL shadow;  // Returns or sets if the if the object has a shadow.
@property double top;  // Returns or sets the distance from the top edge of the object to the top of row 1, on a worksheet, or the top of the chart area, on a chart.
@property (copy, readonly) ExcelRange *topLeftCell;  // Returns a range object that represents the cell that lies under the upper-left corner of the specified object. 
@property BOOL visible;  // Returns or sets if the object is visible.
@property double width;  // Returns or sets  the width of the object.
@property (readonly) NSInteger zOrderPosition;  // Returns the z-order position of the object.

- (void) bringToFront;  // Brings the object to the front of the z-order.
- (void) copyPictureAppearance:(ExcelXlPictureAppearance)appearance format:(ExcelXlCopyPictureFormat)format NS_RETURNS_NOT_RETAINED;  // Copies the selected object to the clipboard as a picture.
- (void) cut;  // Cuts the object to the clipboard.
- (void) saveAsPicturePictureType:(ExcelMsoPictureType)pictureType fileName:(NSString *)fileName;  // Saves the shape in the requested file using the stated graphic format
- (void) sendToBack;  // Sends the object to the back of the z-order.

@end

// Represents the chart title.
@interface ExcelChartTitle : SBObject <ExcelGenericMethods>

- (SBElementArray<ExcelCharacter *> *) characters;

@property BOOL autoScaleFont;  // Returns or sets  if the text in the object changes font size when the object size changes.
@property (copy, readonly) ExcelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy) NSString *caption;  // Returns or sets the chart title text.
@property (copy, readonly) ExcelChartFillFormat *chartFillFormatObject;  // Returns a chart fill format object for this object.
@property (copy, readonly) ExcelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the chart title.
@property (copy) NSString *chartTitleText;  // Returns or sets the text for the specified object.
@property (copy, readonly) ExcelFont *fontObject;  // Returns a font object that represents the font of the specified object.
@property (copy) NSString *formula;  // Returns or sets the object's formula, in A1-style notation and in the language of the macro.
@property (copy) NSString *formulaLocal;  // Returns or sets the formula for the object, using A1-style references in the language of the user.
@property (copy) NSString *formulaR1c1;  // Returns or sets the formula for the object, using R1C1-style notation in the language of the macro.
@property (copy) NSString *formulaR1c1Local;  // Returns or sets the formula for the object, using R1C1-style notation in the language of the user.
@property (readonly) double height;  // Returns the height of the object.
@property ExcelXlHAlign horizontalAlignment;  // Returns or sets the horizontal alignment for the object.
@property BOOL includeInLayout;  // Returns or sets if a chart title will occupy the chart layout space when a chart layout is being determined.
@property (copy, readonly) ExcelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property double leftPosition;  // Returns or sets the left position of the specified object, in points.
@property (copy, readonly) NSString *name;  // Returns the name of the object.
@property ExcelXlOrientation orientation;  // May also be a number value from -90 to 90 degrees.
@property ExcelXlChartElementPosition position;  // Returns or sets the position of the chart title on the chart.
@property ExcelXLDefaultSheetDir readingOrder;  // Returns or sets the reading order for the specified object.
@property BOOL shadow;  // Returns or sets if the object has a shadow.
@property double top;  // Returns or sets the top position of the specified object, in points.
@property ExcelXlVerticalAlignmentTarget verticalAlignment;  // Returns or sets the vertical alignment of the object.
@property (readonly) double width;  // Returns the width of the object.


@end

// Represents a chart in a workbook. The chart can be either an embedded chart, contained in a chart object, or a separate chart sheet.
@interface ExcelChart : SBObject <ExcelGenericMethods>

- (SBElementArray<ExcelShape *> *) shapes;
- (SBElementArray<ExcelArc *> *) arcs;
- (SBElementArray<ExcelAreaGroup *> *) areaGroups;
- (SBElementArray<ExcelBarGroup *> *) barGroups;
- (SBElementArray<ExcelButton *> *) buttons;
- (SBElementArray<ExcelChartGroup *> *) chartGroups;
- (SBElementArray<ExcelChartObject *> *) chartObjects;
- (SBElementArray<ExcelCheckbox *> *) checkboxes;
- (SBElementArray<ExcelColumnGroup *> *) columnGroups;
- (SBElementArray<ExcelDoughnutGroup *> *) doughnutGroups;
- (SBElementArray<ExcelDropdown *> *) dropdowns;
- (SBElementArray<ExcelGroupbox *> *) groupboxes;
- (SBElementArray<ExcelHyperlink *> *) hyperlinks;
- (SBElementArray<ExcelLabel *> *) labels;
- (SBElementArray<ExcelLineGroup *> *) lineGroups;
- (SBElementArray<ExcelLine *> *) lines;
- (SBElementArray<ExcelListbox *> *) listboxes;
- (SBElementArray<ExcelOptionButton *> *) optionButtons;
- (SBElementArray<ExcelOval *> *) ovals;
- (SBElementArray<ExcelPieGroup *> *) pieGroups;
- (SBElementArray<ExcelRadarGroup *> *) radarGroups;
- (SBElementArray<ExcelRectangle *> *) rectangles;
- (SBElementArray<ExcelScrollbar *> *) scrollbars;
- (SBElementArray<ExcelSeries *> *) seriesCollection;
- (SBElementArray<ExcelSpinner *> *) spinners;
- (SBElementArray<ExcelTextbox *> *) textboxes;
- (SBElementArray<ExcelXyGroup *> *) xyGroups;

@property (copy, readonly) ExcelChartGroup *areaThreeDGroup;  // Returns a chart group object that represents the area chart group on a 3-D chart.
@property BOOL autoScaling;  // Returns or sets if Microsoft Excel scales a 3-D chart so that it's closer in size to the equivalent 2-D chart. The chart's right angle axes property must be true.
@property (copy, readonly) ExcelWalls *backWall;  // Returns a walls object that allows the user to individually format the back wall of a 3-D chart.
@property ExcelXlBarShape barShape;  // Returns or sets the shape used with the 3-D bar or column chart.
@property (copy, readonly) ExcelChartGroup *barThreeDGroup;  // Returns a chart group object that represents the bar chart group on a 3-D chart.
@property (copy, readonly) ExcelChartArea *chartAreaObject;  // Returns a chart area object that represents the complete chart area for the chart.
@property NSInteger chartStyle;  // Returns or sets the chart type.
@property (copy, readonly) ExcelChartTitle *chartTitle;  // Returns a chart title object that represents the title of the specified chart.
@property ExcelXlChartType chartType;  // Returns or sets the chart type.
@property (copy, readonly) ExcelChartGroup *columnThreeDGroup;  // Returns a chart group object that represents the column chart group on a 3-D chart.
@property (copy, readonly) ExcelCorners *cornersObject;  // Returns a corners object that represents the corners of a 3-D chart.
@property (copy, readonly) ExcelDataTable *dataTableObject;  // Returns a data table object that represents the chart data table.
@property NSInteger depthPercent;  // Returns or sets the depth of a 3-D chart as a percentage of the chart width, between 20 and 2000 percent.
@property ExcelXlDisplayBlanksAs displayBlanksAs;  // Returns or sets the way that blank cells are plotted on a chart.
@property NSInteger elevation;  // Returns or sets the elevation of the 3-D chart view, in degrees.
@property (readonly) NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property (copy, readonly) ExcelFloor *floorObject;  // Returns a floor object that represents the floor of the 3-D chart.
@property NSInteger gapDepth;  // Returns or sets the distance between the data series in a 3-D chart, as a percentage of the marker width. The value of this property must be between 0 and 500.
@property BOOL hasDataTable;  // Returns or sets if the chart has a data table.
@property BOOL hasLegend;  // Returns or sets if the chart has a legend.
@property BOOL hasTitle;  // Returns or sets if the chart has a title.
@property NSInteger heightPercent;  // Returns or sets the height of a 3-D chart as a percentage of the chart width, between 5 and 500 percent.
@property (copy, readonly) ExcelLegend *legendObject;  // Returns a legend object that represents the legend for the chart.
@property (copy, readonly) ExcelChartGroup *lineThreeDGroup;  // Returns a chart group object that represents the line chart group on a 3-D chart.
@property (copy) NSString *name;  // Returns or sets the name of the chart.
@property (copy, readonly) ExcelChart *next;  // Returns a worksheet object that represents the next sheet.
@property (copy, readonly) ExcelPageSetup *pageSetupObject;  // Returns the page setup object associated with this chart.
@property NSInteger perspective;  // Returns or sets the perspective for the 3-D chart view. Must be between 0 and 100. This property is ignored if the right angle axes property is true.
@property (copy, readonly) ExcelChartGroup *pieThreeDGroup;  // Returns a chart group object that represents the pie chart group on a 3-D chart.
@property (copy, readonly) ExcelPlotArea *plotAreaObject;  // Returns a plot area object that represents the plot area of a chart.
@property ExcelXlRowCol plotBy;  // Returns or sets the way columns or rows are used as data series on the chart.
@property BOOL plotVisibleOnly;  // Returns or sets if only visible cells are plotted. False if both visible and hidden cells are plotted.
@property (copy, readonly) id previous;  // Returns a worksheet object that represents the previous sheet.
@property (readonly) BOOL protectContents;  // Returns true if the contents of the sheet are protected.
@property BOOL protectData;  // Returns or sets if series formulas cannot be modified by the user.
@property (readonly) BOOL protectDrawingObjects;  // Returns true if shapes are protected.
@property BOOL protectFormatting;  // Returns or sets if chart formatting cannot be modified by the user.
@property BOOL protectGoalSeek;  // Returns or sets if the user cannot modify chart data points with mouse actions.
@property BOOL protectSelection;  // Returns or sets if chart elements cannot be selected.
@property (readonly) BOOL protectionMode;  // Returns true if user-interface-only protection is turned on. To turn on user interface protection, use the protect method with the user interface only argument set to true.
@property BOOL rightAngleAxes;  // Returns or sets if the chart axes are at right angles, independent of chart rotation or elevation. Applies only to 3-D line, column, and bar charts.
@property NSInteger rotation;  // The rotation of the 3D chart view.  The value of must be from 0 to 360.
@property (copy, readonly) ExcelTab *sheetTab;  // Returns the sheet tab of the chart sheet
@property BOOL showDataLabelsOverMaximum;  // Returns or sets whether to show the data labels when the value is greater than the maximum value on the value axis.
@property BOOL showWindow;  // Returns or sets if the embedded chart is displayed in a separate window. The Chart object used with this property must refer to an embedded chart.
@property (copy, readonly) ExcelWalls *sideWall;  // Returns a walls object that allows the user to individually format the side wall of a 3-D chart.
@property BOOL sizeWithWindow;  // Returns or sets if Microsoft Excel resizes the chart to match the size of the chart sheet window. False if the chart size isn't attached to the window size. Applies only to chart sheets, doesn't apply to embedded charts.
@property (copy, readonly) ExcelChartGroup *surfaceGroup;  // Returns a chart group object that represents the surface chart group of a 3-D chart.
@property ExcelXlSheetVisibility visible;  // Returns or sets if the chart is visible.
@property BOOL wallsAndGridlinesTwoD;  // Returns or sets if gridlines are drawn two-dimensionally on a 3-D chart.
@property (copy, readonly) ExcelWalls *wallsObject;  // Returns a walls object that represents the walls of the 3-D chart.

- (void) applyCustomChartTypeChartType:(ExcelXlChartGallery)chartType chartName:(NSString *)chartName;  // Applies a standard or custom chart type to a chart or series.
- (void) applyLayoutLayout:(NSInteger)layout chartType:(ExcelXlChartType)chartType;  // Applies a chart layout.
- (ExcelChart *) chartLocationWhere:(ExcelXlChartLocation)where name:(NSString *)name;  // Moves the chart to a new location.
- (void) chartWizardSource:(ExcelRange *)source gallery:(ExcelXlChartType)gallery format:(NSInteger)format plotBy:(ExcelXlRowCol)plotBy categoryLabels:(NSInteger)categoryLabels seriesLabels:(NSInteger)seriesLabels hasLegend:(BOOL)hasLegend title:(NSString *)title categoryTitle:(NSString *)categoryTitle valueTitle:(NSString *)valueTitle extraTitle:(NSString *)extraTitle;  // Modifies the properties of the given chart. You can use this method to quickly format a chart without setting all the individual properties. This method is noninteractive, and it changes only the specified properties.
- (void) checkSpellingCustomDictionary:(NSString *)customDictionary ignoreUppercase:(BOOL)ignoreUppercase alwaysSuggest:(BOOL)alwaysSuggest;  // Checks the spelling of an object.
- (void) clearToMatchStyle;  // Sets the chart formatting to the last predefined chart style applied.
- (void) copyChartAsPictureAppearance:(ExcelXlPictureAppearance)appearance format:(ExcelXlCopyPictureFormat)format outputSize:(ExcelXlPictureAppearance)outputSize NS_RETURNS_NOT_RETAINED;  // Copies the selected object to the clipboard as a picture.
- (void) deselect;  // Cancels the selection for the specified chart.
- (ExcelAxis *) getAxisAxisType:(ExcelXlAxisType)axisType whichAxis:(ExcelXlAxisGroup)whichAxis;  // Returns an object that represents a specified axis object.
- (NSArray<NSAppleEventDescriptor *> *) getChartElementX:(NSInteger)x y:(NSInteger)y;  // Returns information about the chart element at specified X and Y coordinates. This method returns a list  of three items.  Please refer to the documentation on the meaning of the returned values.
- (BOOL) getHasAxisAxisType:(ExcelXlAxisType)axisType axisGroup:(ExcelXlAxisGroup)axisGroup;  // Returns which axes exist on the chart.
- (void) pasteChartFormat:(ExcelXlPasteType)format;  // Pastes chart data from the clipboard into the specified chart.
- (void) printOutFrom:(NSInteger)from to:(NSInteger)to copies:(NSInteger)copies preview:(BOOL)preview activePrinter:(NSString *)activePrinter printToFile:(BOOL)printToFile collate:(BOOL)collate;  // Prints the object
- (void) printPreviewEnableChanges:(BOOL)enableChanges;  // Shows a preview of the object as it would look when printed. This function is deprecated.
- (void) protectChartPassword:(NSString *)password drawingObjects:(BOOL)drawingObjects chartContents:(BOOL)chartContents userInterfaceOnly:(BOOL)userInterfaceOnly;  // Protects a chart so that it cannot be modified.
- (void) refresh;  // Updates the cache of the chart object.
- (void) saveAsFilename:(NSString *)filename fileFormat:(ExcelXlFileFormat)fileFormat password:(NSString *)password writeReservationPassword:(NSString *)writeReservationPassword readOnlyRecommended:(BOOL)readOnlyRecommended createBackup:(BOOL)createBackup addToMostRecentlyUsedList:(BOOL)addToMostRecentlyUsedList saveAsLocalLanguage:(BOOL)saveAsLocalLanguage;  // Saves changes into a different file.
- (void) setBackgroundPicturePictureFileName:(NSString *)pictureFileName;  // Sets the background graphic for a chart.
- (void) setChartElementChartElement:(id)chartElement;  // Sets chart elements on a chart.
- (void) setHasAxisAxisExists:(BOOL)axisExists axisType:(ExcelXlAxisType)axisType axisGroup:(ExcelXlAxisGroup)axisGroup;  // Sets which axes exist on the chart.
- (void) setSourceDataSource:(ExcelRange *)source plotBy:(ExcelXlRowCol)plotBy;  // Sets the source data range for the chart.
- (void) unprotectPassword:(NSString *)password;  // Removes protection from a sheet or workbook. This method has no effect if the sheet or workbook isn't protected.

@end

// A chart sheet is a worksheet that contains only a chart.
@interface ExcelChartSheet : ExcelChart

@property (copy, readonly) ExcelChart *chart;  // Returns the chart for this chart sheet.
@property (readonly) ExcelXlSheetType worksheetType;  // Returns the type of this worksheet.


@end

@interface ExcelColumnGroup : ExcelChartGroup


@end

// Represents the corners of a 3-D chart.
@interface ExcelCorners : SBObject <ExcelGenericMethods>

@property (copy, readonly) NSString *name;  // Returns the name of this object.


@end

// Represents the data label on a chart point or trendline.
@interface ExcelDataLabel : SBObject <ExcelGenericMethods>

- (SBElementArray<ExcelCharacter *> *) characters;

@property BOOL autoScaleFont;  // Returns or sets if the text in the object changes font size when the object size changes.
@property BOOL autoText;  // Returns or sets if the object automatically generates appropriate text based on context.
@property (copy, readonly) ExcelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy) NSString *caption;  // Returns or sets the caption of the object.
@property (copy, readonly) ExcelChartFillFormat *chartFillFormatObject;  // Returns a chart fill format object for this object.
@property (copy, readonly) ExcelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the data label.
@property (copy) NSString *dataLabelText;  // Returns or sets the text for this object.
@property ExcelXlDataLabelsType dataLabelType;  // Returns or sets the type of the data label.
@property (copy, readonly) ExcelFont *fontObject;  // Returns a font object that represents the font of the specified object.
@property (copy) NSString *formula;  // Returns or sets the object's formula, in A1-style notation and in the language of the macro.
@property (copy) NSString *formulaLocal;  // Returns or sets the formula for the object, using A1-style references in the language of the user.
@property (copy) NSString *formulaR1c1;  // Returns or sets the formula for the object, using R1C1-style notation in the language of the macro.
@property (copy) NSString *formulaR1c1Local;  // Returns or sets the formula for the object, using R1C1-style notation in the language of the user.
@property (readonly) double height;  // Returns the height of the object.
@property ExcelXlHAlign horizontalAlignment;  // Returns or sets the horizontal alignment for the object.
@property (copy, readonly) ExcelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property double leftPosition;  // Returns or sets the left position of the specified object, in points.
@property (copy, readonly) NSString *name;  // Returns the name of the object.
@property (copy) NSString *numberFormat;  // Returns or sets the format code for the object.
@property BOOL numberFormatLinked;  // Returns or sets if the number format is linked to the cells, so that the number format changes in the labels when it changes in the cells.
@property (copy) NSString *numberFormatLocal;  // Returns or sets the format code for the object as a string in the language of the user.
@property NSInteger orientation;  // The text orientation.  Can be a number value from -90 to 90 degrees.
@property ExcelXlDataLabelPosition position;  // Returns or sets the position of the specified object.
@property ExcelXLDefaultSheetDir readingOrder;  // Returns or sets the reading order for the specified object.
@property BOOL shadow;  // Returns or sets if the object has a shadow.
@property BOOL showLegendKey;  // Returns or sets if the data label legend key is visible.
@property double top;  // Returns or sets the top position of the specified object, in points.
@property ExcelXlVerticalAlignmentTarget verticalAlignment;  // Returns or sets the vertical alignment of the object.
@property (readonly) double width;  // Returns the width of the object.


@end

// Represents a chart data table.
@interface ExcelDataTable : SBObject <ExcelGenericMethods>

@property BOOL autoScaleFont;  // Returns or sets if the text in the object changes font size when the object size changes.
@property (copy, readonly) ExcelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy, readonly) ExcelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the data table.
@property (copy, readonly) ExcelFont *fontObject;  // Returns a font object that represents the font of the specified object.
@property BOOL hasBorderHorizontal;  // Returns or sets if the chart data table has horizontal cell borders. 
@property BOOL hasBorderOutline;  // Returns or sets if the chart data table has outline borders.
@property BOOL hasBorderVertical;  // Returns or sets if the chart data table has vertical cell borders. 
@property BOOL showLegendKey;  // Returns or sets if the data label legend key is visible.


@end

// Represents a unit label on an axis in the specified chart. Unit labels are useful for charting large values-- for example, in the millions or billions.
@interface ExcelDisplayUnitLabel : SBObject <ExcelGenericMethods>

- (SBElementArray<ExcelCharacter *> *) characters;

@property BOOL autoScaleFont;  // Returns or sets if the text in the object changes font size when the object size changes.
@property (copy, readonly) ExcelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy) NSString *caption;  // Returns or sets the caption for the object.
@property (copy, readonly) ExcelChartFillFormat *chartFillFormatObject;  // Returns a chart fill format object for this object.
@property (copy, readonly) ExcelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the chart title.
@property (copy) NSString *displayLabelUnitText;  // Returns or sets the text for this object.
@property (copy, readonly) ExcelFont *fontObject;  // Returns a font object that represents the font of the specified object.
@property ExcelXlHAlign horizontalAlignment;  // Returns or sets the horizontal alignment for the object.
@property (copy, readonly) ExcelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property double leftPosition;  // Returns or sets the left position of the specified object, in points.
@property (copy, readonly) NSString *name;  // Returns the name of the object.
@property NSInteger orientation;  // The text orientation.  Can be a number value from -90 to 90 degrees.
@property ExcelXlChartElementPosition position;  // Returns or sets the position of the chart title on the chart.
@property ExcelXLDefaultSheetDir readingOrder;  // Returns or sets the reading order for the specified object.
@property BOOL shadow;  // Returns or sets if the object has a shadow.
@property double top;  // Returns or sets the top position of the specified object, in points.
@property ExcelXlVerticalAlignmentTarget verticalAlignment;  // Returns or sets the vertical alignment of the object.


@end

@interface ExcelDoughnutGroup : ExcelChartGroup


@end

// Represents the down bars in a chart group. Down bars connect points on the first series in the chart group with lower values on the last series, the lines go down from the first series.
@interface ExcelDownBars : SBObject <ExcelGenericMethods>

@property (copy, readonly) ExcelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy, readonly) ExcelChartFillFormat *chartFillFormatObject;  // Returns a chart fill format object for this object.
@property (copy, readonly) ExcelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the down bars.
@property (copy, readonly) ExcelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property (copy, readonly) NSString *name;  // Returns the name of the object.


@end

// Represents the drop lines in a chart group. Drop lines connect the points in the chart with the x-axis. Only line and area chart groups can have drop lines.
@interface ExcelDropLines : SBObject <ExcelGenericMethods>

@property (copy, readonly) ExcelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy, readonly) ExcelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the drop lines.
@property (copy, readonly) NSString *name;  // Returns the name of this object.


@end

// Represents the error bars on a chart series. Error bars indicate the degree of uncertainty for chart data.
@interface ExcelErrorBars : SBObject <ExcelGenericMethods>

@property (copy, readonly) ExcelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy, readonly) ExcelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the error bars.
@property ExcelXlEndStyleCap endStyle;  // Returns or sets the end style for the error bars.
@property (copy, readonly) NSString *name;  // Returns the name of the object.


@end

// Represents the floor of a 3-D chart.
@interface ExcelFloor : SBObject <ExcelGenericMethods>

@property (copy, readonly) ExcelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy, readonly) ExcelChartFillFormat *chartFillFormatObject;  // Returns a chart fill format object for this object.
@property (copy, readonly) ExcelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the floor.
@property (copy, readonly) ExcelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property (copy, readonly) NSString *name;  // Returns the name of the object.
@property ExcelXlChartPictureType pictureType;  // Returns or sets the way pictures are displayed on a column or bar picture chart or on the walls and faces of a 3-D chart.
@property NSInteger thickness;  // Returns or sets  the thickness of the floor.


@end

// Represents major or minor gridlines on a chart axis. Gridlines extend the tick marks on a chart axis to make it easier to see the values associated with the data markers.
@interface ExcelGridlines : SBObject <ExcelGenericMethods>

@property (copy, readonly) ExcelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy, readonly) ExcelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the gridlines.
@property (copy, readonly) NSString *name;  // Returns the name of this object.


@end

// Represents the high-low lines in a chart group. High-low lines connect the highest point with the lowest point in every category in the chart group. Only 2-D line groups can have high-low lines.
@interface ExcelHiloLines : SBObject <ExcelGenericMethods>

@property (copy, readonly) ExcelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy, readonly) ExcelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the hilo lines.
@property (copy, readonly) NSString *name;  // Returns the name of this object.


@end

// Represents the interior of an object.
@interface ExcelInterior : SBObject <ExcelGenericMethods>

@property (copy) NSArray<NSNumber *> *color;  // Returns or sets the primary color of the object.
@property ExcelXlColorIndex colorIndex;  // Returns or sets the color of the interior. The color is specified as an index value into the current color palette.
@property BOOL invertIfNegative;  // Returns or sets if Microsoft Excel inverts the pattern in the item when it corresponds to a negative number.
@property (copy, readonly) ExcelLinearGradient *linearGradient;  // Returns or sets the Gradient property of an Interior object of a selection.
@property ExcelXlPattern pattern;  // Returns or sets the interior pattern.
@property (copy) NSArray<NSNumber *> *patternColor;  // Returns or sets the color of the interior pattern as an RGB value.
@property ExcelXlColorIndex patternColorIndex;  // Returns or sets the color of the interior pattern as an index into the current color palette.
@property ExcelXlThemeColor patternThemeColor;  // Returns or sets a theme color pattern for an Interior object.
@property double patternTintAndShade;  // Returns or sets a tint and shade pattern for an Interior object.
@property (copy, readonly) ExcelRectangularGradient *rectangularGradient;  // Returns or sets the Gradient property of an Interior object of a selection.
@property ExcelXlThemeColor themeColor;  // Returns or sets the theme color in the applied color scheme that is associated with the specified object.
@property double tintAndShade;  // Returns or sets a Single that lightens or darkens a color.


@end

// Represents leader lines on a chart. Leader lines connect data labels to data points.
@interface ExcelLeaderLines : SBObject <ExcelGenericMethods>

@property (copy, readonly) ExcelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy, readonly) ExcelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the leader lines.


@end

// Represents a legend entry in a chart legend.
@interface ExcelLegendEntry : SBObject <ExcelGenericMethods>

@property BOOL autoScaleFont;  // Returns or sets if the text in the object changes font size when the object size changes.
@property (readonly) NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property (readonly) double height;  // Returns the height of the object.
@property (readonly) double leftPosition;  // Returns the left position of the specified object, in points.
@property (readonly) double top;  // Returns the top position of the specified object, in points.
@property (readonly) double width;  // Returns the width of the object.


@end

// Represents a legend key in a chart legend. Each legend key is a graphic that visually links a legend entry with its associated series or trendline in the chart.
@interface ExcelLegendKey : SBObject <ExcelGenericMethods>

@property (copy, readonly) ExcelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy, readonly) ExcelChartFillFormat *chartFillFormatObject;  // Returns a chart fill format object for this object.
@property (copy, readonly) ExcelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the legend key.
@property (readonly) double height;  // Returns the height of the object.
@property (copy, readonly) ExcelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property BOOL invertIfNegative;  // Returns or sets if Microsoft Excel inverts the pattern in the item when it corresponds to a negative number.
@property (readonly) double leftPosition;  // Returns the left position of the specified object, in points.
@property (copy) NSArray<NSNumber *> *markerBackgroundColor;  // Returns or sets the marker background color as an RGB value. Applies only to line, scatter, and radar charts.
@property ExcelXlColorIndex markerBackgroundColorIndex;  // Returns or sets the marker background color as an index into the current color palette, or as one of the following constants: color index automatic or color index none. Applies only to line, scatter, and radar charts.
@property (copy) NSArray<NSNumber *> *markerForegroundColor;  // Returns or sets the foreground color of the marker as an RGB value. Applies only to line, scatter, and radar charts.
@property ExcelXlColorIndex markerForegroundColorIndex;  // Returns or sets the marker foreground color as an index into the current color palette, or as one of the following constants: color index automatic or color index none. Applies only to line, scatter, and radar charts.
@property NSInteger markerSize;  // Returns or sets the data-marker size, in points.
@property ExcelXlMarkerStyle markerStyle;  // Returns or sets the marker style for a point or series in a line chart, scatter chart, or radar chart.
@property ExcelXlChartPictureType pictureType;  // Returns or sets the way pictures are displayed on a column or bar picture chart or on the walls and faces of a 3-D chart.
@property double pictureUnit;  // Returns or sets the unit for each picture on the chart if the picture type property is set to chart picture type stack scale, if not, this property is ignored.
@property BOOL shadow;  // Returns or sets if the object has a shadow.
@property BOOL smooth;  // Returns or sets if curve smoothing is turned on for the line chart or scatter chart. Applies only to line and scatter charts.
@property (readonly) double top;  // Returns the top position of the specified object, in points.
@property (readonly) double width;  // Returns the width of the object.


@end

// Represents the legend in a chart. Each chart can have only one legend.
@interface ExcelLegend : SBObject <ExcelGenericMethods>

- (SBElementArray<ExcelLegendEntry *> *) legendEntries;

@property BOOL autoScaleFont;  // Returns or sets if the text in the object changes font size when the object size changes.
@property (copy, readonly) ExcelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy, readonly) ExcelChartFillFormat *chartFillFormatObject;  // Returns a chart fill format object for this object.
@property (copy, readonly) ExcelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the legend.
@property (copy, readonly) ExcelFont *fontObject;  // Returns a font object that represents the font of the specified object.
@property double height;  // Returns or sets the height of the object.
@property BOOL includeInLayout;  // Returns or sets if a legend will occupy the chart layout space when a chart layout is being determined.
@property (copy, readonly) ExcelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property double leftPosition;  // Returns or sets the left position of the specified object, in points.
@property (copy, readonly) NSString *name;  // Returns the name of the object.
@property ExcelXlLegendPosition position;  // Returns or sets the position of the legend on the chart.
@property BOOL shadow;  // Returns or sets if the object has a shadow.
@property double top;  // Returns or sets the top position of the specified object, in points.
@property double width;  // Returns or sets  the width of the object.


@end

@interface ExcelLineGroup : ExcelChartGroup


@end

@interface ExcelPieGroup : ExcelChartGroup


@end

// Represents the plot area of a chart. This is the area where your chart data is plotted.
@interface ExcelPlotArea : SBObject <ExcelGenericMethods>

@property (copy, readonly) ExcelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy, readonly) ExcelChartFillFormat *chartFillFormatObject;  // Returns a chart fill format object for this object.
@property (copy, readonly) ExcelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the plot area.
@property double height;  // Returns or sets the height of the object.
@property (readonly) double insideHeight;  // Returns the inside height of the plot area, in points.
@property (readonly) double insideLeft;  // Returns the distance from the chart edge to the inside left edge of the plot area, in points.
@property (readonly) double insideTop;  // Returns the distance from the chart edge to the inside top edge of the plot area, in points.
@property (readonly) double insideWidth;  // Returns the inside width of the plot area, in points.
@property (copy, readonly) ExcelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property double leftPosition;  // Returns or sets the left position of the specified object, in points.
@property (copy, readonly) NSString *name;  // Returns the name of the object.
@property ExcelXlChartElementPosition position;  // Returns or sets the position of the plot area on the chart.
@property double top;  // Returns or sets the top position of the specified object, in points.
@property double width;  // Returns or sets  the width of the object.


@end

@interface ExcelRadarGroup : ExcelChartGroup


@end

// Represents series lines in a chart group. Series lines connect the data values from each series. Only 2-D stacked bar or column chart groups can have series lines.
@interface ExcelSeriesLines : SBObject <ExcelGenericMethods>

@property (copy, readonly) ExcelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy, readonly) ExcelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the series lines.
@property (copy, readonly) NSString *name;  // Returns the name of this object.


@end

// Represents a single point in a series in a chart.
@interface ExcelSeriesPoint : SBObject <ExcelGenericMethods>

@property BOOL applyPictToEnd;  // Returns or sets if a picture is applied to the end of the point or all points in the series.
@property BOOL applyPictToFront;  // Returns or sets if a picture is applied to the front of the point or all points in the series.
@property BOOL applyPictToSides;  // Returns or sets if a picture is applied to the sides of the point or all points in the series.
@property (copy, readonly) ExcelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy, readonly) ExcelChartFillFormat *chartFillFormatObject;  // Returns a chart fill format object for this object.
@property (copy, readonly) ExcelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the point.
@property (copy, readonly) ExcelDataLabel *dataLabelObject;  // Returns a data label object that represents the data label associated with the point or trendline.
@property NSInteger explosion;  // Returns or sets the explosion value for a pie-chart or doughnut-chart slice. Returns zero if there's no explosion, the tip of the slice is in the center of the pie.
@property BOOL hasDataLabel;  // Returns or sets if the point has a data label.
@property BOOL hasThreeDEffect;  // Returns or sets if a point has a three-dimensional appearance. Applies only to bubble charts.
@property (readonly) double height;  // Returns the height of the object.
@property (copy, readonly) ExcelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property BOOL invertIfNegative;  // Returns or sets if Microsoft Excel inverts the pattern in the item when it corresponds to a negative number.
@property (readonly) double leftPosition;  // Returns the left position of the specified object, in points.
@property (copy) NSArray<NSNumber *> *markerBackgroundColor;  // Returns or sets the marker background color as an RGB value. Applies only to line, scatter, and radar charts.
@property ExcelXlColorIndex markerBackgroundColorIndex;  // Returns or sets the marker background color as an index into the current color palette, or as one of the following constants: color index automatic or color index none. Applies only to line, scatter, and radar charts.
@property (copy) NSArray<NSNumber *> *markerForegroundColor;  // Returns or sets the foreground color of the marker as an RGB value. Applies only to line, scatter, and radar charts.
@property ExcelXlColorIndex markerForegroundColorIndex;  // Returns or sets the marker foreground color as an index into the current color palette, or as one of the following constants: color index automatic or color index none. Applies only to line, scatter, and radar charts.
@property NSInteger markerSize;  // Returns or sets the data-marker size, in points.
@property ExcelXlMarkerStyle markerStyle;  // Returns or sets the marker style for a point or series in a line chart, scatter chart, or radar chart.
@property (copy, readonly) NSString *name;  // Returns the name of the object.
@property ExcelXlChartPictureType pictureType;  // Returns or sets the way pictures are displayed on a column or bar picture chart or on the walls and faces of a 3-D chart.
@property double pictureUnit;  // Returns or sets the unit for each picture on the chart if the picture type property is set to chart picture type stack scale, if not, this property is ignored.
@property BOOL secondaryPlot;  // Returns or sets if the point is in the secondary section of either a pie of pie chart or a bar of pie chart. Applies only to points on pie of pie charts or bar of pie charts.
@property BOOL shadow;  // Returns or sets if the object has a shadow.
@property (readonly) double top;  // Returns the top position of the specified object, in points.
@property (readonly) double width;  // Returns the width of the object.


@end

// Represents a series in a chart.
@interface ExcelSeries : SBObject <ExcelGenericMethods>

- (SBElementArray<ExcelDataLabel *> *) dataLabels;
- (SBElementArray<ExcelSeriesPoint *> *) seriesPoints;
- (SBElementArray<ExcelTrendline *> *) trendlines;

@property BOOL applyPictureToEnd;  // Returns or sets if a picture is applied to the end of the point or all points in the series.
@property BOOL applyPictureToFront;  // Returns or sets if a picture is applied to the front of the point or all points in the series.
@property BOOL applyPictureToSides;  // Returns or sets if a picture is applied to the sides of the point or all points in the series.
@property ExcelXlAxisGroup axisGroup;  // Returns or sets the axis group for the specified series.
@property ExcelXlBarShape barShape;  // Returns or sets the shape used with the 3-D bar or column chart.
@property (copy, readonly) ExcelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy) NSString *bubbleSizes;  // Returns or sets a string in A1-style notation that refers to the worksheet cells containing the size data for the bubble chart. Applies only to bubble charts. 
@property (copy, readonly) ExcelChartFillFormat *chartFillFormatObject;  // Returns a chart fill format object for this object.
@property (copy, readonly) ExcelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the series.
@property ExcelXlChartType chartType;  // Returns or sets the chart type.
@property (copy, readonly) ExcelErrorBars *errorBars;  // Returns an error bars object that represents the error bars for the series.
@property NSInteger explosion;  // Returns or sets the explosion value for a pie-chart or doughnut-chart slice. Returns zero if there's no explosion, the tip of the slice is in the center of the pie.
@property (copy) NSString *formula;  // Returns or sets the object's formula, in A1-style notation and in the language of the macro.
@property (copy) NSString *formulaLocal;  // Returns or sets the formula for the object, using A1-style references in the language of the user.
@property (copy) NSString *formulaR1c1;  // Returns or sets the formula for the object, using R1C1-style notation in the language of the macro.
@property (copy) NSString *formulaR1c1Local;  // Returns or sets the formula for the object, using R1C1-style notation in the language of the user.
@property BOOL hasDataLabels;  // Returns or sets if the series has data labels.
@property BOOL hasErrorBars;  // Returns or set if the series has error bars. This property isn't available for 3-D charts.
@property BOOL hasLeaderLines;  // Returns or sets if the series has leader lines.
@property BOOL hasThreeDEffect;  // Returns or sets if the series has a three-dimensional appearance. Applies only to bubble charts.
@property (copy, readonly) ExcelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property BOOL invertIfNegative;  // Returns or sets if Microsoft Excel inverts the pattern in the item when it corresponds to a negative number.
@property (copy, readonly) ExcelLeaderLines *leaderLines;  // Returns a leader lines object that represents the leader lines for the series.
@property (copy) NSArray<NSNumber *> *markerBackgroundColor;  // Returns or sets the marker background color as an RGB value. Applies only to line, scatter, and radar charts.
@property ExcelXlColorIndex markerBackgroundColorIndex;  // Returns or sets the marker background color as an index into the current color palette, or as one of the following constants: color index automatic or color index none. Applies only to line, scatter, and radar charts.
@property (copy) NSArray<NSNumber *> *markerForegroundColor;  // Returns or sets the foreground color of the marker as an RGB value. Applies only to line, scatter, and radar charts.
@property ExcelXlColorIndex markerForegroundColorIndex;  // Returns or sets the marker foreground color as an index into the current color palette, or as one of the following constants: color index automatic or color index none. Applies only to line, scatter, and radar charts.
@property NSInteger markerSize;  // Returns or sets the data-marker size, in points.
@property ExcelXlMarkerStyle markerStyle;  // Returns or sets the marker style for a point or series in a line chart, scatter chart, or radar chart.
@property (copy) NSString *name;  // Returns or sets the name of the object.
@property ExcelXlChartPictureType pictureType;  // Returns or sets the way pictures are displayed on a column or bar picture chart or on the walls and faces of a 3-D chart.
@property double pictureUnit;  // Returns or sets the unit for each picture on the chart if the picture type property is set to chart picture type stack scale, if not, this property is ignored.
@property (readonly) NSInteger plotColorIndex;  // Returns the plot color index of the series.
@property NSInteger plotOrder;  // Returns or sets the plot order for the selected series within the chart group.
@property ExcelXlCategoryNames seriesValues;  // Returns or sets a list of all the values in the series. This can be a range on a worksheet or a list of constant values.
@property BOOL shadow;  // Returns or sets if the object has a shadow.
@property BOOL smooth;  // Returns or sets if curve smoothing is turned on for the line chart or scatter chart. Applies only to line and scatter charts.
@property ExcelXlCategoryNames xvalues;  // Returns or sets a list of x values for a chart series. The xvalues property can be set to a range on a worksheet or to a list of values.

- (void) errorBarDirection:(ExcelXlErrorBarDirection)direction include:(ExcelXlErrorBarInclude)include type:(ExcelXlErrorBarType)type amount:(double)amount minusValues:(double)minusValues;  // Applies error bars to the series.
- (void) pasteSeries;

@end

// Represents the tick-mark labels associated with tick marks on a chart axis.
@interface ExcelTickLabels : SBObject <ExcelGenericMethods>

@property BOOL autoScaleFont;  // Returns or sets if the text in the object changes font size when the object size changes.
@property (copy, readonly) ExcelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the tick labels.
@property (readonly) NSInteger depth;  // Returns the number of levels of category tick labels.
@property (copy, readonly) ExcelFont *fontObject;  // Returns a font object that represents the font of the specified object.
@property BOOL multiLevel;  // Returns or sets whether an axis is multilevel or not.
@property (copy, readonly) NSString *name;  // Returns the name of the object.
@property (copy) NSString *numberFormat;  // Returns or sets the format code for the object.
@property BOOL numberFormatLinked;  // Returns or sets if the number format is linked to the cells, so that the number format changes in the labels when it changes in the cells.
@property (copy) NSString *numberFormatLocal;  // Returns or sets the format code for the object as a string in the language of the user.
@property NSInteger offset;  // Returns or sets the distance between the levels of labels, and the distance between the first level and the axis line. The value can be an integer percentage from 0 through 1000, relative to the axis label's font size. 
@property ExcelXlTickLabelOrientation orientation;  // The text orientation.  Can be a number value from -90 to 90 degrees.
@property ExcelXLDefaultSheetDir readingOrder;  // Returns or sets the reading order for the specified object.
@property ExcelXlTickHAlign tickAlignment;  // Returns or sets the alignment for the specified tick label.


@end

// Represents a trendline in a chart. A trendline shows the trend, or direction, of data in a series.
@interface ExcelTrendline : SBObject <ExcelGenericMethods>

@property double backward;  // Returns or sets the number of periods, or units on a scatter chart, that the trendline extends backward.
@property (copy, readonly) ExcelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy, readonly) ExcelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the trendline.
@property (copy, readonly) ExcelDataLabel *dataLabelObject;  // Returns a data label object that represents the data label associated with the point or trendline.
@property BOOL displayRSquared;  // Returns or sets if the R-squared value of the trendline is displayed on the chart, in the same data label as the equation. Setting this property to true automatically turns on data labels.
@property BOOL displayEquation;  // Returns or sets if the equation for the trendline is displayed on the chart, in the same data label as the R-squared value. Setting this property to true automatically turns on data labels.
@property (readonly) NSInteger entry_index;  // Returns the index number of the object within the elements of the parent object.
@property double forward;  // Returns or sets the number of periods, or units on a scatter chart, that the trendline extends forward.
@property double intercept;  // Returns or sets the point where the trendline crosses the value axis.
@property BOOL interceptIsAuto;  // Returns or sets if the point where the trendline crosses the value axis is automatically determined by the regression.
@property (copy) NSString *name;  // Returns or sets the name of this object.
@property BOOL nameIsAuto;  // Returns or sets if Microsoft Excel automatically determines the name of the trendline.
@property NSInteger order;  // Returns or sets the trendline order, an integer greater than 1,  when the trendline type is polynomial.
@property NSInteger period;  // Returns or sets the period for the moving-average trendline.
@property ExcelXlTrendlineType trendlineType;  // Returns or sets the type of this trend line


@end

// Represents the up bars in a chart group. Up bars connect points on series one with higher values on the last series in the chart group, the lines go up from series one. Only 2-D line groups that contain at least two series can have up bars.
@interface ExcelUpBars : SBObject <ExcelGenericMethods>

@property (copy, readonly) ExcelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy, readonly) ExcelChartFillFormat *chartFillFormatObject;  // Returns a chart fill format object for this object.
@property (copy, readonly) ExcelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the up bars.
@property (copy, readonly) ExcelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property (copy, readonly) NSString *name;  // Returns the name of this object.


@end

// Represents the walls of a 3-D chart.
@interface ExcelWalls : SBObject <ExcelGenericMethods>

@property (copy, readonly) ExcelBorder *border;  // Returns a border object that represents the border of the object.
@property (copy, readonly) ExcelChartFillFormat *chartFillFormatObject;  // Returns a chart fill format object for this object.
@property (copy, readonly) ExcelChartFormat *chartFormat;  // Returns a chart format object that contains chart formatting properties for the walls.
@property (copy, readonly) ExcelInterior *interiorObject;  // Returns an interior object that represents the interior of the specified object.
@property (copy, readonly) NSString *name;  // Returns or sets the name of the object.
@property ExcelXlChartPictureType pictureType;  // Returns or sets the way pictures are displayed on a column or bar picture chart or on the walls and faces of a 3-D chart.
@property NSInteger pictureUnit;  // Returns or sets the unit for each picture on the chart if the picture type property is set to chart picture type stack scale, if not, this property is ignored.
@property NSInteger thickness;  // Returns or sets  the thickness of the floor.


@end

@interface ExcelXyGroup : ExcelChartGroup


@end


Attribute VB_Name = "MyGeneralOperations"
Option Explicit

' MyGeneralOperations
' Jeff Jenness
' Jenness Enterprises
' http://www.jennessent.com

'  BasicStatsFromArray - GIVEN anArray, Field Name, Table Name, and Application, _
            Returns Sum, Mean, Minimum, Maximum, Range, Count, StDev, Variance, Median, Standard Error of Mean and Mode String
'  BasicStatsFromArraySimple - Simplified version of BasicStatsFromArray, with no progress info and no mode data.
'           GIVEN anArray, optional directive to calculate standard deviation info,
'           Returns Sum, Mean, Minimum, Maximum, Range, Count, StDev, Variance, Median, Standard Error of Mean
'  BasicStatsFromArraySimpleFast - SAME AS BasicStatsFromArraySimple EXCEPT THAT IT IS A SUB AND THEREFORE FILLS VALUES INSTEAD OF
'           RETURNING THEM, AND IT USES NO ARCOBJECTS.
'  BasicStatsFromArray_Weighted - GIVEN 2-Dimensional anArray with Values and Weights, Field Name, Table Name, and Application, _
            Returns Weighted Mean, Weighted StDev and Weighted Variance
'  BasicStatsFromArray_WeightedFast - given 2-Dimensional array with values and weights, fills mean, standard deviation and variance arguments
'  BasicStatsFromVAT - GIVEN 2-Dimensional Arrays (Value and Size, both sorted by Value), Field Name, Table Name, and Application, _
            Returns Sum, Mean, Minimum, Maximum, Range, Count, StDev, Variance, Median, Standard Error of Mean and Mode String
'  CalcStatistics - GIVEN AN ARRAY OF DOUBLES AND AN ARRAY OF BOOLEAN STAT OPTIONS, RETURNS AN ARRAY OF STATISTICS
'  CheckCollectionForKey - GIVEN pCollection and STRING, RETURNS BOOLEAN INDICATING WHETHER COLLECTION HAS THAT KEY OR NOT
'  CheckIfCompressedFGDB - GIVEN pFeatureClass, RETURNS BOOLEAN IF THIS FEATURE CLASS COMES FROM A COMPRESSED FILE GEODATABASE
'  CheckIfFeatureClassExists - GIVEN FEATURE WORKSPACE AND FEATURE CLASS NAME, RETURNS BOOLEAN IF FEATURE CLASS ALREADY EXISTS
'  CheckIfFieldNameExists - GIVEN AN ITable, IFeatureClass, IFields, IStringArray or IVariantArray of Fields, checks to see if a specified field name already exists.
'  CheckIfTableExists - GIVEN FEATURE WORKSPACE AND TABLE NAME, RETURNS BOOLEAN IF TABLE WITH THAT NAME ALREADY EXISTS
'  CheckNumericReal - GIVEN KEYASCII AND TEXTBOX, RESTRICTS INPUT TO ANY REAL NUMBER
'  CheckNumericRealPositive - GIVEN KEYASCII AND TEXTBOX, RESTRICTS INPUT TO ANY POSITIVE REAL NUMBER
'  CheckNumericInteger - GIVEN KEYASCII AND TEXTBOX, RESTRICTS INPUT TO ANY INTEGER
'  CheckNumericIntegerPositive - GIVEN KEYASCII AND TEXTBOX, RESTRICTS INPUT TO ANY POSITIVE INTEGER
'  CheckSpRefDomain - GIVEN SPATIAL REFERENCE, RETURNS BOOLEAN STATING WHETHER IT HAS A VALID XY DOMAIN
'  ClipNumberOfCharacters - GIVEN A STRING, CLIPS OFF THE SPECIFIED NUMBER OF CHARACTERS FROM ETIHER END.
'  ColorToRGB - GIVEN A LONG NUMBER, FILLS THE RED, GREEN AND BLUE COMPONENTS
'  CompareFunctionsInModules - COMPARES FUNCTIONS AND SUBS IN TWO MODULES TO IDENTIFY FUNCTIONS THAT ONLY EXIST IN ONE OF THEM.
'  CompareSpatialReferences - RETURNS A BOOLEAN STATING WHETHER PROJECTION / DATUM ARE SAME.
'  CompareSpatialReferences2 - LIKE CompareSpatialReferences, BUT CORRECTLY LETS YOU CHOOSE TO INCLUDE PRECISION IN COMPARISON
'  ConvertEsriDoubleArrayToVB - GIVEN AN EsriSystem.IDoubleArray, returns a Double().
'  ConvertFClassPathToWSData - GIVEN A PATH NAME, RETURNS APPROPRIATE WORKSPACE, DIRECTORY, FILENAME AND BOOLEANS INDICATING
'           WHETHER THE DIRECTORY AND FILENAME EXIST.
'  ConvertLayoutGraphics - SAMPLE CODE SNIPPET TO CONVERT GRAPHIC POLYGONS DRAWN IN THE LAYOUT INTO SEMI-TRANSPARENT
'           FEATURE LAYERS IN THE DATA FRAME.
'  ConvertLongBinary - GIVEN A LONG AND OPTIONAL NUMBER OF CHARACTERS, RETURNS BINARY REPRESENTATION.
'  ConvertNumberToBullet - GIVEN A NUMBER AND TYPE, RETURNS LETTER OR ROMAN NUMERAL
'  ConvertNumberBase - GIVEN A LONG AND BASE, RETURNS LONG ARRAY CONTAINING NUMERICAL INDEX FOR EACH POSITION
'  ConvertNumberToRoman - GIVEN A LONG AND BOOLEAN FOR UPPERCASE, RETURNS ROMAN NUMERAL EQUIVALENT
'  CountGraphicsByName - GIVEN A NAME AND MAP DOCUMENT, COUNTS ALL GRAPHICS WITH A PARTICULAR NAME
'  CreateDatasetFeatureClass - GIVEN A FEATURE DATASET, FEATURE CLASS NAME, OPTIONAL GOEMETRY TYPE AND FIELD ARRAY, RETURNS A
'           NEW FEATURE CLASS INSIDE THE SPECIFIED FEATURE DATASET:  MODIFIED FROM ESRI SAMPLE
'  CreateFieldAttributeIndex - GIVEN A FIELD NAME AND TABLE/FEATURE CLASS, CREATES AN INDEX FOR THAT FIELD.
'  CreateGDBFeatureClass - GIVEN A FILE OR PERSONAL GEODATABASE, FEATURE CLASS NAME, OPTIONAL GOEMETRY TYPE, FIELD ARRAY, ETC, RETURNS A
'           NEW GEODATABASE FEATURE CLASS:  MODIFIED FROM ESRI SAMPLE
'  CreateGDBTable - GIVEN A FILE OR PERSONAL GEODATABASE, TABLE NAME, FIELD ARRAY, ETC, RETURNS A
'           NEW GEODATABASE TABLE:  MODIFIED FROM ESRI SAMPLE FOR CREATING GEODATABASE FEATURE CLASSES
'  CreateGeneralFeatureClass - RETURNS A NEW FEATURE CLASS IN A FOLDER (I.E. SHAPEFILE) OR IN A FILE / PERSONAL GEODATABASE.
'           CAN SEND IT SPECIFIC INSTRUCTIONS REGARDING NAME AND LOCATION, OR CAN HAVE IT ASK THE USER.
'  CreateGeneralGeographicSpatialReference - GIVEN A FACTORY CODE NUMBER, RETURNS THE PROJECTED COORDINATE SYSTEM
'  CreateGeneralProjectedSpatialReference - GIVEN A FACTORY CODE NUMBER, RETURNS THE GEOGRAPHIC COORDINATE SYSTEM
'  CreateGeneralTable - RETURNS A NEW TABLE IN A FOLDER (I.E. dBASE TABLE) OR IN A FILE / PERSONAL GEODATABASE.
'           CAN SEND IT SPECIFIC INSTRUCTIONS REGARDING NAME AND LOCATION, OR CAN HAVE IT ASK THE USER.
'  CreateInMemoryFeatureClass - GIVEN AN IARRAY OF GEOMETRIES (ALL OF SAME TYPE), RETURNS AN IN-MEMORY FEATURE CLASS.
'  CreateInMemoryFeatureClass_Empty - CREATES AN EMPTY IN-MEMORY FEATURE CLASS THAT CAN BE ADDED TO LIKE ANY OTHER FEATURE CLASS
'  CreateInMemoryFeatureClass2 - GIVEN AN IARRAY OF GEOMETRIES (ALL OF SAME TYPE), WITH OPTIONAL MULTIPLE FIELDS AND PAPP REFERENCE,
'           RETURNS AN IN-MEMORY FEATURE CLASS.
'  CreateInMemoryFeatureClass3 - SAME AS CreateInMemoryFeatureClass2, EXCEPT THAT ADDS FEATURES WITH BUFFER INSTEAD OF CREATING INDIVIDUAL FEATURES.
'           THEREFORE IT RUNS FASTER THAN CreateInMemoryFeatureClass2.  INCLUDES AN OPTIONAL FLUSH VALUE, DEFAULTING TO 500 FEATURES.
'  CreateNAD27_NAD83_GeoTransformationFlagstaffReturns a Geographic Transformation to convert from NAD27 to NAD83, using esriSRGeoTransformation_NAD_1927_TO_NAD_1983_NADCON (United States - lower 48 states)
'  CreateNAD27_WGS84_GeoTransformationFlagstaff:  Returns a Geographic Transformation to convert from NAD27 to WGS84, using esriSRGeoTransformation_NAD1927_To_WGS1984_4 (United States - lower 48 states)
'  CreateNAD83_WGS84_GeoTransformationFlagstaff:  Returns a Geographic Transformation to convert from NAD83 to WGS84, using esriSRGeoTransformation_NAD1983_To_WGS1984_5 (United States - CORS ITRF96)
'  CreateNestedFoldersByPath - GIVEN A PATH NAME, CREATES ANY MISSING FOLDERS IN THAT PATH
'  CreateShapefileFeatureClass - GIVEN A FOLDER, NAME, SPATIAL REFERENCE, GEOMETRY TYPE, OPTIONAL FIELD ARRAY,
'           RETURNS SHAPEFILE FEATURE CLASS.
'  CreateSpatialReferenceWGS84 - RETURNS A WGS84 GEOGRAPHIC COORDINATE SYSTEM
'  CreateSpatialReferenceNAD27 - RETURNS A NAD 1927 GEOGRAPHIC COORDINATE SYSTEM
'  CreateSpatialReferenceNAD83 - RETURNS A NAD 1983 GEOGRAPHIC COORDINATE SYSTEM
'  CreatedBASETableInFolder - GIVEN A FOLDER, NAME, OPTIONAL FIELD ARRAY, CREATEDS dBASE TABLE AND RETURNS ITable OBJECT.
'  CursorToSet_Features - GIVEN A FEATURE CURSOR, RETURNS AN ISET OF FEATURES ALLOWING SET TO BE RUN THROUGH MULTIPLE TIMES
'  CursorToSet_TableRow - GIVEN A CURSOR, RETURNS AN ISET OF ROWS ALLOWING SET TO BE RUN THROUGH MULTIPLE TIMES
'  CursorToVariant_Features - GIVEN A FEATURE CURSOR, RETURNS A VARIANT ARRAY OF FEATURES ALLOWING SET TO BE RUN THROUGH MULTIPLE TIMES
'  CursorToVariant_TableRow - GIVEN A CURSOR, RETURNS A VARIANT ARRAY OF ROWS ALLOWING SET TO BE RUN THROUGH MULTIPLE TIMES
'  DateComponentsFromDate - GIVEN A DATE OBJECT, RETURNS VARIOUS COMPONENTS (MONTH, NUMERIC, LONG AND ABBREVIATED, DAY, YEAR, HOUR, MINUTE, SECOND)
'  DateDiffByDayMonthYear - GIVEN TWO DATES, RETURNS DIFFERENCE IN YEARS, MONTHS AND DAYS.  OPTIONAL TO RETURN ABSOLUTE VALUE IF NEGATIVE
'  Date_MonthNameFromNumber - GIVEN A MONTH NUMBER, RETURNS NAME OR ABBREVIATION IN ENGLISH
'  Date_ReturnMonthNumberFromName - GIVEN A MONTH NAME OR ABBREVIATION, RETURNS A DATE-CONVERTABLE NAME AND MONTH NUMBER
'  DateToYearDecimal - GIVEN A DATE, RETURNS DOUBLE YEAR VALUE (JULY 1, 2010 RETURNS ROUGHLY 2010.5)
'  DeleteGraphicsByGeometry - DELETES ALL GRAPHIC ELEMENTS THAT INTERSECT A GEOMETRY
'  DeleteGraphicsByName - GIVEN A NAME AND MAP DOCUMENT, DELETES ALL GRAPHICS WITH A PARTICULAR NAME
'  DeleteGraphicsByNameByBounds - GIVEN A NAME AND MAP DOCUMENT, DELETES ALL GRAPHICS WITH A PARTICULAR NAME USING BOUNDS OF GRAPHIC SYMBOLOGY
'  DoSpatialQuery - SELECTS ALL FEATURES FROM FIRST FEATURE LAYER THAT INTERSECT SELECTED FEATURES OF SECONDS FEATURE LAYER
'  EnableSelectTool - GIVEN Application, CLICKS THE "SELECT ELEMENTS" TOOL
'  FillFunctionArrayAndCollection - USED BY CompareFunctionsInModules TO FILL ARRAY AND COLLECTION OF FUNCTION AND SUB NAMES
'  FixAtDecimalLevel - TRUNCATES A NUMBER AT A GIVEN DECIMAL LEVEL
'  ForceUppercase - GIVEN KEYASCII AND TEXTBOX, FORCES ENTERED KEYSTROKE TO BE UPPERCASE
'  FormatBySize - QUICK FUNCTION TO ESTIMATE NUMBER OF DECIMAL POINTS BASED ON SIZE OF NUMBER
'  Get_Element_Or_Envelope_Point - USED IN CONJUNCTION WITH Move_Element; THIS FUNCTION RETURNS A POINT AT A PARTICULAR LOCATION IN AN ENVELOPE OR
'           ELEMENT, SUCH AS THE UPPER LEFT CORNER OR THE CENTER OF THE ENVELOPE.  Move_Element HELPS YOU POSITION ONE OBJECT
'           IN RELATION TO ANOTHER, SUCH AS LINED UP ON AN EDGE.
'  GetPart - GIVEN A STRING AND DELIMITER, CLIPS OFF PORTION OF STRING IN FRONT OF DELIMITER AND RETURNS IT.  THE ORIGINAL
'           STRING HAS THE PART AND DELIMITER REMOVED FROM THE FRONT.  USED BY CreateNestedFoldersByPath FUNCTION
'  Graphic_MakeFromGeometry - GIVEN A MAP DOCUMENT, GEOMETRY AND OPTIONAL NAME AND SYMBOLOGY, ADDS GRAPHIC TO MAP.
'  Graphic_MakeFromGeometry - GIVEN A SPECIFIC MAP IN A MAP DOCUMENT, GEOMETRY AND OPTIONAL NAME AND SYMBOLOGY, ADDS GRAPHIC TO MAP.
'  Graphic_ReturnElementFromGeometry - GIVEN MAP DOC, GEOMETRY, OPTIONAL NAME AND OPTIONAL ADD-TO-VIEW, RETURNS THE GRAPHIC
'           ELEMENT
'  Graphic_ReturnElementFromGeometry2 - SAME AS ABOVE, BUT WITH OPTION TO ADD A SYMBOL TO THE GRAPHIC ELEMENT
'  Graphic_ReturnElementFromGeometry3 - SAME AS ABOVE, BUT WITH OPTION TO ELEMENT TO MAP DOC
'  GraphicsSetNameSelected - GIVEN A MAP DOCUMENT AND NAME, ASSIGNS THAT NAME TO ALL SELECTED GRAPHICS
'  HexifyName - GIVEN A WORD, CONVERTS EACH CHARACTER INTO THE HEX EQUIVALENT AND RUNS IT ALL TOGETHER.  USEFUL FOR CASES WHERE
'        WE NEED TO USE A COLLECTION IN A CASE-SENSITIVE MANNER.  USE THIS HEXIFIED-WORD INSTEAD OF THE ACTUAL WORD.
'  ImportASCIIToFileGDB - GIVEN FILENAME FOR CSV FILE, IMPORTS TABLE TO FILE GEODATABASE WHILE ACCEPTING LARGE TEXT FIELDS AND CHECKING
'        ALL RECORDS TO GET CORRECT FIELD TYPE.  NOTE:  ONLY WORKS ON FILES < 2GB
'  IsDimmed - given an array, returns boolean for whether it has been dimmed or not yet.
'  IsFolder_FalseIfCrash - GetAttr crashes on locked files, so this just returns false if file is locked.
'  IsNaN - GIVEN A VARIANT, RETURNS A BOOLEAN INDICATING WHETHER A NUMBER IS A "NOT-A-NUMBER" VALUE
'  IsNormal_FalseIfCrash - GetAttr crashes on locked files, so this just returns false if file is locked.
'  MakeColorRGB - GIVEN RED, GREEN AND BLUE, RETURN ICOLOR
'  MakeColorHSV - GIVEN HUE, SATURATION AND VALUE, RETURN ICOLOR
'  MakeFClassBorderAroundJennessentCompassRose - MAKES A SEMI-TRANSPARENT BORDER AROUND JENNESSENT COMPASS ROSE, SAVED AS FEATURE CLASS
'  MakeFieldNameArray - GIVEN ORIGINAL FEATURE CLASS, NEW FEATURE CLASS, AND IVARIANARRAY OF NEW FIELD NAMES,
'                   RETURN ARRAY CONTAINING INDEX VALUES FOR EACH FIELD NAME
'  MakeFieldNameVarArrayFromTable - SAME AS MakeFieldNameArray, EXCEPT THAT IT TAKES AN ORIGINAL TABLE INSTEAD OF ORIGINAL FEATURE CLASS
'  MakeNorthArrow - MAKES CUSTOM JENNESSENT NORTH ARROW AS GRAPHIC IN LAYOUT
'  MakeRandomNormal - GIVEN MEAN, SD, PLACEHOLDER VARIABLE, OPTIONAL SECOND PLACEHOLDER, REPLACES PLACEHOLDERS WITH RANDOM NUMBERS
'                   FROM NORMAL DISTRIBUTION WITH SPECIFIED MEAN AND STANDARD DEVIATION
'                   USES BOX-MULLER TRANSFORMATION
'  MakeRandomNormalPolar - GIVEN MEAN, SD, PLACEHOLDER VARIABLE, OPTIONAL SECOND PLACEHOLDER, REPLACES PLACEHOLDERS WITH RANDOM NUMBERS
'                   FROM NORMAL DISTRIBUTION WITH SPECIFIED MEAN AND STANDARD DEVIATION
'                   Marsaglia's variation of the Box-Muller Transformation (http://en.wikipedia.org/wiki/Box-Muller_transform)
'  MakeUniquedBASEName - GIVEN A FULL FILE NAME, RETURNS A UNIQUE FILENAME IN THAT FOLDER BY APPENDING COUNTER TO ORIGINAL NAME
'  MakeUniqueGDBFeatureClassName - GIVEN IWorkspace AND NAME STRING, RETURNS A NAME STRING THAT REFLECTS A UNIQUE NAME IN THAT GDB
'  MakeUniqueGDBTableName - GIVEN IWorkspace AND NAME STRING, RETURNS A NAME STRING THAT REFLECTS A UNIQUE TABLE NAME IN THAT GDB
'  MakeUniqueRasterName - GIVEN JenDatasetType, WORKSPACE PATH STRING AND RASTER NAME, RETURNS A UNIQUE RASTER NAME FOR THAT WORKSPACE
'                   OPTIONALLY WILL TRIM TO GRID RESTRICTIONS FOR FILE-BASED RASTER
'  MakeUniqueRasterName2 - GIVEN IWorkspace, RASTER NAME AND OPTIONAL BOOLEAN TO TRIM TO 13 CHARACTERS, RETURNS A UNIQUE RASTER NAME FOR THAT WORKSPACE
'  MakeUniqueShapeFilename - GIVEN A FULL FILE NAME, RETURNS A UNIQUE FILENAME IN THAT FOLDER BY APPENDING COUNTER TO ORIGINAL NAME
'  MakeUniqueShapeFilename2 - GIVEN A WORKSPACE AND FEATURE CLASS NAME, RETURNS A UNIQUE FEATURE CLASS NAME IN THAT WORKSPACE.
'  MakeUniqueDataFrameName - GIVEN A SUGGESTED MAP NAME, APPENDS AN INCREASING NUMBER TO THE END UNTIL IT FINDS A NAME THAT DOES NOT EXIST.
'  Move_Element - MOVES AN ELEMENT ACCORDING TO THE DISTANCE AND DIRECTION BETWEEN TWO REPRESENTATIVE POINTS.  IF THE TWO POINTS
'           ARE SEPARATED BY X HORIZONTAL UNITS AND Y VERTICAL UNITS, THEN THIS FUNCTION WILL MOVE THE GRAPHIC UP Y UNITS AND OVER X UNITS.
'  Move_Geometry - MOVES A GEOMETRY ACCORDING TO THE DISTANCE AND DIRECTION BETWEEN TWO REPRESENTATIVE POINTS.  IF THE TWO POINTS
'           ARE SEPARATED BY X HORIZONTAL UNITS AND Y VERTICAL UNITS, THEN THIS FUNCTION WILL MOVE THE GEOMETRY UP Y UNITS AND OVER X UNITS.
'  OpenFile - GIVEN A DOCUMENT FILENAME AND PATH, OPENS THAT FILE USING THE REGISTERED WINDOWS PROGRAM
'  ReadTextFile - GIVEN A FILENAME, RETURNS A STRING CONTAINING ALL THE TEXT IN THE FILE
'  ReadFile2 - READS TEXT FILE INTO A BYTE ARRAY; RETURNS BYTE ARRAY
'  RemoveKeyFromCollection - GIVEN A COLLECTION AND KEY, REMOVE THE KEY WITHOUT CRASHING CODE IF COLLECTION DOES NOT HAVE THAT KEY
'  ReplaceBadChars - GIVEN A NAME, REPLACES ALL NON-STANDARD CHARACTERS WITH "_" CHARACTERS
'  ReturnAcceptableFieldName - GIVEN SAMPLE FIELD NAME, FORMAT TYPE AND LIST, RETURNS A NAME THAT IS VALID FOR THAT FORMAT AND DOES NOT EXIST IN THE LIST.
'                         LIST CAN BE AN IFields, IStringArray or IVarArray (CONTAINING EITHER STRINGS OR FIELDS)
'  ReturnAcceptableFieldName2 - SAME AS ABOVE, BUT INCLUDES FILE GEODATABASE FIELDS
'  ReturnCurrentMapUnits - GIVE IT AN IMAP, IT RETURNS THE NAME OF THE MAP UNITS
'  ReturnDatasetNamesOrNothing - RETURNS IEnumDatasetName, and returns NOTHING (i.e. will not crash) if no datasets of specified type exist
'  ReturnDatasetTypeName - GIVEN A TYPE (LONG; pDataset.Type), RETURNS THE NAME OF THE TYPE BASED ON ESRI HELP FILES.
'  ReturnDateValFromRow - GIVEN ROW AND INDEX, OPTIONAL BOOLEAN FOR WHETHER IT IS NULL, RETURNS A DATE VALUE
'  ReturnDBASEFieldName - GIVEN SAMPLE FIELD NAME AND LIST, RETURNS A NAME THAT IS VALID FOR DBASE AND DOES NOT EXIST IN THE LIST.
'                         LIST CAN BE AN IFields, IStringArray or IVarArray (CONTAINING EITHER STRINGS OR FIELDS)
'  ReturnDecimalPrecision - GIVEN A NUMBER, RETURNS THE NUMBER OF DECIMAL VALUES INCLUDED IN THE NUMBER.  WILL AUTOMATICALLY TRIM OFF
'                           TRAILING ZEROS, SUCH THAT A VALUE 0F 5.1240 RETURNS "3", A VALUE OF 5.124 RETURNS "3", AND A VALUE OF 5.000 RETURNS "0"
'  ReturnDistanceUnitsName - GIVEN AN esriUnits, RETURNS THE NAME
'  ReturnDoubleValFromRow - GIVEN ROW AND INDEX, OPTIONAL BOOLEAN FOR WHETHER IT IS NULL, RETURNS A DOUBLE VALUE
'  ReturnEsriFieldTypeNameFromNumber - GIVEN AN esriFieldType, RETURNS THE NAME
'  ReturnEsriFieldTypeNameFromNumber_Friendly - GIVEN AN esriFieldType, RETURNS A USER-FRIENDLY NAME
'  ReturnEmptyFClassWithSameSchema - GIVEN A FEATURE CLASS AND WORKSPACE (OR NOTHING IF WANTING IN-MEMORY), RETURNS A NEW FEATURE CLASS WITH SAME FIELDS
'  ReturnFeatureLayersFromGeoDatabase - GIVEN GEODATABASE WORKSPACE AND FEATURE CLASS TYPE, RETURNS VARIANT ARRAY OF FEATURE LAYERS
'  ReturnFeatureClassOrNothing - GIVEN A FEATURE WORKSPACE AND A NAME, RETURNS THE FEATURE CLASS IF ONE EXISTS WITH THAT NAME OR NOTHING
'  ReturnFieldsByType - GIVEN pFields AND JenFieldType ENUMERATION, RETURNS AN IVarARRAY OF FIELDS
'  ReturnFieldsByType2 - SAME AS ReturnFieldsByType, BUT USES BITWISE "AND" OPERATOR TO CHECK SELECTED OPTIONS SO IT IS MORE EFFICIENT.
'  ReturnFilesFromNestedFolders - GIVEN FOLDER PATH AND EXTENSION, RETURNS STRING ARRAY OF FILE PATHS WITH THIS EXTENSION
'  ReturnFilesFromNestedFolders - GIVEN FOLDER PATH AND AND ANY STRING IN THE NAME, RETURNS STRING ARRAY OF FILE PATHS WHERE THE FILENAME
'                                 CONTAINS THIS STRING
'  ReturnFoldersFromNestedFolders - GIVEN FOLDER PATH AND PARTIAL TEXT, RETURNS STRING ARRAY OF FILE PATHS CONTAINING FOLDERS WITH THIS TEXT
'  ReturnGraphicsByName - GIVEN A MAP DOCUMENT AND NAME, RETURNS A COLLECTION CONTAINING THE GEOMETRIES OF THOSE GRAPHICS
'  ReturnGraphicsByNameFromLayout - GIVEN A MAP DOCUMENT AND NAME, RETURNS A COLLECTION CONTAINING EITHER THE GEOMETRIES OR THE ELEMENTS
'  ReturnGraphicsByType - GIVEN A MAP DOCUMENT AND SHAPE TYPE, RETURNS A COLLECTION CONTAINING THE GEOMETRIES OF THOSE GRAPHICS
'  ReturnHistStatsFromDouble - MODIFICATION OF BasicStatsFromArraySimple
'           GIVEN anArray, optional directive to calculate standard deviation info, optional boolean for sample / population,
'                   optional boolean to return sorted array of numbers
'           RETURNS pArray, containing 4 sub-arrays: 1) Double Array of unique values, 2) Long array of counts of values,
'                   Double Array of standard stats (Sum, Mean, Minimum, Maximum, Range, Count, StDev, Variance, Median, Standard Error of Mean)
'                   Optional array of sorted numbers
'  ReturnLayersByType - GIVEN FOCUSMAP AND TYPE, RETURNS IVariantArray OF LAYERS
'  ReturnLayersByType2 - SAME AS ReturnLayersByType, BUT WITH OPTION TO RETURN INVALID LAYERS
'  ReturnLayerByName - GIVEN NAME AND MAP, RETURNS pLayer OR NOTHING
'  ReturnLongValFromRow - GIVEN ROW AND INDEX, OPTIONAL BOOLEAN FOR WHETHER IT IS NULL, RETURNS A LONG VALUE
'  ReturnMapByName - GIVEN NAME AND IMxDocument, RETURNS pMap OR NOTHING
'  ReturnNorthSolidFill - USED BY MakeCompassRose; RETURNS A SOLID FILL SYMBOL
'  ReturnNorthSolidLineSymbol - USED BY MakeCompassRose; RETURNS A SOLID LINE SYMBOL
'  ReturnRasterDatasetOrNothing - GIVEN A RASTER WORKSPACE AND A NAME, RETURNS THE RASTER DATASET IF ONE EXISTS WITH THAT NAME OR NOTHING

'  ReturnSelectedLayers - GIVEN MXDOC, RETURNS IARRAY OF SELECTED LAYERS.  ARRAY IS EMPTY IF NO LAYERS SELECTED.
'  ReturnShapefileLayersFromNestedFolders - GIVEN FOLDER PATH AND FEATURE CLASS TYPE, RETURNS VARIANT ARRAY OF FEATURE LAYERS
'  ReturnShapeTypeFromGeomType - GIVEN AN esriGeometryType, RETURNS A 2-VALUE STRING CONTAINING
'                   THE SINGULAR AND PLURAL VERSIONS OF THE GEOMETRY TYPE NAME
'  ReturnShapeTypeNameFromObject - GIVEN A FEATURE CLASS OR GEOMETRY, RETURNS A 2-VALUE STRING CONTAINING
'                   THE SINGULAR AND PLURAL VERSIONS OF THE GEOMETRY TYPE NAME
'  ReturnStateAbbreviation - GIVEN A STATE NAME (ONLY UNITED STATES), RETURNS THE 2-LETTER ABBREVIATION
'                   IF MANY STATES, THEN QUICKER TO CALL ReturnStateNameCollections
'  ReturnStateNameCollections - RETURNS TWO COLLECTIONS; EACH INDEX EITHER NAME OR ABBREVIATION, AND RETURNS THE OPPOSITE
'  ReturnStringValFromRow - GIVEN ROW AND INDEX, OPTIONAL BOOLEAN FOR WHETHER IT IS NULL, RETURNS A STRING VALUE
'  ReturnStringArrayOfNames - GIVEN IEnumDatasetName, RETURNS STRING ARRAY OF NAMES AND COUNT OF NAMES
'  ReturnTableByName - GIVEN NAME AND MAP, RETURNS pStandaloneTable OR NOTHING
'  ReturnTableOrNothing - GIVEN A FEATURE WORKSPACE AND A NAME, RETURNS THE TABLE IF ONE EXISTS WITH THAT NAME OR NOTHING
'  ReturnTempRasterWorkspace - RETURNS A RASTER WORKSPACE OBJECT LOCATED IN CURRENT TEMP DIRECTORY
'  ReturnTimeStamp - RETURNS STRING FORMATTED AS "YYYYMMDD_HHMMSS", OR SOMETHING LIKE "20081314_083230"
'  ReturnTimeElapsed - GIVEN A START AND END TIME, RETURNS A STRING DESCRIBING THE TIME ELAPSED.
'           Analysis Began: Wednesday, July 26, 2006;  4:39:01 PM
'           Analysis Complete: Wednesday, July 27, 2006;  5:47:22 PM
'           Time Elapsed: 1 day, 1 hour, 8 minutes, 21 seconds...
'  ReturnTimeElapsedFromMilliseconds - GIVEN A MILLISECOND COUNT (FROM GetTickCount API FUNCTION), RETURNS A STRING DESCRIBING THE TIME ELAPSED.
'           Analysis Began: Wednesday, July 26, 2006;  4:39:01 PM
'           Analysis Complete: Wednesday, July 27, 2006;  5:47:22 PM
'           Time Elapsed: 1 day, 1 hour, 8 minutes, 21 seconds...
'  ReturnTimeElapsedRTF - GIVEN A START AND END TIME, RETURNS AN RTF-FORMATTED STRING DESCRIBING THE TIME ELAPSED.
'           Analysis Began: Wednesday, July 26, 2006;  4:39:01 PM
'           Analysis Complete: Wednesday, July 27, 2006;  5:47:22 PM
'           Time Elapsed: 1 day, 1 hour, 8 minutes, 21 seconds...
'  ReturnTitleCase - GIEN A STRING, RETURNS A STRING WITH ALL WORDS UPPERCASE AT FIRST LETTER, LOWERCASE ALL OTHER LETTERS
'  ReturnQuerySpecialCharacters - GIVEN A DATASET AND PLACEHOLDERS, RETURNS SPECIAL DATABASE CHARACTERS
'  ReturnValidFGDBFieldName - GIVEN A STRING, RETURNS A VALID FILE GEODATABASE FIELD NAME VERSION OF THAT STRING.
'  ReturnValidFGDBFieldName2 - SAME AS ABOVE, BUT CHECKS AGAINST LIST OF EXISTING FIELD NAMES
'  SetFeatureSymbols2 - given a layer and a path to a layer file, will apply symbology from layer file to current layer
'  SearchForTextInFolder_and_MakeReport - GIVEN A FOLDER, FILE EXTENSION AND SEARCH TEXT, WILL RETURN A LIST OF ALL FILES IN ALL SUBFOLDERS
'           WITHIN THAT FOLDER THAT CONTAIN THAT TEXT.
'  SetLegendBorderColors: CHANGES FILL SYMBOL BORDER COLOR TO THE SAME COLOR AS SYMBOL.
'  SpacesInFrontOfText:  GIVEN A TEXT STRING AND TOTAL REQUESTED LENGTH, INSERTS SPACES IN FRONT OF ORIGINAL TEXT TO
'           FILL LENGTH.  USEFUL FOR FORMATTING COLUMNS OF NUMBERS IN FIXED-WIDTH-TEXT REPORTS.
'  StringValueInStringArray: RETURNS TRUE OR FALSE DEPENDING IF STRING VALUE EXISTS IN ARRAY
'  TextSlice:  RETURNS TEXT BETWEEN lngStartIndex AND lngEndIndex INCLUSIVE
'  TextSlice2:  Same as TextSlice, but second parameter optional.  If left out, will return single character at lngStartIndex
'  TrimZerosAndDecimas - GIVEN A NUMERIC VARIANT, RETURNS STRING WILL ALL TRAILING ZEROS AND DECIMALS TRIMMED OFF.  OPTIONAL TO ADD COMMAS
'  TrueLayerCount - GIVEN AN IMAP, RETURNS TRUE LAYER COUNT INCLUDING GROUP LAYERS
'  WordifyHex - GIVEN A STRING OF HEX VALUES CONSTRUCTED BY HexifyWord, TURNS IT BACK INTO A READABLE WORD.
'  WriteCodeToDuplicateFields - SIMPLIFIES CODE TO CREATE FIELDS THAT DUPLICATE AN EXISTING TABLE OR FEATURE CLASS, AND ALSO GENERATES
'           INDEX VARIABLES TO IDENTIFY FIELDS
'  WriteTextFile - Writes a string to a specified text file
'  WriteTextFile_SkipError - Special case of 'WriteTextFile' in which it does not write an error.  Intended to be used when ErrorHandling
'           module writes a text file which crashes, causing an infinite loop.

   


Public Declare Function CopyFile Lib "Kernel32" Alias "CopyFileA" _
  (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, _
  ByVal bFailIfExists As Long) As Long
  

Public Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function ShellExecute Lib "shell32.dll" Alias _
           "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, _
           ByVal lpFile As String, ByVal lpParameters As String, _
           ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal _
  hWndInsertAfter As Long, ByVal x As Long, ByVal Y As _
  Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags _
  As Long) As Long

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
  (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Const GWL_HWNDPARENT = (-8)
Private Const INVALID_HANDLE_VALUE As Long = -1
Private Const MAX_PATH As Long = 260

Declare Function GetDesktopWindow Lib "user32.dll" () As Long
Public Declare Function GetTickCount Lib "Kernel32" () As Long

Private Type SECURITY_ATTRIBUTES
   nLength As Long
   lpSecurityDescriptor As Long
   bInheritHandle As Long
End Type

Private Declare Function CreateDirectory Lib "Kernel32" _
   Alias "CreateDirectoryA" _
  (ByVal lpPathName As String, _
   lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long

Public Enum JenNorthArrowColors
  ENUM_Pale_Turquoise = 15658671     ' RGB(175, 238, 238)
  ENUM_Dark_Turquoise = 13749760     ' RGB(0, 206, 209)
  ENUM_Medium_Turquoise = 13422920   ' RGB(72, 209, 204)
  ENUM_Turquoise = 13688896          ' RGB(64, 224, 208)
  ENUM_Black = 0                     ' RGB(0, 0, 0)
  ENUM_Dark_Gray = 1644825           ' RGB(25, 25, 25)
  ENUM_Yellow = 65535                ' RGB(255, 255, 0)
  ENUM_white = 16777215              ' RGB(255, 255, 255)
  ENUM_gray1 = 4276545               ' RGB(65, 65, 65)
  ENUM_gray2 = 8224125               ' RGB(125, 125, 125)
  ENUM_gray3 = 12500670              ' RGB(190, 190, 190)
  ENUM_GrayVeryLight = 15132390      ' RGB(230, 230, 230)
  ENUM_GanadoRed = 2628477           ' RGB(125, 27, 40)
End Enum
   
Public Enum JenFieldTypes
  enum_FieldString = 1
  enum_FieldNumber = 2
  enum_FieldDate = 4
  enum_FieldOID = 8
  ENUM_FieldGeometry = 16
  ENUM_FieldBlob = 32
  ENUM_FieldRaster = 64
  ENUM_FieldGUID = 128
  ENUM_FieldGlobalID = 256
  enum_FieldXML = 512
End Enum

Public Enum JenLayerTypes
  ENUM_jenFeatureLayers = 1
  ENUM_jenRasterLayers = 2
  ENUM_jenStandaloneTables = 4
  ENUM_jenPointLayers = 8
  ENUM_jenPolylineLayers = 16
  ENUM_jenPolygonLayers = 32
  ENUM_jenMultipointLayers = 64
  ENUM_jenTinLayers = 128
  ENUM_jenRastCatalogLayers = 256
End Enum

Public Enum JenDatasetTypes
  ENUM_Shapefile = 1
  ENUM_FileGDB = 2
  ENUM_PersonalGDB = 4
  ENUM_Coverage = 8
  ENUM_SDC_FeatureClass = 16
  ENUM_File_Raster = 32
End Enum

Public Enum JenBulletTypes
  ENUM_Letter_Lowercase = 1
  ENUM_Letter_Uppercase = 2
  ENUM_RomanNumeral_Lowercase = 4
  ENUM_RomanNumeral_Uppercase = 8
End Enum

Public Enum Jen_ElementEnvPoint
  ENUM_Upper_Left = 1
  ENUM_Upper_Right = 2
  ENUM_Lower_Left = 3
  ENUM_Lower_Right = 4
  ENUM_Upper_Center = 5
  ENUM_Lower_Center = 6
  ENUM_Center_Left = 7
  ENUM_Center_Right = 8
  ENUM_Center_Center = 9
  ENUM_By_Percentages = 10
End Enum

Public Enum Jen_AlignColorDialogOption
  ENUM_AlignBeneathRectangle = 1
  ENUM_AlignToRightOfRectangle = 2
End Enum

Const dblPI As Double = 3.14159265358979
Public Sub BasicStatsFromArraySimpleFast(dblSortedArray() As Double, _
    Optional booCalculateVariance As Boolean = False, _
    Optional theSum As Double, Optional theCount As Long, Optional theMinimum As Double, _
    Optional theMaximum As Double, Optional theMean As Double, Optional theMedian As Double, _
    Optional theVariance As Double, Optional theStDev As Double, Optional theStErrMean As Double, _
    Optional theRange As Double)
  
  ' ASSUMES ARRAY IS SORTED!!!! --------------------------------------
  
'  Dim pResponse As esriSystem.IDoubleArray
'  Set pResponse = New esriSystem.DoubleArray
  
  Dim anIndex As Long
  Dim theVal As Double
  
  theSum = 0
  theCount = UBound(dblSortedArray) + 1         ' ARRAY INDEX STARTS AT 0
  theMinimum = dblSortedArray(0)
  theMaximum = dblSortedArray(0)
  
  '  PASS 1:  MINIMUM, MAXIMUM AND SUM --------------------------------------------------
  For anIndex = LBound(dblSortedArray) To UBound(dblSortedArray)
    
    theVal = dblSortedArray(anIndex)
    
    If theVal < theMinimum Then
      theMinimum = theVal
    End If
    If theVal > theMaximum Then
      theMaximum = theVal
    End If
    theSum = theSum + theVal
    
  Next anIndex
  
  theMean = theSum / theCount
  
  Dim theSumSqDev As Double
  Dim theSqDev As Double
  
  '  PASS 2:  STANDARD DEVIATION AND VARIANCE --------------------------------------------------
  If booCalculateVariance Then
    For anIndex = LBound(dblSortedArray) To UBound(dblSortedArray)
                
      theVal = dblSortedArray(anIndex)
      theSqDev = (theVal - theMean) * (theVal - theMean)
      theSumSqDev = theSqDev + theSumSqDev
      
    Next anIndex
  Else
    theSqDev = 0
    theSumSqDev = 0
  End If
    
  If theCount > 1 Then
    If theCount Mod 2 = 0 Then      ' EVEN NUMBER
      theMedian = (dblSortedArray((theCount / 2) - 1) + dblSortedArray(theCount / 2)) / 2
    Else
      theMedian = dblSortedArray((theCount - 1) / 2)
    End If
    
    theVariance = theSumSqDev / (theCount - 1)
    theStDev = Sqr(theVariance)
    theStErrMean = theStDev / (Sqr(theCount))
    
  Else
    If theCount = 1 Then
      theMedian = dblSortedArray(0)
    Else
      theMedian = -9999
    End If
    theVariance = 0
    theStDev = 0
    theStErrMean = 0
  End If

  theRange = theMaximum - theMinimum

  ' OUTPUT ARRAY; VARIANT BECAUSE OF MODE STRING
  '(0) = SUM
  '(1) = MEAN
  '(2) = MINIMUM
  '(3) = MAXIMUM
  '(4) = RANGE
  '(5) = COUNT
  '(6) = STANDARD DEVIATION
  '(7) = VARIANCE
  '(8) = MEDIAN
  '(9) = STANDARD ERROR OF MEAN
  
'  pResponse.Add theSum
'  pResponse.Add theMean
'  pResponse.Add theMinimum
'  pResponse.Add theMaximum
'  pResponse.Add theRange
'  pResponse.Add theCount
'  pResponse.Add theStDev
'  pResponse.Add theVariance
'  pResponse.Add theMedian
'  pResponse.Add theStErrMean
'
'  Set BasicStatsFromArraySimple = pResponse

End Sub
Public Sub BasicStatsFromArraySimpleFast2(dblSortedArray() As Double, _
    Optional booCalculateVariance As Boolean = False, _
    Optional theSum As Double, Optional lngCount As Long, Optional theMinimum As Double, _
    Optional theMaximum As Double, Optional theMean As Double, Optional theMedian As Double, _
    Optional theVariance As Double, Optional theStDev As Double, Optional theStErrMean As Double, _
    Optional theRange As Double, Optional booNeedToSort As Boolean = True, _
    Optional booCalculatePercentiles As Boolean = False, _
    Optional dblPercentileLevelsToCalculate As Variant, Optional dblPercentiles As Variant)
  
  ' NO LONGER ASSUMES ARRAY IS SORTED!!!! --------------------------------------
  If booNeedToSort Then QuickSort.DoubleAscending dblSortedArray, 0, UBound(dblSortedArray)
  
'  Dim pResponse As esriSystem.IDoubleArray
'  Set pResponse = New esriSystem.DoubleArray
  
  Dim anIndex As Long
  Dim theVal As Double
  
  theSum = 0
  lngCount = UBound(dblSortedArray) + 1         ' ARRAY INDEX STARTS AT 0
  theMinimum = dblSortedArray(0)
  theMaximum = dblSortedArray(0)
  
  '  PASS 1:  MINIMUM, MAXIMUM AND SUM --------------------------------------------------
  For anIndex = LBound(dblSortedArray) To UBound(dblSortedArray)
    
    theVal = dblSortedArray(anIndex)
    
    If theVal < theMinimum Then
      theMinimum = theVal
    End If
    If theVal > theMaximum Then
      theMaximum = theVal
    End If
    theSum = theSum + theVal
    
  Next anIndex
  
  theMean = theSum / lngCount
  
  Dim theSumSqDev As Double
  Dim theSqDev As Double
  
  '  PASS 2:  STANDARD DEVIATION AND VARIANCE --------------------------------------------------
  If booCalculateVariance Then
    For anIndex = LBound(dblSortedArray) To UBound(dblSortedArray)
                
      theVal = dblSortedArray(anIndex)
      theSqDev = (theVal - theMean) * (theVal - theMean)
      theSumSqDev = theSqDev + theSumSqDev
      
    Next anIndex
  Else
    theSqDev = 0
    theSumSqDev = 0
  End If
  
  Dim dblM As Double
  Dim dblP As Double
  ' R: QUANTILE TYPE 7
  dblM = 1 - dblP
  ' R: QUANTILE TYPE 8
  dblM = (dblP + 1) / 3
  Dim dblG As Double
  
  Dim lngType As Long
  lngType = 8
  
  ' PERCENTILE ALGORITHM ADAPTED FROM http://en.wikipedia.org/wiki/Percentile
  Dim dblPercentileLevels() As Double
  Dim dblReturnPercentiles() As Double
  Dim lngPercIndex As Long
  Dim dblValIndex As Double
  Dim lngFloor As Long
  Dim lngCeiling As Long
  Dim dblFloorVal As Double
  Dim dblFloorPercentile As Double
  Dim dblCeilingVal As Double
  Dim dblCeilingPercentile As Double
  
  Dim dblPercentile As Double
  If booCalculatePercentiles Then
    dblPercentileLevels = dblPercentileLevelsToCalculate
    ReDim dblReturnPercentiles(UBound(dblPercentileLevels))
  End If
    
  If lngCount > 1 Then
    If lngCount Mod 2 = 0 Then      ' EVEN NUMBER
      theMedian = (dblSortedArray((lngCount / 2) - 1) + dblSortedArray(lngCount / 2)) / 2
    Else
      theMedian = dblSortedArray((lngCount - 1) / 2)
    End If
    
    ' PERCENTILES
    If booCalculatePercentiles Then
      For lngPercIndex = 0 To UBound(dblPercentileLevelsToCalculate)
        dblPercentile = dblPercentileLevelsToCalculate(lngPercIndex)
        
        If lngType = 5 Then
          dblM = 0.5
        ElseIf lngType = 7 Then
          dblM = 1 - dblPercentile
        ElseIf lngType = 8 Then
          dblM = (dblPercentile + 1) / 3
        End If
        
        dblValIndex = (CDbl(lngCount) * dblPercentile) + dblM
        lngFloor = Int(dblValIndex)
        dblG = dblValIndex - CDbl(lngFloor)
        If dblG = 0 Then
          dblReturnPercentiles(lngPercIndex) = dblSortedArray(lngFloor - 1)
        Else
          lngCeiling = lngFloor + 1
          dblFloorVal = dblSortedArray(lngFloor - 1)
          dblCeilingVal = dblSortedArray(lngCeiling - 1)
          
          If lngType = 5 Then
            dblFloorPercentile = (lngFloor - 0.5) / lngCount
            dblCeilingPercentile = (lngCeiling - 0.5) / lngCount
          ElseIf lngType = 7 Then
            dblFloorPercentile = (lngFloor - 1) / (lngCount - 1)
            dblCeilingPercentile = (lngCeiling - 1) / (lngCount - 1)
          ElseIf lngType = 8 Then
            dblFloorPercentile = (lngFloor - (1 / 3)) / (lngCount + (1 / 3))
            dblCeilingPercentile = (lngCeiling - (1 / 3)) / (lngCount + (1 / 3))
          End If
          
'          dblReturnPercentiles(lngPercIndex) = dblFloorVal + (lngCount * (dblPercentile - dblFloorPercentile) * _
              (dblCeilingVal - dblFloorVal))
          ' INTERPOLATION = Beginning Value + ((Range between values) * (percentage between values))
'          dblReturnPercentiles(lngPercIndex) = dblFloorVal + (lngCount * (dblPercentile - dblFloorPercentile) * _
              (dblCeilingVal - dblFloorVal))
          dblReturnPercentiles(lngPercIndex) = dblFloorVal + _
              (((dblPercentile - dblFloorPercentile) / (dblCeilingPercentile - dblFloorPercentile)) * _
              (dblCeilingVal - dblFloorVal))
              
        End If
      Next lngPercIndex
      dblPercentiles = dblReturnPercentiles
    End If
    
    theVariance = theSumSqDev / (lngCount - 1)
    theStDev = Sqr(theVariance)
    theStErrMean = theStDev / (Sqr(lngCount))
    
  Else
    If lngCount = 1 Then
'      dblQuartile1 = dblSortedArray(0)
      theMedian = dblSortedArray(0)
'      dblQuartile3 = dblSortedArray(0)
      If booCalculatePercentiles Then
        For lngPercIndex = 0 To UBound(dblPercentileLevelsToCalculate)
          dblReturnPercentiles(lngPercIndex) = dblSortedArray(0)
        Next lngPercIndex
        dblPercentiles = dblReturnPercentiles
      End If
    Else
'      dblQuartile1 = -9999
      theMedian = -9999
'      dblQuartile3 = -9999
      If booCalculatePercentiles Then
        For lngPercIndex = 0 To UBound(dblPercentileLevelsToCalculate)
          dblReturnPercentiles(lngPercIndex) = -9999
        Next lngPercIndex
        dblPercentiles = dblReturnPercentiles
      End If
    End If
    theVariance = 0
    theStDev = 0
    theStErrMean = 0
  End If

  theRange = theMaximum - theMinimum

  ' OUTPUT ARRAY; VARIANT BECAUSE OF MODE STRING
  '(0) = SUM
  '(1) = MEAN
  '(2) = MINIMUM
  '(3) = MAXIMUM
  '(4) = RANGE
  '(5) = COUNT
  '(6) = STANDARD DEVIATION
  '(7) = VARIANCE
  '(8) = MEDIAN
  '(9) = STANDARD ERROR OF MEAN
  
'  pResponse.Add theSum
'  pResponse.Add theMean
'  pResponse.Add theMinimum
'  pResponse.Add theMaximum
'  pResponse.Add theRange
'  pResponse.Add lngCount
'  pResponse.Add theStDev
'  pResponse.Add theVariance
'  pResponse.Add theMedian
'  pResponse.Add theStErrMean
'
'  Set BasicStatsFromArraySimple = pResponse


  GoTo ClearMemory
ClearMemory:
  Erase dblPercentileLevels
  Erase dblReturnPercentiles
End Sub
Public Function TrimZerosAndDecimals(varNumber As Variant, Optional booAddCommas As Boolean = True, _
    Optional lngMaxNumDecimals As Long = 15) As String

  Dim dblValue As Double
  dblValue = CDbl(varNumber)
  Dim strValue As String
  Dim strZeros As String
  strZeros = String(lngMaxNumDecimals, "0")
  
  If booAddCommas Then
    strValue = CStr(Format(dblValue, "#,##0." & strZeros))
  Else
    strValue = CStr(Format(dblValue, "0." & strZeros))
  End If
  
  Do Until Right(strValue, 1) <> "0" Or Len(strValue) = 0
    strValue = Left(strValue, Len(strValue) - 1)
  Loop
  
  Do Until Right(strValue, 1) <> "." Or Len(strValue) = 0
    strValue = Left(strValue, Len(strValue) - 1)
  Loop
  
  If strValue = "" Then
    strValue = "0"
  End If
  
  TrimZerosAndDecimals = strValue

End Function


Public Function MakeFieldNameVarArray(pOrigFClass As IFeatureClass, pNewFClass As IFeatureClass, _
      pFieldArray As esriSystem.IVariantArray, booFieldArrayContainsNamePairs As Boolean) As esriSystem.IVariantArray
      
  ' MAKE IVariantArray OF FIELD NAME INDICES
  ' STRUCTURE:  ARRAY OF 4-ITEM VARIANT ARRAYS
  '             ITEMS = 0) ORIGINAL FCLASS FIELD NAME
  '                     1) ORIGINAL FCLASS FIELD INDEX
  '                     2) NEW FCLASS FIELD NAME
  '                     3) NEW FCLASS FIELD INDEX
  
  Dim lngIndex As Long
  Dim pField As iField
  Dim strOrigName As String
  Dim strNewName As String
  Dim pNamePair As esriSystem.IStringArray
  Dim pFieldData As esriSystem.IVariantArray
  Set MakeFieldNameVarArray = New varArray
  
  For lngIndex = 0 To pFieldArray.Count - 1
    If booFieldArrayContainsNamePairs Then
      Set pNamePair = pFieldArray.Element(lngIndex)
      strOrigName = pNamePair.Element(0)
      strNewName = pNamePair.Element(1)
    Else
      Set pField = pFieldArray.Element(lngIndex)
      strOrigName = pField.Name
      strNewName = pField.Name
    End If
    Set pFieldData = New esriSystem.varArray
    pFieldData.Add strOrigName
    pFieldData.Add pOrigFClass.FindField(strOrigName)
    pFieldData.Add strNewName
    pFieldData.Add pNewFClass.FindField(strNewName)
    MakeFieldNameVarArray.Add pFieldData
  Next lngIndex

  GoTo ClearMemory

ClearMemory:
  Set pField = Nothing
  Set pNamePair = Nothing
  Set pFieldData = Nothing
End Function

Public Sub MakeRandomNormal(dblMean As Double, dblSD As Double, dblRand1 As Double, Optional dblRand2 As Double = -999)

  Static dblSeed
' Based on the Box-Muller Transformation:.
'          y1 = sqrt( - 2 ln(x1) ) cos( 2 pi x2 )
'          y2 = sqrt( - 2 ln(x1) ) sin( 2 pi x2 )
'        where:
'          x1 = first uniform random number (between 0 and 1)
'          x2 = second uniform random number (between 0 and 1)
'          y1 = first normally distributed random number
'          y2 = second normally distributed random number.
  If dblSeed = 0 Then
    Randomize            ' USE SYSTEM TIMER
  Else
    Randomize dblSeed
  End If
  Dim dblX1 As Double
  Dim dblX2 As Double
  dblX1 = Rnd
  dblX2 = Rnd
  
  dblRand1 = (Sqr(-2 * Log(dblX1)) * Cos(2 * dblPI * dblX2) * dblSD#) + dblMean#
  dblRand2 = (Sqr(-2 * Log(dblX1)) * Sin(2 * dblPI * dblX2) * dblSD#) + dblMean#
  
  dblSeed = dblRand2

End Sub

Public Function ReturnHistStatsFromDouble(pDouble As esriSystem.IDoubleArray, _
            Optional pCounts As esriSystem.IDoubleArray, _
            Optional booVariance As Boolean = False, _
            Optional booIsSample As Boolean = False, _
            Optional booReturnSortedArray As Boolean = False, _
            Optional strName As String = "Statistics") As esriSystem.IVariantArray
    
  Dim dblVals() As Double
  Dim lngIndex As Long
  Dim lngCount As Long
  Dim dblWeight As Double
  
  Dim booHasWeight As Boolean
  booHasWeight = Not pCounts Is Nothing
  
  lngCount = pDouble.Count - 1
  ReDim dblVals(lngCount)
  
  If booHasWeight Then
    Dim dblWeightArray() As Double
    ReDim dblWeightArray(lngCount)
    For lngIndex = 0 To lngCount
      dblVals(lngIndex) = pDouble.Element(lngIndex)
      dblWeightArray(lngIndex) = pCounts.Element(lngIndex)
    Next lngIndex
    
    ' SORT VALUES
    QuickSort.DoubleAscendingWithSizes dblVals, dblWeightArray, 0, lngCount
  Else
    dblWeight = 1
    For lngIndex = 0 To lngCount
      dblVals(lngIndex) = pDouble.Element(lngIndex)
    Next lngIndex
    
    ' SORT VALUES
    QuickSort.DoubleAscending dblVals, 0, lngCount
  End If
    
  ' ASSUMES ARRAY IS SORTED!!!! --------------------------------------
  Dim theSum As Double
  Dim theCount As Double
  Dim theMinimum As Double
  Dim theMaximum As Double
  Dim theMean As Double
  Dim theSumSqDev As Double
  Dim theSqDev As Double
  Dim theMedian As Double
  Dim theVariance As Double
  Dim theStDev As Double
  Dim theStErrMean As Double
  Dim theRange As Double
  Dim dblSumSquaredWeights As Double
  Dim dblCheckVal As Double
  Dim dblHalfCount As Double         ' FOR MEDIAN WHEN USING WEIGHTED VALUES
  Dim booFoundMedian As Boolean      ' FOR MEDIAN WHEN USING WEIGHTED VALUES
  Dim dblRunningCountTotal As Double ' FOR MEDIAN WHEN USING WEIGHTED VALUES
  
  booFoundMedian = False
  
  ' MAKE ARRAYS TO RETURN
  Dim pReturnArray As esriSystem.IDoubleArray
  Dim pCountArray As esriSystem.IDoubleArray  ' "Counts" might be <1 nor non-integer if they are feature lengths or areas
  Dim pResponse As esriSystem.IDoubleArray
  
  Set pReturnArray = New esriSystem.DoubleArray
  Set pCountArray = New esriSystem.DoubleArray
  Set pResponse = New esriSystem.DoubleArray
  
  Dim dblCounter As Long
  Dim dblTestVal As Double
  
  If lngCount = 0 Then
  
    pResponse.Add -999
    pResponse.Add -999
    pResponse.Add -999
    pResponse.Add -999
    pResponse.Add -999
    pResponse.Add -999
    pResponse.Add -999
    pResponse.Add -999
    pResponse.Add -999
    pResponse.Add -999
    pResponse.Add -999   ' FOR dblTrueCount
  
  Else
  
    dblTestVal = dblVals(0)
    If booHasWeight Then
      dblCounter = dblWeightArray(0)
    Else
      dblCounter = 1
    End If
    theMinimum = dblVals(0)
    theMaximum = dblVals(UBound(dblVals))
    theRange = theMaximum - theMinimum
    theSum = dblTestVal
    theCount = dblCounter
    dblSumSquaredWeights = dblCounter ^ 2
'     theCount = lngCount + 1         ' ARRAY INDEX STARTS AT 0
    
    If lngCount > 1 Then
    
      '  PASS 1:  MINIMUM, MAXIMUM HISTOGRAM ARRAYS AND SUM --------------------------------------------------
      For lngIndex = 1 To lngCount
        dblCheckVal = dblVals(lngIndex)
        If booHasWeight Then
          dblWeight = dblWeightArray(lngIndex)  ' OTHERWISE WEIGHT DEFAULTS TO 1
          dblSumSquaredWeights = dblSumSquaredWeights + (dblWeight ^ 2)
        End If
          
        ' BUILD HISTOGRAM ARRAYS
        If dblTestVal = dblCheckVal Then
          dblCounter = dblCounter + dblWeight
        Else
          pReturnArray.Add dblTestVal
          pCountArray.Add dblCounter
          dblTestVal = dblCheckVal
          dblCounter = dblWeight
        End If
        
        ' GENERAL STATS
        theSum = theSum + (dblCheckVal * dblWeight)
        theCount = theCount + dblWeight
      Next lngIndex
      
      ' FOR HISTOGRAM ARRAYS, ADD LAST VALUE
      pReturnArray.Add dblTestVal
      pCountArray.Add dblCounter
      
      ' FOR GENERAL STATS
      theMean = theSum / theCount
      dblHalfCount = theCount / 2
      dblRunningCountTotal = 0
      
      '  PASS 2:  STANDARD DEVIATION AND VARIANCE --------------------------------------------------
      If booVariance Then
        For lngIndex = 0 To lngCount
                    
          dblCheckVal = dblVals(lngIndex)
          
          If booHasWeight Then
            dblWeight = dblWeightArray(lngIndex)
            theSqDev = dblWeight * ((dblCheckVal - theMean) ^ 2)
            If Not booFoundMedian Then
              dblRunningCountTotal = dblRunningCountTotal + dblWeight
              If dblRunningCountTotal >= dblHalfCount Then
                booFoundMedian = True
                theMedian = dblCheckVal
              End If
            End If
          Else
            theSqDev = (dblCheckVal - theMean) ^ 2
          End If
          
          theSumSqDev = theSqDev + theSumSqDev
        Next lngIndex
      Else
        theSqDev = 0
        theSumSqDev = 0
      End If
      
      If booHasWeight Then
        ' THEN MEDIAN IS CALCULATED UP IN PASS 2
      Else
        If theCount Mod 2 = 0 Then      ' EVEN NUMBER
          theMedian = (dblVals((theCount / 2) - 1) + dblVals(theCount / 2)) / 2
        Else
          theMedian = dblVals((theCount - 1) / 2)
        End If
      End If
      
      If booHasWeight Then
        If booIsSample Then
          ' FROM http://en.wikipedia.org/wiki/Weighted_standard_deviation#Weighted_sample_variance
          ' FROM http://pygsl.sourceforge.net/reference/pygsl/node36.html
          theVariance = (theCount / ((theCount ^ 2) - dblSumSquaredWeights)) * theSumSqDev
        Else
          ' FROM http://en.wikipedia.org/wiki/Weighted_standard_deviation#Weighted_sample_variance
          theVariance = theSumSqDev / theCount
        End If
      Else
        If booIsSample Then
          theVariance = theSumSqDev / (theCount - 1)
        Else
          theVariance = theSumSqDev / theCount
        End If
      End If
      
      theStDev = Sqr(theVariance)
      theStErrMean = theStDev / (Sqr(theCount))
      
    Else
      pReturnArray.Add dblTestVal
      pCountArray.Add 1
      
      theMean = dblTestVal
      theMedian = dblTestVal
      theVariance = 0
      theStDev = 0
      theStErrMean = 0
      
    End If
      
    pResponse.Add theSum
    pResponse.Add theMean
    pResponse.Add theMinimum
    pResponse.Add theMaximum
    pResponse.Add theRange
    pResponse.Add theCount
    pResponse.Add theStDev
    pResponse.Add theVariance
    pResponse.Add theMedian
    pResponse.Add theStErrMean
    
  End If
  
  Set ReturnHistStatsFromDouble = New esriSystem.varArray
  ReturnHistStatsFromDouble.Add pReturnArray
  ReturnHistStatsFromDouble.Add pCountArray
  ReturnHistStatsFromDouble.Add pResponse
  If booReturnSortedArray Then
    ReturnHistStatsFromDouble.Add dblVals
  Else
    ReturnHistStatsFromDouble.Add -999
  End If
  ReturnHistStatsFromDouble.Add strName


  GoTo ClearMemory
ClearMemory:
  Erase dblVals
  Erase dblWeightArray
  Set pReturnArray = Nothing
  Set pCountArray = Nothing
  Set pResponse = Nothing
End Function
Public Function ReplaceBadChars(strName As String, Optional booReplacePeriod As Boolean = False, _
    Optional booReplaceBackSlash As Boolean = False, Optional booReplaceSpace As Boolean = True, _
    Optional booStartWithLetter As Boolean = False) As String

  If strName = "" Then
    ReplaceBadChars = "z"
  Else
    Dim strAcceptable As String
    strAcceptable = "abcdefghijklmnopqrstuvwxyz0123456789_"
    If Not booReplacePeriod Then strAcceptable = strAcceptable & "."
    If Not booReplaceBackSlash Then strAcceptable = strAcceptable & "\"
    If Not booReplaceSpace Then strAcceptable = strAcceptable & " "
    
    Dim strOutput As String
    strOutput = strName
    Dim lngIndex As Long
    Dim strChar As String
    For lngIndex = 1 To Len(strOutput)
      strChar = Mid(strOutput, lngIndex, 1)
      If InStr(1, strAcceptable, strChar, vbTextCompare) = 0 Then
        strOutput = Replace(strOutput, strChar, "_")
      End If
    Next lngIndex
    If booStartWithLetter Then
      Dim strLetters As String
      strLetters = "abcdefghijklmnopqrstuvwxyz"
      If InStr(1, strLetters, Left(strOutput, 1), vbTextCompare) = 0 Then
        strOutput = "z" & strOutput
      End If
    End If
    ReplaceBadChars = strOutput
  End If

End Function

Public Function CreateGeneralTable(Optional strDir As String = "", Optional strName As String = "", _
    Optional lngCategory As JenDatasetTypes = ENUM_Shapefile, Optional pAddFields As esriSystem.IVariantArray, _
    Optional strGxDialogCaption As String = "Specify New Table Name:") As ITable

  ' IF booIsFClass = FALSE, THEN THIS FUNCTION WILL CREATE A TABLE INSTEAD
  
  ' THIS FUNCTION ONLY ALLOWS YOU TO WRITE TO A FEATURE DATASET IF IT FORCES USER TO SELECT FEATURE DATASET WITH IGxCatalog.
  ' IF YOU SEND A SPECIFIC PATHNAME AND FILENAME, THIS FUNCTION WILL NOT ALLOW YOU TO SAVE TO A FEATURE DATASET.
  
  Dim pGxDataset As IGxDataset
  Dim pDataset As IDataset
  Dim pWorkspace As IWorkspace
  Dim pFeatureDatasetName As IName
  Dim pFeatureDataset As IFeatureDataset
  Dim booHasFeatureDataset As Boolean
  Dim strFeatDataset As String
  Dim strCategory As String
  Dim strParent As String
  Dim strBaseName As String
  Dim pFeatWS As IFeatureWorkspace
  Dim pWSFact As IWorkspaceFactory2
  Dim pTableName As String
  Dim pTable As ITable
  Dim strWsPathname As String
  Dim pGxObj As IGxObject
  Dim pGxDialog As IGxDialog
  Dim pGxFCFilter As IGxObjectFilter
  Dim pFilterCol As IGxObjectFilterCollection

  ' CHECK IF TABLE ALREADY EXISTS
  Dim booAlreadyExists As Boolean
  booAlreadyExists = True
  Dim lngResponse As VbMsgBoxResult
  
  Do Until booAlreadyExists = False
     
    ' FIRST, CHECK IF DIRECTORY, NAME AND CATEGORY ARE SPECIFIED
    If strDir <> "" And strName <> "" Then
      
      Select Case lngCategory
        Case ENUM_Shapefile
          strCategory = "Shapefile"
          strName = aml_func_mod.SetExtension(strName, "dbf")
          If Right(strDir, 1) <> "\" Then strDir = strDir & "\"
        Case ENUM_FileGDB
          strCategory = "File Geodatabase Feature Class"
        Case ENUM_PersonalGDB
          strCategory = "Personal Geodatabase Feature Class"
        Case Else
          MsgBox "PROBLEM:  This function can only create dBASE Tables, File GeoDatabase Tables, or " & vbCrLf & _
              "Personal GeoDatabase Tables.  Bailing out...", vbOKOnly, "Invalid Output Type:"
          Set CreateGeneralTable = Nothing
          Exit Function
      End Select
      
      strWsPathname = strDir
      booHasFeatureDataset = False
      strBaseName = strName
      booHasFeatureDataset = False
      strFeatDataset = ""
    Else
    
 '     If pGxObj Is Nothing Or IsMissing(pGxObj) Then  ' IF NEED TO ASK USER TO SPECIFY GxObject
      
      Set pGxDialog = New GxDialog
      Set pGxFCFilter = New GxFilterTables
      
      Set pFilterCol = pGxDialog
      pFilterCol.AddFilter pGxFCFilter, True
      pGxDialog.Title = strGxDialogCaption
      pGxDialog.AllowMultiSelect = False
      pGxDialog.ButtonCaption = "Enter"
        
      If Not pGxDialog.DoModalSave(0) Then
          Set CreateGeneralTable = Nothing
          Exit Function 'Exit if user press Cancel
      End If
      
      Set pGxObj = pGxDialog.FinalLocation
      strWsPathname = pGxObj.FullName
      strBaseName = pGxDialog.Name
    
 '     End If
  
  '    Dim pGxCat As IGxCatalog
  '    Set pGxCat = pGxDialog.InternalCatalog
  
      ' CHECK TO SEE IF WE ARE WRITING TO A FEATURE DATASET
      If TypeOf pGxObj Is IGxDataset Then
        Set pGxDataset = pGxObj
          
        Set pDataset = pGxDataset.Dataset
        Set pWorkspace = pDataset.Workspace
        strWsPathname = pWorkspace.PathName
      
        Set pFeatureDatasetName = pDataset.FullName
        
        booHasFeatureDataset = TypeOf pDataset Is IFeatureDataset
        If booHasFeatureDataset Then
          Set pGxObj = pGxObj.Parent  ' BACK OUT IF FEATURE DATASET IF WE ARE IN ONE
        End If
      End If
      
      strCategory = pGxObj.Category
    End If
    
    ' NOW WE KNOW NAME, CATEGORY AND WORKSPACE.  NEXT CREATE WORKSPACE FACTORY
    Select Case strCategory
       Case "Shapefile", "Folder", "dBASE Table"
         Set pWSFact = New ShapefileWorkspaceFactory
         pTableName = strBaseName
       Case "File Geodatabase Feature Class", "File Geodatabase Raster Catalog", _
           "File Geodatabase", "File Geodatabase Feature Dataset", "File Geodatabase Table"
         Set pWSFact = New FileGDBWorkspaceFactory
         pTableName = strBaseName
       Case "Personal Geodatabase Feature Class", "Personal Geodatabase Raster Catalog", _
           "Personal Geodatabase", "Personal Geodatabase Feature Dataset", _
           "Personal Geodatabase Table"
         Set pWSFact = New AccessWorkspaceFactory
         pTableName = strBaseName
'       Case "Polygon Feature Class", "Arc Feature Class", "Point Feature Class", "Route Feature Class", _
'            "Tic Feature Class", "Region Feature Class", "Label Feature Class"
'         Set pWSFact = New ArcInfoWorkspaceFactory
'         If booHasFeatureDataset Then
'           pTableName = strFeatDataset & ":" & strBaseName
'         Else
'           pTableName = strBaseName
'         End If
'       Case "SDC Feature Class"
'         Set pWSFact = New SDCWorkspaceFactory
'         If booHasFeatureDataset Then
'           pTableName = strFeatDataset & ":" & strBaseName
'         Else
'           pTableName = strBaseName
'         End If
       Case Else
         MsgBox "Unexpected Workspace Type (" & strCategory & ")!  Need to write new code for this..."
         Set CreateGeneralTable = Nothing
         Exit Function
    End Select
    
  '  Debug.Print "Full Name = " & strPath & vbCrLf & "Name = " & strName & vbCrLf & _
        "Base Name = " & strBaseName & vbCrLf & "Category = " & strCategory & vbCrLf & "Parent = " & strParent & _
        vbCrLf & "Show String = " & strShowString & vbCrLf & "Save String = " & strSaveString & vbCrLf & _
        "Workspace Pathname = " & pWorkspace.PathName
     
    ' FOR DEBUGGING
  '  Dim frmReportForm As EcoAir_Map.frmReportDialog
  '  Beep
  '  Set frmReportForm = New EcoAir_Map.frmReportDialog
  '  frmReportForm.txtReport.Text = strWsPathName & vbCrLf & strCategory & vbCrLf & pTableName
  '  frmReportForm.Show vbModal
  '  Set frmReportForm = Nothing
  '  MsgBox strWsPathname & vbCrLf & strCategory & vbCrLf & lngCategory
    If Right(strWsPathname, 1) = "\" Or Right(strWsPathname, 1) = "/" Then
      strWsPathname = Left(strWsPathname, Len(strWsPathname) - 1)
    End If
    Set pFeatWS = pWSFact.OpenFromFile(strWsPathname, 0)
        
    booAlreadyExists = CheckIfTableExists(pFeatWS, strBaseName)
    If booAlreadyExists Then
      lngResponse = MsgBox("The Table '" & pTableName & "' already exists!  Do you wish to overwrite it?", _
            vbYesNo, "Table Already Exists:")
      If lngResponse = vbYes Then
        ' DELETE FEATURE CLASS
        Dim pDelete As ITable
        Dim pDelDataset As IDataset
        Set pDelete = pFeatWS.OpenTable(strBaseName)
        Set pDelDataset = pDelete
        If pDelDataset.CanDelete Then
          pDelDataset.DELETE
        Else
          MsgBox "Unable to delete " & pDataset.Name & "!  Bailing out..."
          Set CreateGeneralTable = Nothing
          Exit Function
        End If
        booAlreadyExists = CheckIfTableExists(pFeatWS, strBaseName)
      Else
        Set CreateGeneralTable = Nothing
        Set pGxObj = Nothing
        strName = ""
      End If
    End If
  Loop
  
  ' NOW READY TO CREATE TABLE
  If strCategory = "Shapefile" Or strCategory = "Folder" Or strCategory = "dBASE Table" Then      ' IF dBASE TABLE
    Set pTable = CreatedBASETableInFolder(strWsPathname, strBaseName, pAddFields)
  Else                                                                                            ' IF IN FILE OR PERSONAL GEODATABASE
    Set pTable = CreateGDBTable(pFeatWS, strBaseName, pAddFields)
  End If
  
  Set CreateGeneralTable = pTable

  GoTo ClearMemory

ClearMemory:
  Set pGxDataset = Nothing
  Set pDataset = Nothing
  Set pWorkspace = Nothing
  Set pFeatureDatasetName = Nothing
  Set pFeatureDataset = Nothing
  Set pFeatWS = Nothing
  Set pWSFact = Nothing
  Set pTable = Nothing
  Set pGxObj = Nothing
  Set pGxDialog = Nothing
  Set pGxFCFilter = Nothing
  Set pFilterCol = Nothing
  Set pDelete = Nothing
  Set pDelDataset = Nothing

End Function


Public Function CreateGeneralFeatureClass(pGeomType As esriGeometryType, pSpatialReference As ISpatialReference, _
    Optional strDir As String = "", Optional strName As String = "", Optional lngCategory As JenDatasetTypes = ENUM_Shapefile, _
    Optional pAddFields As esriSystem.IVariantArray, Optional strGxDialogCaption As String = "Specify New Feature Class Name:", _
    Optional booIsInFeatureDataset As Boolean = False, Optional strFeatureDatasetName As String, _
    Optional booIsFClass As Boolean = True, Optional booForceUniqueIDField As Boolean = True, _
    Optional pExtent As IEnvelope, Optional lngOriginalRecCount As Long = -9999) As IFeatureClass
  
'  MsgBox "Geometry Type = " & CStr(pGeomType) & vbCrLf & _
         "Spatial Reference = " & pSpatialReference.Name & vbCrLf & _
         "Directory = " & strDir & vbCrLf & _
         "strName = " & strName & vbCrLf & _
         "Category = " & CStr(lngCategory) & vbCrLf & _
         "Additional Field Count = " & CStr(pAddFields.Count)
  
  ' IF booIsFClass = FALSE, THEN THIS FUNCTION WILL CREATE A TABLE INSTEAD
  
  ' THIS FUNCTION ONLY ALLOWS YOU TO WRITE TO A FEATURE DATASET IF IT FORCES USER TO SELECT FEATURE DATASET WITH IGxCatalog.
  ' IF YOU SEND A SPECIFIC PATHNAME AND FILENAME, THIS FUNCTION WILL NOT ALLOW YOU TO SAVE TO A FEATURE DATASET.
  
  Dim pGxDataset As IGxDataset
  Dim pDataset As IDataset
  Dim pWorkspace As IWorkspace
  Dim pFeatureDatasetName As IName
  Dim pFeatureDataset As IFeatureDataset
  Dim booHasFeatureDataset As Boolean
  Dim strFeatDataset As String
  Dim strCategory As String
  Dim strParent As String
  Dim strBaseName As String
  Dim pFeatWS As IFeatureWorkspace
  Dim pWSFact As IWorkspaceFactory2
  Dim pFClassName As String
  Dim pFClass As IFeatureClass
  Dim strWsPathname As String
  Dim pGxObj As IGxObject
  Dim pGxDialog As IGxDialog
  Dim pGxFCFilter As IGxObjectFilter
  Dim pFilterCol As IGxObjectFilterCollection

  ' CHECK IF FEATURECLASS ALREADY EXISTS
  Dim booAlreadyExists As Boolean
  booAlreadyExists = True
  Dim lngResponse As VbMsgBoxResult
  
  Do Until booAlreadyExists = False
     
    ' FIRST, CHECK IF DIRECTORY, NAME AND CATEGORY ARE SPECIFIED
    If strDir <> "" And strName <> "" Then
      
      Select Case lngCategory
        Case ENUM_Shapefile
          strCategory = "Shapefile"
          strName = aml_func_mod.SetExtension(strName, "shp")
          If Right(strDir, 1) <> "\" Then strDir = strDir & "\"
        Case ENUM_FileGDB
          strCategory = "File Geodatabase Feature Class"
        Case ENUM_PersonalGDB
          strCategory = "Personal Geodatabase Feature Class"
        Case Else
          MsgBox "PROBLEM:  This function can only create shapefiles, file geodatabase feature classes, or " & vbCrLf & _
              "personal geodatabase feature classes.  Bailing out...", vbOKOnly, "Invalid Output Type:"
          Set CreateGeneralFeatureClass = Nothing
          Exit Function
      End Select
      
      strWsPathname = strDir
      booHasFeatureDataset = False
      strBaseName = strName
      booHasFeatureDataset = booIsInFeatureDataset
      strFeatDataset = strFeatureDatasetName
    Else
    
 '     If pGxObj Is Nothing Or IsMissing(pGxObj) Then  ' IF NEED TO ASK USER TO SPECIFY GxObject
      
      Set pGxDialog = New GxDialog
      Set pGxFCFilter = New GxFilterFeatureClasses
      
      Set pFilterCol = pGxDialog
      pFilterCol.AddFilter pGxFCFilter, True
      pGxDialog.Title = strGxDialogCaption
      pGxDialog.AllowMultiSelect = False
      pGxDialog.ButtonCaption = "Enter"
        
      If Not pGxDialog.DoModalSave(0) Then
          Set CreateGeneralFeatureClass = Nothing
          Exit Function 'Exit if user press Cancel
      End If
      
      Set pGxObj = pGxDialog.FinalLocation
      strWsPathname = pGxObj.FullName
      strBaseName = pGxDialog.Name
    
 '     End If
  
  '    Dim pGxCat As IGxCatalog
  '    Set pGxCat = pGxDialog.InternalCatalog
  
      ' CHECK TO SEE IF WE ARE WRITING TO A FEATURE DATASET
      If TypeOf pGxObj Is IGxDataset Then
        Set pGxDataset = pGxObj
          
        Set pDataset = pGxDataset.Dataset
        Set pWorkspace = pDataset.Workspace
        strWsPathname = pWorkspace.PathName
      
''        Set pFeatureClassName = pDataset.FullName
        Set pFeatureDatasetName = pDataset.FullName
        
        booHasFeatureDataset = TypeOf pDataset Is IFeatureDataset
        If booHasFeatureDataset Then
          Set pFeatureDataset = pDataset
          strFeatDataset = pFeatureDataset.Name
        End If
      End If
          
'      strPath = pGxObj.FullName
      strCategory = pGxObj.Category
'      strName = pGxObj.Name
'      strParent = pGxObj.Parent.FullName
'      strBaseName = pGxObj.BaseName
'      strWsPathName = pWorkspace.PathName
    End If
    
    ' NOW WE KNOW NAME, CATEGORY AND WORKSPACE.  NEXT CREATE WORKSPACE FACTORY
    Select Case strCategory
       Case "Shapefile", "Folder"
         Set pWSFact = New ShapefileWorkspaceFactory
         pFClassName = strBaseName
       Case "File Geodatabase Feature Class", "File Geodatabase Raster Catalog", "File Geodatabase", "File Geodatabase Feature Dataset"
         Set pWSFact = New FileGDBWorkspaceFactory
         If booHasFeatureDataset Then
           pFClassName = strFeatDataset & ":" & strBaseName
         Else
           pFClassName = strBaseName
         End If
       Case "Personal Geodatabase Feature Class", "Personal Geodatabase Raster Catalog", "Personal Geodatabase", "Personal Geodatabase Feature Dataset"
         Set pWSFact = New AccessWorkspaceFactory
         If booHasFeatureDataset Then
           pFClassName = strFeatDataset & ":" & strBaseName
         Else
           pFClassName = strBaseName
         End If
'       Case "Polygon Feature Class", "Arc Feature Class", "Point Feature Class", "Route Feature Class", _
'            "Tic Feature Class", "Region Feature Class", "Label Feature Class"
'         Set pWSFact = New ArcInfoWorkspaceFactory
'         If booHasFeatureDataset Then
'           pFClassName = strFeatDataset & ":" & strBaseName
'         Else
'           pFClassName = strBaseName
'         End If
'       Case "SDC Feature Class"
'         Set pWSFact = New SDCWorkspaceFactory
'         If booHasFeatureDataset Then
'           pFClassName = strFeatDataset & ":" & strBaseName
'         Else
'           pFClassName = strBaseName
'         End If
       Case Else
         MsgBox "Unexpected Feature Class Type (" & strCategory & ")!  Need to write new code for this..."
         Set CreateGeneralFeatureClass = Nothing
         Exit Function
    End Select
    
  '  Debug.Print "Full Name = " & strPath & vbCrLf & "Name = " & strName & vbCrLf & _
        "Base Name = " & strBaseName & vbCrLf & "Category = " & strCategory & vbCrLf & "Parent = " & strParent & _
        vbCrLf & "Show String = " & strShowString & vbCrLf & "Save String = " & strSaveString & vbCrLf & _
        "Workspace Pathname = " & pWorkspace.PathName
     
    ' FOR DEBUGGING
  '  Dim frmReportForm As EcoAir_Map.frmReport_modal
  '  Beep
  '  Set frmReportForm = New EcoAir_Map.frmReport_modal
  '  frmReportForm.txtReport.Text = strWsPathName & vbCrLf & strCategory & vbCrLf & pFClassName
  '  frmReportForm.Show vbModal
  '  Set frmReportForm = Nothing
  '  MsgBox strWsPathname & vbCrLf & strCategory & vbCrLf & lngCategory
    If Right(strWsPathname, 1) = "\" Or Right(strWsPathname, 1) = "/" Then
      strWsPathname = Left(strWsPathname, Len(strWsPathname) - 1)
    End If
    If Right(strWsPathname, 1) = "\" Or Right(strWsPathname, 1) = "/" Then
      strWsPathname = Left(strWsPathname, Len(strWsPathname) - 1)
    End If
    Set pFeatWS = pWSFact.OpenFromFile(strWsPathname, 0)
    
    ' GET FEATURE DATASET IF IT EXISTS
    If booHasFeatureDataset Then Set pFeatureDataset = pFeatWS.OpenFeatureDataset(strFeatDataset)
    
    booAlreadyExists = CheckIfFeatureClassExists(pFeatWS, strBaseName)
    If booAlreadyExists Then
      lngResponse = MsgBox("The Feature Class '" & pFClassName & "' already exists!  Do you wish to overwrite it?", _
            vbYesNo, "Feature Class Already Exists:")
      If lngResponse = vbYes Then
        ' DELETE FEATURE CLASS
        Dim pDelete As IFeatureClass
        Dim pDelDataset As IDataset
        Set pDelete = pFeatWS.OpenFeatureClass(strBaseName)
        Set pDelDataset = pDelete
        pDelDataset.DELETE
        booAlreadyExists = CheckIfFeatureClassExists(pFeatWS, strBaseName)
      Else
        Set CreateGeneralFeatureClass = Nothing
        Set pGxObj = Nothing
        strName = ""
      End If
    End If
  Loop
  
  
  ' NOW READY TO CREATE FEATURE CLASS
  If booHasFeatureDataset Then                                                     ' IF IN FEATURE DATASET
    Set pFClass = CreateDatasetFeatureClass(pFeatureDataset, strBaseName, esriFTSimple, pGeomType, pAddFields)
  ElseIf strCategory = "Shapefile" Or strCategory = "Folder" Then                  ' IF SHAPEFILE
    Set pFClass = CreateShapefileFeatureClass(strWsPathname, strBaseName, pSpatialReference, pGeomType, pAddFields, booForceUniqueIDField)
  Else                                                                             ' IF IN FILE OR PERSONAL GEODATABASE
'    Set pFClass = aml_func_mod.createWorkspaceFeatureClass(pFeatWS, strBaseName, esriFTSimple, pGeomType)
    Set pFClass = CreateGDBFeatureClass(pFeatWS, strBaseName, esriFTSimple, pSpatialReference, pGeomType, pAddFields, , , , , _
          lngCategory, pExtent, lngOriginalRecCount)
  End If
  
  Set CreateGeneralFeatureClass = pFClass

ClearMemory:
  Set pGxDataset = Nothing
  Set pDataset = Nothing
  Set pWorkspace = Nothing
  Set pFeatureDatasetName = Nothing
  Set pFeatureDataset = Nothing
  Set pFeatWS = Nothing
  Set pWSFact = Nothing
  Set pFClass = Nothing
  Set pGxObj = Nothing
  Set pGxDialog = Nothing
  Set pGxFCFilter = Nothing
  Set pFilterCol = Nothing
  Set pDelete = Nothing
  Set pDelDataset = Nothing

End Function

Public Function CheckIfTableExists(pWS As IFeatureWorkspace, strName As String) As Boolean

    Dim pTable As ITable
    On Error Resume Next
    Set pTable = pWS.OpenTable(strName)
    If pTable Is Nothing Then
        CheckIfTableExists = False
    Else
        CheckIfTableExists = True
    End If
    If CheckIfTableExists = False Then
      Dim pFClass As ITable
      Set pFClass = pWS.OpenFeatureClass(strName)
      If pFClass Is Nothing Then
          CheckIfTableExists = False
      Else
          CheckIfTableExists = True
      End If
    End If


  GoTo ClearMemory
ClearMemory:
  Set pTable = Nothing
  Set pFClass = Nothing
End Function
Public Function CheckIfFeatureClassExists(pWS As IFeatureWorkspace, strName As String) As Boolean

    Dim pFClass As IFeatureClass
    On Error Resume Next
    Set pFClass = pWS.OpenFeatureClass(strName)
    If pFClass Is Nothing Then
        CheckIfFeatureClassExists = False
    Else
        CheckIfFeatureClassExists = True
    End If
    If CheckIfFeatureClassExists = False Then
      Dim pTable As ITable
      Set pTable = pWS.OpenTable(strName)
      If pTable Is Nothing Then
          CheckIfFeatureClassExists = False
      Else
          CheckIfFeatureClassExists = True
      End If
    End If


  GoTo ClearMemory
ClearMemory:
  Set pFClass = Nothing
  Set pTable = Nothing

End Function

Public Function MakeUniqueGDBTableName(pWS As IWorkspace, strName As String) As String
  
  Dim booTableExists As Boolean
  booTableExists = CheckIfTableExists(pWS, strName)
  Dim theCounter As Long
  theCounter = 1
  Dim strBaseName As String
  Dim strNewName As String
  strBaseName = strName
  strNewName = strName
  Do Until booTableExists = False
    theCounter = theCounter + 1
    strNewName = strBaseName & "_" & CStr(theCounter)
    booTableExists = CheckIfTableExists(pWS, strNewName)
  Loop
  
  MakeUniqueGDBTableName = strNewName

End Function
Public Function MakeUniqueGDBFeatureClassName(pWS As IWorkspace, strName As String) As String
  
  Dim booFCExists As Boolean
  booFCExists = CheckIfFeatureClassExists(pWS, strName)
  Dim theCounter As Long
  theCounter = 1
  Dim strBaseName As String
  Dim strNewName As String
  strBaseName = strName
  strNewName = strName
  Do Until booFCExists = False
    theCounter = theCounter + 1
    strNewName = strBaseName & "_" & CStr(theCounter)
    booFCExists = CheckIfFeatureClassExists(pWS, strNewName)
  Loop
  
  MakeUniqueGDBFeatureClassName = strNewName

End Function

Public Function MakeUniquedBASEName(strFilename As String) As String

  If Not ExistFileDir(strFilename) Then
    MakeUniquedBASEName = strFilename
    Exit Function
  Else
    
    Dim theCounter As Long
    theCounter = 1
    
    Dim theFilename As String
    Dim theBaseName As String
    Dim thePointPos As Long
    Dim theExtension As String
    
    If InStr(1, Right(strFilename, 5), ".", vbTextCompare) > 0 Then
      thePointPos = InStrRev(strFilename, ".", -1, vbTextCompare)
      theExtension = Right(strFilename, Len(strFilename) - (thePointPos - 1))
      theFilename = Left(strFilename, thePointPos - 1)
    Else
      theExtension = ""
      theFilename = strFilename
    End If
    
    theBaseName = theFilename
    
    Do While ExistFileDir(theFilename & theExtension)
      theCounter = theCounter + 1
      theFilename = theBaseName & "_" & CStr(theCounter)
    Loop
    
    MakeUniquedBASEName = theFilename & theExtension
    
  End If

End Function

Public Function MakeUniqueShapeFilename(strFilename As String) As String

  Dim theFilename As String
  Dim theBaseName As String
  Dim thePointPos As Long
  Dim theExtension As String
  
  If InStr(1, Right(strFilename, 5), ".", vbTextCompare) > 0 Then
    thePointPos = InStrRev(strFilename, ".", -1, vbTextCompare)
    theExtension = Right(strFilename, Len(strFilename) - (thePointPos - 1))
    theFilename = Left(strFilename, thePointPos - 1)
  Else
    theExtension = ""
    theFilename = strFilename
  End If
  
  If Not aml_func_mod.ExistFileDir(strFilename) And Not aml_func_mod.ExistFileDir(theFilename) Then
    MakeUniqueShapeFilename = strFilename
    Exit Function
  Else
    
    Dim theCounter As Long
    theCounter = 1
    
    
    theBaseName = theFilename

    Do While aml_func_mod.ExistFileDir(theFilename & theExtension) Or aml_func_mod.ExistFileDir(theFilename)
      theCounter = theCounter + 1
      theFilename = theBaseName & "_" & CStr(theCounter)
    Loop
    
    MakeUniqueShapeFilename = theFilename & theExtension
    
  End If

End Function


Public Function CreateGDBTable(featWorkspace As IFeatureWorkspace, _
                                            Name As String, _
                                            Optional pAddFields As esriSystem.IVariantArray, _
                                            Optional pCLSID As UID, _
                                            Optional pCLSEXT As UID, _
                                            Optional ConfigWord As String = "" _
                                            ) As ITable
 
'' createWorkspaceFeatureClass: simple helper to create a featureclass in a geodatabase workspace.
'' NOTE: when creating a feature class in a workspace it is important to assign the spatial
''       reference to the geometry field.
'' MODIFIED BY JENNESS APRIL 23 2008 FROM ESRI SAMPLE
  
  Set CreateGDBTable = Nothing
  If featWorkspace Is Nothing Then Exit Function
  If Name = "" Then Exit Function
  
  If (pCLSID Is Nothing) Or IsMissing(pCLSID) Then
    Set pCLSID = Nothing
    Set pCLSID = New UID
    pCLSID.Value = "esriGeoDatabase.Object"
  End If
  
  ' CREATE OBJECT ID FIELD; ADD ANY REQUESTED FIELDS IF NECESSARY
  ' establish a fields collection
  Dim pFields As IFields
  Dim pFieldsEdit As IFieldsEdit
  Dim pField As iField
  Dim pFieldEdit As IFieldEdit
  
  Set pFields = New Fields
  Set pFieldsEdit = pFields
  
  '' create the object id field
  Set pField = New Field
  Set pFieldEdit = pField
  pFieldEdit.Name = "Object_ID"
 '   pFieldEdit.AliasName = "object identifier"
  pFieldEdit.Type = esriFieldTypeOID
  pFieldsEdit.AddField pField
 
  ' ADD FIELDS IF REQUESTED
  If Not (pAddFields Is Nothing Or IsMissing(pAddFields)) Then
    Dim lngIndex As Long
    For lngIndex = 0 To pAddFields.Count - 1
      pFieldsEdit.AddField pAddFields.Element(lngIndex)
    Next lngIndex
  End If
  
  ' establish the class extension
  If (pCLSEXT Is Nothing) Or IsMissing(pCLSEXT) Then
    Set pCLSEXT = Nothing
  End If
  
  Set CreateGDBTable = featWorkspace.CreateTable(Name, pFields, pCLSID, pCLSEXT, "")


  GoTo ClearMemory
ClearMemory:
  Set pFields = Nothing
  Set pFieldsEdit = Nothing
  Set pField = Nothing
  Set pFieldEdit = Nothing

End Function

Public Function CreateGDBFeatureClass2(featWorkspace As IFeatureWorkspace, _
                                            Name As String, _
                                            featType As esriFeatureType, _
                                            pSpRef As ISpatialReference, _
                                            Optional geomType As esriGeometryType = esriGeometryPoint, _
                                            Optional pAddFields As esriSystem.IVariantArray, _
                                            Optional pCLSID As UID, _
                                            Optional pCLSEXT As UID, _
                                            Optional ConfigWord As String = "", _
                                            Optional booForceUniqueIDField As Boolean = True, _
                                            Optional lngCategory As JenDatasetTypes, _
                                            Optional pExtent As IEnvelope, _
                                            Optional lngOriginalRecCount As Long = -9999, _
                                            Optional booHasZ As Boolean, _
                                            Optional booHasM As Boolean) As IFeatureClass
 
'' createWorkspaceFeatureClass: simple helper to create a featureclass in a geodatabase workspace.
'' NOTE: when creating a feature class in a workspace it is important to assign the spatial
''       reference to the geometry field.
'' MODIFIED BY JENNESS APRIL 23 2008 FROM ESRI SAMPLE
  
  Set CreateGDBFeatureClass2 = Nothing
  If featWorkspace Is Nothing Then Exit Function
  If Name = "" Then Exit Function
  
  If (pCLSID Is Nothing) Or IsMissing(pCLSID) Then
    Set pCLSID = Nothing
    Set pCLSID = New UID
    
    '' determine the appropriate geometry type corresponding the the feature type
    Select Case featType
      Case esriFTSimple
        pCLSID.Value = "esriGeoDatabase.Feature"
        If geomType = esriGeometryLine Then geomType = esriGeometryPolyline
      Case esriFTSimpleJunction
        geomType = esriGeometryPoint
        pCLSID.Value = "esriGeoDatabase.SimpleJunctionFeature"
      Case esriFTComplexJunction
        pCLSID.Value = "esriGeoDatabase.ComplexJunctionFeature"
      Case esriFTSimpleEdge
        geomType = esriGeometryPolyline
        pCLSID.Value = "esriGeoDatabase.SimpleEdgeFeature"
      Case esriFTComplexEdge
        geomType = esriGeometryPolyline
        pCLSID.Value = "esriGeoDatabase.ComplexEdgeFeature"
      Case esriFTAnnotation
        Exit Function
    End Select
  End If
  
  ' CREATE GEOMETRY AND OBJECT ID FIELD; ADD ANY REQUESTED FIELDS IF NECESSARY
  ' establish a fields collection
  Dim pFields As IFields
  Dim pFieldsEdit As IFieldsEdit
  Dim pGeomDef As IGeometryDef
  Dim pGeomDefEdit As IGeometryDefEdit
  Dim pField As iField
  Dim pFieldEdit As IFieldEdit
  
  Set pFields = New Fields
  Set pFieldsEdit = pFields
  
  '' create the geometry field
  Set pGeomDef = New GeometryDef
  Set pGeomDefEdit = pGeomDef
  
  ' SPATIAL INDEX STUFF
  ' WOULD LIKE THE AVERAGE NUMBER OF FEATURES PER GRID CELL TO BE AROUND 200
  ' GOING TO DEFAULT TO ONE SPATIAL INDEX GRID
  Dim dblIndex0 As Double
  Dim dblIndex1 As Double
  Dim dblIndex2 As Double
  dblIndex0 = 100
  dblIndex1 = 0
  dblIndex2 = 0
  
  If Not pExtent Is Nothing And lngOriginalRecCount > 0 Then
    If Not pExtent.IsEmpty Then
      Dim dblOptimalGridCells As Double
      dblOptimalGridCells = lngOriginalRecCount / 200
      Dim dblX As Double
      Dim dblY As Double
      dblX = pExtent.Width
      dblY = pExtent.Height
      dblIndex0 = Sqr(dblX * dblY / dblOptimalGridCells)
      If dblIndex0 > dblX Then
        dblIndex0 = dblX
      End If
      If dblIndex0 > dblY Then
        dblIndex0 = dblY
      End If
      
      dblIndex1 = 0
      dblIndex2 = 0
    End If
  End If
  
  If Not CheckSpRefDomain(pSpRef) Then
    Dim pSpRefRes As ISpatialReferenceResolution
    Set pSpRefRes = pSpRef
    pSpRefRes.ConstructFromHorizon
  End If
  
  '' assign the geometry definiton properties.
  With pGeomDefEdit
    .GeometryType = geomType
    If lngCategory = ENUM_FileGDB Then       ' FILE GDB
      .GridCount = 1
      .GridSize(0) = dblIndex0
'      .GridSize(1) = dblIndex1
'      .GridSize(2) = dblIndex2
'      .AvgNumPoints = 2
    Else
      .GridCount = 1
      .GridSize(0) = dblIndex0
    End If
    .HasM = booHasM
    .HasZ = booHasZ
    Set .SpatialReference = pSpRef
  End With
  
  Set pField = New Field
  Set pFieldEdit = pField
  
  pFieldEdit.Name = "Shape"
  pFieldEdit.AliasName = "geometry"
  pFieldEdit.Type = esriFieldTypeGeometry
  Set pFieldEdit.GeometryDef = pGeomDef
  pFieldsEdit.AddField pField

  '' create the object id field
'  If booForceUniqueIDField Then
    
    
    Dim strUniqueName As String
    Dim strUniqueBase As String
    strUniqueName = "Object_ID"
    strUniqueBase = strUniqueName
    Dim lngUniqueCounter As Long
    lngUniqueCounter = 1
    Dim lngCounterLength As Long
    Dim lngUniqueIndex As Long
    Dim booNameExists As Boolean
    Dim pCheckField As iField
    booNameExists = True
    If (pAddFields Is Nothing Or IsMissing(pAddFields)) Then
      booNameExists = False
    Else
      Do Until booNameExists = False
        booNameExists = False
        For lngUniqueIndex = 0 To pAddFields.Count - 1
          Set pCheckField = pAddFields.Element(lngUniqueIndex)
          If pCheckField.Name = strUniqueName Then
            booNameExists = True
            Exit For
          End If
        Next lngUniqueIndex
        If booNameExists Then
          lngUniqueCounter = lngUniqueCounter + 1
          lngCounterLength = Len(CStr(lngUniqueCounter))
          strUniqueName = Left(strUniqueBase, 11 - lngCounterLength) & CStr(lngUniqueCounter)
        End If
      Loop
    End If
    
    Set pField = New Field
    Set pFieldEdit = pField
    pFieldEdit.Name = strUniqueName
    pFieldEdit.AliasName = "object identifier"
    pFieldEdit.Type = esriFieldTypeOID
    pFieldsEdit.AddField pField
'  End If
  
  ' ADD FIELDS IF REQUESTED
  If Not (pAddFields Is Nothing Or IsMissing(pAddFields)) Then
    Dim lngIndex As Long
    For lngIndex = 0 To pAddFields.Count - 1
      pFieldsEdit.AddField pAddFields.Element(lngIndex)
    Next lngIndex
  End If
  
  
  ' establish the class extension
  If (pCLSEXT Is Nothing) Or IsMissing(pCLSEXT) Then
    Set pCLSEXT = Nothing
  End If
  
  ' locate the shape field
  Dim strShapeFld As String
  Dim j As Integer
  For j = 0 To pFields.FieldCount - 1
    If pFields.Field(j).Type = esriFieldTypeGeometry Then
      strShapeFld = pFields.Field(j).Name
    End If
  Next
  
  Set CreateGDBFeatureClass2 = featWorkspace.CreateFeatureClass(Name, pFields, pCLSID, _
                             pCLSEXT, featType, strShapeFld, ConfigWord)

  GoTo ClearMemory
ClearMemory:
  Set pFields = Nothing
  Set pFieldsEdit = Nothing
  Set pGeomDef = Nothing
  Set pGeomDefEdit = Nothing
  Set pField = Nothing
  Set pFieldEdit = Nothing
  Set pSpRefRes = Nothing
  Set pCheckField = Nothing

End Function

Public Function CreateGDBFeatureClass(featWorkspace As IFeatureWorkspace, _
                                            strName As String, _
                                            featType As esriFeatureType, _
                                            pSpRef As ISpatialReference, _
                                            Optional geomType As esriGeometryType = esriGeometryPoint, _
                                            Optional pAddFields As esriSystem.IVariantArray, _
                                            Optional pCLSID As UID, _
                                            Optional pCLSEXT As UID, _
                                            Optional ConfigWord As String = "", _
                                            Optional booForceUniqueIDField As Boolean = True, _
                                            Optional lngCategory As JenDatasetTypes, _
                                            Optional pExtent As IEnvelope, _
                                            Optional lngOriginalRecCount As Long = -9999) As IFeatureClass
 
'' createWorkspaceFeatureClass: simple helper to create a featureclass in a geodatabase workspace.
'' NOTE: when creating a feature class in a workspace it is important to assign the spatial
''       reference to the geometry field.
'' MODIFIED BY JENNESS APRIL 23 2008 FROM ESRI SAMPLE
  
  Set CreateGDBFeatureClass = Nothing
  If featWorkspace Is Nothing Then Exit Function
  If strName = "" Then Exit Function
  
  If (pCLSID Is Nothing) Or IsMissing(pCLSID) Then
    Set pCLSID = Nothing
    Set pCLSID = New UID
    
    '' determine the appropriate geometry type corresponding the the feature type
    Select Case featType
      Case esriFTSimple
        pCLSID.Value = "esriGeoDatabase.Feature"
        If geomType = esriGeometryLine Then geomType = esriGeometryPolyline
      Case esriFTSimpleJunction
        geomType = esriGeometryPoint
        pCLSID.Value = "esriGeoDatabase.SimpleJunctionFeature"
      Case esriFTComplexJunction
        pCLSID.Value = "esriGeoDatabase.ComplexJunctionFeature"
      Case esriFTSimpleEdge
        geomType = esriGeometryPolyline
        pCLSID.Value = "esriGeoDatabase.SimpleEdgeFeature"
      Case esriFTComplexEdge
        geomType = esriGeometryPolyline
        pCLSID.Value = "esriGeoDatabase.ComplexEdgeFeature"
      Case esriFTAnnotation
        Exit Function
    End Select
  End If
  
  ' CREATE GEOMETRY AND OBJECT ID FIELD; ADD ANY REQUESTED FIELDS IF NECESSARY
  ' establish a fields collection
  Dim pFields As IFields
  Dim pFieldsEdit As IFieldsEdit
  Dim pGeomDef As IGeometryDef
  Dim pGeomDefEdit As IGeometryDefEdit
  Dim pField As iField
  Dim pFieldEdit As IFieldEdit
  
  Set pFields = New Fields
  Set pFieldsEdit = pFields
  
  '' create the geometry field
  Set pGeomDef = New GeometryDef
  Set pGeomDefEdit = pGeomDef
  
  ' SPATIAL INDEX STUFF
  ' WOULD LIKE THE AVERAGE NUMBER OF FEATURES PER GRID CELL TO BE AROUND 200
  ' GOING TO DEFAULT TO ONE SPATIAL INDEX GRID
  Dim dblIndex0 As Double
  Dim dblIndex1 As Double
  Dim dblIndex2 As Double
  dblIndex0 = 100
  dblIndex1 = 0
  dblIndex2 = 0
  
  If Not pExtent Is Nothing And lngOriginalRecCount > 0 Then
    If Not pExtent.IsEmpty Then
      Dim dblOptimalGridCells As Double
      dblOptimalGridCells = lngOriginalRecCount / 200
      Dim dblX As Double
      Dim dblY As Double
      dblX = pExtent.Width
      dblY = pExtent.Height
      dblIndex0 = Sqr(dblX * dblY / dblOptimalGridCells)
      If dblIndex0 > dblX Then
        dblIndex0 = dblX
      End If
      If dblIndex0 > dblY Then
        dblIndex0 = dblY
      End If
      
      dblIndex1 = 0
      dblIndex2 = 0
    End If
  End If
  
  If Not CheckSpRefDomain(pSpRef) Then
    Dim pSpRefRes As ISpatialReferenceResolution
    Set pSpRefRes = pSpRef
    pSpRefRes.ConstructFromHorizon
  End If
  
  '' assign the geometry definiton properties.
  With pGeomDefEdit
    .GeometryType = geomType
    If lngCategory = ENUM_FileGDB Then       ' FILE GDB
      .GridCount = 1
      .GridSize(0) = dblIndex0
'      .GridSize(1) = dblIndex1
'      .GridSize(2) = dblIndex2
'      .AvgNumPoints = 2
    Else
      .GridCount = 1
      .GridSize(0) = dblIndex0
    End If
    .HasM = False
    .HasZ = False
    Set .SpatialReference = pSpRef
  End With
  
  Set pField = New Field
  Set pFieldEdit = pField
  
  pFieldEdit.Name = "Shape"
  pFieldEdit.AliasName = "geometry"
  pFieldEdit.Type = esriFieldTypeGeometry
  Set pFieldEdit.GeometryDef = pGeomDef
  pFieldsEdit.AddField pField

  '' create the object id field
'  If booForceUniqueIDField Then
    
    
    Dim strUniqueName As String
    Dim strUniqueBase As String
    strUniqueName = "Object_ID"
    strUniqueBase = strUniqueName
    Dim lngUniqueCounter As Long
    lngUniqueCounter = 1
    Dim lngCounterLength As Long
    Dim lngUniqueIndex As Long
    Dim booNameExists As Boolean
    Dim pCheckField As iField
    booNameExists = True
    If (pAddFields Is Nothing Or IsMissing(pAddFields)) Then
      booNameExists = False
    Else
      Do Until booNameExists = False
        booNameExists = False
        For lngUniqueIndex = 0 To pAddFields.Count - 1
          Set pCheckField = pAddFields.Element(lngUniqueIndex)
          If pCheckField.Name = strUniqueName Then
            booNameExists = True
            Exit For
          End If
        Next lngUniqueIndex
        If booNameExists Then
          lngUniqueCounter = lngUniqueCounter + 1
          lngCounterLength = Len(CStr(lngUniqueCounter))
          strUniqueName = Left(strUniqueBase, 11 - lngCounterLength) & CStr(lngUniqueCounter)
        End If
      Loop
    End If
    
    Set pField = New Field
    Set pFieldEdit = pField
    pFieldEdit.Name = strUniqueName
    pFieldEdit.AliasName = "object identifier"
    pFieldEdit.Type = esriFieldTypeOID
    pFieldsEdit.AddField pField
'  End If
  
  ' ADD FIELDS IF REQUESTED
  If Not (pAddFields Is Nothing Or IsMissing(pAddFields)) Then
    Dim lngIndex As Long
    For lngIndex = 0 To pAddFields.Count - 1
      pFieldsEdit.AddField pAddFields.Element(lngIndex)
    Next lngIndex
  End If
  
  
  ' establish the class extension
  If (pCLSEXT Is Nothing) Or IsMissing(pCLSEXT) Then
    Set pCLSEXT = Nothing
  End If
  
  ' locate the shape field
  Dim strShapeFld As String
  Dim j As Integer
  For j = 0 To pFields.FieldCount - 1
    If pFields.Field(j).Type = esriFieldTypeGeometry Then
      strShapeFld = pFields.Field(j).Name
    End If
  Next
  
  Set CreateGDBFeatureClass = featWorkspace.CreateFeatureClass(strName, pFields, pCLSID, _
                             pCLSEXT, featType, strShapeFld, ConfigWord)

  GoTo ClearMemory

ClearMemory:
  Set pFields = Nothing
  Set pFieldsEdit = Nothing
  Set pGeomDef = Nothing
  Set pGeomDefEdit = Nothing
  Set pField = Nothing
  Set pFieldEdit = Nothing
  Set pSpRefRes = Nothing
  Set pCheckField = Nothing
End Function

Private Function CheckSpRefDomain(pSpRef As ISpatialReference) As Boolean
  On Error GoTo ErrHandler

  
  Dim dXmax As Double
  Dim dYmax As Double
  Dim dXmin As Double
  Dim dYmin As Double
  
  'get the xy domain extent of the dataset
  
  pSpRef.GetDomain dXmin, dXmax, dYmin, dYmax
  CheckSpRefDomain = True
  
  Exit Function
ErrHandler:
  CheckSpRefDomain = False


End Function

Public Function CreatedBASETableInFolder(sPath As String, sName As String, Optional pAddFields As esriSystem.IVariantArray) As ITable
  
  If Right(sPath, 4) = ".dbf" Then sPath = ReturnDir(sPath)
  If Right(sName, 4) = ".dbf" Then sName = Left(sName, Len(sName) - 4)
  
  ' Open the folder to contain the shapefile as a workspace
  Dim pFWS As IFeatureWorkspace
  Dim pWorkspaceFactory As IWorkspaceFactory
  Set pWorkspaceFactory = New ShapefileWorkspaceFactory
  
  If Not ExistFileDir(sPath) Then
    MsgBox "Unable to create dBASE Table:" & vbCrLf & _
           sPath & " is not a valid workspace...", , "Failed to Create Table:"
    Set CreatedBASETableInFolder = Nothing
    Exit Function
  End If
  
  Set pFWS = pWorkspaceFactory.OpenFromFile(sPath, 0)
  
'   Set up a simple fields collection
  Dim pFields As IFields
  Dim pFieldsEdit As IFieldsEdit
  Set pFields = New Fields
  Set pFieldsEdit = pFields

  Dim pField As iField
  Dim pFieldEdit As IFieldEdit
  
  Set pField = New Field
  Set pFieldEdit = pField
  With pFieldEdit
      .Precision = 8
      .Name = "Unique_ID"
      .Type = esriFieldTypeInteger
  End With
  pFieldsEdit.AddField pField
  
  ' ADD FIELDS IF REQUESTED
  If Not (pAddFields Is Nothing Or IsMissing(pAddFields)) Then
    Dim lngIndex As Long
    For lngIndex = 0 To pAddFields.Count - 1
      pFieldsEdit.AddField pAddFields.Element(lngIndex)
    Next lngIndex
  End If
    
  ' Create the dBASE Table
  ' (some parameters apply to geodatabase options and can be defaulted as Nothing)
  Dim booFileExists As Boolean
  Dim strCheckString As String
  If Right(sPath, 1) = "\" Then
    strCheckString = sPath & sName & ".dbf"
'    MsgBox sPath & sName & ".dbf" & vbCrLf & "File Exists? " & CStr(Dir(sPath & sName & ".dbf") <> "")
  Else
    strCheckString = sPath & "\" & sName & ".dbf"
'    MsgBox sPath & "\" & sName & ".dbf" & vbCrLf & "File Exists? " & CStr(Dir(sPath & "\" & sName & ".dbf") <> "")
  End If
  
  booFileExists = (Dir(strCheckString) <> "")
  
  If booFileExists Then
    MsgBox "The following file already exists:" & vbCrLf & vbCrLf & strCheckString & vbCrLf & vbCrLf & _
           "Please select a new filename...", , "Duplicate Filename:"
    Set CreatedBASETableInFolder = Nothing
    Exit Function
  End If
  
  Dim pTable As ITable
  Set pTable = pFWS.CreateTable(sName, pFields, Nothing, Nothing, "")
  
  Set CreatedBASETableInFolder = pTable

  GoTo ClearMemory

ClearMemory:
  Set pFWS = Nothing
  Set pWorkspaceFactory = Nothing
  Set pFields = Nothing
  Set pFieldsEdit = Nothing
  Set pField = Nothing
  Set pFieldEdit = Nothing
  Set pTable = Nothing

End Function
    

Public Function CreateShapefileFeatureClass2(sPath As String, sName As String, pSpatialReference As ISpatialReference, _
    pGeomType As esriGeometryType, Optional pAddFields As esriSystem.IVariantArray, _
    Optional booForceUniqueIDField As Boolean = True, Optional booHasZ As Boolean = False, _
    Optional booHasM As Boolean = False) As IFeatureClass                                                  ' Don't include filename!
  
  
  If Right(sPath, 4) = ".shp" Then sPath = ReturnDir(sPath)
  If Right(sName, 4) = ".shp" Then sName = Left(sName, Len(sName) - 4)
  
  ' SET GEOMETRY TYPE, AND EXIT IF NOT ONE OF STANDARD OPTIONS
  Dim pGeomDef As IGeometryDef
  Dim pGeomDefEdit As IGeometryDefEdit
  Set pGeomDef = New GeometryDef
  Set pGeomDefEdit = pGeomDef
  With pGeomDefEdit
    .GeometryType = pGeomType
    .HasM = booHasM
    .HasZ = booHasZ
    Set .SpatialReference = pSpatialReference
  End With
  
  ' Open the folder to contain the shapefile as a workspace
  Dim pFWS As IFeatureWorkspace
  Dim pWorkspaceFactory As IWorkspaceFactory
  Set pWorkspaceFactory = New ShapefileWorkspaceFactory
  
  If Not ExistFileDir(sPath) Then
    MsgBox "Unable to create Feature Class:" & vbCrLf & _
           sPath & " is not a valid workspace...", , "Failed to Create Feature Class:"
    Set CreateShapefileFeatureClass2 = Nothing
    Exit Function
  End If
  
  Set pFWS = pWorkspaceFactory.OpenFromFile(sPath, 0)
  
'   Set up a simple fields collection
  Dim pFields As IFields
  Dim pFieldsEdit As IFieldsEdit
  Set pFields = New Fields
  Set pFieldsEdit = pFields

  Dim pField As iField
  Dim pFieldEdit As IFieldEdit
  
  ' Make the shape field
  ' it will need a geometry definition, with a spatial reference
  Set pField = New Field
  Set pFieldEdit = pField
  pFieldEdit.Name = "Shape"
  pFieldEdit.Type = esriFieldTypeGeometry

  Set pFieldEdit.GeometryDef = pGeomDef
  pFieldsEdit.AddField pField
  
  ' MAKE SURE UNIQUE ID FIELD IS UNIQUELY NAMED
 
  If booForceUniqueIDField Then
    Dim strUniqueName As String
    Dim strUniqueBase As String
    strUniqueName = "Unique_ID"
    strUniqueBase = strUniqueName
    Dim lngUniqueCounter As Long
    lngUniqueCounter = 1
    Dim lngCounterLength As Long
    Dim lngUniqueIndex As Long
    Dim booNameExists As Boolean
    Dim pCheckField As iField
    booNameExists = True
    If (pAddFields Is Nothing Or IsMissing(pAddFields)) Then
      booNameExists = False
    Else
      Do Until booNameExists = False
        booNameExists = False
        For lngUniqueIndex = 0 To pAddFields.Count - 1
          Set pCheckField = pAddFields.Element(lngUniqueIndex)
          If pCheckField.Name = strUniqueName Then
            booNameExists = True
            Exit For
          End If
        Next lngUniqueIndex
        If booNameExists Then
          lngUniqueCounter = lngUniqueCounter + 1
          lngCounterLength = Len(CStr(lngUniqueCounter))
          strUniqueName = Left(strUniqueBase, 10 - lngCounterLength) & CStr(lngUniqueCounter)
        End If
      Loop
    End If
    
    Set pField = New Field
    Set pFieldEdit = pField
    With pFieldEdit
        .Precision = 8
        .Name = strUniqueName
        .Type = esriFieldTypeInteger
    End With
    pFieldsEdit.AddField pField
  End If
  
  ' ADD FIELDS IF REQUESTED
'  Dim strDebugReport As String
'  Dim pDebugField As IField
  If Not (pAddFields Is Nothing Or IsMissing(pAddFields)) Then
    Dim lngIndex As Long
    For lngIndex = 0 To pAddFields.Count - 1
'      Set pDebugField = pAddFields.Element(lngIndex)
'      strDebugReport = strDebugReport & CStr(lngIndex) & "]  Field Name = " & pDebugField.Name & vbCrLf
      pFieldsEdit.AddField pAddFields.Element(lngIndex)
    Next lngIndex
  End If
'  MsgBox CStr(pGeomDef.GeometryType) & vbCrLf & strDebugReport
  ' Create the shapefile
  ' (some parameters apply to geodatabase options and can be defaulted as Nothing)
  Dim booFileExists As Boolean
  Dim strCheckString As String
  If Right(sPath, 1) = "\" Then
    strCheckString = sPath & sName & ".shp"
'    MsgBox sPath & sName & ".shp" & vbCrLf & "File Exists? " & CStr(Dir(sPath & sName & ".shp") <> "")
  Else
    strCheckString = sPath & "\" & sName & ".shp"
'    MsgBox sPath & "\" & sName & ".shp" & vbCrLf & "File Exists? " & CStr(Dir(sPath & "\" & sName & ".shp") <> "")
  End If
  
'  booFileExists = (Dir(strCheckString) <> "")
'  MsgBox strCheckString & vbCrLf & CStr(booFileExists)
  
  If booFileExists Then
    MsgBox "The following file already exists:" & vbCrLf & vbCrLf & strCheckString & vbCrLf & vbCrLf & _
           "Please select a new filename...", , "Duplicate Filename:"
    Set CreateShapefileFeatureClass2 = Nothing
    Exit Function
  End If
  
  Dim pFeatClass As IFeatureClass
  Set pFeatClass = pFWS.CreateFeatureClass(sName, pFields, Nothing, _
                                           Nothing, esriFTSimple, "Shape", "")
                                           
  Set CreateShapefileFeatureClass2 = pFeatClass


  GoTo ClearMemory

ClearMemory:
  Set pGeomDef = Nothing
  Set pGeomDefEdit = Nothing
  Set pFWS = Nothing
  Set pWorkspaceFactory = Nothing
  Set pFields = Nothing
  Set pFieldsEdit = Nothing
  Set pField = Nothing
  Set pFieldEdit = Nothing
  Set pCheckField = Nothing
  Set pFeatClass = Nothing
End Function

  


Public Function CreateShapefileFeatureClass(sPath As String, sName As String, pSpatialReference As ISpatialReference, _
    pGeomType As esriGeometryType, Optional pAddFields As esriSystem.IVariantArray, _
    Optional booForceUniqueIDField As Boolean = True) As IFeatureClass       ' Don't include filename!
  
  If Right(sPath, 4) = ".shp" Then sPath = ReturnDir(sPath)
  If Right(sName, 4) = ".shp" Then sName = Left(sName, Len(sName) - 4)
  
  ' SET GEOMETRY TYPE, AND EXIT IF NOT ONE OF STANDARD OPTIONS
  Dim pGeomDef As IGeometryDef
  Dim pGeomDefEdit As IGeometryDefEdit
  Set pGeomDef = New GeometryDef
  Set pGeomDefEdit = pGeomDef
  With pGeomDefEdit
    .GeometryType = pGeomType
    Set .SpatialReference = pSpatialReference
  End With
  
  ' Open the folder to contain the shapefile as a workspace
  Dim pFWS As IFeatureWorkspace
  Dim pWorkspaceFactory As IWorkspaceFactory
  Set pWorkspaceFactory = New ShapefileWorkspaceFactory
  
  If Not ExistFileDir(sPath) Then
    MsgBox "Unable to create Feature Class:" & vbCrLf & _
           sPath & " is not a valid workspace...", , "Failed to Create Feature Class:"
    Set CreateShapefileFeatureClass = Nothing
    Exit Function
  End If
  
  Set pFWS = pWorkspaceFactory.OpenFromFile(sPath, 0)
  
'   Set up a simple fields collection
  Dim pFields As IFields
  Dim pFieldsEdit As IFieldsEdit
  Set pFields = New Fields
  Set pFieldsEdit = pFields

  Dim pField As iField
  Dim pFieldEdit As IFieldEdit
  
  ' Make the shape field
  ' it will need a geometry definition, with a spatial reference
  Set pField = New Field
  Set pFieldEdit = pField
  pFieldEdit.Name = "Shape"
  pFieldEdit.Type = esriFieldTypeGeometry

  Set pFieldEdit.GeometryDef = pGeomDef
  pFieldsEdit.AddField pField
  
  ' MAKE SURE UNIQUE ID FIELD IS UNIQUELY NAMED
 
  If booForceUniqueIDField Then
    Dim strUniqueName As String
    Dim strUniqueBase As String
    strUniqueName = "Unique_ID"
    strUniqueBase = strUniqueName
    Dim lngUniqueCounter As Long
    lngUniqueCounter = 1
    Dim lngCounterLength As Long
    Dim lngUniqueIndex As Long
    Dim booNameExists As Boolean
    Dim pCheckField As iField
    booNameExists = True
    If (pAddFields Is Nothing Or IsMissing(pAddFields)) Then
      booNameExists = False
    Else
      Do Until booNameExists = False
        booNameExists = False
        For lngUniqueIndex = 0 To pAddFields.Count - 1
          Set pCheckField = pAddFields.Element(lngUniqueIndex)
          If pCheckField.Name = strUniqueName Then
            booNameExists = True
            Exit For
          End If
        Next lngUniqueIndex
        If booNameExists Then
          lngUniqueCounter = lngUniqueCounter + 1
          lngCounterLength = Len(CStr(lngUniqueCounter))
          strUniqueName = Left(strUniqueBase, 10 - lngCounterLength) & CStr(lngUniqueCounter)
        End If
      Loop
    End If
    
    Set pField = New Field
    Set pFieldEdit = pField
    With pFieldEdit
        .Precision = 8
        .Name = strUniqueName
        .Type = esriFieldTypeInteger
    End With
    pFieldsEdit.AddField pField
  End If
  
  ' ADD FIELDS IF REQUESTED
'  Dim strDebugReport As String
'  Dim pDebugField As IField
  If Not (pAddFields Is Nothing Or IsMissing(pAddFields)) Then
    Dim lngIndex As Long
    For lngIndex = 0 To pAddFields.Count - 1
'      Set pDebugField = pAddFields.Element(lngIndex)
'      strDebugReport = strDebugReport & CStr(lngIndex) & "]  Field Name = " & pDebugField.Name & vbCrLf
      pFieldsEdit.AddField pAddFields.Element(lngIndex)
    Next lngIndex
  End If
'  MsgBox CStr(pGeomDef.GeometryType) & vbCrLf & strDebugReport
  ' Create the shapefile
  ' (some parameters apply to geodatabase options and can be defaulted as Nothing)
  Dim booFileExists As Boolean
  Dim strCheckString As String
  If Right(sPath, 1) = "\" Then
    strCheckString = sPath & sName & ".shp"
'    MsgBox sPath & sName & ".shp" & vbCrLf & "File Exists? " & CStr(Dir(sPath & sName & ".shp") <> "")
  Else
    strCheckString = sPath & "\" & sName & ".shp"
'    MsgBox sPath & "\" & sName & ".shp" & vbCrLf & "File Exists? " & CStr(Dir(sPath & "\" & sName & ".shp") <> "")
  End If
  
'  booFileExists = (Dir(strCheckString) <> "")
'  MsgBox strCheckString & vbCrLf & CStr(booFileExists)
  
  If booFileExists Then
    MsgBox "The following file already exists:" & vbCrLf & vbCrLf & strCheckString & vbCrLf & vbCrLf & _
           "Please select a new filename...", , "Duplicate Filename:"
    Set CreateShapefileFeatureClass = Nothing
    Exit Function
  End If
  
  Dim pFeatClass As IFeatureClass
  Set pFeatClass = pFWS.CreateFeatureClass(sName, pFields, Nothing, _
                                           Nothing, esriFTSimple, "Shape", "")
                                           
  Set CreateShapefileFeatureClass = pFeatClass


  GoTo ClearMemory
ClearMemory:
  Set pGeomDef = Nothing
  Set pGeomDefEdit = Nothing
  Set pFWS = Nothing
  Set pWorkspaceFactory = Nothing
  Set pFields = Nothing
  Set pFieldsEdit = Nothing
  Set pField = Nothing
  Set pFieldEdit = Nothing
  Set pCheckField = Nothing
  Set pFeatClass = Nothing
End Function

  
Public Function CreateDatasetFeatureClass(pFDS As IFeatureDataset, _
                                  strName As String, featType As esriFeatureType, _
                                  Optional geomType As esriGeometryType = esriGeometryPoint, _
                                  Optional pAddFields As esriSystem.IVariantArray, _
                                  Optional pCLSID As UID, _
                                  Optional pCLSEXT As UID, _
                                  Optional ConfigWord As String = "" _
                                  ) As IFeatureClass
  
'' createDatasetFeatureClass: simple helper to create a featureclass in a geodatabase Dataset.
'' NOTE: when creating a feature class in a dataset the spatial reference is inherited
'' from the dataset object
'' MODIFIED BY JENNESS APRIL 23 2008 FROM ESRI SAMPLE

  Dim pFields As IFields
  Dim pFieldsEdit As IFieldsEdit
  Dim pGeomDef As IGeometryDef
  Dim pGeomDefEdit As IGeometryDefEdit
  Dim pField As iField
  Dim pFieldEdit As IFieldEdit
  Dim strShapeFld As String
  Dim j As Integer
  
  Set CreateDatasetFeatureClass = Nothing
  If pFDS Is Nothing Then Exit Function
  If strName = "" Then Exit Function
  
  If (pCLSID Is Nothing) Or IsMissing(pCLSID) Then
    Set pCLSID = Nothing
    Set pCLSID = New UID
    
    '' determine the appropriate geometry type corresponding the the feature type
    Select Case featType
      Case esriFTSimple
        pCLSID.Value = "esriGeoDatabase.Feature"
      Case esriFTSimpleJunction
        geomType = esriGeometryPoint
        pCLSID.Value = "esriGeoDatabase.SimpleJunctionFeature"
      Case esriFTComplexJunction
        pCLSID.Value = "esriGeoDatabase.ComplexJunctionFeature"
      Case esriFTSimpleEdge
        geomType = esriGeometryPolyline
        pCLSID.Value = "esriGeoDatabase.SimpleEdgeFeature"
      Case esriFTComplexEdge
        geomType = esriGeometryPolyline
        pCLSID.Value = "esriGeoDatabase.ComplexEdgeFeature"
      Case esriFTAnnotation
        Exit Function
    End Select
  End If
  
  ' CREATE GEOMETRY AND OBJECT ID FIELD; ADD ANY REQUESTED FIELDS IF NECESSARY
  ' establish a fields collection
  Set pFields = New Fields
  Set pFieldsEdit = pFields
  
  '' create the geometry field
  Set pGeomDef = New GeometryDef
  Set pGeomDefEdit = pGeomDef
  
  '' assign the geometry definiton properties.
  With pGeomDefEdit
    .GeometryType = geomType
    .GridCount = 1
    .GridSize(0) = 10
    .AvgNumPoints = 2
    .HasM = False
    .HasZ = False
  End With
  
  Set pField = New Field
  Set pFieldEdit = pField
  
  pFieldEdit.Name = "Shape"
  pFieldEdit.AliasName = "geometry"
  pFieldEdit.Type = esriFieldTypeGeometry
  Set pFieldEdit.GeometryDef = pGeomDef
  pFieldsEdit.AddField pField

  '' create the object id field
  Set pField = New Field
  Set pFieldEdit = pField
  pFieldEdit.Name = "Object_ID"
  pFieldEdit.AliasName = "object identifier"
  pFieldEdit.Type = esriFieldTypeOID
  pFieldsEdit.AddField pField
  
  ' ADD FIELDS IF REQUESTED
  If Not (pAddFields Is Nothing Or IsMissing(pAddFields)) Then
    Dim lngIndex As Long
    For lngIndex = 0 To pAddFields.Count - 1
      pFieldsEdit.AddField pAddFields.Element(lngIndex)
    Next lngIndex
  End If
  
  ' establish the class extension
  If (pCLSEXT Is Nothing) Or IsMissing(pCLSEXT) Then
    Set pCLSEXT = Nothing
  End If
  
  ' locate the shape field
  For j = 0 To pFields.FieldCount - 1
    If pFields.Field(j).Type = esriFieldTypeGeometry Then
      strShapeFld = pFields.Field(j).Name
    End If
  Next
  
  Set CreateDatasetFeatureClass = pFDS.CreateFeatureClass(strName, pFields, pCLSID, pCLSEXT, featType, strShapeFld, ConfigWord)


  GoTo ClearMemory
ClearMemory:
  Set pFields = Nothing
  Set pFieldsEdit = Nothing
  Set pGeomDef = Nothing
  Set pGeomDefEdit = Nothing
  Set pField = Nothing
  Set pFieldEdit = Nothing

End Function
           


Public Function ReturnFieldsByType(pFields As IFields, enumJenFieldTypes As JenFieldTypes) As esriSystem.IVariantArray

  Dim booString As Boolean
  Dim booNumber As Boolean
  Dim booDate As Boolean
  Dim booOID As Boolean
  Dim booGeometry As Boolean
  Dim booBlob As Boolean
  Dim booRaster As Boolean
  Dim booGUID As Boolean
  Dim booGlobalID As Boolean
  Dim booXML As Boolean
  
  Dim strBinary As String
  strBinary = ConvertLongBinary(enumJenFieldTypes, 10)
'  MsgBox "Number = " & enumJenFieldTypes & vbCrLf & "Binary = " & strBinary

  booXML = Mid(strBinary, 1, 1) = "1"
  booGlobalID = Mid(strBinary, 2, 1) = "1"
  booGUID = Mid(strBinary, 3, 1) = "1"
  booRaster = Mid(strBinary, 4, 1) = "1"
  booBlob = Mid(strBinary, 5, 1) = "1"
  booGeometry = Mid(strBinary, 6, 1) = "1"
  booOID = Mid(strBinary, 7, 1) = "1"
  booDate = Mid(strBinary, 8, 1) = "1"
  booNumber = Mid(strBinary, 9, 1) = "1"
  booString = Mid(strBinary, 10, 1) = "1"

  
'  MsgBox "booString = " & booString & vbCrLf & _
      "booNumber = " & booNumber & vbCrLf & _
      "booDate = " & booDate & vbCrLf & _
      "booOID = " & booOID & vbCrLf & _
      "booGeometry = " & booGeometry & vbCrLf & _
      "booBlob = " & booBlob & vbCrLf & _
      "booRaster = " & booRaster & vbCrLf & _
      "booGUID = " & booGUID & vbCrLf & _
      "booGlobalID = " & booGlobalID & vbCrLf & _
      "booXML = " & booXML
  
  Dim pField As iField
  Dim lngIndex As Long
  Dim pType As esriFieldType
  
  Set ReturnFieldsByType = New esriSystem.varArray
    
  For lngIndex = 0 To pFields.FieldCount - 1
    Set pField = pFields.Field(lngIndex)
    pType = pField.Type
    If booXML And (pType = esriFieldTypeXML) Then
      ReturnFieldsByType.Add pField
    ElseIf booGlobalID And (pType = esriFieldTypeGlobalID) Then
      ReturnFieldsByType.Add pField
    ElseIf booGUID And (pType = esriFieldTypeGUID) Then
      ReturnFieldsByType.Add pField
    ElseIf booRaster And (pType = esriFieldTypeRaster) Then
      ReturnFieldsByType.Add pField
    ElseIf booBlob And (pType = esriFieldTypeBlob) Then
      ReturnFieldsByType.Add pField
    ElseIf booGeometry And (pType = esriFieldTypeGeometry) Then
      ReturnFieldsByType.Add pField
    ElseIf booOID And (pType = esriFieldTypeOID) Then
      ReturnFieldsByType.Add pField
    ElseIf booDate And (pType = esriFieldTypeDate) Then
      ReturnFieldsByType.Add pField
    ElseIf booNumber And ((pType = esriFieldTypeDouble) Or (pType = esriFieldTypeInteger) _
      Or (pType = esriFieldTypeSingle) Or (pType = esriFieldTypeSmallInteger)) Then
      ReturnFieldsByType.Add pField
    ElseIf booString And (pType = esriFieldTypeString) Then
      ReturnFieldsByType.Add pField
    End If
  Next lngIndex

  GoTo ClearMemory

ClearMemory:
  Set pField = Nothing
End Function
Public Function ReturnFieldsByType2(pFields As IFields, enumJenFieldTypes As JenFieldTypes) As esriSystem.IVariantArray

  Dim booString As Boolean
  Dim booNumber As Boolean
  Dim booDate As Boolean
  Dim booOID As Boolean
  Dim booGeometry As Boolean
  Dim booBlob As Boolean
  Dim booRaster As Boolean
  Dim booGUID As Boolean
  Dim booGlobalID As Boolean
  Dim booXML As Boolean
    
  booXML = (enumJenFieldTypes And enum_FieldXML) = enum_FieldXML
  booGlobalID = (enumJenFieldTypes And ENUM_FieldGlobalID) = ENUM_FieldGlobalID
  booGUID = (enumJenFieldTypes And ENUM_FieldGUID) = ENUM_FieldGUID
  booRaster = (enumJenFieldTypes And ENUM_FieldRaster) = ENUM_FieldRaster
  booBlob = (enumJenFieldTypes And ENUM_FieldBlob) = ENUM_FieldBlob
  booGeometry = (enumJenFieldTypes And ENUM_FieldGeometry) = ENUM_FieldGeometry
  booOID = (enumJenFieldTypes And enum_FieldOID) = enum_FieldOID
  booDate = (enumJenFieldTypes And enum_FieldDate) = enum_FieldDate
  booNumber = (enumJenFieldTypes And enum_FieldNumber) = enum_FieldNumber
  booString = (enumJenFieldTypes And enum_FieldString) = enum_FieldString
  
'  MsgBox "booString = " & booString & vbCrLf & _
      "booNumber = " & booNumber & vbCrLf & _
      "booDate = " & booDate & vbCrLf & _
      "booOID = " & booOID & vbCrLf & _
      "booGeometry = " & booGeometry & vbCrLf & _
      "booBlob = " & booBlob & vbCrLf & _
      "booRaster = " & booRaster & vbCrLf & _
      "booGUID = " & booGUID & vbCrLf & _
      "booGlobalID = " & booGlobalID & vbCrLf & _
      "booXML = " & booXML
  
  Dim pField As iField
  Dim lngIndex As Long
  Dim pType As esriFieldType
  
  Set ReturnFieldsByType2 = New esriSystem.varArray
    
  For lngIndex = 0 To pFields.FieldCount - 1
    Set pField = pFields.Field(lngIndex)
    pType = pField.Type
    If booXML And (pType = esriFieldTypeXML) Then
      ReturnFieldsByType2.Add pField
    ElseIf booGlobalID And (pType = esriFieldTypeGlobalID) Then
      ReturnFieldsByType2.Add pField
    ElseIf booGUID And (pType = esriFieldTypeGUID) Then
      ReturnFieldsByType2.Add pField
    ElseIf booRaster And (pType = esriFieldTypeRaster) Then
      ReturnFieldsByType2.Add pField
    ElseIf booBlob And (pType = esriFieldTypeBlob) Then
      ReturnFieldsByType2.Add pField
    ElseIf booGeometry And (pType = esriFieldTypeGeometry) Then
      ReturnFieldsByType2.Add pField
    ElseIf booOID And (pType = esriFieldTypeOID) Then
      ReturnFieldsByType2.Add pField
    ElseIf booDate And (pType = esriFieldTypeDate) Then
      ReturnFieldsByType2.Add pField
    ElseIf booNumber And ((pType = esriFieldTypeDouble) Or (pType = esriFieldTypeInteger) _
      Or (pType = esriFieldTypeSingle) Or (pType = esriFieldTypeSmallInteger)) Then
      ReturnFieldsByType2.Add pField
    ElseIf booString And (pType = esriFieldTypeString) Then
      ReturnFieldsByType2.Add pField
    End If
  Next lngIndex


  GoTo ClearMemory
ClearMemory:
  Set pField = Nothing

End Function
Public Function ReturnAcceptableFieldName(strOrigName As String, pFieldSet As IUnknown, Optional booRestrictToDBase As Boolean = False, _
    Optional booRestrictToPersonalGDB As Boolean = False, Optional booRestrictToCoverage As Boolean = False) As String
    
    Dim strName As String
    strName = strOrigName
    
  If booRestrictToDBase Then
    ReturnAcceptableFieldName = ReturnDBASEFieldName(strName, pFieldSet)
  Else
  
    Dim lngIndex As Long
    
    Dim strAcceptable As String
    strAcceptable = "abcdefghijklmnopqrstuvwxyz1234567890_"
    Dim strCharacters As String
    strCharacters = "abcdefghijklmnopqrstuvwxyz"
    
    Dim lngMaxLength As Long
    If booRestrictToCoverage Then
      lngMaxLength = 16
    ElseIf booRestrictToPersonalGDB Then
      lngMaxLength = 52
    Else
      lngMaxLength = 64
    End If
    
    Dim pField As iField
    
    Dim strNewName As String
    Dim strChar As String
    
    ' MAKE SURE SUGGESTED NAME IS VALID FOR dBASE IN GENERAL
    ' CHECK IF FIELD NAME IS EMPTY STRING
    If strName = "" Then strName = "z_Field"
    If StrComp(strName, "date", vbTextCompare) = 0 Then strName = "z_Date"
    If StrComp(strName, "day", vbTextCompare) = 0 Then strName = "z_Day"
    If StrComp(strName, "month", vbTextCompare) = 0 Then strName = "z_Month"
    If StrComp(strName, "table", vbTextCompare) = 0 Then strName = "z_Table"
    If StrComp(strName, "text", vbTextCompare) = 0 Then strName = "z_Text"
    If StrComp(strName, "user", vbTextCompare) = 0 Then strName = "z_User"
    If StrComp(strName, "when", vbTextCompare) = 0 Then strName = "z_When"
    If StrComp(strName, "where", vbTextCompare) = 0 Then strName = "z_Where"
    If StrComp(strName, "year", vbTextCompare) = 0 Then strName = "z_Year"
    If StrComp(strName, "zone", vbTextCompare) = 0 Then strName = "z_Zone"
    If StrComp(strName, "Shape_Length", vbTextCompare) = 0 Then strName = "z_Shape_Length"
    If StrComp(strName, "Shape_Area", vbTextCompare) = 0 Then strName = "z_Shape_Area"
    If StrComp(strName, "OID", vbTextCompare) = 0 Then strName = "z_OID"
    If StrComp(strName, "FID", vbTextCompare) = 0 Then strName = "z_FID"
    If StrComp(strName, "Object_ID", vbTextCompare) = 0 Then strName = "z_Object_ID"
    If StrComp(strName, "Shape", vbTextCompare) = 0 Then strName = "z_Shape"
    If StrComp(strName, "ObjectID", vbTextCompare) = 0 Then strName = "z_ObjectID"
    
    ' CHECK IF FIELD NAME IS PROPER LENGTH
    strName = Left(strName, lngMaxLength)
    
    ' CHECK IF FIELD NAME DOES NOT START WITH A LETTER
    If Not (InStr(1, strCharacters, Left(strName, 1), vbTextCompare) > 0) Then
      strName = Left("z" & strName, lngMaxLength)
    End If
    
    ' CHECK FOR NON_CONFORMING CHARACTERS
    strNewName = ""
    For lngIndex = 1 To Len(strName)
      strChar = Mid(strName, lngIndex, 1)
      If Not (InStr(1, strAcceptable, strChar, vbTextCompare) > 0) Then
        strNewName = strNewName & "_"
      Else
        strNewName = strNewName & strChar
      End If
    Next lngIndex
    strName = strNewName
    
    Dim strTestName As String
      
    ' MAKE SURE FIELD NAME DOES NOT ALREADY EXIST IN LIST
    ' CONVERT OBJECT INTO LIST OF EXISTING NAMES
    Dim pFieldArray As esriSystem.IStringArray
    Set pFieldArray = New esriSystem.strArray
  
    If Not pFieldSet Is Nothing Then
    
      If TypeOf pFieldSet Is IFields Then   ' <--------------------
        
        Dim pFields As IFields
        Set pFields = pFieldSet
        If pFields.FieldCount > 0 Then
          For lngIndex = 0 To pFields.FieldCount - 1
            strTestName = pFields.Field(lngIndex).Name
            pFieldArray.Add strTestName
          Next lngIndex
        End If
        
      ElseIf TypeOf pFieldSet Is esriSystem.IVariantArray Then
        
        Dim pVar As Variant
        Dim pVarArray As esriSystem.IVariantArray
        Set pVarArray = pFieldSet
        If pVarArray.Count > 0 Then
          For lngIndex = 0 To pVarArray.Count - 1
            If VarType(pVarArray.Element(lngIndex)) = vbDataObject Then           ' IS AN IField
              Set pField = pVarArray.Element(lngIndex)
              strTestName = pField.Name
            Else
              strTestName = pVarArray.Element(lngIndex)
            End If
            pFieldArray.Add strTestName
          Next lngIndex
        End If
      
      ElseIf TypeOf pFieldSet Is esriSystem.IStringArray Then
        
        Set pFieldArray = pFieldSet
      
      End If
    End If
    
    ' CHECK CURRENT NAME AGAINST LIST
    If pFieldArray.Count > 0 Then
      Dim lngCounter As Long
      Dim strBaseName As String
      strBaseName = strName
      Dim booFoundConflict As Boolean
      booFoundConflict = True
      Do Until booFoundConflict = False
        booFoundConflict = False
        For lngIndex = 0 To pFieldArray.Count - 1
          strTestName = pFieldArray.Element(lngIndex)
          If StrComp(strName, strTestName, vbTextCompare) = 0 Then
            booFoundConflict = True
            Exit For
          End If
        Next lngIndex
        If booFoundConflict Then
          lngCounter = lngCounter + 1
          strName = Left(strBaseName, lngMaxLength - Len(CStr(lngCounter))) & CStr(lngCounter)
        End If
      Loop
    End If
    
    ReturnAcceptableFieldName = strName
  
  End If
  GoTo ClearMemory
ClearMemory:
  Set pField = Nothing
  Set pFieldArray = Nothing
  Set pFields = Nothing
  pVar = Null
  Set pVarArray = Nothing

End Function

Public Function ReturnDBASEFieldName(strName As String, pFieldSet As Variant) As String

  Dim lngIndex As Long
  
  Dim strAcceptable As String
  strAcceptable = "abcdefghijklmnopqrstuvwxyz1234567890_"
  Dim strCharacters As String
  strCharacters = "abcdefghijklmnopqrstuvwxyz"
  
  Dim pField As iField
  
  Dim strNewName As String
  Dim strChar As String
  
  ' MAKE SURE SUGGESTED NAME IS VALID FOR dBASE IN GENERAL
  ' CHECK IF FIELD NAME IS EMPTY STRING
  If strName = "" Then strName = "z_Field"
  If StrComp(strName, "date", vbTextCompare) = 0 Then strName = "z_Date"
  If StrComp(strName, "day", vbTextCompare) = 0 Then strName = "z_Day"
  If StrComp(strName, "month", vbTextCompare) = 0 Then strName = "z_Month"
  If StrComp(strName, "table", vbTextCompare) = 0 Then strName = "z_Table"
  If StrComp(strName, "text", vbTextCompare) = 0 Then strName = "z_Text"
  If StrComp(strName, "user", vbTextCompare) = 0 Then strName = "z_User"
  If StrComp(strName, "when", vbTextCompare) = 0 Then strName = "z_When"
  If StrComp(strName, "where", vbTextCompare) = 0 Then strName = "z_Where"
  If StrComp(strName, "year", vbTextCompare) = 0 Then strName = "z_Year"
  If StrComp(strName, "zone", vbTextCompare) = 0 Then strName = "z_Zone"
  If StrComp(strName, "Shape_Length", vbTextCompare) = 0 Then strName = "zShape_Len"
  If StrComp(strName, "Shape_Area", vbTextCompare) = 0 Then strName = "zShapeArea"
  If StrComp(strName, "OID", vbTextCompare) = 0 Then strName = "z_OID"
  If StrComp(strName, "FID", vbTextCompare) = 0 Then strName = "z_FID"
  If StrComp(strName, "Object_ID", vbTextCompare) = 0 Then strName = "z_ObjectID"
  If StrComp(strName, "Shape", vbTextCompare) = 0 Then strName = "z_Shape"
  If StrComp(strName, "ObjectID", vbTextCompare) = 0 Then strName = "z_ObjectID"
  
  ' CHECK IF FIELD NAME IS PROPER LENGTH
  strName = Left(strName, 10)
  
  ' CHECK IF FIELD NAME DOES NOT START WITH A LETTER
  If Not (InStr(1, strCharacters, Left(strName, 1), vbTextCompare) > 0) Then
    strName = Left("z" & strName, 10)
  End If
  
  ' CHECK FOR NON_CONFORMING CHARACTERS
  strNewName = ""
  For lngIndex = 1 To Len(strName)
    strChar = Mid(strName, lngIndex, 1)
    If Not (InStr(1, strAcceptable, strChar, vbTextCompare) > 0) Then
      strNewName = strNewName & "_"
    Else
      strNewName = strNewName & strChar
    End If
  Next lngIndex
  strName = strNewName
  
  Dim strTestName As String
    
  ' MAKE SURE FIELD NAME DOES NOT ALREADY EXIST IN LIST
  ' CONVERT OBJECT INTO LIST OF EXISTING NAMES
  Dim pFieldArray As esriSystem.IStringArray
  Set pFieldArray = New esriSystem.strArray
  If Not pFieldSet Is Nothing Then
    If TypeOf pFieldSet Is IFields Then
      
      Dim pFields As IFields
      Set pFields = pFieldSet
      If pFields.FieldCount > 0 Then
        For lngIndex = 0 To pFields.FieldCount - 1
          strTestName = pFields.Field(lngIndex).Name
          pFieldArray.Add strTestName
        Next lngIndex
      End If
      
    ElseIf TypeOf pFieldSet Is esriSystem.IVariantArray Then
      
      Dim pVar As Variant
      Dim pVarArray As esriSystem.IVariantArray
      Set pVarArray = pFieldSet
      If pVarArray.Count > 0 Then
        For lngIndex = 0 To pVarArray.Count - 1
          If VarType(pVarArray.Element(lngIndex)) = vbDataObject Then           ' IS AN IField
            Set pField = pVarArray.Element(lngIndex)
            strTestName = pField.Name
          Else
            strTestName = pVarArray.Element(lngIndex)
          End If
          pFieldArray.Add strTestName
        Next lngIndex
      End If
    
    ElseIf TypeOf pFieldSet Is esriSystem.IStringArray Then
      
      Set pFieldArray = pFieldSet
    
    End If
  End If
  
  ' CHECK CURRENT NAME AGAINST LIST
  If pFieldArray.Count > 0 Then
    Dim lngCounter As Long
    Dim strBaseName As String
    strBaseName = strName
    Dim booFoundConflict As Boolean
    booFoundConflict = True
    Do Until booFoundConflict = False
      booFoundConflict = False
      For lngIndex = 0 To pFieldArray.Count - 1
        strTestName = pFieldArray.Element(lngIndex)
        If StrComp(strName, strTestName, vbTextCompare) = 0 Then
          booFoundConflict = True
          Exit For
        End If
      Next lngIndex
      If booFoundConflict Then
        lngCounter = lngCounter + 1
        strName = Left(strBaseName, 10 - Len(CStr(lngCounter))) & CStr(lngCounter)
      End If
    Loop
  End If
  
  ReturnDBASEFieldName = strName

  GoTo ClearMemory

ClearMemory:
  Set pField = Nothing
  Set pFieldArray = Nothing
  Set pFields = Nothing
  pVar = Null
  Set pVarArray = Nothing

End Function

Public Function CheckIfFieldNameExists(strName As String, pFieldSet As Variant) As Boolean

  Dim lngIndex As Long
  Dim strTestName As String
  Dim pFields As IFields
  Dim pField As iField
  Dim pVarArray As IVariantArray
  Dim pVar As Variant
  Dim pTable As ITable
  Dim pFClass As IFeatureClass
  
  ' MAKE SURE FIELD NAME DOES NOT ALREADY EXIST IN LIST
  ' CONVERT OBJECT INTO LIST OF EXISTING NAMES
  Dim pFieldArray As esriSystem.IStringArray
  Set pFieldArray = New esriSystem.strArray
  
  If TypeOf pFieldSet Is IFields Then
    
    Set pFields = pFieldSet
    If pFields.FieldCount > 0 Then
      For lngIndex = 0 To pFields.FieldCount - 1
        strTestName = pFields.Field(lngIndex).Name
        pFieldArray.Add strTestName
      Next lngIndex
    End If
  
  ElseIf TypeOf pFieldSet Is ITable Then
  
    Set pTable = pFieldSet
    Set pFields = pTable.Fields
    If pFields.FieldCount > 0 Then
      For lngIndex = 0 To pFields.FieldCount - 1
        strTestName = pFields.Field(lngIndex).Name
        pFieldArray.Add strTestName
      Next lngIndex
    End If
  
  ElseIf TypeOf pFieldSet Is IFeatureClass Then
  
    Set pFClass = pFieldSet
    Set pFields = pFClass.Fields
    If pFields.FieldCount > 0 Then
      For lngIndex = 0 To pFields.FieldCount - 1
        strTestName = pFields.Field(lngIndex).Name
        pFieldArray.Add strTestName
      Next lngIndex
    End If
    
  ElseIf TypeOf pFieldSet Is esriSystem.IVariantArray Then
    
    Set pVarArray = pFieldSet
    If pVarArray.Count > 0 Then
      For lngIndex = 0 To pVarArray.Count - 1
        If VarType(pVarArray.Element(lngIndex)) = vbDataObject Then           ' IS AN IField
          Set pField = pVarArray.Element(lngIndex)
          strTestName = pField.Name
        Else
          strTestName = pVarArray.Element(lngIndex)
        End If
        pFieldArray.Add strTestName
      Next lngIndex
    End If
  
  ElseIf TypeOf pFieldSet Is esriSystem.IStringArray Then
    
    Set pFieldArray = pFieldSet
  
  End If
  
  ' CHECK CURRENT NAME AGAINST LIST
  CheckIfFieldNameExists = False
  If pFieldArray.Count > 0 Then
    For lngIndex = 0 To pFieldArray.Count - 1
      strTestName = pFieldArray.Element(lngIndex)
      If StrComp(strName, strTestName, vbTextCompare) = 0 Then
        CheckIfFieldNameExists = True
        Exit For
      End If
    Next lngIndex
  End If
  
  GoTo ClearMemory
ClearMemory:
  Set pFields = Nothing
  Set pField = Nothing
  Set pVarArray = Nothing
  pVar = Null
  Set pTable = Nothing
  Set pFClass = Nothing
  Set pFieldArray = Nothing

End Function



Public Function ReturnLayerByName(strName As String, pMap As IMap) As ILayer

  Set ReturnLayerByName = Nothing
  Dim pLayers As IEnumLayer
  Set pLayers = pMap.Layers(Nothing, True)
  Dim pLayer As ILayer
  pLayers.Reset
  Set pLayer = pLayers.Next
  Do Until pLayer Is Nothing
    If StrComp(pLayer.Name, strName, vbTextCompare) = 0 Then
      Set ReturnLayerByName = pLayer
      Exit Do
    End If
    Set pLayer = pLayers.Next
  Loop
  
  GoTo ClearMemory

ClearMemory:
  Set pLayers = Nothing
  Set pLayer = Nothing

End Function
Public Function ReturnMapByName(strName As String, pMxDoc As IMxDocument) As IMap

  Set ReturnMapByName = Nothing
  Dim pMaps As IMaps
  Set pMaps = pMxDoc.Maps
  Dim lngIndex As Long
  Dim pMap As IMap
  For lngIndex = 0 To pMaps.Count - 1
    Set pMap = pMaps.Item(lngIndex)
    If StrComp(pMap.Name, strName, vbTextCompare) = 0 Then
      Set ReturnMapByName = pMap
      Exit For
    End If
  Next lngIndex
  
  GoTo ClearMemory
ClearMemory:
  Set pMaps = Nothing
  Set pMap = Nothing
End Function

Public Function CheckIfCompressedFGDB(pFeatureClass As IFeatureClass) As Boolean
  ' NOTE:  WILL ONLY RETURN TRUE IF EVERYTHING WORKS; IF FUNCTION CRASHES, THEN
  ' ASSUME THIS IS NOT A FILE GEODATABASE AND THEREFORE CANNOT BE A COMPRESSED FGDB.
  ' FOR EXAMPLE, APPARENTLY PERSONAL GEODATABASES AND SHAPEFILES DON'T HAVE PROPERTY SETS???
  
  On Error GoTo EH
  
  Dim pDataset As IDataset
  Set pDataset = pFeatureClass
  Dim pGeoDataset As IGeoDataset
  Set pGeoDataset = pFeatureClass
  Dim strCategory As String
  strCategory = pDataset.Category
  Dim booIsFGDB As Boolean
  booIsFGDB = InStr(1, strCategory, "file geodatabase", vbTextCompare)
  
  CheckIfCompressedFGDB = False
  
  If booIsFGDB Then      ' IF FILE GEODATABASE, CHECK IF IT IS COMPRESSED
    Dim pPropSet As IPropertySet
    Set pPropSet = pDataset.PropertySet
    Dim varFormat As Variant
    varFormat = pPropSet.GetProperty("Datafile Format")
    If varFormat = 1 Then
      CheckIfCompressedFGDB = True
    Else
      CheckIfCompressedFGDB = False
    End If
  End If
  
  GoTo ClearMemory
  Exit Function
EH:
  CheckIfCompressedFGDB = False


ClearMemory:
  Set pDataset = Nothing
  Set pGeoDataset = Nothing
  Set pPropSet = Nothing
  varFormat = Null

End Function

Public Function TrueLayerCount(pMap As IMap) As Long

  Dim lngLayerCount As Long
  lngLayerCount = 0
  
  If pMap.LayerCount > 0 Then
    Dim pEnumLayer As IEnumLayer
    Set pEnumLayer = pMap.Layers
    pEnumLayer.Reset
    Dim pLayer As IUnknown
    Set pLayer = pEnumLayer.Next
    Do Until pLayer Is Nothing
      lngLayerCount = lngLayerCount + 1
      Set pLayer = pEnumLayer.Next
    Loop
  End If
  
  TrueLayerCount = lngLayerCount


  GoTo ClearMemory
ClearMemory:
  Set pEnumLayer = Nothing
  Set pLayer = Nothing

End Function


Public Function ReturnDistanceUnitsName(lngEsriUnits As esriUnits) As String

  Select Case lngEsriUnits
    Case 0
      ReturnDistanceUnitsName = "Unknown"
    Case 1
      ReturnDistanceUnitsName = "Inches"
    Case 2
      ReturnDistanceUnitsName = "Points"
    Case 3
      ReturnDistanceUnitsName = "Feet"
    Case 4
      ReturnDistanceUnitsName = "Yards"
    Case 5
      ReturnDistanceUnitsName = "Miles"
    Case 6
      ReturnDistanceUnitsName = "Nautical miles"
    Case 7
      ReturnDistanceUnitsName = "Millimeters"
    Case 8
      ReturnDistanceUnitsName = "Centimeters"
    Case 9
      ReturnDistanceUnitsName = "Meters"
    Case 10
      ReturnDistanceUnitsName = "Kilometers"
    Case 11
      ReturnDistanceUnitsName = "Decimal degrees"
    Case 12
      ReturnDistanceUnitsName = "Decimeters"
  End Select

End Function

Public Function ConvertLongBinary(lngNumber As Long, Optional MinimumNumDigits As Long = -1) As String

  Dim strBinary As String
  Dim lngIntermediate As Long
  Dim lngRemainder As Long
  lngIntermediate = lngNumber
  Do Until lngIntermediate = 1
    lngRemainder = lngIntermediate Mod 2
    lngIntermediate = Int(lngIntermediate / 2)
    strBinary = CStr(lngRemainder) & strBinary
  Loop
  strBinary = "1" & strBinary
  
  If Len(strBinary) < MinimumNumDigits Then
    Do While Len(strBinary) < MinimumNumDigits
      strBinary = "0" & strBinary
    Loop
  End If
  
  ConvertLongBinary = strBinary

End Function


Public Function ReturnFilesFromNestedFolders3(ByVal strDir As String, strAnyTextInName As String) As String()
  
'  Set ReturnFilesFromNestedFolders3 = New esriSystem.strArray
  
  Dim strReturn() As String
  
  If Right(strDir, 1) <> "\" Then strDir = strDir & "\"
  
  Dim strOriginalDir As String
  strOriginalDir = strDir
  
  Dim booFoundSubFolders As Boolean
  
  Dim strPathArray() As String
  Dim lngPathArrayIndex As Long
  
'  Dim pPathArray As esriSystem.IStringArray
'  Set pPathArray = New esriSystem.strArray
  
  Dim strFinalArray() As String
  Dim lngFinalArrayIndex As Long
'  Dim pFinalArray As esriSystem.IStringArray
'  Set pFinalArray = New esriSystem.strArray
  Dim pCheckColl As Collection
  Set pCheckColl = New Collection
  
'  pFinalArray.Add strDir
  lngFinalArrayIndex = 0
  ReDim Preserve strFinalArray(lngFinalArrayIndex)
  strFinalArray(lngFinalArrayIndex) = strDir
  pCheckColl.Add True, strDir
  
  Dim strDirName As String
  lngPathArrayIndex = -1
  strDirName = Dir(strDir, vbDirectory)   ' Retrieve the first entry.
  Do While strDirName <> ""   ' Start the loop.
     ' Ignore the current directory and the encompassing directory.
     If strDirName <> "." And strDirName <> ".." Then
        ' Use bitwise comparison to make sure MyName is a directory.
        If IsFolder_FalseIfCrash((strDir & strDirName)) Then
'        If (GetAttr(strDir & strDirName) And vbDirectory) = vbDirectory Then
'           pPathArray.Add strDir & strDirName & "\"
'           pFinalArray.Add strDir & strDirName & "\"
           lngPathArrayIndex = lngPathArrayIndex + 1
           lngFinalArrayIndex = lngFinalArrayIndex + 1
           ReDim Preserve strPathArray(lngPathArrayIndex)
           ReDim Preserve strFinalArray(lngFinalArrayIndex)
           strPathArray(lngPathArrayIndex) = strDir & strDirName & "\"
           strFinalArray(lngPathArrayIndex) = strDir & strDirName & "\"
           pCheckColl.Add False, strDir & strDirName & "\"
        End If   ' it represents a directory.
     End If
     strDirName = Dir   ' Get next entry.
  Loop
  
  booFoundSubFolders = pPathArray.Count > 0
  ' If Not booFoundSubFolders Then Exit Function
  
  Dim strSubFolder As String
  
  Dim booFoundSubHere As Boolean
  Dim strSubArray() As String
  Dim lngSubArrayIndex As Long
'  Dim pSubArray As esriSystem.IStringArray
  
  Dim lngIndex As Long
  
  Do While booFoundSubFolders
    booFoundSubFolders = False
    'Set pSubArray = New esriSystem.strArray
    lngSubArrayIndex = 0
    ReDim strSubArray(lngSubArrayIndex)
'    For lngIndex = 0 To pPathArray.Count - 1
    For lngIndex = 0 To UBound(strPathArray)
'      strSubFolder = pPathArray.Element(lngIndex)
      strSubFolder = strPathArray(lngIndex)

'     If strDirName <> "." And strDirName <> ".." Then
'        lngCounter = lngCounter + 1

      booFoundSubHere = False
      strDirName = Dir(strSubFolder, vbDirectory)   ' Retrieve the first entry.
      Do While strDirName <> ""   ' Start the loop.
         ' Ignore the current directory and the encompassing directory.
         If strDirName <> "." And strDirName <> ".." Then
            ' Use bitwise comparison to make sure MyName is a directory.
            If IsFolder_FalseIfCrash(strSubFolder & strDirName) Then
'            If (GetAttr(strSubFolder & strDirName) And vbDirectory) = vbDirectory Then
'               pSubArray.Add strSubFolder & strDirName & "\"
               lngSubArrayIndex = lngSubArrayIndex + 1
               ReDim Preserve strSubArray(lngSubArrayIndex)
               strSubArray(lngSubArrayIndex) = strSubFolder & strDirName & "\"
               booFoundSubFolders = True
               booFoundSubHere = True
               If Not CheckCollectionForKey(pCheckColl, strSubFolder & strDirName & "\") Then
                 pCheckColl.Add 1, strSubFolder & strDirName & "\"
                 ' pFinalArray.Add strSubFolder & strDirName & "\"
                 
                lngFinalArrayIndex = lngFinalArrayIndex + 1
                ReDim Preserve strFinalArray(lngFinalArrayIndex)
                strFinalArray(lngPathArrayIndex) = strSubFolder & strDirName & "\"
                 
               End If
            End If   ' it represents a directory.
         End If
         strDirName = Dir   ' Get next entry.
      Loop
      
      If Not booFoundSubHere Then
'        pSubArray.Add strSubFolder
        lngSubArrayIndex = lngSubArrayIndex + 1
        ReDim Preserve strSubArray(lngSubArrayIndex)
        strSubArray(lngSubArrayIndex) = strSubFolder
      End If
      
    Next lngIndex
    
    If booFoundSubFolders Then
'      Set pPathArray = pSubArray
      strPathArray = strSubArray
    End If
    
  Loop
  
  Dim strFolders() As String
'  ReDim strFolders(pFinalArray.Count - 1)
  ReDim strFolders(UBound(strFinalArray))
  
  
  For lngIndex = 0 To UBound(strFinalArray)
    strDir = pFinalArray.Element(lngIndex)
    strFolders(lngIndex) = strDir
'    Debug.Print CStr(lngIndex + 1) & "]  " & strDir
  Next lngIndex
  
  QuickSort.StringsAscending strFolders, 0, pFinalArray.Count - 1
  
'  For lngIndex = 0 To UBound(strFolders)
'    strDir = strFolders(lngIndex)
'    Debug.Print CStr(lngIndex + 1) & "]  " & strDir
'  Next lngIndex

  Dim lngCounter As Long
  lngCounter = 0

'  Debug.Print
  
  Dim pFolderFeatLayers As esriSystem.IVariantArray
  Dim pFilenames As esriSystem.IStringArray
  Set pFolderFeatLayers = New esriSystem.varArray
  Set pFilenames = New esriSystem.strArray
  Dim strDirAndFile As String
  
  For lngIndex = 0 To UBound(strFolders)
    strDir = strFolders(lngIndex)
    strDirName = Dir(strDir, vbNormal)   ' Retrieve the first entry.
    lngCounter = lngCounter + 1

    Do While strDirName <> ""   ' Start the loop.
       ' Ignore the current directory and the encompassing directory.
       If strDirName <> "." And strDirName <> ".." Then
          strDirAndFile = strDir & strDirName

          If IsNormal_FalseIfCrash(strDirAndFile) Then
'          If (GetAttr(strDirAndFile) And vbNormal) = vbNormal Then
            'If StrComp(Right(strDirAndFile, Len(strExtensionWithDot)), strExtensionWithDot, vbTextCompare) = 0 Then
            If InStr(1, strDirName, strAnyTextInName, vbTextCompare) > 0 Then
'             Debug.Print "Examining Folder #" & CStr(lngCounter) & ":  " & strDirName
              pFilenames.Add strDirAndFile
'             Debug.Print "  --> " & pDataset.BrowseName
            End If
          End If
       End If
       strDirName = Dir   ' Get next entry.
    Loop
  Next lngIndex

  ' CONFIRM THAT CORRECT DIRECTORY HAS BEEN SELECTED AND THAT IT ACTUALLY HAS POLYLINE SHAPEFILES
  Dim strReturn() As String
  ReDim strReturn(pFilenames.Count - 1)
  For lngIndex = 0 To pFilenames.Count - 1
    strReturn(lngIndex) = pFilenames.Element(lngIndex)
  Next lngIndex
  
  ReturnFilesFromNestedFolders = strReturn
  
  GoTo ClearMemory

ClearMemory:
  Set pPathArray = Nothing
  Set pFinalArray = Nothing
  Set pCheckColl = Nothing
  Set pSubArray = Nothing
  Erase strFolders
  Set pFolderFeatLayers = Nothing
  Set pFilenames = Nothing

End Function






Public Function ReturnLayersByType(pFocusMap As IMap, enumLayerTypes As JenLayerTypes) As esriSystem.IVariantArray

  Dim booFeatureLayers As Boolean
  Dim booRasterLayers As Boolean
  Dim booStandaloneTables As Boolean
  Dim booPointLayers As Boolean
  Dim booPolylineLayers As Boolean
  Dim booPolygonLayers As Boolean
  Dim booMultipointLayers As Boolean
  Dim booTinLayers As Boolean
  Dim booTerrainLayers As Boolean
  Dim booRasterCatalogLayers As Boolean
  
  Dim strBinary As String
  strBinary = ConvertLongBinary(enumLayerTypes, 9)
'  MsgBox "Number = " & enumLayerTypes & vbCrLf & "Binary = " & strBinary
  booFeatureLayers = Mid(strBinary, 9, 1) = "1"
  booRasterLayers = Mid(strBinary, 8, 1) = "1"
  booStandaloneTables = Mid(strBinary, 7, 1) = "1"
  booPointLayers = Mid(strBinary, 6, 1) = "1"
  booPolylineLayers = Mid(strBinary, 5, 1) = "1"
  booPolygonLayers = Mid(strBinary, 4, 1) = "1"
  booMultipointLayers = Mid(strBinary, 3, 1) = "1"
  booTinLayers = Mid(strBinary, 2, 1) = "1"
  booRasterCatalogLayers = Mid(strBinary, 1, 1) = "1"
  
'  MsgBox "booFeatureLayers = " & booFeatureLayers & vbCrLf & _
'      "booRasterLayers = " & booRasterLayers & vbCrLf & _
'      "booStandaloneTables = " & booStandaloneTables & vbCrLf & _
'      "booPointLayers = " & booPointLayers & vbCrLf & _
'      "booPolylineLayers = " & booPolylineLayers & vbCrLf & _
'      "booPolygonLayers = " & booPolygonLayers & vbCrLf & _
'      "booMultipointLayers = " & booMultipointLayers & vbCrLf & _
'      "booTinLayers = " & booTinLayers

  Dim pEnumLayer As IEnumLayer
  Dim pFeatureLayer As IFeatureLayer
  Dim pLayer As IUnknown
  Dim pFeatureClass As IFeatureClass
  Dim pGeometryType As esriGeometryType
  Dim pFeatureLayerForValid As IFeatureLayer
  Dim booOpenDialog As Boolean
  booOpenDialog = False
  Dim pRasterLayer As IRasterLayer
'  Dim pAsLayer As ILayer
    
  Set ReturnLayersByType = New esriSystem.varArray
    
  If (pFocusMap.LayerCount > 0) Then
    Set pEnumLayer = pFocusMap.Layers(, True)
    pEnumLayer.Reset
    
    Set pLayer = pEnumLayer.Next
    Do Until pLayer Is Nothing
'      Set pAsLayer = pLayer
'      MsgBox "Layer Name = " & pAsLayer.Name & vbCrLf & _
'             "Feature Layer:  " & CStr(TypeOf pLayer Is IFeatureLayer) & vbCrLf & _
'             "Raster Layer:  " & CStr(TypeOf pLayer Is IRasterLayer) & vbCrLf & _
'             "TIN Layer:  " & CStr(TypeOf pLayer Is ITinLayer)
             
      If TypeOf pLayer Is IGdbRasterCatalogLayer Then
        Set pFeatureLayerForValid = pLayer
        If pFeatureLayerForValid.Valid Then
          If booRasterCatalogLayers Then ReturnLayersByType.Add pLayer
        End If
            
      ElseIf TypeOf pLayer Is IFeatureLayer Then
        Set pFeatureLayerForValid = pLayer
        ' CHECK IF FEATURE LAYER IS VALID
        If pFeatureLayerForValid.Valid Then
          ' CHECK IF POLYGON LAYER
          Set pFeatureClass = pFeatureLayerForValid.FeatureClass          ' CHECK IF FEATURE LAYER
          pGeometryType = pFeatureClass.ShapeType
          If booFeatureLayers Then
            ReturnLayersByType.Add pLayer
          Else
            If (pGeometryType = esriGeometryPolygon) Then                 ' CHECK IF POLYGON LAYER
              If booPolygonLayers Then ReturnLayersByType.Add pLayer
            ElseIf pGeometryType = esriGeometryPolyline Then              ' CHECK IF POLYLINE LAYER
              If booPolylineLayers Then ReturnLayersByType.Add pLayer
            ElseIf pGeometryType = esriGeometryPoint Then                 ' CHECK IF POINT LAYER
              If booPointLayers Then ReturnLayersByType.Add pLayer
            ElseIf pGeometryType = esriGeometryMultipoint Then            ' CHECK IF MULTIPOINT LAYER
              If booMultipointLayers Then ReturnLayersByType.Add pLayer
            End If
          End If
        End If
      ElseIf TypeOf pLayer Is IRasterLayer Then                           ' CHECK IF RASTER LAYER
        Set pRasterLayer = pLayer
        If pRasterLayer.Valid Then
          If booRasterLayers Then ReturnLayersByType.Add pLayer
        End If
      ElseIf TypeOf pLayer Is ITinLayer Then                              ' CHECK IF TIN LAYER
        Dim pTinLayer As ITinLayer
        Set pTinLayer = pLayer
        If pTinLayer.Valid Then
          If booTinLayers Then ReturnLayersByType.Add pLayer
        End If
      End If
      Set pLayer = pEnumLayer.Next
    Loop
  End If
  
  If booStandaloneTables Then
    Dim pSTCollection As IStandaloneTableCollection
    Set pSTCollection = pFocusMap
    Dim pStTble As IStandaloneTable
    If pSTCollection.StandaloneTableCount > 0 Then
      Dim lngIndex As Long
      For lngIndex = 0 To pSTCollection.StandaloneTableCount - 1
        Set pStTble = pSTCollection.StandaloneTable(lngIndex)
        If booStandaloneTables Then ReturnLayersByType.Add pStTble
      Next lngIndex
    End If
  End If
  
'  Dim strReport As String
'  For lngIndex = 0 To ReturnLayersByType.Count - 1
'    Set pAsLayer = ReturnLayersByType.Element(lngIndex)
'    strReport = strReport & CStr(lngIndex) & "] " & pAsLayer.Name & vbCrLf
'  Next lngIndex
'  MsgBox strReport


  GoTo ClearMemory
ClearMemory:
  Set pEnumLayer = Nothing
  Set pFeatureLayer = Nothing
  Set pLayer = Nothing
  Set pFeatureClass = Nothing
  Set pFeatureLayerForValid = Nothing
  Set pRasterLayer = Nothing
  Set pTinLayer = Nothing
  Set pSTCollection = Nothing
  Set pStTble = Nothing

End Function

Public Function ReturnLayersByType2(pFocusMap As IMap, enumLayerTypes As JenLayerTypes, _
    Optional booIncludeInvalidLayers As Boolean = False) As esriSystem.IVariantArray
  
  ' INCLUDES OPTION TO RETURN INVALID LAYERS

  Dim booFeatureLayers As Boolean
  Dim booRasterLayers As Boolean
  Dim booStandaloneTables As Boolean
  Dim booPointLayers As Boolean
  Dim booPolylineLayers As Boolean
  Dim booPolygonLayers As Boolean
  Dim booMultipointLayers As Boolean
  Dim booTinLayers As Boolean
  Dim booTerrainLayers As Boolean
  Dim booRasterCatalogLayers As Boolean
  
  Dim strBinary As String
  strBinary = ConvertLongBinary(enumLayerTypes, 9)
'  MsgBox "Number = " & enumLayerTypes & vbCrLf & "Binary = " & strBinary
  booFeatureLayers = Mid(strBinary, 9, 1) = "1"
  booRasterLayers = Mid(strBinary, 8, 1) = "1"
  booStandaloneTables = Mid(strBinary, 7, 1) = "1"
  booPointLayers = Mid(strBinary, 6, 1) = "1"
  booPolylineLayers = Mid(strBinary, 5, 1) = "1"
  booPolygonLayers = Mid(strBinary, 4, 1) = "1"
  booMultipointLayers = Mid(strBinary, 3, 1) = "1"
  booTinLayers = Mid(strBinary, 2, 1) = "1"
  booRasterCatalogLayers = Mid(strBinary, 1, 1) = "1"
  
'  MsgBox "booFeatureLayers = " & booFeatureLayers & vbCrLf & _
'      "booRasterLayers = " & booRasterLayers & vbCrLf & _
'      "booStandaloneTables = " & booStandaloneTables & vbCrLf & _
'      "booPointLayers = " & booPointLayers & vbCrLf & _
'      "booPolylineLayers = " & booPolylineLayers & vbCrLf & _
'      "booPolygonLayers = " & booPolygonLayers & vbCrLf & _
'      "booMultipointLayers = " & booMultipointLayers & vbCrLf & _
'      "booTinLayers = " & booTinLayers

  Dim pEnumLayer As IEnumLayer
  Dim pFeatureLayer As IFeatureLayer
  Dim pLayer As IUnknown
  Dim pFeatureClass As IFeatureClass
  Dim pGeometryType As esriGeometryType
  Dim pFeatureLayerForValid As IFeatureLayer
  Dim booOpenDialog As Boolean
  booOpenDialog = False
  Dim pRasterLayer As IRasterLayer
'  Dim pAsLayer As ILayer
    
  Set ReturnLayersByType2 = New esriSystem.varArray
    
  If (pFocusMap.LayerCount > 0) Then
    Set pEnumLayer = pFocusMap.Layers(, True)
    pEnumLayer.Reset
    
    Set pLayer = pEnumLayer.Next
    Do Until pLayer Is Nothing
'      Set pAsLayer = pLayer
'      MsgBox "Layer Name = " & pAsLayer.Name & vbCrLf & _
'             "Feature Layer:  " & CStr(TypeOf pLayer Is IFeatureLayer) & vbCrLf & _
'             "Raster Layer:  " & CStr(TypeOf pLayer Is IRasterLayer) & vbCrLf & _
'             "TIN Layer:  " & CStr(TypeOf pLayer Is ITinLayer)
             
      If TypeOf pLayer Is IGdbRasterCatalogLayer Then
        Set pFeatureLayerForValid = pLayer
        If pFeatureLayerForValid.Valid Or booIncludeInvalidLayers Then
          If booRasterCatalogLayers Then ReturnLayersByType2.Add pLayer
        End If
            
      ElseIf TypeOf pLayer Is IFeatureLayer Then
        Set pFeatureLayerForValid = pLayer
        ' CHECK IF FEATURE LAYER IS VALID
        If pFeatureLayerForValid.Valid Or booIncludeInvalidLayers Then
          ReturnLayersByType2.Add pLayer
'          ' CHECK IF POLYGON LAYER
'          Set pFeatureClass = pFeatureLayerForValid.FeatureClass          ' CHECK IF FEATURE LAYER
'          pGeometryType = pFeatureClass.ShapeType
'          If booFeatureLayers Then
'            ReturnLayersByType2.Add pLayer
'          Else
'            If (pGeometryType = esriGeometryPolygon) Then                 ' CHECK IF POLYGON LAYER
'              If booPolygonLayers Then ReturnLayersByType2.Add pLayer
'            ElseIf pGeometryType = esriGeometryPolyline Then              ' CHECK IF POLYLINE LAYER
'              If booPolylineLayers Then ReturnLayersByType2.Add pLayer
'            ElseIf pGeometryType = esriGeometryPoint Then                 ' CHECK IF POINT LAYER
'              If booPointLayers Then ReturnLayersByType2.Add pLayer
'            ElseIf pGeometryType = esriGeometryMultipoint Then            ' CHECK IF MULTIPOINT LAYER
'              If booMultipointLayers Then ReturnLayersByType2.Add pLayer
'            End If
'          End If
        End If
      ElseIf TypeOf pLayer Is IRasterLayer Then                           ' CHECK IF RASTER LAYER
        Set pRasterLayer = pLayer
        If pRasterLayer.Valid Or booIncludeInvalidLayers Then
          If booRasterLayers Then ReturnLayersByType2.Add pLayer
        End If
      ElseIf TypeOf pLayer Is ITinLayer Then                              ' CHECK IF TIN LAYER
        Dim pTinLayer As ITinLayer
        Set pTinLayer = pLayer
        If pTinLayer.Valid Or booIncludeInvalidLayers Then
          If booTinLayers Then ReturnLayersByType2.Add pLayer
        End If
      End If
      Set pLayer = pEnumLayer.Next
    Loop
  End If
  
  If booStandaloneTables Then
    Dim pSTCollection As IStandaloneTableCollection
    Set pSTCollection = pFocusMap
    Dim pStTble As IStandaloneTable
    If pSTCollection.StandaloneTableCount > 0 Then
      Dim lngIndex As Long
      For lngIndex = 0 To pSTCollection.StandaloneTableCount - 1
        Set pStTble = pSTCollection.StandaloneTable(lngIndex)
        If booStandaloneTables Then ReturnLayersByType2.Add pStTble
      Next lngIndex
    End If
  End If
  
'  Dim strReport As String
'  For lngIndex = 0 To ReturnLayersByType2.Count - 1
'    Set pAsLayer = ReturnLayersByType2.Element(lngIndex)
'    strReport = strReport & CStr(lngIndex) & "] " & pAsLayer.Name & vbCrLf
'  Next lngIndex
'  MsgBox strReport

ClearMemory:
  Set pEnumLayer = Nothing
  Set pFeatureLayer = Nothing
  Set pLayer = Nothing
  Set pFeatureClass = Nothing
  Set pFeatureLayerForValid = Nothing
  Set pRasterLayer = Nothing
  Set pTinLayer = Nothing
  Set pSTCollection = Nothing
  Set pStTble = Nothing

End Function
Public Function BasicStatsFromVAT(anArray() As Double, dblSizeArray() As Double, _
      theFieldName As String, theTableName As String, _
      m_pApp As Application, Optional lngNumberHistBins As Long = -9999) As esriSystem.IVariantArray
  
  
  ' ASSUMES ARRAYS ARE SORTED!!!! --------------------------------------
  
  Dim pResponse As esriSystem.IVariantArray
  Set pResponse = New esriSystem.varArray
  
    ' PROGRESS BAR STUFF
  Dim pSBar As IStatusBar
  Set pSBar = m_pApp.StatusBar
  Dim pPro As IStepProgressor
  Set pPro = pSBar.ProgressBar
  
  ' IF MAKING A HISTOGRAM
  Dim booMakeHistogram As Boolean
  booMakeHistogram = (lngNumberHistBins > 0)
  Dim lngHistCountArray() As Long            ' WILL CONTAIN COUNTS OF VALUES LYING IN EACH HISTOGRAM BIN
  If booMakeHistogram Then
    Dim dblHistLow As Double
    Dim dblHistHigh As Double
    dblHistLow = anArray(0)
    dblHistHigh = anArray(UBound(anArray))
    Dim dblInterval As Double
    dblInterval = (dblHistHigh - dblHistLow) / lngNumberHistBins
    
    ReDim lngHistCountArray(lngNumberHistBins)
    Dim lngHistBinIndex As Long
    lngHistBinIndex = 0
    Dim dblCurrentBinThreshold As Double
    dblCurrentBinThreshold = dblHistLow + dblInterval
  End If
  
  Screen.MousePointer = vbHourglass
  
  If booMakeHistogram Then
    pSBar.ShowProgressBar "Calculating Statistics on field [" & theFieldName & "] in " & theTableName, 1, _
          4 * (UBound(anArray) - 1), 1, True
  Else
    pSBar.ShowProgressBar "Calculating Statistics on field [" & theFieldName & "] in " & theTableName, 1, _
          3 * (UBound(anArray) - 1), 1, True
  End If
  
  Dim FoundMode As Boolean
  FoundMode = False
  
  Dim theModeValList() As Double
  ReDim theModeValList(1, UBound(anArray))
  
  Dim theHighModeCount As Long
  theHighModeCount = 0
  
  Dim anIndex As Long
  Dim theVal As Double
  Dim theSize As Double
  Dim theValTimesSize As Double
  Dim theModeCounter As Long
  Dim theModeIndex As Long
  theModeIndex = -1
  
  anIndex = 0
  pSBar.StepProgressBar
  
  ' IF MAKING HISTOGRAM
  If booMakeHistogram Then
    Do While anIndex <= UBound(anArray)
      theVal = anArray(anIndex)
      theSize = dblSizeArray(anIndex)
      theValTimesSize = theVal * theSize
      If theVal <= dblCurrentBinThreshold Then
        lngHistCountArray(lngHistBinIndex) = lngHistCountArray(lngHistBinIndex) + theSize
      Else
        Do While dblCurrentBinThreshold < theVal
          dblCurrentBinThreshold = dblCurrentBinThreshold + dblInterval
          lngHistBinIndex = lngHistBinIndex + 1
        Loop
        lngHistCountArray(lngHistBinIndex) = theSize
      End If
      anIndex = anIndex + 1
      pSBar.StepProgressBar
    Loop
  End If
  
  '  PASS 1:  MODE --------------------------------------------------
  anIndex = 0
  
  Do While anIndex < UBound(anArray)
    
    theModeCounter = 0
    theVal = anArray(anIndex)
    theSize = dblSizeArray(anIndex)
    
    Do While (anIndex < UBound(anArray))
      ' IF THE NEXT VALUE UP IS DIFFERENT, THEN START NEW COUNT
      If Not (anArray(anIndex + 1) = theVal) Then
        theModeCounter = theSize
        Exit Do
      End If
            
      ' IF NEXT VALUE UP IS THE SAME, THEN ADD NEW SIZE TO CURRENT TALLY.  THEN CONTINUE LOOKING FOR A NEW VALUE.
      theModeCounter = theModeCounter + theSize
      
      anIndex = anIndex + 1
      theSize = dblSizeArray(anIndex)
      
      pSBar.StepProgressBar
    Loop
    
    If theModeCounter > 1 Then
      FoundMode = True
      theModeIndex = theModeIndex + 1
      theModeValList(0, theModeIndex) = theModeCounter
      theModeValList(1, theModeIndex) = theVal
      
'      Debug.Print "Value = " & theModeValList(1, theModeIndex) & "[" & theModeValList(0, theModeIndex) & " cases]"
      
    End If
    
    anIndex = anIndex + 1
    pSBar.StepProgressBar
  Loop
  
  Dim theModeString As String
  
  ' IF ANY VALUE OCCURED > 1 TIME
  Dim theFinalModes() As Double
  If FoundMode Then
    ReDim Preserve theModeValList(1, theModeIndex)
    
    Dim theFinalModeCount As Long
    theFinalModeCount = 0
    Dim theTempCount As Long
    theTempCount = 0
        
    Dim theFinalModeIndex As Long
        
    For anIndex = 0 To theModeIndex
      theTempCount = theModeValList(0, anIndex)
      If theTempCount > theFinalModeCount Then
        theFinalModeCount = theTempCount
        ReDim theFinalModes(0)
        theFinalModes(0) = theModeValList(1, anIndex)
      ElseIf theTempCount = theFinalModeCount Then
        theFinalModeIndex = UBound(theFinalModes) + 1
        ReDim Preserve theFinalModes(theFinalModeIndex)
        theFinalModes(theFinalModeIndex) = theModeValList(1, anIndex)
      End If
    Next anIndex
    
    If UBound(theFinalModes) > 0 Then
      theModeString = UBound(theFinalModes) + 1 & " modes found [" & aml_func_mod.InsertCommas(theFinalModeCount) & " cases each]: Values = "
      For anIndex = 0 To UBound(theFinalModes)
        theModeString = theModeString & theFinalModes(anIndex) & ", "
      Next anIndex
      
      theModeString = aml_func_mod.BasicTrimAvenue(theModeString, "", ", ")
      
    Else
      theModeString = "1 mode found [" & aml_func_mod.InsertCommas(theFinalModeCount) & " cases]: Value = " & theFinalModes(0)
    End If
    
  Else
    theModeString = " < No Mode Found >"
  End If
  
  Dim theSum As Double
  Dim theCount As Double
  Dim theMinimum As Double
  Dim theMaximum As Double
  
  theSum = 0
  theCount = 0
  theMinimum = anArray(0)
  theMaximum = anArray(0)
  
  '  PASS 2:  MINIMUM, MAXIMUM AND SUM --------------------------------------------------
  For anIndex = LBound(anArray) To UBound(anArray)
        
    pSBar.StepProgressBar
    
    theVal = anArray(anIndex)
    theSize = dblSizeArray(anIndex)
    theCount = theCount + theSize
    
    theValTimesSize = theVal * theSize
    
    If theVal < theMinimum Then
      theMinimum = theVal
    End If
    If theVal > theMaximum Then
      theMaximum = theVal
    End If
    theSum = theSum + theValTimesSize
    
  Next anIndex
  
  Dim theMean As Double
  theMean = theSum / theCount
  
  Dim theSumSqDev As Double
  Dim theSqDev As Double
  Dim theMedian As Double
  Dim lngMiddleIndex As Long     ' DON'T HAVE TO WORRY ABOUT DECIMAL COUNTS BECAUSE WORKING WITH VAT.  ALL COUNTS ARE INTEGER.
  Dim theRunningCount As Double
  theRunningCount = 0
  Dim booFoundMedian As Boolean
  booFoundMedian = False
  
  If theCount = 0 Then
    lngMiddleIndex = 0
  ElseIf theCount Mod 2 = 0 Then      ' EVEN NUMBER
'      theMedian = (anArray((theCount / 2) - 1) + anArray(theCount / 2)) / 2
    lngMiddleIndex = (((theCount / 2) - 1) + (theCount / 2)) / 2
  Else
'      theMedian = anArray((theCount - 1) / 2)
    lngMiddleIndex = (theCount - 1) / 2
  End If
  
  '  PASS 2: MEDIAN, STANDARD DEVIATION AND VARIANCE --------------------------------------------------
  For anIndex = LBound(anArray) To UBound(anArray)
        
    pSBar.StepProgressBar
    
    theVal = anArray(anIndex)
    theSize = dblSizeArray(anIndex)
    theRunningCount = theRunningCount + theSize
    If Not booFoundMedian Then
      If theRunningCount >= lngMiddleIndex Then
        theMedian = theVal
        booFoundMedian = True
      End If
    End If
    theSqDev = theSize * ((theVal - theMean) * (theVal - theMean))
    theSumSqDev = theSqDev + theSumSqDev
    
  Next anIndex
  
  Dim theVariance As Double
  Dim theStDev As Double
  Dim theStErrMean As Double
  
  If theCount > 0 Then
    
    theVariance = theSumSqDev / (theCount - 1)
    theStDev = Sqr(theVariance)
    theStErrMean = theStDev / (Sqr(theCount))
    
  Else
    theMedian = -9999
    theVariance = 0
    theStDev = 0
    theStErrMean = 0
  End If

  Dim theRange As Double
  theRange = theMaximum - theMinimum

  ' OUTPUT ARRAY; VARIANT BECAUSE OF MODE STRING
  '(0) = SUM
  '(1) = MEAN
  '(2) = MINIMUM
  '(3) = MAXIMUM
  '(4) = RANGE
  '(5) = COUNT
  '(6) = STANDARD DEVIATION
  '(7) = VARIANCE
  '(8) = MEDIAN
  '(9) = STANDARD ERROR OF MEAN
  '(10) = MODE STRING
  '(11) = DOUBLE ARRAY OF MODE VALUES
  '(12) = BOOLEAN INDICATING WHETHER MODE WAS FOUND
  '(13) = ARRAY OF HISTOGRAM BIN COUNTS
  
  pResponse.Add theSum
  pResponse.Add theMean
  pResponse.Add theMinimum
  pResponse.Add theMaximum
  pResponse.Add theRange
  pResponse.Add theCount
  pResponse.Add theStDev
  pResponse.Add theVariance
  pResponse.Add theMedian
  pResponse.Add theStErrMean
  pResponse.Add theModeString
  pResponse.Add theFinalModes
  pResponse.Add FoundMode
  pResponse.Add lngHistCountArray  ' WILL BE EMPTY ARRAY IF NO HISTOGRAM WAS REQUESTED
  
  pSBar.HideProgressBar
  Screen.MousePointer = vbDefault
  
  Set BasicStatsFromVAT = pResponse


  GoTo ClearMemory
ClearMemory:
  Set pResponse = Nothing
  Set pSBar = Nothing
  Set pPro = Nothing
  Erase lngHistCountArray
  Erase theModeValList
  Erase theFinalModes

End Function

Public Function BasicStatsFromArray_Weighted(anArray() As Double, theFieldName As String, theTableName As String, _
      m_pApp As Application) As esriSystem.IDoubleArray
  
  
  ' REQUIRES 2-DIMENSIONAL INPUT ARRAY, WHERE 1ST VALUE = VALUE AND SECOND VALUE = SIZE
  
  Dim pResponse As esriSystem.IDoubleArray
  Set pResponse = New esriSystem.DoubleArray
  
    ' PROGRESS BAR STUFF
  Dim pSBar As IStatusBar
  Set pSBar = m_pApp.StatusBar
  Dim pPro As IStepProgressor
  Set pPro = pSBar.ProgressBar
    
  Screen.MousePointer = vbHourglass
  
  pSBar.ShowProgressBar "Calculating Statistics on field [" & theFieldName & "] in " & theTableName, 1, _
        2 * (UBound(anArray) - 1), 1, True
  
  Dim anIndex As Long
  Dim theVal As Double
  Dim theWeight As Double
  Dim theWeightedVal As Double
  
  anIndex = 0
  pSBar.StepProgressBar
    
  Dim theSum As Double
  Dim theCount As Double
  
  theCount = 0
  
  '  MINIMUM, MAXIMUM AND SUM --------------------------------------------------
  For anIndex = LBound(anArray, 2) To UBound(anArray, 2)
        
    pSBar.StepProgressBar
    
    theVal = anArray(0, anIndex)
    theWeight = anArray(1, anIndex)
    
    theCount = theCount + theWeight
    theWeightedVal = theVal * theWeight
    theSum = theSum + theWeightedVal
    
  Next anIndex
  
  Dim theMean As Double
  theMean = theSum / theCount
  
  Dim theSumSqDev As Double
  theSumSqDev = 0
  Dim theSqDev As Double
  
  '  PASS 2:  STANDARD DEVIATION AND VARIANCE --------------------------------------------------
  For anIndex = LBound(anArray, 2) To UBound(anArray, 2)
        
    pSBar.StepProgressBar
    
    theVal = anArray(0, anIndex)
    theWeight = anArray(1, anIndex)
    theSqDev = theWeight * ((theVal - theMean) * (theVal - theMean))
    theSumSqDev = theSqDev + theSumSqDev
    
  Next anIndex
  
  Dim theVariance As Double
  Dim theStDev As Double
  
  theVariance = theSumSqDev / theCount
  theStDev = Sqr(theVariance)
  
  pResponse.Add theMean
  pResponse.Add theStDev
  pResponse.Add theVariance
  
  pSBar.HideProgressBar
    
  Screen.MousePointer = vbDefault
  
  Set BasicStatsFromArray_Weighted = pResponse


  GoTo ClearMemory
ClearMemory:
  Set pResponse = Nothing
  Set pSBar = Nothing
  Set pPro = Nothing

End Function
Public Sub BasicStatsFromArray_WeightedFast(dblVals() As Double, dblMean As Double, dblStDev As Double, _
    dblVar As Double)
   
  ' REQUIRES 2-DIMENSIONAL INPUT ARRAY, WHERE 1ST VALUE = VALUE AND SECOND VALUE = SIZE
  Screen.MousePointer = vbHourglass
  
  Dim lngIndex As Long
  Dim dblVal As Double
  Dim dblWeight As Double
  Dim dblWeightedVal As Double
  
  lngIndex = 0
    
  Dim dblSum As Double
  Dim dblCount As Double
  
  dblCount = 0
'  Dim strReport As String
'  strReport = "Value" & Chr(9) & "Weight" & Chr(9) & "CalcMultiply" & vbCrLf
  
  '  MINIMUM, MAXIMUM AND SUM --------------------------------------------------
  For lngIndex = LBound(dblVals, 2) To UBound(dblVals, 2)
        
    dblVal = dblVals(0, lngIndex)
    dblWeight = dblVals(1, lngIndex)
    
    dblCount = dblCount + dblWeight
    dblWeightedVal = dblVal * dblWeight
    dblSum = dblSum + dblWeightedVal
    
'    strReport = strReport & CStr(dblVal) & Chr(9) & CStr(dblWeight) & Chr(9) & CStr(dblWeightedVal) & vbCrLf
    
  Next lngIndex
  
'  Clipboard.Clear
'  Clipboard.SetText strReport
  
  dblMean = dblSum / dblCount
  
  
'  MsgBox "Sum(Slope * Weight) = " & CStr(dblSum) & vbCrLf & _
'         "Sum(Weight) = " & CStr(dblCount) & vbCrLf & _
'         "[Sum(Slope * Weight)]/[Sum(Weight)] = " & CStr(dblMean)
  
  Dim dblSumSqDev As Double
  dblSumSqDev = 0
  Dim theSqDev As Double
  
  '  PASS 2:  STANDARD DEVIATION AND VARIANCE --------------------------------------------------
  For lngIndex = LBound(dblVals, 2) To UBound(dblVals, 2)
    
    dblVal = dblVals(0, lngIndex)
    dblWeight = dblVals(1, lngIndex)
    theSqDev = dblWeight * ((dblVal - dblMean) * (dblVal - dblMean))
    dblSumSqDev = theSqDev + dblSumSqDev
    
  Next lngIndex
    
  dblVar = dblSumSqDev / dblCount
  dblStDev = Sqr(dblVar)
      
  GoTo ClearMemory
ClearMemory:
  Screen.MousePointer = vbDefault

End Sub
Public Function BasicStatsFromArraySimple(anArray() As Double, Optional booVariance As Boolean = False) As esriSystem.IDoubleArray
  
  ' ASSUMES ARRAY IS SORTED!!!! --------------------------------------
  
  Dim pResponse As esriSystem.IDoubleArray
  Set pResponse = New esriSystem.DoubleArray
  
  Dim anIndex As Long
  Dim theVal As Double

  Dim theSum As Double
  Dim theCount As Double
  Dim theMinimum As Double
  Dim theMaximum As Double
  
  theSum = 0
  theCount = UBound(anArray) + 1         ' ARRAY INDEX STARTS AT 0
  theMinimum = anArray(0)
  theMaximum = anArray(0)
  
  '  PASS 1:  MINIMUM, MAXIMUM AND SUM --------------------------------------------------
  For anIndex = LBound(anArray) To UBound(anArray)
    
    theVal = anArray(anIndex)
    
    If theVal < theMinimum Then
      theMinimum = theVal
    End If
    If theVal > theMaximum Then
      theMaximum = theVal
    End If
    theSum = theSum + theVal
    
  Next anIndex
  
  Dim theMean As Double
  theMean = theSum / theCount
  
  Dim theSumSqDev As Double
  Dim theSqDev As Double
  
  '  PASS 2:  STANDARD DEVIATION AND VARIANCE --------------------------------------------------
  If booVariance Then
    For anIndex = LBound(anArray) To UBound(anArray)
                
      theVal = anArray(anIndex)
      theSqDev = (theVal - theMean) * (theVal - theMean)
      theSumSqDev = theSqDev + theSumSqDev
      
    Next anIndex
  Else
    theSqDev = 0
    theSumSqDev = 0
  End If
  
  Dim theMedian As Double
  Dim theVariance As Double
  Dim theStDev As Double
  Dim theStErrMean As Double
  
  If theCount > 1 Then
    If theCount Mod 2 = 0 Then      ' EVEN NUMBER
      theMedian = (anArray((theCount / 2) - 1) + anArray(theCount / 2)) / 2
    Else
      theMedian = anArray((theCount - 1) / 2)
    End If
    
    theVariance = theSumSqDev / (theCount - 1)
    theStDev = Sqr(theVariance)
    theStErrMean = theStDev / (Sqr(theCount))
    
  Else
    If theCount = 1 Then
      theMedian = anArray(0)
    Else
      theMedian = -9999
    End If
    theVariance = 0
    theStDev = 0
    theStErrMean = 0
  End If

  Dim theRange As Double
  theRange = theMaximum - theMinimum

  ' OUTPUT ARRAY; VARIANT BECAUSE OF MODE STRING
  '(0) = SUM
  '(1) = MEAN
  '(2) = MINIMUM
  '(3) = MAXIMUM
  '(4) = RANGE
  '(5) = COUNT
  '(6) = STANDARD DEVIATION
  '(7) = VARIANCE
  '(8) = MEDIAN
  '(9) = STANDARD ERROR OF MEAN
  
  pResponse.Add theSum
  pResponse.Add theMean
  pResponse.Add theMinimum
  pResponse.Add theMaximum
  pResponse.Add theRange
  pResponse.Add theCount
  pResponse.Add theStDev
  pResponse.Add theVariance
  pResponse.Add theMedian
  pResponse.Add theStErrMean
  
  Set BasicStatsFromArraySimple = pResponse

  GoTo ClearMemory

ClearMemory:
  Set pResponse = Nothing
  
End Function

Public Function BasicStatsFromArray(anArray() As Double, theFieldName As String, theTableName As String, _
      m_pApp As Application, Optional lngNumberHistBins As Long = -9999) As esriSystem.IVariantArray
  
  
  ' ASSUMES ARRAY IS SORTED!!!! --------------------------------------
  
  Dim pResponse As esriSystem.IVariantArray
  Set pResponse = New esriSystem.varArray
  
    ' PROGRESS BAR STUFF
  Dim pSBar As IStatusBar
  Set pSBar = m_pApp.StatusBar
  Dim pPro As IStepProgressor
  Set pPro = pSBar.ProgressBar
  
  ' IF MAKING A HISTOGRAM
  Dim booMakeHistogram As Boolean
  booMakeHistogram = (lngNumberHistBins > 0)
  Dim lngHistCountArray() As Long            ' WILL CONTAIN COUNTS OF VALUES LYING IN EACH HISTOGRAM BIN
  If booMakeHistogram Then
    Dim dblHistLow As Double
    Dim dblHistHigh As Double
    dblHistLow = anArray(0)
    dblHistHigh = anArray(UBound(anArray))
    Dim dblInterval As Double
    dblInterval = (dblHistHigh - dblHistLow) / lngNumberHistBins
    
    ReDim lngHistCountArray(lngNumberHistBins)
    Dim lngHistBinIndex As Long
    lngHistBinIndex = 0
    Dim dblCurrentBinThreshold As Double
    dblCurrentBinThreshold = dblHistLow + dblInterval
  End If
  
  Screen.MousePointer = vbHourglass
  
  If booMakeHistogram Then
    pSBar.ShowProgressBar "Calculating Statistics on field [" & theFieldName & "] in " & theTableName, 1, _
          4 * (UBound(anArray) - 1), 1, True
  Else
    pSBar.ShowProgressBar "Calculating Statistics on field [" & theFieldName & "] in " & theTableName, 1, _
          3 * (UBound(anArray) - 1), 1, True
  End If
  
  Dim FoundMode As Boolean
  FoundMode = False
  
  Dim theModeValList() As Double
  ReDim theModeValList(1, UBound(anArray))
  
  Dim theHighModeCount As Long
  theHighModeCount = 0
  
  Dim anIndex As Long
  Dim theVal As Double
  Dim theModeCounter As Long
  Dim theModeIndex As Long
  theModeIndex = -1
  
  anIndex = 0
  pSBar.StepProgressBar
  
  ' IF MAKING HISTOGRAM
  If booMakeHistogram Then
    Do While anIndex <= UBound(anArray)
      theVal = anArray(anIndex)
      If theVal <= dblCurrentBinThreshold Then
        lngHistCountArray(lngHistBinIndex) = lngHistCountArray(lngHistBinIndex) + 1
      Else
        Do While dblCurrentBinThreshold < theVal
          dblCurrentBinThreshold = dblCurrentBinThreshold + dblInterval
          lngHistBinIndex = lngHistBinIndex + 1
        Loop
        lngHistCountArray(lngHistBinIndex) = 1
      End If
      anIndex = anIndex + 1
      pSBar.StepProgressBar
    Loop
  End If
  
  '  PASS 1:  MODE --------------------------------------------------
  anIndex = 0
  
  Do While anIndex < UBound(anArray)
    
    theModeCounter = 1
    theVal = anArray(anIndex)
    
    Do While (anIndex < UBound(anArray))
      If Not (anArray(anIndex + 1) = theVal) Then
        Exit Do
      End If
      anIndex = anIndex + 1
      pSBar.StepProgressBar
      theModeCounter = theModeCounter + 1
    Loop
    
    If theModeCounter > 1 Then
      FoundMode = True
      theModeIndex = theModeIndex + 1
      theModeValList(0, theModeIndex) = theModeCounter
      theModeValList(1, theModeIndex) = theVal
      
'      Debug.Print "Value = " & theModeValList(1, theModeIndex) & "[" & theModeValList(0, theModeIndex) & " cases]"
      
    End If
    
    anIndex = anIndex + 1
    pSBar.StepProgressBar
  Loop
  
  Dim theModeString As String
  
  ' IF ANY VALUE OCCURED > 1 TIME
  Dim theFinalModes() As Double
  If FoundMode Then
    ReDim Preserve theModeValList(1, theModeIndex)
    
    Dim theFinalModeCount As Long
    theFinalModeCount = 0
    Dim theTempCount As Long
    theTempCount = 0
        
    Dim theFinalModeIndex As Long
        
    For anIndex = 0 To theModeIndex
      theTempCount = theModeValList(0, anIndex)
      If theTempCount > theFinalModeCount Then
        theFinalModeCount = theTempCount
        ReDim theFinalModes(0)
        theFinalModes(0) = theModeValList(1, anIndex)
      ElseIf theTempCount = theFinalModeCount Then
        theFinalModeIndex = UBound(theFinalModes) + 1
        ReDim Preserve theFinalModes(theFinalModeIndex)
        theFinalModes(theFinalModeIndex) = theModeValList(1, anIndex)
      End If
    Next anIndex
    
    If UBound(theFinalModes) > 0 Then
      theModeString = UBound(theFinalModes) + 1 & " modes found [" & aml_func_mod.InsertCommas(theFinalModeCount) & " cases each]: Values = "
      For anIndex = 0 To UBound(theFinalModes)
        theModeString = theModeString & theFinalModes(anIndex) & ", "
      Next anIndex
      
      theModeString = aml_func_mod.BasicTrimAvenue(theModeString, "", ", ")
      
    Else
      theModeString = "1 mode found [" & aml_func_mod.InsertCommas(theFinalModeCount) & " cases]: Value = " & theFinalModes(0)
    End If
    
  Else
    theModeString = " < No Mode Found >"
  End If
  
  Dim theSum As Double
  Dim theCount As Double
  Dim theMinimum As Double
  Dim theMaximum As Double
  
  theSum = 0
  theCount = UBound(anArray) + 1         ' ARRAY INDEX STARTS AT 0
  theMinimum = anArray(0)
  theMaximum = anArray(0)
  
  '  PASS 2:  MINIMUM, MAXIMUM AND SUM --------------------------------------------------
  For anIndex = LBound(anArray) To UBound(anArray)
        
    pSBar.StepProgressBar
    
    theVal = anArray(anIndex)
    
    If theVal < theMinimum Then
      theMinimum = theVal
    End If
    If theVal > theMaximum Then
      theMaximum = theVal
    End If
    theSum = theSum + theVal
    
  Next anIndex
  
  Dim theMean As Double
  theMean = theSum / theCount
  
  Dim theSumSqDev As Double
  Dim theSqDev As Double
  
  '  PASS 2:  STANDARD DEVIATION AND VARIANCE --------------------------------------------------
  For anIndex = LBound(anArray) To UBound(anArray)
        
    pSBar.StepProgressBar
    
    theVal = anArray(anIndex)
    theSqDev = (theVal - theMean) * (theVal - theMean)
    theSumSqDev = theSqDev + theSumSqDev
    
  Next anIndex
  
  Dim theMedian As Double
  Dim theVariance As Double
  Dim theStDev As Double
  Dim theStErrMean As Double
  
  If theCount > 1 Then
    If theCount Mod 2 = 0 Then      ' EVEN NUMBER
      theMedian = (anArray((theCount / 2) - 1) + anArray(theCount / 2)) / 2
    Else
      theMedian = anArray((theCount - 1) / 2)
    End If
    
    theVariance = theSumSqDev / (theCount - 1)
    theStDev = Sqr(theVariance)
    theStErrMean = theStDev / (Sqr(theCount))
    
  Else
    If theCount = 1 Then
      theMedian = anArray(0)
    Else
      theMedian = -9999
    End If
    theVariance = 0
    theStDev = 0
    theStErrMean = 0
  End If

  Dim theRange As Double
  theRange = theMaximum - theMinimum

  ' OUTPUT ARRAY; VARIANT BECAUSE OF MODE STRING
  '(0) = SUM
  '(1) = MEAN
  '(2) = MINIMUM
  '(3) = MAXIMUM
  '(4) = RANGE
  '(5) = COUNT
  '(6) = STANDARD DEVIATION
  '(7) = VARIANCE
  '(8) = MEDIAN
  '(9) = STANDARD ERROR OF MEAN
  '(10) = MODE STRING
  '(11) = DOUBLE ARRAY OF MODE VALUES
  '(12) = BOOLEAN INDICATING WHETHER MODE WAS FOUND
  '(13) = ARRAY OF HISTOGRAM BIN COUNTS
  
  pResponse.Add theSum
  pResponse.Add theMean
  pResponse.Add theMinimum
  pResponse.Add theMaximum
  pResponse.Add theRange
  pResponse.Add theCount
  pResponse.Add theStDev
  pResponse.Add theVariance
  pResponse.Add theMedian
  pResponse.Add theStErrMean
  pResponse.Add theModeString
  pResponse.Add theFinalModes
  pResponse.Add FoundMode
  pResponse.Add lngHistCountArray  ' WILL BE EMPTY ARRAY IF NO HISTOGRAM WAS REQUESTED
  
  pSBar.HideProgressBar
  Screen.MousePointer = vbDefault
  
  Set BasicStatsFromArray = pResponse


  GoTo ClearMemory
ClearMemory:
  Set pResponse = Nothing
  Set pSBar = Nothing
  Set pPro = Nothing
  Erase lngHistCountArray
  Erase theModeValList
  Erase theFinalModes

End Function


Public Function CalcStatistics(dblSortedNumbers() As Double, pStatOptions As esriSystem.IVariantArray, _
      Optional booReportProgress As Boolean, Optional pApp As IApplication) As esriSystem.IVariantArray

  ' FAO_WRD.Stat_CalcFieldStats
'
'Dim chkMean As Boolean
'Set chkMean = pStatOptions.Element(0)
'Dim chkSEMean As Boolean
'Set chkSEMean = pStatOptions.Element(1)
'Dim chkCIMean As Boolean
'Set chkCIMean = pStatOptions.Element(2)
'Dim ConLevel As Double
'Set ConLevel = pStatOptions.Element(3)
'Dim chkMinimum As Boolean
'Set chkMinimum = pStatOptions.Element(4)
'Dim chkQ1 As Boolean
'Set chkQ1 = pStatOptions.Element(5)
'Dim chkMedian As Boolean
'Set chkMedian = pStatOptions.Element(6)
'Dim chkQ3 As Boolean
'Set chkQ3 = pStatOptions.Element.Item(7)
'Dim chkMaximum As Boolean
'Set chkMaximum = pStatOptions.Element(8)
'Dim chkVariance As Boolean
'Set chkVariance = pStatOptions.Element(9)
'Dim chkStDev As Boolean
'Set chkStDev = pStatOptions.Element(10)
'Dim chkAvgDev As Boolean
'Set chkAvgDev = pStatOptions.Element(11)
'Dim chkSkewness As Boolean
'Set chkSkewness = pStatOptions.Element(12)
'Dim chkSkewnessFish As Boolean
'Set chkSkewnessFish = pStatOptions.Element(13)
'Dim chkKurtosis As Boolean
'Set chkKurtosis = pStatOptions.Element(14)
'Dim chkKurtosisFish As Boolean
'Set chkKurtosisFish = pStatOptions.Element(15)
'Dim chkCount As Boolean
'Set chkCount = pStatOptions.Element(16)
'Dim chkNumberNull As Boolean
'Set chkNumberNull = pStatOptions.Element(17)
'Dim chkSum As Boolean
'Set chkSum = pStatOptions.Element(18)
'Dim chkRange As Boolean
'Set chkRange = pStatOptions.Element(19)
'Dim chkMode As Boolean
'Set chkMode = pStatOptions.Element(20)
'
'' MakeHistogram = theResults.Get(21)
'
'Dim pResponse As esriSystem.IVariantArray
'Dim anIndex As Long
'For anIndex = 0 To 20
'  pResponse.Add Nothing
'Next anIndex
'
''' UNFORTUNATELY, THE HISTOGRAM FUNCTION WAS WRITTEN SEPARATELY AND THE STATS ARE IN A DIFFERENT ORDER
''theResponse = {nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil}
''theStatsForHistogramScript = {nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil}
''
''If (MakeHistogram) Then
''  chkMean = True
''  chkStDev = True
''  chkCount = True
''  chkMinimum = True
''  chkMaximum = True
''  theResults.Set(0,True)
''  theResults.Set(4,True)
''  theResults.Set(8,True)
''  theResults.Set(10,True)
''  theResults.Set(16,True)
''End
'
''' HELP IT RUN FASTER IF ONLY A FEW OPTIONS ARE CHECKED
'Dim DoQuantiles As Boolean
'Dim DoSkewKurt As Boolean
'Dim MoreThanBasic As Boolean
'Dim lngMultiplier As Long
'lngMultiplier = 1   ' FOR BASIC STATS
'
'' DIMENSION STATISTIC VARIABLES
'Dim theMean As Double
'Dim theSEMean As Double
'Dim LowerCI As Double
'Dim UpperCI As Double
'Dim theMinimum As Double
'Dim theQ1 As Double
'Dim theMedian As Double
'Dim theQ3 As Double
'Dim theMaximum As Double
'Dim theVar As Double
'Dim theStdDev As Double
'Dim theAvgDev As Double
'Dim theSkew As Double
'Dim theFisherSkew As Double
'Dim theKurt As Double
'Dim theFisherKurt As Double
'Dim theCount As Double
'Dim theNumberNull As Double
'Dim theSum As Double
'Dim theRange As Double
'Dim theModeString As String
'
'' BASIC STATS ARE Count, Minimum, Maximum, Mean, Sum
'DoQuantiles = (chkQ1 Or chkMedian Or chkQ3)
'DoSkewKurt = (chkSkewness Or chkSkewnessFish Or chkKurtosis Or chkKurtosisFish)
''MoreThanBasic = (DoQuantiles Or DoSkewKurt Or chkSEMean Or chkCIMean Or chkVariance Or chkStDev Or chkAvgDev Or MakeHistogram)
'MoreThanBasic = (DoQuantiles Or DoSkewKurt Or chkSEMean Or chkCIMean Or chkVariance Or chkStDev Or chkAvgDev)
'
'If chkVariance Or chkStDev Or chkAvgDev Then lngMultiplier = lngMultiplier + 1
'
'If DoQuantiles Then lngMultiplier = lngMultiplier + 1
'If DoSkewKurt Then lngMultiplier = lngMultiplier + 1
'If MoreThanBasic Then lngMultiplier = lngMultiplier + 1
'If chkMode Then lngMultiplier = lngMultiplier + 1
'
'If booReportProgress Then
'    ' PROGRESS BAR STUFF
'  Dim psbar As IStatusBar
'  Set psbar = pApp.StatusBar
'  Dim pPro As IStepProgressor
'  Set pPro = psbar.ProgressBar
'
'  Screen.MousePointer = vbHourglass
'
'  psbar.ShowProgressBar "Calculating Statistics on field [" & theFieldName & "] in " & theTableName, 1, _
'        lngMultiplier * (UBound(dblSortedNumbers) - 1), 1, True
'End If
'
'Dim theCount As Long
'theCount = UBound(dblSortedNumbers) - 1
'
'If chkMode Then  ' ASSUMES NUMBER ARRAY IS SORTED!
'
'  Dim pFoundMode As Boolean
'  Dim theModeValList() As Double
'  ReDim theModeValList(1, UBound(dblSortedNumbers))
'
'  Dim theHighModeCount As Long
'  theHighModeCount = 0
'
'  Dim theVal As Double
'  Dim theModeCounter As Long
'  Dim theModeIndex As Long
'  theModeIndex = -1
'
'  anIndex = 0
'
'  Do While anIndex < UBound(dblSortedNumbers)
'
'    theModeCounter = 1
'    theVal = dblSortedNumbers(anIndex)
'
'    Do While (anIndex < UBound(dblSortedNumbers))
'      If Not (dblSortedNumbers(anIndex + 1) = theVal) Then
'        Exit Do
'      End If
'      anIndex = anIndex + 1
'      If booReportProgress Then psbar.StepProgressBar
'      theModeCounter = theModeCounter + 1
'    Loop
'
'    If theModeCounter > 1 Then
'      FoundMode = True
'      theModeIndex = theModeIndex + 1
'      theModeValList(0, theModeIndex) = theModeCounter
'      theModeValList(1, theModeIndex) = theVal
'
''      Debug.Print "Value = " & theModeValList(1, theModeIndex) & "[" & theModeValList(0, theModeIndex) & " cases]"
'
'    End If
'
'    anIndex = anIndex + 1
'    If booReportProgress Then psbar.StepProgressBar
'  Loop
'
'  Dim theModeString As String
'
'  ' IF ANY VALUE OCCURED > 1 TIME
'  If FoundMode Then
'    ReDim Preserve theModeValList(1, theModeIndex)
'    Dim theFinalModes() As Double
'    Dim theFinalModeCount As Long
'    theFinalModeCount = 0
'    Dim theTempCount As Long
'    theTempCount = 0
'
'    Dim theFinalModeIndex As Long
'
'    For anIndex = 0 To theModeIndex
'      theTempCount = theModeValList(0, anIndex)
'      If theTempCount > theFinalModeCount Then
'        theFinalModeCount = theTempCount
'        ReDim theFinalModes(0)
'        theFinalModes(0) = theModeValList(1, anIndex)
'      ElseIf theTempCount = theFinalModeCount Then
'        theFinalModeIndex = UBound(theFinalModes) + 1
'        ReDim Preserve theFinalModes(theFinalModeIndex)
'        theFinalModes(theFinalModeIndex) = theModeValList(1, anIndex)
'      End If
'    Next anIndex
'
'    If UBound(theFinalModes) > 0 Then
'      theModeString = UBound(theFinalModes) + 1 & " modes found [" & aml_func_mod.InsertCommas(theFinalModeCount) & " cases each]: Values = "
'      For anIndex = 0 To UBound(theFinalModes)
'        theModeString = theModeString & theFinalModes(anIndex) & ", "
'      Next anIndex
'
'      theModeString = aml_func_mod.BasicTrimAvenue(theModeString, "", ", ")
'
'    Else
'      theModeString = "1 mode found [" & aml_func_mod.InsertCommas(theFinalModeCount) & " cases]: Value = " & theFinalModes(0)
'    End If
'
'  Else
'    theModeString = " < No Mode Found >"
'  End If
'End If
'
'theSum = 0
'theCount = UBound(dblSortedNumbers) + 1         ' ARRAY INDEX STARTS AT 0
'theMinimum = dblSortedNumbers(0)
'theMaximum = dblSortedNumbers(0)
'
''  PASS 2:  MINIMUM, MAXIMUM AND SUM --------------------------------------------------
'For anIndex = LBound(dblSortedNumbers) To UBound(dblSortedNumbers)
'
'  If booReportProgress Then psbar.StepProgressBar
'
'  theVal = dblSortedNumbers(anIndex)
'
'  If theVal < theMinimum Then
'    theMinimum = theVal
'  End If
'  If theVal > theMaximum Then
'    theMaximum = theVal
'  End If
'  theSum = theSum + theVal
'
'Next anIndex
'
'Dim theMean As Double
'theMean = theSum / theCount
'
'Dim theSumSqDev As Double
'Dim theSqDev As Double
'
''  PASS 2:  STANDARD DEVIATION AND VARIANCE --------------------------------------------------
'For anIndex = LBound(dblSortedNumbers) To UBound(dblSortedNumbers)
'
'  If booReportProgress Then psbar.StepProgressBar
'
'  theVal = dblSortedNumbers(anIndex)
'  theSqDev = (theVal - theMean) * (theVal - theMean)
'  theSumSqDev = theSqDev + theSumSqDev
'
'Next anIndex
'
'Dim theMedian As Double
'Dim theVariance As Double
'Dim theStDev As Double
'Dim theStErrMean As Double
'
'If theCount > 1 Then
'  If theCount Mod 2 = 0 Then      ' EVEN NUMBER
'    theMedian = (dblSortedNumbers((theCount / 2) - 1) + dblSortedNumbers(theCount / 2)) / 2
'  Else
'    theMedian = dblSortedNumbers((theCount - 1) / 2)
'  End If
'
'  theVariance = theSumSqDev / (theCount - 1)
'  theStDev = Sqr(theVariance)
'  theStErrMean = theStDev / (Sqr(theCount))
'
'Else
'  theMedian = -9999
'  theVariance = 0
'  theStDev = 0
'  theStErrMean = 0
'End If
'
'Dim theRange As Double
'theRange = theMaximum - theMinimum
'
'' OUTPUT ARRAY; VARIANT BECAUSE OF MODE STRING
''(0) = SUM
''(1) = MEAN
''(2) = MINIMUM
''(3) = MAXIMUM
''(4) = RANGE
''(5) = COUNT
''(6) = STANDARD DEVIATION
''(7) = VARIANCE
''(8) = MEDIAN
''(9) = STANDARD ERROR OF MEAN
''(10) = MODE STRING
'
'theStatsArray(0) = theSum
'theStatsArray(1) = theMean
'theStatsArray(2) = theMinimum
'theStatsArray(3) = theMaximum
'theStatsArray(4) = theRange
'theStatsArray(5) = theCount
'theStatsArray(6) = theStDev
'theStatsArray(7) = theVariance
'theStatsArray(8) = theMedian
'theStatsArray(9) = theStErrMean
'theStatsArray(10) = theModeString
'
''If (chkMode) Then
''  FoundMode = False
''  theModeValList = {}
''  theHighModeCount = 0
''  theDictionaryOfValues = Dictionary.Make(theCount)
''End
''
''theList = {}
''theNumberNull = 0
''theSum = 0
''theCounter = 0
''av.ClearStatus
''
''av.ShowMsg ("Calculating Pass 1")
''For Each aRecord In theSelection
''  theCounter = theCounter + 1
''  av.SetStatus ((theCounter / theCount) * 100)
''  theX = theVTab.ReturnValue(theField, aRecord)
''  If (theX.IsNull) Then
''    theNumberNull = theNumberNull + 1
''  Else
''    theList.Add (theX)
''    theSum = theSum + theX
''
''    ' CALCULATE MODE
''    If (chkMode) Then
''      theModeCount = theDictionaryOfValues.Get(theX)
''      If (theModeCount = nil) Then
''        theModeCount = 1
''      Else
''        theModeCount = theModeCount + 1
''        FoundMode = True
''      End
''      theDictionaryOfValues.Set(theX, theModeCount)
''
''      If (theModeCount > theHighModeCount) Then
''        theModeValList = {theX}
''        theHighModeCount = theModeCount
''      ElseIf (theModeCount = theHighModeCount) Then
''        theModeValList.Add (theX)
''      End
''    End
''  End
''End
''av.ClearStatus
''
''If (theList.Count = 0) Then
''  msgBox.Info("No data available to calculate!  Bailing out...", "")
''  return nil
''End
''
''If (chkMode) Then
''  If (FoundMode) Then
''    theModeValList.RemoveDuplicates
''    theModeValList.Sort (True)
''    theModeString = ""
''    For Each aModeVal In theModeValList
''      theModeString = theModeString + aModeVal.AsString + ", "
''    End
''    theModeString = theModeString.Left(theModeString.Count - 2)
''  Else
''    theModeString = " < No Mode Found >"
''  End
''End
''
''theTotalCount = theCount - theNumberNull
''
''theMean = theSum / theTotalCount
''
''If (MoreThanBasic) Then
''  av.ShowMsg ("Second Pass")
''  theCounter = 0
''
''  theAvgDev = 0
''  theMoment2 = 0
''  theMoment3 = 0
''  theMoment4 = 0
''  For Each anX In theList
''    theCounter = theCounter + 1
''    av.SetStatus ((theCounter / theTotalCount) * 100)
''    theDev = anX - theMean
''    theAvgDev = theAvgDev + (theDev.Abs)
''    theMoment2 = theMoment2 + (theDev ^ 2)
''    If (DoSkewKurt) Then
''      theMoment3 = theMoment3 + (theDev ^ 3)
''      theMoment4 = theMoment4 + (theDev ^ 4)
''    End
''  End
''  av.ClearStatus
''
''  ' MEAN, STANDARD DEVIATION, SKEWNESS, KURTOSIS --------------------------------
''  theVar = theMoment2 / (theTotalCount - 1)
''  theStdDev = theVar.Sqrt
''  theAvgDev = theAvgDev / theTotalCount
''
''  theMoment2 = theMoment2 / theTotalCount
''
''  If (DoSkewKurt) Then
''    theMoment3 = theMoment3 / theTotalCount
''    theMoment4 = theMoment4 / theTotalCount
''
''    theSkew = theMoment3 / (theMoment2 ^ (3 / 2))
''    theFisherSkew  = (theSkew*((theTotalCount*(theTotalCount-1)).sqrt))/(theTotalCount-2)
''
''    theKurt = theMoment4 / (theMoment2 ^ 2)
''    theFisherKurt = ((theTotalCount + 1) * (theTotalCount - 1)) / ((theTotalCount - 2) * (theTotalCount - 3))
''    theFisherKurt = theFisherKurt * (theKurt - (3 * (theTotalCount - 1) / (theTotalCount + 1)))
''
''    If (theVar <> 0) Then
''      theSkew = theMoment3 / (theMoment2 ^ (3 / 2))
''      theFisherSkew  = (theSkew*((theTotalCount*(theTotalCount-1)).sqrt))/(theTotalCount-2)
''
''      theKurt = theMoment4 / (theMoment2 ^ 2)
''      theFisherKurt = ((theTotalCount + 1) * (theTotalCount - 1)) / ((theTotalCount - 2) * (theTotalCount - 3))
''      theFisherKurt = theFisherKurt * (theKurt - (3 * (theTotalCount - 1) / (theTotalCount + 1)))
''    Else
''      theSkew = Number.MakeNull
''      theKurt = Number.MakeNull
''    End
''  End
''
''  ' STANDARD ERROR OF MEAN, CONFIDENCE INTERVALS
''  theSEMean = theStdDev / (theTotalCount.Sqrt)
''  If (chkCIMean) Then
''    theAlphaOver2 = 1 - ((1 - ConLevel) / 2)
''    theT = av.Run("FAO_WRD.Stat_IDF_StudentsT", {theAlphaOver2, (theTotalCount-1)})
''    theFactor = theSEMean * theT
''    LowerCI = theMean - theFactor
''    UpperCI = theMean + theFactor
''  End
''
''  ' QUANTILE DATA -----------------------------------------------------------------
''  If (DoQuantiles) Then
''
''    av.ShowMsg ("Calculating Quantile Data...")
''
''    theList.Sort (True)
''    theMinimum = theList.Get(0)
''    theMaximum = theList.Get(theList.Count - 1)
''    theRange = theMaximum - theMinimum
''
''    theListCount = theList.Count
''    theQ1Index = (theListCount + 1) * 0.25
''    theQ2Index = (theListCount + 1) * 0.5
''    theQ3Index = (theListCount + 1) * 0.75
''
''    If (theQ1Index.Round = theQ1Index) Then
''      theQ1 = theList.Get(theQ1Index - 1)
''    Else
''      theFloor = theList.Get((theQ1Index-1).Floor.Max(0))    ' POSSIBLE THAT IT COULD TRY TO GET INDEX -1, SO MAX IT WITH 0
''      theCeiling = theList.Get((theQ1Index-1).Ceiling)
''      theQ1 = (theFloor + theCeiling) / 2
''    End
''
''    If (theQ2Index.Round = theQ2Index) Then
''      theMedian = theList.Get(theQ2Index - 1)
''    Else
''      theFloor = theList.Get((theQ2Index-1).Floor)
''      theCeiling = theList.Get((theQ2Index-1).Ceiling)
''      theMedian = (theFloor + theCeiling) / 2
''    End
''
''    If (theQ3Index.Round = theQ3Index) Then
''      theQ3 = theList.Get(theQ3Index - 1)
''    Else
''      theFloor = theList.Get((theQ3Index-1).Floor)
''      theCeiling = theList.Get((theQ3Index-1).Ceiling.Min(theListCount-1))  ' POSSIBLE THAT IT COULD TRY TO GET INDEX (COUNT+1), SO MIN IT WITH COUNT
''
''      theQ3 = (theFloor + theCeiling) / 2
''    End
''  End
''End   ' END MORE THAN BASIC
''
'
'
'
'If booReportProgress Then
'  psbar.HideProgressBar
'  Screen.MousePointer = vbDefault
'End If
'
'If chkMean Then pResponse.Element(0) = theMean
'If chkSEMean Then pResponse.Element(1) = theSEMean
'If chkCIMean Then pResponse.Element(2) = LowerCI
'If chkCIMean Then pResponse.Element(3) = UpperCI
'If chkMinimum Then pResponse.Element(4) = theMinimum
'If chkQ1 Then pResponse.Element(5) = theQ1
'If chkMedian Then pResponse.Element(6) = theMedian
'If chkQ3 Then pResponse.Element(7) = theQ3
'If chkMaximum Then pResponse.Element(8) = theMaximum
'If chkVariance Then pResponse.Element(9) = theVar
'If chkStDev Then pResponse.Element(10) = theStdDev
'If chkAvgDev Then pResponse.Element(11) = theAvgDev
'If chkSkewness Then pResponse.Element(12) = theSkew
'If chkSkewnessFish Then pResponse.Element(13) = theFisherSkew
'If chkKurtosis Then pResponse.Element(14) = theKurt
'If chkKurtosisFish Then pResponse.Element(15) = theFisherKurt
'If chkCount Then pResponse.Element(16) = theCount
'If chkNumberNull Then pResponse.Element(17) = theNumberNull
'If chkSum Then pResponse.Element(18) = theSum
'If chkRange Then pResponse.Element(19) = theRange
'If chkMode Then pResponse.Element(20) = theModeString
'
'
'

End Function
Public Function Graphic_ReturnElementFromGeometry2(ByRef pMxDoc As IMxDocument, ByRef pGeometry As IGeometry, Optional strName As String, _
    Optional pSymbol As ISymbol, Optional booAddToLayout As Boolean = True) As IElement
  
  
  Dim pGContainer As IGraphicsContainer
  If booAddToLayout Then
    Set pGContainer = pMxDoc.PageLayout
  Else
    Set pGContainer = pMxDoc.FocusMap
  End If
  
  Dim pElement As IElement
  Dim pSpatialReference As ISpatialReference
  Dim pGraphicElement As IGraphicElement
  Dim pElementProperties As IElementProperties
  
  Dim pMarkerElement As IMarkerElement
  Dim pFillElement As IFillShapeElement
  Dim pLineElement As ILineElement
  
  Dim pClone As IClone
  Set pClone = pGeometry
  Dim pNewGeometry As IGeometry
  Set pNewGeometry = pClone.Clone
  
  Dim pGeometryType As esriGeometryType
  pGeometryType = pNewGeometry.GeometryType
  
  'ADD GEOMETRY, NAME AND SPATIAL REFERENCE TO GRAPHIC ELEMENT
  Select Case pGeometryType
    Case 0:
      MsgBox "Null Geometry!  No graphic added..."
    Case 1:
      Set pElement = New MarkerElement
      Set pMarkerElement = pElement
    Case 2:
      Set pElement = New GroupElement
    Case 3, 6, 13, 14, 15, 16:
      Set pElement = New LineElement
      Set pLineElement = pElement
    Case 4, 11:
      Set pElement = New PolygonElement
      Set pFillElement = pElement
    Case 5:
      Set pElement = New RectangleElement
      Set pFillElement = pElement
    Case Else:
      MsgBox "Unexpected Shape Type:  Unable to add graphic..."
  End Select
  
  If pGeometryType = 2 Then
    Dim pGroupElement As IGroupElement2
    Set pGroupElement = New GroupElement
    Dim pSubElement As IElement
    Set pSubElement = New MarkerElement
    Dim lngIndex As Long
    Dim pPtColl As IPointCollection
    Set pPtColl = pNewGeometry
    Dim pPt As IPoint
    For lngIndex = 0 To pPtColl.PointCount - 1
      Set pPt = pPtColl.Point(lngIndex)
      Set pSubElement = New MarkerElement
      pSubElement.Geometry = pPt
      pGroupElement.AddElement pSubElement
    Next lngIndex
    Set pGraphicElement = pGroupElement
    Set pSpatialReference = pGeometry.SpatialReference
    Set pGraphicElement.SpatialReference = pSpatialReference
    Set pElementProperties = pGroupElement
    pElementProperties.Name = strName
    ' ADD GRAPHIC TO GRAPHICS CONTAINER
    pGContainer.AddElement pGroupElement, 0
  Else
    pElement.Geometry = pNewGeometry
    Set pGraphicElement = pElement
    Set pSpatialReference = pGeometry.SpatialReference
    Set pGraphicElement.SpatialReference = pSpatialReference
    Set pElementProperties = pElement
    pElementProperties.Name = strName
    
    If Not pSymbol Is Nothing Then
      Select Case pGeometryType
        Case 1:
          pMarkerElement.Symbol = pSymbol
        Case 3, 6, 13, 14, 15, 16:
          pLineElement.Symbol = pSymbol
        Case 4, 11:
          pFillElement.Symbol = pSymbol
        Case 5:
          pFillElement.Symbol = pSymbol
      End Select
    End If
  
    ' ADD GRAPHIC TO GRAPHICS CONTAINER
    Set Graphic_ReturnElementFromGeometry2 = pElement
    
  End If


  GoTo ClearMemory
ClearMemory:
  Set pGContainer = Nothing
  Set pElement = Nothing
  Set pSpatialReference = Nothing
  Set pGraphicElement = Nothing
  Set pElementProperties = Nothing
  Set pMarkerElement = Nothing
  Set pFillElement = Nothing
  Set pLineElement = Nothing
  Set pClone = Nothing
  Set pNewGeometry = Nothing
  Set pGroupElement = Nothing
  Set pSubElement = Nothing
  Set pPtColl = Nothing
  Set pPt = Nothing
End Function
Public Function Graphic_ReturnElementFromGeometry(ByRef pMxDoc As IMxDocument, ByRef pGeometry As IGeometry, _
    Optional strName As String, Optional AddToView As Boolean) As IElement
  
  
  Dim pGContainer As IGraphicsContainer
  Set pGContainer = pMxDoc.FocusMap
  
  Dim pElement As IElement
  Dim pSpatialReference As ISpatialReference
  Dim pGraphicElement As IGraphicElement
  Dim pElementProperties As IElementProperties
  
  Dim pClone As IClone
  Set pClone = pGeometry
  Dim pNewGeometry As IGeometry
  Set pNewGeometry = pClone.Clone
  
  Dim pGeometryType As esriGeometryType
  pGeometryType = pNewGeometry.GeometryType
  
  'ADD GEOMETRY, NAME AND SPATIAL REFERENCE TO GRAPHIC ELEMENT
  Select Case pGeometryType
    Case 0:
      MsgBox "Null Geometry!  No graphic added..."
    Case 1, 2:
      Set pElement = New MarkerElement
    Case 3, 6, 13, 14, 15, 16:
      Set pElement = New LineElement
    Case 4, 11:
      Set pElement = New PolygonElement
    Case 5:
      Set pElement = New RectangleElement
    Case Else:
      MsgBox "Unexpected Shape Type:  Unable to add graphic..."
      Exit Function
  End Select
  
  
  If pGeometryType = 2 Then
    Dim pGroupElement As IGroupElement2
    Set pGroupElement = New GroupElement
    Dim pSubElement As IElement
    Set pSubElement = New MarkerElement
    Dim lngIndex As Long
    Dim pPtColl As IPointCollection
    Set pPtColl = pNewGeometry
    Dim pPt As IPoint
    For lngIndex = 0 To pPtColl.PointCount - 1
      Set pPt = pPtColl.Point(lngIndex)
      Set pSubElement = New MarkerElement
      pSubElement.Geometry = pPt
      pGroupElement.AddElement pSubElement
    Next lngIndex
    Set pGraphicElement = pGroupElement
    Set pSpatialReference = pGeometry.SpatialReference
    Set pGraphicElement.SpatialReference = pSpatialReference
    Set pElementProperties = pGroupElement
    pElementProperties.Name = strName
    If AddToView Then
      ' ADD GRAPHIC TO GRAPHICS CONTAINER
      pGContainer.AddElement pGroupElement, 0
      'Draw
      pMxDoc.ActiveView.PartialRefresh esriViewGraphics, Nothing, Nothing
    End If

  Else
    pElement.Geometry = pNewGeometry
    Set pGraphicElement = pElement
    Set pSpatialReference = pGeometry.SpatialReference
    Set pGraphicElement.SpatialReference = pSpatialReference
    Set pElementProperties = pElement
    pElementProperties.Name = strName
  
    If AddToView Then
      ' ADD GRAPHIC TO GRAPHICS CONTAINER
      pGContainer.AddElement pElement, 0
      'Draw
      pMxDoc.ActiveView.PartialRefresh esriViewGraphics, Nothing, Nothing
    End If

  End If
  
  
  Set Graphic_ReturnElementFromGeometry = pElement


  GoTo ClearMemory
ClearMemory:
  Set pGContainer = Nothing
  Set pElement = Nothing
  Set pSpatialReference = Nothing
  Set pGraphicElement = Nothing
  Set pElementProperties = Nothing
  Set pClone = Nothing
  Set pNewGeometry = Nothing
  Set pGroupElement = Nothing
  Set pSubElement = Nothing
  Set pPtColl = Nothing
  Set pPt = Nothing

End Function
Public Sub Graphic_MakeFromGeometry(ByRef pMxDoc As IMxDocument, ByRef pGeometry As IGeometry, Optional strName As String, _
    Optional pSymbol As ISymbol, Optional booAddToLayout As Boolean = False)
  
  Dim pGContainer As IGraphicsContainer
  If booAddToLayout Then
    Set pGContainer = pMxDoc.PageLayout
  Else
    Set pGContainer = pMxDoc.FocusMap
  End If
  
  Dim pElement As IElement
  Dim pSpatialReference As ISpatialReference
  Dim pGraphicElement As IGraphicElement
  Dim pElementProperties As IElementProperties
  
  Dim pMarkerElement As IMarkerElement
  Dim pFillElement As IFillShapeElement
  Dim pLineElement As ILineElement
  
  Dim pClone As IClone
  Set pClone = pGeometry
  Dim pNewGeometry As IGeometry
  Set pNewGeometry = pClone.Clone
  
  Dim pGeometryType As esriGeometryType
  pGeometryType = pNewGeometry.GeometryType
  
  'ADD GEOMETRY, NAME AND SPATIAL REFERENCE TO GRAPHIC ELEMENT
  Select Case pGeometryType
    Case 0:
      MsgBox "Null Geometry!  No graphic added..."
    Case 1:
      Set pElement = New MarkerElement
      Set pMarkerElement = pElement
    Case 2:
      Set pElement = New GroupElement
    Case 3, 6, 13, 14, 15, 16:
      Set pElement = New LineElement
      Set pLineElement = pElement
    Case 4, 11:
      Set pElement = New PolygonElement
      Set pFillElement = pElement
    Case 5:
      Set pElement = New RectangleElement
      Set pFillElement = pElement
    Case Else:
      MsgBox "Unexpected Shape Type:  Unable to add graphic..."
  End Select
  
  Dim pGroupElement As IGroupElement2
  Dim pSubElement As IElement
  Dim lngIndex As Long
  Dim pPtColl As IPointCollection
  Dim pPt As IPoint
  
  If pGeometryType = 2 Then
    Set pGroupElement = New GroupElement
    Set pSubElement = New MarkerElement
    Set pPtColl = pNewGeometry
    For lngIndex = 0 To pPtColl.PointCount - 1
      Set pPt = pPtColl.Point(lngIndex)
      Set pSubElement = New MarkerElement
      If Not pSymbol Is Nothing Then
        Set pMarkerElement = pSubElement
        pMarkerElement.Symbol = pSymbol
      End If
      pSubElement.Geometry = pPt
      pGroupElement.AddElement pSubElement
    Next lngIndex
    Set pGraphicElement = pGroupElement
    Set pSpatialReference = pGeometry.SpatialReference
    Set pGraphicElement.SpatialReference = pSpatialReference
    Set pElementProperties = pGroupElement
    pElementProperties.Name = strName
    
    ' ADD GRAPHIC TO GRAPHICS CONTAINER
    pGContainer.AddElement pGroupElement, 0
  Else
    pElement.Geometry = pNewGeometry
    Set pGraphicElement = pElement
    Set pSpatialReference = pGeometry.SpatialReference
    Set pGraphicElement.SpatialReference = pSpatialReference
    Set pElementProperties = pElement
    pElementProperties.Name = strName
    
    If Not pSymbol Is Nothing Then
      Select Case pGeometryType
        Case 1:
          pMarkerElement.Symbol = pSymbol
        Case 3, 6, 13, 14, 15, 16:
          pLineElement.Symbol = pSymbol
        Case 4, 11:
          pFillElement.Symbol = pSymbol
        Case 5:
          pFillElement.Symbol = pSymbol
      End Select
    End If
  
    ' ADD GRAPHIC TO GRAPHICS CONTAINER
    pGContainer.AddElement pElement, 0
  End If

  'Draw
  pMxDoc.ActiveView.PartialRefresh esriViewGraphics, Nothing, Nothing

  GoTo ClearMemory

ClearMemory:
  Set pGContainer = Nothing
  Set pElement = Nothing
  Set pSpatialReference = Nothing
  Set pGraphicElement = Nothing
  Set pElementProperties = Nothing
  Set pMarkerElement = Nothing
  Set pFillElement = Nothing
  Set pLineElement = Nothing
  Set pClone = Nothing
  Set pNewGeometry = Nothing
  Set pGroupElement = Nothing
  Set pSubElement = Nothing
  Set pPtColl = Nothing
  Set pPt = Nothing

End Sub

Public Sub Graphic_MakeFromGeometry_ByMap(ByRef pMap As IMap, ByRef pGeometry As IGeometry, Optional strName As String, _
    Optional pSymbol As ISymbol)

  Dim pActiveView As esriCarto.IActiveView
  
  Dim pGContainer As IGraphicsContainer
  Set pGContainer = pMap
  
  Dim pElement As IElement
  Dim pSpatialReference As ISpatialReference
  Dim pGraphicElement As IGraphicElement
  Dim pElementProperties As IElementProperties
  
  Dim pMarkerElement As IMarkerElement
  Dim pFillElement As IFillShapeElement
  Dim pLineElement As ILineElement
  
  Dim pClone As IClone
  Set pClone = pGeometry
  Dim pNewGeometry As IGeometry
  Set pNewGeometry = pClone.Clone
  
  Dim pGeometryType As esriGeometryType
  pGeometryType = pNewGeometry.GeometryType
  
  'ADD GEOMETRY, NAME AND SPATIAL REFERENCE TO GRAPHIC ELEMENT
  Select Case pGeometryType
    Case 0:
      MsgBox "Null Geometry!  No graphic added..."
    Case 1:
      Set pElement = New MarkerElement
      Set pMarkerElement = pElement
    Case 2:
      Set pElement = New GroupElement
    Case 3, 6, 13, 14, 15, 16:
      Set pElement = New LineElement
      Set pLineElement = pElement
    Case 4, 11:
      Set pElement = New PolygonElement
      Set pFillElement = pElement
    Case 5:
      Set pElement = New RectangleElement
      Set pFillElement = pElement
    Case Else:
      MsgBox "Unexpected Shape Type:  Unable to add graphic..."
  End Select
  
  Dim pGroupElement As IGroupElement2
  Dim pSubElement As IElement
  Dim lngIndex As Long
  Dim pPtColl As IPointCollection
  Dim pPt As IPoint
  
  If pGeometryType = 2 Then
    Set pGroupElement = New GroupElement
    Set pSubElement = New MarkerElement
    Set pPtColl = pNewGeometry
    For lngIndex = 0 To pPtColl.PointCount - 1
      Set pPt = pPtColl.Point(lngIndex)
      Set pSubElement = New MarkerElement
      If Not pSymbol Is Nothing Then
        Set pMarkerElement = pSubElement
        pMarkerElement.Symbol = pSymbol
      End If
      pSubElement.Geometry = pPt
      pGroupElement.AddElement pSubElement
    Next lngIndex
    Set pGraphicElement = pGroupElement
    Set pSpatialReference = pGeometry.SpatialReference
    Set pGraphicElement.SpatialReference = pSpatialReference
    Set pElementProperties = pGroupElement
    pElementProperties.Name = strName
    ' ADD GRAPHIC TO GRAPHICS CONTAINER
    pGContainer.AddElement pGroupElement, 0
  Else
    pElement.Geometry = pNewGeometry
    Set pGraphicElement = pElement
    Set pSpatialReference = pGeometry.SpatialReference
    Set pGraphicElement.SpatialReference = pSpatialReference
    Set pElementProperties = pElement
    pElementProperties.Name = strName
    
    If Not pSymbol Is Nothing Then
      Select Case pGeometryType
        Case 1:
          pMarkerElement.Symbol = pSymbol
        Case 3, 6, 13, 14, 15, 16:
          pLineElement.Symbol = pSymbol
        Case 4, 11:
          pFillElement.Symbol = pSymbol
        Case 5:
          pFillElement.Symbol = pSymbol
      End Select
    End If
  
    ' ADD GRAPHIC TO GRAPHICS CONTAINER
    pGContainer.AddElement pElement, 0
  End If

  'Draw
  Set pActiveView = pMap
  pActiveView.PartialRefresh esriViewGraphics, Nothing, Nothing

  GoTo ClearMemory
ClearMemory:
  Set pActiveView = Nothing
  Set pGContainer = Nothing
  Set pElement = Nothing
  Set pSpatialReference = Nothing
  Set pGraphicElement = Nothing
  Set pElementProperties = Nothing
  Set pMarkerElement = Nothing
  Set pFillElement = Nothing
  Set pLineElement = Nothing
  Set pClone = Nothing
  Set pNewGeometry = Nothing
  Set pGroupElement = Nothing
  Set pSubElement = Nothing
  Set pPtColl = Nothing
  Set pPt = Nothing
End Sub
Public Sub OpenDoc(theDocFilename As String, theDocPath As String)

    ' CHECK IF FILE EXISTS
    Dim booFileExists As Boolean
    Dim theCheckFilename As String
    If Right(theDocPath, 1) = "\" Or Right(theDocPath, 1) = "/" Then
      theCheckFilename = theDocPath & theDocFilename
    Else
      theCheckFilename = theDocPath & "\" & theDocFilename
    End If
    
    booFileExists = Dir(theCheckFilename) <> ""
    
    If booFileExists Then
      Call ShellExecute(0, vbNullString, theDocFilename, vbNullString, theDocPath, 1)
    Else
      MsgBox "Unable to find the following file:" & vbCrLf & vbCrLf & theCheckFilename & vbCrLf & _
            vbCrLf & "Bailing out...", , "Missing File:"
    End If

End Sub

Public Function MakeColorRGB(pRed As Integer, pGreen As Integer, pBlue As Integer) As IColor

  Dim pColor As IRgbColor
  Set pColor = New RgbColor
  pColor.Red = pRed
  pColor.Blue = pBlue
  pColor.Green = pGreen
  pColor.UseWindowsDithering = True
  
  Set MakeColorRGB = pColor

  GoTo ClearMemory

ClearMemory:
  Set pColor = Nothing

End Function

Public Sub ColorToRGB(lngRGB As Long, lngRedToFill As Long, lngGreenToFill As Long, lngBlueToFill As Long)
  
  ' ADAPTED FROM http://www.freevbcode.com/ShowCode.asp?ID=8486
  
  Dim lngColor As Long

  lngColor = lngRGB
  lngBlueToFill = Int(lngColor / &H10000)
  
  lngColor = lngColor - (lngBlueToFill * &H10000)
  lngGreenToFill = Int(lngColor / &H100)
    
  lngColor = lngColor - (lngGreenToFill * &H100)
  lngRedToFill = lngColor

End Sub

Public Function MakeColorHSV(pHue As Integer, pSaturation As Integer, pValue As Integer) As IColor

  Dim pColor As IHsvColor
  Set pColor = New HsvColor
  pColor.Hue = pHue
  pColor.Saturation = pSaturation
  pColor.Value = pValue
  pColor.UseWindowsDithering = True
  
  Set MakeColorHSV = pColor

  GoTo ClearMemory

ClearMemory:
  Set pColor = Nothing

End Function
Public Sub GraphicsSetNameSelected(ByRef pMxDoc As IMxDocument, strName As String)

  Dim pGraphicsContainerSelect As IGraphicsContainerSelect
  
  Set pGraphicsContainerSelect = pMxDoc.FocusMap
  Dim pEnumElement As IEnumElement
  Set pEnumElement = pGraphicsContainerSelect.SelectedElements
  pEnumElement.Reset
  
  Dim pElement As IElement
  Dim pElementProperties As IElementProperties
  
  Set pElement = pEnumElement.Next
  
  While Not pElement Is Nothing
    Set pElementProperties = pElement
    pElementProperties.Name = strName
    Set pElement = pEnumElement.Next
  Wend

  GoTo ClearMemory

ClearMemory:
  Set pGraphicsContainerSelect = Nothing
  Set pEnumElement = Nothing
  Set pElement = Nothing
  Set pElementProperties = Nothing
  
End Sub

Public Function ReturnGraphicsByName(ByRef pMxDoc As IMxDocument, strName As String, _
      Optional AsElements As Boolean) As IArray

  Dim pGraphicsContainer As IGraphicsContainer
  
  Set pGraphicsContainer = pMxDoc.FocusMap
  
  pGraphicsContainer.Reset
  
  Dim pElement As IElement
  Dim pElementProperties As IElementProperties
  
  Set pElement = pGraphicsContainer.Next
  
  Dim pArray As IArray
  Set pArray = New esriSystem.Array
  Dim pGeometry As IGeometry
  Dim pClone As IClone
  
  While Not pElement Is Nothing
    Set pElementProperties = pElement
    If StrComp(pElementProperties.Name, strName, vbTextCompare) = 0 Then
      If AsElements Then
        pArray.Add pElement
      Else
        Set pGeometry = pElement.Geometry
        Set pClone = pGeometry
        pArray.Add pClone.Clone     ' ONLY RETURN A COPY OF THE GEOMETRY; DON'T WANT TO MODIFY ACTUAL GRAPHIC HERE
      End If
    End If
    Set pElement = pGraphicsContainer.Next
    
  Wend
  Set ReturnGraphicsByName = pArray


  GoTo ClearMemory
ClearMemory:
  Set pGraphicsContainer = Nothing
  Set pElement = Nothing
  Set pElementProperties = Nothing
  Set pArray = Nothing
  Set pGeometry = Nothing
  Set pClone = Nothing
End Function

Public Function ReturnGraphicsByType(ByRef pMxDoc As IMxDocument, intGeometryType As esriGeometryType, _
      Optional AsElements As Boolean) As IArray

  Dim pGraphicsContainer As IGraphicsContainer
  
  Set pGraphicsContainer = pMxDoc.FocusMap
  
  pGraphicsContainer.Reset
  
  Dim pElement As IElement
  Dim pElementProperties As IElementProperties
  
  Set pElement = pGraphicsContainer.Next
  
  Dim pArray As IArray
  Set pArray = New esriSystem.Array
  Dim pGeometry As IGeometry
  Dim pClone As IClone
  
  Dim pGeometryType As esriGeometryType
  
  While Not pElement Is Nothing
    Set pElementProperties = pElement
    Set pGeometry = pElement.Geometry
    pGeometryType = pGeometry.GeometryType
    
    If pGeometryType = intGeometryType Then
      If AsElements Then            ' IN THIS CASE RETURN THE ACTUAL GRAPHIC ELEMENT
        pArray.Add pElement
      Else
        Set pClone = pGeometry
        pArray.Add pClone.Clone     ' ONLY RETURN A COPY OF THE GEOMETRY; DON'T WANT TO MODIFY ACTUAL GRAPHIC HERE
      End If
    End If
    Set pElement = pGraphicsContainer.Next
    
  Wend
  Set ReturnGraphicsByType = pArray


  GoTo ClearMemory

ClearMemory:
  Set pGraphicsContainer = Nothing
  Set pElement = Nothing
  Set pElementProperties = Nothing
  Set pArray = Nothing
  Set pGeometry = Nothing
  Set pClone = Nothing

End Function


Public Sub DeleteGraphicsByName(ByRef pMxDoc As IMxDocument, strName As String, _
    Optional booDeleteFromLayout As Boolean = False)

  Dim pGraphicsContainer As IGraphicsContainer
  Dim pActiveView As IActiveView
  
  If booDeleteFromLayout Then
    Set pGraphicsContainer = pMxDoc.PageLayout
  Else
    Set pGraphicsContainer = pMxDoc.FocusMap
  End If
  Set pActiveView = pMxDoc.ActiveView
  Dim pElement As IElement
  Dim pElementProperties As IElementProperties
  Dim pEnvelope As IEnvelope
  
  pGraphicsContainer.Reset
  
  Set pElement = pGraphicsContainer.Next
  
  Dim pDeleteArray As esriSystem.IVariantArray
  Set pDeleteArray = New esriSystem.varArray
  
  While Not pElement Is Nothing
    Set pElementProperties = pElement
    
    If StrComp(pElementProperties.Name, strName, vbTextCompare) = 0 Then
      If (pEnvelope Is Nothing) Then
        Set pEnvelope = pElement.Geometry.Envelope
      Else
        pEnvelope.Union pElement.Geometry.Envelope
      End If
      pDeleteArray.Add pElement
    End If
    Set pElement = pGraphicsContainer.Next
    
  Wend
  
  Dim lngIndex As Long
  If pDeleteArray.Count > 0 Then
    For lngIndex = 0 To pDeleteArray.Count - 1
      Set pElement = pDeleteArray.Element(lngIndex)
      pGraphicsContainer.DeleteElement pElement
    Next lngIndex
  End If
  
  If (Not pEnvelope Is Nothing) Then
    pActiveView.PartialRefresh esriViewGraphics + esriViewGraphicSelection + esriViewGeography, Nothing, pEnvelope
  End If
  
  GoTo ClearMemory

ClearMemory:
  Set pGraphicsContainer = Nothing
  Set pActiveView = Nothing
  Set pElement = Nothing
  Set pElementProperties = Nothing
  Set pEnvelope = Nothing
  Set pDeleteArray = Nothing
End Sub
Public Sub DeleteGraphicsByNameByBounds(ByRef pMxDoc As IMxDocument, strName As String, _
    Optional booDeleteFromLayout As Boolean = False)

  Dim pGraphicsContainer As IGraphicsContainer
  Dim pActiveView As IActiveView
  
  If booDeleteFromLayout Then
    Set pGraphicsContainer = pMxDoc.PageLayout
  Else
    Set pGraphicsContainer = pMxDoc.FocusMap
  End If
  Set pActiveView = pMxDoc.ActiveView
  Dim pElement As IElement
  Dim pElementProperties As IElementProperties
  Dim pEnvelope As IEnvelope
  Dim pTempEnvelope As IEnvelope
  Set pTempEnvelope = New Envelope
  
  pGraphicsContainer.Reset
  
  Set pElement = pGraphicsContainer.Next
  
  Dim pDeleteArray As esriSystem.IVariantArray
  Set pDeleteArray = New esriSystem.varArray
  
  While Not pElement Is Nothing
    Set pElementProperties = pElement
    
    If StrComp(pElementProperties.Name, strName, vbTextCompare) = 0 Then
      If (pEnvelope Is Nothing) Then
        Set pEnvelope = pElement.Geometry.Envelope
        pElement.QueryBounds pMxDoc.ActiveView.ScreenDisplay, pEnvelope
      Else
        pElement.QueryBounds pMxDoc.ActiveView.ScreenDisplay, pTempEnvelope
        pEnvelope.Union pTempEnvelope
      End If
      pDeleteArray.Add pElement
    End If
    Set pElement = pGraphicsContainer.Next
    
  Wend
  
  Dim lngIndex As Long
  If pDeleteArray.Count > 0 Then
    For lngIndex = 0 To pDeleteArray.Count - 1
      Set pElement = pDeleteArray.Element(lngIndex)
      pGraphicsContainer.DeleteElement pElement
    Next lngIndex
  End If
  
  If (Not pEnvelope Is Nothing) Then
    pActiveView.PartialRefresh esriViewGraphics + esriViewGraphicSelection, Nothing, pEnvelope
    DoEvents
  End If
  
  GoTo ClearMemory

ClearMemory:
  Set pGraphicsContainer = Nothing
  Set pActiveView = Nothing
  Set pElement = Nothing
  Set pElementProperties = Nothing
  Set pEnvelope = Nothing
  Set pTempEnvelope = Nothing
  Set pDeleteArray = Nothing

End Sub
Public Function CountGraphicsByName(ByRef pMxDoc As IMxDocument, strName As String, _
      Optional booCountFromLayout As Boolean = False) As Long

  Dim pGraphicsContainer As IGraphicsContainer
  Dim pActiveView As IActiveView
  
  If booCountFromLayout Then
    Set pGraphicsContainer = pMxDoc.PageLayout
  Else
    Set pGraphicsContainer = pMxDoc.FocusMap
  End If
  Set pActiveView = pMxDoc.ActiveView
  Dim pElement As IElement
  Dim pElementProperties As IElementProperties
  
  
  pGraphicsContainer.Reset
  
  Set pElement = pGraphicsContainer.Next
  
  Dim pDeleteArray As esriSystem.IVariantArray
  Set pDeleteArray = New esriSystem.varArray
  
  CountGraphicsByName = 0
  
  While Not pElement Is Nothing
    Set pElementProperties = pElement
    
    If StrComp(pElementProperties.Name, strName, vbTextCompare) = 0 Then
      CountGraphicsByName = CountGraphicsByName + 1
    End If
    Set pElement = pGraphicsContainer.Next
    
  Wend
    
  GoTo ClearMemory

ClearMemory:
  Set pGraphicsContainer = Nothing
  Set pActiveView = Nothing
  Set pElement = Nothing
  Set pElementProperties = Nothing
  Set pDeleteArray = Nothing

End Function

Public Sub DoSpatialQuery(pSourceFeatureLayer As IFeatureLayer, pTargetFeatureLayer As IFeatureLayer, DoAll As Boolean, _
    pSelRelationship As esriSpatialRelEnum)

  Dim pFLayer As IFeatureLayer
   
  ' Specify the polygon layer with currently selected features
  Set pFLayer = pSourceFeatureLayer

  Dim pFeatSel As IFeatureSelection
  Set pFeatSel = pFLayer
  
  Dim pSelSet As ISelectionSet
  
  If DoAll Then
    pFeatSel.SelectFeatures Nothing, esriSelectionResultNew, False
  End If
  
  Set pSelSet = pFeatSel.SelectionSet

  Dim pEnumGeom As IEnumGeometry
  Dim pEnumGeomBind As IEnumGeometryBind

  Set pEnumGeom = New EnumFeatureGeometry
  Set pEnumGeomBind = pEnumGeom
  pEnumGeomBind.BindGeometrySource Nothing, pSelSet

  Dim pGeomFactory As IGeometryFactory
  Set pGeomFactory = New GeometryEnvironment

  Dim pGeom As IGeometry
  Set pGeom = pGeomFactory.CreateGeometryFromEnumerator(pEnumGeom)

  Dim pSpFilter As ISpatialFilter
  Set pSpFilter = New SpatialFilter
  With pSpFilter
    Set .Geometry = pGeom
    .GeometryField = "SHAPE"
    .SpatialRel = pSelRelationship
  End With
  
  Set pFLayer = pTargetFeatureLayer
  Set pFeatSel = pFLayer
  
  pFeatSel.SelectFeatures pSpFilter, esriSelectionResultNew, False

  GoTo ClearMemory

ClearMemory:
  Set pFLayer = Nothing
  Set pFeatSel = Nothing
  Set pSelSet = Nothing
  Set pEnumGeom = Nothing
  Set pEnumGeomBind = Nothing
  Set pGeomFactory = Nothing
  Set pGeom = Nothing
  Set pSpFilter = Nothing

End Sub


Public Function ReturnCurrentMapUnits(pMap As IMap) As String

  Dim intEsriUnits As Integer
  intEsriUnits = pMap.MapUnits
  
  Select Case intEsriUnits
    Case 0
      ReturnCurrentMapUnits = "Unknown"
    Case 1
      ReturnCurrentMapUnits = "Inches"
    Case 2
      ReturnCurrentMapUnits = "Points"
    Case 3
      ReturnCurrentMapUnits = "Feet"
    Case 4
      ReturnCurrentMapUnits = "Yards"
    Case 5
      ReturnCurrentMapUnits = "Miles"
    Case 6
      ReturnCurrentMapUnits = "Nautical Miles"
    Case 7
      ReturnCurrentMapUnits = "Millimeters"
    Case 8
      ReturnCurrentMapUnits = "Centimeters"
    Case 9
      ReturnCurrentMapUnits = "Meters"
    Case 10
      ReturnCurrentMapUnits = "Kilometers"
    Case 11
      ReturnCurrentMapUnits = "Decimal Degrees"
    Case 12
      ReturnCurrentMapUnits = "Decimeters"
  End Select

End Function


Public Sub ForceUppercase(KeyAscii As Integer, txtTextBox As TextBox)
  
  KeyAscii = Asc(UCase(Chr(KeyAscii)))

End Sub

Public Sub CheckNumericReal(KeyAscii As Integer, txtTextBox As TextBox)

Select Case KeyAscii
  Case Is < 8
    KeyAscii = 0
'  Case 8            ' BACKSPACE
  Case 9 To 43
    KeyAscii = 0
  Case 45            ' "-" CHARACTER; INSERT AT BEGINNING OR REMOVE FROM BEGINNING
                     ' RESET CURSOR POSITION TO ORIGINAL CHARACTER LOCATION
    KeyAscii = 0
    Dim strText As String
    Dim lngpos As Long
    strText = txtTextBox.Text
    lngpos = txtTextBox.SelStart
    If Left(strText, 1) = "-" Then
      txtTextBox.Text = Right(strText, Len(strText) - 1)
      If lngpos > 0 Then
        txtTextBox.SelStart = lngpos - 1
      End If
    Else
      txtTextBox.Text = "-" & strText
      txtTextBox.SelStart = lngpos + 1
    End If
  Case 44, 46           ' DECIMAL CHARACTER; ONLY ALLOW ONE PERIOD OR COMMA
    Dim strText2 As String
    strText2 = txtTextBox.Text
    If strText2 = "" Then
      If KeyAscii = 44 Then
        strText2 = "0."
        txtTextBox.SelStart = 2
        KeyAscii = 0
      Else
        strText2 = "0,"
        txtTextBox.SelStart = 2
        KeyAscii = 0
      End If
    ElseIf strText2 = "-" Then
      If KeyAscii = 44 Then
        strText2 = "-0,"
        txtTextBox.SelStart = 3
        KeyAscii = 0
      Else
        strText2 = "-0."
        txtTextBox.SelStart = 3
        KeyAscii = 0
      End If
    Else
      Dim lngPos2 As Long
      lngPos2 = txtTextBox.SelStart
      If Not IsNumeric(Left(strText2, lngPos2) & Chr(KeyAscii) & Right(strText2, Len(strText2) - lngPos2)) Then
        KeyAscii = 0
      End If
    End If
  Case 47          ' "/" CHARACTER
    KeyAscii = 0
  Case Is > 57
    KeyAscii = 0
End Select

End Sub


Public Sub CheckNumericRealPositive(KeyAscii As Integer, txtTextBox As TextBox)

Select Case KeyAscii
  Case Is < 8
    KeyAscii = 0
'  Case 8            ' BACKSPACE
  Case 9 To 43
    KeyAscii = 0
  Case 45            ' "-" CHARACTER
    KeyAscii = 0

  Case 44, 46           ' DECIMAL CHARACTER; ONLY ALLOW ONE PERIOD OR COMMA
    Dim strText2 As String
    strText2 = txtTextBox.Text
    If strText2 = "" Then
      If KeyAscii = 44 Then
        strText2 = "0."
        txtTextBox.SelStart = 2
        KeyAscii = 0
      Else
        strText2 = "0,"
        txtTextBox.SelStart = 2
        KeyAscii = 0
      End If
    Else
      Dim lngPos2 As Long
      lngPos2 = txtTextBox.SelStart
      If Not IsNumeric(Left(strText2, lngPos2) & Chr(KeyAscii) & Right(strText2, Len(strText2) - lngPos2)) Then
        KeyAscii = 0
      End If
    End If
  Case 47          ' "/" CHARACTER
    KeyAscii = 0
  Case Is > 57
    KeyAscii = 0
End Select

End Sub
Public Sub CheckNumericInteger(KeyAscii As Integer, txtTextBox As TextBox)

Select Case KeyAscii
  Case Is < 8
    KeyAscii = 0
'  Case 8            ' BACKSPACE
  Case 9 To 43
    KeyAscii = 0
  Case 45            ' "-" CHARACTER; INSERT AT BEGINNING OR REMOVE FROM BEGINNING
                     ' RESET CURSOR POSITION TO ORIGINAL CHARACTER LOCATION
    KeyAscii = 0
    Dim strText As String
    Dim lngpos As Long
    strText = txtTextBox.Text
    lngpos = txtTextBox.SelStart
    If Left(strText, 1) = "-" Then
      txtTextBox.Text = Right(strText, Len(strText) - 1)
      If lngpos > 0 Then
        txtTextBox.SelStart = lngpos - 1
      End If
    Else
      txtTextBox.Text = "-" & strText
      txtTextBox.SelStart = lngpos + 1
    End If
  Case 44, 46, 47         ' PREVENT PERIOD OR COMMA DECIMAL CHARACTER
    KeyAscii = 0
  Case Is > 57
    KeyAscii = 0
End Select

End Sub

Public Sub CheckNumericIntegerPositive(KeyAscii As Integer, txtTextBox As TextBox)

' ONLY ALLOW NUMBERS AND BACKSPACE
Select Case KeyAscii
  Case Is < 8
    KeyAscii = 0
'  Case 8            ' BACKSPACE
  Case 9 To 47
    KeyAscii = 0
  Case Is > 57
    KeyAscii = 0
End Select

End Sub
Public Function CompareSpatialReferences(ByVal pSourceSR As ISpatialReference, ByVal pTargetSR As ISpatialReference, _
      Optional bSREqual As Boolean, Optional bXYIsEqual As Boolean) As Boolean

  Dim pSourceClone As IClone
  Dim pTargetClone As IClone
  
  Set pSourceClone = pSourceSR
  Set pTargetClone = pTargetSR
  
  'MsgBox "pSourceClone is nothing = " & CStr(pSourceClone Is Nothing) & vbCrLf & _
        "pTargetClone is nothing = " & CStr(pTargetClone Is Nothing)
  
  If pSourceClone Is Nothing And pTargetClone Is Nothing Then
    CompareSpatialReferences = True
  ElseIf pSourceClone Is Nothing Or pTargetClone Is Nothing Then
    CompareSpatialReferences = False
  Else
    
    'Compare the coordinate system component of the spatial reference
    bSREqual = pSourceClone.IsEqual(pTargetClone)
    
    'If the comparison failed, return false and exit
    If Not bSREqual Then
      CompareSpatialReferences = False
      Exit Function
    End If
    
    'We can also compare the XY precision to ensure the spatial references are equal
    Dim pSourceSR2 As ISpatialReference2
    
    Set pSourceSR2 = pSourceSR
    bXYIsEqual = pSourceSR2.IsXYPrecisionEqual(pTargetSR)
    
    'If the comparison failed, return false and exit
    If Not bXYIsEqual Then
      CompareSpatialReferences = False
      Exit Function
    End If
    
    CompareSpatialReferences = True
  End If


  GoTo ClearMemory
ClearMemory:
  Set pSourceClone = Nothing
  Set pTargetClone = Nothing
  Set pSourceSR2 = Nothing
End Function

Public Function ReturnTimeElapsedFromMilliseconds(lngMilliseconds As Long) As String

' SAMPLE CODE
'  Dim datTimeBegan As Long
'  datTimeBegan = GetTickCount()
'  Debug.Print MyGeneralOperations.ReturnTimeElapsedFromMilliseconds(GetTickCount() - datTimeBegan)

  Dim theElapsedTime As Double
  Dim theNumDays As Double
  Dim theNumHours As Double
  Dim theNumMinutes As Double
  Dim theNumSeconds As Double
  
  theElapsedTime = lngMilliseconds / 1000
  
  theNumDays = Int(theElapsedTime / 86400)
  theNumHours = Int((theElapsedTime Mod 86400) / 3600)
  theNumMinutes = Int((theElapsedTime Mod 3600) / 60)
  theNumSeconds = theElapsedTime Mod 60
  
  Dim theDayString As String
  Dim theHourString As String
  Dim theMinString As String
  Dim theSecString As String
  
  If theNumDays = 1 Then
    theDayString = " day"
  Else
    theDayString = " days"
  End If
  
  If theNumHours = 1 Then
    theHourString = " hour"
  Else
    theHourString = " hours"
  End If
  
  If theNumMinutes = 1 Then
    theMinString = " minute"
  Else
    theMinString = " minutes"
  End If
  
  If theNumSeconds = 1 Then
    theSecString = " second..."
  Else
    theSecString = " seconds..."
  End If
  
  Dim theElapsedTimeString As String
  theElapsedTimeString = ""
  If theNumDays > 0 Then
    theElapsedTimeString = theElapsedTimeString & theNumDays & theDayString & ", " & theNumHours & theHourString & ", " & _
    theNumMinutes & theMinString & ", " & theNumSeconds & theSecString
  ElseIf theNumHours > 0 Then
    theElapsedTimeString = theElapsedTimeString & theNumHours & theHourString & ", " & _
    theNumMinutes & theMinString & ", " & theNumSeconds & theSecString
  ElseIf theNumMinutes > 0 Then
    theElapsedTimeString = theElapsedTimeString & _
    theNumMinutes & theMinString & ", " & theNumSeconds & theSecString
  Else
    theElapsedTimeString = theElapsedTimeString & theNumSeconds & theSecString
  End If
  
  ReturnTimeElapsedFromMilliseconds = theElapsedTimeString

End Function
Public Function ReturnTimeElapsed(theTimeBegan As Date, theTimeEnd As Date, Optional strTimeHMS As String) As String
  
  Dim theElapsedTime As Double
  Dim theNumDays As Double
  Dim theNumHours As Double
  Dim theNumMinutes As Double
  Dim theNumSeconds As Double
  
  theElapsedTime = DateDiff("s", theTimeBegan, theTimeEnd)
  
  theNumDays = Int(theElapsedTime / 86400)
  theNumHours = Int((theElapsedTime Mod 86400) / 3600)
  theNumMinutes = Int((theElapsedTime Mod 3600) / 60)
  theNumSeconds = theElapsedTime Mod 60
  
  Dim theDayString As String
  Dim theHourString As String
  Dim theMinString As String
  Dim theSecString As String
  
  If theNumDays = 1 Then
    theDayString = " day"
  Else
    theDayString = " days"
  End If
  
  If theNumHours = 1 Then
    theHourString = " hour"
  Else
    theHourString = " hours"
  End If
  
  If theNumMinutes = 1 Then
    theMinString = " minute"
  Else
    theMinString = " minutes"
  End If
  
  If theNumSeconds = 1 Then
    theSecString = " second..."
  Else
    theSecString = " seconds..."
  End If
  
  Dim theElapsedTimeString As String
  theElapsedTimeString = "Time Elapsed: "
  If theNumDays > 0 Then
    theElapsedTimeString = theElapsedTimeString & theNumDays & theDayString & ", " & theNumHours & theHourString & ", " & _
    theNumMinutes & theMinString & ", " & theNumSeconds & theSecString
  ElseIf theNumHours > 0 Then
    theElapsedTimeString = theElapsedTimeString & theNumHours & theHourString & ", " & _
    theNumMinutes & theMinString & ", " & theNumSeconds & theSecString
  ElseIf theNumMinutes > 0 Then
    theElapsedTimeString = theElapsedTimeString & _
    theNumMinutes & theMinString & ", " & theNumSeconds & theSecString
  Else
    theElapsedTimeString = theElapsedTimeString & theNumSeconds & theSecString
  End If
  
  ReturnTimeElapsed = "Analysis Began: " & Format(theTimeBegan, "long date") & ";  " & Format(theTimeBegan, "long time") & vbCrLf & _
                      "Analysis Complete: " & Format(theTimeEnd, "long date") & ";  " & Format(theTimeEnd, "long time") & vbCrLf & _
                      theElapsedTimeString & vbCrLf & vbCrLf

                      
  strTimeHMS = Replace(theElapsedTimeString, "Time Elapsed: ", "", , , vbTextCompare)
  strTimeHMS = Replace(strTimeHMS, "...", "", , , vbTextCompare)
  strTimeHMS = Trim(strTimeHMS)

End Function


Public Function ReturnTimeElapsedRTF(theTimeBegan As Date, theTimeEnd As Date, lngFontSize As Long) As String

  Dim theElapsedTime As Double
  Dim theNumDays As Double
  Dim theNumHours As Double
  Dim theNumMinutes As Double
  Dim theNumSeconds As Double
  
  theElapsedTime = CDbl(DateDiff("s", theTimeBegan, theTimeEnd))
'  MsgBox "theElapsedTime = " & CStr(theElapsedTime)
  theNumDays = Int(CDbl(theElapsedTime / 86400#))
  theNumHours = Int(CDbl(theElapsedTime Mod 86400#) / 3600#)
  theNumMinutes = Int(CDbl(theElapsedTime Mod 3600#) / 60)
  theNumSeconds = theElapsedTime Mod 60
  
  Dim theDayString As String
  Dim theHourString As String
  Dim theMinString As String
  Dim theSecString As String
  
  If theNumDays = 1 Then
    theDayString = " day"
  Else
    theDayString = " days"
  End If
  
  If theNumHours = 1 Then
    theHourString = " hour"
  Else
    theHourString = " hours"
  End If
  
  If theNumMinutes = 1 Then
    theMinString = " minute"
  Else
    theMinString = " minutes"
  End If
  
  If theNumSeconds = 1 Then
    theSecString = " second..."
  Else
    theSecString = " seconds..."
  End If
  
  Dim theElapsedTimeString As String
  theElapsedTimeString = "Time Elapsed: "
  If theNumDays > 0 Then
    theElapsedTimeString = theElapsedTimeString & theNumDays & theDayString & ", " & theNumHours & theHourString & ", " & _
    theNumMinutes & theMinString & ", " & theNumSeconds & theSecString
  ElseIf theNumHours > 0 Then
    theElapsedTimeString = theElapsedTimeString & theNumHours & theHourString & ", " & _
    theNumMinutes & theMinString & ", " & theNumSeconds & theSecString
  ElseIf theNumMinutes > 0 Then
    theElapsedTimeString = theElapsedTimeString & _
    theNumMinutes & theMinString & ", " & theNumSeconds & theSecString
  Else
    theElapsedTimeString = theElapsedTimeString & theNumSeconds & theSecString
  End If
  
  ReturnTimeElapsedRTF = _
    "\b\fszzzFontSizezzz Analysis Began:\b0\i    zzzTimeBeganzzz\par" & vbCrLf & _
    "\b\i0 Analysis Complete:\b0\i    zzzTimeEndzzz\par" & vbCrLf & _
    "\b\i0 Time ElapseE:\b0\i    zzzTimeElapsedzzz\par" & vbCrLf
  
'  ReturnTimeElapsedRTF = "{\rtf1\ansi\ansicpg1252\deff0\deflang1033{\fonttbl{\f0\fswiss\fcharset0 Arial;}}" & vbCrLf & _
      "{\*\generator Msftedit 5.41.15.1507;}\viewkind4\uc1\pard\b\f0\fszzzFontSizezzz Analysis Began:\b0\i    zzzTimeBeganzzz\par" & vbCrLf & _
      "\b\i0 Analysis Complete:\b0\i    zzzTimeEndzzz\par" & vbCrLf & _
      "\b\i0 Time ElapseE:\b0\i    zzzTimeElapsedzzz\par" & vbCrLf & _
      "}"
  Dim strTimeBegan As String
  strTimeBegan = CStr(Format(theTimeBegan, "long date")) & " at " & CStr(Format(theTimeBegan, "long time"))
  Dim strTimeEnd As String
  strTimeEnd = CStr(Format(theTimeEnd, "long date")) & " at " & CStr(Format(theTimeEnd, "long time"))
  
  ReturnTimeElapsedRTF = Replace(ReturnTimeElapsedRTF, "zzzFontSizezzz", CStr(lngFontSize * 2))
  ReturnTimeElapsedRTF = Replace(ReturnTimeElapsedRTF, "zzzTimeBeganzzz", strTimeBegan)
  ReturnTimeElapsedRTF = Replace(ReturnTimeElapsedRTF, "zzzTimeEndzzz", strTimeEnd)
  ReturnTimeElapsedRTF = Replace(ReturnTimeElapsedRTF, "zzzTimeElapsedzzz", theElapsedTimeString)
  
'  ReturnTimeElapsedRTF = "Analysis Began: " & Format(theTimeBegan, "long date") & ";  " & Format(theTimeBegan, "long time") & vbCrLf & _
                      "Analysis Complete: " & Format(theTimeEnd, "long date") & ";  " & Format(theTimeEnd, "long time") & vbCrLf & _
                      theElapsedTimeString & vbCrLf & vbCrLf

End Function

Public Sub RemoveKeyFromCollection(colCollection As Collection, strKey As String)

  On Error Resume Next
  
  colCollection.Remove strKey

End Sub

Public Function CheckCollectionForKey(colCollection As Collection, strKey As String) As Boolean
  On Error GoTo ErrHandler
  
  CheckCollectionForKey = True
  Dim lngVarType As Long
  lngVarType = VarType(colCollection.Item(strKey))
  
  Exit Function
ErrHandler:
  CheckCollectionForKey = False

End Function

Public Sub EnableSelectTool(pApp As IApplication)

    Dim pUID As New UID
    Dim pCmdItem As ICommandItem
    ' Use the GUID of the Select Elements command
    pUID.Value = "{C22579D1-BC17-11D0-8667-0000F8751720}"
    Set pCmdItem = pApp.Document.CommandBars.Find(pUID)
    pCmdItem.Execute

  GoTo ClearMemory

ClearMemory:
  Set pUID = Nothing
  Set pCmdItem = Nothing

End Sub

Public Function CreateNestedFoldersByPath(ByVal completeDirectory As String) As Long

  'ADAPTED FROM http://vbnet.mvps.org/index.html?code/file/nested.htm
  'Visual Basic File Routines
  'CreateDirectory: Creating Nested Folders
  '
  'Posted:     Saturday September 19, 1998
  'Updated:    Wednesday December 05, 2007
  '
  'Applies to:     VB4-32, VB5, VB6
  'Developed with:     VB5, Windows 98
  'OS restrictions:    None
  'Author:     VBnet - Randy Birch

  'creates nested directories on the drive
  'included in the path by parsing the final
  'directory string into a directory array,
  'and looping through each to create the final path.
  
  'The path could be passed to this method as a
  'pre-filled array, reducing the code.
  
   Dim r As Long
   Dim SA As SECURITY_ATTRIBUTES
   Dim drivePart As String
   Dim newDirectory  As String
   Dim Item As String
   Dim sfolders() As String
   Dim pos As Long
   Dim x As Long
      
  'must have a trailing slash for
  'the GetPart routine below
   If Right$(completeDirectory, 1) <> "\" Then
      completeDirectory = completeDirectory & "\"
   End If
  
  'if there is a drive in the string, get it
  'else, just use nothing - assumes current drive
   pos = InStr(completeDirectory, ":")

   If pos Then
      drivePart = GetPart(completeDirectory, "\")
   Else: drivePart = ""
   End If

  'now get the rest of the items that
  'make up the string
   Do Until completeDirectory = ""

    'strip off one item (i.e. "Files\")
     Item = GetPart(completeDirectory, "\")

    'add it to an array for later use, and
    'if this is the first item (x=0),
    'append the drivepart
     ReDim Preserve sfolders(0 To x) As String

     If x = 0 Then Item = drivePart & Item
     sfolders(x) = Item
     
    'increment the array counter
     x = x + 1

   Loop

  'Now create the directories.
  'Because the first directory is
  '0 in the array, reinitialize x to -1
   x = -1
   
   Do
   
      x = x + 1
     'just keep appending the folders in the
     'array to newDirectory.  When x=0 ,
     'newDirectory is "", so the
     'newDirectory gets assigned drive:\firstfolder.
     
     'Subsequent loops adds the next member of the
     'array to the path, forming a fully qualified
     'path to the new directory.
      newDirectory = newDirectory & sfolders(x)
      
     'the only member of the SA type needed (on
     'a win95/98 system at least)
      SA.nLength = LenB(SA)
      
      Call CreateDirectory(newDirectory, SA)
       
   Loop Until x = UBound(sfolders)
   
  'done. Return x, but add 1 for the 0-based array.
   CreateNestedFoldersByPath = x + 1


  GoTo ClearMemory
ClearMemory:
  Erase sfolders

End Function

Public Function GetPart(startStrg As String, delimiter As String) As String

'takes a string separated by "delimiter",
'splits off 1 item, and shortens the string
'so that the next item is ready for removal.

  Dim C As Integer
  Dim Item As String
  
  C = 1
  
  Do

    If Mid$(startStrg, C, 1) = delimiter Then
      
      Item = Mid$(startStrg, 1, C)
      startStrg = Mid$(startStrg, C + 1, Len(startStrg))
      GetPart = Item
      Exit Function
    
    End If

    C = C + 1

  Loop

End Function

Public Function ReturnTimeStamp() As String

  ReturnTimeStamp = CStr(Format(Now, "yyyymmdd_HhNnSs"))

End Function


Private Function IsNaN(expression As Variant) As Boolean

  On Error Resume Next
  If Not IsNumeric(expression) Then
    IsNaN = False
    Exit Function
  End If
  If (CStr(expression) = "-1.#QNAN") Or (CStr(expression) = "1,#QNAN") Then ' can vary by locale
    IsNaN = True
  Else
    IsNaN = False
  End If

End Function


Public Function MakeUniqueRasterName(pRasterType As JenDatasetTypes, strWSPath As String, strRasterName As String, _
      Optional booTrimForGRID As Boolean) As String
  
  ' ONLY RETURNS ACTUAL RASTER NAME
  
  Dim pWSFact As IWorkspaceFactory
  
  Dim pWorkspace As IWorkspace
  Dim pWorkspace2 As IWorkspace2
  Dim pWsEx As IRasterWorkspaceEx
  Dim pWS As IRasterWorkspace
  
  Dim lngIndex As Long
  lngIndex = 1
  
  Dim strNewName As String
  Dim strBaseName As String
  strNewName = strRasterName
  strBaseName = strRasterName
  
  Dim strIndex As String
  
  Select Case pRasterType
    Case ENUM_FileGDB
      Set pWSFact = New FileGDBWorkspaceFactory
      Set pWsEx = pWSFact.OpenFromFile(strWSPath, 0)
      Set pWorkspace2 = pWsEx
      Do Until Not pWorkspace2.NameExists(esriDTRasterDataset, strNewName)
        lngIndex = lngIndex + 1
        strNewName = strBaseName & "_" & CStr(lngIndex)
      Loop
      MakeUniqueRasterName = strNewName
    Case ENUM_PersonalGDB
      Set pWSFact = New AccessWorkspaceFactory
      Set pWsEx = pWSFact.OpenFromFile(strWSPath, 0)
      Set pWorkspace2 = pWsEx
      Do Until Not pWorkspace2.NameExists(esriDTRasterDataset, strNewName)
        lngIndex = lngIndex + 1
        strNewName = strBaseName & "_" & CStr(lngIndex)
      Loop
      MakeUniqueRasterName = strNewName
    Case ENUM_File_Raster
      
      If booTrimForGRID Then
        ' FOLLOW GRID RESTRICTIONS
        strNewName = Left(ReplaceBadChars(strNewName, True, True, True, False), 13)
        strBaseName = strNewName
      End If
      
      Set pWSFact = New RasterWorkspaceFactory
      Set pWS = pWSFact.OpenFromFile(strWSPath, 0)
      Set pWorkspace = pWS
      
      Dim pEnumDatasetName As IEnumDatasetName
      Set pEnumDatasetName = pWorkspace.DatasetNames(esriDTRasterDataset)
      Dim pDatasetName As IDatasetName
      
      Dim booNameExists As Boolean
      booNameExists = True
      
      Do Until booNameExists = False
        booNameExists = False
                    
        pEnumDatasetName.Reset
        Set pDatasetName = pEnumDatasetName.Next
        
        Do Until pDatasetName Is Nothing
          If StrComp(pDatasetName.Name, strNewName, vbTextCompare) = 0 Then
            booNameExists = True
            Set pDatasetName = Nothing
          Else
            Set pDatasetName = pEnumDatasetName.Next
          End If
        Loop
        
        If booNameExists Then
          lngIndex = lngIndex + 1
          strIndex = "_" & CStr(lngIndex)
          If booTrimForGRID Then
            strNewName = Left(strBaseName, 13 - Len(strIndex)) & strIndex
          Else
            strNewName = strBaseName & strIndex
          End If
        End If
      Loop
      
      MakeUniqueRasterName = strNewName
  End Select

  GoTo ClearMemory

ClearMemory:
  Set pWSFact = Nothing
  Set pWorkspace = Nothing
  Set pWorkspace2 = Nothing
  Set pWsEx = Nothing
  Set pWS = Nothing
  Set pEnumDatasetName = Nothing
  Set pDatasetName = Nothing

End Function

Public Function CursorToSet_Features(pFCursor As IFeatureCursor) As esriSystem.ISet
 
  'http://forums.esri.com/Thread.asp?c=93&f=992&t=79743&mc=16.
  Dim pFeature As IFeature
  Set CursorToSet_Features = New esriSystem.Set
  Set pFeature = pFCursor.NextFeature
  Do While Not pFeature Is Nothing
    CursorToSet_Features.Add pFeature
   Set pFeature = pFCursor.NextFeature
  Loop
 
  CursorToSet_Features.Reset
 
  GoTo ClearMemory
 
ClearMemory:
  Set pFeature = Nothing
End Function
 
Public Function CursorToSet_TableRow(ByVal pCursor As ICursor) As esriSystem.ISet
 
' adapted from VB.NET
' http://forums.esri.com/Thread.asp?c=93&f=992&t=79743&mc=19
 
  Dim pSSet As esriSystem.ISet
  Dim pRow As IRow
 
  'populate the set with the cursor
  Set pSSet = New esriSystem.Set
  Set pRow = pCursor.NextRow
  Do Until pRow Is Nothing
    pSSet.Add pRow
    Set pRow = pCursor.NextRow
  Loop
  pSSet.Reset
 
  Set CursorToSet_TableRow = pSSet
 
 
  GoTo ClearMemory
ClearMemory:
  Set pSSet = Nothing
  Set pRow = Nothing
 
End Function
 
Public Function CursorToVariant_Features(pFCursor As IFeatureCursor) As Variant()
 
  'http://forums.esri.com/Thread.asp?c=93&f=992&t=79743&mc=16.
  Dim pFeature As IFeature
  Dim lngCounter As Long
  Dim varReturn() As Variant
 
  lngCounter = -1
  Set pFeature = pFCursor.NextFeature
  Do While Not pFeature Is Nothing
    lngCounter = lngCounter + 1
    ReDim Preserve varReturn(lngCounter)
    Set varReturn(lngCounter) = pFeature
    Set pFeature = pFCursor.NextFeature
  Loop
 
  CursorToVariant_Features = varReturn
 
  GoTo ClearMemory
 
ClearMemory:
  Erase varReturn
  Set pFeature = Nothing
End Function
 
Public Function CursorToVariant_TableRow(ByVal pCursor As ICursor) As Variant()

' adapted from VB.NET
' http://forums.esri.com/Thread.asp?c=93&f=992&t=79743&mc=19

  Dim pRow As IRow
  Dim lngCounter As Long
  Dim varReturn() As Variant

  lngCounter = -1

  'populate the set with the cursor
  Set pRow = pCursor.NextRow

  Do Until pRow Is Nothing
    lngCounter = lngCounter + 1
    ReDim Preserve varReturn(lngCounter)
    Set varReturn(lngCounter) = pRow
    Set pRow = pCursor.NextRow
  Loop
 
  CursorToVariant_TableRow = varReturn
 
 
  GoTo ClearMemory
ClearMemory:
  Erase varReturn
  Set pRow = Nothing
 
End Function

Public Sub MakeRandomNormalPolar(dblMean As Double, dblSD As Double, dblRand1 As Double, Optional dblRand2 As Double = -999)

'  Static dblSeed
' Based on Marsaglia's variation of the Box-Muller Transformation (http://en.wikipedia.org/wiki/Box-Muller_transform):
'        z1 = u * sqrt((-2 ln(s))/s)
'        z2 = v * sqrt((-2 ln(s))/s)
'      where:
'         s = u^2 + v^2, and is not equal to 0 or > 1
'        x1 = first uniform random number (between -1 and 1)
'        x2 = second uniform random number (between -1 and 1)
'        z1 = first normally distributed random number
'        z2 = second normally distributed random number.

' RANDOMIZE FIRST!
'  If dblSeed = 0 Then
'    Randomize            ' USE SYSTEM TIMER
'  Else
'    Randomize dblSeed
'  End If
  
  Dim dblX1 As Double
  Dim dblX2 As Double
  Dim s As Double
  s = 5
  Do Until s <= 1 And s <> 0
    dblX1 = (Rnd * 2) - 1
    dblX2 = (Rnd * 2) - 1
    s = dblX2 ^ 2 + dblX1 ^ 2
  Loop
  
  dblRand1 = (dblX1 * Sqr(-2 * Log(s) / s) * dblSD#) + dblMean#
  dblRand2 = (dblX2 * Sqr(-2 * Log(s) / s) * dblSD#) + dblMean#
  
'  dblSeed = dblRand2

End Sub

Public Function CreateInMemoryFeatureClass(pGeometryArray As esriSystem.IArray, _
    Optional pValueArray As esriSystem.IVariantArray, Optional pTemplateField As iField) As IFeatureClass

    ' create an inmemory featureclass
    ' ADAPTED FROM KIRK KUYKENDALL
    ' http://forums.esri.com/Thread.asp?c=93&f=993&t=210767
    
    ' ASSUMES ONE OR MORE GEOMETRIES, ALL OF WHICH ARE IN THE SAME PROJECTION
    Dim pGeom As IGeometry
    Set pGeom = pGeometryArray.Element(0)
    
    
    Dim pSpRef As ISpatialReference
    Set pSpRef = pGeom.SpatialReference
    
'    If Not pDomain Is Nothing Then
'      Set pClone = pDomain
'      Set pNewDomain = pClone.Clone
'      pNewDomain.Expand 1.01, 1.01, True
'      pSpRef.SetDomain pNewDomain.XMin, pNewDomain.XMax, pNewDomain.YMin, pNewDomain.YMax
'    End If
    
    Dim pSpRefRes As ISpatialReferenceResolution
    Set pSpRefRes = pSpRef
    pSpRefRes.ConstructFromHorizon
    
    Dim pWSF As IWorkspaceFactory
    Set pWSF = New InMemoryWorkspaceFactory
    
    Dim pName As IName
    Set pName = pWSF.Create("", "inmemory", Nothing, 0)
    Dim pFWS As IFeatureWorkspace
    Set pFWS = pName.Open
            
    Dim pFields As IFields
    Dim pFieldsEdit As IFieldsEdit
    Dim pField As iField
    Dim pFieldEdit As IFieldEdit
   
    Set pFields = New Fields
    Set pFieldsEdit = pFields
      
    '' create the geometry field
    
    Dim pGeomDef As IGeometryDef
    Dim pGeomDefEdit As IGeometryDefEdit
    Set pGeomDef = New GeometryDef
    Set pGeomDefEdit = pGeomDef
    
    '' assign the geometry definiton properties.
    With pGeomDefEdit
      .GeometryType = pGeom.GeometryType
      .GridCount = 1
      .GridSize(0) = 10
      .AvgNumPoints = 2
      .HasM = False
      .HasZ = False
      Set .SpatialReference = pGeom.SpatialReference
    End With
    
    Set pField = New Field
    Set pFieldEdit = pField
    
    pFieldEdit.Name = "Shape"
    pFieldEdit.AliasName = "geometry"
    pFieldEdit.Type = esriFieldTypeGeometry
    Set pFieldEdit.GeometryDef = pGeomDef
    pFieldsEdit.AddField pField
    
    Dim pClone As IClone
    Dim booAddAttribute As Boolean
    booAddAttribute = Not pValueArray Is Nothing And Not pTemplateField Is Nothing
    Dim varVal As Variant
    
    If booAddAttribute Then
      Set pClone = pTemplateField
      Set pField = pClone.Clone
      pFieldsEdit.AddField pField
    Else
      Set pField = New Field
      Set pFieldEdit = pField
      With pFieldEdit
        .Name = "Unique_ID"
        .Type = esriFieldTypeInteger
      End With
      pFieldsEdit.AddField pField
    End If
    
    Dim pCLSID As UID
    Set pCLSID = New UID
    pCLSID.Value = "esriGeoDatabase.Feature"
  
    Dim pInMemFC As IFeatureClass
    Set pInMemFC = pFWS.CreateFeatureClass("In_Memory", pFields, _
                             pCLSID, Nothing, esriFTSimple, _
                             "Shape", "")
    
    Dim lngIDIndex As Long
    
    Dim lngAttIndex As Long
    If booAddAttribute Then
      lngAttIndex = pInMemFC.FindField(pTemplateField.Name)
    Else
      lngIDIndex = pInMemFC.FindField("Unique_ID")
    End If
    
    Dim lngIndex As Long
    Dim pOutFeat As IFeature
    For lngIndex = 0 To pGeometryArray.Count - 1
      Set pClone = pGeometryArray.Element(lngIndex)
      Set pGeom = pClone.Clone
      Set pOutFeat = pInMemFC.CreateFeature
      Set pOutFeat.Shape = pGeom
      If booAddAttribute Then
        varVal = pValueArray.Element(lngIndex)
        pOutFeat.Value(lngAttIndex) = varVal
      Else
        pOutFeat.Value(lngIDIndex) = lngIndex + 1
      End If
      pOutFeat.Store
    Next lngIndex
    
    Set CreateInMemoryFeatureClass = pInMemFC
      
  GoTo ClearMemory
ClearMemory:
  Set pGeom = Nothing
  Set pSpRef = Nothing
  Set pSpRefRes = Nothing
  Set pWSF = Nothing
  Set pName = Nothing
  Set pFWS = Nothing
  Set pFields = Nothing
  Set pFieldsEdit = Nothing
  Set pField = Nothing
  Set pFieldEdit = Nothing
  Set pGeomDef = Nothing
  Set pGeomDefEdit = Nothing
  Set pClone = Nothing
  varVal = Null
  Set pCLSID = Nothing
  Set pInMemFC = Nothing
  Set pOutFeat = Nothing

End Function

Public Function ReturnTempRasterWorkspace() As IWorkspace
    ' ADAPTED FROM ESRI SAMPLE http://edndoc.esri.com/arcobjects/9.2/CPP_VB6_VBA_VCPP_Doc/COM_Samples_Docs/SpatialAnalyst/ImageUpdate/Visual_Basic/ImageUpdate.frm.htm
    ' Given a pathname, returns the raster workspace object for that path
    On Error GoTo erh
    
    ' Create a Rasterworkspace
    Dim sPath As String
    sPath = Environ("TEMP") ' Get temp directory
    
    Dim pWSF As IWorkspaceFactory
    Set pWSF = New RasterWorkspaceFactory
    Dim pWS As IWorkspace
    Set pWS = pWSF.OpenFromFile(sPath, 0)
    Set ReturnTempRasterWorkspace = pWS
    
    GoTo ClearMemory
    Exit Function
erh:
    Set ReturnTempRasterWorkspace = Nothing


ClearMemory:
  Set pWSF = Nothing
  Set pWS = Nothing

End Function
Public Function MakeUniqueRasterName2(pWS As IWorkspace, strBaseName As String, booRestrictTo13CharForGrid As Boolean) As String
  
  Dim pNames As IEnumDatasetName
  Set pNames = pWS.DatasetNames(esriDTAny)
  Dim pDatasetName As IDatasetName
  
  Dim booNameExists As Boolean
  booNameExists = True
  
  Dim lngCounter As Long
  lngCounter = 1
  Dim strCounter As String
  Dim strUniqueName As String
  strUniqueName = strBaseName
  
  Do Until booNameExists = False
    booNameExists = False
    pNames.Reset
    Set pDatasetName = pNames.Next
    Do Until pDatasetName Is Nothing
      Debug.Print "Checking " & pDatasetName.Name
      If UCase(pDatasetName.Name) = UCase(strUniqueName) Then
        booNameExists = True
        Exit Do
      End If
      Set pDatasetName = pNames.Next
    Loop
    If booNameExists = True Then
      lngCounter = lngCounter + 1
      strCounter = CStr(lngCounter)
      If booRestrictTo13CharForGrid Then
        strUniqueName = Left(strBaseName, 12 - Len(strCounter)) & "_" & strCounter
      Else
        strUniqueName = strBaseName & "_" & strCounter
      End If
    End If
  Loop
  MakeUniqueRasterName2 = strUniqueName

  GoTo ClearMemory

ClearMemory:
  Set pNames = Nothing
  Set pDatasetName = Nothing

End Function

Public Function ReturnTableByName(strName As String, pMap As IMap) As IStandaloneTable

  Set ReturnTableByName = Nothing
  Dim pStTabCol As IStandaloneTableCollection
  Set pStTabCol = pMap
  
  Dim pStTab As IStandaloneTable
  Dim lngIndex As Long
  For lngIndex = 0 To pStTabCol.StandaloneTableCount - 1
    Set pStTab = pStTabCol.StandaloneTable(lngIndex)
    If StrComp(pStTab.Name, strName, vbTextCompare) = 0 Then
      Set ReturnTableByName = pStTab
      Exit For
    End If
  Next lngIndex

  GoTo ClearMemory

ClearMemory:
  Set pStTabCol = Nothing
  Set pStTab = Nothing

End Function

Public Function CreateInMemoryFeatureClass2(pGeometryArray As esriSystem.IArray, _
    Optional pValueArray As esriSystem.IVariantArray, Optional pTemplateFields As esriSystem.IVariantArray, _
    Optional pApp As IApplication, Optional strStatusMessage As String = "") As IFeatureClass
    

    ' create an inmemory featureclass
    ' ADAPTED FROM KIRK KUYKENDALL
    ' http://forums.esri.com/Thread.asp?c=93&f=993&t=210767
    
    Dim pSBar As IStatusBar
    Dim pPro As IStepProgressor
    Dim dateRunningTime As Date
    Dim strHeader As String
'    Dim strMessage As String
'    Dim lngTimer As Long
    Dim lngTotalCount As Long
    
    If Not pApp Is Nothing Then
      ' FOR PROGRESS METER
      Set pSBar = pApp.StatusBar
      Set pPro = pSBar.ProgressBar
      dateRunningTime = Now
      strHeader = strStatusMessage
      lngTotalCount = pGeometryArray.Count
      pSBar.ShowProgressBar strStatusMessage, 0, lngTotalCount, 1, True
      pPro.position = 0
    End If
        
    ' ASSUMES ONE OR MORE GEOMETRIES, ALL OF WHICH ARE IN THE SAME PROJECTION
    Dim pGeom As IGeometry
    Set pGeom = pGeometryArray.Element(0)
    
    Dim pSpRef As ISpatialReference
    Set pSpRef = pGeom.SpatialReference
    Dim pClone As IClone
    
'    If Not pDomain Is Nothing Then
'      Set pClone = pDomain
'      Set pNewDomain = pClone.Clone
'      pNewDomain.Expand 1.01, 1.01, True
'      pSpRef.SetDomain pNewDomain.XMin, pNewDomain.XMax, pNewDomain.YMin, pNewDomain.YMax
'    End If
    
    Dim pSpRefRes As ISpatialReferenceResolution
    Set pSpRefRes = pSpRef
    pSpRefRes.ConstructFromHorizon
    Dim pWSF As IWorkspaceFactory
    Set pWSF = New InMemoryWorkspaceFactory
    
    Dim pName As IName
    Set pName = pWSF.Create("", "inmemory", Nothing, 0)
    Dim pFWS As IFeatureWorkspace
    Set pFWS = pName.Open
    
'    Dim pFlds As IFields
        
    Dim pFields As IFields
    Dim pFieldsEdit As IFieldsEdit
    Dim pField As iField
    Dim pFieldEdit As IFieldEdit
   
    Set pFields = New Fields
    Set pFieldsEdit = pFields
      
    '' create the geometry field
    
    Dim pGeomDef As IGeometryDef
    Dim pGeomDefEdit As IGeometryDefEdit
    Set pGeomDef = New GeometryDef
    Set pGeomDefEdit = pGeomDef
    
'    MsgBox pGeom.SpatialReference.Name
    
    '' assign the geometry definiton properties.
    With pGeomDefEdit
      .GeometryType = pGeom.GeometryType
      .GridCount = 1
      .GridSize(0) = 0
      If pGeom.GeometryType = esriGeometryPoint Then
        .AvgNumPoints = 1
      Else
        .AvgNumPoints = 5
      End If
      .HasM = False
      .HasZ = False
      Set .SpatialReference = pGeom.SpatialReference
    End With
    
    Set pField = New Field
    Set pFieldEdit = pField
    
    pFieldEdit.Name = "Shape"
    pFieldEdit.AliasName = "geometry"
    pFieldEdit.Type = esriFieldTypeGeometry
    Set pFieldEdit.GeometryDef = pGeomDef
    pFieldsEdit.AddField pField
    

    Dim booAddAttribute As Boolean
    booAddAttribute = Not pValueArray Is Nothing And Not pTemplateFields Is Nothing
    Dim varVal As Variant
    
    Dim lngIndex As Long
    Dim pSubArray As esriSystem.IVariantArray
    Dim pTemplateField As iField
    
    Dim lngIDIndex() As Long
    Dim pTempField As iField
    If booAddAttribute Then
      For lngIndex = 0 To pTemplateFields.Count - 1
        Set pTemplateField = pTemplateFields.Element(lngIndex)
        Set pClone = pTemplateField
        Set pTempField = pClone.Clone
        pFieldsEdit.AddField pTempField
      Next lngIndex
      ReDim lngIDIndex(pTemplateFields.Count - 1)
    Else
      Set pField = New Field
      Set pFieldEdit = pField
      With pFieldEdit
        .Name = "Unique_ID"
        .Type = esriFieldTypeInteger
      End With
      pFieldsEdit.AddField pField
      ReDim lngIDIndex(0)
    End If
    
    Dim pCLSID As UID
    Set pCLSID = New UID
    pCLSID.Value = "esriGeoDatabase.Feature"
  
    Dim pInMemFC As IFeatureClass
    Set pInMemFC = pFWS.CreateFeatureClass("In_Memory", pFields, _
                             pCLSID, Nothing, esriFTSimple, _
                             "Shape", "")
    
    
    If booAddAttribute Then
      For lngIndex = 0 To pTemplateFields.Count - 1
        Set pTemplateField = pTemplateFields.Element(lngIndex)
        lngIDIndex(lngIndex) = pInMemFC.FindField(pTemplateField.Name)
      Next lngIndex
    Else
      lngIDIndex(0) = pInMemFC.FindField("Unique_ID")
    End If
    
    Dim pOutFeat As IFeature
    Dim lngIndex2 As Long
    
    For lngIndex = 0 To pGeometryArray.Count - 1
      Set pClone = pGeometryArray.Element(lngIndex)
      Set pGeom = pClone.Clone
      Set pOutFeat = pInMemFC.CreateFeature
      Set pOutFeat.Shape = pGeom
      If booAddAttribute Then
        Set pSubArray = pValueArray.Element(lngIndex)
        For lngIndex2 = 0 To pSubArray.Count - 1
          varVal = pSubArray.Element(lngIndex2)
          pOutFeat.Value(lngIDIndex(lngIndex2)) = varVal
        Next lngIndex2
      Else
        pOutFeat.Value(lngIDIndex(0)) = lngIndex + 1
      End If
      pOutFeat.Store
      If Not pApp Is Nothing Then
        pPro.Step
      End If
    Next lngIndex
    
    If Not pApp Is Nothing Then
      pPro.position = 1
      pSBar.HideProgressBar
    End If
    Set CreateInMemoryFeatureClass2 = pInMemFC
      
  GoTo ClearMemory

ClearMemory:
  Set pSBar = Nothing
  Set pPro = Nothing
  Set pGeom = Nothing
  Set pSpRef = Nothing
  Set pClone = Nothing
  Set pSpRefRes = Nothing
  Set pWSF = Nothing
  Set pName = Nothing
  Set pFWS = Nothing
  Set pFields = Nothing
  Set pFieldsEdit = Nothing
  Set pField = Nothing
  Set pFieldEdit = Nothing
  Set pGeomDef = Nothing
  Set pGeomDefEdit = Nothing
  varVal = Null
  Set pSubArray = Nothing
  Set pTemplateField = Nothing
  Erase lngIDIndex
  Set pTempField = Nothing
  Set pCLSID = Nothing
  Set pInMemFC = Nothing
  Set pOutFeat = Nothing

End Function
Public Function MakeFieldNameVarArrayFromTable(pOrigTable As ITable, pNewFClass As IFeatureClass, _
      pFieldArray As esriSystem.IVariantArray, booFieldArrayContainsNamePairs As Boolean) As esriSystem.IVariantArray
      
  ' MAKE IVariantArray OF FIELD NAME INDICES
  ' STRUCTURE:  ARRAY OF 4-ITEM VARIANT ARRAYS
  '             ITEMS = 0) ORIGINAL FCLASS FIELD NAME
  '                     1) ORIGINAL FCLASS FIELD INDEX
  '                     2) NEW FCLASS FIELD NAME
  '                     3) NEW FCLASS FIELD INDEX
  
  Dim lngIndex As Long
  Dim pField As iField
  Dim strOrigName As String
  Dim strNewName As String
  Dim pNamePair As esriSystem.IStringArray
  Dim pFieldData As esriSystem.IVariantArray
  Set MakeFieldNameVarArrayFromTable = New varArray
  
  For lngIndex = 0 To pFieldArray.Count - 1
    If booFieldArrayContainsNamePairs Then
      Set pNamePair = pFieldArray.Element(lngIndex)
      strOrigName = pNamePair.Element(0)
      strNewName = pNamePair.Element(1)
    Else
      Set pField = pFieldArray.Element(lngIndex)
      strOrigName = pField.Name
      strNewName = pField.Name
    End If
    Set pFieldData = New esriSystem.varArray
    pFieldData.Add strOrigName
    pFieldData.Add pOrigTable.FindField(strOrigName)
    pFieldData.Add strNewName
    pFieldData.Add pNewFClass.FindField(strNewName)
    MakeFieldNameVarArrayFromTable.Add pFieldData
  Next lngIndex


  GoTo ClearMemory

ClearMemory:
  Set pField = Nothing
  Set pNamePair = Nothing
  Set pFieldData = Nothing

End Function
Public Sub SetFeatureSymbols2(pFeatureLayer As ILayer, lyrFileName As String)
    
' ADAPTED FROM Kirk Kuykendall
' http://forums.esri.com/Thread.asp?c=93&f=992&t=59083
  
  Dim booScaleSymbols As Boolean
  Dim booShouldDisplayAnnotation As Boolean
  
  Dim pGxFile As IGxFile
  Set pGxFile = New GxLayer
  pGxFile.Path = lyrFileName
  
  Dim pGxLayer As IGxLayer
  Set pGxLayer = pGxFile
  
  Dim pGFLayer As IGeoFeatureLayer
  Dim pFLayerDef As IFeatureLayerDefinition2
  Set pGFLayer = pGxLayer.Layer
  
  Dim strDefQuery As String
  
  If TypeOf pGFLayer Is IFeatureLayerDefinition2 Then
    Set pFLayerDef = pGFLayer
    strDefQuery = pFLayerDef.DefinitionExpression
  End If
  
  booScaleSymbols = pGFLayer.ScaleSymbols
  booShouldDisplayAnnotation = pGFLayer.DisplayAnnotation
  
  ' SET TRANSPARENCY
  Dim pLayerEffects As ILayerEffects
  Dim lngTransp As Double
  Set pLayerEffects = pGFLayer
  lngTransp = pLayerEffects.Transparency
  
  Dim pRenderer As IFeatureRenderer
  Set pRenderer = pGFLayer.Renderer
  
  Dim pAnnotation As IAnnotateLayerPropertiesCollection2
  Set pAnnotation = pGFLayer.AnnotationProperties

  
'    Dim pMxDoc As IMxDocument
'    Set pMxDoc = ThisDocument
  Dim pGFLayer2 As IGeoFeatureLayer
  Dim pFLayerDef2 As IFeatureLayerDefinition2
  
  Set pGFLayer2 = pFeatureLayer
  Set pGFLayer2.Renderer = pRenderer        ' <<------------------------

  Set pFLayerDef2 = pFeatureLayer
  pFLayerDef2.DefinitionExpression = strDefQuery

  Set pLayerEffects = pFeatureLayer
  pLayerEffects.Transparency = lngTransp
  pGFLayer2.ScaleSymbols = booScaleSymbols
  pGFLayer2.DisplayAnnotation = booShouldDisplayAnnotation
  pGFLayer2.AnnotationProperties = pAnnotation
  
'    pMxDoc.ActiveView.Refresh
'    pMxDoc.CurrentContentsView.Refresh pGFLayer


  GoTo ClearMemory

ClearMemory:
  Set pGxFile = Nothing
  Set pGxLayer = Nothing
  Set pGFLayer = Nothing
  Set pFLayerDef = Nothing
  Set pLayerEffects = Nothing
  Set pRenderer = Nothing
  Set pAnnotation = Nothing
  Set pGFLayer2 = Nothing
  Set pFLayerDef2 = Nothing

End Sub

Public Function ReturnQuerySpecialCharacters(pDataset As IDataset, Optional strPrefix As String, _
    Optional strSuffix As String, Optional strWildcardSingleMatch As String, _
    Optional strWildlcardManyMatch As String, Optional strSQLEscapePrefix As String, _
    Optional strSQLEscapeSuffix As String) As Boolean
  On Error GoTo ErrHandler
  

'  Dim pMxDoc As IMxDocument
'  Set pMxDoc = ThisDocument
'  Dim pFLayer As IFeatureLayer
'  Set pFLayer = pMxDoc.FocusMap.Layer(0)
'  Dim pFClass As IFeatureClass
'  Set pFClass = pFLayer.FeatureClass
'
'  Dim strPrefix As String
'  Dim strSuffix As String
'  Dim strWildSingle As String
'  Dim strWildMany As String
'  Dim strEscPrefix As String
'  Dim strEscSuffix As String
'
'  Dim booWorked As Boolean
'  booWorked = ReturnQuerySpecialCharacters(pFClass, strPrefix, strSuffix, strWildSingle, strWildMany, _
'      strEscPrefix, strEscSuffix)
'
'  Dim strPrefixReport As String
'  Dim strSuffixReport As String
'  Dim strWildSingleReport As String
'  Dim strWildManyReport As String
'  Dim strEscPrefixReport As String
'  Dim strEscSuffixReport As String
'
'  If strPrefix = "" Then
'    strPrefixReport = "<EMPTY>"
'  Else
'    strPrefixReport = Asc(strPrefix)
'  End If
'
'  If strSuffix = "" Then
'    strSuffixReport = "<EMPTY>"
'  Else
'    strSuffixReport = Asc(strSuffix)
'  End If
'
'  If strWildSingle = "" Then
'    strWildSingleReport = "<EMPTY>"
'  Else
'    strWildSingleReport = Asc(strWildSingle)
'  End If
'
'  If strWildMany = "" Then
'    strWildManyReport = "<EMPTY>"
'  Else
'    strWildManyReport = Asc(strWildMany)
'  End If
'
'  If strEscPrefix = "" Then
'    strEscPrefixReport = "<EMPTY>"
'  Else
'    strEscPrefixReport = Asc(strEscPrefix)
'  End If
'
'  If strEscSuffix = "" Then
'    strEscSuffixReport = "<EMPTY>"
'  Else
'    strEscSuffixReport = Asc(strEscSuffix)
'  End If
'
'  Debug.Print "strPrefix = " & strPrefix & "(ASCII = " & strPrefixReport & ")" & vbCrLf & _
'              "strSuffix = " & strSuffix & "(ASCII = " & strSuffixReport & ")" & vbCrLf & _
'              "strWildSingle = " & strWildSingle & "(ASCII = " & strWildSingleReport & ")" & vbCrLf & _
'              "strWildMany = " & strWildMany & "(ASCII = " & strWildManyReport & ")" & vbCrLf & _
'              "strEscPrefix = " & strEscPrefix & "(ASCII = " & strEscPrefixReport & ")" & vbCrLf & _
'              "strEscSuffix = " & strEscSuffix & "(ASCII = " & strEscSuffixReport & ")"
  
  ReturnQuerySpecialCharacters = False
  
  Dim pWS As IWorkspace
  If TypeOf pDataset Is IWorkspace Then
    Set pWS = pDataset
  Else
    Set pWS = pDataset.Workspace
  End If
  
  Dim pSQLSyntax As ISQLSyntax
  Set pSQLSyntax = pWS
  strPrefix = pSQLSyntax.GetSpecialCharacter(esriSQL_DelimitedIdentifierPrefix)
  strSuffix = pSQLSyntax.GetSpecialCharacter(esriSQL_DelimitedIdentifierSuffix)
  strWildcardSingleMatch = pSQLSyntax.GetSpecialCharacter(esriSQL_WildcardSingleMatch)
  strWildlcardManyMatch = pSQLSyntax.GetSpecialCharacter(esriSQL_WildcardManyMatch)
  strSQLEscapePrefix = pSQLSyntax.GetSpecialCharacter(esriSQL_EscapeKeyPrefix)
  strSQLEscapeSuffix = pSQLSyntax.GetSpecialCharacter(esriSQL_EscapeKeySuffix)

  ReturnQuerySpecialCharacters = True
  
  GoTo ClearMemory
  Exit Function
ErrHandler:
  
  ReturnQuerySpecialCharacters = False
  
ClearMemory:
  Set pWS = Nothing
  Set pSQLSyntax = Nothing

End Function

Public Function ReturnTitleCase(ByVal strString As String) As String
  
  Do Until InStr(1, strString, "  ") = 0
    strString = Replace(strString, "  ", " ")
  Loop
  
  Dim strSplit() As String
  strSplit = Split(strString, " ")
  Dim lngIndex As Long
  Dim strWord As String
  For lngIndex = 0 To UBound(strSplit)
    strWord = Trim(strSplit(lngIndex))
    If Len(strWord) = 1 Then
      ReturnTitleCase = ReturnTitleCase & UCase(strWord) & " "
    Else
      ReturnTitleCase = ReturnTitleCase & UCase(Left(strWord, 1)) & _
        LCase(Right(strWord, Len(strWord) - 1)) & " "
    End If
  Next lngIndex
  ReturnTitleCase = Left(ReturnTitleCase, Len(ReturnTitleCase) - 1)


  GoTo ClearMemory
ClearMemory:
  Erase strSplit
End Function

Public Function ReadTextFile(strFilename As String) As String

  If Dir(strFilename) <> "" Then
  
    Dim lngFileNumber As Long
    lngFileNumber = FreeFile(0)
'    MsgBox "Trying to open " & strFileName & " as " & CStr(lngFileNumber)
'
'    MsgBox "File Exists = " & CStr(PSN_Cell_Towers.aml_func_mod.ExistFileDir(strFileName))
    
    Dim strFileText As String
    
    Open strFilename For Binary As #lngFileNumber
    
    strFileText = Space$(LOF(lngFileNumber))
    Get #lngFileNumber, , strFileText
    
    Close #lngFileNumber
    
    ReadTextFile = strFileText
  Else
    ReadTextFile = ""
  End If

End Function

Public Function WriteTextFile(strFilename As String, strText As String, Optional booForceOverwrite As Boolean = False, _
    Optional booAppend As Boolean = False) As Boolean
  
  Dim lngFileNumber As Long
  
  If Dir(strFilename) = "" Or booForceOverwrite Then
    lngFileNumber = FreeFile(0)
    
    Open strFilename For Output As #lngFileNumber

    Print #lngFileNumber, strText
    Close #lngFileNumber
    WriteTextFile = True
    
  ElseIf booAppend Then
    lngFileNumber = FreeFile(0)
    
    Open strFilename For Append As #lngFileNumber

    Print #lngFileNumber, strText
    Close #lngFileNumber
    WriteTextFile = True
    
  Else
    ' CONFIRM WHETHER TO OVERWRITE FILE
    Dim lngVBResult As VbMsgBoxResult
    
    lngVBResult = MsgBox("File Already Exists!" & vbCrLf & vbCrLf & "  --> " & strFilename & vbCrLf & vbCrLf & _
        "Click 'OK' to overwrite the file, or 'CANCEL' to quit...", vbOKCancel, "File Exists:")
    If lngVBResult = vbOK Then
      Kill strFilename
      If Dir(strFilename) <> "" Then
        MsgBox "Unable to delete " & strFilename & vbCrLf & vbCrLf & _
          "It may be open in another application.  Please delete this file manually or save the text to a new filename.", , "Unable to Delete File:"
        WriteTextFile = False
      Else
      
        lngFileNumber = FreeFile(0)
        
        Open strFilename For Output As #lngFileNumber
    
        Print #lngFileNumber, strText
        Close #lngFileNumber
        WriteTextFile = True
      End If
    Else
      WriteTextFile = False
    End If
  End If

End Function



Public Function ReturnFilesFromNestedFolders(ByVal strDir As String, strExtensionWithDot As String) As esriSystem.IStringArray
  
  Set ReturnFilesFromNestedFolders = New esriSystem.strArray
    
  If Right(strDir, 1) <> "\" Then strDir = strDir & "\"
  
  Dim strOriginalDir As String
  strOriginalDir = strDir
  
  Dim booFoundSubFolders As Boolean
      
  Dim pPathArray As esriSystem.IStringArray
  Set pPathArray = New esriSystem.strArray
  
  Dim pFinalArray As esriSystem.IStringArray
  Set pFinalArray = New esriSystem.strArray
  Dim pCheckColl As Collection
  Set pCheckColl = New Collection
  
  pFinalArray.Add strDir
  pCheckColl.Add True, strDir
  
  Dim strDirName As String
    
  strDirName = Dir(strDir, vbDirectory)   ' Retrieve the first entry.
  Do While strDirName <> ""   ' Start the loop.
     ' Ignore the current directory and the encompassing directory.
     If strDirName <> "." And strDirName <> ".." Then
        ' Use bitwise comparison to make sure MyName is a directory.
        If IsFolder_FalseIfCrash((strDir & strDirName)) Then
'        If (GetAttr(strDir & strDirName) And vbDirectory) = vbDirectory Then
           pPathArray.Add strDir & strDirName & "\"
           pFinalArray.Add strDir & strDirName & "\"
           pCheckColl.Add False, strDir & strDirName & "\"
        End If   ' it represents a directory.
     End If
     strDirName = Dir   ' Get next entry.
  Loop
  
  booFoundSubFolders = pPathArray.Count > 0
  ' If Not booFoundSubFolders Then Exit Function
  
  Dim strSubFolder As String
  
  Dim booFoundSubHere As Boolean
  Dim pSubArray As esriSystem.IStringArray
  
  Dim lngIndex As Long
  
  Do While booFoundSubFolders
    booFoundSubFolders = False
    Set pSubArray = New esriSystem.strArray
    
    For lngIndex = 0 To pPathArray.Count - 1
      strSubFolder = pPathArray.Element(lngIndex)

'     If strDirName <> "." And strDirName <> ".." Then
'        lngCounter = lngCounter + 1

      booFoundSubHere = False
      strDirName = Dir(strSubFolder, vbDirectory)   ' Retrieve the first entry.
      Do While strDirName <> ""   ' Start the loop.
         ' Ignore the current directory and the encompassing directory.
         If strDirName <> "." And strDirName <> ".." Then
            ' Use bitwise comparison to make sure MyName is a directory.
            If IsFolder_FalseIfCrash(strSubFolder & strDirName) Then
'            If (GetAttr(strSubFolder & strDirName) And vbDirectory) = vbDirectory Then
               pSubArray.Add strSubFolder & strDirName & "\"
               booFoundSubFolders = True
               booFoundSubHere = True
               If Not CheckCollectionForKey(pCheckColl, strSubFolder & strDirName & "\") Then
                 pCheckColl.Add 1, strSubFolder & strDirName & "\"
                 pFinalArray.Add strSubFolder & strDirName & "\"
               End If
            End If   ' it represents a directory.
         End If
         strDirName = Dir   ' Get next entry.
      Loop
      
      If Not booFoundSubHere Then
        pSubArray.Add strSubFolder
      End If
      
    Next lngIndex
    
    If booFoundSubFolders Then
      Set pPathArray = pSubArray
    End If
    
  Loop
  
  Dim strFolders() As String
  ReDim strFolders(pFinalArray.Count - 1)
  
  
  For lngIndex = 0 To pFinalArray.Count - 1
    strDir = pFinalArray.Element(lngIndex)
    strFolders(lngIndex) = strDir
'    Debug.Print CStr(lngIndex + 1) & "]  " & strDir
  Next lngIndex
  
  QuickSort.StringsAscending strFolders, 0, pFinalArray.Count - 1
  
'  For lngIndex = 0 To UBound(strFolders)
'    strDir = strFolders(lngIndex)
'    Debug.Print CStr(lngIndex + 1) & "]  " & strDir
'  Next lngIndex

  Dim lngCounter As Long
  lngCounter = 0

  Debug.Print
  
  Dim pFolderFeatLayers As esriSystem.IVariantArray
  Dim pFilenames As esriSystem.IStringArray
  Set pFolderFeatLayers = New esriSystem.varArray
  Set pFilenames = New esriSystem.strArray
  Dim strDirAndFile As String
  
  For lngIndex = 0 To UBound(strFolders)
    strDir = strFolders(lngIndex)
    strDirName = Dir(strDir, vbNormal)   ' Retrieve the first entry.
    lngCounter = lngCounter + 1

    Do While strDirName <> ""   ' Start the loop.
       ' Ignore the current directory and the encompassing directory.
       If strDirName <> "." And strDirName <> ".." Then
          strDirAndFile = strDir & strDirName

          If IsNormal_FalseIfCrash(strDirAndFile) Then
'          If (GetAttr(strDirAndFile) And vbNormal) = vbNormal Then
            If StrComp(Right(strDirAndFile, Len(strExtensionWithDot)), strExtensionWithDot, vbTextCompare) = 0 Then

'             Debug.Print "Examining Folder #" & CStr(lngCounter) & ":  " & strDirName
              pFilenames.Add strDirAndFile
'             Debug.Print "  --> " & pDataset.BrowseName
            End If
          End If
       End If
       strDirName = Dir   ' Get next entry.
    Loop
  Next lngIndex

  ' CONFIRM THAT CORRECT DIRECTORY HAS BEEN SELECTED AND THAT IT ACTUALLY HAS POLYLINE SHAPEFILES
  
  Set ReturnFilesFromNestedFolders = pFilenames
  
  Set pPathArray = Nothing
  Set pFinalArray = Nothing
  Set pCheckColl = Nothing
  Set pSubArray = Nothing
  Set pFolderFeatLayers = Nothing
  Set pFilenames = Nothing


  GoTo ClearMemory

ClearMemory:
  Set pPathArray = Nothing
  Set pFinalArray = Nothing
  Set pCheckColl = Nothing
  Set pSubArray = Nothing
  Erase strFolders
  Set pFolderFeatLayers = Nothing
  Set pFilenames = Nothing

End Function

Public Function ReturnFilesFromNestedFolders2(ByVal strDir As String, strAnyTextInName As String) As esriSystem.IStringArray
  
  Set ReturnFilesFromNestedFolders2 = New esriSystem.strArray
    
  If Right(strDir, 1) <> "\" Then strDir = strDir & "\"
  
  Dim strOriginalDir As String
  strOriginalDir = strDir
  
  Dim booFoundSubFolders As Boolean
      
  Dim pPathArray As esriSystem.IStringArray
  Set pPathArray = New esriSystem.strArray
  
  Dim pFinalArray As esriSystem.IStringArray
  Set pFinalArray = New esriSystem.strArray
  Dim pCheckColl As Collection
  Set pCheckColl = New Collection
  
  pFinalArray.Add strDir
  pCheckColl.Add True, strDir
  
  Dim strDirName As String
    
  strDirName = Dir(strDir, vbDirectory)   ' Retrieve the first entry.
  Do While strDirName <> ""   ' Start the loop.
     ' Ignore the current directory and the encompassing directory.
     If strDirName <> "." And strDirName <> ".." Then
        ' Use bitwise comparison to make sure MyName is a directory.
        If IsFolder_FalseIfCrash((strDir & strDirName)) Then
'        If (GetAttr(strDir & strDirName) And vbDirectory) = vbDirectory Then
           pPathArray.Add strDir & strDirName & "\"
           pFinalArray.Add strDir & strDirName & "\"
           pCheckColl.Add False, strDir & strDirName & "\"
        End If   ' it represents a directory.
     End If
     strDirName = Dir   ' Get next entry.
  Loop
  
  booFoundSubFolders = pPathArray.Count > 0
  ' If Not booFoundSubFolders Then Exit Function
  
  Dim strSubFolder As String
  
  Dim booFoundSubHere As Boolean
  Dim pSubArray As esriSystem.IStringArray
  
  Dim lngIndex As Long
  
  Do While booFoundSubFolders
    booFoundSubFolders = False
    Set pSubArray = New esriSystem.strArray
    
    For lngIndex = 0 To pPathArray.Count - 1
      strSubFolder = pPathArray.Element(lngIndex)

'     If strDirName <> "." And strDirName <> ".." Then
'        lngCounter = lngCounter + 1

      booFoundSubHere = False
      strDirName = Dir(strSubFolder, vbDirectory)   ' Retrieve the first entry.
      Do While strDirName <> ""   ' Start the loop.
         ' Ignore the current directory and the encompassing directory.
         If strDirName <> "." And strDirName <> ".." Then
            ' Use bitwise comparison to make sure MyName is a directory.
            If IsFolder_FalseIfCrash(strSubFolder & strDirName) Then
'            If (GetAttr(strSubFolder & strDirName) And vbDirectory) = vbDirectory Then
               pSubArray.Add strSubFolder & strDirName & "\"
               booFoundSubFolders = True
               booFoundSubHere = True
               If Not CheckCollectionForKey(pCheckColl, strSubFolder & strDirName & "\") Then
                 pCheckColl.Add 1, strSubFolder & strDirName & "\"
                 pFinalArray.Add strSubFolder & strDirName & "\"
               End If
            End If   ' it represents a directory.
         End If
         strDirName = Dir   ' Get next entry.
      Loop
      
      If Not booFoundSubHere Then
        pSubArray.Add strSubFolder
      End If
      
    Next lngIndex
    
    If booFoundSubFolders Then
      Set pPathArray = pSubArray
    End If
    
  Loop
  
  Dim strFolders() As String
  ReDim strFolders(pFinalArray.Count - 1)
  
  
  For lngIndex = 0 To pFinalArray.Count - 1
    strDir = pFinalArray.Element(lngIndex)
    strFolders(lngIndex) = strDir
'    Debug.Print CStr(lngIndex + 1) & "]  " & strDir
  Next lngIndex
  
  QuickSort.StringsAscending strFolders, 0, pFinalArray.Count - 1
  
'  For lngIndex = 0 To UBound(strFolders)
'    strDir = strFolders(lngIndex)
'    Debug.Print CStr(lngIndex + 1) & "]  " & strDir
'  Next lngIndex

  Dim lngCounter As Long
  lngCounter = 0

'  Debug.Print
  
  Dim pFolderFeatLayers As esriSystem.IVariantArray
  Dim pFilenames As esriSystem.IStringArray
  Set pFolderFeatLayers = New esriSystem.varArray
  Set pFilenames = New esriSystem.strArray
  Dim strDirAndFile As String
  
  For lngIndex = 0 To UBound(strFolders)
    strDir = strFolders(lngIndex)
    strDirName = Dir(strDir, vbNormal)   ' Retrieve the first entry.
    lngCounter = lngCounter + 1

    Do While strDirName <> ""   ' Start the loop.
       ' Ignore the current directory and the encompassing directory.
       If strDirName <> "." And strDirName <> ".." Then
          strDirAndFile = strDir & strDirName

          If IsNormal_FalseIfCrash(strDirAndFile) Then
'          If (GetAttr(strDirAndFile) And vbNormal) = vbNormal Then
            'If StrComp(Right(strDirAndFile, Len(strExtensionWithDot)), strExtensionWithDot, vbTextCompare) = 0 Then
            If InStr(1, strDirName, strAnyTextInName, vbTextCompare) > 0 Then
'             Debug.Print "Examining Folder #" & CStr(lngCounter) & ":  " & strDirName
              pFilenames.Add strDirAndFile
'             Debug.Print "  --> " & pDataset.BrowseName
            End If
          End If
       End If
       strDirName = Dir   ' Get next entry.
    Loop
  Next lngIndex

  ' CONFIRM THAT CORRECT DIRECTORY HAS BEEN SELECTED AND THAT IT ACTUALLY HAS POLYLINE SHAPEFILES
  Set ReturnFilesFromNestedFolders2 = pFilenames
  
  GoTo ClearMemory

ClearMemory:
  Set pPathArray = Nothing
  Set pFinalArray = Nothing
  Set pCheckColl = Nothing
  Set pSubArray = Nothing
  Erase strFolders
  Set pFolderFeatLayers = Nothing
  Set pFilenames = Nothing

End Function



Public Function ReturnShapefileLayersFromNestedFolders(ByVal strDir As String, _
    lngType As esriGeometryType) As esriSystem.IVariantArray
  
  Set ReturnShapefileLayersFromNestedFolders = New esriSystem.varArray
  
  
  If Right(strDir, 1) <> "\" Then strDir = strDir & "\"
  
  Dim strOriginalDir As String
  strOriginalDir = strDir
  
  Dim booFoundSubFolders As Boolean
      
  Dim pPathArray As esriSystem.IStringArray
  Set pPathArray = New esriSystem.strArray
  
  Dim pFinalArray As esriSystem.IStringArray
  Set pFinalArray = New esriSystem.strArray
  Dim pCheckColl As Collection
  Set pCheckColl = New Collection
  
  pFinalArray.Add strDir
  pCheckColl.Add True, strDir
  
  Dim strDirName As String
    
  strDirName = Dir(strDir, vbDirectory)   ' Retrieve the first entry.
  Do While strDirName <> ""   ' Start the loop.
     ' Ignore the current directory and the encompassing directory.
     If strDirName <> "." And strDirName <> ".." Then
        ' Use bitwise comparison to make sure MyName is a directory.
        If IsFolder_FalseIfCrash(strDir & strDirName) Then
'        If (GetAttr(strDir & strDirName) And vbDirectory) = vbDirectory Then
           pPathArray.Add strDir & strDirName & "\"
           pFinalArray.Add strDir & strDirName & "\"
           pCheckColl.Add False, strDir & strDirName & "\"
        End If   ' it represents a directory.
     End If
     strDirName = Dir   ' Get next entry.
  Loop
  
  booFoundSubFolders = pPathArray.Count > 0
  ' If Not booFoundSubFolders Then Exit Function
  
  Dim strSubFolder As String
  
  Dim booFoundSubHere As Boolean
  Dim pSubArray As esriSystem.IStringArray
  
  Dim lngIndex As Long
  
  Do While booFoundSubFolders
    booFoundSubFolders = False
    Set pSubArray = New esriSystem.strArray
    
    For lngIndex = 0 To pPathArray.Count - 1
      strSubFolder = pPathArray.Element(lngIndex)

'     If strDirName <> "." And strDirName <> ".." Then
'        lngCounter = lngCounter + 1

      booFoundSubHere = False
      strDirName = Dir(strSubFolder, vbDirectory)   ' Retrieve the first entry.
      Do While strDirName <> ""   ' Start the loop.
         ' Ignore the current directory and the encompassing directory.
         If strDirName <> "." And strDirName <> ".." Then
            ' Use bitwise comparison to make sure MyName is a directory.
            If IsFolder_FalseIfCrash(strSubFolder & strDirName) Then
'            If (GetAttr(strSubFolder & strDirName) And vbDirectory) = vbDirectory Then
               pSubArray.Add strSubFolder & strDirName & "\"
               booFoundSubFolders = True
               booFoundSubHere = True
               If Not CheckCollectionForKey(pCheckColl, strSubFolder & strDirName & "\") Then
                 pCheckColl.Add 1, strSubFolder & strDirName & "\"
                 pFinalArray.Add strSubFolder & strDirName & "\"
               End If
            End If   ' it represents a directory.
         End If
         strDirName = Dir   ' Get next entry.
      Loop
      
      If Not booFoundSubHere Then
        pSubArray.Add strSubFolder
      End If
      
    Next lngIndex
    
    If booFoundSubFolders Then
      Set pPathArray = pSubArray
    End If
    
  Loop
  
  Dim strFolders() As String
  ReDim strFolders(pFinalArray.Count - 1)
  
  
  For lngIndex = 0 To pFinalArray.Count - 1
    strDir = pFinalArray.Element(lngIndex)
    strFolders(lngIndex) = strDir
'    Debug.Print CStr(lngIndex + 1) & "]  " & strDir
  Next lngIndex
  
  QuickSort.StringsAscending strFolders, 0, pFinalArray.Count - 1
  
'  For lngIndex = 0 To UBound(strFolders)
'    strDir = strFolders(lngIndex)
'    Debug.Print CStr(lngIndex + 1) & "]  " & strDir
'  Next lngIndex

  Dim lngCounter As Long
  lngCounter = 0

  Debug.Print

  Dim pShapefilePaths As esriSystem.IStringArray
  Set pShapefilePaths = New esriSystem.strArray

  Dim pWS As IFeatureWorkspace
  Dim pWSFact As IWorkspaceFactory2
  Set pWSFact = New ShapefileWorkspaceFactory
  Dim pFClass As IFeatureClass
  
  Dim pFolderFeatLayers As esriSystem.IVariantArray
  Dim pShapefileNames As esriSystem.IStringArray
  Set pFolderFeatLayers = New esriSystem.varArray
  Set pShapefileNames = New esriSystem.strArray
  Dim pFLayer As IFeatureLayer
  Dim pDataset As IDataset
  Dim strDirAndFile As String
  
  For lngIndex = 0 To UBound(strFolders)
    strDir = strFolders(lngIndex)
    strDirName = Dir(strDir, vbNormal)   ' Retrieve the first entry.
    lngCounter = lngCounter + 1

    Set pWS = pWSFact.OpenFromFile(strDir, 0) ' Me.hWnd)

    Do While strDirName <> ""   ' Start the loop.
       ' Ignore the current directory and the encompassing directory.
       If strDirName <> "." And strDirName <> ".." Then
          strDirAndFile = strDir & strDirName

          If IsNormal_FalseIfCrash(strDirAndFile) Then
'          If (GetAttr(strDirAndFile) And vbNormal) = vbNormal Then
            If StrComp(Right(strDirAndFile, 4), ".shp", vbTextCompare) = 0 Then

'              Debug.Print "Examining Folder #" & CStr(lngCounter) & ":  " & strDirName

              Set pFClass = pWS.OpenFeatureClass(strDirName)
              If pFClass.ShapeType = lngType Then
                Set pFLayer = New FeatureLayer
                Set pFLayer.FeatureClass = pFClass
                Set pDataset = pFClass
                pFLayer.Name = pDataset.BrowseName
                pFolderFeatLayers.Add pFLayer
                pShapefileNames.Add strDirAndFile
'                Debug.Print "  --> " & pDataset.BrowseName
              End If
            End If
          End If
       End If
       strDirName = Dir   ' Get next entry.
    Loop
  Next lngIndex

  ' CONFIRM THAT CORRECT DIRECTORY HAS BEEN SELECTED AND THAT IT ACTUALLY HAS POLYLINE SHAPEFILES
  Set ReturnShapefileLayersFromNestedFolders = pFolderFeatLayers
  
  GoTo ClearMemory
  
ClearMemory:
  Set pPathArray = Nothing
  Set pFinalArray = Nothing
  Set pCheckColl = Nothing
  Set pSubArray = Nothing
  Erase strFolders
  Set pShapefilePaths = Nothing
  Set pWS = Nothing
  Set pWSFact = Nothing
  Set pFClass = Nothing
  Set pFolderFeatLayers = Nothing
  Set pShapefileNames = Nothing
  Set pFLayer = Nothing
  Set pDataset = Nothing

End Function

Public Function ReturnFeatureLayersFromGeoDatabase(pFeatWS As IFeatureWorkspace, _
    lngType As esriGeometryType) As esriSystem.IVariantArray
  
  On Error GoTo ErrHandler
  
  Set ReturnFeatureLayersFromGeoDatabase = New esriSystem.varArray
  Dim pWS As IWorkspace
  Set pWS = pFeatWS
  
  Dim pDatasets As IEnumDataset
  Set pDatasets = pWS.Datasets(esriDTFeatureClass)
  pDatasets.Reset
  
  Dim pDataset As IDataset
  Dim pFClass As IFeatureClass
  Set pDataset = pDatasets.Next
  Dim pFLayer As IFeatureLayer
  
  Do Until pDataset Is Nothing
    Set pFClass = pDataset
    If pFClass.ShapeType = lngType Then
      Set pFLayer = New FeatureLayer
      Set pFLayer.FeatureClass = pFClass
      pFLayer.Name = pDataset.BrowseName
      ReturnFeatureLayersFromGeoDatabase.Add pFLayer
    End If
    Set pDataset = pDatasets.Next
  Loop
  
  GoTo ClearMemory

  Exit Function

ErrHandler:
  Set ReturnFeatureLayersFromGeoDatabase = New esriSystem.varArray
  

ClearMemory:
  Set pWS = Nothing
  Set pDatasets = Nothing
  Set pDataset = Nothing
  Set pFClass = Nothing
  Set pFLayer = Nothing

End Function

Public Function ReturnRasterDatasetOrNothing(pWS As IRasterWorkspace, strFilename As String) As IRasterDataset
  On Error GoTo ErrHandler
  
  Set ReturnRasterDatasetOrNothing = Nothing
  Set ReturnRasterDatasetOrNothing = pWS.OpenRasterDataset(strFilename)

  Exit Function
ErrHandler:
  Set ReturnRasterDatasetOrNothing = Nothing

End Function
Public Function ReturnFeatureClassOrNothing(pWS As IFeatureWorkspace, strFilename As String) As IFeatureClass
  On Error GoTo ErrHandler
  
  Set ReturnFeatureClassOrNothing = Nothing
  Set ReturnFeatureClassOrNothing = pWS.OpenFeatureClass(strFilename)

  Exit Function
ErrHandler:
  Set ReturnFeatureClassOrNothing = Nothing

End Function
Public Function ReturnTableOrNothing(pWS As IFeatureWorkspace, strFilename As String) As ITable
  On Error GoTo ErrHandler
  
  Set ReturnTableOrNothing = Nothing
  Set ReturnTableOrNothing = pWS.OpenTable(strFilename)

  Exit Function
ErrHandler:
  Set ReturnTableOrNothing = Nothing

End Function

Public Function ConvertNumberToBullet(lngNumber As Long, lngBulletType As JenBulletTypes) As String

  Dim strReturn As String
  If lngNumber < 1 Then
    ConvertNumberToBullet = "Invalid Number..."
  Else
    Select Case lngBulletType
      Case ENUM_Letter_Lowercase, ENUM_Letter_Uppercase
      
        Dim lngMod As Long
        Dim lngRemainder As Long
        
        lngRemainder = lngNumber
        strReturn = ""
        Do While lngRemainder > 0
          lngMod = (lngRemainder - 1) Mod 26
          strReturn = Chr(65 + lngMod) & strReturn
          lngRemainder = Int((lngRemainder - lngMod) / 26)
        Loop
        If lngBulletType = ENUM_Letter_Lowercase Then strReturn = LCase(strReturn)

      
      Case ENUM_RomanNumeral_Lowercase
        strReturn = ConvertNumberToRoman(lngNumber, False)
      
      Case ENUM_RomanNumeral_Uppercase
        strReturn = ConvertNumberToRoman(lngNumber, True)
    End Select
    
    ConvertNumberToBullet = strReturn
  End If

End Function
Function ConvertNumberBase(InputNum As Long, BaseNum As Long) As Long()
  
  ' MODIFIED FROM http://support.microsoft.com/kb/135635
  
  Dim quotient As Long
  Dim remainder As Long
  Dim strCounters As String
  
  Dim lngCounters() As Long
  If BaseNum <= 1 Then
    ReDim lngCounters(0)
    lngCounters(0) = -999
  Else
    
    quotient = InputNum   ' Set quotient to number to convert.
    remainder = InputNum  ' Set remainder to number to convert.
    strCounters = ""
  
    Do While quotient <> 0   ' Loop while quotient is not zero.
  
       ' Store the remainder of the quotient divided by base number in a
       ' variable called remainder.
       remainder = quotient Mod BaseNum
  
       ' Reset quotient variable to the integer value of the quotient
       ' divided by base number.
       quotient = Int(quotient / BaseNum)
  
       ' Reset strCounters to contain remainder and the previous strCounters.
       strCounters = remainder & "," & strCounters
  
    Loop
    If strCounters = "" Then strCounters = "0"
    If Right(strCounters, 1) = "," Then strCounters = Left(strCounters, Len(strCounters) - 1)
    
'    Debug.Print strCounters
    
    Dim strArray() As String
    strArray = Split(strCounters, ",")
    
    ReDim lngCounters(UBound(strArray))
    Dim lngIndex As Long
    For lngIndex = 0 To UBound(strArray)
      lngCounters(lngIndex) = CLng(strArray(lngIndex))
'      Debug.Print "  --> Array Index = " & CStr(lngIndex) & ":  Value = " & CStr(CLng(strArray(lngIndex)))
    Next lngIndex
  End If
  
  ConvertNumberBase = lngCounters   ' Convert strCounters variable to a number.


  GoTo ClearMemory
ClearMemory:
  Erase lngCounters
  Erase strArray

End Function

Public Function ConvertNumberToRoman(ByVal lngNumber As Long, Optional booUppercase As Boolean = False) As String

  ' Formats a number as a roman numeral.
  ' Author: Christian d'Heureuse (www.source-code.biz)
  ' http://www.source-code.biz/snippets/vbasic/7.htm
  ' MODIFIED BY JEFF JENNESS
  
   If lngNumber = 0 Then ConvertNumberToRoman = "0": Exit Function
      ' There is no roman symbol for 0, but we don't want to return an empty string.
   
   ' roman symbols
   Dim r As String
   Dim strThousands As String
   
   If booUppercase Then
     r = "IVXLCDM"
     strThousands = "M"
   Else
     r = "ivxlcdm"
     strThousands = "m"
   End If
   
   Dim i As Long: i = Abs(lngNumber)
   Dim s As String, p As Integer
   For p = 1 To 5 Step 2
      Dim d As Integer: d = i Mod 10: i = i \ 10
      Select Case d                 ' format a decimal digit
         Case 0 To 3: s = String(d, Mid(r, p, 1)) & s
         Case 4:      s = Mid(r, p, 2) & s
         Case 5 To 8: s = Mid(r, p + 1, 1) & String(d - 5, Mid(r, p, 1)) & s
         Case 9:      s = Mid(r, p, 1) & Mid(r, p + 2, 1) & s
         End Select
      Next
   s = String(i, strThousands) & s           ' format thousands
   If lngNumber < 0 Then s = "-" & s        ' insert sign if negative (non-standard)
   ConvertNumberToRoman = s

End Function


Public Sub Move_Element(pElement As IElement, pFromPoint As IPoint, pToPoint As IPoint)
  
  ' EXAMPLE: TO MOVE A GRAPHIC ELEMENT SO IT IS CENTERED IN THE PAGE LAYOUT, DO THE FOLLOWING:
  ' Move_Element pGroupElement, Get_Element_Or_Envelope_Point(pElement, ENUM_Center_Center), _
      Get_Element_Or_Envelope_Point(pLayoutExtent, ENUM_By_Percentages, 0.5, 0.5)
  
    
    Dim pTrans2D As ITransform2D
    Set pTrans2D = pElement

    pTrans2D.Move (pToPoint.x - pFromPoint.x), _
                  (pToPoint.Y - pFromPoint.Y)
    
  GoTo ClearMemory

ClearMemory:
  Set pTrans2D = Nothing

End Sub ' Move_Legend

Public Function Get_Element_Or_Envelope_Point(pElementOrEnvelope As IUnknown, lngAnchorPoint As Jen_ElementEnvPoint, _
  Optional dblXPercent As Double = 0.5, Optional dblYPercent As Double = 0.5, Optional pActiveView As IActiveView) As IPoint

  ' Initialize the output for this procedure...
  Dim theOutput As IPoint
  Set theOutput = Nothing
  Dim pEnv As IEnvelope
  
  If TypeOf pElementOrEnvelope Is IElement Then
    Dim pTemp As IElement
    Set pTemp = pElementOrEnvelope
    If pActiveView Is Nothing Then
      Set pEnv = pTemp.Geometry.Envelope
    Else
      Set pEnv = New Envelope
      pTemp.QueryBounds pActiveView.ScreenDisplay, pEnv
    End If
  Else
    Set pEnv = pElementOrEnvelope
  End If

  Select Case lngAnchorPoint
    Case ENUM_Upper_Left  ' UPPER LEFT CORNER
        Set theOutput = pEnv.UpperLeft
    Case ENUM_Upper_Right  ' UPPER RIGHT CORNER
        Set theOutput = pEnv.UpperRight
    Case ENUM_Lower_Left ' LOWER LEFT CORNER
        Set theOutput = pEnv.LowerLeft
    Case ENUM_Lower_Right  ' LOWER RIGHT CORNER
        Set theOutput = pEnv.LowerRight
    Case ENUM_Upper_Center  ' VERTICAL TOP, HORIZONTAL CENTER
        Set theOutput = New Point
        theOutput.PutCoords ((pEnv.XMax - pEnv.XMin) / 2) + pEnv.XMin, pEnv.YMax
        Set theOutput.SpatialReference = pEnv.SpatialReference
    Case ENUM_Lower_Center  ' VERTICAL BOTTOM, HORIZONTAL CENTER
        Set theOutput = New Point
        theOutput.PutCoords ((pEnv.XMax - pEnv.XMin) / 2) + pEnv.XMin, pEnv.YMin
        Set theOutput.SpatialReference = pEnv.SpatialReference
    Case ENUM_Center_Left  ' VERTICAL CENTER, HORIZONTAL LEFT
        Set theOutput = New Point
        theOutput.PutCoords pEnv.XMin, ((pEnv.YMax - pEnv.YMin) / 2) + pEnv.YMin
        Set theOutput.SpatialReference = pEnv.SpatialReference
    Case ENUM_Center_Right  ' VERTICAL CENTER, HORIZONTAL RIGHT
        Set theOutput = New Point
        theOutput.PutCoords pEnv.XMax, ((pEnv.YMax - pEnv.YMin) / 2) + pEnv.YMin
        Set theOutput.SpatialReference = pEnv.SpatialReference
    Case ENUM_Center_Center  ' VERTICAL CENTER, HORIZONTAL CENTER
        Set theOutput = New Point
        theOutput.PutCoords ((pEnv.XMax - pEnv.XMin) / 2) + pEnv.XMin, ((pEnv.YMax - pEnv.YMin) / 2) + pEnv.YMin
        Set theOutput.SpatialReference = pEnv.SpatialReference
    Case ENUM_By_Percentages
        Set theOutput = New Point
        theOutput.PutCoords ((pEnv.XMax - pEnv.XMin) * dblXPercent) + pEnv.XMin, ((pEnv.YMax - pEnv.YMin) * dblYPercent) + pEnv.YMin
        Set theOutput.SpatialReference = pEnv.SpatialReference
    Case Else
        MsgBox "position not supported: " & CLng(lngAnchorPoint)
  End Select

  ' Return the output for this procedure...
  Set Get_Element_Or_Envelope_Point = theOutput
  
  GoTo ClearMemory

ClearMemory:
  Set theOutput = Nothing
  Set pEnv = Nothing
  Set pTemp = Nothing

End Function ' Get_Element_Or_Envelope_Point

Public Function ReturnGraphicsByNameFromLayout(pMxDoc As IMxDocument, strName As String, _
      Optional AsElements As Boolean) As IArray

  Dim pGraphicsContainer As IGraphicsContainer
  
  Set pGraphicsContainer = pMxDoc.PageLayout
  
  pGraphicsContainer.Reset
  
  Dim pElement As IElement
  Dim pElementProperties As IElementProperties
  
  Set pElement = pGraphicsContainer.Next
  
  Dim pArray As IArray
  Set pArray = New esriSystem.Array
  Dim pGeometry As IGeometry
  Dim pClone As IClone
  
  While Not pElement Is Nothing
    Set pElementProperties = pElement
    If StrComp(pElementProperties.Name, strName, vbTextCompare) = 0 Then
      If AsElements Then
        pArray.Add pElement
      Else
        Set pGeometry = pElement.Geometry
        Set pClone = pGeometry
        pArray.Add pClone.Clone     ' ONLY RETURN A COPY OF THE GEOMETRY; DON'T WANT TO MODIFY ACTUAL GRAPHIC HERE
      End If
    End If
    Set pElement = pGraphicsContainer.Next
    
  Wend
  Set ReturnGraphicsByNameFromLayout = pArray


  GoTo ClearMemory

ClearMemory:
  Set pGraphicsContainer = Nothing
  Set pElement = Nothing
  Set pElementProperties = Nothing
  Set pArray = Nothing
  Set pGeometry = Nothing
  Set pClone = Nothing

End Function
Public Sub Move_Geometry(pGeometry As IGeometry, pFromPoint As IPoint, pToPoint As IPoint)
  
  ' EXAMPLE: TO MOVE A GRAPHIC ELEMENT SO ITS TOP LEFT CORNER TOUCHES ANOTHER POINT, DO THE FOLLOWING:
  ' Move_Geometry pPolygon, Get_Element_Or_Envelope_Point(pElement, ENUM_Upper_Left), _
      pDestinationPoint
  
    
    Dim pTrans2D As ITransform2D
    Set pTrans2D = pGeometry

    pTrans2D.Move (pToPoint.x - pFromPoint.x), _
                  (pToPoint.Y - pFromPoint.Y)
    
  GoTo ClearMemory

ClearMemory:
  Set pTrans2D = Nothing

End Sub ' Move_Legend

Public Function CreateSpatialReferenceWGS84() As ISpatialReference
  
  Dim pWGS84 As IGeographicCoordinateSystem
  Dim pSpatRefFact As ISpatialReferenceFactory
  Set pSpatRefFact = New SpatialReferenceEnvironment
  Set pWGS84 = pSpatRefFact.CreateGeographicCoordinateSystem(esriSRGeoCS_WGS1984)
  Dim pSpRefRes As ISpatialReferenceResolution
  Set pSpRefRes = pWGS84
  pSpRefRes.ConstructFromHorizon
  
  Set CreateSpatialReferenceWGS84 = pWGS84
  
  Set pWGS84 = Nothing
  Set pSpatRefFact = Nothing
  Set pSpRefRes = Nothing


  GoTo ClearMemory
ClearMemory:
  Set pWGS84 = Nothing
  Set pSpatRefFact = Nothing
  Set pSpRefRes = Nothing

End Function

Public Function CreateSpatialReferenceNAD83() As ISpatialReference
  
  Dim pNAD83 As IGeographicCoordinateSystem
  Dim pSpatRefFact As ISpatialReferenceFactory
  Set pSpatRefFact = New SpatialReferenceEnvironment
  Set pNAD83 = pSpatRefFact.CreateGeographicCoordinateSystem(esriSRGeoCS_NAD1983)
  Dim pSpRefRes As ISpatialReferenceResolution
  Set pSpRefRes = pNAD83
  pSpRefRes.ConstructFromHorizon
  
  Set CreateSpatialReferenceNAD83 = pNAD83
  
  Set pNAD83 = Nothing
  Set pSpatRefFact = Nothing
  Set pSpRefRes = Nothing


  GoTo ClearMemory
ClearMemory:
  Set pNAD83 = Nothing
  Set pSpatRefFact = Nothing
  Set pSpRefRes = Nothing

End Function
Public Function CreateSpatialReferenceNAD27() As ISpatialReference
  
  Dim pNAD27 As IGeographicCoordinateSystem
  Dim pSpatRefFact As ISpatialReferenceFactory
  Set pSpatRefFact = New SpatialReferenceEnvironment
  Set pNAD27 = pSpatRefFact.CreateGeographicCoordinateSystem(esriSRGeoCS_NAD1927)
  Dim pSpRefRes As ISpatialReferenceResolution
  Set pSpRefRes = pNAD27
  pSpRefRes.ConstructFromHorizon
  
  Set CreateSpatialReferenceNAD27 = pNAD27
  
  Set pNAD27 = Nothing
  Set pSpatRefFact = Nothing
  Set pSpRefRes = Nothing


  GoTo ClearMemory

ClearMemory:
  Set pNAD27 = Nothing
  Set pSpatRefFact = Nothing
  Set pSpRefRes = Nothing

End Function

Public Function CreateNAD27_WGS84_GeoTransformationFlagstaff() As IGeoTransformation

  Dim pSpRefFact As ISpatialReferenceFactory2
  Set pSpRefFact = New SpatialReferenceEnvironment
  
  Dim pTransformType As esriSRGeoTransformationType
  pTransformType = esriSRGeoTransformation_NAD1927_To_WGS1984_4
  
  Dim pGeoTransform As IGeoTransformation
  Set pGeoTransform = pSpRefFact.CreateGeoTransformation(pTransformType)
  Set CreateNAD27_WGS84_GeoTransformationFlagstaff = pGeoTransform
  
  
'  Dim pFromSpRef As ISpatialReference
'  Dim pToSpRef As ISpatialReference
'  pGeoTransform.GetSpatialReferences pFromSpRef, pToSpRef
'  Debug.Print "CreateNAD27_WGS84_GeoTransformationFlagstaff" & vbCrLf & _
'              "  --> From " & CStr(pFromSpRef.Name) & vbCrLf & _
'              "  --> To " & CStr(pToSpRef.Name)
  Set pSpRefFact = Nothing
  Set pGeoTransform = Nothing


  GoTo ClearMemory
ClearMemory:
  Set pSpRefFact = Nothing
  Set pGeoTransform = Nothing

End Function

Public Function CreateNAD83_WGS84_GeoTransformationFlagstaff() As IGeoTransformation

  Dim pSpRefFact As ISpatialReferenceFactory2
  Set pSpRefFact = New SpatialReferenceEnvironment
  
  Dim pTransformType As esriSRGeoTransformationType
  pTransformType = esriSRGeoTransformation_NAD1983_To_WGS1984_5
  
  Dim pGeoTransform As IGeoTransformation
  Set pGeoTransform = pSpRefFact.CreateGeoTransformation(pTransformType)
  Set CreateNAD83_WGS84_GeoTransformationFlagstaff = pGeoTransform
  
  
'  Dim pFromSpRef As ISpatialReference
'  Dim pToSpRef As ISpatialReference
'  pGeoTransform.GetSpatialReferences pFromSpRef, pToSpRef
'  Debug.Print "CreateNAD83_WGS84_GeoTransformationFlagstaff" & vbCrLf & _
'              "  --> From " & CStr(pFromSpRef.Name) & vbCrLf & _
'              "  --> To " & CStr(pToSpRef.Name)
  
  Set pSpRefFact = Nothing
  Set pGeoTransform = Nothing


  GoTo ClearMemory

ClearMemory:
  Set pSpRefFact = Nothing
  Set pGeoTransform = Nothing

End Function

Public Function CreateNAD27_NAD83_GeoTransformationFlagstaff() As IGeoTransformation

  Dim pSpRefFact As ISpatialReferenceFactory2
  Set pSpRefFact = New SpatialReferenceEnvironment
  
  Dim pTransformType As esriSRGeoTransformation2Type
  pTransformType = esriSRGeoTransformation_NAD_1927_TO_NAD_1983_NADCON
  
  Dim pGeoTransform As IGeoTransformation
  Set pGeoTransform = pSpRefFact.CreateGeoTransformation(pTransformType)
  Set CreateNAD27_NAD83_GeoTransformationFlagstaff = pGeoTransform
  
  
'  Dim pFromSpRef As ISpatialReference
'  Dim pToSpRef As ISpatialReference
'  pGeoTransform.GetSpatialReferences pFromSpRef, pToSpRef
'  Debug.Print "CreateNAD27_NAD83_GeoTransformationFlagstaff" & vbCrLf & _
'              "  --> From " & CStr(pFromSpRef.Name) & vbCrLf & _
'              "  --> To " & CStr(pToSpRef.Name)

  GoTo ClearMemory
  
ClearMemory:
  Set pSpRefFact = Nothing
  Set pGeoTransform = Nothing

End Function

Function MakeESRIColor(pCurrColor As IRgbColor, hParentWind As OLE_HANDLE, booCanceled As Boolean, _
  Optional lngTop As Long = 100, Optional lngLeft As Long = 500, _
  Optional lngRight As Long = 100, Optional lngBottom As Long = 500, _
  Optional lngAlignOption As Jen_AlignColorDialogOption = ENUM_AlignBeneathRectangle) As IColor

  ' ADAPTED FROM:
  ' http://help.arcgis.com/en/sdk/10.0/vba_desktop/componenthelp/index.html#/IColorPalette_Interface/001t0000005v000000/
  ' http://forums.esri.com/thread.asp?c=93&f=1154&t=58833
  
  If lngAlignOption = ENUM_AlignBeneathRectangle Then
    lngBottom = lngBottom + 1
    lngLeft = lngLeft + 2
  Else
    lngRight = lngRight + 3
  End If
  
  Dim pRect As tagRECT
  pRect.Top = lngTop
  pRect.Left = lngLeft
  pRect.Right = lngRight
  pRect.bottom = lngBottom
  
'  Dim pColorSet As ISet
'  Set pColorSet = New esriSystem.Set
'  Dim pRGB As IColor
'  Set pRGB = New RgbColor
'  pRGB.RGB = RGB(pCurrColor.Red, 0, 0)
'  pColorSet.Add pRGB
'  Set pRGB = New RgbColor
'  pRGB.RGB = RGB(0, pCurrColor.Green, 0)
'  pColorSet.Add pRGB
'  Set pRGB = New RgbColor
'  pRGB.RGB = RGB(0, 0, pCurrColor.Blue)
'  pColorSet.Add pRGB
'
'  Dim pCustomColorPalette As ICustomColorPalette
'  Set pCustomColorPalette = New ColorPalette
'  Set pCustomColorPalette.ColorSet = pColorSet

  Dim pMyPalette As IColorPalette
  Set pMyPalette = New ColorPalette ' pCustomColorPalette
  '
  ' Now show the Palette dialog.
  '
  If Not pMyPalette.TrackPopupMenu(pRect, pCurrColor, lngAlignOption = ENUM_AlignToRightOfRectangle, hParentWind) Then
     booCanceled = True
  Else
    '
    ' We can retrieve the selected color from the Color property.
    '
    Set MakeESRIColor = pMyPalette.Color
    booCanceled = False
  End If

  Set pMyPalette = Nothing

  GoTo ClearMemory

ClearMemory:
  Set pMyPalette = Nothing

End Function


Public Function ReturnShapeTypeNameFromObject(pUnknown As IUnknown) As String()
  
  Dim strReturn() As String
  ReDim strReturn(1)
  
  Dim varGeomType As esriGeometryType
  Dim pFClass As IFeatureClass
  Dim pGeom As IGeometry
  
  If TypeOf pUnknown Is IFeatureClass Then
    Set pFClass = pUnknown
    varGeomType = pFClass.ShapeType
    strReturn = ReturnShapeTypeFromGeomType(varGeomType)
  ElseIf TypeOf pUnknown Is IGeometry Then
    Set pGeom = pUnknown
    varGeomType = pGeom.GeometryType
    strReturn = ReturnShapeTypeFromGeomType(varGeomType)
  Else
    strReturn(0) = "Unknown Object Type"
    strReturn(1) = "Unknown Object Type"
  End If
  
  ReturnShapeTypeNameFromObject = strReturn


  GoTo ClearMemory

ClearMemory:
  Erase strReturn
  Set pFClass = Nothing
  Set pGeom = Nothing

End Function

Public Function ReturnShapeTypeFromGeomType(varShapeType As esriGeometryType) As String()

  Dim strReturn() As String
  ReDim strReturn(1)
  
  Select Case varShapeType
    Case esriGeometryNull
      strReturn(0) = "Null"
      strReturn(1) = "Null"
    Case esriGeometryPoint
      strReturn(0) = "Point"
      strReturn(1) = "Points"
    Case esriGeometryMultipoint
      strReturn(0) = "Multipoint"
      strReturn(1) = "Multipoints"
    Case esriGeometryLine
      strReturn(0) = "Line"
      strReturn(1) = "Lines"
    Case esriGeometryCircularArc
      strReturn(0) = "CircularArc"
      strReturn(1) = "CircularArcs"
    Case esriGeometryEllipticArc
      strReturn(0) = "EllipticArc"
      strReturn(1) = "EllipticArcs"
    Case esriGeometryBezier3Curve
      strReturn(0) = "Bezier3Curve"
      strReturn(1) = "Bezier3Curves"
    Case esriGeometryPath
      strReturn(0) = "Path"
      strReturn(1) = "Paths"
    Case esriGeometryPolyline
      strReturn(0) = "Polyline"
      strReturn(1) = "Polylines"
    Case esriGeometryRing
      strReturn(0) = "Ring"
      strReturn(1) = "Rings"
    Case esriGeometryPolygon
      strReturn(0) = "Polygon"
      strReturn(1) = "Polygons"
    Case esriGeometryEnvelope
      strReturn(0) = "Envelope"
      strReturn(1) = "Envelopes"
    Case esriGeometryAny
      strReturn(0) = "Any"
      strReturn(1) = "Any"
    Case esriGeometryBag
      strReturn(0) = "GeometryBag"
      strReturn(1) = "GeometryBagx"
    Case esriGeometryMultiPatch
      strReturn(0) = "MultiPatch"
      strReturn(1) = "MultiPatches"
    Case esriGeometryTriangleStrip
      strReturn(0) = "TriangleStrip"
      strReturn(1) = "TriangleStrips"
    Case esriGeometryTriangleFan
      strReturn(0) = "TriangleFan"
      strReturn(1) = "TriangleFans"
    Case esriGeometryRay
      strReturn(0) = "Ray"
      strReturn(1) = "Rays"
    Case esriGeometrySphere
      strReturn(0) = "Sphere"
      strReturn(1) = "Spheres"
    Case esriGeometryTriangles
      strReturn(0) = "Triangle"
      strReturn(1) = "Triangles"
    Case Else
      strReturn(0) = ""
      strReturn(1) = ""
  End Select
  
  ReturnShapeTypeFromGeomType = strReturn

  GoTo ClearMemory

ClearMemory:
  Erase strReturn

End Function

Public Function ReturnFoldersFromNestedFolders(ByVal strDir As String, strPartialText As String) As esriSystem.IStringArray
  
  Set ReturnFoldersFromNestedFolders = New esriSystem.strArray
    
  If Right(strDir, 1) <> "\" Then strDir = strDir & "\"
  
  Dim strOriginalDir As String
  strOriginalDir = strDir
  
  Dim booFoundSubFolders As Boolean
      
  Dim pPathArray As esriSystem.IStringArray
  Set pPathArray = New esriSystem.strArray
  
  Dim pFinalArray As esriSystem.IStringArray
  Set pFinalArray = New esriSystem.strArray
  Dim pCheckColl As Collection
  Set pCheckColl = New Collection
  
  pFinalArray.Add strDir
  pCheckColl.Add True, strDir
  
  Dim strDirName As String
    
  strDirName = Dir(strDir, vbDirectory)   ' Retrieve the first entry.
  Do While strDirName <> ""   ' Start the loop.
     ' Ignore the current directory and the encompassing directory.
     If strDirName <> "." And strDirName <> ".." Then
        ' Use bitwise comparison to make sure MyName is a directory.
        If IsFolder_FalseIfCrash(strDir & strDirName) Then
'        If (GetAttr(strDir & strDirName) And vbDirectory) = vbDirectory Then
           pPathArray.Add strDir & strDirName & "\"
           pFinalArray.Add strDir & strDirName & "\"
           pCheckColl.Add False, strDir & strDirName & "\"
        End If   ' it represents a directory.
     End If
     strDirName = Dir   ' Get next entry.
  Loop
  
  booFoundSubFolders = pPathArray.Count > 0
  ' If Not booFoundSubFolders Then Exit Function
  
  Dim strSubFolder As String
  
  Dim booFoundSubHere As Boolean
  Dim pSubArray As esriSystem.IStringArray
  
  Dim lngIndex As Long
  
  Do While booFoundSubFolders
    booFoundSubFolders = False
    Set pSubArray = New esriSystem.strArray
    
    For lngIndex = 0 To pPathArray.Count - 1
      strSubFolder = pPathArray.Element(lngIndex)
      
      
'     If strDirName <> "." And strDirName <> ".." Then
'        lngCounter = lngCounter + 1

      booFoundSubHere = False
      strDirName = Dir(strSubFolder, vbDirectory)   ' Retrieve the first entry.
      Do While strDirName <> ""   ' Start the loop.
         ' Ignore the current directory and the encompassing directory.
         If strDirName <> "." And strDirName <> ".." Then
            ' Use bitwise comparison to make sure MyName is a directory.
            If IsFolder_FalseIfCrash(strSubFolder & strDirName) Then
'            If (GetAttr(strSubFolder & strDirName) And vbDirectory) = vbDirectory Then
               pSubArray.Add strSubFolder & strDirName & "\"
               booFoundSubFolders = True
               booFoundSubHere = True
               If Not CheckCollectionForKey(pCheckColl, strSubFolder & strDirName & "\") Then
                 pCheckColl.Add 1, strSubFolder & strDirName & "\"
                 pFinalArray.Add strSubFolder & strDirName & "\"
               End If
            End If   ' it represents a directory.
         End If
         strDirName = Dir   ' Get next entry.
      Loop
      
      If Not booFoundSubHere Then
        pSubArray.Add strSubFolder
      End If
      
    Next lngIndex
    
    If booFoundSubFolders Then
      Set pPathArray = pSubArray
    End If
    
  Loop
  
  Dim strFolders() As String
  ReDim strFolders(pFinalArray.Count - 1)
  
  
  For lngIndex = 0 To pFinalArray.Count - 1
    strDir = pFinalArray.Element(lngIndex)
    strFolders(lngIndex) = strDir
'    Debug.Print CStr(lngIndex + 1) & "]  " & strDir
  Next lngIndex
  
  QuickSort.StringsAscending strFolders, 0, pFinalArray.Count - 1
  
'  For lngIndex = 0 To UBound(strFolders)
'    strDir = strFolders(lngIndex)
'    Debug.Print CStr(lngIndex + 1) & "]  " & strDir
'  Next lngIndex

  Dim lngCounter As Long
  lngCounter = 0

  Debug.Print
  
  Dim pFolderFeatLayers As esriSystem.IVariantArray
  Dim pFilenames As esriSystem.IStringArray
  Set pFolderFeatLayers = New esriSystem.varArray
  Set pFilenames = New esriSystem.strArray
  
  For lngIndex = 0 To UBound(strFolders)
    strDir = strFolders(lngIndex)
    
    If InStr(1, strDir, strPartialText, vbTextCompare) > 0 Then
      pFilenames.Add strDir
    End If
  Next lngIndex

  ' CONFIRM THAT CORRECT DIRECTORY HAS BEEN SELECTED AND THAT IT ACTUALLY HAS POLYLINE SHAPEFILES
  Set ReturnFoldersFromNestedFolders = pFilenames
  
  GoTo ClearMemory

ClearMemory:
  Set pPathArray = Nothing
  Set pFinalArray = Nothing
  Set pCheckColl = Nothing
  Set pSubArray = Nothing
  Erase strFolders
  Set pFolderFeatLayers = Nothing
  Set pFilenames = Nothing

End Function
Public Function ReturnDatasetTypeName(lngType As Long) As String

  Select Case lngType
    Case 1
      ReturnDatasetTypeName = "Any Dataset"
    Case 2
      ReturnDatasetTypeName = "Any Container Dataset"
    Case 3
      ReturnDatasetTypeName = "Any Geo Dataset"
    Case 4
      ReturnDatasetTypeName = "Feature Dataset"
    Case 5
      ReturnDatasetTypeName = "Feature Class"
    Case 6
      ReturnDatasetTypeName = "Planar Graph"
    Case 7
      ReturnDatasetTypeName = "Geometric Network"
    Case 8
      ReturnDatasetTypeName = "Topology"
    Case 9
      ReturnDatasetTypeName = "Text Dataset"
    Case 10
      ReturnDatasetTypeName = "Table Dataset"
    Case 11
      ReturnDatasetTypeName = "Relationship Class"
    Case 12
      ReturnDatasetTypeName = "Raster Dataset"
    Case 13
      ReturnDatasetTypeName = "Raster Band"
    Case 14
      ReturnDatasetTypeName = "Tin Dataset"
    Case 15
      ReturnDatasetTypeName = "CadDrawing Dataset"
    Case 16
      ReturnDatasetTypeName = "Raster Catalog"
    Case 17
      ReturnDatasetTypeName = "Toolbox"
    Case 18
      ReturnDatasetTypeName = "Tool"
    Case 19
      ReturnDatasetTypeName = "Network Dataset"
    Case 20
      ReturnDatasetTypeName = "Terrain dataset"
    Case 21
      ReturnDatasetTypeName = "Feature Class Representation"
    Case 22
      ReturnDatasetTypeName = "Cadastral Fabric"
    Case 23
      ReturnDatasetTypeName = "Schematic Dataset"
    Case 24
      ReturnDatasetTypeName = "Address Locator"
    Case 26
      ReturnDatasetTypeName = "Map"
    Case 27
      ReturnDatasetTypeName = "Layer"
    Case 28
      ReturnDatasetTypeName = "Style"
    Case 29
      ReturnDatasetTypeName = "Mosaic Dataset"
    Case Else
      ReturnDatasetTypeName = "Unknown Type [" & CStr(lngType) & "]"
  End Select

End Function

Public Function ReturnSelectedLayers(pMxDoc As IMxDocument) As esriSystem.IArray

  Set ReturnSelectedLayers = New esriSystem.Array
  
  Dim pUnknown As IUnknown
  Set pUnknown = pMxDoc.SelectedItem
  Dim pSelSet As ISet
  Dim pObj As IUnknown
  
  If Not pUnknown Is Nothing Then
    If TypeOf pUnknown Is ILayer Then
      ReturnSelectedLayers.Add pUnknown
    ElseIf TypeOf pUnknown Is ISet Then
      Set pSelSet = pUnknown
      pSelSet.Reset
      Set pObj = pSelSet.Next
      Do Until pObj Is Nothing
        If TypeOf pObj Is ILayer Then
          ReturnSelectedLayers.Add pObj
        End If
        Set pObj = pSelSet.Next
      Loop
    End If
  End If
    
  GoTo ClearMemory
  
ClearMemory:
  Set pUnknown = Nothing
  Set pSelSet = Nothing
  Set pObj = Nothing

End Function

Public Function SearchForTextInFolder_and_MakeReport(strFolder As String, strSearchString As String, _
    Optional strExtensionWithoutDot As String = "docx") As String
  
  Dim pFolders As IStringArray
  Set pFolders = ReturnFoldersFromNestedFolders(strFolder, "")
  
  Dim pFiles As esriSystem.IStringArray
  Dim pFinalFiles As esriSystem.IStringArray
  Set pFinalFiles = New esriSystem.strArray
  
  Dim lngIndex As Long
  Dim lngIndex2 As Long
  Dim strFile As String
  Dim strDocText As String
  
  For lngIndex = 0 To pFolders.Count - 1
    strFolder = pFolders.Element(lngIndex)
    Debug.Print CStr(lngIndex + 1) & "] " & strFolder
      
    Set pFiles = ReturnFilesFromNestedFolders(strFolder, strExtensionWithoutDot)
    For lngIndex2 = 0 To pFiles.Count - 1
      strFile = pFiles.Element(lngIndex2)

      Debug.Print "  --> " & strFile
      If InStr(1, strExtensionWithoutDot, "doc", vbTextCompare) > 0 Then
'        strDocText = ReturnWordText(strFile)
      Else
        strDocText = MyGeneralOperations.ReadTextFile(strFile)
      End If
      If InStr(1, strDocText, strSearchString, vbTextCompare) > 0 Then
        pFinalFiles.Add strFile
      End If

    Next lngIndex2
  Next lngIndex
  
  SearchForTextInFolder_and_MakeReport = ""
  
  Debug.Print "................................."
  Debug.Print "Found '" & strSearchString & "' in the following documents:"
  If pFinalFiles.Count = 0 Then
    SearchForTextInFolder_and_MakeReport = "Did not find '" & strSearchString & _
        "' in any '" & strExtensionWithoutDot & "' fies in '" & strFolder & "'..."
    Debug.Print SearchForTextInFolder_and_MakeReport
  Else
    SearchForTextInFolder_and_MakeReport = "Found '" & strSearchString & "' in the following documents:" & vbCrLf
    For lngIndex = 0 To pFinalFiles.Count - 1
      strFile = pFinalFiles.Element(lngIndex)
      SearchForTextInFolder_and_MakeReport = SearchForTextInFolder_and_MakeReport & _
          CStr(lngIndex + 1) & "] " & strFile & vbCrLf
      Debug.Print CStr(lngIndex + 1) & "] " & strFile
    Next lngIndex
  End If
  Debug.Print "Done..."
  GoTo ClearMemory
  
ClearMemory:
  Set pFolders = Nothing
  Set pFiles = Nothing
  Set pFinalFiles = Nothing

End Function


Public Function ReturnDecimalPrecision(dblNumber As Double) As Long

  Dim strNumber As String
  strNumber = TrimZerosAndDecimals(dblNumber, False)
  If InStr(1, strNumber, ".") = 0 Then
    ReturnDecimalPrecision = 0
  Else
    ReturnDecimalPrecision = Len(strNumber) - InStrRev(strNumber, ".")
  End If

End Function

Public Function FixAtDecimalLevel(dblNumber As Double, lngNumDecimals As Long)

  ' TRUNCATES A NUMBER AT A GIVEN DECIMAL LEVEL
  FixAtDecimalLevel = CDbl(Fix(dblNumber * (10 ^ lngNumDecimals))) / (10 ^ lngNumDecimals)

End Function

Public Function TextSlice(ByVal strOrigText As String, ByVal lngStartIndex As Long, ByVal lngEndIndex As Long) As String
  
  ' RETURNS TEXT BETWEEN lngStartIndex AND lngEndIndex INCLUSIVE
  Dim lngLength As Long
  lngLength = Len(strOrigText)
  If lngEndIndex >= lngLength Then lngEndIndex = lngLength
  If lngStartIndex < 1 Then lngStartIndex = 1
  If lngStartIndex > lngLength Or lngStartIndex > lngEndIndex Then
    TextSlice = ""
  Else
    TextSlice = Mid(strOrigText, lngStartIndex, lngEndIndex + 1 - lngStartIndex)
  End If

End Function

Public Function TextSlice2(ByVal strOrigText As String, ByVal lngStartIndex As Long, Optional ByVal lngEndIndex As Long = -999) As String
  
  ' RETURNS TEXT BETWEEN lngStartIndex AND lngEndIndex INCLUSIVE
  If lngEndIndex = -999 Then lngEndIndex = lngStartIndex
  
  Dim lngLength As Long
  lngLength = Len(strOrigText)
  If lngEndIndex >= lngLength Then lngEndIndex = lngLength
  If lngStartIndex < 1 Then lngStartIndex = 1
  If lngStartIndex > lngLength Or lngStartIndex > lngEndIndex Then
    TextSlice2 = ""
  Else
    TextSlice2 = Mid(strOrigText, lngStartIndex, lngEndIndex + 1 - lngStartIndex)
  End If

End Function

Public Function ConvertEsriDoubleArrayToVB(pDoubleArray As esriSystem.IDoubleArray) As Double()

  Dim dblReturn() As Double
  If pDoubleArray.Count > 0 Then
    ReDim dblReturn(pDoubleArray.Count - 1)
    Dim lngIndex As Long
    For lngIndex = 0 To pDoubleArray.Count - 1
      dblReturn(lngIndex) = pDoubleArray.Element(lngIndex)
    Next lngIndex
  End If
  ConvertEsriDoubleArrayToVB = dblReturn
  
  GoTo ClearMemory
ClearMemory:
  Erase dblReturn

End Function

Public Function SpacesInFrontOfText(strText As String, lngTotalLength As Long) As String
  
  Dim lngCurrentLength As Long
  lngCurrentLength = Len(strText)
  
  If lngCurrentLength >= lngTotalLength Then
    SpacesInFrontOfText = strText
  Else
    SpacesInFrontOfText = Space(lngTotalLength - lngCurrentLength) & strText
  End If

End Function

Public Function CreateInMemoryFeatureClass3(pGeometryArray As esriSystem.IArray, _
    Optional pValueArray As esriSystem.IVariantArray, Optional pTemplateFields As esriSystem.IVariantArray, _
    Optional pApp As IApplication, Optional strStatusMessage As String = "", _
    Optional lngFlushCount As Long = 500) As IFeatureClass
    

    ' create an inmemory featureclass
    ' ADAPTED FROM KIRK KUYKENDALL
    ' http://forums.esri.com/Thread.asp?c=93&f=993&t=210767
    
    Dim pSBar As IStatusBar
    Dim pPro As IStepProgressor
    Dim dateRunningTime As Date
    Dim strHeader As String
    Dim lngTotalCount As Long
    
    If Not pApp Is Nothing Then
      ' FOR PROGRESS METER
      Set pSBar = pApp.StatusBar
      Set pPro = pSBar.ProgressBar
      dateRunningTime = Now
      strHeader = strStatusMessage
      lngTotalCount = pGeometryArray.Count
      pSBar.ShowProgressBar strStatusMessage, 0, lngTotalCount, 1, True
      pPro.position = 0
    End If
        
    ' ASSUMES ONE OR MORE GEOMETRIES, ALL OF WHICH ARE IN THE SAME PROJECTION
    Dim pGeom As IGeometry
    Set pGeom = pGeometryArray.Element(0)
    
    Dim pSpRef As ISpatialReference
    Set pSpRef = pGeom.SpatialReference
    Dim pClone As IClone
    
'    If Not pDomain Is Nothing Then
'      Set pClone = pDomain
'      Set pNewDomain = pClone.Clone
'      pNewDomain.Expand 1.01, 1.01, True
'      pSpRef.SetDomain pNewDomain.XMin, pNewDomain.XMax, pNewDomain.YMin, pNewDomain.YMax
'    End If
    
    Dim pSpRefRes As ISpatialReferenceResolution
    Set pSpRefRes = pSpRef
    pSpRefRes.ConstructFromHorizon
    Dim pWSF As IWorkspaceFactory
    Set pWSF = New InMemoryWorkspaceFactory
    
    Dim pName As IName
    Set pName = pWSF.Create("", "inmemory", Nothing, 0)
    Dim pFWS As IFeatureWorkspace
    Set pFWS = pName.Open
            
    Dim pFields As IFields
    Dim pFieldsEdit As IFieldsEdit
    Dim pField As iField
    Dim pFieldEdit As IFieldEdit
   
    Set pFields = New Fields
    Set pFieldsEdit = pFields
      
    '' create the geometry field
    
    Dim pGeomDef As IGeometryDef
    Dim pGeomDefEdit As IGeometryDefEdit
    Set pGeomDef = New GeometryDef
    Set pGeomDefEdit = pGeomDef
    
'    MsgBox pGeom.SpatialReference.Name
    
    '' assign the geometry definiton properties.
    With pGeomDefEdit
      .GeometryType = pGeom.GeometryType
      .GridCount = 1
      .GridSize(0) = 0
      If pGeom.GeometryType = esriGeometryPoint Then
        .AvgNumPoints = 1
      Else
        .AvgNumPoints = 5
      End If
      .HasM = False
      .HasZ = False
      Set .SpatialReference = pGeom.SpatialReference
    End With
    
    Set pField = New Field
    Set pFieldEdit = pField
    
    pFieldEdit.Name = "Shape"
    pFieldEdit.AliasName = "geometry"
    pFieldEdit.Type = esriFieldTypeGeometry
    Set pFieldEdit.GeometryDef = pGeomDef
    pFieldsEdit.AddField pField
    

    Dim booAddAttribute As Boolean
    booAddAttribute = Not pValueArray Is Nothing And Not pTemplateFields Is Nothing
    Dim varVal As Variant
    
    Dim lngIndex As Long
    Dim pSubArray As esriSystem.IVariantArray
    Dim pTemplateField As iField
    
    Dim pTempField As iField
    If booAddAttribute Then
      For lngIndex = 0 To pTemplateFields.Count - 1
        Set pTemplateField = pTemplateFields.Element(lngIndex)
        Set pClone = pTemplateField
        Set pTempField = pClone.Clone
        pFieldsEdit.AddField pTempField
      Next lngIndex
      ReDim lngIDIndex(pTemplateFields.Count - 1)
    Else
      Set pField = New Field
      Set pFieldEdit = pField
      With pFieldEdit
        .Name = "Unique_ID"
        .Type = esriFieldTypeInteger
      End With
      pFieldsEdit.AddField pField
      ReDim lngIDIndex(0)
    End If
    
    Dim pCLSID As UID
    Set pCLSID = New UID
    pCLSID.Value = "esriGeoDatabase.Feature"
  
    Dim pInMemFC As IFeatureClass
    Set pInMemFC = pFWS.CreateFeatureClass("In_Memory", pFields, _
                             pCLSID, Nothing, esriFTSimple, _
                             "Shape", "")
    
    
    Dim lngAttIndex As Long
    If booAddAttribute Then
      For lngIndex = 0 To pTemplateFields.Count - 1
        Set pTemplateField = pTemplateFields.Element(lngIndex)
        lngIDIndex(lngIndex) = pInMemFC.FindField(pTemplateField.Name)
      Next lngIndex
    Else
      lngIDIndex(0) = pInMemFC.FindField("Unique_ID")
    End If
    
    Dim pOutFeatBuffer As IFeatureBuffer
    Set pOutFeatBuffer = pInMemFC.CreateFeatureBuffer
    Dim pOutFCursor As IFeatureCursor
    Set pOutFCursor = pInMemFC.Insert(True)
    Dim lngIndex2 As Long
    
    For lngIndex = 0 To pGeometryArray.Count - 1
      Set pClone = pGeometryArray.Element(lngIndex)
      Set pGeom = pClone.Clone
      
      Set pOutFeatBuffer.Shape = pGeom
      
      If booAddAttribute Then
        Set pSubArray = pValueArray.Element(lngIndex)
        For lngIndex2 = 0 To pSubArray.Count - 1
          varVal = pSubArray.Element(lngIndex2)
          pOutFeatBuffer.Value(lngIDIndex(lngIndex2)) = varVal
          
        Next lngIndex2
      Else
        pOutFeatBuffer.Value(lngIDIndex(0)) = lngIndex + 1
      End If
      pOutFCursor.InsertFeature pOutFeatBuffer
      
      If lngIndex Mod lngFlushCount = 0 Then pOutFCursor.Flush
      If Not pApp Is Nothing Then
        pPro.Step
      End If
    Next lngIndex
    pOutFCursor.Flush
    
    If Not pApp Is Nothing Then
      pPro.position = 1
      pSBar.HideProgressBar
    End If
    Set CreateInMemoryFeatureClass3 = pInMemFC
  
  GoTo ClearMemory
ClearMemory:
  Set pSBar = Nothing
  Set pPro = Nothing
  Set pGeom = Nothing
  Set pSpRef = Nothing
  Set pClone = Nothing
  Set pSpRefRes = Nothing
  Set pWSF = Nothing
  Set pName = Nothing
  Set pFWS = Nothing
  Set pFields = Nothing
  Set pFieldsEdit = Nothing
  Set pField = Nothing
  Set pFieldEdit = Nothing
  Set pGeomDef = Nothing
  Set pGeomDefEdit = Nothing
  varVal = Null
  Set pSubArray = Nothing
  Set pTemplateField = Nothing
  Erase lngIDIndex
  Set pTempField = Nothing
  Set pCLSID = Nothing
  Set pInMemFC = Nothing
  Set pOutFeatBuffer = Nothing
  Set pOutFCursor = Nothing


End Function

Public Function ClipNumberOfCharacters(strOrigText As String, lngClipFromRight As Long, lngClipFromLeft As Long) As String
  
  Dim strReturn As String
  If Len(strOrigText) <= lngClipFromRight + lngClipFromLeft Then
    ClipNumberOfCharacters = ""
  Else
    strReturn = Left(strOrigText, Len(strOrigText) - lngClipFromRight)
    strReturn = Right(strReturn, Len(strReturn) - lngClipFromLeft)
    ClipNumberOfCharacters = strReturn
  End If

End Function

Public Function IsFolder_FalseIfCrash(strPath As String) As Boolean
  On Error GoTo ErrHandle
  
  ' GetAttr appears to crash on locked files.
  IsFolder_FalseIfCrash = (GetAttr(strPath) And vbDirectory) = vbDirectory

  Exit Function
ErrHandle:
  IsFolder_FalseIfCrash = False
  
End Function
Public Function IsNormal_FalseIfCrash(strPath As String) As Boolean
  On Error GoTo ErrHandle
  
  ' GetAttr appears to crash on locked files.
  IsNormal_FalseIfCrash = (GetAttr(strPath) And vbNormal) = vbNormal
  
  Exit Function
ErrHandle:
  IsNormal_FalseIfCrash = False
  
End Function

Public Sub DeleteGraphicsByGeometry(ByRef pMxDoc As IMxDocument, pGeom As IGeometry, _
    Optional booDeleteFromLayout As Boolean = False)

  Dim pGraphicsContainer As IGraphicsContainer
  Dim pActiveView As IActiveView

  Dim pRelOp As IRelationalOperator
  Set pRelOp = pGeom

  If booDeleteFromLayout Then
    Set pGraphicsContainer = pMxDoc.PageLayout
  Else
    Set pGraphicsContainer = pMxDoc.FocusMap
  End If
  Set pActiveView = pMxDoc.ActiveView
  Dim pElement As IElement
'  Dim pElementProperties As IElementProperties
  Dim pEnvelope As IEnvelope
  Dim pTempEnvelope As IEnvelope
  Set pTempEnvelope = New Envelope

  pGraphicsContainer.Reset

  Set pElement = pGraphicsContainer.Next

  Dim pDeleteArray As esriSystem.IVariantArray
  Set pDeleteArray = New esriSystem.varArray

  While Not pElement Is Nothing
'    Set pElementProperties = pElement
    If Not pRelOp.Disjoint(pElement.Geometry) Then
      pDeleteArray.Add pElement
      If (pEnvelope Is Nothing) Then
        Set pEnvelope = New Envelope
        pElement.QueryBounds pMxDoc.ActiveView.ScreenDisplay, pEnvelope
      Else
        pElement.QueryBounds pMxDoc.ActiveView.ScreenDisplay, pTempEnvelope
        pEnvelope.Union pTempEnvelope
      End If
    End If
    Set pElement = pGraphicsContainer.Next

  Wend

  Dim lngIndex As Long
  If pDeleteArray.Count > 0 Then
    For lngIndex = 0 To pDeleteArray.Count - 1
      Set pElement = pDeleteArray.Element(lngIndex)
      pGraphicsContainer.DeleteElement pElement
    Next lngIndex
  End If

  If (Not pEnvelope Is Nothing) Then
    pActiveView.PartialRefresh esriViewGraphics + esriViewGraphicSelection, Nothing, pEnvelope
    DoEvents
  End If


  Set pDeleteArray = Nothing
  Set pEnvelope = Nothing
  Set pElement = Nothing
  Set pGraphicsContainer = Nothing
  Set pActiveView = Nothing
'  Set pElementProperties = Nothing
  Set pTempEnvelope = Nothing

End Sub

Public Function DateToYearDecimal(datDate As Date) As Double
  
  Dim lngDayInYear As Long
  Dim lngTotalDays As Long
  Dim booIsLeapYear As Boolean
  
  Dim lngYear As Long
  lngYear = CLng(Year(datDate))
  Dim lngDay As Long
  lngDay = CLng(Day(datDate))
  Dim lngMonth As Long
  lngMonth = CLng(Month(datDate))

  booIsLeapYear = lngYear Mod 4 = 0 And (lngYear Mod 400 = 0 Or lngYear Mod 100 <> 0)
  
  If booIsLeapYear Then
    lngTotalDays = 366
  Else
    lngTotalDays = 365
  End If
  
'  Debug.Print "Year = " & CStr(lngYear)
'  Debug.Print "Day = " & CStr(lngDay)
'  Debug.Print "Month = " & CStr(lngMonth)
'  Debug.Print "Leap Year = " & CStr(booIsLeapYear)
'  Debug.Print "Days in Year = " & CStr(lngTotalDays)
  
  Dim lngDayCounts(11) As Long
  lngDayCounts(0) = 31 ' January
  If booIsLeapYear Then
    lngDayCounts(1) = 29 ' February
  Else
    lngDayCounts(1) = 28 ' February
  End If
  lngDayCounts(2) = 31 ' March
  lngDayCounts(3) = 30 ' April
  lngDayCounts(4) = 31 ' May
  lngDayCounts(5) = 30 ' June
  lngDayCounts(6) = 31 ' July
  lngDayCounts(7) = 31 ' August
  lngDayCounts(8) = 30 ' September
  lngDayCounts(9) = 31 ' October
  lngDayCounts(10) = 30 ' November
  lngDayCounts(11) = 31 ' December
     
  Dim lngIndex As Long
  If lngMonth >= 2 Then
    For lngIndex = 1 To lngMonth - 1
      lngDay = lngDay + lngDayCounts(lngIndex - 1)
    Next lngIndex
  End If
  
'  Debug.Print "Day in Year = " & CStr(lngDay)
  DateToYearDecimal = CDbl(lngYear) + (CDbl(lngDay) / CDbl(lngTotalDays))
  
'  Debug.Print "Decimal Year = " & format(DateToYearDecimal, "0.000")

ClearMemory:
  Erase lngDayCounts

End Function



Public Function DateComponentsFromDate(datDate As String, Optional strLongMonth As String, Optional strShortMonth As String, _
    Optional booCapitalizeMonth As Boolean, Optional lngMonth As Long, Optional lngDay As Long, Optional lngYear As Long, Optional lngHour As Long, _
    Optional lngMinute As Long, Optional lngSecond As Long) As Boolean

  DateComponentsFromDate = False
  lngMonth = Month(datDate)
  lngYear = Year(datDate)
  lngDay = Day(datDate)
  lngHour = Hour(datDate)
  lngMinute = Minute(datDate)
  lngSecond = Second(datDate)
  
  Select Case lngMonth
    Case 1
      strShortMonth = "Jan"
      strLongMonth = "January"
    Case 2
      strShortMonth = "Feb"
      strLongMonth = "February"
    Case 3
      strShortMonth = "Mar"
      strLongMonth = "March"
    Case 4
      strShortMonth = "Apr"
      strLongMonth = "April"
    Case 5
      strShortMonth = "May"
      strLongMonth = "May"
    Case 6
      strShortMonth = "Jun"
      strLongMonth = "June"
    Case 7
      strShortMonth = "Jul"
      strLongMonth = "July"
    Case 8
      strShortMonth = "Aug"
      strLongMonth = "August"
    Case 9
      strShortMonth = "Sep"
      strLongMonth = "September"
    Case 10
      strShortMonth = "Oct"
      strLongMonth = "October"
    Case 11
      strShortMonth = "Nov"
      strLongMonth = "November"
    Case 12
      strShortMonth = "Dec"
      strLongMonth = "December"
  End Select
  
  If booCapitalizeMonth Then
    strShortMonth = UCase(strShortMonth)
    strLongMonth = UCase(strLongMonth)
  End If

  DateComponentsFromDate = True

End Function

Public Function ReturnEsriFieldTypeNameFromNumber(lngFieldType As esriFieldType) As String

  Select Case lngFieldType
    Case esriFieldTypeSmallInteger
      ReturnEsriFieldTypeNameFromNumber = "esriFieldTypeSmallInteger"
    Case esriFieldTypeInteger
      ReturnEsriFieldTypeNameFromNumber = "esriFieldTypeInteger"
    Case esriFieldTypeSingle
      ReturnEsriFieldTypeNameFromNumber = "esriFieldTypeSingle"
    Case esriFieldTypeDouble
      ReturnEsriFieldTypeNameFromNumber = "esriFieldTypeDouble"
    Case esriFieldTypeString
      ReturnEsriFieldTypeNameFromNumber = "esriFieldTypeString"
    Case esriFieldTypeDate
      ReturnEsriFieldTypeNameFromNumber = "esriFieldTypeDate"
    Case esriFieldTypeOID
      ReturnEsriFieldTypeNameFromNumber = "esriFieldTypeOID"
    Case esriFieldTypeGeometry
      ReturnEsriFieldTypeNameFromNumber = "esriFieldTypeGeometry"
    Case esriFieldTypeBlob
      ReturnEsriFieldTypeNameFromNumber = "esriFieldTypeBlob"
    Case esriFieldTypeRaster
      ReturnEsriFieldTypeNameFromNumber = "esriFieldTypeRaster"
    Case esriFieldTypeGUID
      ReturnEsriFieldTypeNameFromNumber = "esriFieldTypeGUID"
    Case esriFieldTypeGlobalID
      ReturnEsriFieldTypeNameFromNumber = "esriFieldTypeGlobalID"
    Case esriFieldTypeXML
      ReturnEsriFieldTypeNameFromNumber = "esriFieldTypeXML"
  End Select

End Function


Public Function Date_MonthNameFromNumber(lngMonth As Long, Optional strLongMonth As String, _
    Optional strShortMonth As String, Optional booCapitalizeMonth As Boolean) As Boolean
  
  Select Case lngMonth
    Case 1
      strShortMonth = "Jan"
      strLongMonth = "January"
    Case 2
      strShortMonth = "Feb"
      strLongMonth = "February"
    Case 3
      strShortMonth = "Mar"
      strLongMonth = "March"
    Case 4
      strShortMonth = "Apr"
      strLongMonth = "April"
    Case 5
      strShortMonth = "May"
      strLongMonth = "May"
    Case 6
      strShortMonth = "Jun"
      strLongMonth = "June"
    Case 7
      strShortMonth = "Jul"
      strLongMonth = "July"
    Case 8
      strShortMonth = "Aug"
      strLongMonth = "August"
    Case 9
      strShortMonth = "Sep"
      strLongMonth = "September"
    Case 10
      strShortMonth = "Oct"
      strLongMonth = "October"
    Case 11
      strShortMonth = "Nov"
      strLongMonth = "November"
    Case 12
      strShortMonth = "Dec"
      strLongMonth = "December"
  End Select
  
  If booCapitalizeMonth Then
    strShortMonth = UCase(strShortMonth)
    strLongMonth = UCase(strLongMonth)
  End If

  Date_MonthNameFromNumber = True

End Function
Public Function Date_ReturnMonthNumberFromName(strName As String, Optional strFullMonthName As String) As Long
  
  Dim strUCaseName As String
  strUCaseName = UCase(strName)
  
  Select Case strUCaseName
    Case "JAN", "JANUARY", "JAN."
      Date_ReturnMonthNumberFromName = 1
      strFullMonthName = "January"
    Case "FEB", "FEBRUARY", "FEB."
      Date_ReturnMonthNumberFromName = 2
      strFullMonthName = "February"
    Case "MAR", "MARCH", "MAR."
      Date_ReturnMonthNumberFromName = 3
      strFullMonthName = "March"
    Case "APR", "APRIL", "APR."
      Date_ReturnMonthNumberFromName = 4
      strFullMonthName = "April"
    Case "MAY"
      Date_ReturnMonthNumberFromName = 5
      strFullMonthName = "May"
    Case "JUN", "JUNE", "JUN."
      Date_ReturnMonthNumberFromName = 6
      strFullMonthName = "June"
    Case "JUL", "JULY", "JUL."
      Date_ReturnMonthNumberFromName = 7
      strFullMonthName = "July"
    Case "AUG", "AUGUST", "AUG."
      Date_ReturnMonthNumberFromName = 8
      strFullMonthName = "August"
    Case "SEP", "SEPTEMBER", "SEP."
      Date_ReturnMonthNumberFromName = 9
      strFullMonthName = "September"
    Case "OCT", "OCTOBER", "OCT."
      Date_ReturnMonthNumberFromName = 10
      strFullMonthName = "October"
    Case "NOV", "NOVEMBER", "NOV."
      Date_ReturnMonthNumberFromName = 11
      strFullMonthName = "November"
    Case "DEC", "DECEMBER", "DEC."
      Date_ReturnMonthNumberFromName = 12
      strFullMonthName = "December"
    Case Else
      MsgBox "Didn't find month '" & strName & "'!"
  End Select

End Function


Public Function WriteCodeToDuplicateFields(pTable As ITable, booDescriptive As Boolean, _
    booMakeFields As Boolean, booDimAndFindOriginalFields As Boolean, _
    strFClassTableVariableName As String, booDimAndFindNewFields As String, _
    strNewFClassTableVariableName As String) As String
   
   
     'Public Sub TestWriteCodeDuplicateFields()
  '
  '  Dim pMxDoc As IMxDocument
  '  Set pMxDoc = ThisDocument
  '
  '  Dim pStTable As IStandaloneTable
  '  Dim pTable As ITable
  '
  '  Set pStTable = MyGeneralOperations.ReturnTableByName("Players", pMxDoc.FocusMap)
  '  Set pTable = pStTable.Table
  '
  '  Dim strReport As String
  '  strReport = WriteCodeToDuplicateFields(pTable, False, True, True, "pOrigTable", True, "pPlayersTable")
  '
  '  Dim pDataObj As New MSForms.DataObject
  '  pDataObj.SetText strReport
  '  pDataObj.PutInClipboard
  '  Set pDataObj = Nothing
  '
  '  Set pMxDoc = Nothing
  '  Set pTable = Nothing
  '  Set pStTable = Nothing
  '
  'End Sub


  Dim strReport As String
  Dim strOrigDimReport As String
  Dim strOrigFindReport As String
  Dim strNewDimReport As String
  Dim strNewFindReport As String
  Dim strDescriptiveReport As String
  Dim pField As iField
  Dim lngIndex As Long
  
  strReport = "  dim pField as ifield" & vbCrLf & _
              "  dim pFieldEdit as ifieldedit" & vbCrLf & _
              "  dim pFieldArray as esrisystem.ivariantarray " & vbCrLf & _
              "  set pfieldarray = new esrisystem.vararray" & vbCrLf & vbCrLf
    
  For lngIndex = 0 To pTable.Fields.FieldCount - 1
    Set pField = pTable.Fields.Field(lngIndex)
    If pField.Type <> esriFieldTypeOID Then
      strDescriptiveReport = strDescriptiveReport & "  " & CStr(lngIndex + 1) & ") " & pField.Name & ": " & vbCrLf
      
      strReport = strReport & "  set pfield = new field" & vbCrLf & _
                              "  set pfieldedit = pfield" & vbCrLf & _
                              "  with pfieldedit" & vbCrLf & _
                              "    .name = """ & pField.Name & """" & vbCrLf & _
                              "    .aliasname = """ & pField.AliasName & """" & vbCrLf & _
                              "    .type = " & ReturnEsriFieldTypeNameFromNumber(pField.Type) & vbCrLf
      If pField.Type = esriFieldTypeDouble Or pField.Type = esriFieldTypeSingle Then
        strReport = strReport & "    .scale = " & pField.Scale & vbCrLf & _
                                "    .precision = " & pField.Precision & vbCrLf
      ElseIf pField.Type = esriFieldTypeString Then
        strReport = strReport & "    .length = " & pField.length & vbCrLf
      End If
      strReport = strReport & "  end with" & vbCrLf & _
                              "  pfieldarray.add pfield " & vbCrLf & vbCrLf
    
      strOrigDimReport = strOrigDimReport & "  dim lng" & pField.Name & "_Index as long" & vbCrLf
      
      strOrigFindReport = strOrigFindReport & "  lng" & pField.Name & "_Index = " & _
          strFClassTableVariableName & ".findfield(""" & pField.Name & """)" & vbCrLf
    
      strNewDimReport = strNewDimReport & "  dim lngNew" & pField.Name & "_Index as long" & vbCrLf
      
      strNewFindReport = strNewFindReport & "  lngNew" & pField.Name & "_Index = " & _
          strNewFClassTableVariableName & ".findfield(""" & pField.Name & """)" & vbCrLf

    End If
  Next lngIndex
  
  If booDescriptive Then
    WriteCodeToDuplicateFields = WriteCodeToDuplicateFields & strDescriptiveReport & vbCrLf
  End If
  
  If booMakeFields Then
    If WriteCodeToDuplicateFields = "" Then
      WriteCodeToDuplicateFields = WriteCodeToDuplicateFields & strReport & vbCrLf
    Else
      WriteCodeToDuplicateFields = WriteCodeToDuplicateFields & "'------------------------------" & vbCrLf & _
          strReport & vbCrLf
    End If
  End If
  
  If booDimAndFindOriginalFields Then
    If WriteCodeToDuplicateFields = "" Then
      WriteCodeToDuplicateFields = WriteCodeToDuplicateFields & strOrigDimReport & vbCrLf & _
          strOrigFindReport & vbCrLf
    Else
      WriteCodeToDuplicateFields = WriteCodeToDuplicateFields & "'------------------------------" & vbCrLf & _
          strOrigDimReport & vbCrLf & strOrigFindReport & vbCrLf
    End If
  End If
  
  If booDimAndFindNewFields Then
    If WriteCodeToDuplicateFields = "" Then
      WriteCodeToDuplicateFields = WriteCodeToDuplicateFields & strNewDimReport & vbCrLf & _
          strNewFindReport & vbCrLf
    Else
      WriteCodeToDuplicateFields = WriteCodeToDuplicateFields & "'------------------------------" & vbCrLf & _
          strNewDimReport & vbCrLf & strNewFindReport & vbCrLf
    End If
  End If
    

  Set pField = Nothing


End Function


Public Sub AddTableToMxDoc(pTable As ITable, pMxDoc As IMxDocument, pApp As IApplication)

  Dim pCountStandaloneTable As IStandaloneTable
  Dim pStandTableColl As IStandaloneTableCollection
  Dim pCountTableWindow As ITableWindow2
  
  Set pStandTableColl = pMxDoc.FocusMap
  
  Set pCountStandaloneTable = New StandaloneTable
  Set pCountStandaloneTable.Table = pTable
  Set pCountTableWindow = New TableWindow
  With pCountTableWindow
    Set .StandaloneTable = pCountStandaloneTable
    Set .Application = pApp
    .TableSelectionAction = esriSelectFeatures
    .ShowAliasNamesInColumnHeadings = True
    .ShowSelected = False
    .Show True
  End With
  pStandTableColl.AddStandaloneTable pCountStandaloneTable
  
  
  GoTo ClearMemory
ClearMemory:
  Set pCountStandaloneTable = Nothing
  Set pStandTableColl = Nothing
  Set pCountTableWindow = Nothing

End Sub

Public Sub ConvertLayoutGraphics()

'  Dim lngTransparency As Long
'  Dim dblOutlineWidth As Double
'  lngTransparency = 30
'  dblOutlineWidth = 2
'
'  ' ----------------------------------------------
'
'  Dim pMxDoc As IMxDocument
'  Set pMxDoc = ThisDocument
'
'  Dim pGeom As IGeometry
'  Dim pArray As esriSystem.IArray
'  Dim lngIndex As Long
'  Dim pPolygon As IPolygon
'
'  Dim pActiveView As IActiveView
'  Dim pDisplay As IScreenDisplay
'  Dim pDisplayTransform As IDisplayTransformation
'
'  Dim pPoint As IPoint
'  Dim pPtColl As IPointCollection
'  Dim pNewPtColl As IPointCollection
'  Dim pNewPolygon As IPolygon
'  Dim lngX As Long
'  Dim lngY As Long
'  Dim lngIndex2 As Long
'  Dim pNewPoint As IPoint
'  Dim pSpRef As ISpatialReference
'  Dim pMapView As IActiveView
'  Set pMapView = pMxDoc.FocusMap
'  Set pSpRef = pMxDoc.FocusMap.SpatialReference
'
'  Dim pNewFClass As IFeatureClass
'  Dim pWS As IWorkspace
'  Dim pWSFact As IWorkspaceFactory
'  Set pWSFact = New FileGDBWorkspaceFactory
'  Set pWS = pWSFact.OpenFromFile("D:\arcGIS_stuff\consultation\Springs_Stewardship_Institute\Range_Maps\Map_Boxes.gdb", 0)
'  Dim strName As String
'  Dim pWS2 As IWorkspace2
'  Set pWS2 = pWS
'  Dim pEnv As IEnvelope
'  Set pEnv = New Envelope
'  Set pEnv.SpatialReference = pSpRef
'  Dim pNewPolys As esriSystem.IArray
'  Dim pNewBuff As IFeatureBuffer
'  Dim pNewCursor As IFeatureCursor
'  Dim lngIDIndex As Long
'  Dim pNewFLayer As IFeatureLayer
'  Dim pDataset As IDataset
'  Dim pRender As ISimpleRenderer
'  Dim pFillSymbol As ISimpleFillSymbol
'  Dim pLineSymbol As ISimpleLineSymbol
'  Dim pWhite As IRgbColor
'  Dim pBlack As IRgbColor
'  Dim pLyr As IGeoFeatureLayer
'  Dim hx As IRendererPropertyPage
'  Dim pLayerEffects As ILayerEffects
'  Dim pNewFLayer2 As IFeatureLayer
'  Dim pGroupLayer As IGroupLayer
'
'  strName = MyGeneralOperations.MakeUniqueGDBFeatureClassName(pWS2, "Map_Boxes")
'
'  Set pArray = MyGeneralOperations.ReturnGraphicsByNameFromLayout(pMxDoc, "Box", False)
'  If pArray.Count > 0 Then
'    Set pNewPolys = New esriSystem.Array
'    For lngIndex = 0 To pArray.Count - 1
'      Set pGeom = pArray.Element(lngIndex)
'
'      If pGeom.GeometryType = esriGeometryPolygon Then
'        Set pActiveView = pMxDoc.PageLayout
'        Set pDisplay = pActiveView.ScreenDisplay
'        Set pDisplayTransform = pDisplay.DisplayTransformation
'
'        Set pPtColl = pGeom
'        Set pNewPolygon = New Polygon
'        Set pNewPolygon.SpatialReference = pSpRef
'        Set pNewPtColl = pNewPolygon
'
'        For lngIndex2 = 0 To pPtColl.PointCount - 1
'          Set pPoint = pPtColl.Point(lngIndex2)
'          pDisplayTransform.FromMapPoint pPoint, lngX, lngY
'          Set pNewPoint = pMapView.ScreenDisplay.DisplayTransformation.ToMapPoint(lngX, lngY)
'          Set pNewPoint.SpatialReference = pSpRef
'          pNewPtColl.AddPoint pNewPoint
'        Next lngIndex2
'        pNewPolygon.Close
'        pNewPolygon.SimplifyPreserveFromTo
'        pNewPolys.Add pNewPolygon
'        pEnv.Union pNewPolygon.Envelope
''        MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pNewPolygon, "Delete_Me"
'      End If
'    Next lngIndex
'
'    Set pNewFClass = MyGeneralOperations.CreateGDBFeatureClass(pWS, strName, esriFTSimple, pSpRef, esriGeometryPolygon, _
'      , , , , True, ENUM_FileGDB, pEnv, pArray.Count)
'    Set pNewBuff = pNewFClass.CreateFeatureBuffer
'    Set pNewCursor = pNewFClass.Insert(True)
'    lngIDIndex = pNewFClass.FindField("Unique_ID")
'
'    For lngIndex = 0 To pNewPolys.Count - 1
'      Set pNewPolygon = pNewPolys.Element(lngIndex)
'      Set pNewBuff.Shape = pNewPolygon
'      pNewCursor.InsertFeature pNewBuff
'    Next lngIndex
'
'    Set pWhite = MyGeneralOperations.MakeColorRGB(255, 255, 255)
'    Set pBlack = MyGeneralOperations.MakeColorRGB(0, 0, 0)
'
'    Set pNewFLayer = New FeatureLayer
'    Set pNewFLayer.FeatureClass = pNewFClass
'    Set pDataset = pNewFClass
'    pNewFLayer.Name = pDataset.BrowseName & " Fill"
'    Set pLyr = pNewFLayer
'    Set pRender = New SimpleRenderer
'    Set pLineSymbol = New SimpleLineSymbol
'    Set pFillSymbol = New SimpleFillSymbol
'    pLineSymbol.Width = 0
'    pLineSymbol.Style = esriSLSNull
'    pFillSymbol.Color = pWhite
'    pFillSymbol.Outline = pLineSymbol
'    pFillSymbol.Style = esriSFSSolid
'    Set pRender.Symbol = pFillSymbol
'    pRender.Label = "Fill"
'    Set pLyr.Renderer = pRender
'    Set hx = New SingleSymbolPropertyPage
'    pLyr.RendererPropertyPageClassID = hx.ClassID
'    Set pLayerEffects = pNewFLayer
'    pLayerEffects.Transparency = lngTransparency
'
'    Set pNewFLayer2 = New FeatureLayer
'    Set pNewFLayer2.FeatureClass = pNewFClass
'    Set pDataset = pNewFClass
'    pNewFLayer2.Name = pDataset.BrowseName & " Outline"
'    Set pLyr = pNewFLayer2
'    Set pRender = New SimpleRenderer
'    Set pLineSymbol = New SimpleLineSymbol
'    Set pFillSymbol = New SimpleFillSymbol
'    pLineSymbol.Width = 2
'    pLineSymbol.Style = esriSLSSolid
'    pLineSymbol.Color = pBlack
'    pFillSymbol.Outline = pLineSymbol
'    pFillSymbol.Style = esriSFSHollow
'    Set pRender.Symbol = pFillSymbol
'    pRender.Label = "Outline"
'    Set pLyr.Renderer = pRender
'    Set hx = New SingleSymbolPropertyPage
'    pLyr.RendererPropertyPageClassID = hx.ClassID
'
'    Set pGroupLayer = New GroupLayer
'    pGroupLayer.Add pNewFLayer
'    pGroupLayer.Add pNewFLayer2
'    pGroupLayer.Name = pDataset.BrowseName
'    pGroupLayer.Expanded = False
'
'    pMxDoc.FocusMap.AddLayer pGroupLayer
'    pMxDoc.UpdateContents
'    pMxDoc.ActiveView.Refresh
'  End If
'
''    Set pLyr = pFLayer
''
''    '** Make the renderer
''    Dim pRender As IUniqueValueRenderer
''    Dim n As Long
''    Set pRender = New UniqueValueRenderer
''
''    Dim pLineSymbol As ISimpleLineSymbol
''    Set pLineSymbol = New SimpleLineSymbol
''    pLineSymbol.Width = 1
''    pLineSymbol.Color = MyGeneralOperations.MakeColorRGB(38, 115, 0)
''
''    '** These properties should be set prior to adding values
''    pRender.FieldCount = 1
''    pRender.Field(0) = "Full_Name"
''    pRender.DefaultSymbol = Nothing
''    pRender.UseDefaultSymbol = False
''
''    Dim pMainSymbol As ISimpleFillSymbol
''    Set pMainSymbol = New SimpleFillSymbol
''    pMainSymbol.Style = esriSFSSolid
''    pMainSymbol.Outline = pLineSymbol
''    pMainSymbol.Color = MyGeneralOperations.MakeColorRGB(0, 200, 0)
''
''    pRender.AddValue strFullName, "Full_Name", pMainSymbol
''
''
''    pRender.ColorScheme = "Custom"
''    pRender.fieldType(0) = True
''    Set pLyr.Renderer = pRender
''    pLyr.DisplayField = "Full_Name"
''
''    '** This makes the layer properties symbology tab show
''    '** show the correct interface.
'
'
'
'ClearMemory:
'  Set pMxDoc = Nothing
'  Set pGeom = Nothing
'  Set pArray = Nothing
'  Set pPolygon = Nothing
'  Set pActiveView = Nothing
'  Set pDisplay = Nothing
'  Set pDisplayTransform = Nothing
'  Set pPoint = Nothing
'  Set pPtColl = Nothing
'  Set pNewPtColl = Nothing
'  Set pNewPolygon = Nothing
'  Set pSpRef = Nothing

End Sub

Public Function CreateGeneralProjectedSpatialReference(lngFactoryID As Long) As ISpatialReference
  
  Dim pGeneralGeo As IProjectedCoordinateSystem
  Dim pSpatRefFact As ISpatialReferenceFactory
  Set pSpatRefFact = New SpatialReferenceEnvironment
  Set pGeneralGeo = pSpatRefFact.CreateProjectedCoordinateSystem(lngFactoryID)
  Dim pSpRefRes As ISpatialReferenceResolution
  Set pSpRefRes = pGeneralGeo
  pSpRefRes.ConstructFromHorizon
  
  Set CreateGeneralProjectedSpatialReference = pGeneralGeo
  
  Set pGeneralGeo = Nothing
  Set pSpatRefFact = Nothing
  Set pSpRefRes = Nothing


  GoTo ClearMemory
ClearMemory:
  Set pGeneralGeo = Nothing
  Set pSpatRefFact = Nothing
  Set pSpRefRes = Nothing

End Function
Public Function CreateGeneralGeographicSpatialReference(lngFactoryID As Long) As ISpatialReference
  
  Dim pGeneralGeo As IGeographicCoordinateSystem
  Dim pSpatRefFact As ISpatialReferenceFactory
  Set pSpatRefFact = New SpatialReferenceEnvironment
  Set pGeneralGeo = pSpatRefFact.CreateGeographicCoordinateSystem(lngFactoryID)
  Dim pSpRefRes As ISpatialReferenceResolution
  Set pSpRefRes = pGeneralGeo
  pSpRefRes.ConstructFromHorizon
  
  Set CreateGeneralGeographicSpatialReference = pGeneralGeo
  
  Set pGeneralGeo = Nothing
  Set pSpatRefFact = Nothing
  Set pSpRefRes = Nothing


  GoTo ClearMemory
ClearMemory:
  Set pGeneralGeo = Nothing
  Set pSpatRefFact = Nothing
  Set pSpRefRes = Nothing

End Function

Public Function ReturnStateNameCollections() As Variant()

'  Dim strText As String
'  Dim pDataObj As New MSForms.DataObject
'  pDataObj.GetFromClipboard
'  strText = pDataObj.GetText
'  Dim strLines() As String
'  Dim strLineSplit() As String
'  Dim strLine As String
  
  Dim pNameKeyColl As New Collection
  Dim pAbbrevKeyColl As New Collection
'  Dim strName As String
'  Dim strAbbrev As String
  
'  Dim strReport As String
'
'  Dim lngIndex As Long
'  Dim lngIndex2 As Long
'  strLines = Split(strText, vbCrLf)
'  For lngIndex = 0 To UBound(strLines)
'    strLine = strLines(lngIndex)
'    strLineSplit = Split(strLine, Chr(9))
''    For lngIndex2 = 1 To Len(strLine)
''      Debug.Print CStr(lngIndex2) & "] " & Mid(strLine, lngIndex2, 1) & "  [chr = " & CStr(Asc(Mid(strLine, lngIndex2, 1))) & "]"
''    Next lngIndex2
'
'    strName = Trim(strLineSplit(0))
'    strAbbrev = Trim(strLineSplit(1))
'    Debug.Print "' " & strAbbrev & ", " & strName & ""
'    Debug.Print "  pNameKeyColl.add """ & strAbbrev & """, """ & strName & """"
'    Debug.Print "  pAbbrevKeyColl.add """ & strName & """, """ & strAbbrev & """" & vbCrLf
'
'    strReport = strReport & "' " & strAbbrev & ", " & strName & "" & vbCrLf & _
'                            "  pNameKeyColl.add """ & strAbbrev & """, """ & strName & """" & vbCrLf & _
'                            "  pAbbrevKeyColl.add """ & strName & """, """ & strAbbrev & """" & vbCrLf & vbCrLf
'  Next lngIndex
'
'  pDataObj.SetText strReport
'  pDataObj.PutInClipboard
'  Set pDataObj = Nothing
  ' AL, Alabama
  pNameKeyColl.Add "AL", "Alabama"
  pAbbrevKeyColl.Add "Alabama", "AL"

' AK, Alaska
  pNameKeyColl.Add "AK", "Alaska"
  pAbbrevKeyColl.Add "Alaska", "AK"

' AZ, Arizona
  pNameKeyColl.Add "AZ", "Arizona"
  pAbbrevKeyColl.Add "Arizona", "AZ"

' AR, Arkansas
  pNameKeyColl.Add "AR", "Arkansas"
  pAbbrevKeyColl.Add "Arkansas", "AR"

' CA, California
  pNameKeyColl.Add "CA", "California"
  pAbbrevKeyColl.Add "California", "CA"

' CO, Colorado
  pNameKeyColl.Add "CO", "Colorado"
  pAbbrevKeyColl.Add "Colorado", "CO"

' CT, Connecticut
  pNameKeyColl.Add "CT", "Connecticut"
  pAbbrevKeyColl.Add "Connecticut", "CT"

' DE, Delaware
  pNameKeyColl.Add "DE", "Delaware"
  pAbbrevKeyColl.Add "Delaware", "DE"

' FL, Florida
  pNameKeyColl.Add "FL", "Florida"
  pAbbrevKeyColl.Add "Florida", "FL"

' GA, Georgia
  pNameKeyColl.Add "GA", "Georgia"
  pAbbrevKeyColl.Add "Georgia", "GA"

' HI, Hawaii
  pNameKeyColl.Add "HI", "Hawaii"
  pAbbrevKeyColl.Add "Hawaii", "HI"

' ID, Idaho
  pNameKeyColl.Add "ID", "Idaho"
  pAbbrevKeyColl.Add "Idaho", "ID"

' IL, Illinois
  pNameKeyColl.Add "IL", "Illinois"
  pAbbrevKeyColl.Add "Illinois", "IL"

' IN, Indiana
  pNameKeyColl.Add "IN", "Indiana"
  pAbbrevKeyColl.Add "Indiana", "IN"

' IA, Iowa
  pNameKeyColl.Add "IA", "Iowa"
  pAbbrevKeyColl.Add "Iowa", "IA"

' KS, Kansas
  pNameKeyColl.Add "KS", "Kansas"
  pAbbrevKeyColl.Add "Kansas", "KS"

' KY, Kentucky
  pNameKeyColl.Add "KY", "Kentucky"
  pAbbrevKeyColl.Add "Kentucky", "KY"

' LA, Louisiana
  pNameKeyColl.Add "LA", "Louisiana"
  pAbbrevKeyColl.Add "Louisiana", "LA"

' ME, Maine
  pNameKeyColl.Add "ME", "Maine"
  pAbbrevKeyColl.Add "Maine", "ME"

' MD, Maryland
  pNameKeyColl.Add "MD", "Maryland"
  pAbbrevKeyColl.Add "Maryland", "MD"

' MA, Massachusetts
  pNameKeyColl.Add "MA", "Massachusetts"
  pAbbrevKeyColl.Add "Massachusetts", "MA"

' MI, Michigan
  pNameKeyColl.Add "MI", "Michigan"
  pAbbrevKeyColl.Add "Michigan", "MI"

' MN, Minnesota
  pNameKeyColl.Add "MN", "Minnesota"
  pAbbrevKeyColl.Add "Minnesota", "MN"

' MS, Mississippi
  pNameKeyColl.Add "MS", "Mississippi"
  pAbbrevKeyColl.Add "Mississippi", "MS"

' MO, Missouri
  pNameKeyColl.Add "MO", "Missouri"
  pAbbrevKeyColl.Add "Missouri", "MO"

' MT, Montana
  pNameKeyColl.Add "MT", "Montana"
  pAbbrevKeyColl.Add "Montana", "MT"

' NE, Nebraska
  pNameKeyColl.Add "NE", "Nebraska"
  pAbbrevKeyColl.Add "Nebraska", "NE"

' NV, Nevada
  pNameKeyColl.Add "NV", "Nevada"
  pAbbrevKeyColl.Add "Nevada", "NV"

' NH, New Hampshire
  pNameKeyColl.Add "NH", "New Hampshire"
  pAbbrevKeyColl.Add "New Hampshire", "NH"

' NJ, New Jersey
  pNameKeyColl.Add "NJ", "New Jersey"
  pAbbrevKeyColl.Add "New Jersey", "NJ"

' NM, New Mexico
  pNameKeyColl.Add "NM", "New Mexico"
  pAbbrevKeyColl.Add "New Mexico", "NM"

' NY, New York
  pNameKeyColl.Add "NY", "New York"
  pAbbrevKeyColl.Add "New York", "NY"

' NC, North Carolina
  pNameKeyColl.Add "NC", "North Carolina"
  pAbbrevKeyColl.Add "North Carolina", "NC"

' ND, North Dakota
  pNameKeyColl.Add "ND", "North Dakota"
  pAbbrevKeyColl.Add "North Dakota", "ND"

' OH, Ohio
  pNameKeyColl.Add "OH", "Ohio"
  pAbbrevKeyColl.Add "Ohio", "OH"

' OK, Oklahoma
  pNameKeyColl.Add "OK", "Oklahoma"
  pAbbrevKeyColl.Add "Oklahoma", "OK"

' OR, Oregon
  pNameKeyColl.Add "OR", "Oregon"
  pAbbrevKeyColl.Add "Oregon", "OR"

' PA, Pennsylvania
  pNameKeyColl.Add "PA", "Pennsylvania"
  pAbbrevKeyColl.Add "Pennsylvania", "PA"

' RI, Rhode Island
  pNameKeyColl.Add "RI", "Rhode Island"
  pAbbrevKeyColl.Add "Rhode Island", "RI"

' SC, South Carolina
  pNameKeyColl.Add "SC", "South Carolina"
  pAbbrevKeyColl.Add "South Carolina", "SC"

' SD, South Dakota
  pNameKeyColl.Add "SD", "South Dakota"
  pAbbrevKeyColl.Add "South Dakota", "SD"

' TN, Tennessee
  pNameKeyColl.Add "TN", "Tennessee"
  pAbbrevKeyColl.Add "Tennessee", "TN"

' TX, Texas
  pNameKeyColl.Add "TX", "Texas"
  pAbbrevKeyColl.Add "Texas", "TX"

' UT, Utah
  pNameKeyColl.Add "UT", "Utah"
  pAbbrevKeyColl.Add "Utah", "UT"

' VT, Vermont
  pNameKeyColl.Add "VT", "Vermont"
  pAbbrevKeyColl.Add "Vermont", "VT"

' VA, Virginia
  pNameKeyColl.Add "VA", "Virginia"
  pAbbrevKeyColl.Add "Virginia", "VA"

' WA, Washington
  pNameKeyColl.Add "WA", "Washington"
  pAbbrevKeyColl.Add "Washington", "WA"

' WV, West Virginia
  pNameKeyColl.Add "WV", "West Virginia"
  pAbbrevKeyColl.Add "West Virginia", "WV"

' WI, Wisconsin
  pNameKeyColl.Add "WI", "Wisconsin"
  pAbbrevKeyColl.Add "Wisconsin", "WI"

' WY, Wyoming
  pNameKeyColl.Add "WY", "Wyoming"
  pAbbrevKeyColl.Add "Wyoming", "WY"

' AS, American Samoa
  pNameKeyColl.Add "AS", "American Samoa"
  pAbbrevKeyColl.Add "American Samoa", "AS"

' DC, District of Columbia
  pNameKeyColl.Add "DC", "District of Columbia"
  pAbbrevKeyColl.Add "District of Columbia", "DC"

' FM, Federated States of Micronesia
  pNameKeyColl.Add "FM", "Federated States of Micronesia"
  pAbbrevKeyColl.Add "Federated States of Micronesia", "FM"

' GU, Guam
  pNameKeyColl.Add "GU", "Guam"
  pAbbrevKeyColl.Add "Guam", "GU"

' MH, Marshall Islands
  pNameKeyColl.Add "MH", "Marshall Islands"
  pAbbrevKeyColl.Add "Marshall Islands", "MH"

' MP, Northern Mariana Islands
  pNameKeyColl.Add "MP", "Northern Mariana Islands"
  pAbbrevKeyColl.Add "Northern Mariana Islands", "MP"

' PW, Palau
  pNameKeyColl.Add "PW", "Palau"
  pAbbrevKeyColl.Add "Palau", "PW"

' PR, Puerto Rico
  pNameKeyColl.Add "PR", "Puerto Rico"
  pAbbrevKeyColl.Add "Puerto Rico", "PR"

' VI, Virgin Islands
  pNameKeyColl.Add "VI", "Virgin Islands"
  pAbbrevKeyColl.Add "Virgin Islands", "VI"

  Dim varReturn() As Variant
  ReDim varReturn(1)
  Set varReturn(0) = pNameKeyColl
  Set varReturn(1) = pAbbrevKeyColl
  ReturnStateNameCollections = varReturn

End Function
Public Function ReturnStateAbbreviation(strState As String) As String
  
  Dim varColls() As Variant
  varColls = ReturnStateNameCollections
  Dim pNameKeyColl As Collection
  Set pNameKeyColl = varColls(0)
  
  If MyGeneralOperations.CheckCollectionForKey(pNameKeyColl, strState) Then
    ReturnStateAbbreviation = pNameKeyColl.Item(strState)
  Else
    ReturnStateAbbreviation = ""
  End If

End Function

Public Function CompareSpatialReferences2(ByVal pSourceSR As ISpatialReference, ByVal pTargetSR As ISpatialReference, _
      Optional bSREqual As Boolean, Optional booAlsoComparePrecision As Boolean = False) As Boolean
  
  
  Dim pSourceClone As IClone
  Dim pTargetClone As IClone
  Dim bXYIsEqual As Boolean
  
  Set pSourceClone = pSourceSR
  Set pTargetClone = pTargetSR
  
  'MsgBox "pSourceClone is nothing = " & CStr(pSourceClone Is Nothing) & vbCrLf & _
        "pTargetClone is nothing = " & CStr(pTargetClone Is Nothing)
  
  If pSourceClone Is Nothing And pTargetClone Is Nothing Then
    CompareSpatialReferences2 = True
  ElseIf pSourceClone Is Nothing Or pTargetClone Is Nothing Then
    CompareSpatialReferences2 = False
  Else
    
    'Compare the coordinate system component of the spatial reference
    bSREqual = pSourceClone.IsEqual(pTargetClone)
    
    'If the comparison failed, return false and exit
    If Not bSREqual Then
      CompareSpatialReferences2 = False
      Exit Function
    End If
    
    'We can also compare the XY precision to ensure the spatial references are equal
    If booAlsoComparePrecision Then
      Dim pSourceSR2 As ISpatialReference2
      
      Set pSourceSR2 = pSourceSR
      bXYIsEqual = pSourceSR2.IsXYPrecisionEqual(pTargetSR)
      
      'If the comparison failed, return false and exit
      If Not bXYIsEqual Then
        CompareSpatialReferences2 = False
        Exit Function
      End If
    End If
    CompareSpatialReferences2 = True
  End If


  GoTo ClearMemory
ClearMemory:
  Set pSourceClone = Nothing
  Set pTargetClone = Nothing
  Set pSourceSR2 = Nothing
End Function
Public Function MakeUniqueDataFrameName(strSuggestName As String, pMxDoc As IMxDocument) As String

  Dim pMaps As IMaps2
  Set pMaps = pMxDoc.Maps
    
  Dim pMap As IMap
  Dim theCounter As Long
  Dim booFoundDuplicate As Boolean
  booFoundDuplicate = True
  
  Dim strMapName As String
  Dim strBaseName As String
  Dim lngIndex As Long
  
  strBaseName = strSuggestName
  strMapName = strBaseName
  theCounter = 1
  
  Do Until booFoundDuplicate = False
    booFoundDuplicate = False
      
    For lngIndex = 0 To pMaps.Count - 1
      Set pMap = pMaps.Item(lngIndex)
      If StrComp(pMap.Name, strMapName, vbTextCompare) = 0 Then
        booFoundDuplicate = True
        
        theCounter = theCounter + 1
        strMapName = strBaseName & "_" & CStr(theCounter)
        Exit For
      End If
    Next lngIndex
  Loop
  
  MakeUniqueDataFrameName = strMapName
  
  GoTo ClearMemory
ClearMemory:
  Set pMaps = Nothing
  Set pMap = Nothing

End Function

Public Sub MakeFClassBorderAroundJennessentCompassRose(pMxDoc As IMxDocument)
 

  Dim lngTransparency As Long
  Dim dblOutlineWidth As Double
  lngTransparency = 30
  dblOutlineWidth = 2

  ' ----------------------------------------------


  Dim pGeom As IGeometry
  Dim pArray As esriSystem.IArray
  Dim lngIndex As Long
  Dim pPolygon As IPolygon

  Dim pActiveView As IActiveView
  Dim pDisplay As IScreenDisplay
  Dim pDisplayTransform As IDisplayTransformation

  Dim pPoint As IPoint
  Dim pPtColl As IPointCollection
  Dim pNewPtColl As IPointCollection
  Dim pNewPolygon As IPolygon
  Dim lngX As Long
  Dim lngY As Long
  Dim lngIndex2 As Long
  Dim pNewPoint As IPoint
  Dim pSpRef As ISpatialReference
  Dim pMapView As IActiveView
  Set pMapView = pMxDoc.FocusMap
  Set pSpRef = pMxDoc.FocusMap.SpatialReference

  Dim pNewFClass As IFeatureClass
  Dim pWS As IWorkspace
  Dim pWSFact As IWorkspaceFactory
  Set pWSFact = New FileGDBWorkspaceFactory
  Set pWS = pWSFact.OpenFromFile("D:\arcGIS_stuff\consultation\Ants_Gary_Beverly\Temp\Temp.gdb", 0)
  Dim strName As String
  Dim pWS2 As IWorkspace2
  Set pWS2 = pWS
  Dim pEnv As IEnvelope
  Set pEnv = New Envelope
  Set pEnv.SpatialReference = pSpRef
  Dim pNewPolys As esriSystem.IArray
  Dim pNewBuff As IFeatureBuffer
  Dim pNewCursor As IFeatureCursor
  Dim lngIDIndex As Long
  Dim pNewFlayer As IFeatureLayer
  Dim pDataset As IDataset
  Dim pRender As ISimpleRenderer
  Dim pFillSymbol As ISimpleFillSymbol
  Dim pLineSymbol As ISimpleLineSymbol
  Dim pWhite As IRgbColor
  Dim pBlack As IRgbColor
  Dim pLyr As IGeoFeatureLayer
  Dim hx As IRendererPropertyPage
  Dim pLayerEffects As ILayerEffects
  Dim pNewFLayer2 As IFeatureLayer
  Dim pGroupLayer As IGroupLayer
  Dim pElement As IElement
  Dim pVisEnv As IEnvelope
  
  Dim pCenter As IPoint
  
  strName = MyGeneralOperations.MakeUniqueGDBFeatureClassName(pWS2, "CompassRose")

  Set pArray = MyGeneralOperations.ReturnGraphicsByNameFromLayout(pMxDoc, "Jennessent Compass Rose", False)
  If pArray.Count > 0 Then
    Set pNewPolys = New esriSystem.Array
    For lngIndex = 0 To pArray.Count - 1
'      Set pElement = pArray.Element(lngIndex)
'      Set pVisEnv = New Envelope
'      pElement.QueryBounds pDisplay, pVisEnv
'      Set pGeom = pVisEnv
      Set pGeom = pArray.Element(lngIndex)

      If pGeom.GeometryType = esriGeometryPolygon Then
        Set pActiveView = pMxDoc.PageLayout
        Set pDisplay = pActiveView.ScreenDisplay
        Set pDisplayTransform = pDisplay.DisplayTransformation

        Set pPtColl = pGeom
        Set pNewPolygon = New Polygon
        Set pNewPolygon.SpatialReference = pSpRef
        Set pNewPtColl = pNewPolygon
        
        Dim pCentroid As IPoint
        

        For lngIndex2 = 0 To pPtColl.PointCount - 1
          Set pPoint = pPtColl.Point(lngIndex2)
          pDisplayTransform.FromMapPoint pPoint, lngX, lngY
          Set pNewPoint = pMapView.ScreenDisplay.DisplayTransformation.ToMapPoint(lngX, lngY)
          Set pNewPoint.SpatialReference = pSpRef
          pNewPtColl.AddPoint pNewPoint
        Next lngIndex2
        pNewPolygon.Close
        pNewPolygon.SimplifyPreserveFromTo
        pNewPolys.Add pNewPolygon
        pEnv.Union pNewPolygon.Envelope
'        MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pNewPolygon, "Delete_Me"
      End If
    Next lngIndex

        
    ' CONVERT TO CIRCLE AROUND COMPASS ROSE
    Set pCenter = New Point
    Set pCenter.SpatialReference = pNewPolygon.SpatialReference
    pCenter.PutCoords pEnv.XMin + (pEnv.Width / 2), pEnv.YMin + (pEnv.Height / 2)
    
    Set pNewPolygon = MyGeometricOperations.CreateCircleAroundPoint(pCenter, pEnv.Width * 0.55, 360)
    pNewPolys.RemoveAll
    pNewPolys.Add pNewPolygon
    Set pEnv = pNewPolygon.Envelope
    
    Set pNewFClass = MyGeneralOperations.CreateGDBFeatureClass(pWS, strName, esriFTSimple, pSpRef, esriGeometryPolygon, _
      , , , , True, ENUM_FileGDB, pEnv, 1)
    Set pNewBuff = pNewFClass.CreateFeatureBuffer
    Set pNewCursor = pNewFClass.Insert(True)
    lngIDIndex = pNewFClass.FindField("Unique_ID")

    For lngIndex = 0 To pNewPolys.Count - 1
      Set pNewPolygon = pNewPolys.Element(lngIndex)
      Set pNewBuff.Shape = pNewPolygon
      pNewCursor.InsertFeature pNewBuff
    Next lngIndex

    Set pWhite = MyGeneralOperations.MakeColorRGB(255, 255, 255)
    Set pBlack = MyGeneralOperations.MakeColorRGB(0, 0, 0)

    Set pNewFlayer = New FeatureLayer
    Set pNewFlayer.FeatureClass = pNewFClass
    Set pDataset = pNewFClass
    pNewFlayer.Name = pDataset.BrowseName & " Fill"
    Set pLyr = pNewFlayer
    Set pRender = New SimpleRenderer
    Set pLineSymbol = New SimpleLineSymbol
    Set pFillSymbol = New SimpleFillSymbol
    pLineSymbol.Width = 0
    pLineSymbol.Style = esriSLSNull
    pFillSymbol.Color = pWhite
    pFillSymbol.Outline = pLineSymbol
    pFillSymbol.Style = esriSFSSolid
    Set pRender.Symbol = pFillSymbol
    pRender.Label = "Fill"
    Set pLyr.Renderer = pRender
    Set hx = New SingleSymbolPropertyPage
    pLyr.RendererPropertyPageClassID = hx.ClassID
    Set pLayerEffects = pNewFlayer
    pLayerEffects.Transparency = lngTransparency

    Set pNewFLayer2 = New FeatureLayer
    Set pNewFLayer2.FeatureClass = pNewFClass
    Set pDataset = pNewFClass
    pNewFLayer2.Name = pDataset.BrowseName & " Outline"
    Set pLyr = pNewFLayer2
    Set pRender = New SimpleRenderer
    Set pLineSymbol = New SimpleLineSymbol
    Set pFillSymbol = New SimpleFillSymbol
    pLineSymbol.Width = 2
    pLineSymbol.Style = esriSLSSolid
    pLineSymbol.Color = pBlack
    pFillSymbol.Outline = pLineSymbol
    pFillSymbol.Style = esriSFSHollow
    Set pRender.Symbol = pFillSymbol
    pRender.Label = "Outline"
    Set pLyr.Renderer = pRender
    Set hx = New SingleSymbolPropertyPage
    pLyr.RendererPropertyPageClassID = hx.ClassID

    Set pGroupLayer = New GroupLayer
    pGroupLayer.Add pNewFlayer
    pGroupLayer.Add pNewFLayer2
    pGroupLayer.Name = pDataset.BrowseName
    pGroupLayer.Expanded = False

    pMxDoc.FocusMap.AddLayer pGroupLayer
    pMxDoc.UpdateContents
    pMxDoc.ActiveView.Refresh
  End If

'    Set pLyr = pFLayer
'
'    '** Make the renderer
'    Dim pRender As IUniqueValueRenderer
'    Dim n As Long
'    Set pRender = New UniqueValueRenderer
'
'    Dim pLineSymbol As ISimpleLineSymbol
'    Set pLineSymbol = New SimpleLineSymbol
'    pLineSymbol.Width = 1
'    pLineSymbol.Color = MyGeneralOperations.MakeColorRGB(38, 115, 0)
'
'    '** These properties should be set prior to adding values
'    pRender.FieldCount = 1
'    pRender.Field(0) = "Full_Name"
'    pRender.DefaultSymbol = Nothing
'    pRender.UseDefaultSymbol = False
'
'    Dim pMainSymbol As ISimpleFillSymbol
'    Set pMainSymbol = New SimpleFillSymbol
'    pMainSymbol.Style = esriSFSSolid
'    pMainSymbol.Outline = pLineSymbol
'    pMainSymbol.Color = MyGeneralOperations.MakeColorRGB(0, 200, 0)
'
'    pRender.AddValue strFullName, "Full_Name", pMainSymbol
'
'
'    pRender.ColorScheme = "Custom"
'    pRender.fieldType(0) = True
'    Set pLyr.Renderer = pRender
'    pLyr.DisplayField = "Full_Name"
'
'    '** This makes the layer properties symbology tab show
'    '** show the correct interface.


  GoTo ClearMemory

ClearMemory:
  Set pMxDoc = Nothing
  Set pGeom = Nothing
  Set pArray = Nothing
  Set pPolygon = Nothing
  Set pActiveView = Nothing
  Set pDisplay = Nothing
  Set pDisplayTransform = Nothing
  Set pPoint = Nothing
  Set pPtColl = Nothing
  Set pNewPtColl = Nothing
  Set pNewPolygon = Nothing
  Set pNewPoint = Nothing
  Set pSpRef = Nothing
  Set pMapView = Nothing
  Set pNewFClass = Nothing
  Set pWS = Nothing
  Set pWSFact = Nothing
  Set pWS2 = Nothing
  Set pEnv = Nothing
  Set pNewPolys = Nothing
  Set pNewBuff = Nothing
  Set pNewCursor = Nothing
  Set pNewFlayer = Nothing
  Set pDataset = Nothing
  Set pRender = Nothing
  Set pFillSymbol = Nothing
  Set pLineSymbol = Nothing
  Set pWhite = Nothing
  Set pBlack = Nothing
  Set pLyr = Nothing
  Set hx = Nothing
  Set pLayerEffects = Nothing
  Set pNewFLayer2 = Nothing
  Set pGroupLayer = Nothing
  Set pElement = Nothing
  Set pVisEnv = Nothing
  Set pCenter = Nothing
  Set pCentroid = Nothing

End Sub





Public Sub MakeNorthArrow(pMxDoc As IMxDocument)
  
  Dim dblRotation As Double
  Dim dblSize As Double
  Dim booMakeColors As Boolean
  Dim booMakeDiamonds As Boolean
  
  ' MAKE OPTIONAL FOR LAYOUT OR NOT
  ' MAKE OPTIONAL WHETHER TO DELETE PREVIOUS GRAPHICS NAMED "Jennessent Compass Rose"
  ' MAKE ROTATION OPTIONAL OR AUTOMATIC
  ' OPTIONAL ARROWHEAD FOR NORTH?
  ' OPTIONAL N,E,S,W?
  
  dblRotation = 0
  dblSize = 2
  booMakeColors = True
  booMakeDiamonds = True
  
'  dblSize = 2.3
  
  Debug.Print "----------------------------------"
  
  Dim dblRatio As Double
  dblRatio = dblSize / 2.3
  
  ' INITIAL VARIABLES
  Dim dblStep As Double
  Dim dblIndex2 As Double
  Dim pPolyline As IPolyline
  Dim pPolygon As IPolygon
  Dim pPolygon2 As IPolygon
  Dim pPolygon3 As IPolygon
  Dim pPolygon4 As IPolygon
  Dim pPolygon5 As IPolygon
  Dim pPolygon6 As IPolygon
  Dim pPtColl As IPointCollection
  Dim pPtColl2 As IPointCollection
  Dim pPtColl3 As IPointCollection
  Dim pPtColl4 As IPointCollection
  Dim pPtColl5 As IPointCollection
  Dim pPtColl6 As IPointCollection
  Dim pClone As IClone
  Dim dblIndex As Double
  Dim pPoint As IPoint
  Dim pOrigin As IPoint
  Dim pSpRef As ISpatialReference
  Set pSpRef = MyGeneralOperations.CreateSpatialReferenceWGS84
  Set pOrigin = New Point
  Set pOrigin.SpatialReference = pSpRef
  pOrigin.PutCoords 0, 0
  Set pClone = pOrigin
  Dim pTopoOp As ITopologicalOperator
  Dim pTransform2D As ITransform2D
  Dim booToggle As Boolean
  
  Dim pGroupElement As IGroupElement3
  Set pGroupElement = New GroupElement
  Dim pElementProperties As IElementProperties
  Set pElementProperties = pGroupElement
  pElementProperties.Name = "Jennessent Compass Rose"
  Dim booIslayout As Boolean
  booIslayout = TypeOf pMxDoc.ActivatedView Is IPageLayout
  
  MyGeneralOperations.DeleteGraphicsByName pMxDoc, "Delete_Me", booIslayout
  MyGeneralOperations.DeleteGraphicsByName pMxDoc, "Jennessent Compass Rose", booIslayout
  
  Dim pElement As IElement
  
  ' MAKE NAVAJO COLORS IF REQUESTED
  If booMakeColors Then
  
    ' NORTH
    Set pPolygon = New Polygon
    If Not booIslayout Then Set pPolygon.SpatialReference = pSpRef
    Set pPtColl = pPolygon
    pPtColl.AddPoint pClone.Clone
    For dblIndex = -22.5 To 22.5 Step 1
      MyGeometricOperations.CalcPointLine pOrigin, 0.95 * dblRatio, dblIndex, pPoint
      pPtColl.AddPoint pPoint
    Next dblIndex
    pPtColl.AddPoint pClone.Clone
    pPolygon.Close
    Set pElement = MyGeneralOperations.Graphic_ReturnElementFromGeometry3(pMxDoc, pPolygon, _
        "Jennessent Compass Rose Navajo Color", ReturnNorthSolidFill(ENUM_gray1, ENUM_Black, 0), booIslayout, False)
    pGroupElement.AddElement pElement
  
    ' EAST
    Set pPolygon = New Polygon
    If Not booIslayout Then Set pPolygon.SpatialReference = pSpRef
    Set pPtColl = pPolygon
    pPtColl.AddPoint pClone.Clone
    For dblIndex = 67.5 To 112.5 Step 1
      MyGeometricOperations.CalcPointLine pOrigin, 0.95 * dblRatio, dblIndex, pPoint
      pPtColl.AddPoint pPoint
    Next dblIndex
    pPtColl.AddPoint pClone.Clone
    pPolygon.Close
    Set pElement = MyGeneralOperations.Graphic_ReturnElementFromGeometry3(pMxDoc, pPolygon, _
        "Jennessent Compass Rose Navajo Color", ReturnNorthSolidFill(ENUM_white, ENUM_Black, 0), booIslayout, False)
    pGroupElement.AddElement pElement
  
    ' SOUTH
    Set pPolygon = New Polygon
    If Not booIslayout Then Set pPolygon.SpatialReference = pSpRef
    Set pPtColl = pPolygon
    pPtColl.AddPoint pClone.Clone
    For dblIndex = 180 - 22.5 To 180 + 22.5 Step 1
      MyGeometricOperations.CalcPointLine pOrigin, 0.95 * dblRatio, dblIndex, pPoint
      pPtColl.AddPoint pPoint
    Next dblIndex
    pPtColl.AddPoint pClone.Clone
    pPolygon.Close
    Set pElement = MyGeneralOperations.Graphic_ReturnElementFromGeometry3(pMxDoc, pPolygon, _
        "Jennessent Compass Rose Navajo Color", ReturnNorthSolidFill(ENUM_Dark_Turquoise, ENUM_Black, 0), booIslayout, False)
    pGroupElement.AddElement pElement
  
    ' WEST
    Set pPolygon = New Polygon
    If Not booIslayout Then Set pPolygon.SpatialReference = pSpRef
    Set pPtColl = pPolygon
    pPtColl.AddPoint pClone.Clone
    For dblIndex = 270 - 22.5 To 270 + 22.5 Step 1
      MyGeometricOperations.CalcPointLine pOrigin, 0.95 * dblRatio, dblIndex, pPoint
      pPtColl.AddPoint pPoint
    Next dblIndex
    pPtColl.AddPoint pClone.Clone
    pPolygon.Close
    Set pElement = MyGeneralOperations.Graphic_ReturnElementFromGeometry3(pMxDoc, pPolygon, _
        "Jennessent Compass Rose Navajo Color", ReturnNorthSolidFill(ENUM_Yellow, ENUM_Black, 0), booIslayout, False)
    pGroupElement.AddElement pElement
  
  End If
  
  
  ' MAKE LINES
  For dblIndex = 0 To 337.5 Step 22.5
    Set pPolyline = New Polyline
    If Not booIslayout Then Set pPolyline.SpatialReference = pSpRef
    Set pPtColl = pPolyline
    Set pPoint = pClone.Clone
    pPtColl.AddPoint pPoint
    MyGeometricOperations.CalcPointLine pOrigin, 0.95 * dblRatio, dblIndex, pPoint
    pPtColl.AddPoint pPoint
    
'    MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pPolyline, "Delete_Me", ReturnNorthSolidLineSymbol(ENUM_Black, 1)
    
    Set pElement = MyGeneralOperations.Graphic_ReturnElementFromGeometry3(pMxDoc, pPolyline, _
        "Jennessent Compass Rose Line", ReturnNorthSolidLineSymbol(ENUM_Black, 1 * dblRatio), booIslayout, False)
    pGroupElement.AddElement pElement
  Next dblIndex
  
  
  
  
  ' MAKE OUTER CIRCLE
  Set pPolygon = New Polygon
  If Not booIslayout Then Set pPolygon.SpatialReference = pSpRef
  Set pPtColl = pPolygon
  For dblIndex = 0 To 360
    MyGeometricOperations.CalcPointLine pOrigin, 1 * dblRatio, dblIndex, pPoint
    pPtColl.AddPoint pPoint
  Next dblIndex
  pPolygon.Close
  
  Set pPolygon2 = New Polygon
  If Not booIslayout Then Set pPolygon2.SpatialReference = pSpRef
  Set pPtColl = pPolygon2
  For dblIndex = 0 To 360
    MyGeometricOperations.CalcPointLine pOrigin, 0.9 * dblRatio, dblIndex, pPoint
    pPtColl.AddPoint pPoint
  Next dblIndex
  pPolygon2.Close
  
  Set pTopoOp = pPolygon
  Set pPolygon = pTopoOp.Difference(pPolygon2)
'  MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pPolygon, "Delete_Me", ReturnNorthSolidFill(ENUM_Black, ENUM_Black, 0)
  
  Set pElement = MyGeneralOperations.Graphic_ReturnElementFromGeometry3(pMxDoc, pPolygon, _
      "Jennessent Compass Rose Outer Circle", ReturnNorthSolidFill(ENUM_Black, ENUM_Black, 0), booIslayout, False)
  pGroupElement.AddElement pElement
  
  ' MAKE OUTER LINE INTERNAL BOXES
  dblStep = 30
  
  For dblIndex = 0 To 360 - dblStep Step dblStep
    booToggle = Not booToggle
    Set pPolygon = New Polygon
    If Not booIslayout Then Set pPolygon.SpatialReference = pSpRef
    Set pPtColl = pPolygon
    For dblIndex2 = dblIndex + 1 To dblIndex + dblStep - 1
      MyGeometricOperations.CalcPointLine pOrigin, 0.98 * dblRatio, dblIndex2, pPoint
      pPtColl.AddPoint pPoint
    Next dblIndex2
    MyGeometricOperations.CalcPointLine pOrigin, 0.92 * dblRatio, dblIndex + dblStep - 1, pPoint
    pPtColl.AddPoint pPoint
    For dblIndex2 = dblIndex + dblStep - 1 To dblIndex + 1 Step -1
      MyGeometricOperations.CalcPointLine pOrigin, 0.92 * dblRatio, dblIndex2, pPoint
      pPtColl.AddPoint pPoint
    Next dblIndex2
    pPolygon.Close
    If booToggle Then
'      MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pPolygon, "Delete_Me", ReturnNorthSolidFill(ENUM_Black, ENUM_Black, 0)
  
      Set pElement = MyGeneralOperations.Graphic_ReturnElementFromGeometry3(pMxDoc, pPolygon, _
          "Jennessent Compass Rose Outer Circle Black Intervals", ReturnNorthSolidFill(ENUM_Black, ENUM_Black, 0), _
          booIslayout, False)
      pGroupElement.AddElement pElement
    Else
'      MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pPolygon, "Delete_Me", ReturnNorthSolidFill(ENUM_White, ENUM_White, 0)
  
      Set pElement = MyGeneralOperations.Graphic_ReturnElementFromGeometry3(pMxDoc, pPolygon, _
          "Jennessent Compass Rose Outer Circle White Intervals", ReturnNorthSolidFill(ENUM_white, ENUM_white, 0), _
          booIslayout, False)
      pGroupElement.AddElement pElement
    End If
  Next dblIndex
  
  ' MAKE INNER TRIANGLES
  For dblIndex = 0 To 315 Step 45
    
    dblIndex2 = dblIndex + 22.5
    Set pPolygon3 = New Polygon
    If Not booIslayout Then Set pPolygon3.SpatialReference = pSpRef
    Set pPtColl3 = pPolygon3
    Set pPoint = pClone.Clone
    pPtColl3.AddPoint pPoint
    MyGeometricOperations.CalcPointLine pOrigin, 0.4 * dblRatio, dblIndex2 - 22.5, pPoint
    pPtColl3.AddPoint pPoint
    MyGeometricOperations.CalcPointLine pOrigin, 0.55 * dblRatio, dblIndex2, pPoint
    pPtColl3.AddPoint pPoint
    pPolygon3.Close
'    MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pPolygon3, "Delete_Me", ReturnNorthSolidFill(ENUM_Black, ENUM_Black, 1)
    
    Set pElement = MyGeneralOperations.Graphic_ReturnElementFromGeometry3(pMxDoc, pPolygon3, _
        "Jennessent Compass Rose 3rd-Order Triangles", ReturnNorthSolidFill(ENUM_Black, ENUM_Black, 1), booIslayout, False)
    pGroupElement.AddElement pElement
    
    Set pPolygon4 = New Polygon
    If Not booIslayout Then Set pPolygon4.SpatialReference = pSpRef
    Set pPtColl4 = pPolygon4
    Set pPoint = pClone.Clone
    pPtColl4.AddPoint pPoint
    MyGeometricOperations.CalcPointLine pOrigin, 0.55 * dblRatio, dblIndex2, pPoint
    pPtColl4.AddPoint pPoint
    MyGeometricOperations.CalcPointLine pOrigin, 0.4 * dblRatio, dblIndex2 + 22.5, pPoint
    pPtColl4.AddPoint pPoint
    pPolygon4.Close
'    MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pPolygon4, "Delete_Me", ReturnNorthSolidFill(ENUM_White, ENUM_Black, 1)
    
    Set pElement = MyGeneralOperations.Graphic_ReturnElementFromGeometry3(pMxDoc, pPolygon4, _
        "Jennessent Compass Rose 3rd-Order Triangles", ReturnNorthSolidFill(ENUM_white, ENUM_Black, 1), booIslayout, False)
    pGroupElement.AddElement pElement
  Next dblIndex
  
  ' MAKE INNER CIRCLE
  Set pPolygon = New Polygon
  If Not booIslayout Then Set pPolygon.SpatialReference = pSpRef
  Set pPtColl = pPolygon
  For dblIndex = 0 To 360
    MyGeometricOperations.CalcPointLine pOrigin, 0.45 * dblRatio, dblIndex, pPoint
    pPtColl.AddPoint pPoint
  Next dblIndex
  pPolygon.Close
'  MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pPolygon, "Delete_Me", ReturnNorthSolidFill(ENUM_Gray3, ENUM_Black, 1)
    
  Set pElement = MyGeneralOperations.Graphic_ReturnElementFromGeometry3(pMxDoc, pPolygon, _
      "Jennessent Compass Rose Center Circle", ReturnNorthSolidFill(ENUM_gray3, ENUM_Black, 1), booIslayout, False)
  pGroupElement.AddElement pElement
  
  For dblIndex = 0 To 270 Step 90
    
    dblIndex2 = dblIndex + 45
    Set pPolygon3 = New Polygon
    If Not booIslayout Then Set pPolygon3.SpatialReference = pSpRef
    Set pPtColl3 = pPolygon3
    Set pPoint = pClone.Clone
    pPtColl3.AddPoint pPoint
    MyGeometricOperations.CalcPointLine pOrigin, 0.25 * dblRatio, dblIndex2 - 45, pPoint
    pPtColl3.AddPoint pPoint
    MyGeometricOperations.CalcPointLine pOrigin, 0.75 * dblRatio, dblIndex2, pPoint
    pPtColl3.AddPoint pPoint
    pPolygon3.Close
'    MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pPolygon3, "Delete_Me", ReturnNorthSolidFill(ENUM_Black, ENUM_Black, 1)
    
    Set pElement = MyGeneralOperations.Graphic_ReturnElementFromGeometry3(pMxDoc, pPolygon3, _
        "Jennessent Compass Rose 2nd-Order Triangles", ReturnNorthSolidFill(ENUM_Black, ENUM_Black, 1), booIslayout, False)
    pGroupElement.AddElement pElement
    
    Set pPolygon4 = New Polygon
    If Not booIslayout Then Set pPolygon4.SpatialReference = pSpRef
    Set pPtColl4 = pPolygon4
    Set pPoint = pClone.Clone
    pPtColl4.AddPoint pPoint
    MyGeometricOperations.CalcPointLine pOrigin, 0.75 * dblRatio, dblIndex2, pPoint
    pPtColl4.AddPoint pPoint
    MyGeometricOperations.CalcPointLine pOrigin, 0.25 * dblRatio, dblIndex2 + 45, pPoint
    pPtColl4.AddPoint pPoint
    pPolygon4.Close
'    MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pPolygon4, "Delete_Me", ReturnNorthSolidFill(ENUM_White, ENUM_Black, 1)
    
    Set pElement = MyGeneralOperations.Graphic_ReturnElementFromGeometry3(pMxDoc, pPolygon4, _
        "Jennessent Compass Rose 2nd-Order Triangles", ReturnNorthSolidFill(ENUM_white, ENUM_Black, 1), booIslayout, False)
    pGroupElement.AddElement pElement
  Next dblIndex
  
  For dblIndex = 0 To 270 Step 90
    Set pPolygon = New Polygon
    If Not booIslayout Then Set pPolygon.SpatialReference = pSpRef
    Set pPtColl = pPolygon
    Set pPoint = pClone.Clone
    pPtColl.AddPoint pPoint
    MyGeometricOperations.CalcPointLine pOrigin, 0.15 * dblRatio, dblIndex - 45, pPoint
    pPtColl.AddPoint pPoint
    MyGeometricOperations.CalcPointLine pOrigin, 0.95 * dblRatio, dblIndex, pPoint
    pPtColl.AddPoint pPoint
    pPolygon.Close
'    MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pPolygon, "Delete_Me", ReturnNorthSolidFill(ENUM_Black, ENUM_Black, 1)
    
    Set pElement = MyGeneralOperations.Graphic_ReturnElementFromGeometry3(pMxDoc, pPolygon, _
        "Jennessent Compass Rose 1st-Order Triangles", ReturnNorthSolidFill(ENUM_Black, ENUM_Black, 1), booIslayout, False)
    pGroupElement.AddElement pElement
    
    Set pPolygon2 = New Polygon
    If Not booIslayout Then Set pPolygon2.SpatialReference = pSpRef
    Set pPtColl2 = pPolygon2
    Set pPoint = pClone.Clone
    pPtColl2.AddPoint pPoint
    MyGeometricOperations.CalcPointLine pOrigin, 0.95 * dblRatio, dblIndex, pPoint
    pPtColl2.AddPoint pPoint
    MyGeometricOperations.CalcPointLine pOrigin, 0.15 * dblRatio, dblIndex + 45, pPoint
    pPtColl2.AddPoint pPoint
    pPolygon2.Close
'    MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pPolygon2, "Delete_Me", ReturnNorthSolidFill(ENUM_White, ENUM_Black, 1)
    
    Set pElement = MyGeneralOperations.Graphic_ReturnElementFromGeometry3(pMxDoc, pPolygon2, _
        "Jennessent Compass Rose 1st-Order Triangles", ReturnNorthSolidFill(ENUM_white, ENUM_Black, 1), booIslayout, False)
    pGroupElement.AddElement pElement
    
  Next dblIndex
  
  Dim pAltOrigin As IPoint
  Dim dblIndex3 As Double
  Dim pWestPoly As IPolygon
  Set pWestPoly = New Polygon
  Set pPtColl = pWestPoly
  If Not booIslayout Then Set pWestPoly.SpatialReference = pSpRef
  Set pPoint = New Point
  pPoint.PutCoords 0, 0
  pPtColl.AddPoint pPoint
  Set pPoint = New Point
  pPoint.PutCoords -2 * dblRatio, 0
  pPtColl.AddPoint pPoint
  Set pPoint = New Point
  pPoint.PutCoords -2 * dblRatio, 2 * dblRatio
  pPtColl.AddPoint pPoint
  Set pPoint = New Point
  pPoint.PutCoords 0, 2 * dblRatio
  pPtColl.AddPoint pPoint
  pWestPoly.Close
  Dim pWestTopoOp As ITopologicalOperator
  Set pWestTopoOp = pWestPoly
    
  ' MAKE OUTER TRIANGLES
  For dblIndex = 0 To 315 Step 45
    Set pPolygon = New Polygon
    If Not booIslayout Then Set pPolygon.SpatialReference = pSpRef
    Set pPtColl = pPolygon
    
    ' NORTH, EAST, SOUTH, WEST ARROWHEADS
    If dblIndex Mod 90 = 0 Then
      MyGeometricOperations.CalcPointLine pOrigin, 1.15 * dblRatio, dblIndex, pPoint
      pPtColl.AddPoint pPoint
      For dblIndex2 = dblIndex + 8 To dblIndex + 3 Step -1
        MyGeometricOperations.CalcPointLine pOrigin, 0.88 * dblRatio, dblIndex2, pPoint
        pPtColl.AddPoint pPoint
      Next dblIndex2
      For dblIndex2 = dblIndex + 5 To dblIndex - 5 Step -1
        MyGeometricOperations.CalcPointLine pOrigin, 0.84 * dblRatio, dblIndex2, pPoint
        pPtColl.AddPoint pPoint
      Next dblIndex2
      For dblIndex2 = dblIndex - 3 To dblIndex - 8 Step -1
        MyGeometricOperations.CalcPointLine pOrigin, 0.88 * dblRatio, dblIndex2, pPoint
        pPtColl.AddPoint pPoint
      Next dblIndex2
      pPolygon.Close
      
      Set pElement = MyGeneralOperations.Graphic_ReturnElementFromGeometry3(pMxDoc, pPolygon, _
          "Jennessent Compass Rose Outer Triangles", ReturnNorthSolidFill(ENUM_GrayVeryLight, ENUM_Black, 1 * dblRatio), booIslayout, False)
      
      ' internal arrowhead
      pGroupElement.AddElement pElement
      Set pTopoOp = pPolygon
      Set pPolygon = pTopoOp.Buffer(-0.02 * dblRatio)
      
      Set pElement = MyGeneralOperations.Graphic_ReturnElementFromGeometry3(pMxDoc, pPolygon, _
          "Jennessent Compass Rose North Triangle", ReturnNorthSolidFill(ENUM_GrayVeryLight, ENUM_Black, 0.3 * dblRatio), booIslayout, False)
      pGroupElement.AddElement pElement
      
      
      If dblIndex = 0 Then
        MyGeometricOperations.CalcPointLine pOrigin, 0.5 * dblRatio, 90, pAltOrigin
        Set pTopoOp = pPolygon
        Set pPolygon = pTopoOp.Buffer(-0.01 * dblRatio)
        Set pTopoOp = pPolygon
        
        For dblIndex2 = 1.25 To 0.9 Step -0.015
          Set pPolyline = New Polyline
          If Not booIslayout Then Set pPolyline.SpatialReference = pSpRef
          Set pPtColl = pPolyline
          
          For dblIndex3 = 325 To 340 Step 0.5
        
            MyGeometricOperations.CalcPointLine pAltOrigin, dblIndex2 * dblRatio, dblIndex3, pPoint
            pPtColl.AddPoint pPoint
                           
          Next dblIndex3
          Set pPolyline = pTopoOp.Intersect(pPolyline, pPolyline.Dimension)
          Set pPolyline = pWestTopoOp.Intersect(pPolyline, pPolyline.Dimension)
          
          Set pElement = MyGeneralOperations.Graphic_ReturnElementFromGeometry3(pMxDoc, pPolyline, _
              "Jennessent Compass Rose North Triangle Lines", ReturnNorthSolidLineSymbol(ENUM_Black, 0.3 * dblRatio), booIslayout, False)
          pGroupElement.AddElement pElement
        Next dblIndex2
      End If
      
    ' ALL OTHER ARROWS
    Else
      
      MyGeometricOperations.CalcPointLine pOrigin, 1.08 * dblRatio, dblIndex, pPoint
      pPtColl.AddPoint pPoint
      For dblIndex2 = dblIndex + 5 To dblIndex - 5 Step -1
        MyGeometricOperations.CalcPointLine pOrigin, 0.88 * dblRatio, dblIndex2, pPoint
        pPtColl.AddPoint pPoint
      Next dblIndex2
  '
  '    MyGeometricOperations.CalcPointLine pOrigin, 0.88, dblIndex - 5, pPoint
  '    pPtColl.AddPoint pPoint
  '    MyGeometricOperations.CalcPointLine pOrigin, 1.08, dblIndex, pPoint
  '    pPtColl.AddPoint pPoint
  '    MyGeometricOperations.CalcPointLine pOrigin, 0.88, dblIndex + 5, pPoint
  '    pPtColl.AddPoint pPoint
      pPolygon.Close
  '    MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pPolygon, "Delete_Me", ReturnNorthSolidFill(ENUM_White, ENUM_Black, 1)
      Set pElement = MyGeneralOperations.Graphic_ReturnElementFromGeometry3(pMxDoc, pPolygon, _
          "Jennessent Compass Rose Outer Triangles", ReturnNorthSolidFill(ENUM_white, ENUM_Black, 1 * dblRatio), booIslayout, False)
      pGroupElement.AddElement pElement
    End If
    
    
  Next dblIndex
  
  
  ' MAKE INNER CIRCLE
  Set pPolygon = New Polygon
  If Not booIslayout Then Set pPolygon.SpatialReference = pSpRef
  Set pPtColl = pPolygon
  For dblIndex = 0 To 360
    MyGeometricOperations.CalcPointLine pOrigin, 0.3 * dblRatio, dblIndex, pPoint
    pPtColl.AddPoint pPoint
  Next dblIndex
  pPolygon.Close
'  MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pPolygon, "Delete_Me", ReturnNorthSolidFill(ENUM_Gray2, ENUM_Black, 1.5)
  
  If booMakeColors Then
    Set pElement = MyGeneralOperations.Graphic_ReturnElementFromGeometry3(pMxDoc, pPolygon, _
        "Jennessent Compass Rose Inner Circle", ReturnNorthSolidFill(ENUM_GanadoRed, ENUM_Black, 1.5 * dblRatio), booIslayout, False)
  Else
    Set pElement = MyGeneralOperations.Graphic_ReturnElementFromGeometry3(pMxDoc, pPolygon, _
        "Jennessent Compass Rose Inner Circle", ReturnNorthSolidFill(ENUM_gray2, ENUM_Black, 1.5 * dblRatio), booIslayout, False)
  End If
  pGroupElement.AddElement pElement

  
  ' MAKE NAVAJO
  Dim dblRowHeight As Double
  Dim dblHorizLength As Double
  Dim dblTopHorizLength As Double
  Dim dblRowAngleLength As Double
  Dim pNewPoint As IPoint
  dblRowHeight = 0.07 * dblRatio
  dblHorizLength = dblRowHeight * (1 + Sqr(3))
  dblTopHorizLength = dblHorizLength * 1.5
  dblRowAngleLength = Sqr(2 * (dblRowHeight ^ 2))
  Set pPolygon = New Polygon
  If Not booIslayout Then Set pPolygon.SpatialReference = pSpRef
  Set pPtColl = pPolygon
  Set pPoint = New Point
  Set pPoint.SpatialReference = pSpRef
  pPoint.PutCoords -dblTopHorizLength / 2, dblRowHeight * 2
  pPtColl.AddPoint pPoint

  MyGeometricOperations.CalcPointLine pPoint, dblTopHorizLength, 90, pNewPoint
  pPtColl.AddPoint pNewPoint
  Set pPoint = pNewPoint
  MyGeometricOperations.CalcPointLine pPoint, dblRowAngleLength, 225, pNewPoint
  pPtColl.AddPoint pNewPoint
  Set pPoint = pNewPoint
  MyGeometricOperations.CalcPointLine pPoint, dblHorizLength, 90, pNewPoint
  pPtColl.AddPoint pNewPoint
  Set pPoint = pNewPoint
  MyGeometricOperations.CalcPointLine pPoint, dblRowAngleLength, 225, pNewPoint
  pPtColl.AddPoint pNewPoint
  Set pPoint = pNewPoint
  MyGeometricOperations.CalcPointLine pPoint, dblRowAngleLength, 135, pNewPoint
  pPtColl.AddPoint pNewPoint
  Set pPoint = pNewPoint
  MyGeometricOperations.CalcPointLine pPoint, dblHorizLength, 270, pNewPoint
  pPtColl.AddPoint pNewPoint
  Set pPoint = pNewPoint
  MyGeometricOperations.CalcPointLine pPoint, dblRowAngleLength, 135, pNewPoint
  pPtColl.AddPoint pNewPoint
  Set pPoint = pNewPoint
  MyGeometricOperations.CalcPointLine pPoint, dblTopHorizLength, 270, pNewPoint
  pPtColl.AddPoint pNewPoint
  Set pPoint = pNewPoint
  MyGeometricOperations.CalcPointLine pPoint, dblRowAngleLength, 45, pNewPoint
  pPtColl.AddPoint pNewPoint
  Set pPoint = pNewPoint
  MyGeometricOperations.CalcPointLine pPoint, dblHorizLength, 270, pNewPoint
  pPtColl.AddPoint pNewPoint
  Set pPoint = pNewPoint
  MyGeometricOperations.CalcPointLine pPoint, dblRowAngleLength, 45, pNewPoint
  pPtColl.AddPoint pNewPoint
  Set pPoint = pNewPoint
  MyGeometricOperations.CalcPointLine pPoint, dblRowAngleLength, 315, pNewPoint
  pPtColl.AddPoint pNewPoint
  Set pPoint = pNewPoint
  MyGeometricOperations.CalcPointLine pPoint, dblHorizLength, 90, pNewPoint
  pPtColl.AddPoint pNewPoint
  Set pPoint = pNewPoint
  MyGeometricOperations.CalcPointLine pPoint, dblRowAngleLength, 315, pNewPoint
  pPtColl.AddPoint pNewPoint
  Set pPoint = pNewPoint
  pPolygon.Close
  ' MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pPolygon, "Delete_Me", ReturnNorthSolidFill(ENUM_White, ENUM_Black, 1.5)
  
  Set pElement = MyGeneralOperations.Graphic_ReturnElementFromGeometry3(pMxDoc, pPolygon, _
      "Jennessent Compass Center Design", ReturnNorthSolidFill(ENUM_white, ENUM_Black, 1 * dblRatio), booIslayout, False)
  pGroupElement.AddElement pElement
  
  If booMakeDiamonds Then
    Dim dblDiamondHeight As Double
    dblDiamondHeight = dblRowHeight * 0.7
    Dim dblDiamondVertDist As Double
    dblDiamondVertDist = dblRowHeight * 3.1

    Set pPolygon = New Polygon
    If Not booIslayout Then Set pPolygon.SpatialReference = pSpRef
    Set pPtColl = pPolygon

    Set pPoint = New Point
    pPoint.PutCoords -dblDiamondHeight * 2, 0
    pPtColl.AddPoint pPoint
    Set pPoint = New Point
    pPoint.PutCoords 0, dblDiamondHeight
    pPtColl.AddPoint pPoint
    Set pPoint = New Point
    pPoint.PutCoords dblDiamondHeight * 2, 0
    pPtColl.AddPoint pPoint
    Set pPoint = New Point
    pPoint.PutCoords 0, -dblDiamondHeight
    pPtColl.AddPoint pPoint

    pPolygon.Close
    If booMakeColors Then
      Set pElement = MyGeneralOperations.Graphic_ReturnElementFromGeometry3(pMxDoc, pPolygon, _
          "Jennessent Compass Rose Diamond1", ReturnNorthSolidFill(ENUM_Black, ENUM_Black, 0.5 * dblRatio), booIslayout, False)
    Else
      Set pElement = MyGeneralOperations.Graphic_ReturnElementFromGeometry3(pMxDoc, pPolygon, _
          "Jennessent Compass Rose Diamond1", ReturnNorthSolidFill(ENUM_gray3, ENUM_Black, 1 * dblRatio), booIslayout, False)
    End If
    pGroupElement.AddElement pElement
    
    Set pPolygon = New Polygon
    If Not booIslayout Then Set pPolygon.SpatialReference = pSpRef
    Set pPtColl = pPolygon

    Set pPoint = New Point
    pPoint.PutCoords -dblDiamondHeight, dblDiamondVertDist
    pPtColl.AddPoint pPoint
    Set pPoint = New Point
    pPoint.PutCoords 0, dblDiamondVertDist + dblDiamondHeight
    pPtColl.AddPoint pPoint
    Set pPoint = New Point
    pPoint.PutCoords dblDiamondHeight, dblDiamondVertDist
    pPtColl.AddPoint pPoint
    Set pPoint = New Point
    pPoint.PutCoords 0, dblDiamondVertDist - dblDiamondHeight
    pPtColl.AddPoint pPoint

    pPolygon.Close
    If booMakeColors Then
      Set pElement = MyGeneralOperations.Graphic_ReturnElementFromGeometry3(pMxDoc, pPolygon, _
          "Jennessent Compass Rose Diamond2", ReturnNorthSolidFill(ENUM_white, ENUM_Black, 0.75 * dblRatio), booIslayout, False)
    Else
      Set pElement = MyGeneralOperations.Graphic_ReturnElementFromGeometry3(pMxDoc, pPolygon, _
          "Jennessent Compass Rose Diamond2", ReturnNorthSolidFill(ENUM_white, ENUM_Black, 0.75 * dblRatio), booIslayout, False)
    End If
    pGroupElement.AddElement pElement
    
    Set pPolygon = New Polygon
    If Not booIslayout Then Set pPolygon.SpatialReference = pSpRef
    Set pPtColl = pPolygon

    Set pPoint = New Point
    pPoint.PutCoords -dblDiamondHeight, -dblDiamondVertDist
    pPtColl.AddPoint pPoint
    Set pPoint = New Point
    pPoint.PutCoords 0, -dblDiamondVertDist + dblDiamondHeight
    pPtColl.AddPoint pPoint
    Set pPoint = New Point
    pPoint.PutCoords dblDiamondHeight, -dblDiamondVertDist
    pPtColl.AddPoint pPoint
    Set pPoint = New Point
    pPoint.PutCoords 0, -dblDiamondVertDist - dblDiamondHeight
    pPtColl.AddPoint pPoint

    pPolygon.Close
    If booMakeColors Then
      Set pElement = MyGeneralOperations.Graphic_ReturnElementFromGeometry3(pMxDoc, pPolygon, _
          "Jennessent Compass Rose Diamond2", ReturnNorthSolidFill(ENUM_white, ENUM_Black, 0.75 * dblRatio), booIslayout, False)
    Else
      Set pElement = MyGeneralOperations.Graphic_ReturnElementFromGeometry3(pMxDoc, pPolygon, _
          "Jennessent Compass Rose Diamond2", ReturnNorthSolidFill(ENUM_white, ENUM_Black, 0.75 * dblRatio), booIslayout, False)
    End If
    pGroupElement.AddElement pElement
    
  End If

  
  Set pTransform2D = pGroupElement
  pTransform2D.Rotate pOrigin, MyGeometricOperations.ConvertRotationCompassDegreesToMathRadians(dblRotation)
  
  Dim pGContainer As IGraphicsContainer
  If booIslayout Then
    Set pGContainer = pMxDoc.PageLayout
  Else
    Set pGContainer = pMxDoc.FocusMap
  End If
  pGContainer.AddElement pGroupElement, 0
  
  pMxDoc.ActiveView.Refresh
  
  Debug.Print "Done..............."
  
  GoTo ClearMemory
ClearMemory:
  Set pMxDoc = Nothing
  Set pPolyline = Nothing
  Set pPolygon = Nothing
  Set pPolygon2 = Nothing
  Set pPolygon3 = Nothing
  Set pPolygon4 = Nothing
  Set pPolygon5 = Nothing
  Set pPolygon6 = Nothing
  Set pPtColl = Nothing
  Set pPtColl2 = Nothing
  Set pPtColl3 = Nothing
  Set pPtColl4 = Nothing
  Set pPtColl5 = Nothing
  Set pPtColl6 = Nothing
  Set pClone = Nothing
  Set pPoint = Nothing
  Set pOrigin = Nothing
  Set pSpRef = Nothing
  Set pTopoOp = Nothing
  Set pTransform2D = Nothing
  Set pGroupElement = Nothing
  Set pElementProperties = Nothing
  Set pElement = Nothing
  Set pAltOrigin = Nothing
  Set pWestPoly = Nothing
  Set pWestTopoOp = Nothing
  Set pNewPoint = Nothing
  Set pGContainer = Nothing

End Sub

Public Function Graphic_ReturnElementFromGeometry3(ByRef pMxDoc As IMxDocument, ByRef pGeometry As IGeometry, Optional strName As String, _
    Optional pSymbol As ISymbol, Optional booElementIsForLayout As Boolean = True, _
    Optional booAlsoAddElementToMapDoc As Boolean = False) As IElement
  
  Dim pMxDocument As esriArcMapUI.IMxDocument
  Dim pActiveView As esriCarto.IActiveView
  
  Dim pGContainer As IGraphicsContainer
  If booElementIsForLayout Then
    Set pGContainer = pMxDoc.PageLayout
  Else
    Set pGContainer = pMxDoc.FocusMap
  End If
  
  Dim pElement As IElement
  Dim pPolygonElement As IPolygonElement
  Dim pSpatialReference As ISpatialReference
  Dim pGraphicElement As IGraphicElement
  Dim pElementProperties As IElementProperties
  
  Dim pMarkerElement As IMarkerElement
  Dim pFillElement As IFillShapeElement
  Dim pLineElement As ILineElement
  
  Dim pClone As IClone
  Set pClone = pGeometry
  Dim pNewGeometry As IGeometry
  Set pNewGeometry = pClone.Clone
  
  Dim pGeometryType As esriGeometryType
  pGeometryType = pNewGeometry.GeometryType
  
  'ADD GEOMETRY, NAME AND SPATIAL REFERENCE TO GRAPHIC ELEMENT
  Select Case pGeometryType
    Case 0:
      MsgBox "Null Geometry!  No graphic added..."
    Case 1:
      Set pElement = New MarkerElement
      Set pMarkerElement = pElement
    Case 2:
      Set pElement = New GroupElement
    Case 3, 6, 13, 14, 15, 16:
      Set pElement = New LineElement
      Set pLineElement = pElement
    Case 4, 11:
      Set pElement = New PolygonElement
      Set pFillElement = pElement
    Case 5:
      Set pElement = New RectangleElement
      Set pFillElement = pElement
    Case Else:
      MsgBox "Unexpected Shape Type:  Unable to add graphic..."
  End Select
  
  If pGeometryType = 2 Then
    Dim pGroupElement As IGroupElement2
    Set pGroupElement = New GroupElement
    Dim pSubElement As IElement
    Set pSubElement = New MarkerElement
    Dim lngIndex As Long
    Dim pPtColl As IPointCollection
    Set pPtColl = pNewGeometry
    Dim pPt As IPoint
    For lngIndex = 0 To pPtColl.PointCount - 1
      Set pPt = pPtColl.Point(lngIndex)
      Set pSubElement = New MarkerElement
      pSubElement.Geometry = pPt
      pGroupElement.AddElement pSubElement
    Next lngIndex
    Set pGraphicElement = pGroupElement
    Set pSpatialReference = pGeometry.SpatialReference
    Set pGraphicElement.SpatialReference = pSpatialReference
    Set pElementProperties = pGroupElement
    pElementProperties.Name = strName
    
    ' ADD GRAPHIC TO GRAPHICS CONTAINER
    If booAlsoAddElementToMapDoc Then pGContainer.AddElement pGroupElement, 0
    
    Set Graphic_ReturnElementFromGeometry3 = pGroupElement
  Else
    pElement.Geometry = pNewGeometry
    Set pGraphicElement = pElement
    Set pSpatialReference = pGeometry.SpatialReference
    Set pGraphicElement.SpatialReference = pSpatialReference
    Set pElementProperties = pElement
    pElementProperties.Name = strName
    
    If Not pSymbol Is Nothing Then
      Select Case pGeometryType
        Case 1:
          pMarkerElement.Symbol = pSymbol
        Case 3, 6, 13, 14, 15, 16:
          pLineElement.Symbol = pSymbol
        Case 4, 11:
          pFillElement.Symbol = pSymbol
        Case 5:
          pFillElement.Symbol = pSymbol
      End Select
    End If
  
    ' ADD GRAPHIC TO GRAPHICS CONTAINER
    
    If booAlsoAddElementToMapDoc Then pGContainer.AddElement pElement, 0
    
    Set Graphic_ReturnElementFromGeometry3 = pElement
    
  End If


  GoTo ClearMemory
ClearMemory:
  Set pMxDocument = Nothing
  Set pActiveView = Nothing
  Set pGContainer = Nothing
  Set pElement = Nothing
  Set pPolygonElement = Nothing
  Set pSpatialReference = Nothing
  Set pGraphicElement = Nothing
  Set pElementProperties = Nothing
  Set pMarkerElement = Nothing
  Set pFillElement = Nothing
  Set pLineElement = Nothing
  Set pClone = Nothing
  Set pNewGeometry = Nothing
  Set pGroupElement = Nothing
  Set pSubElement = Nothing
  Set pPtColl = Nothing
  Set pPt = Nothing

End Function

Public Function ReturnNorthSolidFill(jenMainColor As JenNorthArrowColors, jenOutlineColor As JenNorthArrowColors, _
    dblOutlineWidth As Double, Optional booMakeHollow As Boolean = False) As ISimpleFillSymbol

  Dim pColor As IRgbColor
  Set pColor = New RgbColor
  pColor.RGB = jenMainColor
    
  Dim pOutlineColor As IRgbColor
  Set pOutlineColor = New RgbColor
  pOutlineColor.RGB = jenOutlineColor
  
  Dim pOutline As ILineSymbol
  Set pOutline = New SimpleLineSymbol
  pOutline.Color = pOutlineColor
  pOutline.Width = dblOutlineWidth
  
  Set ReturnNorthSolidFill = New SimpleFillSymbol
  If booMakeHollow Then
    ReturnNorthSolidFill.Style = esriSFSHollow
  Else
    ReturnNorthSolidFill.Color = pColor
  End If
  ReturnNorthSolidFill.Outline = pOutline
  
  Set pColor = Nothing
  Set pOutlineColor = Nothing
  Set pOutline = Nothing

  GoTo ClearMemory
ClearMemory:
  Set pColor = Nothing
  Set pOutlineColor = Nothing
  Set pOutline = Nothing

End Function


Public Function ReturnNorthSolidLineSymbol(jenMainColor As String, dblWidth As Double) As ILineSymbol


  Dim pColor As IRgbColor
  Set pColor = New RgbColor
  pColor.RGB = jenMainColor
    
  Set ReturnNorthSolidLineSymbol = New SimpleLineSymbol
  ReturnNorthSolidLineSymbol.Color = pColor
  ReturnNorthSolidLineSymbol.Width = dblWidth

  GoTo ClearMemory
ClearMemory:
  Set pColor = Nothing

End Function




Public Function ConvertFClassPathToWSData(strPath As String, pWS As IWorkspace, booDirExists As Boolean, _
    strDir As String, strFilename As String, booFileExists As Boolean, booWSOK As Boolean, strErrorDesc As String, _
    Optional lngHWnd As Long = 0) As Boolean

  On err GoTo ErrHandler

'  ' TO RUN:
'  Dim strPath As String
'  Dim pWS As IWorkspace
'  Dim strDir As String
'  Dim strFilename As String
'  Dim booDirExists As Boolean
'  Dim booFileExists As Boolean
'  Dim booWSOK As Boolean
'  Dim booFunctionWorked As Boolean
'  dim strErrorDesc as string
'
'  Debug.Print "-----------------------------"
'  strPath = "D:\arcGIS_stuff\Teaching\Semester_C_2014\Data\Additional_Data\CocRoadStatus_20120425.mdb\ CocRoadStatus_20120425"
'
'  booFunctionWorked = MyGeneralOperations.ConvertFClassPathToWSData( _
'      strPath, pWS, booDirExists, strDir, strFilename, booFileExists, booWSOK, strErrorDesc, 0)
'
'  Debug.Print "Checking '" & strPath & "'"
'  Debug.Print "Function Worked = " & CStr(booFunctionWorked)
'  if not booFunctionWorked then debug.print "----> Error Message:" & vbcrlf & strerrordesc & vbcrlf & "-----------------------------"
'  Debug.Print "Directory = '" & strDir & "'"
'  Debug.Print "Directory Exists = " & CStr(booDirExists)
'  Debug.Print "Filename = '" & strFilename & "'"
'  Debug.Print "Filename Exists = " & CStr(booFileExists)
'  Debug.Print "Created Workspace = " & CStr(Not pWS Is Nothing)
'  If Not pWS Is Nothing Then
'    Debug.Print "  --> Workspace Type = " & CStr(pWS.Type)
'    Debug.Print "  --> Workspace Factory Type = " & pWS.WorkspaceFactory.WorkspaceDescription(False)
'  End If
  
  Set pWS = Nothing
  booWSOK = True
  
  strDir = aml_func_mod.ReturnDir3(strPath, False)
  booDirExists = aml_func_mod.ExistFileDir(strDir)
  
  Dim pWSFact As IWorkspaceFactory
  Dim pWS2 As IWorkspace2
  Dim pFiles As esriSystem.IStringArray
  
  If Not booDirExists Then
    strFilename = ""
    booFileExists = False
  Else
    strFilename = aml_func_mod.ReturnFilename2(strPath)
    If StrComp(Right(strDir, 4), ".mdb", vbTextCompare) = 0 Then  ' Personal Geodatabase
      Set pWSFact = New AccessWorkspaceFactory
      Set pWS = pWSFact.OpenFromFile(strDir, lngHWnd)
      Set pWS2 = pWS
      booFileExists = pWS2.NameExists(esriDTFeatureClass, strFilename)
      
    ElseIf StrComp(Right(strDir, 4), ".gdb", vbTextCompare) = 0 Then  ' File Geodatabase
      Set pWSFact = New FileGDBWorkspaceFactory
      Set pWS = pWSFact.OpenFromFile(strDir, lngHWnd)
      Set pWS2 = pWS
      booFileExists = pWS2.NameExists(esriDTFeatureClass, strFilename)
    
    ElseIf aml_func_mod.PathIsDirectory(strDir) Then                 ' Folder
      Set pFiles = ReturnFilesFromNestedFolders2(strDir, strFilename)
      booFileExists = pFiles.Count > 0
      
      Set pWSFact = New ShapefileWorkspaceFactory
      Set pWS = pWSFact.OpenFromFile(strDir, lngHWnd)
              
    Else
      booWSOK = False
    End If
  End If
  
  ConvertFClassPathToWSData = True
  strErrorDesc = ""
  
  GoTo ClearMemory
  
  Exit Function
ErrHandler:
  ConvertFClassPathToWSData = False
  strErrorDesc = "Error #" & CStr(err.Number) & vbCrLf & "Error Source = " & CStr(err.Source) & vbCrLf & "Description: " & err.Description
  
ClearMemory:
  Set pWSFact = Nothing
  Set pWS2 = Nothing
  Set pFiles = Nothing

End Function


Public Function ReturnEsriFieldTypeNameFromNumber_Friendly(lngFieldType As esriFieldType) As String

  Select Case lngFieldType
    Case esriFieldTypeSmallInteger
      ReturnEsriFieldTypeNameFromNumber_Friendly = "Small Integer"
    Case esriFieldTypeInteger
      ReturnEsriFieldTypeNameFromNumber_Friendly = "Integer"
    Case esriFieldTypeSingle
      ReturnEsriFieldTypeNameFromNumber_Friendly = "Single"
    Case esriFieldTypeDouble
      ReturnEsriFieldTypeNameFromNumber_Friendly = "Double"
    Case esriFieldTypeString
      ReturnEsriFieldTypeNameFromNumber_Friendly = "String"
    Case esriFieldTypeDate
      ReturnEsriFieldTypeNameFromNumber_Friendly = "Date"
    Case esriFieldTypeOID
      ReturnEsriFieldTypeNameFromNumber_Friendly = "OID [Object ID]"
    Case esriFieldTypeGeometry
      ReturnEsriFieldTypeNameFromNumber_Friendly = "Geometry"
    Case esriFieldTypeBlob
      ReturnEsriFieldTypeNameFromNumber_Friendly = "Blob"
    Case esriFieldTypeRaster
      ReturnEsriFieldTypeNameFromNumber_Friendly = "Raster"
    Case esriFieldTypeGUID
      ReturnEsriFieldTypeNameFromNumber_Friendly = "GUID"
    Case esriFieldTypeGlobalID
      ReturnEsriFieldTypeNameFromNumber_Friendly = "Global ID"
    Case esriFieldTypeXML
      ReturnEsriFieldTypeNameFromNumber_Friendly = "XML"
  End Select

End Function


Public Function WriteTextFile_SkipError(strFilename As String, strText As String, Optional booForceOverwrite As Boolean = False) As Boolean

  On Error Resume Next
' INTENDED TO BE USED WHEN ErrorHandling FUNCTION WRITES A TEXT FILE, AND A CRASH IN THIS FUNCTION CAUSES AN INFINITE LOOP
  
  Dim lngFileNumber As Long
  
  If Dir(strFilename) = "" Or booForceOverwrite Then
    lngFileNumber = FreeFile(0)
    
    Open strFilename For Output As #lngFileNumber

    Print #lngFileNumber, strText
    Close #lngFileNumber
    WriteTextFile_SkipError = True
  Else
    ' CONFIRM WHETHER TO OVERWRITE FILE
    Dim lngVBResult As VbMsgBoxResult
    
    lngVBResult = MsgBox("File Already Exists!" & vbCrLf & vbCrLf & "  --> " & strFilename & vbCrLf & vbCrLf & _
        "Click 'OK' to overwrite the file, or 'CANCEL' to quit...", vbOKCancel, "File Exists:")
    If lngVBResult = vbOK Then
      Kill strFilename
      If Dir(strFilename) <> "" Then
        MsgBox "Unable to delete " & strFilename & vbCrLf & vbCrLf & _
          "It may be open in another application.  Please delete this file manually or save the text to a new filename.", , "Unable to Delete File:"
        WriteTextFile_SkipError = False
      Else
      
        lngFileNumber = FreeFile(0)
        
        Open strFilename For Output As #lngFileNumber
    
        Print #lngFileNumber, strText
        Close #lngFileNumber
        WriteTextFile_SkipError = True
      End If
    Else
      WriteTextFile_SkipError = False
    End If
  End If

End Function

Public Sub ImportASCIIToFileGDB(strFilename As String, pWS As IFeatureWorkspace, strDelimChar As String)


' SAMPLE CODE
'  Dim strFileName As String
'  strFileName = "D:\arcGIS_stuff\consultation\Jut_Wynne\Analysis_Files\October_16_2015_Outputs\Mojave_Day_Raster_Data_Oct_16_2015.csv"
'
'  Dim pWS As IFeatureWorkspace
'  Dim pWSFact As IWorkspaceFactory
'  Set pWSFact = New FileGDBWorkspaceFactory
'  Set pWS = pWSFact.OpenFromFile("D:\arcGIS_stuff\consultation\Jut_Wynne\Analysis_Files\October_16_2015_Outputs\test.gdb", 0)
'
'  Call ImportASCIIToFileGDB(strFileName, pWS, ",")
' =========================================================================================
'
' SAMPLE CODE:  BATCH
'
'  Dim pFilenames As esriSystem.IStringArray
'  Set pFilenames = MyGeneralOperations.ReturnFilesFromNestedFolders2("S:\Jeff\MySQL_CSV_Files", "Sep_7")
'
'  Dim pWS As IFeatureWorkspace
'  Dim pWSFact As IWorkspaceFactory
'  Set pWSFact = New FileGDBWorkspaceFactory
'  Set pWS = pWSFact.OpenFromFile("S:\Jeff\SDS_Analysis\MySQL_Sep_7_2015.gdb", 0)
'
'  Dim lngIndex As Long
'  Dim strFile As String
'  For lngIndex = 0 To pFilenames.Count - 1
'    strFile = aml_func_mod.ReturnFilename2(pFilenames.Element(lngIndex))
''    Debug.Print CStr(lngIndex + 1) & ": Importing '" & strFile & "'"
'    Debug.Print CStr(lngIndex + 1) & ": " & Replace(strFile, "_Sep_7_2015.csv", "", , , vbTextCompare)
''    Call ImportASCIIToFileGDB(strFile, pWS, ",")
'  Next lngIndex
'
'
'  Set pFilenames = Nothing
' ===============================================================================================

'  Dim strFilename As String
  Dim strFile As String
  Dim bytFile() As Byte
  Dim booInQuote As Boolean
  
'  strFilename = "S:\Jeff\SDS_Analysis\SDS_Analysis.txt"
    
'  Dim pWS As IFeatureWorkspace
'  Dim pWSFact As IWorkspaceFactory
'  Set pWSFact = New FileGDBWorkspaceFactory
'  Set pWS = pWSFact.OpenFromFile("S:\Jeff\SDS_Analysis\MySQL_Aug_17_2015.gdb", 0)
  
  Dim strNewFile As String
  strNewFile = aml_func_mod.ClipExtension(aml_func_mod.ReturnFilename(strFilename))
  bytFile = ReadFile2(strFilename)
  
  Dim strUBound As String
  strUBound = Format(UBound(bytFile), "#,##0")
  
  Dim lngTab As Long
  Dim bytVal As Byte
'  lngTab = 9
  lngTab = Asc(strDelimChar)
  
  Dim bytQuote As Long
  bytQuote = CByte(Asc(""""))
  
  Dim lngStart As Long
  Dim lngIndex As Long
  Dim strWord As String
  Dim lngWordCount As Long
  lngWordCount = -1
  lngStart = 0
  
  Dim strFirstRow() As String
  Debug.Print "Getting Field Names..."
  
  For lngIndex = 0 To UBound(bytFile)
  
    If lngIndex Mod 500000 = 0 Then
      Debug.Print "  --> " & Format(lngIndex, "#,##0") & " of " & strUBound
    End If
    bytVal = bytFile(lngIndex)
    
    
    If bytVal = bytQuote Then
      booInQuote = Not booInQuote
    End If
    
    If Not booInQuote Then
      If bytVal = lngTab Then
        lngWordCount = lngWordCount + 1
        ReDim Preserve strFirstRow(lngWordCount)
        
        If (Left(strWord, 3) = Chr(34) & Chr(34) & Chr(34)) And (Right(strWord, 3) = Chr(34) & Chr(34) & Chr(34)) Then
          strWord = Left(strWord, Len(strWord) - 2)
          strWord = Right(strWord, Len(strWord) - 2)
        ElseIf (Left(strWord, 1) = Chr(34)) And (Right(strWord, 1) = Chr(34)) Then
          strWord = Left(strWord, Len(strWord) - 1)
          strWord = Right(strWord, Len(strWord) - 1)
        End If
        
        strFirstRow(lngWordCount) = strWord
  '      Debug.Print CStr(lngStart) & " to " & CStr(lngIndex) & ": " & strWord
        strWord = ""
        lngStart = lngIndex + 1
        
      ElseIf bytVal = 13 Then  ' new line
        lngWordCount = lngWordCount + 1
        ReDim Preserve strFirstRow(lngWordCount)
        
        If (Left(strWord, 3) = Chr(34) & Chr(34) & Chr(34)) And (Right(strWord, 3) = Chr(34) & Chr(34) & Chr(34)) Then
          strWord = Left(strWord, Len(strWord) - 2)
          strWord = Right(strWord, Len(strWord) - 2)
        ElseIf (Left(strWord, 1) = Chr(34)) And (Right(strWord, 1) = Chr(34)) Then
          strWord = Left(strWord, Len(strWord) - 1)
          strWord = Right(strWord, Len(strWord) - 1)
        End If
        
        strFirstRow(lngWordCount) = strWord
        Exit For
        
      ElseIf bytVal = 10 Then ' carriage return; ignore
        ' Ignore in this case
        
      Else
        strWord = strWord & Chr(bytVal)
        
      End If
      
    Else
      strWord = strWord & Chr(bytVal)
    End If
    
  Next lngIndex

'  For lngIndex = 0 To UBound(strFirstRow)
'    strWord = strFirstRow(lngIndex)
'    Debug.Print CStr(lngIndex) & "] " & strWord & "  [Text OK = " & CStr(CheckIsAlphanumeric(strWord)) & "]"
'  Next lngIndex
  
  Debug.Print "Parsing into string array..."
  Dim strFinalArray() As String
  Dim lngRowCount As Long
  Dim lngFieldCount As Long
  lngWordCount = -1
  
  lngFieldCount = UBound(strFirstRow)
  lngRowCount = 0
  ReDim strFinalArray(lngFieldCount, lngRowCount)
  
  For lngIndex = 0 To UBound(bytFile)
  
    If lngIndex Mod 1000000 = 0 Then
      Debug.Print "  --> " & Format(lngIndex, "#,##0") & " of " & strUBound
      DoEvents
    End If
    bytVal = bytFile(lngIndex)
    
    If bytVal = bytQuote Then
      booInQuote = Not booInQuote
    End If
    
    If Not booInQuote Then
      If bytVal = lngTab Then
        lngWordCount = lngWordCount + 1
        
        If (Left(strWord, 3) = Chr(34) & Chr(34) & Chr(34)) And (Right(strWord, 3) = Chr(34) & Chr(34) & Chr(34)) Then
          strWord = Left(strWord, Len(strWord) - 2)
          strWord = Right(strWord, Len(strWord) - 2)
        ElseIf (Left(strWord, 1) = Chr(34)) And (Right(strWord, 1) = Chr(34)) Then
          strWord = Left(strWord, Len(strWord) - 1)
          strWord = Right(strWord, Len(strWord) - 1)
        End If
        
        strFinalArray(lngWordCount, lngRowCount) = strWord
        strWord = ""
        
      ElseIf bytVal = 13 Then  ' new line
        lngWordCount = lngWordCount + 1
        
        If (Left(strWord, 3) = Chr(34) & Chr(34) & Chr(34)) And (Right(strWord, 3) = Chr(34) & Chr(34) & Chr(34)) Then
          strWord = Left(strWord, Len(strWord) - 2)
          strWord = Right(strWord, Len(strWord) - 2)
        ElseIf (Left(strWord, 1) = Chr(34)) And (Right(strWord, 1) = Chr(34)) Then
          strWord = Left(strWord, Len(strWord) - 1)
          strWord = Right(strWord, Len(strWord) - 1)
        End If
        
        strFinalArray(lngWordCount, lngRowCount) = strWord
        strWord = ""
        lngWordCount = -1
        lngRowCount = lngRowCount + 1
        ReDim Preserve strFinalArray(lngFieldCount, lngRowCount)
        
      ElseIf bytVal = 10 Then ' carriage return; ignore
        ' Ignore in this case
        
      Else
        strWord = strWord & Chr(bytVal)
        
      End If
      
    Else
      strWord = strWord & Chr(bytVal)
    End If
    
      
  Next lngIndex
  
  Debug.Print "Checking Field Lengths and Types..."
  
  Dim strVal As String
  Dim lngIndex2 As Long
  Dim lngIndexLengths() As Long
  Dim booIsDouble() As Boolean
  Dim booIsLong() As Boolean
  Dim booIsDate() As Boolean
  Dim booFoundSomething() As Boolean
    
  ReDim lngIndexLengths(lngFieldCount)
  ReDim booIsDouble(lngFieldCount)
  ReDim booIsLong(lngFieldCount)
  ReDim booIsDate(lngFieldCount)
  ReDim booFoundSomething(lngFieldCount)
  
  strUBound = Format(UBound(strFinalArray, 2))
  
  For lngIndex = 0 To lngFieldCount
    lngIndexLengths(lngIndex) = 0
    booIsDouble(lngIndex) = True
    booIsLong(lngIndex) = True
    booIsDate(lngIndex) = True
    booFoundSomething(lngIndex) = False
  Next lngIndex
  
  For lngIndex = 1 To UBound(strFinalArray, 2)
    If lngIndex Mod 1000 = 0 Then
      Debug.Print "  --> " & Format(lngIndex, "#,##0") & " of " & strUBound
      DoEvents
    End If
    
    For lngIndex2 = 0 To lngFieldCount
      strVal = Trim(strFinalArray(lngIndex2, lngIndex))
      
'      If lngIndex2 = 0 And Len(strVal) > 13 Then
'        Debug.Print "Here..."
'      End If
      
      lngIndexLengths(lngIndex2) = MyGeometricOperations.MaxLong(lngIndexLengths(lngIndex2), Len(strVal))
      If strVal <> "" Then
        booFoundSomething(lngIndex2) = True
        If IsNumeric(strVal) Then
          If InStr(1, strVal, ".", vbTextCompare) > 0 Then
            booIsLong(lngIndex2) = False
          End If
        Else
          booIsDouble(lngIndex2) = False
          booIsLong(lngIndex2) = False
        End If
        If Not IsDate(strVal) Then booIsDate(lngIndex2) = False
      End If
    Next lngIndex2
  Next lngIndex
  
  Dim pFieldArray As esriSystem.IVariantArray
  Set pFieldArray = New esriSystem.varArray
  Dim pField As iField
  Dim pFieldEdit As IFieldEdit
  
  Dim strFieldReport As String
  For lngIndex = 0 To lngFieldCount
    strVal = strFirstRow(lngIndex)
    If booIsLong(lngIndex) = True Then booIsDouble(lngIndex) = False
    strFieldReport = strFieldReport & CStr(lngIndex + 1) & "] " & strVal & vbCrLf & _
        "  --> Max Length = " & CStr(lngIndexLengths(lngIndex)) & vbCrLf & _
        "  --> Is Long = " & CStr(booIsLong(lngIndex)) & vbCrLf & _
        "  --> Is Double = " & CStr(booIsDouble(lngIndex)) & vbCrLf & _
        "  --> Is Date = " & CStr(booIsDate(lngIndex)) & vbCrLf & _
        "  --> Found Any Value = " & CStr(booFoundSomething(lngIndex)) & vbCrLf & _
        "  --> Is String = " & CStr(Not booIsDate(lngIndex) And Not booIsLong(lngIndex) And Not booIsDouble(lngIndex)) & vbCrLf
      
    Set pField = New Field
    Set pFieldEdit = pField
    With pFieldEdit
      .Name = ReturnValidFGDBFieldName(ReturnTitleCase(strVal))
      If booFoundSomething(lngIndex) = False Then
        .Type = esriFieldTypeString
        .length = 5
      ElseIf booIsLong(lngIndex) Then
        .Type = esriFieldTypeInteger
      ElseIf booIsDouble(lngIndex) Then
        .Type = esriFieldTypeDouble
      ElseIf booIsDate(lngIndex) Then
        .Type = esriFieldTypeDate
      Else
        .Type = esriFieldTypeString
        .length = lngIndexLengths(lngIndex)
      End If
    End With
    pFieldArray.Add pField
      
  Next lngIndex
  
  Dim strTableName As String
  strTableName = MakeUniqueGDBTableName(pWS, _
      aml_func_mod.ClipExtension(aml_func_mod.ReturnFilename(strNewFile)))
  
  Dim pNewTable As ITable
  Set pNewTable = CreateGDBTable(pWS, strTableName, pFieldArray)
  
  Dim varFieldNameIndices() As Variant
  ReDim varFieldNameIndices(1, pFieldArray.Count - 1)
  For lngIndex = 0 To pFieldArray.Count - 1
    Set pField = pFieldArray.Element(lngIndex)
    varFieldNameIndices(0, lngIndex) = pField.Name
    varFieldNameIndices(1, lngIndex) = pNewTable.FindField(pField.Name)
  Next lngIndex
  
  Dim lngFieldIndex As Long
  Dim pBuffer As IRowBuffer
  Dim pCursor As ICursor
  Set pCursor = pNewTable.Insert(True)
  Set pBuffer = pNewTable.CreateRowBuffer
  Dim strFieldName As String
  Dim booFoundAValue As Boolean
  
'  For lngIndex = 1 To UBound(strSplit)
'    If lngIndex Mod 500 = 0 Then
'      pCursor.Flush
'      Debug.Print "Writing row " & CStr(lngIndex)
'    End If
'    strLine = Trim(strSplit(lngIndex))
'
'    If lngIndex >= 22 Then
''      Debug.Print "Here..."
'    End If
    
  For lngIndex = 1 To UBound(strFinalArray, 2)
    If lngIndex Mod 1000 = 0 Then
      Debug.Print "  --> " & Format(lngIndex, "#,##0") & " of " & strUBound
      DoEvents
    End If
    
    booFoundAValue = False
    For lngIndex2 = 0 To lngFieldCount
      strVal = Trim(strFinalArray(lngIndex2, lngIndex))
      
      If strVal <> "" Then booFoundAValue = True
      
      strFieldName = CStr(varFieldNameIndices(0, lngIndex2))
      lngFieldIndex = CLng(varFieldNameIndices(1, lngIndex2))
      
'      Debug.Print "  --> Adding '" & strVal & "' to '" & strFieldName
      
      If booFoundSomething(lngIndex2) = False Then
        pBuffer.Value(lngFieldIndex) = ""
      
      ElseIf booIsLong(lngIndex2) Then
        If strVal = "" Then
          pBuffer.Value(lngFieldIndex) = Null
        Else
          pBuffer.Value(lngFieldIndex) = CLng(strVal)
        End If
      ElseIf booIsDouble(lngIndex2) Then
        If strVal = "" Then
          pBuffer.Value(lngFieldIndex) = Null
        Else
          pBuffer.Value(lngFieldIndex) = CDbl(strVal)
        End If
      ElseIf booIsDate(lngIndex2) Then
        If strVal = "" Then
          pBuffer.Value(lngFieldIndex) = Null
        Else
          pBuffer.Value(lngFieldIndex) = CDate(strVal)
        End If
      Else
        If strVal = "" Then
          pBuffer.Value(lngFieldIndex) = Null
        Else
          pBuffer.Value(lngFieldIndex) = strVal
        End If
      End If
      
    Next lngIndex2
    
    If booFoundAValue Then pCursor.InsertRow pBuffer
    
  Next lngIndex
  pCursor.Flush
  
'  Dim pDataObj As MSForms.DataObject
'  Set pDataObj = New MSForms.DataObject
'  pDataObj.SetText strFieldReport
'  pDataObj.PutInClipboard
'
'  Set pDataObj = Nothing
  
  
  Debug.Print "Done..."
  
  GoTo ClearMemory
ClearMemory:
  Erase bytFile
  Erase strFirstRow
  Erase strFinalArray
  Erase lngIndexLengths
  Erase booIsDouble
  Erase booIsLong
  Erase booIsDate
  Erase booFoundSomething
  Set pFieldArray = Nothing
  Set pField = Nothing
  Set pFieldEdit = Nothing
  Set pNewTable = Nothing
  Erase varFieldNameIndices
  Set pBuffer = Nothing
  Set pCursor = Nothing




End Sub
Private Function ReadFile2(strFilename As String) As Byte()

    Dim lngFile As Integer

    lngFile = FreeFile
    Open strFilename For Binary Access Read As #lngFile
    If LOF(lngFile) > 0 Then
        ReDim ReadFile2(0 To LOF(lngFile) - 1)
        Get lngFile, , ReadFile2
    End If
    Close lngFile

End Function

Public Function CreateFieldAttributeIndex(pTable As ITable, strFieldName As String, _
    Optional strFailReason As String) As Boolean
  
  Dim pDataset As IDataset
  Dim strDatasetName As String
  Set pDataset = pTable
  strDatasetName = pDataset.BrowseName
  
  CreateFieldAttributeIndex = True
  strFailReason = ""
  
  Dim pIndexes As IIndexes
  Set pIndexes = pTable.Indexes
  
  Dim pEnumIndex As IEnumIndex
  Set pEnumIndex = pIndexes.FindIndexesByFieldName(strFieldName)
  
  Dim pFields As IFields
  Dim pField As iField
  Dim pFieldsEdit As IFieldsEdit
  Dim pIndex As IIndex
  Dim pIndexEdit As IIndexEdit
  
  Set pIndex = pEnumIndex.Next
  
  If pIndex Is Nothing Then
    
'    Set pField = pTable.Fields.Field(pTable.FindField(strFieldName))
    If pTable.FindField(strFieldName) = -1 Then
      CreateFieldAttributeIndex = False
      strFailReason = "No Field with name '" & strFieldName & "' [" & strDatasetName & "]"
      Debug.Print "  --> " & strFailReason
    Else
      Set pField = pTable.Fields.Field(pTable.FindField(strFieldName))
      ' Create a fields collection and add the specified field to it.
      Set pFields = New Fields
      Set pFieldsEdit = pFields
      pFieldsEdit.FieldCount = 1
      Set pFieldsEdit.Field(0) = pField
      
      ' Create an index and cast to the IIndexEdit interface.
      Set pIndex = New Index
      Set pIndexEdit = pIndex
      
      ' Set the index's properties, including the associated fields.
      Set pIndexEdit.Fields = pFields
      pIndexEdit.IsAscending = False
      pIndexEdit.IsUnique = False
      pIndexEdit.Name = strFieldName
      
      ' Add the index to the feature class.
      pTable.AddIndex pIndex
      
      Debug.Print "  --> Built Index for '" & strFieldName & "' [" & strDatasetName & "]"
      
    End If
  Else
    CreateFieldAttributeIndex = False
    strFailReason = "Field '" & strFieldName & "' already indexed"
'    Debug.Print "  --> " & strFailReason
  End If
  

ClearMemory:
  Set pIndexes = Nothing
  Set pEnumIndex = Nothing
  Set pFields = Nothing
  Set pField = Nothing
  Set pFieldsEdit = Nothing
  Set pIndex = Nothing
  Set pIndexEdit = Nothing

End Function

Public Function ReturnValidFGDBFieldName(strName As String) As String
  
  Dim strAlphaNumeric As String
  strAlphaNumeric = "0123456789_abcdefghijklmnopqrstuvwxyz"
  
  Dim strChar As String
  Dim lngIndex As Long
  Dim strNewName As String
  
  For lngIndex = 1 To Len(strName)
    strChar = Mid(strName, lngIndex, 1)
    If InStr(1, strAlphaNumeric, strChar, vbTextCompare) > 0 Then
      strNewName = strNewName & strChar
    Else
      strNewName = strNewName & "_"
    End If
  Next lngIndex
  strChar = Left(strNewName, 1)
  If IsNumeric(strChar) Then
    strNewName = "z" & strNewName
  End If
  
  Dim strReserved() As String
  ReDim strReserved(30)
  strReserved(0) = "Add"
  strReserved(1) = "ALTER"
  strReserved(2) = "AND"
  strReserved(3) = "AS"
  strReserved(4) = "Asc"
  strReserved(5) = "BETWEEN"
  strReserved(6) = "BY"
  strReserved(7) = "Column"
  strReserved(8) = "Create"
  strReserved(9) = "Date"
  strReserved(10) = "DELETE"
  strReserved(11) = "DESC"
  strReserved(12) = "DROP"
  strReserved(13) = "Exists"
  strReserved(14) = "FOR"
  strReserved(15) = "From"
  strReserved(16) = "IN"
  strReserved(17) = "Insert"
  strReserved(18) = "INTO"
  strReserved(19) = "IS"
  strReserved(20) = "LIKE"
  strReserved(21) = "NOT"
  strReserved(22) = "Null"
  strReserved(23) = "OR"
  strReserved(24) = "Order"
  strReserved(25) = "SELECT"
  strReserved(26) = "SET"
  strReserved(27) = "Table"
  strReserved(28) = "Update"
  strReserved(29) = "Values"
  strReserved(30) = "WHERE"
  
  For lngIndex = 0 To UBound(strReserved)
    If StrComp(strNewName, strReserved(lngIndex), vbTextCompare) = 0 Then
      strNewName = strNewName & "_"
    End If
  Next lngIndex
  
  ReturnValidFGDBFieldName = strNewName
  Debug.Print "  --> Started with '" & strName & "', ended with '" & ReturnValidFGDBFieldName & "'"

ClearMemory:
  Erase strReserved

End Function

Function IsDimmed(Arr As Variant) As Boolean
  ' Adapted from http://www.vbforums.com/showthread.php?736285-VB6-Returning-Detecting-Empty-Arrays
  On Error GoTo ReturnFalse
  IsDimmed = UBound(Arr) >= LBound(Arr)
  Exit Function
ReturnFalse:
  IsDimmed = False
End Function
Public Function ReturnLongValFromRow(pRow As IRow, lngIndex As Long, Optional booIsNull As Boolean) As Long

  Dim varVal As Variant
  varVal = pRow.Value(lngIndex)
  If IsNull(varVal) Then
    ReturnLongValFromRow = -999
    booIsNull = True
'    MsgBox "Found a null ID value!" & vbCrLf & "pRow.OID = " & CStr(pRow.OID)
  Else
    ReturnLongValFromRow = CLng(varVal)
    booIsNull = False
  End If
  
ClearMemory:
  varVal = Null

End Function
Public Function ReturnDoubleValFromRow(pRow As IRow, lngIndex As Long, Optional booIsNull As Boolean) As Double

  Dim varVal As Variant
  varVal = pRow.Value(lngIndex)
  If IsNull(varVal) Then
    ReturnDoubleValFromRow = -999
    booIsNull = True
'    MsgBox "Found a null ID value!" & vbCrLf & "pRow.OID = " & CStr(pRow.OID)
  Else
    ReturnDoubleValFromRow = CDbl(varVal)
    booIsNull = False
  End If
  
ClearMemory:
  varVal = Null

End Function
Public Function ReturnDateValFromRow(pRow As IRow, lngIndex As Long, Optional booIsNull As Boolean) As Date


  Dim varVal As Variant
  varVal = pRow.Value(lngIndex)
  If IsNull(varVal) Then
    ReturnDateValFromRow = CDate("1/1/1900")
    booIsNull = True
'    MsgBox "Found a null ID value!" & vbCrLf & "pRow.OID = " & CStr(pRow.OID)
  Else
    ReturnDateValFromRow = CDate(varVal)
    booIsNull = False
  End If
  
ClearMemory:
  varVal = Null

End Function
Public Function ReturnStringValFromRow(pRow As IRow, lngIndex As Long, Optional booIsNull As Boolean) As String

  Dim varVal As Variant
  varVal = pRow.Value(lngIndex)
  If IsNull(varVal) Then
    ReturnStringValFromRow = ""
    booIsNull = True
  Else
    ReturnStringValFromRow = Trim(CStr(varVal))
    booIsNull = False
  End If
  
  ReturnStringValFromRow = Replace(ReturnStringValFromRow, "�", "", , , vbTextCompare)
  If Left(ReturnStringValFromRow, 1) = """" And Right(ReturnStringValFromRow, 1) = """" Then
    ReturnStringValFromRow = Left(ReturnStringValFromRow, Len(ReturnStringValFromRow) - 1)
    ReturnStringValFromRow = Right(ReturnStringValFromRow, Len(ReturnStringValFromRow) - 1)
  End If
  
  ReturnStringValFromRow = Replace(ReturnStringValFromRow, "&apos;", "'", , , vbTextCompare)
  
ClearMemory:
  varVal = Null

End Function
Public Sub DateDiffByDayMonthYear(datFirstDate As Date, datSecondDate As Date, Optional lngDays As Long, _
    Optional lngMonths As Long, Optional lngYears As Long, Optional booAbsoluteValueIfNegative As Boolean = True)
    
  ' ADAPTED FROM http://www.vbforums.com/showthread.php?675573-RESOLVED-calculate-age
    
  Dim tmpDate As Date
  Dim tmpDate1 As Date
  Dim tmpDate2 As Date
  tmpDate1 = datFirstDate
  tmpDate2 = datSecondDate
  Dim booSwitch As Boolean
  
  If tmpDate2 <= tmpDate1 Then
    booSwitch = True
    tmpDate = tmpDate1
    tmpDate1 = tmpDate2
    tmpDate2 = tmpDate
'    MsgBox "Invalid date"
'    Exit Sub
  End If
  
  lngYears = DateDiff("yyyy", tmpDate1, tmpDate2) 'year diff between DOB and current date
  tmpDate = DateAdd("yyyy", lngYears, tmpDate1) 'add the year diff to DOB
  
  If tmpDate > tmpDate2 Then
    lngYears = lngYears - 1
    tmpDate = DateAdd("yyyy", -1, tmpDate)
  End If
  
  lngMonths = DateDiff("m", tmpDate, tmpDate2) 'month diff between DOB and current date
  tmpDate = DateAdd("m", lngMonths, tmpDate) 'add the month diff to DOB
  
  If tmpDate > tmpDate2 Then
    lngMonths = lngMonths - 1
    tmpDate = DateAdd("m", -1, tmpDate)
  End If
  
  lngDays = DateDiff("d", tmpDate, tmpDate2)
  
  If booSwitch And Not booAbsoluteValueIfNegative Then lngYears = -lngYears

End Sub

Public Function ReturnValidFGDBFieldName2(strName As String, pFieldSet As IUnknown) As String
  
  Dim strAlphaNumeric As String
  strAlphaNumeric = "0123456789_abcdefghijklmnopqrstuvwxyz"
  
  Dim strChar As String
  Dim lngIndex As Long
  Dim strNewName As String
  
  For lngIndex = 1 To Len(strName)
    strChar = Mid(strName, lngIndex, 1)
    If InStr(1, strAlphaNumeric, strChar, vbTextCompare) > 0 Then
      strNewName = strNewName & strChar
    Else
      strNewName = strNewName & "_"
    End If
  Next lngIndex
  strChar = Left(strNewName, 1)
  If IsNumeric(strChar) Then
    strNewName = "z" & strNewName
  End If
  
  Dim strReserved() As String
  ReDim strReserved(30)
  strReserved(0) = "Add"
  strReserved(1) = "ALTER"
  strReserved(2) = "AND"
  strReserved(3) = "AS"
  strReserved(4) = "Asc"
  strReserved(5) = "BETWEEN"
  strReserved(6) = "BY"
  strReserved(7) = "Column"
  strReserved(8) = "Create"
  strReserved(9) = "Date"
  strReserved(10) = "DELETE"
  strReserved(11) = "DESC"
  strReserved(12) = "DROP"
  strReserved(13) = "Exists"
  strReserved(14) = "FOR"
  strReserved(15) = "From"
  strReserved(16) = "IN"
  strReserved(17) = "Insert"
  strReserved(18) = "INTO"
  strReserved(19) = "IS"
  strReserved(20) = "LIKE"
  strReserved(21) = "NOT"
  strReserved(22) = "Null"
  strReserved(23) = "OR"
  strReserved(24) = "Order"
  strReserved(25) = "SELECT"
  strReserved(26) = "SET"
  strReserved(27) = "Table"
  strReserved(28) = "Update"
  strReserved(29) = "Values"
  strReserved(30) = "WHERE"
  
  For lngIndex = 0 To UBound(strReserved)
    If StrComp(strNewName, strReserved(lngIndex), vbTextCompare) = 0 Then
      strNewName = strNewName & "_"
    End If
  Next lngIndex
  
  Dim lngMaxLength As Long
  lngMaxLength = 64
  strNewName = Left(strNewName, lngMaxLength)
  
  Dim strTestName As String
    
  ' MAKE SURE FIELD NAME DOES NOT ALREADY EXIST IN LIST
  ' CONVERT OBJECT INTO LIST OF EXISTING NAMES
  Dim pFieldArray As esriSystem.IStringArray
  Set pFieldArray = New esriSystem.strArray

  If Not pFieldSet Is Nothing Then
  
    If TypeOf pFieldSet Is IFields Then   ' <--------------------
      
      Dim pFields As IFields
      Set pFields = pFieldSet
      If pFields.FieldCount > 0 Then
        For lngIndex = 0 To pFields.FieldCount - 1
          strTestName = pFields.Field(lngIndex).Name
          pFieldArray.Add strTestName
        Next lngIndex
      End If
      
    ElseIf TypeOf pFieldSet Is esriSystem.IVariantArray Then
      
      Dim pVar As Variant
      Dim pVarArray As esriSystem.IVariantArray
      Dim pField As iField
      Set pVarArray = pFieldSet
      If pVarArray.Count > 0 Then
        For lngIndex = 0 To pVarArray.Count - 1
          If VarType(pVarArray.Element(lngIndex)) = vbDataObject Then           ' IS AN IField
            Set pField = pVarArray.Element(lngIndex)
            strTestName = pField.Name
          Else
            strTestName = pVarArray.Element(lngIndex)
          End If
          pFieldArray.Add strTestName
        Next lngIndex
      End If
    
    ElseIf TypeOf pFieldSet Is esriSystem.IStringArray Then
      
      Set pFieldArray = pFieldSet
    
    End If
  End If
  
  
  ' CHECK CURRENT NAME AGAINST LIST
  If pFieldArray.Count > 0 Then
    Dim lngCounter As Long
    Dim strBaseName As String
    strBaseName = strNewName
    Dim booFoundConflict As Boolean
    booFoundConflict = True
    Do Until booFoundConflict = False
      booFoundConflict = False
      For lngIndex = 0 To pFieldArray.Count - 1
        strTestName = pFieldArray.Element(lngIndex)
        If StrComp(strNewName, strTestName, vbTextCompare) = 0 Then
          booFoundConflict = True
          Exit For
        End If
      Next lngIndex
      If booFoundConflict Then
        lngCounter = lngCounter + 1
        strNewName = Left(strBaseName, lngMaxLength - Len(CStr(lngCounter))) & CStr(lngCounter)
      End If
    Loop
  End If
    
  
  
  ReturnValidFGDBFieldName2 = strNewName
'  Debug.Print "  --> Started with '" & strName & "', ended with '" & ReturnValidFGDBFieldName & "'"

  GoTo ClearMemory
ClearMemory:
  Erase strReserved
  Set pFieldArray = Nothing
  Set pFields = Nothing
  pVar = Null
  Set pVarArray = Nothing
  Set pField = Nothing

End Function
Public Function ReturnAcceptableFieldName2(strOrigName As String, pFieldSet As IUnknown, Optional booRestrictToDBase As Boolean = False, _
    Optional booRestrictToPersonalGDB As Boolean = False, Optional booRestrictToCoverage As Boolean = False, _
    Optional booRestrictToFGDB As Boolean = False) As String
    
    Dim strName As String
    strName = strOrigName

  If booRestrictToDBase Then
    ReturnAcceptableFieldName2 = ReturnDBASEFieldName(strName, pFieldSet)
  ElseIf booRestrictToFGDB Then
    ReturnAcceptableFieldName2 = ReturnValidFGDBFieldName2(strName, pFieldSet)
  
  Else
    Dim lngIndex As Long
    
    Dim strAcceptable As String
    strAcceptable = "abcdefghijklmnopqrstuvwxyz1234567890_"
    Dim strCharacters As String
    strCharacters = "abcdefghijklmnopqrstuvwxyz"
    
    Dim lngMaxLength As Long
    If booRestrictToCoverage Then
      lngMaxLength = 16
    ElseIf booRestrictToPersonalGDB Then
      lngMaxLength = 52
    ElseIf booRestrictToFGDB Then
      lngMaxLength = 64
    Else
      lngMaxLength = 64
    End If
    
    Dim pField As iField
    
    Dim strNewName As String
    Dim strChar As String
    
    ' MAKE SURE SUGGESTED NAME IS VALID FOR dBASE IN GENERAL
    ' CHECK IF FIELD NAME IS EMPTY STRING
    If strName = "" Then strName = "z_Field"
    If StrComp(strName, "date", vbTextCompare) = 0 Then strName = "z_Date"
    If StrComp(strName, "day", vbTextCompare) = 0 Then strName = "z_Day"
    If StrComp(strName, "month", vbTextCompare) = 0 Then strName = "z_Month"
    If StrComp(strName, "table", vbTextCompare) = 0 Then strName = "z_Table"
    If StrComp(strName, "text", vbTextCompare) = 0 Then strName = "z_Text"
    If StrComp(strName, "user", vbTextCompare) = 0 Then strName = "z_User"
    If StrComp(strName, "when", vbTextCompare) = 0 Then strName = "z_When"
    If StrComp(strName, "where", vbTextCompare) = 0 Then strName = "z_Where"
    If StrComp(strName, "year", vbTextCompare) = 0 Then strName = "z_Year"
    If StrComp(strName, "zone", vbTextCompare) = 0 Then strName = "z_Zone"
    If StrComp(strName, "Shape_Length", vbTextCompare) = 0 Then strName = "z_Shape_Length"
    If StrComp(strName, "Shape_Area", vbTextCompare) = 0 Then strName = "z_Shape_Area"
    If StrComp(strName, "OID", vbTextCompare) = 0 Then strName = "z_OID"
    If StrComp(strName, "FID", vbTextCompare) = 0 Then strName = "z_FID"
    If StrComp(strName, "Object_ID", vbTextCompare) = 0 Then strName = "z_Object_ID"
    If StrComp(strName, "Shape", vbTextCompare) = 0 Then strName = "z_Shape"
    If StrComp(strName, "ObjectID", vbTextCompare) = 0 Then strName = "z_ObjectID"
    
    ' CHECK IF FIELD NAME IS PROPER LENGTH
    strName = Left(strName, lngMaxLength)
    
    ' CHECK IF FIELD NAME DOES NOT START WITH A LETTER
    If Not (InStr(1, strCharacters, Left(strName, 1), vbTextCompare) > 0) Then
      strName = Left("z" & strName, lngMaxLength)
    End If
    
    ' CHECK FOR NON_CONFORMING CHARACTERS
    strNewName = ""
    For lngIndex = 1 To Len(strName)
      strChar = Mid(strName, lngIndex, 1)
      If Not (InStr(1, strAcceptable, strChar, vbTextCompare) > 0) Then
        strNewName = strNewName & "_"
      Else
        strNewName = strNewName & strChar
      End If
    Next lngIndex
    strName = strNewName
    
    Dim strTestName As String
      
    ' MAKE SURE FIELD NAME DOES NOT ALREADY EXIST IN LIST
    ' CONVERT OBJECT INTO LIST OF EXISTING NAMES
    Dim pFieldArray As esriSystem.IStringArray
    Set pFieldArray = New esriSystem.strArray
  
    If Not pFieldSet Is Nothing Then
    
      If TypeOf pFieldSet Is IFields Then   ' <--------------------
        
        Dim pFields As IFields
        Set pFields = pFieldSet
        If pFields.FieldCount > 0 Then
          For lngIndex = 0 To pFields.FieldCount - 1
            strTestName = pFields.Field(lngIndex).Name
            pFieldArray.Add strTestName
          Next lngIndex
        End If
        
      ElseIf TypeOf pFieldSet Is esriSystem.IVariantArray Then
        
        Dim pVar As Variant
        Dim pVarArray As esriSystem.IVariantArray
        Set pVarArray = pFieldSet
        If pVarArray.Count > 0 Then
          For lngIndex = 0 To pVarArray.Count - 1
            If VarType(pVarArray.Element(lngIndex)) = vbDataObject Then           ' IS AN IField
              Set pField = pVarArray.Element(lngIndex)
              strTestName = pField.Name
            Else
              strTestName = pVarArray.Element(lngIndex)
            End If
            pFieldArray.Add strTestName
          Next lngIndex
        End If
      
      ElseIf TypeOf pFieldSet Is esriSystem.IStringArray Then
        
        Set pFieldArray = pFieldSet
      
      End If
    End If
    
    ' CHECK CURRENT NAME AGAINST LIST
    If pFieldArray.Count > 0 Then
      Dim lngCounter As Long
      Dim strBaseName As String
      strBaseName = strName
      Dim booFoundConflict As Boolean
      booFoundConflict = True
      Do Until booFoundConflict = False
        booFoundConflict = False
        For lngIndex = 0 To pFieldArray.Count - 1
          strTestName = pFieldArray.Element(lngIndex)
          If StrComp(strName, strTestName, vbTextCompare) = 0 Then
            booFoundConflict = True
            Exit For
          End If
        Next lngIndex
        If booFoundConflict Then
          lngCounter = lngCounter + 1
          strName = Left(strBaseName, lngMaxLength - Len(CStr(lngCounter))) & CStr(lngCounter)
        End If
      Loop
    End If
    
    ReturnAcceptableFieldName2 = strName
  
  End If
  GoTo ClearMemory
ClearMemory:
  Set pField = Nothing
  Set pFieldArray = Nothing
  Set pFields = Nothing
  pVar = Null
  Set pVarArray = Nothing

End Function
Public Function CreateInMemoryFeatureClass_Empty(pTemplateFields As esriSystem.IVariantArray, strName As String, _
    pSpRef As ISpatialReference, lngGeometryType As esriGeometryType, booHasM As Boolean, booHasZ As Boolean) As IFeatureClass
    

    ' create an inmemory featureclass
    ' ADAPTED FROM KIRK KUYKENDALL
    ' http://forums.esri.com/Thread.asp?c=93&f=993&t=210767
            
    Dim pClone As IClone
    Dim pSpRefRes As ISpatialReferenceResolution
    Set pSpRefRes = pSpRef
    pSpRefRes.ConstructFromHorizon
    Dim pWSF As IWorkspaceFactory
    Set pWSF = New InMemoryWorkspaceFactory
    
    Dim pName As IName
    Set pName = pWSF.Create("", "inmemory", Nothing, 0)
    Dim pFWS As IFeatureWorkspace
    Set pFWS = pName.Open
            
    Dim pFields As IFields
    Dim pFieldsEdit As IFieldsEdit
    Dim pField As iField
    Dim pFieldEdit As IFieldEdit
   
    Set pFields = New Fields
    Set pFieldsEdit = pFields
      
    '' create the geometry field
    
    Dim pGeomDef As IGeometryDef
    Dim pGeomDefEdit As IGeometryDefEdit
    Set pGeomDef = New GeometryDef
    Set pGeomDefEdit = pGeomDef
    
'    MsgBox pGeom.SpatialReference.Name
    
    '' assign the geometry definiton properties.
    With pGeomDefEdit
      .GeometryType = lngGeometryType
      .GridCount = 1
      .GridSize(0) = 0
      If lngGeometryType = esriGeometryPoint Then
        .AvgNumPoints = 1
      Else
        .AvgNumPoints = 5
      End If
      .HasM = booHasM
      .HasZ = booHasZ
      Set .SpatialReference = pSpRef
    End With
    
    Set pField = New Field
    Set pFieldEdit = pField
    
    pFieldEdit.Name = "Shape"
    pFieldEdit.AliasName = "geometry"
    pFieldEdit.Type = esriFieldTypeGeometry
    Set pFieldEdit.GeometryDef = pGeomDef
    pFieldsEdit.AddField pField
    

    Dim booAddAttribute As Boolean
    booAddAttribute = Not pTemplateFields Is Nothing
    Dim varVal As Variant
    
    Dim lngIndex As Long
    Dim pTemplateField As iField
    
    Dim pTempField As iField
    If booAddAttribute Then
      For lngIndex = 0 To pTemplateFields.Count - 1
        Set pTemplateField = pTemplateFields.Element(lngIndex)
        Set pClone = pTemplateField
        Set pTempField = pClone.Clone
        pFieldsEdit.AddField pTempField
      Next lngIndex
      ReDim lngIDIndex(pTemplateFields.Count - 1)
    Else
      Set pField = New Field
      Set pFieldEdit = pField
      With pFieldEdit
        .Name = "Unique_ID"
        .Type = esriFieldTypeInteger
      End With
      pFieldsEdit.AddField pField
      ReDim lngIDIndex(0)
    End If
    
    Dim pCLSID As UID
    Set pCLSID = New UID
    pCLSID.Value = "esriGeoDatabase.Feature"
  
    Dim pInMemFC As IFeatureClass
    Set pInMemFC = pFWS.CreateFeatureClass(strName, pFields, _
                             pCLSID, Nothing, esriFTSimple, _
                             "Shape", "")
    
    
    Set CreateInMemoryFeatureClass_Empty = pInMemFC
  
  GoTo ClearMemory
ClearMemory:
  Set pClone = Nothing
  Set pSpRefRes = Nothing
  Set pWSF = Nothing
  Set pName = Nothing
  Set pFWS = Nothing
  Set pFields = Nothing
  Set pFieldsEdit = Nothing
  Set pField = Nothing
  Set pFieldEdit = Nothing
  Set pGeomDef = Nothing
  Set pGeomDefEdit = Nothing
  varVal = Null
  Set pTemplateField = Nothing
  Set pTempField = Nothing
  Set pCLSID = Nothing
  Set pInMemFC = Nothing



End Function

Public Function MakeUniqueShapeFilename2(pWS As IFeatureWorkspace, strFClassName As String, booIncludeDotSHP As Boolean) As String

  Dim strTestName As String
  strTestName = strFClassName
  If StrComp(Right(strTestName, 4), ".shp", vbTextCompare) = 0 Then
    strTestName = Left(strTestName, Len(strTestName) - 4)
  End If
  
  Dim strNames() As String
  Dim lngCount As Long
  Dim strBaseName As String
  Dim lngCounter As Long
  Dim strNewName As String
  strNames = ReturnStringArrayOfNames(pWS, esriDTFeatureClass, lngCount)
  If lngCount = 0 Then
    MakeUniqueShapeFilename2 = strTestName
  Else
    strNewName = strTestName
    strBaseName = strTestName
    lngCounter = 1
    Do Until StringValueInStringArray(strNewName, strNames, vbTextCompare) = False Or lngCounter = 1000
      lngCounter = lngCounter + 1
      strNewName = strBaseName & "_" & CStr(lngCounter)
    Loop
    MakeUniqueShapeFilename2 = strNewName
  End If
  
  If booIncludeDotSHP Then MakeUniqueShapeFilename2 = MakeUniqueShapeFilename2 & ".shp"
  
End Function

Public Function StringValueInStringArray(strValue As String, strArray() As String, lngCaseSensitive As VbCompareMethod) As Boolean
  
  StringValueInStringArray = False
  Dim lngIndex As Long
  If IsDimmed(strArray) Then
    For lngIndex = 0 To UBound(strArray)
      If StrComp(strValue, strArray(lngIndex), lngCaseSensitive) = 0 Then
        StringValueInStringArray = True
        Exit For
      End If
    Next lngIndex
  End If

End Function

Public Function ReturnStringArrayOfNames(pWS As IWorkspace, lngDatasetType As esriDatasetType, lngCount As Long) As String()

  Dim pNames As IEnumDatasetName
  Set pNames = ReturnDatasetNamesOrNothing(pWS, lngDatasetType)
  Dim lngIndex As Long
  Dim pDatasetName As IDatasetName
  lngCount = 0
  Dim strReturn() As String

  If pNames Is Nothing Then
    Exit Function
  Else
    pNames.Reset
    Set pDatasetName = pNames.Next
    Do Until pDatasetName Is Nothing
      lngCount = lngCount + 1
      ReDim Preserve strReturn(lngCount - 1)
      strReturn(lngCount - 1) = pDatasetName.Name
      Set pDatasetName = pNames.Next
    Loop
  End If

  ReturnStringArrayOfNames = strReturn

  GoTo ClearMemory
ClearMemory:
  Set pNames = Nothing

End Function

Public Function ReturnDatasetNamesOrNothing(pWS As IWorkspace, lngDatasetType As esriDatasetType) As IEnumDatasetName
  On Error GoTo ErrHandler

  Set ReturnDatasetNamesOrNothing = pWS.DatasetNames(lngDatasetType)

  Exit Function
ErrHandler:
  Set ReturnDatasetNamesOrNothing = Nothing
End Function

Public Sub ExtensionNotRegisteredMessage(strName As String, strRegisterBatchFileName As String)

'  If m_pThisExt Is Nothing Then
'    Call MyGeneralOperations.ExtensionNotRegisteredMessage( _
'      "Jenness Enterprises CodeHelper ArcGIS Tools", "Register_Jennessent_Tools.bat")
'    GoTo ClearMemory
'  End If

  Dim strReport As String
  strReport = "Problem!" & vbCrLf & vbCrLf & _
    "The extension '" & strName & "' does not appear to be properly registered with ArcGIS.  " & vbCrLf & vbCrLf & _
    "If you are running ArcGIS 10.x, do you recall seeing the 'Registration Succeeded' message " & vbCrLf & _
    "when you ran the Registration batch file (" & strRegisterBatchFileName & ")?  If not, then " & vbCrLf & _
    "this is almost certainly the problem." & vbCrLf & vbCrLf & _
    "Please shut down ArcMap completely, then run '" & strRegisterBatchFileName & "' again." & vbCrLf & _
    "If you are running Windows Vista or newer, then be sure to right-click on the batch " & vbCrLf & _
    "file and choose 'Run as Administrator'." & vbCrLf & vbCrLf & _
    "If that does not solve the problem, then shut down ArcMap, then open your 'Task Manager'" & vbCrLf & _
    "window and look for any instances of 'ArcMap.exe' processes running. Sometimes ArcMap does " & vbCrLf & _
    "not shut down completely, and in this case you will need to shut down the ArcMap.exe process " & vbCrLf & _
    "manually.  After shutting it down, run '" & strRegisterBatchFileName & "' again as described above." & vbCrLf & _
    "Then restart ArcMap."
    
  MsgBox strReport, vbExclamation, "Tool Not Registered:"

End Sub

Public Sub RefreshTableWindows_Jennessent(pStTableOrLayer As IUnknown, pApp As IApplication)

  Dim pTableWindow As ITableWindow2
  Set pTableWindow = New TableWindow
  Set pTableWindow.Application = pApp
  If TypeOf pStTableOrLayer Is IStandaloneTable Then
    Set pTableWindow = pTableWindow.FindViaStandaloneTable(pStTableOrLayer)
  ElseIf TypeOf pStTableOrLayer Is ILayer Then
    Set pTableWindow = pTableWindow.FindViaLayer(pStTableOrLayer)
  Else
    GoTo ClearMemory
  End If
  
  Dim pTableControl As ITableControl3
  If Not pTableWindow Is Nothing Then
    pTableWindow.Refresh
    
    Set pTableControl = pTableWindow.TableControl
    If Not pTableControl Is Nothing Then
      pTableControl.Redraw
    End If
  End If
  
  GoTo ClearMemory
ClearMemory:
  Set pTableWindow = Nothing
  Set pTableControl = Nothing



End Sub

Public Function FormatBySize(dblVal As Double, Optional booInsertCommas As Boolean = False) As String

  If Abs(dblVal) > 1000000 Then
    FormatBySize = Format(dblVal, "0")
  ElseIf Abs(dblVal) > 1000 Then
    FormatBySize = Format(dblVal, "0.00")
  ElseIf Abs(dblVal) > 10 Then
    FormatBySize = Format(dblVal, "0.0000")
  Else
    FormatBySize = Format(dblVal, "0.0000000")
  End If
  FormatBySize = TrimZerosAndDecimals(FormatBySize, booInsertCommas)

End Function

Public Function CompareFunctionsInModules(strName1 As String, strModule1 As String, _
    strName2 As String, strModule2 As String) As String

'  ' SAMPLE CODE
'  Dim strReport As String
'  Dim strName1 As String
'  Dim strName2 As String
'  Dim strModule1 As String
'  Dim strModule2 As String
'
'  strName1 = "D:\arcGIS_stuff\general_tools\Registration\MyGeometricOperations.bas"
'  strName2 = "D:\arcGIS_stuff\consultation\USGS\Cross_Section\Code\MyGeometricOperations.bas"
'  strModule1 = MyGeneralOperations.ReadTextFile(strName1)
'  strModule2 = MyGeneralOperations.ReadTextFile(strName2)
'
'  strReport = CompareFunctionsInModules(strName1, strModule1, strName2, strModule2)
'  Debug.Print "-------------------------"
'  Debug.Print strReport

  Dim strReport As String
  Dim pColl1 As New Collection
  Dim pColl2 As New Collection
  Dim strArray1() As String
  Dim strArray2() As String

  Call FillFunctionArrayAndCollection(strModule1, strArray1, pColl1)
  Call FillFunctionArrayAndCollection(strModule2, strArray2, pColl2)
  
  Dim lngIndex As Long
  Dim strName As String
  Dim lngCounter As Long
  
  lngCounter = 0
  strReport = "Functions existing only in " & strName1 & ":" & vbCrLf
  For lngIndex = 0 To UBound(strArray1)
    strName = strArray1(lngIndex)
    If Not CheckCollectionForKey(pColl2, strName) Then
      lngCounter = lngCounter + 1
      strReport = strReport & "  " & CStr(lngCounter) & "] " & strName & vbCrLf
    End If
  Next lngIndex
  
  strReport = strReport & vbCrLf
  lngCounter = 0
  strReport = strReport & "Functions existing only in " & strName2 & ":" & vbCrLf
  For lngIndex = 0 To UBound(strArray2)
    strName = strArray2(lngIndex)
    If Not CheckCollectionForKey(pColl1, strName) Then
      lngCounter = lngCounter + 1
      strReport = strReport & "  " & CStr(lngCounter) & "] " & strName & vbCrLf
    End If
  Next lngIndex
  
  CompareFunctionsInModules = strReport
  
  GoTo ClearMemory
ClearMemory:
  Set pColl1 = Nothing
  Set pColl2 = Nothing
  Erase strArray1
  Erase strArray2

End Function

Public Sub FillFunctionArrayAndCollection(strModule As String, strFunctionNames() As String, _
    pFunctionItems As Collection)
    
  Dim strLines() As String
  strLines = Split(strModule, vbCrLf)
  Dim lngIndex As Long
  Dim strLine As String
  Dim strLineSplit() As String
  Dim strName As String
  Dim lngCounter As Long
  
  lngCounter = -1
  For lngIndex = 0 To UBound(strLines)
    strLine = strLines(lngIndex)
    strLine = Replace(strLine, "Public ", "", , , vbTextCompare)
    strLine = Replace(strLine, "Private ", "", , , vbTextCompare)
    strLine = Replace(strLine, "Friend ", "", , , vbTextCompare)
    strLine = Trim(strLine)
    If Left(strLine, 4) = "Sub " Or Left(strLine, 9) = "Function " Then
'      Debug.Print strLine
      strLineSplit = Split(strLine, " ")
      strLineSplit = Split(strLineSplit(1), "(", , vbTextCompare)
      strName = strLineSplit(0)
      
      lngCounter = lngCounter + 1
      ReDim Preserve strFunctionNames(lngCounter)
      strFunctionNames(lngCounter) = strName
      pFunctionItems.Add True, strName
    End If
  Next lngIndex
  

  GoTo ClearMemory
ClearMemory:
  Erase strLines
  Erase strLineSplit

End Sub

Public Sub SetLegendBorderColors(pMxDoc As IMxDocument, pFLayer As IFeatureLayer)
    
'  Dim pMxDoc As IMxDocument
'  Set pMxDoc = ThisDocument
'
'  Dim pFLayer As IFeatureLayer
'  Set pFLayer = MyGeneralOperations.ReturnLayerByName("tl_2014_us_aiannh selection", pMxDoc.FocusMap)
'
'  Call SetLegendBorderColors(pMxDoc, pFLayer)
'
'  Debug.Print "Done..."
'
'ClearMemory:
'  Set pMxDoc = Nothing
'  Set pFLayer = Nothing
  
  Dim pGFLayer As IGeoFeatureLayer
  Set pGFLayer = pFLayer
  
  Dim pRenderer As IFeatureRenderer
  Set pRenderer = pGFLayer.Renderer
  If pFLayer.FeatureClass.ShapeType <> esriGeometryPolygon Then
    MsgBox "Feature Layer '" & pFLayer.Name & "' is not a polygon feature class..."
    Exit Sub
  End If
  
  Dim pRendClasses As IRendererClasses
  Dim lngCount As Long
  Dim pClass As IClass
  Dim pSimpleRender As ISimpleRenderer
  Dim pPolygon As IPolygon
  Dim pFill As ISimpleFillSymbol
  Dim pLegendInfo As ILegendInfo
  Dim pLegendGroup As ILegendGroup
  Dim lngIndex As Long
  Dim pLegendClass As ILegendClass
  Dim lngIndex2 As Long
  Dim pOutline As ILineSymbol
  
  If TypeOf pRenderer Is ISimpleRenderer Then
    Set pSimpleRender = pRenderer
    Set pFill = pSimpleRender.Symbol
    Set pOutline = pFill.Outline
    pOutline.Color = pFill.Color
    pFill.Outline = pOutline
    pMxDoc.UpdateContents
    Set pGFLayer.Renderer = pRenderer
    pMxDoc.ActiveView.Refresh
    
  ElseIf TypeOf pRenderer Is IClassBreaksRenderer Or TypeOf pRenderer Is IUniqueValueRenderer Or _
      TypeOf pRenderer Is IUniqueValueRenderer Then
    Set pRendClasses = pRenderer
    Set pLegendInfo = pFLayer
    For lngIndex = 0 To pLegendInfo.LegendGroupCount - 1
      Set pLegendGroup = pLegendInfo.LegendGroup(lngIndex)
      For lngIndex2 = 0 To pLegendGroup.ClassCount - 1
        Set pLegendClass = pLegendGroup.Class(lngIndex2)
        Set pFill = pLegendClass.Symbol
        Set pOutline = pFill.Outline
        pOutline.Color = pFill.Color
        pFill.Outline = pOutline
      Next lngIndex2
    Next lngIndex
    
    Set pGFLayer.Renderer = pRenderer
    
    pMxDoc.UpdateContents
    pMxDoc.ActiveView.Refresh
  End If
 
ClearMemory:
  Set pGFLayer = Nothing
  Set pRenderer = Nothing
  Set pRendClasses = Nothing
  Set pClass = Nothing
  Set pSimpleRender = Nothing
  Set pPolygon = Nothing
  Set pFill = Nothing
  Set pLegendInfo = Nothing
  Set pLegendGroup = Nothing
  Set pLegendClass = Nothing

End Sub


Public Function ReturnEmptyFClassWithSameSchema(pFClass As IFeatureClass, pWS_NothingForInMemory As IWorkspace, _
    varFieldIndexArray() As Variant, strName As String, booHasFields As Boolean, _
    Optional lngForceGeometryType As esriGeometryType = esriGeometryAny) As IFeatureClass
  
  ' varFieldIndexArray WILL HAVE 4 COLUMNS AND ANY NUMBER OR ROWS.
  ' COLUMN 0 = SOURCE FIELD NAME
  ' COLUMN 1 = SOURCE FIELD INDEX
  ' COLUMN 2 = NEW FIELD NAME
  ' COLUMN 3 = NEW FIELD INDEX
  
  Dim pFields As IFields
  Set pFields = pFClass.Fields
  Dim booIsShapefile As Boolean
  Dim booIsAccess As Boolean
  Dim booIsFGDB As Boolean
  Dim booIsInMem As Boolean
  Dim lngCategory As JenDatasetTypes
  
  If lngForceGeometryType = esriGeometryAny Then
    lngForceGeometryType = pFClass.ShapeType
  End If
  
  If pWS_NothingForInMemory Is Nothing Then
    booIsInMem = True
  Else
    booIsInMem = False
    booIsShapefile = ReturnWorkspaceFactoryType(pWS_NothingForInMemory.WorkspaceFactory.GetClassID) = "Esri Shapefile Workspace Factory"
    booIsAccess = ReturnWorkspaceFactoryType(pWS_NothingForInMemory.WorkspaceFactory.GetClassID) = "Esri Access Workspace Factory"
    booIsFGDB = ReturnWorkspaceFactoryType(pWS_NothingForInMemory.WorkspaceFactory.GetClassID) = "File GeoDatabase Workspace Factory"
  End If
  
  If booIsAccess Then
    lngCategory = ENUM_PersonalGDB
  ElseIf booIsFGDB Then
    lngCategory = ENUM_FileGDB
  End If
    
  Dim lngIndex As Long
  Dim pSrcField As iField
  Dim pNewField As iField
  Dim pNewFieldEdit As IFieldEdit
  Dim pClone As IClone
  
  Dim pNewFieldArray As esriSystem.IVariantArray
  Set pNewFieldArray = New esriSystem.varArray
  
  Dim lngCounter As Long
  lngCounter = -1
  Dim varReturnArray() As Variant
  
  For lngIndex = 0 To pFields.FieldCount - 1
    Set pSrcField = pFields.Field(lngIndex)
    If Not pSrcField.Type = esriFieldTypeOID And pSrcField.Type <> esriFieldTypeGeometry And _
        StrComp(Left(pSrcField.Name, 6), "Shape_", vbTextCompare) <> 0 Then
      Set pClone = pSrcField
      Set pNewField = pClone.Clone
      Set pNewFieldEdit = pNewField
      With pNewFieldEdit
        If booIsShapefile Then
          .Name = MyGeneralOperations.ReturnAcceptableFieldName2(pSrcField.Name, pNewFieldArray, booIsShapefile, booIsAccess, False, booIsFGDB)
          .IsNullable = False
          If pSrcField.Type = esriFieldTypeDouble Then
            .Precision = 16
            .Scale = 6
          End If
        Else
          .IsNullable = True
        End If
      End With
      pNewFieldArray.Add pNewField
      
      lngCounter = lngCounter + 1
      ReDim Preserve varReturnArray(3, lngCounter)
      varReturnArray(0, lngCounter) = pSrcField.Name
      varReturnArray(1, lngCounter) = lngIndex
      varReturnArray(2, lngCounter) = pNewField.Name
      
    End If
  Next lngIndex
  
  Dim pDataset As IDataset
  Dim pGeoDataset As IGeoDataset
  Set pDataset = pFClass
  Set pGeoDataset = pFClass
  Dim pGeomDef As IGeometryDef
  Set pGeomDef = pFClass.Fields.Field(pFClass.FindField(pFClass.ShapeFieldName)).GeometryDef
  
  Dim pNewFClass As IFeatureClass
'  If booIsInMem Then
'    Set pNewFClass = MyGeneralOperations.CreateInMemoryFeatureClass_Empty(pNewFieldArray, strName, pGeoDataset.SpatialReference, _
'        pGeomDef.GeometryType, pGeomDef.HasM, pGeomDef.HasZ)
'  ElseIf booIsFGDB Or booIsAccess Then
'    Set pNewFClass = MyGeneralOperations.CreateGDBFeatureClass2(pWS_NothingForInMemory, strName, esriFTSimple, pGeoDataset.SpatialReference, _
'        pGeomDef.GeometryType, pNewFieldArray, , , , False, lngCategory, pGeoDataset.Extent, , pGeomDef.HasZ, pGeomDef.HasM)
'  ElseIf booIsShapefile Then
'    Set pNewFClass = MyGeneralOperations.CreateShapefileFeatureClass2(pWS_NothingForInMemory.PathName, strName, _
'        pGeoDataset.SpatialReference, pGeomDef.GeometryType, pNewFieldArray, False, pGeomDef.HasZ, pGeomDef.HasM)
'  Else
'    MsgBox "No code written for this workspace type!"
'    GoTo ClearMemory
'  End If

  ' SPECIAL FOR MARGARET'S PROJECT
  Dim pEnv As IEnvelope
  Set pEnv = New Envelope
  pEnv.PutCoords -5, -5, 5, 5
  Dim pSpRef As ISpatialReference
  Set pSpRef = MyGeneralOperations.CreateGeneralProjectedSpatialReference(26912)
  
  If booIsInMem Then
    Set pNewFClass = MyGeneralOperations.CreateInMemoryFeatureClass_Empty(pNewFieldArray, strName, pSpRef, _
        lngForceGeometryType, pGeomDef.HasM, pGeomDef.HasZ)
  ElseIf booIsFGDB Or booIsAccess Then
    Set pNewFClass = MyGeneralOperations.CreateGDBFeatureClass2(pWS_NothingForInMemory, strName, esriFTSimple, pSpRef, _
        lngForceGeometryType, pNewFieldArray, , , , False, lngCategory, pEnv, , pGeomDef.HasZ, pGeomDef.HasM)
  ElseIf booIsShapefile Then
    Set pNewFClass = MyGeneralOperations.CreateShapefileFeatureClass2(pWS_NothingForInMemory.PathName, strName, _
        pSpRef, lngForceGeometryType, pNewFieldArray, False, pGeomDef.HasZ, pGeomDef.HasM)
  Else
    MsgBox "No code written for this workspace type!"
    GoTo ClearMemory
  End If
  
  booHasFields = lngCounter > -1
  If booHasFields Then
    For lngIndex = 0 To lngCounter
      varReturnArray(3, lngIndex) = pNewFClass.FindField(CStr(varReturnArray(2, lngIndex)))
    Next lngIndex
  End If
  varFieldIndexArray = varReturnArray
  Set ReturnEmptyFClassWithSameSchema = pNewFClass
  
  GoTo ClearMemory
ClearMemory:
  Set pFields = Nothing
  Set pSrcField = Nothing
  Set pNewField = Nothing
  Set pNewFieldEdit = Nothing
  Set pClone = Nothing
  Set pNewFieldArray = Nothing
  Erase varReturnArray
  Set pDataset = Nothing
  Set pGeoDataset = Nothing
  Set pGeomDef = Nothing
  Set pNewFClass = Nothing

End Function

Public Function HexifyName(strName As String) As String
  Dim strChar As String
  Dim lngIndex As Long
  Dim strReturn As String
  For lngIndex = 1 To Len(strName)
    strChar = CStr(Hex(Asc(Mid(strName, lngIndex, 1))))
    strChar = String(3 - Len(strChar), "0") & strChar
    strReturn = strReturn & strChar
  Next lngIndex
  HexifyName = strReturn
  
End Function

Public Function WordifyHex(strHexed As String) As String

  Dim lngIndex As Long
  lngIndex = 1
  Dim strReturn As String
  
  Do Until lngIndex > Len(strHexed)
    strReturn = strReturn & Chr(CLng("&H" & Mid(strHexed, lngIndex, 3)))
    
    lngIndex = lngIndex + 3
  Loop

  WordifyHex = strReturn
End Function

Public Function ReturnDomainCollectionFromField(pField As iField) As Collection
 
'  Dim pMxDoc As IMxDocument
'  Dim pFLayer As IFeatureLayer
'  Dim pFClass As IFeatureClass
'  Dim pFields As IFields
'
'  Set pMxDoc = ThisDocument
'  Set pFLayer = MyGeneralOperations.ReturnLayerByName("PADUS 1.4 Combined", pMxDoc.FocusMap)
'  Set pFClass = pFLayer.FeatureClass
'  Set pFields = pFClass.Fields
'
'  Dim pColl As Collection
'  Set pColl = ReturnDomainCollectionFromField(pFields.Field(pFields.FindField("Own_Name")))
'
'  Dim strTest As String
'  strTest = "SDOL"
'  Debug.Print strTest & " --> " & pColl.Item(CStr(strTest))
'
'ClearMemory:
'  Set pMxDoc = Nothing
'  Set pFLayer = Nothing
'  Set pFClass = Nothing
'  Set pFields = Nothing
'  Set pColl = Nothing
 
  
  Dim pDomain As IDomain
  Set pDomain = pField.Domain
  Dim pCVDomain As ICodedValueDomain2
  Dim lngIndex As Long
  Dim varCode As Variant
  Dim varValue As Variant
  Dim pReturn As New Collection
 
  If Not pDomain Is Nothing Then
    If TypeOf pDomain Is ICodedValueDomain Then
      Set pCVDomain = pDomain
      If pCVDomain.CodeCount >= 1 Then
       
        For lngIndex = 0 To pCVDomain.CodeCount - 1
          varCode = pCVDomain.Name(lngIndex)
          varValue = pCVDomain.Value(lngIndex)
          If Not IsNull(varCode) And Not IsNull(varValue) Then
            pReturn.Add CStr(varCode), varValue
          End If
        Next lngIndex
       
        If pReturn.Count = 0 Then
          Set ReturnDomainCollectionFromField = Nothing
          GoTo ClearMemory
        Else
          Set ReturnDomainCollectionFromField = pReturn
          GoTo ClearMemory
        End If
      Else
        Set ReturnDomainCollectionFromField = Nothing
        GoTo ClearMemory
      End If
    Else
      Set ReturnDomainCollectionFromField = Nothing
      GoTo ClearMemory
    End If
  Else
    Set ReturnDomainCollectionFromField = Nothing
    GoTo ClearMemory
  End If
   
  GoTo ClearMemory
ClearMemory:
  Set pDomain = Nothing
  Set pCVDomain = Nothing
  varCode = Null
  Set varValue = Nothing
  Set pReturn = Nothing
 
End Function
Public Function ExportToCSV(pFLayerOrStandaloneTable As IUnknown, strFilename As String, _
    booOverwrite As Boolean, booForceUniqueFilename As Boolean, booOnlySelected As Boolean, _
    booSucceeded As Boolean, Optional varFieldNames As Variant = Null, _
    Optional pApp As IApplication = Nothing, Optional booSkipOID As Boolean = True) As String
  
  Dim booShowProgress As Boolean
  Dim pSBar As IStatusBar
  Dim pProg As IStepProgressor
  booShowProgress = Not pApp Is Nothing
  
  
  booSucceeded = True
  Dim lngCount As Long
  Dim lngWriteInterval As Long
  Dim lngCounter As Long
  
  Dim pStTable As IStandaloneTable
  Dim pRowSel As ITableSelection
  Dim pTable As ITable
  Dim pCursor As ICursor
  Dim pRow As IRow
  
  Dim pFLayer As IFeatureLayer
  Dim pFeatSel As IFeatureSelection
  Dim pFClass As IFeatureClass
  Dim pFCursor As IFeatureCursor
  Dim pFeature As IFeature
  
  Dim pSelSet As ISelectionSet
  Dim strDataName As String
  Dim pFields As IFields
  
  If Not (TypeOf pFLayerOrStandaloneTable Is IStandaloneTable Or _
          TypeOf pFLayerOrStandaloneTable Is IFeatureLayer Or _
          TypeOf pFLayerOrStandaloneTable Is IFeatureClass Or _
          TypeOf pFLayerOrStandaloneTable Is ITable) Then
    booSucceeded = False
    ExportToCSV = "Invalid Data Type:  This function requires either a feature class or table"
    GoTo ClearMemory
  End If
  
  Dim booIsTable As Boolean
  Dim pDataset As IDataset
  
  If TypeOf pFLayerOrStandaloneTable Is IStandaloneTable Then
    booIsTable = True
    Set pStTable = pFLayerOrStandaloneTable
    strDataName = pStTable.Name
    Set pTable = pStTable.Table
    Set pFields = pTable.Fields
    If booOnlySelected Then
      Set pRowSel = pStTable
      Set pSelSet = pRowSel.SelectionSet
      lngCount = pSelSet.Count
      If lngCount = 0 Then
        booSucceeded = False
        ExportToCSV = "No rows selected in table"
        GoTo ClearMemory
      End If
      pSelSet.Search Nothing, False, pCursor
    Else
      Set pCursor = pTable.Search(Nothing, False)
      lngCount = pTable.RowCount(Nothing)
    End If
  ElseIf TypeOf pFLayerOrStandaloneTable Is ITable Then
    booIsTable = True
    Set pTable = pFLayerOrStandaloneTable
    Set pDataset = pTable
    strDataName = pDataset.BrowseName
    Set pFields = pTable.Fields
    Set pCursor = pTable.Search(Nothing, False)
    lngCount = pTable.RowCount(Nothing)
  ElseIf TypeOf pFLayerOrStandaloneTable Is IFeatureLayer Then
    booIsTable = False
    Set pFLayer = pFLayerOrStandaloneTable
    strDataName = pFLayer.Name
    Set pFClass = pFLayer.FeatureClass
    Set pFields = pFClass.Fields
    If booOnlySelected Then
      Set pFeatSel = pFLayer
      Set pSelSet = pFeatSel.SelectionSet
      lngCount = pSelSet.Count
      If lngCount = 0 Then
        booSucceeded = False
        ExportToCSV = "No features selected in feature class"
        GoTo ClearMemory
      End If
      pSelSet.Search Nothing, False, pFCursor
    Else
      Set pFCursor = pFClass.Search(Nothing, False)
      lngCount = pFClass.FeatureCount(Nothing)
    End If
  ElseIf TypeOf pFLayerOrStandaloneTable Is IFeatureClass Then
    booIsTable = True
    Set pFClass = pFLayerOrStandaloneTable
    Set pDataset = pFClass
    strDataName = pDataset.BrowseName
    Set pTable = pFClass
    Set pFields = pTable.Fields
    Set pCursor = pTable.Search(Nothing, False)
    lngCount = pTable.RowCount(Nothing)
  End If
  
  If aml_func_mod.ExistFileDir(strFilename) Then
    If booForceUniqueFilename Then
      strFilename = MakeUniqueFilename(strFilename)
    ElseIf booOverwrite Then
      Kill strFilename
    Else
      ExportToCSV = "File Exists"
      booSucceeded = False
      GoTo ClearMemory
    End If
  End If
    
  Dim lngFieldCounter As Long
  Dim varFieldIndexes() As Variant
  Dim lngIndex As Long
  Dim pField As iField
  Dim strFieldName As String
  Dim lngFieldIndex As Long
  Dim strReport As String
  Dim lngFileNumber As Long
  Dim lngFieldType As esriFieldType
  
  lngFieldCounter = -1
  If IsNull(varFieldNames) Then
    For lngIndex = 0 To pTable.Fields.FieldCount - 1
      Set pField = pFields.Field(lngIndex)
      If pField.Type <> esriFieldTypeBlob And pField.Type <> esriFieldTypeRaster And _
              pField.Type <> esriFieldTypeGeometry And _
              (pField.Type <> esriFieldTypeOID Or Not booSkipOID) Then
                
        lngFieldCounter = lngFieldCounter + 1
        ReDim Preserve varFieldIndexes(2, lngFieldCounter)
        varFieldIndexes(0, lngFieldCounter) = pField.AliasName
        varFieldIndexes(1, lngFieldCounter) = lngIndex
        varFieldIndexes(2, lngFieldCounter) = pField.Type
        
        strReport = strReport & aml_func_mod.QuoteString(pField.AliasName) & ","
        
      End If
    Next lngIndex
  Else
    For lngIndex = 0 To UBound(varFieldNames)
      strFieldName = varFieldNames(lngIndex)
      lngFieldIndex = pFields.FindField(strFieldName)
      If lngFieldIndex = -1 Then lngFieldIndex = pFields.FindFieldByAliasName(strFieldName)
      If lngFieldIndex = -1 Then
        ExportToCSV = "Unable to find Attribute Field '" & strFieldName & "'"
        booSucceeded = False
        GoTo ClearMemory
      End If
      Set pField = pFields.Field(lngFieldIndex)
      
      If pField.Type <> esriFieldTypeBlob And pField.Type <> esriFieldTypeRaster And _
              pField.Type <> esriFieldTypeGeometry Then
        lngFieldCounter = lngFieldCounter + 1
        ReDim Preserve varFieldIndexes(2, lngFieldCounter)
        varFieldIndexes(0, lngFieldCounter) = pField.Name
        varFieldIndexes(1, lngFieldCounter) = lngFieldIndex
        varFieldIndexes(2, lngFieldCounter) = pField.Type
        
        strReport = strReport & aml_func_mod.QuoteString(pField.AliasName) & ","
      End If
    Next
  End If
  
  If lngFieldCounter = -1 Then
    ExportToCSV = "No Exportable Fields Found"
    booSucceeded = False
    GoTo ClearMemory
  End If
  
  If Right(strReport, 1) = "," Then strReport = Left(strReport, Len(strReport) - 1)
  lngFileNumber = FreeFile(0)
  Open strFilename For Output As #lngFileNumber
  Print #lngFileNumber, strReport
  Close #lngFileNumber
  
  If lngCount < 1000 Then
    lngWriteInterval = 1
  ElseIf lngCount > 1000000 Then
    lngWriteInterval = 100
  Else
    lngWriteInterval = lngCount / 1000
  End If
  
  
  If booShowProgress Then
    Set pSBar = pApp.StatusBar
    Set pProg = pSBar.ProgressBar
    pSBar.ShowProgressBar "Exporting '" & strDataName & "' to " & strFilename & "...", 0, lngCount, 1, True
    pProg.position = 0
  End If
  
  strReport = ""
  If booIsTable Then
    Set pRow = pCursor.NextRow
    Do Until pRow Is Nothing
      lngCounter = lngCounter + 1
      If booShowProgress Then pProg.Step
      For lngIndex = 0 To UBound(varFieldIndexes, 2)
        lngFieldIndex = varFieldIndexes(1, lngIndex)
        lngFieldType = varFieldIndexes(2, lngIndex)
        If IsNull(pRow.Value(lngFieldIndex)) Then
          strReport = strReport & ","
        ElseIf lngFieldType = esriFieldTypeString Or lngFieldType = esriFieldTypeXML Then
          strReport = strReport & aml_func_mod.QuoteString(pRow.Value(lngFieldIndex)) & ","
        ElseIf lngFieldType = esriFieldTypeDouble Then
          strReport = strReport & Format(pRow.Value(lngFieldIndex), "0.000000000000") & ","
        ElseIf lngFieldType = esriFieldTypeSingle Then
          strReport = strReport & Format(pRow.Value(lngFieldIndex), "0.000000") & ","
        ElseIf lngFieldType = esriFieldTypeDate Then
          strReport = strReport & Format(pRow.Value(lngFieldIndex), "yyyy/mm/dd, hh:Nn:Ss, dddd") & ","
        ElseIf lngFieldType = esriFieldTypeInteger Or lngFieldType = esriFieldTypeSmallInteger Then
          strReport = strReport & Format(pRow.Value(lngFieldIndex), "0") & ","
        Else
          strReport = strReport & CStr(pRow.Value(lngFieldIndex)) & ","
        End If
      Next lngIndex
      
      strReport = Left(strReport, Len(strReport) - 1)
      If lngCounter >= lngWriteInterval Then
        lngCounter = 0
        lngFileNumber = FreeFile(0)
        Open strFilename For Append As #lngFileNumber
        Print #lngFileNumber, strReport
        Close #lngFileNumber
        strReport = ""
        DoEvents
      Else
        strReport = strReport & vbCrLf
      End If
      
      Set pRow = pCursor.NextRow
    Loop
  Else
    Set pFeature = pFCursor.NextFeature
    Do Until pFeature Is Nothing
      If booShowProgress Then pProg.Step
      lngCounter = lngCounter + 1
      For lngIndex = 0 To UBound(varFieldIndexes, 2)
        lngFieldIndex = varFieldIndexes(1, lngIndex)
        lngFieldType = varFieldIndexes(2, lngIndex)
        If lngFieldType = esriFieldTypeString Or lngFieldType = esriFieldTypeXML Then
          strReport = strReport & aml_func_mod.QuoteString(pFeature.Value(lngFieldIndex)) & ","
        ElseIf lngFieldType = esriFieldTypeDouble Then
          strReport = strReport & Format(pFeature.Value(lngFieldIndex), "0.000000000000") & ","
        ElseIf lngFieldType = esriFieldTypeSingle Then
          strReport = strReport & Format(pFeature.Value(lngFieldIndex), "0.000000") & ","
        ElseIf lngFieldType = esriFieldTypeDate Then
          strReport = strReport & Format(pFeature.Value(lngFieldIndex), "yyyy/mm/dd, hh:Nn:Ss, dddd") & ","
        ElseIf lngFieldType = esriFieldTypeInteger Or lngFieldType = esriFieldTypeSmallInteger Then
          strReport = strReport & Format(pFeature.Value(lngFieldIndex), "0") & ","
        Else
          strReport = strReport & CStr(pFeature.Value(lngFieldIndex)) & ","
        End If
      Next lngIndex
      
      strReport = Left(strReport, Len(strReport) - 1)
      If lngCounter >= lngWriteInterval Then
        lngCounter = 0
        lngFileNumber = FreeFile(0)
        Open strFilename For Append As #lngFileNumber
        Print #lngFileNumber, strReport
        Close #lngFileNumber
        strReport = ""
        DoEvents
      Else
        strReport = strReport & vbCrLf
      End If
    
      Set pFeature = pFCursor.NextFeature
    Loop
  End If
  
  If strReport <> "" Then
    lngFileNumber = FreeFile(0)
    Open strFilename For Append As #lngFileNumber
    Print #lngFileNumber, strReport
    Close #lngFileNumber
  End If
  
  If booShowProgress Then
    pSBar.HideProgressBar
    pProg.position = 0
  End If
  
  booSucceeded = True
  ExportToCSV = "Succeeded"
  
  GoTo ClearMemory
  Exit Function
  
ClearMemory:
  Set pStTable = Nothing
  Set pRowSel = Nothing
  Set pTable = Nothing
  Set pCursor = Nothing
  Set pRow = Nothing
  Set pFLayer = Nothing
  Set pFeatSel = Nothing
  Set pFClass = Nothing
  Set pFCursor = Nothing
  Set pFeature = Nothing
  Set pSelSet = Nothing
  Erase varFieldIndexes
  Set pField = Nothing


  
End Function
Public Function ExportToCSV_SpecialCases(pFLayerOrStandaloneTable As IUnknown, strFilename As String, _
    booOverwrite As Boolean, booForceUniqueFilename As Boolean, booOnlySelected As Boolean, _
    booSucceeded As Boolean, Optional varFieldNames As Variant = Null, _
    Optional pApp As IApplication = Nothing, Optional booSkipOID As Boolean = True, _
    Optional booCreateAreaAndCentroidFields = False, Optional pPlotLocColl As Collection) As String
  
  Dim booShowProgress As Boolean
  Dim pSBar As IStatusBar
  Dim pProg As IStepProgressor
  booShowProgress = Not pApp Is Nothing
  
  
  booSucceeded = True
  Dim lngCount As Long
  Dim lngWriteInterval As Long
  Dim lngCounter As Long
  
  Dim pStTable As IStandaloneTable
  Dim pRowSel As ITableSelection
  Dim pTable As ITable
  Dim pCursor As ICursor
  Dim pRow As IRow
  
  Dim pFLayer As IFeatureLayer
  Dim pFeatSel As IFeatureSelection
  Dim pFClass As IFeatureClass
  Dim pFCursor As IFeatureCursor
  Dim pFeature As IFeature
  
  Dim pSelSet As ISelectionSet
  Dim strDataName As String
  Dim pFields As IFields
  Dim dblEasting As Double
  Dim dblNorthing As Double
  Dim lngQuadratIDIndex As Long
  Dim strQuad As String
  
  If TypeOf pFLayerOrStandaloneTable Is IStandaloneTable Or _
          TypeOf pFLayerOrStandaloneTable Is ITable And Not TypeOf pFLayerOrStandaloneTable Is IFeatureClass Then booCreateAreaAndCentroidFields = False
  
  If Not (TypeOf pFLayerOrStandaloneTable Is IStandaloneTable Or _
          TypeOf pFLayerOrStandaloneTable Is IFeatureLayer Or _
          TypeOf pFLayerOrStandaloneTable Is IFeatureClass Or _
          TypeOf pFLayerOrStandaloneTable Is ITable) Then
    booSucceeded = False
    ExportToCSV_SpecialCases = "Invalid Data Type:  This function requires either a feature class or table"
    GoTo ClearMemory
  End If
  
  Dim booIsTable As Boolean
  Dim pDataset As IDataset
  
  If TypeOf pFLayerOrStandaloneTable Is IFeatureLayer Then
    booIsTable = False
    Set pFLayer = pFLayerOrStandaloneTable
    strDataName = pFLayer.Name
    Set pFClass = pFLayer.FeatureClass
    Set pFields = pFClass.Fields
    If booOnlySelected Then
      Set pFeatSel = pFLayer
      Set pSelSet = pFeatSel.SelectionSet
      lngCount = pSelSet.Count
      If lngCount = 0 Then
        booSucceeded = False
        ExportToCSV_SpecialCases = "No features selected in feature class"
        GoTo ClearMemory
      End If
      pSelSet.Search Nothing, False, pFCursor
    Else
      Set pFCursor = pFClass.Search(Nothing, False)
      lngCount = pFClass.FeatureCount(Nothing)
    End If
  ElseIf TypeOf pFLayerOrStandaloneTable Is IFeatureClass Then
    booIsTable = False
    Set pFClass = pFLayerOrStandaloneTable
    Set pDataset = pFClass
    strDataName = pDataset.BrowseName
    Set pTable = pFClass
    Set pFields = pTable.Fields
'    Set pCursor = pTable.Search(Nothing, False)
    Set pFCursor = pFClass.Search(Nothing, False)
    lngCount = pTable.RowCount(Nothing)
  ElseIf TypeOf pFLayerOrStandaloneTable Is IStandaloneTable Then
    booIsTable = True
    Set pStTable = pFLayerOrStandaloneTable
    strDataName = pStTable.Name
    Set pTable = pStTable.Table
    Set pFields = pTable.Fields
    If booOnlySelected Then
      Set pRowSel = pStTable
      Set pSelSet = pRowSel.SelectionSet
      lngCount = pSelSet.Count
      If lngCount = 0 Then
        booSucceeded = False
        ExportToCSV_SpecialCases = "No rows selected in table"
        GoTo ClearMemory
      End If
      pSelSet.Search Nothing, False, pCursor
    Else
      Set pCursor = pTable.Search(Nothing, False)
      lngCount = pTable.RowCount(Nothing)
    End If
  ElseIf TypeOf pFLayerOrStandaloneTable Is ITable Then
    booIsTable = True
    Set pTable = pFLayerOrStandaloneTable
    Set pDataset = pTable
    strDataName = pDataset.BrowseName
    Set pFields = pTable.Fields
    Set pCursor = pTable.Search(Nothing, False)
    lngCount = pTable.RowCount(Nothing)
  End If
  
  If aml_func_mod.ExistFileDir(strFilename) Then
    If booForceUniqueFilename Then
      strFilename = MakeUniqueFilename(strFilename)
    ElseIf booOverwrite Then
      Kill strFilename
    Else
      ExportToCSV_SpecialCases = "File Exists"
      booSucceeded = False
      GoTo ClearMemory
    End If
  End If
    
  Dim lngFieldCounter As Long
  Dim varFieldIndexes() As Variant
  Dim lngIndex As Long
  Dim pField As iField
  Dim strFieldName As String
  Dim lngFieldIndex As Long
  Dim strReport As String
  Dim lngFileNumber As Long
  Dim lngFieldType As esriFieldType
  
  lngFieldCounter = -1
  If IsNull(varFieldNames) Then
    For lngIndex = 0 To pTable.Fields.FieldCount - 1
      Set pField = pFields.Field(lngIndex)
      If pField.Type <> esriFieldTypeBlob And pField.Type <> esriFieldTypeRaster And _
              pField.Type <> esriFieldTypeGeometry And _
              (pField.Type <> esriFieldTypeOID Or Not booSkipOID) Then
                
        lngFieldCounter = lngFieldCounter + 1
        ReDim Preserve varFieldIndexes(2, lngFieldCounter)
        varFieldIndexes(0, lngFieldCounter) = pField.AliasName
        varFieldIndexes(1, lngFieldCounter) = lngIndex
        varFieldIndexes(2, lngFieldCounter) = pField.Type
        
        strReport = strReport & aml_func_mod.QuoteString(pField.AliasName) & ","
        
      End If
    Next lngIndex
  Else
    For lngIndex = 0 To UBound(varFieldNames)
      strFieldName = varFieldNames(lngIndex)
      lngFieldIndex = pFields.FindField(strFieldName)
      If lngFieldIndex = -1 Then lngFieldIndex = pFields.FindFieldByAliasName(strFieldName)
      If lngFieldIndex = -1 Then
        ExportToCSV_SpecialCases = "Unable to find Attribute Field '" & strFieldName & "'"
        booSucceeded = False
        GoTo ClearMemory
      End If
      Set pField = pFields.Field(lngFieldIndex)
      
      If pField.Type <> esriFieldTypeBlob And pField.Type <> esriFieldTypeRaster And _
              pField.Type <> esriFieldTypeGeometry Then
        lngFieldCounter = lngFieldCounter + 1
        ReDim Preserve varFieldIndexes(2, lngFieldCounter)
        varFieldIndexes(0, lngFieldCounter) = pField.Name
        varFieldIndexes(1, lngFieldCounter) = lngFieldIndex
        varFieldIndexes(2, lngFieldCounter) = pField.Type
        
        strReport = strReport & aml_func_mod.QuoteString(pField.AliasName) & ","
      End If
    Next
  End If
  
  If booCreateAreaAndCentroidFields Then strReport = strReport & _
      """Easting_NAD83_UTM_Zone_12"",""Northing_NAD83_UTM_Zone_12""," & _
      """X_Coord_In_Quadrat"",""Y_Coord_In_Quadrat"",""Sq_Cm"""
  
  If lngFieldCounter = -1 Then
    ExportToCSV_SpecialCases = "No Exportable Fields Found"
    booSucceeded = False
    GoTo ClearMemory
  End If
  
  If Right(strReport, 1) = "," Then strReport = Left(strReport, Len(strReport) - 1)
  lngFileNumber = FreeFile(0)
  Open strFilename For Output As #lngFileNumber
  Print #lngFileNumber, strReport
  Close #lngFileNumber
  
  If lngCount < 1000 Then
    lngWriteInterval = 1
  ElseIf lngCount > 1000000 Then
    lngWriteInterval = 100
  Else
    lngWriteInterval = lngCount / 1000
  End If
  
  Dim pPoly As IPolygon
  Dim pArea As IArea
  
  If booShowProgress Then
    Set pSBar = pApp.StatusBar
    Set pProg = pSBar.ProgressBar
    pSBar.ShowProgressBar "Exporting '" & strDataName & "' to " & strFilename & "...", 0, lngCount, 1, True
    pProg.position = 0
  End If
  
  Dim booInclude As Boolean
  Dim varArray() As Variant
  
  strReport = ""
  If booIsTable Then
    Set pRow = pCursor.NextRow
    Do Until pRow Is Nothing
      lngCounter = lngCounter + 1
      If booShowProgress Then pProg.Step
      For lngIndex = 0 To UBound(varFieldIndexes, 2)
        lngFieldIndex = varFieldIndexes(1, lngIndex)
        lngFieldType = varFieldIndexes(2, lngIndex)
        If lngFieldType = esriFieldTypeString Or lngFieldType = esriFieldTypeXML Then
          strReport = strReport & aml_func_mod.QuoteString(pRow.Value(lngFieldIndex)) & ","
        ElseIf lngFieldType = esriFieldTypeDouble Then
          strReport = strReport & Format(pRow.Value(lngFieldIndex), "0.000000000000") & ","
        ElseIf lngFieldType = esriFieldTypeSingle Then
          strReport = strReport & Format(pRow.Value(lngFieldIndex), "0.000000") & ","
        ElseIf lngFieldType = esriFieldTypeDate Then
          strReport = strReport & Format(pRow.Value(lngFieldIndex), "yyyy/mm/dd, hh:Nn:Ss, dddd") & ","
        ElseIf lngFieldType = esriFieldTypeInteger Or lngFieldType = esriFieldTypeSmallInteger Then
          strReport = strReport & Format(pRow.Value(lngFieldIndex), "0") & ","
        Else
          strReport = strReport & CStr(pRow.Value(lngFieldIndex)) & ","
        End If
      Next lngIndex
      
      strReport = Left(strReport, Len(strReport) - 1)
      If lngCounter >= lngWriteInterval Then
        lngCounter = 0
        lngFileNumber = FreeFile(0)
        Open strFilename For Append As #lngFileNumber
        Print #lngFileNumber, strReport
        Close #lngFileNumber
        strReport = ""
        DoEvents
      Else
        strReport = strReport & vbCrLf
      End If
      
      Set pRow = pCursor.NextRow
    Loop
  Else
    lngQuadratIDIndex = pFCursor.FindField("Quadrat")
    Set pFeature = pFCursor.NextFeature
    Do Until pFeature Is Nothing
      If booShowProgress Then pProg.Step
      lngCounter = lngCounter + 1
      
      booInclude = True
      If booCreateAreaAndCentroidFields Then
        strQuad = Trim(CStr(pFeature.Value(lngQuadratIDIndex)))
        strQuad = Replace(strQuad, "*", "")
        strQuad = Trim(strQuad)
        If strQuad = "6" Then
          DoEvents
        End If
        
  '      pCollection_To_Fill.Add Array(dblEasting, dblNorthing, strName, strComment, strNote, str2016, str2017), strQuad
        varArray = pPlotLocColl.Item(strQuad)
        dblEasting = varArray(0)
        dblNorthing = varArray(1)
        
        Set pPoly = pFeature.ShapeCopy
        If pPoly.IsEmpty Then
          booInclude = False
        Else
          Set pArea = pPoly
        End If
      End If
      
      If booInclude Then
        For lngIndex = 0 To UBound(varFieldIndexes, 2)
          lngFieldIndex = varFieldIndexes(1, lngIndex)
          lngFieldType = varFieldIndexes(2, lngIndex)
          If lngFieldType = esriFieldTypeString Or lngFieldType = esriFieldTypeXML Then
            strReport = strReport & aml_func_mod.QuoteString(pFeature.Value(lngFieldIndex)) & ","
          ElseIf lngFieldType = esriFieldTypeDouble Then
            strReport = strReport & Format(pFeature.Value(lngFieldIndex), "0.000000000000") & ","
          ElseIf lngFieldType = esriFieldTypeSingle Then
            strReport = strReport & Format(pFeature.Value(lngFieldIndex), "0.000000") & ","
          ElseIf lngFieldType = esriFieldTypeDate Then
            strReport = strReport & Format(pFeature.Value(lngFieldIndex), "yyyy/mm/dd, hh:Nn:Ss, dddd") & ","
          ElseIf lngFieldType = esriFieldTypeInteger Or lngFieldType = esriFieldTypeSmallInteger Then
            strReport = strReport & Format(pFeature.Value(lngFieldIndex), "0") & ","
          Else
            strReport = strReport & CStr(pFeature.Value(lngFieldIndex)) & ","
          End If
        Next lngIndex
        
        If booCreateAreaAndCentroidFields Then
        
'          ' REVISED TO ANCHOR TO UPPER LEFT CORNER OF QUADRAT
'          If dblCentroidX <> 0 And dblCentroidY <> 1 Then
'            Set pTransform = pPolygon
'            pTransform.Move dblCentroidX, dblCentroidY - 1
'          End If
          strReport = strReport & Format(pArea.Centroid.x, "0.000000") & "," & _
              Format(pArea.Centroid.Y, "0.000000") & "," & Format(pArea.Centroid.x - dblEasting, "0.000000") & "," & _
              Format(pArea.Centroid.Y - dblNorthing + 1, "0.000000") & "," & Format(pArea.Area * 10000, "0.000000") & ","
        End If
        
        strReport = Left(strReport, Len(strReport) - 1)
        If lngCounter >= lngWriteInterval Then
          lngCounter = 0
          lngFileNumber = FreeFile(0)
          Open strFilename For Append As #lngFileNumber
          Print #lngFileNumber, strReport
          Close #lngFileNumber
          strReport = ""
          DoEvents
        Else
          strReport = strReport & vbCrLf
        End If
      End If
    
      Set pFeature = pFCursor.NextFeature
    Loop
  End If
  
  If strReport <> "" Then
    lngFileNumber = FreeFile(0)
    Open strFilename For Append As #lngFileNumber
    Print #lngFileNumber, strReport
    Close #lngFileNumber
  End If
  
  If booShowProgress Then
    pSBar.HideProgressBar
    pProg.position = 0
  End If
  
  booSucceeded = True
  ExportToCSV_SpecialCases = "Succeeded"
  
  GoTo ClearMemory
  Exit Function
  
ClearMemory:
  Set pStTable = Nothing
  Set pRowSel = Nothing
  Set pTable = Nothing
  Set pCursor = Nothing
  Set pRow = Nothing
  Set pFLayer = Nothing
  Set pFeatSel = Nothing
  Set pFClass = Nothing
  Set pFCursor = Nothing
  Set pFeature = Nothing
  Set pSelSet = Nothing
  Erase varFieldIndexes
  Set pField = Nothing


  
End Function
Public Function FindFieldByNameOrAlias(pFields As IFields, strFieldName As String) As Long
  Dim lngIndex As Long
  Dim pField As iField
  FindFieldByNameOrAlias = -1
  For lngIndex = 0 To pFields.FieldCount - 1
    Set pField = pFields.Field(lngIndex)
    If StrComp(pField.Name, strFieldName, vbTextCompare) = 0 Then
      FindFieldByNameOrAlias = lngIndex
      Exit For
    ElseIf StrComp(pField.Name, strFieldName, vbTextCompare) = 0 Then
      FindFieldByNameOrAlias = lngIndex
      Exit For
    End If
  Next lngIndex
    
End Function
Public Function CheckCollectionForKey_CaseSensitive(colCollection As Collection, strKey As String, _
    strAllKeys() As String, booCaseSensitive As Boolean) As Boolean
  On Error GoTo ErrHandler
  
  CheckCollectionForKey_CaseSensitive = False
  Dim lngVarType As Long
  Dim lngIndex As Long
  If Not IsDimmed(strAllKeys) Then
    ' CheckCollectionForKey_CaseSensitive = False
  Else
    For lngIndex = 0 To UBound(strAllKeys)
      Debug.Print "Checking '" & strAllKeys(lngIndex)
      If booCaseSensitive Then
        If StrComp(strKey, strAllKeys(lngIndex), vbBinaryCompare) = 0 Then
          CheckCollectionForKey_CaseSensitive = True
          Exit For
        End If
      Else
        If StrComp(strKey, strAllKeys(lngIndex), vbTextCompare) = 0 Then
          CheckCollectionForKey_CaseSensitive = True
          Exit For
        End If
      End If
    Next lngIndex
  End If
    
  Exit Function
ErrHandler:
  CheckCollectionForKey_CaseSensitive = False

End Function

Public Function ReturnMonthNameFromNumber(lngMonth As Long, booInvalid As Boolean, _
    Optional booAbbreviated As Boolean = False) As String
  
  booInvalid = False
  Select Case lngMonth
    Case 1
      ReturnMonthNameFromNumber = "January"
    Case 2
      ReturnMonthNameFromNumber = "February"
    Case 3
      ReturnMonthNameFromNumber = "March"
    Case 4
      ReturnMonthNameFromNumber = "April"
    Case 5
      ReturnMonthNameFromNumber = "May"
    Case 6
      ReturnMonthNameFromNumber = "June"
    Case 7
      ReturnMonthNameFromNumber = "July"
    Case 8
      ReturnMonthNameFromNumber = "August"
    Case 9
      ReturnMonthNameFromNumber = "September"
    Case 10
      ReturnMonthNameFromNumber = "October"
    Case 11
      ReturnMonthNameFromNumber = "November"
    Case 12
      ReturnMonthNameFromNumber = "December"
  End Select
  
  If booAbbreviated Then ReturnMonthNameFromNumber = Left(ReturnMonthNameFromNumber, 3)

  GoTo ClearMemory
ClearMemory:

End Function

Public Function CreateOrReturnFileGeodatabase(strPath As String) As IWorkspace

  Dim pWSName As IWorkspaceName
  Dim pName As IName
  Dim pWS As IWorkspace
  Dim pWSFact As IWorkspaceFactory
  Set pWSFact = New FileGDBWorkspaceFactory
  
  If pWSFact.IsWorkspace(strPath) Then
    Set pWS = pWSFact.OpenFromFile(strPath, 0)
  ElseIf pWSFact.IsWorkspace(strPath & ".gdb") Then
    Set pWS = pWSFact.OpenFromFile(strPath & ".gdb", 0)
  Else
    Set pWSName = pWSFact.Create(aml_func_mod.ReturnDir3(strPath, False), _
        aml_func_mod.ReturnFilename2(strPath), Nothing, 0)
    Set pName = pWSName
    Set pWS = pName.Open
  End If
  
  Set CreateOrReturnFileGeodatabase = pWS
  
      GoTo ClearMemory

ClearMemory:
  Set pWSName = Nothing
  Set pName = Nothing
  Set pWS = Nothing
  Set pWSFact = Nothing


  
  

End Function






Public Function ReturnWorkspaceTypeName(lngType As esriWorkspaceType) As String

  If lngType = esriFileSystemWorkspace Then
    ReturnWorkspaceTypeName = "File System Workspace"
  ElseIf lngType = esriLocalDatabaseWorkspace Then
    ReturnWorkspaceTypeName = "Local Database Workspace"
  ElseIf lngType = esriRemoteDatabaseWorkspace Then
    ReturnWorkspaceTypeName = "Remote Database Workspace"
  Else
    ReturnWorkspaceTypeName = "Unknown Workspace Type"
  End If

End Function

Public Function ReturnWorkspaceFactoryType(strClassID As String) As String
  
    
  Dim pColl As New Collection
  pColl.Add "Esri Access Workspace Factory", "{DD48C96A-D92A-11D1-AA81-00C04FA33A15}"
  pColl.Add "Workspace factory used to create workspace objects for ArcInfo coverages and Info tables", "{1D887452-D9F2-11D1-AA81-00C04FA33A15}"
  pColl.Add "Esri Cad Workspace Factory", "{9E2C27CE-62C6-11D2-9AED-00C04FA33299}"
  pColl.Add "Excel Workspace Factory", "{30F6F271-852B-4EE8-BD2D-099F51D6B238}"
  pColl.Add "FeatureService workspace factory", "{C81194E7-4DAA-418B-8C83-2942E65D2B8C}"
  pColl.Add "File GeoDatabase Workspace Factory", "{71FE75F0-EA0C-4406-873E-B7D53748AE7E}"
  pColl.Add "GeoRSS workspace factory", "{894AF6A1-157A-47BA-BDEC-3CF98D4CCE74}"
  pColl.Add "InMemory workspace factory", "{7F2BC55C-B902-43D0-A566-AA47EA9FDA2C}"
  pColl.Add "Esri LasDataset workspace-factory component", "{7AB01D9A-FDFE-4DFB-9209-86603EE9AEC6}"
  pColl.Add "OleDB Workspace Factory", "{59158055-3171-11D2-AA94-00C04FA37849}"
  pColl.Add "Esri PC ARC/INFO Workspace Factory", "{6DE812D2-9AB6-11D2-B0D7-0000F8780820}"
  pColl.Add "Provides access to members that control creation of raster workspaces", "{4C91D963-3390-11D2-8D25-0000F8780535}"
  pColl.Add "Esri SDC workspace factory", "{34DAE34F-DBE2-409C-8F85-DDBB46138011}"
  pColl.Add "Esri SDE Workspace Factory", "{D9B4FA40-D6D9-11D1-AA81-00C04FA33A15}"
  pColl.Add "Esri Shapefile Workspace Factory", "{A06ADB96-D95C-11D1-AA81-00C04FA33A15}"
  pColl.Add "Sql workspace factory", "{5297187B-FD2B-4A5F-8991-EB3F6F1CA502}"
  pColl.Add "Text File Workspace Factory", "{72CE59EC-0BE8-11D4-AE03-00C04FA33A15}"
  pColl.Add "Esri TIN workspace factory is used to access TINs on disk", "{AD4E89D9-00A5-11D2-B1CA-00C04F8EDEFF}"
  pColl.Add "Workspace Factory used to open toolbox workspaces", "{E9231B31-2A34-4729-8DE2-12CF39674B1B}"
  pColl.Add "Esri VPF Workspace Factory", "{397847F9-C865-11D3-9B56-00C04FA33299}"

  If CheckCollectionForKey(pColl, strClassID) Then
    ReturnWorkspaceFactoryType = pColl.Item(strClassID)
  Else
    ReturnWorkspaceFactoryType = "<- Unknown Workspace Factory Type ->"
  End If
  
  Set pColl = Nothing

End Function

Public Sub ApplyUniqueValueRenderer(pLayer As ILayer)
     
     '** Paste into VBA
     '** Creates a UniqueValuesRenderer and applies it to first layer in the map.
     '** Layer must have "Name" field
 
     Dim pApp As Application
     Dim pDoc As IMxDocument
     Set pDoc = ThisDocument
     Dim pMap As IMap
     Set pMap = pDoc.FocusMap
    
     Dim pFClass As IFeatureClass
    
     Dim pFLayer As IFeatureLayer
     Set pFLayer = pLayer
     Set pFClass = pFLayer.FeatureClass
     Dim pLyr As IGeoFeatureLayer
     Set pLyr = pFLayer
     
     Dim pFeatCls As IFeatureClass
     Set pFeatCls = pFLayer.FeatureClass
     Dim pQueryFilter As IQueryFilter
     Set pQueryFilter = New QueryFilter 'empty supports: SELECT *
     Dim pFeatCursor As IFeatureCursor
     Set pFeatCursor = pFeatCls.Search(pQueryFilter, False)
 
     '** Make the color ramp we will use for the symbols in the renderer
     Dim rx As IRandomColorRamp
     Set rx = New RandomColorRamp
     rx.MinSaturation = 20
     rx.MaxSaturation = 50
     rx.MinValue = 15
     rx.MaxValue = 70
     rx.StartHue = 1
     rx.EndHue = 360
     rx.UseSeed = True
     rx.Seed = 43
     
     '** Make the renderer
     Dim pRender As IUniqueValueRenderer, n As Long
     Set pRender = New UniqueValueRenderer
     
     Dim symd As ISimpleMarkerSymbol
     Set symd = New SimpleMarkerSymbol
     symd.Style = esriSMSCircle
     symd.size = 8
     
     Dim pSymDFill As ISimpleFillSymbol
     Set pSymDFill = New SimpleFillSymbol
     pSymDFill.Style = esriSFSSolid
     pSymDFill.Outline.Width = 0.5
     
     '** These properties should be set prior to adding values
     pRender.FieldCount = 1
     pRender.Field(0) = "Species"
     If pFClass.ShapeType = esriGeometryPolygon Then
       pRender.DefaultSymbol = pSymDFill
     Else
       pRender.DefaultSymbol = symd
     End If
     pRender.UseDefaultSymbol = False
     
     Dim pFeat As IFeature
     n = pFeatCls.FeatureCount(pQueryFilter)
     '** Loop through the features
     Dim i As Integer
     i = 0
     Dim ValFound As Boolean
     Dim NoValFound As Boolean
     Dim uh As Integer
     Dim pFields As IFields
     Dim iField As Integer
     Dim symx As ISimpleMarkerSymbol
     Dim pSymXFill As ISimpleFillSymbol
     Dim x As String
  
  
     Set pFields = pFeatCursor.Fields
     iField = pFields.FindField("Species")
     Do Until i = n
         If pFClass.ShapeType = esriGeometryPolygon Then
           Set pSymXFill = New SimpleFillSymbol
           pSymXFill.Style = esriSFSSolid
           pSymXFill.Outline.Width = 0.5
         Else
           Set symx = New SimpleMarkerSymbol
           symx.Style = esriSMSCircle
           symx.size = 8
         End If
         Set pFeat = pFeatCursor.NextFeature
         x = pFeat.Value(iField) '*new Cory*
         '** Test to see if we've already added this value
         '** to the renderer, if not, then add it.
         ValFound = False
         For uh = 0 To (pRender.ValueCount - 1)
           If pRender.Value(uh) = x Then
             NoValFound = True
             Exit For
           End If
         Next uh
         If Not ValFound Then
             If pFClass.ShapeType = esriGeometryPolygon Then
               pRender.AddValue x, "Species", pSymXFill
               pRender.Symbol(x) = pSymXFill
             Else
               pRender.AddValue x, "Species", symx
               pRender.Symbol(x) = symx
             End If
             pRender.Label(x) = x
         End If
         i = i + 1
     Loop
     
     '** now that we know how many unique values there are
     '** we can size the color ramp and assign the colors.
     rx.size = pRender.ValueCount
     rx.CreateRamp (True)
     Dim RColors As IEnumColors, ny As Long
     Set RColors = rx.Colors
     Dim pSymYFill As ISimpleFillSymbol
     Dim jsy As ISimpleMarkerSymbol
     Dim xv As String
     RColors.Reset
     For ny = 0 To (pRender.ValueCount - 1)
         xv = pRender.Value(ny)
         If xv <> "" Then
             If pFClass.ShapeType = esriGeometryPolygon Then
               Set pSymYFill = pRender.Symbol(xv)
               pSymYFill.Color = RColors.Next
               pRender.Symbol(xv) = pSymYFill
             Else
               Set jsy = pRender.Symbol(xv)
               jsy.Color = RColors.Next
               pRender.Symbol(xv) = jsy
             End If
         End If
     Next ny
 
     '** If you didn't use a color ramp that was predefined
     '** in a style, you need to use "Custom" here, otherwise
     '** use the name of the color ramp you chose.
     pRender.ColorScheme = "Custom"
     pRender.fieldType(0) = True
     Set pLyr.Renderer = pRender
     pLyr.DisplayField = "Species"
 
     '** This makes the layer properties symbology tab show
     '** show the correct interface.
     Dim hx As IRendererPropertyPage
     Set hx = New UniqueValuePropertyPage
     pLyr.RendererPropertyPageClassID = hx.ClassID
     
ClearMemory:
  Set pApp = Nothing
  Set pDoc = Nothing
  Set pMap = Nothing
  Set pFClass = Nothing
  Set pFLayer = Nothing
  Set pLyr = Nothing
  Set pFeatCls = Nothing
  Set pQueryFilter = Nothing
  Set pFeatCursor = Nothing
  Set rx = Nothing
  Set symd = Nothing
  Set pSymDFill = Nothing
  Set pFeat = Nothing
  Set pFields = Nothing
  Set symx = Nothing
  Set pSymXFill = Nothing
  Set pSymYFill = Nothing
  Set jsy = Nothing
  Set hx = Nothing



End Sub
Public Function CreateShapefileFeatureClass3(pWS As IWorkspace, sName As String, pSpatialReference As ISpatialReference, _
    pGeomType As esriGeometryType, Optional pAddFields As esriSystem.IVariantArray, _
    Optional booForceUniqueIDField As Boolean = True, Optional booHasZ As Boolean = False, _
    Optional booHasM As Boolean = False) As IFeatureClass                                                  ' Don't include filename!
    
  ' SET GEOMETRY TYPE, AND EXIT IF NOT ONE OF STANDARD OPTIONS
  Dim pGeomDef As IGeometryDef
  Dim pGeomDefEdit As IGeometryDefEdit
  Set pGeomDef = New GeometryDef
  Set pGeomDefEdit = pGeomDef
  With pGeomDefEdit
    .GeometryType = pGeomType
    .HasM = booHasM
    .HasZ = booHasZ
    Set .SpatialReference = pSpatialReference
  End With
  
  ' Open the folder to contain the shapefile as a workspace
  Dim pFWS As IFeatureWorkspace
  
  Set pFWS = pWS
  
'   Set up a simple fields collection
  Dim pFields As IFields
  Dim pFieldsEdit As IFieldsEdit
  Set pFields = New Fields
  Set pFieldsEdit = pFields

  Dim pField As iField
  Dim pFieldEdit As IFieldEdit
  
  ' Make the shape field
  ' it will need a geometry definition, with a spatial reference
  Set pField = New Field
  Set pFieldEdit = pField
  pFieldEdit.Name = "Shape"
  pFieldEdit.Type = esriFieldTypeGeometry

  Set pFieldEdit.GeometryDef = pGeomDef
  pFieldsEdit.AddField pField
  
  ' MAKE SURE UNIQUE ID FIELD IS UNIQUELY NAMED
 
  If booForceUniqueIDField Then
    Dim strUniqueName As String
    Dim strUniqueBase As String
    strUniqueName = "Unique_ID"
    strUniqueBase = strUniqueName
    Dim lngUniqueCounter As Long
    lngUniqueCounter = 1
    Dim lngCounterLength As Long
    Dim lngUniqueIndex As Long
    Dim booNameExists As Boolean
    Dim pCheckField As iField
    booNameExists = True
    If (pAddFields Is Nothing Or IsMissing(pAddFields)) Then
      booNameExists = False
    Else
      Do Until booNameExists = False
        booNameExists = False
        For lngUniqueIndex = 0 To pAddFields.Count - 1
          Set pCheckField = pAddFields.Element(lngUniqueIndex)
          If pCheckField.Name = strUniqueName Then
            booNameExists = True
            Exit For
          End If
        Next lngUniqueIndex
        If booNameExists Then
          lngUniqueCounter = lngUniqueCounter + 1
          lngCounterLength = Len(CStr(lngUniqueCounter))
          strUniqueName = Left(strUniqueBase, 10 - lngCounterLength) & CStr(lngUniqueCounter)
        End If
      Loop
    End If
    
    Set pField = New Field
    Set pFieldEdit = pField
    With pFieldEdit
        .Precision = 8
        .Name = strUniqueName
        .Type = esriFieldTypeInteger
    End With
    pFieldsEdit.AddField pField
  End If
  
  ' ADD FIELDS IF REQUESTED
'  Dim strDebugReport As String
'  Dim pDebugField As IField
  If Not (pAddFields Is Nothing Or IsMissing(pAddFields)) Then
    Dim lngIndex As Long
    For lngIndex = 0 To pAddFields.Count - 1
'      Set pDebugField = pAddFields.Element(lngIndex)
'      strDebugReport = strDebugReport & CStr(lngIndex) & "]  Field Name = " & pDebugField.Name & vbCrLf
      pFieldsEdit.AddField pAddFields.Element(lngIndex)
    Next lngIndex
  End If
'  MsgBox CStr(pGeomDef.GeometryType) & vbCrLf & strDebugReport
  ' Create the shapefile
  ' (some parameters apply to geodatabase options and can be defaulted as Nothing)
  Dim booFileExists As Boolean
  Dim strCheckString As String
  If Right(sPath, 1) = "\" Then
    strCheckString = sPath & sName & ".shp"
'    MsgBox sPath & sName & ".shp" & vbCrLf & "File Exists? " & CStr(Dir(sPath & sName & ".shp") <> "")
  Else
    strCheckString = sPath & "\" & sName & ".shp"
'    MsgBox sPath & "\" & sName & ".shp" & vbCrLf & "File Exists? " & CStr(Dir(sPath & "\" & sName & ".shp") <> "")
  End If
  
'  booFileExists = (Dir(strCheckString) <> "")
'  MsgBox strCheckString & vbCrLf & CStr(booFileExists)
  
  If booFileExists Then
    MsgBox "The following file already exists:" & vbCrLf & vbCrLf & strCheckString & vbCrLf & vbCrLf & _
           "Please select a new filename...", , "Duplicate Filename:"
    Set CreateShapefileFeatureClass3 = Nothing
    Exit Function
  End If
  
  Dim pFeatClass As IFeatureClass
  Set pFeatClass = pFWS.CreateFeatureClass(sName, pFields, Nothing, _
                                           Nothing, esriFTSimple, "Shape", "")
                                           
  Set CreateShapefileFeatureClass3 = pFeatClass


  GoTo ClearMemory

ClearMemory:
  Set pGeomDef = Nothing
  Set pGeomDefEdit = Nothing
  Set pFWS = Nothing
  Set pFields = Nothing
  Set pFieldsEdit = Nothing
  Set pField = Nothing
  Set pFieldEdit = Nothing
  Set pCheckField = Nothing
  Set pFeatClass = Nothing

End Function

Public Function CopyFeatureClass(pFClass As IFeatureClass, pNewWS As IWorkspace, _
    Optional strNewName As String = "", Optional booOverwrite As Boolean = False, _
    Optional strReasonForFail As String, Optional booWorked As Boolean) As IFeatureClass
  
  Dim pMxDoc As IMxDocument
  Dim pApp As IApplication
  Dim pSBar As IStatusBar
  Dim pProg As IStepProgressor
  
  Set pMxDoc = ThisDocument
  Set pApp = Application
  Set pSBar = pApp.StatusBar
  Set pProg = pSBar.ProgressBar
  
  Dim lngFeatureCount As Long
  Dim lngCounter As Long
  
  Dim pDataset As IDataset
  Dim pGeoDataset As IGeoDataset
  Dim pNewDataset As IDataset
  Dim pNewFeatWS As IFeatureWorkspace
  Dim booOutIsShapefile As Boolean
  
  booOutIsShapefile = MyGeneralOperations.ReturnWorkspaceFactoryType( _
      pNewWS.WorkspaceFactory.GetClassID) = "{A06ADB96-D95C-11D1-AA81-00C04FA33A15}"
  
  Set pDataset = pFClass
  Set pGeoDataset = pFClass
  
  If strNewName = "" Then strNewName = pDataset.BrowseName
  
  If MyGeneralOperations.CheckIfFeatureClassExists(pWS, strNewName) Then
    If booOverwrite Then
      Set pNewDataset = pNewFeatWS.OpenFeatureClass(strNewName)
      If pNewDataset.CanDelete Then
        pNewDataset.DELETE
      Else
        strReasonForFail = "Failed to delete existing Feature Class named '" & strNewName & "'."
        booWorked = False
        Set CopyFeatureClass = Nothing
        GoTo ClearMemory
      End If
    Else
      strReasonForFail = "Feature Class named '" & strNewName & "' already exists!"
      booWorked = False
      Set CopyFeatureClass = Nothing
      GoTo ClearMemory
    End If
  End If

  Dim lngIndex As Long
  
  ' varFieldIndexArray WILL HAVE 4 COLUMNS AND ANY NUMBER OR ROWS.
  ' COLUMN 0 = SOURCE FIELD NAME
  ' COLUMN 1 = SOURCE FIELD INDEX
  ' COLUMN 2 = NEW FIELD NAME
  ' COLUMN 3 = NEW FIELD INDEX
  Dim varFieldIndexes() As Variant
  Set pNewFClass = MyGeneralOperations.ReturnEmptyFClassWithSameSchema(pFClass, pNewWS, varFieldIndexes, strNewName, True)
    
  lngFeatureCount = pFClass.FeatureCount(Nothing)
  lngCounter = 0
  
  Dim pSrcFCursor As IFeatureCursor
  Dim pSrcFeature As IFeature
  Dim pDestFCursor As IFeatureCursor
  Dim pDestFBuffer As IFeatureBuffer
  
  Set pSrcFCursor = pFClass.Search(Nothing, False)
  Set pSrcFeature = pSrcFCursor.NextFeature
  Set pDestFCursor = pNewFClass.Insert(True)
  Set pDestFBuffer = pNewFClass.CreateFeatureBuffer
  
  pSBar.ShowProgressBar "Copying '" & pDataset.BrowseName & "' to '" & strNewName & "'...", 0, lngfeaturecou, 1, True
  pProg.position = 0
  
  Do Until pSrcFeature Is Nothing
    pProg.Step
    lngCounter = lngCounter + 1
    If lngCounter Mod 500 = 0 Then
      pDestFCursor.Flush
      pProg.Message = "Copying '" & pDataset.BrowseName & "' to '" & strNewName & "'...[" & _
          Format(lngCounter, "#,##0") & " of " & Format(lngFeatureCount, "#,##0") & "]"
      DoEvents
    End If
    Set pDestFBuffer.Shape = pSrcFeature.ShapeCopy
    For lngIndex = 0 To UBound(varFieldIndexes, 2)  ' WILL FAIL IF WRITING NULL VALUES TO SHAPEFILE
      pDestFBuffer.Value(varFieldIndexes(3, lngIndex)) = pSrcFeature.Value(varFieldIndexes(1, lngIndex))
    Next lngIndex
    
    pDestFCursor.InsertFeature pDestFBuffer
    Set pSrcFeature = pSrcFCursor.NextFeature
  Loop
  
  pDestFCursor.Flush
  pProg.position = 0
  pSBar.HideProgressBar
  
  Set CopyFeatureClass = pNewFClass
  booWorked = True
  strReasonForFail = ""

ClearMemory:
  Set pMxDoc = Nothing
  Set pApp = Nothing
  Set pSBar = Nothing
  Set pProg = Nothing
  Set pDataset = Nothing
  Set pGeoDataset = Nothing
  Set pNewDataset = Nothing
  Set pNewFeatWS = Nothing
  Erase varFieldIndexes
  Set pSrcFCursor = Nothing
  Set pSrcFeature = Nothing
  Set pDestFCursor = Nothing
  Set pDestFBuffer = Nothing




End Function
Public Function FindFieldCreateIfNecessary(strFieldName As String, pTable As ITable, _
    Optional booCreateIfMissing As Boolean = False, _
    Optional lngFieldType As esriFieldType = -999, Optional strAlias As String = "", Optional lngPrecision As Long = -999, _
    Optional lngScale As Long = -999, Optional lngLength As Long = -999) As Long
    

  Dim pField As iField
  Dim pFieldEdit As IFieldEdit
  Dim lngReturn As Long
  
  lngReturn = pTable.FindField(strFieldName)
  
  If lngReturn = -1 And booCreateIfMissing Then
    If lngFieldType = -999 Then
      MsgBox "Field must be created, but no field type specified!" & vbCrLf & "Returning -1 value..."
      FindFieldCreateIfNecessary = -1
      GoTo ClearMemory
    End If
    
    Set pField = New Field
    Set pFieldEdit = pField
    With pFieldEdit
      .Name = strFieldName
      If strAlias <> "" Then .AliasName = strAlias
      .Type = lngFieldType
      If lngPrecision <> -999 Then .Precision = lngPrecision
      If lngScale <> -999 Then .Scale = lngScale
      If lngLength <> -999 Then .length = lngLength
        
    End With
    pTable.AddField pField
    lngReturn = pTable.FindField(strFieldName)
  End If
  
  FindFieldCreateIfNecessary = lngReturn
  
ClearMemory:
  Set pField = Nothing
  Set pFieldEdit = Nothing

End Function
Public Function CreateNewField(strFieldName As String, lngType As esriFieldType, _
    Optional strAlias As String = "", Optional lngLength As Long = -999, _
    Optional dblPrecision As Double = -999, Optional dblScale As Double = -999) As iField
    
  Dim pReturn As iField
  Dim pFieldEdit As IFieldEdit
  
  Set pReturn = New Field
  Set pFieldEdit = pReturn
  With pFieldEdit
    .Name = strFieldName
    If strAlias = "" Then .AliasName = strFieldName
    .Type = lngType
    If lngType = esriFieldTypeString Then
      If lngLength = -999 Then
        .length = 255
      Else
        .length = lngLength
      End If
    End If
    If lngType = esriFieldTypeDouble Then
      If dblPrecision <> -999 Then .Precision = dblPrecision
      If dblScale <> -999 Then .Scale = dblScale
    End If
  End With
        
  
  Set CreateNewField = pReturn
  Set pReturn = Nothing
  Set pFieldEdit = Nothing
    
End Function


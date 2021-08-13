Attribute VB_Name = "Metadata_Functions"
Option Explicit

' MyGeneralOperations
' Jeff Jenness
' Jenness Enterprises
' http://www.jennessent.com
'
'                           AddCitationDates: (Public Function) Adds Date Created, Date Published and Date Revised.  Date Published is required by FGDC standard
'        AddContact_CitationResponsibleParty: (Public Function) Adds Citation Responsible Party; ORIGINATOR role required by FGDC standard
'                        AddContact_Metadata: (Public Function) Adds Metadata Contact Name and Role, required by FGDC standard
'          AddContact_ResourcePointOfContact: (Public Function) Adds Resource Point of Contact, required by FGDC standard
'              AddDetailsForObjectDefinition: (Public Function) Sets the Object Definition in the "Fields" section, which describes
'                                              briefly what the objects contain.  I'm not sure if this only applies to feature classes
'                         AddFieldAttributes: (Public Function) Adds attribute field descriptions
'                  AddMetadataUseLimitations: (Public Function) Sets or adds Use Limitations.
'                          AddNewGeoProcStep: (Public Function) Adds new geoprocessing step
'                          AddNewLineageStep: (Public Function) Adds new lineage step
'                   AddResourceDetailsStatus: (Public Function) Adds Resource Detains Status (Completed, Historical Archive, Ongoing, Proposed, etc.)
'                     AddResourceMaintenance: (Public Function) Adds Maintenance Interval (Continual, Daily, Weekly, Irregular, Not Planned, Unknown, etc.)
'               CheckIfContactAlreadyPresent: (Private Function) Internal function to see if a specified contact already exists.
' CopyAttributesFromSourceToDestinationField: (Public Function) Copies field descriptions from an existing field to a new field.
'                CopyLineageOrGeoprocHistory: (Public Function) Copies all lineage steps or geoprocessing steps from a source dataset to this new one.
'                 ExtractAttributesFromField: (Public Function) Fills in several field metadata attributes in string variables.
'         InsertExistingGeoProcessingHistory: (Private Function) Internal function to splice in a set of geoprocessing steps into new metadata.
'               InsertExistingLineageHistory: (Private Function) Internal function to splice in a lineage history into new metadata
'                          ReturnAddressType: (Private Function) Internal function to return address code from enumeration (Physical, Postal, Both)
'     ReturnAllMetadataPropertiesFromDataset: (Public Function) Given a dataset, returns text string listing all XPath properties in metadata.
'             ReturnAttributeFieldXPathIndex: (Private Function) Internal function to return the XPath Index value of a specified field.
'               ReturnDataElementFromDataset: (Public Function) Create a Data Element from a Dataset
'ReturnExistingLineageOrGeoprocessingHistory: (Private Function) Internal function to copy existing lineage or geoprocessing history from a dataset
'             ReturnExistingMetadataKeyWords: (Public Function) Extract keywords from an existing dataset, in several keyword classes.
'                 ReturnGxDatasetFromDataset: (Public Function) Create a GXDataset from a Dataset.
'                    ReturnLargestIndexValue: (Private Function) Internal function to find the highest current XPath Index
'                      ReturnMaintenanceCode: (Private Function) Internal function to return Maintenance Code from enumeration (Continual, Daily, etc.)
'           ReturnMetadataPropSetFromDataset: (Public Function) Mostly internal function to return an IPropSet of dataset metadata.
'         ReturnMetadataXMLStringFromDataset: (Public Function) Given a dataset, returns text string containing all XML data.
'                  ReturnPropertyFromPropSet: (Private Function) Used with "ExtractAttributesFromField":  Returns the string property value or "" if none found.
'                         ReturnRoleCDString: (Private Function) Internal Function to return Contact Role Code from enumeration (Author, Publisher, Originator, etc.)
'                         ReturnStatusString: (Private Function) Internal function to return Metadata Status code from enumeration (Completed, Ongoing, etc.)
'                               SaveMetadata: (Public Sub) Saves current metadata
'                        SetMetadataAbstract: (Public Function) Adds or Replaces Metadata Abstract/Description
'                         SetMetadataCredits: (Public Function) Sets metadata credits.
'                   SetMetadataFormatVersion: (Public Function) Sets current version; optional method to save current ArcGIS version.
'                        SetMetadataKeyWords: (Public Function) Sets metadata keywords
'                         SetMetadataPurpose: (Public Function) Sets metadata purpose/Summary
'                 SynchronizeMetadataPropSet: (Public Function) Synchronizes/regenerates metadata with current version of data.  Use if you have added an attribute
'                                             field after initial metadata creation and before setting attributes of field.


Public Enum JenMetadataRoleCDValues
  JenMetadata_ResourceProvider = 1
  JenMetadata_Custodian = 2
  JenMetadata_Owner = 3
  JenMetadata_User = 4
  JenMetadata_Distributor = 5
  JenMetadata_Originator = 6
  JenMetadata_PointOfContact = 7
  JenMetadata_PrincipalInvestigator = 8
  JenMetadata_Processor = 9
  JenMetadata_Publisher = 10
  JenMetadata_Author = 11
  JenMetadata_Collaborator = 12
  JenMetadata_Editor = 13
  JenMetadata_Mediator = 14
  JenMetadata_RightsHolder = 15
End Enum

Public Enum JenMetadataAddressTypeValues
  JenMetadata_Postal = 1
  JenMetadata_Physical = 2
  JenMetadata_both = 3
  JenMetadata_Skip = 4
End Enum

Public Enum JenMetadataLineageOrGeoprocValues
  JenMetadata_Lineage = 1
  JenMetadata_GeoprocessingHistory = 2
End Enum

Public Enum JenMetadataStatusValues
  JenMetadata_Completed = 1
  JenMetadata_HistoricalArchive = 2
  JenMetadata_Obsolete = 3
  JenMetadata_Ongoing = 4
  JenMetadata_Planned = 5
  JenMetadata_Required = 6
  JenMetadata_UnderDevelopment = 7
  JenMetadata_Proposed = 8
End Enum

Public Enum JenMetadataMaintenanceCodes
  JenMetadata_Maint_Continual = 1
  JenMetadata_Maint_Daily = 2
  JenMetadata_Maint_Weekly = 3
  JenMetadata_Maint_Fortnightly = 4
  JenMetadata_Maint_Monthly = 5
  JenMetadata_Maint_Quarterly = 6
  JenMetadata_Maint_BiAnnually = 7
  JenMetadata_Maint_Annually = 8
  JenMetadata_Maint_AsNeeded = 9
  JenMetadata_Maint_Irregular = 10
  JenMetadata_Maint_NotPlanned = 11
  JenMetadata_Maint_Unknown = 12
  JenMetadata_Maint_SemiMonthly = 13
End Enum

Public Function SampleCodeToSetMetadataProperties()
  
  '-----------------------------------------------------------------------------------------------------------------------
  ' WRITE METADATA =======================================================================================================
  '-----------------------------------------------------------------------------------------------------------------------
  
  Dim pFClass As IFeatureClass
  Dim pWS As IFeatureWorkspace
  Dim pWSFact As IWorkspaceFactory
  Set pWSFact = New ShapefileWorkspaceFactory
  Set pWS = pWSFact.OpenFromFile("D:\arcGIS_stuff\consultation\az_linkages\Crystal_Krause_Btlnck_Batch\Outputs\Output_Folder", 0)
  Set pFClass = pWS.OpenFeatureClass("abc_Def")
  
  Dim pDataset As IDataset
  Set pDataset = pFClass
  
  Debug.Print "------------------------------"
  Debug.Print "Feature Class Name = " & pDataset.BrowseName
  
  Dim pPropSet As IPropertySet
  Set pPropSet = Metadata_Functions.ReturnMetadataPropSetFromDataset(pDataset)
  
  ' SYNCHRONIZE METADATA
  Dim strResponse As String
  strResponse = Metadata_Functions.SynchronizeMetadataPropSet(pDataset)
  Debug.Print "Synchronization: " & strResponse
  
  Dim strAbstract As String
  strAbstract = "This dataset represents points along the bottleneck route.  This bottleneck route " & _
    "describes the path between the two habitat blocks 'aaa' and 'bbb', within the corridor polygon 'ccc'., " & _
    "which follows the route with the widest narrow point."
  strResponse = Metadata_Functions.SetMetadataAbstract(pDataset, strAbstract)
  Debug.Print "Saving Abstract: " & strResponse
    
  Dim strPurpose As String
  strPurpose = "Point dataset of points along route with widest bottleneck, with corridor width values at each point."
  strResponse = Metadata_Functions.SetMetadataPurpose(pDataset, strPurpose)
  Debug.Print "Saving Purpose: " & strResponse
  
  Dim pKeyWords As esriSystem.IStringArray
  Dim pIncludeThemeKeys As esriSystem.IStringArray
  Dim pIncludeSearchKeys As esriSystem.IStringArray
  Dim pIncludeDescKeys As esriSystem.IStringArray
  Dim pIncludeStratKeys As esriSystem.IStringArray
  Dim pIncludeThemeSlashThemekeys As esriSystem.IStringArray
  Dim pIncludePlaceKeys As esriSystem.IStringArray
  Dim pIncludeTemporalKeys As esriSystem.IStringArray
  
  Set pIncludeThemeKeys = New esriSystem.strArray
  Set pIncludeSearchKeys = New esriSystem.strArray
  Set pIncludeDescKeys = New esriSystem.strArray
  Set pIncludeStratKeys = New esriSystem.strArray
  Set pIncludeThemeSlashThemekeys = New esriSystem.strArray
  Set pIncludePlaceKeys = New esriSystem.strArray
  Set pIncludeTemporalKeys = New esriSystem.strArray
  
  pIncludeThemeSlashThemekeys.Add "theme_slash_theme"
  pIncludeTemporalKeys.Add "temporal"
  
  Dim pCombinedKeyWords As esriSystem.IStringArray
  Dim pFLayer As IFeatureLayer
  Dim pMxDoc As IMxDocument
  ' Set pMxDoc = ThisDocument
  Set pFLayer = MyGeneralOperations.ReturnLayerByName("mroads", pMxDoc.FocusMap)
  Dim booSucceeded As Boolean
  Dim lngIndex As Long
  Set pCombinedKeyWords = Metadata_Functions.ReturnExistingMetadataKeyWords(pFLayer.FeatureClass, _
      pKeyWords, booSucceeded, pIncludeThemeKeys, pIncludeSearchKeys, pIncludeDescKeys, pIncludeStratKeys, _
      pIncludeThemeSlashThemekeys, _
      pIncludePlaceKeys, pIncludeTemporalKeys)
  Debug.Print "Extracting keywords: " & UCase(CStr(booSucceeded))
'  If booSucceeded Then
'    Debug.Print "Combined..."
'    If pCombinedKeyWords.Count > 0 Then
'      For lngIndex = 0 To pCombinedKeyWords.Count - 1
'        Debug.Print "  --> " & CStr(lngIndex + 1) & "] " & pCombinedKeyWords.Element(lngIndex)
'      Next lngIndex
'    Else
'      Debug.Print "Nothing in 'pCombinedKeyWords'..."
'    End If
'    Debug.Print "Theme..."
'    If pIncludeThemeKeys.Count > 0 Then
'      For lngIndex = 0 To pIncludeThemeKeys.Count - 1
'        Debug.Print "  --> " & CStr(lngIndex + 1) & "] " & pIncludeThemeKeys.Element(lngIndex)
'      Next lngIndex
'    Else
'      Debug.Print "-- Nothing in 'pIncludeThemeKeys'..."
'    End If
'    Debug.Print "Search..."
'    If pIncludeSearchKeys.Count > 0 Then
'      For lngIndex = 0 To pIncludeSearchKeys.Count - 1
'        Debug.Print "  --> " & CStr(lngIndex + 1) & "] " & pIncludeSearchKeys.Element(lngIndex)
'      Next lngIndex
'    Else
'      Debug.Print "-- Nothing in 'pIncludeSearchKeys'..."
'    End If
'    Debug.Print "Desc..."
'    If pIncludeDescKeys.Count > 0 Then
'      For lngIndex = 0 To pIncludeDescKeys.Count - 1
'        Debug.Print "  --> " & CStr(lngIndex + 1) & "] " & pIncludeDescKeys.Element(lngIndex)
'      Next lngIndex
'    Else
'      Debug.Print "-- Nothing in 'pIncludeDescKeys'..."
'    End If
'    Debug.Print "Strat..."
'    If pIncludeStratKeys.Count > 0 Then
'      For lngIndex = 0 To pIncludeStratKeys.Count - 1
'        Debug.Print "  --> " & CStr(lngIndex + 1) & "] " & pIncludeStratKeys.Element(lngIndex)
'      Next lngIndex
'    Else
'      Debug.Print "-- Nothing in 'pIncludeStratKeys'..."
'    End If
'    Debug.Print "ThemeSlashTheme..."
'    If pIncludeThemeSlashThemekeys.Count > 0 Then
'      For lngIndex = 0 To pIncludeThemeSlashThemekeys.Count - 1
'        Debug.Print "  --> " & CStr(lngIndex + 1) & "] " & pIncludeThemeSlashThemekeys.Element(lngIndex)
'      Next lngIndex
'    Else
'      Debug.Print "-- Nothing in 'pIncludeThemeSlashThemekeys'..."
'    End If
'    Debug.Print "Place..."
'    If pIncludePlaceKeys.Count > 0 Then
'      For lngIndex = 0 To pIncludePlaceKeys.Count - 1
'        Debug.Print "  --> " & CStr(lngIndex + 1) & "] " & pIncludePlaceKeys.Element(lngIndex)
'      Next lngIndex
'    Else
'      Debug.Print "-- Nothing in 'pIncludePlaceKeys'..."
'    End If
'    Debug.Print "Temporal..."
'    If pIncludeTemporalKeys.Count > 0 Then
'      For lngIndex = 0 To pIncludeTemporalKeys.Count - 1
'        Debug.Print "  --> " & CStr(lngIndex + 1) & "] " & pIncludeTemporalKeys.Element(lngIndex)
'      Next lngIndex
'    Else
'      Debug.Print "-- Nothing in 'pIncludeTemporalKeys'..."
'    End If
'  End If
      
  
  pIncludeThemeKeys.Add "Bottleneck"
  pIncludeThemeKeys.Add "Width"
  pIncludeThemeKeys.Add "Corridor"
  
  For lngIndex = 0 To pIncludeThemeKeys.Count - 1
    pIncludeSearchKeys.Add pIncludeThemeKeys.Element(lngIndex)
    pIncludeDescKeys.Add pIncludeThemeKeys.Element(lngIndex)
  Next lngIndex
  
  strResponse = Metadata_Functions.SetMetadataKeyWords(pDataset, pIncludeThemeKeys, pIncludeSearchKeys, _
        pIncludeDescKeys, pIncludeStratKeys, pIncludeThemeSlashThemekeys, pIncludePlaceKeys, pIncludeTemporalKeys)
  Debug.Print "Saving Keywords: " & strResponse
  
  Dim strAllProps As String
  Dim strXML As String
  
  strXML = Metadata_Functions.ReturnMetadataXMLStringFromDataset(pDataset)
  strAllProps = Metadata_Functions.ReturnAllMetadataPropertiesFromDataset(pDataset)
  
  Clipboard.Clear
  Clipboard.SetText strXML
  Clipboard.SetText strAllProps
  
'  Dim DataObj As New MSForms.DataObject
''  DataObj.SetText strXML
'  DataObj.SetText strAllProps
'  DataObj.PutInClipboard
'  Set DataObj = Nothing
  
  
  ' SET PROCESS STEP
  Dim strDescription As String
  Dim strName As String
  strDescription = "Generated points along bottleneck route."
  strName = "Sample..." & aml_func_mod.GetTheUserName
  strResponse = Metadata_Functions.AddNewLineageStep(pDataset, strDescription, Now, JenMetadata_Processor, _
      strName, "Sample Organization Name", "Sample Position Name", "1-234-567-8901", _
      "Sample Street", "Sample City", "Sample State", "Sample Zip", "Sample Country", "me@jennessent.com", _
      JenMetadata_Physical)
  Debug.Print "Lineage Successful = " & strResponse
  
  ' REPLACE PROCESS STEP
  Dim pSourceFClass As IFeatureClass
  Set pFLayer = MyGeneralOperations.ReturnLayerByName("Roads_for_Metadata", pMxDoc.FocusMap)
  Set pSourceFClass = pFLayer.FeatureClass
'  strResponse = CopyLineageOrGeoprocHistory(pSourceFClass, pDataset, True, JenMetadata_Lineage)
'  Debug.Print "Lineage Transfer Successful = " & strResponse
  
  ' REPLACE GEOPROCESSING HISTORY
'  strResponse = Metadata_Functions.ACopyLineageOrGeoprocHistory(pSourceFClass, pDataset, False, JenMetadata_GeoprocessingHistory)
'  Debug.Print "Geoproc History Transfer Successful = " & strResponse
  
  ' ADD NEW GEOPROCESSING EVENT
  strResponse = Metadata_Functions.AddNewGeoProcStep(pDataset, "Details on bottleneck analysis...", "c:/BatchBottleneck tool", _
      Now, "Corridor Designer Extension Batch Bottleneck Tool", False)
  Debug.Print "Added new Geoprocessing Event = " & strResponse
  
  ' ADD NEW METADATA CONTACTS
  strResponse = Metadata_Functions.AddContact_Metadata(pDataset, JenMetadata_PointOfContact, _
    "Jeff1", False, "Jenness1", "Analyst1", "111-222-3345", "3020 N. Schevene", "Flagstaff", "AZ", _
    "86004", "USA", "me1@jennessent.com", JenMetadata_Postal)
  Debug.Print "Added Metadata Contact 'Point of Contact' = " & strResponse
  strResponse = Metadata_Functions.AddContact_Metadata(pDataset, JenMetadata_Processor, _
    "Jeff1", False, "Jenness1", "Analyst1", "111-222-3345", "3020 N. Schevene", "Flagstaff", "AZ", _
    "86004", "USA", "me1@jennessent.com", JenMetadata_Postal)
  Debug.Print "Added Metadata Contact 'Processor' = " & strResponse
  strResponse = Metadata_Functions.AddContact_Metadata(pDataset, JenMetadata_Custodian, _
    "Jeff5", True, "Jenness1", "Analyst1", "111-222-3345", "3020 N. Schevene", "Flagstaff", "AZ", _
    "86004", "USA", "me1@jennessent.com", JenMetadata_Postal)
  Debug.Print "Added Metadata Contact 'Custodian' = " & strResponse
  strResponse = Metadata_Functions.AddContact_Metadata(pDataset, JenMetadata_Author, _
    "Jeff5", False, "Jenness1", "Analyst1", "111-222-3345", "3020 N. Schevene", "Flagstaff", "AZ", _
    "86004", "USA", "me1@jennessent.com", JenMetadata_Postal)
  Debug.Print "Added Metadata Contact 'Author' = " & strResponse
  strResponse = Metadata_Functions.AddContact_Metadata(pDataset, JenMetadata_PointOfContact, _
    "Jeff1", False, "Jenness1", "Analyst1", "111-222-3345", "3020 N. Schevene", "Flagstaff", "AZ", _
    "86004", "USA", "me1@jennessent.com", JenMetadata_Physical)
  Debug.Print "Added Metadata Contact 'Point of Contact' with Physical address = " & strResponse
  strResponse = Metadata_Functions.AddContact_Metadata(pDataset, JenMetadata_PointOfContact, _
    "Jeff1", True, "Jenness1", "Analyst1", "111-222-3345", "3020 N. Schevene", "Flagstaff", "AZ", _
    "86004", "USA", "me1@jennessent.com", JenMetadata_both)
  Debug.Print "Added Metadata Contact 'Point of Contact' with Both address = " & strResponse
  
  ' ADD NEW CITATION CONTACTS; SHOULD HAVE ORIGINATOR
  strResponse = Metadata_Functions.AddContact_CitationResponsibleParty(pDataset, JenMetadata_PointOfContact, _
    "Jeff1", False, "Jenness1", "Analyst1", "111-222-3345", "3020 N. Schevene", "Flagstaff", "AZ", _
    "86004", "USA", "me1@jennessent.com", JenMetadata_Postal)
  Debug.Print "Added CITATION Contact 'Point of Contact' = " & strResponse
  strResponse = Metadata_Functions.AddContact_CitationResponsibleParty(pDataset, JenMetadata_Processor, _
    "Jeff1", False, "Jenness1", "Analyst1", "111-222-3345", "3020 N. Schevene", "Flagstaff", "AZ", _
    "86004", "USA", "me1@jennessent.com", JenMetadata_Postal)
  Debug.Print "Added CITATION Contact 'Processor' = " & strResponse
  strResponse = Metadata_Functions.AddContact_CitationResponsibleParty(pDataset, JenMetadata_Originator, _
    "Jeff1", True, "Jenness1", "Analyst1", "111-222-3345", "3020 N. Schevene", "Flagstaff", "AZ", _
    "86004", "USA", "me1@jennessent.com", JenMetadata_Postal)
  Debug.Print "Added CITATION Contact 'Custodian' = " & strResponse
  
  ' ADD NEW RESOURCE CONTACTS
  strResponse = Metadata_Functions.AddContact_ResourcePointOfContact(pDataset, JenMetadata_PointOfContact, _
    "Jeff1", True, "Jenness1", "Analyst1", "111-222-3345", "3020 N. Schevene", "Flagstaff", "AZ", _
    "86004", "USA", "me1@jennessent.com", JenMetadata_Postal)
  Debug.Print "Added RESOURCE Contact 'Point of Contact' = " & strResponse
  strResponse = Metadata_Functions.AddContact_ResourcePointOfContact(pDataset, JenMetadata_Processor, _
    "Jeff1", False, "Jenness1", "Analyst1", "111-222-3345", "3020 N. Schevene", "Flagstaff", "AZ", _
    "86004", "USA", "me1@jennessent.com", JenMetadata_Postal)
  Debug.Print "Added RESOURCE Contact 'Processor' = " & strResponse
  strResponse = Metadata_Functions.AddContact_ResourcePointOfContact(pDataset, JenMetadata_Custodian, _
    "Jeff1", False, "Jenness1", "Analyst1", "111-222-3345", "3020 N. Schevene", "Flagstaff", "AZ", _
    "86004", "USA", "me1@jennessent.com", JenMetadata_Postal)
  Debug.Print "Added RESOURCE Contact 'Custodian' = " & strResponse
  
  ' ADD CITATION DATES
  Dim datCreated As Date
  Dim datPublished As Date
  
  datCreated = Now
  datPublished = DateAdd("s", 1234567, datCreated)
  strResponse = Metadata_Functions.AddCitationDates(pDataset, datCreated, datPublished)
  Debug.Print "Added Citation Dates = " & strResponse
  
  ' SET RESOURCE STATUS
  strResponse = Metadata_Functions.AddResourceDetailsStatus(pDataset, JenMetadata_Ongoing)
  Debug.Print "Added Resource Status = " & strResponse
  
  ' SET MAINTENANCE STATUS
  strResponse = Metadata_Functions.AddResourceMaintenance(pDataset, JenMetadata_Maint_Daily)
  Debug.Print "Added Resource Maintenance = " & strResponse
  
  ' COPY OVER COMPLICATED FIELD INFO
  Dim pGCRoads As IFeatureClass
  Set pFLayer = MyGeneralOperations.ReturnLayerByName("Roads_for_Metadata", pMxDoc.FocusMap)
  Set pGCRoads = pFLayer.FeatureClass
  strResponse = Metadata_Functions.CopyAttributesFromSourceToDestinationField(pGCRoads, "BLM_OBS_USE1", _
      pDataset, "text_field")
  Debug.Print "Copy Complex Field = " & strResponse
  
  Dim pMRoads As IFeatureClass
  Set pFLayer = MyGeneralOperations.ReturnLayerByName("mroads", pMxDoc.FocusMap)
  Set pMRoads = pFLayer.FeatureClass
  strResponse = Metadata_Functions.CopyAttributesFromSourceToDestinationField(pMRoads, "FCC", _
      pDataset, "FCC")
  Debug.Print "Copy Complex Field 2 = " & strResponse
  
  ' ADD UDOM (Unrepresentable Domain) FIELD INFO
  strResponse = Metadata_Functions.AddFieldAttributes(pDataset, "forudom", _
      "Description of an 'Unrepresentable Domain' field...", "World Book Encyclopedia", _
      , , , , , , , "Values represent area of polygons as calculated by spheroidal means...")
  Debug.Print "Add Unrepresentable Domain Field info = " & strResponse
  
  ' ADD EDOM (Enumerated Domain) FIELD INFO
  Dim varData() As Variant
  ReDim varData(0)
  Dim strList() As String
  ReDim strList(2, 5)
  strList(0, 0) = "Item #1"
  strList(1, 0) = "Item #1 Description"
  strList(2, 0) = "Item #1 Source"
  strList(0, 1) = "Item #2"
  strList(1, 1) = "Item #2 Description"
  strList(2, 1) = "Item #2 Source"
  strList(0, 2) = "Item #3"
  strList(1, 2) = "Item #3 Description"
  strList(2, 2) = "Item #3 Source"
  strList(0, 3) = "Item #4"
  strList(1, 3) = "Item #4 Description"
  strList(2, 3) = "Item #4 Source"
  strList(0, 4) = "Item #5"
  strList(1, 4) = "Item #5 Description"
  strList(2, 4) = "Item #5 Source"
  strList(0, 5) = "Item #6"
  strList(1, 5) = "Item #6 Description"
  strList(2, 5) = "Item #6 Source"
  varData(0) = strList
  
  strResponse = Metadata_Functions.AddFieldAttributes(pDataset, "foredom", _
      "Description of an 'Enumerated Domain' field...", "World Book Encyclopedia", _
      , , , , , , varData)
  Debug.Print "Add Enumerated Domain Field info = " & strResponse
  
  ' ADD RDOM (Range Domain) FIELD INFO
  strResponse = Metadata_Functions.AddFieldAttributes(pDataset, "forrdom", _
      "Description of an 'Range Domain' field...", "World Book Encyclopedia", _
       "0", "100", "54", "Meters", "1.234", "1.111")
  Debug.Print "Add Enumerated Domain Field info = " & strResponse
  
  ' ADD CODESET FIELD INFO
  strResponse = Metadata_Functions.AddFieldAttributes(pDataset, "forcodeset", _
      "Description of an 'Codeset Domain' field...", "World Book Encyclopedia", _
      , , , , , , , , "Census Feature Classification Codes (also called ""FCC"")", _
      "Department of Commerce, Census Bureau")
  Debug.Print "Add Codeset Domain Field info = " & strResponse
  
  ' ADD ALL FOUR DOMAINS FIELD INFO
  strResponse = Metadata_Functions.AddFieldAttributes(pDataset, "foralldom", _
      "Description of an 'Range Domain' field...", "World Book Encyclopedia", _
      "0", "100", "54", "Meters", "1.234", "1.111", varData, _
      "Values represent area of polygons as calculated by spheroidal means...", _
      "Census Feature Classification Codes (also called ""FCC"")", _
      "Department of Commerce, Census Bureau", True)
  Debug.Print "Add All Three Domain Field info = " & strResponse
  
  ' SET METADATA FORMAT VERSION
  Dim strFormatVersion As String
  Dim lngVersion As Long
  lngVersion = aml_func_mod.ReturnArcGISVersionAlt2(pMxDoc, strFormatVersion)
'  strResponse = Metadata_Functions.SetMetadataFormatVersion(pDataset, "Created in ArcGIS " & strFormatVersion)
  strResponse = Metadata_Functions.SetMetadataFormatVersion(pDataset, , True, pMxDoc)
  Debug.Print "Added Format Version = " & strResponse
  
  
  ' RESYNCHRONIZE METADATA
  strResponse = Metadata_Functions.SynchronizeMetadataPropSet(pDataset)
  Debug.Print "ReSynchronization: " & strResponse
  
ClearMemory:
  Set pFClass = Nothing
  Set pWS = Nothing
  Set pWSFact = Nothing
  Set pDataset = Nothing
  Set pPropSet = Nothing
  Set pKeyWords = Nothing
  Set pIncludeThemeKeys = Nothing
  Set pIncludeSearchKeys = Nothing
  Set pIncludeDescKeys = Nothing
  Set pIncludeStratKeys = Nothing
  Set pIncludeThemeSlashThemekeys = Nothing
  Set pIncludePlaceKeys = Nothing
  Set pIncludeTemporalKeys = Nothing
  Set pCombinedKeyWords = Nothing
  Set pFLayer = Nothing
  Set pMxDoc = Nothing
  Set pSourceFClass = Nothing
  Set pGCRoads = Nothing
  Set pMRoads = Nothing
  Erase varData
  Erase strList


End Function

Public Function AddDetailsForObjectDefinition(pDataset As IDataset, strShortDefinition As String, _
    strDefinitionSource As String) As String

  On Error GoTo ErrHandler
  
'  strResponse = Metadata_Functions.AddDetailsForObjectDefinition(pDataset, _
    "Points representing corridor midpoints", "Corridor Designer")
'  Debug.Print "Added Resource Status = " & strResponse
  
  AddDetailsForObjectDefinition = "Succeeded"
  
  Dim pPropSet As IPropertySet
  Set pPropSet = ReturnMetadataPropSetFromDataset(pDataset)
  
  pPropSet.SetProperty "eainfo/detailed/enttyp/enttypd", strShortDefinition
  pPropSet.SetProperty "eainfo/detailed/enttyp/enttypds", strDefinitionSource
  
  
  Metadata_Functions.SaveMetadata pDataset, pPropSet
  
  GoTo ClearMemory
  Exit Function
  
ErrHandler:
  AddDetailsForObjectDefinition = "Failed"

ClearMemory:
  Set pPropSet = Nothing
  
End Function




Public Function CopyAttributesFromSourceToDestinationField(pSourceDataset As IDataset, _
  strSourceFieldName As String, pDestinationDataset As IDataset, strDestinationFieldName As String, _
  Optional booFailed As Boolean) As String
  

  On Error GoTo ErrHandler
  ' GENERALLY CALLED BY OTHER FUNCTIONS
  booFailed = False
  CopyAttributesFromSourceToDestinationField = "Succeeded"
  Dim pOrigPropSet As IPropertySet
  Set pOrigPropSet = ReturnMetadataPropSetFromDataset(pSourceDataset)
  Dim pNewPropSet As IPropertySet
  Set pNewPropSet = ReturnMetadataPropSetFromDataset(pDestinationDataset)
    
  ' GET SOURCE INDEX NUMBER
  Dim lngSourceFieldIndex As Long
  lngSourceFieldIndex = ReturnAttributeFieldXPathIndex(pSourceDataset, strSourceFieldName, booFailed)
  If lngSourceFieldIndex = -1 Then
    If booFailed Then
      CopyAttributesFromSourceToDestinationField = "ReturnAttributeFieldXPathIndex Failed"
    Else
      CopyAttributesFromSourceToDestinationField = "No Source Field Found"
    End If
    GoTo ClearMemory
  End If
    
  ' GET DESTINATION INDEX NUMBER
  Dim lngDestinationFieldIndex As Long
  lngDestinationFieldIndex = ReturnAttributeFieldXPathIndex(pDestinationDataset, strDestinationFieldName, booFailed)
  If lngDestinationFieldIndex = -1 Then
    If booFailed Then
      CopyAttributesFromSourceToDestinationField = "ReturnAttributeFieldXPathIndex Failed"
    Else
      CopyAttributesFromSourceToDestinationField = "No Destination Field Found"
    End If
    GoTo ClearMemory
  End If
  
  Dim strXOrigName As String
  Dim strXOrigNameSource As String
  Dim strXNewName As String
  Dim strXNewNameSource As String
  
'  Dim strXDomMin As String
'  Dim strXDomMax As String
'  Dim strXDomUnits As String
'  Dim strXDom As String
'  Dim strXUDom As String
'  Dim strXEDom As String
  
  ' GET CURRENT DESCRIPTION
  strXOrigName = "eainfo/detailed/attr[" & CStr(lngSourceFieldIndex) & "]/attrdef"         ' DESCRIPTION OF FIELD
  strXOrigNameSource = "eainfo/detailed/attr[" & CStr(lngSourceFieldIndex) & "]/attrdefs"   ' DESCRIPTION SOURCE
  strXNewName = "eainfo/detailed/attr[" & CStr(lngDestinationFieldIndex) & "]/attrdef"         ' DESCRIPTION OF FIELD
  strXNewNameSource = "eainfo/detailed/attr[" & CStr(lngDestinationFieldIndex) & "]/attrdefs"   ' DESCRIPTION SOURCE
  
'  strXDomMin = "eainfo/detailed/attr[" & CStr(lngFieldIndex) & "]/attrdomv/rdom/rdommin"  ' MINIMUM VALUE
'  strXDomMax = "eainfo/detailed/attr[" & CStr(lngFieldIndex) & "]/attrdomv/rdom/rdommax"  ' MAXIMUM VALUE
'  strXDomUnits = "eainfo/detailed/attr[" & CStr(lngFieldIndex) & "]/attrdomv/rdom/attrunit"   ' UNITS
'  strXDom = "eainfo/detailed/attr[" & CStr(lngFieldIndex) & "]/attrdomv"
'  strXUDom = "eainfo/detailed/attr[" & CStr(lngFieldIndex) & "]/attrdomv/udom"  ' DESCRIPTION OF VALUES
'  strXUDom = "eainfo/detailed/attr[" & CStr(lngFieldIndex) & "]/attrdomv/edom"  ' LIST OF VALUES
                                                                          ' /edomv = Value
                                                                          ' /edomvd = Description of Value
                                                                          ' /edomvds = Enumerated domain value definition source
  
  ' WRITE DESCRIPTION AND SOURCE
  Dim varVal As Variant
  Dim strDescription As String
  Dim strDescriptionSource As String
  
  varVal = pOrigPropSet.GetProperty(strXOrigName)
  If Not IsEmpty(varVal) Then
    strDescription = CStr(varVal(0))
  Else
    strDescription = "<-- No Description Found -->"
  End If
  varVal = pOrigPropSet.GetProperty(strXOrigNameSource)
  If Not IsEmpty(varVal) Then
    strDescriptionSource = CStr(varVal(0))
  Else
    strDescriptionSource = "<-- No Description Source Found -->"
  End If
  
  pNewPropSet.SetProperty strXNewName, strDescription
  pNewPropSet.SetProperty strXNewNameSource, strDescriptionSource
  
  ' SET DOMAIN INFO
  Dim strOrigFClassXName As String
  Dim strNewFClassXName As String
  Dim varProperty As Variant
  Dim varSubProperty As Variant
  Dim lngIndex3 As Long
  Dim lngIndex4 As Long
  Dim strSubXName1 As String
  Dim strSubXName2 As String
  Dim strSubXName3 As String
  Dim varSubProp1 As Variant
  Dim varSubProp2 As Variant
  Dim strValue As String
  Dim varCheckPropertyPresent As Variant
  
  varProperty = Null
  varSubProperty = Null
  varSubProp1 = Null
  varSubProp2 = Null
  varCheckPropertyPresent = Null
  varVal = Null
  
'                strOrigFCLassXName = "eainfo/detailed/attr[" & CStr(lngIndex) & "]/attrdomv"
'                varProperty = pBLM_AZStrip_UTM12_NAD83_PropSet.GetProperty(strOrigFCLassXName)
'                varSubProperty = varProperty(0)
'                For lngIndex2 = 0 To UBound(varProperty)
'                  Debug.Print "       " & CStr(lngIndex2) & "] " & CStr(varProperty(lngIndex2))
'                Next lngIndex2
              
  ' GET DOMAIN ATTRIBUTES FROM ORIGINAL FCLASS
  strOrigFClassXName = "eainfo/detailed/attr[" & CStr(lngSourceFieldIndex) & "]/attrdomv"
  strNewFClassXName = "eainfo/detailed/attr[" & CStr(lngDestinationFieldIndex) & "]/attrdomv"
  varProperty = pOrigPropSet.GetProperty(strOrigFClassXName)

  ' ADD DOMAIN ATTRIBUTES TO NEW FCLASS
  varCheckPropertyPresent = pNewPropSet.GetProperty(strNewFClassXName)
  If Not IsEmpty(varCheckPropertyPresent) Then
    pNewPropSet.RemoveProperty strNewFClassXName
  End If
  If Not IsEmpty(varProperty) Then
    For lngIndex3 = 0 To UBound(varProperty)
      strSubXName1 = "eainfo/detailed/attr[" & CStr(lngSourceFieldIndex) & "]/attrdomv/" & _
          varProperty(lngIndex3) & "[" & CStr(lngIndex3) & "]"
      varSubProp1 = pOrigPropSet.GetProperty(strSubXName1)
  
      For lngIndex4 = 0 To UBound(varSubProp1)
        strSubXName2 = "eainfo/detailed/attr[" & CStr(lngSourceFieldIndex) & "]/attrdomv/" & _
            varProperty(lngIndex3) & "[" & CStr(lngIndex3) & "]/" & varSubProp1(lngIndex4)
        varSubProp2 = pOrigPropSet.GetProperty(strSubXName2)
  
        ' WRITE VALUE BACK TO NEW FCLASS
        strSubXName3 = "eainfo/detailed/attr[" & CStr(lngDestinationFieldIndex) & "]/attrdomv/" & _
            varProperty(lngIndex3) & "[" & CStr(lngIndex3) & "]/" & varSubProp1(lngIndex4)
        strValue = varSubProp2(0)
        pNewPropSet.SetProperty strSubXName3, strValue
  
      Next lngIndex4
    Next lngIndex3
  End If
  
  ' COPY ANY SUBTYPE INFO OVER
'  Dim lngSubTypeIndex As Long
'  Dim lngStFieldIndex As Long
'  Dim lngMaxSubTypeIndex As Long
'  Dim lngMaxStFieldIndex As Long
'
'  lngMaxSubTypeIndex = ReturnLargestIndexValue("", pSourceDataset)
  
'  pPropSet.RemoveProperty "eainfo/detailed/subtype"
'  pPropSet.SetProperty "eainfo/detailed/subtype[0]/stname", "Administrative"
'  pPropSet.SetProperty "eainfo/detailed/subtype[0]/stcode", "1"
'  pPropSet.SetProperty "eainfo/detailed/subtype[0]/stfield[0]/stfldnm", "ROADS_SurfaceMaterial"
'  pPropSet.SetProperty "eainfo/detailed/subtype[0]/stfield[0]/stflddd/domname", "SurfaceMaterialDomain"
'  pPropSet.SetProperty "eainfo/detailed/subtype[0]/stfield[0]/stflddd/domdesc", "Description"
'  pPropSet.SetProperty "eainfo/detailed/subtype[0]/stfield[0]/stflddd/domtype", "Coded Value"
'  pPropSet.SetProperty "eainfo/detailed/subtype[0]/stfield[0]/stflddd/mrgtype", "Default value"
'  pPropSet.SetProperty "eainfo/detailed/subtype[0]/stfield[0]/stflddd/splttype", "Default value"
'  pPropSet.SetProperty "eainfo/detailed/subtype[0]/stfield[0]/stflddd/domowner", ""
'  pPropSet.SetProperty "eainfo/detailed/subtype[0]/stfield[0]/stflddd/domfldtp", "String"
'  pPropSet.SetProperty "eainfo/detailed/subtype[0]/stfield[1]/stfldnm", "ROADS_Status"
'  pPropSet.SetProperty "eainfo/detailed/subtype[0]/stfield[1]/stflddd/domname", "RoadStatusDomain"
'  pPropSet.SetProperty "eainfo/detailed/subtype[0]/stfield[1]/stflddd/domdesc", "Description"
'  pPropSet.SetProperty "eainfo/detailed/subtype[0]/stfield[1]/stflddd/domtype", "Coded Value"
'  pPropSet.SetProperty "eainfo/detailed/subtype[0]/stfield[1]/stflddd/mrgtype", "Default value"
'  pPropSet.SetProperty "eainfo/detailed/subtype[0]/stfield[1]/stflddd/splttype", "Default value"
'  pPropSet.SetProperty "eainfo/detailed/subtype[0]/stfield[1]/stflddd/domowner", ""
'  pPropSet.SetProperty "eainfo/detailed/subtype[0]/stfield[1]/stflddd/domfldtp", "String"
'  pPropSet.SetProperty "eainfo/detailed/subtype[0]/stfield[2]/stfldnm", "ROADS_Compendium"
'  pPropSet.SetProperty "eainfo/detailed/subtype[0]/stfield[2]/stflddd/domname", "CompendiumDomain"
'  pPropSet.SetProperty "eainfo/detailed/subtype[0]/stfield[2]/stflddd/domdesc", "Description"
'  pPropSet.SetProperty "eainfo/detailed/subtype[0]/stfield[2]/stflddd/domtype", "Coded Value"
'  pPropSet.SetProperty "eainfo/detailed/subtype[0]/stfield[2]/stflddd/mrgtype", "Default value"
'  pPropSet.SetProperty "eainfo/detailed/subtype[0]/stfield[2]/stflddd/splttype", "Default value"
'  pPropSet.SetProperty "eainfo/detailed/subtype[0]/stfield[2]/stflddd/domowner", ""
'  pPropSet.SetProperty "eainfo/detailed/subtype[0]/stfield[2]/stflddd/domfldtp", "String"
'  pPropSet.SetProperty "eainfo/detailed/subtype[0]/stfield[3]/stfldnm", "WildernessRecommendation"
'  pPropSet.SetProperty "eainfo/detailed/subtype[0]/stfield[3]/stflddd/domname", "WildernessRecommendDomain"
'  pPropSet.SetProperty "eainfo/detailed/subtype[0]/stfield[3]/stflddd/domdesc", "Description"
'  pPropSet.SetProperty "eainfo/detailed/subtype[0]/stfield[3]/stflddd/domtype", "Coded Value"
'  pPropSet.SetProperty "eainfo/detailed/subtype[0]/stfield[3]/stflddd/mrgtype", "Default value"
'  pPropSet.SetProperty "eainfo/detailed/subtype[0]/stfield[3]/stflddd/splttype", "Default value"
'  pPropSet.SetProperty "eainfo/detailed/subtype[0]/stfield[3]/stflddd/domowner", ""
'  pPropSet.SetProperty "eainfo/detailed/subtype[0]/stfield[3]/stflddd/domfldtp", "String"
'  pPropSet.SetProperty "eainfo/detailed/subtype[0]/stfield[4]/stfldnm", "WildernessMap_1980"
'  pPropSet.SetProperty "eainfo/detailed/subtype[0]/stfield[4]/stflddd/domname", "WildMap1980Domain"
'  pPropSet.SetProperty "eainfo/detailed/subtype[0]/stfield[4]/stflddd/domdesc", "Description"
'  pPropSet.SetProperty "eainfo/detailed/subtype[0]/stfield[4]/stflddd/domtype", "Coded Value"
'  pPropSet.SetProperty "eainfo/detailed/subtype[0]/stfield[4]/stflddd/mrgtype", "Default value"
'  pPropSet.SetProperty "eainfo/detailed/subtype[0]/stfield[4]/stflddd/splttype", "Default value"
'  pPropSet.SetProperty "eainfo/detailed/subtype[0]/stfield[4]/stflddd/domowner", ""
'  pPropSet.SetProperty "eainfo/detailed/subtype[0]/stfield[4]/stflddd/domfldtp", "String"
'  pPropSet.SetProperty "eainfo/detailed/subtype[1]/stname", "Closed"
'  pPropSet.SetProperty "eainfo/detailed/subtype[1]/stcode", "2"
'  pPropSet.SetProperty "eainfo/detailed/subtype[1]/stfield[0]/stfldnm", "ROADS_SurfaceMaterial"
'  pPropSet.SetProperty "eainfo/detailed/subtype[1]/stfield[0]/stflddd/domname", "SurfaceMaterialDomain"
'  pPropSet.SetProperty "eainfo/detailed/subtype[1]/stfield[0]/stflddd/domdesc", "Description"
'  pPropSet.SetProperty "eainfo/detailed/subtype[1]/stfield[0]/stflddd/domtype", "Coded Value"
'  pPropSet.SetProperty "eainfo/detailed/subtype[1]/stfield[0]/stflddd/mrgtype", "Default value"
'  pPropSet.SetProperty "eainfo/detailed/subtype[1]/stfield[0]/stflddd/splttype", "Default value"
'  pPropSet.SetProperty "eainfo/detailed/subtype[1]/stfield[0]/stflddd/domowner", ""
'  pPropSet.SetProperty "eainfo/detailed/subtype[1]/stfield[0]/stflddd/domfldtp", "String"
'  pPropSet.SetProperty "eainfo/detailed/subtype[1]/stfield[1]/stfldnm", "ROADS_Status"
'  pPropSet.SetProperty "eainfo/detailed/subtype[1]/stfield[1]/stflddd/domname", "RoadStatusDomain"
'  pPropSet.SetProperty "eainfo/detailed/subtype[1]/stfield[1]/stflddd/domdesc", "Description"
'  pPropSet.SetProperty "eainfo/detailed/subtype[1]/stfield[1]/stflddd/domtype", "Coded Value"
'  pPropSet.SetProperty "eainfo/detailed/subtype[1]/stfield[1]/stflddd/mrgtype", "Default value"
'  pPropSet.SetProperty "eainfo/detailed/subtype[1]/stfield[1]/stflddd/splttype", "Default value"
'  pPropSet.SetProperty "eainfo/detailed/subtype[1]/stfield[1]/stflddd/domowner", ""
'  pPropSet.SetProperty "eainfo/detailed/subtype[1]/stfield[1]/stflddd/domfldtp", "String"
'  pPropSet.SetProperty "eainfo/detailed/subtype[1]/stfield[2]/stfldnm", "ROADS_Compendium"
'  pPropSet.SetProperty "eainfo/detailed/subtype[1]/stfield[2]/stflddd/domname", "CompendiumDomain"
'  pPropSet.SetProperty "eainfo/detailed/subtype[1]/stfield[2]/stflddd/domdesc", "Description"
'  pPropSet.SetProperty "eainfo/detailed/subtype[1]/stfield[2]/stflddd/domtype", "Coded Value"
'  pPropSet.SetProperty "eainfo/detailed/subtype[1]/stfield[2]/stflddd/mrgtype", "Default value"
'  pPropSet.SetProperty "eainfo/detailed/subtype[1]/stfield[2]/stflddd/splttype", "Default value"
'  pPropSet.SetProperty "eainfo/detailed/subtype[1]/stfield[2]/stflddd/domowner", ""
'  pPropSet.SetProperty "eainfo/detailed/subtype[1]/stfield[2]/stflddd/domfldtp", "String"
'  pPropSet.SetProperty "eainfo/detailed/subtype[1]/stfield[3]/stfldnm", "WildernessRecommendation"
'  pPropSet.SetProperty "eainfo/detailed/subtype[1]/stfield[3]/stflddd/domname", "WildernessRecommendDomain"
'  pPropSet.SetProperty "eainfo/detailed/subtype[1]/stfield[3]/stflddd/domdesc", "Description"
'  pPropSet.SetProperty "eainfo/detailed/subtype[1]/stfield[3]/stflddd/domtype", "Coded Value"
'  pPropSet.SetProperty "eainfo/detailed/subtype[1]/stfield[3]/stflddd/mrgtype", "Default value"
'  pPropSet.SetProperty "eainfo/detailed/subtype[1]/stfield[3]/stflddd/splttype", "Default value"
'  pPropSet.SetProperty "eainfo/detailed/subtype[1]/stfield[3]/stflddd/domowner", ""
'  pPropSet.SetProperty "eainfo/detailed/subtype[1]/stfield[3]/stflddd/domfldtp", "String"
'  pPropSet.SetProperty "eainfo/detailed/subtype[1]/stfield[4]/stfldnm", "WildernessMap_1980"
'  pPropSet.SetProperty "eainfo/detailed/subtype[1]/stfield[4]/stflddd/domname", "WildMap1980Domain"
'  pPropSet.SetProperty "eainfo/detailed/subtype[1]/stfield[4]/stflddd/domdesc", "Description"
'  pPropSet.SetProperty "eainfo/detailed/subtype[1]/stfield[4]/stflddd/domtype", "Coded Value"
'  pPropSet.SetProperty "eainfo/detailed/subtype[1]/stfield[4]/stflddd/mrgtype", "Default value"
'  pPropSet.SetProperty "eainfo/detailed/subtype[1]/stfield[4]/stflddd/splttype", "Default value"
'  pPropSet.SetProperty "eainfo/detailed/subtype[1]/stfield[4]/stflddd/domowner", ""
'  pPropSet.SetProperty "eainfo/detailed/subtype[1]/stfield[4]/stflddd/domfldtp", "String"
'  pPropSet.SetProperty "eainfo/detailed/subtype[2]/stname", "Public"
'  pPropSet.SetProperty "eainfo/detailed/subtype[2]/stcode", "3"
'  pPropSet.SetProperty "eainfo/detailed/subtype[2]/stfield[0]/stfldnm", "ROADS_SurfaceMaterial"
'  pPropSet.SetProperty "eainfo/detailed/subtype[2]/stfield[0]/stflddd/domname", "SurfaceMaterialDomain"
'  pPropSet.SetProperty "eainfo/detailed/subtype[2]/stfield[0]/stflddd/domdesc", "Description"
'  pPropSet.SetProperty "eainfo/detailed/subtype[2]/stfield[0]/stflddd/domtype", "Coded Value"
'  pPropSet.SetProperty "eainfo/detailed/subtype[2]/stfield[0]/stflddd/mrgtype", "Default value"
'  pPropSet.SetProperty "eainfo/detailed/subtype[2]/stfield[0]/stflddd/splttype", "Default value"
'  pPropSet.SetProperty "eainfo/detailed/subtype[2]/stfield[0]/stflddd/domowner", ""
'  pPropSet.SetProperty "eainfo/detailed/subtype[2]/stfield[0]/stflddd/domfldtp", "String"
'  pPropSet.SetProperty "eainfo/detailed/subtype[2]/stfield[1]/stfldnm", "ROADS_Status"
'  pPropSet.SetProperty "eainfo/detailed/subtype[2]/stfield[1]/stflddd/domname", "RoadStatusDomain"
'  pPropSet.SetProperty "eainfo/detailed/subtype[2]/stfield[1]/stflddd/domdesc", "Description"
'  pPropSet.SetProperty "eainfo/detailed/subtype[2]/stfield[1]/stflddd/domtype", "Coded Value"
'  pPropSet.SetProperty "eainfo/detailed/subtype[2]/stfield[1]/stflddd/mrgtype", "Default value"
'  pPropSet.SetProperty "eainfo/detailed/subtype[2]/stfield[1]/stflddd/splttype", "Default value"
'  pPropSet.SetProperty "eainfo/detailed/subtype[2]/stfield[1]/stflddd/domowner", ""
'  pPropSet.SetProperty "eainfo/detailed/subtype[2]/stfield[1]/stflddd/domfldtp", "String"
'  pPropSet.SetProperty "eainfo/detailed/subtype[2]/stfield[2]/stfldnm", "ROADS_Compendium"
'  pPropSet.SetProperty "eainfo/detailed/subtype[2]/stfield[2]/stflddd/domname", "CompendiumDomain"
'  pPropSet.SetProperty "eainfo/detailed/subtype[2]/stfield[2]/stflddd/domdesc", "Description"
'  pPropSet.SetProperty "eainfo/detailed/subtype[2]/stfield[2]/stflddd/domtype", "Coded Value"
'  pPropSet.SetProperty "eainfo/detailed/subtype[2]/stfield[2]/stflddd/mrgtype", "Default value"
'  pPropSet.SetProperty "eainfo/detailed/subtype[2]/stfield[2]/stflddd/splttype", "Default value"
'  pPropSet.SetProperty "eainfo/detailed/subtype[2]/stfield[2]/stflddd/domowner", ""
'  pPropSet.SetProperty "eainfo/detailed/subtype[2]/stfield[2]/stflddd/domfldtp", "String"
'  pPropSet.SetProperty "eainfo/detailed/subtype[2]/stfield[3]/stfldnm", "WildernessRecommendation"
'  pPropSet.SetProperty "eainfo/detailed/subtype[2]/stfield[3]/stflddd/domname", "WildernessRecommendDomain"
'  pPropSet.SetProperty "eainfo/detailed/subtype[2]/stfield[3]/stflddd/domdesc", "Description"
'  pPropSet.SetProperty "eainfo/detailed/subtype[2]/stfield[3]/stflddd/domtype", "Coded Value"
'  pPropSet.SetProperty "eainfo/detailed/subtype[2]/stfield[3]/stflddd/mrgtype", "Default value"
'  pPropSet.SetProperty "eainfo/detailed/subtype[2]/stfield[3]/stflddd/splttype", "Default value"
'  pPropSet.SetProperty "eainfo/detailed/subtype[2]/stfield[3]/stflddd/domowner", ""
'  pPropSet.SetProperty "eainfo/detailed/subtype[2]/stfield[3]/stflddd/domfldtp", "String"
'  pPropSet.SetProperty "eainfo/detailed/subtype[2]/stfield[4]/stfldnm", "WildernessMap_1980"
'  pPropSet.SetProperty "eainfo/detailed/subtype[2]/stfield[4]/stflddd/domname", "WildMap1980Domain"
'  pPropSet.SetProperty "eainfo/detailed/subtype[2]/stfield[4]/stflddd/domdesc", "Description"
'  pPropSet.SetProperty "eainfo/detailed/subtype[2]/stfield[4]/stflddd/domtype", "Coded Value"
'  pPropSet.SetProperty "eainfo/detailed/subtype[2]/stfield[4]/stflddd/mrgtype", "Default value"
'  pPropSet.SetProperty "eainfo/detailed/subtype[2]/stfield[4]/stflddd/splttype", "Default value"
'  pPropSet.SetProperty "eainfo/detailed/subtype[2]/stfield[4]/stflddd/domowner", ""
'  pPropSet.SetProperty "eainfo/detailed/subtype[2]/stfield[4]/stflddd/domfldtp", "String"

  Metadata_Functions.SaveMetadata pDestinationDataset, pNewPropSet
  
  GoTo ClearMemory
  Exit Function
  
ErrHandler:
  booFailed = True
  CopyAttributesFromSourceToDestinationField = "Failed"
  
ClearMemory:

  Set pOrigPropSet = Nothing
  Set pNewPropSet = Nothing
  varVal = Null
  varProperty = Null
  varSubProperty = Null
  varSubProp1 = Null
  varSubProp2 = Null
  varCheckPropertyPresent = Null




'
'
'    ' ----------------------------------------------------------------------------------------------
'    ' <<<<<<<<<<<<  BLM_AZStrip_UTM12_NAD83  >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'    '-----------------------------------------------------------------------------------------------
'    If strFieldName = "BLM_Add_Attribute" Then
'      strDescription = """TRUE"" or ""FALSE"", whether there was a polyline feature from the 'BLM_AZStrip_UTM12_NAD83' " & _
'          "feature class within 100m of the centroid of this polyline.  If so, then this '" & _
'          "BLM_AZStrip_UTM12_NAD83' feature would be considered a candidate for extracting attribute values " & _
'          "to describe this feature."
'      If Not IsEmpty(varProperty) Then
'        pPropSet.RemoveProperty strXName
'      End If
'      pPropSet.SetProperty strXName, strDescription
'
'      ' SET SOURCE IF IT IS NOT ALREADY SET
'      varSource = pPropSet.GetProperty(strXNameSource)
'      If IsEmpty(varSource) Then
'        pPropSet.SetProperty strXNameSource, "Lab of Landscape Ecology and Conservation Biology; " & _
'            "School of Earth Sciences and Environmental Sustainability; College of Engineering, " & _
'            "Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011"
'      End If
'    End If
'    If strFieldName = "BLM_Attribute_Source" Then
'      strDescription = "If 'BLM_Add_Attribute' above = ""TRUE"", then this field will contain the name of the " & _
'          "BLM feature class (BLM_AZStrip_UTM12_NAD83).  Otherwise it should be Null."
'      If Not IsEmpty(varProperty) Then
'        pPropSet.RemoveProperty strXName
'      End If
'      pPropSet.SetProperty strXName, strDescription
'
'      ' SET SOURCE IF IT IS NOT ALREADY SET
'      varSource = pPropSet.GetProperty(strXNameSource)
'      If IsEmpty(varSource) Then
'        pPropSet.SetProperty strXNameSource, "Lab of Landscape Ecology and Conservation Biology; " & _
'            "School of Earth Sciences and Environmental Sustainability; College of Engineering, " & _
'            "Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011"
'      End If
'    End If
'    If strFieldName = "BLM_Attribute_Distance" Then
'      strDescription = "If 'BLM_Add_Attribute' above = ""TRUE"", then this field will contain the distance " & _
'          "(in meters) from the centroid of this polyline to the nearest BLM feature.  Otherwise it should be Null.  " & _
'          "Values should always be <= 100m."
'      If Not IsEmpty(varProperty) Then
'        pPropSet.RemoveProperty strXName
'      End If
'      pPropSet.SetProperty strXName, strDescription
'
'      ' SET SOURCE IF IT IS NOT ALREADY SET
'      varSource = pPropSet.GetProperty(strXNameSource)
'      If IsEmpty(varSource) Then
'        pPropSet.SetProperty strXNameSource, "Lab of Landscape Ecology and Conservation Biology; " & _
'            "School of Earth Sciences and Environmental Sustainability; College of Engineering, " & _
'            "Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011"
'      End If
'
'      pPropSet.RemoveProperty strXDom
'      pPropSet.SetProperty strXDomMin, "0"
'      pPropSet.SetProperty strXDomMax, "100"
'      pPropSet.SetProperty strXDomUnits, "Meters"
'    End If
'    If strFieldName = "BLM_ObjectID" Then
'      strDescription = "If 'BLM_Add_Attribute' above = ""TRUE"", then this field will contain the OBJECTID " & _
'          "value of the nearest BLM feature, from the BLM_AZStrip_UTM12_NAD83 feature class."
'      If Not IsEmpty(varProperty) Then
'        pPropSet.RemoveProperty strXName
'      End If
'      pPropSet.SetProperty strXName, strDescription
'
'      ' SET SOURCE IF IT IS NOT ALREADY SET
'      varSource = pPropSet.GetProperty(strXNameSource)
'      If IsEmpty(varSource) Then
'        pPropSet.SetProperty strXNameSource, "Lab of Landscape Ecology and Conservation Biology; " & _
'            "School of Earth Sciences and Environmental Sustainability; College of Engineering, " & _
'            "Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011"
'      End If
'    End If
'
'
'    For lngBLMIndex = 0 To UBound(strBLMObsArray, 1)
'      If strFieldName = strBLMObsArray(lngBLMIndex, 0) Then '  "BLM_OBS_USE1", "BLM_OBS_USE2" or "BLM_OBS_USE3"
'
'        ' FIND ORIGINAL FIELD METADATA
'        For lngIndex = 0 To pBLM_AZStrip_UTM12_NAD83_FClass.Fields.FieldCount
'          strNames(0) = "eainfo/detailed/attr[" & CStr(lngIndex) & "]"
'          pBLM_AZStrip_UTM12_NAD83_PropSet.GetProperties strNames, varVals
'          varSubVals = varVals(0)
'          If Not IsEmpty(varSubVals) Then
'            ' GET FIELD NAME
'            strOrigFCLassXName = "eainfo/detailed/attr[" & CStr(lngIndex) & "]/attrlabl"
'            varProperty = pBLM_AZStrip_UTM12_NAD83_PropSet.GetProperty(strOrigFCLassXName)
'            If Not IsEmpty(varProperty) Then
'              varSubProperty = varProperty(0)
'              If IsEmpty(varSubProperty) Then
'                strOrigFieldName = "<-- No Field Name -->"
'              Else
'                strOrigFieldName = CStr(varSubProperty)
'              End If
'            Else
'              strOrigFieldName = "<-- No Field Name -->"
'            End If
'
'            If strOrigFieldName = strBLMObsArray(lngBLMIndex, 1) Then ' "OBS_USE1", "OBS_USE2", or "OBS_USE3"
'
'
'              If lngBLMIndex <= 4 Then
'                strOrigFCLassXName = "eainfo/detailed/attr[" & CStr(lngIndex) & "]/attrdomv"
'                varProperty = pBLM_AZStrip_UTM12_NAD83_PropSet.GetProperty(strOrigFCLassXName)
'                varSubProperty = varProperty(0)
''                For lngIndex2 = 0 To UBound(varProperty)
''                  Debug.Print "       " & CStr(lngIndex2) & "] " & CStr(varProperty(lngIndex2))
''                Next lngIndex2
'
'                ' GET DOMAIN ATTRIBUTES FROM ORIGINAL FCLASS
'                strOrigFCLassXName = "eainfo/detailed/attr[" & CStr(lngIndex) & "]/attrdomv"
'                varProperty = pBLM_AZStrip_UTM12_NAD83_PropSet.GetProperty(strOrigFCLassXName)
'
'                ' ADD DOMAIN ATTRIBUTES TO NEW FCLASS
'                varCheckPropertyPresent = pPropSet.GetProperty("eainfo/detailed/attr[" & CStr(lngFieldIndex) & "]/attrdomv")
'                If Not IsEmpty(varCheckPropertyPresent) Then
'                  pPropSet.RemoveProperty "eainfo/detailed/attr[" & CStr(lngFieldIndex) & "]/attrdomv"
'                End If
'                For lngIndex3 = 0 To UBound(varProperty)
'                  strSubXName1 = "eainfo/detailed/attr[" & CStr(lngIndex) & "]/attrdomv/" & _
'                      varProperty(lngIndex3) & "[" & CStr(lngIndex3) & "]"
'                  varSubProp1 = pBLM_AZStrip_UTM12_NAD83_PropSet.GetProperty(strSubXName1)
'
'                  For lngIndex4 = 0 To UBound(varSubProp1)
'                    strSubXName2 = "eainfo/detailed/attr[" & CStr(lngIndex) & "]/attrdomv/" & _
'                        varProperty(lngIndex3) & "[" & CStr(lngIndex3) & "]/" & varSubProp1(lngIndex4)
'                    varSubProp2 = pBLM_AZStrip_UTM12_NAD83_PropSet.GetProperty(strSubXName2)
'
'                    ' WRITE VALUE BACK TO NEW FCLASS
'                    strSubXName3 = "eainfo/detailed/attr[" & CStr(lngFieldIndex) & "]/attrdomv/" & _
'                        varProperty(lngIndex3) & "[" & CStr(lngIndex3) & "]/" & varSubProp1(lngIndex4)
'                    strValue = varSubProp2(0)
'                    strValue = Replace(strValue, "20003", "2003")
'                    strValue = Replace(strValue, "Dicitonary", "Dictionary")
'                    pPropSet.SetProperty strSubXName3, strValue
'
'                  Next lngIndex4
'                Next lngIndex3
'              Else
''                strOrigFCLassXName = "eainfo/detailed/attr[" & CStr(lngIndex) & "]"
''                varProperty = pBLM_AZStrip_UTM12_NAD83_PropSet.GetProperty(strOrigFCLassXName)
''                varSubProperty = varProperty(0)
''                Debug.Print "Examining " & strOrigFieldName & "..."
''                For lngIndex2 = 0 To UBound(varProperty)
''                  varSubProp1 = pBLM_AZStrip_UTM12_NAD83_PropSet.GetProperty(strOrigFCLassXName & "/" & _
''                        CStr(varProperty(lngIndex2)))
''                  Debug.Print "       " & CStr(lngIndex2) & "] " & CStr(varProperty(lngIndex2)) & " = " & _
''                        CStr(varSubProp1(0))
''                Next lngIndex2
'              End If
'
'              If strOrigFieldName = "ROAD_NO_" Then
'                Debug.Print "Here..."
'              End If
'
'              ' GET DEFINITION FROM ORIGINAL FCLASS
'              varProperty = pBLM_AZStrip_UTM12_NAD83_PropSet.GetProperty("eainfo/detailed/attr[" & CStr(lngIndex) & "]/attrdef")
'              If Not IsEmpty(varProperty) Then
'                varSubProperty = varProperty(0)
'                If IsEmpty(varSubProperty) Then
'                  strDescription = "<-- No Description -->"
'                Else
'                  strDescription = CStr(varSubProperty)
'                End If
'              Else
'                strDescription = "<-- No Description -->"
'              End If
'              strDescription = Replace(strDescription, "  [Imported from TransportationLine_UTM12_NAD83]", "", , , vbTextCompare)
'              strDescription = strDescription & _
'                  "  This field should only have a value if the 'BLM_Add_Attribute' value = ""TRUE"".  " & _
'                  "[Imported from BLM_AZStrip_UTM12_NAD83, Field """ & strBLMObsArray(lngBLMIndex, 1) & """]"
'
'              ' SET FIELD DEFINITION OF NEW FCLASS
'              varCheckPropertyPresent = pPropSet.GetProperty("eainfo/detailed/attr[" & CStr(lngFieldIndex) & "]/attrdef")
'              If Not IsEmpty(varCheckPropertyPresent) Then
'                pPropSet.RemoveProperty "eainfo/detailed/attr[" & CStr(lngFieldIndex) & "]/attrdef"
'              End If
'              pPropSet.SetProperty "eainfo/detailed/attr[" & CStr(lngFieldIndex) & "]/attrdef", strDescription
'
'              ' GET DESCRIPTION SOURCE FROM ORIGINAL FCLASS
'              varProperty = pBLM_AZStrip_UTM12_NAD83_PropSet.GetProperty("eainfo/detailed/attr[" & CStr(lngIndex) & "]/attrdefs")
'              If Not IsEmpty(varProperty) Then
'                varSubProperty = varProperty(0)
'                If IsEmpty(varSubProperty) Then
'                  strDescriptionSource = "BLM_AZStrip_UTM12_NAD83"
'                Else
'                  strDescriptionSource = CStr(varSubProperty)
'                End If
'              Else
'                strDescriptionSource = "BLM_AZStrip_UTM12_NAD83"
'              End If
'              strDescriptionSource = Replace(strDescriptionSource, "Dicitonary", "Dictionary", , , vbTextCompare)
'              strDescriptionSource = Replace(strDescriptionSource, "20003", "2003", , , vbTextCompare)
'
'              ' SET DESCRIPTION SOURCE DEFINITION OF NEW FCLASS
'              varCheckPropertyPresent = pPropSet.GetProperty("eainfo/detailed/attr[" & CStr(lngFieldIndex) & "]/attrdefs")
'              If Not IsEmpty(varCheckPropertyPresent) Then
'                pPropSet.RemoveProperty "eainfo/detailed/attr[" & CStr(lngFieldIndex) & "]/attrdefs"
'              End If
'
'              ' ONLY ADD THIS NEW ONE IF THERE WAS ONE IN THE ORIGINAL FCLASS
'              If strDescriptionSource <> "" Then
'                pPropSet.SetProperty "eainfo/detailed/attr[" & CStr(lngFieldIndex) & "]/attrdefs", strDescriptionSource
'              End If
'
'              Exit For
'            End If
'          End If
'        Next lngIndex
'      End If
'    Next lngBLMIndex
'  End If











End Function
Public Function SetMetadataFormatVersion(pDataset As IDataset, _
  Optional strFormatVersion As String, _
  Optional booInsertArcGISVersionAutomatically As Boolean = False, _
  Optional pMxDoc As IMxDocument) As String
    On Error GoTo ErrHandler
'
'  Dim strFormatVersion As String
'  Dim lngVersion As Long
'  lngVersion = aml_func_mod.ReturnArcGISVersionAlt(pMxDoc, strFormatVersion)
'  strResponse = Metadata_Functions.SetMetadataFormatVersion(pDataset, "Created in ArcGIS " & strFormatVersion)
'  Debug.Print "Added Format Version = " & strResponse
  
  SetMetadataFormatVersion = "Succeeded"
  
  Dim pPropSet As IPropertySet
  Set pPropSet = ReturnMetadataPropSetFromDataset(pDataset)
  
  Dim strAutoVersion As String
  Dim lngAutoVersion As Long
  If booInsertArcGISVersionAutomatically Then
    If pMxDoc Is Nothing Then
      MsgBox "Map Document required to generate version!"
      strAutoVersion = strFormatVersion
    Else
      lngAutoVersion = aml_func_mod.ReturnArcGISVersionAlt2(pMxDoc, strAutoVersion)
      strAutoVersion = "Created in ArcGIS Version " & strAutoVersion
    End If
    pPropSet.RemoveProperty "distInfo/distFormat/formatVer"
    pPropSet.SetProperty "distInfo/distFormat/formatVer", strAutoVersion
  Else
    pPropSet.RemoveProperty "distInfo/distFormat/formatVer"
    pPropSet.SetProperty "distInfo/distFormat/formatVer", strFormatVersion
  End If
  
  Metadata_Functions.SaveMetadata pDataset, pPropSet
  
  GoTo ClearMemory
  Exit Function
  
ErrHandler:
  SetMetadataFormatVersion = "Failed"

ClearMemory:
  Set pPropSet = Nothing

End Function

Public Function AddFieldAttributes(pDataset As IDataset, strFieldName As String, _
  strFieldDescription As String, strFieldDescriptionSource As String, _
  Optional strRDOMFieldMin As String, Optional strRDOMFieldMax As String, _
  Optional strRDOMFieldMean As String, Optional strRDOMFieldUnit As String, _
  Optional strRDOMFieldStDev As String, Optional strRDOMFieldMinResolution As String, _
  Optional varEDOMArrayOfList_ValueDescSource As Variant = Null, _
  Optional strUDOM_DescriptionOfValues As String, _
  Optional strCodesetNameOfList As String, Optional strCodesetSource As String, _
  Optional booClearExistingFieldInfoFirst As Boolean = True) As String
  
  On Error GoTo ErrHandler
  AddFieldAttributes = "Succeeded"
  
  ' rdom = RANGE DOMAIN
  ' edim = ENUMERATED DOMAIN
  ' udom = UNREPRESENTABLE DOMAIN
  ' codesetd = CODESET DOMAIN
  
'  strResponse = Metadata_Functions.AddResourceMaintenance(pDataset, JenMetadata_Ongoing)
'  Debug.Print "Added Resource Status = " & strResponse

  ' IF varStringArrayOfValueDescSourceList EXISTS, IT SHOULD CONTAIN A STRING
  ' ARRAY WITH DIMENSIONS (2,X), WHERE X IS NUMBER OF ELEMENTS IN LIST.
  ' THIS IS ZERO-BASED, SO "2" MEANS 3 ATTRIBUTES PER ELEMENT (LIST ITEM, DESCRIPTION, SOURCE)
  Dim booAddList As Boolean
  booAddList = False
  
  Dim strListArray() As String
  If Not IsNull(varEDOMArrayOfList_ValueDescSource) Then
    strListArray = varEDOMArrayOfList_ValueDescSource(0)
    If UBound(strListArray, 1) <> 2 Then
      MsgBox "Array has incorrect dimensions.  Skipping this item..."
    Else
      booAddList = True
    End If
  End If
     
  Dim pPropSet As IPropertySet
  Set pPropSet = ReturnMetadataPropSetFromDataset(pDataset)
  Dim pXMLPropSet As IXmlPropertySet2
  Set pXMLPropSet = pPropSet
  
  ' GET INDEX NUMBER
  Dim lngFieldIndex As Long
  Dim booFailed As Boolean
  lngFieldIndex = ReturnAttributeFieldXPathIndex(pDataset, strFieldName, booFailed)
  If lngFieldIndex = -1 Then
    If booFailed Then
      AddFieldAttributes = "ReturnAttributeFieldXPathIndex Failed"
    Else
      AddFieldAttributes = "No Field Found"
    End If
    GoTo ClearMemory
  End If
    
  Dim strXName As String
  Dim strXNameSource As String
  Dim strXDomMin As String
  Dim strXDomMax As String
  Dim strXDomUnits As String
  Dim strXDom As String
  Dim strXRDom As String
  Dim strXUDom As String
  Dim strXEDom As String
  Dim strXCodesetDom As String
  
  ' see http://resources.arcgis.com/en/help/main/10.1/index.html#//003t00000037000000
  ' GET CURRENT DESCRIPTION
  strXName = "eainfo/detailed/attr[" & CStr(lngFieldIndex) & "]/attrdef"         ' DESCRIPTION OF FIELD
  strXNameSource = "eainfo/detailed/attr[" & CStr(lngFieldIndex) & "]/attrdefs"   ' DESCRIPTION SOURCE
  strXRDom = "eainfo/detailed/attr[" & CStr(lngFieldIndex) & "]/attrdomv/rdom"  ' RANGE DOMAIN IN GENERAL
  strXDomMin = "eainfo/detailed/attr[" & CStr(lngFieldIndex) & "]/attrdomv/rdom/rdommin"  ' MINIMUM VALUE
  strXDomMax = "eainfo/detailed/attr[" & CStr(lngFieldIndex) & "]/attrdomv/rdom/rdommax"  ' MAXIMUM VALUE
  strXDomUnits = "eainfo/detailed/attr[" & CStr(lngFieldIndex) & "]/attrdomv/rdom/attrunit"   ' UNITS
  strXDom = "eainfo/detailed/attr[" & CStr(lngFieldIndex) & "]/attrdomv"
  strXUDom = "eainfo/detailed/attr[" & CStr(lngFieldIndex) & "]/attrdomv/udom"  ' DESCRIPTION OF VALUES
  strXEDom = "eainfo/detailed/attr[" & CStr(lngFieldIndex) & "]/attrdomv/edom"  ' LIST OF VALUES
  strXCodesetDom = "eainfo/detailed/attr[" & CStr(lngFieldIndex) & "]/attrdomv/codesetd"  ' GENERAL CODESET DOMAIN
                                                                          ' /edomv = Value
                                                                          ' /edomvd = Description of Value
                                                                          ' /edomvds = Enumerated domain value definition source
  ' rdom = RANGE DOMAIN
  ' edim = ENUMERATED DOMAIN
  ' udom = UNREPRESENTABLE DOMAIN
  ' codesetd = CODESET DOMAIN
  
  If booClearExistingFieldInfoFirst Then
    pPropSet.RemoveProperty strXName
    pPropSet.RemoveProperty strXNameSource
    pPropSet.RemoveProperty strXRDom
    pPropSet.RemoveProperty strXUDom
    pPropSet.RemoveProperty strXEDom
    pPropSet.RemoveProperty strXCodesetDom
  End If
  
  If Trim(strFieldDescription) <> "" Then pPropSet.SetProperty strXName, Trim(strFieldDescription)
  If Trim(strFieldDescriptionSource) <> "" Then pPropSet.SetProperty strXNameSource, Trim(strFieldDescriptionSource)
  If Trim(strRDOMFieldMin) <> "" Then pPropSet.SetProperty strXRDom & "/rdommin", Trim(strRDOMFieldMin)
  If Trim(strRDOMFieldMax) <> "" Then pPropSet.SetProperty strXRDom & "/rdommax", Trim(strRDOMFieldMax)
  If Trim(strRDOMFieldMean) <> "" Then pPropSet.SetProperty strXRDom & "/rdommean", Trim(strRDOMFieldMean)
  If Trim(strRDOMFieldUnit) <> "" Then pPropSet.SetProperty strXRDom & "/attrunit", Trim(strRDOMFieldUnit)
  If Trim(strRDOMFieldStDev) <> "" Then pPropSet.SetProperty strXRDom & "/rdomstdv", Trim(strRDOMFieldStDev)
  If Trim(strRDOMFieldMinResolution) <> "" Then pPropSet.SetProperty strXRDom & "/attrmres", Trim(strRDOMFieldMinResolution)
  If Trim(strUDOM_DescriptionOfValues) <> "" Then pPropSet.SetProperty strXUDom, strUDOM_DescriptionOfValues
  If Trim(strCodesetNameOfList) <> "" Then pPropSet.SetProperty strXCodesetDom & "/codesetn", strCodesetNameOfList
  If Trim(strCodesetSource) <> "" Then pPropSet.SetProperty strXCodesetDom & "/codesets", strCodesetSource
  
  Dim lngIndex As Long
  Dim strValue As String
  Dim strDescription As String
  Dim strSource As String
  Dim lngCounter As Long
  lngCounter = -1
  
  If booAddList Then
    For lngIndex = 0 To UBound(strListArray, 2)
      strValue = Trim(strListArray(0, lngIndex))
      strDescription = Trim(strListArray(1, lngIndex))
      strSource = Trim(strListArray(2, lngIndex))
      
      If strValue <> "" Or strDescription <> "" Or strSource <> "" Then
        lngCounter = lngCounter + 1
        If strValue <> "" Then pPropSet.SetProperty strXEDom & "[" & CStr(lngCounter) & "]/edomv", strValue
        If strDescription <> "" Then pPropSet.SetProperty strXEDom & "[" & CStr(lngCounter) & "]/edomvd", strDescription
        If strSource <> "" Then pPropSet.SetProperty strXEDom & "[" & CStr(lngCounter) & "]/edomvds", strSource
      End If
    Next lngIndex
  End If
  
'  strFieldDescription As String, strFieldDescriptionSource As String, _
'  Optional strRDOMFieldMin As String, Optional strRDOMFieldMax As String, _
'  Optional strRDOMFieldMean As String, Optional strRDOMFieldUnit As String, _
'  Optional strRDOMFieldStDev As String, Optional strRDOMFieldMinResolution As String, _
'  Optional varEDOMArrayOfList_ValueDescSource As Variant = Null, _
'  Optional strUDOM_DescriptionOfValues As String
  
  
  Metadata_Functions.SaveMetadata pDataset, pPropSet
  
  GoTo ClearMemory
  Exit Function
  
ErrHandler:
  AddFieldAttributes = "Failed"

ClearMemory:
  Erase strListArray
  Set pPropSet = Nothing
  Set pXMLPropSet = Nothing

  
  
End Function



Private Function ReturnAttributeFieldXPathIndex(pDataset As IDataset, strFieldName As String, _
    Optional booFailed As Boolean) As Long

  On Error GoTo ErrHandler
  ' GENERALLY CALLED BY OTHER FUNCTIONS
  booFailed = False
  
  Dim pPropSet As IPropertySet
  Set pPropSet = ReturnMetadataPropSetFromDataset(pDataset)
  
  ' GET INDEX NUMBER
  Dim lngFieldIndex As Long
  Dim booFoundField As Boolean
  Dim varVals As Variant
  varVals = Array("placeholder")
  Dim varName As Variant
  Dim strTestName As String
  
  booFoundField = False
  Dim lngIndex As Long
  lngFieldIndex = -1
  lngIndex = -1
  Do Until IsEmpty(varVals)
    lngIndex = lngIndex + 1
    varVals = pPropSet.GetProperty("eainfo/detailed/attr[" & CStr(lngIndex) & "]")
    If Not IsEmpty(varVals) Then
      varName = pPropSet.GetProperty("eainfo/detailed/attr[" & CStr(lngIndex) & "]/attrlabl")
      If Not IsEmpty(varName) Then
        strTestName = CStr(varName(0))
'        Debug.Print CStr(lngIndex) & "] " & strTestName
        If StrComp(Trim(strTestName), Trim(strFieldName), vbTextCompare) = 0 Then
          booFoundField = True
          lngFieldIndex = lngIndex
          Exit Do
        End If
      End If
    End If
  Loop
  
  If booFoundField Then
    ReturnAttributeFieldXPathIndex = lngFieldIndex
  Else
    ReturnAttributeFieldXPathIndex = -1
  End If
  
  GoTo ClearMemory
  Exit Function
  
ErrHandler:
  booFailed = True
  ReturnAttributeFieldXPathIndex = -1
  
ClearMemory:
  Set pPropSet = Nothing
  varVals = Null
  varName = Null
  
End Function



Public Function AddResourceMaintenance(pDataset As IDataset, enumMaintCode As JenMetadataMaintenanceCodes) As String

  On Error GoTo ErrHandler
  
'  strResponse = Metadata_Functions.AddResourceMaintenance(pDataset, JenMetadata_Ongoing)
'  Debug.Print "Added Resource Status = " & strResponse
  
  AddResourceMaintenance = "Succeeded"
  
  Dim pPropSet As IPropertySet
  Set pPropSet = ReturnMetadataPropSetFromDataset(pDataset)
  Dim pXMLPropSet As IXmlPropertySet2
  Set pXMLPropSet = pPropSet
  
  Dim strMaintenance As String
  strMaintenance = ReturnMaintenanceCode(enumMaintCode)
  pXMLPropSet.SetAttribute "dataIdInfo/resMaint/maintFreq/MaintFreqCd", "value", strMaintenance, esriXSPAAddOrReplace
  
  Metadata_Functions.SaveMetadata pDataset, pPropSet
  
  GoTo ClearMemory
  Exit Function
  
ErrHandler:
  AddResourceMaintenance = "Failed"

ClearMemory:
  Set pPropSet = Nothing
  Set pXMLPropSet = Nothing
End Function
Public Function AddResourceDetailsStatus(pDataset As IDataset, enumJenStatus As JenMetadataStatusValues) As String

  On Error GoTo ErrHandler
  
'  strResponse = Metadata_Functions.AddResourceDetailsStatus(pDataset, JenMetadata_Ongoing)
'  Debug.Print "Added Resource Status = " & strResponse
  
  AddResourceDetailsStatus = "Succeeded"
  
  Dim pPropSet As IPropertySet
  Set pPropSet = ReturnMetadataPropSetFromDataset(pDataset)
  Dim pXMLPropSet As IXmlPropertySet2
  Set pXMLPropSet = pPropSet
  
  Dim strStatus As String
  strStatus = ReturnStatusString(enumJenStatus)
  pXMLPropSet.SetAttribute "dataIdInfo/idStatus/ProgCd", "value", strStatus, esriXSPAAddOrReplace
  
  Metadata_Functions.SaveMetadata pDataset, pPropSet
  
  GoTo ClearMemory
  Exit Function
  
ErrHandler:
  AddResourceDetailsStatus = "Failed"

ClearMemory:
  Set pPropSet = Nothing
  Set pXMLPropSet = Nothing
  
End Function

Public Function AddCitationDates(pDataset As IDataset, _
  Optional datCreated As Date = CDate(0), _
  Optional datPublished As Date = CDate(0), _
  Optional datRevised As Date = CDate(0)) As String
  
  On Error GoTo ErrHandler
  
'  Dim datCreated As Date
'  Dim datPublished As Date
'  datCreated = Now
'  datPublished = DateAdd("s", 1234567, datCreated)
'  strResponse = Metadata_Functions.AddCitationDates(pDataset, datCreated, datPublished)
'  Debug.Print "Added Citation Dates = " & strResponse

  AddCitationDates = "Succeeded"
  
  Dim pPropSet As IPropertySet
  Set pPropSet = ReturnMetadataPropSetFromDataset(pDataset)
  
  If Not datCreated = CDate(0) Then
    pPropSet.SetProperty "dataIdInfo/idCitation/date/createDate", Format(datCreated, "yyyy-mm-ddTHh:Nn:Ss")
  End If
  If Not datPublished = CDate(0) Then
    pPropSet.SetProperty "dataIdInfo/idCitation/date/pubDate", Format(datPublished, "yyyy-mm-ddTHh:Nn:Ss")
  End If
  If Not datRevised = CDate(0) Then
    pPropSet.SetProperty "dataIdInfo/idCitation/date/reviseDate", Format(datRevised, "yyyy-mm-ddTHh:Nn:Ss")
  End If

  Metadata_Functions.SaveMetadata pDataset, pPropSet
  
  GoTo ClearMemory
  Exit Function
  
ErrHandler:
  AddCitationDates = "Failed"

ClearMemory:
  Set pPropSet = Nothing
  
End Function



Public Function AddContact_Metadata(pDataset As IDataset, _
    enumJenRole As JenMetadataRoleCDValues, _
    strIndividualName As String, _
    Optional booGetIndividualNameAutomatically As Boolean = False, _
    Optional strOrganizationName As String = "", _
    Optional strPositionName As String = "", _
    Optional strVoiceNumber As String = "", _
    Optional strAddressStreet As String = "", _
    Optional strAddressCity As String = "", _
    Optional strAddressState As String = "", _
    Optional strAddressZip As String = "", _
    Optional strAddressCountry As String = "", _
    Optional strAddressEmail As String = "", _
    Optional enumJenAddressType As JenMetadataAddressTypeValues = JenMetadata_Skip, _
    Optional booSkipIfAlreadyPresent As Boolean = True) As String
    
    On Error GoTo ErrHandler
    
  ' NEED CONTACT FOR METADATA, CITATION RESPONSIBLE PARTY AND RESOURCE POINT OF CONTACT
  ' ADD NEW METADATA CONTACTS
  '  strName = aml_func_mod.GetTheUserName
  '  strResponse = Metadata_Functions.AddContact_Metadata(pDataset, JenMetadata_PointOfContact, _
  '    "Jeff Jenness or strName", False, "Jenness Enterprises", "Geoanalyst and Programmer", "1-928-607-4635", _
  '    "3020 N. Schevene Blvd.", "Flagstaff", "AZ", "86004", "USA", "jeffj@jennessent.com", _
  "    JenMetadata_Postal, True)
  ' Debug.Print "Added Metadata Contact 'Custodian' = " & strResponse
  
  AddContact_Metadata = "Succeeded"
  
  If booGetIndividualNameAutomatically Then strIndividualName = aml_func_mod.GetTheUserName
  
  Dim pPropSet As IPropertySet
  Set pPropSet = ReturnMetadataPropSetFromDataset(pDataset)
  Dim pXMLPropSet As IXmlPropertySet2
  Set pXMLPropSet = pPropSet
  
  Dim strRole As String
  strRole = ReturnRoleCDString(enumJenRole)
  
  Dim strAddressType As String
  strAddressType = ReturnAddressType(enumJenAddressType)
    
  Dim lngXIndex As Long
  Dim strXPath As String
  strXPath = "mdContact"
  lngXIndex = ReturnLargestIndexValue(strXPath, pDataset) + 1
  
  Dim booContactAlreadyPresent As Boolean
  If booSkipIfAlreadyPresent Then
    booContactAlreadyPresent = CheckIfContactAlreadyPresent(pDataset, _
      enumJenRole, strIndividualName, strOrganizationName, _
       strPositionName, strVoiceNumber, strAddressStreet, strAddressCity, _
       strAddressState, strAddressZip, strAddressCountry, strAddressEmail, _
       enumJenAddressType, lngXIndex, pPropSet, pXMLPropSet, strXPath)
    
    If booContactAlreadyPresent Then
      AddContact_Metadata = "Metadata Contact Already Present"
      GoTo ClearMemory
    End If
  End If
  
  pPropSet.SetProperty "mdContact[" & CStr(lngXIndex) & "]/rpIndName", strIndividualName
  If strOrganizationName <> "" Then
    pPropSet.SetProperty "mdContact[" & CStr(lngXIndex) & "]/rpOrgName", strOrganizationName
  End If
  If strPositionName <> "" Then
    pPropSet.SetProperty "mdContact[" & CStr(lngXIndex) & "]/rpPosName", strPositionName
  End If
  If strVoiceNumber <> "" Then
    pPropSet.SetProperty "mdContact[" & CStr(lngXIndex) & "]/rpCntInfo/cntPhone/voiceNum", strVoiceNumber
  End If
  If strAddressStreet <> "" Then
    pPropSet.SetProperty "mdContact[" & CStr(lngXIndex) & "]/rpCntInfo/cntAddress/delPoint", strAddressStreet
  End If
  If strAddressCity <> "" Then
    pPropSet.SetProperty "mdContact[" & CStr(lngXIndex) & "]/rpCntInfo/cntAddress/city", strAddressCity
  End If
  If strAddressState <> "" Then
    pPropSet.SetProperty "mdContact[" & CStr(lngXIndex) & "]/rpCntInfo/cntAddress/adminArea", strAddressState
  End If
  If strAddressZip <> "" Then
    pPropSet.SetProperty "mdContact[" & CStr(lngXIndex) & "]/rpCntInfo/cntAddress/postCode", strAddressZip
  End If
  If strAddressCountry <> "" Then
    pPropSet.SetProperty "mdContact[" & CStr(lngXIndex) & "]/rpCntInfo/cntAddress/country", strAddressCountry
  End If
  If strAddressEmail <> "" Then
    pPropSet.SetProperty "mdContact[" & CStr(lngXIndex) & "]/rpCntInfo/cntAddress/eMailAdd", strAddressEmail
  End If
  
  pPropSet.SetProperty "mdContact[" & CStr(lngXIndex) & "]/role/RoleCd", ""
  pXMLPropSet.SetAttribute "mdContact[" & CStr(lngXIndex) & "]/role/RoleCd", "value", strRole, esriXSPAAddOrReplace
  If enumJenAddressType <> JenMetadata_Skip Then
    pXMLPropSet.SetAttribute "mdContact[" & CStr(lngXIndex) & _
         "]/rpCntInfo/cntAddress", "addressType", strAddressType, esriXSPAAddOrReplace
  End If

  Metadata_Functions.SaveMetadata pDataset, pPropSet

  GoTo ClearMemory
  Exit Function
  
ErrHandler:
  AddContact_Metadata = "Failed"

ClearMemory:
  Set pPropSet = Nothing
  Set pXMLPropSet = Nothing
End Function

Private Function CheckIfContactAlreadyPresent(pDataset As IDataset, _
    enumJenRole As JenMetadataRoleCDValues, _
    strIndividualName As String, _
    strOrganizationName As String, _
    strPositionName As String, _
    strVoiceNumber As String, _
    strAddressStreet As String, _
    strAddressCity As String, _
    strAddressState As String, _
    strAddressZip As String, _
    strAddressCountry As String, _
    strAddressEmail As String, _
    enumJenAddressType As JenMetadataAddressTypeValues, _
    lngMaxIndex As Long, _
    pPropSet As IPropertySet, _
    pXMLPropSet As IXmlPropertySet2, _
    strXPath As String, _
    Optional booFailed As Boolean) As Boolean
    
    On Error GoTo ErrHandler
  
  ' GENERALLY CALLED BY OTHER FUNCTIONS
  
  Dim varVals As Variant
'  varVals = Array("placeholder")
'
'  ReturnLargestIndexValue = -1
'  Do Until IsEmpty(varVals)
'    ReturnLargestIndexValue = ReturnLargestIndexValue + 1
'    varVals = pPropSet.GetProperty(strXPath & "[" & CStr(ReturnLargestIndexValue) & "]")
'  Loop
  
  booFailed = False
  
  Dim booFoundDuplicate As Boolean
  booFoundDuplicate = False
  Dim booFoundDuplicateInStep As Boolean
  Dim strTestVal As String
  Dim lngIndex As Long
  
  For lngIndex = 0 To lngMaxIndex
    booFoundDuplicateInStep = True
    
    ' NAME
    varVals = pPropSet.GetProperty(strXPath & "[" & CStr(lngIndex) & "]/rpIndName")
    If IsEmpty(varVals) Then
      If strIndividualName <> "" Then booFoundDuplicateInStep = False
    Else
      strTestVal = CStr(varVals(0))
      If Trim(strIndividualName) <> Trim(strTestVal) Then booFoundDuplicateInStep = False
    End If
    
    If booFoundDuplicateInStep Then
      ' ORGANIZATION
      varVals = pPropSet.GetProperty(strXPath & "[" & CStr(lngIndex) & "]/rpOrgName")
      If IsEmpty(varVals) Then
        If strOrganizationName <> "" Then booFoundDuplicateInStep = False
      Else
        strTestVal = CStr(varVals(0))
        If Trim(strOrganizationName) <> Trim(strTestVal) Then booFoundDuplicateInStep = False
      End If
    End If
    
    If booFoundDuplicateInStep Then
      ' POSITION
      varVals = pPropSet.GetProperty(strXPath & "[" & CStr(lngIndex) & "]/rpPosName")
      If IsEmpty(varVals) Then
        If strPositionName <> "" Then booFoundDuplicateInStep = False
      Else
        strTestVal = CStr(varVals(0))
        If Trim(strPositionName) <> Trim(strTestVal) Then booFoundDuplicateInStep = False
      End If
    End If
      
    If booFoundDuplicateInStep Then
      ' VOICE NUMBER
      varVals = pPropSet.GetProperty(strXPath & "[" & CStr(lngIndex) & "]/rpCntInfo/cntPhone/voiceNum")
      If IsEmpty(varVals) Then
        If strVoiceNumber <> "" Then booFoundDuplicateInStep = False
      Else
        strTestVal = CStr(varVals(0))
        If Trim(strVoiceNumber) <> Trim(strTestVal) Then booFoundDuplicateInStep = False
      End If
    End If
      
    If booFoundDuplicateInStep Then
      ' ADDRESS STREET
      varVals = pPropSet.GetProperty(strXPath & "[" & CStr(lngIndex) & "]/rpCntInfo/cntAddress/delPoint")
      If IsEmpty(varVals) Then
        If strAddressStreet <> "" Then booFoundDuplicateInStep = False
      Else
        strTestVal = CStr(varVals(0))
        If Trim(strAddressStreet) <> Trim(strTestVal) Then booFoundDuplicateInStep = False
      End If
    End If
      
    If booFoundDuplicateInStep Then
      ' ADDRESS CITY
      varVals = pPropSet.GetProperty(strXPath & "[" & CStr(lngIndex) & "]/rpCntInfo/cntAddress/city")
      If IsEmpty(varVals) Then
        If strAddressCity <> "" Then booFoundDuplicateInStep = False
      Else
        strTestVal = CStr(varVals(0))
        If Trim(strAddressCity) <> Trim(strTestVal) Then booFoundDuplicateInStep = False
      End If
    End If
      
    If booFoundDuplicateInStep Then
      ' ADDRESS STATE
      varVals = pPropSet.GetProperty(strXPath & "[" & CStr(lngIndex) & "]/rpCntInfo/cntAddress/adminArea")
      If IsEmpty(varVals) Then
        If strAddressState <> "" Then booFoundDuplicateInStep = False
      Else
        strTestVal = CStr(varVals(0))
        If Trim(strAddressState) <> Trim(strTestVal) Then booFoundDuplicateInStep = False
      End If
    End If
      
    If booFoundDuplicateInStep Then
      ' ADDRESS ZIP
      varVals = pPropSet.GetProperty(strXPath & "[" & CStr(lngIndex) & "]/rpCntInfo/cntAddress/postCode")
      If IsEmpty(varVals) Then
        If strAddressZip <> "" Then booFoundDuplicateInStep = False
      Else
        strTestVal = CStr(varVals(0))
        If Trim(strAddressZip) <> Trim(strTestVal) Then booFoundDuplicateInStep = False
      End If
    End If
      
    If booFoundDuplicateInStep Then
      ' ADDRESS COUNTRY
      varVals = pPropSet.GetProperty(strXPath & "[" & CStr(lngIndex) & "]/rpCntInfo/cntAddress/country")
      If IsEmpty(varVals) Then
        If strAddressCountry <> "" Then booFoundDuplicateInStep = False
      Else
        strTestVal = CStr(varVals(0))
        If Trim(strAddressCountry) <> Trim(strTestVal) Then booFoundDuplicateInStep = False
      End If
    End If
      
    If booFoundDuplicateInStep Then
      ' ADDRESS EMAIL
      varVals = pPropSet.GetProperty(strXPath & "[" & CStr(lngIndex) & "]/rpCntInfo/cntAddress/eMailAdd")
      If IsEmpty(varVals) Then
        If strAddressEmail <> "" Then booFoundDuplicateInStep = False
      Else
        strTestVal = CStr(varVals(0))
        If Trim(strAddressEmail) <> Trim(strTestVal) Then booFoundDuplicateInStep = False
      End If
    End If
      
    If booFoundDuplicateInStep Then
      ' ROLE
      pXMLPropSet.GetAttribute strXPath & "[" & CStr(lngIndex) & "]/role/RoleCd", "value", varVals
      If IsEmpty(varVals) Then
        booFoundDuplicateInStep = False
      Else
        strTestVal = CStr(varVals(0))
        If Trim(ReturnRoleCDString(enumJenRole)) <> Trim(strTestVal) Then booFoundDuplicateInStep = False
      End If
    End If
      
    If booFoundDuplicateInStep Then
      ' ADDRESS TYPE
      pXMLPropSet.GetAttribute strXPath & "[" & CStr(lngIndex) & "]/rpCntInfo/cntAddress", "addressType", varVals
      If IsEmpty(varVals) Then
        If enumJenAddressType <> JenMetadata_Skip Then booFoundDuplicateInStep = False
      Else
        strTestVal = CStr(varVals(0))
        If Trim(ReturnAddressType(enumJenAddressType)) <> Trim(strTestVal) Then booFoundDuplicateInStep = False
      End If
    End If
    
  
'  pPropSet.SetProperty "mdContact[" & CStr(lngXIndex) & "]/rpIndName", strIndividualName
'  If strOrganizationName <> "" Then
'    pPropSet.SetProperty "mdContact[" & CStr(lngXIndex) & "]/rpOrgName", strOrganizationName
'  End If
'  If strPositionName <> "" Then
'    pPropSet.SetProperty "mdContact[" & CStr(lngXIndex) & "]/rpPosName", strPositionName
'  End If
'  If strVoiceNumber <> "" Then
'    pPropSet.SetProperty "mdContact[" & CStr(lngXIndex) & "]/rpCntInfo/cntPhone/voiceNum", strVoiceNumber
'  End If
'  If strAddressStreet <> "" Then
'    pPropSet.SetProperty "mdContact[" & CStr(lngXIndex) & "]/rpCntInfo/cntAddress/delPoint", strAddressStreet
'  End If
'  If strAddressCity <> "" Then
'    pPropSet.SetProperty "mdContact[" & CStr(lngXIndex) & "]/rpCntInfo/cntAddress/city", strAddressCity
'  End If
'  If strAddressState <> "" Then
'    pPropSet.SetProperty "mdContact[" & CStr(lngXIndex) & "]/rpCntInfo/cntAddress/adminArea", strAddressState
'  End If
'  If strAddressZip <> "" Then
'    pPropSet.SetProperty "mdContact[" & CStr(lngXIndex) & "]/rpCntInfo/cntAddress/postCode", strAddressZip
'  End If
'  If strAddressCountry <> "" Then
'    pPropSet.SetProperty "mdContact[" & CStr(lngXIndex) & "]/rpCntInfo/cntAddress/country", strAddressCountry
'  End If
'  If strAddressEmail <> "" Then
'    pPropSet.SetProperty "mdContact[" & CStr(lngXIndex) & "]/rpCntInfo/cntAddress/eMailAdd", strAddressEmail
'  End If
'
'  pPropSet.SetProperty "mdContact[" & CStr(lngXIndex) & "]/role/RoleCd", ""
'  pXMLPropSet.SetAttribute "mdContact[" & CStr(lngXIndex) & "]/role/RoleCd", "value", strRole, esriXSPAAddOrReplace
'  If enumJenAddressType <> JenMetadata_Skip Then
'    pXMLPropSet.SetAttribute "mdContact[" & CStr(lngXIndex) & _
'         "]/rpCntInfo/cntAddress", "addressType", strAddressType, esriXSPAAddOrReplace
'  End If
'
'  Metadata_Functions.SaveMetadata pDataset, pPropSet

    
    
    If booFoundDuplicateInStep Then
      booFoundDuplicate = True
      Exit For
    End If
  Next lngIndex
  
  CheckIfContactAlreadyPresent = booFoundDuplicate
  
  GoTo ClearMemory
  Exit Function
  
ErrHandler:
  CheckIfContactAlreadyPresent = False
  booFailed = True
  
ClearMemory:
  varVals = Null
  
End Function

Public Function AddContact_CitationResponsibleParty(pDataset As IDataset, _
    enumJenRole As JenMetadataRoleCDValues, _
    strIndividualName As String, _
    Optional booGetIndividualNameAutomatically As Boolean = False, _
    Optional strOrganizationName As String = "", _
    Optional strPositionName As String = "", _
    Optional strVoiceNumber As String = "", _
    Optional strAddressStreet As String = "", _
    Optional strAddressCity As String = "", _
    Optional strAddressState As String = "", _
    Optional strAddressZip As String = "", _
    Optional strAddressCountry As String = "", _
    Optional strAddressEmail As String = "", _
    Optional enumJenAddressType As JenMetadataAddressTypeValues = JenMetadata_Skip, _
    Optional booSkipIfAlreadyPresent As Boolean = True) As String
    
  On Error GoTo ErrHandler
    
  ' NEED CONTACT FOR METADATA, CITATION RESPONSIBLE PARTY AND RESOURCE POINT OF CONTACT
    ' ADD NEW CITATION CONTACTS; SHOULD HAVE ORIGINATOR
  '  strName = Linkages.aml_func_mod.GetTheUserName
  '  strResponse = Metadata_Functions.AddContact_CitationResponsibleParty(pDataset, JenMetadata_PointOfContact, _
  '    "Jeff Jenness or strName", False, "Jenness Enterprises", "Geoanalyst and Programmer", "1-928-607-4635", _
  '    "3020 N. Schevene Blvd.", "Flagstaff", "AZ", "86004", "USA", "jeffj@jennessent.com", JenMetadata_Postal)
  ' Debug.Print "Added Citation Contact = " & strResponse
    
  AddContact_CitationResponsibleParty = "Succeeded"
  
  If booGetIndividualNameAutomatically Then strIndividualName = aml_func_mod.GetTheUserName
    
  Dim pPropSet As IPropertySet
  Set pPropSet = ReturnMetadataPropSetFromDataset(pDataset)
  Dim pXMLPropSet As IXmlPropertySet2
  Set pXMLPropSet = pPropSet
  
  Dim strRole As String
  strRole = ReturnRoleCDString(enumJenRole)
  
  Dim strAddressType As String
  strAddressType = ReturnAddressType(enumJenAddressType)
    
  Dim lngXIndex As Long
  Dim strXPath As String
  strXPath = "dataIdInfo/idCitation/citRespParty"
  lngXIndex = ReturnLargestIndexValue(strXPath, pDataset) + 1
      
  Dim booContactAlreadyPresent As Boolean
  If booSkipIfAlreadyPresent Then
    booContactAlreadyPresent = CheckIfContactAlreadyPresent(pDataset, _
      enumJenRole, strIndividualName, strOrganizationName, _
       strPositionName, strVoiceNumber, strAddressStreet, strAddressCity, _
       strAddressState, strAddressZip, strAddressCountry, strAddressEmail, _
       enumJenAddressType, lngXIndex, pPropSet, pXMLPropSet, strXPath)
    
    If booContactAlreadyPresent Then
      AddContact_CitationResponsibleParty = "Citation Contact Already Present"
      GoTo ClearMemory
    End If
  End If
  
  pPropSet.SetProperty "dataIdInfo/idCitation/citRespParty[" & CStr(lngXIndex) & "]/rpIndName", strIndividualName
  If strOrganizationName <> "" Then
    pPropSet.SetProperty "dataIdInfo/idCitation/citRespParty[" & CStr(lngXIndex) & "]/rpOrgName", strOrganizationName
  End If
  If strPositionName <> "" Then
    pPropSet.SetProperty "dataIdInfo/idCitation/citRespParty[" & CStr(lngXIndex) & "]/rpPosName", strPositionName
  End If
  If strVoiceNumber <> "" Then
    pPropSet.SetProperty "dataIdInfo/idCitation/citRespParty[" & CStr(lngXIndex) & "]/rpCntInfo/cntPhone/voiceNum", strVoiceNumber
  End If
  If strAddressStreet <> "" Then
    pPropSet.SetProperty "dataIdInfo/idCitation/citRespParty[" & CStr(lngXIndex) & "]/rpCntInfo/cntAddress/delPoint", strAddressStreet
  End If
  If strAddressCity <> "" Then
    pPropSet.SetProperty "dataIdInfo/idCitation/citRespParty[" & CStr(lngXIndex) & "]/rpCntInfo/cntAddress/city", strAddressCity
  End If
  If strAddressState <> "" Then
    pPropSet.SetProperty "dataIdInfo/idCitation/citRespParty[" & CStr(lngXIndex) & "]/rpCntInfo/cntAddress/adminArea", strAddressState
  End If
  If strAddressZip <> "" Then
    pPropSet.SetProperty "dataIdInfo/idCitation/citRespParty[" & CStr(lngXIndex) & "]/rpCntInfo/cntAddress/postCode", strAddressZip
  End If
  If strAddressCountry <> "" Then
    pPropSet.SetProperty "dataIdInfo/idCitation/citRespParty[" & CStr(lngXIndex) & "]/rpCntInfo/cntAddress/country", strAddressCountry
  End If
  If strAddressEmail <> "" Then
    pPropSet.SetProperty "dataIdInfo/idCitation/citRespParty[" & CStr(lngXIndex) & "]/rpCntInfo/cntAddress/eMailAdd", strAddressEmail
  End If
  
  pPropSet.SetProperty "dataIdInfo/idCitation/citRespParty[" & CStr(lngXIndex) & "]/role/RoleCd", ""
  pXMLPropSet.SetAttribute "dataIdInfo/idCitation/citRespParty[" & CStr(lngXIndex) & "]/role/RoleCd", "value", strRole, esriXSPAAddOrReplace
  If enumJenAddressType <> JenMetadata_Skip Then
    pXMLPropSet.SetAttribute "dataIdInfo/idCitation/citRespParty[" & CStr(lngXIndex) & _
         "]/rpCntInfo/cntAddress", "addressType", strAddressType, esriXSPAAddOrReplace
  End If

  Metadata_Functions.SaveMetadata pDataset, pPropSet

  GoTo ClearMemory
  Exit Function
  
ErrHandler:
  AddContact_CitationResponsibleParty = "Failed"

ClearMemory:
  Set pPropSet = Nothing
  Set pXMLPropSet = Nothing
End Function

Public Function AddContact_ResourcePointOfContact(pDataset As IDataset, _
    enumJenRole As JenMetadataRoleCDValues, _
    strIndividualName As String, _
    Optional booGetIndividualNameAutomatically As Boolean = False, _
    Optional strOrganizationName As String = "", _
    Optional strPositionName As String = "", _
    Optional strVoiceNumber As String = "", _
    Optional strAddressStreet As String = "", _
    Optional strAddressCity As String = "", _
    Optional strAddressState As String = "", _
    Optional strAddressZip As String = "", _
    Optional strAddressCountry As String = "", _
    Optional strAddressEmail As String = "", _
    Optional enumJenAddressType As JenMetadataAddressTypeValues = JenMetadata_Skip, _
    Optional booSkipIfAlreadyPresent As Boolean = True) As String
    
  On Error GoTo ErrHandler
    
  ' NEED CONTACT FOR METADATA, CITATION RESPONSIBLE PARTY AND RESOURCE POINT OF CONTACT
    ' ADD NEW CITATION CONTACTS; SHOULD HAVE ORIGINATOR
  '  strName = aml_func_mod.GetTheUserName
  '  strResponse = Metadata_Functions.AddContact_ResourcePointOfContact(pDataset, JenMetadata_PointOfContact, _
  '    "Jeff Jenness or strName", False,  "Jenness Enterprises", "Geoanalyst and Programmer", "1-928-607-4635", _
  '    "3020 N. Schevene Blvd.", "Flagstaff", "AZ", "86004", "USA", "jeffj@jennessent.com", JenMetadata_Postal)
  ' Debug.Print "Added Resource Contact = " & strResponse
  
  AddContact_ResourcePointOfContact = "Succeeded"
  
  If booGetIndividualNameAutomatically Then strIndividualName = aml_func_mod.GetTheUserName
  
  Dim pPropSet As IPropertySet
  Set pPropSet = ReturnMetadataPropSetFromDataset(pDataset)
  Dim pXMLPropSet As IXmlPropertySet2
  Set pXMLPropSet = pPropSet
  
  Dim strRole As String
  strRole = ReturnRoleCDString(enumJenRole)
  
  Dim strAddressType As String
  strAddressType = ReturnAddressType(enumJenAddressType)
    
  Dim lngXIndex As Long
  Dim strXPath As String
  strXPath = "dataIdInfo/idPoC"
  lngXIndex = ReturnLargestIndexValue(strXPath, pDataset) + 1
    
  Dim booContactAlreadyPresent As Boolean
  If booSkipIfAlreadyPresent Then
    booContactAlreadyPresent = CheckIfContactAlreadyPresent(pDataset, _
      enumJenRole, strIndividualName, strOrganizationName, _
       strPositionName, strVoiceNumber, strAddressStreet, strAddressCity, _
       strAddressState, strAddressZip, strAddressCountry, strAddressEmail, _
       enumJenAddressType, lngXIndex, pPropSet, pXMLPropSet, strXPath)
    
    If booContactAlreadyPresent Then
      AddContact_ResourcePointOfContact = "Resource Contact Already Present"
      GoTo ClearMemory
    End If
  End If
  
  pPropSet.SetProperty "dataIdInfo/idPoC[" & CStr(lngXIndex) & "]/rpIndName", strIndividualName
  If strOrganizationName <> "" Then
    pPropSet.SetProperty "dataIdInfo/idPoC[" & CStr(lngXIndex) & "]/rpOrgName", strOrganizationName
  End If
  If strPositionName <> "" Then
    pPropSet.SetProperty "dataIdInfo/idPoC[" & CStr(lngXIndex) & "]/rpPosName", strPositionName
  End If
  If strVoiceNumber <> "" Then
    pPropSet.SetProperty "dataIdInfo/idPoC[" & CStr(lngXIndex) & "]/rpCntInfo/cntPhone/voiceNum", strVoiceNumber
  End If
  If strAddressStreet <> "" Then
    pPropSet.SetProperty "dataIdInfo/idPoC[" & CStr(lngXIndex) & "]/rpCntInfo/cntAddress/delPoint", strAddressStreet
  End If
  If strAddressCity <> "" Then
    pPropSet.SetProperty "dataIdInfo/idPoC[" & CStr(lngXIndex) & "]/rpCntInfo/cntAddress/city", strAddressCity
  End If
  If strAddressState <> "" Then
    pPropSet.SetProperty "dataIdInfo/idPoC[" & CStr(lngXIndex) & "]/rpCntInfo/cntAddress/adminArea", strAddressState
  End If
  If strAddressZip <> "" Then
    pPropSet.SetProperty "dataIdInfo/idPoC[" & CStr(lngXIndex) & "]/rpCntInfo/cntAddress/postCode", strAddressZip
  End If
  If strAddressCountry <> "" Then
    pPropSet.SetProperty "dataIdInfo/idPoC[" & CStr(lngXIndex) & "]/rpCntInfo/cntAddress/country", strAddressCountry
  End If
  If strAddressEmail <> "" Then
    pPropSet.SetProperty "dataIdInfo/idPoC[" & CStr(lngXIndex) & "]/rpCntInfo/cntAddress/eMailAdd", strAddressEmail
  End If
  
  pPropSet.SetProperty "dataIdInfo/idPoC[" & CStr(lngXIndex) & "]/role/RoleCd", ""
  pXMLPropSet.SetAttribute "dataIdInfo/idPoC[" & CStr(lngXIndex) & "]/role/RoleCd", "value", strRole, esriXSPAAddOrReplace
  If enumJenAddressType <> JenMetadata_Skip Then
    pXMLPropSet.SetAttribute "dataIdInfo/idPoC[" & CStr(lngXIndex) & _
         "]/rpCntInfo/cntAddress", "addressType", strAddressType, esriXSPAAddOrReplace
  End If

  Metadata_Functions.SaveMetadata pDataset, pPropSet

  GoTo ClearMemory
  Exit Function
  
ErrHandler:
  AddContact_ResourcePointOfContact = "Failed"

ClearMemory:
  Set pPropSet = Nothing
  Set pXMLPropSet = Nothing
End Function


Public Function AddNewGeoProcStep(pDataset As IDataset, _
    strDescription As String, _
    strToolSource As String, _
    datDate As Date, _
    strProcessName As String, _
    booShouldExport As Boolean) As String
  
  On Error GoTo ErrHandler
  
'  ' ADD NEW GEOPROCESSING EVENT
'  strResponse = AddNewGeoProcStep(pDataset, "NOTE:  This is not Python code! " & _
'      "  Parameters used in analysis...", app.Path & "\" & App.EXEName & ".dll\Tool_Name", _
'      Now, "Extension Name, then Tool Name", False)
'  Debug.Print "Added new Geoprocessing Event = " & strResponse
  
  AddNewGeoProcStep = "Succeeded"
  
  Dim pPropSet As IPropertySet
  Set pPropSet = ReturnMetadataPropSetFromDataset(pDataset)
  Dim pXMLPropSet As IXmlPropertySet2
  Set pXMLPropSet = pPropSet
  
  Dim strDate As String
  Dim strTime As String
  
  strDate = Format(datDate, "yyyymmdd")
  strTime = Format(datDate, "HhNnSs")
  
  Dim strExport As String
  If booShouldExport Then
    strExport = "True"
  Else
    strExport = ""
  End If
  
  Dim lngLineageIndex As Long
  Dim strXPath As String
  strXPath = "Esri/DataProperties/lineage/Process"
  lngLineageIndex = ReturnLargestIndexValue(strXPath, pDataset) + 1
    
  pPropSet.SetProperty "Esri/DataProperties/lineage/Process[" & CStr(lngLineageIndex) & _
      "]", strDescription

  pXMLPropSet.SetAttribute "Esri/DataProperties/lineage/Process[" & CStr(lngLineageIndex) & _
      "]", "ToolSource", strToolSource, esriXSPAAddOrReplace
  pXMLPropSet.SetAttribute "Esri/DataProperties/lineage/Process[" & CStr(lngLineageIndex) & _
      "]", "Date", strDate, esriXSPAAddOrReplace
  pXMLPropSet.SetAttribute "Esri/DataProperties/lineage/Process[" & CStr(lngLineageIndex) & _
      "]", "Time", strTime, esriXSPAAddOrReplace
  pXMLPropSet.SetAttribute "Esri/DataProperties/lineage/Process[" & CStr(lngLineageIndex) & _
      "]", "Name", strProcessName, esriXSPAAddOrReplace
  pXMLPropSet.SetAttribute "Esri/DataProperties/lineage/Process[" & CStr(lngLineageIndex) & _
      "]", "export", strExport, esriXSPAAddOrReplace
  
  Metadata_Functions.SaveMetadata pDataset, pPropSet

  GoTo ClearMemory
  Exit Function
  
ErrHandler:
  AddNewGeoProcStep = "Failed"

ClearMemory:
  Set pPropSet = Nothing
  Set pXMLPropSet = Nothing
  
End Function
Public Function CopyLineageOrGeoprocHistory(pSourceDataset As IDataset, pRecipientDataset As IDataset, _
    booEraseExistingInRecipient As Boolean, enumLineageOrGeoproc As JenMetadataLineageOrGeoprocValues) As String
    
'  strResponse = CopyLineageOrGeoprocHistory(pSourceDataset, pRecipientDataset, False, JenMetadata_GeoprocessingHistory)
'  Debug.Print "Geoproc History Transfer Successful = " & strResponse

  Dim strSourceXML As String
  strSourceXML = ReturnExistingLineageOrGeoprocessingHistory(pSourceDataset, enumLineageOrGeoproc)
  
  If enumLineageOrGeoproc = JenMetadata_Lineage Then
    CopyLineageOrGeoprocHistory = InsertExistingLineageHistory(pRecipientDataset, strSourceXML, booEraseExistingInRecipient)
  Else
    CopyLineageOrGeoprocHistory = InsertExistingGeoProcessingHistory(pRecipientDataset, strSourceXML, booEraseExistingInRecipient)
  End If
  
End Function

Private Function ReturnExistingLineageOrGeoprocessingHistory(pDataset As IDataset, _
     enumLineageOrGeoproc As JenMetadataLineageOrGeoprocValues) As String
  On Error GoTo ErrHandler:
  
  ' GENERALLY CALLED BY:
  '  CopyLineageOrGeoprocHistory
  '  InsertExistingGeoProcessingHistory
  '  InsertExistingLineageHistory
  
  Dim pPropSet As IPropertySet
  Set pPropSet = ReturnMetadataPropSetFromDataset(pDataset)
  Dim pXMLPropSet As IXmlPropertySet2
  Set pXMLPropSet = pPropSet
  Dim strXPath As String
  If enumLineageOrGeoproc = JenMetadata_Lineage Then
    strXPath = "dqInfo/dataLineage"
  Else
    strXPath = "Esri/DataProperties/lineage"
  End If
  
  Dim strXML As String
  strXML = pXMLPropSet.GetXml(strXPath)
  ReturnExistingLineageOrGeoprocessingHistory = strXML
  
'  Debug.Print strXML
  
  GoTo ClearMemory
  Exit Function

ErrHandler:
  ReturnExistingLineageOrGeoprocessingHistory = ""

ClearMemory:
  Set pPropSet = Nothing
  Set pXMLPropSet = Nothing
End Function

Private Function InsertExistingGeoProcessingHistory(pDataset As IDataset, strLineageHistory As String, _
    booEraseExistingInRecipient As Boolean) As String
  On Error GoTo ErrHandler:
  
  ' GENERALLY CALLED BY:
  '  CopyLineageOrGeoprocHistory
  
  InsertExistingGeoProcessingHistory = "Succeeded"
    
  Dim pPropSet As IPropertySet
  Set pPropSet = ReturnMetadataPropSetFromDataset(pDataset)
  Dim pXMLPropSet As IXmlPropertySet2
  Set pXMLPropSet = pPropSet
  
  Dim strCurrentLineage As String
  
  Dim strFullXML As String
  Dim lngStartSplice As Long
  Dim lngEndSplice As Long
  Dim lngInsertPoint As Long
  Dim strTempXML As String
  Dim strIntermXML As String
  
  strFullXML = pXMLPropSet.GetXml("")
  
  If booEraseExistingInRecipient Then
    lngStartSplice = InStr(1, strFullXML, "<lineage>", vbTextCompare)
    If lngStartSplice = 0 Then
      pPropSet.SetProperty "Esri/DataProperties/lineage/Process[0]", "placeholder"
      strFullXML = pXMLPropSet.GetXml("")
      lngStartSplice = InStr(1, strFullXML, "<lineage>", vbTextCompare)
    End If
    
    lngEndSplice = InStr(1, strFullXML, "</lineage>", vbTextCompare)
    
    lngInsertPoint = lngStartSplice
    
    strTempXML = strFullXML
    Do Until lngStartSplice = 0
      strTempXML = Left(strTempXML, lngStartSplice - 1) & Right(strTempXML, Len(strTempXML) - lngEndSplice - 9)
      lngStartSplice = InStr(1, strTempXML, "<lineage>", vbTextCompare)
      lngEndSplice = InStr(1, strTempXML, "</lineage>", vbTextCompare)
    Loop
    
    strTempXML = Left(strTempXML, lngInsertPoint - 1) & strLineageHistory & _
        Right(strTempXML, Len(strTempXML) - lngInsertPoint + 1)
  Else
    
    ' GET XML LINEAGE TEXT FROM CURRENT DATASET
    strCurrentLineage = ReturnExistingLineageOrGeoprocessingHistory(pDataset, JenMetadata_GeoprocessingHistory)
    lngInsertPoint = InStr(1, strCurrentLineage, "</lineage>", vbTextCompare)
    
    ' IF IT DOESN'T EXIST, CREATE NEW INSERT POINT BASED ON FULL XML STRING AND PASTE IN LINEAGE
    If lngInsertPoint = 0 Then   ' IF THERE IS NO "Lineage" IN METADATA
      pPropSet.SetProperty "Esri/DataProperties/lineage/Process[0]", "placeholder"
      strFullXML = pXMLPropSet.GetXml("")
      lngStartSplice = InStr(1, strFullXML, "<lineage>", vbTextCompare)
        
      lngEndSplice = InStr(1, strFullXML, "</lineage>", vbTextCompare)
      
      lngInsertPoint = lngStartSplice
      
      strTempXML = strFullXML
      Do Until lngStartSplice = 0
        strTempXML = Left(strTempXML, lngStartSplice - 1) & Right(strTempXML, Len(strTempXML) - lngEndSplice - 9)
        lngStartSplice = InStr(1, strTempXML, "<lineage>", vbTextCompare)
        lngEndSplice = InStr(1, strTempXML, "</lineage>", vbTextCompare)
      Loop
    
      strTempXML = Left(strTempXML, lngInsertPoint - 1) & strLineageHistory & _
          Right(strTempXML, Len(strTempXML) - lngInsertPoint + 1)
    Else
      
      ' IF IT DOES EXIST, THEN NEED TO TRIM OFF "dataLineage" TAGS AND STICK HISTORY INSIDE OF CURRENT LINEAGE
      strLineageHistory = Replace(strLineageHistory, "<lineage>", "", , , vbBinaryCompare)
      strLineageHistory = Replace(strLineageHistory, "</lineage>", "", , , vbBinaryCompare)
      
      strIntermXML = Left(strCurrentLineage, lngInsertPoint - 1) & strLineageHistory & _
          Right(strCurrentLineage, Len(strCurrentLineage) - lngInsertPoint + 1)
      
      ' NOW REPLACE CURRENT LINEAGE IN FULL XML TEXT
      lngStartSplice = InStr(1, strFullXML, "<lineage>", vbTextCompare)
      If lngStartSplice = 0 Then
        pPropSet.SetProperty "Esri/DataProperties/lineage/Process[0]", "placeholder"
        strFullXML = pXMLPropSet.GetXml("")
        lngStartSplice = InStr(1, strFullXML, "<lineage>", vbTextCompare)
      End If
      
      lngEndSplice = InStr(1, strFullXML, "</lineage>", vbTextCompare)
      
      lngInsertPoint = lngStartSplice
      strTempXML = strFullXML
      Do Until lngStartSplice = 0
        strTempXML = Left(strTempXML, lngStartSplice - 1) & Right(strTempXML, Len(strTempXML) - lngEndSplice - 9)
        lngStartSplice = InStr(1, strTempXML, "<lineage>", vbTextCompare)
        lngEndSplice = InStr(1, strTempXML, "</lineage>", vbTextCompare)
      Loop
      
      strTempXML = Left(strTempXML, lngInsertPoint - 1) & strIntermXML & _
          Right(strTempXML, Len(strTempXML) - lngInsertPoint + 1)
        
    End If
  End If
  
  pXMLPropSet.SetXml strTempXML
    
  
  Metadata_Functions.SaveMetadata pDataset, pPropSet
  
  GoTo ClearMemory
  Exit Function

ErrHandler:
  InsertExistingGeoProcessingHistory = "Failed"

ClearMemory:
  Set pPropSet = Nothing
  Set pXMLPropSet = Nothing
  
End Function

Private Function InsertExistingLineageHistory(pDataset As IDataset, strLineageHistory As String, _
    booEraseExistingInRecipient As Boolean) As String
  On Error GoTo ErrHandler:
  
  ' GENERALLY CALLED BY:
  '  CopyLineageOrGeoprocHistory
  
  InsertExistingLineageHistory = "Succeeded"
    
  Dim pPropSet As IPropertySet
  Set pPropSet = ReturnMetadataPropSetFromDataset(pDataset)
  Dim pXMLPropSet As IXmlPropertySet2
  Set pXMLPropSet = pPropSet
  
  Dim strCurrentLineage As String
  
  Dim strFullXML As String
  Dim lngStartSplice As Long
  Dim lngEndSplice As Long
  Dim lngInsertPoint As Long
  Dim strTempXML As String
  Dim strIntermXML As String
  
  strFullXML = pXMLPropSet.GetXml("")
  
  If booEraseExistingInRecipient Then
    lngStartSplice = InStr(1, strFullXML, "<dataLineage>", vbTextCompare)
    If lngStartSplice = 0 Then
      pPropSet.SetProperty "dqInfo/dataLineage/prcStep[0]/stepDesc", "placeholder"
      strFullXML = pXMLPropSet.GetXml("")
      lngStartSplice = InStr(1, strFullXML, "<dataLineage>", vbTextCompare)
    End If
    
    lngEndSplice = InStr(1, strFullXML, "</dataLineage>", vbTextCompare)
    
    lngInsertPoint = lngStartSplice
    
    strTempXML = strFullXML
    Do Until lngStartSplice = 0
      strTempXML = Left(strTempXML, lngStartSplice - 1) & Right(strTempXML, Len(strTempXML) - lngEndSplice - 13)
      lngStartSplice = InStr(1, strTempXML, "<dataLineage>", vbTextCompare)
      lngEndSplice = InStr(1, strTempXML, "</dataLineage>", vbTextCompare)
    Loop
    
    strTempXML = Left(strTempXML, lngInsertPoint - 1) & strLineageHistory & _
        Right(strTempXML, Len(strTempXML) - lngInsertPoint + 1)
  Else
    
    ' GET XML LINEAGE TEXT FROM CURRENT DATASET
    strCurrentLineage = ReturnExistingLineageOrGeoprocessingHistory(pDataset, JenMetadata_Lineage)
    lngInsertPoint = InStr(1, strCurrentLineage, "</dataLineage>", vbTextCompare)
    
    ' IF IT DOESN'T EXIST, CREATE NEW INSERT POINT BASED ON FULL XML STRING AND PASTE IN LINEAGE
    If lngInsertPoint = 0 Then
      
      strTempXML = pXMLPropSet.GetXml("")
      lngInsertPoint = InStr(1, strTempXML, "</metadata>", vbTextCompare)
      strTempXML = Left(strTempXML, lngInsertPoint - 1) & "<dqInfo>" & strLineageHistory & "</dqInfo>" & _
          Right(strTempXML, Len(strTempXML) - lngInsertPoint + 1)
    Else
      
      
      ' IF IT DOES EXIST, THEN NEED TO TRIM OFF "dataLineage" TAGS AND STICK HISTORY INSIDE OF CURRENT LINEAGE
      strLineageHistory = Replace(strLineageHistory, "<dataLineage>", "", , , vbTextCompare)
      strLineageHistory = Replace(strLineageHistory, "</dataLineage>", "", , , vbTextCompare)
      
      strIntermXML = Left(strCurrentLineage, lngInsertPoint - 1) & strLineageHistory & _
          Right(strCurrentLineage, Len(strCurrentLineage) - lngInsertPoint + 1)
      
      ' NOW REPLACE CURRENT LINEAGE IN FULL XML TEXT
      lngStartSplice = InStr(1, strFullXML, "<dataLineage>", vbTextCompare)
      If lngStartSplice = 0 Then
        pPropSet.SetProperty "dqInfo/dataLineage/prcStep[0]/stepDesc", "placeholder"
        strFullXML = pXMLPropSet.GetXml("")
        lngStartSplice = InStr(1, strFullXML, "<dataLineage>", vbTextCompare)
      End If
      
      lngEndSplice = InStr(1, strFullXML, "</dataLineage>", vbTextCompare)
      
      lngInsertPoint = lngStartSplice
      strTempXML = strFullXML
      Do Until lngStartSplice = 0
        strTempXML = Left(strTempXML, lngStartSplice - 1) & Right(strTempXML, Len(strTempXML) - lngEndSplice - 13)
        lngStartSplice = InStr(1, strTempXML, "<dataLineage>", vbTextCompare)
        lngEndSplice = InStr(1, strTempXML, "</dataLineage>", vbTextCompare)
      Loop
      
      strTempXML = Left(strTempXML, lngInsertPoint - 1) & strIntermXML & _
          Right(strTempXML, Len(strTempXML) - lngInsertPoint + 1)
        
    End If
  End If
  
  pXMLPropSet.SetXml strTempXML
      
  Metadata_Functions.SaveMetadata pDataset, pPropSet
  
  GoTo ClearMemory
  Exit Function

ErrHandler:
  InsertExistingLineageHistory = "Failed"

ClearMemory:
  Set pPropSet = Nothing
  Set pXMLPropSet = Nothing
  
End Function

Public Function AddNewLineageStep(pDataset As IDataset, _
    strDescription As String, _
    datDate As Date, _
    enumJenRole As JenMetadataRoleCDValues, _
    strIndividualName As String, _
    Optional booGetIndividualNameAutomatically As Boolean = False, _
    Optional strOrganizationName As String = "", _
    Optional strPositionName As String = "", _
    Optional strVoiceNumber As String = "", _
    Optional strAddressStreet As String = "", _
    Optional strAddressCity As String = "", _
    Optional strAddressState As String = "", _
    Optional strAddressZip As String = "", _
    Optional strAddressCountry As String = "", _
    Optional strAddressEmail As String = "", _
    Optional enumJenAddressType As JenMetadataAddressTypeValues = JenMetadata_Skip) As String
  
  On Error GoTo ErrHandler
    
'  strDescription = "Lengthy description of process, along with selected parameter values."
'  strName = aml_func_mod.GetTheUserName
'  strResponse = Metadata_Functions.AddNewLineageStep(pDataset, strDescription, Now, JenMetadata_PointOfContact, _
'      "Jeff Jenness or strName", "Jenness Enterprises", "Geoanalyst and Programmer", "1-928-607-4635", _
'      "3020 N. Schevene Blvd.", "Flagstaff", "AZ", "86004", "USA", "jeffj@jennessent.com", JenMetadata_Postal)
'  Debug.Print "Lineage Successful = " & strResponse
  
  AddNewLineageStep = "Succeeded"
  
  If booGetIndividualNameAutomatically Then strIndividualName = aml_func_mod.GetTheUserName
  
  Dim pPropSet As IPropertySet
  Set pPropSet = ReturnMetadataPropSetFromDataset(pDataset)
  Dim pXMLPropSet As IXmlPropertySet2
  Set pXMLPropSet = pPropSet
  
  Dim strRole As String
  strRole = ReturnRoleCDString(enumJenRole)
  
  Dim strAddressType As String
  strAddressType = ReturnAddressType(enumJenAddressType)
    
  Dim lngLineageIndex As Long
  Dim strXPath As String
  strXPath = "dqInfo/dataLineage/prcStep"
  lngLineageIndex = ReturnLargestIndexValue(strXPath, pDataset) + 1
  
'  Dim varVals As Variant
'  varVals = Array("placeholder")
'
'  lngLineageIndex = -1
'  Do Until IsEmpty(varVals)
'    lngLineageIndex = lngLineageIndex + 1
'    varVals = pPropSet.GetProperty("dqInfo/dataLineage/prcStep[" & CStr(lngLineageIndex) & "]")
'  Loop
  
  
  
  ' RoleCD values:
  '
  ' 006 = Originator
  ' 007 = Point of Contact
  ' 009 = Processor
  ' 010 = Publisher    ....   "A cited responsible party is added with the publisher " & _
  '                           "role to contain all information pertaining to the publisher. " & _
  '                           "The contents of the publication place element is placed in the " & _
  '                           "address's delivery point element. This may not be correct in all cases."

  
  ' AscTypeCD values:
  ' 001 = Cross Reference
  ' 002 - LargerWorkCitation
  
  ' MaintFreqCD values:
  
  ' RestrictCD values:
  
  ' ProgCD values:
  
  ' ClasscationCD values:  (sic)
  
  ' PresFormCD values:
  
  ' MedNameCD values:
  
  ' SpatRepTypCD values:  "Values will be either vector or grid; point = vector."
  
  ' GeoObjTypCD values:  "Related information with ArcGIS terminology recorded at " & _\
  '     "/metadata/spdoinfo/ptvctinf/esriterm, with feature type at efeageom/@code and feature count in efeacnt."

  ' TopoLevCD values:  "ArcGIS-related topology information is recorded at " & _
  '     "/metadata/spdoinfo/ptvctinf/esriterm/esritopo for all vector data formats.
  
  pPropSet.RemoveProperty "dqInfo/dataLineage/prcStep[" & CStr(lngLineageIndex) & "]/stepDesc"
  pPropSet.RemoveProperty "dqInfo/dataLineage/prcStep[" & CStr(lngLineageIndex) & "]/stepDateTm"
  pPropSet.RemoveProperty "dqInfo/dataLineage/prcStep[" & CStr(lngLineageIndex) & "]/stepProc/rpIndName"
  pPropSet.RemoveProperty "dqInfo/dataLineage/prcStep[" & CStr(lngLineageIndex) & "]/stepProc/rpOrgName"
  pPropSet.RemoveProperty "dqInfo/dataLineage/prcStep[" & CStr(lngLineageIndex) & "]/stepProc/rpPosName"
  pPropSet.RemoveProperty "dqInfo/dataLineage/prcStep[" & CStr(lngLineageIndex) & "]/stepProc/rpCntInfo/cntPhone/voiceNum"
  pPropSet.RemoveProperty "dqInfo/dataLineage/prcStep[" & CStr(lngLineageIndex) & "]/stepProc/rpCntInfo/cntAddress/delPoint"
  pPropSet.RemoveProperty "dqInfo/dataLineage/prcStep[" & CStr(lngLineageIndex) & "]/stepProc/rpCntInfo/cntAddress/city"
  pPropSet.RemoveProperty "dqInfo/dataLineage/prcStep[" & CStr(lngLineageIndex) & "]/stepProc/rpCntInfo/cntAddress/adminArea"
  pPropSet.RemoveProperty "dqInfo/dataLineage/prcStep[" & CStr(lngLineageIndex) & "]/stepProc/rpCntInfo/cntAddress/postCode"
  pPropSet.RemoveProperty "dqInfo/dataLineage/prcStep[" & CStr(lngLineageIndex) & "]/stepProc/rpCntInfo/cntAddress/country"
  pPropSet.RemoveProperty "dqInfo/dataLineage/prcStep[" & CStr(lngLineageIndex) & "]/stepProc/rpCntInfo/cntAddress/eMailAdd"
  pPropSet.RemoveProperty "dqInfo/dataLineage/prcStep[" & CStr(lngLineageIndex) & "]/stepProc/role/RoleCd"
  pPropSet.RemoveProperty "dqInfo/dataLineage/prcStep[" & CStr(lngLineageIndex) & "]/stepProc/displayName"
    
  pPropSet.SetProperty "dqInfo/dataLineage/prcStep[" & CStr(lngLineageIndex) & _
      "]/stepDesc", strDescription
  pPropSet.SetProperty "dqInfo/dataLineage/prcStep[" & CStr(lngLineageIndex) & _
       "]/stepDateTm", Format(datDate, "yyyy-mm-ddTHh:Nn:Ss") ' "2013-09-01T00:00:00"
  pPropSet.SetProperty "dqInfo/dataLineage/prcStep[" & CStr(lngLineageIndex) & _
       "]/stepProc/rpIndName", strIndividualName
  If strOrganizationName <> "" Then
    pPropSet.SetProperty "dqInfo/dataLineage/prcStep[" & CStr(lngLineageIndex) & _
         "]/stepProc/rpOrgName", strOrganizationName
  End If
  If strPositionName <> "" Then
    pPropSet.SetProperty "dqInfo/dataLineage/prcStep[" & CStr(lngLineageIndex) & _
         "]/stepProc/rpPosName", strPositionName
  End If
  If strVoiceNumber <> "" Then
    pPropSet.SetProperty "dqInfo/dataLineage/prcStep[" & CStr(lngLineageIndex) & _
         "]/stepProc/rpCntInfo/cntPhone/voiceNum", strVoiceNumber
  End If
  If strAddressStreet <> "" Then
    pPropSet.SetProperty "dqInfo/dataLineage/prcStep[" & CStr(lngLineageIndex) & _
         "]/stepProc/rpCntInfo/cntAddress/delPoint", strAddressStreet
  End If
  If strAddressCity <> "" Then
    pPropSet.SetProperty "dqInfo/dataLineage/prcStep[" & CStr(lngLineageIndex) & _
         "]/stepProc/rpCntInfo/cntAddress/city", strAddressCity
  End If
  If strAddressState <> "" Then
    pPropSet.SetProperty "dqInfo/dataLineage/prcStep[" & CStr(lngLineageIndex) & _
         "]/stepProc/rpCntInfo/cntAddress/adminArea", strAddressState
  End If
  If strAddressZip <> "" Then
    pPropSet.SetProperty "dqInfo/dataLineage/prcStep[" & CStr(lngLineageIndex) & _
         "]/stepProc/rpCntInfo/cntAddress/postCode", strAddressZip
  End If
  If strAddressCountry <> "" Then
    pPropSet.SetProperty "dqInfo/dataLineage/prcStep[" & CStr(lngLineageIndex) & _
         "]/stepProc/rpCntInfo/cntAddress/country", strAddressCountry
  End If
  If strAddressEmail <> "" Then
    pPropSet.SetProperty "dqInfo/dataLineage/prcStep[" & CStr(lngLineageIndex) & _
         "]/stepProc/rpCntInfo/cntAddress/eMailAdd", strAddressEmail
  End If
  
  pPropSet.SetProperty "dqInfo/dataLineage/prcStep[" & CStr(lngLineageIndex) & _
       "]/stepProc/role/RoleCd", ""
       
  pPropSet.SetProperty "dqInfo/dataLineage/prcStep[" & CStr(lngLineageIndex) & _
       "]/stepProc/displayName", strIndividualName

  pXMLPropSet.SetAttribute "dqInfo/dataLineage/prcStep[" & CStr(lngLineageIndex) & _
      "]/stepProc/role/RoleCd", "value", strRole, esriXSPAAddOrReplace
  
  If enumJenAddressType <> JenMetadata_Skip Then
    pXMLPropSet.SetAttribute "dqInfo/dataLineage/prcStep[" & CStr(lngLineageIndex) & _
         "]/stepProc/rpCntInfo/cntAddress", "addressType", strAddressType, esriXSPAAddOrReplace
  End If

  Metadata_Functions.SaveMetadata pDataset, pPropSet

  GoTo ClearMemory
  Exit Function
  
ErrHandler:
  AddNewLineageStep = "Failed"

ClearMemory:
  Set pPropSet = Nothing
  Set pXMLPropSet = Nothing
  
End Function

Private Function ReturnLargestIndexValue(strXPath As String, pDataset As IDataset, _
    Optional booFailed As Boolean) As Long
  
  On Error GoTo ErrHandler
  
  ' GENERALLY CALLED BY OTHER FUNCTIONS
  booFailed = False
  Dim pPropSet As IPropertySet
  Set pPropSet = ReturnMetadataPropSetFromDataset(pDataset)
  
  Dim varVals As Variant
  varVals = Array("placeholder")
  
  ReturnLargestIndexValue = -1
  Do Until IsEmpty(varVals)
    ReturnLargestIndexValue = ReturnLargestIndexValue + 1
    varVals = pPropSet.GetProperty(strXPath & "[" & CStr(ReturnLargestIndexValue) & "]")
  Loop
  ReturnLargestIndexValue = ReturnLargestIndexValue - 1
    
  GoTo ClearMemory
  Exit Function
  
ErrHandler:
  ReturnLargestIndexValue = -1
  booFailed = True
  
ClearMemory:
  Set pPropSet = Nothing
  varVals = Null
  
End Function
Public Function ReturnAllMetadataPropertiesFromDataset(pDataset As IDataset, _
    Optional booFailed As Boolean) As String

  On Error GoTo ErrHandler
  
'  SAMPLE CODE
'  Dim pFLayer As IFeatureLayer
'  Dim pMxDoc As IMxDocument
'  Set pMxDoc = ThisDocument
'  Set pFLayer = mygeneralOperations.ReturnLayerByName("Roads_for_Metadata", pMxDoc.FocusMap)
'  Dim pFClass As IFeatureClass
'
'  Set pFClass = pFLayer.FeatureClass
'  Dim pDataset As IDataset
'  Set pDataset = pFClass
'
'  Dim strAllProps As String
'  Dim strXML As String
'
'  strXML = Metadata_Functions.ReturnMetadataXMLStringFromDataset(pDataset)
'  strAllProps = Metadata_Functions.ReturnAllMetadataPropertiesFromDataset(pDataset)
'
'  Dim DataObj As New MSForms.DataObject
''  DataObj.SetText strXML
'  DataObj.SetText strAllProps
'  DataObj.PutInClipboard
'  Set DataObj = Nothing
  
  booFailed = False
  Dim pPropSet As IPropertySet
  Set pPropSet = ReturnMetadataPropSetFromDataset(pDataset)

  Dim varVals As Variant
  Dim varNames As Variant
  pPropSet.GetAllProperties varNames, varVals
  
  Dim strProps As String
  Dim lngIndex As Long
  
  ' ALL PROPERTIES ========================================
  
  For lngIndex = 0 To UBound(varVals)
    If lngIndex <> UBound(varVals) Then

      strProps = strProps & CStr(lngIndex) & "] " & varNames(lngIndex) & vbCrLf
      If VarType(varVals(lngIndex)) = vbDataObject Then
        strProps = strProps & "  --> *** DATA OBJECT ***" & vbCrLf
      Else
        strProps = strProps & "  --> " & varVals(lngIndex) & vbCrLf
      End If
    End If

  Next lngIndex
  
  ReturnAllMetadataPropertiesFromDataset = strProps
  
  GoTo ClearMemory
  Exit Function
  
ErrHandler:
  ReturnAllMetadataPropertiesFromDataset = "ReturnAllMetadataPropertiesFromDataset Failed"
  booFailed = True
  
ClearMemory:
  Set pPropSet = Nothing
  varVals = Null
  varNames = Null
  
End Function

Public Function ReturnMetadataXMLStringFromDataset(pDataset As IDataset, _
    Optional booFailed As Boolean) As String

  On Error GoTo ErrHandler
  
'  SAMPLE CODE
'  Dim pFLayer As IFeatureLayer
'  Dim pMxDoc As IMxDocument
'  Set pMxDoc = ThisDocument
'  Set pFLayer = mygeneralOperations.ReturnLayerByName("Roads_for_Metadata", pMxDoc.FocusMap)
'  Dim pFClass As IFeatureClass
'
'  Set pFClass = pFLayer.FeatureClass
'  Dim pDataset As IDataset
'  Set pDataset = pFClass
'
'  Dim strAllProps As String
'  Dim strXML As String
'
'  strXML = Metadata_Functions.ReturnMetadataXMLStringFromDataset(pDataset)
'  strAllProps = Metadata_Functions.ReturnAllMetadataPropertiesFromDataset(pDataset)
'
'  Dim DataObj As New MSForms.DataObject
''  DataObj.SetText strXML
'  DataObj.SetText strAllProps
'  DataObj.PutInClipboard
'  Set DataObj = Nothing
  
  booFailed = False
  
  Dim pPropSet As IPropertySet
  Set pPropSet = ReturnMetadataPropSetFromDataset(pDataset)
  Dim pXMLPropSet As IXmlPropertySet2
  Set pXMLPropSet = pPropSet
  
  Dim strXML As String
  strXML = pXMLPropSet.GetXml("")
  ReturnMetadataXMLStringFromDataset = strXML
    
  GoTo ClearMemory
  Exit Function
  
ErrHandler:
  ReturnMetadataXMLStringFromDataset = "ReturnMetadataXMLStringFromDataset Failed"
  booFailed = True
  
ClearMemory:
  Set pPropSet = Nothing
  Set pXMLPropSet = Nothing
  
End Function

Public Function ReturnDataElementFromDataset(pDataset As IDataset, Optional booFailed As Boolean) As IDataElement

  On Error GoTo ErrHandler
  
  Dim pName As IName
  Set pName = pDataset.FullName
'  Debug.Print pName.NameString
  
  Dim pDEUtils As IDEUtilities
  Set pDEUtils = New DEUtilities
  
  Set ReturnDataElementFromDataset = pDEUtils.MakeDataElementFromNameObject(pName)
  
  GoTo ClearMemory
  Exit Function
  
ErrHandler:
  Set ReturnDataElementFromDataset = Nothing
  booFailed = True
  
ClearMemory:
  Set pName = Nothing
  Set pDEUtils = Nothing
  
End Function

Public Function SetMetadataKeyWords(pDataset As IDataset, _
  Optional pIncludeThemeKeys As esriSystem.IStringArray, _
  Optional pIncludeSearchKeys As esriSystem.IStringArray, _
  Optional pIncludeDescKeys As esriSystem.IStringArray, _
  Optional pIncludeStratKeys As esriSystem.IStringArray, _
  Optional pIncludeThemeSlashThemekeys As esriSystem.IStringArray, _
  Optional pIncludePlaceKeys As esriSystem.IStringArray, _
  Optional pIncludeTemporalKeys As esriSystem.IStringArray) As String   ', _
   ' lngReplaceOrAdd As esriXmlSetPropertyAction) As String
  
  On Error GoTo ErrHandler
  
'  Dim pKeyWords As esriSystem.IStringArray
'  Dim pIncludeThemeKeys As esriSystem.IStringArray
'  Dim pIncludeSearchKeys As esriSystem.IStringArray
'  Dim pIncludeDescKeys As esriSystem.IStringArray
'  Dim pIncludeStratKeys As esriSystem.IStringArray
'  Dim pIncludeThemeSlashThemekeys As esriSystem.IStringArray
'  Dim pIncludePlaceKeys As esriSystem.IStringArray
'  Dim pIncludeTemporalKeys As esriSystem.IStringArray
'
'  Set pIncludeThemeKeys = New esriSystem.strArray
'  Set pIncludeSearchKeys = New esriSystem.strArray
'  Set pIncludeDescKeys = New esriSystem.strArray
'  Set pIncludeStratKeys = New esriSystem.strArray
'  Set pIncludeThemeSlashThemekeys = New esriSystem.strArray
'  Set pIncludePlaceKeys = New esriSystem.strArray
'  Set pIncludeTemporalKeys = New esriSystem.strArray
'
'  pIncludeThemeKeys.Add "Keyword 1"
'  pIncludeThemeKeys.Add "Keyword 2"
'  pIncludeThemeKeys.Add "Keyword 3"
'  pIncludeSearchKeys.Add "Keyword 1"
'  pIncludeSearchKeys.Add "Keyword 2"
'  pIncludeSearchKeys.Add "Keyword 3"
'  pIncludeDescKeys.Add "Keyword 1"
'  pIncludeDescKeys.Add "Keyword 2"
'  pIncludeDescKeys.Add "Keyword 3"
'
'  pIncludePlaceKeys.Add "Flagstaff"
'  pIncludeThemeSlashThemekeys.Add "theme_slash_theme"
'  pIncludeTemporalKeys.Add "temporal"
'
'  strResponse = Metadata_Functions.SetMetadataKeyWords(pDataset, pIncludeThemeKeys, pIncludeSearchKeys, _
'        pIncludeDescKeys, pIncludeStratKeys, pIncludeThemeSlashThemekeys, pIncludePlaceKeys, pIncludeTemporalKeys)
'  Debug.Print "Saving Keywords: " & strResponse
  
  
  SetMetadataKeyWords = "Succeeded"

  Dim strThemeKeywordsXPath As String
  Dim strSearchKeywordsXPath As String
  Dim strDescKeywordsXPath As String
  Dim strStratKeywordsXPath As String
  Dim strThemeSlashThemekeyKeywordsXPath As String
  Dim strPlaceKeywordsXPath As String
  Dim strTemporalKeywordsXPath As String
  
  strThemeKeywordsXPath = "dataIdInfo/themeKeys/keyword"
  strSearchKeywordsXPath = "dataIdInfo/searchKeys/keyword"
  strDescKeywordsXPath = "dataIdInfo/descKeys/keyword"
  strStratKeywordsXPath = "dataIdInfo/StratKeys/keyword"
  strThemeSlashThemekeyKeywordsXPath = "idinfo/keywords/theme/themekey"
  strPlaceKeywordsXPath = "idinfo/keywords/place/placekey"
  strTemporalKeywordsXPath = "idinfo/keywords/temporal/tempkey"
  
  Dim lngIndex As Long
  Dim strValue As String
  
  Dim pPropSet As IPropertySet
  Set pPropSet = ReturnMetadataPropSetFromDataset(pDataset)
  
  If Not pIncludeThemeKeys Is Nothing Then
    If pIncludeThemeKeys.Count > 0 Then
      pPropSet.RemoveProperty strThemeKeywordsXPath
      For lngIndex = 0 To pIncludeThemeKeys.Count - 1
        strValue = pIncludeThemeKeys.Element(lngIndex)
        pPropSet.SetProperty strThemeKeywordsXPath & "[" & CStr(lngIndex) & "]", strValue
    '    Debug.Print "  " & CStr(lngIndex) & "] " & strValue
      Next lngIndex
    End If
  End If
  
  If Not pIncludeStratKeys Is Nothing Then
    If pIncludeStratKeys.Count > 0 Then
      pPropSet.RemoveProperty strStratKeywordsXPath
      For lngIndex = 0 To pIncludeStratKeys.Count - 1
        strValue = pIncludeStratKeys.Element(lngIndex)
        pPropSet.SetProperty strStratKeywordsXPath & "[" & CStr(lngIndex) & "]", strValue
      Next lngIndex
    End If
  End If
  
  If Not pIncludeSearchKeys Is Nothing Then
    If pIncludeSearchKeys.Count > 0 Then
      pPropSet.RemoveProperty strSearchKeywordsXPath
      For lngIndex = 0 To pIncludeThemeKeys.Count - 1
        strValue = pIncludeSearchKeys.Element(lngIndex)
        pPropSet.SetProperty strSearchKeywordsXPath & "[" & CStr(lngIndex) & "]", strValue
      Next lngIndex
    End If
  End If
  
  If Not pIncludeDescKeys Is Nothing Then
    If pIncludeDescKeys.Count > 0 Then
      pPropSet.RemoveProperty strDescKeywordsXPath
      For lngIndex = 0 To pIncludeDescKeys.Count - 1
        strValue = pIncludeDescKeys.Element(lngIndex)
        pPropSet.SetProperty strDescKeywordsXPath & "[" & CStr(lngIndex) & "]", strValue
      Next lngIndex
    End If
  End If
  
  If Not pIncludeThemeSlashThemekeys Is Nothing Then
    If pIncludeThemeSlashThemekeys.Count > 0 Then
      pPropSet.RemoveProperty strThemeSlashThemekeyKeywordsXPath
      For lngIndex = 0 To pIncludeThemeSlashThemekeys.Count - 1
        strValue = pIncludeThemeSlashThemekeys.Element(lngIndex)
        pPropSet.SetProperty strThemeSlashThemekeyKeywordsXPath & "[" & CStr(lngIndex) & "]", strValue
      Next lngIndex
    End If
  End If
  
  If Not pIncludePlaceKeys Is Nothing Then
    If pIncludePlaceKeys.Count > 0 Then
      pPropSet.RemoveProperty strPlaceKeywordsXPath
      For lngIndex = 0 To pIncludePlaceKeys.Count - 1
        strValue = pIncludePlaceKeys.Element(lngIndex)
        pPropSet.SetProperty strPlaceKeywordsXPath & "[" & CStr(lngIndex) & "]", strValue
      Next lngIndex
    End If
  End If
  
  If Not pIncludeTemporalKeys Is Nothing Then
    If pIncludeTemporalKeys.Count > 0 Then
      pPropSet.RemoveProperty strTemporalKeywordsXPath
      For lngIndex = 0 To pIncludeTemporalKeys.Count - 1
        strValue = pIncludeTemporalKeys.Element(lngIndex)
        pPropSet.SetProperty strTemporalKeywordsXPath & "[" & CStr(lngIndex) & "]", strValue
      Next lngIndex
    End If
  End If
  
  
  Metadata_Functions.SaveMetadata pDataset, pPropSet
  
  GoTo ClearMemory
  Exit Function
  
ErrHandler:
  SetMetadataKeyWords = "Failed"

ClearMemory:
  Set pPropSet = Nothing
  
End Function


Public Function ReturnExistingMetadataKeyWords(pDataset As IDataset, _
  pKeyWordsToInsertToArray As esriSystem.IStringArray, booSucceeded As Boolean, _
  Optional pIncludeThemeKeys As esriSystem.IStringArray, _
  Optional pIncludeSearchKeys As esriSystem.IStringArray, _
  Optional pIncludeDescKeys As esriSystem.IStringArray, _
  Optional pIncludeStratKeys As esriSystem.IStringArray, _
  Optional pIncludeThemeSlashThemekeys As esriSystem.IStringArray, _
  Optional pIncludePlaceKeys As esriSystem.IStringArray, _
  Optional pIncludeTemporalKeys As esriSystem.IStringArray) As esriSystem.IStringArray
  
  On Error GoTo ErrHandler
  
  ' SAMPLE CODE
'  Dim pKeyWords As esriSystem.IStringArray
'  Dim pIncludeThemeKeys As esriSystem.IStringArray
'  Dim pIncludeSearchKeys As esriSystem.IStringArray
'  Dim pIncludeDescKeys As esriSystem.IStringArray
'  Dim pIncludeStratKeys As esriSystem.IStringArray
'  Dim pIncludeThemeSlashThemekeys As esriSystem.IStringArray
'  Dim pIncludePlaceKeys As esriSystem.IStringArray
'  Dim pIncludeTemporalKeys As esriSystem.IStringArray
'
'  Set pIncludeThemeKeys = New esriSystem.strArray
'  Set pIncludeSearchKeys = New esriSystem.strArray
'  Set pIncludeDescKeys = New esriSystem.strArray
'  Set pIncludeStratKeys = New esriSystem.strArray
'  Set pIncludeThemeSlashThemekeys = New esriSystem.strArray
'  Set pIncludePlaceKeys = New esriSystem.strArray
'  Set pIncludeTemporalKeys = New esriSystem.strArray
'
'  pIncludeThemeSlashThemekeys.Add "theme_slash_theme"
'  pIncludeTemporalKeys.Add "temporal"
'
'  Dim pCombinedKeyWords As esriSystem.IStringArray
'  Dim pFLayer As IFeatureLayer
'  Dim pMxDoc As IMxDocument
'  Set pMxDoc = ThisDocument
'  Set pFLayer = mygeneralOperations.ReturnLayerByName("mroads", pMxDoc.FocusMap)
'  Dim booSucceeded As Boolean
'  Dim lngIndex As Long
'  Set pCombinedKeyWords = Metadata_Functions.ReturnExistingMetadataKeyWords(pFLayer.FeatureClass, _
'      pKeyWords, booSucceeded, pIncludeThemeKeys, pIncludesearchKeys, pIncludeDescKeys, pIncludeStratKeys, _
'      pIncludeThemeSlashThemekeys, _
'      pIncludePlaceKeys, pIncludeTemporalKeys)
'  Debug.Print "Extracting keywords: " & UCase(CStr(booSucceeded))
'  If booSucceeded Then
'    Debug.Print "Combined..."
'    If pCombinedKeyWords.Count > 0 Then
'      For lngIndex = 0 To pCombinedKeyWords.Count - 1
'        Debug.Print "  --> " & CStr(lngIndex + 1) & "] " & pCombinedKeyWords.Element(lngIndex)
'      Next lngIndex
'    Else
'      Debug.Print "Nothing in 'pCombinedKeyWords'..."
'    End If
'    Debug.Print "Theme..."
'    If pIncludeThemeKeys.Count > 0 Then
'      For lngIndex = 0 To pIncludeThemeKeys.Count - 1
'        Debug.Print "  --> " & CStr(lngIndex + 1) & "] " & pIncludeThemeKeys.Element(lngIndex)
'      Next lngIndex
'    Else
'      Debug.Print "-- Nothing in 'pIncludeThemeKeys'..."
'    End If
'    Debug.Print "Search..."
'    If pIncludeSearchKeys.Count > 0 Then
'      For lngIndex = 0 To pIncludeSearchKeys.Count - 1
'        Debug.Print "  --> " & CStr(lngIndex + 1) & "] " & pIncludeSearchKeys.Element(lngIndex)
'      Next lngIndex
'    Else
'      Debug.Print "-- Nothing in 'pIncludeSearchKeys'..."
'    End If
'    Debug.Print "Desc..."
'    If pIncludeDescKeys.Count > 0 Then
'      For lngIndex = 0 To pIncludeDescKeys.Count - 1
'        Debug.Print "  --> " & CStr(lngIndex + 1) & "] " & pIncludeDescKeys.Element(lngIndex)
'      Next lngIndex
'    Else
'      Debug.Print "-- Nothing in 'pIncludeDescKeys'..."
'    End If
'    Debug.Print "Strat..."
'    If pIncludeStratKeys.Count > 0 Then
'      For lngIndex = 0 To pIncludeStratKeys.Count - 1
'        Debug.Print "  --> " & CStr(lngIndex + 1) & "] " & pIncludeStratKeys.Element(lngIndex)
'      Next lngIndex
'    Else
'      Debug.Print "-- Nothing in 'pIncludeStratKeys'..."
'    End If
'    Debug.Print "ThemeSlashTheme..."
'    If pIncludeThemeSlashThemekeys.Count > 0 Then
'      For lngIndex = 0 To pIncludeThemeSlashThemekeys.Count - 1
'        Debug.Print "  --> " & CStr(lngIndex + 1) & "] " & pIncludeThemeSlashThemekeys.Element(lngIndex)
'      Next lngIndex
'    Else
'      Debug.Print "-- Nothing in 'pIncludeThemeSlashThemekeys'..."
'    End If
'    Debug.Print "Place..."
'    If pIncludePlaceKeys.Count > 0 Then
'      For lngIndex = 0 To pIncludePlaceKeys.Count - 1
'        Debug.Print "  --> " & CStr(lngIndex + 1) & "] " & pIncludePlaceKeys.Element(lngIndex)
'      Next lngIndex
'    Else
'      Debug.Print "-- Nothing in 'pIncludePlaceKeys'..."
'    End If
'    Debug.Print "Temporal..."
'    If pIncludeTemporalKeys.Count > 0 Then
'      For lngIndex = 0 To pIncludeTemporalKeys.Count - 1
'        Debug.Print "  --> " & CStr(lngIndex + 1) & "] " & pIncludeTemporalKeys.Element(lngIndex)
'      Next lngIndex
'    Else
'      Debug.Print "-- Nothing in 'pIncludeTemporalKeys'..."
'    End If
'  End If
  
  
  ' THIS WILL ADD KEY WORDS TO VARIOUS STRING ARRAYS IF ANY EXIST
  
  booSucceeded = True
  
  Dim strThemeKeywordsXPath As String
  Dim strSearchKeywordsXPath As String
  Dim strDescKeywordsXPath As String
  Dim strStratKeywordsXPath As String
  Dim strThemeSlashThemekeyKeywordsXPath As String
  Dim strPlaceKeywordsXPath As String
  Dim strTemporalKeywordsXPath As String
  Dim varAtts As Variant
  Dim strValue As String
  
  Dim lngIndex As Long
  Set ReturnExistingMetadataKeyWords = New esriSystem.strArray
  Dim pKeyWordColl As New Collection
  Dim pThemeKeyWordColl As New Collection
  Dim pSearchKeyWordColl As New Collection
  Dim pStratKeyWordColl As New Collection
  Dim pDescKeyWordColl As New Collection
  Dim pThemeSlashThemeKeyWordColl As New Collection
  Dim pPlaceKeyWordColl As New Collection
  Dim pTemporalKeyWordColl As New Collection
  
  If Not pKeyWordsToInsertToArray Is Nothing Then
    If pKeyWordsToInsertToArray.Count > 0 Then
      For lngIndex = 0 To pKeyWordsToInsertToArray.Count - 1
        strValue = pKeyWordsToInsertToArray.Element(lngIndex)
        If Not MyGeneralOperations.CheckCollectionForKey(pKeyWordColl, strValue) Then
          ReturnExistingMetadataKeyWords.Add strValue
          pKeyWordColl.Add True, strValue
        End If
      Next lngIndex
    End If
  End If
    
  Dim pPropSet As IPropertySet
  Set pPropSet = ReturnMetadataPropSetFromDataset(pDataset)
    
  strThemeKeywordsXPath = "dataIdInfo/themeKeys/keyword"
  strSearchKeywordsXPath = "dataIdInfo/searchKeys/keyword"
  strStratKeywordsXPath = "dataIdInfo/StratKeys/keyword"
  strDescKeywordsXPath = "dataIdInfo/descKeys/keyword"
  strThemeSlashThemekeyKeywordsXPath = "idinfo/keywords/theme/themekey"
  strPlaceKeywordsXPath = "idinfo/keywords/place/placekey"
  strTemporalKeywordsXPath = "idinfo/keywords/temporal/tempkey"
      
   
'    Debug.Print "Keywords for '" & strFClassName & "'"
  If Not pIncludeStratKeys Is Nothing Then
    varAtts = pPropSet.GetProperty(strStratKeywordsXPath)
    If Not IsEmpty(varAtts) Then
      For lngIndex = 0 To UBound(varAtts)
        strValue = CStr(varAtts(lngIndex))
  '        Debug.Print "  " & CStr(lngIndex) & "] " & strValue
        
        If Not MyGeneralOperations.CheckCollectionForKey(pKeyWordColl, strValue) Then
          ReturnExistingMetadataKeyWords.Add strValue
          pKeyWordColl.Add True, strValue
        End If
        If Not MyGeneralOperations.CheckCollectionForKey(pStratKeyWordColl, strValue) Then
          pIncludeStratKeys.Add strValue
          pStratKeyWordColl.Add True, strValue
        End If
      Next lngIndex
    End If
  End If
  
  If Not pIncludeThemeKeys Is Nothing Then
    varAtts = pPropSet.GetProperty(strThemeKeywordsXPath)
    If Not IsEmpty(varAtts) Then
      For lngIndex = 0 To UBound(varAtts)
        strValue = CStr(varAtts(lngIndex))
  '        Debug.Print "  " & CStr(lngIndex) & "] " & strValue
        
        If Not MyGeneralOperations.CheckCollectionForKey(pKeyWordColl, strValue) Then
          ReturnExistingMetadataKeyWords.Add strValue
          pKeyWordColl.Add True, strValue
        End If
        If Not MyGeneralOperations.CheckCollectionForKey(pThemeKeyWordColl, strValue) Then
          pIncludeThemeKeys.Add strValue
          pThemeKeyWordColl.Add True, strValue
        End If
      Next lngIndex
    End If
  End If
  
  If Not pIncludeSearchKeys Is Nothing Then
    varAtts = pPropSet.GetProperty(strSearchKeywordsXPath)
    If Not IsEmpty(varAtts) Then
      For lngIndex = 0 To UBound(varAtts)
        strValue = CStr(varAtts(lngIndex))
  '        Debug.Print "  " & CStr(lngIndex) & "] " & strValue
        
        If Not MyGeneralOperations.CheckCollectionForKey(pKeyWordColl, strValue) Then
          ReturnExistingMetadataKeyWords.Add strValue
          pKeyWordColl.Add True, strValue
        End If
        If Not MyGeneralOperations.CheckCollectionForKey(pSearchKeyWordColl, strValue) Then
          pIncludeSearchKeys.Add strValue
          pSearchKeyWordColl.Add True, strValue
        End If
      Next lngIndex
    End If
  End If
  
  If Not pIncludeDescKeys Is Nothing Then
    varAtts = pPropSet.GetProperty(strDescKeywordsXPath)
    If Not IsEmpty(varAtts) Then
      For lngIndex = 0 To UBound(varAtts)
        strValue = CStr(varAtts(lngIndex))
  '        Debug.Print "  " & CStr(lngIndex) & "] " & strValue
        
        If Not MyGeneralOperations.CheckCollectionForKey(pKeyWordColl, strValue) Then
          ReturnExistingMetadataKeyWords.Add strValue
          pKeyWordColl.Add True, strValue
        End If
        If Not MyGeneralOperations.CheckCollectionForKey(pSearchKeyWordColl, strValue) Then
          pIncludeDescKeys.Add strValue
          pDescKeyWordColl.Add True, strValue
        End If
      Next lngIndex
    End If
  End If
  
  If Not pIncludeThemeSlashThemekeys Is Nothing Then
    varAtts = pPropSet.GetProperty(strThemeSlashThemekeyKeywordsXPath)
    If Not IsEmpty(varAtts) Then
      For lngIndex = 0 To UBound(varAtts)
        strValue = CStr(varAtts(lngIndex))
  '        Debug.Print "  " & CStr(lngIndex) & "] " & strValue
        
        If Not MyGeneralOperations.CheckCollectionForKey(pKeyWordColl, strValue) Then
          ReturnExistingMetadataKeyWords.Add strValue
          pKeyWordColl.Add True, strValue
        End If
        If Not MyGeneralOperations.CheckCollectionForKey(pThemeSlashThemeKeyWordColl, strValue) Then
          pIncludeThemeSlashThemekeys.Add strValue
          pThemeSlashThemeKeyWordColl.Add True, strValue
        End If
      Next lngIndex
    End If
  End If
  
  If Not pIncludePlaceKeys Is Nothing Then
    varAtts = pPropSet.GetProperty(strPlaceKeywordsXPath)
    If Not IsEmpty(varAtts) Then
      For lngIndex = 0 To UBound(varAtts)
        strValue = CStr(varAtts(lngIndex))
  '        Debug.Print "  " & CStr(lngIndex) & "] " & strValue
        
        If Not MyGeneralOperations.CheckCollectionForKey(pKeyWordColl, strValue) Then
          ReturnExistingMetadataKeyWords.Add strValue
          pKeyWordColl.Add True, strValue
        End If
        If Not MyGeneralOperations.CheckCollectionForKey(pPlaceKeyWordColl, strValue) Then
          pIncludePlaceKeys.Add strValue
          pPlaceKeyWordColl.Add True, strValue
        End If
      Next lngIndex
    End If
  End If
  
  If Not pIncludeTemporalKeys Is Nothing Then
    varAtts = pPropSet.GetProperty(strTemporalKeywordsXPath)
    If Not IsEmpty(varAtts) Then
      For lngIndex = 0 To UBound(varAtts)
        strValue = CStr(varAtts(lngIndex))
  '        Debug.Print "  " & CStr(lngIndex) & "] " & strValue
        
        If Not MyGeneralOperations.CheckCollectionForKey(pKeyWordColl, strValue) Then
          ReturnExistingMetadataKeyWords.Add strValue
          pKeyWordColl.Add True, strValue
        End If
        If Not MyGeneralOperations.CheckCollectionForKey(pTemporalKeyWordColl, strValue) Then
          pIncludeTemporalKeys.Add strValue
          pTemporalKeyWordColl.Add True, strValue
        End If
      Next lngIndex
    End If
  End If
    
  GoTo ClearMemory
  Exit Function
  
ErrHandler:
  booSucceeded = False
  
ClearMemory:
  varAtts = Null
  Set pKeyWordColl = Nothing
  Set pThemeKeyWordColl = Nothing
  Set pSearchKeyWordColl = Nothing
  Set pStratKeyWordColl = Nothing
  Set pDescKeyWordColl = Nothing
  Set pThemeSlashThemeKeyWordColl = Nothing
  Set pPlaceKeyWordColl = Nothing
  Set pTemporalKeyWordColl = Nothing
  Set pPropSet = Nothing
  
End Function

Public Function SetMetadataPurpose(pDataset As IDataset, strPurpose As String) As String ', _
   ' lngReplaceOrAdd As esriXmlSetPropertyAction) As String
  
  On Error GoTo ErrHandler
    
'  Dim strPurpose As String
'  strPurpose = "Point dataset of points along route with widest bottleneck, with corridor width values at each point."
'  strResponse = Metadata_Functions.SetMetadataPurpose(pDataset, strPurpose)
'  Debug.Print "Saving Purpose: " & strResponse
  
  SetMetadataPurpose = "Succeeded"

  Dim strPurposeXPath As String
  strPurposeXPath = "dataIdInfo/idPurp"
  
  Dim pPropSet As IPropertySet
  Set pPropSet = ReturnMetadataPropSetFromDataset(pDataset)
  
  pPropSet.SetProperty strPurposeXPath, strPurpose
  
  Metadata_Functions.SaveMetadata pDataset, pPropSet
  
  GoTo ClearMemory
  Exit Function
  
ErrHandler:
  SetMetadataPurpose = "Failed"

ClearMemory:
  Set pPropSet = Nothing
End Function

Public Function SetMetadataAbstract(pDataset As IDataset, strAbstract As String) As String ', _
   ' lngReplaceOrAdd As esriXmlSetPropertyAction) As String
  
  On Error GoTo ErrHandler
    
'  Dim strResponse As String
'  Dim strAbstract As String
'  strAbstract = "This dataset represents points along the bottleneck route.  This bottleneck route " & _
'    "describes the path between the two habitat blocks 'aaa' and 'bbb', within the corridor polygon 'ccc'., " & _
'    "which follows the route with the widest narrow point."
'  strResponse = Metadata_Functions.SetMetadataAbstract(pDataset, strAbstract)
'  Debug.Print "Saving Abstract: " & strResponse
  
  SetMetadataAbstract = "Succeeded"

  Dim strAbstractXPath As String
  strAbstractXPath = "idinfo/descript/abstract"
  Dim strDescriptionXPath As String
  strDescriptionXPath = "dataIdInfo/idAbs"
  
  Dim pPropSet As IPropertySet
  Set pPropSet = ReturnMetadataPropSetFromDataset(pDataset)
  
  pPropSet.SetProperty strAbstractXPath, strAbstract
  pPropSet.SetProperty strDescriptionXPath, strAbstract
  
  Metadata_Functions.SaveMetadata pDataset, pPropSet
  
  GoTo ClearMemory
  Exit Function
  
ErrHandler:
  SetMetadataAbstract = "Failed"

ClearMemory:
  Set pPropSet = Nothing
  
End Function
Public Function ReturnGxDatasetFromDataset(pDataset As IDataset, _
    Optional booFailed As Boolean) As IGxDataset
  
  On Error GoTo ErrHandler
  
  ' GENERALLY CALLED BY OTHER FUNCTIONS
  
  booFailed = False
  
  Dim pName As IName
  Set pName = pDataset.FullName
'  Debug.Print pName.NameString
  
  Dim pGxDataset As IGxDataset
  Set pGxDataset = New GxDataset
  Set pGxDataset.DatasetName = pName
  
  Set ReturnGxDatasetFromDataset = pGxDataset
      
  GoTo ClearMemory
  Exit Function
  
ErrHandler:
  Set ReturnGxDatasetFromDataset = Nothing
  booFailed = True
  
ClearMemory:
  Set pName = Nothing
  Set pGxDataset = Nothing
  
End Function

Public Function ReturnMetadataPropSetFromDataset(pDataset As IDataset, _
    Optional booFailed As Boolean) As IPropertySet
    
  On Error GoTo ErrHandler
  
  ' GENERALLY CALLED BY OTHER FUNCTIONS
    
'  Dim pPropSet As IPropertySet
'  Set pPropSet = Metadata_Functions.ReturnMetadataPropSetFromDataset(pFClass)
'  Dim varNames As Variant
'  Dim varVals As Variant
'  pPropSet.GetAllProperties varNames, varVals
'
'  Dim lngIndex As Long
'  For lngIndex = 0 To UBound(varNames)
'    Debug.Print "Name " & CStr(lngIndex) & ": " & CStr(varNames(lngIndex)) & "  --> Value = " & CStr(varVals(lngIndex))
'  Next lngIndex
'
  Dim pGxDataset As IGxDataset
  Set pGxDataset = ReturnGxDatasetFromDataset(pDataset)
'  Debug.Print "Browse name = " & pGxDataset.Dataset.BrowseName
    
  Dim pMetaData As IMetadata
  Set pMetaData = pGxDataset
  
'  Debug.Print pMetaData Is Nothing
  
  Dim pMetadataEdit As IMetadataEdit
  Set pMetadataEdit = pGxDataset
'  Debug.Print "Can Edit = " & pMetadataEdit.CanEditMetadata
  
  Dim pPropSet As IPropertySet
  Set pPropSet = pMetaData.Metadata
  
'  Debug.Print "Property Count = " & CStr(pPropSet.Count)
  Dim pXMLPropSet As IXmlPropertySet2
  Set pXMLPropSet = pMetaData.Metadata
  Set pPropSet = pXMLPropSet
'  Debug.Print "Property Set is New = " & CStr(pXMLPropSet.IsNew)
  
  Dim pXMLPropSet2 As IXmlPropertySet2
  If pMetadataEdit.CanEditMetadata Then
    If pXMLPropSet.IsNew Then
      Set pXMLPropSet2 = pXMLPropSet
      pXMLPropSet2.InitExisting
      pMetaData.Metadata = pPropSet
    End If
  End If
  
  Set ReturnMetadataPropSetFromDataset = pPropSet
'  Debug.Print "Re-Check Property Set is New = " & CStr(pXMLPropSet.IsNew)
        
  GoTo ClearMemory
  Exit Function
  
ErrHandler:
  Set ReturnMetadataPropSetFromDataset = Nothing
  booFailed = True
  
ClearMemory:
  Set pGxDataset = Nothing
  Set pMetaData = Nothing
  Set pMetadataEdit = Nothing
  Set pPropSet = Nothing
  Set pXMLPropSet = Nothing
  Set pXMLPropSet2 = Nothing
  
End Function

Public Function SynchronizeMetadataPropSet(pDataset As IDataset) As String
    
  On Error GoTo ErrHandler
  
'  SYNCHRONIZE Metadata
'  strResponse = Metadata_Functions.SynchronizeMetadataPropSet(pDataset)
'  Debug.Print "ReSynchronization: " & strResponse
'
  ' SHOULD UPDATE ANY NEW CHANGES TO THE METADATA, LIKE ADDING NEW FIELDS
  
  SynchronizeMetadataPropSet = "Succeeded"
  
  Dim pGxDataset As IGxDataset
  Set pGxDataset = ReturnGxDatasetFromDataset(pDataset)
'  Debug.Print "Browse name = " & pGxDataset.Dataset.BrowseName
    
  Dim pMetaData As IMetadata
  Set pMetaData = pGxDataset
  
'  Debug.Print pMetaData Is Nothing
  
  Dim pMetadataEdit As IMetadataEdit
  Set pMetadataEdit = pGxDataset
'  Debug.Print "Can Edit = " & pMetadataEdit.CanEditMetadata
  
  If pMetadataEdit.CanEditMetadata Then
    pMetaData.SYNCHRONIZE esriMSAAccessed, 0
  Else
    SynchronizeMetadataPropSet = "Unable to synchronize; Metadata not editable..."
  End If
          
  GoTo ClearMemory
  Exit Function
  
ErrHandler:
  SynchronizeMetadataPropSet = "Failed"
  
ClearMemory:
  Set pGxDataset = Nothing
  Set pMetaData = Nothing
  Set pMetadataEdit = Nothing
  
End Function

Public Sub SaveMetadata(pDataset As IDataset, pPropSet As IPropertySet, _
    Optional booFailed As Boolean)
  On Error GoTo ErrHandler
  
  ' GENERALLY CALLED FROM OTHER FUNCTIONS
  booFailed = False
  Dim pGxDataset As IGxDataset
  Set pGxDataset = ReturnGxDatasetFromDataset(pDataset)
'  Debug.Print "Browse name = " & pGxDataset.Dataset.BrowseName
    
  Dim pMetaData As IMetadata
  Set pMetaData = pGxDataset
  
  Dim pMetadataEdit As IMetadataEdit
  Set pMetadataEdit = pMetaData
  
  If pMetadataEdit.CanEditMetadata Then
    pMetaData.Metadata = pPropSet
  End If
  
  GoTo ClearMemory
  Exit Sub
  
ErrHandler:
  booFailed = True
  
ClearMemory:
  Set pGxDataset = Nothing
  Set pMetaData = Nothing
  Set pMetadataEdit = Nothing
  
End Sub


Private Function ReturnAddressType(enumJenAddressType As JenMetadataAddressTypeValues) As String
  
  On Error GoTo ErrHandler
  ' GENERALLY CALLED BY OTHER FUNCTIONS
  
  Select Case enumJenAddressType
    Case JenMetadata_Postal
      ReturnAddressType = "postal"
    Case JenMetadata_Physical
      ReturnAddressType = "physical"
    Case JenMetadata_both
      ReturnAddressType = "both"
    Case Else
      ReturnAddressType = "skip"
  End Select
  
  Exit Function
  
ErrHandler:
  ReturnAddressType = ""
  
End Function

Private Function ReturnMaintenanceCode(enumJenMaintenanceCode As JenMetadataMaintenanceCodes) As String
  
  On Error GoTo ErrHandler
  ' GENERALLY CALLED BY OTHER FUNCTIONS
      
  Select Case enumJenMaintenanceCode
    Case JenMetadata_Maint_Continual
      ReturnMaintenanceCode = "001"
    Case JenMetadata_Maint_Daily
      ReturnMaintenanceCode = "002"
    Case JenMetadata_Maint_Weekly
      ReturnMaintenanceCode = "003"
    Case JenMetadata_Maint_Fortnightly
      ReturnMaintenanceCode = "004"
    Case JenMetadata_Maint_Monthly
      ReturnMaintenanceCode = "005"
    Case JenMetadata_Maint_Quarterly
      ReturnMaintenanceCode = "006"
    Case JenMetadata_Maint_BiAnnually
      ReturnMaintenanceCode = "007"
    Case JenMetadata_Maint_Annually
      ReturnMaintenanceCode = "008"
    Case JenMetadata_Maint_AsNeeded
      ReturnMaintenanceCode = "009"
    Case JenMetadata_Maint_Irregular
      ReturnMaintenanceCode = "010"
    Case JenMetadata_Maint_NotPlanned
      ReturnMaintenanceCode = "011"
    Case JenMetadata_Maint_Unknown
      ReturnMaintenanceCode = "012"
    Case JenMetadata_Maint_SemiMonthly
      ReturnMaintenanceCode = "013"
  End Select
    
  Exit Function
  
ErrHandler:
  ReturnMaintenanceCode = ""
  
End Function


Private Function ReturnStatusString(enumJenStatus As JenMetadataStatusValues) As String

  On Error GoTo ErrHandler
  ' GENERALLY CALLED BY OTHER FUNCTIONS
  Select Case enumJenStatus
    Case JenMetadata_Completed
      ReturnStatusString = "001"
    Case JenMetadata_HistoricalArchive
      ReturnStatusString = "002"
    Case JenMetadata_Obsolete
      ReturnStatusString = "003"
    Case JenMetadata_Ongoing
      ReturnStatusString = "004"
    Case JenMetadata_Planned
      ReturnStatusString = "005"
    Case JenMetadata_Required
      ReturnStatusString = "006"
    Case JenMetadata_UnderDevelopment
      ReturnStatusString = "007"
    Case JenMetadata_Proposed
      ReturnStatusString = "008"
  End Select
  
  Exit Function
  
ErrHandler:
  ReturnStatusString = ""
End Function

Private Function ReturnRoleCDString(enumJenRole As JenMetadataRoleCDValues) As String
  On Error GoTo ErrHandler
  ' GENERALLY CALLED BY OTHER FUNCTIONS
  
  Select Case enumJenRole
    Case JenMetadata_ResourceProvider
      ReturnRoleCDString = "001"
    Case JenMetadata_Custodian
      ReturnRoleCDString = "002"
    Case JenMetadata_Owner
      ReturnRoleCDString = "003"
    Case JenMetadata_User
      ReturnRoleCDString = "004"
    Case JenMetadata_Distributor
      ReturnRoleCDString = "005"
    Case JenMetadata_Originator
      ReturnRoleCDString = "006"
    Case JenMetadata_PointOfContact
      ReturnRoleCDString = "007"
    Case JenMetadata_PrincipalInvestigator
      ReturnRoleCDString = "008"
    Case JenMetadata_Processor
      ReturnRoleCDString = "009"
    Case JenMetadata_Publisher
      ReturnRoleCDString = "010"
    Case JenMetadata_Author
      ReturnRoleCDString = "011"
    Case JenMetadata_Collaborator
      ReturnRoleCDString = "012"
    Case JenMetadata_Editor
      ReturnRoleCDString = "013"
    Case JenMetadata_Mediator
      ReturnRoleCDString = "014"
    Case JenMetadata_RightsHolder
      ReturnRoleCDString = "015"
    Case Else
      ReturnRoleCDString = ""
  End Select
  
  Exit Function
  
ErrHandler:
  ReturnRoleCDString = ""
End Function


Public Sub SampleCode()
'  BELOW COPIED FROM D:\arcGIS_stuff\consultation\NAU_GCNP\Data\Transportation\Transportation_2_nov_24.mxd
'
'Dim strBaseString As String
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "' BELOW COPIED FROM D:\arcGIS_stuff\consultation\NAU_GCNP\Data\Transportation\Transportation_2_nov_24.mxd" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "Public Sub ReplaceOrAddGeneralMetadataElements(pPropSet As IPropertySet, _" & vbNewLine
'  strBaseString = strBaseString & "    pFClassArray As esriSystem.IArray)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  Dim pRoads_KNF_UTM12_NAD83_FClass As IFeatureClass" & vbNewLine
'  strBaseString = strBaseString & "  Dim pTIGER_2012_Trans_Near_Park_UTM12_NAD83_FClass As IFeatureClass" & vbNewLine
'  strBaseString = strBaseString & "  Dim pBLM_AZStrip_UTM12_NAD83_FClass As IFeatureClass" & vbNewLine
'  strBaseString = strBaseString & "  Dim pRoads_UTM12_NAD83_FClass As IFeatureClass" & vbNewLine
'  strBaseString = strBaseString & "  Dim pTrails_UTM12_NAD83_FClass As IFeatureClass" & vbNewLine
'  strBaseString = strBaseString & "  Dim pRoutes_UTM12_NAD83_FClass As IFeatureClass" & vbNewLine
'  strBaseString = strBaseString & "  Dim pUtah_Roads_Near_AOI_UTM12_NAD83_FClass As IFeatureClass" & vbNewLine
'  strBaseString = strBaseString & "  Dim pTrails_KNF_UTM12_NAD83_FClass As IFeatureClass" & vbNewLine
'  strBaseString = strBaseString & "  Dim pTransportationLine_UTM12_NAD83_FClass As IFeatureClass" & vbNewLine
'  strBaseString = strBaseString & "   " & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  Dim pRoads_KNF_UTM12_NAD83_PropSet As IPropertySet" & vbNewLine
'  strBaseString = strBaseString & "  Dim pTIGER_2012_Trans_Near_Park_UTM12_NAD83_PropSet As IPropertySet" & vbNewLine
'  strBaseString = strBaseString & "  Dim pBLM_AZStrip_UTM12_NAD83_PropSet As IPropertySet" & vbNewLine
'  strBaseString = strBaseString & "  Dim pRoads_UTM12_NAD83_PropSet As IPropertySet" & vbNewLine
'  strBaseString = strBaseString & "  Dim pTrails_UTM12_NAD83_PropSet As IPropertySet" & vbNewLine
'  strBaseString = strBaseString & "  Dim pRoutes_UTM12_NAD83_PropSet As IPropertySet" & vbNewLine
'  strBaseString = strBaseString & "  Dim pUtah_Roads_Near_AOI_UTM12_NAD83_PropSet As IPropertySet" & vbNewLine
'  strBaseString = strBaseString & "  Dim pTrails_KNF_UTM12_NAD83_PropSet As IPropertySet" & vbNewLine
'  strBaseString = strBaseString & "  Dim pTransportationLine_UTM12_NAD83_PropSet As IPropertySet" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  Dim pXMLPropSet As IXmlPropertySet" & vbNewLine
'  strBaseString = strBaseString & "  Set pXMLPropSet = pPropSet" & vbNewLine
'  strBaseString = strBaseString & "  Dim pSubArray As esriSystem.IArray" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  Set pSubArray = pFClassArray.Element(0)" & vbNewLine
'  strBaseString = strBaseString & "  Set pRoads_KNF_UTM12_NAD83_FClass = pSubArray.Element(0)" & vbNewLine
'  strBaseString = strBaseString & "  Set pRoads_KNF_UTM12_NAD83_PropSet = pSubArray.Element(1)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  Set pSubArray = pFClassArray.Element(1)" & vbNewLine
'  strBaseString = strBaseString & "  Set pTIGER_2012_Trans_Near_Park_UTM12_NAD83_FClass = pSubArray.Element(0)" & vbNewLine
'  strBaseString = strBaseString & "  Set pTIGER_2012_Trans_Near_Park_UTM12_NAD83_PropSet = pSubArray.Element(1)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "   " & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  Set pSubArray = pFClassArray.Element(2)" & vbNewLine
'  strBaseString = strBaseString & "  Set pBLM_AZStrip_UTM12_NAD83_FClass = pSubArray.Element(0)" & vbNewLine
'  strBaseString = strBaseString & "  Set pBLM_AZStrip_UTM12_NAD83_PropSet = pSubArray.Element(1)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  Set pSubArray = pFClassArray.Element(3)" & vbNewLine
'  strBaseString = strBaseString & "  Set pRoads_UTM12_NAD83_FClass = pSubArray.Element(0)" & vbNewLine
'  strBaseString = strBaseString & "  Set pRoads_UTM12_NAD83_PropSet = pSubArray.Element(1)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  Set pSubArray = pFClassArray.Element(4)" & vbNewLine
'  strBaseString = strBaseString & "  Set pTrails_UTM12_NAD83_FClass = pSubArray.Element(0)" & vbNewLine
'  strBaseString = strBaseString & "  Set pTrails_UTM12_NAD83_PropSet = pSubArray.Element(1)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  Set pSubArray = pFClassArray.Element(5)" & vbNewLine
'  strBaseString = strBaseString & "  Set pRoutes_UTM12_NAD83_FClass = pSubArray.Element(0)" & vbNewLine
'  strBaseString = strBaseString & "  Set pRoutes_UTM12_NAD83_PropSet = pSubArray.Element(1)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  Set pSubArray = pFClassArray.Element(6)" & vbNewLine
'  strBaseString = strBaseString & "  Set pUtah_Roads_Near_AOI_UTM12_NAD83_FClass = pSubArray.Element(0)" & vbNewLine
'  strBaseString = strBaseString & "  Set pUtah_Roads_Near_AOI_UTM12_NAD83_PropSet = pSubArray.Element(1)" & vbNewLine
'  strBaseString = strBaseString & "   " & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  Set pSubArray = pFClassArray.Element(7)" & vbNewLine
'  strBaseString = strBaseString & "  Set pTrails_KNF_UTM12_NAD83_FClass = pSubArray.Element(0)" & vbNewLine
'  strBaseString = strBaseString & "  Set pTrails_KNF_UTM12_NAD83_PropSet = pSubArray.Element(1)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  Set pSubArray = pFClassArray.Element(8)" & vbNewLine
'  strBaseString = strBaseString & "  Set pTransportationLine_UTM12_NAD83_FClass = pSubArray.Element(0)" & vbNewLine
'  strBaseString = strBaseString & "  Set pTransportationLine_UTM12_NAD83_PropSet = pSubArray.Element(1)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  ' PUT TRANSPORTATION LINE FIRST FOR THIS FUNCTION" & vbNewLine
'  strBaseString = strBaseString & "  Dim lngIndex As Long" & vbNewLine
'  strBaseString = strBaseString & "  Dim pNewFClassArray As esriSystem.IArray" & vbNewLine
'  strBaseString = strBaseString & "  Set pNewFClassArray = New esriSystem.Array" & vbNewLine
'  strBaseString = strBaseString & "  pNewFClassArray.Add pFClassArray.Element(8)" & vbNewLine
'  strBaseString = strBaseString & "  For lngIndex = 0 To 7" & vbNewLine
'  strBaseString = strBaseString & "    pNewFClassArray.Add pFClassArray.Element(lngIndex)" & vbNewLine
'  strBaseString = strBaseString & "  Next lngIndex" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  Dim strAbstract As String" & vbNewLine
'  strBaseString = strBaseString & "  Dim strSummaryString As String" & vbNewLine
'  strBaseString = strBaseString & "  Dim strDescriptionString  As String" & vbNewLine
'  strBaseString = strBaseString & "  Dim strCreditString As String" & vbNewLine
'  strBaseString = strBaseString & "  Dim strUseLimitationsString As String" & vbNewLine
'  strBaseString = strBaseString & "  Dim strUseLimitationsString2 As String" & vbNewLine
'  strBaseString = strBaseString & "  Dim strThemeKeywords As String" & vbNewLine
'  strBaseString = strBaseString & "  Dim strSearchKeywords As String" & vbNewLine
'  strBaseString = strBaseString & "  Dim strSupplementalInformation As String" & vbNewLine
'  strBaseString = strBaseString & "  Dim strResponsiblePartyName As String" & vbNewLine
'  strBaseString = strBaseString & "  Dim strOtherConstraints As String" & vbNewLine
'  strBaseString = strBaseString & "  Dim strValue As String" & vbNewLine
'  strBaseString = strBaseString & "  Dim varAtts As Variant" & vbNewLine
'  strBaseString = strBaseString & "  Dim lngIndex2 As Long" & vbNewLine
'  strBaseString = strBaseString & "  Dim pContributorPropSet As IPropertySet" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  Dim lngPropSetIndex As Long" & vbNewLine
'  strBaseString = strBaseString & "  Dim strFClassNames(8) As String" & vbNewLine
'  strBaseString = strBaseString & "  strFClassNames(0) = ""TransportationLine_UTM12_NAD83""" & vbNewLine
'  strBaseString = strBaseString & "  strFClassNames(1) = ""KNF_UTM12_NAD83""" & vbNewLine
'  strBaseString = strBaseString & "  strFClassNames(2) = ""TIGER_2012_Trans_Near_Park_UTM12_NAD83""" & vbNewLine
'  strBaseString = strBaseString & "  strFClassNames(3) = ""BLM_AZStrip_UTM12_NAD83""" & vbNewLine
'  strBaseString = strBaseString & "  strFClassNames(4) = ""Roads_UTM12_NAD83""" & vbNewLine
'  strBaseString = strBaseString & "  strFClassNames(5) = ""Trails_UTM12_NAD83""" & vbNewLine
'  strBaseString = strBaseString & "  strFClassNames(6) = ""Routes_UTM12_NAD83""" & vbNewLine
'  strBaseString = strBaseString & "  strFClassNames(7) = ""Utah_Roads_Near_AOI_UTM12_NAD83""" & vbNewLine
'  strBaseString = strBaseString & "  strFClassNames(8) = ""Trails_KNF_UTM12_NAD83""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  Dim strFClassName As String" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "   " & vbNewLine
'  strBaseString = strBaseString & "  strAbstract = ""idinfo/descript/abstract""" & vbNewLine
'  strBaseString = strBaseString & "  strSummaryString = ""dataIdInfo/idPurp""" & vbNewLine
'  strBaseString = strBaseString & "  strDescriptionString = ""dataIdInfo/idAbs""" & vbNewLine
'  strBaseString = strBaseString & "  strCreditString = ""dataIdInfo/idCredit""" & vbNewLine
'  strBaseString = strBaseString & "  strUseLimitationsString = ""dataIdInfo/resConst/Consts/useLimit""" & vbNewLine
'  strBaseString = strBaseString & "  strUseLimitationsString2 = ""idinfo/useconst""" & vbNewLine
'  strBaseString = strBaseString & "  strThemeKeywords = ""dataIdInfo/themeKeys/keyword""" & vbNewLine
'  strBaseString = strBaseString & "  strSearchKeywords = ""dataIdInfo/searchKeys/keyword""" & vbNewLine
'  strBaseString = strBaseString & "  strSupplementalInformation = ""dataIdInfo/suppInfo""" & vbNewLine
'  strBaseString = strBaseString & "  strResponsiblePartyName = ""dataIdInfo/idCitation/citRespParty/rpOrgName""" & vbNewLine
'  strBaseString = strBaseString & "  strOtherConstraints = ""dataIdInfo/resConst/LegConsts/othConsts""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "'  Data Quality?  in dqInfo/report/measDesc" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "   " & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  ' --------------------------------------------------------------------------------" & vbNewLine
'  strBaseString = strBaseString & "  ' <><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>" & vbNewLine
'  strBaseString = strBaseString & "  ' --------------------------------------------------------------------------------" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  ' ---------------------------------------------------------------------------------------------" & vbNewLine
'  strBaseString = strBaseString & "  '  478] mdContact/rpIndName  '    --> Mark Nebel" & vbNewLine
'  strBaseString = strBaseString & "  '  479] mdContact/rpOrgName  '    --> Grand Canyon National Park" & vbNewLine
'  strBaseString = strBaseString & "  '  480] mdContact/rpPosName  '    --> GIS Coordinator" & vbNewLine
'  strBaseString = strBaseString & "  '  481] mdContact/rpCntInfo/cntPhone/voiceNum  '    --> 928-638-7451" & vbNewLine
'  strBaseString = strBaseString & "  '  482] mdContact/rpCntInfo/cntAddress/delPoint  '    --> 1824 S. Thompson St" & vbNewLine
'  strBaseString = strBaseString & "  '  483] mdContact/rpCntInfo/cntAddress/delPoint  '    --> Suite 200" & vbNewLine
'  strBaseString = strBaseString & "  '  484] mdContact/rpCntInfo/cntAddress/city  '    --> Flagstaff" & vbNewLine
'  strBaseString = strBaseString & "  '  485] mdContact/rpCntInfo/cntAddress/adminArea  '    --> AZ" & vbNewLine
'  strBaseString = strBaseString & "  '  486] mdContact/rpCntInfo/cntAddress/postCode  '    --> 86001" & vbNewLine
'  strBaseString = strBaseString & "  '  487] mdContact/rpCntInfo/cntAddress/country  '    --> UAS" & vbNewLine
'  strBaseString = strBaseString & "  '  488] mdContact/rpCntInfo/cntAddress/eMailAdd  '    --> mark_nebel@nps.gov" & vbNewLine
'  strBaseString = strBaseString & "  '  489] mdContact/role/RoleCd  '    -->" & vbNewLine
'  strBaseString = strBaseString & "  '" & vbNewLine
'  strBaseString = strBaseString & "  '  Jill M. Rundall" & vbNewLine
'  strBaseString = strBaseString & "  '  Operations manager/Spatial analyst" & vbNewLine
'  strBaseString = strBaseString & "  '  Lab of Landscape Ecology and Conservation Biology" & vbNewLine
'  strBaseString = strBaseString & "  '  Ph: 928-523-4730" & vbNewLine
'  strBaseString = strBaseString & "  '  Fax: 928-523-7078" & vbNewLine
'  strBaseString = strBaseString & "  'E:   Jill.Rundall@ nau.edu" & vbNewLine
'  strBaseString = strBaseString & "  '" & vbNewLine
'  strBaseString = strBaseString & "  '  School of Earth Sciences and Environmental Sustainability" & vbNewLine
'  strBaseString = strBaseString & "  '  College of Engineering, Forestry and Natural Sciences" & vbNewLine
'  strBaseString = strBaseString & "  '  Northern Arizona University, Flagstaff, AZ 86011" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "   " & vbNewLine
'  strBaseString = strBaseString & "'  ' PRIMARY CONTACT INFORMATION" & vbNewLine
'  strBaseString = strBaseString & "'  varAtts = pPropSet.GetProperty(""placeholder"")" & vbNewLine
'  strBaseString = strBaseString & "'  Debug.Print ""placeholder:  Count = "" & CStr(UBound(varAtts) + 1)" & vbNewLine
'  strBaseString = strBaseString & "'  varAtts = pPropSet.GetProperty(""mdContact/rpIndName"")" & vbNewLine
'  strBaseString = strBaseString & "'  Debug.Print ""mdContact/rpIndName:  Count = "" & CStr(UBound(varAtts) + 1)" & vbNewLine
'  strBaseString = strBaseString & "'  varAtts = pPropSet.GetProperty(""mdContact/rpOrgName"")" & vbNewLine
'  strBaseString = strBaseString & "'  Debug.Print ""mdContact/rpOrgName:  Count = "" & CStr(UBound(varAtts) + 1)" & vbNewLine
'  strBaseString = strBaseString & "'  varAtts = pPropSet.GetProperty(""mdContact/rpPosName"")" & vbNewLine
'  strBaseString = strBaseString & "'  Debug.Print ""mdContact/rpPosName:  Count = "" & CStr(UBound(varAtts) + 1)" & vbNewLine
'  strBaseString = strBaseString & "'  varAtts = pPropSet.GetProperty(""mdContact/rpCntInfo/cntPhone/voiceNum"")" & vbNewLine
'  strBaseString = strBaseString & "'  Debug.Print ""mdContact/rpCntInfo/cntPhone/voiceNum:  Count = "" & CStr(UBound(varAtts) + 1)" & vbNewLine
'  strBaseString = strBaseString & "'  varAtts = pPropSet.GetProperty(""mdContact/rpCntInfo/cntAddress/delPoint"")" & vbNewLine
'  strBaseString = strBaseString & "'  Debug.Print ""mdContact/rpCntInfo/cntAddress/delPoint:  Count = "" & CStr(UBound(varAtts) + 1)" & vbNewLine
'  strBaseString = strBaseString & "'  varAtts = pPropSet.GetProperty(""mdContact/rpCntInfo/cntAddress/delPoint"")" & vbNewLine
'  strBaseString = strBaseString & "'  Debug.Print ""mdContact/rpCntInfo/cntAddress/delPoint:  Count = "" & CStr(UBound(varAtts) + 1)" & vbNewLine
'  strBaseString = strBaseString & "'  varAtts = pPropSet.GetProperty(""mdContact/rpCntInfo/cntAddress/city"")" & vbNewLine
'  strBaseString = strBaseString & "'  Debug.Print ""mdContact/rpCntInfo/cntAddress/city:  Count = "" & CStr(UBound(varAtts) + 1)" & vbNewLine
'  strBaseString = strBaseString & "'  varAtts = pPropSet.GetProperty(""mdContact/rpCntInfo/cntAddress/adminArea"")" & vbNewLine
'  strBaseString = strBaseString & "'  Debug.Print ""mdContact/rpCntInfo/cntAddress/adminArea:  Count = "" & CStr(UBound(varAtts) + 1)" & vbNewLine
'  strBaseString = strBaseString & "'  varAtts = pPropSet.GetProperty(""mdContact/rpCntInfo/cntAddress/postCode"")" & vbNewLine
'  strBaseString = strBaseString & "'  Debug.Print ""mdContact/rpCntInfo/cntAddress/postCode:  Count = "" & CStr(UBound(varAtts) + 1)" & vbNewLine
'  strBaseString = strBaseString & "'  varAtts = pPropSet.GetProperty(""mdContact/rpCntInfo/cntAddress/country"")" & vbNewLine
'  strBaseString = strBaseString & "'  Debug.Print ""mdContact/rpCntInfo/cntAddress/country:  Count = "" & CStr(UBound(varAtts) + 1)" & vbNewLine
'  strBaseString = strBaseString & "'  varAtts = pPropSet.GetProperty(""mdContact/rpCntInfo/cntAddress/eMailAdd"")" & vbNewLine
'  strBaseString = strBaseString & "'  Debug.Print ""mdContact/rpCntInfo/cntAddress/eMailAdd:  Count = "" & CStr(UBound(varAtts) + 1)" & vbNewLine
'  strBaseString = strBaseString & "'  varAtts = pPropSet.GetProperty(""mdContact/role/RoleCd"")" & vbNewLine
'  strBaseString = strBaseString & "'  Debug.Print ""mdContact/role/RoleCd:  Count = "" & CStr(UBound(varAtts) + 1)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "'  pPropSet.RemoveProperty ""mdContact/rpIndName""" & vbNewLine
'  strBaseString = strBaseString & "'  pPropSet.RemoveProperty ""mdContact/rpOrgName""" & vbNewLine
'  strBaseString = strBaseString & "'  pPropSet.RemoveProperty ""mdContact/rpPosName""" & vbNewLine
'  strBaseString = strBaseString & "'  pPropSet.RemoveProperty ""mdContact/rpCntInfo/cntPhone/voiceNum""" & vbNewLine
'  strBaseString = strBaseString & "'  pPropSet.RemoveProperty ""mdContact/rpCntInfo/cntAddress/delPoint""" & vbNewLine
'  strBaseString = strBaseString & "'  pPropSet.RemoveProperty ""mdContact/rpCntInfo/cntAddress/city""" & vbNewLine
'  strBaseString = strBaseString & "'  pPropSet.RemoveProperty ""mdContact/rpCntInfo/cntAddress/adminArea""" & vbNewLine
'  strBaseString = strBaseString & "'  pPropSet.RemoveProperty ""mdContact/rpCntInfo/cntAddress/postCode""" & vbNewLine
'  strBaseString = strBaseString & "'  pPropSet.RemoveProperty ""mdContact/rpCntInfo/cntAddress/country""" & vbNewLine
'  strBaseString = strBaseString & "'  pPropSet.RemoveProperty ""mdContact/rpCntInfo/cntAddress/eMailAdd""" & vbNewLine
'  strBaseString = strBaseString & "'  pPropSet.RemoveProperty ""mdContact/role/RoleCd""" & vbNewLine
'  strBaseString = strBaseString & "'" & vbNewLine
'  strBaseString = strBaseString & "'" & vbNewLine
'  strBaseString = strBaseString & "'  pPropSet.SetProperty ""mdContact/rpIndName"", ""Jill M. Rundall""" & vbNewLine
'  strBaseString = strBaseString & "'  pPropSet.SetProperty ""mdContact/rpOrgName[0]"", ""Lab of Landscape Ecology and Conservation Biology""" & vbNewLine
'  strBaseString = strBaseString & "'  pPropSet.SetProperty ""mdContact/rpOrgName[1]"", ""School of Earth Sciences and Environmental Sustainability""" & vbNewLine
'  strBaseString = strBaseString & "'  pPropSet.SetProperty ""mdContact/rpOrgName[2]"", ""College of Engineering, Forestry and Natural Sciences""" & vbNewLine
'  strBaseString = strBaseString & "'  pPropSet.SetProperty ""mdContact/rpOrgName[3]"", ""Northern Arizona University""" & vbNewLine
'  strBaseString = strBaseString & "'  pPropSet.SetProperty ""mdContact/rpPosName"", ""Operations manager/Spatial analyst""" & vbNewLine
'  strBaseString = strBaseString & "'  pPropSet.SetProperty ""mdContact/rpCntInfo/cntPhone/voiceNum"", ""928-523-4730""" & vbNewLine
'  strBaseString = strBaseString & "'  pPropSet.SetProperty ""mdContact/rpCntInfo/cntAddress/delPoint[0]"", ""602 S Humphreys""" & vbNewLine
'  strBaseString = strBaseString & "'  pPropSet.SetProperty ""mdContact/rpCntInfo/cntAddress/delPoint[1]"", ""PO Box: 5694""" & vbNewLine
'  strBaseString = strBaseString & "'  pPropSet.SetProperty ""mdContact/rpCntInfo/cntAddress/city"", ""Flagstaff""" & vbNewLine
'  strBaseString = strBaseString & "'  pPropSet.SetProperty ""mdContact/rpCntInfo/cntAddress/adminArea"", ""AZ""" & vbNewLine
'  strBaseString = strBaseString & "'  pPropSet.SetProperty ""mdContact/rpCntInfo/cntAddress/postCode"", ""86011""" & vbNewLine
'  strBaseString = strBaseString & "'  pPropSet.SetProperty ""mdContact/rpCntInfo/cntAddress/country"", ""USA""" & vbNewLine
'  strBaseString = strBaseString & "'  pPropSet.SetProperty ""mdContact/rpCntInfo/cntAddress/eMailAdd"", ""Jill.Rundall@nau.edu""" & vbNewLine
'  strBaseString = strBaseString & "'  pPropSet.SetProperty ""mdContact/role/RoleCd"", """"" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.RemoveProperty ""distInfo/distFormat/formatVer""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""distInfo/distFormat/formatVer"", 10.1" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "'  varAtts = pPropSet.GetProperty(""dataIdInfo/idCitation/citRespParty/rpOrgName"")" & vbNewLine
'  strBaseString = strBaseString & "'  Debug.Print ""dataIdInfo/idCitation/citRespParty/rpOrgName:  Count = "" & CStr(UBound(varAtts) + 1)" & vbNewLine
'  strBaseString = strBaseString & "'  varAtts = pPropSet.GetProperty(""dataIdInfo/idCitation/citRespParty/role/RoleCd"")" & vbNewLine
'  strBaseString = strBaseString & "'  Debug.Print ""dataIdInfo/idCitation/citRespParty/role/RoleCd:  Count = "" & CStr(UBound(varAtts) + 1)" & vbNewLine
'  strBaseString = strBaseString & "'  varAtts = pPropSet.GetProperty(""dataIdInfo/idCitation/presForm/PresFormCd"")" & vbNewLine
'  strBaseString = strBaseString & "'  Debug.Print ""dataIdInfo/idCitation/presForm/PresFormCd:  Count = "" & CStr(UBound(varAtts) + 1)" & vbNewLine
'  strBaseString = strBaseString & "'  varAtts = pPropSet.GetProperty(""dataIdInfo/idCitation/resTitle"")" & vbNewLine
'  strBaseString = strBaseString & "'  Debug.Print ""dataIdInfo/idCitation/resTitle:  Count = "" & CStr(UBound(varAtts) + 1)" & vbNewLine
'  strBaseString = strBaseString & "'  varAtts = pPropSet.GetProperty(""dataIdInfo/idStatus/ProgCd"")" & vbNewLine
'  strBaseString = strBaseString & "'  Debug.Print ""dataIdInfo/idStatus/ProgCd:  Count = "" & CStr(UBound(varAtts) + 1)" & vbNewLine
'  strBaseString = strBaseString & "'  varAtts = pPropSet.GetProperty(""dataIdInfo/idPoC/rpIndName"")" & vbNewLine
'  strBaseString = strBaseString & "'  Debug.Print ""dataIdInfo/idPoC/rpIndName:  Count = "" & CStr(UBound(varAtts) + 1)" & vbNewLine
'  strBaseString = strBaseString & "'  varAtts = pPropSet.GetProperty(""dataIdInfo/idPoC/rpOrgName"")" & vbNewLine
'  strBaseString = strBaseString & "'  Debug.Print ""dataIdInfo/idPoC/rpOrgName:  Count = "" & CStr(UBound(varAtts) + 1)" & vbNewLine
'  strBaseString = strBaseString & "'  varAtts = pPropSet.GetProperty(""dataIdInfo/idPoC/rpPosName"")" & vbNewLine
'  strBaseString = strBaseString & "'  Debug.Print ""dataIdInfo/idPoC/rpPosName:  Count = "" & CStr(UBound(varAtts) + 1)" & vbNewLine
'  strBaseString = strBaseString & "'  varAtts = pPropSet.GetProperty(""dataIdInfo/idPoC/rpCntInfo/cntPhone/voiceNum"")" & vbNewLine
'  strBaseString = strBaseString & "'  Debug.Print ""dataIdInfo/idPoC/rpCntInfo/cntPhone/voiceNum:  Count = "" & CStr(UBound(varAtts) + 1)" & vbNewLine
'  strBaseString = strBaseString & "'  varAtts = pPropSet.GetProperty(""dataIdInfo/idPoC/rpCntInfo/cntAddress/delPoint"")" & vbNewLine
'  strBaseString = strBaseString & "'  Debug.Print ""dataIdInfo/idPoC/rpCntInfo/cntAddress/delPoint:  Count = "" & CStr(UBound(varAtts) + 1)" & vbNewLine
'  strBaseString = strBaseString & "'  varAtts = pPropSet.GetProperty(""dataIdInfo/idPoC/rpCntInfo/cntAddress/city"")" & vbNewLine
'  strBaseString = strBaseString & "'  Debug.Print ""dataIdInfo/idPoC/rpCntInfo/cntAddress/city:  Count = "" & CStr(UBound(varAtts) + 1)" & vbNewLine
'  strBaseString = strBaseString & "'  varAtts = pPropSet.GetProperty(""dataIdInfo/idPoC/rpCntInfo/cntAddress/adminArea"")" & vbNewLine
'  strBaseString = strBaseString & "'  Debug.Print ""dataIdInfo/idPoC/rpCntInfo/cntAddress/adminArea:  Count = "" & CStr(UBound(varAtts) + 1)" & vbNewLine
'  strBaseString = strBaseString & "'  varAtts = pPropSet.GetProperty(""dataIdInfo/idPoC/rpCntInfo/cntAddress/postCode"")" & vbNewLine
'  strBaseString = strBaseString & "'  Debug.Print ""dataIdInfo/idPoC/rpCntInfo/cntAddress/postCode:  Count = "" & CStr(UBound(varAtts) + 1)" & vbNewLine
'  strBaseString = strBaseString & "'  varAtts = pPropSet.GetProperty(""dataIdInfo/idPoC/rpCntInfo/cntAddress/country"")" & vbNewLine
'  strBaseString = strBaseString & "'  Debug.Print ""dataIdInfo/idPoC/rpCntInfo/cntAddress/country:  Count = "" & CStr(UBound(varAtts) + 1)" & vbNewLine
'  strBaseString = strBaseString & "'  varAtts = pPropSet.GetProperty(""dataIdInfo/idPoC/rpCntInfo/cntAddress/eMailAdd"")" & vbNewLine
'  strBaseString = strBaseString & "'  Debug.Print ""dataIdInfo/idPoC/rpCntInfo/cntAddress/eMailAdd:  Count = "" & CStr(UBound(varAtts) + 1)" & vbNewLine
'  strBaseString = strBaseString & "'  varAtts = pPropSet.GetProperty(""dataIdInfo/idPoC/role/RoleCd"")" & vbNewLine
'  strBaseString = strBaseString & "'  Debug.Print ""dataIdInfo/idPoC/role/RoleCd:  Count = "" & CStr(UBound(varAtts) + 1)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  ' LINEAGE =================================================================" & vbNewLine
'  strBaseString = strBaseString & "  Dim strCompiledData As String" & vbNewLine
'  strBaseString = strBaseString & "  strCompiledData = ""Compiled transportation features and attributes "" & _" & vbNewLine
'  strBaseString = strBaseString & "    ""from multiple feature classes (""""TransportationLine_UTM12_NAD83"""", """"Roads_KNF_UTM12_NAD83"""", "" & _" & vbNewLine
'  strBaseString = strBaseString & "    """"""BLM_AZStrip_UTM12_NAD83"""", """"Roads_UTM12_NAD83"""", """"UTAH_Roads_Near_AOI_UTM12_NAD83"""", "" & _" & vbNewLine
'  strBaseString = strBaseString & "    """"""Routes_UTM12_NAD83"""", """"Trails_UTM12_NAD83"""", """"TIGER_2012_Trans_Near_Park_UTM12_NAD83"""" "" & _" & vbNewLine
'  strBaseString = strBaseString & "    ""and """"UTAH_Roads_Near_AOI_UTM12_NAD83"""") into a single comprehensive transportation dataset "" & _" & vbNewLine
'  strBaseString = strBaseString & "    ""covering the Grand Canyon region.  Also identified the 4 closest features from each of these "" & _" & vbNewLine
'  strBaseString = strBaseString & "    ""feature classes to each compiled feature, and used these 4 closest features to estimate "" & _" & vbNewLine
'  strBaseString = strBaseString & "    ""the name, type, surface type and use level of each feature.""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  Dim strCompiledMetadata As String" & vbNewLine
'  strBaseString = strBaseString & "  strCompiledMetadata = ""Developed metadata that describes this compiled transportation feature class, "" & _" & vbNewLine
'  strBaseString = strBaseString & "    ""including imported metadata from the various contributor feature classes "" & _" & vbNewLine
'  strBaseString = strBaseString & "    ""(""""TransportationLine_UTM12_NAD83"""", """"Roads_KNF_UTM12_NAD83"""", "" & _" & vbNewLine
'  strBaseString = strBaseString & "    """"""BLM_AZStrip_UTM12_NAD83"""", """"Roads_UTM12_NAD83"""", """"UTAH_Roads_Near_AOI_UTM12_NAD83"""", "" & _" & vbNewLine
'  strBaseString = strBaseString & "    """"""Routes_UTM12_NAD83"""", """"Trails_UTM12_NAD83"""", """"TIGER_2012_Trans_Near_Park_UTM12_NAD83"""" "" & _" & vbNewLine
'  strBaseString = strBaseString & "    ""and """"UTAH_Roads_Near_AOI_UTM12_NAD83"""")""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  Dim varVals As Variant" & vbNewLine
'  strBaseString = strBaseString & "  Dim varPerson As Variant" & vbNewLine
'  strBaseString = strBaseString & "  Dim varDesc As Variant" & vbNewLine
'  strBaseString = strBaseString & "  Dim strStepDesc As String" & vbNewLine
'  strBaseString = strBaseString & "  Dim strStepPerson As String" & vbNewLine
'  strBaseString = strBaseString & "  Dim lngLineageCounter As Long" & vbNewLine
'  strBaseString = strBaseString & "  Dim lngCompiledDataIndex As Long" & vbNewLine
'  strBaseString = strBaseString & "  Dim lngCompiledMetadataIndex As Long" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  lngCompiledDataIndex = -999" & vbNewLine
'  strBaseString = strBaseString & "  lngCompiledMetadataIndex = -999" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  varVals = Array(""placeholder"")" & vbNewLine
'  strBaseString = strBaseString & "   " & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  lngLineageCounter = -1" & vbNewLine
'  strBaseString = strBaseString & "  Do Until IsEmpty(varVals)" & vbNewLine
'  strBaseString = strBaseString & "    lngLineageCounter = lngLineageCounter + 1" & vbNewLine
'  strBaseString = strBaseString & "    varVals = pPropSet.GetProperty(""dqInfo/dataLineage/prcStep["" & CStr(lngLineageCounter) & ""]"")" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "    If Not IsEmpty(varVals) Then" & vbNewLine
'  strBaseString = strBaseString & "      varPerson = pPropSet.GetProperty(""dqInfo/dataLineage/prcStep["" & CStr(lngLineageCounter) & ""]/stepProc/rpIndName"")" & vbNewLine
'  strBaseString = strBaseString & "      strStepPerson = CStr(varPerson(0))" & vbNewLine
'  strBaseString = strBaseString & "      varDesc = pPropSet.GetProperty(""dqInfo/dataLineage/prcStep["" & CStr(lngLineageCounter) & ""]/stepDesc"")" & vbNewLine
'  strBaseString = strBaseString & "      strStepDesc = CStr(varDesc(0))" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      If InStr(1, strStepPerson, ""Jenness"", vbTextCompare) > 0 And _" & vbNewLine
'  strBaseString = strBaseString & "         InStr(1, strStepDesc, "" Also identified the 4 closest features "", vbTextCompare) > 0 Then" & vbNewLine
'  strBaseString = strBaseString & "            lngCompiledDataIndex = lngLineageCounter" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      If InStr(1, strStepPerson, ""Jenness"", vbTextCompare) > 0 And _" & vbNewLine
'  strBaseString = strBaseString & "         InStr(1, strStepDesc, ""from the various contributor feature classes"", vbTextCompare) > 0 Then" & vbNewLine
'  strBaseString = strBaseString & "            lngCompiledMetadataIndex = lngLineageCounter" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "'    pXMLPropSet.GetAttribute ""dqInfo/dataLineage/prcStep["" & CStr(lngIndex) & _" & vbNewLine
'  strBaseString = strBaseString & "'        ""]/stepProc/role/RoleCd"", ""value"", varAtts" & vbNewLine
'  strBaseString = strBaseString & "'    Debug.Print lngIndex & ""] "" & varVals2(0) & "":  Value = "" & CStr(varAtts(0)) & _" & vbNewLine
'  strBaseString = strBaseString & "'        "" [n = "" & CStr(UBound(varAtts)) & ""]""" & vbNewLine
'  strBaseString = strBaseString & "'    lngIndex = lngIndex + 1" & vbNewLine
'  strBaseString = strBaseString & "'    varVals = pPropSet.GetProperty(""dqInfo/dataLineage/prcStep["" & CStr(lngLineageCounter) & ""]"")" & vbNewLine
'  strBaseString = strBaseString & "  Loop" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  If lngCompiledDataIndex = -999 And lngCompiledMetadataIndex = -999 Then" & vbNewLine
'  strBaseString = strBaseString & "    lngCompiledDataIndex = lngLineageCounter" & vbNewLine
'  strBaseString = strBaseString & "    lngCompiledMetadataIndex = lngLineageCounter + 1" & vbNewLine
'  strBaseString = strBaseString & "  ElseIf lngCompiledDataIndex = -999 Then" & vbNewLine
'  strBaseString = strBaseString & "    lngCompiledDataIndex = lngLineageCounter" & vbNewLine
'  strBaseString = strBaseString & "  ElseIf lngCompiledMetadataIndex = -999 Then" & vbNewLine
'  strBaseString = strBaseString & "    lngCompiledMetadataIndex = lngLineageCounter" & vbNewLine
'  strBaseString = strBaseString & "  End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  Debug.Print ""Data Index = "" & CStr(lngCompiledDataIndex)" & vbNewLine
'  strBaseString = strBaseString & "  Debug.Print ""MetaData Index = "" & CStr(lngCompiledMetadataIndex)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "'  636] dqInfo/dataLineage/prcStep/stepDesc    --> Made edits consistent with changes to Road and Trail feature classes related to West Boundary Road (W-1), Walden Trail Access Road (W-1B), and the Waldron Trail. Confirmed Datum transformation." & vbNewLine
'  strBaseString = strBaseString & "'  637] dqInfo/dataLineage/prcStep/stepDateTm    --> 2013-03-11T00:00:00" & vbNewLine
'  strBaseString = strBaseString & "'  638] dqInfo/dataLineage/prcStep/stepProc/rpIndName    --> Mark Nebel" & vbNewLine
'  strBaseString = strBaseString & "'  639] dqInfo/dataLineage/prcStep/stepProc/rpOrgName    --> Grand Canyon National Park" & vbNewLine
'  strBaseString = strBaseString & "'  640] dqInfo/dataLineage/prcStep/stepProc/rpPosName    --> GIS Coordinator" & vbNewLine
'  strBaseString = strBaseString & "'  641] dqInfo/dataLineage/prcStep/stepProc/rpCntInfo/cntPhone/voiceNum    --> 928-638-7451" & vbNewLine
'  strBaseString = strBaseString & "'  642] dqInfo/dataLineage/prcStep/stepProc/rpCntInfo/cntAddress/delPoint    --> 1824 S. Thompson St" & vbNewLine
'  strBaseString = strBaseString & "'  643] dqInfo/dataLineage/prcStep/stepProc/rpCntInfo/cntAddress/delPoint    --> Suite 200" & vbNewLine
'  strBaseString = strBaseString & "'  644] dqInfo/dataLineage/prcStep/stepProc/rpCntInfo/cntAddress/city    --> Flagstaff" & vbNewLine
'  strBaseString = strBaseString & "'  645] dqInfo/dataLineage/prcStep/stepProc/rpCntInfo/cntAddress/adminArea    --> AZ" & vbNewLine
'  strBaseString = strBaseString & "'  646] dqInfo/dataLineage/prcStep/stepProc/rpCntInfo/cntAddress/postCode    --> 86001" & vbNewLine
'  strBaseString = strBaseString & "'  647] dqInfo/dataLineage/prcStep/stepProc/rpCntInfo/cntAddress/country    --> UAS" & vbNewLine
'  strBaseString = strBaseString & "'  648] dqInfo/dataLineage/prcStep/stepProc/rpCntInfo/cntAddress/eMailAdd    --> mark_nebel@nps.gov" & vbNewLine
'  strBaseString = strBaseString & "'  649] dqInfo/dataLineage/prcStep/stepProc/role/RoleCd    -->" & vbNewLine
'  strBaseString = strBaseString & "'  650] dqInfo/dataLineage/prcStep/stepProc/displayName    --> Mark Nebel" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.RemoveProperty ""dqInfo/dataLineage/prcStep["" & CStr(lngCompiledDataIndex) & ""]/stepDesc""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.RemoveProperty ""dqInfo/dataLineage/prcStep["" & CStr(lngCompiledDataIndex) & ""]/stepDateTm""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.RemoveProperty ""dqInfo/dataLineage/prcStep["" & CStr(lngCompiledDataIndex) & ""]/stepProc/rpIndName""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.RemoveProperty ""dqInfo/dataLineage/prcStep["" & CStr(lngCompiledDataIndex) & ""]/stepProc/rpOrgName""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.RemoveProperty ""dqInfo/dataLineage/prcStep["" & CStr(lngCompiledDataIndex) & ""]/stepProc/rpPosName""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.RemoveProperty ""dqInfo/dataLineage/prcStep["" & CStr(lngCompiledDataIndex) & ""]/stepProc/rpCntInfo/cntPhone/voiceNum""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.RemoveProperty ""dqInfo/dataLineage/prcStep["" & CStr(lngCompiledDataIndex) & ""]/stepProc/rpCntInfo/cntAddress/delPoint""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.RemoveProperty ""dqInfo/dataLineage/prcStep["" & CStr(lngCompiledDataIndex) & ""]/stepProc/rpCntInfo/cntAddress/city""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.RemoveProperty ""dqInfo/dataLineage/prcStep["" & CStr(lngCompiledDataIndex) & ""]/stepProc/rpCntInfo/cntAddress/adminArea""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.RemoveProperty ""dqInfo/dataLineage/prcStep["" & CStr(lngCompiledDataIndex) & ""]/stepProc/rpCntInfo/cntAddress/postCode""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.RemoveProperty ""dqInfo/dataLineage/prcStep["" & CStr(lngCompiledDataIndex) & ""]/stepProc/rpCntInfo/cntAddress/country""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.RemoveProperty ""dqInfo/dataLineage/prcStep["" & CStr(lngCompiledDataIndex) & ""]/stepProc/rpCntInfo/cntAddress/eMailAdd""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.RemoveProperty ""dqInfo/dataLineage/prcStep["" & CStr(lngCompiledDataIndex) & ""]/stepProc/role/RoleCd""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.RemoveProperty ""dqInfo/dataLineage/prcStep["" & CStr(lngCompiledDataIndex) & ""]/stepProc/displayName""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""dqInfo/dataLineage/prcStep["" & CStr(lngCompiledDataIndex) & _" & vbNewLine
'  strBaseString = strBaseString & "      ""]/stepDesc"", strCompiledData" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""dqInfo/dataLineage/prcStep["" & CStr(lngCompiledDataIndex) & _" & vbNewLine
'  strBaseString = strBaseString & "       ""]/stepDateTm"", ""2013-09-01T00:00:00""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""dqInfo/dataLineage/prcStep["" & CStr(lngCompiledDataIndex) & _" & vbNewLine
'  strBaseString = strBaseString & "       ""]/stepProc/rpIndName"", ""Jeff Jenness""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""dqInfo/dataLineage/prcStep["" & CStr(lngCompiledDataIndex) & _" & vbNewLine
'  strBaseString = strBaseString & "       ""]/stepProc/rpOrgName"", ""Jenness Enterprises""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""dqInfo/dataLineage/prcStep["" & CStr(lngCompiledDataIndex) & _" & vbNewLine
'  strBaseString = strBaseString & "       ""]/stepProc/rpPosName"", ""GIS Analyst""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""dqInfo/dataLineage/prcStep["" & CStr(lngCompiledDataIndex) & _" & vbNewLine
'  strBaseString = strBaseString & "       ""]/stepProc/rpCntInfo/cntPhone/voiceNum"", ""928-607-4638""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""dqInfo/dataLineage/prcStep["" & CStr(lngCompiledDataIndex) & _" & vbNewLine
'  strBaseString = strBaseString & "       ""]/stepProc/rpCntInfo/cntAddress/delPoint"", ""3020 N. Schevene Blvd.""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""dqInfo/dataLineage/prcStep["" & CStr(lngCompiledDataIndex) & _" & vbNewLine
'  strBaseString = strBaseString & "       ""]/stepProc/rpCntInfo/cntAddress/city"", ""Flagstaff""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""dqInfo/dataLineage/prcStep["" & CStr(lngCompiledDataIndex) & _" & vbNewLine
'  strBaseString = strBaseString & "       ""]/stepProc/rpCntInfo/cntAddress/adminArea"", ""AZ""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""dqInfo/dataLineage/prcStep["" & CStr(lngCompiledDataIndex) & _" & vbNewLine
'  strBaseString = strBaseString & "       ""]/stepProc/rpCntInfo/cntAddress/postCode"", ""86004""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""dqInfo/dataLineage/prcStep["" & CStr(lngCompiledDataIndex) & _" & vbNewLine
'  strBaseString = strBaseString & "       ""]/stepProc/rpCntInfo/cntAddress/country"", ""UNITED STATES""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""dqInfo/dataLineage/prcStep["" & CStr(lngCompiledDataIndex) & _" & vbNewLine
'  strBaseString = strBaseString & "       ""]/stepProc/rpCntInfo/cntAddress/eMailAdd"", ""jeffj@jennessent.com""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""dqInfo/dataLineage/prcStep["" & CStr(lngCompiledDataIndex) & _" & vbNewLine
'  strBaseString = strBaseString & "       ""]/stepProc/role/RoleCd"", """"" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""dqInfo/dataLineage/prcStep["" & CStr(lngCompiledDataIndex) & _" & vbNewLine
'  strBaseString = strBaseString & "       ""]/stepProc/displayName"", ""Jeff Jenness""" & vbNewLine
'  strBaseString = strBaseString & "  pXMLPropSet.SetAttribute ""dqInfo/dataLineage/prcStep["" & CStr(lngCompiledDataIndex) & _" & vbNewLine
'  strBaseString = strBaseString & "      ""]/stepProc/role/RoleCd"", ""value"", ""009"", esriXSPAAddOrReplace" & vbNewLine
'  strBaseString = strBaseString & "  pXMLPropSet.SetAttribute ""dqInfo/dataLineage/prcStep["" & CStr(lngCompiledDataIndex) & _" & vbNewLine
'  strBaseString = strBaseString & "       ""]/stepProc/rpCntInfo/cntAddress"", ""addressType"", ""both"", esriXSPAAddOrReplace" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.RemoveProperty ""dqInfo/dataLineage/prcStep["" & CStr(lngCompiledMetadataIndex) & ""]/stepDesc""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.RemoveProperty ""dqInfo/dataLineage/prcStep["" & CStr(lngCompiledMetadataIndex) & ""]/stepDateTm""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.RemoveProperty ""dqInfo/dataLineage/prcStep["" & CStr(lngCompiledMetadataIndex) & ""]/stepProc/rpIndName""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.RemoveProperty ""dqInfo/dataLineage/prcStep["" & CStr(lngCompiledMetadataIndex) & ""]/stepProc/rpOrgName""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.RemoveProperty ""dqInfo/dataLineage/prcStep["" & CStr(lngCompiledMetadataIndex) & ""]/stepProc/rpPosName""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.RemoveProperty ""dqInfo/dataLineage/prcStep["" & CStr(lngCompiledMetadataIndex) & ""]/stepProc/rpCntInfo/cntPhone/voiceNum""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.RemoveProperty ""dqInfo/dataLineage/prcStep["" & CStr(lngCompiledMetadataIndex) & ""]/stepProc/rpCntInfo/cntAddress/delPoint""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.RemoveProperty ""dqInfo/dataLineage/prcStep["" & CStr(lngCompiledMetadataIndex) & ""]/stepProc/rpCntInfo/cntAddress/city""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.RemoveProperty ""dqInfo/dataLineage/prcStep["" & CStr(lngCompiledMetadataIndex) & ""]/stepProc/rpCntInfo/cntAddress/adminArea""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.RemoveProperty ""dqInfo/dataLineage/prcStep["" & CStr(lngCompiledMetadataIndex) & ""]/stepProc/rpCntInfo/cntAddress/postCode""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.RemoveProperty ""dqInfo/dataLineage/prcStep["" & CStr(lngCompiledMetadataIndex) & ""]/stepProc/rpCntInfo/cntAddress/country""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.RemoveProperty ""dqInfo/dataLineage/prcStep["" & CStr(lngCompiledMetadataIndex) & ""]/stepProc/rpCntInfo/cntAddress/eMailAdd""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.RemoveProperty ""dqInfo/dataLineage/prcStep["" & CStr(lngCompiledMetadataIndex) & ""]/stepProc/role/RoleCd""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.RemoveProperty ""dqInfo/dataLineage/prcStep["" & CStr(lngCompiledMetadataIndex) & ""]/stepProc/displayName""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""dqInfo/dataLineage/prcStep["" & CStr(lngCompiledMetadataIndex) & _" & vbNewLine
'  strBaseString = strBaseString & "      ""]/stepDesc"", strCompiledMetadata" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""dqInfo/dataLineage/prcStep["" & CStr(lngCompiledMetadataIndex) & _" & vbNewLine
'  strBaseString = strBaseString & "       ""]/stepDateTm"", ""2013-11-18T00:00:00""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""dqInfo/dataLineage/prcStep["" & CStr(lngCompiledMetadataIndex) & _" & vbNewLine
'  strBaseString = strBaseString & "       ""]/stepProc/rpIndName"", ""Jeff Jenness""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""dqInfo/dataLineage/prcStep["" & CStr(lngCompiledMetadataIndex) & _" & vbNewLine
'  strBaseString = strBaseString & "       ""]/stepProc/rpOrgName"", ""Jenness Enterprises""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""dqInfo/dataLineage/prcStep["" & CStr(lngCompiledMetadataIndex) & _" & vbNewLine
'  strBaseString = strBaseString & "       ""]/stepProc/rpPosName"", ""GIS Analyst""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""dqInfo/dataLineage/prcStep["" & CStr(lngCompiledMetadataIndex) & _" & vbNewLine
'  strBaseString = strBaseString & "       ""]/stepProc/rpCntInfo/cntPhone/voiceNum"", ""928-607-4638""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""dqInfo/dataLineage/prcStep["" & CStr(lngCompiledMetadataIndex) & _" & vbNewLine
'  strBaseString = strBaseString & "       ""]/stepProc/rpCntInfo/cntAddress/delPoint"", ""3020 N. Schevene Blvd.""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""dqInfo/dataLineage/prcStep["" & CStr(lngCompiledMetadataIndex) & _" & vbNewLine
'  strBaseString = strBaseString & "       ""]/stepProc/rpCntInfo/cntAddress/city"", ""Flagstaff""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""dqInfo/dataLineage/prcStep["" & CStr(lngCompiledMetadataIndex) & _" & vbNewLine
'  strBaseString = strBaseString & "       ""]/stepProc/rpCntInfo/cntAddress/adminArea"", ""AZ""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""dqInfo/dataLineage/prcStep["" & CStr(lngCompiledMetadataIndex) & _" & vbNewLine
'  strBaseString = strBaseString & "       ""]/stepProc/rpCntInfo/cntAddress/postCode"", ""86004""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""dqInfo/dataLineage/prcStep["" & CStr(lngCompiledMetadataIndex) & _" & vbNewLine
'  strBaseString = strBaseString & "       ""]/stepProc/rpCntInfo/cntAddress/country"", ""UNITED STATES""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""dqInfo/dataLineage/prcStep["" & CStr(lngCompiledMetadataIndex) & _" & vbNewLine
'  strBaseString = strBaseString & "       ""]/stepProc/rpCntInfo/cntAddress/eMailAdd"", ""jeffj@jennessent.com""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""dqInfo/dataLineage/prcStep["" & CStr(lngCompiledMetadataIndex) & _" & vbNewLine
'  strBaseString = strBaseString & "       ""]/stepProc/role/RoleCd"", """"" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""dqInfo/dataLineage/prcStep["" & CStr(lngCompiledMetadataIndex) & _" & vbNewLine
'  strBaseString = strBaseString & "       ""]/stepProc/displayName"", ""Jeff Jenness""" & vbNewLine
'  strBaseString = strBaseString & "  pXMLPropSet.SetAttribute ""dqInfo/dataLineage/prcStep["" & CStr(lngCompiledMetadataIndex) & _" & vbNewLine
'  strBaseString = strBaseString & "      ""]/stepProc/role/RoleCd"", ""value"", ""009"", esriXSPAAddOrReplace" & vbNewLine
'  strBaseString = strBaseString & "  pXMLPropSet.SetAttribute ""dqInfo/dataLineage/prcStep["" & CStr(lngCompiledMetadataIndex) & _" & vbNewLine
'  strBaseString = strBaseString & "       ""]/stepProc/rpCntInfo/cntAddress"", ""addressType"", ""both"", esriXSPAAddOrReplace" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.RemoveProperty ""dataIdInfo/idCitation/citRespParty/rpIndName""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.RemoveProperty ""dataIdInfo/idCitation/citRespParty/rpOrgName""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.RemoveProperty ""dataIdInfo/idCitation/citRespParty/rpPosName""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.RemoveProperty ""dataIdInfo/idCitation/citRespParty/role/RoleCd""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.RemoveProperty ""dataIdInfo/idCitation/citRespParty/rpCntInfo/cntAddress/delPoint""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.RemoveProperty ""dataIdInfo/idCitation/citRespParty/rpCntInfo/cntAddress/city""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.RemoveProperty ""dataIdInfo/idCitation/citRespParty/rpCntInfo/cntAddress/adminArea""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.RemoveProperty ""dataIdInfo/idCitation/citRespParty/rpCntInfo/cntAddress/postCode""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.RemoveProperty ""dataIdInfo/idCitation/citRespParty/rpCntInfo/cntAddress/eMailAdd""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.RemoveProperty ""dataIdInfo/idCitation/citRespParty/rpCntInfo/cntAddress/country""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.RemoveProperty ""dataIdInfo/idCitation/citRespParty/rpCntInfo/cntPhone/voiceNum""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.RemoveProperty ""dataIdInfo/idCitation/date/pubDate""" & vbNewLine
'  strBaseString = strBaseString & "'  pPropSet.RemoveProperty ""dataIdInfo/idCitation/presForm/PresFormCd""" & vbNewLine
'  strBaseString = strBaseString & "'  pPropSet.RemoveProperty ""dataIdInfo/idCitation/resTitle""" & vbNewLine
'  strBaseString = strBaseString & "'  pPropSet.RemoveProperty ""dataIdInfo/idStatus/ProgCd""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.RemoveProperty ""dataIdInfo/idPoC/rpIndName""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.RemoveProperty ""dataIdInfo/idPoC/rpOrgName""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.RemoveProperty ""dataIdInfo/idPoC/rpPosName""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.RemoveProperty ""dataIdInfo/idPoC/rpCntInfo/cntPhone/voiceNum""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.RemoveProperty ""dataIdInfo/idPoC/rpCntInfo/cntAddress/delPoint""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.RemoveProperty ""dataIdInfo/idPoC/rpCntInfo/cntAddress/city""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.RemoveProperty ""dataIdInfo/idPoC/rpCntInfo/cntAddress/adminArea""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.RemoveProperty ""dataIdInfo/idPoC/rpCntInfo/cntAddress/postCode""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.RemoveProperty ""dataIdInfo/idPoC/rpCntInfo/cntAddress/country""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.RemoveProperty ""dataIdInfo/idPoC/rpCntInfo/cntAddress/eMailAdd""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.RemoveProperty ""dataIdInfo/idPoC/role/RoleCd""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  '  474] dataIdInfo/idCitation/citRespParty/rpOrgName  '    --> Grand Canyon National Park" & vbNewLine
'  strBaseString = strBaseString & "  '  475] dataIdInfo/idCitation/citRespParty/role/RoleCd  '    -->" & vbNewLine
'  strBaseString = strBaseString & "  '  476] dataIdInfo/idCitation/presForm/PresFormCd  '    -->" & vbNewLine
'  strBaseString = strBaseString & "  '  477] dataIdInfo/idCitation/resTitle  '    --> Transportation_Test_Metadata" & vbNewLine
'  strBaseString = strBaseString & "  '  478] dataIdInfo/idStatus/ProgCd  '    -->" & vbNewLine
'  strBaseString = strBaseString & "  '  479] dataIdInfo/idPoC/rpIndName  '    --> Mark Nebel" & vbNewLine
'  strBaseString = strBaseString & "  '  480] dataIdInfo/idPoC/rpOrgName  '    --> Grand Canyon National Park" & vbNewLine
'  strBaseString = strBaseString & "  '  481] dataIdInfo/idPoC/rpPosName  '    --> GIS Coordinator" & vbNewLine
'  strBaseString = strBaseString & "  '  482] dataIdInfo/idPoC/rpCntInfo/cntPhone/voiceNum  '    --> 928-638-7451" & vbNewLine
'  strBaseString = strBaseString & "  '  483] dataIdInfo/idPoC/rpCntInfo/cntAddress/delPoint  '    --> 1824 S Thompson St" & vbNewLine
'  strBaseString = strBaseString & "  '  484] dataIdInfo/idPoC/rpCntInfo/cntAddress/delPoint  '    --> Suite 200" & vbNewLine
'  strBaseString = strBaseString & "  '  485] dataIdInfo/idPoC/rpCntInfo/cntAddress/city  '    --> Flagstaff" & vbNewLine
'  strBaseString = strBaseString & "  '  486] dataIdInfo/idPoC/rpCntInfo/cntAddress/adminArea  '    --> AZ" & vbNewLine
'  strBaseString = strBaseString & "  '  487] dataIdInfo/idPoC/rpCntInfo/cntAddress/postCode  '    --> 86001" & vbNewLine
'  strBaseString = strBaseString & "  '  488] dataIdInfo/idPoC/rpCntInfo/cntAddress/country  '    --> US" & vbNewLine
'  strBaseString = strBaseString & "  '  489] dataIdInfo/idPoC/rpCntInfo/cntAddress/eMailAdd  '    --> mark_nebel@nps.gov" & vbNewLine
'  strBaseString = strBaseString & "  '  490] dataIdInfo/idPoC/role/RoleCd  '    -->" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  ' FOR RoleCD:  ""006"" = ""Originator"";   ""007"" = ""Point of Contact""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""dataIdInfo/idCitation/citRespParty/rpIndName"", ""Jill M. Rundall""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""dataIdInfo/idCitation/citRespParty/rpOrgName[0]"", ""Lab of Landscape Ecology and Conservation Biology""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""dataIdInfo/idCitation/citRespParty/rpOrgName[1]"", ""School of Earth Sciences and Environmental Sustainability""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""dataIdInfo/idCitation/citRespParty/rpOrgName[2]"", ""College of Engineering, Forestry and Natural Sciences""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""dataIdInfo/idCitation/citRespParty/rpOrgName[3]"", ""Northern Arizona University""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""dataIdInfo/idCitation/citRespParty/rpPosName"", ""Operations manager/Spatial analyst""" & vbNewLine
'  strBaseString = strBaseString & "  pXMLPropSet.SetAttribute ""dataIdInfo/idCitation/citRespParty/role/RoleCd"", ""value"", ""006"", esriXSPAAddOrReplace" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""dataIdInfo/idCitation/citRespParty/rpCntInfo/cntAddress/delPoint[0]"", ""602 S Humphreys""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""dataIdInfo/idCitation/citRespParty/rpCntInfo/cntAddress/delPoint[1]"", ""P.O Box 5694""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""dataIdInfo/idCitation/citRespParty/rpCntInfo/cntAddress/city"", ""Flagstaff""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""dataIdInfo/idCitation/citRespParty/rpCntInfo/cntAddress/adminArea"", ""AZ""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""dataIdInfo/idCitation/citRespParty/rpCntInfo/cntAddress/postCode"", ""86011""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""dataIdInfo/idCitation/citRespParty/rpCntInfo/cntAddress/eMailAdd"", ""Jill.Rundall@nau.edu""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""dataIdInfo/idCitation/citRespParty/rpCntInfo/cntAddress/country"", ""UNITED STATES""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""dataIdInfo/idCitation/citRespParty/rpCntInfo/cntPhone/voiceNum"", ""928-523-4730""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""dataIdInfo/idCitation/date/pubDate"", ""2013-11-17T00:00:00""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "'  pPropSet.SetProperty ""dataIdInfo/idCitation/presForm/PresFormCd""" & vbNewLine
'  strBaseString = strBaseString & "'  pPropSet.SetProperty ""dataIdInfo/idCitation/citRespParty/role/RoleCd"", """"" & vbNewLine
'  strBaseString = strBaseString & "'  pPropSet.SetProperty ""dataIdInfo/idCitation/presForm/PresFormCd"", """"" & vbNewLine
'  strBaseString = strBaseString & "'  pPropSet.SetProperty ""dataIdInfo/idCitation/resTitle"", """"" & vbNewLine
'  strBaseString = strBaseString & "'  pPropSet.SetProperty ""dataIdInfo/idStatus/ProgCd"", """"" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""dataIdInfo/idPoC/rpIndName"", ""Jill M. Rundall""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""dataIdInfo/idPoC/rpOrgName[0]"", ""Lab of Landscape Ecology and Conservation Biology""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""dataIdInfo/idPoC/rpOrgName[1]"", ""School of Earth Sciences and Environmental Sustainability""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""dataIdInfo/idPoC/rpOrgName[2]"", ""College of Engineering, Forestry and Natural Sciences""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""dataIdInfo/idPoC/rpOrgName[3]"", ""Northern Arizona University""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""dataIdInfo/idPoC/rpPosName"", ""Operations manager/Spatial analyst""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""dataIdInfo/idPoC/rpCntInfo/cntPhone/voiceNum"", ""928-523-4730""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""dataIdInfo/idPoC/rpCntInfo/cntAddress/delPoint[0]"", ""602 S Humphreys""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""dataIdInfo/idPoC/rpCntInfo/cntAddress/delPoint[1]"", ""P.O Box 5694""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""dataIdInfo/idPoC/rpCntInfo/cntAddress/city"", ""Flagstaff""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""dataIdInfo/idPoC/rpCntInfo/cntAddress/adminArea"", ""AZ""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""dataIdInfo/idPoC/rpCntInfo/cntAddress/postCode"", ""86011""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""dataIdInfo/idPoC/rpCntInfo/cntAddress/country"", ""UNITED STATES""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""dataIdInfo/idPoC/rpCntInfo/cntAddress/eMailAdd"", ""Jill.Rundall@nau.edu""" & vbNewLine
'  strBaseString = strBaseString & "  pXMLPropSet.SetAttribute ""dataIdInfo/idPoC/role/RoleCd"", ""value"", ""007"", esriXSPAAddOrReplace" & vbNewLine
'  strBaseString = strBaseString & "'  pPropSet.SetProperty ""dataIdInfo/idPoC/role/RoleCd"", """"" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.RemoveProperty ""mdContact""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "'  pPropSet.RemoveProperty ""mdContact/rpIndName""" & vbNewLine
'  strBaseString = strBaseString & "'  pPropSet.RemoveProperty ""mdContact/rpOrgName""" & vbNewLine
'  strBaseString = strBaseString & "'  pPropSet.RemoveProperty ""mdContact/rpPosName""" & vbNewLine
'  strBaseString = strBaseString & "'  pPropSet.RemoveProperty ""mdContact/rpCntInfo/cntPhone/voiceNum""" & vbNewLine
'  strBaseString = strBaseString & "'  pPropSet.RemoveProperty ""mdContact/rpCntInfo/cntAddress/delPoint""" & vbNewLine
'  strBaseString = strBaseString & "'  pPropSet.RemoveProperty ""mdContact/rpCntInfo/cntAddress/city""" & vbNewLine
'  strBaseString = strBaseString & "'  pPropSet.RemoveProperty ""mdContact/rpCntInfo/cntAddress/adminArea""" & vbNewLine
'  strBaseString = strBaseString & "'  pPropSet.RemoveProperty ""mdContact/rpCntInfo/cntAddress/postCode""" & vbNewLine
'  strBaseString = strBaseString & "'  pPropSet.RemoveProperty ""mdContact/rpCntInfo/cntAddress/country""" & vbNewLine
'  strBaseString = strBaseString & "'  pPropSet.RemoveProperty ""mdContact/rpCntInfo/cntAddress/eMailAdd""" & vbNewLine
'  strBaseString = strBaseString & "'  pPropSet.RemoveProperty ""mdContact/role/RoleCd""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""mdContact[0]/rpIndName"", ""Jill M. Rundall""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""mdContact[0]/rpOrgName[0]"", ""Lab of Landscape Ecology and Conservation Biology""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""mdContact[0]/rpOrgName[1]"", ""School of Earth Sciences and Environmental Sustainability""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""mdContact[0]/rpOrgName[2]"", ""College of Engineering, Forestry and Natural Sciences""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""mdContact[0]/rpOrgName[3]"", ""Northern Arizona University""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""mdContact[0]/rpPosName"", ""Operations manager/Spatial analyst""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""mdContact[0]/rpCntInfo/cntPhone/voiceNum"", ""928-523-4730""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""mdContact[0]/rpCntInfo/cntAddress/delPoint[0]"", ""602 S Humphreys""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""mdContact[0]/rpCntInfo/cntAddress/delPoint[1]"", ""P.O Box 5694""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""mdContact[0]/rpCntInfo/cntAddress/city"", ""Flagstaff""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""mdContact[0]/rpCntInfo/cntAddress/adminArea"", ""AZ""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""mdContact[0]/rpCntInfo/cntAddress/postCode"", ""86011""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""mdContact[0]/rpCntInfo/cntAddress/country"", ""UNITED STATES""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""mdContact[0]/rpCntInfo/cntAddress/eMailAdd"", ""Jill.Rundall@nau.edu""" & vbNewLine
'  strBaseString = strBaseString & "  pXMLPropSet.SetAttribute ""mdContact[0]/role/RoleCd"", ""value"", ""007"", esriXSPAAddOrReplace" & vbNewLine
'  strBaseString = strBaseString & "  pXMLPropSet.SetAttribute ""mdContact[0]/rpCntInfo/cntAddress"", ""addressType"", ""both"", esriXSPAAddOrReplace" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""mdContact[1]/rpIndName"", ""Jeff Jenness""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""mdContact[1]/rpOrgName"", ""Jenness Enterprises""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""mdContact[1]/rpPosName"", ""GIS Analyst, Application Designer""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""mdContact[1]/rpCntInfo/cntPhone/voiceNum"", ""928-607-4638""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""mdContact[1]/rpCntInfo/cntAddress/delPoint"", ""3020 N. Schevene Blvd.""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""mdContact[1]/rpCntInfo/cntAddress/city"", ""Flagstaff""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""mdContact[1]/rpCntInfo/cntAddress/adminArea"", ""AZ""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""mdContact[1]/rpCntInfo/cntAddress/postCode"", ""86001""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""mdContact[1]/rpCntInfo/cntAddress/country"", ""UNITED STATES""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""mdContact[1]/rpCntInfo/cntAddress/eMailAdd"", ""jeffj@jennessent.com""" & vbNewLine
'  strBaseString = strBaseString & "  pXMLPropSet.SetAttribute ""mdContact[1]/role/RoleCd"", ""value"", ""006"", esriXSPAAddOrReplace" & vbNewLine
'  strBaseString = strBaseString & "  pXMLPropSet.SetAttribute ""mdContact[1]/rpCntInfo/cntAddress"", ""addressType"", ""both"", esriXSPAAddOrReplace" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "'  pPropSet.SetProperty ""mdContact/role/RoleCd"", """"" & vbNewLine
'  strBaseString = strBaseString & "  '  478] mdContact/rpIndName  '    --> Mark Nebel" & vbNewLine
'  strBaseString = strBaseString & "  '  479] mdContact/rpOrgName  '    --> Grand Canyon National Park" & vbNewLine
'  strBaseString = strBaseString & "  '  480] mdContact/rpPosName  '    --> GIS Coordinator" & vbNewLine
'  strBaseString = strBaseString & "  '  481] mdContact/rpCntInfo/cntPhone/voiceNum  '    --> 928-638-7451" & vbNewLine
'  strBaseString = strBaseString & "  '  482] mdContact/rpCntInfo/cntAddress/delPoint  '    --> 1824 S. Thompson St" & vbNewLine
'  strBaseString = strBaseString & "  '  483] mdContact/rpCntInfo/cntAddress/delPoint  '    --> Suite 200" & vbNewLine
'  strBaseString = strBaseString & "  '  484] mdContact/rpCntInfo/cntAddress/city  '    --> Flagstaff" & vbNewLine
'  strBaseString = strBaseString & "  '  485] mdContact/rpCntInfo/cntAddress/adminArea  '    --> AZ" & vbNewLine
'  strBaseString = strBaseString & "  '  486] mdContact/rpCntInfo/cntAddress/postCode  '    --> 86001" & vbNewLine
'  strBaseString = strBaseString & "  '  487] mdContact/rpCntInfo/cntAddress/country  '    --> UAS" & vbNewLine
'  strBaseString = strBaseString & "  '  488] mdContact/rpCntInfo/cntAddress/eMailAdd  '    --> mark_nebel@nps.gov" & vbNewLine
'  strBaseString = strBaseString & "  '  489] mdContact/role/RoleCd  '    -->" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  '" & vbNewLine
'  strBaseString = strBaseString & "  '  Jill M. Rundall" & vbNewLine
'  strBaseString = strBaseString & "  '  Operations manager/Spatial analyst" & vbNewLine
'  strBaseString = strBaseString & "  '  Lab of Landscape Ecology and Conservation Biology" & vbNewLine
'  strBaseString = strBaseString & "  '  Ph: 928-523-4730" & vbNewLine
'  strBaseString = strBaseString & "  '  Fax: 928-523-7078" & vbNewLine
'  strBaseString = strBaseString & "  'E:   Jill.Rundall@ nau.edu" & vbNewLine
'  strBaseString = strBaseString & "  '" & vbNewLine
'  strBaseString = strBaseString & "  '  School of Earth Sciences and Environmental Sustainability" & vbNewLine
'  strBaseString = strBaseString & "  '  College of Engineering, Forestry and Natural Sciences" & vbNewLine
'  strBaseString = strBaseString & "  '  Northern Arizona University, Flagstaff, AZ 86011" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  ' ---------------------------------------------------------------------------------------------" & vbNewLine
'  strBaseString = strBaseString & "  ' SUPPLEMENTAL INFORMATION" & vbNewLine
'  strBaseString = strBaseString & "  Dim strOtherConstraintsData As String" & vbNewLine
'  strBaseString = strBaseString & "  strOtherConstraintsData = ""Additional Constraints for each "" & _" & vbNewLine
'  strBaseString = strBaseString & "    ""of the contributor feature classes are listed below:"" & vbCrLf & vbCrLf" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  For lngPropSetIndex = 0 To pNewFClassArray.Count - 1" & vbNewLine
'  strBaseString = strBaseString & "    strFClassName = strFClassNames(lngPropSetIndex)" & vbNewLine
'  strBaseString = strBaseString & "    Set pSubArray = pNewFClassArray.Element(lngPropSetIndex)" & vbNewLine
'  strBaseString = strBaseString & "    Set pContributorPropSet = pSubArray.Element(1)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "'    Debug.Print ""Descriptions for '"" & strFClassName & ""'""" & vbNewLine
'  strBaseString = strBaseString & "    varAtts = pContributorPropSet.GetProperty(strOtherConstraints)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "    If IsEmpty(varAtts) Then" & vbNewLine
'  strBaseString = strBaseString & "'      varAtts = pContributorPropSet.GetProperty(strOtherConstraints)" & vbNewLine
'  strBaseString = strBaseString & "'" & vbNewLine
'  strBaseString = strBaseString & "'      If IsEmpty(varAtts) Then" & vbNewLine
'  strBaseString = strBaseString & "        strOtherConstraintsData = strOtherConstraintsData & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Additional Constraints for "" & strFClassName & _" & vbNewLine
'  strBaseString = strBaseString & "            ""  [-- No Additional Constraints Available --]"" & vbCrLf & vbCrLf" & vbNewLine
'  strBaseString = strBaseString & "'      Else" & vbNewLine
'  strBaseString = strBaseString & "'        For lngIndex = 0 To UBound(varAtts)" & vbNewLine
'  strBaseString = strBaseString & "'          strValue = CStr(varAtts(lngIndex))" & vbNewLine
'  strBaseString = strBaseString & "'          strValue = RemoveTags(strValue)" & vbNewLine
'  strBaseString = strBaseString & "''          Debug.Print ""  "" & CStr(lngIndex) & ""] "" & strValue" & vbNewLine
'  strBaseString = strBaseString & "'" & vbNewLine
'  strBaseString = strBaseString & "'          If InStr(1, strValue, ""DIV STYLE"", vbTextCompare) > 0 Then" & vbNewLine
'  strBaseString = strBaseString & "'            strOtherConstraintsData = strOtherConstraintsData & _" & vbNewLine
'  strBaseString = strBaseString & "'                ""<P><SPAN STYLE=""""font-weight:bold;"""">Use Limitations for "" & strFClassName & _" & vbNewLine
'  strBaseString = strBaseString & "'                "":  </SPAN><SPAN> "" & strValue & ""</SPAN></P>""" & vbNewLine
'  strBaseString = strBaseString & "'          Else" & vbNewLine
'  strBaseString = strBaseString & "'" & vbNewLine
'  strBaseString = strBaseString & "'            strOtherConstraintsData = strOtherConstraintsData & _" & vbNewLine
'  strBaseString = strBaseString & "'                ""<P><SPAN STYLE=""""font-weight:bold;"""">Use Limitations for "" & strFClassName & _" & vbNewLine
'  strBaseString = strBaseString & "'                "":  </SPAN><SPAN> "" & strValue & ""</SPAN></P>""" & vbNewLine
'  strBaseString = strBaseString & "'          End If" & vbNewLine
'  strBaseString = strBaseString & "'" & vbNewLine
'  strBaseString = strBaseString & "'        Next lngIndex" & vbNewLine
'  strBaseString = strBaseString & "'      End If" & vbNewLine
'  strBaseString = strBaseString & "    Else" & vbNewLine
'  strBaseString = strBaseString & "      For lngIndex = 0 To UBound(varAtts)" & vbNewLine
'  strBaseString = strBaseString & "        strValue = CStr(varAtts(lngIndex))" & vbNewLine
'  strBaseString = strBaseString & "        strValue = RemoveTags(strValue)" & vbNewLine
'  strBaseString = strBaseString & "'        Debug.Print ""  "" & CStr(lngIndex) & ""] "" & strValue" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "        strOtherConstraintsData = strOtherConstraintsData & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Additional Constraints for "" & strFClassName & _" & vbNewLine
'  strBaseString = strBaseString & "            "":  "" & strValue & vbCrLf & vbCrLf" & vbNewLine
'  strBaseString = strBaseString & "      Next lngIndex" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "  Next lngPropSetIndex" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  strOtherConstraintsData = Replace(strOtherConstraintsData, ""TrasnportationLine"", ""TransportationLine"", , , vbTextCompare)" & vbNewLine
'  strBaseString = strBaseString & "'  strSupplemental = strSupplemental & ""</DIV></DIV>""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  varAtts = pPropSet.GetProperty(strOtherConstraints)" & vbNewLine
'  strBaseString = strBaseString & "  If Not IsEmpty(varAtts) Then" & vbNewLine
'  strBaseString = strBaseString & "    pPropSet.RemoveProperty strOtherConstraints" & vbNewLine
'  strBaseString = strBaseString & "  End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  varAtts = pPropSet.GetProperty(strOtherConstraints)" & vbNewLine
'  strBaseString = strBaseString & "  If Not IsEmpty(varAtts) Then" & vbNewLine
'  strBaseString = strBaseString & "    If UBound(varAtts) = 0 Then" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strOtherConstraints, strOtherConstraintsData" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "  Else" & vbNewLine
'  strBaseString = strBaseString & "    pPropSet.SetProperty strOtherConstraints, strOtherConstraintsData" & vbNewLine
'  strBaseString = strBaseString & "  End If" & vbNewLine
'  strBaseString = strBaseString & "'" & vbNewLine
'  strBaseString = strBaseString & "'  Debug.Print ""Final Description:""" & vbNewLine
'  strBaseString = strBaseString & "'  Debug.Print strFullDescription" & vbNewLine
'  strBaseString = strBaseString & "  'pPropSet.SetProperty strSummaryString & strfinalsummary" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  ' ---------------------------------------------------------------------------------------------" & vbNewLine
'  strBaseString = strBaseString & "  ' SUPPLEMENTAL INFORMATION" & vbNewLine
'  strBaseString = strBaseString & "  Dim strSupplemental As String" & vbNewLine
'  strBaseString = strBaseString & "  strSupplemental = ""Supplemental Information for each "" & _" & vbNewLine
'  strBaseString = strBaseString & "    ""of the contributor feature classes are listed below:"" & vbCrLf & vbCrLf" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  For lngPropSetIndex = 0 To pNewFClassArray.Count - 1" & vbNewLine
'  strBaseString = strBaseString & "    strFClassName = strFClassNames(lngPropSetIndex)" & vbNewLine
'  strBaseString = strBaseString & "    Set pSubArray = pNewFClassArray.Element(lngPropSetIndex)" & vbNewLine
'  strBaseString = strBaseString & "    Set pContributorPropSet = pSubArray.Element(1)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "'    Debug.Print ""Descriptions for '"" & strFClassName & ""'""" & vbNewLine
'  strBaseString = strBaseString & "    varAtts = pContributorPropSet.GetProperty(strSupplementalInformation)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "    If IsEmpty(varAtts) Then" & vbNewLine
'  strBaseString = strBaseString & "'      varAtts = pContributorPropSet.GetProperty(strUseLimitationsString2)" & vbNewLine
'  strBaseString = strBaseString & "'" & vbNewLine
'  strBaseString = strBaseString & "'      If IsEmpty(varAtts) Then" & vbNewLine
'  strBaseString = strBaseString & "        strSupplemental = strSupplemental & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Supplemental Information for "" & strFClassName & _" & vbNewLine
'  strBaseString = strBaseString & "            ""  [-- No Supplemental Information Available --]"" & vbCrLf & vbCrLf" & vbNewLine
'  strBaseString = strBaseString & "'      Else" & vbNewLine
'  strBaseString = strBaseString & "'        For lngIndex = 0 To UBound(varAtts)" & vbNewLine
'  strBaseString = strBaseString & "'          strValue = CStr(varAtts(lngIndex))" & vbNewLine
'  strBaseString = strBaseString & "'          strValue = RemoveTags(strValue)" & vbNewLine
'  strBaseString = strBaseString & "''          Debug.Print ""  "" & CStr(lngIndex) & ""] "" & strValue" & vbNewLine
'  strBaseString = strBaseString & "'" & vbNewLine
'  strBaseString = strBaseString & "'          If InStr(1, strValue, ""DIV STYLE"", vbTextCompare) > 0 Then" & vbNewLine
'  strBaseString = strBaseString & "'            strSupplemental = strSupplemental & _" & vbNewLine
'  strBaseString = strBaseString & "'                ""<P><SPAN STYLE=""""font-weight:bold;"""">Use Limitations for "" & strFClassName & _" & vbNewLine
'  strBaseString = strBaseString & "'                "":  </SPAN><SPAN> "" & strValue & ""</SPAN></P>""" & vbNewLine
'  strBaseString = strBaseString & "'          Else" & vbNewLine
'  strBaseString = strBaseString & "'" & vbNewLine
'  strBaseString = strBaseString & "'            strSupplemental = strSupplemental & _" & vbNewLine
'  strBaseString = strBaseString & "'                ""<P><SPAN STYLE=""""font-weight:bold;"""">Use Limitations for "" & strFClassName & _" & vbNewLine
'  strBaseString = strBaseString & "'                "":  </SPAN><SPAN> "" & strValue & ""</SPAN></P>""" & vbNewLine
'  strBaseString = strBaseString & "'          End If" & vbNewLine
'  strBaseString = strBaseString & "'" & vbNewLine
'  strBaseString = strBaseString & "'        Next lngIndex" & vbNewLine
'  strBaseString = strBaseString & "'      End If" & vbNewLine
'  strBaseString = strBaseString & "    Else" & vbNewLine
'  strBaseString = strBaseString & "      For lngIndex = 0 To UBound(varAtts)" & vbNewLine
'  strBaseString = strBaseString & "        strValue = CStr(varAtts(lngIndex))" & vbNewLine
'  strBaseString = strBaseString & "        strValue = RemoveTags(strValue)" & vbNewLine
'  strBaseString = strBaseString & "'        Debug.Print ""  "" & CStr(lngIndex) & ""] "" & strValue" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "        strSupplemental = strSupplemental & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Supplemental Information for "" & strFClassName & _" & vbNewLine
'  strBaseString = strBaseString & "            "":  "" & strValue & vbCrLf & vbCrLf" & vbNewLine
'  strBaseString = strBaseString & "      Next lngIndex" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "  Next lngPropSetIndex" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  strSupplemental = Replace(strSupplemental, ""TrasnportationLine"", ""TransportationLine"", , , vbTextCompare)" & vbNewLine
'  strBaseString = strBaseString & "'  strSupplemental = strSupplemental & ""</DIV></DIV>""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  varAtts = pPropSet.GetProperty(strSupplementalInformation)" & vbNewLine
'  strBaseString = strBaseString & "  If Not IsEmpty(varAtts) Then" & vbNewLine
'  strBaseString = strBaseString & "    pPropSet.RemoveProperty strSupplementalInformation" & vbNewLine
'  strBaseString = strBaseString & "  End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  varAtts = pPropSet.GetProperty(strSupplementalInformation)" & vbNewLine
'  strBaseString = strBaseString & "  If Not IsEmpty(varAtts) Then" & vbNewLine
'  strBaseString = strBaseString & "    If UBound(varAtts) = 0 Then" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strSupplementalInformation, strSupplemental" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "  Else" & vbNewLine
'  strBaseString = strBaseString & "    pPropSet.SetProperty strSupplementalInformation, strSupplemental" & vbNewLine
'  strBaseString = strBaseString & "  End If" & vbNewLine
'  strBaseString = strBaseString & "'" & vbNewLine
'  strBaseString = strBaseString & "'  Debug.Print ""Final Description:""" & vbNewLine
'  strBaseString = strBaseString & "'  Debug.Print strFullDescription" & vbNewLine
'  strBaseString = strBaseString & "  'pPropSet.SetProperty strSummaryString & strfinalsummary" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  ' ---------------------------------------------------------------------------------------------" & vbNewLine
'  strBaseString = strBaseString & "  ' USE LIMITATIONS" & vbNewLine
'  strBaseString = strBaseString & "  Dim strFullLimitations As String" & vbNewLine
'  strBaseString = strBaseString & "  strFullLimitations = ""<DIV STYLE=""""text-align:Left;""""><DIV><P><SPAN>Use Limitations for each "" & _" & vbNewLine
'  strBaseString = strBaseString & "    ""of the contributor feature classes are listed below:</SPAN></P>""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  For lngPropSetIndex = 0 To pNewFClassArray.Count - 1" & vbNewLine
'  strBaseString = strBaseString & "    strFClassName = strFClassNames(lngPropSetIndex)" & vbNewLine
'  strBaseString = strBaseString & "    Set pSubArray = pNewFClassArray.Element(lngPropSetIndex)" & vbNewLine
'  strBaseString = strBaseString & "    Set pContributorPropSet = pSubArray.Element(1)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "'    Debug.Print ""Descriptions for '"" & strFClassName & ""'""" & vbNewLine
'  strBaseString = strBaseString & "    varAtts = pContributorPropSet.GetProperty(strUseLimitationsString)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "    If IsEmpty(varAtts) Then" & vbNewLine
'  strBaseString = strBaseString & "      varAtts = pContributorPropSet.GetProperty(strUseLimitationsString2)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varAtts) Then" & vbNewLine
'  strBaseString = strBaseString & "        strFullLimitations = strFullLimitations & _" & vbNewLine
'  strBaseString = strBaseString & "            ""<P><SPAN STYLE=""""font-weight:bold;"""">Use Limitations for "" & strFClassName & _" & vbNewLine
'  strBaseString = strBaseString & "            "":</SPAN><SPAN> &lt;-- No Use Limitations Available --&gt;</SPAN></P>"" & vbCrLf" & vbNewLine
'  strBaseString = strBaseString & "      Else" & vbNewLine
'  strBaseString = strBaseString & "        For lngIndex = 0 To UBound(varAtts)" & vbNewLine
'  strBaseString = strBaseString & "          strValue = CStr(varAtts(lngIndex))" & vbNewLine
'  strBaseString = strBaseString & "          strValue = RemoveTags(strValue)" & vbNewLine
'  strBaseString = strBaseString & "'          Debug.Print ""  "" & CStr(lngIndex) & ""] "" & strValue" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "          If InStr(1, strValue, ""DIV STYLE"", vbTextCompare) > 0 Then" & vbNewLine
'  strBaseString = strBaseString & "            strFullLimitations = strFullLimitations & _" & vbNewLine
'  strBaseString = strBaseString & "                ""<P><SPAN STYLE=""""font-weight:bold;"""">Use Limitations for "" & strFClassName & _" & vbNewLine
'  strBaseString = strBaseString & "                "":  </SPAN><SPAN> "" & strValue & ""</SPAN></P>"" & vbCrLf" & vbNewLine
'  strBaseString = strBaseString & "          Else" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "            strFullLimitations = strFullLimitations & _" & vbNewLine
'  strBaseString = strBaseString & "                ""<P><SPAN STYLE=""""font-weight:bold;"""">Use Limitations for "" & strFClassName & _" & vbNewLine
'  strBaseString = strBaseString & "                "":  </SPAN><SPAN> "" & strValue & ""</SPAN></P>"" & vbCrLf" & vbNewLine
'  strBaseString = strBaseString & "          End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "        Next lngIndex" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    Else" & vbNewLine
'  strBaseString = strBaseString & "      For lngIndex = 0 To UBound(varAtts)" & vbNewLine
'  strBaseString = strBaseString & "        strValue = CStr(varAtts(lngIndex))" & vbNewLine
'  strBaseString = strBaseString & "        strValue = RemoveTags(strValue)" & vbNewLine
'  strBaseString = strBaseString & "'        Debug.Print ""  "" & CStr(lngIndex) & ""] "" & strValue" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "        strFullLimitations = strFullLimitations & _" & vbNewLine
'  strBaseString = strBaseString & "            ""<P><SPAN STYLE=""""font-weight:bold;"""">Use Limitations for "" & strFClassName & _" & vbNewLine
'  strBaseString = strBaseString & "            "":  </SPAN><SPAN> "" & strValue & ""</SPAN></P>"" & vbCrLf" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      Next lngIndex" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "  Next lngPropSetIndex" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  strFullLimitations = Replace(strFullLimitations, ""TrasnportationLine"", ""TransportationLine"", , , vbTextCompare)" & vbNewLine
'  strBaseString = strBaseString & "  strFullLimitations = strFullLimitations & ""</DIV></DIV>""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  varAtts = pPropSet.GetProperty(strUseLimitationsString)" & vbNewLine
'  strBaseString = strBaseString & "  If Not IsEmpty(varAtts) Then" & vbNewLine
'  strBaseString = strBaseString & "    pPropSet.RemoveProperty strUseLimitationsString" & vbNewLine
'  strBaseString = strBaseString & "  End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  varAtts = pPropSet.GetProperty(strUseLimitationsString)" & vbNewLine
'  strBaseString = strBaseString & "  If Not IsEmpty(varAtts) Then" & vbNewLine
'  strBaseString = strBaseString & "    If UBound(varAtts) = 0 Then" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strUseLimitationsString, strFullLimitations" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "  Else" & vbNewLine
'  strBaseString = strBaseString & "    pPropSet.SetProperty strUseLimitationsString, strFullLimitations" & vbNewLine
'  strBaseString = strBaseString & "  End If" & vbNewLine
'  strBaseString = strBaseString & "'" & vbNewLine
'  strBaseString = strBaseString & "'  Debug.Print ""Final Description:""" & vbNewLine
'  strBaseString = strBaseString & "'  Debug.Print strFullDescription" & vbNewLine
'  strBaseString = strBaseString & "  'pPropSet.SetProperty strSummaryString & strfinalsummary" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  ' ---------------------------------------------------------------------------------------------" & vbNewLine
'  strBaseString = strBaseString & "  ' SET CREDITS" & vbNewLine
'  strBaseString = strBaseString & "  Dim strFullCredits As String" & vbNewLine
'  strBaseString = strBaseString & "  strFullCredits = ""<DIV STYLE=""""text-align:Left;""""><DIV>""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  For lngPropSetIndex = 0 To pNewFClassArray.Count - 1" & vbNewLine
'  strBaseString = strBaseString & "    strFClassName = strFClassNames(lngPropSetIndex)" & vbNewLine
'  strBaseString = strBaseString & "    Set pSubArray = pNewFClassArray.Element(lngPropSetIndex)" & vbNewLine
'  strBaseString = strBaseString & "    Set pContributorPropSet = pSubArray.Element(1)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "'    Debug.Print ""Credits for '"" & strFClassName & ""'""" & vbNewLine
'  strBaseString = strBaseString & "    varAtts = pContributorPropSet.GetProperty(strCreditString)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "    If IsEmpty(varAtts) Then" & vbNewLine
'  strBaseString = strBaseString & "'      varAtts = pContributorPropSet.GetProperty(strAbstract)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "'      If IsEmpty(varAtts) Then" & vbNewLine
'  strBaseString = strBaseString & "        strFullCredits = strFullCredits & _" & vbNewLine
'  strBaseString = strBaseString & "            ""<P><SPAN STYLE=""""font-weight:bold;"""">Credits for "" & strFClassName & _" & vbNewLine
'  strBaseString = strBaseString & "            "":</SPAN><SPAN> [-- No Credits Available --].       </SPAN></P>""" & vbNewLine
'  strBaseString = strBaseString & "'      Else" & vbNewLine
'  strBaseString = strBaseString & "'        For lngIndex = 0 To UBound(varAtts)" & vbNewLine
'  strBaseString = strBaseString & "'          strValue = CStr(varAtts(lngIndex))" & vbNewLine
'  strBaseString = strBaseString & "'          strValue = RemoveTags(strValue)" & vbNewLine
'  strBaseString = strBaseString & "'          Debug.Print ""  "" & CStr(lngIndex) & ""] "" & strValue" & vbNewLine
'  strBaseString = strBaseString & "'" & vbNewLine
'  strBaseString = strBaseString & "'          If InStr(1, strValue, ""DIV STYLE"", vbTextCompare) > 0 Then" & vbNewLine
'  strBaseString = strBaseString & "'            strFullDescription = strFullDescription & _" & vbNewLine
'  strBaseString = strBaseString & "'                ""<P><SPAN STYLE=""""font-weight:bold;"""">Description for "" & strFClassName & _" & vbNewLine
'  strBaseString = strBaseString & "'                "":  </SPAN><SPAN> "" & strValue & ""</SPAN></P>""" & vbNewLine
'  strBaseString = strBaseString & "'          Else" & vbNewLine
'  strBaseString = strBaseString & "'" & vbNewLine
'  strBaseString = strBaseString & "'            strFullDescription = strFullDescription & _" & vbNewLine
'  strBaseString = strBaseString & "'                ""<P><SPAN STYLE=""""font-weight:bold;"""">Description for "" & strFClassName & _" & vbNewLine
'  strBaseString = strBaseString & "'                "":  </SPAN><SPAN> "" & strValue & ""</SPAN></P>""" & vbNewLine
'  strBaseString = strBaseString & "'          End If" & vbNewLine
'  strBaseString = strBaseString & "'" & vbNewLine
'  strBaseString = strBaseString & "'        Next lngIndex" & vbNewLine
'  strBaseString = strBaseString & "'      End If" & vbNewLine
'  strBaseString = strBaseString & "    Else" & vbNewLine
'  strBaseString = strBaseString & "      For lngIndex = 0 To UBound(varAtts)" & vbNewLine
'  strBaseString = strBaseString & "        strValue = CStr(varAtts(lngIndex))" & vbNewLine
'  strBaseString = strBaseString & "        strValue = RemoveTags(strValue)" & vbNewLine
'  strBaseString = strBaseString & "'        Debug.Print ""  "" & CStr(lngIndex) & ""] "" & strValue" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "        strFullCredits = strFullCredits & _" & vbNewLine
'  strBaseString = strBaseString & "            ""<P><SPAN STYLE=""""font-weight:bold;"""">Credits for "" & strFClassName & _" & vbNewLine
'  strBaseString = strBaseString & "            "":  </SPAN><SPAN> "" & strValue & "".       </SPAN></P>""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      Next lngIndex" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "  Next lngPropSetIndex" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  strFullCredits = Replace(strFullCredits, ""TrasnportationLine"", ""TransportationLine"", , , vbTextCompare)" & vbNewLine
'  strBaseString = strBaseString & "  strFullCredits = strFullCredits & ""</DIV></DIV>""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  varAtts = pPropSet.GetProperty(strCreditString)" & vbNewLine
'  strBaseString = strBaseString & "  If Not IsEmpty(varAtts) Then" & vbNewLine
'  strBaseString = strBaseString & "    pPropSet.RemoveProperty strCreditString" & vbNewLine
'  strBaseString = strBaseString & "  End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  varAtts = pPropSet.GetProperty(strCreditString)" & vbNewLine
'  strBaseString = strBaseString & "  If Not IsEmpty(varAtts) Then" & vbNewLine
'  strBaseString = strBaseString & "    If UBound(varAtts) = 0 Then" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strCreditString, strFullCredits" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "  Else" & vbNewLine
'  strBaseString = strBaseString & "    pPropSet.SetProperty strCreditString, strFullCredits" & vbNewLine
'  strBaseString = strBaseString & "  End If" & vbNewLine
'  strBaseString = strBaseString & "'" & vbNewLine
'  strBaseString = strBaseString & "'  Debug.Print ""Final Credits:""" & vbNewLine
'  strBaseString = strBaseString & "'  Debug.Print strFullCredits" & vbNewLine
'  strBaseString = strBaseString & "  'pPropSet.SetProperty strSummaryString & strfinalsummary" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "   " & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  ' ---------------------------------------------------------------------------------------------" & vbNewLine
'  strBaseString = strBaseString & "  ' SET DESCRIPTIONS" & vbNewLine
'  strBaseString = strBaseString & "  Dim strFullDescription As String" & vbNewLine
'  strBaseString = strBaseString & "  strFullDescription = ""<DIV STYLE=""""text-align:Left;""""><DIV><P><SPAN>This feature class "" & _" & vbNewLine
'  strBaseString = strBaseString & "    ""compiles transportation features and attributes from "" & _" & vbNewLine
'  strBaseString = strBaseString & "    ""multiple feature classes (""""TransportationLine_UTM12_NAD83"""""" & _" & vbNewLine
'  strBaseString = strBaseString & "    """"""Roads_KNF_UTM12_NAD83"""", """"BLM_AZStrip_UTM12_NAD83"""", """"Roads_UTM12_NAD83"""", "" & _" & vbNewLine
'  strBaseString = strBaseString & "    """"""UTAH_Roads_Near_AOI_UTM12_NAD83"""", """"Routes_UTM12_NAD83"""", """"Trails_UTM12_NAD83"""", "" & _" & vbNewLine
'  strBaseString = strBaseString & "    """"""TIGER_2012_Trans_Near_Park_UTM12_NAD83"""" and """"UTAH_Roads_Near_AOI_UTM12_NAD83"""") into a "" & _" & vbNewLine
'  strBaseString = strBaseString & "    ""single comprehensive transportation dataset covering the Grand Canyon region.  Descriptions "" & _" & vbNewLine
'  strBaseString = strBaseString & "    ""for these contributor feature classes are listed below:</SPAN></P>""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  For lngPropSetIndex = 0 To pNewFClassArray.Count - 1" & vbNewLine
'  strBaseString = strBaseString & "    strFClassName = strFClassNames(lngPropSetIndex)" & vbNewLine
'  strBaseString = strBaseString & "    Set pSubArray = pNewFClassArray.Element(lngPropSetIndex)" & vbNewLine
'  strBaseString = strBaseString & "    Set pContributorPropSet = pSubArray.Element(1)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "'    Debug.Print ""Descriptions for '"" & strFClassName & ""'""" & vbNewLine
'  strBaseString = strBaseString & "    varAtts = pContributorPropSet.GetProperty(strDescriptionString)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "    If IsEmpty(varAtts) Then" & vbNewLine
'  strBaseString = strBaseString & "      varAtts = pContributorPropSet.GetProperty(strAbstract)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varAtts) Then" & vbNewLine
'  strBaseString = strBaseString & "        strFullDescription = strFullDescription & _" & vbNewLine
'  strBaseString = strBaseString & "            ""<P><SPAN STYLE=""""font-weight:bold;"""">Description for "" & strFClassName & _" & vbNewLine
'  strBaseString = strBaseString & "            "":</SPAN><SPAN> &lt;-- No Description Available --&gt; </SPAN></P>""" & vbNewLine
'  strBaseString = strBaseString & "      Else" & vbNewLine
'  strBaseString = strBaseString & "        For lngIndex = 0 To UBound(varAtts)" & vbNewLine
'  strBaseString = strBaseString & "          strValue = CStr(varAtts(lngIndex))" & vbNewLine
'  strBaseString = strBaseString & "          strValue = RemoveTags(strValue)" & vbNewLine
'  strBaseString = strBaseString & "'          Debug.Print ""  "" & CStr(lngIndex) & ""] "" & strValue" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "          If InStr(1, strValue, ""DIV STYLE"", vbTextCompare) > 0 Then" & vbNewLine
'  strBaseString = strBaseString & "            strFullDescription = strFullDescription & _" & vbNewLine
'  strBaseString = strBaseString & "                ""<P><SPAN STYLE=""""font-weight:bold;"""">Description for "" & strFClassName & _" & vbNewLine
'  strBaseString = strBaseString & "                "":  </SPAN><SPAN> "" & strValue & ""</SPAN></P>""" & vbNewLine
'  strBaseString = strBaseString & "          Else" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "            strFullDescription = strFullDescription & _" & vbNewLine
'  strBaseString = strBaseString & "                ""<P><SPAN STYLE=""""font-weight:bold;"""">Description for "" & strFClassName & _" & vbNewLine
'  strBaseString = strBaseString & "                "":  </SPAN><SPAN> "" & strValue & ""</SPAN></P>""" & vbNewLine
'  strBaseString = strBaseString & "          End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "        Next lngIndex" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    Else" & vbNewLine
'  strBaseString = strBaseString & "      For lngIndex = 0 To UBound(varAtts)" & vbNewLine
'  strBaseString = strBaseString & "        strValue = CStr(varAtts(lngIndex))" & vbNewLine
'  strBaseString = strBaseString & "        strValue = RemoveTags(strValue)" & vbNewLine
'  strBaseString = strBaseString & "'        Debug.Print ""  "" & CStr(lngIndex) & ""] "" & strValue" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "        strFullDescription = strFullDescription & _" & vbNewLine
'  strBaseString = strBaseString & "            ""<P><SPAN STYLE=""""font-weight:bold;"""">Description for "" & strFClassName & _" & vbNewLine
'  strBaseString = strBaseString & "            "":  </SPAN><SPAN> "" & strValue & ""</SPAN></P>""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      Next lngIndex" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "  Next lngPropSetIndex" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  strFullDescription = Replace(strFullDescription, ""TrasnportationLine"", ""TransportationLine"", , , vbTextCompare)" & vbNewLine
'  strBaseString = strBaseString & "  strFullDescription = strFullDescription & ""</DIV></DIV>""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  varAtts = pPropSet.GetProperty(strDescriptionString)" & vbNewLine
'  strBaseString = strBaseString & "  If Not IsEmpty(varAtts) Then" & vbNewLine
'  strBaseString = strBaseString & "    pPropSet.RemoveProperty strDescriptionString" & vbNewLine
'  strBaseString = strBaseString & "  End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  varAtts = pPropSet.GetProperty(strDescriptionString)" & vbNewLine
'  strBaseString = strBaseString & "  If Not IsEmpty(varAtts) Then" & vbNewLine
'  strBaseString = strBaseString & "    If UBound(varAtts) = 0 Then" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strDescriptionString, strFullDescription" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "  Else" & vbNewLine
'  strBaseString = strBaseString & "    pPropSet.SetProperty strDescriptionString, strFullDescription" & vbNewLine
'  strBaseString = strBaseString & "  End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "'" & vbNewLine
'  strBaseString = strBaseString & "'  Debug.Print ""Final Description:""" & vbNewLine
'  strBaseString = strBaseString & "'  Debug.Print strFullDescription" & vbNewLine
'  strBaseString = strBaseString & "  'pPropSet.SetProperty strSummaryString & strfinalsummary" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  ' ---------------------------------------------------------------------------------------------" & vbNewLine
'  strBaseString = strBaseString & "   ' SET SUMMARIES" & vbNewLine
'  strBaseString = strBaseString & "  Dim pSummaries As esriSystem.IStringArray" & vbNewLine
'  strBaseString = strBaseString & "  Set pSummaries = New esriSystem.strArray" & vbNewLine
'  strBaseString = strBaseString & "  Dim strFullSummary As String" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  For lngPropSetIndex = 0 To pNewFClassArray.Count - 1" & vbNewLine
'  strBaseString = strBaseString & "    strFClassName = strFClassNames(lngPropSetIndex)" & vbNewLine
'  strBaseString = strBaseString & "    Set pSubArray = pNewFClassArray.Element(lngPropSetIndex)" & vbNewLine
'  strBaseString = strBaseString & "    Set pContributorPropSet = pSubArray.Element(1)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "'    Debug.Print ""Summaries for '"" & strFClassName & ""'""" & vbNewLine
'  strBaseString = strBaseString & "    varAtts = pContributorPropSet.GetProperty(strSummaryString)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "    If Not IsEmpty(varAtts) Then" & vbNewLine
'  strBaseString = strBaseString & "      For lngIndex = 0 To UBound(varAtts)" & vbNewLine
'  strBaseString = strBaseString & "        strValue = CStr(varAtts(lngIndex))" & vbNewLine
'  strBaseString = strBaseString & "'        Debug.Print ""  "" & CStr(lngIndex) & ""] "" & strValue" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "        strFullSummary = strFullSummary & ""<b>Summary for "" & strFClassName & "":</b>"" & _" & vbNewLine
'  strBaseString = strBaseString & "            vbLf & ""<p>"" & strValue & ""</p>"" & vbLf & vbLf" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      Next lngIndex" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "  Next lngPropSetIndex" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  varAtts = pPropSet.GetProperty(strSummaryString)" & vbNewLine
'  strBaseString = strBaseString & "  If Not IsEmpty(varAtts) Then" & vbNewLine
'  strBaseString = strBaseString & "    pPropSet.RemoveProperty strSummaryString" & vbNewLine
'  strBaseString = strBaseString & "  End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  ' ****** Summary can't be broken up into paragraphs!  Rewrite a single one from scratch..." & vbNewLine
'  strBaseString = strBaseString & "  strFullSummary = ""<DIV STYLE=""""text-align:Left;""""><DIV><P><SPAN>"" & _" & vbNewLine
'  strBaseString = strBaseString & "    ""This feature class compiles transportation features and attributes from "" & _" & vbNewLine
'  strBaseString = strBaseString & "    ""multiple feature classes into a single comprehensive transportation dataset covering the "" & _" & vbNewLine
'  strBaseString = strBaseString & "    ""Grand Canyon region.  Features are extracted primarily from the Grand Canyon National Park "" & _" & vbNewLine
'  strBaseString = strBaseString & "    """"""TransportationLine_UTM12_NAD83"""" feature class, and supplemented with features from "" & _" & vbNewLine
'  strBaseString = strBaseString & "    """"""Roads_KNF_UTM12_NAD83"""", """"BLM_AZStrip_UTM12_NAD83"""", """"Roads_UTM12_NAD83"""", "" & _" & vbNewLine
'  strBaseString = strBaseString & "    """"""UTAH_Roads_Near_AOI_UTM12_NAD83"""", """"Routes_UTM12_NAD83"""", """"Trails_UTM12_NAD83"""", "" & _" & vbNewLine
'  strBaseString = strBaseString & "    """"""TIGER_2012_Trans_Near_Park_UTM12_NAD83"""" and """"UTAH_Roads_Near_AOI_UTM12_NAD83"""".</SPAN></P>"" & _" & vbNewLine
'  strBaseString = strBaseString & "    ""<P><SPAN>Additional "" & _" & vbNewLine
'  strBaseString = strBaseString & "    ""attributes describing the transportation feature name, type, surface type and use level "" & _" & vbNewLine
'  strBaseString = strBaseString & "    ""were extracted from these feature classes based on proximity to the centroid of each feature.  "" & _" & vbNewLine
'  strBaseString = strBaseString & "    ""See descriptions of each attribute field for more detailed definitions of the attributes and "" & _" & vbNewLine
'  strBaseString = strBaseString & "    ""how the values were determined..</SPAN></P>"" & _" & vbNewLine
'  strBaseString = strBaseString & "    ""<P><SPAN>Process Steps and Lineage detailed below are extracted from "" & _" & vbNewLine
'  strBaseString = strBaseString & "    """"""TransportationLine_UTM12_NAD83""""</SPAN></P></DIV></DIV>""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  strFullSummary = Replace(strFullSummary, ""TrasnportationLine"", ""TransportationLine"", , , vbTextCompare)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  varAtts = pPropSet.GetProperty(strSummaryString)" & vbNewLine
'  strBaseString = strBaseString & "  If Not IsEmpty(varAtts) Then" & vbNewLine
'  strBaseString = strBaseString & "    If UBound(varAtts) = 0 Then" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strSummaryString, strFullSummary" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "  Else" & vbNewLine
'  strBaseString = strBaseString & "    pPropSet.SetProperty strSummaryString, strFullSummary" & vbNewLine
'  strBaseString = strBaseString & "  End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "'  Debug.Print ""Final Summary:""" & vbNewLine
'  strBaseString = strBaseString & "'  Debug.Print strFullSummary" & vbNewLine
'  strBaseString = strBaseString & "  'pPropSet.SetProperty strSummaryString & strfinalsummary" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  ' SET KEYWORDS" & vbNewLine
'  strBaseString = strBaseString & "  Dim pKeyWords As esriSystem.IStringArray" & vbNewLine
'  strBaseString = strBaseString & "  Set pKeyWords = New esriSystem.strArray" & vbNewLine
'  strBaseString = strBaseString & "  Dim pKeyWordColl As Collection" & vbNewLine
'  strBaseString = strBaseString & "  Set pKeyWordColl = New Collection" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  For lngPropSetIndex = 0 To pNewFClassArray.Count - 1" & vbNewLine
'  strBaseString = strBaseString & "    strFClassName = strFClassNames(lngPropSetIndex)" & vbNewLine
'  strBaseString = strBaseString & "    Set pSubArray = pNewFClassArray.Element(lngPropSetIndex)" & vbNewLine
'  strBaseString = strBaseString & "    Set pContributorPropSet = pSubArray.Element(1)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "    If Not VDH_Football.MyGeneralOperations.CheckCollectionForKey(pKeyWordColl, strFClassName) Then" & vbNewLine
'  strBaseString = strBaseString & "      pKeyWords.Add strFClassName" & vbNewLine
'  strBaseString = strBaseString & "      pKeyWordColl.Add True, strFClassName" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "'    Debug.Print ""Keywords for '"" & strFClassName & ""'""" & vbNewLine
'  strBaseString = strBaseString & "    varAtts = pContributorPropSet.GetProperty(strThemeKeywords)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "    If Not IsEmpty(varAtts) Then" & vbNewLine
'  strBaseString = strBaseString & "      For lngIndex = 0 To UBound(varAtts)" & vbNewLine
'  strBaseString = strBaseString & "        strValue = CStr(varAtts(lngIndex))" & vbNewLine
'  strBaseString = strBaseString & "'        Debug.Print ""  "" & CStr(lngIndex) & ""] "" & strValue" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "        If Not VDH_Football.MyGeneralOperations.CheckCollectionForKey(pKeyWordColl, strValue) Then" & vbNewLine
'  strBaseString = strBaseString & "          pKeyWords.Add strValue" & vbNewLine
'  strBaseString = strBaseString & "          pKeyWordColl.Add True, strValue" & vbNewLine
'  strBaseString = strBaseString & "        End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      Next lngIndex" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "  Next lngPropSetIndex" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  varAtts = pPropSet.GetProperty(strThemeKeywords)" & vbNewLine
'  strBaseString = strBaseString & "  If Not IsEmpty(varAtts) Then" & vbNewLine
'  strBaseString = strBaseString & "    pPropSet.RemoveProperty strThemeKeywords" & vbNewLine
'  strBaseString = strBaseString & "    pPropSet.RemoveProperty strSearchKeywords" & vbNewLine
'  strBaseString = strBaseString & "  End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "'  Debug.Print ""Final Keywords:""" & vbNewLine
'  strBaseString = strBaseString & "  For lngIndex = 0 To pKeyWords.Count - 1" & vbNewLine
'  strBaseString = strBaseString & "    strValue = pKeyWords.Element(lngIndex)" & vbNewLine
'  strBaseString = strBaseString & "    pPropSet.SetProperty strThemeKeywords & ""["" & CStr(lngIndex) & ""]"", strValue" & vbNewLine
'  strBaseString = strBaseString & "    pPropSet.SetProperty strSearchKeywords & ""["" & CStr(lngIndex) & ""]"", strValue" & vbNewLine
'  strBaseString = strBaseString & "'    Debug.Print ""  "" & CStr(lngIndex) & ""] "" & strValue" & vbNewLine
'  strBaseString = strBaseString & "  Next lngIndex" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "'" & vbNewLine
'  strBaseString = strBaseString & "'  ' CHECK ABSTRACTS" & vbNewLine
'  strBaseString = strBaseString & "'  Debug.Print ""Abstracts...........""" & vbNewLine
'  strBaseString = strBaseString & "'  For lngPropSetIndex = 0 To pNewFClassArray.Count - 1" & vbNewLine
'  strBaseString = strBaseString & "'    strFClassName = strFClassNames(lngPropSetIndex)" & vbNewLine
'  strBaseString = strBaseString & "'    Set pSubArray = pNewFClassArray.Element(lngPropSetIndex)" & vbNewLine
'  strBaseString = strBaseString & "'    Set pContributorPropSet = pSubArray.Element(1)" & vbNewLine
'  strBaseString = strBaseString & "'" & vbNewLine
'  strBaseString = strBaseString & "'    Debug.Print ""Abstracts for '"" & strFClassName & ""'""" & vbNewLine
'  strBaseString = strBaseString & "'    varAtts = pContributorPropSet.GetProperty(strAbstract)" & vbNewLine
'  strBaseString = strBaseString & "'" & vbNewLine
'  strBaseString = strBaseString & "'    If IsEmpty(varAtts) Then" & vbNewLine
'  strBaseString = strBaseString & "'      Debug.Print ""..No Abstract""" & vbNewLine
'  strBaseString = strBaseString & "'    Else" & vbNewLine
'  strBaseString = strBaseString & "'      For lngIndex = 0 To UBound(varAtts)" & vbNewLine
'  strBaseString = strBaseString & "'        strValue = CStr(varAtts(lngIndex))" & vbNewLine
'  strBaseString = strBaseString & "'        Debug.Print ""..Abstract =  "" & CStr(lngIndex) & ""] "" & strValue" & vbNewLine
'  strBaseString = strBaseString & "'" & vbNewLine
'  strBaseString = strBaseString & "'      Next lngIndex" & vbNewLine
'  strBaseString = strBaseString & "'    End If" & vbNewLine
'  strBaseString = strBaseString & "'  Next lngPropSetIndex" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "ClearMemory:" & vbNewLine
'  strBaseString = strBaseString & "  Set pRoads_KNF_UTM12_NAD83_FClass = Nothing" & vbNewLine
'  strBaseString = strBaseString & "  Set pTIGER_2012_Trans_Near_Park_UTM12_NAD83_FClass = Nothing" & vbNewLine
'  strBaseString = strBaseString & "  Set pBLM_AZStrip_UTM12_NAD83_FClass = Nothing" & vbNewLine
'  strBaseString = strBaseString & "  Set pRoads_UTM12_NAD83_FClass = Nothing" & vbNewLine
'  strBaseString = strBaseString & "  Set pTrails_UTM12_NAD83_FClass = Nothing" & vbNewLine
'  strBaseString = strBaseString & "  Set pRoutes_UTM12_NAD83_FClass = Nothing" & vbNewLine
'  strBaseString = strBaseString & "  Set pUtah_Roads_Near_AOI_UTM12_NAD83_FClass = Nothing" & vbNewLine
'  strBaseString = strBaseString & "  Set pTrails_KNF_UTM12_NAD83_FClass = Nothing" & vbNewLine
'  strBaseString = strBaseString & "  Set pTransportationLine_UTM12_NAD83_FClass = Nothing" & vbNewLine
'  strBaseString = strBaseString & "  Set pRoads_KNF_UTM12_NAD83_PropSet = Nothing" & vbNewLine
'  strBaseString = strBaseString & "  Set pTIGER_2012_Trans_Near_Park_UTM12_NAD83_PropSet = Nothing" & vbNewLine
'  strBaseString = strBaseString & "  Set pBLM_AZStrip_UTM12_NAD83_PropSet = Nothing" & vbNewLine
'  strBaseString = strBaseString & "  Set pRoads_UTM12_NAD83_PropSet = Nothing" & vbNewLine
'  strBaseString = strBaseString & "  Set pTrails_UTM12_NAD83_PropSet = Nothing" & vbNewLine
'  strBaseString = strBaseString & "  Set pRoutes_UTM12_NAD83_PropSet = Nothing" & vbNewLine
'  strBaseString = strBaseString & "  Set pUtah_Roads_Near_AOI_UTM12_NAD83_PropSet = Nothing" & vbNewLine
'  strBaseString = strBaseString & "  Set pTrails_KNF_UTM12_NAD83_PropSet = Nothing" & vbNewLine
'  strBaseString = strBaseString & "  Set pTransportationLine_UTM12_NAD83_PropSet = Nothing" & vbNewLine
'  strBaseString = strBaseString & "  Set pXMLPropSet = Nothing" & vbNewLine
'  strBaseString = strBaseString & "  Set pSubArray = Nothing" & vbNewLine
'  strBaseString = strBaseString & "  Set pNewFClassArray = Nothing" & vbNewLine
'  strBaseString = strBaseString & "  Set pContributorPropSet = Nothing" & vbNewLine
'  strBaseString = strBaseString & "  Erase strFClassNames" & vbNewLine
'  strBaseString = strBaseString & "  Set pXMLPropSet = Nothing" & vbNewLine
'  strBaseString = strBaseString & "  Set pSummaries = Nothing" & vbNewLine
'  strBaseString = strBaseString & "  Set pKeyWords = Nothing" & vbNewLine
'  strBaseString = strBaseString & "  Set pKeyWordColl = Nothing" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "End Sub" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "Public Sub TestRemoveTags()" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  Dim strtest As String" & vbNewLine
'  strBaseString = strBaseString & "  strtest = ""Hello <> Dolly<>test""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "'  strTest = """"" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  Debug.Print RemoveTags(strtest)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "End Sub" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "Public Function RemoveTags(strString) As String" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  Dim lngIndex As Long" & vbNewLine
'  strBaseString = strBaseString & "  Dim lngEndIndex As Long" & vbNewLine
'  strBaseString = strBaseString & "  RemoveTags = strString" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  lngIndex = InStr(1, RemoveTags, ""<"")" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  Do Until lngIndex = 0" & vbNewLine
'  strBaseString = strBaseString & "    lngEndIndex = InStr(lngIndex, RemoveTags, "">"")" & vbNewLine
'  strBaseString = strBaseString & "    RemoveTags = Left(RemoveTags, lngIndex - 1) & Right(RemoveTags, Len(RemoveTags) - lngEndIndex)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "    lngIndex = InStr(1, RemoveTags, ""<"")" & vbNewLine
'  strBaseString = strBaseString & "  Loop" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "End Function" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "Public Sub ReplaceOrAddDescription(pPropSet As IPropertySet, lngFieldIndex As Long, _" & vbNewLine
'  strBaseString = strBaseString & "    pFClassArray As esriSystem.IArray)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "'  Debug.Print ""Set definition sources for all fields!""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  Dim pRoads_KNF_UTM12_NAD83_FClass As IFeatureClass" & vbNewLine
'  strBaseString = strBaseString & "  Dim pTIGER_2012_Trans_Near_Park_UTM12_NAD83_FClass As IFeatureClass" & vbNewLine
'  strBaseString = strBaseString & "  Dim pBLM_AZStrip_UTM12_NAD83_FClass As IFeatureClass" & vbNewLine
'  strBaseString = strBaseString & "  Dim pRoads_UTM12_NAD83_FClass As IFeatureClass" & vbNewLine
'  strBaseString = strBaseString & "  Dim pTrails_UTM12_NAD83_FClass As IFeatureClass" & vbNewLine
'  strBaseString = strBaseString & "  Dim pRoutes_UTM12_NAD83_FClass As IFeatureClass" & vbNewLine
'  strBaseString = strBaseString & "  Dim pUtah_Roads_Near_AOI_UTM12_NAD83_FClass As IFeatureClass" & vbNewLine
'  strBaseString = strBaseString & "  Dim pTrails_KNF_UTM12_NAD83_FClass As IFeatureClass" & vbNewLine
'  strBaseString = strBaseString & "  Dim pTransportationLine_UTM12_NAD83_FClass As IFeatureClass" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  Dim pRoads_KNF_UTM12_NAD83_PropSet As IPropertySet" & vbNewLine
'  strBaseString = strBaseString & "  Dim pTIGER_2012_Trans_Near_Park_UTM12_NAD83_PropSet As IPropertySet" & vbNewLine
'  strBaseString = strBaseString & "  Dim pBLM_AZStrip_UTM12_NAD83_PropSet As IPropertySet" & vbNewLine
'  strBaseString = strBaseString & "  Dim pRoads_UTM12_NAD83_PropSet As IPropertySet" & vbNewLine
'  strBaseString = strBaseString & "  Dim pTrails_UTM12_NAD83_PropSet As IPropertySet" & vbNewLine
'  strBaseString = strBaseString & "  Dim pRoutes_UTM12_NAD83_PropSet As IPropertySet" & vbNewLine
'  strBaseString = strBaseString & "  Dim pUtah_Roads_Near_AOI_UTM12_NAD83_PropSet As IPropertySet" & vbNewLine
'  strBaseString = strBaseString & "  Dim pTrails_KNF_UTM12_NAD83_PropSet As IPropertySet" & vbNewLine
'  strBaseString = strBaseString & "  Dim pTransportationLine_UTM12_NAD83_PropSet As IPropertySet" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  Dim pSubArray As esriSystem.IArray" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  Set pSubArray = pFClassArray.Element(0)" & vbNewLine
'  strBaseString = strBaseString & "  Set pRoads_KNF_UTM12_NAD83_FClass = pSubArray.Element(0)" & vbNewLine
'  strBaseString = strBaseString & "  Set pRoads_KNF_UTM12_NAD83_PropSet = pSubArray.Element(1)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  Set pSubArray = pFClassArray.Element(1)" & vbNewLine
'  strBaseString = strBaseString & "  Set pTIGER_2012_Trans_Near_Park_UTM12_NAD83_FClass = pSubArray.Element(0)" & vbNewLine
'  strBaseString = strBaseString & "  Set pTIGER_2012_Trans_Near_Park_UTM12_NAD83_PropSet = pSubArray.Element(1)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  Set pSubArray = pFClassArray.Element(2)" & vbNewLine
'  strBaseString = strBaseString & "  Set pBLM_AZStrip_UTM12_NAD83_FClass = pSubArray.Element(0)" & vbNewLine
'  strBaseString = strBaseString & "  Set pBLM_AZStrip_UTM12_NAD83_PropSet = pSubArray.Element(1)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  Set pSubArray = pFClassArray.Element(3)" & vbNewLine
'  strBaseString = strBaseString & "  Set pRoads_UTM12_NAD83_FClass = pSubArray.Element(0)" & vbNewLine
'  strBaseString = strBaseString & "  Set pRoads_UTM12_NAD83_PropSet = pSubArray.Element(1)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  Set pSubArray = pFClassArray.Element(4)" & vbNewLine
'  strBaseString = strBaseString & "  Set pTrails_UTM12_NAD83_FClass = pSubArray.Element(0)" & vbNewLine
'  strBaseString = strBaseString & "  Set pTrails_UTM12_NAD83_PropSet = pSubArray.Element(1)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  Set pSubArray = pFClassArray.Element(5)" & vbNewLine
'  strBaseString = strBaseString & "  Set pRoutes_UTM12_NAD83_FClass = pSubArray.Element(0)" & vbNewLine
'  strBaseString = strBaseString & "  Set pRoutes_UTM12_NAD83_PropSet = pSubArray.Element(1)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  Set pSubArray = pFClassArray.Element(6)" & vbNewLine
'  strBaseString = strBaseString & "  Set pUtah_Roads_Near_AOI_UTM12_NAD83_FClass = pSubArray.Element(0)" & vbNewLine
'  strBaseString = strBaseString & "  Set pUtah_Roads_Near_AOI_UTM12_NAD83_PropSet = pSubArray.Element(1)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  Set pSubArray = pFClassArray.Element(7)" & vbNewLine
'  strBaseString = strBaseString & "  Set pTrails_KNF_UTM12_NAD83_FClass = pSubArray.Element(0)" & vbNewLine
'  strBaseString = strBaseString & "  Set pTrails_KNF_UTM12_NAD83_PropSet = pSubArray.Element(1)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  Set pSubArray = pFClassArray.Element(8)" & vbNewLine
'  strBaseString = strBaseString & "  Set pTransportationLine_UTM12_NAD83_FClass = pSubArray.Element(0)" & vbNewLine
'  strBaseString = strBaseString & "  Set pTransportationLine_UTM12_NAD83_PropSet = pSubArray.Element(1)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  Dim lngIndex As Long" & vbNewLine
'  strBaseString = strBaseString & "  Dim lngIndex2 As Long" & vbNewLine
'  strBaseString = strBaseString & "  Dim varVals As Variant" & vbNewLine
'  strBaseString = strBaseString & "  Dim varSubVals As Variant" & vbNewLine
'  strBaseString = strBaseString & "  Dim strNames(0) As String" & vbNewLine
'  strBaseString = strBaseString & "  Dim strFieldName As String" & vbNewLine
'  strBaseString = strBaseString & "  Dim strDescription As String" & vbNewLine
'  strBaseString = strBaseString & "  Dim varProperty As Variant" & vbNewLine
'  strBaseString = strBaseString & "  Dim varSubProperty As Variant" & vbNewLine
'  strBaseString = strBaseString & "  Dim strXName As String" & vbNewLine
'  strBaseString = strBaseString & "  Dim strProps As String" & vbNewLine
'  strBaseString = strBaseString & "  Dim strNewXName As String" & vbNewLine
'  strBaseString = strBaseString & "  Dim strOrigFClassXName As String" & vbNewLine
'  strBaseString = strBaseString & "  Dim strOrigFieldName As String" & vbNewLine
'  strBaseString = strBaseString & "  Dim varCheckPropertyPresent As Variant" & vbNewLine
'  strBaseString = strBaseString & "  Dim varSubProp1 As Variant" & vbNewLine
'  strBaseString = strBaseString & "  Dim varSubProp2 As Variant" & vbNewLine
'  strBaseString = strBaseString & "  Dim lngIndex3 As Long" & vbNewLine
'  strBaseString = strBaseString & "  Dim lngIndex4 As Long" & vbNewLine
'  strBaseString = strBaseString & "  Dim strSubXName1 As String" & vbNewLine
'  strBaseString = strBaseString & "  Dim strSubXName2 As String" & vbNewLine
'  strBaseString = strBaseString & "  Dim strSubXName3 As String" & vbNewLine
'  strBaseString = strBaseString & "  Dim strDescriptionSource As String" & vbNewLine
'  strBaseString = strBaseString & "  Dim strValue As String" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  Dim strBLMObsArray() As String" & vbNewLine
'  strBaseString = strBaseString & "  ReDim strBLMObsArray(11, 1)" & vbNewLine
'  strBaseString = strBaseString & "  strBLMObsArray(0, 0) = ""BLM_OBS_USE1""" & vbNewLine
'  strBaseString = strBaseString & "  strBLMObsArray(0, 1) = ""OBS_USE1""" & vbNewLine
'  strBaseString = strBaseString & "  strBLMObsArray(1, 0) = ""BLM_OBS_USE2""" & vbNewLine
'  strBaseString = strBaseString & "  strBLMObsArray(1, 1) = ""OBS_USE2""" & vbNewLine
'  strBaseString = strBaseString & "  strBLMObsArray(2, 0) = ""BLM_OBS_USE3""" & vbNewLine
'  strBaseString = strBaseString & "  strBLMObsArray(2, 1) = ""OBS_USE3""" & vbNewLine
'  strBaseString = strBaseString & "  strBLMObsArray(3, 0) = ""BLM_Use_Level""" & vbNewLine
'  strBaseString = strBaseString & "  strBLMObsArray(3, 1) = ""USE_LEVEL""" & vbNewLine
'  strBaseString = strBaseString & "  strBLMObsArray(4, 0) = ""BLM_Surface_PR""" & vbNewLine
'  strBaseString = strBaseString & "  strBLMObsArray(4, 1) = ""SURFACE_PR""" & vbNewLine
'  strBaseString = strBaseString & "  strBLMObsArray(5, 0) = ""BLM_Road_No""" & vbNewLine
'  strBaseString = strBaseString & "  strBLMObsArray(5, 1) = ""ROAD_NO_""" & vbNewLine
'  strBaseString = strBaseString & "  strBLMObsArray(6, 0) = ""BLM_Road_Name""" & vbNewLine
'  strBaseString = strBaseString & "  strBLMObsArray(6, 1) = ""ROAD_NAME""" & vbNewLine
'  strBaseString = strBaseString & "  strBLMObsArray(7, 0) = ""BLM_Straying_P""" & vbNewLine
'  strBaseString = strBaseString & "  strBLMObsArray(7, 1) = ""STRAYING_P""" & vbNewLine
'  strBaseString = strBaseString & "  strBLMObsArray(8, 0) = ""BLM_Purpose""" & vbNewLine
'  strBaseString = strBaseString & "  strBLMObsArray(8, 1) = ""PURPOSE""" & vbNewLine
'  strBaseString = strBaseString & "  strBLMObsArray(9, 0) = ""BLM_Route_ID""" & vbNewLine
'  strBaseString = strBaseString & "  strBLMObsArray(9, 1) = ""ROUTE_ID""" & vbNewLine
'  strBaseString = strBaseString & "  strBLMObsArray(10, 0) = ""BLM_RMP92""" & vbNewLine
'  strBaseString = strBaseString & "  strBLMObsArray(10, 1) = ""RMP92""" & vbNewLine
'  strBaseString = strBaseString & "  strBLMObsArray(11, 0) = ""BLM_E_Map""" & vbNewLine
'  strBaseString = strBaseString & "  strBLMObsArray(11, 1) = ""E_MAP""" & vbNewLine
'  strBaseString = strBaseString & "  Dim lngBLMIndex As Long" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  Dim strKNF_RoadArray() As String" & vbNewLine
'  strBaseString = strBaseString & "  ReDim strKNF_RoadArray(6, 1)" & vbNewLine
'  strBaseString = strBaseString & "  strKNF_RoadArray(0, 0) = ""KNF_ObjectID""" & vbNewLine
'  strBaseString = strBaseString & "  strKNF_RoadArray(0, 1) = ""OBJECTID""" & vbNewLine
'  strBaseString = strBaseString & "  strKNF_RoadArray(1, 0) = ""KNF_ID""" & vbNewLine
'  strBaseString = strBaseString & "  strKNF_RoadArray(1, 1) = ""ID""" & vbNewLine
'  strBaseString = strBaseString & "  strKNF_RoadArray(2, 0) = ""KNF_Name""" & vbNewLine
'  strBaseString = strBaseString & "  strKNF_RoadArray(2, 1) = ""NAME""" & vbNewLine
'  strBaseString = strBaseString & "  strKNF_RoadArray(3, 0) = ""KNF_Lanes""" & vbNewLine
'  strBaseString = strBaseString & "  strKNF_RoadArray(3, 1) = ""LANES""" & vbNewLine
'  strBaseString = strBaseString & "  strKNF_RoadArray(4, 0) = ""KNF_Surface_Type""" & vbNewLine
'  strBaseString = strBaseString & "  strKNF_RoadArray(4, 1) = ""SURFACE_TY""" & vbNewLine
'  strBaseString = strBaseString & "  strKNF_RoadArray(5, 0) = ""KNF_Oper_Maint""" & vbNewLine
'  strBaseString = strBaseString & "  strKNF_RoadArray(5, 1) = ""OPER_MAINT""" & vbNewLine
'  strBaseString = strBaseString & "  strKNF_RoadArray(6, 0) = ""KNF_Objective""" & vbNewLine
'  strBaseString = strBaseString & "  strKNF_RoadArray(6, 1) = ""OBJECTIVE_""" & vbNewLine
'  strBaseString = strBaseString & "  Dim lngKNFIndex As Long" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  Dim strTIG_RoadArray() As String" & vbNewLine
'  strBaseString = strBaseString & "  ReDim strTIG_RoadArray(2, 1)" & vbNewLine
'  strBaseString = strBaseString & "  strTIG_RoadArray(0, 0) = ""TIG_MTFCC""" & vbNewLine
'  strBaseString = strBaseString & "  strTIG_RoadArray(0, 1) = ""MTFCC""" & vbNewLine
'  strBaseString = strBaseString & "  strTIG_RoadArray(1, 0) = ""TIG_FULLNAME""" & vbNewLine
'  strBaseString = strBaseString & "  strTIG_RoadArray(1, 1) = ""FULLNAME""" & vbNewLine
'  strBaseString = strBaseString & "  strTIG_RoadArray(2, 0) = ""TIG_MT_Fclass""" & vbNewLine
'  strBaseString = strBaseString & "  strTIG_RoadArray(2, 1) = ""MT_FClass""" & vbNewLine
'  strBaseString = strBaseString & "  Dim lngTIGIndex As Long" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  Dim strROADS_RoadArray() As String" & vbNewLine
'  strBaseString = strBaseString & "  ReDim strROADS_RoadArray(4, 1)" & vbNewLine
'  strBaseString = strBaseString & "  strROADS_RoadArray(0, 0) = ""ROADS_RoadName""" & vbNewLine
'  strBaseString = strBaseString & "  strROADS_RoadArray(0, 1) = ""RoadName""" & vbNewLine
'  strBaseString = strBaseString & "  strROADS_RoadArray(1, 0) = ""ROADS_FMSS_ID""" & vbNewLine
'  strBaseString = strBaseString & "  strROADS_RoadArray(1, 1) = ""FMSS_ID""" & vbNewLine
'  strBaseString = strBaseString & "  strROADS_RoadArray(2, 0) = ""ROADS_SurfaceMaterial""" & vbNewLine
'  strBaseString = strBaseString & "  strROADS_RoadArray(2, 1) = ""SurfaceMaterial""" & vbNewLine
'  strBaseString = strBaseString & "  strROADS_RoadArray(3, 0) = ""ROADS_Status""" & vbNewLine
'  strBaseString = strBaseString & "  strROADS_RoadArray(3, 1) = ""Status""" & vbNewLine
'  strBaseString = strBaseString & "  strROADS_RoadArray(4, 0) = ""ROADS_Compendium""" & vbNewLine
'  strBaseString = strBaseString & "  strROADS_RoadArray(4, 1) = ""Compendium""" & vbNewLine
'  strBaseString = strBaseString & "  Dim lngRoads_Index As Long" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  Dim strTrails_RoadArray() As String" & vbNewLine
'  strBaseString = strBaseString & "  ReDim strTrails_RoadArray(2, 1)" & vbNewLine
'  strBaseString = strBaseString & "  strTrails_RoadArray(0, 0) = ""TRAILS_TrailName""" & vbNewLine
'  strBaseString = strBaseString & "  strTrails_RoadArray(0, 1) = ""TrailName""" & vbNewLine
'  strBaseString = strBaseString & "  strTrails_RoadArray(1, 0) = ""TRAILS_FMSS_ID""" & vbNewLine
'  strBaseString = strBaseString & "  strTrails_RoadArray(1, 1) = ""FMSS_ID""" & vbNewLine
'  strBaseString = strBaseString & "  strTrails_RoadArray(2, 0) = ""TRAILS_SurfaceMaterial""" & vbNewLine
'  strBaseString = strBaseString & "  strTrails_RoadArray(2, 1) = ""SurfaceMaterial""" & vbNewLine
'  strBaseString = strBaseString & "  Dim lngTrails_Index As Long" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  Dim strRoutes_RoadArray() As String" & vbNewLine
'  strBaseString = strBaseString & "  ReDim strRoutes_RoadArray(2, 1)" & vbNewLine
'  strBaseString = strBaseString & "  strRoutes_RoadArray(0, 0) = ""ROUTES_RouteName""" & vbNewLine
'  strBaseString = strBaseString & "  strRoutes_RoadArray(0, 1) = ""RouteName""" & vbNewLine
'  strBaseString = strBaseString & "  Dim lngRoutes_Index As Long" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  Dim strUtah_RoadArray() As String" & vbNewLine
'  strBaseString = strBaseString & "  ReDim strUtah_RoadArray(3, 1)" & vbNewLine
'  strBaseString = strBaseString & "  strUtah_RoadArray(0, 0) = ""UTAH_FullName""" & vbNewLine
'  strBaseString = strBaseString & "  strUtah_RoadArray(0, 1) = ""FULLNAME""" & vbNewLine
'  strBaseString = strBaseString & "  strUtah_RoadArray(1, 0) = ""UTAH_Alias1""" & vbNewLine
'  strBaseString = strBaseString & "  strUtah_RoadArray(1, 1) = ""ALIAS1""" & vbNewLine
'  strBaseString = strBaseString & "  strUtah_RoadArray(2, 0) = ""UTAH_SurfType""" & vbNewLine
'  strBaseString = strBaseString & "  strUtah_RoadArray(2, 1) = ""SURFTYPE""" & vbNewLine
'  strBaseString = strBaseString & "  strUtah_RoadArray(3, 0) = ""UTAH_SurfWidth""" & vbNewLine
'  strBaseString = strBaseString & "  strUtah_RoadArray(3, 1) = ""SURFWIDTH""" & vbNewLine
'  strBaseString = strBaseString & "  Dim lngUtah_Index As Long" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  Dim strKNFTR_RoadArray() As String" & vbNewLine
'  strBaseString = strBaseString & "  ReDim strKNFTR_RoadArray(6, 1)" & vbNewLine
'  strBaseString = strBaseString & "  strKNFTR_RoadArray(0, 0) = ""KNFTR_ID""" & vbNewLine
'  strBaseString = strBaseString & "  strKNFTR_RoadArray(0, 1) = ""ID""" & vbNewLine
'  strBaseString = strBaseString & "  strKNFTR_RoadArray(1, 0) = ""KNFTR_Name""" & vbNewLine
'  strBaseString = strBaseString & "  strKNFTR_RoadArray(1, 1) = ""NAME""" & vbNewLine
'  strBaseString = strBaseString & "  strKNFTR_RoadArray(2, 0) = ""KNFTR_Trail_Type""" & vbNewLine
'  strBaseString = strBaseString & "  strKNFTR_RoadArray(2, 1) = ""TRAIL_TYPE""" & vbNewLine
'  strBaseString = strBaseString & "  strKNFTR_RoadArray(3, 0) = ""KNFTR_Designed_Use""" & vbNewLine
'  strBaseString = strBaseString & "  strKNFTR_RoadArray(3, 1) = ""DESIGNED_U""" & vbNewLine
'  strBaseString = strBaseString & "  strKNFTR_RoadArray(4, 0) = ""KNFTR_Trail_Class""" & vbNewLine
'  strBaseString = strBaseString & "  strKNFTR_RoadArray(4, 1) = ""TRAIL_CLAS""" & vbNewLine
'  strBaseString = strBaseString & "  strKNFTR_RoadArray(5, 0) = ""KNFTR_Trail_Surface""" & vbNewLine
'  strBaseString = strBaseString & "  strKNFTR_RoadArray(5, 1) = ""TRAIL_SURF""" & vbNewLine
'  strBaseString = strBaseString & "  Dim lngKNFTR_Index As Long" & vbNewLine
'  strBaseString = strBaseString & "  Dim strXNameSource As String" & vbNewLine
'  strBaseString = strBaseString & "  Dim strXDomMin As String" & vbNewLine
'  strBaseString = strBaseString & "  Dim strXDomMax As String" & vbNewLine
'  strBaseString = strBaseString & "  Dim strXDomUnits As String" & vbNewLine
'  strBaseString = strBaseString & "  Dim strXDom As String" & vbNewLine
'  strBaseString = strBaseString & "  Dim strXUDom As String" & vbNewLine
'  strBaseString = strBaseString & "  Dim varSource As Variant" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  ' GET FIELD NAME" & vbNewLine
'  strBaseString = strBaseString & "  strXName = ""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrlabl""" & vbNewLine
'  strBaseString = strBaseString & "  varProperty = pPropSet.GetProperty(strXName)" & vbNewLine
'  strBaseString = strBaseString & "  If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "    varSubProperty = varProperty(0)" & vbNewLine
'  strBaseString = strBaseString & "    If IsEmpty(varSubProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "      strFieldName = ""<-- No Field Name -->""" & vbNewLine
'  strBaseString = strBaseString & "    Else" & vbNewLine
'  strBaseString = strBaseString & "      strFieldName = CStr(varSubProperty)" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "  Else" & vbNewLine
'  strBaseString = strBaseString & "    strFieldName = ""<-- No Field Name -->""" & vbNewLine
'  strBaseString = strBaseString & "  End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  ' GET CURRENT DESCRIPTION" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  strXName = ""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdef""" & vbNewLine
'  strBaseString = strBaseString & "  strXNameSource = ""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdefs""" & vbNewLine
'  strBaseString = strBaseString & "  strXDomMin = ""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdomv/rdom/rdommin""" & vbNewLine
'  strBaseString = strBaseString & "  strXDomMax = ""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdomv/rdom/rdommax""" & vbNewLine
'  strBaseString = strBaseString & "  strXDomUnits = ""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdomv/rdom/attrunit""" & vbNewLine
'  strBaseString = strBaseString & "  strXDom = ""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdomv""" & vbNewLine
'  strBaseString = strBaseString & "  strXUDom = ""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdomv/udom""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  varProperty = pPropSet.GetProperty(strXName)" & vbNewLine
'  strBaseString = strBaseString & "  If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "    varSubProperty = varProperty(0)" & vbNewLine
'  strBaseString = strBaseString & "    If IsEmpty(varSubProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""<-- No Description -->""" & vbNewLine
'  strBaseString = strBaseString & "    Else" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = CStr(varSubProperty)" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "  Else" & vbNewLine
'  strBaseString = strBaseString & "    strDescription = ""<-- No Description -->""" & vbNewLine
'  strBaseString = strBaseString & "  End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  ' FIRST CHECK FOR ORIGINAL TRANSPORTATION LINE FIELDS" & vbNewLine
'  strBaseString = strBaseString & "  If pTransportationLine_UTM12_NAD83_FClass.FindField(strFieldName) > -1 Then" & vbNewLine
'  strBaseString = strBaseString & "    If InStr(1, strDescription, ""Imported from TransportationLine_UTM12_NAD83"", vbTextCompare) = 0 Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = strDescription & ""  [Imported from TransportationLine_UTM12_NAD83]""" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "    ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "    varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "    If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXNameSource, ""TransportationLine_UTM12_NAD83""" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  Else" & vbNewLine
'  strBaseString = strBaseString & "    ' OTHERWISE GET CORRECT FIELD ATTRIBUTE DESCRIPTIONS" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "    '-----------------------------------------------------------------------------------------------" & vbNewLine
'  strBaseString = strBaseString & "    ' <<<<<<<<<<<<  FINAL ATTRIBUTES  >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>" & vbNewLine
'  strBaseString = strBaseString & "    '-----------------------------------------------------------------------------------------------" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""Final_Use_Level"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = """"""Use Level"""" indicates the transportation quality and accessibility of this "" & _" & vbNewLine
'  strBaseString = strBaseString & "        ""feature, such as distinguishing between paved roads and 4-wheel drive tracks.  "" & _" & vbNewLine
'  strBaseString = strBaseString & "        """"""Final_Use_Level"""" is the most likely use level of this feature, based on 'use level' values from "" & _" & vbNewLine
'  strBaseString = strBaseString & "        ""BLM_AZStrip_UTM12_NAD83, Roads_UTM12_NAD83, Roads_KNF_UTM12_NAD83, UTAH_Roads_Near_AOI_UTM12_NAD83, "" & _" & vbNewLine
'  strBaseString = strBaseString & "        ""and TIGER_2012_Trans_Near_Park_UTM12_NAD83."" & vbCrLf & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "        ""The Final Use Level value is extracted from the 'Use Level' field of the closest feature that "" & _" & vbNewLine
'  strBaseString = strBaseString & "        ""does not have a Null or 'No Value' name value.  See descriptions of the attribute fields "" & _" & vbNewLine
'  strBaseString = strBaseString & "        """"""Use_Level_1"""", """"Use_Level_2"""", """"Use_Level_3"""" and """"Use_Level_4"""" for details."" & vbCrLf & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "        ""'Nearby Features' are defined as features from any of the potential contributor "" & _" & vbNewLine
'  strBaseString = strBaseString & "        ""feature classes that are <= 10 meters from the centroid of this feature."" & vbCrLf & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "        ""If no nearby features have non-Null 'Use Level' values, then this value will be Null.""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.RemoveProperty strXUDom" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXUDom, ""Values extracted from 'use level' values obtained from "" & _" & vbNewLine
'  strBaseString = strBaseString & "        """"""BLM_AZStrip_UTM12_NAD83"""", """"Roads_UTM12_NAD83"""", """"Roads_KNF_UTM12_NAD83"""", "" & _" & vbNewLine
'  strBaseString = strBaseString & "        """"""UTAH_Roads_Near_AOI_UTM12_NAD83"""", and """"TIGER_2012_Trans_Near_Park_UTM12_NAD83"""".""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""Final_Use_Level_Source"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""The name of the feature class used to extract the most likely Use Level "" & _" & vbNewLine
'  strBaseString = strBaseString & "        ""of this feature.  See the definition of the attribute field """"Final_Use_Level"""" above "" & _" & vbNewLine
'  strBaseString = strBaseString & "        ""for details."" & vbCrLf & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "        ""Possible values are "" & _" & vbNewLine
'  strBaseString = strBaseString & "        ""BLM_AZStrip_UTM12_NAD83, Roads_UTM12_NAD83, Roads_KNF_UTM12_NAD83, UTAH_Roads_Near_AOI_UTM12_NAD83, "" & _" & vbNewLine
'  strBaseString = strBaseString & "        ""or TIGER_2012_Trans_Near_Park_UTM12_NAD83.""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.RemoveProperty strXUDom" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXUDom, ""Values selected from Feature Class Names "" & _" & vbNewLine
'  strBaseString = strBaseString & "        """"""BLM_AZStrip_UTM12_NAD83"""", """"Roads_UTM12_NAD83"""", """"Roads_KNF_UTM12_NAD83"""", "" & _" & vbNewLine
'  strBaseString = strBaseString & "        """"""UTAH_Roads_Near_AOI_UTM12_NAD83"""", and """"TIGER_2012_Trans_Near_Park_UTM12_NAD83"""".""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""Final_Surface_Type"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = """"""Surface_Type"""" indicates the construction material of this feature, "" & _" & vbNewLine
'  strBaseString = strBaseString & "        ""such as distinguishing betweeen paved, gravel and dirt roads.  "" & _" & vbNewLine
'  strBaseString = strBaseString & "        ""The """"Final_Surface_Type"""" is the most likely surface type of this feature, based on "" & _" & vbNewLine
'  strBaseString = strBaseString & "        ""Surface Type values from Roads_UTM12_NAD83, BLM_AZStrip_UTM12_NAD83, "" & _" & vbNewLine
'  strBaseString = strBaseString & "        ""Roads_KNF_UTM12_NAD83, TIGER_2012_Trans_Near_Park_UTM12_NAD83, Utah_Roads_Near_AOI_UTM12_NAD83 and "" & _" & vbNewLine
'  strBaseString = strBaseString & "        ""TransportationLine_UTM12_NAD83."" & vbCrLf & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "        ""If the 'Surface_Type' value from Transportation_Line_UTM12_NAD83 is Null or 'No Value', then the "" & _" & vbNewLine
'  strBaseString = strBaseString & "        ""Final Surface Type value is extracted from the 'Type' field of the closest feature that "" & _" & vbNewLine
'  strBaseString = strBaseString & "        ""does not have a Null or 'No Value' Type value.  See descriptions of the attribute fields "" & _" & vbNewLine
'  strBaseString = strBaseString & "        """"""SurfaceType_1"""", """"SurfaceType_2"""", """"SurfaceType_3"""" and """"SurfaceType_4"""" for details."" & vbCrLf & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "        ""'Nearby Features' are defined as features from any of the potential contributor "" & _" & vbNewLine
'  strBaseString = strBaseString & "        ""feature classes that are <= 10 meters from the centroid of this feature."" & vbCrLf & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "        ""If no nearby features have non-Null 'Surface Type' values, then this value will be Null.""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.RemoveProperty strXUDom" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXUDom, ""Values extracted from 'use level' values obtained from "" & _" & vbNewLine
'  strBaseString = strBaseString & "        """"""BLM_AZStrip_UTM12_NAD83"""", """"Roads_UTM12_NAD83"""", """"Roads_KNF_UTM12_NAD83"""", "" & _" & vbNewLine
'  strBaseString = strBaseString & "        """"""UTAH_Roads_Near_AOI_UTM12_NAD83"""", """"TIGER_2012_Trans_Near_Park_UTM12_NAD83"""" and "" & _" & vbNewLine
'  strBaseString = strBaseString & "        """"""TransportationLine_UTM12_NAD83."""".""" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""Final_Surface_Type_Source"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""The name of the feature class used to extract the most likely Surface Type "" & _" & vbNewLine
'  strBaseString = strBaseString & "        ""of this feature.  See the definition of the attribute field """"Final_Surface_Type"""" above "" & _" & vbNewLine
'  strBaseString = strBaseString & "        ""for details."" & vbCrLf & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "        ""Possible values are Roads_UTM12_NAD83, BLM_AZStrip_UTM12_NAD83, "" & _" & vbNewLine
'  strBaseString = strBaseString & "        ""Roads_KNF_UTM12_NAD83, TIGER_2012_Trans_Near_Park_UTM12_NAD83, Utah_Roads_Near_AOI_UTM12_NAD83 and "" & _" & vbNewLine
'  strBaseString = strBaseString & "        ""TransportationLine_UTM12_NAD83.""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.RemoveProperty strXUDom" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXUDom, ""Values selected from Feature Class Names "" & _" & vbNewLine
'  strBaseString = strBaseString & "      """"""BLM_AZStrip_UTM12_NAD83"""", """"Roads_UTM12_NAD83"""", """"Roads_KNF_UTM12_NAD83"""", "" & _" & vbNewLine
'  strBaseString = strBaseString & "        """"""UTAH_Roads_Near_AOI_UTM12_NAD83"""", """"TIGER_2012_Trans_Near_Park_UTM12_NAD83"""" and "" & _" & vbNewLine
'  strBaseString = strBaseString & "        """"""TransportationLine_UTM12_NAD83."""".""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""Final_Type"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = """"""Type"""" indicates the primary use or intended purpose of this feature, "" & _" & vbNewLine
'  strBaseString = strBaseString & "        ""such as distinguishing betweeen roads and hiking trails."" & _" & vbNewLine
'  strBaseString = strBaseString & "        ""The """"Final_Type"""" is the most likely type of this feature, based on Type values from "" & _" & vbNewLine
'  strBaseString = strBaseString & "        ""Transportation_Line_UTM12_NAD83 or extracted from nearby features from "" & _" & vbNewLine
'  strBaseString = strBaseString & "        ""BLM_AZStrip_UTM12_NAD83, Roads_KNF_UTM12_NAD83, Trails_KNF_UTM12_NAD83, "" & _" & vbNewLine
'  strBaseString = strBaseString & "        ""and TIGER_2012_Trans_Near_Park_UTM12_NAD83."" & vbCrLf & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "        ""If the 'Type' value from Transportation_Line_UTM12_NAD83 is Null or 'No Value', then the "" & _" & vbNewLine
'  strBaseString = strBaseString & "        ""Final Type value is extracted from the 'Type' field of the closest feature that "" & _" & vbNewLine
'  strBaseString = strBaseString & "        ""does not have a Null or 'No Value' Type value.  See descriptions of the attribute fields "" & _" & vbNewLine
'  strBaseString = strBaseString & "        """"""Type_1"""", """"Type_2"""", """"Type_3"""" and """"Type_4"""" for details."" & vbCrLf & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "        ""'Nearby Features' are defined as features from any of the potential contributor "" & _" & vbNewLine
'  strBaseString = strBaseString & "        ""feature classes that are <= 10 meters from the centroid of this feature."" & vbCrLf & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "        ""If no nearby features have non-Null 'Type' values, then this value will be Null.""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.RemoveProperty strXUDom" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXUDom, ""Values extracted from 'use level' values obtained from "" & _" & vbNewLine
'  strBaseString = strBaseString & "        """"""BLM_AZStrip_UTM12_NAD83"""", """"Trails_KNF_UTM12_NAD83"""", """"Roads_KNF_UTM12_NAD83"""", "" & _" & vbNewLine
'  strBaseString = strBaseString & "        """"""Transportation_Line_UTM12_NAD83"""", and """"TIGER_2012_Trans_Near_Park_UTM12_NAD83"""".""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""Final_Type_Source"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""The name of the feature class used to extract the most likely Type "" & _" & vbNewLine
'  strBaseString = strBaseString & "        ""of this feature.  See the definition of the attribute field """"Final_Type"""" above "" & _" & vbNewLine
'  strBaseString = strBaseString & "        ""for details."" & vbCrLf & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "        ""Possible values are Transportation_Line_UTM12_NAD83, "" & _" & vbNewLine
'  strBaseString = strBaseString & "        ""BLM_AZStrip_UTM12_NAD83, Roads_KNF_UTM12_NAD83, Trails_KNF_UTM12_NAD83, "" & _" & vbNewLine
'  strBaseString = strBaseString & "        ""or TIGER_2012_Trans_Near_Park_UTM12_NAD83.""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.RemoveProperty strXUDom" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXUDom, ""Values selected from Feature Class Names "" & _" & vbNewLine
'  strBaseString = strBaseString & "        """"""BLM_AZStrip_UTM12_NAD83"""", """"Trails_KNF_UTM12_NAD83"""", """"Roads_KNF_UTM12_NAD83"""", "" & _" & vbNewLine
'  strBaseString = strBaseString & "        """"""Transportation_Line_UTM12_NAD83"""", and """"TIGER_2012_Trans_Near_Park_UTM12_NAD83"""".""" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""Final_Name"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""The most likely name of this feature, based on name values from "" & _" & vbNewLine
'  strBaseString = strBaseString & "        ""Transportation_Line_UTM12_NAD83 or extracted from nearby features from "" & _" & vbNewLine
'  strBaseString = strBaseString & "        ""BLM_AZStrip_UTM12_NAD83, Roads_UTM12_NAD83, Routes_UTM12_NAD83, Trails_UTM12_NAD83, "" & _" & vbNewLine
'  strBaseString = strBaseString & "        ""Roads_KNF_UTM12_NAD83, Trails_KNF_UTM12_NAD83, UTAH_Roads_Near_AOI_UTM12_NAD83, "" & _" & vbNewLine
'  strBaseString = strBaseString & "        ""and TIGER_2012_Trans_Near_Park_UTM12_NAD83."" & vbCrLf & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "        ""The Final Name value is extracted from the 'Name' field of the closest feature that "" & _" & vbNewLine
'  strBaseString = strBaseString & "        ""does not have a Null or 'No Value' name value.  See descriptions of the attribute fields "" & _" & vbNewLine
'  strBaseString = strBaseString & "        """"""Name_1"""", """"Name_2"""", """"Name_3"""" and """"Name_4"""" for details."" & vbCrLf & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "        ""'Nearby Features' are defined as features from any of the potential contributor "" & _" & vbNewLine
'  strBaseString = strBaseString & "        ""feature classes that are <= 10 meters from the centroid of this feature."" & vbCrLf & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "        ""If no nearby features have non-Null 'Name' values, then this value will be Null.""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.RemoveProperty strXUDom" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXUDom, ""Values extracted from 'use level' values obtained from "" & _" & vbNewLine
'  strBaseString = strBaseString & "        """"""BLM_AZStrip_UTM12_NAD83"""", """"Roads_UTM12_NAD83"""", """"Roads_KNF_UTM12_NAD83"""", "" & _" & vbNewLine
'  strBaseString = strBaseString & "        """"""Routes_UTM12_NAD83"""", """"Trails_UTM12_NAD83"""", """"Trails_KNF_UTM12_NAD83"""", "" & _" & vbNewLine
'  strBaseString = strBaseString & "        """"""UTAH_Roads_Near_AOI_UTM12_NAD83"""", """"TIGER_2012_Trans_Near_Park_UTM12_NAD83"""" and "" & _" & vbNewLine
'  strBaseString = strBaseString & "        """"""TransportationLine_UTM12_NAD83."""".""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""Final_Name_Source"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""The name of the feature class used to extract the most likely name "" & _" & vbNewLine
'  strBaseString = strBaseString & "        ""of this feature.  See the definition of the attribute field """"Final_Name"""" above "" & _" & vbNewLine
'  strBaseString = strBaseString & "        ""for details."" & vbCrLf & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "        ""Possible values are Transportation_Line_UTM12_NAD83, "" & _" & vbNewLine
'  strBaseString = strBaseString & "        ""BLM_AZStrip_UTM12_NAD83, Roads_UTM12_NAD83, Routes_UTM12_NAD83, Trails_UTM12_NAD83, "" & _" & vbNewLine
'  strBaseString = strBaseString & "        ""Roads_KNF_UTM12_NAD83, Trails_KNF_UTM12_NAD83, UTAH_Roads_Near_AOI_UTM12_NAD83, "" & _" & vbNewLine
'  strBaseString = strBaseString & "        ""or TIGER_2012_Trans_Near_Park_UTM12_NAD83.""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "   " & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.RemoveProperty strXUDom" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXUDom, ""Values selected from Feature Class Names "" & _" & vbNewLine
'  strBaseString = strBaseString & "        """"""BLM_AZStrip_UTM12_NAD83"""", """"Roads_UTM12_NAD83"""", """"Roads_KNF_UTM12_NAD83"""", "" & _" & vbNewLine
'  strBaseString = strBaseString & "        """"""Routes_UTM12_NAD83"""", """"Trails_UTM12_NAD83"""", """"Trails_KNF_UTM12_NAD83"""", "" & _" & vbNewLine
'  strBaseString = strBaseString & "        """"""UTAH_Roads_Near_AOI_UTM12_NAD83"""", """"TIGER_2012_Trans_Near_Park_UTM12_NAD83"""" and "" & _" & vbNewLine
'  strBaseString = strBaseString & "        """"""TransportationLine_UTM12_NAD83."""".""" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "    '-----------------------------------------------------------------------------------------------" & vbNewLine
'  strBaseString = strBaseString & "    ' <<<<<<<<<<<<  NAME  >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>" & vbNewLine
'  strBaseString = strBaseString & "    '-----------------------------------------------------------------------------------------------" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""Name_1"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""The Name or Identifying Code of this feature.  This attribute field contains "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""the Name value from the closest feature from 9 potential contributor feature "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""classes, where distance is defined as the distance from the centerpoint of this feature to the "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""nearest point on the nearby feature.  The maximum distance considered was 10 meters.  This value "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""should be Null if there were no features within 10m."" & vbCrLf & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""This attribute value was drawn from different attribute fields depending on the feature class:"" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from BLM_AZStrip_UTM12_NAD83, then the attribute field """"Road_Name""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"BLM_Road_Name"""" above for definitions."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from Roads_UTM12_NAD83, then the attribute field """"RoadName""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"ROADS_RoadName"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from Routes_UTM12_NAD83, then the attribute field """"RouteName""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"ROUTES_RouteName"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from Trails_UTM12_NAD83, then the attribute field """"TrailName""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"TRAILS_TrailName"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from Roads_KNF_UTM12_NAD83, then the attribute field """"NAME""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"KNF_Name"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from Trails_KNF_UTM12_NAD83, then the attribute field """"NAME""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"KNFTR_Name"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from UTAH_Roads_Near_AOI_UTM12_NAD83, then the attribute field """"FullName""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"UTAH_FullName"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from TIGER_2012_Trans_Near_Park_UTM12_NAD83, then the attribute field """"FULLNAME""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"TIG_FULLNAME"""" above for definition."" & vbCrLf" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = strDescription & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from TransportationLine_UTM12_NAD83, then the attribute fields """"RoadName"""", "" & _" & vbNewLine
'  strBaseString = strBaseString & "          """"""RouteName"""", and/or """"TrailName""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See descriptions of """"RoadName"""", """"RouteName"""" and """"TrailName"""" above for definitions."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          vbCrLf & ""If the nearest feature was from the TransportationLine_UTM12_NAD83 feature class, and if "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""all three attribute fields """"RoadName"""", """"RouteName"""" and """"TrailName"""" were empty or null, then "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""this feature was ignored and not considered as a candidate to assign a name to this feature.  "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""Features from any other nearby feature classes were included regardless of whether the Name values "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""were empty or not.""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""Name_1_Source"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""This is the name of the feature class containing the closest feature to this feature.  "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""The 9 potential contributer feature classes are Roads_KNF_UTM12_NAD83, BLM_AZStrip_UTM12_NAD83, "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""Roads_UTM12_NAD83, UTAH_Roads_Near_AOI_UTM12_NAD83, Routes_UTM12_NAD83, Trails_UTM12_NAD83, "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""TIGER_2012_Trans_Near_Park_UTM12_NAD83, UTAH_Roads_Near_AOI_UTM12_NAD83 and "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""TransportationLine_UTM12_NAD83.  "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""The maximum distance considered was 10 meters."" & vbCrLf & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""If this attribute value is extracted from the TransportationLine_UTM12_NAD83 feature class, then this "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""value is appended with either """":RoadName"""", """":RouteName"""" or """":TrailName"""" to indicate which "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""attribute field the attribute came from.""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""Name_1_Distance"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""This is the distance in meters from the centroid of this feature to any part of the "" & _" & vbNewLine
'  strBaseString = strBaseString & "           ""closest feature from one of 9 potential contributor feature classes, up to a maximum distance "" & _" & vbNewLine
'  strBaseString = strBaseString & "           ""of 10 meters.  "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""The 9 potential contributer feature classes are Roads_KNF_UTM12_NAD83, BLM_AZStrip_UTM12_NAD83, "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""Roads_UTM12_NAD83, UTAH_Roads_Near_AOI_UTM12_NAD83, Routes_UTM12_NAD83, Trails_UTM12_NAD83, "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""TIGER_2012_Trans_Near_Park_UTM12_NAD83, UTAH_Roads_Near_AOI_UTM12_NAD83 and "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""TransportationLine_UTM12_NAD83.""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.RemoveProperty strXDom" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomMin, ""0""" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomMax, ""10""" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomUnits, ""Meters""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""Name_2"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""The Name or Identifying Code of this feature.  This attribute field contains "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""the Name value from the 2nd-closest feature from 9 potential contributor feature "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""classes, where distance is defined as the distance from the centerpoint of this feature to the "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""nearest point on the nearby feature.  The maximum distance considered was 10 meters.  This value "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""should be Null if there was no 2nd-closest feature within 10m."" & vbCrLf & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""This attribute value was drawn from different attribute fields depending on the feature class:"" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from BLM_AZStrip_UTM12_NAD83, then the attribute field """"Road_Name""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"BLM_Road_Name"""" above for definitions."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from Roads_UTM12_NAD83, then the attribute field """"RoadName""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"ROADS_RoadName"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from Routes_UTM12_NAD83, then the attribute field """"RouteName""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"ROUTES_RouteName"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from Trails_UTM12_NAD83, then the attribute field """"TrailName""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"TRAILS_TrailName"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from Roads_KNF_UTM12_NAD83, then the attribute field """"NAME""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"KNF_Name"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from Trails_KNF_UTM12_NAD83, then the attribute field """"NAME""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"KNFTR_Name"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from UTAH_Roads_Near_AOI_UTM12_NAD83, then the attribute field """"FullName""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"UTAH_FullName"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from TIGER_2012_Trans_Near_Park_UTM12_NAD83, then the attribute field """"FULLNAME""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"TIG_FULLNAME"""" above for definition."" & vbCrLf" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = strDescription & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from TransportationLine_UTM12_NAD83, then the attribute fields """"RoadName"""", "" & _" & vbNewLine
'  strBaseString = strBaseString & "          """"""RouteName"""", and/or """"TrailName""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See descriptions of """"RoadName"""", """"RouteName"""" and """"TrailName"""" above for definitions."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          vbCrLf & ""If the 2nd-nearest feature was from the TransportationLine_UTM12_NAD83 feature class, and if "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""all three attribute fields """"RoadName"""", """"RouteName"""" and """"TrailName"""" were empty or null, then "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""this feature was ignored and not considered as a candidate to assign a name to this feature.  "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""Features from any other nearby feature classes were included regardless of whether the Name values "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""were empty or not.""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""Name_2_Source"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""This is the name of the feature class containing the 2nd-closest feature to this feature.  "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""The 9 potential contributer feature classes are Roads_KNF_UTM12_NAD83, BLM_AZStrip_UTM12_NAD83, "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""Roads_UTM12_NAD83, UTAH_Roads_Near_AOI_UTM12_NAD83, Routes_UTM12_NAD83, Trails_UTM12_NAD83, "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""TIGER_2012_Trans_Near_Park_UTM12_NAD83, UTAH_Roads_Near_AOI_UTM12_NAD83 and "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""TransportationLine_UTM12_NAD83.  "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""The maximum distance considered was 10 meters."" & vbCrLf & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""If this attribute value is extracted from the TransportationLine_UTM12_NAD83 feature class, then this "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""value is appended with either """":RoadName"""", """":RouteName"""" or """":TrailName"""" to indicate which "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""attribute field the attribute came from.""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""Name_2_Distance"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""This is the distance in meters from the centroid of this feature to any part of the "" & _" & vbNewLine
'  strBaseString = strBaseString & "           ""2nd-closest feature from one of 9 potential contributor feature classes, up to a maximum distance "" & _" & vbNewLine
'  strBaseString = strBaseString & "           ""of 10 meters.  "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""The 9 potential contributer feature classes are Roads_KNF_UTM12_NAD83, BLM_AZStrip_UTM12_NAD83, "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""Roads_UTM12_NAD83, UTAH_Roads_Near_AOI_UTM12_NAD83, Routes_UTM12_NAD83, Trails_UTM12_NAD83, "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""TIGER_2012_Trans_Near_Park_UTM12_NAD83, UTAH_Roads_Near_AOI_UTM12_NAD83 and "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""TransportationLine_UTM12_NAD83.""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.RemoveProperty strXDom" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomMin, ""0""" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomMax, ""10""" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomUnits, ""Meters""" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""Name_3"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""The Name or Identifying Code of this feature.  This attribute field contains "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""the Name value from the 3rd-closest feature from 9 potential contributor feature "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""classes, where distance is defined as the distance from the centerpoint of this feature to the "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""nearest point on the nearby feature.  The maximum distance considered was 10 meters.  This value "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""should be Null if there was no 3rd-closest feature within 10m."" & vbCrLf & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""This attribute value was drawn from different attribute fields depending on the feature class:"" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from BLM_AZStrip_UTM12_NAD83, then the attribute field """"Road_Name""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"BLM_Road_Name"""" above for definitions."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from Roads_UTM12_NAD83, then the attribute field """"RoadName""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"ROADS_RoadName"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from Routes_UTM12_NAD83, then the attribute field """"RouteName""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"ROUTES_RouteName"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from Trails_UTM12_NAD83, then the attribute field """"TrailName""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"TRAILS_TrailName"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from Roads_KNF_UTM12_NAD83, then the attribute field """"NAME""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"KNF_Name"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from Trails_KNF_UTM12_NAD83, then the attribute field """"NAME""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"KNFTR_Name"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from UTAH_Roads_Near_AOI_UTM12_NAD83, then the attribute field """"FullName""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"UTAH_FullName"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from TIGER_2012_Trans_Near_Park_UTM12_NAD83, then the attribute field """"FULLNAME""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"TIG_FULLNAME"""" above for definition."" & vbCrLf" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = strDescription & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from TransportationLine_UTM12_NAD83, then the attribute fields """"RoadName"""", "" & _" & vbNewLine
'  strBaseString = strBaseString & "          """"""RouteName"""", and/or """"TrailName""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See descriptions of """"RoadName"""", """"RouteName"""" and """"TrailName"""" above for definitions."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          vbCrLf & ""If the 3rd-nearest feature was from the TransportationLine_UTM12_NAD83 feature class, and if "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""all three attribute fields """"RoadName"""", """"RouteName"""" and """"TrailName"""" were empty or null, then "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""this feature was ignored and not considered as a candidate to assign a name to this feature.  "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""Features from any other nearby feature classes were included regardless of whether the Name values "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""were empty or not.""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""Name_3_Source"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""This is the name of the feature class containing the 3rd-closest feature to this feature.  "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""The 9 potential contributer feature classes are Roads_KNF_UTM12_NAD83, BLM_AZStrip_UTM12_NAD83, "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""Roads_UTM12_NAD83, UTAH_Roads_Near_AOI_UTM12_NAD83, Routes_UTM12_NAD83, Trails_UTM12_NAD83, "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""TIGER_2012_Trans_Near_Park_UTM12_NAD83, UTAH_Roads_Near_AOI_UTM12_NAD83 and "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""TransportationLine_UTM12_NAD83.  "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""The maximum distance considered was 10 meters."" & vbCrLf & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""If this attribute value is extracted from the TransportationLine_UTM12_NAD83 feature class, then this "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""value is appended with either """":RoadName"""", """":RouteName"""" or """":TrailName"""" to indicate which "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""attribute field the attribute came from.""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""Name_3_Distance"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""This is the distance in meters from the centroid of this feature to any part of the "" & _" & vbNewLine
'  strBaseString = strBaseString & "           ""3rd-closest feature from one of 9 potential contributor feature classes, up to a maximum distance "" & _" & vbNewLine
'  strBaseString = strBaseString & "           ""of 10 meters.  "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""The 9 potential contributer feature classes are Roads_KNF_UTM12_NAD83, BLM_AZStrip_UTM12_NAD83, "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""Roads_UTM12_NAD83, UTAH_Roads_Near_AOI_UTM12_NAD83, Routes_UTM12_NAD83, Trails_UTM12_NAD83, "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""TIGER_2012_Trans_Near_Park_UTM12_NAD83, UTAH_Roads_Near_AOI_UTM12_NAD83 and "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""TransportationLine_UTM12_NAD83.""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.RemoveProperty strXDom" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomMin, ""0""" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomMax, ""10""" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomUnits, ""Meters""" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""Name_4"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""The Name or Identifying Code of this feature.  This attribute field contains "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""the Name value from the 4th-closest feature from 9 potential contributor feature "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""classes, where distance is defined as the distance from the centerpoint of this feature to the "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""nearest point on the nearby feature.  The maximum distance considered was 10 meters.  This value "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""should be Null if there was no 3rd-closest feature within 10m."" & vbCrLf & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""This attribute value was drawn from different attribute fields depending on the feature class:"" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from BLM_AZStrip_UTM12_NAD83, then the attribute field """"Road_Name""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"BLM_Road_Name"""" above for definitions."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from Roads_UTM12_NAD83, then the attribute field """"RoadName""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"ROADS_RoadName"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from Routes_UTM12_NAD83, then the attribute field """"RouteName""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"ROUTES_RouteName"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from Trails_UTM12_NAD83, then the attribute field """"TrailName""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"TRAILS_TrailName"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from Roads_KNF_UTM12_NAD83, then the attribute field """"NAME""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"KNF_Name"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from Trails_KNF_UTM12_NAD83, then the attribute field """"NAME""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"KNFTR_Name"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from UTAH_Roads_Near_AOI_UTM12_NAD83, then the attribute field """"FullName""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"UTAH_FullName"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from TIGER_2012_Trans_Near_Park_UTM12_NAD83, then the attribute field """"FULLNAME""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"TIG_FULLNAME"""" above for definition."" & vbCrLf" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = strDescription & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from TransportationLine_UTM12_NAD83, then the attribute fields """"RoadName"""", "" & _" & vbNewLine
'  strBaseString = strBaseString & "          """"""RouteName"""", and/or """"TrailName""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See descriptions of """"RoadName"""", """"RouteName"""" and """"TrailName"""" above for definitions."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          vbCrLf & ""If the 4th-nearest feature was from the TransportationLine_UTM12_NAD83 feature class, and if "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""all three attribute fields """"RoadName"""", """"RouteName"""" and """"TrailName"""" were empty or null, then "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""this feature was ignored and not considered as a candidate to assign a name to this feature.  "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""Features from any other nearby feature classes were included regardless of whether the Name values "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""were empty or not.""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""Name_4_Source"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""This is the name of the feature class containing the 4th-closest feature to this feature.  "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""The 9 potential contributer feature classes are Roads_KNF_UTM12_NAD83, BLM_AZStrip_UTM12_NAD83, "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""Roads_UTM12_NAD83, UTAH_Roads_Near_AOI_UTM12_NAD83, Routes_UTM12_NAD83, Trails_UTM12_NAD83, "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""TIGER_2012_Trans_Near_Park_UTM12_NAD83, UTAH_Roads_Near_AOI_UTM12_NAD83 and "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""TransportationLine_UTM12_NAD83.  "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""The maximum distance considered was 10 meters."" & vbCrLf & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""If this attribute value is extracted from the TransportationLine_UTM12_NAD83 feature class, then this "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""value is appended with either """":RoadName"""", """":RouteName"""" or """":TrailName"""" to indicate which "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""attribute field the attribute came from.""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""Name_4_Distance"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""This is the distance in meters from the centroid of this feature to any part of the "" & _" & vbNewLine
'  strBaseString = strBaseString & "           ""4th-closest feature from one of 9 potential contributor feature classes, up to a maximum distance "" & _" & vbNewLine
'  strBaseString = strBaseString & "           ""of 10 meters.  "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""The 9 potential contributer feature classes are Roads_KNF_UTM12_NAD83, BLM_AZStrip_UTM12_NAD83, "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""Roads_UTM12_NAD83, UTAH_Roads_Near_AOI_UTM12_NAD83, Routes_UTM12_NAD83, Trails_UTM12_NAD83, "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""TIGER_2012_Trans_Near_Park_UTM12_NAD83, UTAH_Roads_Near_AOI_UTM12_NAD83 and "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""TransportationLine_UTM12_NAD83.""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.RemoveProperty strXDom" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomMin, ""0""" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomMax, ""10""" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomUnits, ""Meters""" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "    ' ----------------------------------------------------------------------------------------------" & vbNewLine
'  strBaseString = strBaseString & "    ' <<<<<<<<<<<<  Use Level  >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>" & vbNewLine
'  strBaseString = strBaseString & "    '-----------------------------------------------------------------------------------------------" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""Use_Level_1"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = """"""Use Level"""" indicates the transportation quality and accessibility of this feature, such as "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""distinguishing between paved roads and 4-wheel drive tracks.  This attribute field contains "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""the Use Level value from the closest feature from 5 potential contributor feature "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""classes, where distance is defined as the distance from the centerpoint of this feature to the "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""nearest point on the nearby feature.  The maximum distance considered was 10 meters.  This value "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""should be Null if there were no features within 10m."" & vbCrLf & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""This attribute value was drawn from different attribute fields depending on the feature class:"" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from BLM_AZStrip_UTM12_NAD83, then the attribute field """"E_Map""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"BLM_E_Map"""" above for definitions."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from Roads_UTM12_NAD83, then the attribute field """"Status""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"ROADS_Status"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from Roads_KNF_UTM12_NAD83, then the attribute field """"OPER_MAINT""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"KNF_Oper_Maint"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from UTAH_Roads_Near_AOI_UTM12_NAD83, then the attribute field """"SurfType""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"UTAH_SurfType"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from TIGER_2012_Trans_Near_Park_UTM12_NAD83, then the attribute field """"MT_FCLASS""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"TIG_MT_Fclass"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""The feature classes Trails_KNF_UTM12_NAD83, Trails_UTM12_NAD83, Routes_UTM12_NAD83 and "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""TransportationLine_UTM12_NAD83 did not have good attribute values to describe Use Level.""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""Use_Level_1_Source"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""This is the name of the feature class containing the closest feature to this feature.  "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""The 5 potential contributer feature classes are Roads_KNF_UTM12_NAD83, BLM_AZStrip_UTM12_NAD83, "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""Roads_UTM12_NAD83, UTAH_Roads_Near_AOI_UTM12_NAD83 and TIGER_2012_Trans_Near_Park_UTM12_NAD83.  "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""The maximum distance considered was 10 meters.""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""Use_Level_1_Distance"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""This is the distance in meters from the centroid of this feature to any part of the "" & _" & vbNewLine
'  strBaseString = strBaseString & "           ""closest feature from one of 5 potential contributor feature classes, up to a maximum distance "" & _" & vbNewLine
'  strBaseString = strBaseString & "           ""of 10 meters.  "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""The 5 potential contributer feature classes are Roads_KNF_UTM12_NAD83, BLM_AZStrip_UTM12_NAD83, "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""Roads_UTM12_NAD83, UTAH_Roads_Near_AOI_UTM12_NAD83 and TIGER_2012_Trans_Near_Park_UTM12_NAD83.""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.RemoveProperty strXDom" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomMin, ""0""" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomMax, ""10""" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomUnits, ""Meters""" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""Use_Level_2"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = """"""Use Level"""" indicates the transportation quality and accessibility of this feature, such as "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""distinguishing between paved roads and 4-wheel drive tracks.  This attribute field contains "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""the Use Level value from the 2nd-closest feature from 5 potential contributor feature "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""classes, where distance is defined as the distance from the centerpoint of this feature to the "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""nearest point on the nearby feature.  The maximum distance considered was 10 meters.  This value "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""should be Null if there was no 2nd-closest feature within 10m."" & vbCrLf & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""This attribute value was drawn from different attribute fields depending on the feature class:"" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from BLM_AZStrip_UTM12_NAD83, then the attribute field """"E_Map""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"BLM_E_Map"""" above for definitions."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from Roads_UTM12_NAD83, then the attribute field """"Status""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"ROADS_Status"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from Roads_KNF_UTM12_NAD83, then the attribute field """"OPER_MAINT""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"KNF_Oper_Maint"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from UTAH_Roads_Near_AOI_UTM12_NAD83, then the attribute field """"SurfType""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"UTAH_SurfType"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from TIGER_2012_Trans_Near_Park_UTM12_NAD83, then the attribute field """"MT_FCLASS""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"TIG_MT_Fclass"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""The feature classes Trails_KNF_UTM12_NAD83, Trails_UTM12_NAD83, Routes_UTM12_NAD83 and "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""TransportationLine_UTM12_NAD83 did not have good attribute values to describe Use Level.""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""Use_Level_2_Source"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""This is the name of the feature class containing the 2nd-closest feature to this feature.  "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""The 5 potential contributer feature classes are Roads_KNF_UTM12_NAD83, BLM_AZStrip_UTM12_NAD83, "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""Roads_UTM12_NAD83, UTAH_Roads_Near_AOI_UTM12_NAD83 and TIGER_2012_Trans_Near_Park_UTM12_NAD83.  "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""The maximum distance considered was 10 meters.""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""Use_Level_2_Distance"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""This is the distance in meters from the centroid of this feature to any part of the "" & _" & vbNewLine
'  strBaseString = strBaseString & "           ""2nd-closest feature from one of 5 potential contributor feature classes, up to a maximum distance "" & _" & vbNewLine
'  strBaseString = strBaseString & "           ""of 10 meters.  "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""The 5 potential contributer feature classes are Roads_KNF_UTM12_NAD83, BLM_AZStrip_UTM12_NAD83, "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""Roads_UTM12_NAD83, UTAH_Roads_Near_AOI_UTM12_NAD83 and TIGER_2012_Trans_Near_Park_UTM12_NAD83.""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.RemoveProperty strXDom" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomMin, ""0""" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomMax, ""10""" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomUnits, ""Meters""" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""Use_Level_3"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = """"""Use Level"""" indicates the transportation quality and accessibility of this feature, such as "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""distinguishing between paved roads and 4-wheel drive tracks.  This attribute field contains "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""the Use Level value from the 3rd-closest feature from 5 potential contributor feature "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""classes, where distance is defined as the distance from the centerpoint of this feature to the "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""nearest point on the nearby feature.  The maximum distance considered was 10 meters.  This value "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""should be Null if there was no 3rd-closest feature within 10m."" & vbCrLf & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""This attribute value was drawn from different attribute fields depending on the feature class:"" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from BLM_AZStrip_UTM12_NAD83, then the attribute field """"E_Map""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"BLM_E_Map"""" above for definitions."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from Roads_UTM12_NAD83, then the attribute field """"Status""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"ROADS_Status"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from Roads_KNF_UTM12_NAD83, then the attribute field """"OPER_MAINT""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"KNF_Oper_Maint"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from UTAH_Roads_Near_AOI_UTM12_NAD83, then the attribute field """"SurfType""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"UTAH_SurfType"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from TIGER_2012_Trans_Near_Park_UTM12_NAD83, then the attribute field """"MT_FCLASS""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"TIG_MT_Fclass"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""The feature classes Trails_KNF_UTM12_NAD83, Trails_UTM12_NAD83, Routes_UTM12_NAD83 and "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""TransportationLine_UTM12_NAD83 did not have good attribute values to describe Use Level.""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""Use_Level_3_Source"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""This is the name of the feature class containing the 3rd-closest feature to this feature.  "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""The 5 potential contributer feature classes are Roads_KNF_UTM12_NAD83, BLM_AZStrip_UTM12_NAD83, "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""Roads_UTM12_NAD83, UTAH_Roads_Near_AOI_UTM12_NAD83 and TIGER_2012_Trans_Near_Park_UTM12_NAD83.  "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""The maximum distance considered was 10 meters.""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""Use_Level_3_Distance"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""This is the distance in meters from the centroid of this feature to any part of the "" & _" & vbNewLine
'  strBaseString = strBaseString & "           ""3rd-closest feature from one of 5 potential contributor feature classes, up to a maximum distance "" & _" & vbNewLine
'  strBaseString = strBaseString & "           ""of 10 meters.  "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""The 5 potential contributer feature classes are Roads_KNF_UTM12_NAD83, BLM_AZStrip_UTM12_NAD83, "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""Roads_UTM12_NAD83, UTAH_Roads_Near_AOI_UTM12_NAD83 and TIGER_2012_Trans_Near_Park_UTM12_NAD83.""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.RemoveProperty strXDom" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomMin, ""0""" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomMax, ""10""" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomUnits, ""Meters""" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""Use_Level_4"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = """"""Use Level"""" indicates the transportation quality and accessibility of this feature, such as "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""distinguishing between paved roads and 4-wheel drive tracks.  This attribute field contains "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""the Use Level value from the 4th-closest feature from 5 potential contributor feature "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""classes, where distance is defined as the distance from the centerpoint of this feature to the "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""nearest point on the nearby feature.  The maximum distance considered was 10 meters.  This value "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""should be Null if there was no 4th-closest feature within 10m."" & vbCrLf & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""This attribute value was drawn from different attribute fields depending on the feature class:"" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from BLM_AZStrip_UTM12_NAD83, then the attribute field """"E_Map""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"BLM_E_Map"""" above for definitions."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from Roads_UTM12_NAD83, then the attribute field """"Status""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"ROADS_Status"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from Roads_KNF_UTM12_NAD83, then the attribute field """"OPER_MAINT""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"KNF_Oper_Maint"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from UTAH_Roads_Near_AOI_UTM12_NAD83, then the attribute field """"SurfType""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"UTAH_SurfType"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from TIGER_2012_Trans_Near_Park_UTM12_NAD83, then the attribute field """"MT_FCLASS""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"TIG_MT_Fclass"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""The feature classes Trails_KNF_UTM12_NAD83, Trails_UTM12_NAD83, Routes_UTM12_NAD83 and "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""TransportationLine_UTM12_NAD83 did not have good attribute values to describe Use Level.""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""Use_Level_4_Source"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""This is the name of the feature class containing the 4th-closest feature to this feature.  "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""The 5 potential contributer feature classes are Roads_KNF_UTM12_NAD83, BLM_AZStrip_UTM12_NAD83, "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""Roads_UTM12_NAD83, UTAH_Roads_Near_AOI_UTM12_NAD83 and TIGER_2012_Trans_Near_Park_UTM12_NAD83.  "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""The maximum distance considered was 10 meters.""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "   " & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""Use_Level_4_Distance"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""This is the distance in meters from the centroid of this feature to any part of the "" & _" & vbNewLine
'  strBaseString = strBaseString & "           ""4th-closest feature from one of 5 potential contributor feature classes, up to a maximum distance "" & _" & vbNewLine
'  strBaseString = strBaseString & "           ""of 10 meters.  "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""The 5 potential contributer feature classes are Roads_KNF_UTM12_NAD83, BLM_AZStrip_UTM12_NAD83, "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""Roads_UTM12_NAD83, UTAH_Roads_Near_AOI_UTM12_NAD83 and TIGER_2012_Trans_Near_Park_UTM12_NAD83.""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.RemoveProperty strXDom" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomMin, ""0""" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomMax, ""10""" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomUnits, ""Meters""" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "    ' ----------------------------------------------------------------------------------------------" & vbNewLine
'  strBaseString = strBaseString & "    ' <<<<<<<<<<<<  Final Types  >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>" & vbNewLine
'  strBaseString = strBaseString & "    '-----------------------------------------------------------------------------------------------" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""Type_1"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""This is the Type value from the TransportationLine_UTM12_NAD83 feature "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""class attribute field """"PrimaryUse"""".  This attribute field ONLY has values for features "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""INSIDE the Grand Canyon NP.  If this feature is outside the park boundaries, or if this "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""feature was imported from another feature "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""class, then this value will be either '<Null>' or '<-- No Value -->'.""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""Type_2"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = """"""Type"""" indicates the primary use or intended purpose of this feature, such as "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""distinguishing betweeen roads and hiking trails.  This attribute field contains "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""the Type value from the closest feature from 5 potential contributor feature "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""classes, where distance is defined as the distance from the centerpoint of this feature to the "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""nearest point on the nearby feature.  The maximum distance considered was 10 meters.  This value "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""should be Null if there were no features within 10m."" & vbCrLf & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""This attribute value was drawn from different attribute fields depending on the feature class:"" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from BLM_AZStrip_UTM12_NAD83, then this value is concatenated from the three attribute fields "" & _" & vbNewLine
'  strBaseString = strBaseString & "          """"""OBS_USE1"""", """"OBS_USE2"""" and """"OBS_USE3""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See descriptions of """"BLM_OBS_USE1"""", """"BLM_OBS_USE2"""" and """"BLM_OBS_USE3"""" above for definitions."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from Roads_KNF_UTM12_NAD83, then the attribute field """"Lanes""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"KNF_Lanes"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from Trails_KNF_UTM12_NAD83, then the attribute field """"Designed_U""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"KNFTR_Designed_Use"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from TIGER_2012_Trans_Near_Park_UTM12_NAD83, then the attribute field """"MT_FCLASS""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"TIG_MT_Fclass"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from TransportationLine_UTM12_NAD83, then the attribute field """"PrimaryUse""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"PrimaryUse"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""The feature classes Roads_UTM12_NAD83, Trails_UTM12_NAD83, Routes_UTM12_NAD83 and "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""Utah_Roads_Near_AOI_UTM12_NAD83 did not have good attribute values to describe Type.""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""Type_2_Source"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""This is the name of the feature class containing the closest feature to this feature.  "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""The 5 potential contributer feature classes are Roads_KNF_UTM12_NAD83, BLM_AZStrip_UTM12_NAD83, "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""Trails_KNF_UTM12_NAD83, TIGER_2012_Trans_Near_Park_UTM12_NAD83 and TransportationLine_UTM12_NAD83.  "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""The maximum distance considered was 10 meters.""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""Type_2_Distance"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""This is the distance in meters from the centroid of this feature to any part of the "" & _" & vbNewLine
'  strBaseString = strBaseString & "           ""closest feature from one of 4 potential contributor feature classes, up to a maximum distance "" & _" & vbNewLine
'  strBaseString = strBaseString & "           ""of 10 meters.  "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""The 5 potential contributer feature classes are Roads_KNF_UTM12_NAD83, BLM_AZStrip_UTM12_NAD83, "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""Trails_KNF_UTM12_NAD83, TIGER_2012_Trans_Near_Park_UTM12_NAD83 and TransportationLine_UTM12_NAD83.""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.RemoveProperty strXDom" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomMin, ""0""" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomMax, ""10""" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomUnits, ""Meters""" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""Type_3"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = """"""Type"""" indicates the primary use or intended purpose of this feature, such as "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""distinguishing betweeen roads and hiking trails.  "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""This attribute field contains "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""the Type value from the 2nd-closest feature from 5 potential contributor feature "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""classes, where distance is defined as the distance from the centerpoint of this feature to the "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""nearest point on the nearby feature.  The maximum distance considered was 10 meters.  This value "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""should be Null if there was no 2nd-closest feature within 10m."" & vbCrLf & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""This attribute value was drawn from different attribute fields depending on the feature class:"" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from BLM_AZStrip_UTM12_NAD83, then this value is concatenated from the three attribute fields "" & _" & vbNewLine
'  strBaseString = strBaseString & "          """"""OBS_USE1"""", """"OBS_USE2"""" and """"OBS_USE3""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See descriptions of """"BLM_OBS_USE1"""", """"BLM_OBS_USE2"""" and """"BLM_OBS_USE3"""" above for definitions."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from Roads_KNF_UTM12_NAD83, then the attribute field """"Lanes""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"KNF_Lanes"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from Trails_KNF_UTM12_NAD83, then the attribute field """"Designed_U""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"KNFTR_Designed_Use"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from TIGER_2012_Trans_Near_Park_UTM12_NAD83, then the attribute field """"MT_FCLASS""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"TIG_MT_Fclass"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from TransportationLine_UTM12_NAD83, then the attribute field """"PrimaryUse""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"PrimaryUse"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""The feature classes Roads_UTM12_NAD83, Trails_UTM12_NAD83, Routes_UTM12_NAD83 and "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""Utah_Roads_Near_AOI_UTM12_NAD83 did not have good attribute values to describe Type.""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""Type_3_Source"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""This is the name of the feature class containing the 2nd-closest feature to this feature.  "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""The 5 potential contributer feature classes are Roads_KNF_UTM12_NAD83, BLM_AZStrip_UTM12_NAD83, "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""Trails_KNF_UTM12_NAD83, TIGER_2012_Trans_Near_Park_UTM12_NAD83 and TransportationLine_UTM12_NAD83.  "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""The maximum distance considered was 10 meters.""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""Type_3_Distance"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""This is the distance in meters from the centroid of this feature to any part of the "" & _" & vbNewLine
'  strBaseString = strBaseString & "           ""2nd-closest feature from one of 4 potential contributor feature classes, up to a maximum distance "" & _" & vbNewLine
'  strBaseString = strBaseString & "           ""of 10 meters."" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""The 5 potential contributer feature classes are Roads_KNF_UTM12_NAD83, BLM_AZStrip_UTM12_NAD83, "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""Trails_KNF_UTM12_NAD83, TIGER_2012_Trans_Near_Park_UTM12_NAD83 and TransportationLine_UTM12_NAD83.""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.RemoveProperty strXDom" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomMin, ""0""" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomMax, ""10""" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomUnits, ""Meters""" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""Type_4"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = """"""Type"""" indicates the primary use or intended purpose of this feature, such as "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""distinguishing betweeen roads and hiking trails.  "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""This attribute field contains "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""the Type value from the 3rd-closest feature from 5 potential contributor feature "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""classes, where distance is defined as the distance from the centerpoint of this feature to the "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""nearest point on the nearby feature.  The maximum distance considered was 10 meters.  This value "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""should be Null if there was no 3rd-closest feature within 10m."" & vbCrLf & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""This attribute value was drawn from different attribute fields depending on the feature class:"" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from BLM_AZStrip_UTM12_NAD83, then this value is concatenated from the three attribute fields "" & _" & vbNewLine
'  strBaseString = strBaseString & "          """"""OBS_USE1"""", """"OBS_USE2"""" and """"OBS_USE3""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See descriptions of """"BLM_OBS_USE1"""", """"BLM_OBS_USE2"""" and """"BLM_OBS_USE3"""" above for definitions."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from Roads_KNF_UTM12_NAD83, then the attribute field """"Lanes""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"KNF_Lanes"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from Trails_KNF_UTM12_NAD83, then the attribute field """"Designed_U""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"KNFTR_Designed_Use"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from TIGER_2012_Trans_Near_Park_UTM12_NAD83, then the attribute field """"MT_FCLASS""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"TIG_MT_Fclass"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from TransportationLine_UTM12_NAD83, then the attribute field """"PrimaryUse""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"PrimaryUse"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""The feature classes Roads_UTM12_NAD83, Trails_UTM12_NAD83, Routes_UTM12_NAD83 and "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""Utah_Roads_Near_AOI_UTM12_NAD83 did not have good attribute values to describe Type.""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""Type_4_Source"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""This is the name of the feature class containing the 3rd-closest feature to this feature.  "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""The 5 potential contributer feature classes are Roads_KNF_UTM12_NAD83, BLM_AZStrip_UTM12_NAD83, "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""Trails_KNF_UTM12_NAD83, TIGER_2012_Trans_Near_Park_UTM12_NAD83 and TransportationLine_UTM12_NAD83.  "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""The maximum distance considered was 10 meters.""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""Type_4_Distance"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""This is the distance in meters from the centroid of this feature to any part of the "" & _" & vbNewLine
'  strBaseString = strBaseString & "           ""3rd-closest feature from one of 4 potential contributor feature classes, up to a maximum distance "" & _" & vbNewLine
'  strBaseString = strBaseString & "           ""of 10 meters."" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""The 5 potential contributer feature classes are Roads_KNF_UTM12_NAD83, BLM_AZStrip_UTM12_NAD83, "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""Trails_KNF_UTM12_NAD83, TIGER_2012_Trans_Near_Park_UTM12_NAD83 and TransportationLine_UTM12_NAD83.""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.RemoveProperty strXDom" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomMin, ""0""" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomMax, ""10""" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomUnits, ""Meters""" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "   " & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "    ' ----------------------------------------------------------------------------------------------" & vbNewLine
'  strBaseString = strBaseString & "    ' <<<<<<<<<<<<  Surface Types  >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>" & vbNewLine
'  strBaseString = strBaseString & "    '-----------------------------------------------------------------------------------------------" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""SurfaceType_1"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = """"""Surface_Type"""" indicates the construction material of this feature, "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""such as distinguishing betweeen paved, gravel and dirt roads.  "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""This is the Surface Type value from the TransportationLine_UTM12_NAD83 feature "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""class attribute field """"SurfaceMaterial"""".  This attribute field ONLY has values for features "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""INSIDE the Grand Canyon NP.  If this feature is outside the park boundaries, or if this "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""feature was imported from another feature "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""class, then this value will be either '<Null>' or '<-- No Value -->'.""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""SurfaceType_2"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = """"""Surface_Type"""" indicates the construction material of this feature, "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""such as distinguishing betweeen paved, gravel and dirt roads.  "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""This is the Surface Type value from the closest feature from 6 potential contributor feature "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""classes, where distance is defined as the distance from the centerpoint of this feature to the "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""nearest point on the nearby feature.  The maximum distance considered was 10 meters.  This value "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""should be Null if there were no features within 10m."" & vbCrLf & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""This attribute value was drawn from different attribute fields depending on the feature class:"" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from Roads_UTM12_NAD83, then the attribute field """"SurfaceMaterial""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"ROADS_SurfaceMaterial"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from BLM_AZStrip_UTM12_NAD83, then the attribute field """"Surface_PR""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"BLM_Surface_PR"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from Roads_KNF_UTM12_NAD83, then the attribute field """"SURFACE_TY""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"KNF_Surface_Type"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from TIGER_2012_Trans_Near_Park_UTM12_NAD83, then the attribute field """"MT_FCLASS""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"TIG_MT_Fclass"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from Utah_Roads_Near_AOI_UTM12_NAD83, then the attribute field """"SurfType""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"UTAH_SurfType"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from TransportationLine_UTM12_NAD83, then the attribute field """"SurfaceMaterial""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"SurfaceMaterial"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""The feature classes Trails_KNF_UTM12_NAD83, Trails_UTM12_NAD83 and Routes_UTM12_NAD83 did not "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""have good attribute values to describe Surface Type.""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""SurfaceType_2_Source"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""This is the name of the feature class containing the closest feature to this feature.  "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""The 6 potential contributer feature classes are Roads_UTM12_NAD83, BLM_AZStrip_UTM12_NAD83, "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""Roads_KNF_UTM12_NAD83, TIGER_2012_Trans_Near_Park_UTM12_NAD83, Utah_Roads_Near_AOI_UTM12_NAD83 and "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""TransportationLine_UTM12_NAD83.  "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""The maximum distance considered was 10 meters.""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""SurfaceType_2_Distance"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""This is the distance in meters from the centroid of this feature to any part of the "" & _" & vbNewLine
'  strBaseString = strBaseString & "           ""closest feature from one of 5 potential contributor feature classes, up to a maximum distance "" & _" & vbNewLine
'  strBaseString = strBaseString & "           ""of 10 meters.  "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""The 6 potential contributer feature classes are Roads_UTM12_NAD83, BLM_AZStrip_UTM12_NAD83, "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""Roads_KNF_UTM12_NAD83, TIGER_2012_Trans_Near_Park_UTM12_NAD83, Utah_Roads_Near_AOI_UTM12_NAD83 and "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""TransportationLine_UTM12_NAD83.""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.RemoveProperty strXDom" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomMin, ""0""" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomMax, ""10""" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomUnits, ""Meters""" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""SurfaceType_3"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = """"""Surface_Type"""" indicates the construction material of this feature, "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""such as distinguishing betweeen paved, gravel and dirt roads.  "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""This is the Surface Type value from the second closest feature from 6 potential contributor feature "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""classes, where distance is defined as the distance from the centerpoint of this feature to the "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""nearest point on the nearby feature.  The maximum distance considered was 10 meters.  This value "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""should be Null if there was no 2nd-closest feature within 10m."" & vbCrLf & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""This attribute value was drawn from different attribute fields depending on the feature class:"" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from Roads_UTM12_NAD83, then the attribute field """"SurfaceMaterial""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"ROADS_SurfaceMaterial"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from BLM_AZStrip_UTM12_NAD83, then the attribute field """"Surface_PR""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"BLM_Surface_PR"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from Roads_KNF_UTM12_NAD83, then the attribute field """"SURFACE_TY""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"KNF_Surface_Type"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from TIGER_2012_Trans_Near_Park_UTM12_NAD83, then the attribute field """"MT_FCLASS""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"TIG_MT_Fclass"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from Utah_Roads_Near_AOI_UTM12_NAD83, then the attribute field """"SurfType""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"UTAH_SurfType"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from TransportationLine_UTM12_NAD83, then the attribute field """"SurfaceMaterial""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"SurfaceMaterial"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""The feature classes Trails_KNF_UTM12_NAD83, Trails_UTM12_NAD83 and Routes_UTM12_NAD83 did not "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""have good attribute values to describe Surface Type.""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""SurfaceType_3_Source"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""This is the name of the feature class containing the second closest feature to this feature.  "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""The 6 potential contributer feature classes are Roads_UTM12_NAD83, BLM_AZStrip_UTM12_NAD83, "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""Roads_KNF_UTM12_NAD83, TIGER_2012_Trans_Near_Park_UTM12_NAD83, Utah_Roads_Near_AOI_UTM12_NAD83 and "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""TransportationLine_UTM12_NAD83.""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""SurfaceType_3_Distance"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""This is the distance in meters from the centroid of this feature to any part of the "" & _" & vbNewLine
'  strBaseString = strBaseString & "           ""second closest feature from one of 5 potential contributor feature classes, up to a maximum distance "" & _" & vbNewLine
'  strBaseString = strBaseString & "           ""of 10 meters.  "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""The 6 potential contributer feature classes are Roads_UTM12_NAD83, BLM_AZStrip_UTM12_NAD83, "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""Roads_KNF_UTM12_NAD83, TIGER_2012_Trans_Near_Park_UTM12_NAD83, Utah_Roads_Near_AOI_UTM12_NAD83 and "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""TransportationLine_UTM12_NAD83.""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.RemoveProperty strXDom" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomMin, ""0""" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomMax, ""10""" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomUnits, ""Meters""" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""SurfaceType_4"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = """"""Surface_Type"""" indicates the construction material of this feature, "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""such as distinguishing betweeen paved, gravel and dirt roads.  "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""This is the Surface Type value from the 3rd-closest feature from 6 potential contributor feature "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""classes, where distance is defined as the distance from the centerpoint of this feature to the "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""nearest point on the nearby feature.  The maximum distance considered was 10 meters.  This value "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""should be Null if there was no 3rd-closest feature within 10m."" & vbCrLf & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""This attribute value was drawn from different attribute fields depending on the feature class:"" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from Roads_UTM12_NAD83, then the attribute field """"SurfaceMaterial""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"ROADS_SurfaceMaterial"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from BLM_AZStrip_UTM12_NAD83, then the attribute field """"Surface_PR""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"BLM_Surface_PR"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from Roads_KNF_UTM12_NAD83, then the attribute field """"SURFACE_TY""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"KNF_Surface_Type"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from TIGER_2012_Trans_Near_Park_UTM12_NAD83, then the attribute field """"MT_FCLASS""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"TIG_MT_Fclass"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from Utah_Roads_Near_AOI_UTM12_NAD83, then the attribute field """"SurfType""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"UTAH_SurfType"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          "" - If from TransportationLine_UTM12_NAD83, then the attribute field """"SurfaceMaterial""""."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""    - See description of """"SurfaceMaterial"""" above for definition."" & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""The feature classes Trails_KNF_UTM12_NAD83, Trails_UTM12_NAD83 and Routes_UTM12_NAD83 did not "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""have good attribute values to describe Surface Type.""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""SurfaceType_4_Source"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""This is the name of the feature class containing the 3rd closest feature to this feature.  "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""The 6 potential contributer feature classes are Roads_UTM12_NAD83, BLM_AZStrip_UTM12_NAD83, "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""Roads_KNF_UTM12_NAD83, TIGER_2012_Trans_Near_Park_UTM12_NAD83, Utah_Roads_Near_AOI_UTM12_NAD83 and "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""TransportationLine_UTM12_NAD83.  "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""The maximum distance considered was 10 meters.""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""SurfaceType_4_Distance"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""This is the distance in meters from the centroid of this feature to any part of the "" & _" & vbNewLine
'  strBaseString = strBaseString & "           ""3rd-closest feature from one of 5 potential contributor feature classes, up to a maximum distance "" & _" & vbNewLine
'  strBaseString = strBaseString & "           ""of 10 meters.  "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""The 6 potential contributer feature classes are Roads_UTM12_NAD83, BLM_AZStrip_UTM12_NAD83, "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""Roads_KNF_UTM12_NAD83, TIGER_2012_Trans_Near_Park_UTM12_NAD83, Utah_Roads_Near_AOI_UTM12_NAD83 and "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""TransportationLine_UTM12_NAD83.""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.RemoveProperty strXDom" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomMin, ""0""" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomMax, ""10""" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomUnits, ""Meters""" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "    ' ----------------------------------------------------------------------------------------------" & vbNewLine
'  strBaseString = strBaseString & "    ' <<<<<<<<<<<<  Nearest  >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>" & vbNewLine
'  strBaseString = strBaseString & "    '-----------------------------------------------------------------------------------------------" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""Nearest_Source"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""This feature class is derived primarily from the polyline features of the "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""GCNP feature class """"TransportationLine_UTM12_NAD83"""", but the attributes for these features "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""are extracted from 8 other feature classes based on proximity ([1] BLM_AZStrip_UTM12_NAD83, [2] "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""TIGER_2012_Trans_Near_Park_UTM12_NAD83, "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""[3] Roads_KNF_UTM12_NAD83, [4] Utah_Roads_Near_AOI_UTM12_NAD83, [5] Trails_KNF_UTM12_NAD83, "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""[6] Roads_UTM12_NAD83, [7] Trails_UTM12_NAD83 or [8] Routes_UTM12_NAD83)."" & vbCrLf & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""This attribute field shows the name of the feature class containing the closest feature to the "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""centroid of this feature, up to a maximum distance of 100m.  "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""This closest feature is used as a candidate to extract attribute values from.""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""Nearest_OID"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""This feature class is derived primarily from the polyline features of the "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""GCNP feature class """"TransportationLine_UTM12_NAD83"""", but the attributes for these features "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""are extracted from 8 other feature classes based on proximity ([1] BLM_AZStrip_UTM12_NAD83, [2] "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""TIGER_2012_Trans_Near_Park_UTM12_NAD83, "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""[3] Roads_KNF_UTM12_NAD83, [4] Utah_Roads_Near_AOI_UTM12_NAD83, [5] Trails_KNF_UTM12_NAD83, "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""[6] Roads_UTM12_NAD83, [7] Trails_UTM12_NAD83 or [8] Routes_UTM12_NAD83)."" & vbCrLf & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""This attribute field shows the OID value of the closest feature from these 8 feature classes to the "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""centroid of this feature, up to a maximum distance of 100m.  "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""This closest feature is used as a candidate to extract attribute values from.""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""Nearest_Distance"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""This feature class is derived primarily from the polyline features of the "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""GCNP feature class """"TransportationLine_UTM12_NAD83"""", but the attributes for these features "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""are extracted from 8 other feature classes based on proximity ([1] BLM_AZStrip_UTM12_NAD83, [2] "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""TIGER_2012_Trans_Near_Park_UTM12_NAD83, "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""[3] Roads_KNF_UTM12_NAD83, [4] Utah_Roads_Near_AOI_UTM12_NAD83, [5] Trails_KNF_UTM12_NAD83, "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""[6] Roads_UTM12_NAD83, [7] Trails_UTM12_NAD83 or [8] Routes_UTM12_NAD83)."" & vbCrLf & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""This attribute field shows the distance in meters to the closest feature from these 8 feature classes to the "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""centroid of this feature, up to a maximum distance of 100m.  "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""This closest feature is used as a candidate to extract attribute values from.""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.RemoveProperty strXDom" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomMin, ""0""" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomMax, ""100""" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomUnits, ""Meters""" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "    ' ----------------------------------------------------------------------------------------------" & vbNewLine
'  strBaseString = strBaseString & "    ' <<<<<<<<<<<<  Ownership  >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>" & vbNewLine
'  strBaseString = strBaseString & "    '-----------------------------------------------------------------------------------------------" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "   ' 2 OWNERSHIP FIELDS" & vbNewLine
'  strBaseString = strBaseString & "   ' IF ARIZONA, THEN 1ST OWNERSHIP FIELD IS FROM ""OWN.SHP"" FIELD [DESC_]" & vbNewLine
'  strBaseString = strBaseString & "   '                  2ND OWNERSHIP FIELD IS FROM ""OWN.SHP"" FIELD [CATEGORY]" & vbNewLine
'  strBaseString = strBaseString & "   ' IF UTAH, THEN 1ST OWNERSHIP FIELD IS FROM ""LandOwnership"" FIELD [UT_LGD]; ALIAS ""UTAH BLM LEGEND""" & vbNewLine
'  strBaseString = strBaseString & "   '               2ND OWNERSHIP FIELD IS FROM ""LandOwnership"" FIELD [ADMIN]" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""State"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""The State that this feature lies in.  This value will be """"UTAH"""" if the "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""feature was extracted from either the Utah Roads feature class (Utah_Roads_New_AOI_UTM12_NAD83) or "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""from the Utah portion of the TIGER_2012_Trans_Near_Park_UTM12_NAD83 feature class.  Otherwise "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""this value will be """"Arizona"""".""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""Ownership_1"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""This value contains the agency or owner who is responsible for the polygon this "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""feature lies in. Therefore, this value is not guaranteed to show the true """"owner"""" of the polyline, "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""but rather the most likely responsible party."" & vbCrLf & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""If the transportation feature is in Arizona, then the """"Ownership_1"""" is extracted from the "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""attribute field """"DESC_"""" in the Arizona Land Ownership feature class """"OWN""""."" & vbCrLf & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""If the transportation feature is in Utah, then the """"Ownership_1"""" is extracted from the "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""attribute field """"UT_LGD"""", alias """"UTAH BLM LEGEND"""" in the Utah Land Ownership feature "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""class """"LandOwnership"""".""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""Ownership_2"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""This value contains the agency or owner who is responsible for the polygon this "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""feature lies in. Therefore, this value is not guaranteed to show the true """"owner"""" of the polyline, "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""but rather the most likely responsible party."" & vbCrLf & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""If the transportation feature is in Arizona, then the """"Ownership_2"""" is extracted from the "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""attribute field """"CATEGORY"""" in the Arizona Land Ownership feature class """"OWN""""."" & vbCrLf & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""If the transportation feature is in Utah, then the """"Ownership_1"""" is extracted from the "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""attribute field """"UT_LGD"""", alias """"UTAH BLM LEGEND"""" in the Utah Land Ownership feature "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""class """"ADMIN"""".""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "    ' ----------------------------------------------------------------------------------------------" & vbNewLine
'  strBaseString = strBaseString & "    ' <<<<<<<<<<<<  ADD_Feature, etc.  >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>" & vbNewLine
'  strBaseString = strBaseString & "    '-----------------------------------------------------------------------------------------------" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""Add_Feature"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = """"""TRUE"""" or """"FALSE"""", where """"True"""" indicates that this feature came from one of the "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""ancillary feature classes ([1] BLM_AZStrip_UTM12_NAD83, [2] TIGER_2012_Trans_Near_Park_UTM12_NAD83, "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""[3] Roads_KNF_UTM12_NAD83, [4] Utah_Roads_Near_AOI_UTM12_NAD83, [5] Trails_KNF_UTM12_NAD83, "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""[6] Roads_UTM12_NAD83, [7] Trails_UTM12_NAD83 or [8] Routes_UTM12_NAD83), and """"FALSE"""" indicates "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""that this feature came from the original TransportationLine_UTM12_NAD83 feature class.""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""Add_Feature_Source"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""The name of the source feature class that this feature was extracted from.  Possible "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""values are TransportationLine_UTM12_NAD83, BLM_AZStrip_UTM12_NAD83, "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""TIGER_2012_Trans_Near_Park_UTM12_NAD83, Roads_KNF_UTM12_NAD83, Utah_Roads_Near_AOI_UTM12_NAD83, "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""[5] Trails_KNF_UTM12_NAD83, Roads_UTM12_NAD83, Trails_UTM12_NAD83 or Routes_UTM12_NAD83.""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""Add_Date"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""The date that this feature was extracted from the ancillary feature class and added to "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""this feature class.""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""Add_Name"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""The person who added this feature.  JJENNESS = Jeff Jenness, jeffj@jennessent.com""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""Add_OID"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""The Object ID value of the feature in the original feature class.""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""Geometry_Modified"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""All features in this feature class were imported from one of nine source feature "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""classes (([1] BLM_AZStrip_UTM12_NAD83, [2] TIGER_2012_Trans_Near_Park_UTM12_NAD83, "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""[3] Roads_KNF_UTM12_NAD83, [4] Utah_Roads_Near_AOI_UTM12_NAD83, [5] Trails_KNF_UTM12_NAD83, "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""[6] Roads_UTM12_NAD83, [7] Trails_UTM12_NAD83, [8] Routes_UTM12_NAD83 or [9] "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""TransportationLine_UTM12_NAD83).  This attribute field contains """"TRUE"""" or """"FALSE"""" values "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""indicating whether the geometry was modified after being imported into this feature class. "" & vbCrLf & vbCrLf & _" & vbNewLine
'  strBaseString = strBaseString & "          ""In general, the most likely reason a geometry would be altered from its original state would "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""be to make it match topologically with the other features.""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "    ' ----------------------------------------------------------------------------------------------" & vbNewLine
'  strBaseString = strBaseString & "    ' <<<<<<<<<<<<  Trails_KNF_UTM12_NAD83  >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>" & vbNewLine
'  strBaseString = strBaseString & "    '-----------------------------------------------------------------------------------------------" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""KNFTR_Add_Attribute"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = """"""TRUE"""" or """"FALSE"""", whether there was a polyline feature from the 'Trails_KNF_UTM12_NAD83' "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""feature class within 100m of the centroid of this polyline.  If so, then this '"" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""Trails_KNF_UTM12_NAD83' feature would be considered a candidate for extracting attribute values "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""to describe this feature.""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""KNFTR_Attribute_Source"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""If 'KNFTR_Add_Attribute' above = """"TRUE"""", then this field will contain the name of the "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""Kaibab NF Trails feature class (Trails_KNF_UTM12_NAD83).  Otherwise it should be Null.""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""KNFTR_Attribute_Distance"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""If 'KNFTR_Add_Attribute' above = """"TRUE"""", then this field will contain the distance "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""(in meters) from the centroid of this polyline to the nearest Kaibab NF Trails feature.  Otherwise it should be Null.  "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""Values should always be <= 100m.""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.RemoveProperty strXDom" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomMin, ""0""" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomMax, ""100""" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomUnits, ""Meters""" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""KNFTR_ObjectID"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""If 'KNFTR_Add_Attribute' above = """"TRUE"""", then this field will contain the OBJECTID "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""value of the nearest Kaibab NF Trails feature, from the Trails_KNF_UTM12_NAD83 feature class.""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "    For lngKNFTR_Index = 0 To UBound(strKNFTR_RoadArray, 1)" & vbNewLine
'  strBaseString = strBaseString & "      If strFieldName = strKNFTR_RoadArray(lngKNFTR_Index, 0) Then '  ""BLM_OBS_USE1"", ""BLM_OBS_USE2"" or ""BLM_OBS_USE3""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "        ' FIND ORIGINAL FIELD METADATA" & vbNewLine
'  strBaseString = strBaseString & "        For lngIndex = 0 To pTrails_KNF_UTM12_NAD83_FClass.Fields.FieldCount" & vbNewLine
'  strBaseString = strBaseString & "          strNames(0) = ""eainfo/detailed/attr["" & CStr(lngIndex) & ""]""" & vbNewLine
'  strBaseString = strBaseString & "          pTrails_KNF_UTM12_NAD83_PropSet.GetProperties strNames, varVals" & vbNewLine
'  strBaseString = strBaseString & "          varSubVals = varVals(0)" & vbNewLine
'  strBaseString = strBaseString & "          If Not IsEmpty(varSubVals) Then" & vbNewLine
'  strBaseString = strBaseString & "            ' GET FIELD NAME" & vbNewLine
'  strBaseString = strBaseString & "            strOrigFClassXName = ""eainfo/detailed/attr["" & CStr(lngIndex) & ""]/attrlabl""" & vbNewLine
'  strBaseString = strBaseString & "            varProperty = pTrails_KNF_UTM12_NAD83_PropSet.GetProperty(strOrigFClassXName)" & vbNewLine
'  strBaseString = strBaseString & "            If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "              varSubProperty = varProperty(0)" & vbNewLine
'  strBaseString = strBaseString & "              If IsEmpty(varSubProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "                strOrigFieldName = ""<-- No Field Name -->""" & vbNewLine
'  strBaseString = strBaseString & "              Else" & vbNewLine
'  strBaseString = strBaseString & "                strOrigFieldName = CStr(varSubProperty)" & vbNewLine
'  strBaseString = strBaseString & "              End If" & vbNewLine
'  strBaseString = strBaseString & "            Else" & vbNewLine
'  strBaseString = strBaseString & "              strOrigFieldName = ""<-- No Field Name -->""" & vbNewLine
'  strBaseString = strBaseString & "            End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "            If strOrigFieldName = strKNFTR_RoadArray(lngKNFTR_Index, 1) Then ' ""OBS_USE1"", ""OBS_USE2"", or ""OBS_USE3""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "              ' CHECK IF DOMAIN INFO" & vbNewLine
'  strBaseString = strBaseString & "              strOrigFClassXName = ""eainfo/detailed/attr["" & CStr(lngIndex) & ""]/attrdomv""" & vbNewLine
'  strBaseString = strBaseString & "              varProperty = pTrails_KNF_UTM12_NAD83_PropSet.GetProperty(strOrigFClassXName)" & vbNewLine
'  strBaseString = strBaseString & "              If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "                varSubProperty = varProperty(0)" & vbNewLine
'  strBaseString = strBaseString & "'                For lngIndex2 = 0 To UBound(varProperty)" & vbNewLine
'  strBaseString = strBaseString & "'                  Debug.Print ""       "" & CStr(lngIndex2) & ""] "" & CStr(varProperty(lngIndex2))" & vbNewLine
'  strBaseString = strBaseString & "'                Next lngIndex2" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "                ' GET DOMAIN ATTRIBUTES FROM ORIGINAL FCLASS" & vbNewLine
'  strBaseString = strBaseString & "                strOrigFClassXName = ""eainfo/detailed/attr["" & CStr(lngIndex) & ""]/attrdomv""" & vbNewLine
'  strBaseString = strBaseString & "                varProperty = pTrails_KNF_UTM12_NAD83_PropSet.GetProperty(strOrigFClassXName)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "                ' ADD DOMAIN ATTRIBUTES TO NEW FCLASS" & vbNewLine
'  strBaseString = strBaseString & "                varCheckPropertyPresent = pPropSet.GetProperty(""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdomv"")" & vbNewLine
'  strBaseString = strBaseString & "                If Not IsEmpty(varCheckPropertyPresent) Then" & vbNewLine
'  strBaseString = strBaseString & "                  pPropSet.RemoveProperty ""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdomv""" & vbNewLine
'  strBaseString = strBaseString & "                End If" & vbNewLine
'  strBaseString = strBaseString & "                For lngIndex3 = 0 To UBound(varProperty)" & vbNewLine
'  strBaseString = strBaseString & "                  strSubXName1 = ""eainfo/detailed/attr["" & CStr(lngIndex) & ""]/attrdomv/"" & _" & vbNewLine
'  strBaseString = strBaseString & "                      varProperty(lngIndex3) & ""["" & CStr(lngIndex3) & ""]""" & vbNewLine
'  strBaseString = strBaseString & "                  varSubProp1 = pTrails_KNF_UTM12_NAD83_PropSet.GetProperty(strSubXName1)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "                  If UBound(varSubProp1) = 0 Then" & vbNewLine
'  strBaseString = strBaseString & "                    strSubXName3 = ""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdomv/"" & _" & vbNewLine
'  strBaseString = strBaseString & "                        varProperty(lngIndex3) & ""["" & CStr(lngIndex3) & ""]""" & vbNewLine
'  strBaseString = strBaseString & "                    strValue = varSubProp1(0)" & vbNewLine
'  strBaseString = strBaseString & "                    strValue = Replace(strValue, ""20003"", ""2003"")" & vbNewLine
'  strBaseString = strBaseString & "                    strValue = Replace(strValue, ""Dicitonary"", ""Dictionary"")" & vbNewLine
'  strBaseString = strBaseString & "                    pPropSet.SetProperty strSubXName3, strValue" & vbNewLine
'  strBaseString = strBaseString & "                  Else" & vbNewLine
'  strBaseString = strBaseString & "                    For lngIndex4 = 0 To UBound(varSubProp1)" & vbNewLine
'  strBaseString = strBaseString & "                      strSubXName2 = ""eainfo/detailed/attr["" & CStr(lngIndex) & ""]/attrdomv/"" & _" & vbNewLine
'  strBaseString = strBaseString & "                          varProperty(lngIndex3) & ""["" & CStr(lngIndex3) & ""]/"" & varSubProp1(lngIndex4)" & vbNewLine
'  strBaseString = strBaseString & "                      varSubProp2 = pTrails_KNF_UTM12_NAD83_PropSet.GetProperty(strSubXName2)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "                      ' WRITE VALUE BACK TO NEW FCLASS" & vbNewLine
'  strBaseString = strBaseString & "                      strSubXName3 = ""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdomv/"" & _" & vbNewLine
'  strBaseString = strBaseString & "                          varProperty(lngIndex3) & ""["" & CStr(lngIndex3) & ""]/"" & varSubProp1(lngIndex4)" & vbNewLine
'  strBaseString = strBaseString & "                      strValue = varSubProp2(0)" & vbNewLine
'  strBaseString = strBaseString & "                      strValue = Replace(strValue, ""20003"", ""2003"")" & vbNewLine
'  strBaseString = strBaseString & "                      strValue = Replace(strValue, ""Dicitonary"", ""Dictionary"")" & vbNewLine
'  strBaseString = strBaseString & "                      pPropSet.SetProperty strSubXName3, strValue" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "                    Next lngIndex4" & vbNewLine
'  strBaseString = strBaseString & "                  End If" & vbNewLine
'  strBaseString = strBaseString & "                Next lngIndex3" & vbNewLine
'  strBaseString = strBaseString & "              Else" & vbNewLine
'  strBaseString = strBaseString & "'                strOrigFCLassXName = ""eainfo/detailed/attr["" & CStr(lngIndex) & ""]""" & vbNewLine
'  strBaseString = strBaseString & "'                varProperty = pTrails_KNF_UTM12_NAD83_PropSet.GetProperty(strOrigFCLassXName)" & vbNewLine
'  strBaseString = strBaseString & "'                varSubProperty = varProperty(0)" & vbNewLine
'  strBaseString = strBaseString & "'                Debug.Print ""Examining "" & strOrigFieldName & ""...""" & vbNewLine
'  strBaseString = strBaseString & "'                For lngIndex2 = 0 To UBound(varProperty)" & vbNewLine
'  strBaseString = strBaseString & "'                  varSubProp1 = pTrails_KNF_UTM12_NAD83_PropSet.GetProperty(strOrigFCLassXName & ""/"" & _" & vbNewLine
'  strBaseString = strBaseString & "'                        CStr(varProperty(lngIndex2)))" & vbNewLine
'  strBaseString = strBaseString & "'                  Debug.Print ""       "" & CStr(lngIndex2) & ""] "" & CStr(varProperty(lngIndex2)) & "" = "" & _" & vbNewLine
'  strBaseString = strBaseString & "'                        CStr(varSubProp1(0))" & vbNewLine
'  strBaseString = strBaseString & "'                Next lngIndex2" & vbNewLine
'  strBaseString = strBaseString & "              End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "              If strOrigFieldName = ""ROAD_NO_"" Then" & vbNewLine
'  strBaseString = strBaseString & "                Debug.Print ""Here...""" & vbNewLine
'  strBaseString = strBaseString & "              End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "              ' GET DEFINITION FROM ORIGINAL FCLASS" & vbNewLine
'  strBaseString = strBaseString & "              varProperty = pTrails_KNF_UTM12_NAD83_PropSet.GetProperty(""eainfo/detailed/attr["" & CStr(lngIndex) & ""]/attrdef"")" & vbNewLine
'  strBaseString = strBaseString & "              If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "                varSubProperty = varProperty(0)" & vbNewLine
'  strBaseString = strBaseString & "                If IsEmpty(varSubProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "                  strDescription = ""<-- No Description -->""" & vbNewLine
'  strBaseString = strBaseString & "                Else" & vbNewLine
'  strBaseString = strBaseString & "                  strDescription = CStr(varSubProperty)" & vbNewLine
'  strBaseString = strBaseString & "                  strDescription = Trim(strDescription)" & vbNewLine
'  strBaseString = strBaseString & "                  If Right(strDescription, 1) <> ""."" And Right(strDescription, 1) <> ""!"" And _" & vbNewLine
'  strBaseString = strBaseString & "                      Right(strDescription, 1) <> ""?"" Then" & vbNewLine
'  strBaseString = strBaseString & "                    strDescription = strDescription & "".""" & vbNewLine
'  strBaseString = strBaseString & "                  End If" & vbNewLine
'  strBaseString = strBaseString & "                End If" & vbNewLine
'  strBaseString = strBaseString & "              Else" & vbNewLine
'  strBaseString = strBaseString & "                strDescription = ""<-- No Description -->""" & vbNewLine
'  strBaseString = strBaseString & "              End If" & vbNewLine
'  strBaseString = strBaseString & "              strDescription = Replace(strDescription, ""  [Imported from TransportationLine_UTM12_NAD83]"", """", , , vbTextCompare)" & vbNewLine
'  strBaseString = strBaseString & "              strDescription = strDescription & _" & vbNewLine
'  strBaseString = strBaseString & "                  ""  This field should only have a value if the 'KNFTR_Add_Attribute' value = """"TRUE"""".  "" & _" & vbNewLine
'  strBaseString = strBaseString & "                  ""[Imported from Trails_KNF_UTM12_NAD83, Field """""" & strKNFTR_RoadArray(lngKNFTR_Index, 1) & """"""]""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "              ' SET FIELD DEFINITION OF NEW FCLASS" & vbNewLine
'  strBaseString = strBaseString & "              varCheckPropertyPresent = pPropSet.GetProperty(""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdef"")" & vbNewLine
'  strBaseString = strBaseString & "              If Not IsEmpty(varCheckPropertyPresent) Then" & vbNewLine
'  strBaseString = strBaseString & "                pPropSet.RemoveProperty ""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdef""" & vbNewLine
'  strBaseString = strBaseString & "              End If" & vbNewLine
'  strBaseString = strBaseString & "              pPropSet.SetProperty ""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdef"", strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "              ' GET DESCRIPTION SOURCE FROM ORIGINAL FCLASS" & vbNewLine
'  strBaseString = strBaseString & "              varProperty = pTrails_KNF_UTM12_NAD83_PropSet.GetProperty(""eainfo/detailed/attr["" & CStr(lngIndex) & ""]/attrdefs"")" & vbNewLine
'  strBaseString = strBaseString & "              If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "                varSubProperty = varProperty(0)" & vbNewLine
'  strBaseString = strBaseString & "                If IsEmpty(varSubProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "                  strDescriptionSource = ""Trails_KNF_UTM12_NAD83""" & vbNewLine
'  strBaseString = strBaseString & "                Else" & vbNewLine
'  strBaseString = strBaseString & "                  strDescriptionSource = CStr(varSubProperty)" & vbNewLine
'  strBaseString = strBaseString & "                End If" & vbNewLine
'  strBaseString = strBaseString & "              Else" & vbNewLine
'  strBaseString = strBaseString & "                strDescriptionSource = ""Trails_KNF_UTM12_NAD83""" & vbNewLine
'  strBaseString = strBaseString & "              End If" & vbNewLine
'  strBaseString = strBaseString & "              strDescriptionSource = Replace(strDescriptionSource, ""Dicitonary"", ""Dictionary"", , , vbTextCompare)" & vbNewLine
'  strBaseString = strBaseString & "              strDescriptionSource = Replace(strDescriptionSource, ""20003"", ""2003"", , , vbTextCompare)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "              ' SET DESCRIPTION SOURCE DEFINITION OF NEW FCLASS" & vbNewLine
'  strBaseString = strBaseString & "              varCheckPropertyPresent = pPropSet.GetProperty(""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdefs"")" & vbNewLine
'  strBaseString = strBaseString & "              If Not IsEmpty(varCheckPropertyPresent) Then" & vbNewLine
'  strBaseString = strBaseString & "                pPropSet.RemoveProperty ""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdefs""" & vbNewLine
'  strBaseString = strBaseString & "              End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "              ' ONLY ADD THIS NEW ONE IF THERE WAS ONE IN THE ORIGINAL FCLASS" & vbNewLine
'  strBaseString = strBaseString & "              If strDescriptionSource <> """" Then" & vbNewLine
'  strBaseString = strBaseString & "                pPropSet.SetProperty ""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdefs"", strDescriptionSource" & vbNewLine
'  strBaseString = strBaseString & "              End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "              Exit For" & vbNewLine
'  strBaseString = strBaseString & "            End If" & vbNewLine
'  strBaseString = strBaseString & "          End If" & vbNewLine
'  strBaseString = strBaseString & "        Next lngIndex" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    Next lngKNFTR_Index" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "   " & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "    ' ----------------------------------------------------------------------------------------------" & vbNewLine
'  strBaseString = strBaseString & "    ' <<<<<<<<<<<<  Utah_Roads_Near_AOI_UTM12_NAD83  >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>" & vbNewLine
'  strBaseString = strBaseString & "    '-----------------------------------------------------------------------------------------------" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""UTAH_Add_Attribute"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = """"""TRUE"""" or """"FALSE"""", whether there was a polyline feature from the 'Utah_Roads_Near_AOI_UTM12_NAD83' "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""feature class within 100m of the centroid of this polyline.  If so, then this '"" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""Utah_Roads_Near_AOI_UTM12_NAD83' feature would be considered a candidate for extracting attribute values "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""to describe this feature.""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""UTAH_Attribute_Source"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""If 'UTAH_Add_Attribute' above = """"TRUE"""", then this field will contain the name of the "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""Utah Roads feature class (Utah_Roads_Near_AOI_UTM12_NAD83).  Otherwise it should be Null.""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""UTAH_Attribute_Distance"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""If 'UTAH_Add_Attribute' above = """"TRUE"""", then this field will contain the distance "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""(in meters) from the centroid of this polyline to the nearest Utah Roads feature.  Otherwise it should be Null.  "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""Values should always be <= 100m.""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.RemoveProperty strXDom" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomMin, ""0""" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomMax, ""100""" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomUnits, ""Meters""" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""UTAH_ObjectID"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""If 'UTAH_Add_Attribute' above = """"TRUE"""", then this field will contain the OBJECTID "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""value of the nearest Utah Roads feature, from the Utah_Roads_Near_AOI_UTM12_NAD83 feature class.""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "    For lngUtah_Index = 0 To UBound(strUtah_RoadArray, 1)" & vbNewLine
'  strBaseString = strBaseString & "      If strFieldName = strUtah_RoadArray(lngUtah_Index, 0) Then '  ""BLM_OBS_USE1"", ""BLM_OBS_USE2"" or ""BLM_OBS_USE3""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "        ' FIND ORIGINAL FIELD METADATA" & vbNewLine
'  strBaseString = strBaseString & "        For lngIndex = 0 To pUtah_Roads_Near_AOI_UTM12_NAD83_FClass.Fields.FieldCount" & vbNewLine
'  strBaseString = strBaseString & "          strNames(0) = ""eainfo/detailed/attr["" & CStr(lngIndex) & ""]""" & vbNewLine
'  strBaseString = strBaseString & "          pUtah_Roads_Near_AOI_UTM12_NAD83_PropSet.GetProperties strNames, varVals" & vbNewLine
'  strBaseString = strBaseString & "          varSubVals = varVals(0)" & vbNewLine
'  strBaseString = strBaseString & "          If Not IsEmpty(varSubVals) Then" & vbNewLine
'  strBaseString = strBaseString & "            ' GET FIELD NAME" & vbNewLine
'  strBaseString = strBaseString & "            strOrigFClassXName = ""eainfo/detailed/attr["" & CStr(lngIndex) & ""]/attrlabl""" & vbNewLine
'  strBaseString = strBaseString & "            varProperty = pUtah_Roads_Near_AOI_UTM12_NAD83_PropSet.GetProperty(strOrigFClassXName)" & vbNewLine
'  strBaseString = strBaseString & "            If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "              varSubProperty = varProperty(0)" & vbNewLine
'  strBaseString = strBaseString & "              If IsEmpty(varSubProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "                strOrigFieldName = ""<-- No Field Name -->""" & vbNewLine
'  strBaseString = strBaseString & "              Else" & vbNewLine
'  strBaseString = strBaseString & "                strOrigFieldName = CStr(varSubProperty)" & vbNewLine
'  strBaseString = strBaseString & "              End If" & vbNewLine
'  strBaseString = strBaseString & "            Else" & vbNewLine
'  strBaseString = strBaseString & "              strOrigFieldName = ""<-- No Field Name -->""" & vbNewLine
'  strBaseString = strBaseString & "            End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "            If strOrigFieldName = strUtah_RoadArray(lngUtah_Index, 1) Then ' ""OBS_USE1"", ""OBS_USE2"", or ""OBS_USE3""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "              ' CHECK IF DOMAIN INFO" & vbNewLine
'  strBaseString = strBaseString & "              strOrigFClassXName = ""eainfo/detailed/attr["" & CStr(lngIndex) & ""]/attrdomv""" & vbNewLine
'  strBaseString = strBaseString & "              varProperty = pUtah_Roads_Near_AOI_UTM12_NAD83_PropSet.GetProperty(strOrigFClassXName)" & vbNewLine
'  strBaseString = strBaseString & "              If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "                varSubProperty = varProperty(0)" & vbNewLine
'  strBaseString = strBaseString & "'                For lngIndex2 = 0 To UBound(varProperty)" & vbNewLine
'  strBaseString = strBaseString & "'                  Debug.Print ""       "" & CStr(lngIndex2) & ""] "" & CStr(varProperty(lngIndex2))" & vbNewLine
'  strBaseString = strBaseString & "'                Next lngIndex2" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "                ' GET DOMAIN ATTRIBUTES FROM ORIGINAL FCLASS" & vbNewLine
'  strBaseString = strBaseString & "                strOrigFClassXName = ""eainfo/detailed/attr["" & CStr(lngIndex) & ""]/attrdomv""" & vbNewLine
'  strBaseString = strBaseString & "                varProperty = pUtah_Roads_Near_AOI_UTM12_NAD83_PropSet.GetProperty(strOrigFClassXName)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "                ' ADD DOMAIN ATTRIBUTES TO NEW FCLASS" & vbNewLine
'  strBaseString = strBaseString & "                varCheckPropertyPresent = pPropSet.GetProperty(""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdomv"")" & vbNewLine
'  strBaseString = strBaseString & "                If Not IsEmpty(varCheckPropertyPresent) Then" & vbNewLine
'  strBaseString = strBaseString & "                  pPropSet.RemoveProperty ""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdomv""" & vbNewLine
'  strBaseString = strBaseString & "                End If" & vbNewLine
'  strBaseString = strBaseString & "                For lngIndex3 = 0 To UBound(varProperty)" & vbNewLine
'  strBaseString = strBaseString & "                  strSubXName1 = ""eainfo/detailed/attr["" & CStr(lngIndex) & ""]/attrdomv/"" & _" & vbNewLine
'  strBaseString = strBaseString & "                      varProperty(lngIndex3) & ""["" & CStr(lngIndex3) & ""]""" & vbNewLine
'  strBaseString = strBaseString & "                  varSubProp1 = pUtah_Roads_Near_AOI_UTM12_NAD83_PropSet.GetProperty(strSubXName1)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "                  If UBound(varSubProp1) = 0 Then" & vbNewLine
'  strBaseString = strBaseString & "                    strSubXName3 = ""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdomv/"" & _" & vbNewLine
'  strBaseString = strBaseString & "                        varProperty(lngIndex3) & ""["" & CStr(lngIndex3) & ""]""" & vbNewLine
'  strBaseString = strBaseString & "                    strValue = varSubProp1(0)" & vbNewLine
'  strBaseString = strBaseString & "                    strValue = Replace(strValue, ""20003"", ""2003"")" & vbNewLine
'  strBaseString = strBaseString & "                    strValue = Replace(strValue, ""Dicitonary"", ""Dictionary"")" & vbNewLine
'  strBaseString = strBaseString & "                    pPropSet.SetProperty strSubXName3, strValue" & vbNewLine
'  strBaseString = strBaseString & "                  Else" & vbNewLine
'  strBaseString = strBaseString & "                    For lngIndex4 = 0 To UBound(varSubProp1)" & vbNewLine
'  strBaseString = strBaseString & "                      strSubXName2 = ""eainfo/detailed/attr["" & CStr(lngIndex) & ""]/attrdomv/"" & _" & vbNewLine
'  strBaseString = strBaseString & "                          varProperty(lngIndex3) & ""["" & CStr(lngIndex3) & ""]/"" & varSubProp1(lngIndex4)" & vbNewLine
'  strBaseString = strBaseString & "                      varSubProp2 = pUtah_Roads_Near_AOI_UTM12_NAD83_PropSet.GetProperty(strSubXName2)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "                      ' WRITE VALUE BACK TO NEW FCLASS" & vbNewLine
'  strBaseString = strBaseString & "                      strSubXName3 = ""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdomv/"" & _" & vbNewLine
'  strBaseString = strBaseString & "                          varProperty(lngIndex3) & ""["" & CStr(lngIndex3) & ""]/"" & varSubProp1(lngIndex4)" & vbNewLine
'  strBaseString = strBaseString & "                      strValue = varSubProp2(0)" & vbNewLine
'  strBaseString = strBaseString & "                      strValue = Replace(strValue, ""20003"", ""2003"")" & vbNewLine
'  strBaseString = strBaseString & "                      strValue = Replace(strValue, ""Dicitonary"", ""Dictionary"")" & vbNewLine
'  strBaseString = strBaseString & "                      pPropSet.SetProperty strSubXName3, strValue" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "                    Next lngIndex4" & vbNewLine
'  strBaseString = strBaseString & "                  End If" & vbNewLine
'  strBaseString = strBaseString & "                Next lngIndex3" & vbNewLine
'  strBaseString = strBaseString & "              Else" & vbNewLine
'  strBaseString = strBaseString & "'                strOrigFCLassXName = ""eainfo/detailed/attr["" & CStr(lngIndex) & ""]""" & vbNewLine
'  strBaseString = strBaseString & "'                varProperty = pUtah_Roads_Near_AOI_UTM12_NAD83_PropSet.GetProperty(strOrigFCLassXName)" & vbNewLine
'  strBaseString = strBaseString & "'                varSubProperty = varProperty(0)" & vbNewLine
'  strBaseString = strBaseString & "'                Debug.Print ""Examining "" & strOrigFieldName & ""...""" & vbNewLine
'  strBaseString = strBaseString & "'                For lngIndex2 = 0 To UBound(varProperty)" & vbNewLine
'  strBaseString = strBaseString & "'                  varSubProp1 = pUtah_Roads_Near_AOI_UTM12_NAD83_PropSet.GetProperty(strOrigFCLassXName & ""/"" & _" & vbNewLine
'  strBaseString = strBaseString & "'                        CStr(varProperty(lngIndex2)))" & vbNewLine
'  strBaseString = strBaseString & "'                  Debug.Print ""       "" & CStr(lngIndex2) & ""] "" & CStr(varProperty(lngIndex2)) & "" = "" & _" & vbNewLine
'  strBaseString = strBaseString & "'                        CStr(varSubProp1(0))" & vbNewLine
'  strBaseString = strBaseString & "'                Next lngIndex2" & vbNewLine
'  strBaseString = strBaseString & "              End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "              If strOrigFieldName = ""ROAD_NO_"" Then" & vbNewLine
'  strBaseString = strBaseString & "                Debug.Print ""Here...""" & vbNewLine
'  strBaseString = strBaseString & "              End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "              ' GET DEFINITION FROM ORIGINAL FCLASS" & vbNewLine
'  strBaseString = strBaseString & "              varProperty = pUtah_Roads_Near_AOI_UTM12_NAD83_PropSet.GetProperty(""eainfo/detailed/attr["" & CStr(lngIndex) & ""]/attrdef"")" & vbNewLine
'  strBaseString = strBaseString & "              If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "                varSubProperty = varProperty(0)" & vbNewLine
'  strBaseString = strBaseString & "                If IsEmpty(varSubProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "                  strDescription = ""<-- No Description -->""" & vbNewLine
'  strBaseString = strBaseString & "                Else" & vbNewLine
'  strBaseString = strBaseString & "                  strDescription = CStr(varSubProperty)" & vbNewLine
'  strBaseString = strBaseString & "                  strDescription = Trim(strDescription)" & vbNewLine
'  strBaseString = strBaseString & "                  If Right(strDescription, 1) <> ""."" And Right(strDescription, 1) <> ""!"" And _" & vbNewLine
'  strBaseString = strBaseString & "                      Right(strDescription, 1) <> ""?"" Then" & vbNewLine
'  strBaseString = strBaseString & "                    strDescription = strDescription & "".""" & vbNewLine
'  strBaseString = strBaseString & "                  End If" & vbNewLine
'  strBaseString = strBaseString & "                End If" & vbNewLine
'  strBaseString = strBaseString & "              Else" & vbNewLine
'  strBaseString = strBaseString & "                strDescription = ""<-- No Description -->""" & vbNewLine
'  strBaseString = strBaseString & "              End If" & vbNewLine
'  strBaseString = strBaseString & "              strDescription = Replace(strDescription, ""  [Imported from TransportationLine_UTM12_NAD83]"", """", , , vbTextCompare)" & vbNewLine
'  strBaseString = strBaseString & "              strDescription = strDescription & _" & vbNewLine
'  strBaseString = strBaseString & "                  ""  This field should only have a value if the 'UTAH_Add_Attribute' value = """"TRUE"""".  "" & _" & vbNewLine
'  strBaseString = strBaseString & "                  ""[Imported from Utah_Roads_Near_AOI_UTM12_NAD83, Field """""" & strUtah_RoadArray(lngUtah_Index, 1) & """"""]""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "              ' SET FIELD DEFINITION OF NEW FCLASS" & vbNewLine
'  strBaseString = strBaseString & "              varCheckPropertyPresent = pPropSet.GetProperty(""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdef"")" & vbNewLine
'  strBaseString = strBaseString & "              If Not IsEmpty(varCheckPropertyPresent) Then" & vbNewLine
'  strBaseString = strBaseString & "                pPropSet.RemoveProperty ""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdef""" & vbNewLine
'  strBaseString = strBaseString & "              End If" & vbNewLine
'  strBaseString = strBaseString & "              pPropSet.SetProperty ""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdef"", strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "              ' GET DESCRIPTION SOURCE FROM ORIGINAL FCLASS" & vbNewLine
'  strBaseString = strBaseString & "              varProperty = pUtah_Roads_Near_AOI_UTM12_NAD83_PropSet.GetProperty(""eainfo/detailed/attr["" & CStr(lngIndex) & ""]/attrdefs"")" & vbNewLine
'  strBaseString = strBaseString & "              If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "                varSubProperty = varProperty(0)" & vbNewLine
'  strBaseString = strBaseString & "                If IsEmpty(varSubProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "                  strDescriptionSource = ""Utah_Roads_Near_AOI_UTM12_NAD83""" & vbNewLine
'  strBaseString = strBaseString & "                Else" & vbNewLine
'  strBaseString = strBaseString & "                  strDescriptionSource = CStr(varSubProperty)" & vbNewLine
'  strBaseString = strBaseString & "                End If" & vbNewLine
'  strBaseString = strBaseString & "              Else" & vbNewLine
'  strBaseString = strBaseString & "                strDescriptionSource = ""Utah_Roads_Near_AOI_UTM12_NAD83""" & vbNewLine
'  strBaseString = strBaseString & "              End If" & vbNewLine
'  strBaseString = strBaseString & "              strDescriptionSource = Replace(strDescriptionSource, ""Dicitonary"", ""Dictionary"", , , vbTextCompare)" & vbNewLine
'  strBaseString = strBaseString & "              strDescriptionSource = Replace(strDescriptionSource, ""20003"", ""2003"", , , vbTextCompare)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "              ' SET DESCRIPTION SOURCE DEFINITION OF NEW FCLASS" & vbNewLine
'  strBaseString = strBaseString & "              varCheckPropertyPresent = pPropSet.GetProperty(""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdefs"")" & vbNewLine
'  strBaseString = strBaseString & "              If Not IsEmpty(varCheckPropertyPresent) Then" & vbNewLine
'  strBaseString = strBaseString & "                pPropSet.RemoveProperty ""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdefs""" & vbNewLine
'  strBaseString = strBaseString & "              End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "              ' ONLY ADD THIS NEW ONE IF THERE WAS ONE IN THE ORIGINAL FCLASS" & vbNewLine
'  strBaseString = strBaseString & "              If strDescriptionSource <> """" Then" & vbNewLine
'  strBaseString = strBaseString & "                pPropSet.SetProperty ""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdefs"", strDescriptionSource" & vbNewLine
'  strBaseString = strBaseString & "              End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "              Exit For" & vbNewLine
'  strBaseString = strBaseString & "            End If" & vbNewLine
'  strBaseString = strBaseString & "          End If" & vbNewLine
'  strBaseString = strBaseString & "        Next lngIndex" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    Next lngUtah_Index" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "    ' ----------------------------------------------------------------------------------------------" & vbNewLine
'  strBaseString = strBaseString & "    ' <<<<<<<<<<<<  Routes_UTM12_NAD83  >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>" & vbNewLine
'  strBaseString = strBaseString & "    '-----------------------------------------------------------------------------------------------" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""ROUTES_Add_Attribute"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = """"""TRUE"""" or """"FALSE"""", whether there was a polyline feature from the 'Routes_UTM12_NAD83' "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""feature class within 100m of the centroid of this polyline.  If so, then this '"" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""Routes_UTM12_NAD83' feature would be considered a candidate for extracting attribute values "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""to describe this feature.""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""ROUTES_Attribute_Source"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""If 'ROUTES_Add_Attribute' above = """"TRUE"""", then this field will contain the name of the "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""GCNP Routes feature class (Routes_UTM12_NAD83).  Otherwise it should be Null.""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""ROUTES_Attribute_Distance"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""If 'ROUTES_Add_Attribute' above = """"TRUE"""", then this field will contain the distance "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""(in meters) from the centroid of this polyline to the nearest GCNP Routes feature.  Otherwise it should be Null.  "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""Values should always be <= 100m.""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.RemoveProperty strXDom" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomMin, ""0""" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomMax, ""100""" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomUnits, ""Meters""" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""ROUTES_ObjectID"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""If 'ROUTES_Add_Attribute' above = """"TRUE"""", then this field will contain the OBJECTID "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""value of the nearest GCNP Routes feature, from the Routes_UTM12_NAD83 feature class.""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "    For lngRoutes_Index = 0 To UBound(strRoutes_RoadArray, 1)" & vbNewLine
'  strBaseString = strBaseString & "      If strFieldName = strRoutes_RoadArray(lngRoutes_Index, 0) Then '  ""BLM_OBS_USE1"", ""BLM_OBS_USE2"" or ""BLM_OBS_USE3""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "        ' FIND ORIGINAL FIELD METADATA" & vbNewLine
'  strBaseString = strBaseString & "        For lngIndex = 0 To pRoutes_UTM12_NAD83_FClass.Fields.FieldCount" & vbNewLine
'  strBaseString = strBaseString & "          strNames(0) = ""eainfo/detailed/attr["" & CStr(lngIndex) & ""]""" & vbNewLine
'  strBaseString = strBaseString & "          pRoutes_UTM12_NAD83_PropSet.GetProperties strNames, varVals" & vbNewLine
'  strBaseString = strBaseString & "          varSubVals = varVals(0)" & vbNewLine
'  strBaseString = strBaseString & "          If Not IsEmpty(varSubVals) Then" & vbNewLine
'  strBaseString = strBaseString & "            ' GET FIELD NAME" & vbNewLine
'  strBaseString = strBaseString & "            strOrigFClassXName = ""eainfo/detailed/attr["" & CStr(lngIndex) & ""]/attrlabl""" & vbNewLine
'  strBaseString = strBaseString & "            varProperty = pRoutes_UTM12_NAD83_PropSet.GetProperty(strOrigFClassXName)" & vbNewLine
'  strBaseString = strBaseString & "            If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "              varSubProperty = varProperty(0)" & vbNewLine
'  strBaseString = strBaseString & "              If IsEmpty(varSubProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "                strOrigFieldName = ""<-- No Field Name -->""" & vbNewLine
'  strBaseString = strBaseString & "              Else" & vbNewLine
'  strBaseString = strBaseString & "                strOrigFieldName = CStr(varSubProperty)" & vbNewLine
'  strBaseString = strBaseString & "              End If" & vbNewLine
'  strBaseString = strBaseString & "            Else" & vbNewLine
'  strBaseString = strBaseString & "              strOrigFieldName = ""<-- No Field Name -->""" & vbNewLine
'  strBaseString = strBaseString & "            End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "            If strOrigFieldName = strRoutes_RoadArray(lngRoutes_Index, 1) Then ' ""OBS_USE1"", ""OBS_USE2"", or ""OBS_USE3""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "              ' CHECK IF DOMAIN INFO" & vbNewLine
'  strBaseString = strBaseString & "              strOrigFClassXName = ""eainfo/detailed/attr["" & CStr(lngIndex) & ""]/attrdomv""" & vbNewLine
'  strBaseString = strBaseString & "              varProperty = pRoutes_UTM12_NAD83_PropSet.GetProperty(strOrigFClassXName)" & vbNewLine
'  strBaseString = strBaseString & "              If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "                varSubProperty = varProperty(0)" & vbNewLine
'  strBaseString = strBaseString & "'                For lngIndex2 = 0 To UBound(varProperty)" & vbNewLine
'  strBaseString = strBaseString & "'                  Debug.Print ""       "" & CStr(lngIndex2) & ""] "" & CStr(varProperty(lngIndex2))" & vbNewLine
'  strBaseString = strBaseString & "'                Next lngIndex2" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "                ' GET DOMAIN ATTRIBUTES FROM ORIGINAL FCLASS" & vbNewLine
'  strBaseString = strBaseString & "                strOrigFClassXName = ""eainfo/detailed/attr["" & CStr(lngIndex) & ""]/attrdomv""" & vbNewLine
'  strBaseString = strBaseString & "                varProperty = pRoutes_UTM12_NAD83_PropSet.GetProperty(strOrigFClassXName)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "                ' ADD DOMAIN ATTRIBUTES TO NEW FCLASS" & vbNewLine
'  strBaseString = strBaseString & "                varCheckPropertyPresent = pPropSet.GetProperty(""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdomv"")" & vbNewLine
'  strBaseString = strBaseString & "                If Not IsEmpty(varCheckPropertyPresent) Then" & vbNewLine
'  strBaseString = strBaseString & "                  pPropSet.RemoveProperty ""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdomv""" & vbNewLine
'  strBaseString = strBaseString & "                End If" & vbNewLine
'  strBaseString = strBaseString & "                For lngIndex3 = 0 To UBound(varProperty)" & vbNewLine
'  strBaseString = strBaseString & "                  strSubXName1 = ""eainfo/detailed/attr["" & CStr(lngIndex) & ""]/attrdomv/"" & _" & vbNewLine
'  strBaseString = strBaseString & "                      varProperty(lngIndex3) & ""["" & CStr(lngIndex3) & ""]""" & vbNewLine
'  strBaseString = strBaseString & "                  varSubProp1 = pRoutes_UTM12_NAD83_PropSet.GetProperty(strSubXName1)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "                  If UBound(varSubProp1) = 0 Then" & vbNewLine
'  strBaseString = strBaseString & "                    strSubXName3 = ""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdomv/"" & _" & vbNewLine
'  strBaseString = strBaseString & "                        varProperty(lngIndex3) & ""["" & CStr(lngIndex3) & ""]""" & vbNewLine
'  strBaseString = strBaseString & "                    strValue = varSubProp1(0)" & vbNewLine
'  strBaseString = strBaseString & "                    strValue = Replace(strValue, ""20003"", ""2003"")" & vbNewLine
'  strBaseString = strBaseString & "                    strValue = Replace(strValue, ""Dicitonary"", ""Dictionary"")" & vbNewLine
'  strBaseString = strBaseString & "                    pPropSet.SetProperty strSubXName3, strValue" & vbNewLine
'  strBaseString = strBaseString & "                  Else" & vbNewLine
'  strBaseString = strBaseString & "                    For lngIndex4 = 0 To UBound(varSubProp1)" & vbNewLine
'  strBaseString = strBaseString & "                      strSubXName2 = ""eainfo/detailed/attr["" & CStr(lngIndex) & ""]/attrdomv/"" & _" & vbNewLine
'  strBaseString = strBaseString & "                          varProperty(lngIndex3) & ""["" & CStr(lngIndex3) & ""]/"" & varSubProp1(lngIndex4)" & vbNewLine
'  strBaseString = strBaseString & "                      varSubProp2 = pRoutes_UTM12_NAD83_PropSet.GetProperty(strSubXName2)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "                      ' WRITE VALUE BACK TO NEW FCLASS" & vbNewLine
'  strBaseString = strBaseString & "                      strSubXName3 = ""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdomv/"" & _" & vbNewLine
'  strBaseString = strBaseString & "                          varProperty(lngIndex3) & ""["" & CStr(lngIndex3) & ""]/"" & varSubProp1(lngIndex4)" & vbNewLine
'  strBaseString = strBaseString & "                      strValue = varSubProp2(0)" & vbNewLine
'  strBaseString = strBaseString & "                      strValue = Replace(strValue, ""20003"", ""2003"")" & vbNewLine
'  strBaseString = strBaseString & "                      strValue = Replace(strValue, ""Dicitonary"", ""Dictionary"")" & vbNewLine
'  strBaseString = strBaseString & "                      pPropSet.SetProperty strSubXName3, strValue" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "                    Next lngIndex4" & vbNewLine
'  strBaseString = strBaseString & "                  End If" & vbNewLine
'  strBaseString = strBaseString & "                Next lngIndex3" & vbNewLine
'  strBaseString = strBaseString & "              Else" & vbNewLine
'  strBaseString = strBaseString & "'                strOrigFCLassXName = ""eainfo/detailed/attr["" & CStr(lngIndex) & ""]""" & vbNewLine
'  strBaseString = strBaseString & "'                varProperty = pRoutes_UTM12_NAD83_PropSet.GetProperty(strOrigFCLassXName)" & vbNewLine
'  strBaseString = strBaseString & "'                varSubProperty = varProperty(0)" & vbNewLine
'  strBaseString = strBaseString & "'                Debug.Print ""Examining "" & strOrigFieldName & ""...""" & vbNewLine
'  strBaseString = strBaseString & "'                For lngIndex2 = 0 To UBound(varProperty)" & vbNewLine
'  strBaseString = strBaseString & "'                  varSubProp1 = pRoutes_UTM12_NAD83_PropSet.GetProperty(strOrigFCLassXName & ""/"" & _" & vbNewLine
'  strBaseString = strBaseString & "'                        CStr(varProperty(lngIndex2)))" & vbNewLine
'  strBaseString = strBaseString & "'                  Debug.Print ""       "" & CStr(lngIndex2) & ""] "" & CStr(varProperty(lngIndex2)) & "" = "" & _" & vbNewLine
'  strBaseString = strBaseString & "'                        CStr(varSubProp1(0))" & vbNewLine
'  strBaseString = strBaseString & "'                Next lngIndex2" & vbNewLine
'  strBaseString = strBaseString & "              End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "              If strOrigFieldName = ""ROAD_NO_"" Then" & vbNewLine
'  strBaseString = strBaseString & "                Debug.Print ""Here...""" & vbNewLine
'  strBaseString = strBaseString & "              End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "              ' GET DEFINITION FROM ORIGINAL FCLASS" & vbNewLine
'  strBaseString = strBaseString & "              varProperty = pRoutes_UTM12_NAD83_PropSet.GetProperty(""eainfo/detailed/attr["" & CStr(lngIndex) & ""]/attrdef"")" & vbNewLine
'  strBaseString = strBaseString & "              If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "                varSubProperty = varProperty(0)" & vbNewLine
'  strBaseString = strBaseString & "                If IsEmpty(varSubProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "                  strDescription = ""<-- No Description -->""" & vbNewLine
'  strBaseString = strBaseString & "                Else" & vbNewLine
'  strBaseString = strBaseString & "                  strDescription = CStr(varSubProperty)" & vbNewLine
'  strBaseString = strBaseString & "                  strDescription = Trim(strDescription)" & vbNewLine
'  strBaseString = strBaseString & "                  If Right(strDescription, 1) <> ""."" And Right(strDescription, 1) <> ""!"" And _" & vbNewLine
'  strBaseString = strBaseString & "                      Right(strDescription, 1) <> ""?"" Then" & vbNewLine
'  strBaseString = strBaseString & "                    strDescription = strDescription & "".""" & vbNewLine
'  strBaseString = strBaseString & "                  End If" & vbNewLine
'  strBaseString = strBaseString & "                End If" & vbNewLine
'  strBaseString = strBaseString & "              Else" & vbNewLine
'  strBaseString = strBaseString & "                strDescription = ""<-- No Description -->""" & vbNewLine
'  strBaseString = strBaseString & "              End If" & vbNewLine
'  strBaseString = strBaseString & "              strDescription = Replace(strDescription, ""  [Imported from TransportationLine_UTM12_NAD83]"", """", , , vbTextCompare)" & vbNewLine
'  strBaseString = strBaseString & "              strDescription = strDescription & _" & vbNewLine
'  strBaseString = strBaseString & "                  ""  This field should only have a value if the 'ROUTES_Add_Attribute' value = """"TRUE"""".  "" & _" & vbNewLine
'  strBaseString = strBaseString & "                  ""[Imported from Routes_UTM12_NAD83, Field """""" & strRoutes_RoadArray(lngRoutes_Index, 1) & """"""]""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "              ' SET FIELD DEFINITION OF NEW FCLASS" & vbNewLine
'  strBaseString = strBaseString & "              varCheckPropertyPresent = pPropSet.GetProperty(""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdef"")" & vbNewLine
'  strBaseString = strBaseString & "              If Not IsEmpty(varCheckPropertyPresent) Then" & vbNewLine
'  strBaseString = strBaseString & "                pPropSet.RemoveProperty ""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdef""" & vbNewLine
'  strBaseString = strBaseString & "              End If" & vbNewLine
'  strBaseString = strBaseString & "              pPropSet.SetProperty ""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdef"", strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "              ' GET DESCRIPTION SOURCE FROM ORIGINAL FCLASS" & vbNewLine
'  strBaseString = strBaseString & "              varProperty = pRoutes_UTM12_NAD83_PropSet.GetProperty(""eainfo/detailed/attr["" & CStr(lngIndex) & ""]/attrdefs"")" & vbNewLine
'  strBaseString = strBaseString & "              If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "                varSubProperty = varProperty(0)" & vbNewLine
'  strBaseString = strBaseString & "                If IsEmpty(varSubProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "                  strDescriptionSource = ""Routes_UTM12_NAD83""" & vbNewLine
'  strBaseString = strBaseString & "                Else" & vbNewLine
'  strBaseString = strBaseString & "                  strDescriptionSource = CStr(varSubProperty)" & vbNewLine
'  strBaseString = strBaseString & "                End If" & vbNewLine
'  strBaseString = strBaseString & "              Else" & vbNewLine
'  strBaseString = strBaseString & "                strDescriptionSource = ""Routes_UTM12_NAD83""" & vbNewLine
'  strBaseString = strBaseString & "              End If" & vbNewLine
'  strBaseString = strBaseString & "              strDescriptionSource = Replace(strDescriptionSource, ""Dicitonary"", ""Dictionary"", , , vbTextCompare)" & vbNewLine
'  strBaseString = strBaseString & "              strDescriptionSource = Replace(strDescriptionSource, ""20003"", ""2003"", , , vbTextCompare)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "              ' SET DESCRIPTION SOURCE DEFINITION OF NEW FCLASS" & vbNewLine
'  strBaseString = strBaseString & "              varCheckPropertyPresent = pPropSet.GetProperty(""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdefs"")" & vbNewLine
'  strBaseString = strBaseString & "              If Not IsEmpty(varCheckPropertyPresent) Then" & vbNewLine
'  strBaseString = strBaseString & "                pPropSet.RemoveProperty ""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdefs""" & vbNewLine
'  strBaseString = strBaseString & "              End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "              ' ONLY ADD THIS NEW ONE IF THERE WAS ONE IN THE ORIGINAL FCLASS" & vbNewLine
'  strBaseString = strBaseString & "              If strDescriptionSource <> """" Then" & vbNewLine
'  strBaseString = strBaseString & "                pPropSet.SetProperty ""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdefs"", strDescriptionSource" & vbNewLine
'  strBaseString = strBaseString & "              End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "              Exit For" & vbNewLine
'  strBaseString = strBaseString & "            End If" & vbNewLine
'  strBaseString = strBaseString & "          End If" & vbNewLine
'  strBaseString = strBaseString & "        Next lngIndex" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    Next lngRoutes_Index" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "    ' ----------------------------------------------------------------------------------------------" & vbNewLine
'  strBaseString = strBaseString & "    ' <<<<<<<<<<<<  Trails_UTM12_NAD83  >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>" & vbNewLine
'  strBaseString = strBaseString & "    '-----------------------------------------------------------------------------------------------" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""TRAILS_Add_Attribute"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = """"""TRUE"""" or """"FALSE"""", whether there was a polyline feature from the 'Trails_UTM12_NAD83' "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""feature class within 100m of the centroid of this polyline.  If so, then this '"" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""Trails_UTM12_NAD83' feature would be considered a candidate for extracting attribute values "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""to describe this feature.""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""TRAILS_Attribute_Source"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""If 'TRAILS_Add_Attribute' above = """"TRUE"""", then this field will contain the name of the "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""GCNP Trails feature class (Trails_UTM12_NAD83).  Otherwise it should be Null.""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""TRAILS_Attribute_Distance"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""If 'TRAILS_Add_Attribute' above = """"TRUE"""", then this field will contain the distance "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""(in meters) from the centroid of this polyline to the nearest GCNP Trails feature.  Otherwise it should be Null.  "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""Values should always be <= 100m.""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.RemoveProperty strXDom" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomMin, ""0""" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomMax, ""100""" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomUnits, ""Meters""" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""TRAILS_ObjectID"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""If 'TRAILS_Add_Attribute' above = """"TRUE"""", then this field will contain the OBJECTID "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""value of the nearest GCNP Trails feature, from the Trails_UTM12_NAD83 feature class.""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "    For lngTrails_Index = 0 To UBound(strTrails_RoadArray, 1)" & vbNewLine
'  strBaseString = strBaseString & "      If strFieldName = strTrails_RoadArray(lngTrails_Index, 0) Then '  ""BLM_OBS_USE1"", ""BLM_OBS_USE2"" or ""BLM_OBS_USE3""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "        ' FIND ORIGINAL FIELD METADATA" & vbNewLine
'  strBaseString = strBaseString & "        For lngIndex = 0 To pTrails_UTM12_NAD83_FClass.Fields.FieldCount" & vbNewLine
'  strBaseString = strBaseString & "          strNames(0) = ""eainfo/detailed/attr["" & CStr(lngIndex) & ""]""" & vbNewLine
'  strBaseString = strBaseString & "          pTrails_UTM12_NAD83_PropSet.GetProperties strNames, varVals" & vbNewLine
'  strBaseString = strBaseString & "          varSubVals = varVals(0)" & vbNewLine
'  strBaseString = strBaseString & "          If Not IsEmpty(varSubVals) Then" & vbNewLine
'  strBaseString = strBaseString & "            ' GET FIELD NAME" & vbNewLine
'  strBaseString = strBaseString & "            strOrigFClassXName = ""eainfo/detailed/attr["" & CStr(lngIndex) & ""]/attrlabl""" & vbNewLine
'  strBaseString = strBaseString & "            varProperty = pTrails_UTM12_NAD83_PropSet.GetProperty(strOrigFClassXName)" & vbNewLine
'  strBaseString = strBaseString & "            If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "              varSubProperty = varProperty(0)" & vbNewLine
'  strBaseString = strBaseString & "              If IsEmpty(varSubProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "                strOrigFieldName = ""<-- No Field Name -->""" & vbNewLine
'  strBaseString = strBaseString & "              Else" & vbNewLine
'  strBaseString = strBaseString & "                strOrigFieldName = CStr(varSubProperty)" & vbNewLine
'  strBaseString = strBaseString & "              End If" & vbNewLine
'  strBaseString = strBaseString & "            Else" & vbNewLine
'  strBaseString = strBaseString & "              strOrigFieldName = ""<-- No Field Name -->""" & vbNewLine
'  strBaseString = strBaseString & "            End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "            If strOrigFieldName = strTrails_RoadArray(lngTrails_Index, 1) Then ' ""OBS_USE1"", ""OBS_USE2"", or ""OBS_USE3""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "              ' CHECK IF DOMAIN INFO" & vbNewLine
'  strBaseString = strBaseString & "              strOrigFClassXName = ""eainfo/detailed/attr["" & CStr(lngIndex) & ""]/attrdomv""" & vbNewLine
'  strBaseString = strBaseString & "              varProperty = pTrails_UTM12_NAD83_PropSet.GetProperty(strOrigFClassXName)" & vbNewLine
'  strBaseString = strBaseString & "              If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "                varSubProperty = varProperty(0)" & vbNewLine
'  strBaseString = strBaseString & "'                For lngIndex2 = 0 To UBound(varProperty)" & vbNewLine
'  strBaseString = strBaseString & "'                  Debug.Print ""       "" & CStr(lngIndex2) & ""] "" & CStr(varProperty(lngIndex2))" & vbNewLine
'  strBaseString = strBaseString & "'                Next lngIndex2" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "                ' GET DOMAIN ATTRIBUTES FROM ORIGINAL FCLASS" & vbNewLine
'  strBaseString = strBaseString & "                strOrigFClassXName = ""eainfo/detailed/attr["" & CStr(lngIndex) & ""]/attrdomv""" & vbNewLine
'  strBaseString = strBaseString & "                varProperty = pTrails_UTM12_NAD83_PropSet.GetProperty(strOrigFClassXName)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "                ' ADD DOMAIN ATTRIBUTES TO NEW FCLASS" & vbNewLine
'  strBaseString = strBaseString & "                varCheckPropertyPresent = pPropSet.GetProperty(""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdomv"")" & vbNewLine
'  strBaseString = strBaseString & "                If Not IsEmpty(varCheckPropertyPresent) Then" & vbNewLine
'  strBaseString = strBaseString & "                  pPropSet.RemoveProperty ""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdomv""" & vbNewLine
'  strBaseString = strBaseString & "                End If" & vbNewLine
'  strBaseString = strBaseString & "                For lngIndex3 = 0 To UBound(varProperty)" & vbNewLine
'  strBaseString = strBaseString & "                  strSubXName1 = ""eainfo/detailed/attr["" & CStr(lngIndex) & ""]/attrdomv/"" & _" & vbNewLine
'  strBaseString = strBaseString & "                      varProperty(lngIndex3) & ""["" & CStr(lngIndex3) & ""]""" & vbNewLine
'  strBaseString = strBaseString & "                  varSubProp1 = pTrails_UTM12_NAD83_PropSet.GetProperty(strSubXName1)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "                  If UBound(varSubProp1) = 0 Then" & vbNewLine
'  strBaseString = strBaseString & "                    strSubXName3 = ""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdomv/"" & _" & vbNewLine
'  strBaseString = strBaseString & "                        varProperty(lngIndex3) & ""["" & CStr(lngIndex3) & ""]""" & vbNewLine
'  strBaseString = strBaseString & "                    strValue = varSubProp1(0)" & vbNewLine
'  strBaseString = strBaseString & "                    strValue = Replace(strValue, ""20003"", ""2003"")" & vbNewLine
'  strBaseString = strBaseString & "                    strValue = Replace(strValue, ""Dicitonary"", ""Dictionary"")" & vbNewLine
'  strBaseString = strBaseString & "                    pPropSet.SetProperty strSubXName3, strValue" & vbNewLine
'  strBaseString = strBaseString & "                  Else" & vbNewLine
'  strBaseString = strBaseString & "                    For lngIndex4 = 0 To UBound(varSubProp1)" & vbNewLine
'  strBaseString = strBaseString & "                      strSubXName2 = ""eainfo/detailed/attr["" & CStr(lngIndex) & ""]/attrdomv/"" & _" & vbNewLine
'  strBaseString = strBaseString & "                          varProperty(lngIndex3) & ""["" & CStr(lngIndex3) & ""]/"" & varSubProp1(lngIndex4)" & vbNewLine
'  strBaseString = strBaseString & "                      varSubProp2 = pTrails_UTM12_NAD83_PropSet.GetProperty(strSubXName2)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "                      ' WRITE VALUE BACK TO NEW FCLASS" & vbNewLine
'  strBaseString = strBaseString & "                      strSubXName3 = ""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdomv/"" & _" & vbNewLine
'  strBaseString = strBaseString & "                          varProperty(lngIndex3) & ""["" & CStr(lngIndex3) & ""]/"" & varSubProp1(lngIndex4)" & vbNewLine
'  strBaseString = strBaseString & "                      strValue = varSubProp2(0)" & vbNewLine
'  strBaseString = strBaseString & "                      strValue = Replace(strValue, ""20003"", ""2003"")" & vbNewLine
'  strBaseString = strBaseString & "                      strValue = Replace(strValue, ""Dicitonary"", ""Dictionary"")" & vbNewLine
'  strBaseString = strBaseString & "                      pPropSet.SetProperty strSubXName3, strValue" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "                    Next lngIndex4" & vbNewLine
'  strBaseString = strBaseString & "                  End If" & vbNewLine
'  strBaseString = strBaseString & "                Next lngIndex3" & vbNewLine
'  strBaseString = strBaseString & "              Else" & vbNewLine
'  strBaseString = strBaseString & "'                strOrigFCLassXName = ""eainfo/detailed/attr["" & CStr(lngIndex) & ""]""" & vbNewLine
'  strBaseString = strBaseString & "'                varProperty = pTrails_UTM12_NAD83_PropSet.GetProperty(strOrigFCLassXName)" & vbNewLine
'  strBaseString = strBaseString & "'                varSubProperty = varProperty(0)" & vbNewLine
'  strBaseString = strBaseString & "'                Debug.Print ""Examining "" & strOrigFieldName & ""...""" & vbNewLine
'  strBaseString = strBaseString & "'                For lngIndex2 = 0 To UBound(varProperty)" & vbNewLine
'  strBaseString = strBaseString & "'                  varSubProp1 = pTrails_UTM12_NAD83_PropSet.GetProperty(strOrigFCLassXName & ""/"" & _" & vbNewLine
'  strBaseString = strBaseString & "'                        CStr(varProperty(lngIndex2)))" & vbNewLine
'  strBaseString = strBaseString & "'                  Debug.Print ""       "" & CStr(lngIndex2) & ""] "" & CStr(varProperty(lngIndex2)) & "" = "" & _" & vbNewLine
'  strBaseString = strBaseString & "'                        CStr(varSubProp1(0))" & vbNewLine
'  strBaseString = strBaseString & "'                Next lngIndex2" & vbNewLine
'  strBaseString = strBaseString & "              End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "              If strOrigFieldName = ""ROAD_NO_"" Then" & vbNewLine
'  strBaseString = strBaseString & "                Debug.Print ""Here...""" & vbNewLine
'  strBaseString = strBaseString & "              End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "              ' GET DEFINITION FROM ORIGINAL FCLASS" & vbNewLine
'  strBaseString = strBaseString & "              varProperty = pTrails_UTM12_NAD83_PropSet.GetProperty(""eainfo/detailed/attr["" & CStr(lngIndex) & ""]/attrdef"")" & vbNewLine
'  strBaseString = strBaseString & "              If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "                varSubProperty = varProperty(0)" & vbNewLine
'  strBaseString = strBaseString & "                If IsEmpty(varSubProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "                  strDescription = ""<-- No Description -->""" & vbNewLine
'  strBaseString = strBaseString & "                Else" & vbNewLine
'  strBaseString = strBaseString & "                  strDescription = CStr(varSubProperty)" & vbNewLine
'  strBaseString = strBaseString & "                  strDescription = Trim(strDescription)" & vbNewLine
'  strBaseString = strBaseString & "                  If Right(strDescription, 1) <> ""."" And Right(strDescription, 1) <> ""!"" And _" & vbNewLine
'  strBaseString = strBaseString & "                      Right(strDescription, 1) <> ""?"" Then" & vbNewLine
'  strBaseString = strBaseString & "                    strDescription = strDescription & "".""" & vbNewLine
'  strBaseString = strBaseString & "                  End If" & vbNewLine
'  strBaseString = strBaseString & "                End If" & vbNewLine
'  strBaseString = strBaseString & "              Else" & vbNewLine
'  strBaseString = strBaseString & "                strDescription = ""<-- No Description -->""" & vbNewLine
'  strBaseString = strBaseString & "              End If" & vbNewLine
'  strBaseString = strBaseString & "              strDescription = Replace(strDescription, ""  [Imported from TransportationLine_UTM12_NAD83]"", """", , , vbTextCompare)" & vbNewLine
'  strBaseString = strBaseString & "              strDescription = strDescription & _" & vbNewLine
'  strBaseString = strBaseString & "                  ""  This field should only have a value if the 'TRAILS_Add_Attribute' value = """"TRUE"""".  "" & _" & vbNewLine
'  strBaseString = strBaseString & "                  ""[Imported from Trails_UTM12_NAD83, Field """""" & strTrails_RoadArray(lngTrails_Index, 1) & """"""]""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "              ' SET FIELD DEFINITION OF NEW FCLASS" & vbNewLine
'  strBaseString = strBaseString & "              varCheckPropertyPresent = pPropSet.GetProperty(""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdef"")" & vbNewLine
'  strBaseString = strBaseString & "              If Not IsEmpty(varCheckPropertyPresent) Then" & vbNewLine
'  strBaseString = strBaseString & "                pPropSet.RemoveProperty ""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdef""" & vbNewLine
'  strBaseString = strBaseString & "              End If" & vbNewLine
'  strBaseString = strBaseString & "              pPropSet.SetProperty ""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdef"", strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "              ' GET DESCRIPTION SOURCE FROM ORIGINAL FCLASS" & vbNewLine
'  strBaseString = strBaseString & "              varProperty = pTrails_UTM12_NAD83_PropSet.GetProperty(""eainfo/detailed/attr["" & CStr(lngIndex) & ""]/attrdefs"")" & vbNewLine
'  strBaseString = strBaseString & "              If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "                varSubProperty = varProperty(0)" & vbNewLine
'  strBaseString = strBaseString & "                If IsEmpty(varSubProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "                  strDescriptionSource = ""Trails_UTM12_NAD83""" & vbNewLine
'  strBaseString = strBaseString & "                Else" & vbNewLine
'  strBaseString = strBaseString & "                  strDescriptionSource = CStr(varSubProperty)" & vbNewLine
'  strBaseString = strBaseString & "                End If" & vbNewLine
'  strBaseString = strBaseString & "              Else" & vbNewLine
'  strBaseString = strBaseString & "                strDescriptionSource = ""Trails_UTM12_NAD83""" & vbNewLine
'  strBaseString = strBaseString & "              End If" & vbNewLine
'  strBaseString = strBaseString & "              strDescriptionSource = Replace(strDescriptionSource, ""Dicitonary"", ""Dictionary"", , , vbTextCompare)" & vbNewLine
'  strBaseString = strBaseString & "              strDescriptionSource = Replace(strDescriptionSource, ""20003"", ""2003"", , , vbTextCompare)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "              ' SET DESCRIPTION SOURCE DEFINITION OF NEW FCLASS" & vbNewLine
'  strBaseString = strBaseString & "              varCheckPropertyPresent = pPropSet.GetProperty(""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdefs"")" & vbNewLine
'  strBaseString = strBaseString & "              If Not IsEmpty(varCheckPropertyPresent) Then" & vbNewLine
'  strBaseString = strBaseString & "                pPropSet.RemoveProperty ""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdefs""" & vbNewLine
'  strBaseString = strBaseString & "              End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "              ' ONLY ADD THIS NEW ONE IF THERE WAS ONE IN THE ORIGINAL FCLASS" & vbNewLine
'  strBaseString = strBaseString & "              If strDescriptionSource <> """" Then" & vbNewLine
'  strBaseString = strBaseString & "                pPropSet.SetProperty ""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdefs"", strDescriptionSource" & vbNewLine
'  strBaseString = strBaseString & "              End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "              Exit For" & vbNewLine
'  strBaseString = strBaseString & "            End If" & vbNewLine
'  strBaseString = strBaseString & "          End If" & vbNewLine
'  strBaseString = strBaseString & "        Next lngIndex" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    Next lngTrails_Index" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "    ' ----------------------------------------------------------------------------------------------" & vbNewLine
'  strBaseString = strBaseString & "    ' <<<<<<<<<<<<  Roads_UTM12_NAD83  >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>" & vbNewLine
'  strBaseString = strBaseString & "    '-----------------------------------------------------------------------------------------------" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""ROADS_Add_Attribute"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = """"""TRUE"""" or """"FALSE"""", whether there was a polyline feature from the 'Roads_UTM12_NAD83' "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""feature class within 100m of the centroid of this polyline.  If so, then this '"" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""Roads_UTM12_NAD83' feature would be considered a candidate for extracting attribute values "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""to describe this feature.""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""ROADS_Attribute_Source"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""If 'ROADS_Add_Attribute' above = """"TRUE"""", then this field will contain the name of the "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""GCNP Roads feature class (Roads_UTM12_NAD83).  Otherwise it should be Null.""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "   " & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""ROADS_Attribute_Distance"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""If 'ROADS_Add_Attribute' above = """"TRUE"""", then this field will contain the distance "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""(in meters) from the centroid of this polyline to the nearest GCNP Roads feature.  Otherwise it should be Null.  "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""Values should always be <= 100m.""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.RemoveProperty strXDom" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomMin, ""0""" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomMax, ""100""" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomUnits, ""Meters""" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""ROADS_ObjectID"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""If 'ROADS_Add_Attribute' above = """"TRUE"""", then this field will contain the OBJECTID "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""value of the nearest GCNP Roads feature, from the Roads_UTM12_NAD83 feature class.""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "    For lngRoads_Index = 0 To UBound(strROADS_RoadArray, 1)" & vbNewLine
'  strBaseString = strBaseString & "      If strFieldName = strROADS_RoadArray(lngRoads_Index, 0) Then '  ""BLM_OBS_USE1"", ""BLM_OBS_USE2"" or ""BLM_OBS_USE3""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "        ' FIND ORIGINAL FIELD METADATA" & vbNewLine
'  strBaseString = strBaseString & "        For lngIndex = 0 To pRoads_UTM12_NAD83_FClass.Fields.FieldCount" & vbNewLine
'  strBaseString = strBaseString & "          strNames(0) = ""eainfo/detailed/attr["" & CStr(lngIndex) & ""]""" & vbNewLine
'  strBaseString = strBaseString & "          pRoads_UTM12_NAD83_PropSet.GetProperties strNames, varVals" & vbNewLine
'  strBaseString = strBaseString & "          varSubVals = varVals(0)" & vbNewLine
'  strBaseString = strBaseString & "          If Not IsEmpty(varSubVals) Then" & vbNewLine
'  strBaseString = strBaseString & "            ' GET FIELD NAME" & vbNewLine
'  strBaseString = strBaseString & "            strOrigFClassXName = ""eainfo/detailed/attr["" & CStr(lngIndex) & ""]/attrlabl""" & vbNewLine
'  strBaseString = strBaseString & "            varProperty = pRoads_UTM12_NAD83_PropSet.GetProperty(strOrigFClassXName)" & vbNewLine
'  strBaseString = strBaseString & "            If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "              varSubProperty = varProperty(0)" & vbNewLine
'  strBaseString = strBaseString & "              If IsEmpty(varSubProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "                strOrigFieldName = ""<-- No Field Name -->""" & vbNewLine
'  strBaseString = strBaseString & "              Else" & vbNewLine
'  strBaseString = strBaseString & "                strOrigFieldName = CStr(varSubProperty)" & vbNewLine
'  strBaseString = strBaseString & "              End If" & vbNewLine
'  strBaseString = strBaseString & "            Else" & vbNewLine
'  strBaseString = strBaseString & "              strOrigFieldName = ""<-- No Field Name -->""" & vbNewLine
'  strBaseString = strBaseString & "            End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "            If strOrigFieldName = strROADS_RoadArray(lngRoads_Index, 1) Then ' ""OBS_USE1"", ""OBS_USE2"", or ""OBS_USE3""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "              ' CHECK IF DOMAIN INFO" & vbNewLine
'  strBaseString = strBaseString & "              strOrigFClassXName = ""eainfo/detailed/attr["" & CStr(lngIndex) & ""]/attrdomv""" & vbNewLine
'  strBaseString = strBaseString & "              varProperty = pRoads_UTM12_NAD83_PropSet.GetProperty(strOrigFClassXName)" & vbNewLine
'  strBaseString = strBaseString & "              If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "                varSubProperty = varProperty(0)" & vbNewLine
'  strBaseString = strBaseString & "'                For lngIndex2 = 0 To UBound(varProperty)" & vbNewLine
'  strBaseString = strBaseString & "'                  Debug.Print ""       "" & CStr(lngIndex2) & ""] "" & CStr(varProperty(lngIndex2))" & vbNewLine
'  strBaseString = strBaseString & "'                Next lngIndex2" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "                ' GET DOMAIN ATTRIBUTES FROM ORIGINAL FCLASS" & vbNewLine
'  strBaseString = strBaseString & "                strOrigFClassXName = ""eainfo/detailed/attr["" & CStr(lngIndex) & ""]/attrdomv""" & vbNewLine
'  strBaseString = strBaseString & "                varProperty = pRoads_UTM12_NAD83_PropSet.GetProperty(strOrigFClassXName)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "                ' ADD DOMAIN ATTRIBUTES TO NEW FCLASS" & vbNewLine
'  strBaseString = strBaseString & "                varCheckPropertyPresent = pPropSet.GetProperty(""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdomv"")" & vbNewLine
'  strBaseString = strBaseString & "                If Not IsEmpty(varCheckPropertyPresent) Then" & vbNewLine
'  strBaseString = strBaseString & "                  pPropSet.RemoveProperty ""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdomv""" & vbNewLine
'  strBaseString = strBaseString & "                End If" & vbNewLine
'  strBaseString = strBaseString & "                For lngIndex3 = 0 To UBound(varProperty)" & vbNewLine
'  strBaseString = strBaseString & "                  strSubXName1 = ""eainfo/detailed/attr["" & CStr(lngIndex) & ""]/attrdomv/"" & _" & vbNewLine
'  strBaseString = strBaseString & "                      varProperty(lngIndex3) & ""["" & CStr(lngIndex3) & ""]""" & vbNewLine
'  strBaseString = strBaseString & "                  varSubProp1 = pRoads_UTM12_NAD83_PropSet.GetProperty(strSubXName1)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "                  If UBound(varSubProp1) = 0 Then" & vbNewLine
'  strBaseString = strBaseString & "                    strSubXName3 = ""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdomv/"" & _" & vbNewLine
'  strBaseString = strBaseString & "                        varProperty(lngIndex3) & ""["" & CStr(lngIndex3) & ""]""" & vbNewLine
'  strBaseString = strBaseString & "                    strValue = varSubProp1(0)" & vbNewLine
'  strBaseString = strBaseString & "                    strValue = Replace(strValue, ""20003"", ""2003"")" & vbNewLine
'  strBaseString = strBaseString & "                    strValue = Replace(strValue, ""Dicitonary"", ""Dictionary"")" & vbNewLine
'  strBaseString = strBaseString & "                    pPropSet.SetProperty strSubXName3, strValue" & vbNewLine
'  strBaseString = strBaseString & "                  Else" & vbNewLine
'  strBaseString = strBaseString & "                    For lngIndex4 = 0 To UBound(varSubProp1)" & vbNewLine
'  strBaseString = strBaseString & "                      strSubXName2 = ""eainfo/detailed/attr["" & CStr(lngIndex) & ""]/attrdomv/"" & _" & vbNewLine
'  strBaseString = strBaseString & "                          varProperty(lngIndex3) & ""["" & CStr(lngIndex3) & ""]/"" & varSubProp1(lngIndex4)" & vbNewLine
'  strBaseString = strBaseString & "                      varSubProp2 = pRoads_UTM12_NAD83_PropSet.GetProperty(strSubXName2)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "                      ' WRITE VALUE BACK TO NEW FCLASS" & vbNewLine
'  strBaseString = strBaseString & "                      strSubXName3 = ""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdomv/"" & _" & vbNewLine
'  strBaseString = strBaseString & "                          varProperty(lngIndex3) & ""["" & CStr(lngIndex3) & ""]/"" & varSubProp1(lngIndex4)" & vbNewLine
'  strBaseString = strBaseString & "                      strValue = varSubProp2(0)" & vbNewLine
'  strBaseString = strBaseString & "                      strValue = Replace(strValue, ""20003"", ""2003"")" & vbNewLine
'  strBaseString = strBaseString & "                      strValue = Replace(strValue, ""Dicitonary"", ""Dictionary"")" & vbNewLine
'  strBaseString = strBaseString & "                      pPropSet.SetProperty strSubXName3, strValue" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "                    Next lngIndex4" & vbNewLine
'  strBaseString = strBaseString & "                  End If" & vbNewLine
'  strBaseString = strBaseString & "                Next lngIndex3" & vbNewLine
'  strBaseString = strBaseString & "              Else" & vbNewLine
'  strBaseString = strBaseString & "'                strOrigFCLassXName = ""eainfo/detailed/attr["" & CStr(lngIndex) & ""]""" & vbNewLine
'  strBaseString = strBaseString & "'                varProperty = pRoads_UTM12_NAD83_PropSet.GetProperty(strOrigFCLassXName)" & vbNewLine
'  strBaseString = strBaseString & "'                varSubProperty = varProperty(0)" & vbNewLine
'  strBaseString = strBaseString & "'                Debug.Print ""Examining "" & strOrigFieldName & ""...""" & vbNewLine
'  strBaseString = strBaseString & "'                For lngIndex2 = 0 To UBound(varProperty)" & vbNewLine
'  strBaseString = strBaseString & "'                  varSubProp1 = pRoads_UTM12_NAD83_PropSet.GetProperty(strOrigFCLassXName & ""/"" & _" & vbNewLine
'  strBaseString = strBaseString & "'                        CStr(varProperty(lngIndex2)))" & vbNewLine
'  strBaseString = strBaseString & "'                  Debug.Print ""       "" & CStr(lngIndex2) & ""] "" & CStr(varProperty(lngIndex2)) & "" = "" & _" & vbNewLine
'  strBaseString = strBaseString & "'                        CStr(varSubProp1(0))" & vbNewLine
'  strBaseString = strBaseString & "'                Next lngIndex2" & vbNewLine
'  strBaseString = strBaseString & "              End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "              If strOrigFieldName = ""ROAD_NO_"" Then" & vbNewLine
'  strBaseString = strBaseString & "                Debug.Print ""Here...""" & vbNewLine
'  strBaseString = strBaseString & "              End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "              ' GET DEFINITION FROM ORIGINAL FCLASS" & vbNewLine
'  strBaseString = strBaseString & "              varProperty = pRoads_UTM12_NAD83_PropSet.GetProperty(""eainfo/detailed/attr["" & CStr(lngIndex) & ""]/attrdef"")" & vbNewLine
'  strBaseString = strBaseString & "              If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "                varSubProperty = varProperty(0)" & vbNewLine
'  strBaseString = strBaseString & "                If IsEmpty(varSubProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "                  strDescription = ""<-- No Description -->""" & vbNewLine
'  strBaseString = strBaseString & "                Else" & vbNewLine
'  strBaseString = strBaseString & "                  strDescription = CStr(varSubProperty)" & vbNewLine
'  strBaseString = strBaseString & "                  strDescription = Trim(strDescription)" & vbNewLine
'  strBaseString = strBaseString & "                  If Right(strDescription, 1) <> ""."" And Right(strDescription, 1) <> ""!"" And _" & vbNewLine
'  strBaseString = strBaseString & "                      Right(strDescription, 1) <> ""?"" Then" & vbNewLine
'  strBaseString = strBaseString & "                    strDescription = strDescription & "".""" & vbNewLine
'  strBaseString = strBaseString & "                  End If" & vbNewLine
'  strBaseString = strBaseString & "                End If" & vbNewLine
'  strBaseString = strBaseString & "              Else" & vbNewLine
'  strBaseString = strBaseString & "                strDescription = ""<-- No Description -->""" & vbNewLine
'  strBaseString = strBaseString & "              End If" & vbNewLine
'  strBaseString = strBaseString & "              strDescription = Replace(strDescription, ""  [Imported from TransportationLine_UTM12_NAD83]"", """", , , vbTextCompare)" & vbNewLine
'  strBaseString = strBaseString & "              strDescription = strDescription & _" & vbNewLine
'  strBaseString = strBaseString & "                  ""  This field should only have a value if the 'ROADS_Add_Attribute' value = """"TRUE"""".  "" & _" & vbNewLine
'  strBaseString = strBaseString & "                  ""[Imported from Roads_UTM12_NAD83, Field """""" & strROADS_RoadArray(lngRoads_Index, 1) & """"""]""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "              ' SET FIELD DEFINITION OF NEW FCLASS" & vbNewLine
'  strBaseString = strBaseString & "              varCheckPropertyPresent = pPropSet.GetProperty(""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdef"")" & vbNewLine
'  strBaseString = strBaseString & "              If Not IsEmpty(varCheckPropertyPresent) Then" & vbNewLine
'  strBaseString = strBaseString & "                pPropSet.RemoveProperty ""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdef""" & vbNewLine
'  strBaseString = strBaseString & "              End If" & vbNewLine
'  strBaseString = strBaseString & "              pPropSet.SetProperty ""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdef"", strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "              ' GET DESCRIPTION SOURCE FROM ORIGINAL FCLASS" & vbNewLine
'  strBaseString = strBaseString & "              varProperty = pRoads_UTM12_NAD83_PropSet.GetProperty(""eainfo/detailed/attr["" & CStr(lngIndex) & ""]/attrdefs"")" & vbNewLine
'  strBaseString = strBaseString & "              If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "                varSubProperty = varProperty(0)" & vbNewLine
'  strBaseString = strBaseString & "                If IsEmpty(varSubProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "                  strDescriptionSource = ""Roads_UTM12_NAD83""" & vbNewLine
'  strBaseString = strBaseString & "                Else" & vbNewLine
'  strBaseString = strBaseString & "                  strDescriptionSource = CStr(varSubProperty)" & vbNewLine
'  strBaseString = strBaseString & "                End If" & vbNewLine
'  strBaseString = strBaseString & "              Else" & vbNewLine
'  strBaseString = strBaseString & "                strDescriptionSource = ""Roads_UTM12_NAD83""" & vbNewLine
'  strBaseString = strBaseString & "              End If" & vbNewLine
'  strBaseString = strBaseString & "              strDescriptionSource = Replace(strDescriptionSource, ""Dicitonary"", ""Dictionary"", , , vbTextCompare)" & vbNewLine
'  strBaseString = strBaseString & "              strDescriptionSource = Replace(strDescriptionSource, ""20003"", ""2003"", , , vbTextCompare)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "              ' SET DESCRIPTION SOURCE DEFINITION OF NEW FCLASS" & vbNewLine
'  strBaseString = strBaseString & "              varCheckPropertyPresent = pPropSet.GetProperty(""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdefs"")" & vbNewLine
'  strBaseString = strBaseString & "              If Not IsEmpty(varCheckPropertyPresent) Then" & vbNewLine
'  strBaseString = strBaseString & "                pPropSet.RemoveProperty ""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdefs""" & vbNewLine
'  strBaseString = strBaseString & "              End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "              ' ONLY ADD THIS NEW ONE IF THERE WAS ONE IN THE ORIGINAL FCLASS" & vbNewLine
'  strBaseString = strBaseString & "              If strDescriptionSource <> """" Then" & vbNewLine
'  strBaseString = strBaseString & "                pPropSet.SetProperty ""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdefs"", strDescriptionSource" & vbNewLine
'  strBaseString = strBaseString & "              End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "              Exit For" & vbNewLine
'  strBaseString = strBaseString & "            End If" & vbNewLine
'  strBaseString = strBaseString & "          End If" & vbNewLine
'  strBaseString = strBaseString & "        Next lngIndex" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    Next lngRoads_Index" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "    ' ----------------------------------------------------------------------------------------------" & vbNewLine
'  strBaseString = strBaseString & "    ' <<<<<<<<<<<<  TIGER_2012_Trans_Near_Park_UTM12_NAD83  >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>" & vbNewLine
'  strBaseString = strBaseString & "    '-----------------------------------------------------------------------------------------------" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""TIG_Add_Attribute"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = """"""TRUE"""" or """"FALSE"""", whether there was a polyline feature from the 'TIGER_2012_Trans_Near_Park_UTM12_NAD83' "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""feature class within 100m of the centroid of this polyline.  If so, then this '"" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""TIGER_2012_Trans_Near_Park_UTM12_NAD83' feature would be considered a candidate for extracting attribute values "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""to describe this feature.""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""TIG_Attribute_Source"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""If 'TIG_Add_Attribute' above = """"TRUE"""", then this field will contain the name of the "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""TIGER 2012 Transportation feature class (TIGER_2012_Trans_Near_Park_UTM12_NAD83).  Otherwise it should be Null.""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""TIG_Attribute_Distance"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""If 'TIG_Add_Attribute' above = """"TRUE"""", then this field will contain the distance "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""(in meters) from the centroid of this polyline to the nearest TIGER 2012 Transportation feature.  Otherwise it should be Null.  "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""Values should always be <= 100m.""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.RemoveProperty strXDom" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomMin, ""0""" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomMax, ""100""" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomUnits, ""Meters""" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""TIG_ObjectID"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""If 'TIG_Add_Attribute' above = """"TRUE"""", then this field will contain the OBJECTID "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""value of the nearest TIGER 2012 Transportation feature, from the TIGER_2012_Trans_Near_Park_UTM12_NAD83 feature class.""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "    For lngTIGIndex = 0 To UBound(strTIG_RoadArray, 1)" & vbNewLine
'  strBaseString = strBaseString & "      If strFieldName = strTIG_RoadArray(lngTIGIndex, 0) Then '  ""BLM_OBS_USE1"", ""BLM_OBS_USE2"" or ""BLM_OBS_USE3""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "        ' FIND ORIGINAL FIELD METADATA" & vbNewLine
'  strBaseString = strBaseString & "        For lngIndex = 0 To pTIGER_2012_Trans_Near_Park_UTM12_NAD83_FClass.Fields.FieldCount" & vbNewLine
'  strBaseString = strBaseString & "          strNames(0) = ""eainfo/detailed/attr["" & CStr(lngIndex) & ""]""" & vbNewLine
'  strBaseString = strBaseString & "          pTIGER_2012_Trans_Near_Park_UTM12_NAD83_PropSet.GetProperties strNames, varVals" & vbNewLine
'  strBaseString = strBaseString & "          varSubVals = varVals(0)" & vbNewLine
'  strBaseString = strBaseString & "          If Not IsEmpty(varSubVals) Then" & vbNewLine
'  strBaseString = strBaseString & "            ' GET FIELD NAME" & vbNewLine
'  strBaseString = strBaseString & "            strOrigFClassXName = ""eainfo/detailed/attr["" & CStr(lngIndex) & ""]/attrlabl""" & vbNewLine
'  strBaseString = strBaseString & "            varProperty = pTIGER_2012_Trans_Near_Park_UTM12_NAD83_PropSet.GetProperty(strOrigFClassXName)" & vbNewLine
'  strBaseString = strBaseString & "            If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "              varSubProperty = varProperty(0)" & vbNewLine
'  strBaseString = strBaseString & "              If IsEmpty(varSubProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "                strOrigFieldName = ""<-- No Field Name -->""" & vbNewLine
'  strBaseString = strBaseString & "              Else" & vbNewLine
'  strBaseString = strBaseString & "                strOrigFieldName = CStr(varSubProperty)" & vbNewLine
'  strBaseString = strBaseString & "              End If" & vbNewLine
'  strBaseString = strBaseString & "            Else" & vbNewLine
'  strBaseString = strBaseString & "              strOrigFieldName = ""<-- No Field Name -->""" & vbNewLine
'  strBaseString = strBaseString & "            End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "            If strOrigFieldName = strTIG_RoadArray(lngTIGIndex, 1) Then ' ""OBS_USE1"", ""OBS_USE2"", or ""OBS_USE3""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "              ' CHECK IF DOMAIN INFO" & vbNewLine
'  strBaseString = strBaseString & "              strOrigFClassXName = ""eainfo/detailed/attr["" & CStr(lngIndex) & ""]/attrdomv""" & vbNewLine
'  strBaseString = strBaseString & "              varProperty = pTIGER_2012_Trans_Near_Park_UTM12_NAD83_PropSet.GetProperty(strOrigFClassXName)" & vbNewLine
'  strBaseString = strBaseString & "              If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "                varSubProperty = varProperty(0)" & vbNewLine
'  strBaseString = strBaseString & "'                For lngIndex2 = 0 To UBound(varProperty)" & vbNewLine
'  strBaseString = strBaseString & "'                  Debug.Print ""       "" & CStr(lngIndex2) & ""] "" & CStr(varProperty(lngIndex2))" & vbNewLine
'  strBaseString = strBaseString & "'                Next lngIndex2" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "                ' GET DOMAIN ATTRIBUTES FROM ORIGINAL FCLASS" & vbNewLine
'  strBaseString = strBaseString & "                strOrigFClassXName = ""eainfo/detailed/attr["" & CStr(lngIndex) & ""]/attrdomv""" & vbNewLine
'  strBaseString = strBaseString & "                varProperty = pTIGER_2012_Trans_Near_Park_UTM12_NAD83_PropSet.GetProperty(strOrigFClassXName)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "                ' ADD DOMAIN ATTRIBUTES TO NEW FCLASS" & vbNewLine
'  strBaseString = strBaseString & "                varCheckPropertyPresent = pPropSet.GetProperty(""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdomv"")" & vbNewLine
'  strBaseString = strBaseString & "                If Not IsEmpty(varCheckPropertyPresent) Then" & vbNewLine
'  strBaseString = strBaseString & "                  pPropSet.RemoveProperty ""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdomv""" & vbNewLine
'  strBaseString = strBaseString & "                End If" & vbNewLine
'  strBaseString = strBaseString & "                For lngIndex3 = 0 To UBound(varProperty)" & vbNewLine
'  strBaseString = strBaseString & "                  strSubXName1 = ""eainfo/detailed/attr["" & CStr(lngIndex) & ""]/attrdomv/"" & _" & vbNewLine
'  strBaseString = strBaseString & "                      varProperty(lngIndex3) & ""["" & CStr(lngIndex3) & ""]""" & vbNewLine
'  strBaseString = strBaseString & "                  varSubProp1 = pTIGER_2012_Trans_Near_Park_UTM12_NAD83_PropSet.GetProperty(strSubXName1)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "                  If UBound(varSubProp1) = 0 Then" & vbNewLine
'  strBaseString = strBaseString & "                    strSubXName3 = ""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdomv/"" & _" & vbNewLine
'  strBaseString = strBaseString & "                        varProperty(lngIndex3) & ""["" & CStr(lngIndex3) & ""]""" & vbNewLine
'  strBaseString = strBaseString & "                    strValue = varSubProp1(0)" & vbNewLine
'  strBaseString = strBaseString & "                    strValue = Replace(strValue, ""20003"", ""2003"")" & vbNewLine
'  strBaseString = strBaseString & "                    strValue = Replace(strValue, ""Dicitonary"", ""Dictionary"")" & vbNewLine
'  strBaseString = strBaseString & "                    pPropSet.SetProperty strSubXName3, strValue" & vbNewLine
'  strBaseString = strBaseString & "                  Else" & vbNewLine
'  strBaseString = strBaseString & "                    For lngIndex4 = 0 To UBound(varSubProp1)" & vbNewLine
'  strBaseString = strBaseString & "                      strSubXName2 = ""eainfo/detailed/attr["" & CStr(lngIndex) & ""]/attrdomv/"" & _" & vbNewLine
'  strBaseString = strBaseString & "                          varProperty(lngIndex3) & ""["" & CStr(lngIndex3) & ""]/"" & varSubProp1(lngIndex4)" & vbNewLine
'  strBaseString = strBaseString & "                      varSubProp2 = pTIGER_2012_Trans_Near_Park_UTM12_NAD83_PropSet.GetProperty(strSubXName2)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "                      ' WRITE VALUE BACK TO NEW FCLASS" & vbNewLine
'  strBaseString = strBaseString & "                      strSubXName3 = ""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdomv/"" & _" & vbNewLine
'  strBaseString = strBaseString & "                          varProperty(lngIndex3) & ""["" & CStr(lngIndex3) & ""]/"" & varSubProp1(lngIndex4)" & vbNewLine
'  strBaseString = strBaseString & "                      strValue = varSubProp2(0)" & vbNewLine
'  strBaseString = strBaseString & "                      strValue = Replace(strValue, ""20003"", ""2003"")" & vbNewLine
'  strBaseString = strBaseString & "                      strValue = Replace(strValue, ""Dicitonary"", ""Dictionary"")" & vbNewLine
'  strBaseString = strBaseString & "                      pPropSet.SetProperty strSubXName3, strValue" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "                    Next lngIndex4" & vbNewLine
'  strBaseString = strBaseString & "                  End If" & vbNewLine
'  strBaseString = strBaseString & "                Next lngIndex3" & vbNewLine
'  strBaseString = strBaseString & "              Else" & vbNewLine
'  strBaseString = strBaseString & "'                strOrigFCLassXName = ""eainfo/detailed/attr["" & CStr(lngIndex) & ""]""" & vbNewLine
'  strBaseString = strBaseString & "'                varProperty = pTIGER_2012_Trans_Near_Park_UTM12_NAD83_PropSet.GetProperty(strOrigFCLassXName)" & vbNewLine
'  strBaseString = strBaseString & "'                varSubProperty = varProperty(0)" & vbNewLine
'  strBaseString = strBaseString & "'                Debug.Print ""Examining "" & strOrigFieldName & ""...""" & vbNewLine
'  strBaseString = strBaseString & "'                For lngIndex2 = 0 To UBound(varProperty)" & vbNewLine
'  strBaseString = strBaseString & "'                  varSubProp1 = pTIGER_2012_Trans_Near_Park_UTM12_NAD83_PropSet.GetProperty(strOrigFCLassXName & ""/"" & _" & vbNewLine
'  strBaseString = strBaseString & "'                        CStr(varProperty(lngIndex2)))" & vbNewLine
'  strBaseString = strBaseString & "'                  Debug.Print ""       "" & CStr(lngIndex2) & ""] "" & CStr(varProperty(lngIndex2)) & "" = "" & _" & vbNewLine
'  strBaseString = strBaseString & "'                        CStr(varSubProp1(0))" & vbNewLine
'  strBaseString = strBaseString & "'                Next lngIndex2" & vbNewLine
'  strBaseString = strBaseString & "              End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "              If strOrigFieldName = ""ROAD_NO_"" Then" & vbNewLine
'  strBaseString = strBaseString & "                Debug.Print ""Here...""" & vbNewLine
'  strBaseString = strBaseString & "              End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "              ' GET DEFINITION FROM ORIGINAL FCLASS" & vbNewLine
'  strBaseString = strBaseString & "              varProperty = pTIGER_2012_Trans_Near_Park_UTM12_NAD83_PropSet.GetProperty(""eainfo/detailed/attr["" & CStr(lngIndex) & ""]/attrdef"")" & vbNewLine
'  strBaseString = strBaseString & "              If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "                varSubProperty = varProperty(0)" & vbNewLine
'  strBaseString = strBaseString & "                If IsEmpty(varSubProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "                  strDescription = ""<-- No Description -->""" & vbNewLine
'  strBaseString = strBaseString & "                Else" & vbNewLine
'  strBaseString = strBaseString & "                  strDescription = CStr(varSubProperty)" & vbNewLine
'  strBaseString = strBaseString & "                  strDescription = Trim(strDescription)" & vbNewLine
'  strBaseString = strBaseString & "                  If Right(strDescription, 1) <> ""."" And Right(strDescription, 1) <> ""!"" And _" & vbNewLine
'  strBaseString = strBaseString & "                      Right(strDescription, 1) <> ""?"" Then" & vbNewLine
'  strBaseString = strBaseString & "                    strDescription = strDescription & "".""" & vbNewLine
'  strBaseString = strBaseString & "                  End If" & vbNewLine
'  strBaseString = strBaseString & "                End If" & vbNewLine
'  strBaseString = strBaseString & "              Else" & vbNewLine
'  strBaseString = strBaseString & "                strDescription = ""<-- No Description -->""" & vbNewLine
'  strBaseString = strBaseString & "              End If" & vbNewLine
'  strBaseString = strBaseString & "              strDescription = Replace(strDescription, ""  [Imported from TransportationLine_UTM12_NAD83]"", """", , , vbTextCompare)" & vbNewLine
'  strBaseString = strBaseString & "              strDescription = strDescription & _" & vbNewLine
'  strBaseString = strBaseString & "                  ""  This field should only have a value if the 'TIG_Add_Attribute' value = """"TRUE"""".  "" & _" & vbNewLine
'  strBaseString = strBaseString & "                  ""[Imported from TIGER_2012_Trans_Near_Park_UTM12_NAD83, Field """""" & strTIG_RoadArray(lngTIGIndex, 1) & """"""]""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "              ' SET FIELD DEFINITION OF NEW FCLASS" & vbNewLine
'  strBaseString = strBaseString & "              varCheckPropertyPresent = pPropSet.GetProperty(""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdef"")" & vbNewLine
'  strBaseString = strBaseString & "              If Not IsEmpty(varCheckPropertyPresent) Then" & vbNewLine
'  strBaseString = strBaseString & "                pPropSet.RemoveProperty ""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdef""" & vbNewLine
'  strBaseString = strBaseString & "              End If" & vbNewLine
'  strBaseString = strBaseString & "              pPropSet.SetProperty ""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdef"", strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "              ' GET DESCRIPTION SOURCE FROM ORIGINAL FCLASS" & vbNewLine
'  strBaseString = strBaseString & "              varProperty = pTIGER_2012_Trans_Near_Park_UTM12_NAD83_PropSet.GetProperty(""eainfo/detailed/attr["" & CStr(lngIndex) & ""]/attrdefs"")" & vbNewLine
'  strBaseString = strBaseString & "              If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "                varSubProperty = varProperty(0)" & vbNewLine
'  strBaseString = strBaseString & "                If IsEmpty(varSubProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "                  strDescriptionSource = ""TIGER_2012_Trans_Near_Park_UTM12_NAD83""" & vbNewLine
'  strBaseString = strBaseString & "                Else" & vbNewLine
'  strBaseString = strBaseString & "                  strDescriptionSource = CStr(varSubProperty)" & vbNewLine
'  strBaseString = strBaseString & "                End If" & vbNewLine
'  strBaseString = strBaseString & "              Else" & vbNewLine
'  strBaseString = strBaseString & "                strDescriptionSource = ""TIGER_2012_Trans_Near_Park_UTM12_NAD83""" & vbNewLine
'  strBaseString = strBaseString & "              End If" & vbNewLine
'  strBaseString = strBaseString & "              strDescriptionSource = Replace(strDescriptionSource, ""Dicitonary"", ""Dictionary"", , , vbTextCompare)" & vbNewLine
'  strBaseString = strBaseString & "              strDescriptionSource = Replace(strDescriptionSource, ""20003"", ""2003"", , , vbTextCompare)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "              ' SET DESCRIPTION SOURCE DEFINITION OF NEW FCLASS" & vbNewLine
'  strBaseString = strBaseString & "              varCheckPropertyPresent = pPropSet.GetProperty(""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdefs"")" & vbNewLine
'  strBaseString = strBaseString & "              If Not IsEmpty(varCheckPropertyPresent) Then" & vbNewLine
'  strBaseString = strBaseString & "                pPropSet.RemoveProperty ""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdefs""" & vbNewLine
'  strBaseString = strBaseString & "              End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "              ' ONLY ADD THIS NEW ONE IF THERE WAS ONE IN THE ORIGINAL FCLASS" & vbNewLine
'  strBaseString = strBaseString & "              If strDescriptionSource <> """" Then" & vbNewLine
'  strBaseString = strBaseString & "                pPropSet.SetProperty ""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdefs"", strDescriptionSource" & vbNewLine
'  strBaseString = strBaseString & "              End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "              Exit For" & vbNewLine
'  strBaseString = strBaseString & "            End If" & vbNewLine
'  strBaseString = strBaseString & "          End If" & vbNewLine
'  strBaseString = strBaseString & "        Next lngIndex" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    Next lngTIGIndex" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "    ' ----------------------------------------------------------------------------------------------" & vbNewLine
'  strBaseString = strBaseString & "    ' <<<<<<<<<<<<  Roads_KNF_UTM12_NAD83  >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>" & vbNewLine
'  strBaseString = strBaseString & "    '-----------------------------------------------------------------------------------------------" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""KNF_Add_Attribute"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = """"""TRUE"""" or """"FALSE"""", whether there was a polyline feature from the 'Roads_KNF_UTM12_NAD83' "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""feature class within 100m of the centroid of this polyline.  If so, then this '"" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""Roads_KNF_UTM12_NAD83' feature would be considered a candidate for extracting attribute values "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""to describe this feature.""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""KNF_Attribute_Source"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""If 'KNF_Add_Attribute' above = """"TRUE"""", then this field will contain the name of the "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""Kaibab NF Roads feature class (Roads_KNF_UTM12_NAD83).  Otherwise it should be Null.""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""KNF_Attribute_Distance"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""If 'KNF_Add_Attribute' above = """"TRUE"""", then this field will contain the distance "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""(in meters) from the centroid of this polyline to the nearest Kaibab NF Road feature.  Otherwise it should be Null.  "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""Values should always be <= 100m.""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "   " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.RemoveProperty strXDom" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomMin, ""0""" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomMax, ""100""" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomUnits, ""Meters""" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""KNF_ObjectID_1"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""If 'KNF_Add_Attribute' above = """"TRUE"""", then this field will contain the OBJECTID_1 "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""value of the nearest Kaibab NF Road feature, from the Roads_KNF_UTM12_NAD83 feature class.""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "    For lngKNFIndex = 0 To UBound(strKNF_RoadArray, 1)" & vbNewLine
'  strBaseString = strBaseString & "      If strFieldName = strKNF_RoadArray(lngKNFIndex, 0) Then '  ""BLM_OBS_USE1"", ""BLM_OBS_USE2"" or ""BLM_OBS_USE3""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "        ' FIND ORIGINAL FIELD METADATA" & vbNewLine
'  strBaseString = strBaseString & "        For lngIndex = 0 To pRoads_KNF_UTM12_NAD83_FClass.Fields.FieldCount" & vbNewLine
'  strBaseString = strBaseString & "          strNames(0) = ""eainfo/detailed/attr["" & CStr(lngIndex) & ""]""" & vbNewLine
'  strBaseString = strBaseString & "          pRoads_KNF_UTM12_NAD83_PropSet.GetProperties strNames, varVals" & vbNewLine
'  strBaseString = strBaseString & "          varSubVals = varVals(0)" & vbNewLine
'  strBaseString = strBaseString & "          If Not IsEmpty(varSubVals) Then" & vbNewLine
'  strBaseString = strBaseString & "            ' GET FIELD NAME" & vbNewLine
'  strBaseString = strBaseString & "            strOrigFClassXName = ""eainfo/detailed/attr["" & CStr(lngIndex) & ""]/attrlabl""" & vbNewLine
'  strBaseString = strBaseString & "            varProperty = pRoads_KNF_UTM12_NAD83_PropSet.GetProperty(strOrigFClassXName)" & vbNewLine
'  strBaseString = strBaseString & "            If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "              varSubProperty = varProperty(0)" & vbNewLine
'  strBaseString = strBaseString & "              If IsEmpty(varSubProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "                strOrigFieldName = ""<-- No Field Name -->""" & vbNewLine
'  strBaseString = strBaseString & "              Else" & vbNewLine
'  strBaseString = strBaseString & "                strOrigFieldName = CStr(varSubProperty)" & vbNewLine
'  strBaseString = strBaseString & "              End If" & vbNewLine
'  strBaseString = strBaseString & "            Else" & vbNewLine
'  strBaseString = strBaseString & "              strOrigFieldName = ""<-- No Field Name -->""" & vbNewLine
'  strBaseString = strBaseString & "            End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "            If strOrigFieldName = strKNF_RoadArray(lngKNFIndex, 1) Then ' ""OBS_USE1"", ""OBS_USE2"", or ""OBS_USE3""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "              ' CHECK IF DOMAIN INFO" & vbNewLine
'  strBaseString = strBaseString & "              strOrigFClassXName = ""eainfo/detailed/attr["" & CStr(lngIndex) & ""]/attrdomv""" & vbNewLine
'  strBaseString = strBaseString & "              varProperty = pRoads_KNF_UTM12_NAD83_PropSet.GetProperty(strOrigFClassXName)" & vbNewLine
'  strBaseString = strBaseString & "              If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "                varSubProperty = varProperty(0)" & vbNewLine
'  strBaseString = strBaseString & "'                For lngIndex2 = 0 To UBound(varProperty)" & vbNewLine
'  strBaseString = strBaseString & "'                  Debug.Print ""       "" & CStr(lngIndex2) & ""] "" & CStr(varProperty(lngIndex2))" & vbNewLine
'  strBaseString = strBaseString & "'                Next lngIndex2" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "                ' GET DOMAIN ATTRIBUTES FROM ORIGINAL FCLASS" & vbNewLine
'  strBaseString = strBaseString & "                strOrigFClassXName = ""eainfo/detailed/attr["" & CStr(lngIndex) & ""]/attrdomv""" & vbNewLine
'  strBaseString = strBaseString & "                varProperty = pRoads_KNF_UTM12_NAD83_PropSet.GetProperty(strOrigFClassXName)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "                ' ADD DOMAIN ATTRIBUTES TO NEW FCLASS" & vbNewLine
'  strBaseString = strBaseString & "                varCheckPropertyPresent = pPropSet.GetProperty(""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdomv"")" & vbNewLine
'  strBaseString = strBaseString & "                If Not IsEmpty(varCheckPropertyPresent) Then" & vbNewLine
'  strBaseString = strBaseString & "                  pPropSet.RemoveProperty ""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdomv""" & vbNewLine
'  strBaseString = strBaseString & "                End If" & vbNewLine
'  strBaseString = strBaseString & "                For lngIndex3 = 0 To UBound(varProperty)" & vbNewLine
'  strBaseString = strBaseString & "                  strSubXName1 = ""eainfo/detailed/attr["" & CStr(lngIndex) & ""]/attrdomv/"" & _" & vbNewLine
'  strBaseString = strBaseString & "                      varProperty(lngIndex3) & ""["" & CStr(lngIndex3) & ""]""" & vbNewLine
'  strBaseString = strBaseString & "                  varSubProp1 = pRoads_KNF_UTM12_NAD83_PropSet.GetProperty(strSubXName1)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "                  If UBound(varSubProp1) = 0 Then" & vbNewLine
'  strBaseString = strBaseString & "                    strSubXName3 = ""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdomv/"" & _" & vbNewLine
'  strBaseString = strBaseString & "                        varProperty(lngIndex3) & ""["" & CStr(lngIndex3) & ""]""" & vbNewLine
'  strBaseString = strBaseString & "                    strValue = varSubProp1(0)" & vbNewLine
'  strBaseString = strBaseString & "                    strValue = Replace(strValue, ""20003"", ""2003"")" & vbNewLine
'  strBaseString = strBaseString & "                    strValue = Replace(strValue, ""Dicitonary"", ""Dictionary"")" & vbNewLine
'  strBaseString = strBaseString & "                    pPropSet.SetProperty strSubXName3, strValue" & vbNewLine
'  strBaseString = strBaseString & "                  Else" & vbNewLine
'  strBaseString = strBaseString & "                    For lngIndex4 = 0 To UBound(varSubProp1)" & vbNewLine
'  strBaseString = strBaseString & "                      strSubXName2 = ""eainfo/detailed/attr["" & CStr(lngIndex) & ""]/attrdomv/"" & _" & vbNewLine
'  strBaseString = strBaseString & "                          varProperty(lngIndex3) & ""["" & CStr(lngIndex3) & ""]/"" & varSubProp1(lngIndex4)" & vbNewLine
'  strBaseString = strBaseString & "                      varSubProp2 = pRoads_KNF_UTM12_NAD83_PropSet.GetProperty(strSubXName2)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "                      ' WRITE VALUE BACK TO NEW FCLASS" & vbNewLine
'  strBaseString = strBaseString & "                      strSubXName3 = ""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdomv/"" & _" & vbNewLine
'  strBaseString = strBaseString & "                          varProperty(lngIndex3) & ""["" & CStr(lngIndex3) & ""]/"" & varSubProp1(lngIndex4)" & vbNewLine
'  strBaseString = strBaseString & "                      strValue = varSubProp2(0)" & vbNewLine
'  strBaseString = strBaseString & "                      strValue = Replace(strValue, ""20003"", ""2003"")" & vbNewLine
'  strBaseString = strBaseString & "                      strValue = Replace(strValue, ""Dicitonary"", ""Dictionary"")" & vbNewLine
'  strBaseString = strBaseString & "                      pPropSet.SetProperty strSubXName3, strValue" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "                    Next lngIndex4" & vbNewLine
'  strBaseString = strBaseString & "                  End If" & vbNewLine
'  strBaseString = strBaseString & "                Next lngIndex3" & vbNewLine
'  strBaseString = strBaseString & "              Else" & vbNewLine
'  strBaseString = strBaseString & "'                strOrigFCLassXName = ""eainfo/detailed/attr["" & CStr(lngIndex) & ""]""" & vbNewLine
'  strBaseString = strBaseString & "'                varProperty = pRoads_KNF_UTM12_NAD83_PropSet.GetProperty(strOrigFCLassXName)" & vbNewLine
'  strBaseString = strBaseString & "'                varSubProperty = varProperty(0)" & vbNewLine
'  strBaseString = strBaseString & "'                Debug.Print ""Examining "" & strOrigFieldName & ""...""" & vbNewLine
'  strBaseString = strBaseString & "'                For lngIndex2 = 0 To UBound(varProperty)" & vbNewLine
'  strBaseString = strBaseString & "'                  varSubProp1 = pRoads_KNF_UTM12_NAD83_PropSet.GetProperty(strOrigFCLassXName & ""/"" & _" & vbNewLine
'  strBaseString = strBaseString & "'                        CStr(varProperty(lngIndex2)))" & vbNewLine
'  strBaseString = strBaseString & "'                  Debug.Print ""       "" & CStr(lngIndex2) & ""] "" & CStr(varProperty(lngIndex2)) & "" = "" & _" & vbNewLine
'  strBaseString = strBaseString & "'                        CStr(varSubProp1(0))" & vbNewLine
'  strBaseString = strBaseString & "'                Next lngIndex2" & vbNewLine
'  strBaseString = strBaseString & "              End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "              If strOrigFieldName = ""ROAD_NO_"" Then" & vbNewLine
'  strBaseString = strBaseString & "                Debug.Print ""Here...""" & vbNewLine
'  strBaseString = strBaseString & "              End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "              ' GET DEFINITION FROM ORIGINAL FCLASS" & vbNewLine
'  strBaseString = strBaseString & "              varProperty = pRoads_KNF_UTM12_NAD83_PropSet.GetProperty(""eainfo/detailed/attr["" & CStr(lngIndex) & ""]/attrdef"")" & vbNewLine
'  strBaseString = strBaseString & "              If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "                varSubProperty = varProperty(0)" & vbNewLine
'  strBaseString = strBaseString & "                If IsEmpty(varSubProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "                  strDescription = ""<-- No Description -->""" & vbNewLine
'  strBaseString = strBaseString & "                Else" & vbNewLine
'  strBaseString = strBaseString & "                  strDescription = CStr(varSubProperty)" & vbNewLine
'  strBaseString = strBaseString & "                End If" & vbNewLine
'  strBaseString = strBaseString & "              Else" & vbNewLine
'  strBaseString = strBaseString & "                strDescription = ""<-- No Description -->""" & vbNewLine
'  strBaseString = strBaseString & "              End If" & vbNewLine
'  strBaseString = strBaseString & "              strDescription = Replace(strDescription, ""  [Imported from TransportationLine_UTM12_NAD83]"", """", , , vbTextCompare)" & vbNewLine
'  strBaseString = strBaseString & "              strDescription = strDescription & _" & vbNewLine
'  strBaseString = strBaseString & "                  ""  This field should only have a value if the 'KNF_Add_Attribute' value = """"TRUE"""".  "" & _" & vbNewLine
'  strBaseString = strBaseString & "                  ""[Imported from Roads_KNF_UTM12_NAD83, Field """""" & strKNF_RoadArray(lngKNFIndex, 1) & """"""]""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "              ' SET FIELD DEFINITION OF NEW FCLASS" & vbNewLine
'  strBaseString = strBaseString & "              varCheckPropertyPresent = pPropSet.GetProperty(""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdef"")" & vbNewLine
'  strBaseString = strBaseString & "              If Not IsEmpty(varCheckPropertyPresent) Then" & vbNewLine
'  strBaseString = strBaseString & "                pPropSet.RemoveProperty ""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdef""" & vbNewLine
'  strBaseString = strBaseString & "              End If" & vbNewLine
'  strBaseString = strBaseString & "              pPropSet.SetProperty ""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdef"", strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "              ' GET DESCRIPTION SOURCE FROM ORIGINAL FCLASS" & vbNewLine
'  strBaseString = strBaseString & "              varProperty = pRoads_KNF_UTM12_NAD83_PropSet.GetProperty(""eainfo/detailed/attr["" & CStr(lngIndex) & ""]/attrdefs"")" & vbNewLine
'  strBaseString = strBaseString & "              If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "                varSubProperty = varProperty(0)" & vbNewLine
'  strBaseString = strBaseString & "                If IsEmpty(varSubProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "                  strDescriptionSource = ""Roads_KNF_UTM12_NAD83""" & vbNewLine
'  strBaseString = strBaseString & "                Else" & vbNewLine
'  strBaseString = strBaseString & "                  strDescriptionSource = CStr(varSubProperty)" & vbNewLine
'  strBaseString = strBaseString & "                End If" & vbNewLine
'  strBaseString = strBaseString & "              Else" & vbNewLine
'  strBaseString = strBaseString & "                strDescriptionSource = ""Roads_KNF_UTM12_NAD83""" & vbNewLine
'  strBaseString = strBaseString & "              End If" & vbNewLine
'  strBaseString = strBaseString & "              strDescriptionSource = Replace(strDescriptionSource, ""Dicitonary"", ""Dictionary"", , , vbTextCompare)" & vbNewLine
'  strBaseString = strBaseString & "              strDescriptionSource = Replace(strDescriptionSource, ""20003"", ""2003"", , , vbTextCompare)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "              ' SET DESCRIPTION SOURCE DEFINITION OF NEW FCLASS" & vbNewLine
'  strBaseString = strBaseString & "              varCheckPropertyPresent = pPropSet.GetProperty(""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdefs"")" & vbNewLine
'  strBaseString = strBaseString & "              If Not IsEmpty(varCheckPropertyPresent) Then" & vbNewLine
'  strBaseString = strBaseString & "                pPropSet.RemoveProperty ""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdefs""" & vbNewLine
'  strBaseString = strBaseString & "              End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "              ' ONLY ADD THIS NEW ONE IF THERE WAS ONE IN THE ORIGINAL FCLASS" & vbNewLine
'  strBaseString = strBaseString & "              If strDescriptionSource <> """" Then" & vbNewLine
'  strBaseString = strBaseString & "                pPropSet.SetProperty ""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdefs"", strDescriptionSource" & vbNewLine
'  strBaseString = strBaseString & "              End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "              Exit For" & vbNewLine
'  strBaseString = strBaseString & "            End If" & vbNewLine
'  strBaseString = strBaseString & "          End If" & vbNewLine
'  strBaseString = strBaseString & "        Next lngIndex" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    Next lngKNFIndex" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "    ' ----------------------------------------------------------------------------------------------" & vbNewLine
'  strBaseString = strBaseString & "    ' <<<<<<<<<<<<  BLM_AZStrip_UTM12_NAD83  >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>" & vbNewLine
'  strBaseString = strBaseString & "    '-----------------------------------------------------------------------------------------------" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""BLM_Add_Attribute"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = """"""TRUE"""" or """"FALSE"""", whether there was a polyline feature from the 'BLM_AZStrip_UTM12_NAD83' "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""feature class within 100m of the centroid of this polyline.  If so, then this '"" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""BLM_AZStrip_UTM12_NAD83' feature would be considered a candidate for extracting attribute values "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""to describe this feature.""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""BLM_Attribute_Source"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""If 'BLM_Add_Attribute' above = """"TRUE"""", then this field will contain the name of the "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""BLM feature class (BLM_AZStrip_UTM12_NAD83).  Otherwise it should be Null.""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""BLM_Attribute_Distance"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""If 'BLM_Add_Attribute' above = """"TRUE"""", then this field will contain the distance "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""(in meters) from the centroid of this polyline to the nearest BLM feature.  Otherwise it should be Null.  "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""Values should always be <= 100m.""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.RemoveProperty strXDom" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomMin, ""0""" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomMax, ""100""" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXDomUnits, ""Meters""" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & "    If strFieldName = ""BLM_ObjectID"" Then" & vbNewLine
'  strBaseString = strBaseString & "      strDescription = ""If 'BLM_Add_Attribute' above = """"TRUE"""", then this field will contain the OBJECTID "" & _" & vbNewLine
'  strBaseString = strBaseString & "          ""value of the nearest BLM feature, from the BLM_AZStrip_UTM12_NAD83 feature class.""" & vbNewLine
'  strBaseString = strBaseString & "      If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.RemoveProperty strXName" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "      pPropSet.SetProperty strXName, strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "      ' SET SOURCE IF IT IS NOT ALREADY SET" & vbNewLine
'  strBaseString = strBaseString & "      varSource = pPropSet.GetProperty(strXNameSource)" & vbNewLine
'  strBaseString = strBaseString & "      If IsEmpty(varSource) Then" & vbNewLine
'  strBaseString = strBaseString & "        pPropSet.SetProperty strXNameSource, ""Lab of Landscape Ecology and Conservation Biology; "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""School of Earth Sciences and Environmental Sustainability; College of Engineering, "" & _" & vbNewLine
'  strBaseString = strBaseString & "            ""Forestry and Natural Sciences; Northern Arizona University; Flagstaff, AZ 86011""" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "    For lngBLMIndex = 0 To UBound(strBLMObsArray, 1)" & vbNewLine
'  strBaseString = strBaseString & "      If strFieldName = strBLMObsArray(lngBLMIndex, 0) Then '  ""BLM_OBS_USE1"", ""BLM_OBS_USE2"" or ""BLM_OBS_USE3""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "        ' FIND ORIGINAL FIELD METADATA" & vbNewLine
'  strBaseString = strBaseString & "        For lngIndex = 0 To pBLM_AZStrip_UTM12_NAD83_FClass.Fields.FieldCount" & vbNewLine
'  strBaseString = strBaseString & "          strNames(0) = ""eainfo/detailed/attr["" & CStr(lngIndex) & ""]""" & vbNewLine
'  strBaseString = strBaseString & "          pBLM_AZStrip_UTM12_NAD83_PropSet.GetProperties strNames, varVals" & vbNewLine
'  strBaseString = strBaseString & "          varSubVals = varVals(0)" & vbNewLine
'  strBaseString = strBaseString & "          If Not IsEmpty(varSubVals) Then" & vbNewLine
'  strBaseString = strBaseString & "            ' GET FIELD NAME" & vbNewLine
'  strBaseString = strBaseString & "            strOrigFClassXName = ""eainfo/detailed/attr["" & CStr(lngIndex) & ""]/attrlabl""" & vbNewLine
'  strBaseString = strBaseString & "            varProperty = pBLM_AZStrip_UTM12_NAD83_PropSet.GetProperty(strOrigFClassXName)" & vbNewLine
'  strBaseString = strBaseString & "            If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "              varSubProperty = varProperty(0)" & vbNewLine
'  strBaseString = strBaseString & "              If IsEmpty(varSubProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "                strOrigFieldName = ""<-- No Field Name -->""" & vbNewLine
'  strBaseString = strBaseString & "              Else" & vbNewLine
'  strBaseString = strBaseString & "                strOrigFieldName = CStr(varSubProperty)" & vbNewLine
'  strBaseString = strBaseString & "              End If" & vbNewLine
'  strBaseString = strBaseString & "            Else" & vbNewLine
'  strBaseString = strBaseString & "              strOrigFieldName = ""<-- No Field Name -->""" & vbNewLine
'  strBaseString = strBaseString & "            End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "            If strOrigFieldName = strBLMObsArray(lngBLMIndex, 1) Then ' ""OBS_USE1"", ""OBS_USE2"", or ""OBS_USE3""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "              If lngBLMIndex <= 4 Then" & vbNewLine
'  strBaseString = strBaseString & "                strOrigFClassXName = ""eainfo/detailed/attr["" & CStr(lngIndex) & ""]/attrdomv""" & vbNewLine
'  strBaseString = strBaseString & "                varProperty = pBLM_AZStrip_UTM12_NAD83_PropSet.GetProperty(strOrigFClassXName)" & vbNewLine
'  strBaseString = strBaseString & "                varSubProperty = varProperty(0)" & vbNewLine
'  strBaseString = strBaseString & "'                For lngIndex2 = 0 To UBound(varProperty)" & vbNewLine
'  strBaseString = strBaseString & "'                  Debug.Print ""       "" & CStr(lngIndex2) & ""] "" & CStr(varProperty(lngIndex2))" & vbNewLine
'  strBaseString = strBaseString & "'                Next lngIndex2" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "                ' GET DOMAIN ATTRIBUTES FROM ORIGINAL FCLASS" & vbNewLine
'  strBaseString = strBaseString & "                strOrigFClassXName = ""eainfo/detailed/attr["" & CStr(lngIndex) & ""]/attrdomv""" & vbNewLine
'  strBaseString = strBaseString & "                varProperty = pBLM_AZStrip_UTM12_NAD83_PropSet.GetProperty(strOrigFClassXName)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "                ' ADD DOMAIN ATTRIBUTES TO NEW FCLASS" & vbNewLine
'  strBaseString = strBaseString & "                varCheckPropertyPresent = pPropSet.GetProperty(""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdomv"")" & vbNewLine
'  strBaseString = strBaseString & "                If Not IsEmpty(varCheckPropertyPresent) Then" & vbNewLine
'  strBaseString = strBaseString & "                  pPropSet.RemoveProperty ""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdomv""" & vbNewLine
'  strBaseString = strBaseString & "                End If" & vbNewLine
'  strBaseString = strBaseString & "                For lngIndex3 = 0 To UBound(varProperty)" & vbNewLine
'  strBaseString = strBaseString & "                  strSubXName1 = ""eainfo/detailed/attr["" & CStr(lngIndex) & ""]/attrdomv/"" & _" & vbNewLine
'  strBaseString = strBaseString & "                      varProperty(lngIndex3) & ""["" & CStr(lngIndex3) & ""]""" & vbNewLine
'  strBaseString = strBaseString & "                  varSubProp1 = pBLM_AZStrip_UTM12_NAD83_PropSet.GetProperty(strSubXName1)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "                  For lngIndex4 = 0 To UBound(varSubProp1)" & vbNewLine
'  strBaseString = strBaseString & "                    strSubXName2 = ""eainfo/detailed/attr["" & CStr(lngIndex) & ""]/attrdomv/"" & _" & vbNewLine
'  strBaseString = strBaseString & "                        varProperty(lngIndex3) & ""["" & CStr(lngIndex3) & ""]/"" & varSubProp1(lngIndex4)" & vbNewLine
'  strBaseString = strBaseString & "                    varSubProp2 = pBLM_AZStrip_UTM12_NAD83_PropSet.GetProperty(strSubXName2)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "                    ' WRITE VALUE BACK TO NEW FCLASS" & vbNewLine
'  strBaseString = strBaseString & "                    strSubXName3 = ""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdomv/"" & _" & vbNewLine
'  strBaseString = strBaseString & "                        varProperty(lngIndex3) & ""["" & CStr(lngIndex3) & ""]/"" & varSubProp1(lngIndex4)" & vbNewLine
'  strBaseString = strBaseString & "                    strValue = varSubProp2(0)" & vbNewLine
'  strBaseString = strBaseString & "                    strValue = Replace(strValue, ""20003"", ""2003"")" & vbNewLine
'  strBaseString = strBaseString & "                    strValue = Replace(strValue, ""Dicitonary"", ""Dictionary"")" & vbNewLine
'  strBaseString = strBaseString & "                    pPropSet.SetProperty strSubXName3, strValue" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "                  Next lngIndex4" & vbNewLine
'  strBaseString = strBaseString & "                Next lngIndex3" & vbNewLine
'  strBaseString = strBaseString & "              Else" & vbNewLine
'  strBaseString = strBaseString & "'                strOrigFCLassXName = ""eainfo/detailed/attr["" & CStr(lngIndex) & ""]""" & vbNewLine
'  strBaseString = strBaseString & "'                varProperty = pBLM_AZStrip_UTM12_NAD83_PropSet.GetProperty(strOrigFCLassXName)" & vbNewLine
'  strBaseString = strBaseString & "'                varSubProperty = varProperty(0)" & vbNewLine
'  strBaseString = strBaseString & "'                Debug.Print ""Examining "" & strOrigFieldName & ""...""" & vbNewLine
'  strBaseString = strBaseString & "'                For lngIndex2 = 0 To UBound(varProperty)" & vbNewLine
'  strBaseString = strBaseString & "'                  varSubProp1 = pBLM_AZStrip_UTM12_NAD83_PropSet.GetProperty(strOrigFCLassXName & ""/"" & _" & vbNewLine
'  strBaseString = strBaseString & "'                        CStr(varProperty(lngIndex2)))" & vbNewLine
'  strBaseString = strBaseString & "'                  Debug.Print ""       "" & CStr(lngIndex2) & ""] "" & CStr(varProperty(lngIndex2)) & "" = "" & _" & vbNewLine
'  strBaseString = strBaseString & "'                        CStr(varSubProp1(0))" & vbNewLine
'  strBaseString = strBaseString & "'                Next lngIndex2" & vbNewLine
'  strBaseString = strBaseString & "              End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "              If strOrigFieldName = ""ROAD_NO_"" Then" & vbNewLine
'  strBaseString = strBaseString & "                Debug.Print ""Here...""" & vbNewLine
'  strBaseString = strBaseString & "              End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "              ' GET DEFINITION FROM ORIGINAL FCLASS" & vbNewLine
'  strBaseString = strBaseString & "              varProperty = pBLM_AZStrip_UTM12_NAD83_PropSet.GetProperty(""eainfo/detailed/attr["" & CStr(lngIndex) & ""]/attrdef"")" & vbNewLine
'  strBaseString = strBaseString & "              If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "                varSubProperty = varProperty(0)" & vbNewLine
'  strBaseString = strBaseString & "                If IsEmpty(varSubProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "                  strDescription = ""<-- No Description -->""" & vbNewLine
'  strBaseString = strBaseString & "                Else" & vbNewLine
'  strBaseString = strBaseString & "                  strDescription = CStr(varSubProperty)" & vbNewLine
'  strBaseString = strBaseString & "                End If" & vbNewLine
'  strBaseString = strBaseString & "              Else" & vbNewLine
'  strBaseString = strBaseString & "                strDescription = ""<-- No Description -->""" & vbNewLine
'  strBaseString = strBaseString & "              End If" & vbNewLine
'  strBaseString = strBaseString & "              strDescription = Replace(strDescription, ""  [Imported from TransportationLine_UTM12_NAD83]"", """", , , vbTextCompare)" & vbNewLine
'  strBaseString = strBaseString & "              strDescription = strDescription & _" & vbNewLine
'  strBaseString = strBaseString & "                  ""  This field should only have a value if the 'BLM_Add_Attribute' value = """"TRUE"""".  "" & _" & vbNewLine
'  strBaseString = strBaseString & "                  ""[Imported from BLM_AZStrip_UTM12_NAD83, Field """""" & strBLMObsArray(lngBLMIndex, 1) & """"""]""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "              ' SET FIELD DEFINITION OF NEW FCLASS" & vbNewLine
'  strBaseString = strBaseString & "              varCheckPropertyPresent = pPropSet.GetProperty(""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdef"")" & vbNewLine
'  strBaseString = strBaseString & "              If Not IsEmpty(varCheckPropertyPresent) Then" & vbNewLine
'  strBaseString = strBaseString & "                pPropSet.RemoveProperty ""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdef""" & vbNewLine
'  strBaseString = strBaseString & "              End If" & vbNewLine
'  strBaseString = strBaseString & "              pPropSet.SetProperty ""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdef"", strDescription" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "              ' GET DESCRIPTION SOURCE FROM ORIGINAL FCLASS" & vbNewLine
'  strBaseString = strBaseString & "              varProperty = pBLM_AZStrip_UTM12_NAD83_PropSet.GetProperty(""eainfo/detailed/attr["" & CStr(lngIndex) & ""]/attrdefs"")" & vbNewLine
'  strBaseString = strBaseString & "              If Not IsEmpty(varProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "                varSubProperty = varProperty(0)" & vbNewLine
'  strBaseString = strBaseString & "                If IsEmpty(varSubProperty) Then" & vbNewLine
'  strBaseString = strBaseString & "                  strDescriptionSource = ""BLM_AZStrip_UTM12_NAD83""" & vbNewLine
'  strBaseString = strBaseString & "                Else" & vbNewLine
'  strBaseString = strBaseString & "                  strDescriptionSource = CStr(varSubProperty)" & vbNewLine
'  strBaseString = strBaseString & "                End If" & vbNewLine
'  strBaseString = strBaseString & "              Else" & vbNewLine
'  strBaseString = strBaseString & "                strDescriptionSource = ""BLM_AZStrip_UTM12_NAD83""" & vbNewLine
'  strBaseString = strBaseString & "              End If" & vbNewLine
'  strBaseString = strBaseString & "              strDescriptionSource = Replace(strDescriptionSource, ""Dicitonary"", ""Dictionary"", , , vbTextCompare)" & vbNewLine
'  strBaseString = strBaseString & "              strDescriptionSource = Replace(strDescriptionSource, ""20003"", ""2003"", , , vbTextCompare)" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "              ' SET DESCRIPTION SOURCE DEFINITION OF NEW FCLASS" & vbNewLine
'  strBaseString = strBaseString & "              varCheckPropertyPresent = pPropSet.GetProperty(""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdefs"")" & vbNewLine
'  strBaseString = strBaseString & "              If Not IsEmpty(varCheckPropertyPresent) Then" & vbNewLine
'  strBaseString = strBaseString & "                pPropSet.RemoveProperty ""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdefs""" & vbNewLine
'  strBaseString = strBaseString & "              End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "              ' ONLY ADD THIS NEW ONE IF THERE WAS ONE IN THE ORIGINAL FCLASS" & vbNewLine
'  strBaseString = strBaseString & "              If strDescriptionSource <> """" Then" & vbNewLine
'  strBaseString = strBaseString & "                pPropSet.SetProperty ""eainfo/detailed/attr["" & CStr(lngFieldIndex) & ""]/attrdefs"", strDescriptionSource" & vbNewLine
'  strBaseString = strBaseString & "              End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "              Exit For" & vbNewLine
'  strBaseString = strBaseString & "            End If" & vbNewLine
'  strBaseString = strBaseString & "          End If" & vbNewLine
'  strBaseString = strBaseString & "        Next lngIndex" & vbNewLine
'  strBaseString = strBaseString & "      End If" & vbNewLine
'  strBaseString = strBaseString & "    Next lngBLMIndex" & vbNewLine
'  strBaseString = strBaseString & "  End If" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.RemoveProperty ""eainfo/detailed/subtype""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[0]/stname"", ""Administrative""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[0]/stcode"", ""1""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[0]/stfield[0]/stfldnm"", ""ROADS_SurfaceMaterial""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[0]/stfield[0]/stflddd/domname"", ""SurfaceMaterialDomain""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[0]/stfield[0]/stflddd/domdesc"", ""Description""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[0]/stfield[0]/stflddd/domtype"", ""Coded Value""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[0]/stfield[0]/stflddd/mrgtype"", ""Default value""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[0]/stfield[0]/stflddd/splttype"", ""Default value""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[0]/stfield[0]/stflddd/domowner"", """"" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[0]/stfield[0]/stflddd/domfldtp"", ""String""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[0]/stfield[1]/stfldnm"", ""ROADS_Status""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[0]/stfield[1]/stflddd/domname"", ""RoadStatusDomain""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[0]/stfield[1]/stflddd/domdesc"", ""Description""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[0]/stfield[1]/stflddd/domtype"", ""Coded Value""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[0]/stfield[1]/stflddd/mrgtype"", ""Default value""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[0]/stfield[1]/stflddd/splttype"", ""Default value""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[0]/stfield[1]/stflddd/domowner"", """"" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[0]/stfield[1]/stflddd/domfldtp"", ""String""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[0]/stfield[2]/stfldnm"", ""ROADS_Compendium""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[0]/stfield[2]/stflddd/domname"", ""CompendiumDomain""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[0]/stfield[2]/stflddd/domdesc"", ""Description""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[0]/stfield[2]/stflddd/domtype"", ""Coded Value""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[0]/stfield[2]/stflddd/mrgtype"", ""Default value""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[0]/stfield[2]/stflddd/splttype"", ""Default value""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[0]/stfield[2]/stflddd/domowner"", """"" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[0]/stfield[2]/stflddd/domfldtp"", ""String""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[0]/stfield[3]/stfldnm"", ""WildernessRecommendation""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[0]/stfield[3]/stflddd/domname"", ""WildernessRecommendDomain""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[0]/stfield[3]/stflddd/domdesc"", ""Description""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[0]/stfield[3]/stflddd/domtype"", ""Coded Value""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[0]/stfield[3]/stflddd/mrgtype"", ""Default value""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[0]/stfield[3]/stflddd/splttype"", ""Default value""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[0]/stfield[3]/stflddd/domowner"", """"" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[0]/stfield[3]/stflddd/domfldtp"", ""String""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[0]/stfield[4]/stfldnm"", ""WildernessMap_1980""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[0]/stfield[4]/stflddd/domname"", ""WildMap1980Domain""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[0]/stfield[4]/stflddd/domdesc"", ""Description""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[0]/stfield[4]/stflddd/domtype"", ""Coded Value""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[0]/stfield[4]/stflddd/mrgtype"", ""Default value""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[0]/stfield[4]/stflddd/splttype"", ""Default value""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[0]/stfield[4]/stflddd/domowner"", """"" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[0]/stfield[4]/stflddd/domfldtp"", ""String""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[1]/stname"", ""Closed""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[1]/stcode"", ""2""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[1]/stfield[0]/stfldnm"", ""ROADS_SurfaceMaterial""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[1]/stfield[0]/stflddd/domname"", ""SurfaceMaterialDomain""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[1]/stfield[0]/stflddd/domdesc"", ""Description""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[1]/stfield[0]/stflddd/domtype"", ""Coded Value""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[1]/stfield[0]/stflddd/mrgtype"", ""Default value""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[1]/stfield[0]/stflddd/splttype"", ""Default value""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[1]/stfield[0]/stflddd/domowner"", """"" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[1]/stfield[0]/stflddd/domfldtp"", ""String""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[1]/stfield[1]/stfldnm"", ""ROADS_Status""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[1]/stfield[1]/stflddd/domname"", ""RoadStatusDomain""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[1]/stfield[1]/stflddd/domdesc"", ""Description""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[1]/stfield[1]/stflddd/domtype"", ""Coded Value""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[1]/stfield[1]/stflddd/mrgtype"", ""Default value""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[1]/stfield[1]/stflddd/splttype"", ""Default value""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[1]/stfield[1]/stflddd/domowner"", """"" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[1]/stfield[1]/stflddd/domfldtp"", ""String""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[1]/stfield[2]/stfldnm"", ""ROADS_Compendium""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[1]/stfield[2]/stflddd/domname"", ""CompendiumDomain""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[1]/stfield[2]/stflddd/domdesc"", ""Description""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[1]/stfield[2]/stflddd/domtype"", ""Coded Value""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[1]/stfield[2]/stflddd/mrgtype"", ""Default value""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[1]/stfield[2]/stflddd/splttype"", ""Default value""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[1]/stfield[2]/stflddd/domowner"", """"" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[1]/stfield[2]/stflddd/domfldtp"", ""String""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[1]/stfield[3]/stfldnm"", ""WildernessRecommendation""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[1]/stfield[3]/stflddd/domname"", ""WildernessRecommendDomain""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[1]/stfield[3]/stflddd/domdesc"", ""Description""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[1]/stfield[3]/stflddd/domtype"", ""Coded Value""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[1]/stfield[3]/stflddd/mrgtype"", ""Default value""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[1]/stfield[3]/stflddd/splttype"", ""Default value""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[1]/stfield[3]/stflddd/domowner"", """"" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[1]/stfield[3]/stflddd/domfldtp"", ""String""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[1]/stfield[4]/stfldnm"", ""WildernessMap_1980""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[1]/stfield[4]/stflddd/domname"", ""WildMap1980Domain""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[1]/stfield[4]/stflddd/domdesc"", ""Description""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[1]/stfield[4]/stflddd/domtype"", ""Coded Value""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[1]/stfield[4]/stflddd/mrgtype"", ""Default value""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[1]/stfield[4]/stflddd/splttype"", ""Default value""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[1]/stfield[4]/stflddd/domowner"", """"" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[1]/stfield[4]/stflddd/domfldtp"", ""String""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[2]/stname"", ""Public""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[2]/stcode"", ""3""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[2]/stfield[0]/stfldnm"", ""ROADS_SurfaceMaterial""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[2]/stfield[0]/stflddd/domname"", ""SurfaceMaterialDomain""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[2]/stfield[0]/stflddd/domdesc"", ""Description""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[2]/stfield[0]/stflddd/domtype"", ""Coded Value""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[2]/stfield[0]/stflddd/mrgtype"", ""Default value""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[2]/stfield[0]/stflddd/splttype"", ""Default value""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[2]/stfield[0]/stflddd/domowner"", """"" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[2]/stfield[0]/stflddd/domfldtp"", ""String""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[2]/stfield[1]/stfldnm"", ""ROADS_Status""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[2]/stfield[1]/stflddd/domname"", ""RoadStatusDomain""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[2]/stfield[1]/stflddd/domdesc"", ""Description""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[2]/stfield[1]/stflddd/domtype"", ""Coded Value""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[2]/stfield[1]/stflddd/mrgtype"", ""Default value""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[2]/stfield[1]/stflddd/splttype"", ""Default value""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[2]/stfield[1]/stflddd/domowner"", """"" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[2]/stfield[1]/stflddd/domfldtp"", ""String""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[2]/stfield[2]/stfldnm"", ""ROADS_Compendium""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[2]/stfield[2]/stflddd/domname"", ""CompendiumDomain""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[2]/stfield[2]/stflddd/domdesc"", ""Description""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[2]/stfield[2]/stflddd/domtype"", ""Coded Value""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[2]/stfield[2]/stflddd/mrgtype"", ""Default value""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[2]/stfield[2]/stflddd/splttype"", ""Default value""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[2]/stfield[2]/stflddd/domowner"", """"" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[2]/stfield[2]/stflddd/domfldtp"", ""String""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[2]/stfield[3]/stfldnm"", ""WildernessRecommendation""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[2]/stfield[3]/stflddd/domname"", ""WildernessRecommendDomain""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[2]/stfield[3]/stflddd/domdesc"", ""Description""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[2]/stfield[3]/stflddd/domtype"", ""Coded Value""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[2]/stfield[3]/stflddd/mrgtype"", ""Default value""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[2]/stfield[3]/stflddd/splttype"", ""Default value""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[2]/stfield[3]/stflddd/domowner"", """"" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[2]/stfield[3]/stflddd/domfldtp"", ""String""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[2]/stfield[4]/stfldnm"", ""WildernessMap_1980""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[2]/stfield[4]/stflddd/domname"", ""WildMap1980Domain""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[2]/stfield[4]/stflddd/domdesc"", ""Description""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[2]/stfield[4]/stflddd/domtype"", ""Coded Value""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[2]/stfield[4]/stflddd/mrgtype"", ""Default value""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[2]/stfield[4]/stflddd/splttype"", ""Default value""" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[2]/stfield[4]/stflddd/domowner"", """"" & vbNewLine
'  strBaseString = strBaseString & "  pPropSet.SetProperty ""eainfo/detailed/subtype[2]/stfield[4]/stflddd/domfldtp"", ""String""" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "ClearMemory:" & vbNewLine
'  strBaseString = strBaseString & "  Set pRoads_KNF_UTM12_NAD83_FClass = Nothing" & vbNewLine
'  strBaseString = strBaseString & "  Set pTIGER_2012_Trans_Near_Park_UTM12_NAD83_FClass = Nothing" & vbNewLine
'  strBaseString = strBaseString & "  Set pBLM_AZStrip_UTM12_NAD83_FClass = Nothing" & vbNewLine
'  strBaseString = strBaseString & "  Set pRoads_UTM12_NAD83_FClass = Nothing" & vbNewLine
'  strBaseString = strBaseString & "  Set pTrails_UTM12_NAD83_FClass = Nothing" & vbNewLine
'  strBaseString = strBaseString & "  Set pRoutes_UTM12_NAD83_FClass = Nothing" & vbNewLine
'  strBaseString = strBaseString & "  Set pUtah_Roads_Near_AOI_UTM12_NAD83_FClass = Nothing" & vbNewLine
'  strBaseString = strBaseString & "  Set pTrails_KNF_UTM12_NAD83_FClass = Nothing" & vbNewLine
'  strBaseString = strBaseString & "  Set pTransportationLine_UTM12_NAD83_FClass = Nothing" & vbNewLine
'  strBaseString = strBaseString & "  Set pRoads_KNF_UTM12_NAD83_PropSet = Nothing" & vbNewLine
'  strBaseString = strBaseString & "  Set pTIGER_2012_Trans_Near_Park_UTM12_NAD83_PropSet = Nothing" & vbNewLine
'  strBaseString = strBaseString & "  Set pBLM_AZStrip_UTM12_NAD83_PropSet = Nothing" & vbNewLine
'  strBaseString = strBaseString & "  Set pRoads_UTM12_NAD83_PropSet = Nothing" & vbNewLine
'  strBaseString = strBaseString & "  Set pTrails_UTM12_NAD83_PropSet = Nothing" & vbNewLine
'  strBaseString = strBaseString & "  Set pRoutes_UTM12_NAD83_PropSet = Nothing" & vbNewLine
'  strBaseString = strBaseString & "  Set pUtah_Roads_Near_AOI_UTM12_NAD83_PropSet = Nothing" & vbNewLine
'  strBaseString = strBaseString & "  Set pTrails_KNF_UTM12_NAD83_PropSet = Nothing" & vbNewLine
'  strBaseString = strBaseString & "  Set pTransportationLine_UTM12_NAD83_PropSet = Nothing" & vbNewLine
'  strBaseString = strBaseString & "  Set pSubArray = Nothing" & vbNewLine
'  strBaseString = strBaseString & "  Erase strNames" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine
'  strBaseString = strBaseString & "End Sub" & vbNewLine
'  strBaseString = strBaseString & " " & vbNewLine

End Sub
Public Function ExtractAttributesFromField(pDataset As IDataset, strFieldName As String, _
  strFieldDescription As String, strFieldDescriptionSource As String, _
  Optional strRDOMFieldMin As String, Optional strRDOMFieldMax As String, _
  Optional strRDOMFieldMean As String, Optional strRDOMFieldUnit As String, _
  Optional strRDOMFieldStDev As String, Optional strRDOMFieldMinResolution As String, _
  Optional strUDOM_DescriptionOfValues As String, _
  Optional strCodesetNameOfList As String, Optional strCodesetSource As String) As String

  On Error GoTo ErrHandler
  ExtractAttributesFromField = "Succeeded"

  ' rdom = RANGE DOMAIN
  ' edim = ENUMERATED DOMAIN
  ' udom = UNREPRESENTABLE DOMAIN
  ' codesetd = CODESET DOMAIN

'  strResponse = Metadata_Functions.AddResourceMaintenance(pDataset, JenMetadata_Ongoing)
'  Debug.Print "Added Resource Status = " & strResponse

  ' IF varStringArrayOfValueDescSourceList EXISTS, IT SHOULD CONTAIN A STRING
  ' ARRAY WITH DIMENSIONS (2,X), WHERE X IS NUMBER OF ELEMENTS IN LIST.
  ' THIS IS ZERO-BASED, SO "2" MEANS 3 ATTRIBUTES PER ELEMENT (LIST ITEM, DESCRIPTION, SOURCE)
  Dim booAddList As Boolean
  booAddList = False

'  Dim strListArray() As String
'  If Not IsNull(varEDOMArrayOfList_ValueDescSource) Then
'    strListArray = varEDOMArrayOfList_ValueDescSource(0)
'    If UBound(strListArray, 1) <> 2 Then
'      MsgBox "Array has incorrect dimensions.  Skipping this item..."
'    Else
'      booAddList = True
'    End If
'  End If

  Dim pPropSet As IPropertySet
  Set pPropSet = ReturnMetadataPropSetFromDataset(pDataset)
  Dim pXMLPropSet As IXmlPropertySet2
  Set pXMLPropSet = pPropSet

  ' GET INDEX NUMBER
  Dim lngFieldIndex As Long
  Dim booFailed As Boolean
  lngFieldIndex = ReturnAttributeFieldXPathIndex(pDataset, strFieldName, booFailed)
  If lngFieldIndex = -1 Then
    If booFailed Then
      ExtractAttributesFromField = "ReturnAttributeFieldXPathIndex Failed"
    Else
      ExtractAttributesFromField = "No Field Found"
    End If
    GoTo ClearMemory
  End If

  Dim strXName As String
  Dim strXNameSource As String
  Dim strXDomMin As String
  Dim strXDomMax As String
  Dim strXDomUnits As String
  Dim strXDom As String
  Dim strXRDom As String
  Dim strXUDom As String
  Dim strXEDom As String
  Dim strXCodesetDom As String

  ' see http://resources.arcgis.com/en/help/main/10.1/index.html#//003t00000037000000
  ' GET CURRENT DESCRIPTION
  strXName = "eainfo/detailed/attr[" & CStr(lngFieldIndex) & "]/attrdef"         ' DESCRIPTION OF FIELD
  strXNameSource = "eainfo/detailed/attr[" & CStr(lngFieldIndex) & "]/attrdefs"   ' DESCRIPTION SOURCE
  strXRDom = "eainfo/detailed/attr[" & CStr(lngFieldIndex) & "]/attrdomv/rdom"  ' RANGE DOMAIN IN GENERAL
  strXDomMin = "eainfo/detailed/attr[" & CStr(lngFieldIndex) & "]/attrdomv/rdom/rdommin"  ' MINIMUM VALUE
  strXDomMax = "eainfo/detailed/attr[" & CStr(lngFieldIndex) & "]/attrdomv/rdom/rdommax"  ' MAXIMUM VALUE
  strXDomUnits = "eainfo/detailed/attr[" & CStr(lngFieldIndex) & "]/attrdomv/rdom/attrunit"   ' UNITS
  strXDom = "eainfo/detailed/attr[" & CStr(lngFieldIndex) & "]/attrdomv"
  strXUDom = "eainfo/detailed/attr[" & CStr(lngFieldIndex) & "]/attrdomv/udom"  ' DESCRIPTION OF VALUES
  strXEDom = "eainfo/detailed/attr[" & CStr(lngFieldIndex) & "]/attrdomv/edom"  ' LIST OF VALUES
  strXCodesetDom = "eainfo/detailed/attr[" & CStr(lngFieldIndex) & "]/attrdomv/codesetd"  ' GENERAL CODESET DOMAIN
                                                                          ' /edomv = Value
                                                                          ' /edomvd = Description of Value
                                                                          ' /edomvds = Enumerated domain value definition source
  ' rdom = RANGE DOMAIN
  ' edim = ENUMERATED DOMAIN
  ' udom = UNREPRESENTABLE DOMAIN
  ' codesetd = CODESET DOMAIN
'  Dim varVal As Variant
  
'  varVal = pPropSet.GetProperty(strXName)
'  strFieldDescription = varVal(0)
'  varVal = pPropSet.GetProperty(strXNameSource)
'  strFieldDescriptionSource = varVal(0)
'  varVal = pPropSet.GetProperty(strXRDom & "/rdommin")
'  strRDOMFieldMin = varVal(0)
'  varVal = pPropSet.GetProperty(strXRDom & "/rdommax")
'  strRDOMFieldMax = varVal(0)
'  varVal = pPropSet.GetProperty(strXRDom & "/rdommean")
'  strRDOMFieldMean = varVal(0)
'  varVal = pPropSet.GetProperty(strXRDom & "/attrunit")
'  strRDOMFieldUnit = varVal(0)
'  varVal = pPropSet.GetProperty(strXRDom & "/rdomstdv")
'  strRDOMFieldStDev = varVal(0)
'  varVal = pPropSet.GetProperty(strXRDom & "/attrmres")
'  strRDOMFieldMinResolution = varVal(0)
'  varVal = pPropSet.GetProperty(strXUDom)
'  strUDOM_DescriptionOfValues = varVal(0)
'  varVal = pPropSet.GetProperty(strXCodesetDom & "/codesetn")
'  strCodesetNameOfList = varVal(0)
'  varVal = pPropSet.GetProperty(strXCodesetDom & "/codesets")
'  strCodesetSource = varVal(0)
  
  strFieldDescription = ReturnPropertyFromPropSet(pPropSet, strXName)
  strFieldDescriptionSource = ReturnPropertyFromPropSet(pPropSet, strXNameSource)
  strRDOMFieldMin = ReturnPropertyFromPropSet(pPropSet, strXRDom & "/rdommin")
  strRDOMFieldMax = ReturnPropertyFromPropSet(pPropSet, strXRDom & "/rdommax")
  strRDOMFieldMean = ReturnPropertyFromPropSet(pPropSet, strXRDom & "/rdommean")
  strRDOMFieldUnit = ReturnPropertyFromPropSet(pPropSet, strXRDom & "/attrunit")
  strRDOMFieldStDev = ReturnPropertyFromPropSet(pPropSet, strXRDom & "/rdomstdv")
  strRDOMFieldMinResolution = ReturnPropertyFromPropSet(pPropSet, strXRDom & "/attrmres")
  strUDOM_DescriptionOfValues = ReturnPropertyFromPropSet(pPropSet, strXUDom)
  strCodesetNameOfList = ReturnPropertyFromPropSet(pPropSet, strXCodesetDom & "/codesetn")
  strCodesetSource = ReturnPropertyFromPropSet(pPropSet, strXCodesetDom & "/codesets")

'  strFieldDescription = pPropSet.GetProperty(strXName)
'  strFieldDescriptionSource = pPropSet.GetProperty(strXNameSource)
'  strRDOMFieldMin = pPropSet.GetProperty(strXRDom & "/rdommin")
'  strRDOMFieldMax = pPropSet.GetProperty(strXRDom & "/rdommax")
'  strRDOMFieldMean = pPropSet.GetProperty(strXRDom & "/rdommean")
'  strRDOMFieldUnit = pPropSet.GetProperty(strXRDom & "/attrunit")
'  strRDOMFieldStDev = pPropSet.GetProperty(strXRDom & "/rdomstdv")
'  strRDOMFieldMinResolution = pPropSet.GetProperty(strXRDom & "/attrmres")
'  strUDOM_DescriptionOfValues = pPropSet.GetProperty(strXUDom)
'  strCodesetNameOfList = pPropSet.GetProperty(strXCodesetDom & "/codesetn")
'  strCodesetSource = pPropSet.GetProperty(strXCodesetDom & "/codesets")

'  Dim lngIndex As Long
'  Dim strValue As String
'  Dim strDescription As String
'  Dim strSource As String
'  Dim lngCounter As Long
'  lngCounter = -1
'
'  If booAddList Then
'    For lngIndex = 0 To UBound(strListArray, 2)
'      strValue = Trim(strListArray(0, lngIndex))
'      strDescription = Trim(strListArray(1, lngIndex))
'      strSource = Trim(strListArray(2, lngIndex))
'
'      If strValue <> "" Or strDescription <> "" Or strSource <> "" Then
'        lngCounter = lngCounter + 1
'        If strValue <> "" Then pPropSet.SetProperty strXEDom & "[" & CStr(lngCounter) & "]/edomv", strValue
'        If strDescription <> "" Then pPropSet.SetProperty strXEDom & "[" & CStr(lngCounter) & "]/edomvd", strDescription
'        If strSource <> "" Then pPropSet.SetProperty strXEDom & "[" & CStr(lngCounter) & "]/edomvds", strSource
'      End If
'    Next lngIndex
'  End If

'  strFieldDescription As String, strFieldDescriptionSource As String, _
'  Optional strRDOMFieldMin As String, Optional strRDOMFieldMax As String, _
'  Optional strRDOMFieldMean As String, Optional strRDOMFieldUnit As String, _
'  Optional strRDOMFieldStDev As String, Optional strRDOMFieldMinResolution As String, _
'  Optional varEDOMArrayOfList_ValueDescSource As Variant = Null, _
'  Optional strUDOM_DescriptionOfValues As String

  GoTo ClearMemory
  Exit Function

ErrHandler:
  ExtractAttributesFromField = "Failed"

ClearMemory:
'  Erase strListArray
  Set pPropSet = Nothing
  Set pXMLPropSet = Nothing



End Function

Private Function ReturnPropertyFromPropSet(pPropSet As IPropertySet, strName As String) As String
  On Error GoTo ErrHandler
  Dim varVal As Variant
  varVal = pPropSet.GetProperty(strName)
  If VarType(varVal(0)) = vbString Then
    ReturnPropertyFromPropSet = CStr(varVal(0))
  Else
    ReturnPropertyFromPropSet = ""
  End If
  
  GoTo ClearMemory
  Exit Function
ErrHandler:
  ReturnPropertyFromPropSet = ""

ClearMemory:
  varVal = Null

End Function


Public Function AddMetadataUseLimitations(pDataset As IDataset, strUseLimitations As String) As String ', _
   ' lngReplaceOrAdd As esriXmlSetPropertyAction) As String

  On Error GoTo ErrHandler

'  Dim strResponse As String
'  Dim strUseLimitations As String
'  strUseLimitations = "This dataset represents points along the bottleneck route.  This bottleneck route " & _
'    "describes the path between the two habitat blocks 'aaa' and 'bbb', within the corridor polygon 'ccc'., " & _
'    "which follows the route with the widest narrow point."
'  strResponse = Metadata_Functions.AddMetadataUseLimitations(pDataset, strUseLimitations)
'  Debug.Print "Saving UseLimitations: " & strResponse

  AddMetadataUseLimitations = "Succeeded"
  
  strUseLimitations = "<DIV STYLE=""text-align:Left;""><DIV><P><SPAN>" & strUseLimitations & "</SPAN></P></DIV></DIV>"
  
  Dim strUseLimitationsXPath As String
  strUseLimitationsXPath = "dataIdInfo/resConst/Consts/useLimit"

'  Dim lngLineageIndex As Long
  Dim pPropSet As IPropertySet
  Set pPropSet = ReturnMetadataPropSetFromDataset(pDataset)
  
'  lngLineageIndex = ReturnLargestIndexValue(strUseLimitationsXPath, pDataset) + 1
'
'  pPropSet.RemoveProperty "dataIdInfo/resConst/Consts/useLimit[" & CStr(lngLineageIndex) & "]"
'  pPropSet.SetProperty "dataIdInfo/resConst/Consts/useLimit[[" & CStr(lngLineageIndex) & "]", strUseLimitations

  pPropSet.SetProperty strUseLimitationsXPath, strUseLimitations

  Metadata_Functions.SaveMetadata pDataset, pPropSet

  GoTo ClearMemory
  Exit Function

ErrHandler:
  AddMetadataUseLimitations = "Failed"

ClearMemory:
  Set pPropSet = Nothing

End Function
Public Function SetMetadataCredits(pDataset As IDataset, strCredits As String) As String ', _
   ' lngReplaceOrAdd As esriXmlSetPropertyAction) As String

  On Error GoTo ErrHandler

'  Dim strResponse As String
'  Dim strCredits As String
'  strCredits = "This dataset represents points along the bottleneck route.  This bottleneck route " & _
'    "describes the path between the two habitat blocks 'aaa' and 'bbb', within the corridor polygon 'ccc'., " & _
'    "which follows the route with the widest narrow point."
'  strResponse = Metadata_Functions.SetMetadataCredits(pDataset, strCredits)
'  Debug.Print "Saving Credits: " & strResponse

  SetMetadataCredits = "Succeeded"

  Dim strCreditsXPath As String
  strCreditsXPath = "dataIdInfo/idCredit"

  Dim pPropSet As IPropertySet
  Set pPropSet = ReturnMetadataPropSetFromDataset(pDataset)

  pPropSet.SetProperty strCreditsXPath, strCredits

  Metadata_Functions.SaveMetadata pDataset, pPropSet

  GoTo ClearMemory
  Exit Function

ErrHandler:
  SetMetadataCredits = "Failed"

ClearMemory:
  Set pPropSet = Nothing

End Function




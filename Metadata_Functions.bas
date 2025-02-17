Attribute VB_Name = "Metadata_Functions"
Option Explicit

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

Public Function SetMetadataFormatVersion(pDataset As IDataset, _
  Optional strFormatVersion As String, _
  Optional booInsertArcGISVersionAutomatically As Boolean = False, _
  Optional pMxDoc As IMxDocument) As String
    On Error GoTo ErrHandler

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
  booFailed = False

  Dim pPropSet As IPropertySet
  Set pPropSet = ReturnMetadataPropSetFromDataset(pDataset)

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

  "    JenMetadata_Postal, True)

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

  Dim varVals As Variant

  booFailed = False

  Dim booFoundDuplicate As Boolean
  booFoundDuplicate = False
  Dim booFoundDuplicateInStep As Boolean
  Dim strTestVal As String
  Dim lngIndex As Long

  For lngIndex = 0 To lngMaxIndex
    booFoundDuplicateInStep = True

    varVals = pPropSet.GetProperty(strXPath & "[" & CStr(lngIndex) & "]/rpIndName")
    If IsEmpty(varVals) Then
      If strIndividualName <> "" Then booFoundDuplicateInStep = False
    Else
      strTestVal = CStr(varVals(0))
      If Trim(strIndividualName) <> Trim(strTestVal) Then booFoundDuplicateInStep = False
    End If

    If booFoundDuplicateInStep Then
      varVals = pPropSet.GetProperty(strXPath & "[" & CStr(lngIndex) & "]/rpOrgName")
      If IsEmpty(varVals) Then
        If strOrganizationName <> "" Then booFoundDuplicateInStep = False
      Else
        strTestVal = CStr(varVals(0))
        If Trim(strOrganizationName) <> Trim(strTestVal) Then booFoundDuplicateInStep = False
      End If
    End If

    If booFoundDuplicateInStep Then
      varVals = pPropSet.GetProperty(strXPath & "[" & CStr(lngIndex) & "]/rpPosName")
      If IsEmpty(varVals) Then
        If strPositionName <> "" Then booFoundDuplicateInStep = False
      Else
        strTestVal = CStr(varVals(0))
        If Trim(strPositionName) <> Trim(strTestVal) Then booFoundDuplicateInStep = False
      End If
    End If

    If booFoundDuplicateInStep Then
      varVals = pPropSet.GetProperty(strXPath & "[" & CStr(lngIndex) & "]/rpCntInfo/cntPhone/voiceNum")
      If IsEmpty(varVals) Then
        If strVoiceNumber <> "" Then booFoundDuplicateInStep = False
      Else
        strTestVal = CStr(varVals(0))
        If Trim(strVoiceNumber) <> Trim(strTestVal) Then booFoundDuplicateInStep = False
      End If
    End If

    If booFoundDuplicateInStep Then
      varVals = pPropSet.GetProperty(strXPath & "[" & CStr(lngIndex) & "]/rpCntInfo/cntAddress/delPoint")
      If IsEmpty(varVals) Then
        If strAddressStreet <> "" Then booFoundDuplicateInStep = False
      Else
        strTestVal = CStr(varVals(0))
        If Trim(strAddressStreet) <> Trim(strTestVal) Then booFoundDuplicateInStep = False
      End If
    End If

    If booFoundDuplicateInStep Then
      varVals = pPropSet.GetProperty(strXPath & "[" & CStr(lngIndex) & "]/rpCntInfo/cntAddress/city")
      If IsEmpty(varVals) Then
        If strAddressCity <> "" Then booFoundDuplicateInStep = False
      Else
        strTestVal = CStr(varVals(0))
        If Trim(strAddressCity) <> Trim(strTestVal) Then booFoundDuplicateInStep = False
      End If
    End If

    If booFoundDuplicateInStep Then
      varVals = pPropSet.GetProperty(strXPath & "[" & CStr(lngIndex) & "]/rpCntInfo/cntAddress/adminArea")
      If IsEmpty(varVals) Then
        If strAddressState <> "" Then booFoundDuplicateInStep = False
      Else
        strTestVal = CStr(varVals(0))
        If Trim(strAddressState) <> Trim(strTestVal) Then booFoundDuplicateInStep = False
      End If
    End If

    If booFoundDuplicateInStep Then
      varVals = pPropSet.GetProperty(strXPath & "[" & CStr(lngIndex) & "]/rpCntInfo/cntAddress/postCode")
      If IsEmpty(varVals) Then
        If strAddressZip <> "" Then booFoundDuplicateInStep = False
      Else
        strTestVal = CStr(varVals(0))
        If Trim(strAddressZip) <> Trim(strTestVal) Then booFoundDuplicateInStep = False
      End If
    End If

    If booFoundDuplicateInStep Then
      varVals = pPropSet.GetProperty(strXPath & "[" & CStr(lngIndex) & "]/rpCntInfo/cntAddress/country")
      If IsEmpty(varVals) Then
        If strAddressCountry <> "" Then booFoundDuplicateInStep = False
      Else
        strTestVal = CStr(varVals(0))
        If Trim(strAddressCountry) <> Trim(strTestVal) Then booFoundDuplicateInStep = False
      End If
    End If

    If booFoundDuplicateInStep Then
      varVals = pPropSet.GetProperty(strXPath & "[" & CStr(lngIndex) & "]/rpCntInfo/cntAddress/eMailAdd")
      If IsEmpty(varVals) Then
        If strAddressEmail <> "" Then booFoundDuplicateInStep = False
      Else
        strTestVal = CStr(varVals(0))
        If Trim(strAddressEmail) <> Trim(strTestVal) Then booFoundDuplicateInStep = False
      End If
    End If

    If booFoundDuplicateInStep Then
      pXMLPropSet.GetAttribute strXPath & "[" & CStr(lngIndex) & "]/role/RoleCd", "value", varVals
      If IsEmpty(varVals) Then
        booFoundDuplicateInStep = False
      Else
        strTestVal = CStr(varVals(0))
        If Trim(ReturnRoleCDString(enumJenRole)) <> Trim(strTestVal) Then booFoundDuplicateInStep = False
      End If
    End If

    If booFoundDuplicateInStep Then
      pXMLPropSet.GetAttribute strXPath & "[" & CStr(lngIndex) & "]/rpCntInfo/cntAddress", "addressType", varVals
      If IsEmpty(varVals) Then
        If enumJenAddressType <> JenMetadata_Skip Then booFoundDuplicateInStep = False
      Else
        strTestVal = CStr(varVals(0))
        If Trim(ReturnAddressType(enumJenAddressType)) <> Trim(strTestVal) Then booFoundDuplicateInStep = False
      End If
    End If

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

Public Function SetMetadataKeyWords(pDataset As IDataset, _
  Optional pIncludeThemeKeys As esriSystem.IStringArray, _
  Optional pIncludeSearchKeys As esriSystem.IStringArray, _
  Optional pIncludeDescKeys As esriSystem.IStringArray, _
  Optional pIncludeStratKeys As esriSystem.IStringArray, _
  Optional pIncludeThemeSlashThemekeys As esriSystem.IStringArray, _
  Optional pIncludePlaceKeys As esriSystem.IStringArray, _
  Optional pIncludeTemporalKeys As esriSystem.IStringArray) As String   ', _

  On Error GoTo ErrHandler

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

  If Not pIncludeStratKeys Is Nothing Then
    varAtts = pPropSet.GetProperty(strStratKeywordsXPath)
    If Not IsEmpty(varAtts) Then
      For lngIndex = 0 To UBound(varAtts)
        strValue = CStr(varAtts(lngIndex))

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

  On Error GoTo ErrHandler

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

  On Error GoTo ErrHandler

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

  booFailed = False

  Dim pName As IName
  Set pName = pDataset.FullName

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

  Dim pGxDataset As IGxDataset
  Set pGxDataset = ReturnGxDatasetFromDataset(pDataset)

  Dim pMetaData As IMetadata
  Set pMetaData = pGxDataset

  Dim pMetadataEdit As IMetadataEdit
  Set pMetadataEdit = pGxDataset

  Dim pPropSet As IPropertySet
  Set pPropSet = pMetaData.Metadata

  Dim pXMLPropSet As IXmlPropertySet2
  Set pXMLPropSet = pMetaData.Metadata
  Set pPropSet = pXMLPropSet

  Dim pXMLPropSet2 As IXmlPropertySet2
  If pMetadataEdit.CanEditMetadata Then
    If pXMLPropSet.IsNew Then
      Set pXMLPropSet2 = pXMLPropSet
      pXMLPropSet2.InitExisting
      pMetaData.Metadata = pPropSet
    End If
  End If

  Set ReturnMetadataPropSetFromDataset = pPropSet

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

  SynchronizeMetadataPropSet = "Succeeded"

  Dim pGxDataset As IGxDataset
  Set pGxDataset = ReturnGxDatasetFromDataset(pDataset)

  Dim pMetaData As IMetadata
  Set pMetaData = pGxDataset

  Dim pMetadataEdit As IMetadataEdit
  Set pMetadataEdit = pGxDataset

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

  booFailed = False
  Dim pGxDataset As IGxDataset
  Set pGxDataset = ReturnGxDatasetFromDataset(pDataset)

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

Public Function AddMetadataUseLimitations(pDataset As IDataset, strUseLimitations As String) As String ', _

  On Error GoTo ErrHandler

  AddMetadataUseLimitations = "Succeeded"

  strUseLimitations = "<DIV STYLE=""text-align:Left;""><DIV><P><SPAN>" & strUseLimitations & "</SPAN></P></DIV></DIV>"

  Dim strUseLimitationsXPath As String
  strUseLimitationsXPath = "dataIdInfo/resConst/Consts/useLimit"

  Dim pPropSet As IPropertySet
  Set pPropSet = ReturnMetadataPropSetFromDataset(pDataset)

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

  On Error GoTo ErrHandler

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



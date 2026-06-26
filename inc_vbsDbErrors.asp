<%
'***********************************************
sub drawErrorBoxBegin()
'***********************************************
%>
  <br>
    <center>
      <table border=4>
        <tr bgcolor=#cea6c6>
          <td align=center>
            <font face="Arial, Helvetica, sans-serif" size=3>
              <b>
<%
end sub

'***********************************************
sub drawErrorBoxEnd()
'***********************************************
%>
              </b>
            </font>
      </table>
    </center>
<%
end sub

'***********************************************
sub drawErrorBox( strError )
'***********************************************
  drawErrorBoxBegin
%>
    VBSdb Error: <%=strError%>
<%
  drawErrorBoxEnd
end sub

'***********************************************
sub drawDebugBoxBodyIntroduction( objVbsDb )
'***********************************************
  response.write "Debugging values:<br>"
end sub

'***********************************************
sub drawDebugLine( strPropertyName , objVbsDb )
'***********************************************
  response.write strPropertyName & ": " & server.htmlEncode( cStr( objVbsDb( strPropertyName ) ) ) & "<br>"
end sub

'***********************************************
sub vbsDbDrawScriptEngineVersion()
'***********************************************
  response.write "Script Engine version: " & _
    server.htmlEncode( ScriptEngineMajorVersion & "." & ScriptEngineMinorVersion & "." & ScriptEngineBuildVersion ) & "<br>"
end sub

'***********************************************
sub vbsDbDrawAdoVersion( objVbsDb )
'***********************************************
  response.write "ADO version: " & server.htmlEncode( cStr( objVbsDb( "Connection" ).version ) ) & "<br>"
end sub

'***********************************************
sub vbsDbDrawEnvironmentVersions( objVbsDb )
'***********************************************
  vbsDbDrawScriptEngineVersion
  vbsDbDrawAdoVersion objVbsDb
end sub

'***********************************************
sub vbsDbDrawDebugForProperties( objVbsDb )
'***********************************************
  drawDebugLine "Dsn" , objVbsDb
  drawDebugLine "MdbPath" , objVbsDb
  drawDebugLine "Sql" , objVbsDb
  drawDebugLine "SqlWithFilterWhereClause" , objVbsDb
  drawDebugLine "ViewMode" , objVbsDb
  drawDebugLine "ViewNavigationButtons" , objVbsDb
  drawDebugLine "ViewNavigationPosition" , objVbsDb
  drawDebugLine "ViewNavigationBGColor" , objVbsDb
  drawDebugLine "ViewNavigationDisabledFGColor" , objVbsDb
  drawDebugLine "ViewNavigationFGColor" , objVbsDb
  drawDebugLine "ViewPosition" , objVbsDb
  drawDebugLine "ViewFieldFormats" , objVbsDb
  drawDebugLine "GridPageSize" , objVbsDb
  drawDebugLine "GridFields" , objVbsDb
  drawDebugLine "GridHideFields" , objVbsDb
  drawDebugLine "GridShowIndex" , objVbsDb
  drawDebugLine "GridUnselectedIndexFGColor" , objVbsDb
  drawDebugLine "GridSelectedIndexBGColor" , objVbsDb
  drawDebugLine "GridSelectedIndexFGColor" , objVbsDb
  drawDebugLine "GridTableTag" , objVbsDb
  drawDebugLine "GridHorizontalStripeBGColor" , objVbsDb
  drawDebugLine "GridUpdateButtons" , objVbsDb
  drawDebugLine "GridDeleteButtons" , objVbsDb
  drawDebugLine "GridTemplate" , objVbsDb
  drawDebugLine "FormExactAfterEdit" , objVbsDb
  drawDebugLine "FormFields" , objVbsDb
  drawDebugLine "FormHideFields" , objVbsDb
  drawDebugLine "FormTableTag" , objVbsDb
  drawDebugLine "FormTemplate" , objVbsDb
  drawDebugLine "InputSelectFields" , objVbsDb
  drawDebugLine "InputValidateDateFields" , objVbsDb
  drawDebugLine "InputPasswordFields" , objVbsDb
  drawDebugLine "SearchFields" , objVbsDb
  drawDebugLine "SearchHideFields" , objVbsDb
  drawDebugLine "SearchAliasFields" , objVbsDb
  drawDebugLine "SearchOperators" , objVbsDb
  drawDebugLine "SearchScreenRepeat" , objVbsDb
  drawDebugLine "SearchTemplate" , objVbsDb
  drawDebugLine "EditMemoFields" , objVbsDb
  drawDebugLine "EditTableName" , objVbsDb
  drawDebugLine "EditKeyFields" , objVbsDb
  drawDebugLine "EditFields" , objVbsDb
  drawDebugLine "EditAutoincrement" , objVbsDb
  drawDebugLine "EditReadOnlyFields" , objVbsDb
  drawDebugLine "EditHideFields" , objVbsDb
  drawDebugLine "EditAddFieldDefaults" , objVbsDb
  drawDebugLine "EditAddScreenRepeat" , objVbsDb
  drawDebugLine "EditAddSingle" , objVbsDb
  drawDebugLine "EditAddTemplate" , objVbsDb
  drawDebugLine "EditDeleteTemplate" , objVbsDb
  drawDebugLine "EditAddFieldDefaults" , objVbsDb
  drawDebugLine "EditUpdateTemplate" , objVbsDb
  drawDebugLine "EditValidateRequiredFields" , objVbsDb
  drawDebugLine "EditValidateRegExp" , objVbsDb
  drawDebugLine "GlobalReset" , objVbsDb
  drawDebugLine "GlobalBooleanFields" , objVbsDb
  drawDebugLine "GlobalFieldHeaders" , objVbsDb
  drawDebugLine "GlobalTableBGColor" , objVbsDb
  drawDebugLine "GlobalTableFGColor" , objVbsDb
  drawDebugLine "GlobalHeaderBGColor" , objVbsDb
  drawDebugLine "GlobalHeaderFGColor" , objVbsDb
  drawDebugLine "GlobalImageDir" , objVbsDb
  drawDebugLine "GlobalId" , objVbsDb
  drawDebugLine "GlobalDbType" , objVbsDb
  drawDebugLine "GlobalLanguage" , objVbsDb
  drawDebugLine "GlobalCustomText" , objVbsDb
  drawDebugLine "GlobalCSS" , objVbsDb
  drawDebugLine "GlobalStartScreen" , objVbsDb
  drawDebugLine "GlobalFalseText" , objVbsDb
  drawDebugLine "GlobalTrueText" , objVbsDb
  drawDebugLine "GlobalDateFormatValidDateSeparators" , objVbsDb
  drawDebugLine "GlobalDateFormatValidTimeSeparators" , objVbsDb
  drawDebugLine "GlobalDateFormat" , objVbsDb
  drawDebugLine "GlobalSeparators" , objVbsDb
  drawDebugLine "GlobalOleDbProvider" , objVbsDb
  drawDebugLine "GlobalPagePositioning" , objVbsDb
  drawDebugLine "GlobalQuerystringPreserve" , objVbsDb
  drawDebugLine "FormAbsolutePosition" , objVbsDb
  drawDebugLine "OrderBy" , objVbsDb
  drawDebugLine "FilterWhereClause" , objVbsDb
end sub

'***********************************************
sub vbsDbDrawVBSdbVersion( objVbsDb )
'***********************************************
  response.write "VBSdb Version: " & constVBSdbVersion & "<br>"
end sub

'***********************************************
sub vbsDbDrawInterestingMethodValues( objVbsDb )
'***********************************************
  response.write "VbsDbGetScreenType: " & server.htmlEncode( VbsDbGetScreenType( objVbsDb ) ) & "<br>"
  response.write "VbsDbGetLastAction: " & server.htmlEncode( VbsDbGetLastAction( objVbsDb ) ) & "<br>"
end sub

'***********************************************
sub drawDebugBoxBodyLines( objVbsDb )
'***********************************************
  vbsDbDrawEnvironmentVersions objVbsDb
  vbsDbDrawVBSdbVersion objVbsDb
  vbsDbDrawDebugForProperties objVbsDb
  vbsDbDrawInterestingMethodValues objVbsDb
end sub

'***********************************************
sub drawDebugBoxBody( objVbsDb )
'***********************************************
  drawDebugBoxBodyIntroduction objVbsDb
  drawDebugBoxBodyLines objVbsDb
end sub

'***********************************************
sub drawDebugBox( objVbsDb )
'***********************************************
  drawErrorBoxBegin
  drawDebugBoxBody objVbsDb
  drawErrorBoxEnd
end sub

'***********************************************
'***********************************************
sub drawErrorWithDebugging( strError , objVbsDb )
'***********************************************
'***********************************************
  drawErrorBox strError
  drawDebugBox objVbsDb
  VbsDbClose objVbsDb
  response.end
end sub

'***********************************************
'***********************************************
sub drawError( strError , objVbsDb )
'***********************************************
'***********************************************
  drawErrorBox strError
  VbsDbClose objVbsDb
  response.end
end sub

'***********************************************
'***********************************************
function funBolIsAPropertyName( strItemName , objVbsDb )
'***********************************************
'***********************************************
  dim strUCaseItemName
  strUCaseItemName = uCase( strItemName )
  funBolIsAPropertyName = _
    ( strUCaseItemName = "DSN" ) or _
    ( strUCaseItemName = "MDBPATH" ) or _
    ( strUCaseItemName = "SQL" ) or _
    ( strUCaseItemName = "VIEWMODE" ) or _
    ( strUCaseItemName = "VIEWNAVIGATIONBUTTONS" ) or _
    ( strUCaseItemName = "VIEWNAVIGATIONPOSITION" ) or _
    ( strUCaseItemName = "VIEWNAVIGATIONBGCOLOR" ) or _
    ( strUCaseItemName = "VIEWNAVIGATIONDISABLEDFGCOLOR" ) or _
    ( strUCaseItemName = "VIEWNAVIGATIONFGCOLOR" ) or _
    ( strUCaseItemName = "VIEWPOSITION" ) or _
    ( strUCaseItemName = "VIEWFIELDFORMATS" ) or _
    ( strUCaseItemName = "GRIDPAGESIZE" ) or _
    ( strUCaseItemName = "GRIDFIELDS" ) or _
    ( strUCaseItemName = "GRIDHIDEFIELDS" ) or _
    ( strUCaseItemName = "GRIDSHOWINDEX" ) or _
    ( strUCaseItemName = "GRIDUNSELECTEDINDEXFGCOLOR" ) or _
    ( strUCaseItemName = "GRIDSELECTEDINDEXBGCOLOR" ) or _
    ( strUCaseItemName = "GRIDSELECTEDINDEXFGCOLOR" ) or _
    ( strUCaseItemName = "GRIDTABLETAG" ) or _
    ( strUCaseItemName = "GRIDHORIZONTALSTRIPEBGCOLOR" ) or _
    ( strUCaseItemName = "GRIDUPDATEBUTTONS" ) or _
    ( strUCaseItemName = "GRIDDELETEBUTTONS" ) or _
    ( strUCaseItemName = "GRIDTEMPLATE" ) or _
    ( strUCaseItemName = "FORMEXACTAFTEREDIT" ) or _
    ( strUCaseItemName = "FORMFIELDS" ) or _
    ( strUCaseItemName = "FORMHIDEFIELDS" ) or _
    ( strUCaseItemName = "FORMTABLETAG" ) or _
    ( strUCaseItemName = "FORMTEMPLATE" ) or _
    ( strUCaseItemName = "INPUTSELECTFIELDS" ) or _
    ( strUCaseItemName = "INPUTENUMERATEDFIELDS" ) or _
    ( strUCaseItemName = "INPUTVALIDATEDATEFIELDS" ) or _
    ( strUCaseItemName = "INPUTPASSWORDFIELDS" ) or _
    ( strUCaseItemName = "SEARCHFIELDS" ) or _
    ( strUCaseItemName = "SEARCHHIDEFIELDS" ) or _
    ( strUCaseItemName = "SEARCHALIASFIELDS" ) or _
    ( strUCaseItemName = "SEARCHOPERATORS" ) or _
    ( strUCaseItemName = "SEARCHSCREENREPEAT" ) or _
    ( strUCaseItemName = "SEARCHTEMPLATE" ) or _
    ( strUCaseItemName = "EDITMEMOFIELDS" ) or _
    ( strUCaseItemName = "EDITTABLENAME" ) or _
    ( strUCaseItemName = "EDITKEYFIELDS" ) or _
    ( strUCaseItemName = "EDITFIELDS" ) or _
    ( strUCaseItemName = "EDITAUTOINCREMENT" ) or _
    ( strUCaseItemName = "EDITREADONLYFIELDS" ) or _
    ( strUCaseItemName = "EDITHIDEFIELDS" ) or _
    ( strUCaseItemName = "EDITADDFIELDDEFAULTS" ) or _
    ( strUCaseItemName = "EDITADDSCREENREPEAT" ) or _
    ( strUCaseItemName = "EDITADDSINGLE" ) or _
    ( strUCaseItemName = "EDITADDTEMPLATE" ) or _
    ( strUCaseItemName = "EDITDELETETEMPLATE" ) or _
    ( strUCaseItemName = "EDITUPDATEFIELDDEFAULTS" ) or _
    ( strUCaseItemName = "EDITUPDATETEMPLATE" ) or _
    ( strUCaseItemName = "EDITVALIDATEREQUIREDFIELDS" ) or _
    ( strUCaseItemName = "EDITVALIDATEREGEXP" ) or _
    ( strUCaseItemName = "EDITERRORMESSAGES" ) or _
    ( strUCaseItemName = "EDITFORMTAMPERINGCONTROL" ) or _
    ( strUCaseItemName = "GLOBALRESET" ) or _
    ( strUCaseItemName = "GLOBALBOOLEANFIELDS" ) or _
    ( strUCaseItemName = "GLOBALFIELDHEADERS" ) or _
    ( strUCaseItemName = "GLOBALTABLEBGCOLOR" ) or _
    ( strUCaseItemName = "GLOBALTABLEFGCOLOR" ) or _
    ( strUCaseItemName = "GLOBALHEADERBGCOLOR" ) or _
    ( strUCaseItemName = "GLOBALHEADERFGCOLOR" ) or _
    ( strUCaseItemName = "GLOBALIMAGEDIR" ) or _
    ( strUCaseItemName = "GLOBALID" ) or _
    ( strUCaseItemName = "GLOBALDBTYPE" ) or _
    ( strUCaseItemName = "GLOBALLANGUAGE" ) or _
    ( strUCaseItemName = "GLOBALCUSTOMTEXT" ) or _
    ( strUCaseItemName = "GLOBALCSS" ) or _
    ( strUCaseItemName = "GLOBALSTARTSCREEN" ) or _
    ( strUCaseItemName = "GLOBALFALSETEXT" ) or _
    ( strUCaseItemName = "GLOBALTRUETEXT" ) or _
    ( strUCaseItemName = "GLOBALSEPARATORS" ) or _
    ( strUCaseItemName = "GLOBALOLEDBPROVIDER" ) or _
    ( strUCaseItemName = "GLOBALPAGEPOSITIONING" ) or _
    ( strUCaseItemName = "GLOBALQUERYSTRINGPRESERVE" ) or _
    ( strUCaseItemName = "GLOBALDATEFORMATVALIDDATESEPARATORS" ) or _
    ( strUCaseItemName = "GLOBALDATEFORMATVALIDTIMESEPARATORS" ) or _
    ( strUCaseItemName = "GLOBALDATEFORMAT" )
end function

'***********************************************
sub vbsDbCheckIfStrItemNameIsAPropertyName( strItemName , objVbsDb )
'***********************************************
  if not funBolIsAPropertyName( strItemName , objVbsDb ) then
    drawError strItemName & " is not a supported property. Check typo mistakes.<br>" &_
      "Attention: property names are not case sensitive." , objVbsDb
  end if
end sub

'***********************************************
sub vbsDbCheckErrorUnsupportedProperty( objVbsDb )
'***********************************************
  dim strItemName
  for each strItemName in objVbsDb
    vbsDbCheckIfStrItemNameIsAPropertyName strItemName , objVbsDb
  next
end sub

'***********************************************
function funBolIsProVersion()
'***********************************************
  funBolIsProVersion = ( inStr( uCase( constVBSdbVersion ) , "PRO" ) > 0 )
end function

'***********************************************
sub vbsDbCheckErrorProOnlyPropertyDrawError( strPropertyName , objVbsDb )
'***********************************************
  drawError strPropertyName & " property is not supported by VBSdb Free. It requires VBSdb Professional." , objVbsDb
end sub

'***********************************************
sub vbsDbCheckErrorProOnlyPropertyGridTemplate( objVbsDb )
'***********************************************
  if ( not funBolIsProVersion() ) and ( objVbsDb( "GridTemplate" ) <> "" ) then
    vbsDbCheckErrorProOnlyPropertyDrawError "GridTemplate" , objVbsDb
  end if
end sub

'***********************************************
sub vbsDbCheckErrorProOnlyPropertyFormTemplate( objVbsDb )
'***********************************************
  if ( not funBolIsProVersion() ) and ( objVbsDb( "FormTemplate" ) <> "" ) then
    vbsDbCheckErrorProOnlyPropertyDrawError "FormTemplate" , objVbsDb
  end if
end sub

'***********************************************
sub vbsDbCheckErrorProOnlyPropertySearchTemplate( objVbsDb )
'***********************************************
  if ( not funBolIsProVersion() ) and ( objVbsDb( "SearchTemplate" ) <> "" ) then
    vbsDbCheckErrorProOnlyPropertyDrawError "SearchTemplate" , objVbsDb
  end if
end sub

'***********************************************
sub vbsDbCheckErrorProOnlyPropertyEditAddTemplate( objVbsDb )
'***********************************************
  if ( not funBolIsProVersion() ) and ( objVbsDb( "EditAddTemplate" ) <> "" ) then
    vbsDbCheckErrorProOnlyPropertyDrawError "EditAddTemplate" , objVbsDb
  end if
end sub

'***********************************************
sub vbsDbCheckErrorProOnlyPropertyEditUpdateTemplate( objVbsDb )
'***********************************************
  if ( not funBolIsProVersion() ) and ( objVbsDb( "EditUpdateTemplate" ) <> "" ) then
    vbsDbCheckErrorProOnlyPropertyDrawError "EditUpdateTemplate" , objVbsDb
  end if
end sub

'***********************************************
sub vbsDbCheckErrorProOnlyPropertyEditDeleteTemplate( objVbsDb )
'***********************************************
  if ( not funBolIsProVersion() ) and ( objVbsDb( "EditDeleteTemplate" ) <> "" ) then
    vbsDbCheckErrorProOnlyPropertyDrawError "EditDeleteTemplate" , objVbsDb
  end if
end sub

'***********************************************
sub vbsDbCheckErrorProOnlyPropertyInputValidateDateFields( objVbsDb )
'***********************************************
  if ( not funBolIsProVersion() ) and ( objVbsDb( "InputValidateDateFields" ) <> "" ) then
    vbsDbCheckErrorProOnlyPropertyDrawError "InputValidateDateFields" , objVbsDb
  end if
end sub

'***********************************************
sub vbsDbCheckErrorProOnlyPropertyInputEnumeratedFields( objVbsDb )
'***********************************************
  if ( not funBolIsProVersion() ) and ( objVbsDb( "InputEnumeratedFields" ) <> "" ) then
    vbsDbCheckErrorProOnlyPropertyDrawError "InputEnumeratedFields" , objVbsDb
  end if
end sub

'***********************************************
sub vbsDbCheckErrorProOnlyPropertyEditValidateRegExp( objVbsDb )
'***********************************************
  if ( not funBolIsProVersion() ) and ( objVbsDb( "EditValidateRegExp" ) <> "" ) then
    vbsDbCheckErrorProOnlyPropertyDrawError "EditValidateRegExp" , objVbsDb
  end if
end sub

'***********************************************
sub vbsDbCheckErrorProOnlyProperties( objVbsDb )
'***********************************************
  vbsDbCheckErrorProOnlyPropertyGridTemplate objVbsDb
  vbsDbCheckErrorProOnlyPropertyFormTemplate objVbsDb
  vbsDbCheckErrorProOnlyPropertySearchTemplate objVbsDb
  vbsDbCheckErrorProOnlyPropertyEditAddTemplate objVbsDb
  vbsDbCheckErrorProOnlyPropertyEditUpdateTemplate objVbsDb
  vbsDbCheckErrorProOnlyPropertyEditDeleteTemplate objVbsDb
  vbsDbCheckErrorProOnlyPropertyInputValidateDateFields objVbsDb
  vbsDbCheckErrorProOnlyPropertyInputEnumeratedFields objVbsDb
  vbsDbCheckErrorProOnlyPropertyEditValidateRegExp objVbsDb
end sub

'***********************************************
sub vbsDbCheckMandatoryDsnOrMdbPath( objVbsDb )
'***********************************************
  if ( objVbsDb( "Dsn" ) = "" ) and ( objVbsDb( "MdbPath" ) = "" ) then
    ' user didn't define any connection information
    drawError "Neither Dsn nor MdbPath property was defined. One of the two has to be " & _
      "set, to provide VBSdb with the proper connection information." , objVbsDb
  end if
end sub

'***********************************************
sub vbsDbCheckMandatorySql( objVbsDb )
'***********************************************
  if objVbsDb( "Sql" ) = "" then
    ' user didn't define the sql string
    drawError "Sql was not defined. Sql is mandatory." , objVbsDb
  end if
end sub

'***********************************************
sub vbsDbCheckMandatoryProperties( objVbsDb )
'***********************************************
  vbsDbCheckMandatoryDsnOrMdbPath objVbsDb
  vbsDbCheckMandatorySql objVbsDb
end sub

'***********************************************
sub vbsDbCheckErrorPropertyValueTypeIsString( strPropertyName , objVbsDb )
'***********************************************
  if varType( objVbsDb( strPropertyName ) ) <> 8 then
    ' strPropertyName is a string property, but its value is not a string
    drawError strPropertyName & " wrong type. " & strPropertyName & " is a string valued property. Please, " & _
      "control " & strPropertyName & " assignment." , objVbsDb
  end if
end sub

'***********************************************
sub vbsDbCheckErrorPropertyValueTypeIsBoolean( strPropertyName , objVbsDb )
'***********************************************
  if varType( objVbsDb( strPropertyName ) ) <> 11 then
    ' strPropertyName is a boolean property, but its value is not a boolean
    drawError strPropertyName & " wrong type. " & strPropertyName & " is a boolan property. Please, " & _
      "control " & strPropertyName & " assignment." , objVbsDb
  end if
end sub

'***********************************************
sub vbsDbCheckGlobalSeparators( objVbsDb )
'***********************************************
  vbsDbCheckErrorPropertyValueTypeIsString "GlobalSeparators" , objVbsDb
end sub

'***********************************************
sub vbsDbCheckFormExactAfterEdit( objVbsDb )
'***********************************************
  vbsDbCheckErrorPropertyValueTypeIsBoolean "FormExactAfterEdit" , objVbsDb
end sub

'***********************************************
sub vbsDbCheckViewPosition( objVbsDb )
'***********************************************
  vbsDbCheckErrorPropertyValueTypeIsBoolean "ViewPosition" , objVbsDb
end sub

'***********************************************
sub vbsDbCheckEditAddScreenRepeat( objVbsDb )
'***********************************************
  vbsDbCheckErrorPropertyValueTypeIsBoolean "EditAddScreenRepeat" , objVbsDb
end sub

'***********************************************
sub vbsDbCheckGridDeleteButtons( objVbsDb )
'***********************************************
  vbsDbCheckErrorPropertyValueTypeIsBoolean "GridDeleteButtons" , objVbsDb
  if objVbsDb( "GridDeleteButtons" ) and ( objVbsDb( "GlobalImageDir" ) = "" ) then
    drawError "GridDeleteButtons property. If you want to use GridDeleteButtons, you have to set the GlobalImageDir property too." , objVbsDb
  end if
  if objVbsDb( "GridDeleteButtons" ) and ( objVbsDb( "EditTableName" ) = "" ) then
    drawError "GridDeleteButtons property. If you want to use GridDeleteButtons, you have to set the EditTableName property too." , objVbsDb
  end if
  if objVbsDb( "GridDeleteButtons" ) and ( objVbsDb( "EditKeyFields" ) = "" ) then
    drawError "GridDeleteButtons property. If you want to use GridDeleteButtons, you have to set the EditKeyFields property too." , objVbsDb
  end if
end sub

'***********************************************
sub vbsDbCheckGridUpdateButtons( objVbsDb )
'***********************************************
  vbsDbCheckErrorPropertyValueTypeIsBoolean "GridUpdateButtons" , objVbsDb
  if objVbsDb( "GridUpdateButtons" ) and ( objVbsDb( "GlobalImageDir" ) = "" ) then
    drawError "GridUpdateButtons property. If you want to use GridUpdateButtons, you have to set the GlobalImageDir property too." , objVbsDb
  end if
  if objVbsDb( "GridUpdateButtons" ) and ( objVbsDb( "EditTableName" ) = "" ) then
    drawError "GridUpdateButtons property. If you want to use GridUpdateButtons, you have to set the EditTableName property too." , objVbsDb
  end if
  if objVbsDb( "GridUpdateButtons" ) and ( objVbsDb( "EditKeyFields" ) = "" ) then
    drawError "GridUpdateButtons property. If you want to use GridUpdateButtons, you have to set the EditKeyFields property too." , objVbsDb
  end if
end sub

'***********************************************
sub vbsDbCheckViewNavigationButtons( objVbsDb )
'***********************************************
  if ( funBolInDbNavigationItem( "ADD" , objVbsDb )    or _
       funBolInDbNavigationItem( "UPDATE" , objVbsDb ) or _
       funBolInDbNavigationItem( "DELETE" , objVbsDb ) ) and _
     ( objVbsDb( "EditTableName" ) = "" ) then
    ' ViewNavigationButtons property enables database editing, but no edit table has been defined
    drawError "ViewNavigationButtons property cannot contain 'Add', 'Update' or 'Delete' if the EditTableName property is not set." , objVbsDb
  end if
  if ( funBolInDbNavigationItem( "UPDATE" , objVbsDb ) or _
       funBolInDbNavigationItem( "DELETE" , objVbsDb ) ) and _
     ( objVbsDb( "EditKeyFields" ) = "" ) then
    ' ViewNavigationButtons property enables existing record editing, but no edit key fields have been defined
    drawError "ViewNavigationButtons property cannot contain 'Update' or 'Delete' if the EditKeyFields property is not set." , objVbsDb
  end if
end sub

'***********************************************
sub vbsDbCheckEditAddSingle( objVbsDb )
'***********************************************
  vbsDbCheckErrorPropertyValueTypeIsBoolean "EditAddSingle" , objVbsDb
  if ( objVbsDb( "EditAddSingle" ) ) and ( trim( objVbsDb( "EditTableName" ) = "" ) ) then
    drawError "EditAddSingle property. If you want to set EditAddSingle, you have to set the EditTableName property too." , objVbsDb
  end if
end sub

'***********************************************
sub vbsDbCheckEditKeyFields( objVbsDb )
'***********************************************
  if ( objVbsDb( "EditKeyFields" ) <> "" ) and _
      ( not funBolUpdateIsGranted( objVbsDb ) ) and ( not funBolDeleteIsGranted( objVbsDb ) ) then
    drawError "EditKeyFields property. EditKeyFields is set, but VBSdb neither has been configured to " & _
      "update nor to delete records. EditKeyFields is meaningful only if at least one of the following is true:" & _
      "<ul>" & _
      "<li>the ViewNavigationButtons property contains the Update button" & _
      "<li>the ViewNavigationButtons property contains the Delete button" & _
      "<li>the GridUpdateButtons property is set to true" & _
      "<li>the GridDeleteButtons property is set to true" & _
      "<ul>" , objVbsDb
  end if
  if ( objVbsDb( "EditKeyFields" ) <> "" ) and ( trim( objVbsDb( "EditTableName" ) = "" ) ) then
    drawError "EditKeyFields property. If you want to use EditKeyFields, you have to set the EditTableName property too." , objVbsDb
  end if
end sub

'***********************************************
sub vbsDbCheckEditTableName( objVbsDb )
'***********************************************
  if ( objVbsDb( "EditTableName" ) <> "" ) and _
      ( not funBolAddIsGranted( objVbsDb ) ) and _
      ( not funBolUpdateIsGranted( objVbsDb ) ) and _
      ( not funBolDeleteIsGranted( objVbsDb ) ) then
    drawError "EditTableName property. EditTableName is set, but VBSdb neither has been configured to " & _
      "add nor to update nor to delete records. EditTableName is meaningful only if at least one of the following is true:" & _
      "<ul>" & _
      "<li>the ViewNavigationButtons property contains the Add button" & _
      "<li>the ViewNavigationButtons property contains the Update button" & _
      "<li>the ViewNavigationButtons property contains the Delete button" & _
      "<li>the GridUpdateButtons property is set to true" & _
      "<li>the GridDeleteButtons property is set to true" & _
      "<li>the EditAddSingle property is set to true" & _
      "<ul>" , objVbsDb
  end if
end sub

'***********************************************
sub vbsDbCheckViewMode( objVbsDb )
'***********************************************
  if objVbsDb( "ViewMode" ) <> "GRID" and objVbsDb( "ViewMode" ) <> "FORM" and objVbsDb( "ViewMode" ) <> "GRID-FORM" and objVbsDb( "ViewMode" ) <> "GRID|FORM" then
    ' the developer inserted a bad VieMode value
    drawError "ViewMode property, bad assignment. Valid values are ""GRID"",""FORM"",""GRID-FORM"" and ""GRID|FORM""" , objVbsDb
  end if
end sub

'***********************************************
sub vbsDbCheckSearchScreenRepeat( objVbsDb )
'***********************************************
  vbsDbCheckErrorPropertyValueTypeIsBoolean "SearchScreenRepeat" , objVbsDb
end sub

'***********************************************
sub vbsDbCheckGlobalReset( objVbsDb )
'***********************************************
  vbsDbCheckErrorPropertyValueTypeIsBoolean "GlobalReset" , objVbsDb
end sub

'***********************************************
sub vbsDbCheckGlobalStartScreen( objVbsDb )
'***********************************************
  if ( uCase( objVbsDb( "GlobalStartScreen" ) ) = "ADD" ) and ( objVbsDb( "EditTableName" ) = "" ) then
    drawError "when GlobalStartScreen is set to 'Add', EditTableName must be defined." , objVbsDb
  end if
end sub

'***********************************************
sub vbsDbCheckColorAndCSS( strPropertyName , strDefaultValue , objVbsDb )
'***********************************************
  if ( uCase( objVbsDb( strPropertyName ) ) <> uCase( strDefaultValue ) ) and _
    objVbsDb( "GlobalCSS" ) <> "" then
    ' a color related property has been assigned, but the GlobalCSS property have been set
    drawError strPropertyName & " has been assigned, but GlobalCSS has been assigned too.<br><br>"  & _
      "When the GlobalCSS property is assigned, the color related properties are not considered " & _
      "at all. To customize colors, modify the pointed CSS style sheet file.<br><br>Please, " & _
      "delete the " & strPropertyName & " assignment (or delete the GlobalCSS assignment)." , objVbsDb
  end if
end sub

'***********************************************
sub vbsDbCheckColorProperty( strColorPropertyName , strDefaultValue , objVbsDb )
'***********************************************
  vbsDbCheckErrorPropertyValueTypeIsString strColorPropertyName , objVbsDb
  vbsDbCheckColorAndCSS strColorPropertyName , strDefaultValue , objVbsDb
end sub

'***********************************************
sub vbsDbCheckColorProperties( objVbsDb )
'***********************************************
  vbsDbCheckColorProperty "ViewNavigationBGColor" , constDefaultViewNavigationBGColor , objVbsDb
  vbsDbCheckColorProperty "ViewNavigationDisabledFGColor" , constDefaultViewNavigationDisabledFGColor , objVbsDb
  vbsDbCheckColorProperty "ViewNavigationFGColor" , constDefaultViewNavigationFGColor , objVbsDb
  vbsDbCheckColorProperty "GridHorizontalStripeBGColor" , constDefaultGridHorizontalStripeBGColor , objVbsDb
  vbsDbCheckColorProperty "GridUnselectedIndexFGColor" , constDefaultGridUnselectedIndexFGColor , objVbsDb
  vbsDbCheckColorProperty "GridSelectedIndexBGColor" , constDefaultGridSelectedIndexBGColor , objVbsDb
  vbsDbCheckColorProperty "GridSelectedIndexFGColor" , constDefaultGridSelectedIndexFGColor , objVbsDb
  vbsDbCheckColorProperty "GlobalTableBGColor" , constDefaultGlobalTableBGColor , objVbsDb
  vbsDbCheckColorProperty "GlobalTableFGColor" , constDefaultGlobalTableFGColor , objVbsDb
  vbsDbCheckColorProperty "GlobalHeaderBGColor" , constDefaultGlobalHeaderBGColor , objVbsDb
  vbsDbCheckColorProperty "GlobalHeaderFGColor" , constDefaultGlobalHeaderFGColor , objVbsDb
end sub

'***********************************************
sub vbsDbCheckErrorWrongPropertyAssignment( objVbsDb )
'***********************************************
  vbsDbCheckGlobalSeparators objVbsDb
  vbsDbCheckFormExactAfterEdit objVbsDb
  vbsDbCheckViewPosition objVbsDb
  vbsDbCheckEditAddScreenRepeat objVbsDb
  vbsDbCheckGridUpdateButtons objVbsDb
  vbsDbCheckGridDeleteButtons objVbsDb
  vbsDbCheckViewNavigationButtons objVbsDb
  vbsDbCheckEditAddSingle objVbsDb
  vbsDbCheckEditKeyFields objVbsDb
  vbsDbCheckEditTableName objVbsDb
  vbsDbCheckViewMode objVbsDb
  vbsDbCheckSearchScreenRepeat objVbsDb
  vbsDbCheckGlobalReset objVbsDb
  vbsDbCheckGlobalStartScreen objVbsDb
  vbsDbCheckColorProperties objVbsDb
end sub

'***********************************************
'***********************************************
sub vbsDbCheckErrorPropertyAssignment( objVbsDb )
'***********************************************
'***********************************************
  vbsDbCheckErrorUnsupportedProperty objVbsDb
  vbsDbCheckErrorProOnlyProperties objVbsDb
  vbsDbCheckMandatoryProperties objVbsDb
  vbsDbCheckErrorWrongPropertyAssignment objVbsDb
  'vbsDbCheckErrorBadPropertyValueTypes objVbsDb
  'vbsDbCheckSpecificProperties objVbsDb
end sub

'***********************************************
function funStrFieldNameToDisplay( strFieldName , strAssociatedDictionary , objVbsDb )
'***********************************************
  if varType( objVbsDb( strAssociatedDictionary )( strFieldName ) ) = vbString then
    if uCase( strFieldName ) = uCase( objVbsDb( strAssociatedDictionary )( strFieldName ) ) then
      ' there is no parameter, for this property, for this field
      funStrFieldNameToDisplay = objVbsDb( strAssociatedDictionary )( strFieldName )
    else
      ' for this field, in this property, there is some parameter
      funStrFieldNameToDisplay = uCase( strFieldName )
    end if
  else
    ' for this field, in this property, there is some parameter
    funStrFieldNameToDisplay = uCase( strFieldName )
  end if
end function

'***********************************************
function funStrErrorBadFieldForSqlRecordSet( strPropertyName , strFieldNameToDisplay , strFieldName , objVbsDb )
'***********************************************
  funStrErrorBadFieldForSqlRecordSet = ""
  on error resume next
  isObject( objVbsDb( "SqlRecordSet" )( strFieldName ) )
  if err.number <> 0 then
    ' the developer has written a bad field for the grid
    if inStr( strFieldName , "," ) > 0 then
      ' the developer has used ',' as a separator
      funStrErrorBadFieldForSqlRecordSet = strPropertyName & " property is wrong: do not use ',' as a separator. Use ';' instead."
    else
      ' the developer has used a bad field name
      funStrErrorBadFieldForSqlRecordSet = strPropertyName & " property is wrong: " & _
        strFieldNameToDisplay & " is not a field in the sql query.<br>" & _
        "<br><br><br><br>" & _
        "Attention: if your Sql is a Join and " & strFieldNameToDisplay & _
        " is contained in more than one table of your query, you may have to disambiguate " & _
        "the field name " & _
        " using the tableName.fieldName notation (it depends on the kind of " & _
        "Ole Db provider you are using, if you really need to disambiguate this way)."
    end if
  end if
end function

'***********************************************
sub vbsDbCheckSqlRecordSetFieldsErrorsCheckField( strPropertyName , strFieldNameToDisplay , strFieldName , objVbsDb )
'***********************************************
  dim strErrorMessage
  strErrorMessage = funStrErrorBadFieldForSqlRecordSet( strPropertyName , strFieldNameToDisplay , strFieldName , objVbsDb )
  if strErrorMessage <> "" then
    ' strFieldName is not a good field name for the Sql query
    drawError strErrorMessage , objVbsDb
  end if
end sub

'***********************************************
sub vbsDbCheckSqlRecordSetFieldsErrors( strPropertyName , strAssociatedDictionary , objVbsDb )
'***********************************************
  dim strFieldName , strFieldNameToDisplay
  for each strFieldName in objVbsDb( strAssociatedDictionary )
    strFieldNameToDisplay = funStrFieldNameToDisplay( strFieldName , strAssociatedDictionary , objVbsDb )
    vbsDbCheckSqlRecordSetFieldsErrorsCheckField strPropertyName , strFieldNameToDisplay , strFieldName , objVbsDb
  next
end sub

'***********************************************
sub vbsDbCheckGridFieldsErrors( objVbsDb )
'***********************************************
  vbsDbCheckSqlRecordSetFieldsErrors "GridFields" , "DictionaryGridFields" , objVbsDb
end sub

'***********************************************
sub vbsDbCheckGridHideFieldsErrors( objVbsDb )
'***********************************************
  vbsDbCheckSqlRecordSetFieldsErrors "GridHideFields" , "DictionaryGridHideFields" , objVbsDb
end sub

'***********************************************
sub vbsDbCheckFormFieldsErrors( objVbsDb )
'***********************************************
  vbsDbCheckSqlRecordSetFieldsErrors "FormFields" , "DictionaryFormFields" , objVbsDb
end sub

'***********************************************
sub vbsDbCheckFormHideFieldsErrors( objVbsDb )
'***********************************************
  vbsDbCheckSqlRecordSetFieldsErrors "FormHideFields" , "DictionaryFormHideFields" , objVbsDb
end sub

'***********************************************
sub vbsDbCheckSearchFieldsErrors( objVbsDb )
'***********************************************
  vbsDbCheckSqlRecordSetFieldsErrors "SearchFields" , "DictionarySearchFields" , objVbsDb
end sub

'***********************************************
sub vbsDbCheckSearchHideFieldsErrors( objVbsDb )
'***********************************************
  vbsDbCheckSqlRecordSetFieldsErrors "SearchHideFields" , "DictionarySearchHideFields" , objVbsDb
end sub

'***********************************************
sub vbsDbCheckViewFieldFormats( objVbsDb )
'***********************************************
  vbsDbCheckSqlRecordSetFieldsErrors "ViewFieldFormats" , "DictionaryViewFieldFormats" , objVbsDb
end sub

'***********************************************
sub vbsDbCheckInputSelectFieldsErrorsForSqlRecordSet( objVbsDb )
'***********************************************
  vbsDbCheckSqlRecordSetFieldsErrors "InputSelectFields" , "DictionaryInputSelectFields" , objVbsDb
end sub

'***********************************************
sub vbsDbCheckEditKeyFieldsErrorsForSqlRecordSet( objVbsDb )
'***********************************************
  vbsDbCheckSqlRecordSetFieldsErrors "EditKeyFields" , "DictionaryEditKeyFields" , objVbsDb
end sub

'***********************************************
sub vbsDbCheckErrorBadFieldsForSqlRecordSet( objVbsDb )
'***********************************************
  vbsDbCheckGridFieldsErrors objVbsDb
  vbsDbCheckGridHideFieldsErrors objVbsDb
  vbsDbCheckFormFieldsErrors objVbsDb
  vbsDbCheckFormHideFieldsErrors objVbsDb
  vbsDbCheckSearchFieldsErrors objVbsDb
  vbsDbCheckSearchHideFieldsErrors objVbsDb
  'vbsDbCheckSearchAliasFieldsErrors objVbsDb    ' cannot be checked here, it should be checked before assigning the SqlRecordSet
  vbsDbCheckViewFieldFormats objVbsDb
  vbsDbCheckInputSelectFieldsErrorsForSqlRecordSet objVbsDb
  vbsDbCheckEditKeyFieldsErrorsForSqlRecordSet objVbsDb
end sub

'***********************************************
function funStrErrorBadFieldForEditTableRecordSetActually( strPropertyName , strFieldNameToDisplay , strFieldName , objVbsDb )
'***********************************************
  funStrErrorBadFieldForEditTableRecordSetActually = ""
  if not funBolAdoxIsEditField( strFieldName , objVbsDb ) then
    ' the developer has written a bad field for the grid
    if inStr( strFieldName , "," ) > 0 then
      ' the developer has used ',' as a separator
      funStrErrorBadFieldForEditTableRecordSetActually = strPropertyName & " property is wrong: do not use ',' as a separator. Use ';' instead."
    else
      ' the developer has used a bad field name
      funStrErrorBadFieldForEditTableRecordSetActually = strPropertyName & " property is wrong: " & strFieldNameToDisplay & " is not a field in the edit table"
    end if
  end if
end function

'***********************************************
function funStrErrorBadFieldForEditTableRecordSet( strPropertyName , strFieldNameToDisplay , strFieldName , objVbsDb )
'***********************************************
  if objVbsDb( "EditTableName" ) = "" then
    funStrErrorBadFieldForEditTableRecordSet = ""
  else
    funStrErrorBadFieldForEditTableRecordSet = funStrErrorBadFieldForEditTableRecordSetActually( strPropertyName , strFieldNameToDisplay , strFieldName , objVbsDb )
  end if
end function

'***********************************************
sub vbsDbCheckEditTableRecordSetFieldsErrorsCheckField( strPropertyName , strFieldNameToDisplay , strFieldName , objVbsDb )
'***********************************************
  dim strErrorMessage
  strErrorMessage = funStrErrorBadFieldForEditTableRecordSet( strPropertyName , strFieldNameToDisplay , strFieldName , objVbsDb )
  if strErrorMessage <> "" then
    ' strFieldName is not a good field name for the edit table
    drawError strErrorMessage , objVbsDb
  end if
end sub

'***********************************************
sub vbsDbCheckEditTableRecordSetFieldsErrors( strPropertyName , strAssociatedDictionary , objVbsDb )
'***********************************************
  dim strFieldName , strFieldNameToDisplay
  for each strFieldName in objVbsDb( strAssociatedDictionary )
    strFieldNameToDisplay = funStrFieldNameToDisplay( strFieldName , strAssociatedDictionary , objVbsDb )
    vbsDbCheckEditTableRecordSetFieldsErrorsCheckField strPropertyName , strFieldNameToDisplay , strFieldName , objVbsDb
  next
end sub

'***********************************************
sub vbsDbCheckEditFieldsErrors( objVbsDb )
'***********************************************
  vbsDbCheckEditTableRecordSetFieldsErrors "EditFields" , "DictionaryEditFields" , objVbsDb
end sub

'***********************************************
sub vbsDbCheckEditAddFieldDefaultsErrors( objVbsDb )
'***********************************************
  vbsDbCheckEditTableRecordSetFieldsErrors "EditAddFieldDefaults" , "DictionaryEditAddFieldDefaults" , objVbsDb
end sub

'***********************************************
sub vbsDbCheckEditUpdateFieldDefaultsErrors( objVbsDb )
'***********************************************
  vbsDbCheckEditTableRecordSetFieldsErrors "EditUpdateFieldDefaults" , "DictionaryEditUpdateFieldDefaults" , objVbsDb
end sub

'***********************************************
sub vbsDbCheckEditMemoFieldsErrors( objVbsDb )
'***********************************************
  vbsDbCheckEditTableRecordSetFieldsErrors "EditMemoFields" , "DictionaryEditMemoFields" , objVbsDb
end sub

'***********************************************
sub vbsDbCheckInputSelectFieldsErrorsForEditTableRecordSet( objVbsDb )
'***********************************************
  'vbsDbCheckEditTableRecordSetFieldsErrors "InputSelectFields" , "DictionaryInputSelectFields" , objVbsDb
end sub

'***********************************************
sub vbsDbCheckEditKeyFieldsErrorsForEditTableRecordSet( objVbsDb )
'***********************************************
  vbsDbCheckEditTableRecordSetFieldsErrors "EditKeyFields" , "DictionaryEditKeyFields" , objVbsDb
end sub

'***********************************************
sub vbsDbCheckEditValidateRequiredFieldsErrorsForEditTableRecordSet( objVbsDb )
'***********************************************
  vbsDbCheckEditTableRecordSetFieldsErrors "EditValidateRequiredFields" , "DictionaryEditValidateRequiredFields" , objVbsDb
end sub

'***********************************************
sub vbsDbCheckErrorBadFieldsForEditTableRecordSetActually( objVbsDb )
'***********************************************
  vbsDbCheckEditFieldsErrors objVbsDb
  vbsDbCheckEditAddFieldDefaultsErrors objVbsDb
  vbsDbCheckEditUpdateFieldDefaultsErrors objVbsDb
  vbsDbCheckEditMemoFieldsErrors objVbsDb
  vbsDbCheckInputSelectFieldsErrorsForEditTableRecordSet objVbsDb
  vbsDbCheckEditKeyFieldsErrorsForEditTableRecordSet objVbsDb
  vbsDbCheckEditValidateRequiredFieldsErrorsForEditTableRecordSet objVbsDb
end sub

'***********************************************
sub vbsDbCheckErrorBadFieldsForEditTableRecordSet( objVbsDb )
'***********************************************
  if objVbsDb( "EditTableName" ) <> "" then
    vbsDbCheckErrorBadFieldsForEditTableRecordSetActually objVbsDb
  end if
end sub

'***********************************************
sub vbsDbCheckErrorBadFieldForSqlRecordSetAndEditTableRecordSet( strPropertyName , strFieldNameToDisplay , strFieldName , objVbsDb )
'***********************************************
  dim strErrorMessageForSqlRecordSet , strErrorMessageForEditTableRecordSet
  strErrorMessageForSqlRecordSet = funStrErrorBadFieldForSqlRecordSet( strPropertyName , strFieldNameToDisplay , strFieldName , objVbsDb )
  strErrorMessageForEditTableRecordSet = funStrErrorBadFieldForEditTableRecordSet( strPropertyName , strFieldNameToDisplay , strFieldName , objVbsDb )
  if ( strErrorMessageForSqlRecordSet <> "" ) and ( strErrorMessageForEditTableRecordSet <> "" )then
    ' strFieldName is not a good field name for the Sql query and it is not a good field for the edit table
    drawError strErrorMessageForSqlRecordSet & "<br>" & strErrorMessageForEditTableRecordSet , objVbsDb
    if ( strErrorMessageForSqlRecordSet <> "" ) then
      ' strFieldName is not a good field name for the Sql query
      drawError strErrorMessageForSqlRecordSet , objVbsDb
    end if
    if strErrorMessageForEditTableRecordSet <> "" then
      ' strFieldName is not a good field name for the edit table
      drawError strErrorMessageForEditTableRecordSet , objVbsDb
    end if
  end if
end sub

'***********************************************
sub vbsDbCheckErrorBadPropertyFieldsForSqlRecordSetAndEditTableRecordSet( strPropertyName , strAssociatedDictionary , objVbsDb )
'***********************************************
  dim strFieldName , strFieldNameToDisplay
  for each strFieldName in objVbsDb( strAssociatedDictionary )
    strFieldNameToDisplay = funStrFieldNameToDisplay( strFieldName , strAssociatedDictionary , objVbsDb )
    vbsDbCheckErrorBadFieldForSqlRecordSetAndEditTableRecordSet strPropertyName , strFieldNameToDisplay , strFieldName , objVbsDb
  next
end sub

'***********************************************
sub vbsDbCheckGlobalFieldHeadersErrors( objVbsDb )
'***********************************************
  vbsDbCheckErrorBadPropertyFieldsForSqlRecordSetAndEditTableRecordSet "GlobalFieldHeaders" , "DictionaryGlobalFieldHeaders" , objVbsDb
end sub

'***********************************************
sub vbsDbCheckGlobalBooleanFieldsErrors( objVbsDb )
'***********************************************
  vbsDbCheckErrorBadPropertyFieldsForSqlRecordSetAndEditTableRecordSet "GlobalBooleanFields" , "DictionaryGlobalBooleanFields" , objVbsDb
end sub

'***********************************************
sub vbsDbCheckErrorBadFieldsForSqlRecordSetAndEditTableRecordSet( objVbsDb )
'***********************************************
  vbsDbCheckGlobalFieldHeadersErrors objVbsDb
  vbsDbCheckGlobalBooleanFieldsErrors objVbsDb
end sub

'***********************************************
'***********************************************
sub vbsDbCheckErrorBadFieldNames( objVbsDb )
'***********************************************
'***********************************************
  vbsDbCheckErrorBadFieldsForSqlRecordSet objVbsDb
  vbsDbCheckErrorBadFieldsForEditTableRecordSet objVbsDb
  vbsDbCheckErrorBadFieldsForSqlRecordSetAndEditTableRecordSet objVbsDb
end sub

'***********************************************
sub vbsDbCheckErrorProOnlyMethodDrawError( strMethodError , objVbsDb )
'***********************************************
  drawError strMethodError & " method is not supported by VBSdb Free. It requires VBSdb Professional." , objVbsDb
end sub

'***********************************************
'***********************************************
sub vbsDbCheckErrorProOnlyMethodVbsDbGetOutputFieldValue( objVbsDb )
'***********************************************
'***********************************************
  if ( not funBolIsProVersion() ) then
    vbsDbCheckErrorProOnlyMethodDrawError "VbsDbGetOutputFieldValue" , objVbsDb
  end if
end sub
%>
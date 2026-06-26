<%
'***********************************************
function funVarVbsDbGetOutputFieldValueWithPositionedSqlRecordSet( strFieldName , objVbsDb )
'***********************************************
  dim objField , bolFieldNameExists
  bolFieldNameExists = false
  for each objField in objVbsDb( "SqlRecordSet" ).fields
    if uCase( objField.name ) = uCase( strFieldName ) then
      bolFieldNameExists = true
      funVarVbsDbGetOutputFieldValueWithPositionedSqlRecordSet = objField.value
    end if
  next
  if not bolFieldNameExists then
    ' the sql recordset is not empty, but the requested field name is not in the sql recordset
    drawError "VbsDbGetOutputFieldValue request error. " & strFieldName & " is not a valid field name. Please, check the Sql property." , objVbsDb
  end if
end function

'***********************************************
sub vbsDbGetOutputFieldValueResetSqlRecordSet( bolEndOfFile , lngAbsolutePosition , byRef objVbsDb )
'***********************************************
  if bolEndOfFile then
    ' objVbsDb( "SqlRecordSet" ) was at the end of file
    if ( objVbsDb( "SqlRecordSet" ).recordCount > 0 ) then
      ' objVbsDb( "SqlRecordSet" ) is not empty
      ' reset objVbsDb( "SqlRecordSet" ) to the End Of File
      objVbsDb( "SqlRecordSet" ).absolutePosition = objVbsDb( "SqlRecordSet" ).recordCount
      objVbsDb( "SqlRecordSet" ).moveNext
    end if
  else
    ' objVbsDb( "SqlRecordSet" ) was at lngAbsolutePosition position
    objVbsDb( "SqlRecordSet" ).absolutePosition = lngAbsolutePosition
  end if
end sub

'***********************************************
function funVarVbsDbGetOutputFieldValue_forViewScreen( strFieldName , byRef objVbsDb )
'***********************************************
  dim bolEndOfFile , lngAbsolutePosition
  bolEndOfFile = objVbsDb( "SqlRecordSet" ).eof
  lngAbsolutePosition = objVbsDb( "SqlRecordSet" ).absolutePosition
  
  objVbsDb( "SqlRecordSet" ).absolutePosition = objVbsDb( "FormAbsolutePosition" )
  funVarVbsDbGetOutputFieldValue_forViewScreen = _
    funVarVbsDbGetOutputFieldValueWithPositionedSqlRecordSet( strFieldName , objVbsDb )
  
  'objVbsDb( "SqlRecordSet" ).absolutePosition = lngAbsolutePosition ' resets recordset position
  vbsDbGetOutputFieldValueResetSqlRecordSet bolEndOfFile , lngAbsolutePosition , objVbsDb 
end function

'***********************************************
function VbsDbGetOutputFieldValue_forViewScreen( strFieldName , byRef objVbsDb )
'***********************************************
  if ( objVbsDb( "SqlRecordSet" ).recordCount <> 0 ) then 
    ' objVbsDb( "SqlRecordSet" ) is not empty
    VbsDbGetOutputFieldValue_forViewScreen = _
      funVarVbsDbGetOutputFieldValue_forViewScreen( strFieldName , objVbsDb )
  else
    ' objVbsDb( "SqlRecordSet" ) is empty
    VbsDbGetOutputFieldValue_forViewScreen = constVbsDbGetOutputFieldValueEmptySqlRecordSet
  end if
end function

'***********************************************
function VbsDbGetOutputFieldValue_forEditScreen( strFieldName , byRef objVbsDb )
'***********************************************
  on error resume next
  VbsDbGetOutputFieldValue_forEditScreen = _
    objVbsDb( "EditTableRecordset" )( strFieldName )
  if err.number <> 0 then
    drawError "VbsDbGetOutputFieldValue method. " & _
    "The field " & strFieldName & " is not a field in the Edit table.<br><br>" & _
    "If you wish to get the value of a field of your Sql query, that is not a field in " & _
    "your edit table, you can easily manage this error by invoking the " & _
    "VbsDbGetScreenType method to check " & _
    "the current screen type.", objVbsDb
  end if
end function

'***********************************************
function VbsDbGetOutputFieldValueActually( strFieldName , byRef objVbsDb )
'***********************************************
  if VbsDbGetScreenType( objVbsDb ) = "View" then
    VbsDbGetOutputFieldValueActually = _
      VbsDbGetOutputFieldValue_forViewScreen( strFieldName , objVbsDb )
  else
    ' the current screen is an either an Update screen or a Delete screen
    VbsDbGetOutputFieldValueActually = _
      VbsDbGetOutputFieldValue_forEditScreen( strFieldName , objVbsDb )
  end if
end function

'***********************************************
sub vbsDbCheckErrorVbsDbGetOutputFieldValueInvokedBeforeInvokingTheVbsDbMethod( objVbsDb )
'***********************************************
  if not isObject( objVbsDb( "SqlRecordSet" ) ) then
    ' probably, the VbsDbGetOutputFieldValue method is invoked before the VbsDb method is invoked
    drawError "VbsDbGetOutputFieldValue method. " & _
    "Please, perform the following debug steps:" & _
    "<div align=left>" & _
    "<ul><li>Control you invoke the VbsDb method before you invoke the VbsDbGetOutputFieldValue method</li>" & _
    "</ul>" & _
    "</div>" , objVbsDb
  end if
end sub

'***********************************************
'sub vbsDbCheckErrorVbsDbGetOutputFieldValueInvokedInWrongScreenMode( objVbsDb )
'***********************************************
'  if VbsDbGetScreenType( objVbsDb ) = "Search" or VbsDbGetScreenType( objVbsDb ) = "Add" then
'    drawError _
'      "The VbsDbGetOutputFieldValue method cannot be invoked when the " & _
'      "VBSdb object is either in a Search screen or in an Add screen: " & _
'      "no record is selected, so invoking the VbsDbGetOutputFieldValue in such screens " & _
'      "has no sense.<br><br>" & _
'      "You can easily manage this, by invoking the VbsDbGetScreenType method to check " & _
'      "the current screen type." , objVbsDb
'  end if
'end sub

'***********************************************
'***********************************************
function VbsDbGetOutputFieldValue( strFieldName , byRef objVbsDb )
'***********************************************
'***********************************************
  vbsDbCheckErrorProOnlyMethodVbsDbGetOutputFieldValue objVbsDb
  vbsDbCheckErrorVbsDbGetOutputFieldValueInvokedBeforeInvokingTheVbsDbMethod objVbsDb
  if VbsDbGetScreenType( objVbsDb ) = "Search" or VbsDbGetScreenType( objVbsDb ) = "Add" then
    VbsDbGetOutputFieldValue = constVbsDbGetOutputFieldValue_inWrongScreenType
  else
    VbsDbGetOutputFieldValue = VbsDbGetOutputFieldValueActually( strFieldName , objVbsDb )
  end if
end function
%>
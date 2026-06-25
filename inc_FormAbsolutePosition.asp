<%
'***********************************************
function funLngGridAbsolutePage( objVbsDb )
'***********************************************
  if objVbsDb( "FormAbsolutePosition" ) mod objVbsDb( "GridPageSize" ) = 0 then
    ' the record shown in the form is the last one of the page
    funLngGridAbsolutePage = objVbsDb( "FormAbsolutePosition" ) \ objVbsDb( "GridPageSize" )
  else
    ' the record shown in the form is not the last one of the page
    funLngGridAbsolutePage = ( objVbsDb( "FormAbsolutePosition" ) \ objVbsDb( "GridPageSize" ) ) + 1
  end if
end function

'***********************************************
function funBolObjFieldValueIsEqualToTheInsertedValue( objField , objVbsDb )
'***********************************************
  funBolObjFieldValueIsEqualToTheInsertedValue = _
    ( funVarDbEditFieldFormattedValue( objField.name , _
        funStrConvertNullToEmptyString( objField.value ) , objVbsDb ) _
      = _
      funStrDbEditFieldRequestFormattedValue( objField.name , objVbsDb ) )
end function

'***********************************************
function funBolEditableFieldNotInConflict( objField , objVbsDb )
'***********************************************
  if not funBolEditableEditField( objField.name , objVbsDb ) then
    funBolEditableFieldNotInConflict = true
  else
    funBolEditableFieldNotInConflict = _
      funBolObjFieldValueIsEqualToTheInsertedValue( objField , objVbsDb )
  end if
end function

'***********************************************
function funStrConvertNullToNullString( varValue )
'***********************************************
  if isNull( varValue ) then
    funStrConvertNullToNullString = "null"
  elseif varType( varValue ) = constVarTypeBoolean then
    if varValue then
      funStrConvertNullToEmptyString = "true"
    else
      funStrConvertNullToEmptyString = "false"
    end if
  else
    funStrConvertNullToNullString = cStr( varValue )
  end if
end function

'***********************************************
function funStrEditFieldFormattedValueForFormAbsolutePosition( objField , objVbsDb )
'***********************************************
  funStrEditFieldFormattedValueForFormAbsolutePosition = _
    funVarDbEditFieldFormattedValue( objField.name , _
    objVbsDb( "DictionaryEditAddFieldDefaults" )( uCase( objField.name ) ) , objVbsDb )
  if inStr( funStrEditFieldFormattedValueForFormAbsolutePosition , "'" ) = 1 then
    ' funVarDbEditFieldFormattedValue added a starting ' and an ending '
    funStrEditFieldFormattedValueForFormAbsolutePosition = _
      mid( funStrEditFieldFormattedValueForFormAbsolutePosition , 2 , _
        len( funStrEditFieldFormattedValueForFormAbsolutePosition ) - 2 )
  end if
end function

'***********************************************
function funBolObjFieldValueIsEqualToTheEditAddDefaultValue( objField , objVbsDb )
'***********************************************
  funBolObjFieldValueIsEqualToTheEditAddDefaultValue = _
    ( funStrConvertNullToNullString( objField.value ) = _
      funStrEditFieldFormattedValueForFormAbsolutePosition( objField , objVbsDb ) )
  if isDate( objField.value ) _
    and isDate( objVbsDb( "DictionaryEditAddFieldDefaults" )( objField.name ) ) then
    funBolObjFieldValueIsEqualToTheEditAddDefaultValue = _
      year( objField.value ) = year( objVbsDb( "DictionaryEditAddFieldDefaults" )( objField.name ) ) and _
      month( objField.value ) = year( objVbsDb( "DictionaryEditAddFieldDefaults" )( objField.name ) ) and _
      day( objField.value ) = year( objVbsDb( "DictionaryEditAddFieldDefaults" )( objField.name ) ) and _
      hour( objField.value ) = hour( objVbsDb( "DictionaryEditAddFieldDefaults" )( objField.name ) ) and _
      minute( objField.value ) = minute( objVbsDb( "DictionaryEditAddFieldDefaults" )( objField.name ) ) and _
      second( objField.value ) = second( objVbsDb( "DictionaryEditAddFieldDefaults" )( objField.name ) )
  end if
end function

'***********************************************
function funBolNotInConflictHandlingEditAddFieldDefaults( objField , objVbsDb )
'***********************************************
  if funBolEditableEditField( objField.name , objVbsDb ) or _
    ( not objVbsDb( "DictionaryEditAddFieldDefaults" ).exists( uCase( objField.name ) ) ) then
    funBolNotInConflictHandlingEditAddFieldDefaults = true
  else
    funBolNotInConflictHandlingEditAddFieldDefaults = _
      funBolObjFieldValueIsEqualToTheEditAddDefaultValue( objField , objVbsDb )
  end if
end function

'***********************************************
function funBolStrEditFieldNameValueIsNotInConflictWithTheRecordAdded( objField , objVbsDb )
'***********************************************
  funBolStrEditFieldNameValueIsNotInConflictWithTheRecordAdded = _
    funBolEditableFieldNotInConflict( objField , objVbsDb ) and _
    funBolNotInConflictHandlingEditAddFieldDefaults( objField , objVbsDb )
end function

'***********************************************
function funBolCurrentFieldValueIsNotInConflictWithTheRecordAdded( objField , objVbsDb )
'***********************************************
  if not funBolAdoxIsEditField( objField.name , objVbsDb ) then
    ' objField is not a field of the edit table
    funBolCurrentFieldValueIsNotInConflictWithTheRecordAdded = true
  else
    funBolCurrentFieldValueIsNotInConflictWithTheRecordAdded = _
      ( ( not objVbsDb( "DictionaryEditFields" ).exists( uCase( objField.name ) ) ) and _
        ( not objVbsDb( "DictionaryEditAddFieldDefaults" ).exists( uCase( objField.name ) ) ) or _
        funBolStrEditFieldNameValueIsNotInConflictWithTheRecordAdded( objField , objVbsDb ) )
  end if
end function

'***********************************************
function funBolCurrentRecordIsTheOneAdded( objVbsDb )
'***********************************************
  dim intFieldIndex
  intFieldIndex = 0
  funBolCurrentRecordIsTheOneAdded = true
  while ( intFieldIndex < objVbsDb( "SqlRecordSet" ).fields.count ) and _
          funBolCurrentRecordIsTheOneAdded
    funBolCurrentRecordIsTheOneAdded = funBolCurrentRecordIsTheOneAdded and _
      funBolCurrentFieldValueIsNotInConflictWithTheRecordAdded( _
        objVbsDb( "SqlRecordSet" ).fields( intFieldIndex ) , objVbsDb )
    intFieldIndex = intFieldIndex + 1
  wend
end function

'***********************************************
sub setSqlRecordSetForFormAbsolutePositionForApplyAddNotEmpty( byRef objVbsDb )
'***********************************************
  dim bolDoLoop
  bolDoLoop = true
  while bolDoLoop
    if funBolCurrentRecordIsTheOneAdded( objVbsDb ) then
      ' the current recordset record is the one the user has just added
      bolDoLoop = false
    else
      objVbsDb( "SqlRecordSet" ).moveNext
      bolDoLoop = not objVbsDb( "SqlRecordSet" ).eof
    end if
  wend
end sub

'***********************************************
function funLngFormAbsolutePositionForApplyAddNotEmptyWithSqlRecordSetAlreadySet( objVbsDb )
'***********************************************
  if not objVbsDb( "SqlRecordSet" ).eof then
    ' objVbsDb( "SqlRecordSet" ) points to the record just addedd
    funLngFormAbsolutePositionForApplyAddNotEmptyWithSqlRecordSetAlreadySet = _
      objVbsDb( "SqlRecordSet" ).absolutePosition
  else
    funLngFormAbsolutePositionForApplyAddNotEmptyWithSqlRecordSetAlreadySet = _
      objVbsDb( "SqlRecordSet" ).recordCount
  end if
end function

'***********************************************
function funLngFormAbsolutePositionForApplyAddNotEmpty( byRef objVbsDb )
'***********************************************
  if objVbsDb( "SqlRecordSet" ).recordCount > 0 then
    objVbsDb( "SqlRecordSet" ).moveFirst
  end if
  setSqlRecordSetForFormAbsolutePositionForApplyAddNotEmpty( objVbsDb )
  funLngFormAbsolutePositionForApplyAddNotEmpty = _
    funLngFormAbsolutePositionForApplyAddNotEmptyWithSqlRecordSetAlreadySet( objVbsDb )
  if objVbsDb( "SqlRecordSet" ).recordCount > 0 then
    objVbsDb( "SqlRecordSet" ).moveFirst
  end if
end function

'***********************************************
function funLngFormAbsolutePositionForApplyAddActually( byRef objVbsDb )
'***********************************************
  if objVbsDb( "SqlRecordSet" ).eof then
    funLngFormAbsolutePositionForApplyAddActually = 1
  else
    funLngFormAbsolutePositionForApplyAddActually = _
      funLngFormAbsolutePositionForApplyAddNotEmpty( objVbsDb )
  end if
end function

'***********************************************
function funLngFormAbsolutePositionForApplyAdd( byRef objVbsDb )
'***********************************************
  if objVbsDb( "FormExactAfterEdit" ) then
    funLngFormAbsolutePositionForApplyAdd = _
      funLngFormAbsolutePositionForApplyAddActually( objVbsDb )
  else
    funLngFormAbsolutePositionForApplyAdd = objVbsDb( "SqlRecordSet" ).recordCount
  end if
end function

'***********************************************
function funBolCurrentRecordIsTheOneUpdated( byRef objVbsDb )
'***********************************************
  funBolCurrentRecordIsTheOneUpdated = _
    ( funStrVbsDbEditWhere( objVbsDb ) = objVbsDb( "RequestVBSdbEditWhere" ) )
end function

'***********************************************
function funLngFormAbsolutePositionForApplyUpdateNotEmpty( byRef objVbsDb )
'***********************************************
  while ( not objVbsDb( "SqlRecordSet" ).eof ) and _
    ( not funBolCurrentRecordIsTheOneUpdated( objVbsDb ) )
    objVbsDb( "SqlRecordSet" ).moveNext
  wend
  if not objVbsDb( "SqlRecordSet" ).eof then
    ' objVbsDb( "SqlRecordSet" ) points to the just updated record
    funLngFormAbsolutePositionForApplyUpdateNotEmpty = _
      objVbsDb( "SqlRecordSet" ).absolutePosition
  end if
end function

'***********************************************
function funLngFormAbsolutePositionForApplyUpdateActually( byRef objVbsDb )
'***********************************************
  if objVbsDb( "SqlRecordSet" ).eof then
    funLngFormAbsolutePositionForApplyUpdateActually = 1
  else
    objVbsDb( "SqlRecordSet" ).moveFirst
    funLngFormAbsolutePositionForApplyUpdateActually = _
      funLngFormAbsolutePositionForApplyUpdateNotEmpty( objVbsDb )
  end if
  'objVbsDb( "SqlRecordSet" ).find( objVbsDb( "RequestVBSdbEditWhere" ) )
end function

'***********************************************
function funLngFormAbsolutePositionForApplyUpdate( byRef objVbsDb )
'***********************************************
  if objVbsDb( "FormExactAfterEdit" ) then
    funLngFormAbsolutePositionForApplyUpdate = _
      funLngFormAbsolutePositionForApplyUpdateActually( objVbsDb )
  else
    funLngFormAbsolutePositionForApplyUpdate = objVbsDb( "FormAbsolutePosition" )
  end if
end function

'***********************************************
'***********************************************
sub vbsDbSetFormAbsolutePosition( byRef objVbsDb )
'***********************************************
'***********************************************
  select case objVbsDb( "RequestVBSdbClickClass" )
	  case "VBSdbGoToGridRow" , "VBSdbCancel"
	    objVbsDb( "FormAbsolutePosition" ) = objVbsDb( "RequestVBSGridIndex" )
  	case "VBSdbGridFirst"
  	  objVbsDb( "FormAbsolutePosition" ) = 1
  	case "VBSdbGridPrev"
  	  objVbsDb( "FormAbsolutePosition" ) = ( ( funLngGridAbsolutePage( objVbsDb ) - 2 ) * objVbsDb( "GridPageSize" ) ) + 1
  	case "VBSdbGridNext"
  	  objVbsDb( "FormAbsolutePosition" ) = ( funLngGridAbsolutePage( objVbsDb ) * objVbsDb( "GridPageSize" ) ) + 1
  	case "VBSdbGridLast"
  	  objVbsDb( "FormAbsolutePosition" ) = objVbsDb( "SqlRecordSet" ).recordCount
  	  objVbsDb( "FormAbsolutePosition" ) = ( ( funLngGridAbsolutePage( objVbsDb ) - 1 ) * objVbsDb( "GridPageSize" ) ) + 1
  	case "VBSdbApplyFilter"
  	  objVbsDb( "FormAbsolutePosition" ) = 1
	  case "VBSdbApplyAdd"
      objVbsDb( "FormAbsolutePosition" ) = funLngFormAbsolutePositionForApplyAdd( objVbsDb )
	  case "VBSdbApplyUpdate"
      objVbsDb( "FormAbsolutePosition" ) = funLngFormAbsolutePositionForApplyUpdate( objVbsDb )
  end select
  if objVbsDb( "FormAbsolutePosition" ) > objVbsDb( "SqlRecordSet" ).recordCount then
    ' the last record was deleted
    objVbsDb( "FormAbsolutePosition" ) = funNumMax( objVbsDb( "SqlRecordSet" ).recordCount , 1 )
  end if
  if objVbsDb( "FormAbsolutePosition" ) < 1 then
    ' there has been a 'prev' button leading to the first page, and then a browser refresh
    objVbsDb( "FormAbsolutePosition" ) = 1
  end if
end sub

'***********************************************
'***********************************************
sub vbsDbSetGridAbsolutePage( byRef objVbsDb )
'***********************************************
'***********************************************
  objVbsDb( "GridAbsolutePage" ) = funLngGridAbsolutePage( objVbsDb )		
end sub
%>
<%
'***********************************************
'***********************************************
function funBolEditableEditField( strEditFieldName , objVbsDb )
'***********************************************
'***********************************************
  funBolEditableEditField = ( _
    objVbsDb( "DictionaryEditFields" ).exists( uCase( strEditFieldName ) ) and _
    ( not funBolEditFieldIsHideField( strEditFieldName , objVbsDb ) ) and _
    ( not funBolEditFieldIsReadOnly( strEditFieldName , objVbsDb ) ) )
end function

'***********************************************
function funStrEditFieldNamesListForEditableEditFields( objVbsDb )
'***********************************************
  dim strEditFieldName
  funStrEditFieldNamesListForEditableEditFields = ""
  for each strEditFieldName in objVbsDb( "DictionaryEditFields" )
    if funBolEditableEditField( strEditFieldName , objVbsDb ) then
  	  funStrEditFieldNamesListForEditableEditFields = _
  	    funStrEditFieldNamesListForEditableEditFields & _
  	    funStrGlobalDbTypeFieldName( strEditFieldName , objVbsDb ) & " , "
  	end if
  next
end function

'***********************************************
function funStrEditFieldNamesListForEditFieldDefaults( objVbsDb )
'***********************************************
  if objVbsDb( "RequestVBSdbClickClass" ) = "VBSdbApplyAdd" then
    ' user asked to add a record
    funStrEditFieldNamesListForEditFieldDefaults = _
      funStrEditFieldNamesListForEditAddFieldDefaults( objVbsDb )
  else
    ' user asked to update a record
    funStrEditFieldNamesListForEditFieldDefaults = _
      funStrEditFieldNamesListForEditUpdateFieldDefaults( objVbsDb )
  end if
end function

'***********************************************
'***********************************************
function funStrEditFieldNamesList( objVbsDb )
'***********************************************
'***********************************************
  funStrEditFieldNamesList = funStrEditFieldNamesListForEditableEditFields( objVbsDb )
  funStrEditFieldNamesList = funStrEditFieldNamesList & funStrEditFieldNamesListForEditFieldDefaults( objVbsDb )
  funStrEditFieldNamesList = cutFinal( funStrEditFieldNamesList , " , " )
end function

'***********************************************
function funVarNumericEditFieldFormattedValueActually( strFieldValue )
'***********************************************
  funVarNumericEditFieldFormattedValueActually = replace( strFieldValue , "," , "." )
end function

'***********************************************
function funVarNumericEditFieldFormattedValue( strFieldName , strFieldValue , objVbsDb )
'***********************************************
  if trim( strFieldValue ) = "" then
    ' user inserted an empty numeric field
    funVarNumericEditFieldFormattedValue = "null"
  else
    if isNumeric( strFieldValue ) then
      funVarNumericEditFieldFormattedValue = funVarNumericEditFieldFormattedValueActually( strFieldValue )
    else
      drawError strFieldName & " has not been assigned a numeric value!" , objVbsDb
    end if
  end if
end function

'***********************************************
function funBolStringEditFieldProperFormattedValueIsNull( strFieldName , strControlValue , objVbsDb )
'***********************************************
  funBolStringEditFieldProperFormattedValueIsNull = _
    ( funBolAdoxFieldIsNullable( strFieldName , objVbsDb ) and _
    ( isNull( strControlValue ) or ( strControlValue = "" ) ) )
end function

'***********************************************
function funStrReverseHtmlEncode( strControlValue )
'***********************************************
  funStrReverseHtmlEncode = replace( strControlValue , "&quot;" , """" )
  funStrReverseHtmlEncode = replace( funStrReverseHtmlEncode , "&amp;" , "&" )
  funStrReverseHtmlEncode = replace( funStrReverseHtmlEncode , "&lt;" , "<" )
  funStrReverseHtmlEncode = replace( funStrReverseHtmlEncode , "&gt;" , ">" )
  'funStrReverseHtmlEncode = replace( funStrReverseHtmlEncode , "<br>" , chr( 10 ) )
end function

'***********************************************
function funStrHandleSingleQuoteForUpdate( strControlValue )
'***********************************************
  funStrHandleSingleQuoteForUpdate = replace( strControlValue , "'" , "''" )
end function

'***********************************************
function funStrStringEditFieldFormattedValueNonEmpty( strControlValue )
'***********************************************
  funStrStringEditFieldFormattedValueNonEmpty = funStrReverseHtmlEncode( strControlValue )
  funStrStringEditFieldFormattedValueNonEmpty = funStrHandleSingleQuoteForUpdate( funStrStringEditFieldFormattedValueNonEmpty )
end function

'***********************************************
function funStrStringEditFieldFormattedValue( strFieldName , strControlValue , objVbsDb )
'***********************************************
  if funBolStringEditFieldProperFormattedValueIsNull( strFieldName , strControlValue , objVbsDb ) then
    ' the field value must become null
    funStrStringEditFieldFormattedValue = "null"
  else
    funStrStringEditFieldFormattedValue = "'" & _
      funStrStringEditFieldFormattedValueNonEmpty( strControlValue ) & "'"
  end if
end function

'***********************************************
'***********************************************
function funVarDbEditFieldFormattedValue( strFieldName , strFieldValue , objVbsDb )
'***********************************************
'***********************************************
  'select case objVbsDb( "EditTableRecordSet" ).fields( strFieldName ).type
  select case funIntAdoxFieldType( strFieldName , objVbsDb )
  	case vbsDbAdSmallInt, vbsDbAdInteger, vbsDbAdDecimal, vbsDbAdSingle, vbsDbAdDouble, vbsDbAdCurrency, _
  	  vbsDbAdTinyInt, vbsDbAdUnsignedTinyInt, vbsDbAdUnsignedSmallInt, vbsDbAdUnsignedInt , vbsDbAdBigInt, vbsDbAdUnsignedBigInt, _
  	  vbsDbAdNumeric , vbsDbAdVarNumeric , vbsDbAdGUID
  	  ' 2 Byte Integer, 4 Byte Integer, Currency
  			if funBolInputFieldIsBoolean( strFieldName , objVbsDb ) then
  				' the field is defined boolean using the GlobalBooleanFields property
					funVarDbEditFieldFormattedValue = funStrDbEditBooleanFieldFormattedValue( strFieldValue , objVbsDb )
  			else
  				' the field is not defined boolean using the GlobalBooleanFields property
					funVarDbEditFieldFormattedValue = funVarNumericEditFieldFormattedValue( strFieldName , strFieldValue , objVbsDb )
				end if
  	case vbsDbAdBoolean 	
  	  ' Boolean True/False
  	  funVarDbEditFieldFormattedValue = funStrDbEditBooleanFieldFormattedValue( strFieldValue , objVbsDb )
  	case vbsDbAdChar, vbsDbAdWChar, vbsDbAdVarChar, vbsDbAdLongVarChar, vbsDbAdVarWChar, vbsDbAdLongVarWChar ' Char, WChar, VarChar, LongVarChar (Memo), VarWChar (Unicode String), LongVarWChar
  	  ' Char, WChar, VarChar, LongVarChar (Memo), VarWChar (Unicode String), LongVarWChar
  	  funVarDbEditFieldFormattedValue = funStrStringEditFieldFormattedValue( strFieldName , strFieldValue , objVbsDb)
  	case vbsDbAdVarBinary , vbsDbAdLongVarBinary 	
  	  ' ole object
  	  drawError "Function funVarDbEditFieldFormattedValue: field " &_
  	    strFieldName & " is an OLE object field. It cannot be used as an edit field. " & _
  	    "Check these properties:<div align=left><ul><li><b>EditFields</b></li><li>" & _
  	    "<b>EditHideFields</b></li><li><b>Sql</b></li></ul></div>" , objVbsDb
  	case vbsDbAdDate, vbsDbAdDBDate, vbsDbAdDBTime, vbsDbAdDBTimeStamp
  	  ' Date / Time 
  	  'funVarDbEditFieldFormattedValue = "#" & month( strFieldValue ) & "/" & day( strFieldValue ) & "/" & year( strFieldValue ) & " " & hour( strFieldValue ) & "." & minute( strFieldValue ) & "." & second( strFieldValue ) & "#"
  	  funVarDbEditFieldFormattedValue = funStrGlobalDbTypeDateFieldFormattedValue( strFieldValue , objVbsDb )
  	case else
  	  drawError "Function funVarDbEditFieldFormattedValue. Field " & strFieldName & ": type " &_
  	    funIntAdoxFieldType( strFieldName , objVbsDb ) &_
  	    " is not covered. Please, signal this bug message to <a href='mailto:support@vbsdb.com'>support@vbsdb.com</a>." , objVbsDb
  end select
end function

'***********************************************
function funStrDbEditFieldRequestFormattedValue( strFieldName , objVbsDb )
'***********************************************
  dim strFormFieldName
  strFormFieldName = funStrFormFieldName( strFieldName , objVbsDb )
  funStrDbEditFieldRequestFormattedValue = funVarDbEditFieldFormattedValue( strFieldName , cStr( request.form( strFormFieldName ) ) , objVbsDb )
end function

'***********************************************
function funStrDbEditFieldsRequestValuesListForEditableEditFields( objVbsDb )
'***********************************************
  dim strEditFieldName
  funStrDbEditFieldsRequestValuesListForEditableEditFields = ""
  for each strEditFieldName in objVbsDb( "DictionaryEditFields" )
    if funBolEditableEditField( strEditFieldName , objVbsDb ) then
      ' strEditFieldName is neither a read only field nor an hide field
      funStrDbEditFieldsRequestValuesListForEditableEditFields = funStrDbEditFieldsRequestValuesListForEditableEditFields & _
              funStrDbEditFieldRequestFormattedValue( strEditFieldName , objVbsDb ) & constRequestValueListDelimiter
    end if
  next
end function

'***********************************************
function funStrDbEditFieldsRequestValuesListForEditFieldDefaults( objVbsDb )
'***********************************************
  if objVbsDb( "RequestVBSdbClickClass" ) = "VBSdbApplyAdd" then
    ' user asked to add a record
    funStrDbEditFieldsRequestValuesListForEditFieldDefaults = _
      funStrDbEditFieldsRequestValuesListForEditAddFieldDefaults( objVbsDb )
  else
    ' user asked to update a record
    funStrDbEditFieldsRequestValuesListForEditFieldDefaults = _
      funStrDbEditFieldsRequestValuesListForEditUpdateFieldDefaults( objVbsDb )
  end if
end function

'***********************************************
function funStrDbEditFieldsRequestValuesList( objVbsDb )
'***********************************************
  funStrDbEditFieldsRequestValuesList = funStrDbEditFieldsRequestValuesListForEditableEditFields( objVbsDb )
  funStrDbEditFieldsRequestValuesList = funStrDbEditFieldsRequestValuesList & funStrDbEditFieldsRequestValuesListForEditFieldDefaults( objVbsDb )
  funStrDbEditFieldsRequestValuesList =  cutFinal( funStrDbEditFieldsRequestValuesList , constRequestValueListDelimiter )
end function

'***********************************************
function funStrSqlVbsDbEditDbAdd( objVbsDb )
'***********************************************
  funStrSqlVbsDbEditDbAdd = "insert into " &_
    funStrGlobalDbTypeEditTableName( objVbsDb ) &_
    " (" & funStrEditFieldNamesList( objVbsDb ) & ")" &_
    " values " & _
    " (" & replace( funStrDbEditFieldsRequestValuesList( objVbsDb ) , constRequestValueListDelimiter , " , " ) & ")"
end function

'***********************************************
sub vbsDbEditDbAddActually( objVbsDb )
'***********************************************
  dim strSql
  strSql = funStrSqlVbsDbEditDbAdd( objVbsDb )
  on error resume next
  objVbsDb( "Connection" ).execute strSql
  if err.number<>0 then
    vbsDbEditErrorMessageForAdd strSql , objVbsDb
  end if
end sub

'***********************************************
sub VbsDbEditDbAdd( objVbsDb )
'***********************************************
  'if funBolAddIsGranted( objVbsDb ) then
    vbsDbEditDbAddActually objVbsDb
  'end if
end sub

'***********************************************
function funStrSqlVbsDbUpdateSetClauseLoopWithArrays( arrStrEditFieldNamesToBeTrimmedList , arrStrDbEditFieldsRequestValuesList )
'***********************************************
  dim intArrayIndex
  funStrSqlVbsDbUpdateSetClauseLoopWithArrays = ""
  for intArrayIndex = 0 to uBound( arrStrEditFieldNamesToBeTrimmedList )
  	funStrSqlVbsDbUpdateSetClauseLoopWithArrays = funStrSqlVbsDbUpdateSetClauseLoopWithArrays &_
  		trim( arrStrEditFieldNamesToBeTrimmedList( intArrayIndex ) ) & "=" &_
  		arrStrDbEditFieldsRequestValuesList( intArrayIndex ) & " , "
  next
  funStrSqlVbsDbUpdateSetClauseLoopWithArrays = cutFinal( funStrSqlVbsDbUpdateSetClauseLoopWithArrays , " , " )
end function

'***********************************************
function funStrSqlVbsDbUpdateSetClauseLoop( strEditFieldNamesList , strDbEditFieldsRequestValuesList )
'***********************************************
  dim arrStrEditFieldNamesToBeTrimmedList , arrStrDbEditFieldsRequestValuesList
  arrStrEditFieldNamesToBeTrimmedList = split( strEditFieldNamesList , "," )  ' each field name will have to be trimmed, because of the two blanks next to each comma
  arrStrDbEditFieldsRequestValuesList = funArrStrSplit( strDbEditFieldsRequestValuesList , constRequestValueListDelimiter )
  funStrSqlVbsDbUpdateSetClauseLoop = funStrSqlVbsDbUpdateSetClauseLoopWithArrays( arrStrEditFieldNamesToBeTrimmedList ,  arrStrDbEditFieldsRequestValuesList )
end function

'***********************************************
function funStrSqlVbsDbUpdateSetClause( objVbsDb )
'***********************************************
  dim strEditFieldNamesList , strDbEditFieldsRequestValuesList
  strEditFieldNamesList = funStrEditFieldNamesList( objVbsDb )
  strDbEditFieldsRequestValuesList = funStrDbEditFieldsRequestValuesList( objVbsDb )
  funStrSqlVbsDbUpdateSetClause = funStrSqlVbsDbUpdateSetClauseLoop( strEditFieldNamesList , strDbEditFieldsRequestValuesList )
end function

'***********************************************
function funStrSqlVbsDbEditDbUpdate( objVbsDb )
'***********************************************
  funStrSqlVbsDbEditDbUpdate = "update " &_
    funStrGlobalDbTypeEditTableName( objVbsDb ) &_
    " set " & funStrSqlVbsDbUpdateSetClause( objVbsDb ) &_
    " where " & objVbsDb( "RequestVBSdbEditWhere" )
end function

'***********************************************
sub VbsDbEditDbUpdateActually( objVbsDb )
'***********************************************
	dim strSql , strErrorMessage
	strSql = funStrSqlVbsDbEditDbUpdate( objVbsDb )
  on error resume next
	objVbsDb( "Connection" ).execute strSql
  if err.number<>0 then
    vbsDbEditErrorMessageForUpdate strSql , objVbsDb
  end if
end sub

'***********************************************
sub VbsDbEditDbUpdate( objVbsDb )
'***********************************************
  'if funBolUpdateIsGranted( objVbsDb ) then
    VbsDbEditDbUpdateActually objVbsDb
  'end if
end sub

'***********************************************
function funStrSqlVbsDbEditDbDelete( objVbsDb )
'***********************************************
  funStrSqlVbsDbEditDbDelete = "delete from " &_
    funStrGlobalDbTypeEditTableName( objVbsDb ) &_
    " where " & objVbsDb( "RequestVBSdbEditWhere" )
end function

'***********************************************
sub VbsDbEditDbDeleteActually( objVbsDb )
'***********************************************
  dim strSql
  strSql = funStrSqlVbsDbEditDbDelete( objVbsDb )
  on error resume next
  objVbsDb( "Connection" ).execute strSql
  if err.number<>0 then
    vbsDbEditErrorMessageForDelete strSql , objVbsDb
  end if
end sub

'***********************************************
sub VbsDbEditDbDelete( objVbsDb )
'***********************************************
  'if funBolDeleteIsGranted( objVbsDb ) then
    VbsDbEditDbDeleteActually objVbsDb
  'end if
end sub

'***********************************************
sub VbsDbEditDbActually( objVbsDb )
'***********************************************
  select case VbsDbGetLastAction( objVbsDb )
  	case "SubmitAdd"
  	  VbsDbEditDbAdd objVbsDb
  	case "SubmitUpdate"
  	  VbsDbEditDbUpdate objVbsDb
  	case "SubmitDelete"
  	  VbsDbEditDbDelete objVbsDb
  end select
end sub

'***********************************************
'***********************************************
sub VbsDbEditDb( objVbsDb )
'***********************************************
'***********************************************
  dim strFormTamperingError
  strFormTamperingError = funStrFormTamperingControlForEditDb( objVbsDb )
  if strFormTamperingError = "" then
    ' form tampering was not detected
    VbsDbEditDbActually objVbsDb
  else
    ' form tampering was detected
    drawError "Database editing is not granted. It results to be a form tampering issue. " & _
              "The specific error is:<br> " & strFormTamperingError & _
              "<br>If this is not the result for a form tampering, you can avoid it setting the EditFormTamperingControl " & _
              "property to false. " & _
              "If this is not the result for a form tampering, please report this message to " & _
              "the VBSdb support team" , objVbsDb
  end if
end sub
%>
<%
'***********************************************
'***********************************************
function funNumMax( i1 , i2 )
'***********************************************
'***********************************************
  if i1 > i2 then
    funNumMax = i1
  else
    funNumMax = i2
  end if
end function

'***********************************************
'***********************************************
function funNumMin( i1 , i2 )
'***********************************************
'***********************************************
  if i1 < i2 then
    funNumMin = i1
  else
    funNumMin = i2
  end if
end function

'***********************************************
'***********************************************
function cutFinal( strToBeCut , strSubStrToCut )
'***********************************************
'***********************************************
	if uCase( right( strToBeCut , len( strSubStrToCut ) ) ) = uCase( strSubStrToCut ) then
		cutFinal = left( strToBeCut , len( strToBeCut ) - len( strSubStrToCut ) )
	else
		cutFinal = strToBeCut
	end if
end function 

'***********************************************
function funStrFieldDefaultValueMayBeEmptyForUpdateOrDeleteNotNullForNonEditUpdateFieldDefaults( strFieldName , objVbsDb )
'***********************************************
  if funBolAdoxEditFiledIsDate( strFieldName , objVbsDb ) then
    funStrFieldDefaultValueMayBeEmptyForUpdateOrDeleteNotNullForNonEditUpdateFieldDefaults = _
      funStrInputControlDateFormat( strFieldName , objVbsDb )
      'cStr( formatDateTime( objVbsDb( "EditTableRecordSet" ).fields( strFieldName ).value , vbsDbVbLongDate ) )
  else
    funStrFieldDefaultValueMayBeEmptyForUpdateOrDeleteNotNullForNonEditUpdateFieldDefaults = _
      cStr( objVbsDb( "EditTableRecordSet" ).fields( strFieldName ).value )
  end if
end function

'***********************************************
function funStrFieldDefaultValueMayBeEmptyForUpdateOrDeleteForNull( strFieldName , objVbsDb )
'***********************************************
  if ( VbsDbGetScreenType( objVbsDb ) = "Update" ) and ( objVbsDb( "DictionaryEditUpdateFieldDefaults" ).exists( strFieldName ) ) then
    ' current screen is an update screen and strFieldName is defined in the EditUpdateFieldDefaults property
    funStrFieldDefaultValueMayBeEmptyForUpdateOrDeleteForNull = _
      objVbsDb( "DictionaryEditUpdateFieldDefaults" )( strFieldName )
  else
    funStrFieldDefaultValueMayBeEmptyForUpdateOrDeleteForNull = ""
  end if
end function

'***********************************************
function funStrFieldDefaultValueMayBeEmptyForUpdateOrDeleteNotNull( strFieldName , objVbsDb )
'***********************************************
  if ( VbsDbGetScreenType( objVbsDb ) = "Update" ) and ( objVbsDb( "DictionaryEditUpdateFieldDefaults" ).exists( strFieldName ) ) then
    ' current screen is an update screen and strFieldName is defined in the EditUpdateFieldDefaults property
    funStrFieldDefaultValueMayBeEmptyForUpdateOrDeleteNotNull = _
      objVbsDb( "DictionaryEditUpdateFieldDefaults" )( strFieldName )
  else
    funStrFieldDefaultValueMayBeEmptyForUpdateOrDeleteNotNull = _
      funStrFieldDefaultValueMayBeEmptyForUpdateOrDeleteNotNullForNonEditUpdateFieldDefaults( strFieldName , objVbsDb )
  end if
end function

'***********************************************
function funFieldDefaultValueMayBeEmpty( strFieldName , objVbsDb )
'***********************************************
  select case VbsDbGetScreenType( objVbsDb )
  	case "Search"
  	  funFieldDefaultValueMayBeEmpty = ""
  	case "Add"
  	  funFieldDefaultValueMayBeEmpty = objVbsDb( "DictionaryEditAddFieldDefaults" )( strFieldName )
  	case "Delete","Update"
  	  if isNull( objVbsDb( "EditTableRecordSet" ).fields( strFieldName ).value ) then
  	    ' current field value (in edit table) is null
  	    funFieldDefaultValueMayBeEmpty = ""
  	    funFieldDefaultValueMayBeEmpty = funStrFieldDefaultValueMayBeEmptyForUpdateOrDeleteForNull( strFieldName , objVbsDb )
  	  else
  	    ' current field value (in edit table) is not null
  	    funFieldDefaultValueMayBeEmpty = funStrFieldDefaultValueMayBeEmptyForUpdateOrDeleteNotNull( strFieldName , objVbsDb )
  	  end if
  end select
  if isNull( funFieldDefaultValueMayBeEmpty ) then
    funFieldDefaultValueMayBeEmpty = ""
  end if
  funFieldDefaultValueMayBeEmpty = server.HTMLEncode( funFieldDefaultValueMayBeEmpty )
end function

'***********************************************
'***********************************************
function funFieldDefaultValue( strFieldName , objVbsDb )
'***********************************************
'***********************************************
  funFieldDefaultValue = funFieldDefaultValueMayBeEmpty( strFieldName , objVbsDb )
  if funBolInputFieldIsReadOnly( strFieldName , objVbsDb ) and _
    ( funFieldDefaultValue = "" ) then
    ' current field default value is empty or null, and it is a read only field
    funFieldDefaultValue = "&nbsp;"     ' avoids the bad layout for empty input read only fields
  end if
end function

'***********************************************
'***********************************************
function funArrStrSplit( strStringToSplit , strDelimiters )
'***********************************************
'***********************************************
  dim intStartPosition , intNextFoundPosition , intArrayUpperBound , arrStrSplit()
  intStartPosition = 1
  intArrayUpperBound = -1
  intNextFoundPosition = inStr( intStartPosition , strStringToSplit , strDelimiters )
  while intNextFoundPosition > 0
    intArrayUpperBound = intArrayUpperBound + 1
    redim preserve arrStrSplit( intArrayUpperBound )
    arrStrSplit( uBound( arrStrSplit ) ) = mid( strStringToSplit , intStartPosition , intNextFoundPosition - intStartPosition )
    intStartPosition = intNextFoundPosition + len( strDelimiters )
    intNextFoundPosition = inStr( intStartPosition , strStringToSplit , strDelimiters )
  wend
  redim preserve arrStrSplit( intArrayUpperBound + 1 )
  arrStrSplit( uBound( arrStrSplit ) ) = mid( strStringToSplit , intStartPosition )
  funArrStrSplit = arrStrSplit
end function

'***********************************************
'***********************************************
function funStrFieldWhereClauseForFieldOfTypeAdGUID( strFieldName , strFieldValue , objVbsDb )
'***********************************************
'***********************************************
  dim strTransformedFieldValue
  strTransformedFieldValue = replace( strFieldValue , "{" , "'" )
  strTransformedFieldValue = replace( strTransformedFieldValue , "}" , "'" )
  funStrFieldWhereClauseForFieldOfTypeAdGUID = _
    funStrGlobalDbTypeFieldName( strFieldName , objVbsDb ) & _
      funStrVbsDbSearchOperator( strFieldName , objVbsDb ) & _
      strTransformedFieldValue
end function

'***********************************************
function funStrEditFieldWhereClauseActually( intFieldType , strFieldName , strFieldValue , objVbsDb )
'***********************************************
  select case intFieldType
  	case vbsDbAdGUID
  	  ' SqlServer unique identifier
      funStrEditFieldWhereClauseActually = funStrFieldWhereClauseForFieldOfTypeAdGUID( strFieldName , strFieldValue , objVbsDb )
  	case vbsDbAdSmallInt, vbsDbAdInteger, vbsDbAdDecimal, vbsDbAdSingle, vbsDbAdDouble, vbsDbAdCurrency, _
  	  vbsDbAdTinyInt, vbsDbAdUnsignedTinyInt, vbsDbAdUnsignedSmallInt, vbsDbAdUnsignedInt , vbsDbAdBigInt, vbsDbAdUnsignedBigInt, _
  	  vbsDbAdNumeric, vbsDbAdVarNumeric
  	  ' numeric
      funStrEditFieldWhereClauseActually = funStrGlobalDbTypeFieldName( strFieldName , objVbsDb ) & "=" & cStr( strFieldValue )
  	case vbsDbAdBoolean	
  	  ' boolean
       funStrEditFieldWhereClauseActually = funStrGlobalDbTypeFieldName( strFieldName , objVbsDb ) & "=" & funStrGlobalDbTypeBooleanFieldFormattedValue( strFieldValue , objVbsDb )
 	  case vbsDbAdChar, vbsDbAdWChar, vbsDbAdVarChar, vbsDbAdLongVarChar, vbsDbAdVarWChar, vbsDbAdLongVarWChar, vbsDbAdVarBinary
  	  ' strings
  	  funStrEditFieldWhereClauseActually = funStrGlobalDbTypeFieldName( strFieldName , objVbsDb ) & "='" & _
  	    replace( strFieldValue , "'" , "''" ) & "'"
  	case vbsDbAdDate, vbsDbAdDBDate, vbsDbAdDBTime, vbsDbAdDBTimeStamp
  	  ' dates/time
  	  funStrEditFieldWhereClauseActually = funStrGlobalDbTypeFieldName( strFieldName , objVbsDb ) & "=" & funStrGlobalDbTypeDateFieldFormattedValue( strFieldValue , objVbsDb )
  	case else
  	  drawError "Function funStrFieldWhereClause: type " &_
  	    intFieldType &_
  	    " is not covered. Please, signal this bug message to <a href='mailto:support@vbsdb.com'>support@vbsdb.com</a>." , objVbsDb
  end select
end function

'***********************************************
function funStrEditFieldWhereClause( intFieldType , strFieldName , strFieldValue , objVbsDb )
'***********************************************
  if isNull( strFieldValue ) then
    funStrEditFieldWhereClause = funStrGlobalDbTypeFieldName( strFieldName , objVbsDb ) & " is null"
  else
    funStrEditFieldWhereClause = funStrEditFieldWhereClauseActually( intFieldType , strFieldName , strFieldValue , objVbsDb )
  end if
end function

'***********************************************
'***********************************************
function funStrVbsDbEditWhere( objVbsDb )
'***********************************************
'***********************************************
  dim strFieldName , strVbsDbEditWhere
  strVbsDbEditWhere = ""
  for each strFieldName in objVbsDb( "DictionaryEditKeyFields" )
    'strVbsDbEditWhere = strVbsDbEditWhere & _
    '  "(" & strFieldName  & "=" &_
    '  funStrVbsDbEditWhereFieldValue( strFieldName , objVbsDb )  & ") and"
    strVbsDbEditWhere = strVbsDbEditWhere & _
      "(" & funStrEditFieldWhereClause( funVarSqlRecordSetFieldType( strFieldName , objVbsDb ) , _
        strFieldName , funStrVbsDbEditWhereFieldValue( strFieldName , objVbsDb ) , objVbsDb ) & _
      ") and "
  next
  funStrVbsDbEditWhere = cutFinal( strVbsDbEditWhere , " and " )
end function

'***********************************************
'***********************************************
function funBolJustArrived( objVbsDb )
'***********************************************
'***********************************************
  funBolJustArrived = ( _
    ( ( inStr( replace( cStr( request.serverVariables( "HTTP_REFERER" ) ) , "%20" , " " ) , _
      cStr( request.serverVariables( "URL" ) ) ) = 0 ) and _
      ( VbsDbGetLastAction( objVbsDb ) = "NonVbsDb" ) ) or _
      objVbsDb( "GlobalReset" ) )
end function

'***********************************************
'***********************************************
function funStrConvertNullToEmptyString( varToBeConverted )
'***********************************************
'***********************************************
  if isNull( varToBeConverted ) then
    funStrConvertNullToEmptyString = ""
  elseif varType( varToBeConverted ) = constVarTypeBoolean then
    if varToBeConverted then
      funStrConvertNullToEmptyString = "true"
    else
      funStrConvertNullToEmptyString = "false"
    end if
  else
    funStrConvertNullToEmptyString = cStr( varToBeConverted )
  end if
end function

'***********************************************
'***********************************************
function funBolIsOdd( intVar )
'***********************************************
'***********************************************
  funBolIsOdd = ( ( intVar mod 2 ) = 1 )
end function

'***********************************************
'***********************************************
function vbsDbGetClickClass( objVbsDb )
'***********************************************
'***********************************************
  vbsDbGetClickClass = cStr( request( "VBSdbClickClass_" & objVbsDb( "GlobalId" ) ) )
end function

'***********************************************
function funIntSqlFromBegin( objVbsDb )
'***********************************************
  funIntSqlFromBegin = inStr( 1 , objVbsDb( "Sql" ) , " FROM " , 1 )
end function

'***********************************************
function funIntSqlClauseBegin( strClauseReservedName , objVbsDb )
'***********************************************
  dim intFromBegin , intClauseBegin
  intFromBegin = funIntSqlFromBegin( objVbsDb )
  intClauseBegin = inStr( intFromBegin , objVbsDb( "Sql" ) , strClauseReservedName , 1 )
  if intClauseBegin < intFromBegin then
    ' the sql clause is not present
    funIntSqlClauseBegin = len( objVbsDb( "Sql" ) ) + 1
  else
    ' the sql clause is present
    funIntSqlClauseBegin = intClauseBegin
  end if
end function

'***********************************************
function funIntSqlForBrowseBegin( strSql , objVbsDb )
'***********************************************
  funIntSqlForBrowseBegin = funIntSqlClauseBegin( " FOR " , objVbsDb )
end function

'***********************************************
function funIntSqlComputeBegin( strSql , objVbsDb )
'***********************************************
  funIntSqlComputeBegin = funIntSqlClauseBegin( " COMPUTE " , objVbsDb )
end function

'***********************************************
function funIntSqlOrderByBegin( strSql , objVbsDb )
'***********************************************
  funIntSqlOrderByBegin = funIntSqlClauseBegin( " ORDER " , objVbsDb )
end function

'***********************************************
function funIntSqlHavingBegin( strSql , objVbsDb )
'***********************************************
  funIntSqlHavingBegin = funIntSqlClauseBegin( " HAVING " , objVbsDb )
end function

'***********************************************
function funIntSqlGroupByBegin( strSql , objVbsDb )
'***********************************************
  funIntSqlGroupByBegin = funIntSqlClauseBegin( " GROUP " , objVbsDb )
end function

'***********************************************
function funIntSqlFinalSemicolon( strSql , objVbsDb )
'***********************************************
  if right( objVbsDb( "Sql" ) , 1 ) = ";" then
    ' the last Sql character is a semicolon
    funIntSqlFinalSemicolon = len( objVbsDb( "Sql" ) )
  else
    ' the last Sql character is not a semicolon
    funIntSqlFinalSemicolon = len( objVbsDb( "Sql" ) ) + 1
  end if
end function

'***********************************************
function funIntSqlWhereClauseEnd( strSql , objVbsDb )
'***********************************************
  dim intFromBegin
  funIntSqlWhereClauseEnd = funIntSqlForBrowseBegin( strSql , objVbsDb ) - 1
  funIntSqlWhereClauseEnd = funNumMin( funIntSqlWhereClauseEnd , funIntSqlComputeBegin( strSql , objVbsDb ) - 1 )
  funIntSqlWhereClauseEnd = funNumMin( funIntSqlWhereClauseEnd , funIntSqlOrderByBegin( strSql , objVbsDb ) - 1 )
  funIntSqlWhereClauseEnd = funNumMin( funIntSqlWhereClauseEnd , funIntSqlHavingBegin( strSql , objVbsDb ) - 1 )
  funIntSqlWhereClauseEnd = funNumMin( funIntSqlWhereClauseEnd , funIntSqlGroupByBegin( strSql , objVbsDb ) - 1 )
  funIntSqlWhereClauseEnd = funNumMin( funIntSqlWhereClauseEnd , funIntSqlFinalSemicolon( strSql , objVbsDb ) - 1 )
end function

'***********************************************
function funStrSqlInsertWhereClauseBegin( strSql , strWhereClauseToBeInserted , objVbsDb )
'***********************************************
  funStrSqlInsertWhereClauseBegin = _
    left( strSql , funIntSqlWhereClauseEnd( strSql , objVbsDb ) )
end function

'***********************************************
function funBolSqlWhereClauseIsAlreadyInSql( strSql , objVbsDb )
'***********************************************
  dim intFromBegin , intClauseBegin
  intFromBegin = funIntSqlFromBegin( objVbsDb )
  intClauseBegin = inStr( intFromBegin , objVbsDb( "Sql" ) , " WHERE " , 1 )
  funBolSqlWhereClauseIsAlreadyInSql = ( intFromBegin < intClauseBegin )
end function

'***********************************************
function funStrSqlInsertWhereClauseBody( strSql , strWhereClauseToBeInserted , objVbsDb )
'***********************************************
  if strWhereClauseToBeInserted = "" then
    funStrSqlInsertWhereClauseBody = ""
  else
    if funBolSqlWhereClauseIsAlreadyInSql( strSql , objVbsDb ) then
      ' the Sql property already contains a where clause
      funStrSqlInsertWhereClauseBody = " and " & strWhereClauseToBeInserted
    else
      ' the Sql property doesn't contain any where clause
      funStrSqlInsertWhereClauseBody = " where " & strWhereClauseToBeInserted
    end if
  end if
end function

'***********************************************
function funStrSqlInsertWhereClauseEnd( strSql , strWhereClauseToBeInserted , objVbsDb )
'***********************************************
  funStrSqlInsertWhereClauseEnd = _
    mid( objVbsDb( "Sql" ) , funIntSqlWhereClauseEnd( strSql , objVbsDb ) + 1 )
end function

'***********************************************
'***********************************************
function funStrSqlInsertWhereClause( strSql , strWhereClauseToBeInserted , objVbsDb )
'***********************************************
'***********************************************
  funStrSqlInsertWhereClause = _
    funStrSqlInsertWhereClauseBegin( strSql , strWhereClauseToBeInserted , objVbsDb ) & _
    funStrSqlInsertWhereClauseBody( strSql , strWhereClauseToBeInserted , objVbsDb ) & _
    funStrSqlInsertWhereClauseEnd( strSql , strWhereClauseToBeInserted , objVbsDb )
end function
%>
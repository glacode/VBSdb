<%
'***********************************************
'***********************************************
function funBolAdoxFieldIsNullable( strColumnName , objVbsDb )
'***********************************************
'***********************************************
  dim intAttributes
  intAttributes = objVbsDb( "AdoxCatalog" ).tables( objVbsDb( "EditTableName" ) ).columns( strColumnName ).attributes
  funBolAdoxFieldIsNullable = ( ( intAttributes = 2 ) or ( intAttributes = 3 ) ) ' check adox documentation for this issue
end function

'***********************************************
'***********************************************
function funLngAdoxFieldDefinedSize( strColumnName , objVbsDb )
'***********************************************
'***********************************************
  funLngAdoxFieldDefinedSize = objVbsDb( "AdoxCatalog" ).tables( objVbsDb( "EditTableName" ) ).columns( strColumnName ).definedSize
end function

'***********************************************
function funBolAdoxFieldAutoincrementNoAdoxSupportForAccess( objColumn , objVbsDb )
'***********************************************
  funBolAdoxFieldAutoincrementNoAdoxSupportForAccess = ( ( objColumn.attributes = 1 ) and ( objColumn.type = 3 ) )  ' check the ADO documentation: it assumes that, if a field is numeric and not nullable, then it is an autonumber field; it doesn't work for SqlServer 2000; if it doesn't work, use EditAutoincrement
end function

'***********************************************
function funBolAdoxFieldAutoincrementActuallyForAccessAndSqlServerNoAdoxSupport( objColumn , objVbsDb )
'***********************************************  
  if ( uCase( objVbsDb( "GlobalDbType" ) ) = "ACCESS" ) then
    funBolAdoxFieldAutoincrementActuallyForAccessAndSqlServerNoAdoxSupport = _
      funBolAdoxFieldAutoincrementNoAdoxSupportForAccess( objColumn , objVbsDb )
  else
    ' server type is SqlServer
    funBolAdoxFieldAutoincrementActuallyForAccessAndSqlServerNoAdoxSupport = _
      uCase( objColumn.name ) = uCase( objVbsDb( "EditSlqServerAutoIncrement" ) )
  end if
end function

'***********************************************
function funBolAdoxFieldAutoincrementActuallyForAccessAndSqlServer( objColumn , objVbsDb )
'***********************************************
  on error resume next
  funBolAdoxFieldAutoincrementActuallyForAccessAndSqlServer = _
    objColumn.properties( "AutoIncrement" )
  if err.number <> 0 then
    ' the provider doesn't support the AutoIncrement property
    funBolAdoxFieldAutoincrementActuallyForAccessAndSqlServer = _
      funBolAdoxFieldAutoincrementActuallyForAccessAndSqlServerNoAdoxSupport( objColumn , objVbsDb )
  end if
end function

'***********************************************
function funBolAdoxFieldAutoincrementActually( objColumn , objVbsDb )
'***********************************************  
  if ( uCase( objVbsDb( "GlobalDbType" ) ) = "ACCESS" ) or ( uCase( objVbsDb( "GlobalDbType" ) ) = "SQLSERVER" ) then
    funBolAdoxFieldAutoincrementActually = _
      funBolAdoxFieldAutoincrementActuallyForAccessAndSqlServer( objColumn , objVbsDb )
  else
    funBolAdoxFieldAutoincrementActually = false
  end if
end function

'***********************************************
'***********************************************
function funBolAdoxFieldAutoincrement( objColumn , objVbsDb )
'***********************************************
'***********************************************
  if objVbsDb( "EditAutoincrement" ) <> constAutomaticSearchForAutoincrement then
    ' the developer has explicitly identified the autonumber field
    funBolAdoxFieldAutoincrement = _
      ( ucase( objColumn.name ) = uCase( trim( objVbsDb( "EditAutoincrement" ) ) ) )
  else
    ' the developer didn't explicitly identify a field as an autoincrement field
    funBolAdoxFieldAutoincrement = _
      ( ( objColumn.type = 72 ) and uCase( objVbsDb( "GlobalDbType" ) = "SqlServer" ) ) or _
      funBolAdoxFieldAutoincrementActually( objColumn , objVbsDb )
  end if
end function

'***********************************************
'***********************************************
function funBolAdoxEditFieldIsOleObjectType( objColumn , objVbsDb )
'***********************************************
'***********************************************
  funBolAdoxEditFieldIsOleObjectType = ( ( objColumn.type = vbsDbAdVarBinary ) or ( objColumn.type = vbsDbAdLongVarBinary ) )
end function

'***********************************************
'***********************************************
function funBolAdoxColumnIsBoolean( strColumnName , objVbsDb )
'***********************************************
'***********************************************
  funBolAdoxColumnIsBoolean = _
    ( objVbsDb( "AdoxCatalog" ).tables( objVbsDb( "EditTableName" ) ).columns( strColumnName ).type = _
    vbsDbAdBoolean )
end function

'***********************************************
'***********************************************
function funBolAdoxIsEditField( strFieldName , objVbsDb )
'***********************************************
'***********************************************
  on error resume next
  isObject( objVbsDb( "AdoxCatalog" ).tables( objVbsDb( "EditTableName" ) ).columns( strFieldName ) )
  funBolAdoxIsEditField = ( err.number <= 0 )
end function

'***********************************************
'***********************************************
function funIntAdoxFieldType( strColumnName , objVbsDb )
'***********************************************
'***********************************************
  funIntAdoxFieldType = objVbsDb( "AdoxCatalog" ).tables( objVbsDb( "EditTableName" ) ).columns( strColumnName ).type
  funIntAdoxFieldType = funIntCheckAdoxFieldTypeForOracleDatabase( funIntAdoxFieldType , objVbsDb )
end function

'***********************************************
'***********************************************
function funBolAdoxEditFieldIsString( strFieldName , objVbsDb )
'***********************************************
'***********************************************
  dim intFieldType
  intFieldType = funIntAdoxFieldType( strFieldName , objVbsDb )
  funBolAdoxEditFieldIsString = _
    ( intFieldType = vbsDbAdChar ) or _
    ( intFieldType = vbsDbAdVarChar ) or _
    ( intFieldType = vbsDbAdLongVarChar ) or _
    ( intFieldType = vbsDbAdWChar ) or _
    ( intFieldType = vbsDbAdVarWChar ) or _
    ( intFieldType = vbsDbAdLongVarWChar )
end function

'***********************************************
'***********************************************
function funBolAdoxEditFiledIsDate( strFieldName , objVbsDb )
'***********************************************
'***********************************************
  dim intFieldType
  intFieldType = funIntAdoxFieldType( strFieldName , objVbsDb )
  funBolAdoxEditFiledIsDate = _
    ( intFieldType = vbsDbAdDBDate ) or _
    ( intFieldType = vbsDbAdDBTime ) or _
    ( intFieldType = vbsDbAdDBTimeStamp )
end function

'***********************************************
'***********************************************
function funObjAdoxEditTableColumns( byRef objVbsDb )
'***********************************************
'***********************************************
  on error resume next
  set funObjAdoxEditTableColumns = objVbsDb( "AdoxCatalog" ).tables( objVbsDb( "EditTableName" ) ).columns
  if err.number<>0 then
    drawErrorWithDebugging "EditTableName assignment error. Check your EditTableName property assignment." , objVbsDb
  end if
end function

'***********************************************
'***********************************************
sub vbsDbSetAdoxCatalog( byRef objVbsDb )
'***********************************************
'***********************************************
  set objVbsDb( "AdoxCatalog" ) = server.createObject( "ADOX.Catalog" )
  objVbsDb( "AdoxCatalog" ).activeConnection = objVbsDb( "Connection" )
end sub
%>
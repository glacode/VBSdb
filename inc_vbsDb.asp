<%
' session( "VbsDbSessionFormAbsolutePosition_" & objVbsDb( "GlobalId" ) ) denotes the grid position of the record to be shown in the form
' session( "VbsDbSessionOrderBy_" & objVbsDb( "GlobalId" ) )	denotes the ORDER BY clause if the user asked to order the grid by a certain column
' session( "VbsDbFilterWhereClause_" & objVbsDb( "GlobalId" ) ) denotes the additional WHERE clause, if the user asked to apply a filter to the view mode

'***********************************************
'***********************************************
sub VbsDbNew( byRef objVbsDb )
'***********************************************
'***********************************************
  set objVbsDb = server.createobject( "scripting.dictionary" )
  objVbsDb.compareMode = 1

  ' public input properties
  objVbsDb( "MdbPath" ) = ""
  objVbsDb( "Dsn" ) = ""
  objVbsDb( "Sql" ) = ""
  objVbsDb( "ViewMode" ) = "GRID-FORM"		' accepted values are "GRID", "FORM", "GRID-FORM", "GRID|FORM"
  objVbsDb( "ViewNavigationButtons" ) = "FIRST;NEXT;PREV;LAST;REMOVEFILTER;SEARCH"   ' possible values are first;next;prev;last;removeFilter;search;add;update;delete
  objVbsDb( "ViewNavigationPosition" ) = "BOTTOM"       ' possible values are TOP,BOTTOM,BOTH,NONE
  objVbsDb( "ViewNavigationBGColor" ) = constDefaultViewNavigationBGColor
  objVbsDb( "ViewNavigationDisabledFGColor" ) = constDefaultViewNavigationDisabledFGColor
  objVbsDb( "ViewNavigationFGColor" ) = constDefaultViewNavigationFGColor
  objVbsDb( "ViewPosition" ) = false
  objVbsDb( "ViewFieldFormats" ) = ""	    ' "fieldName|fieldFormatMacro; repeat..."
  'objVbsDb( "ViewEmptySqlMessage" ) = "No record found. Press 'Add New' to start to populate the table."
  'objVbsDb( "ViewEmptyFilterMessage" ) = "No record found. Press 'Search' to change the filter or to reset it."
  objVbsDb( "GridPageSize" ) = 5
  objVbsDb( "GridFields" ) = ""
  objVbsDb( "GridHideFields" ) = ""
  objVbsDb( "GridShowIndex" ) = true          ' if it is true, the grid shows an additional column indexing each record with an integer number
  objVbsDb( "GridHorizontalStripeBGColor" ) = constDefaultGridHorizontalStripeBGColor
  objVbsDb( "GridUnselectedIndexFGColor" ) = constDefaultGridUnselectedIndexFGColor
  objVbsDb( "GridSelectedIndexBGColor" ) = constDefaultGridSelectedIndexBGColor
  objVbsDb( "GridSelectedIndexFGColor" ) = constDefaultGridSelectedIndexFGColor
  objVbsDb( "GridTableTag" ) = "cellSpacing=1 cellPadding=2 border=0"
  objVbsDb( "GridUpdateButtons" ) = false
  objVbsDb( "GridDeleteButtons" ) = false
  objVbsDb( "GridTemplate" ) = ""
  objVbsDb( "FormFields" ) = ""
  objVbsDb( "FormHideFields" ) = ""           ' not yet implemented
  objVbsDb( "FormTableTag" ) = "cellSpacing=1 cellPadding=2 border=0"
  objVbsDb( "FormTemplate" ) = ""
  objVbsDb( "FormExactAfterEdit" ) = true
  objVbsDb( "InputEnumeratedFields" ) = ""
  objVbsDb( "InputSelectFields" ) = ""
  objVbsDb( "InputValidateDateFields" ) = ""
  objVbsDb( "SearchFields" ) = ""
  objVbsDb( "SearchHideFields" ) = ""
  objVbsDb( "SearchAliasFields" ) = ""
  objVbsDb( "SearchOperators" ) = ""        ' "fieldName|searchOperator; repeat..."
  objVbsDb( "SearchScreenRepeat" ) = false  ' true iff the screen must remeain of search type
  objVbsDb( "EditMemoFields" ) = ""               ' "fieldName|editModeNumRows|editModeNumCols|editModeMaxChars; repeat..."
  objVbsDb( "EditTableName" ) = ""
  objVbsDb( "EditKeyFields" ) = ""
  objVbsDb( "EditFields" ) = ""               ' "fieldName; repeat..."
  objVbsDb( "EditReadOnlyFields" ) = ""       ' "fieldName; repeat..."
  objVbsDb( "EditAutoincrement" ) = constAutomaticSearchForAutoincrement ' either "" or "fieldName"
  objVbsDb( "EditAddFieldDefaults" ) = ""     ' "fieldName|fieldDefaultValueInAddScreen; repeat..."
  objVbsDb( "EditUpdateFieldDefaults" ) = ""  ' "fieldName|fieldDefaultValueInUpdateScreen; repeat..."
  objVbsDb( "EditAddScreenRepeat" ) = false   ' true iff the screen must remeain of add type
  objVbsDb( "EditAddSingle" ) = false         ' true iff the user wants to add a single record to a table
  objVbsDb( "EditHideFields" ) = ""         
  objVbsDb( "EditValidateRegExp" ) = ""     ' "fieldName|strRegularExpression|strMsgBoxMessage; repeat..."
  objVbsDb( "EditValidateRequiredFields" ) = ""   ' "fieldName|strMsgBoxMessage; repeat..."
  objVbsDb( "EditErrorMessages" ) = "default"     ' possible values are default, custom and none
  objVbsDb( "EditFormTamperingControl" ) = true   ' true iff the site developer wants to check possible visitor's form tampering
  objVbsDb( "GlobalBooleanFields" ) = ""
  objVbsDb( "GlobalFieldHeaders" ) = ""
  objVbsDb( "GlobalTableBGColor" ) = constDefaultGlobalTableBGColor
  objVbsDb( "GlobalTableFGColor" ) = constDefaultGlobalTableFGColor
  objVbsDb( "GlobalHeaderBGColor" ) = constDefaultGlobalHeaderBGColor
  objVbsDb( "GlobalHeaderFGColor" ) = constDefaultGlobalHeaderFGColor
  objVbsDb( "GlobalTrueText" ) = "Yes"
  objVbsDb( "GlobalFalseText" ) = "No"
  objVbsDb( "GlobalImageDir" ) = ""
  objVbsDb( "GlobalStartScreen" ) = "VIEW"  ' accepted values are "VIEW","SEARCH","ADD"
  objVbsDb( "GlobalId" ) = 0
  objVbsDb( "GlobalDbType" ) = "ACCESS"     ' accepted values are ACCESS,SQLSERVER,FOXPRO,ORACLE,OTHER
  objVbsDb( "GlobalLanguage" ) = "english"
  objVbsDb( "GlobalOleDbProvider" ) = ""    ' accepted values are "JET.4.0","JET.3.51","ODBC"
  objVbsDb( "GlobalPagePositioning" ) = "Top"   ' accepted values are "Top", "VBSdb"
  objVbsDb( "GlobalQuerystringPreserve" ) = false
  objVbsDb( "GlobalReset" ) = false
  objVbsDb( "GlobalCSS" ) = ""
  objVbsDb( "GlobalSeparators" ) = ""       ' "propertyName1 primarySeparator1 secondarySeparator1 propertyName2 primarySeparator2 secondarySeparator2 ... repeat ..." 
  
  vbsDbSetVbsDbGetDateFormatRelatedProperties objVbsDb
end sub

'***********************************************
sub vbsDbInitializeOutputPropertiesAndPrivateProperties( byRef objVbsDb )
'***********************************************
  ' private properties
  objVbsDb( "Connection" ) = ""	            ' contiene la connessione al database
  objVbsDb( "AdoxCatalog" ) = ""	        ' connection catalog
  objVbsDb( "SqlRecordSet" ) = ""	        ' se in view mode, contiene il record set che deve essere visualizzato
  objVbsDb( "FormAbsolutePosition" )	= 1 ' posizione del record corrente nel record set, visualizzato nel form (se e' il primo record, vale 1)
  objVbsDb( "GridAbsolutePage" )	= 1     ' numero di pagina visualizzata
  objVbsDb( "GridTemplateRowTemplate" )	= ""
  objVbsDb( "FormTemplateContent" )	= ""
  objVbsDb( "OrderBy" ) = "" 
  objVbsDb( "FilterWhereClause" ) = "" 
  objVbsDb( "SqlWithFilterWhereClause" ) = "" 
  objVbsDb( "ScreenTitle" ) = ""
  objVbsDb( "InputFormLeftHeader" ) = ""
  objVbsDb( "InputFormRightHeader" ) = ""
  objVbsDb( "InputCancelContent" ) = ""
  objVbsDb( "InputCancelTitle" ) = ""
  objVbsDb( "InputSubmitValue" ) = ""
  objVbsDb( "InputResetValue" ) = ""
  objVbsDb( "InputTemplateContent" ) = ""
  objVbsDb( "EditTableRecordSet" ) = ""               ' used to find the type of the fields updated; se in add mode o in view mode, contiene il primo record della tabella; se in update o in delete mode, contiene il record che deve essere modificato o cancellato; se in view mode e si viene da una schermata di edit, viene utilizzato per determinare i tipi dei campi da modificare
  objVbsDb( "EditSlqServerAutoIncrement" ) = ""
  objVbsDb( "EditErrorNumber" ) = 0
  objVbsDb( "EditErrorDescription" ) = ""
  'objVbsDb( "EditValidate" ) = false                 ' true iff there is at least an edit field to be validated
  objVbsDb( "EditMemoFieldMaxCharExists" ) = false    ' true iff there is at least an edit memo field with a maximum number of chars to be edited
  objVbsDb( "ViewNavigationFirst" ) = true
  objVbsDb( "ViewNavigationPrev" ) = true
  objVbsDb( "ViewNavigationNext" ) = true
  objVbsDb( "ViewNavigationLast" ) = true
  objVbsDb( "ViewNavigationSearch" ) = true
  objVbsDb( "ViewNavigationAdd" ) = false
  objVbsDb( "ViewNavigationUpdate" ) = false
  objVbsDb( "ViewNavigationDelete" ) = false
  objVbsDb( "RequestVBSdbClickClass" ) = ""						' possible values are VBSdbGoToGridRow, VBSdbGridFirst, VBSdbGridPrev, VBSdbGridNext, VBSdbGridLast, VBSdbEditUpdate, VBSdbApplyFilter, VBSdbApplyAdd, VBSdbApplyUpdate, VBSdbApplyDelete
  objVbsDb( "RequestVBSGridIndex" ) = ""					' number of the record to jump to in the record set (0 if it is the first record in the record set)
  objVbsDb( "RequestVBSdbGridSort" ) = ""				' if not empty, it contains the field to use to sort the grid
  objVbsDb( "RequestVBSdbEditWhere" ) = ""			' if not empty, it contains the where clause to get the record to be edited
  
  objVbsDb( "DictionaryGlobalFieldHeaders" ) = ""			            
  objVbsDb( "DictionaryGlobalBooleanFields" ) = ""			            
  objVbsDb( "DictionarySearchFields" ) = ""         
  objVbsDb( "DictionarySearchHideFields" ) = ""
  objVbsDb( "DictionaryEditFields" ) = ""         
  objVbsDb( "DictionaryEditHideFields" ) = ""		            
  objVbsDb( "DictionaryEditReadOnlyFields" ) = ""			            
  objVbsDb( "DictionaryEditAddFieldDefaults" ) = ""
  objVbsDb( "DictionaryEditUpdateFieldDefaults" ) = ""
  objVbsDb( "DictionaryEditKeyFields" ) = ""
  objVbsDb( "DictionaryEditMemoFields" ) = ""          
  objVbsDb( "DictionaryEditValidateFields" ) = ""          
  objVbsDb( "DictionaryGridFields" ) = ""           ' grid fields to be shown in view screen mode			            
  objVbsDb( "DictionaryGridHideFields" ) = ""
  objVbsDb( "DictionaryFormFields" ) = ""           ' form fields to be shown in view screen mode			            
  objVbsDb( "DictionaryFormHideFields" ) = ""
  objVbsDb( "DictionaryInputFields" ) = ""			' fields to be shown in input screen mode            
  objVbsDb( "DictionaryInputSelectFields" ) = ""			            
  objVbsDb( "DictionaryInputValidateDateFields" ) = ""			            
  objVbsDb( "DictionaryViewFieldFormats" ) = ""

  objVbsDb( "DictionaryGlobalCustomText" ) = ""
  objVbsDb( "DictionaryLanguage" ) = ""
  objVbsDb( "DictionaryInputValidateDateFormatPosition" ) = ""
end sub

'***********************************************
sub vbsDbReset( byRef objVbsDb )
'***********************************************
  session( "VbsDbSessionFormAbsolutePosition_" & objVbsDb( "GlobalId" ) ) = ""
  session( "VbsDbSessionOrderBy_" & objVbsDb( "GlobalId" ) ) = ""
  session( "VbsDbFilterWhereClause_" & objVbsDb( "GlobalId" ) ) = ""
end sub

'***********************************************
sub vbsDbHandleReset( byRef objVbsDb )
'***********************************************
  if funBolJustArrived( objVbsDb ) then
  	' the last page was not the current one
  	vbsDbReset objVbsDb
  end if
end sub

'***********************************************
sub vbsDbSetPropertiesToUpcase( byRef objVbsDb )
'***********************************************
  dim strPropertyName
  for each strPropertyName in objVbsDb
  	if ( strPropertyName <> "Sql" ) and _
  	   ( strPropertyName <> "Dsn" ) and _
  	   ( strPropertyName <> "GlobalBooleanFields" ) and _
  	   ( strPropertyName <> "GlobalFieldHeaders" ) and _
  	   ( strPropertyName <> "GlobalTrueText" ) and _
  	   ( strPropertyName <> "GlobalFalseText" ) and _
  	   ( strPropertyName <> "GlobalCustomText" ) and _
  	   ( strPropertyName <> "GridFields" ) and _
  	   ( strPropertyName <> "GridHideFields" ) and _
  	   ( strPropertyName <> "FormFields" ) and _
  	   ( strPropertyName <> "FormHideFields" ) and _
  	   ( strPropertyName <> "GlobalDateFormat" ) and _
  	   ( strPropertyName <> "InputEnumeratedFields" ) and _
  	   ( strPropertyName <> "InputValidateDateFields" ) and _
  	   ( strPropertyName <> "InputSelectFields" ) and _
  	   ( strPropertyName <> "SearchFields" ) and _
  	   ( strPropertyName <> "SearchHideFields" ) and _
  	   ( strPropertyName <> "SearchAliasFields" ) and _
  	   ( strPropertyName <> "EditFields" ) and _
  	   ( strPropertyName <> "EditHideFields" ) and _
  	   ( strPropertyName <> "EditTableName" ) and _
  	   ( strPropertyName <> "EditKeyFields" ) and _
  	   ( strPropertyName <> "ViewFieldFormats" ) and _
  	   ( strPropertyName <> "EditAddFieldDefaults" ) and _
  	   ( strPropertyName <> "EditUpdateFieldDefaults" ) and _
  	   ( strPropertyName <> "EditValidateRequiredFields" ) and _
  	   ( strPropertyName <> "EditValidateRegExp" ) and _
  	   ( varType( objVbsDb( strPropertyName ) ) = vbString ) then
          ' current item is a string
          objVbsDb( strPropertyName ) = uCase( objVbsDb( strPropertyName ) )
  	end if
  next	
end sub

'***********************************************
sub vbsDbNormalizeSeparatorsForSemicolonProperty( strPropertyName , byRef objVbsDb )
'***********************************************
  ' for those properties where the developer used "," or "|" instead of the right separator, that is ";"
  objVbsDb( strPropertyName ) = replace( objVbsDb( strPropertyName ) , "," , ";" )
  objVbsDb( strPropertyName ) = replace( objVbsDb( strPropertyName ) , "|" , ";" )
end sub

'***********************************************
sub vbsDbNormalizeSeparatorsForSemicolonProperties( byRef objVbsDb )
'***********************************************
  vbsDbNormalizeSeparatorsForSemicolonProperty "GlobalBooleanFields" , objVbsDb
  vbsDbNormalizeSeparatorsForSemicolonProperty "GridFields" , objVbsDb
  vbsDbNormalizeSeparatorsForSemicolonProperty "GridHideFields" , objVbsDb
  vbsDbNormalizeSeparatorsForSemicolonProperty "FormFields" , objVbsDb
  vbsDbNormalizeSeparatorsForSemicolonProperty "FormHideFields" , objVbsDb
  vbsDbNormalizeSeparatorsForSemicolonProperty "SearchFields" , objVbsDb
  vbsDbNormalizeSeparatorsForSemicolonProperty "SearchHideFields" , objVbsDb
  vbsDbNormalizeSeparatorsForSemicolonProperty "EditFields" , objVbsDb
  vbsDbNormalizeSeparatorsForSemicolonProperty "EditKeyFields" , objVbsDb
end sub

'***********************************************
sub vbsDbNormalizeSeparatorsForProperties( byRef objVbsDb )
'***********************************************
  vbsDbNormalizeSeparatorsForSemicolonProperties objVbsDb
  vbsDbNormalizeSeparatorsForSearchOperators objVbsDb
  vbsDbNormalizeEditMemoFields objVbsDb
end sub

'***********************************************
sub vbsDbHandlePropertySintaxes( byRef objVbsDb )
'***********************************************
  vbsDbNormalizeSeparatorsForProperties objVbsDb
  vbsDbSetPropertiesToUpcase objVbsDb
end sub

'***********************************************
sub vbsDbSetOrderByWithDbRequestVBSdbGridSort( byRef objVbsDb )
'***********************************************
  if ( objVbsDb( "OrderBy" ) = _
      ( funStrGlobalDbTypeFieldName( objVbsDb( "RequestVBSdbGridSort" ) , objVbsDb ) & " ASC" ) ) then
  	' the column clicked was the one which the grid had already been ordered on (ascending), by means of an header click
  	objVbsDb( "OrderBy" ) = _
  	  funStrGlobalDbTypeFieldName( objVbsDb( "RequestVBSdbGridSort" ) , objVbsDb ) & " DESC"
  else
  	' the column clicked was not the one which the grid had already been ordered on (ascending), by means of an header click
  	objVbsDb( "OrderBy" ) = _
  	  funStrGlobalDbTypeFieldName( objVbsDb( "RequestVBSdbGridSort" ) , objVbsDb ) & " ASC"
  end if
end sub

'***********************************************
sub vbsDbSetOrderBy( byRef objVbsDb )
'***********************************************
  if objVbsDb( "RequestVBSdbGridSort" ) <> "" then
  	' the user clicked on a grid header, in order to sort the grid according to the corresponding field
  	vbsDbSetOrderByWithDbRequestVBSdbGridSort objVbsDb
  end if
end sub

'***********************************************
sub vbsDbSetEditTableRecordSet( byRef objVbsDb )
'***********************************************
  if objVbsDb( "EditTableName" ) <> "" then
    vbsDbSetEditTableRecordSetWithDbEditTableNotEmpty objVbsDb
  end if
end sub

'***********************************************
function funStrVbsDbConnection_strMapPath( objVbsDb )
'***********************************************
  dim strError
  on error resume next
  funStrVbsDbConnection_strMapPath = server.mapPath( objVbsDb( "MdbPath" ) )
  if err.number <> 0 then
    ' server.mapPath( objVbsDb ) is not a string. Bad Server.MapPath argument
    strError = "MdbPath error.<br><br>" & _
      "Your MdbPath assignment is '" & objVbsDb( "MdbPath" ) & _
      "'.<br><br>" & _
      "Returned error description is: '" & err.description & "'<br><br>" & _
      "Please, control you are not assigning a physical path to the MdbPath property: you have to use a virtual path."
    drawError strError , objVbsDb
  end if    
end function

'***********************************************
function funStrVbsDbConnection_Dsn( objVbsDb )
'***********************************************
  funStrVbsDbConnection_Dsn = trim( objVbsDb( "Dsn" ) )
end function

'***********************************************
function funStrVbsDbConnection_jet40( objVbsDb )
'***********************************************
  dim strMapPath
  strMapPath = funStrVbsDbConnection_strMapPath( objVbsDb )
  funStrVbsDbConnection_jet40 = _
    "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strMapPath
end function

'***********************************************
function funStrVbsDbConnection_odbc( objVbsDb )
'***********************************************
  dim strMapPath
  strMapPath = funStrVbsDbConnection_strMapPath( objVbsDb )
  funStrVbsDbConnection_odbc = _
    "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & strMapPath & ";"
end function

'***********************************************
function funStrVbsDbConnection_jet351( objVbsDb )
'***********************************************
  dim strMapPath
  strMapPath = funStrVbsDbConnection_strMapPath( objVbsDb )
  funStrVbsDbConnection_jet351 = _
    "Provider=Microsoft.Jet.OLEDB.3.51;Data Source=" & strMapPath
end function

'***********************************************
function funStrErrorWrongDsn( strConnection , objVbsDb )
'***********************************************
  funStrErrorWrongDsn = _
    "Dsn error. To connect to database, VBSdb is trying the following connection string:<br><br>" & _
    strConnection & "<br><br>" & _
    "Returned error description is: '" & err.description & "'<br><br>" & _
    "Please, control the data source name " & objVbsDb( "Dsn" ) & _
    " is properly configured on your system and it is properly spelled."
end function

'***********************************************
function funStrErrorWrongOleDbProvider( strVbsDbConnectionString , objVbsDb )
'***********************************************
  ' funStrVbsDbConnection( objVbsDb ) is a string.
  funStrErrorWrongOleDbProvider = _
    "MdbPath error. To connect to database, VBSdb is trying this connection string:<br><br>" & _
    strVbsDbConnectionString & "<br>" & _
    "Returned error description is: '" & err.description & "'<br><br>" & _
    "Please, perform the following debug steps:" & _
    "<div align=left>" & _
    "<ul><li>the Microsoft Access file '" & objVbsDb( "MdbPath" ) & _
    "' is properly spelled and it actually exists at the specified location</li>" & _
    "<li>if you are setting the GlobalOleDbProvider property, control the corresponding " & _
    "OleDb provider is available on your server</li>" & _
    "</ul>"
end function

'***********************************************
sub vbsDbSetConnection_tryTheBestOleDbProviderPossible( byRef strErrorMessage , byRef objVbsDb )
'***********************************************
  on error resume next
  objVbsDb( "Connection" ).open funStrVbsDbConnection_jet40( objVbsDb )
  if err.number<>0 then
    err.clear
    objVbsDb( "Connection" ).open funStrVbsDbConnection_odbc( objVbsDb )
    if err.number<>0 then
      err.clear
      objVbsDb( "Connection" ).open funStrVbsDbConnection_jet351( objVbsDb )
      if err.number<>0 then
        err.clear
        objVbsDb( "Connection" ).open funStrVbsDbConnection_jet40( objVbsDb )
        strErrorMessage = _
          funStrErrorWrongOleDbProvider( funStrVbsDbConnection_jet40( objVbsDb ) , objVbsDb )
      end if
    end if
  end if
end sub

'***********************************************
sub vbsDbSetConnection( byRef objVbsDb )
'***********************************************
  dim strErrorMessage
  strErrorMessage = ""
  set objVbsDb( "Connection" ) = Server.CreateObject( "ADODB.Connection" )
  on error resume next
  if objVbsDb( "Dsn" ) <> "" then
    objVbsDb( "Connection" ).open funStrVbsDbConnection_Dsn( objVbsDb )
    if err.number<>0 then
      strErrorMessage = funStrErrorWrongDsn( funStrVbsDbConnection_Dsn( objVbsDb ) , objVbsDb )
    end if
  elseif objVbsDb( "GlobalOleDbProvider" ) = "JET.4.0" then
    objVbsDb( "Connection" ).open funStrVbsDbConnection_jet40( objVbsDb )
    if err.number<>0 then
      strErrorMessage = _
        funStrErrorWrongOleDbProvider( funStrVbsDbConnection_jet40( objVbsDb ) , objVbsDb )
    end if
  elseif objVbsDb( "GlobalOleDbProvider" ) = "JET.3.51" then
    objVbsDb( "Connection" ).open funStrVbsDbConnection_jet351( objVbsDb )
    if err.number<>0 then
      strErrorMessage = _
        funStrErrorWrongOleDbProvider( funStrVbsDbConnection_jet351( objVbsDb ) , objVbsDb )
    end if
  elseif objVbsDb( "GlobalOleDbProvider" ) = "ODBC" then
    objVbsDb( "Connection" ).open funStrVbsDbConnection_odbc( objVbsDb )
'response.write "  1 - " & unStrVbsDbConnection_odbc( objVbsDb )
'response.end
    if err.number<>0 then
      strErrorMessage = funStrErrorWrongOleDbProvider( funStrVbsDbConnection_odbc( objVbsDb ) , objVbsDb )
    end if
  else
    ' objVbsDb( "GlobalOleDbProvider" ) = ""
    vbsDbSetConnection_tryTheBestOleDbProviderPossible strErrorMessage , objVbsDb  
  end if
  if strErrorMessage <> "" then
    drawError strErrorMessage , objVbsDb
  end if
end sub

'***********************************************
sub vbsDbSetEditTableRecordSetWithDbEditTableNotEmpty( byRef objVbsDb )
'***********************************************
  dim strSql , strErrorMessage
  if ( VbsDbGetScreenType( objVbsDb ) = "Update" ) or ( VbsDbGetScreenType( objVbsDb ) = "Delete" ) then
    ' we are in an update screen
	  strSql = "select " & funStrGlobalDbTypeSelectTop1( objVbsDb ) & "* from " & funStrGlobalDbTypeEditTableName( objVbsDb ) & " where " & objVbsDb( "RequestVBSdbEditWhere" )
    set objVbsDb( "EditTableRecordSet" ) = Server.CreateObject("ADODB.Recordset")
    on error resume next
    objVbsDb( "EditTableRecordSet" ).Open strSql , objVbsDb( "Connection" ) , vbsDbAdOpenKeySet , vbsDbAdLockReadOnly , vbsDbAdCmdText
    if err.number<>0 then
      strErrorMessage = "Error opening the Edit table. VBSdb is trying this Sql query:<br><br>" & _
        strSql & "<br><br>" & _
        "Returned error description is: '" & err.description & "'<br><br>" & _
        "Please, perform the following debug steps:" & _
        "<div align=left>" & _
        "<ul><li>control the edit table is not exclusively opened for database managment</li>" & _
        "<li>control the edit table name is correct</li>" & _
        "<li>try the Sql query displayed above directly against your database</li>" & _
        "<li>check the Sql property assignment</li>" & _
        "<li>check the SearchAliasFields property if a filter has been applied and an ambiguous field name is detected</li>" & _
        "</ul>" & _
        "</div>"
      drawError strErrorMessage , objVbsDb
    end if
  end if
end sub

'***********************************************
sub vbsDbSetDictionaryEditFieldsActually( byRef objVbsDb )
'***********************************************
  vbsDbSetDictionaryObjectFromList "EditFields" , objVbsDb
end sub

''***********************************************
'sub vbsDbHandleEditFieldsIsEmptyActually( byRef objVbsDb )
''***********************************************
' we don't use this function anymore, because ADOX with the Jet provider returns
' columns in alphabetical order, not in definition order
'  dim objColumn
'  for each objColumn in funObjAdoxEditTableColumns( objVbsDb )
'    objVbsDb( "DictionaryEditFields" ).add uCase( objColumn.name ) , objColumn.name
'  next
'end sub

'***********************************************
sub vbsDbHandleEditFieldsIsEmptyActually_createOrderedArray( byRef arrColumns , byRef objVbsDb )
'***********************************************
  dim objRs , lngNumColumns
  lngNumColumns = 0
  set objRs = Server.CreateObject("ADODB.Recordset")
  objRs.cursorLocation = vbsDbAdUseClient
  set objRs = objVbsDb( "Connection" ).OpenSchema( vbsDbAdSchemaTables, _
			 array( Empty , Empty , objVbsDb( "EditTableName" ) ) )
  while not objRs.eof
    if objRs( "Ordinal_Position" ).value > lngNumColumns then
      lngNumColumns = objRs( "Ordinal_Position" ).value
    end if
    objRs.moveNext
  wend
  redim preserve arrColumns( lngNumColumns )
  objRs.moveFirst
  while not objRs.eof
    arrColumns( objRs( "Ordinal_Position" ).value ) = objRs( "Column_Name" ).value
    objRs.moveNext
  wend
  objRs.close
  set objRs = nothing
end sub

'***********************************************
sub vbsDbHandleEditFieldsIsEmptyActually_forAccessJet( byRef objVbsDb )
'***********************************************
  ' we don't use funObjAdoxEditTableColumns() because the Jet provider returns
  ' columns in alphabetical order, not in definition order
  dim arrColumns() , lngIndex
  vbsDbHandleEditFieldsIsEmptyActually_createOrderedArray arrColumns , objVbsDb
  for lngIndex = 1 to uBound( arrColumns )
    objVbsDb( "DictionaryEditFields" ).add _
      uCase( arrColumns( lngIndex ) ) , arrColumns( lngIndex )
  next
end sub

'***********************************************
sub vbsDbHandleEditFieldsIsEmptyActually_forNonAccessJet( byRef objVbsDb )
'***********************************************
  dim objColumn
  for each objColumn in funObjAdoxEditTableColumns( objVbsDb )
    objVbsDb( "DictionaryEditFields" ).add uCase( objColumn.name ) , objColumn.name
  next
end sub

'***********************************************
sub vbsDbHandleEditFieldsIsEmptyActually( byRef objVbsDb )
'***********************************************
  if uCase( objVbsDb( "GlobalDbType" ) ) = "ACCESS" and uCase( objVbsDb( "GlobalOleDbProvider" ) ) <> "ODBC" then
    vbsDbHandleEditFieldsIsEmptyActually_forAccessJet objVbsDb
  else
    vbsDbHandleEditFieldsIsEmptyActually_forNonAccessJet objVbsDb
  end if
end sub

'***********************************************
sub vbsDbHandleEditFieldsIsEmpty( byRef objVbsDb )
'***********************************************
  if objVbsDb( "EditFields" ) = "" then
    vbsDbHandleEditFieldsIsEmptyActually objVbsDb
  end if 
end sub

'***********************************************
sub vbsDbSetDictionaryEditFields( byRef objVbsDb )
'***********************************************
  vbsDbSetDictionaryEditFieldsActually objVbsDb
  vbsDbHandleEditFieldsIsEmpty objVbsDb  ' if no edit field has been defined, all are addedd
end sub

'***********************************************
function funStrErrorMessageForSqlRecordSetFilterErrorFilterCondition( objVbsDb )
'***********************************************
  funStrErrorMessageForSqlRecordSetFilterErrorFilterCondition = "error applying this Filter Condition:<br><br>" & _
    cStr( objVbsDb( "FilterWhereClause" ) ) & "<br><br>" & _
    "Returned error description is: '" & err.description & "'<br><br>" & _
    "Please, perform the following debug steps:" & _
    "<div align=left>" & _
    "<ul>" & _
    "<li>control user inserted values are not of the wrong type. To avoid this kind of error, just add " & _
    "Javascript validation constraints.</li>" & _
    "</ul>" & _
    "</div>" & _
    "If none of the above solves, please, " & _
    "contact VBSdb support, providing " &_
    "them with this message and the debugging values below. Thanks."
end function

'***********************************************
function funStrErrorMessageForSqlRecordSetFilterErrorE_Fail( objVbsDb )
'***********************************************
  funStrErrorMessageForSqlRecordSetFilterErrorE_Fail = "Data Provider error.<br><br>" & _
    "<br><br>" & _
    "Returned error description is: '" & err.description & "'<br><br>" & _
    "This error does not depend on VBSdb: this error arises when ADO, for some unpredictable reason, falls in an error state.<br><br>" & _
    "Developers usually define the E_FAIL message an ADO bug, but we don't know if Microsoft would accept such definition :)<br><br>" & _
    "Please, perform the following workaround steps:" & _
    "<div align=left>" & _
    "<ul>" & _
    "<li>update your the MDAC version of your system. It's free of charge and it can be downloaded from Microsoft's site.</li>" & _
    "<li>often, this error arises when right joins, left joins and outer joins are used. In some cases, some ADO installations doesn't treat such Sql " &_
    "statements properly. Try to use an equivalent Sql query, avoiding right joins. " & _
    "For example, instead of<div align=center>'select * from Suppliers right join Products on Suppliers.SupplierId=Products.SupplierId'</div></li>" & _
    "<div align=center>use</div> <div align=center>'select * from Suppliers,Products where Suppliers.SupplierId=Products.SupplierId or Products.SupplierId is null'</div><br></li>" & _
    "</ul>" & _
    "</div>" & _
    "If none of the above solves, please, " & _
    "contact VBSdb support, providing " &_
    "them with this message and the debugging values below. Thanks."
end function

'***********************************************
function funStrErrorMessageForSqlRecordSetFilterUnexpectedError( objVbsDb )
'***********************************************
  funStrErrorMessageForSqlRecordSetFilterUnexpectedError = "Unexpected VBSdb error.<br><br>" & _
    "<br><br>" & _
    "If none of the above solves, please, " & _
    "Please, contact the VBSdb support, providing " &_
    "them with this message and the debugging values below. Thanks."
end function

'***********************************************
function funStrErrorMessageForSqlRecordSetFilter( objVbsDb )
'***********************************************
  if trim( objVbsDb( "FilterWhereClause" ) ) <> "" then
    ' the error should be determined by an error in the inserted filtering values
    funStrErrorMessageForSqlRecordSetFilter = _
      funStrErrorMessageForSqlRecordSetFilterErrorFilterCondition( objVbsDb )
  elseif inStr( uCase( err.description ) , "E_FAIL" ) > 0 then
    ' the error should be determined by a bad ADO behavior
    funStrErrorMessageForSqlRecordSetFilter = _
      funStrErrorMessageForSqlRecordSetFilterErrorE_Fail( objVbsDb )
  else
    ' this error is unexpected
    funStrErrorMessageForSqlRecordSetFilter = _
      funStrErrorMessageForSqlRecordSetFilterUnexpectedError( objVbsDb )
  end if
end function

'***********************************************
'sub vbsDbSetSqlRecordSetFilter( byRef objVbsDb )
'***********************************************
'  dim strErrorMessage
'  on error resume next
'  objVbsDb( "SqlRecordSet" ).Filter = cStr( objVbsDb( "FilterWhereClause" ) )
'  if err.number<>0 then
'    strErrorMessage = funStrErrorMessageForSqlRecordSetFilter( objVbsDb )
'    drawErrorWithDebugging strErrorMessage , objVbsDb
'  end if
'end sub

'***********************************************
sub vbsDbSetSqlRecordSetSort( byRef objVbsDb )
'***********************************************
  if objVbsDb( "OrderBy" ) <> "" then
    objVbsDb( "SqlRecordSet" ).Sort = objVbsDb( "OrderBy" )
  end if
end sub

'***********************************************
sub vbsDbSetDictionaryObjectForGridFields( byRef objVbsDb )
'***********************************************
  vbsDbSetDictionaryObjectFromList "GridFields" , objVbsDb
end sub

'***********************************************
sub vbsDbSetDictionaryObjectForFormFields( byRef objVbsDb )
'***********************************************
  vbsDbSetDictionaryObjectFromList "FormFields" , objVbsDb
end sub

'***********************************************
sub vbsDbSetDictionaryObjectForSearchFields( byRef objVbsDb )
'***********************************************
  vbsDbSetDictionaryObjectFromList "SearchFields" , objVbsDb
end sub

'***********************************************
sub vbsDbSetDictionaryObjectsBasedOnListPropertiesThatDependOnSqlRecordset( byRef objVbsDb )
'***********************************************
  vbsDbSetDictionaryObjectForGridFields objVbsDb
  vbsDbSetDictionaryObjectForFormFields objVbsDb
  vbsDbSetDictionaryObjectForSearchFields objVbsDb
end sub

'***********************************************
sub vbsDbHandleSqlRecordsetRelatedEmptyProperty( strPropertyName , byRef objVbsDb )
'***********************************************
  dim objField
  for each objField in objVbsDb( "SqlRecordSet" ).fields
    if not objVbsDb( strPropertyName ).exists( uCase( objField.name ) ) then
      ' current field has not already been added; if it should happen a field to
      ' be already added, probably it would be a join query, with a repeated field: in
      ' such a case, VbsDb considers it just once
      objVbsDb( strPropertyName ).add uCase( objField.name ) , objField.name
    end if
  next
end sub

'***********************************************
sub vbsDbHandleGridFieldsIsEmpty( byRef objVbsDb )
'***********************************************
  if objVbsDb( "GridFields" ) = "" then
    vbsDbHandleSqlRecordsetRelatedEmptyProperty "DictionaryGridFields" , objVbsDb
  end if 
end sub

'***********************************************
sub vbsDbHandleFormFieldsIsEmpty( byRef objVbsDb )
'***********************************************
  if objVbsDb( "FormFields" ) = "" then
    vbsDbHandleSqlRecordsetRelatedEmptyProperty "DictionaryFormFields" , objVbsDb
  end if 
end sub

'***********************************************
sub vbsDbHandleSearchFieldsIsEmpty( byRef objVbsDb )
'***********************************************
  if objVbsDb( "SearchFields" ) = "" then
    vbsDbHandleSqlRecordsetRelatedEmptyProperty "DictionarySearchFields" , objVbsDb
  end if 
end sub

'***********************************************
sub vbsDbHandleDictionaryObjectsIfListPropertiesThatDependOnSqlRecordsetAreEmpty( byRef objVbsDb )
'***********************************************
  vbsDbHandleGridFieldsIsEmpty objVbsDb
  vbsDbHandleFormFieldsIsEmpty objVbsDb
  vbsDbHandleSearchFieldsIsEmpty objVbsDb
end sub

'***********************************************
sub vbsDbHandleGridHideFieldsFormHideFieldsAndSearchHideFields( byRef objVbsDb )
'***********************************************
  vbsDbHandleGridHideFields objVbsDb
  vbsDbHandleFormHideFields objVbsDb
  vbsDbHandleSearchHideFields objVbsDb
end sub

'***********************************************
sub vbsDbSetDictionaryObjectsForListPropertiesThatDependOnSqlRecordset( byRef objVbsDb )
'***********************************************
  vbsDbSetDictionaryObjectsBasedOnListPropertiesThatDependOnSqlRecordset objVbsDb
  vbsDbHandleDictionaryObjectsIfListPropertiesThatDependOnSqlRecordsetAreEmpty objVbsDb
  vbsDbHandleGridHideFieldsFormHideFieldsAndSearchHideFields objVbsDb
end sub

'***********************************************
sub vbsDbSetSqlRecordPageSize( byRef objVbsDb )
'***********************************************
  objVbsDb( "SqlRecordSet" ).PageSize = objVbsDb( "GridPageSize" )
end sub

'***********************************************
sub vbsDbSetSqlRecordSetRelatedProperties( byRef objVbsDb )
'***********************************************
  vbsDbSetSqlRecordSet objVbsDb
  vbsDbSetDictionaryObjectsForListPropertiesThatDependOnSqlRecordset objVbsDb
  'vbsDbCheckErrorBadFieldsForSqlRecordSet objVbsDb
  'vbsDbSetSqlRecordSetFilter objVbsDb
  'vbsDbSetSqlWithFilterWhereClause objVbsDb
  vbsDbSetSqlRecordSetSort objVbsDb
  vbsDbSetSqlRecordPageSize objVbsDb
end sub

'***********************************************
sub vbsDbSetNavigationProperties( byRef objVbsDb )
'***********************************************
  objVbsDb( "ViewNavigationFirst" ) = funBolInDbNavigationItem( "FIRST" , objVbsDb )
  objVbsDb( "ViewNavigationPrev" ) = funBolInDbNavigationItem( "PREV" , objVbsDb )
  objVbsDb( "ViewNavigationNext" ) = funBolInDbNavigationItem( "NEXT" , objVbsDb )
  objVbsDb( "ViewNavigationLast" ) = funBolInDbNavigationItem( "LAST" , objVbsDb )
  objVbsDb( "ViewNavigationRemoveFilter" ) = funBolInDbNavigationItem( "REMOVEFILTER" , objVbsDb )
  objVbsDb( "ViewNavigationSearch" ) = funBolInDbNavigationItem( "SEARCH" , objVbsDb )
  objVbsDb( "ViewNavigationAdd" ) = funBolInDbNavigationItem( "ADD" , objVbsDb )
  objVbsDb( "ViewNavigationUpdate" ) = funBolInDbNavigationItem( "UPDATE" , objVbsDb )
  objVbsDb( "ViewNavigationDelete" ) = funBolInDbNavigationItem( "DELETE" , objVbsDb )
end sub

'***********************************************
sub vbsDbGetSessionVariables( byRef objVbsDb )
'***********************************************
  if session( "VbsDbSessionFormAbsolutePosition_" & objVbsDb( "GlobalId" ) ) <> "" then
  	' it is not the first time this page was called in this session
  	objVbsDb( "FormAbsolutePosition" ) = session( "VbsDbSessionFormAbsolutePosition_" & objVbsDb( "GlobalId" ) )
  end if
  objVbsDb( "OrderBy" ) = session( "VbsDbSessionOrderBy_" & objVbsDb( "GlobalId" ) )
  'objVbsDb( "EditWhere" ) = session( "sv_editWhere_" & objVbsDb( "GlobalId" ) )
  objVbsDb( "FilterWhereClause" ) = session( "VbsDbFilterWhereClause_" & objVbsDb( "GlobalId" ) )
end sub

'***********************************************
sub vbsDbSetRequestProperties( byRef objVbsDb )
'***********************************************
  objVbsDb( "RequestVBSdbClickClass" ) = vbsDbGetClickClass( objVbsDb )
  objVbsDb( "RequestVBSGridIndex" ) = cLng( request( "VBSdbIndex_" & objVbsDb( "GlobalId" ) ) )
  objVbsDb( "RequestVBSdbGridSort" ) = cStr( request( "VBSdbGridSort_" & objVbsDb( "GlobalId" ) ) )
  objVbsDb( "RequestVBSdbEditWhere" ) = _
    replace( cStr( request( "VBSdbEditWhere_" & objVbsDb( "GlobalId" ) ) ) , constStrReplacementForQuotesInEditWhereConditions , """" )
end sub

'***********************************************
sub vbsDbGetInputVariables( byRef objVbsDb )
'***********************************************
  vbsDbGetSessionVariables objVbsDb
  vbsDbSetRequestProperties objVbsDb
end sub

'***********************************************
sub vbsDbInitializeVariables( byRef objVbsDb )
'***********************************************
  vbsDbInitializeOutputPropertiesAndPrivateProperties objVbsDb
  vbsDbHandleReset objVbsDb                         ' resets session variables if we come to this page from a different page
  vbsDbGetInputVariables objVbsDb
end sub

'***********************************************
sub vbsDbManageDictionariesRelatedToEditFields( byRef objVbsDb )
'***********************************************
  vbsDbSetDictionaryEditKeyFields objVbsDb
  if objVbsDb( "EditTableName" ) <> "" then
    vbsDbSetDictionaryEditFields objVbsDb
    vbsDbHandleEditHideFields objVbsDb
    vbsDbSetDictionaryEditReadOnlyFields objVbsDb
    vbsDbAddAdditionalFieldsToDictionaryEditReadOnlyFields objVbsDb
    vbsDbSetDictionaryEditAddFieldDefaults objVbsDb
    vbsDbSetDictionaryEditUpdateFieldDefaults objVbsDb
    vbsDbSetDictionaryObjectForEditMemoFields objVbsDb
    'vbsDbCheckErrorBadFieldsForEditTableRecordSet objVbsDb
  end if
end sub

'***********************************************
sub vbsDbSetVariablesNeededForFormTamperingControls( byRef objVbsDb )
'***********************************************
  vbsDbSetDictionaryInputSelectFields objVbsDb
  vbsDbSetDictionaryInputEnumeratedFields objVbsDb
  'vbsDbSetDictionarySearchAliasFields objVbsDb
end sub

'***********************************************
sub vbsDbSetPropertiesNeededForEditDb( byRef objVbsDb )
'***********************************************
  vbsDbSetGlobalSeparators objVbsDb
  vbsDbSetNavigationProperties objVbsDb
  vbsDbSetConnection objVbsDb
  vbsDbSetDictionaryGlobalBooleanFields objVbsDb
  vbsDbSetAdoxCatalog objVbsDb
  vbsDbManageDictionariesRelatedToEditFields objVbsDb
  vbsDbSetVariablesNeededForFormTamperingControls objVbsDb
  'vbsDbSetInputValidateDateRelatedProperties objVbsDb
end sub

'***********************************************
sub vbsDbManageProgramVariablesBeforeEditDb( byRef objVbsDb )
'***********************************************
  vbsDbHandlePropertySintaxes objVbsDb
  vbsDbCheckErrorPropertyAssignment objVbsDb       ' check if the user has assigned an unsupported property name
  vbsDbInitializeVariables objVbsDb
  vbsDbSetPropertiesNeededForEditDb objVbsDb
end sub

'***********************************************
sub vbsDbSetDictionaryEditKeyFields( byRef objVbsDb )
'***********************************************
  vbsDbSetDictionaryObjectFromList "EditKeyFields" , objVbsDb
end sub

'***********************************************
sub vbsDbSetDictionaryObjectsForListPropertiesThatDoNotDependOnSqlRecordSet( byRef objVbsDb )
'***********************************************
  vbsDbSetDictionaryObjectForViewFieldFormats objVbsDb
  vbsDbSetDictionaryGlobalFieldHeaders objVbsDb
  'vbsDbSetDictionaryInputSelectFields objVbsDb
  'vbsDbSetDictionaryInputEnumeratedFields objVbsDb
  vbsDbSetDictionaryEditValidateRequiredFields objVbsDb
  vbsDbSetDictionaryEditValidateRegExp objVbsDb
  vbsDbSetDictionaryGlobalCustomText objVbsDb
  vbsDbSetDictionaryInputPasswordFields objVbsDb
  vbsDbSetDictionaryLanguage objVbsDb
end sub

'***********************************************
sub vbsDbSetGlobalImageDir( byRef objVbsDb )
'***********************************************
  if ( objVbsDb( "GlobalImageDir" ) <> "" ) and ( right( objVbsDb( "GlobalImageDir" ) , 1 ) <> "/" ) then
    ' objVbsDb( "GlobalImageDir" ) does not end with a '/' simbol
    objVbsDb( "GlobalImageDir" ) = objVbsDb( "GlobalImageDir" ) & "/"
  end if
end sub

'***********************************************
function funBolInDbNavigationItem( strStringToSearch , objVbsDb )
'***********************************************
	funBolInDbNavigationItem = ( inStr( uCase( objVbsDb( "ViewNavigationButtons" ) ) , uCase( strStringToSearch ) ) > 0 )
end function

'***********************************************
sub vbsDbSetScreenTypeRelatedProperties( byRef objVbsDb )
'***********************************************
  select case VbsDbGetScreenType( objVbsDb )
  	case "Search"
  	  ' it is an add screen
  	  objVbsDb( "ScreenTitle" ) = funStrTranslate( "searchScreenTitle" , objVbsDb )
  	  objVbsDb( "InputFormLeftHeader" ) = funStrTranslate( "searchFormLeftHeader" , objVbsDb )
  	  objVbsDb( "InputFormRightHeader" ) = funStrTranslate( "searchFormRightHeader" , objVbsDb )
      objVbsDb( "InputCancelContent" ) = funStrTranslate( "searchCancelContent" , objVbsDb )
      objVbsDb( "InputCancelTitle" ) = funStrTranslate( "searchCancelTitle" , objVbsDb )
  	  objVbsDb( "InputSubmitValue" ) = funStrTranslate( "searchSubmit" , objVbsDb )
  	  objVbsDb( "InputResetValue" ) = funStrTranslate( "searchReset" , objVbsDb )
  	  set objVbsDb( "DictionaryInputFields" ) = objVbsDb( "DictionarySearchFields" )
  	case "Add"
  	  ' it is an add screen
  	  objVbsDb( "ScreenTitle" ) = funStrTranslate( "addScreenTitle" , objVbsDb )
   	  objVbsDb( "InputFormLeftHeader" ) = funStrTranslate( "addFormLeftHeader" , objVbsDb )
 	    objVbsDb( "InputFormRightHeader" ) = funStrTranslate( "addFormRightHeader" , objVbsDb )
      objVbsDb( "InputCancelContent" ) = funStrTranslate( "addCancelContent" , objVbsDb )
      objVbsDb( "InputCancelTitle" ) = funStrTranslate( "addCancelTitle" , objVbsDb )
  	  objVbsDb( "InputSubmitValue" ) = funStrTranslate( "addSubmit" , objVbsDb )
  	  objVbsDb( "InputResetValue" ) = funStrTranslate( "addReset" , objVbsDb )
  	  set objVbsDb( "DictionaryInputFields" ) = objVbsDb( "DictionaryEditFields" )
  	case "Update"
  	  ' it is an update screen
  	  objVbsDb( "ScreenTitle" ) = funStrTranslate( "updateScreenTitle" , objVbsDb )
  	  objVbsDb( "InputFormLeftHeader" ) = funStrTranslate( "updateFormLeftHeader" , objVbsDb )
  	  objVbsDb( "InputFormRightHeader" ) = funStrTranslate( "updateFormRightHeader" , objVbsDb )
      objVbsDb( "InputCancelContent" ) = funStrTranslate( "updateCancelContent" , objVbsDb )
      objVbsDb( "InputCancelTitle" ) = funStrTranslate( "updateCancelTitle" , objVbsDb )
  	  objVbsDb( "InputSubmitValue" ) = funStrTranslate( "updateSubmit" , objVbsDb )
  	  objVbsDb( "InputResetValue" ) = funStrTranslate( "updateReset" , objVbsDb )
  	  set objVbsDb( "DictionaryInputFields" ) = objVbsDb( "DictionaryEditFields" )
  	case "Delete"
  	  ' it is an update screen
  	  objVbsDb( "ScreenTitle" ) = funStrTranslate( "deleteScreenTitle" , objVbsDb )
  	  objVbsDb( "InputFormLeftHeader" ) = funStrTranslate( "deleteFormLeftHeader" , objVbsDb )
  	  objVbsDb( "InputFormRightHeader" ) = funStrTranslate( "deleteFormRightHeader" , objVbsDb )
      objVbsDb( "InputCancelContent" ) = funStrTranslate( "deleteCancelContent" , objVbsDb )
      objVbsDb( "InputCancelTitle" ) = funStrTranslate( "deleteCancelTitle" , objVbsDb )
  	  objVbsDb( "InputSubmitValue" ) = funStrTranslate( "deleteSubmit" , objVbsDb )
  	  objVbsDb( "InputResetValue" ) = funStrTranslate( "deleteReset" , objVbsDb )
  	  set objVbsDb( "DictionaryInputFields" ) = objVbsDb( "DictionaryEditFields" )
  end select
end sub

'***********************************************
sub vbsDbBuildMiscellaneousProperties( byRef objVbsDb )
'***********************************************
  'vbsDbChangeFormAbsolutePosition objVbsDb
  vbsDbSetGlobalImageDir objVbsDb
  'vbsDbSetEditValidate objVbsDb
  vbsDbSetScreenTypeRelatedProperties objVbsDb
  vbsDbSetFormAbsolutePosition objVbsDb
  vbsDbSetGridAbsolutePage objVbsDb
  vbsDbSetGridTemplateRelatedProperties objVbsDb
  vbsDbSetFormTemplateRelatedProperties objVbsDb
  vbsDbSetInputTemplateRelatedProperties objVbsDb
  vbsDbSetEditSlqServerAutoIncrement objVbsDb
end sub

'***********************************************
sub vbsDbSetSqlRecordSetUnrelatedProperties( byRef objVbsDb )
'***********************************************
  vbsDbSetDictionaryObjectsForListPropertiesThatDoNotDependOnSqlRecordSet objVbsDb
  vbsDbBuildMiscellaneousProperties objVbsDb
end sub

'***********************************************
sub vbsDbSetPropertiesAfterEditDb( byRef objVbsDb )
'***********************************************
  vbsDbSetDictionarySearchAliasFields objVbsDb
  vbsDbSetDictionarySearchOperators objVbsDb
  vbsDbSetSqlWithFilterWhereClause objVbsDb
  vbsDbSetOrderBy objVbsDb
  vbsDbSetEditTableRecordSet objVbsDb
  vbsDbSetSqlRecordSetRelatedProperties objVbsDb
  vbsDbSetSqlRecordSetUnrelatedProperties objVbsDb
  vbsDbCheckErrorBadFieldNames objVbsDb
  vbsDbSetInputValidateDateRelatedProperties objVbsDb
end sub

'***********************************************
sub vbsDbSetSessionVariables( byRef objVbsDb )
'***********************************************
  session( "VbsDbSessionFormAbsolutePosition_" & objVbsDb( "GlobalId" ) ) = objVbsDb( "FormAbsolutePosition" )
  session( "VbsDbSessionOrderBy_" & objVbsDb( "GlobalId" ) ) = objVbsDb( "OrderBy" )
  session( "VbsDbFilterWhereClause_" & objVbsDb( "GlobalId" ) ) = objVbsDb( "FilterWhereClause" )
end sub

'***********************************************
sub vbsDbManageProgramVariablesAfterEditDb( byRef objVbsDb )
'***********************************************
  vbsDbSetPropertiesAfterEditDb objVbsDb
  'vbsDbCheckErrorWrongPropertyAssignment objVbsDb   ' check if the user has assigned an unsupported property value
  vbsDbSetSessionVariables objVbsDb
end sub

'***********************************************
sub vbsDbManageProgramVariablesAndDataBaseData( byRef objVbsDb )
'***********************************************
  vbsDbManageProgramVariablesBeforeEditDb objVbsDb
  vbsDbEditDb objVbsDb								' updates the database (if needed)
  'if objVbsDb( "EditErrorNumber" ) = 0 then
    ' there has been no error while updating the database
    vbsDbManageProgramVariablesAfterEditDb objVbsDb
  'end if
end sub

'***********************************************
function funBolVbsDbDrawActually( objVbsDb )
'***********************************************
  funBolVbsDbDrawActually = not ( objVbsDb( "EditAddSingle" ) and _
    ( ( uCase( VbsDbGetLastAction( objVbsDb ) ) = "SUBMITADD" ) or _
      ( uCase( VbsDbGetLastAction( objVbsDb ) ) = "CANCEL" ) ) )
end function

'***********************************************
'***********************************************
sub VbsDb( byRef objVbsDb )
'***********************************************
'***********************************************
  vbsDbManageProgramVariablesAndDataBaseData objVbsDb
  if funBolVbsDbDrawActually( objVbsDb ) then
    vbsDbDraw objVbsDb
  end if
end sub

'***********************************************
'***********************************************
sub VbsDbClose( byRef objVbsDb )
'***********************************************
'***********************************************
  on error resume next
  set objVbsDb( "AdoxCatalog" ) = nothing
  objVbsDb( "AdoxCatalog" ) = ""
  if isObject( objVbsDb( "SqlRecordSet" ) ) then
    objVbsDb( "SqlRecordSet" ).close
    set objVbsDb( "SqlRecordSet" ) = nothing
  end if
  objVbsDb( "SqlRecordSet" ) = ""
  if isObject( objVbsDb( "EditTableRecordSet" ) ) then
    objVbsDb( "EditTableRecordSet" ).close
    set objVbsDb( "EditTableRecordSet" ) = nothing
  end if
  objVbsDb( "EditTableRecordSet" ) = ""
  if isObject( objVbsDb( "Connection" ) ) then
    objVbsDb( "Connection" ).close
    set objVbsDb( "Connection" ) = nothing
  end if
  objVbsDb( "Connection" ) = ""
  vbsDbDictionaryClose( objVbsDb )
end sub
%>
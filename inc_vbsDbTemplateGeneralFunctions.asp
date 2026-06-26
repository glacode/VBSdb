<%
'***********************************************
'***********************************************
function funStrTextFileContent( strTextFileName , objVbsDb )
'***********************************************
'***********************************************
  dim objFileSystemObject , objTextStream , strPhysicalPath , strErrorMessage
  set objFileSystemObject = CreateObject( "Scripting.FileSystemObject" )
  on error resume next
  strPhysicalPath = server.mapPath( strTextFileName )
  if err.number<>0 then
    strErrorMessage = "Error trying to determine the physical path for this Template file:<br><br>" & _
      strTextFileName & "<br><br>" & _
      "Returned error description is: '" & err.description & "'<br><br>" & _
      "Please, perform the following debugging controls:" & _
      "<div align=left>" & _
      "<ul><li>don't assign a physical path to the Template property. Use a virtual path instead</li>" & _
      "<li>don't use the server.mapPath method for the Template property assignment</li>" & _
      "</ul>" & _
      "</div>"
      drawErrorWithDebugging strErrorMessage , objVbsDb
  end if
  set objTextStream = objFileSystemObject.openTextFile( server.mapPath( strTextFileName ) , 1 , false , 0 ) ' readOnly, don't create, ASCII
  if err.number<>0 then
    strErrorMessage = "Error opening this Template file:<br><br>" & _
      server.mapPath( strTextFileName ) & "<br><br>" & _
      "Returned error description is: '" & err.description & "'<br><br>" & _
      "Please, perform the following debugging controls:" & _
      "<div align=left>" & _
      "<ul><li>control the displayed file exists on the server</li>" & _
      "<li>control the internet user has the permission to read that file</li>" & _
      "<li>control the template property assignment</li>" & _
      "</ul>" & _
      "</div>"
      drawErrorWithDebugging strErrorMessage , objVbsDb
  end if
  funStrTextFileContent = objTextStream.readAll
  set objTextStream = nothing
  set objFileSystemObject = nothing
end function

'***********************************************
'***********************************************
function funStrCutSpacesWithinElement( strElementName , strStrinWithElements )
'***********************************************
'***********************************************
  funStrCutSpacesWithinElement = strStrinWithElements
  while inStr( 1 , funStrCutSpacesWithinElement , "<" & strElementName & "> " , 1 ) > 0
    funStrCutSpacesWithinElement = replace( funStrCutSpacesWithinElement , "<" & strElementName & "> " , "<" & strElementName & ">" , 1 , -1 , 1 )
  wend
  while inStr( 1 , funStrCutSpacesWithinElement , " </" & strElementName & ">" , 1 ) > 0
    funStrCutSpacesWithinElement = replace( funStrCutSpacesWithinElement , " </" & strElementName & ">" , "</" & strElementName & ">" , 1 , -1 , 1 )
  wend
end function

'***********************************************
function funStrPrivateTemplateVariableName( strTemplateType )
'***********************************************
  if strTemplateType = "Grid" then
    ' a Grid template is being computed
    funStrPrivateTemplateVariableName = "GridTemplateRowTemplate"
  else
    ' a Form template is being computed
    funStrPrivateTemplateVariableName = "FormTemplateContent"
  end if
end function

'***********************************************
function funStrVbsDbFieldParameterValue( strFieldName , objVbsDb )
'***********************************************
  select case VbsDbGetScreenType( objVbsDb )
    case "View"
      funStrVbsDbFieldParameterValue = funStrConvertNullToEmptyString( objVbsDb( "SqlRecordSet" ).fields( strFieldName ).value )
    case "Update" , "Delete"
      funStrVbsDbFieldParameterValue = funStrConvertNullToEmptyString( objVbsDb( "EditTableRecordSet" ).fields( strFieldName ).value )
  end select
  funStrVbsDbFieldParameterValue = replace( funStrVbsDbFieldParameterValue , """" , """" & """" )
  funStrVbsDbFieldParameterValue = replace( funStrVbsDbFieldParameterValue , chr( 13 ) & chr( 10 ) , "<br>" )
  funStrVbsDbFieldParameterValue = """" & funStrVbsDbFieldParameterValue & """"
end function

'***********************************************
function funStrTemplateTransformationWithoutBadSpacesTransformFieldValues( strTemplateType , strTemplateWithoutBadSpaces , objVbsDb )
'***********************************************
  dim strFieldName
  funStrTemplateTransformationWithoutBadSpacesTransformFieldValues = strTemplateWithoutBadSpaces
  for each strFieldName in objVbsDb( "Dictionary" & strTemplateType & "Fields" )
  	funStrTemplateTransformationWithoutBadSpacesTransformFieldValues = _
  	  replace( funStrTemplateTransformationWithoutBadSpacesTransformFieldValues , _
  	  "<VbsDbFieldParameter>" & strFieldName & "</VbsDbFieldParameter>" , _
      funStrVbsDbFieldParameterValue( strFieldName , objVbsDb ) , 1 , -1 , 1 )   ' last one is needed to perform a text comparison, so that grid template file field bookmarks become case insensitive
  	funStrTemplateTransformationWithoutBadSpacesTransformFieldValues = _
  	  replace( funStrTemplateTransformationWithoutBadSpacesTransformFieldValues , "<VbsDbFieldValue>" & strFieldName & "</VbsDbFieldValue>" , _
      funStrFieldValueForViewScreen( strFieldName , objVbsDb ) , 1 , -1 , 1 )   ' last one is needed to perform a text comparison, so that grid template file field bookmarks become case insensitive
  next
end function

'***********************************************
sub vbsDbTemplateTransformationPerformNextExecuteActually( lngNextExecuteBegin , lngNextExecuteEnd , byRef strTemplateWithoutBadSpaces , byRef objVbsDb )
'***********************************************
  dim strExecuteContent , varExecuteReturnValue , strErrorMessage
  strExecuteContent = mid( strTemplateWithoutBadSpaces , lngNextExecuteBegin + 14 , lngNextExecuteEnd - lngNextExecuteBegin - 14 )
  on error resume next
  execute "varExecuteReturnValue = " & strExecuteContent
  if err.number <> 0 then
    strErrorMessage = "Error trying to execute a server side formatting.<br><br>" & _
      "Returned error description is: '" & err.description & "'<br><br>" & _
      "The error arised trying to execute the following server side formatting statement:<br><br>" & _
      "<center>" & "varExecuteReturnValue = " & funStrConvertNullToEmptyString( strExecuteContent ) & "</center>" & _
      "Please, perform the following debugging controls:" & _
      "<div align=left>" & _
      "<ul><li>check the &lt;VbsDbExecute&gt; elements in your Grid and Form templates</li>" & _
      "<li>check the &lt;VbsDbExecute&gt; elements in your ViewFieldFormats property</li>" & _
      "<li>contol the invoked function is actually either defined or included in your page</li>" & _
      "<li>control your script engine version is 5.0 or higher (see below)</li>" & _
      "</ul>" & _
      "</div>"
    drawErrorWithDebugging strErrorMessage , objVbsDb
  end if
  strTemplateWithoutBadSpaces = _
    left( strTemplateWithoutBadSpaces , lngNextExecuteBegin - 1 ) & _
    varExecuteReturnValue & _
    mid( strTemplateWithoutBadSpaces , lngNextExecuteEnd + 15 )
end sub

'***********************************************
sub vbsDbTemplateTransformationPerformNextExecute( byRef lngParsePosition , byRef strTemplateWithoutBadSpaces , byRef objVbsDb)
'***********************************************
  dim lngNextExecuteBegin , lngNextExecuteEnd
  lngNextExecuteBegin = inStr( lngParsePosition + 1 , strTemplateWithoutBadSpaces , "<VbsDbExecute>" , 1 )
  if lngNextExecuteBegin > 0 then
    ' the begin of a new item to be executed has been found
    lngParsePosition = lngNextExecuteBegin - 1
    lngNextExecuteEnd = inStr( lngNextExecuteBegin + 1 , strTemplateWithoutBadSpaces , "</VbsDbExecute>" , 1 )
    if lngNextExecuteEnd > 0 then
      ' the end of the new item to be executed has been found
      vbsDbTemplateTransformationPerformNextExecuteActually lngNextExecuteBegin , lngNextExecuteEnd , strTemplateWithoutBadSpaces , objVbsDb
    else
      ' the end of the new item to be executed has not been found
      lngParsePosition = len( strTemplateWithoutBadSpaces ) + 1
    end if
  else
    ' the begin of a new item to be executed has not been found
    lngParsePosition = len( strTemplateWithoutBadSpaces ) + 1
  end if
end sub

'***********************************************
function funStrTemplateTransformationWithoutBadSpacesPerformExecutes( strTemplateWithoutBadSpaces , objVbsDb )
'***********************************************
  dim lngParsePosition
  lngParsePosition = 0
  while lngParsePosition < len( strTemplateWithoutBadSpaces )
    vbsDbTemplateTransformationPerformNextExecute lngParsePosition , strTemplateWithoutBadSpaces , objVbsDb
  wend
  funStrTemplateTransformationWithoutBadSpacesPerformExecutes = strTemplateWithoutBadSpaces
end function

'***********************************************
function funStrTemplateTransformationWithoutBadSpaces( strTemplateType , strTemplateWithoutBadSpaces , objVbsDb )
'***********************************************
  funStrTemplateTransformationWithoutBadSpaces = _
    funStrTemplateTransformationWithoutBadSpacesTransformFieldValues( strTemplateType , strTemplateWithoutBadSpaces , objVbsDb )
  funStrTemplateTransformationWithoutBadSpaces = _
    funStrTemplateTransformationWithoutBadSpacesPerformExecutes( funStrTemplateTransformationWithoutBadSpaces , objVbsDb )
end function

'***********************************************
'***********************************************
function funStrVbsDbTemplateTransformation( strTemplateType , objVbsDb )
'***********************************************
'***********************************************
  funStrVbsDbTemplateTransformation = funStrCutSpacesWithinElement( "VbsDbFieldValue" , objVbsDb( funStrPrivateTemplateVariableName( strTemplateType ) ) )
  funStrVbsDbTemplateTransformation = funStrTemplateTransformationWithoutBadSpaces( strTemplateType , funStrVbsDbTemplateTransformation , objVbsDb )
end function
%>
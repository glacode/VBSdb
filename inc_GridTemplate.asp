<%
'***********************************************
'***********************************************
sub vbsDbDrawTemplateGridTableBegin( objVbsDb )
'***********************************************
'***********************************************
%>
    <table id="GridTemplate">
<%
end sub

'***********************************************
'***********************************************
sub vbsDbDrawTemplateGridTableEnd()
'***********************************************
'***********************************************
%>
  </table>
<%
end sub

'***********************************************
function funLngElementContentBeginPosition( strElementName , strString , objVbsDb )
'***********************************************
  funLngElementContentBeginPosition = inStr( 1 , strString , "<" & strElementName & ">" , 1 )
  if funLngElementContentBeginPosition = 0 then
    drawError "GridTemplate Error. &ltVbsDbGridRow&gt; not found" , objVbsDb
  else
    funLngElementContentBeginPosition = funLngElementContentBeginPosition + len( strElementName ) + 2
  end if
end function

'***********************************************
function funLngElementContentEndPosition( strElementName , strString , objVbsDb )
'***********************************************
  funLngElementContentEndPosition = inStr( 1 , strString , "</VbsDbGridRow>" , 1 )
  if funLngElementContentEndPosition = 0 then
    drawError "GridTemplate Error. &/ltVbsDbGridRow&gt; not found" , objVbsDb
  else
    funLngElementContentEndPosition = funLngElementContentEndPosition - 1
  end if
end function

'***********************************************
function funStrStringElementContent( strElementName , strString , objVbsDb )
'***********************************************
  dim lngElementContentBeginPosition , lngElementContentEndPosition , lngElementContentLength
  lngElementContentBeginPosition = funLngElementContentBeginPosition( strElementName , strString , objVbsDb )
  lngElementContentEndPosition = funLngElementContentEndPosition( strElementName , strString , objVbsDb )
  if lngElementContentBeginPosition >= lngElementContentEndPosition then
    drawError "GridTemplate Error. &ltVbsDbGridRow&gt; is not before &lt/VbsDbGridRow&gt;" , objVbsDb
  else
    lngElementContentLength = lngElementContentEndPosition - lngElementContentBeginPosition + 1
    funStrStringElementContent = mid( strString , lngElementContentBeginPosition , lngElementContentLength )
  end if
end function

'***********************************************
function funStrTemplateElementContent( strElementName , strTemplateFileName , objVbsDb )
'***********************************************
  dim strTemplateFileContent
  strTemplateFileContent = funStrTextFileContent( strTemplateFileName , objVbsDb )
  funStrTemplateElementContent = funStrStringElementContent( strElementName , strTemplateFileContent , objVbsDb )
end function

'***********************************************
sub vbsDbSetGridTemplateRowTemplate( objVbsDb )
'***********************************************
  objVbsDb( "GridTemplateRowTemplate" ) = funStrTemplateElementContent( "VbsDbGridRow" , objVbsDb( "GridTemplate" ) , objVbsDb )
end sub

'***********************************************
'***********************************************
sub vbsDbSetGridTemplateRelatedProperties( objVbsDb )
'***********************************************
'***********************************************
  if objVbsDb( "GridTemplate" ) <> "" then
    ' the user wants to dispalay a custom grid
    vbsDbSetGridTemplateRowTemplate objVbsDb
  end if
end sub

'***********************************************
'function funStrTemplateGridRowTransformationWithoutBadSpaces( strGridRowTemplateWithoutBadSpaces , objVbsDb )
'***********************************************
'  dim strFieldName
'  funStrTemplateGridRowTransformationWithoutBadSpaces = strGridRowTemplateWithoutBadSpaces
'  for each strFieldName in objVbsDb( "DictionaryGridFields" )
'  	funStrTemplateGridRowTransformationWithoutBadSpaces = replace( funStrTemplateGridRowTransformationWithoutBadSpaces , "<VbsDbFieldValue>" & strFieldName & "</VbsDbFieldValue>" , _
'    funStrFieldValueForViewScreen( strFieldName , objVbsDb ) , 1 , -1 , 1 )   ' last one is needed to perform a text comparison, so that grid template file field bookmarks become case insensitive
'  next
'end function

'***********************************************
'function funStrTemplateGridRowTransformation( objVbsDb )
'***********************************************
'  funStrTemplateGridRowTransformation = funStrCutSpacesWithinElement( "VbsDbFieldValue" , objVbsDb( "GridTemplateRowTemplate" ) )
'  funStrTemplateGridRowTransformation = funStrTemplateGridRowTransformationWithoutBadSpaces( funStrTemplateGridRowTransformation , objVbsDb )
'end function

'***********************************************
'***********************************************
sub vbsDbDrawTemplateGridRow( objVbsDb )
'***********************************************
'***********************************************
%>
  <tr>
    <td>
      <%=funStrVbsDbTemplateTransformation( "Grid" , objVbsDb )%>
    </td>
  </tr>
<%
end sub
%>
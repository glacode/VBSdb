<%
'***********************************************
sub vbsDbDrawStandardFormTableBeginActually( objVbsDb )
'***********************************************
  if objVbsDb( "FormTableTag" ) <> "" then
%>
    <table <%=objVbsDb( "FormTableTag" )%> id="FormBorders">
<%
  else
    ' objVbsDb( "FormTableTag" ) is empty
%>
    <table id="FormBorders">
<%
  end if
end sub

'***********************************************
sub vbsDbDrawStandardFormTableBegin( objVbsDb )
'***********************************************
  vbsDbBeginDrawForBorders "FormBorders"
  vbsDbDrawStandardFormTableBeginActually objVbsDb
end sub

'***********************************************
sub vbsDbDrawStandardFormRowDbFieldHeader( objField , objVbsDb )
'***********************************************
%>
  <td class="FormLeftColumn">
  	<%=funStrFieldHeader( objField.name , objVbsDb )%>
  </td>
<%
end sub

'***********************************************
sub vbsDbDrawStandardFormRowDbFieldData( objField , objVbsDb )
'***********************************************
%>
  <td class="FormRightColumn">
    <%=funStrFieldValueForViewScreen( objField.name , objVbsDb )%>
  </td>
<%
end sub

'***********************************************
sub vbsDbDrawStandardFormRowDbField( objField , objVbsDb )
'***********************************************
%>
  <tr class="Form">
<%
  	vbsDbDrawStandardFormRowDbFieldHeader objField , objVbsDb
  	vbsDbDrawStandardFormRowDbFieldData objField , objVbsDb
%>
  </tr>
<%
end sub

'***********************************************
sub vbsDbDrawStandardFormRow( strFieldName , objVbsDb )
'***********************************************
  dim objField
  set objField = objVbsDb( "SqlRecordSet" ).fields( strFieldName )
  vbsDbDrawStandardFormRowDbField objField , objVbsDb
end sub

'***********************************************
sub vbsDbDrawStandardFormRowsForCurrentRecord( objVbsDb )
'***********************************************
  dim strFieldName
  for each strFieldName in objVbsDb( "DictionaryFormFields" )
  	vbsDbDrawStandardFormRow strFieldName , objVbsDb
  next
end sub

'***********************************************
sub vbsDbDrawStandardFormTableBody( objVbsDb )
'***********************************************
	vbsDbDrawStandardFormRowsForCurrentRecord objVbsDb
end sub

'***********************************************
sub vbsDbDrawStandardFormTableEnd()
'***********************************************
%>
  </table>
<%
  vbsDbEndDrawForBorders
end sub

'***********************************************
sub vbsDbDrawStandardForm( objVbsDb )
'***********************************************
  vbsDbDrawStandardFormTableBegin objVbsDb
  vbsDbDrawStandardFormTableBody objVbsDb
  vbsDbDrawStandardFormTableEnd
end sub

'***********************************************
sub vbsDbDrawFormActually( objVbsDb )
'***********************************************
  if objVbsDb( "FormTemplate" ) = "" then
    ' no form template has been specified
    vbsDbDrawStandardForm objVbsDb
  else
    ' a form template has been specified
    vbsDbDrawTemplateForm objVbsDb
  end if
end sub

'***********************************************
sub vbsDbDrawForm( objVbsDb )
'***********************************************
  if objVbsDb( "SqlRecordSet" ).recordCount > 0 then
    ' the query returned at least one record
	  objVbsDb( "SqlRecordSet" ).absolutePosition = objVbsDb( "FormAbsolutePosition" )
	  vbsDbDrawFormActually objVbsDb
  end if
end sub
%>
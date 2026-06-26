<%
'***********************************************
function funStrVbsDbViewEmptyMessage( objVbsDb )
'***********************************************
  if objVbsDb( "FilterWhereClause" ) <> "" then
    ' the sql query gives no result, but the user added some filter
    'funStrVbsDbViewEmptyMessage = objVbsDb( "ViewEmptyFilterMessage" )
    funStrVbsDbViewEmptyMessage = funStrTranslate( "viewScreenEmptyFilter" , objVbsDb )
  else
    ' the sql query gives no result and the user is not applying any filter
    'funStrVbsDbViewEmptyMessage = objVbsDb( "ViewEmptySqlMessage" )
    funStrVbsDbViewEmptyMessage = funStrTranslate( "viewScreenEmptySql" , objVbsDb )
  end if
end function

'***********************************************
sub vbsDbDrawViewEmptyMessage( objVbsDb )
'***********************************************
%>
  <div class=ViewMessage>
    <%=funStrVbsDbViewEmptyMessage( objVbsDb )%>
  </div>
<%
end sub

'***********************************************
sub vbsDbDrawGridTableBeginActually( objVbsDb )
'***********************************************
  if objVbsDb( "GridTableTag" ) <> "" then
%>
    <table <%=objVbsDb( "GridTableTag" )%> id="Grid">
<%
  else
    ' objVbsDb( "GridTableTag" ) is empty
%>
    <table id="Grid">
<%
  end if
end sub

'***********************************************
sub vbsDbDrawStandardGridTableBegin( objVbsDb )
'***********************************************
  vbsDbBeginDrawForBorders "GridBorders"
  vbsDbDrawGridTableBeginActually objVbsDb
end sub

'***********************************************
sub vbsDbDrawGridTableBegin( objVbsDb )
'***********************************************
  if objVbsDb( "GridTemplate" ) = "" then
    vbsDbDrawStandardGridTableBegin objVbsDb
  else
    vbsDbDrawTemplateGridTableBegin objVbsDb    
  end if
end sub

'***********************************************
sub vbsDbDrawGridIndexHeader( byRef objVbsDb )
'***********************************************
  if objVbsDb( "GridShowIndex" ) then
  	' the grid index column has to be shown
%>
  	<th class="GridIndex">&nbsp;</th>
<%
  end if
end sub

'***********************************************
function funBolSqlFieldIsSortable( strFieldName , objVbsDb )
'***********************************************
  funBolSqlFieldIsSortable = ( ( not ( funVarSqlRecordSetFieldType( strFieldName , objVbsDb ) = vbsDbAdLongVarChar ) ) and _
                               ( not ( funVarSqlRecordSetFieldType( strFieldName , objVbsDb ) = vbsDbAdLongVarWChar ) ) )
end function

'***********************************************
function funStrGridHeaderAHRef( strFieldName , objVbsDb )
'***********************************************
  funStrGridHeaderAHRef = request.serverVariables( "url" ) &_
  	"?vbsDbGridSort_" & objVbsDb( "GlobalId" ) & "=" & strFieldName
end function

'***********************************************
sub vbsDbDrawGridHeadersCycleDrawCurrentHeaderSortable( strFieldName , objVbsDb )
'***********************************************
%>
  <th class="GridHeader">
  	<a href="<%=funStrVbsDbGlobalQuerystringPreserve( funStrGridHeaderAHRef( strFieldName , objVbsDb ), objVbsDb )%>"
  	  class="GridHeader">
  	  	<%=funStrFieldHeader( objVbsDb( "DictionaryGridFields" )( strFieldName ) , objVbsDb )%></a></th>
<%
end sub

'***********************************************
sub vbsDbDrawGridHeadersCycleDrawCurrentHeaderNonSortable( strFieldName , objVbsDb )
'***********************************************
%>
  <th class="GridHeader">
  	<%=funStrFieldHeader( objVbsDb( "DictionaryGridFields" )( strFieldName ) , objVbsDb )%></th>
<%
end sub

'***********************************************
sub vbsDbDrawGridHeadersCycleDrawCurrentHeader( strFieldName , objVbsDb )
'***********************************************
  if funBolSqlFieldIsSortable( strFieldName , objVbsDb ) then
    ' strFieldName is not a memo field
    vbsDbDrawGridHeadersCycleDrawCurrentHeaderSortable strFieldName , objVbsDb
  else
    ' strFieldName is a memo field
    vbsDbDrawGridHeadersCycleDrawCurrentHeaderNonSortable strFieldName , objVbsDb
  end if
end sub

'***********************************************
sub vbsDbDrawGridHeadersCycle( byRef objVbsDb )
'***********************************************
  dim strFieldName
  for each strFieldName in objVbsDb( "DictionaryGridFields" )
  	vbsDbDrawGridHeadersCycleDrawCurrentHeader strFieldName , objVbsDb
  next
end sub

'***********************************************
sub vbsDbDrawGridHeaders( byRef objVbsDb )
'***********************************************
%>
  <tr id=GridHeader>
<%
  	vbsDbDrawGridEditColumnHeader objVbsDb
  	vbsDbDrawGridIndexHeader objVbsDb
  	vbsDbDrawGridHeadersCycle objVbsDb
%>
  </tr>
<%
end sub

'***********************************************
function funStrGridRowAHRef( byRef objVbsDb )
'***********************************************
  funStrGridRowAHRef = request.serverVariables( "url" ) &_
  	"?VBSdbClickClass_" & objVbsDb( "GlobalId" ) & "=VBSdbGoToGridRow" &_
  	"&VBSdbIndex_" & objVbsDb( "GlobalId" ) & "=" &_
  	objVbsDb( "SqlRecordSet" ).absolutePosition
end function

'***********************************************
function funBolDrawGridCurrentRecordIsSelected( objVbsDb )
'***********************************************
  funBolDrawGridCurrentRecordIsSelected = _
    objVbsDb( "SqlRecordSet" ).absolutePosition = objVbsDb( "FormAbsolutePosition" )
end function

'***********************************************
function funStrCSSAttributeForGridIndexTd( objVbsDb )
'***********************************************
  if funBolDrawGridCurrentRecordIsSelected( objVbsDb ) then
  	' the current Grid record is the one that will be displayed in the detailed Form
  	funStrCSSAttributeForGridIndexTd = "class='GridSelectedIndex'"
  else
  	' the current Grid record is not the one that will be displayed in the detailed Form
  	funStrCSSAttributeForGridIndexTd = "class='GridUnselectedIndex'"
  end if
end function

'***********************************************
function funStrCSSAttributeForGridIndexAnchor( objVbsDb )
'***********************************************
  if funBolDrawGridCurrentRecordIsSelected( objVbsDb ) then
  	' the current Grid record is the one that will be displayed in the detailed Form
  	funStrCSSAttributeForGridIndexAnchor = "class='GridSelectedIndex'"
  else
  	' the current Grid record is not the one that will be displayed in the detailed Form
  	funStrCSSAttributeForGridIndexAnchor = "class='GridUnselectedIndex'"
  end if
end function

'***********************************************
sub vbsDbDrawGridRowFirstCellActually( byRef objVbsDb )
'***********************************************
%>
	<td <%=funStrCSSAttributeForGridIndexTd( objVbsDb )%>>
		<a href="<%=funStrVbsDbGlobalQuerystringPreserve( funStrGridRowAHRef( objVbsDb ) , objVbsDb )%>"
		<%=funStrCSSAttributeForGridIndexAnchor( objVbsDb )%>>
			<%=objVbsDb( "SqlRecordSet" ).absolutePosition%></a></td>			
<%
end sub

'***********************************************
sub vbsDbDrawGridRowFirstCell( byRef objVbsDb )
'***********************************************
  if 	objVbsDb( "GridShowIndex" ) then
  	' the grid index column has to be shown
  	vbsDbDrawGridRowFirstCellActually objVbsDb
  end if
end sub

'***********************************************
'function funStrDbColorHorizontalStripeColor( byRef objVbsDb )
'***********************************************
'  if objVbsDb( "GridHorizontalStripeBGColor" ) <> "" then
  	' objVbsDb( "GridHorizontalStripeBGColor" ) was defined
'  	funStrDbColorHorizontalStripeColor = objVbsDb( "GridHorizontalStripeBGColor" )
'  else
  	' objVbsDb( "GridHorizontalStripeBGColor" ) was defined
'  	funStrDbColorHorizontalStripeColor = objVbsDb( "GlobalTableBGColor" )
'  end if	
'end function

'***********************************************
'***********************************************
function funBolOddGridRow( objVbsDb )
'***********************************************
'***********************************************
  funBolOddGridRow = ( ( ( objVbsDb( "SqlRecordSet" ).absolutePosition - _
    ( objVbsDb( "GridAbsolutePage" ) - 1 ) * objVbsDb( "GridPageSize" ) ) mod 2 ) = 1 )
end function

'***********************************************
'function funStrDbDrawGridRowCurrentCellBGColor( objVbsDb )
'***********************************************
'  if funBolOddGridRow( objVbsDb ) then
   	' it is an odd row for the current grid page
'  	funStrDbDrawGridRowCurrentCellBGColor = objVbsDb( "GlobalTableBGColor" )
' else
  	' it is an even row for the current grid page
'  	funStrDbDrawGridRowCurrentCellBGColor = funStrDbColorHorizontalStripeColor( objVbsDb )
'  end if
'end function

'***********************************************
function funStrDbDrawCSSClassGridRow( objVbsDb )
'***********************************************
  if funBolOddGridRow( objVbsDb ) and funBolDrawGridCurrentRecordIsSelected( objVbsDb ) then
   	' it is an odd row for the current grid page and it is the selected row
  	funStrDbDrawCSSClassGridRow = "GridOddRowSelected"
 elseif funBolOddGridRow( objVbsDb ) and ( not funBolDrawGridCurrentRecordIsSelected( objVbsDb ) ) then
   	' it is an odd row for the current grid page and it is not the selected row
  	funStrDbDrawCSSClassGridRow = "GridOddRowUnselected"
 elseif ( not funBolOddGridRow( objVbsDb ) ) and ( funBolDrawGridCurrentRecordIsSelected( objVbsDb ) ) then
   	' it is an even row for the current grid page and it is the selected row
  	funStrDbDrawCSSClassGridRow = "GridEvenRowSelected"
 elseif ( not funBolOddGridRow( objVbsDb ) ) and ( not funBolDrawGridCurrentRecordIsSelected( objVbsDb ) ) then
   	' it is an even row for the current grid page and it is not the selected row
  	funStrDbDrawCSSClassGridRow = "GridEvenRowUnselected"
  end if
end function

'***********************************************
function funStrFieldValueForViewScreenActually( strFieldName , objVbsDb )
'***********************************************
  if funBolSqlFieldIsBoolean( strFieldName , objVbsDb ) then
    ' current sql recordset fiel is of boolean type
    funStrFieldValueForViewScreenActually = funStrGlobalBooleanText( strFieldName , objVbsDb )
  else
    ' current sql recordset fiel is not of boolean type
    funStrFieldValueForViewScreenActually = funStrVbsDbViewFieldFormat( strFieldName , objVbsDb )
  end if
end function

'***********************************************
'***********************************************
function funStrFieldValueForViewScreen( strFieldName , objVbsDb )
'***********************************************
'***********************************************
  funStrFieldValueForViewScreen = funStrFieldValueForViewScreenActually( strFieldName , objVbsDb )
  if funStrFieldValueForViewScreen = "" then
    ' the field value is the empty string
    funStrFieldValueForViewScreen = "&nbsp;"
  end if
end function

'***********************************************
sub vbsDbDrawGridRowCurrentCell( strFieldName , objVbsDb )
'***********************************************
%>
  <td class="<%=funStrDbDrawCSSClassGridRow( objVbsDb )%>">
  	<%=funStrFieldValueForViewScreen( strFieldName , objVbsDb )%>
  </td>
<%
end sub

'***********************************************
sub vbsDbDrawGridRowFieldCells( byRef objVbsDb )
'***********************************************
  dim strFieldName
  for each strFieldName in objVbsDb( "DictionaryGridFields" )
  	vbsDbDrawGridRowCurrentCell strFieldName , objVbsDb
  next
end sub

'***********************************************
sub vbsDbDrawStandardGridRow( byRef objVbsDb )
'***********************************************
%>
  <tr class="<%=funStrDbDrawCSSClassGridRow( objVbsDb )%>">
<%
  	vbsDbDrawGridRowEditButtons objVbsDb
  	vbsDbDrawGridRowFirstCell objVbsDb
  	vbsDbDrawGridRowFieldCells objVbsDb
%>
  </tr>
<%
end sub

'***********************************************
sub vbsDbDrawGridRow( byRef objVbsDb )
'***********************************************
  if objVbsDb( "GridTemplate" ) = "" then
    ' the user wants a standard grid
    vbsDbDrawStandardGridRow objVbsDb
  else
    ' the user wants a customized grid
    vbsDbDrawTemplateGridRow objVbsDb
  end if
end sub

'***********************************************
sub vbsDbDrawGridRows( byRef objVbsDb )
'***********************************************
  if not ( objVbsDb( "SqlRecordSet" ).recordCount=0 ) then
    ' the query returned at least one record
    objVbsDb( "SqlRecordSet" ).absolutePage = objVbsDb( "GridAbsolutePage" )
  end if
  while ( not objVbsDb( "SqlRecordSet" ).eof ) and _
    ( objVbsDb( "SqlRecordSet" ).absolutePosition < ( objVbsDb( "GridAbsolutePage" ) * objVbsDb( "GridPageSize" ) + 1 ) )
  	vbsDbDrawGridRow objVbsDb
  	objVbsDb( "SqlRecordSet" ).moveNext
  wend
end sub

'***********************************************
sub vbsDbDrawGridTableBody( byRef objVbsDb )
'***********************************************
  if objVbsDb( "GridTemplate" ) = "" then
    ' the user wants a standard grid
    vbsDbDrawGridHeaders objVbsDb
  end if
  vbsDbDrawGridRows objVbsDb
end sub

'***********************************************
sub vbsDbDrawGridTableEndActually()
'***********************************************
%>
  </table>
<%
end sub

'***********************************************
sub vbsDbDrawStandardGridTableEnd()
'***********************************************
  vbsDbDrawGridTableEndActually
  vbsDbEndDrawForBorders
end sub

'***********************************************
sub vbsDbDrawGridTableEnd( objVbsDb )
'***********************************************
  if objVbsDb( "GridTemplate" ) = "" then
    vbsDbDrawStandardGridTableEnd
  else
    vbsDbDrawTemplateGridTableEnd    
  end if
end sub

'***********************************************
sub vbsDbDrawGrid( byRef objVbsDb )
'***********************************************
  vbsDbDrawGridTableBegin objVbsDb
  vbsDbDrawGridTableBody objVbsDb
  vbsDbDrawGridTableEnd objVbsDb
end sub
%>
<%
'***********************************************
function funIntFirstRecordInThePage( objVbsDb )
'***********************************************
  funIntFirstRecordInThePage = objVbsDb( "SqlRecordSet" ).pageSize * _
    ( objVbsDb( "SqlRecordSet" ).absolutePage - 1 ) + 1
end function

'***********************************************
function funIntLastRecordInThePage( objVbsDb )
'***********************************************
  funIntLastRecordInThePage = funIntFirstRecordInThePage( objVbsDb ) + _
    objVbsDb( "SqlRecordSet" ).pageSize - 1
  funIntLastRecordInThePage = funNumMin( funIntLastRecordInThePage , _
                                   objVbsDb( "SqlRecordSet" ).recordCount )
end function

'***********************************************
function funStrViewPosition( intFirstRecordInThePage , intLastRecordInThePage , intRecordCount , intAbsolutePage , intPageCount , objVbsDb )
'***********************************************
  funStrViewPosition = funStrTranslate( "viewPosition" , objVbsDb )
  funStrViewPosition = replace( funStrViewPosition , "<FirstRecordInThePage>" , intFirstRecordInThePage , _
    1 , -1 , 1 )
  funStrViewPosition = replace( funStrViewPosition , "<LastRecordInThePage>" , intLastRecordInThePage , _
    1 , -1 , 1 )
  funStrViewPosition = replace( funStrViewPosition , "<TotalRecords>" , intRecordCount , _
    1 , -1 , 1 )
  funStrViewPosition = replace( funStrViewPosition , "<CurrentPage>" , intAbsolutePage , _
    1 , -1 , 1 )
  funStrViewPosition = replace( funStrViewPosition , "<TotalPages>" , intPageCount , _
    1 , -1 , 1 )
end function

'***********************************************
sub vbsDbDrawViewPositionHTML( intFirstRecordInThePage , intLastRecordInThePage , intRecordCount , intCurrentPage , intLastPage , objVbsDb )
'***********************************************
  vbsDbBeginDrawForBorders "ViewPositionBorders"
%>
  <table cellspacing="1" cellpadding="2" border="0">
    <tr>
      <td id="ViewPosition">
        <%=funStrViewPosition( intFirstRecordInThePage , intLastRecordInThePage , intRecordCount , intCurrentPage , intLastPage , objVbsDb )%>
      </td>
    </tr>
  </table>
<%
  vbsDbEndDrawForBorders
end sub

'***********************************************
sub vbsDbDrawViewPositionActually( objVbsDb )
'***********************************************
  dim intFirstRecordInThePage , intLastRecordInThePage , intRecordCount , intCurrentPage , intLastPage
  intFirstRecordInThePage = funIntFirstRecordInThePage( objVbsDb )
  intLastRecordInThePage = funIntLastRecordInThePage( objVbsDb )
  intCurrentPage = objVbsDb( "SqlRecordSet" ).absolutePage
  intLastPage = objVbsDb( "SqlRecordSet" ).pageCount
  intRecordCount = objVbsDb( "SqlRecordSet" ).recordCount
  vbsDbDrawViewPositionHTML intFirstRecordInThePage , intLastRecordInThePage , intRecordCount , intCurrentPage , intLastPage , objVbsDb
end sub

'***********************************************
'***********************************************
sub vbsDbDrawViewPosition( objVbsDb )
'***********************************************
'***********************************************
  if objVbsDb( "ViewPosition" ) and ( objVbsDb( "SqlRecordSet" ).recordCount > 0 ) then
    objVbsDb( "SqlRecordSet" ).absolutePosition = objVbsDb( "FormAbsolutePosition" )
    vbsDbDrawViewPositionActually objVbsDb
  end if
end sub
%>
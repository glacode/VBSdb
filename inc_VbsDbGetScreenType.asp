<%
'***********************************************
function funStrGetScreenTypeHandlingGlobalStartScreen( objVbsDb )
'***********************************************
  if objVbsDb( "SearchScreenRepeat" ) then
    ' the user wants this VBSdb object to stay in search screen
    funStrGetScreenTypeHandlingGlobalStartScreen = "Search"
  elseif objVbsDb( "EditAddSingle" ) then
    ' the user wants this VBSdb object to add a single record
    funStrGetScreenTypeHandlingGlobalStartScreen = "Add"
  else
    select case uCase( objVbsDb( "GlobalStartScreen" ) )
      case "VIEW"
        funStrGetScreenTypeHandlingGlobalStartScreen = "View"
      case "SEARCH"
        funStrGetScreenTypeHandlingGlobalStartScreen = "Search"
      case "ADD"
        funStrGetScreenTypeHandlingGlobalStartScreen = "Add"
      case "UPDATE"
        funStrGetScreenTypeHandlingGlobalStartScreen = "Update"
      case "DELETE"
        funStrGetScreenTypeHandlingGlobalStartScreen = "Delete"
      case else
        funStrGetScreenTypeHandlingGlobalStartScreen = "View"
    end select
    'funStrGetScreenTypeHandlingGlobalStartScreen = objVbsDb( "GlobalStartScreen" )
  end if
end function

'***********************************************
function funStrGetScreenTypeReturningHere( objVbsDb )
'***********************************************
  if ( not objVbsDb( "SearchScreenRepeat" ) ) and _
    ( not objVbsDb( "EditAddSingle" ) ) and _
    ( ( inStr( "VBSdbGoToGridRow**VBSdbGridFirst**VBSdbGridPrev**VBSdbGridNext**VBSdbGridLast**VBSdbRemoveFilter**VBSdbApplyFilter**VBSdbApplyUpdate**VBSdbApplyDelete**VBSdbCancel" , _
               vbsDbGetClickClass( objVbsDb ) ) > 0 ) or _
    ( ( not objVbsDb( "EditAddScreenRepeat" ) ) and  _
    ( vbsDbGetClickClass( objVbsDb ) = "VBSdbApplyAdd" ) ) ) then
    funStrGetScreenTypeReturningHere = "View"
  elseif ( inStr( vbsDbGetClickClass( objVbsDb ) , "VBSdbFilter" ) > 0 ) or _
    objVbsDb( "SearchScreenRepeat" ) then
    funStrGetScreenTypeReturningHere = "Search"
  elseif ( inStr( vbsDbGetClickClass( objVbsDb ) , "VBSdbEditAdd" ) > 0 ) _
    or ( objVbsDb( "EditAddScreenRepeat" ) and  vbsDbGetClickClass( objVbsDb ) = "VBSdbApplyAdd" ) _
    or ( objVbsDb( "EditAddSingle" ) and  vbsDbGetClickClass( objVbsDb ) <> "VBSdbApplyAdd" ) then
    funStrGetScreenTypeReturningHere = "Add"
  elseif inStr( vbsDbGetClickClass( objVbsDb ) , "VBSdbEditUpdate" ) > 0 then
    funStrGetScreenTypeReturningHere = "Update"
  elseif inStr( vbsDbGetClickClass( objVbsDb ) , "VBSdbEditDelete" ) > 0 then
    funStrGetScreenTypeReturningHere = "Delete"
  else
    ' EditAddSingle = true and Add just submitted
    funStrGetScreenTypeReturningHere = "NoScreen"
  end if
end function

'***********************************************
'***********************************************
function VbsDbGetScreenType( objVbsDb )
'***********************************************
'***********************************************
  if funBolJustArrived( objVbsDb ) then
  	' the last page was not the current one
  	VbsDbGetScreenType = funStrGetScreenTypeHandlingGlobalStartScreen( objVbsDb )
  else
  	' the last page was the current one
  	VbsDbGetScreenType = funStrGetScreenTypeReturningHere( objVbsDb )
  end if
end function
%>
<%
'***********************************************
'***********************************************
function VbsDbGetLastAction( objVbsDb )
'***********************************************
'***********************************************
  select case uCase( vbsDbGetClickClass( objVbsDb ) )
    case uCase( "VBSdbGoToGridRow" )
      VbsDbGetLastAction = "GoToGridRow"
    case uCase( "VBSdbGridFirst" )
      VbsDbGetLastAction = "GoToFirstGridPage"
    case uCase( "VBSdbGridPrev" )
      VbsDbGetLastAction = "GoToPrevGridPage"
    case uCase( "VBSdbGridNext" )
      VbsDbGetLastAction = "GoToNextGridPage"
    case uCase( "VBSdbGridLast" )
      VbsDbGetLastAction = "GoToLastGridPage"
    case uCase( "VBSdbRemoveFilter" )
      VbsDbGetLastAction = "RemoveFilter"
    case uCase( "VBSdbFilter" )
      VbsDbGetLastAction = "GoToSearch"
    case uCase( "VBSdbEditAdd" )
      VbsDbGetLastAction = "GoToAdd"
    case uCase( "VBSdbEditUpdate" )
      VbsDbGetLastAction = "GoToUpdate"
    case uCase( "VBSdbEditDelete" )
      VbsDbGetLastAction = "GoToDelete"
    case uCase( "VBSdbCancel" )
      VbsDbGetLastAction = "Cancel"
    case uCase( "VBSdbApplyFilter" )
      VbsDbGetLastAction = "SubmitSearch"
    case uCase( "VBSdbApplyAdd" )
      VbsDbGetLastAction = "SubmitAdd"
    case uCase( "VBSdbApplyUpdate" )
      VbsDbGetLastAction = "SubmitUpdate"
    case uCase( "VBSdbApplyDelete" )
      VbsDbGetLastAction = "SubmitDelete"
    case ""
      VbsDbGetLastAction = "NonVbsDb"
  end select
end function
%>
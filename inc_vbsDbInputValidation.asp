<%
'***********************************************
'***********************************************
function funBolJavascriptValidation( byRef objVbsDb )
'***********************************************
'***********************************************
  if ( VbsDbGetScreenType( objVbsDb ) = "Search" ) then
    ' search screen
    funBolJavascriptValidation = ( objVbsDb( "InputValidateDateFields" ) <> "" )
  elseif ( ( VbsDbGetScreenType( objVbsDb ) = "Add" ) or ( VbsDbGetScreenType( objVbsDb ) = "Update" ) ) then
    ' edit screen
    funBolJavascriptValidation = ( objVbsDb( "InputValidateDateFields" ) <> "" ) or _
                                 ( objVbsDb( "EditValidateRequiredFields" ) <> "" )
  end if
end function

'***********************************************
sub drawJavascriptFunValidateNumeric()
'***********************************************
end sub

'***********************************************
sub drawJavascriptFunValidateEMail()
'***********************************************
end sub

'***********************************************
sub drawValidateJavascriptFunctionActuallyDrawValidateFunctions( objVbsDb )
'***********************************************
  drawVbsDbValidateRequired objVbsDb
  drawVbsDbValidateRegExp objVbsDb
  drawJavascriptFunValidateNumeric
  drawJavascriptFunValidateEMail
  drawJavascriptFunValidateDate objVbsDb
end sub

'***********************************************
sub drawValidateFunctionCallLines( objVbsDb )
'***********************************************
  drawValidateFunctionCallLinesForEditValidateRequiredFields objVbsDb
  drawValidateFunctionCallLinesForInputValidateDateFields objVbsDb
  drawValidateFunctionCallLinesForEditValidateRegExp objVbsDb
end sub

'***********************************************
sub drawValidateJavascriptFunctionActuallyDrawValidateFunctionCalls( objVbsDb )
'***********************************************
%>
  function EditValidate(form)
  {
  	var message = ""
  	var more_message = ""
  	var showmsg = "no"
  	var x
  	x = form.elements.length
  	x = x - 1
<%
    drawValidateFunctionCallLines objVbsDb
%>
  	//This code will prevent a submit if data is incoorect
  	if (message > "")
  		{
  			// originale vecchio alert("The following form field(s) were incomplete or incorrect:\n\n" + message + "\n\n Please complete or correct the form and submit again.")
  			alert("<%=funStrTranslate( "javascriptValidationBegin" , objVbsDb)%>" + "\n\n" + message + "\n" + "<%=funStrTranslate( "javascriptValidationEnd" , objVbsDb)%>")
  		}
  	else
  		{
  			form.submit()
  		}
  }
<%
end sub

'***********************************************
sub drawValidateJavascriptFunctionActually( objVbsDb )
'***********************************************
%>
  <script language="javascript">
  var new_fieldname = ""
<%
  drawValidateJavascriptFunctionActuallyDrawValidateFunctions objVbsDb
  drawValidateJavascriptFunctionActuallyDrawValidateFunctionCalls objVbsDb
%>
  // End -->
  </script>
<%
end sub

'***********************************************
'***********************************************
sub drawValidateJavascriptFunction( objVbsDb )
'***********************************************
'***********************************************
  if funBolJavascriptValidation( objVbsDb ) then
    ' there is at least an input field to be validated
    drawValidateJavascriptFunctionActually objVbsDb
  end if
end sub

%>
<%
'----------------------------------------------------------------------
' Module	 : KudzuLib_SPACING.asp - KudzuASP SPACING Library
' Author	 : Andrew F. Friedl @ TriLogic Industries, LLC
' Created	 : 2011.02.17
' Revised	 : 2011.02.17
' Version	 : 1.0.0
' Copyright: 2006-2011 TriLogic Industries, LLC
' License  : Full license is granted for personal or commercial use
'          : as long as this header remains intact.
'----------:-----------------------------------------------------------
'          : Oh Mary conceived without sin,
'          : pray for use who have recourse to thee.
'----------------------------------------------------------------------

'----------------------------------------------------------------------
' Tag Handlers
'----------------------------------------------------------------------
Class CTPSpace
	Public Sub HandleTag(vNode)
		Dim Count: Count = 1
		If vNode.ParamCount > 0 Then
			Count = CLng(vNode.ParamItem(1))
		End If
		vNode.Engine.ContentAppend String(Count, " ")
	End Sub
End Class
Class CTPTab
	Public Sub HandleTag(vNode)
		Dim Count: Count = 1
		If vNode.ParamCount > 0 Then
			Count = CLng(vNode.ParamItem(1))
		End If
		vNode.Engine.ContentAppend String(Count, vbTab)
	End Sub
End Class
Class CTPCrlf
	Public Sub HandleTag(vNode)
		Dim Count: Count = 1
		If vNode.ParamCount > 0 Then
			Count = CLng(vNode.ParamItem(1))
		End If
		vNode.Engine.ContentAppend String(Count, vbCrLf)
	End Sub
End Class

'----------------------------------------------------------------------
' Library Import Subroutine
'----------------------------------------------------------------------
Sub KudzuLibImport_SPACING(libName)
	Dim thisLib: Set thisLib = KudzuLib.libGet(libName)
	thisLib.SetTag "space", New CTPSpace
	thisLib.SetTag "tab", New CTPTab
	thisLib.SetTag "crlf", New CTPCrLf
End Sub
%>

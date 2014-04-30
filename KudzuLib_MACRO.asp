<%
'----------------------------------------------------------------------
' Module	: KudzuLib_MACRO.asp - KudzuASP MACRO Library
' Author	: Andrew F. Friedl @ TriLogic Industries, LLC
' Created	: 2011.02.17
' Revised	: 2011.02.17
' Version	: 1.0.0
' Copyright : 2006-2011 TriLogic Industries, LLC
' License   : Full license is granted for personal or commercial use
'           : as long as this header remains intact.
'-----------:----------------------------------------------------------
'           : Oh Mary conceived without sin,
'           : pray for use who have recourse to thee.
'----------------------------------------------------------------------

'----------------------------------------------------------------------
' Tag Handlers
'----------------------------------------------------------------------
Class CTPMacro
	Public Sub HandleTag(vNode)
		Dim idx, qty: qty=0
		If vNode.ParamCount < 1 Then
			vNode.AppendError "macro|replayQty]"
		End If
		vNode.Engine.PutValue vNode.ParamItem(1), vNode
		If vNode.ParamCount = 2 Then qty = CInt(vNode.ParamItem(2))
		For idx = 1 To qty
			vNode.StackPush
			vNode.EvalNodes
			vNode.StackPop
		Next
	End Sub
End Class

Class CTPReplay
	Public Sub HandleTag(vNode)
		Dim idx, qty: qty = 1
		If vNode.ParamCount < 1 Then
			vNode.AppendError "macro[|replayQty]"
		End If
		If vNode.ParamCount = 2 Then qty = CInt(vNode.ParamItem(2))
		Set mNode = vNode.Engine.GetObjectValue(vNode.ParamItem(1))
		For idx = 1 To qty
			mNode.StackPush
			mNode.EvalNodes
			mNode.StackPop
		Next
	End Sub
End Class

'----------------------------------------------------------------------
' Library Import Subroutine
'----------------------------------------------------------------------
Sub KudzuLibImport_MACRO(libName)
	Dim thisLib: Set thisLib = KudzuLib.libGet(libName)
	thisLib.SetTag "macro", New CTPMacro
	thisLib.SetTag "replay", New CTPReplay
End Sub
%>
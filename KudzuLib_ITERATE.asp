<%
'----------------------------------------------------------------------
' Module	 : KudzuLib_Iterate.asp - KudzuASP Interation Library
' Author	 : Andrew F. Friedl @ TriLogic Industries, LLC
' Created	 : 2011.02.17
' Revised	 : 2011.02.17
' Version	 : 1.0.0
' Copyright: 2006-2011 TriLogic Industries, LLC
' License  : Full license is granted for personal or commercial use
'          : as long as this header remains intact.
'----------:-----------------------------------------------------------
'          : Oh Mary conceived without sin,
'          : pray for us who have recourse to thee.
'----------------------------------------------------------------------

'----------------------------------------------------------------------
' Iterators Tag Handlers
'----------------------------------------------------------------------
Class CTPForArray
	Sub HandleTag( vNode )
		Dim Idx, vParam, sFlag, oIter
		If vNode.ParamCount < 2 Then
			vNode.AppendTagError "value_id|iterator_id"
			Exit Sub
		End If
		If Not vNode.Engine.HasIterator( vNode.ParamItem(2) ) Then
			vNode.AppendTagError "undefined iterator (" & vNode.ParamItem(2) & ")"
			Exit Sub
		End If
		Set oIter = vNode.Engine.GetIterator( vNode.ParamItem(2) )
		vParam = vNode.EvalParamString( vNode.ParamItem(1) )
		If Not IsArray(vParam) Then
			vNode.AppendTagError "not an array (" & vNode.ParamItem(1) & ")"
			Exit Sub
		End If
		vNode.StackPush
		For Idx = 1 To UBound(vParam)
			If Not oIter.GoItem(vNode,vParam(Idx)) Then
				Exit For
			End If
			vNode.StackPush
			vNode.EvalNodes
			vNode.StackPop
		Next
		vNode.StackPop
	End Sub
End Class

Class CTPForEach
	Sub HandleTag( vNode )
		Dim vItem, vParam, sFlag, oIter
		If vNode.ParamCount < 2 Then
			vNode.AppendTagError "value_id|iterator_id"
			Exit Sub
		End If
		If Not vNode.Engine.HasIterator( vNode.ParamItem(2) ) Then
			vNode.AppendTagError "undefined iterator (" & vNode.ParamItem(2) & ")"
			Exit Sub
		End If
		Set oIter = vNode.Engine.GetIterator( vNode.ParamItem(2) )
		vNode.StackPush
	    On Error Resume Next
		For Each vItem In vNode.EvalParamObject( vNode.ParamItem(1) )
			If Err.Number <> 0 Then
				vNode.AppendTagError Err.Description
				Exit For
			End If
			If Not oIter.GoItem( vNode, vItem ) Then
				Exit For
			End If
			vNode.StackPush
			vNode.EvalNodes
			vNode.StackPop
		Next
		vNode.StackPop
	End Sub
End Class

Class CTPIterate
	Sub HandleTag( vNode )
		If vNode.ParamCount < 1 Then
			vNode.AppendTagError "iterator_id"
			Exit Sub
		End If
		If Not vNode.Engine.HasIterator( vNode.ParamItem(1) ) Then
			vNode.AppendTagError "undefined iterator (" & vNode.ParamItem(1) & ")"
			Exit Sub
		End If		
		Dim oIter, bMore, bUseTag
		Set oIter = vNode.Engine.GetIterator( vNode.ParamItem(1) )
		bMore = oIter.GoFirst( vNode )
		vNode.StackPush
		While bMore
			vNode.StackPush
			vNode.EvalNodes
			vNode.StackPop
			bMore = oIter.GoNext( vNode )
		Wend
		vNode.StackPop
	End Sub
End Class

'----------------------------------------------------------------------
' Predefined Iterators - Iterate Command
'----------------------------------------------------------------------
Class CIterateOverI
	Function GoFirst( vNode )
		GoFirst = GoFirstI( vNode )
	End Function
	Function GoNext( vNode )
		GoNext = GoNextI( vNode )
	End Function
End Class
Class CIterateOverJ
	Function GoFirst( vNode )
		GoFirst = GoFirstJ( vNode )
	End Function
	Function GoNext( vNode )
		GoNext = GoNextJ( vNode )
	End Function
End Class
Class CIterateOverK
	Function GoFirst( vNode )
		GoFirst = GoFirstK( vNode )
	End Function
	Function GoNext( vNode )
		GoNext = GoNextK( vNode )
	End Function
End Class
Class CIterateERR
	Function GoFirst( vNode )
		GoFirst = GoFirstERR( vNode )
	End Function
	Function GoNext( vNode )
		GoNext = GoNextERR( vNode )
	End Function
End Class
Class CIterateNIL
	Function GoFirst( vNode )
		GoFirst = False
	End Function
End Class

'----------------------------------------------------------------------
' Predefinied Iterators - ForArray & ForEach Command
'----------------------------------------------------------------------
Class CForItemI
	Function GoItem( vNode, vParam )
		GoItem = OnItemI( vNode, vParam )
	End Function
End Class
Class CForItemJ
	Function GoItem( vNode, vParam )
		GoItem = OnItemJ( vNode, vParam )
	End Function
End Class
Class CForItemK
	Function GoItem( vNode, vParam )
		GoItem = OnItemK( vNode, vParam )
	End Function
End Class

'----------------------------------------------------------------------
' Library Import Subroutine
'----------------------------------------------------------------------
Sub KudzuLibImport_ITERATE(libName)
	Dim thisLib: Set thisLib = KudzuLib.libGet(libName)
	thisLib.SetTag "ForArray", New CTPForArray
	thisLib.SetTag "ForEach", New CTPForEach
	thisLib.SetTag "Iterate", New CTPIterate
End Sub
%>

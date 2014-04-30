<%
'----------------------------------------------------------------------
' Module	 : KudzuLib_TREE.asp - KudzuASP TREE Library
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
'
' This component exposes the ability of the template engine to generate a tree of
' data nodes using the templated content for output.  Upon execution this component
' retrieves the named value from the engine which must be of Type 'CTPTree'.  For each
' node in the tree, this componet will attempt to evaluate a templated node corresponding
' to the tree node type.
' WARNING: Self referential trees will cause an infinite loop.
'
Class CTPTree
	Sub HandleTag( vNode )
		Dim vTree
		If vNode.ParamCount < 1 Then
			vNode.AppendTagError "tree value?"
			Exit Sub
		End If
		On Error Resume Next
		Set vTree = vNode.Engine.GetObjectValue(vNode.ParamItem(1))
		If vTree.NodeCount = 0 Then
			EvalNode vNode, "EmptyTree"
			vNode.StackPop
			Exit Sub
		End If
		EvalTreeNode vNode, vTree
		If Err.Number <> 0 Then
			vNode.AppendTagError Err.Description
		End If
	End Sub
	Sub EvalTreeNode( vNode, vTreeNode )
		Dim Idx, tNode
		vNode.StackPush
		If LCase(vTreeNode.NodeType) = "tree" Then
			EvalTreeNodeAttr vNode, vTreeNode
			EvalNode vNode, "header"
			EvalTreeChildNodes vNode, vTreeNode
			EvalTreeNodeAttr vNode, vTreeNode
			EvalNode vNode, "footer"
		Else
			Set tNode = vNode.LocateNode( vTreeNode.NodeType )
			If tNode Is Nothing Then
				vNode.AppendTagError "missing tree node(" & vTreeNode.NodeType & "," & vTree.NodeType & ")"
			Else
				EvalTreeNodeAttr vNode, vTreeNode
				EvalNode tNode, "header"
				EvalNode tNode, "Body"
				EvalTreeChildNodes vNode, vTreeNode
				EvalTreeNodeAttr vNode, vTreeNode
				EvalNode tNode, "footer"				
			End If
		End If
		vNode.StackPop		
	End Sub
	Sub EvalTreeChildNodes( vNode, vTreeNode )
		Dim Idx
		For Idx = 1 To vTreeNode.NodeCount
			EvalTreeNode vNode, vTreeNode.Node(Idx)
		Next
	End Sub
	Sub EvalTreeNodeAttr( vNode, vTreeNode )
		Dim arrAttr, Idx
		arrAttr = vTreeNode.AttributeKeys()
		For Idx = 1 To UBound(arrAttr)
			vNode.Engine.PutValue arrAttr(Idx), vTreeNode.Attribute(arrAttr(Idx))
		Next
	End Sub
	Sub EvalNode( vNode, node_id )
		Dim mNode
		Set mNode = vNode.LocateNode( node_id )
		If mNode Is Nothing Then Exit Sub
		mNode.EvalNodes
	End Sub
End Class

'----------------------------------------------------------------------
' Support Code
'----------------------------------------------------------------------
Class CTPTreeNode
	Dim mType, mAttr, mNodes
	
	Sub Class_Initialize()
		mType = "TREE"
		Set mAttr = CreateObject("Scripting.Dictionary")
		mNodes = Array("")
	End Sub

	Property Get NodeType()
		NodeType = mType
	End Property
	Property Let NodeType(Value)
		mType = Value
	End Property

	Property Get Attributes()
		Set Attributes = mAttr
	End Property
	Property Get AttributeKeys()
		AttributeKeys = mAttr.Keys
	End Property
	Property Get HasAttribute(sName)
		HasAttribute = mAttr.Exists(sName)
	End Property
	Property Get Attribute(sName)
		If HasAttribute(LCase(sName)) Then
			Attribute = mAttr(sName)
		Else
			Attribute = ""
		End If
	End Property
	Property Let Attribute(sName,sValue)
		If mAttr.Exists(LCase(sName)) Then
			mAttr.Remove(LCase(sName))
		End If
		mAttr.Add LCase(sName),sValue
	End Property

	Property Get NodeCount()
		NodeCount = UBound(mNodes)
	End Property
	Property Get Nodes()
		Nodes = mNodes
	End Property
	Property Get Node(Index)
		Set Node = mNodes(Index)
	End Property
	Function AddNode( oNode )
		Redim Preserve mNodes(Ubound(mNodes)+1)
		Set mNodes(Ubound(mNodes)) = oNode
		Set AddNode = oNode
	End Function

	Function ToArray()
		Dim mResult, mArr, Idx
		mResult = Array(mType,mAttr,Array(""))
		mArr = Array("")
		Redim mArr(UBound(mNodes))
		For Idx = 1 To UBound(mNodes)
			mArr(Idx) = mNodes(Idx).ToArray()
		Next
		mResult(2) = mArr
		ToArray = mResult
	End Function
	Sub FromArray( mArray )
		Dim mArr, Idx, oNode
		mType(0) = mArray(0) ' type
		Set mAttr(1) = mArray(1) ' attributes
		mArr = mArray(2)
		Redim mNodes(Ubound(mArr))
		For Idx = 1 To Ubound(mArr)
			set mNodes(Idx) = New CTreeNode
			mNodes(Idx).FromArray mArr(Idx)
		Next
	End Sub

End Class

'----------------------------------------------------------------------
' Library Import Subroutine
'----------------------------------------------------------------------
Sub KudzuLibImport_TREE(libName)
	Dim thisLib: Set thisLib = KudzuLib.libGet(libName)
	thisLib.SetTag "Tree", New CTPTree
End Sub
%>

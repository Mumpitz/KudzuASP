<%
'----------------------------------------------------------------------
' Module	 : _kudzu.asp - Kudzu ASP Template Engine
' Author	 : Andrew F. Friedl @ TriLogic Industries, LLC
' Created	 : 2006.05.09
' Revised	 : 2014.04.30
' Version	 : 1.6.0
' Copyright: 2006-2014 TriLogic Industries, LLC
' License  : Full license is granted for personal or commercial use
'          : as long as this header remains intact.
'----------:-----------------------------------------------------------
'          : Oh Mary conceived without sin,
'          : pray for use who have recourse to thee.
'----------------------------------------------------------------------

'----------------------------------------------------------------------
' Kudzu Library Code
'----------------------------------------------------------------------
Dim KudzuLib: Set KudzuLib = New CKudzuLib

Function KudzuReadLib(PluginFileName)
  Dim fso, txt
  Set fso = Server.CreateObject("Scripting.FileSystemObject")
  Set txt = fso.OpenTextFile(PluginFileName)
  KudzuReadLib = txt.ReadAll
  txt.Close: Set txt=Nothing: Set fso=Nothing
End Function

Sub KudzuLoadLib(PluginFile)
	Dim strSource, objPlugin, lenSource
	strSource = KudzuReadLib(PluginFile)
	lenSource = Len(strSource)
	If Left(strSource,2) = CHR(60) & "%" Then
		strSource = Right(strSource,lenSource-2)
		lenSource = lenSource - 2
	End If
	If Right(strSource,2) = "%" & CHR(62) Then
		strSource = Left(strSource,lenSource-2)
	End If
	ExecuteGlobal strSource
End Sub

Class CKudzuLibItem
	Dim tags, libName, libDesc, libFile, libFunc, libVers
	Sub Class_Initialize()
		set tags = CreateObject("Scripting.Dictionary")
		libName = "": libDesc = "": libFile = ""
		libFunc = "": libVers = "0.0"
	End Sub
	Sub setTag(tag,obj)
		tag = UCase(tag)
		If tags.Exists(tag) Then tags.Remove tag
		tags.Add tag, obj
	End Sub
	Sub getTag(tag)
		tag = UCase(tag)
		If tags.Exists(tag) Then
			Set getTag = tags(tag)
		Else
			Set getTag = Nothing
		End If
	End Sub
	Sub libInit(sPath,sName)		
		libName = UCASE(sName)
		libFile = sPath & "KudzuLib_" & libName & ".asp"
		libFunc = "KudzuLibImport_" & libName
	End Sub
	Sub setTags(eng)
		Dim idx, keys: keys = tags.Keys
		For idx = 0 To Ubound(keys)
			eng.SetHandler keys(idx), tags(keys(idx))
		Next
	End Sub
	Sub libImport()
		KudzuLoadLib libFile
		ExecuteGlobal libFunc & """" & libName & """"
	End Sub
End Class

Class CKudzuLib
	Dim tagLibs, libPath, libCurr
	Sub Class_Initialize()
		Set tagLibs = CreateObject("Scripting.Dictionary")
		libPath = Array(Server.MapPath(".")&"\")
		Set libCurr = nothing
	End Sub
	Function libExists(libName)
		libExists = tagLibs.Exists(UCase(libName))
	End Function
	Function libFind(libName)
		Dim idx, pth, fso: Set fso = CreateObject("Scripting.FileSystemObject")
		libName = UCase(libName)
		For idx = 0 To Ubound(libPath)
			pth = libPath(idx) & "KudzuLib_" & libName & ".asp"
			If fso.FileExists(pth) Then
				libFind = libPath(idx)
				Exit Function
			End If
		Next
		libFind = ""
	End Function
	Function libImport(libName)
		Dim path
		If libExists(libName) Then
			Set libImport = libGet(libName)
			Exit Function
		End If
		libName = UCase(libName)
		path = libFind(libName)
		If path = "" Then 
			Set libImport = Nothing
			Exit Function
		End If
		Set libCurr = new CKudzuLibItem
		libCurr.libInit path, libName
		If tagLibs.Exists(libName) Then tagLibs.Remove libName
		tagLibs.Add libName, libCurr
		libCurr.libImport
		Set libImport = libCurr
	End Function
	Function libGet(libName)
		Set libGet = tagLibs(UCase(libName))
	End Function
	Function libSetTags(libName,eng)
		libImport libName
		Dim libObj: Set libObj = libGet(libName)
		If (libObj Is Nothing) Then
			libSetTags = False
		Else
			libObj.setTags(eng)
			libSetTags = True
		End If
	End Function
	Sub libPathPush(sPath)
		Redim Preserve libPath(UBound(libPath)+1)
		libPath(Ubound(libPath)) = sPath
	End Sub
	Sub libPathPop()
		If (Ubound(libPath)<1) Then Exit Sub
		Redim Preserve libPath(Ubound(libPath)-1)
	End Sub
	Sub setLibPath(sPath)
		libPath(0) = sPath
	End Sub
End Class

'----------------------------------------------------------------------
' Tag Handlers
'----------------------------------------------------------------------
Class CTPFlush
	Sub HandleTag(vNode)
		vNode.Engine.ContentFlush
	End Sub
End Class
Class CTPImport
	Sub HandleTag(vNode)
		If vNode.ParamCount < 1 Then
			vNode.AppendTagError "libName[|libName2]*"
			Exit Sub
		End If
		Dim idx
		For idx=1 To vNode.ParamCount
			vNode.Engine.libImport vNode.ParamItem(idx)
		Next
	End Sub
End Class
Class CTPProfiler
	Sub HandleTag(vNode)
		Dim timeSpan
		vNode.Engine.StopTime = Timer()
		timeSpan = vNode.Engine.StopTime - vNode.Engine.StartTime
		vNode.Engine.PutValue "PageTime", FormatNumber(timeSpan,4)
		vNode.StackPush
		vNode.EvalNodes
		vNode.Engine.ContentReplaceFields
		vNode.StackPop
	End Sub
End Class
Class CTPExecute
	Sub HandleTag(vNode)
		Dim sExpr
		If vNode.ParamCount < 1 Then
			vNode.AppendTagError "expr"
			Exit Sub
		End If
		sExpr = vNode.ParamItem(1)
		If Left(sExpr, 1) = "?" Then
			sExpr = Right(sExpr, Len(sExpr) - 1)
			sExpr = vNode.EvalParamString(sExpr)
		End If
		If sExpr = "" Then
			vNode.AppendTagError "empty url"
			Exit Sub
		End If
		On Error Resume Next
		vNode.Engine.ContentFlush
		vNode.StackPush
		Server.Execute sExpr
		If Err.Number <> 0 Then
			vNode.AppendTagError Err.Description
		End If
		vNode.StackPop
	End Sub
End Class
Class CTPIIf
	Sub HandleTag(vNode)
		If vNode.ParamCount < 3 Then
			vNode.AppendTagError "value_id|true_id|false_id"
			Exit Sub
		End If
		If Not CBool(vNode.EvalParamString(vNode.ParamItem(1))) Then
			Exit Sub
		End If
		vNode.StackPush
		If CBool(vNode.EvalParamString(vNode.ParamItem(1))) Then
			vNode.Engine.ContentAppend vNode.EvalParamString(vNode.ParamItem(2))
		Else
			vNode.Engine.ContentAppend vNode.EvalParamString(vNode.ParamItem(3))
		End If
		vNode.StackPop
	End Sub
End Class
Class CTPIfTrue
	Sub HandleTag(vNode)
		If vNode.ParamCount < 1 Then
			vNode.AppendTagError "value_id"
			Exit Sub
		End If
		If Not CBool(vNode.EvalParamString(vNode.ParamItem(1))) Then
			Exit Sub
		End If
		vNode.StackPush
		vNode.EvalNodes
		vNode.StackPop
	End Sub
End Class
Class CTPIfFalse
	Sub HandleTag(vNode)
		If vNode.ParamCount < 1 Then
			vNode.AppendTagError "value_id"
			Exit Sub
		End If
		If CBool(vNode.EvalParamString(vNode.ParamItem(1))) Then
			Exit Sub
		End If
		vNode.StackPush
		vNode.EvalNodes
		vNode.StackPop
	End Sub
End Class
Class CTPIfThen
	Sub HandleTag(vNode)
		If vNode.ParamCount < 1 Then
			vNode.AppendTagError "value_id"
			Exit Sub
		End If
		If CBool(vNode.EvalParamString(vNode.ParamItem(1))) Then
			EvalNode vNode, "then"
		else
			EvalNode vNode, "else"
		End If
	End Sub
	Sub EvalNode(vNode, sNode)
		Dim mNode
		Set mNode = vNode.LocateNode(sNode)
		If mNode Is Nothing Then Exit Sub
		mNode.StackPush
		mNode.EvalNodes
		mNode.StackPop
	End Sub
End Class
Class CTPIgnore
	Sub HandleTag(vNode)
	End Sub
End Class
Class CTPReplace
	Sub HandleTag(vNode)
		If vNode.ParamCount < 1 Then
			vNode.AppendTagError "value_id"
			Exit Sub
		End If
		If vNode.ParamItem(1) = "" Then Exit Sub
		vNode.StackPush
		On Error Resume Next
		vNode.Engine.ContentAppend vNode.EvalParamString(vNode.ParamItem(1))
		vNode.StackPop
	End Sub
End Class
Class CTPSubst
	Sub HandleTag(vNode)
		vNode.StackPush
		vNode.StackPush
		vNode.EvalNodes
		vNode.StackPop
		vNode.Engine.ContentReplaceFields
		vNode.StackPop
	End Sub
End Class
Class CTPCase
	Sub HandleTag(vNode)
		If vNode.ParamCount < 1 Then
			vNode.AppendTagError "value_id"
			Exit Sub
		End If
		Dim NodeName
		NodeName = vNode.EvalParamString(vNode.ParamItem(1))
		EvalCase vNode, NodeName
	End Sub
	Sub EvalCase(vNode, sNode)
		Dim mNode
		Set mNode = vNode.LocateNode(sNode)
		If mNode Is Nothing Then
			Set mNode = vNode.LocateNode("else")
		End If
		If mNode Is Nothing Then Exit Sub
		mNode.StackPush
		mNode.EvalNodes
		mNode.StackPop
	End Sub
End Class
Class CTPDecr
	Public Sub HandleTag(vNode)
		Dim ValName, CurValue
		If vNode.ParamCount < 1 Then
			vNode.AppendError "value_id"
		End If
		ValName = vNode.ParamItem(1)
		If vNode.Engine.HasValue(ValName) Then
			CurValue = vNode.Engine.GetValue(ValName)
			CurValue = CLng(CurValue) - 1
		Else
			If vNode.ParamCount > 1 Then
				CurValue = CLng(vNode.ParamItem(2))
			Else
				CurValue = 0
			End If
		End If
		vNode.Engine.PutValue ValName, CurValue
	End Sub
End Class
Class CTPIncr
	Public Sub HandleTag(vNode)
		Dim ValName, CurValue
		If vNode.ParamCount < 1 Then
			vNode.AppendError "value_id"
		End If
		ValName = vNode.ParamItem(1)
		If vNode.Engine.HasValue(ValName) Then
			CurValue = vNode.Engine.GetValue(ValName)
			CurValue = CLng(CurValue) + 1
		Else
			If vNode.ParamCount > 1 Then
				CurValue = CLng(vNode.ParamItem(2))
			Else
				CurValue = 0
			End If
		End If
		vNode.Engine.PutValue ValName, CurValue
	End Sub
End Class
Class CTPSetValue
	Public Sub HandleTag(vNode)
		Dim ValName, CurValue, Idx, Jdx
		If vNode.ParamCount < 1 Then
			vNode.AppendError "value_id??values&&|..."
		End If
		For Idx = 1 To vNode.ParamCount
			ValName = vNode.ParamItem(Idx)
			CurValue = Split(ValName, "??")
			' name of the value to be set
			ValName = CurValue(0)
			If UBound(CurValue) > 0 Then
				CurValue = CurValue(1)
				CurValue = Split(CurValue, "&&")
				If UBound(CurValue) < 1 Then
					CurValue = CurValue(0)
					CurValue = vNode.Engine.ReplaceFields(CStr(CurValue))
				Else
					ReDim Preserve CurValue(UBound(CurValue) + 1)
					For Jdx = UBound(CurValue) To 1 Step -1
						CurValue(Jdx) = CurValue(Jdx - 1)
						If Left(CurValue(Jdx), 1) = "$" Then
							CurValue(Jdx) = vNode.Engine.ReplaceFields(CStr(CurValue(Jdx)))
						End If
					Next
					CurValue(0) = ""
				End If
			Else
				CurValue = ""
			End If
			
			If Left(ValName, 1) = "$" Then
				ValName = vNode.Engine.ReplaceFields(ValName)
			End If
			
			' put the most recent value
			vNode.Engine.PutValue ValName, CurValue
		Next
	End Sub
End Class
Class CTPUnSetValue
	Public Sub HandleTag(vNode)
		Dim ValName, CurValue, Idx, Jdx
		If vNode.ParamCount < 1 Then
			vNode.AppendError "value_id|..."
		End If
		For Idx = 1 To vNode.ParamCount
			ValName = vNode.ParamItem(Idx)
			vNode.Engine.UnPutValue ValName
		Next
	End Sub
End Class
Class CTPCycle
	Public Sub HandleTag(vNode)
		Dim ValName, CurValue, Idx, Jdx
		If vNode.ParamCount < 3 Then
			vNode.AppendError "value_id|Alt1|Alt2..."
		End If
		ValName = vNode.ParamItem(1)
		CurValue = vNode.ParamItem(2)
		If vNode.Engine.HasValue(ValName) Then
			CurValue = vNode.Engine.GetValue(ValName)
			Jdx = 2
			For Idx = 2 To vNode.ParamCount
				If vNode.ParamItem(Idx) = CurValue Then
					Jdx = Idx
				End If
			Next
			If Jdx >= vNode.ParamCount Then
				CurValue = vNode.ParamItem(2)
			Else
				CurValue = vNode.ParamItem(Jdx + 1)
			End If
		End If
		vNode.Engine.PutValue ValName, CurValue
	End Sub
End Class

'----------------------------------------------------------------------
' Class Template Node
'----------------------------------------------------------------------
Class CTemplateNode
	Dim iid
	Dim mStartTag, mStopTag, mContent, mEvalProc
	Dim mEngine, mNodes, mParams
	Sub Class_Initialize()
		mStartTag="": mStopTag="": mContent="": mEvalProc=""
		mNodes = Array(""): mParams = Array("")
		Set mEngine = Nothing
	End Sub
	Property Get Engine()
		Set Engine = mEngine
	End Property
	Property Set Engine(Value)
		Dim Idx
		Set mEngine = Value
		For Idx = 1 To Ubound(mNodes)
			Set mNodes(Idx).Engine = mEngine
		Next
	End Property
	Property Get ID()
		ID = iid
	End Property
	Property Let ID(Value)
		iid = LCase(Value)
	End Property
	Property Get StartTag()
		StartTag = mStartTag
	End Property
	Property Let StartTag(Value)
		mStartTag = LCase(Value)
	End Property
	Property Get StopTag()
		StopTag = mStopTag
	End Property
	Property Let StopTag(Value)
		mStopTag = LCase(Value)
	End Property
	Property Get Content()
		Content = mContent
	End Property
	Property Let Content(Value)
		mContent = Value
	End Property
	Property Get EvalProc()
		EvalProc = mEvalProc
	End Property
	Property Let EvalProc(Value)
		mEvalProc = Value
	End Property
	Public Property Get NodeItem(Index)
		Set NodeItem = mNodes(Index)
	End Property
	Property Get NodeCount()
		NodeCount = UBound(mNodes)
	End Property
	Sub NodeAppend(vNode)
		ReDim Preserve mNodes(UBound(mNodes) + 1)
		Set mNodes(UBound(mNodes)) = vNode
	End Sub
	Public Property Get ParamCount()
		ParamCount = Ubound(mParams)
	End Property
	Property Get ParamItem(Index)
		ParamItem = mParams(Index)
	End Property
	Sub ParamAdd(sValue)
		Redim Preserve mParams(Ubound(mParams)+1)
		mParams(Ubound(mParams)) = sValue
	End Sub
	Sub EvalProcString()
		On Error Goto 0
		If mEvalProc <> "" Then
			If mEngine.HasHandler(mEvalProc) Then
				mEngine.GetHandler(mEvalProc).HandleTag Me
			Else
				mEngine.GetHandler("CustomHandler").HandleTag Me
			End If
		End If
	End Sub
	Function EvalParamString(sParam)
		If mEngine.HasValue(sParam) Then
			EvalParamString = mEngine.GetValue(sParam)
		Else
			EvalParamString = Eval(sParam)
		End If
	End Function
	Function EvalParamObject(sParam)
		If mEngine.HasValue(sParam) Then
			Set EvalParamObject = mEngine.GetObjectValue(sParam)
		Else
			Set EvalParamObject = Eval(sParam)
		End If
	End Function
	Sub EvalNode()
		If iid = "" Or mEvalProc = "" Then
			mEngine.ContentAppend mStartTag
			mEngine.ContentAppend mContent
			EvalNodes
			mEngine.ContentAppend mStopTag
		Else
			On Error Resume Next
			EvalProcString
			If Err.Number <> 0 Then
				mEngine.ContentAppend Err.Description
			End If
		End If
	End Sub
	Sub EvalNodeID(node_id)
		Dim mNode
		Set mNode = LocateNode(node_id)
		If mNode Is Nothing Then Exit Sub
		mNode.EvalNodes
	End Sub
	Sub EvalNodes()
		Dim Idx
		For Idx = 1 To UBound(mNodes)
			mNodes(Idx).EvalNode
		Next
	End Sub
	Sub StackPush()
		mEngine.ContentPush
	End Sub
	Sub StackPop()
		mEngine.ContentAppend mEngine.ContentPop()
	End Sub
	Sub AppendText(sText)
		Dim vNode: Set vNode = New CTemplateNode
		vNode.Content = sText
		NodeAppend vNode
	End Sub
	Sub AppendError(sMsg)
		AppendContent "<i><b>Error:</b>" & mEvalProc
		If sMsg <> "" Then AppendContent "|" & sMsg
		AppendContent "</i>"
	End Sub
	Sub AppendContent(sContent)
		mEngine.ContentAppend sContent
	End Sub
	Sub AppendTagError(sMSg)
		StackPush
		AppendError sMsg
		StackPop
	End Sub
	Function LocateNode(sID)
		Dim Idx
		For Idx = 1 to Ubound(mNodes)
			If LCase(mNodes(Idx).ID) = LCase(sID) Then
				Set LocateNode = mNodes(Idx)
				Exit Function
			End If
		Next
		Set LocateNode = Nothing
	End Function
End Class

'----------------------------------------------------------------------
' Template Compiler
'----------------------------------------------------------------------
Class CTemplateCompiler
	Dim mRxID, mRxTag, mRxDir
	Dim mParseLevel, mParseStack
	Dim RGX_ID, RGX_TAG, RGX_DIR
	Dim mDebug, mScrubTags
	Dim mParent

	Sub Class_Initialize()
		InitParseStack
		RGX_TAG = "<!--\[[^\]]+\]-->"
		Set mRxTag = new RegExp
		mRxTag.IgnoreCase = True
		mRxTag.Global = True
		mRxTag.MultiLine = False
		mRxTag.Pattern = RGX_TAG
		mDebug = False
		mScrubTags = True
		Set mParent = Nothing
	End Sub

	Property Get Debug()
		Debug = mDebug
	End Property
	Property Let Debug(Value)
		mDebug = Value
	End Property
	Property Get ScrubTags()
		ScrubTags = mScrubTags
	End Property
	Property Let ScrubTags(Value)
		mScrubTags = Value
	End Property
	Property Get Parent
		Set Parent = mParent
	End Property
	Property Set Parent(node)
		Set mParent = node
	End Property
	Sub InitParseStack()
		mParseStack = Array(New CTemplateNode)
		Set mParseStack(0).Engine = Me
		mParseStack(0).ID="_root"
		mParseLevel = 0
	End Sub
	Sub ParsePush(Value)
		mParseLevel = mParseLevel + 1
		ReDim Preserve mParseStack(mParseLevel)
		Set mParseStack(mParseLevel) = Value
	End Sub
	Function ParsePop()
		Set ParsePop = mParseStack(mParseLevel)
		mParseLevel = mParseLevel - 1
		ReDim Preserve mParseStack(mParseLevel)
	End Function
	Function ParsePeek()
		Set ParsePeek = mParseStack(mParseLevel)
	End Function
	Property Get ParseLevel()
		ParseLevel = mParseLevel
	End Property
	Function ParseFile(sFileName)
		Dim inp, tso, fso
		Set fso = Server.CreateObject("Scripting.FileSystemObject")
		On Error Resume next
		Set tso = fso.OpenTextFile(sFileName, 1, False)
		If Err.Number <> 0 Then
			Response.Write Err.Description & "<BR />"
			Response.Write sFileName
			Response.End
		End If
		inp = tso.ReadAll: tso.Close
		Set tso = Nothing: Set fso = Nothing
		Set ParseFile = ParseTemplate(inp)
	End Function
	Function ParseTemplate(sTemplate)
		Dim Matches, Index, LastOffset, Obj, LenTemplate
		InitParseStack
		Set Matches = mRxTag.Execute(sTemplate)
		LastOffset = 1
		LenTemplate = Len(sTemplate)
		For Index = 0 To Matches.Count - 1
			If Matches(Index).FirstIndex - LastOffset >= 0 Then
				ParsePeek().AppendText Mid(sTemplate, LastOffset, (Matches(Index).FirstIndex - LastOffset) + 1)
			End If
			HandleTagMatch Matches(Index)
			LastOffset = Matches(Index).FirstIndex + Matches(Index).Length + 1
		Next
		If LastOffset < LenTemplate Then
			ParsePeek().AppendText Right(sTemplate, (LenTemplate - LastOffset) + 1)
		End If
		Set ParseTemplate = mParseStack(0)
	End Function
	Function IsFormalEndTag(sValue)
		IsFormalEndTag = (Left(sValue, 6) = "<!--[/")
	End Function
	Function IsTermedTag(sValue)
		IsTermedTag = (Right(sValue, 5) = "/]-->")
	End Function
	Sub HandleTagMatch(oMatch)
		If IsFormalEndTag(oMatch.Value) Then
			ParseEndTag oMatch
		ElseIf IsTermedTag(oMatch.Value) Then
			ParseTermedTag oMatch
		Else
			ParseBeginTag oMatch
		End If
	End Sub
	Sub ParseTagProperties(oMatch, oNode, setID)
		Dim M, SP, Idx, Temp
		If mDebug Or Not mScrubTags Then
			oNode.StartTag = oMatch.Value
		Else
			oNode.StartTag = ""
		End If
		' STRIP INITIAL "<!--[/?" AND FINAL "/]-->"
		Temp = oMatch.Value
		Temp = Mid(Temp,6,Len(Temp)-9)
		If Left(Temp,1) = "/" Then
			Temp = Right(Temp,Len(Temp)-1)
		ElseIf Right(Temp,1) = "/" Then
			Temp = Left(Temp,Len(Temp)-1)
		End If
		SP = Split(Temp,"|")
		oNode.ID = SP(0)
		oNode.EvalProc = SP(0)
		For Idx = 1 To UBound(SP)
			oNode.ParamAdd SP(Idx)
		Next
	End Sub
	Sub ParseBeginTag(oMatch)
		Dim oNode
		If mDebug Then
			WriteOutput Right("000" & mParseLevel, 3) & ": " & DString(mParseLevel," | ") & Server.HtmlEncode(oMatch.Value) & "<br>" & vbCrLf
		End If
		Set oNode = New CTemplateNode
		Set oNode.Engine = Me
		ParseTagProperties oMatch, oNode, True
		ParsePush oNode
	End Sub
	Sub ParseEndTag(oMatch)
		Dim oNode, ID: ID = oMatch.Value
		ID = LCase(Trim(Mid(ID, 7, Len(ID) - 10)))
		While (ParsePeek().ID <> ID And ParseLevel > 0)
			Set oNode = ParsePop()
			ParsePeek().NodeAppend oNode
		Wend
		If ParsePeek().ID = ID Then
			Set oNode = ParsePeek()
			oNode.StopTag = oMatch.Value
			If ParseLevel > 0 Then
				ParsePop
				ParsePeek().NodeAppend oNode
			End If
		End If
		If mDebug Then
			WriteOutput Right("000" & mParseLevel, 3) & ": " & DString(mParseLevel," | ") & Server.HtmlEncode(oMatch.Value) & "<br>" & vbCrLf
		End If
	End Sub
	Sub ParseTermedTag(oMatch)
		Dim oNode: Set oNode = New CTemplateNode
		Set oNode.Engine = Me
		ParseTagProperties oMatch, oNode, False
		ParsePeek().NodeAppend oNode
		If mDebug Then
			WriteOutput Right("00000" & mParseLevel, 5) & ": " & DString(mParseLevel," | ") & Server.HtmlEncode(oMatch.Value) & "<br>" & vbCrLf
		End If
	End Sub
	Sub WriteOutput(sContent)
		If mParent Is Nothing Then
			Response.Write sContent
		Else
			mParent.ENGINE.ContentAppend sContent
		End IF
	End Sub
	Function DString(iCount, sRep)
		Dim IDx, Result
		For Idx = 1 To iCount
			Result = Result & sRep
		NExt
		DString = Result
	End Function
End Class

'----------------------------------------------------------------------
' Runtime Template Engine Factory
'----------------------------------------------------------------------
Class CKudzuEngineFactory
	Function CreateEngine()
		Set CreateEngine = New CTemplateEngine
	End Function
	Function CreateChildEngine(vNode)
		Set CreateChildEngine = CreateEngine()
		Set CreateChildEngine.Parent = vNode
	End Function
End Class

'----------------------------------------------------------------------
' Runtime Template Engine - BEGINS
'----------------------------------------------------------------------
Class CTemplateEngine
	Dim mRxFld
	Dim mRunStack, mRunLevel
	Dim mNodeTree
	Dim mDebug, mScrubTags, mVersion
	Dim mHandlers, mValues, mIterators
	Dim mTimeStart, mTimeEnd
	Dim mParent, mFactory
	Sub Class_Initialize()
		mTimeStart = Timer()
		mTimeEnd = mTimeStart
		Set mRxFld = new RegExp
		mRxFld.IgnoreCase = True
		mRxFld.Global = True
		mRxFld.MultiLine = False
		mRxFld.Pattern = "\{\{[^}]+\}\}"
		mDebug = False
		mScrubTags = True
		mVersion = "1.0.2a"
		Set mParent = Nothing
		Set mHandlers = Server.CreateObject("Scripting.Dictionary")
		Set mValues = Server.CreateObject("Scripting.Dictionary")
		Set mIterators = Server.CreateObject("Scripting.Dictionary")
		InstallHandlers
		PutValue "Kudzu_Version", mVersion
		ClassReset
	End Sub
	Sub InstallHandlers()
		SetHandler "Case", New CTPCase
		SetHandler "Cycle", New CTPCycle
		SetHandler "Decr", New CTPDecr
		SetHandler "Execute", New CTPExecute
		SetHandler "Flush", New CTPFlush
		SetHandler "IIf", New CTPSubst
		SetHandler "If", New CTPIfThen
		SetHandler "IfTrue", New CTPIfTrue
		SetHandler "IfFalse", New CTPIfFalse
		SetHandler "Ignore", New CTPIgnore
		SetHandler "Import", New CTPImport
		SetHandler "Incr", New CTPIncr
		SetHandler "Profiler", New CTPProfiler
		SetHandler "Replace", New CTPReplace
		SetHandler "SetValue", New CTPSetValue
		SetHandler "Subst", New CTPSubst
		SetHandler "UnSetValue", New CTPUnSetValue
	End Sub
	Property Let StartTime(Value)
		mTimeStart = Value
	End Property
	Property Get StartTime()
		StartTime = mTimeStart
	End Property
	Property Let StopTime(Value)
		mTimeEnd = Value
	End Property
	Property Get StopTime()
		StopTime = mTimeEnd
	End Property
	Property Let Debug(Value)
		mDebug = Value
	End Property
	Property Get Debug()
		Debug = mDebug
	End Property
	Property Get ScrubTags()
		ScrubTags = mScrubTags
	End Property
	Property Let ScrubTags(Value)
		mScrubTags = Value
	End Property
	Sub ClassReset()
		Set mNodeTree = New CTemplateNode
		Set mNodeTree.Engine = Me
		mNodeTree.Content = "No Content"
		mRunStack = Array("")
		mRunLevel = 0
	End Sub
	Property Get Parent
		Set Parent = mParent
	End Property
	Property Set Parent(node)
		Set mParent = node
	End Property

	Property Get Factory
		Set Factory = mFactory
	End Property
	Property Set Factory(oFactory)
		Set mFactory = oFactory
	End Property
	Property Get Iterators()
		Set Iterators = mIterators
	End Property
	Function HasIterator(sName)
		HasIterator = mIterators.Exists(LCase(sName))
	End Function
	Function GetIterator(sName)
		Set GetIterator = mIterators(LCase(sName))
	End Function
	Sub PutIterator(sName, oIter)
		Dim sKey: sKey = LCase(sName)
		If mIterators.Exists(sKey) Then
			mIterators.Remove sKey
		End If
		mIterators.Add sKey, oIter
	End Sub
	Property Get Handlers()
		Set Handlers = mHandlers
	End Property
	Function GetHandler(sName)
		Set GetHandler = mHandlers(LCase(sName))
	End Function
	Sub SetHandler(sName, oHandler)
		Dim sKey: sKey = LCase(sName)
		If mHandlers.Exists(sKey) Then
			mHandlers.Remove(sKey)
		End If
		If Not oHandler Is Nothing Then
			mHandlers.Add sKey, oHandler
		End If
	End Sub
	Function HasHandler(sName)
		HasHandler = mHandlers.Exists(LCase(sName))
	End Function
	Property Get Values()
		Set Values = mValues
	End Property
	Function GetValue(sName)
		GetValue = mValues(LCase(sName))
	End Function
	Function GetObjectValue(sName)
		Set GetObjectValue = mValues(LCase(sName))
	End Function
	Property Get HasValue(sName)
		HasValue = mValues.Exists(LCase(sName))
	End Property
	Sub PutValue(sName, vValue)
		Dim sKey: sKey = LCase(sName)
		If mValues.Exists(sKey) Then
			mValues.Remove sKey
		End If
		mValues.Add sKey, vValue
	End Sub
	Sub ParseFile(sFileName)
		Dim T_COMPILER
		mTimeStart = Timer()
		Set T_COMPILER = New CTemplateCompiler
		T_COMPILER.Debug = mDebug
		T_COMPILER.ScrubTags = mScrubTags
		Set T_COMPILER.Parent = mParent
		Me.ClassReset
		Set mNodeTree = T_COMPILER.ParseFile(sFileName)
		If mDebug Then Exit Sub
		Set mNodeTree.Engine = Me
	End Sub
	Sub ParseTemplate(sTemplate)
		Dim T_COMPILER
		mTimeStart = Timer()
		Set T_COMPILER = New CTemplateCompiler
		T_COMPILER.Debug = mDebug
		T_COMPILER.ScrubTags = mScrubTags
		Set T_COMPILER.Parent = mParent
		Me.ClassReset
		Set mNodeTree = T_COMPILER.ParseTemplate(sTemplate)
		If mDebug Then Exit Sub
		Set mNodeTree.Engine = Me
	End Sub
	Sub EvalTemplate()
		If mDebug Then Exit Sub
		mNodeTree.EvalNode
		WriteOutput mRunStack(0)
		mTimeEnd = Timer()
		If mParent Is Nothing Then
			WriteOutput VbCrLf & "<!-- Rendered by Kudzu " & mVersion
			WriteOutput ", " & FormatNumber(mTimeEnd - mTimeStart,4) & " seconds -->"
		End If
	End Sub
	Sub WriteOutput(sContent)
		If mParent Is Nothing Then
			Response.Write sContent
		Else
			ME.Parent.ENGINE.ContentAppend sContent
		End IF
	End Sub
	Function ContentLevel()
		ContentLevel = mRunLevel
	End Function
	Sub ContentPush()
		mRunLevel = mRunLevel + 1
		ReDim Preserve mRunStack(mRunLevel)
		mRunStack(mRunLevel) = ""
	End Sub
	Property Get Content()
		Content = mRunStack(mRunLevel)
	End Property
	Property Let Content(Value)
		mRunStack(mRunLevel) = Value
	End Property
	Function ContentPop()
		ContentPop = mRunStack(mRunLevel)
		If mRunLevel < 1 Then Exit Function
		mRunLevel = mRunLevel - 1
		ReDim Preserve mRunStack(mRunLevel)
	End Function
	Sub ContentAppend(sText)
		mRunStack(mRunLevel) = mRunStack(mRunLevel) & sText
	End Sub
	Sub ContentFlush()
		Dim Idx
		For Idx = 0 To mRunLevel
			WriteOutput mRunStack(Idx)
			mRunStack(Idx) = ""
		Next
		If Not mParent is Nothing Then
			mParent.Engine.ContentFlush
		End If
	End Sub
	Sub ContentReplaceFields()
		Me.Content = ReplaceFields(Me.Content)
	End Sub
	Function GetFieldTags(sValue)
		Set GetFieldTags = mRxFld.Execute(sValue)
	End Function
	Function EvalParamString(sParam)
		Select Case Left(sParam,1)
		Case "?" ' RequestString or Form Value
			EvalParamString = CStr(Request(Right(sParam,Len(sParam)-1)))
		Case "%" ' Session Variable
			EvalParamString = CStr(Session(Right(sParam,Len(sParam)-1)))
		Case "^" ' Application Level Variable - causes locking
			Application.Lock
			EvalParamString = Application(Right(sParam,Len(sParam)-1))
			Application.Unlock
		Case "#" ' Simple Cookie
			EvalParamString = Request.Cookies(Right(sParam,Len(sParam)-1))
		Case "&" ' Server Variable
			EvalParamString = CStr(Request.ServerVariables(Right(sParam,Len(sParam)-1)))
		Case "$" ' user has TranslateKudzuValue() Method available
			EvalParamString = TranslateKudzuVariable(Right(sParam,Len(sParam)-1))
		Case "!" ' an explicit VBScript Eval()
			EvalParamString = Eval(sParam)
		Case Else ' from the collection
			EvalParamString = GetValue(sParam)
		End Select
	End Function
	Function ReplaceFields(sValue)
		Dim LastOffset, Idx, M
		Dim Result, Tag
		Result = "": LastOffset = 1
		Set M = GetFieldTags(sValue)
		If M.Count > 0 Then
			On Error Resume Next
			For Idx = 0 to M.Count - 1
				If M(Idx).FirstIndex - LastOffset >= 0 Then
					Result = Result & Mid(sValue, LastOffset, (M(Idx).FirstIndex - LastOffset) + 1)
				End If
				Tag = Trim(Mid(M(Idx).Value, 3, M(Idx).Length - 4))
				Result = Result & EvalParamString(Tag)
				LastOffset = M(Idx).FirstIndex + M(Idx).Length + 1
			Next
			If LastOffset < Len(sValue) Then
				Result = Result & Right(sValue, (Len(sValue)-LastOffset) + 1)
			End If
		Else
			Result = sValue
		End If
		ReplaceFields = Result
	End Function
	Function libImport(libName)
		KudzuLIB.libImport libName
		libImport = KudzuLIB.libSetTags(libName,Me)
	End Function
	Function getLibrary()
		Set getLibrary = KudzuLIB
	End Function
	Function setLibPath(libPath)
		KudzuLIB.setLibPath libPath
	End Function
End Class
%>
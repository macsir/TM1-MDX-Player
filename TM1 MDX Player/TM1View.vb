'Imports TM1API
'Imports Excel = Microsoft.Office.Interop.Excel
'Option Explicit On

Module TM1View

    Dim hUser As Long
    Dim hServer As Long

    'Declare Function TM1_API2HAN Lib "tm1.xll" () As Long
    'The VB declarations were originally developed for VB6, where a Long, for example, was a 32-bit integer. In VB.NET, a Long is a 64-bit integer.
    'Declare Function TM1_API2HAN Lib "tm1.xll" () As Integer

    Const ErrNoConnection As String = "No connection  to TM1 Server"
    Const ErrNoAPI As String = "No connection to TM1 API."
    Public Function getServerHandle(hUser As Long, sServer As String) As Long
        getServerHandle = 0

        If sServer = "" Then
            Exit Function
        End If

        getServerHandle = TM1SystemServerHandle(hUser, sServer)

        If getServerHandle = 0 Then
            handleError(ErrNoConnection)
        End If

    End Function
    Public Function getUserHandle() As Long

        'getUserHandle = TM1_API2HAN()
        getUserHandle = TM1SystemOpen()

        If getUserHandle = 0 Then
            handleError(ErrNoAPI)
        End If
    End Function

    Public Sub handleError(sError As String)
        Debug.Print(sError)
    End Sub

    '********************************************************************************************
    ' Create mdx string of TM1 View
    ' This could then be used in mdx reports in Excel or on the web
    ' Could be enhanced to pickup subsets in dimensions and then do crossjoins to be more dynamic
    '********************************************************************************************
    Function TM1MDXView(p_strServer As String, strCube As String, strView As String) As String

        Dim hCube As Long, hView As Long
        Dim vNRows As Long, vNCols As Long, vRowSubsetArray As Long, vColSubsetArray As Long
        Dim NRows As Long, NCols As Long, NColDims As Long, NRowDims As Long
        Dim i As Integer, j As Integer
        Dim hColPool As Long, hRowPool As Long, hValPool As Long
        Dim vElement As Long, vName As Long, vRet As Long, vParent As Long
        Dim strRow As String, strCol As String, strParent As String
        Dim vColSubs(16) As Long, vRowSubs(16) As Long
        Dim vColDims(,) As String, vRowDims(,) As String
        Dim kr As Integer
        Dim strMDXRows As String, strMDXCols As String, strMDXFrom As String


        hUser = getUserHandle()

        If hUser = 0 Then
            Exit Function
        End If

        hServer = getServerHandle(hUser, p_strServer)

        If hServer = 0 Then
            Exit Function
        End If

        ' Create a Pool Handle
        hValPool = TM1ValPoolCreate(hUser)
        hCube = TM1ObjectListHandleByNameGet(hValPool, hServer, TM1ServerCubes(), TM1ValString(hValPool, strCube, 0))
        hView = TM1ObjectListHandleByNameGet(hValPool, hCube, TM1CubeViews(), TM1ValString(hValPool, strView, 0))

        TM1ViewArrayConstruct(hValPool, hView)

        vNRows = TM1ViewArrayRowsNof(hValPool, hView)
        vNCols = TM1ViewArrayColumnsNof(hValPool, hView)
        vRowSubsetArray = TM1ObjectPropertyGet(hValPool, hView, TM1ViewRowSubsets())
        vColSubsetArray = TM1ObjectPropertyGet(hValPool, hView, TM1ViewColumnSubsets())

        NRows = TM1ValIndexGet(hUser, vNRows)
        NCols = TM1ValIndexGet(hUser, vNCols)
        NColDims = TM1ValArrayMaxSize(hUser, vColSubsetArray)
        NRowDims = TM1ValArrayMaxSize(hUser, vRowSubsetArray)

        If (NRows = NColDims And Not NCols = NRowDims) Then
            Exit Function
        End If

        ' Get subsets
        For i = 1 To NColDims
            vColSubs(i) = TM1ValArrayGet(hUser, vColSubsetArray, i)
        Next i

        For i = 1 To NRowDims
            vRowSubs(i) = TM1ValArrayGet(hUser, vRowSubsetArray, i)
        Next i

        ' Val Pools
        hColPool = TM1ValPoolCreate(hUser)
        hRowPool = TM1ValPoolCreate(hUser)

        ' Get Col Elements
        If NColDims <> 0 Then
            ReDim vColDims(NColDims, NCols)
            For i = NRowDims + 1 To NCols
                For j = 1 To NColDims
                    strParent = ""
                    vElement = TM1ObjectListHandleByIndexGet(hValPool, vColSubs(j), TM1SubsetElements(), TM1ViewArrayValueGet(hValPool, hView, TM1ValIndex(hValPool, i), TM1ValIndex(hValPool, j)))
                    vName = TM1ObjectPropertyGet(hColPool, vElement, TM1ObjectName())
                    vParent = TM1ObjectPropertyGet(hValPool, TM1ObjectPropertyGet(hValPool, vElement, TM1ObjectParent()), TM1ObjectName())
                    TM1ValStringGet_VB(hUser, vParent, strParent, 100)
                    vColDims(j, i - 1) = Trim(Mid(strParent, 1, TM1ValStringMaxSize(hUser, vParent)))
                Next j
            Next i
        Else
            TM1MDXView = "No Column Dimension in view"
            Exit Function
        End If

        ' Get Row Elements
        If NRowDims <> 0 Then
            ReDim vRowDims(NRowDims, NRows)
            For i = NColDims + 1 To NRows
                For j = 1 To NRowDims
                    strParent = ""
                    vElement = TM1ObjectListHandleByIndexGet(hValPool, vRowSubs(j), TM1SubsetElements(), TM1ViewArrayValueGet(hValPool, hView, TM1ValIndex(hValPool, j), TM1ValIndex(hValPool, i)))
                    vName = TM1ObjectPropertyGet(hRowPool, vElement, TM1ObjectName())
                    vParent = TM1ObjectPropertyGet(hValPool, TM1ObjectPropertyGet(hValPool, vElement, TM1ObjectParent()), TM1ObjectName())
                    TM1ValStringGet_VB(hUser, vParent, strParent, 100)
                    vRowDims(j, i - 1) = Trim(Mid(strParent, 1, TM1ValStringMaxSize(hUser, vParent)))
                Next j
            Next i
        Else
            TM1MDXView = "No Row Dimension in view"
            Exit Function
        End If

        'tracking variables
        kr = 0

        'Zero Supression on?
        vRet = TM1ObjectPropertyGet(hValPool, hView, TM1ViewSuppressZeroes)
        If TM1ValBoolGet(hUser, vRet) = 1 Then
            strMDXRows = "SELECT NON EMPTY{"
            strMDXCols = "NON EMPTY {"
        Else
            strMDXRows = "SELECT {"
            strMDXCols = "{"
        End If

        ' Row Elements
        For i = NColDims + 1 To NRows
            strMDXRows = strMDXRows & "("
            For j = 1 To NRowDims
                strRow = ""
                vName = TM1ValPoolGet(hRowPool, kr)
                TM1ValStringGet_VB(hUser, vName, strRow, 100)
                strMDXRows = strMDXRows & "[" & Trim(vRowDims(j, i - 1)) & "].[" & Trim(Mid(strRow, 1, TM1ValStringMaxSize(hUser, vName))) & "],"
                Debug.Print(strMDXRows)
                kr = kr + 1
            Next j
            strMDXRows = Left(strMDXRows, Len(strMDXRows) - 1) & "),"
        Next i
        ' Remove extra comma
        strMDXRows = Left(strMDXRows, Len(strMDXRows) - 1) & "} ON ROWS,"

        'tracking variables
        kr = 0

        ' Col Elements
        For i = NRowDims + 1 To NCols
            strMDXCols = strMDXCols & "("
            For j = 1 To NColDims
                strCol = ""
                vName = TM1ValPoolGet(hColPool, kr)
                TM1ValStringGet_VB(hUser, vName, strCol, 100)
                strMDXCols = strMDXCols & "[" & Trim(vColDims(j, i - 1)) & "].[" & Trim(Mid(strCol, 1, TM1ValStringMaxSize(hUser, vName))) & "],"
                kr = kr + 1
            Next j
            strMDXCols = Left(strMDXCols, Len(strMDXCols) - 1) & "),"
        Next i
        ' Remove extra comma
        strMDXCols = Left(strMDXCols, Len(strMDXCols) - 1) & "} ON COLUMNS"

        strMDXFrom = "FROM [" & strCube & "]"

        TM1ValPoolDestroy(hColPool)
        TM1ValPoolDestroy(hRowPool)

        TM1ViewArrayDestroy(hValPool, hView)

        TM1ValPoolDestroy(hValPool)

        TM1MDXView = strMDXRows & vbCrLf & strMDXCols & vbCrLf & strMDXFrom & vbCrLf & TM1MDXViewTitles(p_strServer, strCube, strView)

    End Function
    '********************************************************************************************
    ' Create where clause of mdx string in TM1 View
    '********************************************************************************************
    Function TM1MDXViewTitles(p_strServer As String, strCube As String, strView As String)

        Dim vTitleArray As Long, vTitleElemArray As Long, vTitle As Long, vParent As Long
        Dim hCube As Long, hView As Long, hTitle As Long, hValPool As Long
        Dim iViewNoTitles As Integer, i As Integer
        Dim strTitle As String, strParent As String, strMDXWhere As String
        Dim arrTitles() As Long, arrTitleMembers() As Long

        hUser = getUserHandle()

        If hUser = 0 Then
            Exit Function
        End If

        hServer = getServerHandle(hUser, p_strServer)

        If hServer = 0 Then
            Exit Function
        End If

        ' Create a Pool Handle
        hValPool = TM1ValPoolCreate(hUser)

        hCube = TM1ObjectListHandleByNameGet(hValPool, hServer, TM1ServerCubes(), TM1ValString(hValPool, strCube, 0))
        hView = TM1ObjectListHandleByNameGet(hValPool, hCube, TM1CubeViews(), TM1ValString(hValPool, strView, 0))
        'This gets you the array of Title subset handles
        vTitleArray = TM1ObjectPropertyGet(hValPool, hView, TM1ViewTitleSubsets())
        'This gets you the array of Title member selections (index of element in subset)
        vTitleElemArray = TM1ObjectPropertyGet(hValPool, hView, TM1ViewTitleElements())
        iViewNoTitles = TM1ValArrayMaxSize(hUser, vTitleArray)

        'Store the array of subset handles
        ReDim arrTitles(iViewNoTitles)
        For i = 1 To iViewNoTitles
            arrTitles(i) = TM1ValArrayGet(hUser, vTitleArray, i)
        Next i
        'Store the array of subset member indexes
        ReDim arrTitleMembers(iViewNoTitles)
        For i = 1 To iViewNoTitles
            arrTitleMembers(i) = TM1ValArrayGet(hUser, vTitleElemArray, i)
        Next i

        strMDXWhere = "WHERE ("
        'Now get the member names
        For i = 1 To iViewNoTitles
            hTitle = TM1ObjectListHandleByIndexGet(hValPool, arrTitles(i), TM1SubsetElements(), arrTitleMembers(i))
            vTitle = TM1ObjectPropertyGet(hValPool, hTitle, TM1ObjectName())
            vParent = TM1ObjectPropertyGet(hValPool, TM1ObjectPropertyGet(hValPool, hTitle, TM1ObjectParent()), TM1ObjectName())
            If TM1ValType(hUser, vTitle) = TM1ValTypeString() Then
                TM1ValStringGet_VB(hUser, vTitle, strTitle, 100)
                TM1ValStringGet_VB(hUser, vParent, strParent, 100)
                strMDXWhere = strMDXWhere & "[" & Trim(Mid(strParent, 1, TM1ValStringMaxSize(hUser, vParent))) & "].[" & Trim(Mid(strTitle, 1, TM1ValStringMaxSize(hUser, vTitle))) & "],"
            Else
                strMDXWhere = ""
            End If
        Next i
        ' Remove extra comma
        strMDXWhere = Left(strMDXWhere, Len(strMDXWhere) - 1) & ")"

        TM1MDXViewTitles = strMDXWhere

        TM1ValPoolDestroy(hValPool)

    End Function
End Module

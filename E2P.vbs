'
' Excel-2-Powerpoint Bridge
'
' (c) 2016 Peter Postmann
'
' This project is licensed under the terms of the MIT license
'

Dim ConsoleMode, colNamedArguments, saveJob, _ 
xlApp, xlBook, ppApp, ppPpt, ppFileName, _ 
xlFileName, xlTargetSheetName, xlTargetSheet, xlHeaderRow, xlFirstDataRow, xlFirstCol, xlIndexName, _
xlFirstCol_str, xlHeaderRow_str, xlFirstDataRow_str, ppTemplateSlideID_str, ppInsertAfterSlide_str, _
dictMapping, dictIDs, index, dictCommon, dictUpdated, dictShapeData, dictSlideIndex, _
ppTemplateSlideID, ppTemplateSlide, ppInsertAfterSlide

' Success?
Dim success
success = false
Set xlApp  = Nothing

'Declare progressbar and percentage complete
Dim pb
Dim pbCount
Dim pbTotal

'Setup the initial progress bar
Set pb = New ProgressBar

' Don't Display File/Open Dialog if started with cscript
ConsoleMode = false

' Get Arguments
Set colNamedArguments  = WScript.Arguments.Named
xlFileName             = colNamedArguments.Item("xlFileName")
xlTargetSheetName      = colNamedArguments.Item("xlTargetSheetName")
xlFirstCol_str         = colNamedArguments.Item("xlFirstCol")
xlHeaderRow_str        = colNamedArguments.Item("xlHeaderRow")
xlFirstDataRow_str     = colNamedArguments.Item("xlFirstDataRow")
ppFileName             = colNamedArguments.Item("ppFileName")
ppTemplateSlideID_str  = colNamedArguments.Item("ppTemplateSlideID")
ppInsertAfterSlide_str = colNamedArguments.Item("ppInsertAfterSlide")
saveJob                = colNamedArguments.Item("saveJob")
 
Do

    ' Get Excel file name
    If xlFileName = "" Then 
        xlFileName = SelectFile("Select Excel File")
    End IF
    If xlFileName = "" Then Exit Do

    ' Open workbook
    Set xlApp  = CreateObject("Excel.Application")
        xlApp.Visible = True
    Set xlBook = xlApp.Workbooks.Open(xlFileName,,True)
        
    ' Get target sheet name and open it    
    If xlTargetSheetName = "" Then 
        xlTargetSheetName = xlBook.ActiveSheet.Name
        xlTargetSheetName = UserInput("Target sheet", "Target Sheet", xlTargetSheetName)
    End If    
    Set xlTargetSheet = xlGetTargetSheet(xlBook, xlTargetSheetName)
    If xlTargetSheet Is Nothing Then Exit Do
    
    ' Get ColumnOffset
    If xlFirstCol_str = "" Then
        xlFirstCol = CInt(UserInput("First column", "", 1))
    Else
        xlFirstCol = CInt(xlFirstCol_str)
    End If
    If xlFirstCol = "" Then Exit Do
    
    ' Get header row
    If xlHeaderRow_str = "" Then
        xlHeaderRow =  xlGetHeaderRow(xlTargetSheet, xlFirstCol)
        xlHeaderRow = CInt(UserInput("Header row", "", xlHeaderRow))
    Else
        xlHeaderRow = CInt(xlHeaderRow_str)
    End If
    If xlHeaderRow = "" Then Exit Do
    
    ' Get first data row
    If xlFirstDataRow_str = "" Then
        xlFirstDataRow = CInt(UserInput("First data row", "", xlHeaderRow + 1))
    Else
        xlFirstDataRow = CInt(xlFirstDataRow_str)
    End If
    If xlFirstDataRow = "" Then Exit Do
    
    ' Get name of ID row
    xlIndexName = "{{" & xlTargetSheet.Cells(xlHeaderRow, xlFirstCol).Value & "}}"
    
    ' Create mapping {{header}} --> column number
    Set dictMapping = xlGetMapping(xlTargetSheet, xlHeaderRow, xlFirstCol)
    If dictMapping Is Nothing Then Exit Do
    
    ' Create mapping {ID} --> row number
    Set dictIDs = xlGetIDs(xlTargetSheet, xlFirstDataRow, xlFirstCol)
    If dictIDs Is Nothing Then Exit Do
    
    ' Create mapping {common_infos} --> data
    Set dictCommon = xlGetCommon(xlBook)
    If dictCommon Is Nothing Then Exit Do
    
    ' Get Powerpoiunt file name
    If ppFileName = "" Then 
        ppFileName = SelectFile("Select Powerpoint File")
    End If
    If ppFileName = "" Then Exit Do

    ' Open presentation
    Set ppApp = CreateObject("Powerpoint.Application")
    Set ppPpt = ppApp.Presentations.Open(ppFileName)

    ' Get template slide
    If ppTemplateSlideID_str = "" Then 
        ppTemplateSlideID = CInt(UserInput("Template Slide", "", 2))
    Else
        ppTemplateSlideID = CInt(ppTemplateSlideID_str)
    End If
    Set ppTemplateSlide = ppPpt.Slides(ppTemplateSlideID)
    If ppTemplateSlide Is Nothing Then Exit Do
        
    ' Get insert position
    If ppInsertAfterSlide_str = "" Then 
        ppInsertAfterSlide = ppPpt.Slides.Count
        ppInsertAfterSlide = UserInput("Insert new Slides after", "", ppInsertAfterSlide)
    Else
        ppInsertAfterSlide = CInt(ppInsertAfterSlide_str)
    End If
    If ppInsertAfterSlide <= 0 Then Exit Do
    
    If saveJob = "" Then        
        saveJob = YesNoDialog("Save Job?", "Save Job")
    End If
    
    If saveJob = "y" Then
    
        Set objFSO=CreateObject("Scripting.FileSystemObject")

        ' How to write file
        outFile = xlFileName & ".bat"
        Set objFile = objFSO.CreateTextFile(outFile,True)
        objFile.Write "wscript " & wscript.scriptname _
                                 & " /xlFileName:" & chr(34) & xlFileName & chr(34) _
                                 & " /xlTargetSheetName:" & chr(34) & xlTargetSheetName & chr(34) _
                                 & " /xlFirstCol:" & xlFirstCol _
                                 & " /xlHeaderRow:" & xlHeaderRow _
                                 & " /xlFirstDataRow:" & xlFirstDataRow _
                                 & " /ppFileName:" & chr(34) & ppFileName & chr(34) _
                                 & " /ppTemplateSlideID:" & ppTemplateSlideID _
                                 & " /ppInsertAfterSlide:" & ppInsertAfterSlide _ 
                                 & " /saveJob:n" _
                                 & vbCrLf
        objFile.Close
    End If
    
    ' Monitor updated SlideIndex: SlideIndex --> true/false
    Set dictUpdated    = CreateObject("Scripting.Dictionary")
    
    ' Create mapping {ID} --> Shapes    
    Set dictSlideIndex = CreateObject("Scripting.Dictionary")
        
    ' Create mapping {ID} --> Shapes    
    Set dictShapeData  = CreateObject("Scripting.Dictionary")
    
    ' Setup ProgressBar
    pb.Show()
    pb.NextTaks "Taks 1 of 2", "Parsing Presentation", ppPpt.Slides.Count
    
    ' Loop throug all Slides
    For Each Slide in ppPpt.Slides 
    If Slide.SlideIndex <> ppTemplateSlideID Then
    
        Dim dictShapes, TargetID
        
        Set dictShapes = ppGetShapes(Slide.Shapes)
            TargetID   = ""
        
        ' Check for each Shape
        For Each Shape in dictShapes.Items
                        
            ' if binding exists
            IF Shape.Name = xlIndexName Then
                
                If TargetID <> "" Then
                    WScript.Echo ("Fatal: Dublicated shape name! This should never happen.")           
                    Exit Do
                End If
                
                ' and request update
                TargetID = Shape.TextEffect.Text
            End If
            
            ' and updated common information --> updated
            If dictCommon.Exists(Shape.Name) Then
                Shape.TextEffect.Text = dictCommon(Shape.Name)
            End If
            
        Next
        
        ' Save target information        
        If Not dictUpdated.Exists(TargetID) Then
        
            IF TargetID <> "" Then
                dictUpdated.add    TargetID, false
                dictSlideIndex.add TargetID, Slide.SlideIndex
                dictShapeData.add  TargetID, dictShapes
            End If
            
        Else        
            WScript.Echo ("Warning: Dublicated binding. Skipping slide " & Slide.SlideIndex)  
        End If
        
    End If   
    
        pb.NextStep()
    
    Next
    
    ' Setup ProgressBar
    pb.NextTaks "Taks 2 of 2", "Copying Data", dictIDs.Count
    
    ' Loop throug data       
    For xlRowNum = xlFirstDataRow To dictIDs.Count + xlFirstDataRow - 1
    
        Dim Slide, Shapes
        
        TargetId = CStr(xlTargetSheet.Cells(xlRowNum, xlFirstCol).Value)
             
        ' Get referance or create new slide
        If dictUpdated.Exists(TargetId) Then
            Set Slide  = ppPpt.Slides(dictSlideIndex(TargetId))
            Set Shapes = dictShapeData(TargetId)
        Else      
            Set Slide  = ppTemplateSlide.Duplicate
            ppInsertAfterSlide = ppInsertAfterSlide + 1
            Slide.MoveTo(ppInsertAfterSlide)
            Set Shapes = ppGetShapes(Slide.Shapes)
        End If
        
        ' Check for each shape 
        For Each Shape in Shapes.Items
        
            ' if mapping exists
            If dictMapping.Exists(Shape.Name) Then
                
                ' check if target has TextEffect property
                ' https://msdn.microsoft.com/en-us/library/aa432678(v=office.12).aspx
                If  ISObject(Shape.TextEffect) Then                      
                       
                    Dim xlColumnNum, Cell, CellVarType
                    
                    xlColumnNum = dictMapping(Shape.Name)
                    Cell        = xlTargetSheet.Cells(xlRowNum, xlColumnNum)
                    CellVarType = VarType(Cell)
                    
                    IF  CellVarType = 0  Or _ 
                        CellVarType = 5  Or _ 
                        CellVarType = 8 Or _ 
                        CellVarType = 17 Then    
                        
                        ' and update Data
                        Shape.TextEffect.Text = xlTargetSheet.Cells(xlRowNum, dictMapping(Shape.Name))
                        
                    Else
                        WScript.Echo("Warning: Incompatible cell type '" & CellVarType & "' in Cell (" & xlRowNum & "," & xlColumnNum & ")!")                         
                    End If
                Else
                    WScript.Echo("Warning: Incompatible shape type (" & Shape.Type & ") for data binding '" & Shape.Name & "'!") 
                End If
                
            End IF
            
        Next 
        
        dictUpdated(TargetId) = true

        pb.NextStep()        
        
    Next
       
    ' Check for missing links 
    For Each Key in dictUpdated.Keys
    
        If Not dictUpdated(Key) Then
            WScript.Echo("Warning: ID " & Key & " on Slide " & dictSlideIndex(Key) & " not found in data source!")          
        End If
    Next    
    
    success = true
    
Exit Do
Loop

pb.Close()

If Not xlApp Is Nothing Then xlApp.Quit
If success Then WScript.Echo("Finished")


Set xlBook = Nothing
Set xlApp  = Nothing

Set ppPpt  = Nothing
Set ppApp  = Nothing

WScript.Quit

'
' xlGetTargetSheet
'
' Return reference of target sheet by name or first sheet 
'
Function xlGetTargetSheet(Book, TargetSheetName)

    Dim TargetSheet

    If TargetSheetName = "" Then
        Set TargetSheet = Book.Sheets(1)
    Else
        For Each Sheet in Book.Sheets
            If Sheet.Name = TargetSheetName Then
                Set TargetSheet = Sheet
                Exit For
            End If
        Next
    End If
    
    If TargetSheet Is Nothing Then
        WScript.Echo "Error: Sheet '" & TargetSheetName & "' not found!"
    End If
    
    Set xlGetTargetSheet = TargetSheet

End Function

'
' xlGetHeaderRow
'
' Return number of first non-empty row in TargetSheet and asume it's the table header
'
Function xlGetHeaderRow(TargetSheet, FirstCol)

    Dim RowNum, targetCells
    
    RowNum = 1
    
    targetCells = TargetSheet.Cells(RowNum, FirstCol)
    
    Do While TargetSheet.Cells(RowNum, FirstCol).Value = ""     
        RowNum = RowNum + 1
        If RowNum > 100 Then Exit Do
    Loop
    
    If RowNum > 100 Then 
        xlGetHeaderRow = ""
    Else
        xlGetHeaderRow = RowNum
    End If

End Function

'
' xlGetMapping
'
' Return a dictonary which maps headings to column numbers
'
Function xlGetMapping(TargetSheet, HeaderRow, FirstCol)

    Dim Mapping, ColNum

    Set Mapping = CreateObject("Scripting.Dictionary")
    ColNum      = FirstCol

    Do    
        Dim Value
    
        ' Get cell value
        Value = CStr(TargetSheet.Cells(HeaderRow, ColNum).Value)
        
        ' Loop until cell is empty
        If Value = "" Then
            Exit Do
        End If

        ' Check for duplicated keys
        If Mapping.Exists("{{" & Value & "}}") Then
        
            Dim numCopy, newName
            
            numCopy = 1
            
            ' Rename until key is unique
            Do 
                newName = Value & "_" & numCopy
                
                If Not Mapping.Exists("{{" & newName & "}}") Then
                    Exit Do
                End If
                
                numCopy = numCopy + 1
            Loop
        
            WScript.Echo("Warning: Mapping dublicated heading on column " & ColNum & " to '" & newName & "'.")
             
            Value = newName
        End If
    
        ' Add heading->column
        Mapping.Add "{{" & Value & "}}", ColNum
        
        ColNum = ColNum + 1
    Loop
        
    Set xlGetMapping = Mapping
    
    If ColNum = FirstCol Then
        WScript.Echo "Error: No headings found!"
        Set xlGetMapping = Nothing
    End If
    
End Function

'
' xlGetCommon
'
' Return a dictonary whith common information
'
Function xlGetCommon(Book)

    Dim dictCommon

    Set dictCommon = CreateObject("Scripting.Dictionary")
    
    dictCommon.Add "{source}", Book.Name    
    
    Set xlGetCommon = dictCommon
    
End Function

Function xlGetIDs(TargetSheet, FirstDataRow, FirstCol)

    Dim dictIDs, RowNum, errDoublicatedId, TargetID

    Set dictIDs = CreateObject("Scripting.Dictionary")
        RowNum  = FirstDataRow
        errDoublicatedId = false
        
    Do         
        TargetID = CStr(TargetSheet.Cells(RowNum, FirstCol).Value)
        
        If TargetID = "" Then
            Exit Do
        End If
    
        If dictIDs.Exists(TargetID) Then
            errDoublicatedId = true
            Exit Do
        End If
    
        dictIDs.Add TargetID, RowNum
    
        RowNum = RowNum + 1
    Loop
    
    Set xlGetIDs = dictIDs
    
    If RowNum = 1 Then
        WScript.Echo "Error: No data found!"
        Set xlGetIDs = Nothing
    End If 
    
    If errDoublicatedId Then
        WScript.Echo "Error: Dublicated ID '" & TargetID & "' in Row '" & RowNum & "'!"
        Set xlGetIDs = Nothing
    End If
    
End Function

Function isConsoleScript()
    If UCase( Right( WScript.FullName, 12 ) ) = "\CSCRIPT.EXE" Then
        isConsoleScript = True
    Else
        isConsoleScript = False
    End If
End Function

'
' UserInput
'
' Return input from user or default valie
'
Function UserInput(PromptText, Title, DefaultValue)
' This function prompts the user for some input.
' When the script runs in CSCRIPT.EXE, StdIn is used,
' otherwise the VBScript InputBox( ) function is used.
' myPrompt is the the text used to prompt the user for input.
' The function returns the input typed either on StdIn or in InputBox( ).
' Written by Rob van der Woude
' http://www.robvanderwoude.com

    If isConsoleScript() Then
        Dim DefaultText
        
        If DefaultValue <> "" Then
            DefaultText = " [" & DefaultValue & "]"
        Else
            DefaultText = ""
        End If
            
        WScript.StdOut.Write PromptText & DefaultText & ": "
        UserInput = WScript.StdIn.ReadLine        
        
        If UserInput = "" Then
            UserInput = DefaultValue
        End If
    Else
        ' If not, use InputBox( )
        UserInput = InputBox(PromptText, Title, DefaultValue)
    End If
End Function

'
' SelectFileDialog
'
' Return file name from file-open dialog
'
Function SelectFileDialog()
    ' File Browser via HTA
    ' Author:   Rudi Degrande, modifications by Denis St-Pierre and Rob van der Woude
    ' Features: Works in Windows Vista and up (Should also work in XP).
    '           Fairly fast.
    '           All native code/controls (No 3rd party DLL/ XP DLL).
    ' Caveats:  Cannot define default starting folder.
    '           Uses last folder used with MSHTA.EXE stored in Binary in [HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\ComDlg32].
    '           Dialog title says "Choose file to upload".
    ' Source:   https://social.technet.microsoft.com/Forums/scriptcenter/en-US/a3b358e8-15ae-4ba3-bca5-ec349df65ef6/windows7-vbscript-open-file-dialog-box-fakepath?forum=ITCG

    SelectFileDialog = ""
        
    If ConsoleMode And isConsoleScript() Then
        WScript.StdOut.Write "Select File: "
        SelectFileDialog = WScript.StdIn.ReadLine     
    Else
    
        Dim objExec, strMSHTA, wshShell


        ' For use in HTAs as well as "plain" VBScript:
        strMSHTA = "mshta.exe ""about:" & "<" & "input type=file id=FILE>" _
                 & "<" & "script>FILE.click();new ActiveXObject('Scripting.FileSystemObject')" _
                 & ".GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);" & "<" & "/script>"""
        ' For use in "plain" VBScript only:
        ' strMSHTA = "mshta.exe ""about:<input type=file id=FILE>" _
        '          & "<script>FILE.click();new ActiveXObject('Scripting.FileSystemObject')" _
        '          & ".GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);</script>"""

        Set wshShell = CreateObject( "WScript.Shell" )
        Set objExec = wshShell.Exec( strMSHTA )

        SelectFileDialog = objExec.StdOut.ReadLine( )

        Set objExec = Nothing
        Set wshShell = Nothing
        
    End If
    
End Function

'
' YesNoDialog
'
' Return user response yes or no
'
Function YesNoDialog(PromptText, Title)

    YesNoDialog = ""

    If ConsoleMode And isConsoleScript() Then
        WScript.StdOut.Write PromptText & "[y/n]: "
        YesNoDialog = LCase(WScript.StdIn.ReadLine)   
    Else
        result = MsgBox (PromptText, vbYesNo, Title)

        Select Case result
        Case vbYes
            YesNoDialog = "y"
        Case vbNo
            YesNoDialog = "n"
        End Select
        
    End If
End Function

'
' getFullPath
'
' Return full path if files exists
'
Function GetFullPath(FileName)
   
    Dim fso   
    
    Set fso = CreateObject("Scripting.FileSystemObject")

    If (fso.FileExists(FileName)) Then
        GetFullPath = fso.GetAbsolutePathName(FileName)
    Else
        GetFullPath = ""
    End If
End Function

'
' SelectFile
'
' Return full path of target file or empty string on error
'
Function SelectFile(PromptText)

    Dim FileName

    If PromptText <> "" Then
        WScript.Echo PromptText
    End If
    
    FileName = SelectFileDialog()

    If FileName = "" Then
        WScript.Echo "Error: No file selected!"
    Else
        FileName = GetFullPath(FileName)

        If FileName = "" Then
            WScript.Echo "Error: File '" & FileName & "' does not exist!"
        End If
    End If
    
    SelectFile = FileName
    
End Function

'
' ppGetShapes
'
' Return dictonary with all shapes including nested shapes
'
Function ppGetShapes(Shapes)

    Set dictShapes  = CreateObject("Scripting.Dictionary")
    Set dictShapes  = ppGetShapesRecursive(dictShapes, Shapes)
    Set ppGetShapes = dictShapes
    
End Function

'
' ppGetShapes
'
' Return dictonary with all shapes including nested shapes (recursive)
'
Function ppGetShapesRecursive(Data, Shapes)

    For Each Shape in Shapes
           
        ' If shape is group element
        If Shape.Type = 6 Then ' msoGroup
            Set Data = ppGetShapesRecursive(Data, Shape.GroupItems)
        Else
        
            ' Check for data bindings
            If Left(Shape.Name, 1) = "{" And Right(Shape.Name, 1) = "}" Then
                Data.Add Data.Count + 1, Shape 
            End If
            
        End If
        
    Next

    Set ppGetShapesRecursive = Data
    
End Function

'
' ProgressBar
'
' Source: http://www.northatlantawebdesign.com/index.php/2009/07/16/simple-vbscript-progress-bar/
'
Class ProgressBar
    Private m_PercentComplete
    Private m_CurrentStep
    Private m_ProgressBar
    Private m_Total
    Private m_Count
    Private m_Title
    Private m_Text
    Private m_StatusBarText
     
    'Initialize defaults
    Private Sub ProgessBar_Initialize
        Set m_ProgressBar = Nothing
        m_PercentComplete = 0
        m_CurrentStep = 0
        m_Title = "Progress"
        m_Text = ""
    End Sub
     
    Public Function SetTitle(pTitle)
        m_Title = pTitle
    End Function
     
    Public Function SetText(pText)
        m_Text = pText
    End Function
     
    Public Function Update(percentComplete)
        m_PercentComplete = percentComplete
        UpdateProgressBar()
    End Function
     
    Public Function Show()
        Set m_ProgressBar = CreateObject("InternetExplorer.Application")
        'in code, the colon acts as a line feed
        m_ProgressBar.navigate2 "about:blank" : m_ProgressBar.width = 315 : m_ProgressBar.height = 40 : m_ProgressBar.toolbar = false : m_ProgressBar.menubar = false : m_ProgressBar.statusbar = false : m_ProgressBar.visible = True
        m_ProgressBar.document.write "<body Scroll=no style='margin:0px;padding:0px;'><div style='text-align:center;'><span name='pc' id='pc'>0</span></div>"
        m_ProgressBar.document.write "<div id='statusbar' name='statusbar' style='border:1px solid blue;line-height:10px;height:10px;color:blue;'></div>"
        m_ProgressBar.document.write "<div style='text-align:center'><span id='text' name='text'></span></div>"
    End Function
    
    Public Function NextTaks(pTitle, pText, pbTotal)
    
        SetTitle(pTitle)
        SetText(pText)
        
        m_Total = pbTotal 
        m_Count = pbCount    
        m_PercentComplete = 0 
        
        UpdateProgressBarStep()    
    End Function
    
    Public Function NextStep()     
        m_Count = m_Count + 1  
        UpdateProgressBarStep()       
    End Function
     
    Public Function Close()
        If ISObject(m_ProgressBar) Then
            m_ProgressBar.quit
            Set m_ProgressBar = Nothing
        End If
    End Function
     
    Private Function UpdateProgressBar()
        If m_PercentComplete = 0 Then
        m_StatusBarText = ""
        End If
        For n = m_CurrentStep to m_PercentComplete - 1
        m_StatusBarText = m_StatusBarText & "|"
        m_ProgressBar.Document.GetElementById("statusbar").InnerHtml = m_StatusBarText
        m_ProgressBar.Document.title = n & "% Complete : " & m_Title
        m_ProgressBar.Document.GetElementById("pc").InnerHtml = n & "% Complete : " & m_Title
        wscript.sleep 10
        Next
        m_ProgressBar.Document.GetElementById("statusbar").InnerHtml = m_StatusBarText
        m_ProgressBar.Document.title = m_PercentComplete & "% Complete : " & m_Title
        m_ProgressBar.Document.GetElementById("pc").InnerHtml = m_PercentComplete & "% Complete : " & m_Title
        m_ProgressBar.Document.GetElementById("text").InnerHtml = m_Text
        m_CurrentStep = m_PercentComplete
    End Function
    
    
    Private Function UpdateProgressBarStep()
        
        m_PercentComplete = Int((m_Count/m_Total)*100)
    
        If m_PercentComplete = 0 Then
            m_StatusBarText = ""
        End If
        
        For n = m_CurrentStep to m_PercentComplete - 1
            m_StatusBarText = m_StatusBarText & "|"
            m_ProgressBar.Document.GetElementById("statusbar").InnerHtml = m_StatusBarText
            m_ProgressBar.Document.title = n & "% Complete : " & m_Title
            m_ProgressBar.Document.GetElementById("pc").InnerHtml = "Step " & m_Count & " of " & m_Total & " : " & m_Title  
            wscript.sleep 10
        Next
        
        m_ProgressBar.Document.GetElementById("statusbar").InnerHtml = m_StatusBarText
        m_ProgressBar.Document.title = n & "% Complete : " & m_Title
        m_ProgressBar.Document.GetElementById("pc").InnerHtml = "Step " & m_Count & " of " & m_Total & " : " & m_Title 
        m_ProgressBar.Document.GetElementById("text").InnerHtml = m_Text
        m_CurrentStep = m_PercentComplete
    End Function
 
End Class

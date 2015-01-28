Sub exportGEXF()
    Dim allTasks As Tasks
    Dim theTask As Task
    Dim taskDependency As taskDependency
    Dim textToExport As String
    Dim linkArray() As String
    Dim linkArrayLength As Integer
    Dim arrayCounter As Integer
    Dim filePath As String
    Dim theCounter As Integer
    Dim edgeSourceTarget As String
    Dim parentTask As String
    Dim edgeSource As String
            
    textToExport = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "UTF-8" & Chr(34) & "?>" & vbCrLf _
                    & "<gexf xmlns=" & Chr(34) & "http://www.gexf.net/1.2draft" & Chr(34) & " version=" & Chr(34) & "1.2" & Chr(34) & ">" & vbCrLf _
                    & "<meta lastmodifieddate=" & Chr(34) & "2009-03-20" & Chr(34) & ">" & vbCrLf _
                    & vbTab & "<creator>Wolfgang Geithner</creator>" & vbCrLf _
                    & vbTab & "<description>Cryring Project Dependencies</description>" & vbCrLf _
                    & "</meta>" & vbCrLf _
                    & "<graph mode=" & Chr(34) & "static" & Chr(34) & " defaultedgetype=" & Chr(34) & "directed" & Chr(34) & ">" & vbCrLf _
                    & "<attributes class=""node"">" & vbCrLf _
                    & vbTab & "<attribute id=""0"" title=""status"" type=""string""/>" & vbCrLf _
                    & "</attributes>" & vbCrLf _
                    & vbTab & "<nodes>" & vbCrLf _
    
    'First create nodes
    For Each theTask In ActiveProject.Tasks
        
        If theTask.Status = 0 Then
            textToExport = textToExport & vbTab & vbTab & "<node id=""" & theTask.UniqueID & """ label=" & Chr(34) & Chr(34) & " pid=""" & theTask.OutlineParent & """ >" & vbCrLf
        Else
            textToExport = textToExport & vbTab & vbTab & "<node id=""" & theTask.UniqueID & """ label=""" & theTask.Name & """ pid=""" & theTask.OutlineParent & """ >" & vbCrLf
        End If
        textToExport = textToExport & vbTab & vbTab & vbTab & "<attvalues>" & vbCrLf _
                                    & vbTab & vbTab & vbTab & vbTab & "<attvalue for=""0"" value=""" & CStr(theTask.Status) & """ />" & vbCrLf _
                                    & vbTab & vbTab & vbTab & "</attvalues>" & vbCrLf _
                                    & vbTab & vbTab & "</node>" & vbCrLf
    Next theTask
    
    textToExport = textToExport & vbTab & "</nodes>" & vbCrLf _
                                & vbTab & "<edges>" & vbCrLf _
    
    '************************ Then create edges *******************************++
    theCounter = 0
    
    For Each theTask In ActiveProject.Tasks
        If theTask.taskDependencies.Count > 0 Then
            'For Each taskDependency In theTask.taskDependencies
                'First analyze predecessors
                If InStr(theTask.taskDependencies.Parent.UniqueIDPredecessors, ",") = 0 And theTask.taskDependencies.Parent.UniqueIDPredecessors <> "" Then
                    If InStr(theTask.taskDependencies.Parent.UniqueIDPredecessors, "+") <> 0 Then
                        'Filter tasks which have prolongations: <taskID>AA/EA+# <time unit>
                        edgeSource = Left(theTask.taskDependencies.Parent.UniqueIDPredecessors, 4)
                    Else
                        edgeSource = theTask.taskDependencies.Parent.UniqueIDPredecessors
                    End If
                    edgeSourceTarget = """ source=""" & edgeSource & """ target=""" & theTask.UniqueID
                    If InStr(textToExport, edgeSourceTarget) = 0 Then
                    'If source -> target is not already included in edge list
                        textToExport = textToExport & vbTab & vbTab & "<edge id=""" & CStr(theCounter) & edgeSourceTarget & """ />" & vbCrLf
                        theCounter = theCounter + 1
                    End If
                Else
                    linkArray = Split(theTask.taskDependencies.Parent.UniqueIDPredecessors, ",")
                    linkArrayLength = UBound(linkArray)
                    If linkArrayLength > 0 Then
                        For arrayCounter = 0 To linkArrayLength
                            If linkArray(arrayCounter) <> CStr(theTask.UniqueID) Then
                                If InStr(linkArray(arrayCounter), "+") <> 0 Then
                                    'Filter tasks which have prolongations: <taskID>AA/EA+# <time unit>
                                    edgeSource = Left(linkArray(arrayCounter), 4)
                                Else
                                    edgeSource = linkArray(arrayCounter)
                                End If
                                edgeSourceTarget = """ source=""" & edgeSource & """ target=""" & theTask.UniqueID
                                If InStr(textToExport, edgeSourceTarget) = 0 Then
                                    textToExport = textToExport & vbTab & vbTab & "<edge id=""" & CStr(theCounter) & edgeSourceTarget & """ />" & vbCrLf
                                    theCounter = theCounter + 1
                                End If
                            End If
                        Next
                    End If
                End If
                
                'Then analyze Successors
                'linkArray = Split(taskDependency.To.UniqueIDSuccessors, ",")
                'linkArrayLength = UBound(linkArray)
                'If linkArrayLength > 0 Then
                '    For arrayCounter = 0 To linkArrayLength - 1
                '        If linkArray(arrayCounter) <> CStr(theTask.UniqueID) Then
                '            textToExport = textToExport & "<edge id=" & Chr(34) & CStr(theCounter) & Chr(34) & " source=" & Chr(34) & linkArray(arrayCounter) & Chr(34) & " target=" & Chr(34) & theTask.UniqueID & Chr(34) & " />" & vbCrLf
                '            theCounter = theCounter + 1
                '        End If
                '    Next
                'End If
            'Next taskDependency
        End If
    Next theTask
    
    'Trailing XML
    textToExport = textToExport & vbTab & "</edges>" & vbCrLf _
                                & "</graph>" & vbCrLf _
                                & "</gexf>"
    
    'Some cleaning
    textToExport = Replace(textToExport, "&", "+")
    
    'Write to file
    filePath = "F:\Cryring\dependencies.gexf"
    Open filePath For Output As #1
    Print #1, textToExport
    Close #1
    
End Sub

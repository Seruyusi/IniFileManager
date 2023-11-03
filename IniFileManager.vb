Imports System.IO

Public Module IniFileManager

    Public Sub WriteValue(ByVal iniFilePath As String, ByVal section As String, ByVal key As String, ByVal value As String)
        Dim lines As List(Of String) = ReadAllLines(iniFilePath)

        Dim sectionFound As Boolean = False
        Dim keyFound As Boolean = False

        For i As Integer = 0 To lines.Count - 1
            Dim currentLine As String = lines(i)

            If currentLine.StartsWith("[") AndAlso currentLine.EndsWith("]") AndAlso currentLine.Substring(1, currentLine.Length - 2) = section Then
                sectionFound = True
            End If

            If sectionFound AndAlso currentLine.StartsWith(key & "=") Then
                lines(i) = key & "=" & value
                keyFound = True
            End If
        Next

        If Not sectionFound Then
            lines.Add("[" & section & "]")
            lines.Add(key & "=" & value)
        ElseIf sectionFound AndAlso Not keyFound Then
            lines.Add(key & "=" & value)
        End If

        WriteAllLines(iniFilePath, lines)
    End Sub

    Public Sub RemoveKey(ByVal iniFilePath As String, ByVal section As String, ByVal key As String)
        Dim lines As List(Of String) = ReadAllLines(iniFilePath)

        Dim sectionFound As Boolean = False
        Dim keyFound As Boolean = False

        For i As Integer = 0 To lines.Count - 1
            Dim currentLine As String = lines(i)

            If currentLine.StartsWith("[") AndAlso currentLine.EndsWith("]") AndAlso currentLine.Substring(1, currentLine.Length - 2) = section Then
                sectionFound = True
            End If

            If sectionFound AndAlso currentLine.StartsWith(key & "=") Then
                lines.RemoveAt(i)
                keyFound = True
                Exit For ' Exiting the loop since the key is removed
            End If
        Next

        If sectionFound AndAlso keyFound Then
            WriteAllLines(iniFilePath, lines)
        End If
    End Sub

    Public Function ReadValue(ByVal iniFilePath As String, ByVal section As String, ByVal key As String) As String
        Dim lines As List(Of String) = ReadAllLines(iniFilePath)

        Dim sectionFound As Boolean = False
        Dim value As String = Nothing

        For Each line As String In lines
            If line.StartsWith("[") AndAlso line.EndsWith("]") AndAlso line.Substring(1, line.Length - 2) = section Then
                sectionFound = True
            End If

            If sectionFound AndAlso line.StartsWith(key & "=") Then
                value = line.Substring(key.Length + 1)
                Exit For ' Exiting the loop since the value is found
            End If
        Next

        Return value
    End Function

    Private Function ReadAllLines(ByVal filePath As String) As List(Of String)
        Dim lines As New List(Of String)

        If File.Exists(filePath) Then
            lines.AddRange(File.ReadAllLines(filePath))
        End If

        Return lines
    End Function

    Private Sub WriteAllLines(ByVal filePath As String, ByVal lines As List(Of String))
        File.WriteAllLines(filePath, lines)
    End Sub

End Module

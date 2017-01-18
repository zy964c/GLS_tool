Imports INFITF
Imports System.IO

Module Module1
    Dim catia As INFITF.Application = CreateObject("CATIA.Application")
    Sub Main()
        Dim documents = catia.Documents
        Dim carm_pn1 As String
        Dim dict As New Dictionary(Of String, String)
        For k As Integer = 1 To documents.Count
            Dim doc_name As String = documents.Item(k).Name
            If doc_name.Contains("CA836Z") And doc_name.Contains(".CATPart") Then
                carm_pn1 = doc_name
                'Debug.Print(doc_name)
                Dim carm_doc = documents.Item(carm_pn1)
                Dim carm_part = carm_doc.Part
                Dim hybrid_bodies = carm_part.HybridBodies
                Dim joint As String = "Construction Geometry"
                Dim jd As Object
                Try
                    jd = hybrid_bodies.Item(joint)
                Catch
                    Continue For
                End Try
                Dim hybrid_bodies1 = jd.HybridBodies
                Dim flagname As String = "flagnote"
                Dim flagnote As Object
                Try
                    flagnote = hybrid_bodies1.Item(flagname)
                Catch
                    Continue For
                End Try
                Dim strPath As String = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase)

                'If hybrid_bodies1.Count > 0 Then
                'For i As Integer = 1 To hybrid_bodies1.Count
                Dim hybrid_shapes = flagnote.HybridShapes
                If hybrid_shapes.Count > 0 Then
                    For j As Integer = 1 To (hybrid_shapes.Count)
                        Dim point = hybrid_shapes.Item(j)
                        Dim PointComponentsArray(2)
                        point.GetCoordinates(PointComponentsArray)
                        For Each coord As Integer In PointComponentsArray
                            'Debug.Print(coord / 25.4)
                        Next
                        Dim Result As String = String.Join(", ", (PointComponentsArray))
                        Dim q As Integer = 0
                        Dim init_name As String = point.name
                        While dict.ContainsKey(point.name)
                            q += 1
                            point.name = init_name & "." & q.ToString
                        End While
                        Try
                            dict.Add(point.Name, Result)
                        Catch
                            Continue For
                        End Try

                        Result = ""

                    Next
                End If

                Dim json As String = Newtonsoft.Json.JsonConvert.SerializeObject(dict, Newtonsoft.Json.Formatting.Indented)
                strPath = strPath.Replace("file:\", "")
                System.IO.File.WriteAllText(strPath & "\flagnote_" & carm_pn1.Replace(".CATPart", "") & ".txt", json)
                dict.Clear()
            End If

        Next
    End Sub

End Module


Imports INFITF
Imports System.IO

Module jd2

    Dim catia As INFITF.Application = CreateObject("CATIA.Application")
    Sub jd()
        Dim strPath As String = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase)
        Dim documents = catia.Documents
        Dim carm_pn As String
        Dim dict As New Dictionary(Of String, String)
        For k As Integer = 1 To documents.Count
            Dim doc_name As String = documents.Item(k).Name
            If doc_name.Contains("CA836Z") And doc_name.Contains(".CATPart") Then
                carm_pn = doc_name
                'Debug.Print(doc_name)
                Dim carm_doc = documents.Item(carm_pn)
                Dim carm_part = carm_doc.Part
                Dim hybrid_bodies = carm_part.HybridBodies
                Dim joint As String = "Joint Definitions"
                Dim jd = hybrid_bodies.Item(joint)
                Dim hybrid_bodies1 = jd.HybridBodies
                If hybrid_bodies1.Count > 0 Then
                    For i As Integer = 1 To hybrid_bodies1.Count
                        Dim hybrid_shapes = hybrid_bodies1.Item(i).HybridShapes
                        If hybrid_shapes.Count > 0 Then
                            For j As Integer = 1 To (hybrid_shapes.Count)
                                Dim point = hybrid_shapes.Item(j)
                                Dim point_name As String = point.Name
                                Dim vector As String = "FIDV"
                                Dim check As Boolean = point_name.Contains(vector)
                                If check Then Continue For
                                Dim PointComponentsArray(2)
                                point.GetCoordinates(PointComponentsArray)
                                For Each coord As Integer In PointComponentsArray
                                    'Debug.Print(coord / 25.4)
                                Next
                                Dim Result As String = String.Join(", ", (PointComponentsArray))
                                Try
                                    dict.Add(point.Name, Result)
                                Catch
                                    Continue For
                                End Try

                                Result = ""

                            Next
                        End If
                    Next
                End If
                Dim json As String = Newtonsoft.Json.JsonConvert.SerializeObject(dict, Newtonsoft.Json.Formatting.Indented)
                Dim strPath_short = strPath.Replace("file:\", "")
                Dim strPath_temp = strPath_short & "\temp"
                If (Not System.IO.Directory.Exists(strPath_temp)) Then
                    System.IO.Directory.CreateDirectory(strPath_temp)
                End If
                System.IO.File.WriteAllText(strPath_temp & "\" & carm_pn.Replace(".CATPart", "") & ".txt", json)
                dict.Clear()

            End If
        Next
    End Sub

End Module

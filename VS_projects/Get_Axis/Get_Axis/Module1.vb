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
            If doc_name.Contains("seed_ctr") Then
                carm_pn1 = doc_name
                'Debug.Print(doc_name)
                Dim carm_doc = documents.Item(carm_pn1)
                Dim carm_part = carm_doc.Part
                Dim axis_systems = carm_part.AxisSystems

                If axis_systems.Count > 0 Then
                    For j As Integer = 1 To (axis_systems.Count)
                        Dim axis = axis_systems.Item(j)
                        Dim SummaryArray(11)
                        Dim AxisOriginComponentsArray(2)
                        Dim AxisXComponentsArray(2)
                        Dim AxisYComponentsArray(2)
                        Dim AxisZComponentsArray(2)

                        axis.GetOrigin(AxisOriginComponentsArray)
                        axis.GetXAxis(AxisXComponentsArray)
                        axis.GetYAxis(AxisYComponentsArray)
                        axis.GetZAxis(AxisZComponentsArray)

                        Array.Copy(AxisOriginComponentsArray, SummaryArray, AxisOriginComponentsArray.Length)
                        Array.Copy(AxisXComponentsArray, 0, SummaryArray, AxisOriginComponentsArray.Length, AxisXComponentsArray.Length)
                        Array.Copy(AxisYComponentsArray, 0, SummaryArray, AxisOriginComponentsArray.Length + AxisXComponentsArray.Length, AxisYComponentsArray.Length)
                        Array.Copy(AxisZComponentsArray, 0, SummaryArray, AxisOriginComponentsArray.Length + AxisXComponentsArray.Length + AxisYComponentsArray.Length, AxisZComponentsArray.Length)

                        For a As Integer = 0 To SummaryArray.GetUpperBound(0)
                            SummaryArray(a) = Math.Round(SummaryArray(a), 4)
                        Next


                        'For Each coord As Integer In AxisOriginComponentsArray
                        'Debug.Print(coord)
                        'Next
                        'For Each coord As Integer In AxisXComponentsArray
                        'Debug.Print(coord)
                        'Next
                        'For Each coord As Integer In AxisYComponentsArray
                        'Debug.Print(coord)
                        'Next
                        'For Each coord As Integer In AxisZComponentsArray
                        'Debug.Print(coord)
                        'Next
                        Dim Result As String = String.Join(", ", (SummaryArray))
                        dict.Add(axis.Name, Result)
                        Result = ""

                    Next
                End If


            End If

        Next
        Dim json As String = Newtonsoft.Json.JsonConvert.SerializeObject(dict, Newtonsoft.Json.Formatting.Indented)
        System.IO.File.WriteAllText("Y:\My Documents\Python Scripts\axis.txt", json)
        'dict.Clear()

    End Sub


End Module

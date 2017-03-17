Imports INFITF
Imports System.IO

Module Module1
    Dim catia As INFITF.Application = CreateObject("CATIA.Application")
    Sub Main()
        Dim documents = catia.Documents
        Dim dict As New Dictionary(Of String, String)
        For k As Integer = 1 To documents.Count
            Dim doc_name As String = documents.Item(k).Name
            If doc_name.Contains("seed_ctr") Then
                Dim carm_doc = documents.Item(doc_name)
                Dim carm_cameras = carm_doc.Cameras
                For l As Integer = 1 To carm_cameras.Count
                    Dim viewpoint = carm_cameras.Item(l)
                    Dim vp3d = viewpoint.Viewpoint3D
                    Dim SummaryArray(8)
                    Dim CameraCoordArray(2)
                    Dim CameraSightDirArray(2)
                    Dim CameraUpDirArray(2)
                    vp3d.GetOrigin(CameraCoordArray)
                    vp3d.GetSightDirection(CameraSightDirArray)
                    vp3d.GetUpDirection(CameraUpDirArray)
                    CameraCoordArray(0) = (CameraCoordArray(0) - 699.0 * 25.4)

                    Array.Copy(CameraCoordArray, SummaryArray, CameraCoordArray.Length)
                    Array.Copy(CameraSightDirArray, 0, SummaryArray, CameraCoordArray.Length, CameraSightDirArray.Length)
                    Array.Copy(CameraUpDirArray, 0, SummaryArray, CameraCoordArray.Length + CameraSightDirArray.Length, CameraUpDirArray.Length)

                    Dim Result As String = String.Join(", ", (SummaryArray))
                    dict.Add(viewpoint.Name, Result)
                    Result = ""

                Next
            End If
        Next
        Dim json As String = Newtonsoft.Json.JsonConvert.SerializeObject(dict, Newtonsoft.Json.Formatting.Indented)
        System.IO.File.WriteAllText("Y:\My Documents\Python Scripts\cameras.txt", json)
    End Sub

End Module

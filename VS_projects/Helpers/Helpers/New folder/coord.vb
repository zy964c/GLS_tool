
Imports INFITF
Imports System.IO

Module Module1
    Sub Main()

        Dim catia As INFITF.Application = CreateObject("CATIA.Application")
        Dim list1 As New List(Of VariantType)
        'Dim dict As New Dictionary(Of String, List(Of VariantType))
        Dim dict As New Dictionary(Of String, String)
        Dim productDocument1 = catia.ActiveDocument
        Dim product = productDocument1.Product
        Dim collection = product.Products
        Dim strPath As String = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase)

        'Dim dict As New Dictionary(Of String, Array)

        For i As Integer = 1 To collection.Count
            'Debug.Print(collection.Item(i).Name)
            Dim product_inside = collection.Item(i)
            Dim collection_inside = product_inside.Products

            If collection_inside.Count > 0 Then

                For j As Integer = 1 To collection_inside.Count
                    Dim product_current = collection_inside.Item(j)
                    'Debug.Print(product_current.Name)
                    Dim oAxisComponentsArray(11)
                    product_current.Position.GetComponents(oAxisComponentsArray)
                    'For Each elem As Object In oAxisComponentsArray
                    'list1.Add(elem)
                    'Next
                    Dim Result As String = String.Join(", ", (oAxisComponentsArray))
                    Try
                        dict.Add(product_current.Name, Result)
                    Catch
                        Continue For
                    End Try

                    Result = ""
                    For Each coord As Integer In oAxisComponentsArray
                        'Debug.Print(coord)
                    Next

                Next

            End If

        Next

        Dim list2 As New List(Of String)(dict.Keys)
        Dim str As String
        For Each str In list2
            'Debug.Print("{0}, {1}", str, dict.Item(str))
        Next

        Dim json As String = Newtonsoft.Json.JsonConvert.SerializeObject(dict, Newtonsoft.Json.Formatting.Indented)
        'Debug.Print(json)
        'Console.WriteLine(json)
        'Console.ReadLine()
        'System.IO.File.WriteAllText("C:\Temp\GLS\coord.txt", json)
        'Debug.Print(strPath)
        strPath = strPath.Replace("file:\", "")
        'Debug.Print(strPath)
        System.IO.File.WriteAllText(strPath & "\coord.txt", json)

    End Sub
End Module


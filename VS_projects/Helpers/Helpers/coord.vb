
Imports INFITF
Imports System.IO

Module coord1

    Sub coord()

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
                    Dim collection_inside1 = product_current.Products
                    If collection_inside1.Count > 0 Then

                        For k As Integer = 1 To collection_inside1.Count
                            Dim product_current1 = collection_inside1.Item(k)
                            Dim collection_inside2 = product_current1.Products
                            If collection_inside2.Count > 0 Then
                                For m As Integer = 1 To collection_inside2.Count
                                    Dim product_current2 = collection_inside2.Item(m)
                                    Dim collection_inside3 = product_current2.Products
                                    Dim oAxisComponentsArray2(11)
                                    product_current1.Position.GetComponents(oAxisComponentsArray2)
                                    Dim Result2 As String = String.Join(", ", (oAxisComponentsArray2))
                                    Try
                                        'dict.Add(product_current.Name & "_" & product_current1.Name, Result1)
                                        dict.Add(product_current2.Name, Result2)
                                    Catch
                                        Continue For
                                    End Try
                                    Result2 = ""
                                    For Each coord As Integer In oAxisComponentsArray2
                                        'Debug.Print(coord)
                                    Next
                                Next
                            End If
                            Dim oAxisComponentsArray1(11)
                            product_current.Position.GetComponents(oAxisComponentsArray1)
                            Dim Result1 As String = String.Join(", ", (oAxisComponentsArray1))
                            Try
                                'dict.Add(product_current.Name & "_" & product_current1.Name, Result1)
                                dict.Add(product_current1.Name, Result1)
                            Catch
                                Continue For
                            End Try
                            Result1 = ""
                            For Each coord As Integer In oAxisComponentsArray1
                                'Debug.Print(coord)
                            Next
                        Next
                    End If
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
        Debug.Print(json)
        'Console.WriteLine(json)
        'Console.ReadLine()
        'System.IO.File.WriteAllText("C:\Temp\GLS\coord.txt", json)
        'Debug.Print(strPath)
        strPath = strPath.Replace("file:\", "")
        strPath = strPath & "\temp"
        If (Not System.IO.Directory.Exists(strPath)) Then
            System.IO.Directory.CreateDirectory(strPath)
        End If
        'Debug.Print(strPath)
        System.IO.File.WriteAllText(strPath & "\coord.txt", json)

    End Sub
End Module


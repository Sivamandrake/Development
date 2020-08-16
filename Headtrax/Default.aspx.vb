Imports Excel = Microsoft.Office.Interop.Excel
Public Class _Default
    Inherits Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load

    End Sub

    Protected Sub btnSubmit_Click(sender As Object, e As EventArgs) Handles btnSubmit.Click

        Dim MyAlias() As String = txtAlias.Text.Split(",")
        Dim people As who.PeopleStore = New who.PeopleStore()
        people.Credentials = System.Net.CredentialCache.DefaultCredentials
        Dim result As who.PersonContext = New who.PersonContext()
        Dim excelApp As Excel.Application
        Dim excelDoc As Excel.Workbook
        Dim excelSheet As Excel.Worksheet
        Dim range As Excel.Range

        ' Start Excel and get Application object.
        excelApp = CreateObject("Excel.Application")
        excelApp.Visible = True

        excelDoc = excelApp.Workbooks.Add
        excelSheet = excelDoc.ActiveSheet

        Dim DS As DataSet = New DataSet()
        Dim count As New Integer
        Dim i As New Integer
        Dim counter As New Integer

        Try

            counter = 0
            For Each Item As String In MyAlias
                Try

                    result = people.FindPersonContextByAlias(Item)

                    count = result.Managers.Length

                    For i = 0 To count - 1
                        Try
                            'list.Items.Add(result.Managers(i).Name.ToString())

                            'list.Items.Add(result.Managers(i).Alias.ToString())

                            'list.Items.Add(result.Managers(i).Department.ToString())

                            'list.Items.Add(result.Managers(i).Title.ToString())


                            excelSheet.Cells(counter + 1, 1).Value = result.Managers(i).Name.ToString()

                            excelSheet.Cells(counter + 1, 2).Value = result.Managers(i).Alias.ToString()

                            'excelSheet.Cells(counter + 1, 2).Value = result.Managers(i).SamAccountName.ToString()

                            excelSheet.Cells(counter + 1, 3).Value = result.Managers(i).Department.ToString()

                            excelSheet.Cells(counter + 1, 4).Value = result.Managers(i).Title.ToString()

                            excelSheet.Cells(counter + 1, 5).Value = result.Managers(i).Office.ToString()

                            excelSheet.Cells(counter + 1, 6).Value = "L" + (count - i).ToString()

                            counter += 1


                        Catch ex As Exception
                            Continue For
                        End Try
                    Next

                    'list.Items.Add(result.Person.Name.ToString())
                    'list.Items.Add(result.Person.Alias.ToString())
                    'list.Items.Add(result.Person.Department.ToString())
                    'list.Items.Add(result.Person.Title.ToString())

                    excelSheet.Cells(counter + 1, 1).Value = result.Person.Name.ToString()

                    excelSheet.Cells(counter + 1, 2).Value = result.Person.Alias.ToString()

                    'excelSheet.Cells(counter + 1, 2).Value = result.Person.SamAccountName.ToString()

                    excelSheet.Cells(counter + 1, 3).Value = result.Person.Department.ToString()

                    excelSheet.Cells(counter + 1, 4).Value = result.Person.Title.ToString()

                    excelSheet.Cells(counter + 1, 5).Value = result.Person.Office.ToString()

                    excelSheet.Cells(counter + 1, 6).Value = "L0"

                    counter += 1

                Catch ex As Exception
                    Continue For
                End Try


            Next


            excelDoc = Nothing
            excelSheet = Nothing
            excelApp.Quit()
            excelApp = Nothing

        Catch ex As Exception

        End Try
    End Sub
End Class
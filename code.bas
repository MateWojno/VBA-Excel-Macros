Sub transform_algorithm_1_and_2()
'
' transform Macro
'

'
    ActiveWorkbook.Queries.Add Name:="RaportProdukcji", Formula:= _
        "let" & Chr(13) & "" & Chr(10) & "    Source = Access.Database(File.Contents(""C:\Users\PC\Desktop\w-db_files\wdb.mdb""), [CreateNavigationProperties=true])," & Chr(13) & "" & Chr(10) & "    _RaportProdukcji = Source{[Schema="""",Item=""RaportProdukcji""]}[Data]," & Chr(13) & "" & Chr(10) & "    #""Removed Other Columns"" = Table.SelectColumns(_RaportProdukcji,{""nr_raportu"", ""data"", ""kod_receptury"", ""nazwa_receptury"", ""zamowiono"", ""wyrodu" & _
        "kowano"", ""zamowiono_colosc"", ""wyslano"", ""samochod"", ""samochod_kierowca"", ""pompa"", ""pompa_kierowca"", ""klient"", ""klient2"", ""budowa"", ""budowa2""})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    #""Removed Other Columns"""
        ActiveWorkbook.Worksheets.Add
        ActiveSheet.Name = "wdb"
     With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=RaportProdukcji;Extended Properties=""""" _
        , Destination:=Range("$A$1")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [RaportProdukcji]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.DisplayName = "RaportProdukcji"
        .refresh BackgroundQuery:=False
    End With
End Sub

' te makro bierze plik bazodanowy wczytuje go i transformuje - jest to odpowiednik algorytmu_1 i algorytmu_2



Sub read_algorythm_1()

    ActiveWorkbook.Queries.Add Name:="RaportProdukcji", Formula:= _
        "let" & Chr(13) & "" & Chr(10) & "    Source = Access.Database(File.Contents(""C:\Users\PC\Desktop\w-db_files\wdb.mdb""), [CreateNavigationProperties=true])," & Chr(13) & "" & Chr(10) & "    _RaportProdukcji = Source{[Schema="""",Item=""RaportProdukcji""]}[Data]" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    _RaportProdukcji"
    ActiveWorkbook.Worksheets.Add
    ActiveSheet.Name = "wdb"
    With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=RaportProdukcji;Extended Properties=""""" _
        , Destination:=Range("$A$1")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [RaportProdukcji]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.DisplayName = "RaportProdukcji"
        .refresh BackgroundQuery:=False
    End With
End Sub

' te makro bierze plik bazodanowy i wczytuje go - jest to odpowiednik algorytmu_1


Sub refresh_algorythm_3()
    ActiveWorkbook.Queries("RaportProdukcji").Delete
    Sheets("wdb").Select
    ActiveWindow.SelectedSheets.Delete
End Sub

' powinno usunac query i "wdb"

  
  Sub Picture10_Click()
MsgBox ("Pierwszy od lewej przycisk - wczytywanie bazy, drugi -  usuwanie bazy")
End Sub

    
    

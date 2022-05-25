Attribute VB_Name = "Mod_MasterToCatalogue"
Sub ImportarDatosMasterData()

On Error GoTo error

'Declaración variables libro origen
Dim wbLibroOrigen As Workbook
Dim wsHojaOrigenCategory As Worksheet
Dim wsHojaOrigenFamily As Worksheet
Dim wsHojaOrigenSubFamily As Worksheet
Dim wsHojaOrigenItems As Worksheet
Dim wsHojaOrigenSuppliers As Worksheet
Dim wsHojaOrigenLinkArtSup As Worksheet
Dim wsHojaOrigenDataMaster As Worksheet

'Declaracion variables del libro destino
Dim wbLibroDestino As Workbook
Dim wsHojaDestinoCategory As Worksheet
Dim wsHojaDestinoFamily As Worksheet
Dim wsHojaDestinoSubFamily As Worksheet
Dim wsHojaDestinoArticles As Worksheet
Dim wsHojaDestinoSuppliers As Worksheet
Dim wsHojaDestinoLinkArtSup As Worksheet
Dim wsHojaDestinoDataMaster As Worksheet

'Declaracion variable para obtener la ruta del archivo
Dim ruta As String

'Ruta del archivo excel con el master data
ruta = abrirArchivo
Set wbLibroOrigen = Workbooks.Open(ruta)


'Datos destino
Set wbLibroDestino = Workbooks(ThisWorkbook.Name)
Set wsHojaDestinoCategory = wbLibroDestino.Worksheets("CATEGORY")
Set wsHojaDestinoFamily = wbLibroDestino.Worksheets("FAMILY")
Set wsHojaDestinoSubFamily = wbLibroDestino.Worksheets("SUB-FAMILY")
Set wsHojaDestinoArticles = wbLibroDestino.Worksheets("ARTICLES")
Set wsHojaDestinoSuppliers = wbLibroDestino.Worksheets("SUPPLIERS")
Set wsHojaDestinoLinkArtSup = wbLibroDestino.Worksheets("Link art-sup")
Set wsHojaDestinoDataMaster = wbLibroDestino.Worksheets("DATA_MASTER")




'Comprobamos que el archivo contiene una hoja llamada categoria e items para asegurarnos que es Master Data
If BuscarHoja("CATEGORY", wbLibroOrigen) And BuscarHoja("ITEMS", wbLibroOrigen) Then
    
    Set wsHojaOrigenCategory = wbLibroOrigen.Worksheets("CATEGORY")
    Set wsHojaOrigenFamily = wbLibroOrigen.Worksheets("FAMILY")
    Set wsHojaOrigenSubFamily = wbLibroOrigen.Worksheets("SUB-FAMILY")
    Set wsHojaOrigenItems = wbLibroOrigen.Worksheets("ITEMS")
    Set wsHojaOrigenSuppliers = wbLibroOrigen.Worksheets("SUPPLIERS")
    Set wsHojaOrigenLinkArtSup = wbLibroOrigen.Worksheets("Link art-sup")
    Set wsHojaOrigenDataMaster = wbLibroOrigen.Worksheets("DATA_MASTER")

Else
    'Si el archivo no es Master Data lanzamos error, cerramos el archivo y salimos de la macro
    MsgBox "El archivo elegido no es de tipo Master Data", vbInformation
    Workbooks(wbLibroOrigen.Name).Close Savechanges:=False
    Exit Sub

End If



'Obtenemos la ultima fila con datos de cada hoja de la master data
uFilaCategory = wsHojaOrigenCategory.Range("A" & Rows.Count).End(xlUp).Row
uFilaFamily = wsHojaOrigenFamily.Range("A" & Rows.Count).End(xlUp).Row
uFilaSubFamily = wsHojaOrigenSubFamily.Range("A" & Rows.Count).End(xlUp).Row
uFilaItems = wsHojaOrigenItems.Range("A" & Rows.Count).End(xlUp).Row
uFilaSuppliers = wsHojaOrigenSuppliers.Range("A" & Rows.Count).End(xlUp).Row
uFilaDataMasterB = wsHojaOrigenDataMaster.Range("B" & Rows.Count).End(xlUp).Row
uFilaDataMasterC = wsHojaOrigenDataMaster.Range("C" & Rows.Count).End(xlUp).Row
uFilaDataMasterF = wsHojaOrigenDataMaster.Range("F" & Rows.Count).End(xlUp).Row




'Copiamos los datos de la hojas del master data y los pegamos en las hojas del libro catalogue

 
Call copiarPegarRango(wsHojaDestinoCategory, wsHojaOrigenCategory.Range("A2:B" & uFilaCategory), wsHojaDestinoCategory.Range("A2"))

Call copiarPegarRango(wsHojaDestinoFamily, wsHojaOrigenFamily.Range("A2:C" & uFilaFamily), wsHojaDestinoFamily.Range("A2"))

Call copiarPegarRango(wsHojaDestinoSubFamily, wsHojaOrigenSubFamily.Range("A2:C" & uFilaSubFamily), wsHojaDestinoSubFamily.Range("A2"))

Call copiarPegarRango(wsHojaDestinoArticles, wsHojaOrigenItems.Range("A2:F" & uFilaItems), wsHojaDestinoArticles.Range("A2"))

Call ReemplazarTipoArt(wsHojaOrigenItems, wsHojaDestinoArticles)

Call copiarPegarRango(wsHojaDestinoArticles, wsHojaOrigenItems.Range("I2:J" & uFilaItems), wsHojaDestinoArticles.Range("H2"))

Call copiarPegarRango(wsHojaDestinoSuppliers, wsHojaOrigenSuppliers.Range("A2:M" & uFilaSuppliers), wsHojaDestinoSuppliers.Range("A2"))

Call copiarPegarRango(wsHojaDestinoLinkArtSup, wsHojaOrigenItems.Range("B2:B" & uFilaItems), wsHojaDestinoLinkArtSup.Range("B2"))

Call copiarPegarRango(wsHojaDestinoLinkArtSup, wsHojaOrigenItems.Range("Y2:Y" & uFilaItems), wsHojaDestinoLinkArtSup.Range("A2"))

Call copiarPegarRango(wsHojaDestinoLinkArtSup, wsHojaOrigenItems.Range("Z2:AF" & uFilaItems), wsHojaDestinoLinkArtSup.Range("C2"))

Call copiarPegarRango(wsHojaDestinoDataMaster, wsHojaOrigenDataMaster.Range("B2:B" & uFilaDataMasterB), wsHojaDestinoDataMaster.Range("B2"))

Call copiarPegarRango(wsHojaDestinoDataMaster, wsHojaOrigenDataMaster.Range("C2:C" & uFilaDataMasterC), wsHojaDestinoDataMaster.Range("C2"))

Call copiarPegarRango(wsHojaDestinoDataMaster, wsHojaOrigenDataMaster.Range("F2:F" & uFilaDataMasterF), wsHojaDestinoDataMaster.Range("I2"))



'Posicionamos el foco en la primera hoja
wsHojaDestinoCategory.Activate

error:
If wbLibroOrigen Is Nothing Then
    MsgBox "no ha seleccionado ningún archivo", vbInformation
Else
    Workbooks(wbLibroOrigen.Name).Close Savechanges:=False
End If


End Sub

Function IngresarRuta() As String

'Ingresar la ruta de la Master Data
MsgBox "Elige la ruta del archivo que contiene la Master Data", vbInformation

Dim i As Long

With Application.FileDialog(msoFileDialogFilePicker)
    .Show
    
    For i = 1 To .SelectedItems.Count
    IngresarRuta = .SelectedItems(i)
    Next i
    
End With


End Function


Function BuscarHoja(nombreHoja As String, libro As Workbook) As Boolean

'Funcion para comprobar que existan las hojas de la Master Data

Dim IsExist As Boolean
IsExist = False
For i = 1 To libro.Sheets.Count
    If libro.Sheets(i).Name = nombreHoja Then
        IsExist = True
        Exit For
    End If
Next
BuscarHoja = IsExist
End Function

Function abrirArchivo() As String

    Dim fd As Office.FileDialog
   
     
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
     
    With fd
     
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx?", 1
        .Title = "Seleccione un archivo excel con la Master Data"
        .AllowMultiSelect = False
     
        If .Show = True Then
     
            abrirArchivo = .SelectedItems(1)
     
        End If
     
    End With
End Function

Sub copiarPegarRango(hojaDestino As Worksheet, rangoOrigen As Range, rangoDestino As Range)

'funcion para copiar y pegar los datos y deselccionar los rangos de celdas pegadas

rangoOrigen.Copy
rangoDestino.PasteSpecial xlPasteValues

hojaDestino.Activate
hojaDestino.Range("A1").Select


End Sub



Sub ReemplazarTipoArt(hojaOrigen As Worksheet, hojaDestino As Worksheet)

'funcion para coger los datos de los tipos de artículo,traducirlos si es necesario  y pegar en destino

Dim miMatriz As Variant

uFila = hojaOrigen.Range("G" & Rows.Count).End(xlUp).Row
miMatriz = hojaOrigen.Range("G2:G" & uFila).Value

For i = LBound(miMatriz) To UBound(miMatriz)
    Item = miMatriz(i, 1)
        Select Case Item
                Case "ARTICLE"
                    miMatriz(i, 1) = "ARTICULO"
                Case "SHOP", "BOUTIQUE"
                    miMatriz(i, 1) = "TIENDA"
                Case "SERVICES"
                    miMatriz(i, 1) = "SERVICIOS"
                Case "CARBURANT", "FUEL", "FORECOURT"
                    miMatriz(i, 1) = "COMBUSTIBLE"
                    
        End Select
            
Next i

hojaDestino.Range("G2:G" & uFila) = miMatriz

End Sub

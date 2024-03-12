Imports DotSpatial.Controls
Imports DotSpatial.Data
Imports DotSpatial.Symbology
Imports DotSpatial.Topology


Public Enum MapShapeType
    Point
    Line
    Polygon
End Enum

Public Class Form1

    Private ShapeType As MapShapeType

#Region "Point shapefile class level variable"
    'the new point feature set
    Dim pointF As New FeatureSet(FeatureType.Point)
    'the id of point
    Dim pointID As Integer = 0
    'to differentiate the right and left mouse click
    Dim pointmouseClick As Boolean = False
#End Region
#Region "Polyline shapefile class level variables"
    'the line layer
    Dim lineLayer As MapLineLayer
    'the line feature set
    Dim lineF As New FeatureSet(FeatureType.Line)
    Dim lineID As Integer = 0
    'boolean variable for first time mouse click
    Dim firstClick As Boolean = False
    'It controls the drawing the polyline after the polyline saved operation.
    Dim linemouseClick As Boolean = False
#End Region
#Region "Polygon shapefile class level variables"
    'the polygon feature set
    Dim polygonF As New FeatureSet(FeatureType.Polygon)
    Dim polygonID As Integer = 0
    Dim polygonmouseClick As Boolean = False
#End Region

    Private Sub CreatePointShapefileToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CreatePointShapefileToolStripMenuItem.Click
        'Change the cursor style
        map1.Cursor = Cursors.Cross

        'set the shape type to the class level string variable
        'we are going to use this variable in a select case statement
        ShapeType = MapShapeType.Point

        'set projection
        pointF.Projection = map1.Projection

        'initialize the featureSet attribute table
        Dim column As New DataColumn("ID")
        pointF.DataTable.Columns.Add(column)

        'add the featureSet as a map layer
        Dim pointLayer As MapPointLayer = map1.Layers.Add(pointF)

        'Create a new symbolizer
        Dim symbol As New PointSymbolizer(Color.Red, DotSpatial.Symbology.PointShape.Ellipse, 3)

        'Set the symbolizer to the point layer
        pointLayer.Symbolizer = symbol

        'Set the legendText as point
        pointLayer.LegendText = "point"

        'Set left mouse click as true
        pointmouseClick = True

    End Sub

    Private Sub Map1_MouseDown(sender As Object, e As MouseEventArgs) Handles Map1.MouseDown
        Select Case ShapeType
            Case MapShapeType.Point
                If e.Button = MouseButtons.Left Then
                    If (pointmouseClick) Then
                        'This method is used to convert the screen coordinate to map coordinate
                        'e.location is the mouse click point on the map control
                        Dim coord As Coordinate = map1.PixelToProj(e.Location)
                        'Create a new point
                        'Input parameter is clicked point coordinate
                        Dim point As New DotSpatial.Topology.Point(coord)
                        'Add the point into the Point Feature
                        'assigning the point feature to IFeature because via it only we can set the attributes.
                        Dim currentFeature As IFeature = pointF.AddFeature(point)
                        'increase the point id
                        pointID = pointID + 1
                        'set the ID attribute
                        currentFeature.DataRow("ID") = pointID
                        'refresh the map
                        map1.ResetBuffer()
                    End If
                Else
                    map1.Cursor = Cursors.Default
                    pointmouseClick = False
                End If

            Case MapShapeType.Line
                If e.Button = MouseButtons.Left Then
                    'left click - fill array of coordinates
                    'coordinate of clicked point
                    Dim coord As Coordinate = map1.PixelToProj(e.Location)
                    If (linemouseClick) Then
                        'first time left click - create empty line feature
                        If firstClick Then
                            'Create a new List called lineArray.
                            'In List we need not define the size and also
                            'Here this list will store the Coordinates
                            'We are going to store the mouse click coordinates into this array.
                            Dim lineArray As New List(Of Coordinate)
                            'Create an instance for LineString class.
                            'We need to pass collection of list coordinates
                            Dim lineGeometry As New LineString(lineArray)
                            'Add the linegeometry to line feature
                            Dim lineFeature As IFeature = lineF.AddFeature(lineGeometry)
                            'add first coordinate to the line feature
                            lineFeature.Coordinates.Add(coord)
                            'set the line feature attribute
                            lineID = lineID + 1
                            lineFeature.DataRow("LineID") = lineID
                            firstClick = False
                        Else
                            'second or more clicks - add points to the existing feature
                            Dim existingFeature As IFeature = lineF.Features(lineF.Features.Count - 1)
                            existingFeature.Coordinates.Add(coord)
                            'refresh the map if line has 2 or more points
                            If existingFeature.Coordinates.Count >= 2 Then
                                lineF.InitializeVertices()
                                map1.ResetBuffer()
                            End If
                        End If
                    End If
                Else
                    'right click - reset first mouse click
                    firstClick = True
                    map1.ResetBuffer()
                End If

            Case MapShapeType.Polygon
                If e.Button = MouseButtons.Left Then
                    'left click - fill array of coordinates
                    Dim coord As Coordinate = map1.PixelToProj(e.Location)
                    If (polygonmouseClick) Then
                        'first time left click - create empty line feature
                        If firstClick Then
                            'Create a new List called polygonArray.
                            'Here this list will store the Coordinates
                            'We are going to store the mouse click coordinates into this array.
                            Dim polygonArray As New List(Of Coordinate)
                            'Create an instance for LinearRing class.
                            'We pass the polygon List to the constructor of this class
                            Dim polygonGeometry As New LinearRing(polygonArray)
                            'Add the polygonGeometry instance to PolygonFeature
                            Dim polygonFeature As IFeature = polygonF.AddFeature(polygonGeometry)
                            'add first coordinate to the polygon feature
                            polygonFeature.Coordinates.Add(coord)
                            'set the polygon feature attribute
                            polygonID = polygonID + 1
                            polygonFeature.DataRow("ID") = polygonID
                            firstClick = False
                        Else
                            'second or more clicks - add points to the existing feature
                            Dim existingFeature As IFeature = polygonF.Features(polygonF.Features.Count - 1)
                            existingFeature.Coordinates.Add(coord)
                            'refresh the map if line has 2 or more points
                            If existingFeature.Coordinates.Count >= 3 Then
                                'refresh the map
                                polygonF.InitializeVertices()
                                map1.ResetBuffer()
                            End If
                        End If
                    End If
                Else
                    'right click - reset first mouse click
                    firstClick = True
                    map1.Cursor = Cursors.Arrow
                End If
        End Select


    End Sub

    Private Sub SavePointToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SavePointToolStripMenuItem.Click
        pointF.SaveAs("point.shp", True)
        MsgBox("The point shapefile has been saved.")
        Map1.Cursor = Cursors.Arrow
    End Sub

    Private Sub SavePointToolStripMenuItem_MouseDown(sender As Object, e As MouseEventArgs) Handles SavePointToolStripMenuItem.MouseDown

    End Sub

    Private Sub CreatePolylineShapefileToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CreatePolylineShapefileToolStripMenuItem.Click
        'initialize polyline feature set
        map1.Cursor = Cursors.Cross
        'set shape type
        ShapeType = MapShapeType.Line
        'set projection
        lineF.Projection = map1.Projection
        'initialize the featureSet attribute table
        Dim column As New DataColumn("LineID")
        If Not (lineF.DataTable.Columns.Contains("LineID")) Then
            lineF.DataTable.Columns.Add(column)
        End If
        'add the featureSet as map layer
        lineLayer = map1.Layers.Add(lineF)
        Dim symbol As New LineSymbolizer(Color.Blue, 2)
        lineLayer.Symbolizer = symbol
        lineLayer.LegendText = "line"
        firstClick = True
        linemouseClick = True
    End Sub

    Private Sub SaveToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveToolStripMenuItem.Click
        lineF.SaveAs("c:\MW\line.shp", True)
        MsgBox("The line shapefile has been saved in C:\MW\")
        Map1.Cursor = Cursors.Arrow
        linemouseClick = False
    End Sub

    Private Sub CreatePolygonShapefileToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CreatePolygonShapefileToolStripMenuItem.Click
       'initialize polyline feature set
        map1.Cursor = Cursors.Cross
        'set shape type
        ShapeType = MapShapeType.Polygon
        'set projection
        polygonF.Projection = map1.Projection
        'initialize the featureSet attribute table
        Dim column As New DataColumn("ID")
        polygonF.DataTable.Columns.Add(column)
        'add the featureSet as map layer
        Dim polygonLayer As MapPolygonLayer = map1.Layers.Add(polygonF)
        'set a symbology
        Dim symbol As New PolygonSymbolizer(Color.Green)
        polygonLayer.Symbolizer = symbol
        'add the legend text
        polygonLayer.LegendText = "polygon"
        firstClick = True
        polygonmouseClick = True

    End Sub

    Private Sub SaveToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles SaveToolStripMenuItem1.Click
        polygonF.SaveAs("c:\MW\polygon.shp", True)
        MsgBox("The polygon shapefile has been saved in C:\MW\")
        Map1.Cursor = Cursors.Arrow
        polygonmouseClick = False
    End Sub

    Private Sub ViewAttributeTableToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ViewAttributeTableToolStripMenuItem.Click
        'Declare a datatable
        Dim dt As DataTable
        If Map1.Layers.Count > 0 Then
            Dim stateLayer As MapPolygonLayer
            stateLayer = CType(Map1.Layers(0), MapPolygonLayer)
            If stateLayer Is Nothing Then
                MessageBox.Show("The layer is not a polygon layer.")
            Else
                'Get the shapefile's attribute table to our datatable dt
                dt = stateLayer.DataSet.DataTable
                'Set the datagridview datasource from datatable dt
                dgvAttributeTable.DataSource = dt
            End If
        Else
            MessageBox.Show("Please add a layer to the map.")
        End If
    End Sub

    Private Sub DeleteAColumnInTheAttributeTableToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeleteAColumnInTheAttributeTableToolStripMenuItem.Click
        'Declare a datatable
        Dim dt As DataTable
        If Map1.Layers.Count > 0 Then
            Dim stateLayer As MapPolygonLayer
            stateLayer = CType(Map1.Layers(0), MapPolygonLayer)
            If stateLayer Is Nothing Then
                MessageBox.Show("The layer is not a polygon layer.")
            Else
                'Get the shapefile's attribute table to our datatable dt
                dt = stateLayer.DataSet.DataTable
                'Remove a column
                dt.Columns.Remove("PercentMale")
                'Set the datagridview datasource from datatable dt
                dgvAttributeTable.DataSource = dt
            End If
        Else
            MessageBox.Show("Please add a layer to the map.")
        End If
    End Sub

    Private Sub ExportAttributeTableToExcelToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExportAttributeTableToExcelToolStripMenuItem.Click
        'Declare a datatable
        Dim dt As DataTable
        If Map1.Layers.Count > 0 Then
            Dim stateLayer As MapPolygonLayer
            stateLayer = CType(Map1.Layers(0), MapPolygonLayer)
            If stateLayer Is Nothing Then
                MessageBox.Show("The layer is not a polygon layer.")
            Else
                'Get the shapefile's attribute table to our datatable dt
                dt = stateLayer.DataSet.DataTable
                'Call the sub ExportToExcel 
                'This sub procedure expects a datatable as an input
                ExportToExcel(dt)
            End If
        Else
            MessageBox.Show("Please add a layer to the map.")
        End If
    End Sub
    ''' <summary>
 ''' This sub method is used to create an excel worksheet from the attribute table
 ''' </summary>
 ''' <param name="objDT">attribute table as a datatable input</param>
 ''' <remarks> Click the COM tab of the Add Reference dialog box, and find Microsoft Excel 14 Object 
Library.</remarks>
Private Sub PrintMapToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PrintMapToolStripMenuItem.Click
        Dim frm As New DotSpatial.Controls.LayoutForm
 frm.MapControl = Map1
 frm.Show()
    End Sub

    Private Sub LoadToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LoadToolStripMenuItem.Click
        Map1.AddLayer()
    End Sub

    Private Sub ExportToExcel(ByVal objDT As DataTable)
        Dim Excel As Object = CreateObject("Excel.Application")
        Dim strFilename As String
        Dim intCol, intRow As Integer
        'path for storing excel datasheet
        Dim strPath As String = "C:\2009 Falls\"
If Excel Is Nothing Then
 MsgBox("It appears that Excel is not installed on this machine. This operation requires MS Excel 
to be installed on this machine.", MsgBoxStyle.Critical)
 Return
 End If
 Try
 With Excel
 .SheetsInNewWorkbook = 1
 .Workbooks.Add()
 .Worksheets(1).Select()
 .cells(1, 1).value = "Attribute table" 'Heading of the excel file
 .cells(1, 1).EntireRow.Font.Bold = True
 'Add the column names from the attribute table to excel worksheet
 Dim intI As Integer = 1
 For intCol = 0 To objDT.Columns.Count - 1
 .cells(2, intI).value = objDT.Columns(intCol).ColumnName
 .cells(2, intI).EntireRow.Font.Bold = True
 intI += 1
 Next
 'Add the data row values from the attribute table to ecxel worksheet
 intI = 3
 Dim intK As Integer = 1
 For intCol = 0 To objDT.Columns.Count - 1
 intI = 3
 For intRow = 0 To objDT.Rows.Count - 1
 .Cells(intI, intK).Value = objDT.Rows(intRow).ItemArray(intCol)
 intI += 1
 Next
 intK += 1
 Next
 If Mid$(strPath, strPath.Length, 1) <> "\" Then
 strPath = strPath & "\"
 End If
 strFilename = strPath & "AttributeTable.xls"
 .ActiveCell.Worksheet.SaveAs(strFilename)
 End With
 System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel)
 Excel = Nothing
 MsgBox("Data's are exported to Excel Succesfully in '" & strFilename & "'", 
MsgBoxStyle.Information)
 Catch ex As Exception
 MsgBox(ex.Message)
 End Try
 ' The excel is created and opened for insert value. We most close this excel using this system
 Dim pro() As Process = System.Diagnostics.Process.GetProcessesByName("EXCEL")
 For Each i As Process In pro
 i.Kill()
 Next
 End Sub

End Class

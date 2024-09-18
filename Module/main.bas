Attribute VB_Name = "main"
Option Explicit
Public aryPush As ArrayPush
Public cmn As Common
Public cdict As CreateDict
Public pcs As ProcessControl
Public timeStr As timeStr

Const grpTimeTableStr As String = "_TimeTable"
Const grpDetailStr As String = "_Detail"

Sub setup()
    Set aryPush = New ArrayPush
    Set cmn = New Common
    Set cdict = New CreateDict
    Set pcs = New ProcessControl
    Set timeStr = New timeStr
    
End Sub


'/**********************
' * 時間軸を作る
' **********************/
Public Sub makeTimeAxis(sheetName As String)
    Call setup
    Dim config As Dictionary
    Dim tbl As Variant
    Dim stg As Dictionary
    Dim startMinute As Integer
    Dim endMinute As Integer
    Dim squareTime As Integer
    Dim squareHeight As Integer
    Dim timeAxisArray() As String
    ReDim timeAxisArray(0)
    Dim timeAxisLineArray() As String
    ReDim timeAxisLineArray(0)
    Dim grp As Shape
    Dim grp_line As Shape
    Dim i As Integer
   
    Set config = cdict.createCellMappingDict("config")
    tbl = cdict.createCellComparisonList("tbl")
    Set stg = cdict.createNestedDict("stg", 2)
    
    startMinute = timeStr.timeStrToMinute(config("tbl_start"))
    endMinute = timeStr.timeStrToMinute(config("tbl_end"))
    squareTime = config("square_time")
    squareHeight = config("square_height")
    
    For i = 0 To (endMinute - startMinute) / squareTime
        Dim name As String
        name = "timeAxis" & Format(i, "00")
        With ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, _
                0, _
                i * squareHeight, _
                config("time_axis_width"), _
                squareHeight)
            .name = name
            '表示文字の指定
            .TextFrame.Characters.Text = timeStr.MinuteToTimeStr(startMinute + i * squareTime)
            '図形内テキストのフォントカラーを指定する
            .TextFrame.Characters.Font.ColorIndex = xlAutomatic
            .TextFrame.Characters.Font.Size = config("time_axis_font")
            .TextFrame.Characters.Font.Bold = msoTrue
            '図形内のテキストを中央位置かつ上揃えにする
            .TextFrame.HorizontalAlignment = xlHAlignCenter
            .TextFrame.VerticalAlignment = xlVAlignTop
            '図形の枠線を表示しない
            .Line.Visible = msoFalse
            '図形の塗りつぶし色なし
            .Fill.Visible = msoFalse
        End With
        
        Call aryPush.ArrayPush(timeAxisArray, name)
        
    Next i
    
    For i = 0 To ((endMinute - startMinute) / squareTime)
        name = "timeAxisLine" & Format(i, "00")
        With ActiveSheet.Shapes.AddShape(msoConnectorStraight, _
                100, _
                i * squareHeight, _
                100, _
                0)
            .name = name
            '図形の枠線の色を指定する
            .Line.ForeColor.RGB = RGB(128, 128, 128)
            .Line.Weight = 1
            .Line.Transparency = 0.3
        End With
        
        Call aryPush.ArrayPush(timeAxisLineArray, name)
        
    Next i
    
    Set grp = Worksheets(sheetName).Shapes.Range(timeAxisArray).group
    grp.name = "time_axis"
    
    Set grp_line = Worksheets(sheetName).Shapes.Range(timeAxisLineArray).group
    grp_line.name = "time_axis_line"
End Sub


'/**********************
' * 列をつくる
' **********************/

Public Sub makeTimeTable(stage_id As String, sheetName As String)
    Call setup
    Dim config As Dictionary
    Dim tblArray As Variant
    Dim stg As Dictionary
    Dim tblDict As Dictionary
    Dim groupArray() As String
    ReDim groupArray(0)
    Dim name As String
    
    Set config = cdict.createCellMappingDict("config")
    tblArray = cdict.createCellComparisonList("tbl")
    Set stg = cdict.createNestedDict("stg", 1)
    Set tblDict = cdict.createCellListDict("tbl", 2)
    
    Dim startTimeAxisMinute As Integer
    Dim endTimeAxisMinute As Integer
    startTimeAxisMinute = timeStr.timeStrToMinute(config("tbl_start"))
    endTimeAxisMinute = timeStr.timeStrToMinute(config("tbl_end"))
    
    name = stage_id & "_base"
    With ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, _
            0, _
            config("textbox_size_plus"), _
            config("group_width"), _
            ((endTimeAxisMinute - startTimeAxisMinute) / config("square_time") + config("tbl_margin_bottom_ratio")) * config("square_height") + config("textbox_size_plus") _
            )
        .name = name
        '図形内テキストのフォントカラーを指定する
        .TextFrame.Characters.Font.ColorIndex = xlAutomatic
        .TextFrame.Characters.Font.Size = config("stage_font")
        '図形内のテキストを中央位置かつ上揃えにする
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignTop
        '図形の枠線の色を指定する
        .Line.ForeColor.RGB = RGB(0, 0, 0)
        .Line.Weight = 1
        .Line.Transparency = 1
        '図形の塗りつぶし色を指定する
        .Fill.ForeColor.RGB = RGB(0, 0, 0)
        .Fill.Transparency = 1
        .Adjustments.Item(1) = 0.001
    End With
    Call aryPush.ArrayPush(groupArray, name)
    
    Dim group As Variant
    For Each group In tblDict(stage_id)
        Dim startGroupMinute As Integer
        Dim endGroupMinute As Integer
        Dim timeStartEnd As String
        startGroupMinute = timeStr.timeStrToMinute(group("start_time"))
        endGroupMinute = timeStr.timeStrToMinute(group("end_time"))
        timeStartEnd = group("start_time") & " ～ " & group("end_time")
        '下地
        name = stage_id & "_" & group("group_name") & "_" & "under"
        With ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, _
                config("group_width") * (1 - config("group_width_per")) * 0.5, _
                (startGroupMinute - startTimeAxisMinute) * config("square_height") / config("square_time") + config("group_margin") + config("textbox_size_plus"), _
                config("group_width") * config("group_width_per"), _
                (endGroupMinute - startGroupMinute) * config("square_height") / config("square_time") - config("group_margin") * 2 _
                )
            .name = name
            '表示文字の指定
            .TextFrame.Characters.Text = ""
            '図形の枠線の色を指定する
            .Line.ForeColor.RGB = RGB(0, 0, 0)
            .Line.Weight = 1
            .Line.Transparency = 1
            '図形の塗りつぶし色を指定する
            .Fill.ForeColor.RGB = RGB(255, 255, 255)
            .Fill.Transparency = 0
            .Adjustments.Item(1) = 0.07
            '図形の余白をなくす
            .TextFrame2.MarginTop = 0
            .TextFrame2.MarginBottom = 0
        End With
        Call aryPush.ArrayPush(groupArray, name)
        
        '透明テキストボックス
        name = stage_id & "_" & group("group_name") & "_" & "text"
        With ActiveSheet.Shapes.AddShape(msoTextOrientationHorizontal, _
                config("group_width") * (1 - config("group_width_per")) * 0.5, _
                (startGroupMinute - startTimeAxisMinute) * config("square_height") / config("square_time"), _
                config("group_width") * config("group_width_per"), _
                (endGroupMinute - startGroupMinute) * config("square_height") / config("square_time") + config("textbox_size_plus") * 2 _
                )
            .name = name
            '表示文字の指定
            .TextFrame.Characters.Text = group("group_name") & vbLf & timeStartEnd
            .TextFrame.Characters.Font.Size = config("group_font")
            ' group("group_name") を黒字に設定
            With .TextFrame.Characters(1, Len(group("group_name"))).Font
                .Size = config("group_font")
                .Bold = msoTrue
                .Color = RGB(0, 0, 0)
            End With
            ' "10:00 ～ 10:30" を水色に設定
            With .TextFrame.Characters(Len(group("group_name")) + 2, Len(timeStartEnd)).Font
                .Size = config("group_time_font")
                .Bold = msoTrue
                .Color = RGB(0, 200, 200)
            End With
            '図形内のテキストを中央位置にする
            .TextFrame.HorizontalAlignment = xlHAlignCenter
            .TextFrame.VerticalAlignment = xlVAlignCenter
            '図形の枠線の色を指定する
            .Line.ForeColor.RGB = RGB(0, 0, 0)
            .Line.Weight = 1
            .Line.Transparency = 1
            '図形の塗りつぶし色を指定する
            .Fill.ForeColor.RGB = RGB(255, 255, 255)
            .Fill.Transparency = 1
            '図形の余白をなくす
            .TextFrame2.MarginTop = 0
            .TextFrame2.MarginBottom = 0
        End With
        Call aryPush.ArrayPush(groupArray, name)
        
        '上(クリックイベントも埋め込む)
        name = stage_id & "_" & group("group_name") & "_" & "up"
        With ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, _
                config("group_width") * (1 - config("group_width_per")) * 0.5, _
                (startGroupMinute - startTimeAxisMinute) * config("square_height") / config("square_time") + config("group_margin") + config("textbox_size_plus"), _
                config("group_width") * config("group_width_per"), _
                (endGroupMinute - startGroupMinute) * config("square_height") / config("square_time") - config("group_margin") * 2 _
                )
            .name = name
            '表示文字の指定
            .TextFrame.Characters.Text = ""
            ' group("group_name") を黒字に設定
            '図形の枠線の色を指定する
            .Line.ForeColor.RGB = RGB(0, 0, 0)
            .Line.Weight = 1
            .Line.Transparency = 0
            '図形の塗りつぶし色を指定する
            .Fill.ForeColor.RGB = RGB(255, 255, 255)
            .Fill.Transparency = 1
            .Adjustments.Item(1) = 0.07
            '図形の余白をなくす
            .TextFrame2.MarginTop = 0
            .TextFrame2.MarginBottom = 0
        End With
        Call aryPush.ArrayPush(groupArray, name)
        
        Worksheets(sheetName).Shapes(name).OnAction = "Shape_DoubleClick"
    Next group
    
    Dim grp As Shape
    Set grp = Worksheets(sheetName).Shapes.Range(groupArray).group
    grp.name = stage_id & grpTimeTableStr
    
    
    
End Sub

Public Sub makeStageDetail(stage_id As String, sheetName As String)
    Call setup
    Dim config As Dictionary
    Dim tblArray As Variant
    Dim stg As Dictionary
    Dim tblDict As Dictionary
    Dim groupArray() As String
    ReDim groupArray(0)
    Dim name As String

    Set config = cdict.createCellMappingDict("config")
    tblArray = cdict.createCellComparisonList("tbl")
    Set stg = cdict.createNestedDict("stg", 1)
    Set tblDict = cdict.createCellListDict("tbl", 2)
    
    Dim i As Integer
    i = 0
    Dim value As Variant
    For Each value In Array("stage_name", "stage_place")
        name = stage_id & "_" & value
        With ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, _
                config("group_width") * (1 - config("group_width_per")) * 0.5, _
                config("stage_detail_top") + i * (config("stage_interval") * 1 + config("stage_height") * 1), _
                config("group_width") * config("group_width_per"), _
                config("stage_height"))
            .name = name
            '表示文字の指定
            .TextFrame.Characters.Text = stg(stage_id)(value)
            '図形内テキストのフォントカラーを指定する
            .TextFrame.Characters.Font.ColorIndex = xlAutomatic
            .TextFrame.Characters.Font.Size = config("stage_font")
            .TextFrame.Characters.Font.Bold = msoTrue
            '図形内のテキストを中央位置かつ上揃えにする
            .TextFrame.HorizontalAlignment = xlHAlignCenter
            .TextFrame.VerticalAlignment = xlVAlignTop
            '図形の枠線の色を指定する
            .Line.ForeColor.RGB = RGB(0, 0, 0)
            .Line.Weight = 1
            .Line.Transparency = 0
            '図形の塗りつぶし色を指定する
            .Fill.ForeColor.RGB = RGB(255, 255, 255)
            .Fill.Transparency = 0
            .Adjustments.Item(1) = 0.07
        End With
        i = i + 1
        Call aryPush.ArrayPush(groupArray, name)
        
    Next
    
    
    
    name = stage_id & "_detail_base"
    With ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, _
            0, _
            0, _
            config("group_width"), _
            config("stage_detail_top") * 2 + config("stage_interval") * 1 + config("stage_height") * 2)
        .name = name
        '図形内テキストのフォントカラーを指定する
        .TextFrame.Characters.Font.ColorIndex = xlAutomatic
        .TextFrame.Characters.Font.Size = config("time_axis_font")
        '図形内のテキストを中央位置かつ上揃えにする
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignTop
        '図形の枠線の色を指定する
        .Line.ForeColor.RGB = RGB(0, 0, 0)
        .Line.Weight = 1
        .Line.Transparency = 1
        '図形の塗りつぶし色を指定する
        .Fill.ForeColor.RGB = RGB(0, 0, 0)
        .Fill.Transparency = 1
        .Adjustments.Item(1) = 0.001
    End With
    Call aryPush.ArrayPush(groupArray, name)
    
    
    Dim grp As Shape
    Set grp = Worksheets(sheetName).Shapes.Range(groupArray).group
    grp.name = stage_id & grpDetailStr
    
    
    
End Sub


Public Sub start()
    Call setup
    Call pcs.StartProcess
    Dim sheetName As String
    sheetName = "timetable"
    Call cmn.delWorkSheet(sheetName)
    Call cmn.makeNewWorkSheet(sheetName)
    Worksheets(sheetName).Select
    Call makeTimeAxis(sheetName)
    Dim stg As Dictionary
    Set stg = cdict.createNestedDict("stg", 1)
    
    Dim stage_id As Variant
    Dim i As Integer
    For Each stage_id In stg
        Call test(CStr(stage_id), i, sheetName)
        i = i + 1
        Worksheets(sheetName).Shapes(CStr(stage_id) & grpTimeTableStr).Ungroup
    Next stage_id
    
    Call pcs.EndProcess
End Sub




Public Sub test(stage_id As String, i As Integer, sheetName As String)
    Call setup
    Dim config As Dictionary
    Set config = cdict.createCellMappingDict("config")
    Dim stg As Dictionary
    Set stg = cdict.createNestedDict("stg", 1)
    Dim groupArray() As String
    ReDim groupArray(0)
    
    Call makeStageDetail(stage_id, sheetName)
    Call makeTimeTable(stage_id, sheetName)
    
    Dim ta As Shape
    Dim tt As Shape
    Dim d As Shape
    Dim tal As Shape
    
    Set ta = Worksheets(sheetName).Shapes("time_axis")
    Set tt = Worksheets(sheetName).Shapes(stage_id & grpTimeTableStr)
    Set d = Worksheets(sheetName).Shapes(stage_id & grpDetailStr)
    Set tal = Worksheets(sheetName).Shapes("time_axis_line")
    
    ta.Top = d.Top + d.Height + 5
    tt.Top = ta.Top + config("time_axis_font") - config("textbox_size_plus")
    tal.Top = ta.Top + config("time_axis_font")
    
    tt.Left = ta.Width + i * (tt.Width + config("stage_table_margin"))
    d.Left = ta.Width + i * (tt.Width + config("stage_table_margin"))
    tal.Left = ta.Width
    tal.Width = tt.Left - ta.Width + tt.Width
    
    Dim name As String
    name = stage_id & "_all_base"
    With ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, _
            d.Left, _
            0, _
            config("group_width"), _
            tt.Top + tt.Height _
            )
        .name = name
        '図形内テキストのフォントカラーを指定する
        .TextFrame.Characters.Font.ColorIndex = xlAutomatic
        .TextFrame.Characters.Font.Size = config("time_axis_font")
        '図形内のテキストを中央位置かつ上揃えにする
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignTop
        '図形の枠線の色を指定する
        .Line.ForeColor.RGB = RGB(stg(stage_id)("R"), stg(stage_id)("G"), stg(stage_id)("B"))
        .Line.Weight = 1
        .Line.Transparency = 0.7
        '図形の塗りつぶし色を指定する
        .Fill.ForeColor.RGB = RGB(stg(stage_id)("R"), stg(stage_id)("G"), stg(stage_id)("B"))
        .Fill.Transparency = 0.9
        .Adjustments.Item(1) = 0.001
    End With
    
    Worksheets(sheetName).Shapes(name).ZOrder msoSendToBack
    
    
    
End Sub

Sub アクティブシートの全Shapeを選択する_For_Each_Next()
    Call setup
    Dim shp As Shape
    Dim testshp As Shape
    Dim sheetName As String
    sheetName = "timetable"
    For Each shp In Worksheets(sheetName).Shapes
        
    Next
  
    Dim groupArray() As Integer
    ReDim groupArray(0)
    Call aryPush.ArrayPush(groupArray, 1)
    Call aryPush.ArrayPush(groupArray, 2)
    Call aryPush.ArrayPush(groupArray, 3)
  
  
    Dim grp As Shape
    
End Sub


Public Sub Shape_DoubleClick()
    Call setup
    Dim config As Dictionary
    Set config = cdict.createCellMappingDict("config")
    Dim shpName As String
    Dim shp As Shape
    ' クリックされた図形の名前を取得
    shpName = Application.Caller
    ' アクティブなシートから図形を取得
    Set shp = ActiveSheet.Shapes(shpName)
    
    Dim name As String
    name = shpName & "_selected"
    With ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, _
            shp.Left - config("selected_margin_lr"), _
            shp.Top - config("selected_margin_tb"), _
            shp.Width + 2 * config("selected_margin_lr"), _
            shp.Height + 2 * config("selected_margin_tb") _
            )
        .name = name
        '図形内テキストのフォントカラーを指定する
        .TextFrame.Characters.Font.ColorIndex = xlAutomatic
        .TextFrame.Characters.Font.Size = config("time_axis_font")
        '図形内のテキストを中央位置かつ上揃えにする
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignTop
        '図形の枠線の色を指定する
        .Line.ForeColor.RGB = RGB(0, 200, 200)
        .Line.Weight = 3
        .Line.Transparency = 0.7
        '図形の塗りつぶし色を指定する
        .Fill.ForeColor.RGB = RGB(0, 200, 200)
        .Fill.Transparency = 0.7
        .Adjustments.Item(1) = 0.1
    End With
    ActiveSheet.Shapes(name).OnAction = "Shape_Click"
    
End Sub


Public Sub Shape_Click()
    Call setup
    Dim config As Dictionary
    Set config = cdict.createCellMappingDict("config")
    Dim shpName As String
    Dim shp As Shape
    ' クリックされた図形の名前を取得
    shpName = Application.Caller
    ' アクティブなシートから図形を取得
    Set shp = ActiveSheet.Shapes(shpName)
    
    If shp.Fill.Transparency = 1 Then
        shp.Delete
        Exit Sub
    End If
    
    shp.Line.ForeColor.RGB = RGB(0, 0, 0)
    shp.Line.Transparency = 0
    shp.Fill.Transparency = 1

    
End Sub

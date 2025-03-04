
;NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

;★一般★

#^+t:: ; Ctrl+Shift+Tで実行


inputbox var
Switch var
{
    case "720":
        filePath := "\\000lib01.canon.net\aaa109286\B77-3331-000.csv" ;    ; ファイルが存在するか確認
    case "710":

    Default:
}

    ; ファイルが存在するか確認
    if !FileExist(filePath)
    {
        MsgBox, 48, Error, File not found: %filePath%
        return
    }

    ; ファイルの内容を読み取る
    FileRead, fileContent, %filePath%
    if (StrLen(fileContent) > 10000) ; 10,000文字を超える場合にトリミング
{
    fileContent := SubStr(fileContent, 1, 10000) . "`n... (truncated)"
}


    ; GUIを作成
    Gui, +AlwaysOnTop +Resize +ToolWindow
    Gui, Margin, 10, 10
    Gui, Add, Edit, x0 y0 w500 h1000 ReadOnly +VScroll, %fileContent% ; スクロール可能なテキストボックス

    ; マウス位置に表示
    MouseGetPos, mouseX, mouseY
    Gui, Show, x%mouseX% y%mouseY% w520 h420, Scrollable Text
return




;★キー代用系
ScrollLock::Return
Insert::Return
vk1D & H::Send,{Blind}{Left}
vk1D & S::Send,{Blind}{Left}
vk1D & J::Send,{Blind}{Down}
vk1D & F::Send,{Blind}{Down}
vk1D & K::Send,{Blind}{Up}
vk1D & D::Send,{Blind}{Up}
vk1D & L::Send,{Blind}{Right}
vk1D & G::Send,{Blind}{Right}
vk1D & Y::Send,{Blind}{HOME}
vk1D & O::Send,{Blind}{END}
vk1D & U::Send,{Blind}{PgUp}
vk1D & I::Send,{Blind}{PgDn}

vk1D & R::      ;ウィンドウ最大化
    WinGet, tmp, MinMax, A
    If tmp = 1
        WinRestore, A
    Else
        WinMaximize, A
    Return

vk1D & E::Send, {RWin Down}{Shift Down}{left}{Shift Up}{RWin Up}    ;ウィンウを左右に
vk1D & W::PostMessage, 274, 61472, 0, , A       ;ウィンドウ最小化


;vk1D & C::Send,{Blind}{PrintScreen}
vk1D & Z::Send,{Blind}{Delete}
vk1D & B::Send,{Blind}{Backspace}
vk1D & 2::Send,{F2}
vk1D & Q::!F4
!q::!F4
vk1D & space::-
vk1D & vk1C::send, {esc}
vk1D & A::Send,{Blind}{sc01C}   ;enter

vk1D::
    KeyWait, vk1D, D T0.3
    sleep, 50

;★
    if (ErrorLevel = 0)
    {
        Send,{sc01C}
    }
    else
    {}
    Return

;★入力省略系
::d/::
    FormatTime, now,, yyyy/MM/dd
    Clipboard = %now%
    Send,^v
    return

::d;::
    FormatTime, now,, yyyyMMdd
    Clipboard = %now%
    Send,^v
    return

::m/::
    FormatTime, now,, M/d
    Clipboard = %now%
    Send,^v
    return

::t;::
    FormatTime, now,, HH:mm:ss
    SendInput %now%
    return

::ij;::
    buban = IJ21157-0
    Clipboard = %buban%
    Send,^v
    return

::ba;::
    buban = BA5-
    Clipboard = %buban%
    Send,^v
    return

::f;::
    from = from:
    Clipboard = %from%
    Send,^v
    return

vk1D & 8::
        backup := ClipboardAll
        ClipWait, 2
        Clipboard := "()"
        ClipWait, 2
        Send, ^v
        Sleep,60
        Clipboard := backup
        Send, {Left}
Return


vk1C & G::  ;googleで開く
    InputBox, sword, Google Search, , ,500,130
    if ErrorLevel <> 0
    {}
    else
    {
        NewStr := RegExReplace(sword, "\s+", "+")
        Sleep 3
                ; Run, "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" -new-tab "http://www.google.co.jp/search?hl=ja&lr=lang_ja&ie=UTF-8&q=%sword%"
        Run, "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" -new-tab "http://www.google.co.jp/search?hl=ja&lr=lang_ja&ie=UTF-8&q=%sword%"
    }
    return

; ホットキーを指定してbatファイルを実行
#+A::
    {
        Path := "C:\Users\aaa109286\Documents"
        ; 実行するbatファイルの名前（拡張子込み）

        batFileName := "ClipboardImage2File.bat"
        ; 完全なパスを組み立て
        batFilePath := Path . "\" . batFileName

        if FileExist(batFilePath)
        {
        Run, %batFilePath%, %Path%, Hide
        ToolTip now making
        SetTimer, RemoveToolTip, 1000
        }
        else
        {
        MsgBox, Error: %batFileName% does not exist
        }
    return
    }

vk1D & v::  ;リンクを開く
    if(GetKeyState("Shift","P"))
    {
        clipboard =
        Send, ^c
        ClipWait

        stringreplace clipboard,clipboard,file:///,, All
        stringreplace clipboard,clipboard,<,, All
        stringreplace clipboard,clipboard,>,, All
        stringreplace clipboard,clipboard,`r`n,, All
        stringreplace clipboard,clipboard,`r,, All
        stringreplace clipboard,clipboard,",, All
        stringreplace clipboard,clipboard, %A_Space%,, All
        stringreplace clipboard,clipboard, %A_Tab%,, All
        text := clipboard
        ;最後の文字列\の位置を返す
        temp := % InStr(clipboard,"\",false,0,1)
        ;  Return strPath
        ;上記位置より前の文字列をクリップボードに保存
        stringleft path, text, %temp%
        Run explorer.exe %path%
    }
    else if(GetKeyState("Ctrl","P"))
    {
        Send,{F4}
    }
    else
    {
        clipboard =
        Send, ^c
        ClipWait
        stringreplace clipboard,clipboard,file:///,, All
        stringreplace clipboard,clipboard,<,, All
        stringreplace clipboard,clipboard,>,, All
        stringreplace clipboard,clipboard, `r`n,, All
        stringreplace clipboard,clipboard, `r ,, All
        stringreplace clipboard,clipboard,",, All
        stringreplace clipboard,clipboard, %A_Space%,, All
        stringreplace clipboard,clipboard, %A_Tab%,, All
        ;msgbox %clipboard%
        Run explorer.exe %clipboard%
    }
    return

;【Excel以外】Shift+マウスホイールで水平スクロール
+WheelUp::
    MouseGetPos,,,id, fcontrol,1
    Loop 2 ;数字を大きくするとそれだけ大きくスクロール
    SendMessage, 0x114, 0, 0, %fcontrol%, ahk_id %id%
    Return

+WheelDown::
    MouseGetPos,,,id, fcontrol,1
    Loop 2 ;数字を大きくするとそれだけ大きくスクロール
    SendMessage, 0x114, 1, 0, %fcontrol%, ahk_id %id%
    Return

;【Excel】Shift+マウスホイールで水平スクロール
#IfWinActive ahk_exe EXCEL.EXE
+WheelUp::
    SetScrollLockState, On
    SendInput {Left 1}  ;数字を大きくすると大きくスクロールする
    SetScrollLockState, Off
    Return

+WheelDown::
    SetScrollLockState, On
    SendInput {Right 1} ;数字を大きくすると大きくスクロールする
    SetScrollLockState, Off
    Return

#IfWinActive


;---------------------------------------------------;

~XButton1 & WheelDown::Send {Alt down}{Tab}
~RButton & WheelDown::Send {Ctrl down}{Tab}{Ctrl up}


;引用貼り付け
;Win+Ctrl+v
^#v::
    StrRefMsg := Clipboard
    Clipboard = %StrRefMsg%
    Send, ^v
return

;---------------------------------------------------;
IMEGetstateOFF(){
DetectHiddenWindows, ON
WinGet, vcurrentwindow, ID, A
vgetdefault := DllCall("imm32.dll\ImmGetDefaultIMEWnd", "Uint", vcurrentwindow)
vimestate := DllCall("user32.dll\SendMessageA", "UInt", vgetdefault, "UInt", 0x0283, "Int", 0x0005, "Int", 0)
DetectHiddenWindows, Off
If (vimestate=0) ;imeがoffなら
{
;Offだから何もしない
}
else
{
Send, {vkf3}
}
return
}

IMEGetstateON(){
DetectHiddenWindows, ON
WinGet, vcurrentwindow, ID, A
vgetdefault := DllCall("imm32.dll\ImmGetDefaultIMEWnd", "Uint", vcurrentwindow)
vimestate := DllCall("user32.dll\SendMessageA", "UInt", vgetdefault, "UInt", 0x0283, "Int", 0x0005, "Int", 0)
DetectHiddenWindows, Off
If (vimestate=0) ;imeがoffなら
        {
        Send, {vkf3}
        }
else
        {
        ;Onだから何もしない
        }
return
}


; コピーしたらツールチップを表示
OnClipboardChange:
if (A_EventInfo = 1) {
;テキストの場合は内容を表示
;長いテキストの場合は長いテキストと表示
    StringLen, length, Clipboard
    if (length>200) {
        ToolTip 長いテキストをコピーしました。
        SetTimer, RemoveToolTip, 1000
    }
    else {
;短いテキストの場合は内容を表示
        ToolTip %length% %Clipboard%
        SetTimer, RemoveToolTip, 1000
    }
    }
    else if (A_EventInfo = 2) {
;テキスト以外の場合は適当に表示
        ToolTip テキストでないものをコピーしました。
        SetTimer, RemoveToolTip, 1000
}
return

RemoveToolTip:
SetTimer, RemoveToolTip, Off
ToolTip
return

;★teams★
TeamsMicToggle.ahk
vk1D & M::
    ToggleMuteTeams()
    Return
ToggleMuteTeams()
{
    SetTitleMatchMode,Regex
    IfWinNotExist, Microsoft Teams ahk_exe Teams.exe
    return

    WinGet, active_id, ID, A
    WinActivate, Microsoft Teams$ ahk_exe Teams.exe
    Send, ^+m
    Sleep, 100
    WinActivate, ahk_id %active_id%
}

#SingleInstance force

^+w:: ; Ctrl + Shift + W でファイル名の列幅を自動調整
; エクスプローラーのウィンドウを取得
hwnd := WinActive("ahk_class CabinetWClass")
if !hwnd {
MsgBox, 48, Explorer, Explorer window is not active!
return
}

; シェルウィンドウのリストビュー（SysListView32）を取得
for window in ComObjCreate("Shell.Application").Windows {
if (window.hwnd == hwnd) {
ControlGet, listHwnd, Hwnd ,, SysListView32, ahk_id %hwnd%
if listHwnd {
; LVM_SETCOLUMNWIDTH メッセージを送信（0 はファイル名の列）
SendMessage, 4126, 0, -1,, ahk_id %listHwnd%
MsgBox, 64, Explorer, Column width adjusted!
} else {
MsgBox, 16, Error, Could not find list view control!
}
return
}
}
return





;★powerpoint★

;#IfWinActive,ahk_exe POWERPNT.EXE
#IfWinActive, ahk_class PPTFrameClass

^+!G::send,!{n}{p}{1}{d}        ;画像の挿入
+^F::send,!{h}{g}{f}    ;前面
+^!F::send,!{h}{g}{b}   ;背面
+^C::send,!{h}{g}{a}{c}      ;水平方向中央揃え
+^!C::send,!{h}{g}{a}{m}      ;垂直方向中央揃え
+^T::send,!{j}{p}{v}{a}{4}      ;垂直方向中央揃え

;WinActivate, ahk_class screenClass ahk_exe POWERPNT.EXE
vk1D & WheelDown::
    Tooltip
    ppt := ComObjActive("PowerPoint.Application")
    SetFormat, FloatFast, 0.2
    Current_Thickness := ppt.ActiveWindow.Selection.ShapeRange.Line.Weight
    ppt.ActiveWindow.Selection.ShapeRange.Line.Weight:= Current_Thickness - 0.25
    Tooltip, % "Width: " Current_Thickness - 0.25
    SetTimer, RemoveToolTip, 1000
Return


vk1D & WheelUp::
;WinActivate, ahk_class screenClass ahk_exe POWERPNT.EXE
    Tooltip
    ppt := ComObjActive("PowerPoint.Application")
    SetFormat, FloatFast, 0.2
    Current_Thickness := ppt.ActiveWindow.Selection.ShapeRange.Line.Weight
    ppt.ActiveWindow.Selection.ShapeRange.Line.Weight:= Current_Thickness + 0.25
    Tooltip, % "Width: " Current_Thickness + 0.25
    SetTimer, RemoveToolTip, 1000
Return

/*
+#I::
    ppt := ComObjActive("PowerPoint.Application")  ; PowerPointのインスタンスを取得
    slide := ppt.ActiveWindow.View.Slide  ; アクティブなスライドを取得

    ; テキストボックスを追加 (位置とサイズを指定)
    textbox := slide.Shapes.AddTextbox(1, 50, 500, 200, 75) ; (テキストの種類, x, y, 幅, 高さ)
    textbox.TextFrame.TextRange.Text := "ここにテキストを入力" ; テキストを設定
    textbox.TextFrame.AutoSize := 1

    ; テキストボックスを選択
    textbox.Select
    ppt.ActiveWindow.Selection.ShapeRange.TextFrame.TextRange.Select
Return
*/


+^Z::   ;ラインを追加
    ppt := ComObjActive("PowerPoint.Application")
    slide := ppt.ActiveWindow.View.Slide
    line := slide.Shapes.AddLine(100, 100, 100, 300) ; (x1, y1, x2, y2)
    line.Line.ForeColor.RGB := 0x0000FF ; 線の色を赤に設定
    line.Line.EndArrowheadStyle := 3
    line.Select
    Return

+^B::   ;ラインを黒
    ppt := ComObjActive("PowerPoint.Application")
    selection := ppt.ActiveWindow.Selection
    ppt.ActiveWindow.selection.ShapeRange.Line.ForeColor.RGB := 0x000000
    Return

+^W::   ;ラインを白
    ppt := ComObjActive("PowerPoint.Application")
    selection := ppt.ActiveWindow.Selection
    ppt.ActiveWindow.selection.ShapeRange.Line.ForeColor.RGB := 0xFFFFFF
    Return


+^R::   ;ラインを赤
    ppt := ComObjActive("PowerPoint.Application")
    selection := ppt.ActiveWindow.Selection
    ppt.ActiveWindow.selection.ShapeRange.Line.ForeColor.RGB := 0x0000FF
    Return

+^!d::  ;ラインを点線
    ppt := ComObjActive("PowerPoint.Application")
    if(pptApp.activewindow.selection.type = 2)
    {
        SetFormat, FloatFast, 0.2
        Current_DashStyle := ppt.ActiveWindow.Selection.ShapeRange.Line.DashStyle
        if (Current_DashStyle = 1){
            ppt.ActiveWindow.Selection.ShapeRange.Line.DashStyle := 4
        }
        else if (Current_DashStyle = 4){
            ppt.ActiveWindow.Selection.ShapeRange.Line.DashStyle := 1
        }
    }
    Return

+^p::   ;図形の塗りつぶし
    Tooltip
    ppt := ComObjActive("PowerPoint.Application")
    SetFormat, FloatFast, 0.2
    shape := ppt.ActiveWindow.Selection.ShapeRange(1)

    if (shape.Fill.Transparency = 1){
        shape.Fill.Transparency := 0
        shape.Fill.ForeColor.RGB := 0xFFFFFF
    }
    else {
        shape.Fill.Transparency := 1
        }
    Return

+^V::
    Tooltip
    ppt := ComObjActive("PowerPoint.Application")
    SetFormat, FloatFast, 0.2
    shape := ppt.ActiveWindow.Selection.ShapeRange(1)
    if (shape.Line.Visible){
        shape.Line.Visible := False
    }
    else {
        shape.Line.Visible := True
    }
    Return

+^A::   ;ラインの先端を矢印に
    ppt := ComObjActive("PowerPoint.Application")
    SetSelectedLineArrowhead()

    SetSelectedLineArrowhead(){
    global ppt
    selectedShape :=ppt.ActiveWindow.Selection.ShapeRange
    if (selectedShape.Count > 0 && selectedShape.Type = 9) {
            currentArrowStyle := selectedShape.Line.EndArrowheadStyle
            if(currentArrowStyle = 1){
                    selectedShape.Line.EndArrowheadStyle := 3
            }
            else {
                    selectedShape.Line.EndArrowheadStyle := 1
            }
    }
    else {
            Msgbox, No line is selected or the selected shape is not a line!
    }
    }
    Return

+^X:: ; 選択された図形を90度回転
        ppt := ComObjActive("PowerPoint.Application")  ; PowerPointのインスタンスを取得
        selection := ppt.ActiveWindow.Selection  ; 現在の選択を取得

        ; 選択されているものがあるか確認
            selection.ShapeRange.Rotation += 90 ; 90度回転
    Return

+^!X:: ; 選択された図形を90度回転
        ppt := ComObjActive("PowerPoint.Application")  ; PowerPointのインスタンスを取得
        selection := ppt.ActiveWindow.Selection  ; 現在の選択を取得

        ; 選択されているものがあるか確認
            selection.ShapeRange.Rotation += 5 ; 5度回転
    Return



+^I::   ; PowerPointを起動し、選択中の図形にアクセス
        pptApp := ComObjActive("PowerPoint.Application")
        shape := pptApp.ActiveWindow.Selection.ShapeRange(1)

        ; 図形の位置とサイズを取得
        shapeLeft := shape.Left        ; 図形の左端の位置
        shapeTop := shape.Top          ; 図形の上端の位置
        shapeWidth := shape.Width      ; 図形の幅
        shapeHeight := shape.Height    ; 図形の高さ

        ; テキストボックスの幅と高さを設定
        textBoxWidth := 200            ; テキストボックスの幅
        textBoxHeight := 50            ; テキストボックスの高さ

        ; テキストボックスを図形の下側中央に配置
        textBoxLeft := shapeLeft + (shapeWidth / 2) - (textBoxWidth / 2)  ; 図形の中央に水平配置
        textBoxTop := shapeTop + shapeHeight + 5  ; 図形の下端から少し下に配置

        ; テキストボックスを追加
        slide := pptApp.ActiveWindow.View.Slide
        textBox := slide.Shapes.AddTextbox(1, textBoxLeft, textBoxTop, textBoxWidth, textBoxHeight)
        textBox.TextFrame.TextRange.Text := "edit this text"

        ; フォントやサイズを変更する
        textBox.TextFrame.TextRange.Font.Name := "Arial"
        textBox.TextFrame.TextRange.Font.Size := 12

        ; テキストボックスのサイズをテキストに合わせる
        textBox.TextFrame.AutoSize := 1 ; msoAutoSizeShapeToFitText = 1

        textBox.select()
        pptApp.ActiveWindow.Selection.TextRange.select()
        return

+^!R::
    ; PowerPointを起動し、スライドにアクセス
    pptApp := ComObjActive("PowerPoint.Application")
    slide := pptApp.ActiveWindow.View.Slide

    ; 長方形を指定の位置に追加 (左、上、幅、高さ)
    rect := slide.Shapes.AddShape(1, 100, 100, 300, 150)  ; msoShapeRectangle = 1

    ; 塗りつぶしなし
    rect.Fill.Visible := False

    ; 枠線の色と太さを設定
    rect.Line.Visible := True
    rect.Line.ForeColor.RGB := 0x0000FF  ; 赤色 (RGB: 255, 0, 0)
    rect.Line.Weight := 3  ; 枠線の太さ3ポイント
    return


+^Y::
    ; PowerPointを起動して現在のプレゼンテーションにアクセス
    pptApp := ComObjActive("PowerPoint.Application")
    presentation := pptApp.ActivePresentation

    ; プレゼンテーションが保存されているか確認
    if (presentation.Path = "")
    {
        MsgBox, 48, エラー, プレゼンテーションが保存されていません。
    }
    else
    {
        ; プレゼンテーションのファイルパスを取得
        filePath := presentation.FullName

        ; クリップボードにファイルパスをコピー
        Clipboard := filePath
        ;MsgBox, 64, 完了, ファイルパスがクリップボードにコピーされました: `n%filePath%
        Tooltip, %filePath%
    }
    return

+^E::
    ; 拡大率の設定（10%拡大）
    scaleFactor := 1.1  ; 拡大したい場合は1.1、縮小したい場合は0.9など

    ; PowerPointを起動して現在の選択にアクセス
    pptApp := ComObjActive("PowerPoint.Application")
    shapes := pptApp.ActiveWindow.Selection.ShapeRange

    ; 図形が選択されているか確認
    if shapes.Count > 0
    {
        ; 選択されたすべての図形に対して拡大率を変更
        Loop, % shapes.Count
        {
            shape := shapes.Item(A_Index)  ; 図形をItemメソッドで参照

            ; 現在の幅と高さを取得
            currentWidth := shape.Width
            currentHeight := shape.Height

            ; 拡大率に基づいて新しい幅と高さを計算
            shape.Width := currentWidth * scaleFactor
            shape.Height := currentHeight * scaleFactor
        }

        ; 拡大率のメッセージを表示
        newScalePercent := Round((scaleFactor * 100) - 100) ; 拡大率を計算
        ;MsgBox, 64, 完了, 選択された図形が拡大されました: " . newScalePercent . "%"
    }
    else
    {
        MsgBox, 48, エラー, 図形が選択されていません。
    }
    return


+^!S:: ; Ctrl+Shift+Fでフォントサイズ変更
    pptApp := ComObjActive("PowerPoint.Application")

    ; 現在の選択がテキストか確認
    if (pptApp.ActiveWindow.Selection.Type = 3) ; 3 = ppSelectionText
    {
        InputBox, fontSize, Font Size, Enter the font size:, , 300, 150
        if ErrorLevel
            return  ; キャンセルされた場合

        ; 入力値が有効な数値か確認
        if (fontSize >= 1 && fontSize <= 999)
        {
            ; 選択されたテキストのフォントサイズを変更
            pptApp.ActiveWindow.Selection.TextRange.Font.Size := fontSize
            ; MsgBox, 64, Completed, Font size set to %fontSize%.
        }
        else
        {
            MsgBox, 48, Error, Invalid font size. Please enter a value between 1 and 999.
        }
    }
    else
    {
        MsgBox, 48, Error, Please select some text first.
    }
return

;#IfWinActive, ahk_exe POWERPNT.EXE
#IfWinActive, ahk_class

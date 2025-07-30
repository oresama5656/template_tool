#NoEnv
#Persistent
#SingleInstance Force
SendMode Input
SetTitleMatchMode 2
SetWorkingDir %A_ScriptDir%

global currentTemplate := ""
global editMode := ""
global originalName := ""
global buttonCount := 0

Gui, Add, Text,, Alt+Vで貼り付け
buttonCount := 0

; templates.txtが存在しない場合は作成
IfNotExist, templates.txt
{
    FileAppend,, templates.txt
}

Loop, Read, templates.txt
{
    if (RegExMatch(A_LoopReadLine, "^(.*?)::", match))
    {
        label := match1
        buttonCount++
        ; 奇数番目（1,3,5...）は左列、偶数番目（2,4,6...）は右列
        if (Mod(buttonCount, 2) = 1) {
            ; 左列（奇数番目）
            Gui, Add, Button, gSetTemplate vBtn%label% w45 x10, %label%
        } else {
            ; 右列（偶数番目）- 前のボタンの右側に配置
            Gui, Add, Button, gSetTemplate vBtn%label% w45 x+5, %label%
        }
    }
}
Gui, Add, Text, vCurrentTemplateText w100 x10, 現在：なし
Gui, Add, Button, gEditTemplates w40 x10, 編集
Gui, Add, Text, gDragWindow w40 h20 Center x+5, 移動≡
Gui, Show, x0 y0 AutoSize, テンプレ貼付ランチャー
; 常に最前面に設定
Gui, +AlwaysOnTop
return

; 選択ボタン押下時
SetTemplate:
GuiControlGet, buttonLabel, , %A_GuiControl%
content := GetTemplateContent(buttonLabel)
currentTemplate := content
GuiControl,, CurrentTemplateText, 現在：%buttonLabel%
return

; Alt+Vで実行
!v::
if (currentTemplate = "")
{
    MsgBox, テンプレートが選択されていません。
    return
}

; [TAB]が含まれている場合の処理
if (InStr(currentTemplate, "[TAB]")) {
    ; 改行で囲まれた[TAB]がある場合はブロック別処理
    if (InStr(currentTemplate, "`n[TAB]`n")) {
        ; [TAB]で分割してブロック処理
        blocks := StrSplit(currentTemplate, "`n[TAB]`n")
        for index, block in blocks {
            ; 空のブロックはスキップ
            if (Trim(block) = "")
                continue
                
            ; ブロックを一括貼り付け（先頭末尾の空白行のみ削除、改行は保持）
            cleanBlock := RegExReplace(block, "^[\r\n]+|[\r\n]+$", "")
            ; 改行文字をWindows形式に統一
            cleanBlock := RegExReplace(cleanBlock, "\r?\n", "`r`n")
            ; ブロック内の[TAB]をタブ文字に変換
            cleanBlock := RegExReplace(cleanBlock, "\[TAB\]", "`t")
            Clipboard := cleanBlock
            ClipWait, 1
            Send ^v
            Sleep 200
            
            ; 最後のブロックでなければTabキーを送信
            if (index < blocks.MaxIndex()) {
                Send {Tab}
                Sleep 200
            }
        }
    } else {
        ; 同一行の[TAB]をタブキーとして送信
        parts := StrSplit(currentTemplate, "[TAB]")
        for index, part in parts {
            ; パートを貼り付け
            if (Trim(part) != "") {
                processedPart := RegExReplace(part, "\r?\n", "`r`n")
                Clipboard := processedPart
                ClipWait, 1
                Send ^v
                Sleep 100
            }
            
            ; 最後のパートでなければTabキーを送信
            if (index < parts.MaxIndex()) {
                Send {Tab}
                Sleep 100
            }
        }
    }
} else {
    ; [TAB]がない場合は一括コピペ
    ; 改行文字をWindows形式に統一
    processedTemplate := RegExReplace(currentTemplate, "\r?\n", "`r`n")
    Clipboard := processedTemplate
    ClipWait, 1
    Send ^v
}
return

; テンプレート読み込み関数
GetTemplateContent(label) {
    content := ""
    reading := false
    Loop, Read, templates.txt
    {
        line := A_LoopReadLine
        if (RegExMatch(line, "^(.*?)::", match)) {
            if (reading)
                break
            if (match1 = label) {
                reading := true
                continue
            }
        } else if (reading) {
            ; 改行文字を明示的に追加
            if (content != "")
                content .= "`r`n"
            content .= line
        }
    }
    return content
}

; 編集ボタン押下時
EditTemplates:
Gui, 2:Destroy
Gui, 2:Add, Text,, テンプレート一覧：
Gui, 2:Add, ListBox, vTemplateList w300 h200 gSelectTemplate

; テンプレート一覧を読み込み
Loop, Read, templates.txt
{
    if (RegExMatch(A_LoopReadLine, "^(.*?)::", match))
    {
        GuiControl, 2:, TemplateList, %match1%
    }
}

Gui, 2:Add, Button, gAddTemplate w80 h30, 追加
Gui, 2:Add, Button, gEditTemplate w80 h30 x+10, 編集
Gui, 2:Add, Button, gDeleteTemplate w80 h30 x+10, 削除
Gui, 2:Add, Button, gOpenAI w80 h30 x+10, AI
Gui, 2:Add, Button, gMoveUp w80 h30 x10, 上へ
Gui, 2:Add, Button, gMoveDown w80 h30 x+10, 下へ
Gui, 2:Add, Button, gCloseEditor w80 h30 x+10, 閉じる

Gui, 2:Show, , テンプレート編集(４文字以内推奨)
return

; テンプレート選択時
SelectTemplate:
return

; 新規追加
AddTemplate:
Gui, 3:Destroy
Gui, 3:Add, Text,, テンプレート名：
Gui, 3:Add, Edit, vTemplateName w300
Gui, 3:Add, Text,, テンプレート内容：
Gui, 3:Add, Edit, vTemplateContent w300 h200 Multi
Gui, 3:Add, Button, gInsertTab w100 h25, タブキー挿入
Gui, 3:Add, Button, gSaveTemplate w80 h30, 保存
Gui, 3:Add, Button, gCancelForm w80 h30 x+10, キャンセル

editMode := "追加"
originalName := ""

Gui, 3:Show, , テンプレート追加
return

; 編集
EditTemplate:
Gui, 2:Submit, NoHide
GuiControlGet, selectedTemplate, 2:, TemplateList
if (selectedTemplate = "") {
    MsgBox, テンプレートを選択してください。
    return
}
content := GetTemplateContent(selectedTemplate)

Gui, 3:Destroy
Gui, 3:Add, Text,, テンプレート名：
Gui, 3:Add, Edit, vTemplateName w300, %selectedTemplate%
Gui, 3:Add, Text,, テンプレート内容：
Gui, 3:Add, Edit, vTemplateContent w300 h200 Multi, %content%
Gui, 3:Add, Button, gInsertTab w100 h25, タブキー挿入
Gui, 3:Add, Button, gSaveTemplate w80 h30, 保存
Gui, 3:Add, Button, gCancelForm w80 h30 x+10, キャンセル

editMode := "編集"
originalName := selectedTemplate

Gui, 3:Show, , テンプレート編集
return

; 削除
DeleteTemplate:
Gui, 2:Submit, NoHide
GuiControlGet, selectedTemplate, 2:, TemplateList
if (selectedTemplate = "") {
    MsgBox, テンプレートを選択してください。
    return
}
MsgBox, 4, 確認, %selectedTemplate% を削除しますか？
IfMsgBox Yes
{
    DeleteTemplateFromFile(selectedTemplate)
    ; リストを再読み込み
    RefreshTemplateList()
}
return

; 編集画面を閉じる
CloseEditor:
Gui, 2:Destroy
; メイン画面のボタンのみ更新（位置は保持）
RefreshMainGuiButtons()
return

; テンプレートを上へ移動
MoveUp:
Gui, 2:Submit, NoHide
GuiControlGet, selectedTemplate, 2:, TemplateList
if (selectedTemplate = "") {
    MsgBox, テンプレートを選択してください。
    return
}
MoveTemplateUp(selectedTemplate)
; リストを再読み込み
RefreshTemplateList()
; 移動したテンプレートを再選択
GuiControl, 2:Choose, TemplateList, %selectedTemplate%
return

; テンプレートを下へ移動
MoveDown:
Gui, 2:Submit, NoHide
GuiControlGet, selectedTemplate, 2:, TemplateList
if (selectedTemplate = "") {
    MsgBox, テンプレートを選択してください。
    return
}
MoveTemplateDown(selectedTemplate)
; リストを再読み込み
RefreshTemplateList()
; 移動したテンプレートを再選択
GuiControl, 2:Choose, TemplateList, %selectedTemplate%
return

; テンプレートリストを更新する関数
RefreshTemplateList() {
    GuiControl, 2:, TemplateList, |
    Loop, Read, templates.txt
    {
        if (RegExMatch(A_LoopReadLine, "^(.*?)::", match))
        {
            GuiControl, 2:, TemplateList, %match1%
        }
    }
}

; テンプレート保存
SaveTemplate:
Gui, 3:Submit
if (TemplateName = "") {
    MsgBox, テンプレート名を入力してください。
    return
}

; 新規追加時の重複チェック
if (editMode = "追加") {
    Loop, Read, templates.txt
    {
        if (RegExMatch(A_LoopReadLine, "^(.*?)::", match) && match1 = TemplateName) {
            MsgBox, そのテンプレート名は既に存在します。
            return
        }
    }
}

if (editMode = "編集") {
    ; 既存テンプレートを元の位置で置き換え
    ReplaceTemplateInFile(originalName, TemplateName, TemplateContent)
} else {
    ; 新しいテンプレートを追加
    SaveTemplateToFile(TemplateName, TemplateContent)
}

; 編集画面のリストを更新
RefreshTemplateList()

Gui, 3:Destroy
return

; タブキー挿入
InsertTab:
; フォーカスをテンプレート内容の入力欄に設定
ControlFocus, Edit2, A
; カーソル位置に[TAB]を挿入
ControlSend, Edit2, [TAB], A
return

; フォームキャンセル
CancelForm:
Gui, 3:Destroy
return

; テンプレートをファイルに保存
SaveTemplateToFile(name, content) {
    ; 改行文字をファイル保存用に統一
    content := RegExReplace(content, "\r?\n", "`n")
    
    ; ファイルが存在するかチェック
    FileGetSize, fileSize, templates.txt
    if (fileSize > 0) {
        ; ファイルが空でない場合は改行を追加
        FileAppend, `n%name%::`n%content%`n, templates.txt
    } else {
        ; ファイルが空の場合は改行なしで開始
        FileAppend, %name%::`n%content%`n, templates.txt
    }
}

; テンプレートをファイルから削除
DeleteTemplateFromFile(name) {
    FileRead, templateContent, templates.txt
    newContent := ""
    reading := false
    skipEmptyLines := false
    
    Loop, Parse, templateContent, `n, `r
    {
        line := A_LoopField
        if (RegExMatch(line, "^(.*?)::", match)) {
            if (match1 = name) {
                reading := true
                skipEmptyLines := true
                continue
            } else {
                reading := false
                skipEmptyLines := false
            }
        }
        
        if (!reading) {
            ; 削除したテンプレートの後の空行をスキップ
            if (skipEmptyLines && Trim(line) = "") {
                continue
            }
            skipEmptyLines := false
            newContent .= line . "`n"
        }
    }
    
    ; 末尾の余分な改行を削除して、最後に1つの改行を追加
    newContent := RTrim(newContent, "`n") . "`n"
    
    FileDelete, templates.txt
    FileAppend, %newContent%, templates.txt
}

; メインGUIのボタンのみ更新
RefreshMainGuiButtons() {
    global
    ; 現在の位置を記憶
    WinGetPos, currentX, currentY,,, テンプレ貼付ランチャー
    
    ; メインGUIのみを再構築（より確実な方法）
    Gui, 1:Destroy
    
    ; 新しいメインGUIを作成
    Gui, 1:Add, Text,, Alt+Vで貼り付け
    buttonCount := 0
    Loop, Read, templates.txt
    {
        if (RegExMatch(A_LoopReadLine, "^(.*?)::", match))
        {
            label := match1
            buttonCount++
            ; 奇数番目（1,3,5...）は左列、偶数番目（2,4,6...）は右列
            if (Mod(buttonCount, 2) = 1) {
                ; 左列（奇数番目）
                Gui, 1:Add, Button, gSetTemplate vBtn%label% w45 x10, %label%
            } else {
                ; 右列（偶数番目）- 前のボタンの右側に配置
                Gui, 1:Add, Button, gSetTemplate vBtn%label% w45 x+5, %label%
            }
        }
    }
    Gui, 1:Add, Text, vCurrentTemplateText w100 x10, 現在：なし
    Gui, 1:Add, Button, gEditTemplates w40 x10, 編集
    Gui, 1:Add, Text, gDragWindow w40 h20 Center x+5, 移動≡
    
    ; 現在選択されているテンプレートをリセット
    currentTemplate := ""
    
    ; 元の位置で表示
    if (currentX != "" && currentY != "") {
        Gui, 1:Show, x%currentX% y%currentY% AutoSize, テンプレ貼付ランチャー
    } else {
        Gui, 1:Show, x0 y0 AutoSize, テンプレ貼付ランチャー
    }
    ; 常に最前面に設定
    Gui, 1:+AlwaysOnTop
}

; メインGUI更新
RefreshMainGui() {
    RefreshMainGuiButtons()
}

; ドラッグ用アイコンクリック時
DragWindow:
PostMessage, 0xA1, 2,,, A
return

GuiClose:
ExitApp

2GuiClose:
Gui, 2:Destroy
return

3GuiClose:
Gui, 3:Destroy
return

; テンプレートを元の位置で置き換え
ReplaceTemplateInFile(oldName, newName, newContent) {
    ; 改行文字をファイル保存用に統一
    newContent := RegExReplace(newContent, "\r?\n", "`n")
    
    FileRead, templateContent, templates.txt
    newFileContent := ""
    reading := false
    
    Loop, Parse, templateContent, `n, `r
    {
        line := A_LoopField
        if (RegExMatch(line, "^(.*?)::", match)) {
            if (match1 = oldName) {
                ; 古いテンプレートの開始位置で新しいテンプレートを挿入
                newFileContent .= newName . "::" . "`n" . newContent . "`n"
                reading := true
                continue
            } else {
                reading := false
            }
        }
        
        if (!reading) {
            newFileContent .= line . "`n"
        }
    }
    
    ; 末尾の余分な改行を削除して、最後に1つの改行を追加
    newFileContent := RTrim(newFileContent, "`n") . "`n"
    
    FileDelete, templates.txt
    FileAppend, %newFileContent%, templates.txt
}

; テンプレートを上へ移動
MoveTemplateUp(templateName) {
    FileRead, templateContent, templates.txt
    templates := []
    currentTemplateData := ""
    reading := false
    
    ; テンプレートをすべて配列に読み込み
    Loop, Parse, templateContent, `n, `r
    {
        line := A_LoopField
        if (RegExMatch(line, "^(.*?)::", match)) {
            if (reading && currentTemplateData != "") {
                templates.Push(currentTemplateData)
            }
            currentTemplateData := line . "`n"
            reading := true
        } else if (reading) {
            currentTemplateData .= line . "`n"
        }
    }
    ; 最後のテンプレートを追加
    if (reading && currentTemplateData != "") {
        templates.Push(currentTemplateData)
    }
    
    ; 指定されたテンプレートを見つけて上に移動
    targetIndex := 0
    Loop, % templates.Length()
    {
        if (RegExMatch(templates[A_Index], "^(.*?)::", match) && match1 = templateName) {
            targetIndex := A_Index
            break
        }
    }
    
    ; 最初の要素でなければ上に移動
    if (targetIndex > 1) {
        temp := templates[targetIndex]
        templates[targetIndex] := templates[targetIndex - 1]
        templates[targetIndex - 1] := temp
        
        ; ファイルに書き戻し
        newContent := ""
        Loop, % templates.Length()
        {
            newContent .= templates[A_Index]
        }
        newContent := RTrim(newContent, "`n") . "`n"
        
        FileDelete, templates.txt
        FileAppend, %newContent%, templates.txt
    }
}

; テンプレートを下へ移動
MoveTemplateDown(templateName) {
    FileRead, templateContent, templates.txt
    templates := []
    currentTemplateData := ""
    reading := false
    
    ; テンプレートをすべて配列に読み込み
    Loop, Parse, templateContent, `n, `r
    {
        line := A_LoopField
        if (RegExMatch(line, "^(.*?)::", match)) {
            if (reading && currentTemplateData != "") {
                templates.Push(currentTemplateData)
            }
            currentTemplateData := line . "`n"
            reading := true
        } else if (reading) {
            currentTemplateData .= line . "`n"
        }
    }
    ; 最後のテンプレートを追加
    if (reading && currentTemplateData != "") {
        templates.Push(currentTemplateData)
    }
    
    ; 指定されたテンプレートを見つけて下に移動
    targetIndex := 0
    Loop, % templates.Length()
    {
        if (RegExMatch(templates[A_Index], "^(.*?)::", match) && match1 = templateName) {
            targetIndex := A_Index
            break
        }
    }
    
    ; 最後の要素でなければ下に移動
    if (targetIndex > 0 && targetIndex < templates.Length()) {
        temp := templates[targetIndex]
        templates[targetIndex] := templates[targetIndex + 1]
        templates[targetIndex + 1] := temp
        
        ; ファイルに書き戻し
        newContent := ""
        Loop, % templates.Length()
        {
            newContent .= templates[A_Index]
        }
        newContent := RTrim(newContent, "`n") . "`n"
        
        FileDelete, templates.txt
        FileAppend, %newContent%, templates.txt
    }
}

; AIボタン押下時
OpenAI:
; prompt.txtの内容をクリップボードにコピー
FileRead, promptContent, prompt.txt
if ErrorLevel {
    MsgBox, prompt.txtファイルが見つかりません。
    return
}
Clipboard := promptContent
ClipWait, 1

; Geminiのウェブサイトを開く
Run, https://gemini.google.com/?hl=ja

; ユーザーに通知
MsgBox, 4160, 通知, prompt.txtの内容をクリップボードにコピーしました。`n`n1. Geminiのチャット欄に貼り付けてください`n2. 生成されたテンプレートをtemplates.txtに貼り付けてください`n3. 編集画面を閉じて再度開くとボタンに反映されます, 5
return

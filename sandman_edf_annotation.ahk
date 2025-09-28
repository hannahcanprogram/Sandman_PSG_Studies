#Requires AutoHotkey v2.0
#SingleInstance Force
SetTitleMatchMode(2)
SetWorkingDir(A_ScriptDir)

UpdateStatus(row, status, extra := "") {
    ;xlsx  := "D:\PSG_test\test_paths.xlsx"
    xlsx  := "E:\Machine_01\Sandman\Metadata\group4_A.xlsx"
    sheet := "Sheet1"
    extra2 := StrReplace(extra, '"', "''")
    cmd := A_ComSpec
        . ' /c python "' A_ScriptDir '\update_status.py" '
        . '"' xlsx '" "' sheet '" ' row ' ' status ' "' extra2 '"'
    RunWait(cmd, , "Hide")
}

EnsureFocus(ctrl, win) {
    Loop 3 {
        ControlFocus(ctrl, win)
        Sleep 80
        WinActivate(win)
        Sleep 60
    }
}

IsOnDrives(win) {
    txt := WinGetText(win)
    return InStr(txt, "Path:") && InStr(txt, "Media Type")
}

getListCtrl(title) {
    for ctrl in WinGetControls(title)
        if RegExMatch(ctrl, "i)^SysListView\d+$")
            return ctrl
    return ""
}

ClearAndSet(ctrl, text, winTitle) {
    WinActivate winTitle
    if !WinWaitActive(winTitle, , 2)
        throw Error("window is not activate: " winTitle)

    ControlFocus ctrl, winTitle
    Sleep 500

    if !ControlSetText("", ctrl, winTitle) {
        ; WM_SETTEXT to remove
        SendMessage 0x000C, 0, 0, ctrl, winTitle  ; WM_SETTEXT, wParam=0, lParam=NULL
    }
    Sleep 500

    if !ControlSetText(text, ctrl, winTitle) {
        ; WM_SETTEXT writes
        SendMessage 0x000C, 0, StrPtr(text), ctrl, winTitle
    }
    Sleep 500

    cur := ControlGetText(ctrl, winTitle)
    if (cur != text)
        throw Error("failed writes: " ctrl " expect=" text " actual=" cur)
}

SendToTree(dmSpec, keys) {
    ; 1) SysTreeView321 / Tree1
    for ctrl in ["SysTreeView321", "Tree1"] {
        try {
            ControlFocus(ctrl, dmSpec)
            Sleep 120
            ControlSend(ctrl, keys, dmSpec)
            return
        }
        catch {
            ; ingnore
        }
    }

    ; 2) plan b, click tree empty box
    container := "AfxWnd90u1"
    x := y := w := h := 0
    try ControlGetPos(&x, &y, &w, &h, container, dmSpec)
    catch {
        WinActivate(dmSpec)
        WinWaitActive(dmSpec,, 3)
        MouseGetPos(&mx, &my)
        Click "Left", 100, 160   ; screen coordinates, updates when needed
        Sleep 120
        Send keys
        return
    }

    cx := 40, cy := 60 
    ControlClick(container, dmSpec,,, 1, "NA x" cx " y" cy)
    Sleep 120
    Send keys
}

DirHasLock(dir) {
    for pat in ["*.ldb", "*.laccdb"]          ; Jet/ACE 两种锁
        Loop Files, dir "\" pat, "F"
            return true
    return false
}

WaitNoLock(dir, timeoutMs := 60000, pollMs := 250) {
    deadline := A_TickCount + timeoutMs
    while (A_TickCount < deadline) {
        if !DirHasLock(dir)
            return true
        Sleep pollMs
    }
    return false
}

; --- 等文件大小稳定 ---
WaitFileStable(fp, stableMs := 4000, timeoutMs := 1800000, pollMs := 500) {
    start := A_TickCount
    lastSize := -1
    lastChange := A_TickCount
    while (A_TickCount - start < timeoutMs) {
        if FileExist(fp) {
            size := FileGetSize(fp, "Raw")
            if (size = lastSize) {
                if (A_TickCount - lastChange >= stableMs)
                    return true   ; 存在且连续 stableMs 未增长
            } else {
                lastSize := size
                lastChange := A_TickCount
            }
        } else {
            ; 如果进度窗已经消失且文件还不存在，多半是失败/取消
            if !WinExist("Convert File")
                return false
        }
        Sleep pollMs
    }
    return false
}



; -------------------- Main Loop --------------------
Loop {
    try {
        ; ---------------------------Step 1: get next psg path and name from get_next.py-------------------------------------------------

        cmd := A_ComSpec ' /c python "' A_ScriptDir '\get_next.py" "E:\Machine_01\Sandman\Metadata\group4_A.xlsx" "Sheet1" > "' A_ScriptDir '\result.json"'
        RunWait(cmd, , "Hide")

        result := FileRead(A_ScriptDir "\result.json", "UTF-8")

        path := RegExReplace(result, '.*"path":\s*"(.*?)".*', "$1")
        name := RegExReplace(result, '.*"name":\s*"(.*?)".*', "$1")
        result := FileRead(A_ScriptDir "\result.json", "UTF-8")

        if (Trim(result) = "")
        {
            MsgBox "failed reading result.json or empty, process ended", "error", "Iconx"
            ExitApp
        }

        if RegExMatch(result, '^\s*\{\s*\}\s*$')
        {
            MsgBox "Excel no records, process ended."
            ExitApp
        }

        if RegExMatch(result, 's)"path"\s*:\s*"na"')
        {
            MsgBox 'checked path == "na", process ended'
            ExitApp
        }

        m := ""

        ; path
        if !RegExMatch(result, 's)"path"\s*:\s*"(.*?)"', &m)
            throw Error("path not found in JSON: " result)
        path := m[1]
        path := StrReplace(path, '\"', '"')
        path := StrReplace(path, '\\', '\')

        ; name
        if !RegExMatch(result, 's)"name"\s*:\s*"(.*?)"', &m)
            throw Error("name not found in JSON: " result)
        name := m[1]

        ; row
        if !RegExMatch(result, 's)"row"\s*:\s*(\d+)', &m)
            throw Error("row not found in JSON: " result)
        row := Number(m[1])

        if (Trim(path) = "" || row <= 0)
        {
            MsgBox "no valid path/row found in JSON, process ended", "error", "Iconx"
            ExitApp
        }

        ;MsgBox "Next path: " path "`nNext name: " name, "Debug Info"

        ; ----------------------------Step 2: Activate Sandman Navigation Screen---------------------------------------------------------
        sandTitle := "Sandman Navigation Screen ahk_class #32770"
        cfgWin    := "Configuration ahk_class #32770"
        WinWait(sandTitle,,30), WinActivate(sandTitle), WinWaitActive(sandTitle,,5)


        ; ---------- Try #1: Focus + Space ----------
        EnsureFocus("Button4", sandTitle)
        ControlSend("{Space}", "Button4", sandTitle)

        if !WinWait(cfgWin,,1) {
            PostMessage(0x00F5, 0, 0, "Button4", sandTitle)  ; BM_CLICK
        }

        ; Try #2 ControlClick
        if !WinWait(cfgWin,,1) {
            Loop 3 {
                ControlFocus("Button4", sandTitle)
                Sleep 250
                ControlClick("Button4", sandTitle,,,1,"NA")
                if WinWait(cfgWin,,1)
                    break
                Sleep 250
            }
        }

        if !WinExist(cfgWin)
            throw Error("failed to click Configuration")
    
        ;------------------------------Step 3: map to drives in Configuration------------------------------------------------------------
        tabCtrl   := "SysTabControl321" ; 
        confTitle := "Configuration ahk_class #32770"   ; or "ahk_exe Config.exe"
        
        WinActivate(confTitle), WinWaitActive(confTitle,,5)

        EnsureFocus(tabCtrl, confTitle)

        ;Site Information → General → Drives，tab twice→
        ;Send "{Right 2}"
        ;Sleep 100

        ControlSend("{Home}{Right 2}", tabCtrl, confTitle)
        Sleep 100
        if IsOnDrives(confTitle)
            goto __tab_ok

        ; TCM_SETCURSEL（0：Site Info=0, General=1, Drives=2）
        PostMessage(0x130C, 2, 0, tabCtrl, confTitle) ; TCM_SETCURSEL
        Sleep 100
        ControlClick(tabCtrl, confTitle,,,1,"NA x160 y12") 
        Sleep 100
        if IsOnDrives(confTitle)
            goto __tab_ok
        
        __tab_ok: ; in Drives

        ;-----------------------Delete-------------------------
        listCtrl := getListCtrl(confTitle)
        if !listCtrl {
            MsgBox "could not find control list (SysListViewNN)。`ncontrols：`n`n" . WinGetControls(confTitle).Join("`n")
            ExitApp
        }

        EnsureFocus(listCtrl, confTitle)
        Sleep 100

        ControlClick listCtrl, confTitle,,, 1, "NA"
        Sleep 250
        ControlSend "{Home}", listCtrl, confTitle
        Sleep 250
        ControlSend "{Down}", listCtrl, confTitle
        Sleep 250
        ControlSend "{Space}", listCtrl, confTitle
        Sleep 250
        
        EnsureFocus("Button4", confTitle) ; Button4: Delete
        Send("!d")
        Sleep 100

        ;-----------------------Add-------------------------
        path := StrReplace(path, "\\", "\")  ; \\ -> \
        path := StrReplace(path, '\"', '"')  ; \" -> "
        ClearAndSet("Edit1", path, confTitle)   ; Path
        ClearAndSet("Edit2", name, confTitle)   ; Name

        ;ControlClick "Button3", confTitle
        EnsureFocus("Button3", confTitle) ; Button3: ADD
        ControlClick "Button3", confTitle
        Sleep 500


        ;ControlClick "Button8", confTitle
        EnsureFocus("Button8", confTitle) ; Button8: OK

        Send("!o")
        Sleep 500

        ;------------------------------Step 4: Data Management-------------------------------------------------------------------------

        sandTitle := "Sandman Navigation Screen ahk_class #32770"
        dmSpec    := "Sandman Elite (Data Management) ahk_exe Data Management.exe"

        WinWait(sandTitle,,30), WinActivate(sandTitle), WinWaitActive(sandTitle,,5)

        ; Data Management（Button3 ——> dmSpec
        EnsureFocus("Button3", sandTitle)
        ControlSend("{Space}", "Button3", sandTitle)

        WinWait(dmSpec,,45), WinActivate(dmSpec), WinWaitActive(dmSpec,,30)

        
        SplitPath(path, , &lockDir)
        ; 如果实际锁文件在上一级目录，就再执行一次 SplitPath(lockDir, , &lockDir)

        if !WaitNoLock(lockDir, 60000) {
            UpdateStatus(row, "skip_lock", "lock timeout at " lockDir)
            continue
        }

        WinActivate(dmSpec)
        WinWaitActive(dmSpec,, 3)

        EnsureFocus("SysTreeView321", dmSpec)

        Send("{Ctrl up}{Shift up}{Alt up}")

        ; ↑Down →Right ↓Down
        ControlSend("{Down}",  "SysTreeView321", dmSpec)
        Sleep 150
        ControlSend("{Right}", "SysTreeView321", dmSpec)
        Sleep 150
        ControlSend("{Down}",  "SysTreeView321", dmSpec)

        ; Down -> Right -> Down
        ;SendToTree(dmSpec, "{Down}")
        ;Sleep 250
        ;SendToTree(dmSpec, "{Right}")
        ;Sleep 250
        ;SendToTree(dmSpec, "{Down}")
        ;Sleep 250

        Loop 7 {
            Send "{Tab}"
            Sleep 250
        }
        ;Send "{Enter}"   ; Try #1 convert
        Send("!o") 

        ; Try #2 
        dlg   := "Select Destination ahk_class #32770"
        ;if !WinWait(dlg,,2) {
        ;    PostMessage(0x00F5, 0, 0, "Button6", dmSpec)  ; BM_CLICK
        ;}

        ;----------------------------- Select Destination-------------------------------
        WinWait(dlg,,45), WinActivate(dlg), WinWaitActive(dlg,,5)
        if !WinExist(dlg)
            throw Error("failed to click Convert")
        ;Button6 is the destination folder
        tries := 0
        while (tries < 10) {
            EnsureFocus("Button6", dlg)
            ControlClick("Button6", dlg)
            Sleep 500
            focused := ControlGetFocus(dlg)
            if (focused = "Button6") {
                break   ;
            }
            tries++
        }
        Loop 2 {
            Send "{Tab}"
            Sleep 250
        }
        Send "{Enter}"

        ;---------------------------Select Target Data File Format----------------------
        dlg2 := "Select Target Data File Format ahk_exe Data Management.exe"
        WinWait(dlg2,,45), WinActivate(dlg2), WinWaitActive(dlg2,,5)
        Loop 2 {
            Send "{Tab}"
            Sleep 120
        }
        Send("{Up}") ; Select EDF File Format
        Sleep(150)
        Send("{Enter}") ; click OK

        ;----------------------------Enter EDF Filename---------------------------------
        dlg3 := "Enter EDF Filename ahk_exe Data Management.exe"
        WinWait(dlg3,,45), WinActivate(dlg3), WinWaitActive(dlg3,,5)
        Sleep(1000)     ;keep default file name

        ControlSetText("", "Edit1", dlg3)
        Sleep(1000) 
        ;ControlSetText(name, "Edit1", dlg3)
        ControlSetText(name . ".REC", "Edit1", dlg3)
        Sleep(1000)

        Send("{Enter}") ; ControlClick("Button1", dlg3)


        ;------------------------Resample channels and start converting-----------------
        dlg4 := "Resample Channels ahk_exe Data Management.exe"
        WinWait(dlg4,,45), WinActivate(dlg4), WinWaitActive(dlg4,,5)
        Sleep(5000)                  ; wait for loading dataset
        Send("{Enter}")              ; click OK

        ;------------------------Handle Data Management Warning-------------------------
        ;dlg5 := "Data Management ahk_exe Data Management.exe"
        ;WinWait(dlg5,,45), WinActivate(dlg5), WinWaitActive(dlg5,,5)
        ;Sleep(300)
        ;Send("{Enter}")               ; click OK
        ;Sleep(1000)

        ;---------------------------Convert File-----------------------------------------
        WinWait("Convert File", , 1800)
        sleep(2000)

        WinWaitClose("Convert File", , 1800)
        sleep(2000)
        

        ;---------------------------------------Step 5: View--------------------------------------------------------------
        dmSpec    := "Sandman Elite (Data Management) ahk_exe Data Management.exe"
        WinWait(dmSpec,,45), WinActivate(dmSpec), WinWaitActive(dmSpec,,5)

        SendToTree(dmSpec, "{Down}")
        Sleep 250
        SendToTree(dmSpec, "{Right}")
        Sleep 250
        SendToTree(dmSpec, "{Down}")
        Sleep 250
        Loop 2 {
            Send "{Tab}"
            Sleep 250
        }
        ;Send "{Enter}"
        
        ;WinActivate(dmSpec), WinWaitActive(dmSpec,,5)
        sleep 1000
        Send("!v")
        ;Sleep 1000

        ;dlg   := "Analysis ahk_class #32770"
        ;if !WinWait(dlg,,1) {
        ;    PostMessage(0x00F5, 0, 0, "Button1", dmSpec)  ; BM_CLICK   
        ;}

        Sleep 3000
        ; -------------------Analysis Warning----------------
        WinWait("Analysis", , 60), WinActivate("Analysis"), WinWaitActive("Analysis",,5)
        Sleep 1000
        Send("!o")
        Sleep 3000

        ; -----------------Select Score----------------------
        WinWait("Select Score", , 60), WinActivate("Select Score"), WinWaitActive("Select Score",,5)
        Send "{Up 2}"
        Sleep 1000
        Send "{Enter}"
        Sleep 6000

        ; ---------------------Analysis Warning 2--------------------------
        WinWait("Analysis", , 60), WinActivate("Analysis"), WinWaitActive("Analysis",,5)
        Sleep 2000
        Send "{Enter}"
        Sleep 3000

        ; ==========
        DismissAnalysisWarnings(timeout := 2.0, maxLoops := 5) {
            title := "Analysis ahk_class #32770"

            Loop maxLoops {
                if !WinWait(title, , timeout)
                    break

                WinActivate(title)
                WinWaitActive(title, , 1)

                Send("{Ctrl up}{Shift up}{Alt up}")

                try {
                    if !ControlClick("Button1", title) {
                        ; 如果 Button1
                        Send("{Enter}")
                    }
                } catch {
                    Send("{Enter}")
                }

                Sleep 300
            }
        }
        DismissAnalysisWarnings( timeout := 2.0, maxLoops := 5 )
        Sleep 2000

        ; ---------------- Display Event Window --------------------------
        Send "{F10}"
        Sleep 1000
        Loop 6 {
            Send "{Right}"
            Sleep 500
        }
        Send "{Enter}"
        Sleep 1000
        Loop 11 {
            Send "{Down}"
            Sleep 500
        }
        Send "{Enter}"
        Sleep 5000

        ; ---------------- Display All Event Data ---------
        WinActivate("ahk_exe Analysis.exe"), WinWaitActive("ahk_exe Analysis.exe",,5)
        ControlFocus "AfxFrameOrView90u2", "ahk_exe Analysis.exe"

        winTitle := "Event Window"
        WinGetClientPos &x, &y, &w, &h, winTitle
        cx := x + w//2
        cy := y + h//2

        CoordMode "Mouse", "Window"
        MouseMove cx, cy
        Sleep 1000
        Click "Right"
        Sleep 1000

        if WinWait("ahk_class #32768",, 2) {
            BlockInput "MouseMove"
            Send "{Down 2}"
            Sleep 1000
            Send "{Enter}"
            BlockInput "MouseMoveOff"
        }
        Click "Right"
        Sleep 1000

        if WinWait("ahk_class #32768",, 2) {
            BlockInput "MouseMove" 
            Send "{Up 2}"
            Sleep 1000
            Send "{Enter}"
            BlockInput "MouseMoveOff"
        }
        Sleep 1500

        ; -----------------Save Exported Data---------------
        dlg6 := "Save Exported Data"
        WinWait(dlg6, , 60), WinActivate(dlg6), WinWaitActive(dlg6,,5)
        name := RegExReplace(result, '.*"name":\s*"(.*?)".*', "$1")

        ok := false
        try {
            ControlFocus "Edit1", dlg6
            Sleep 250
            ControlSetText "Edit1", name, dlg6
            ok := true
        } catch {
            ok := false
        }
        if !ok {
            Send "!n"
            Sleep 250
            ; A
            oldClip := A_Clipboard
            A_Clipboard := ""
            A_Clipboard := name
            ClipWait 1
            Send "^a"
            Sleep 80
            Send "^v"
            Sleep 80
            A_Clipboard := oldClip

            ; B
            ;Send "^a"
            ;Sleep 80
            ;SendText name
        }

        Send "{Tab 5}"
        Sleep 300
        Send "{Enter}"
        Sleep 2000
        Send "{Enter}" 
        Sleep 1000

        Send "{Tab 11}"
        Sleep 1000
        Send "{Enter}"

        WinWait("Export Data", , 60)
        sleep(2000)
        WinWaitClose("Export Data", , 1800)

        UpdateStatus(row, "success", name)
        ;---------------------------------------Step 6: Exit--------------------------------------------------------------
        WinActivate("ahk_exe Analysis.exe"), WinWaitActive("ahk_exe Analysis.exe",,5) ; Sandman Analysis
        WinClose("ahk_exe Analysis.exe")
        Sleep 1000
        Send "{Enter}"

        WinActivate "Sandman Elite (Data Management)"
        WinWaitActive "Sandman Elite (Data Management)"
        Loop 10 {
            Send "{Tab}"
            Sleep 100
        }
        Send "{Enter}"

        ; ==========
        CloseResidualDialogs(titles := ["Analysis","Data Management"], timeout := 1.5, maxLoops := 6) {
            for t in titles {
                dlg := t " ahk_class #32770" 
                Loop maxLoops {
                    if !WinWait(dlg, , timeout)
                        break
                    WinActivate(dlg)
                    WinWaitActive(dlg, , 1)
                    Send("{Ctrl up}{Shift up}{Alt up}")
                    if !ControlClick("Button1", dlg)
                        Send("{Enter}")
                    WinWaitClose(dlg, , 2)
                    Sleep 120
                }
            }
        }
        CloseResidualDialogs(["Analysis","Data Management"], 2.0, 5) 
        if ProcessExist("Analysis.exe")
            ProcessWaitClose("Analysis.exe", 2)

        sandTitle := "Sandman Navigation Screen"
        if !WinWait(sandTitle ' ahk_class #32770',, 10)
        {
            MsgBox "10 s Sandman Navigation Screen ", "Error", "Iconx"
            ExitApp
        }
        WinActivate(sandTitle ' ahk_class #32770')
        WinWaitActive(sandTitle ' ahk_class #32770',, 3)
    
    } catch as e {
        UpdateStatus(row, "error", e.Message)
        Sleep 500
        continue
    }
}
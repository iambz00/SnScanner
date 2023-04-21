#NoEnv
#NoTrayIcon
#SingleInstance Force

;
; AHK file - UTF8 + BOM
; INI file - UTF8 (w/o BOM)
;

FormatTime, start_date, , yyyyMMdd_HHmmss
StartDir := A_WorkingDir
DefaultOutputPath := StrReplace(StartDir . "\", "\\", "\") . "output_" . start_date . ".xlsx"

SetWorkingDir, %A_ScriptDir%
CoordMode, Pixel, Client
CoordMode, Mouse, Client

Global App := {}
App.FullName := "Serial Number Scanner GUI"
App.ShortName := "SnScanner"
App.Version := "20230425"
App.WinTitle := Format("{1} v{2}", App.FullName, App.Version)

App.BinPath := A_ScriptDir . "\sccore"
;@Ahk2Exe-IgnoreBegin
App.BinPath := A_ScriptDir . "\deploy\sccore"
;@Ahk2Exe-IgnoreEnd

App.Scanner := App.BinPath . "\sccore.exe"
App.Font := App.BinPath . "\D2Coding-01.ttf"
App.IniFile := A_ScriptDir . "\" . App.ShortName . ".ini"

;App.sb := New StatusBar()

Global IniFile := App.IniFile

Global hGui
Global hwndCtrl_Log

; Build Gui
Gui, MainWindow:New, hwndhGui MinSize ;, Resize
Gui, Font,, Malgun Gothic
Gui, Margin, 8, 8

Gui, Add, Text, xm+8 y+12 w88 Right, % "이미지 경로"
Gui, Add, Edit, x+8 yp-3 w400 h22 vImageDir,
Gui, Add, Button, x+2 yp-1 w48 hp+2 gSetImageDir, 찾기
Gui, Add, Text, xm+8 y+6 w88 Right, % "Tesseract 경로"
Gui, Add, Edit, x+8 yp-3 w400 h22 vTesseractPath, 
Gui, Add, Button, x+2 yp-1 w48 hp+2 gSetTesseractPath,  찾기
Gui, Add, Text, xm+8 y+6 w88 Right, % "출력파일 경로"
Gui, Add, Edit, x+8 yp-3 w400 h22 vOutputPath, % DefaultOutputPath
Gui, Add, Button, x+2 yp-1 w48 hp+2 gSetOutputPath,  찾기
Gui, Add, Text, xm+8 y+6 w88 Right, % "검출 패턴"
Gui, Add, Edit, x+8 yp-3 w400 ReadOnly vPattern,

Gui, Add, Radio, vPatternGroup1 gSetPattern1 xm+104 y+8, 시리얼번호
Gui, Add, Radio, vPatternGroup2 gSetPattern2 x+2, MAC 주소
Gui, Add, Radio, vPatternGroup3 gSetPattern3 x+2, 기타
Gui, Add, Edit, x+2 yp-4 w240 vUserPattern,

Gui, Add, CheckBox, xm+88 y+8 vInteractOption, 검출 영역 확인하며 진행(-i)
;Gui, Add, CheckBox, xm+84 y+10 vExecResult, 완료 후 결과파일 열기(엑셀)

Gui, Add, Button, xm+16 yp-20 w60 h36 gScan, 스 캔

;@Ahk2Exe-IgnoreBegin
Gui, Add, Button, x+240 yp+16 w60 h20 gRestart, 재시작
;@Ahk2Exe-IgnoreEnd
;Gui, Add, Text, Section xs, 로그
;Gui, Add, Edit,  w520 r20 HwndhwndCtrl_Log ReadOnly vCtrl_Log

Gui, Show,, % App.WinTitle

GoSub Initialize

Log("SnScanner")
Log(DefaultOutputPath)

Return


SetImageDir:
    GuiControlGet, image_dir, , ImageDir
    FileSelectFolder, dir, *%A_WorkingDir%, 2, 이미지가 있는 폴더를 지정'
    if (dir) {
        GuiControl,, ImageDir, % dir
    }
Return

SetTesseractPath:
    GuiControlGet, tesseract_path, , TesseractPath
    FileSelectFile, file, 3, %tesseract_path%, Tesseract 실행 파일 지정, tesseract.exe
    if (file) {
        GuiControl,, TesseractPath, % file
    }
Return

SetOutputPath:
    GuiControlGet, output_path, , OutputPath
    FileSelectFile, file, 2, %output_path%, 결과파일 지정, *.xlsx
    if (file) {
        file .= ".xlsx"
        file := StrReplace(file, ".xlsx.xlsx", ".xlsx")
        GuiControl,, OutputPath, % file
    }
Return

SetPattern1:
    App.PatternGroup := 1
    GuiControl,, Pattern, % "R[A-Z0-9]{10}"
Return

SetPattern2:
    App.PatternGroup := 2
    GuiControl,, Pattern, % "..[.:]..[.:]..[.:]..[.:]..[.:].."
Return

SetPattern3:
    App.PatternGroup := 3
    GuiControlGet, user_pattern, , UserPattern
    GuiControl,, Pattern, % user_pattern
Return

Initialize:
    GoSub LoadSettings
Return

Scan:
    GuiControlGet, image_dir, , ImageDir
    GuiControlGet, tesseract_path, , TesseractPath
    GuiControlGet, output_path, , OutputPath
    GuiControlGet, search_pattern, , Pattern
    GuiControlGet, interact, , InteractOption
    GuiControlGet, open_result, , ExecResult
    if (tesseract_path) {
        tesseract_path := Format("-t ""{1}"" ", tesseract_path)
    }
    if (output_path) {
        output_path := Format("-o ""{1}"" ", output_path)
    }
    if (search_pattern) {
        search_pattern := Format("-p ""{1}"" ", StrReplace(search_pattern, "\", "\\"))
    }
    if (interact > 0) {
        interact := "-i "
    } else {
        interact := ""
    }
    command := Format(App.Scanner . " {2}{3}{4}{5}""{1}""", image_dir, tesseract_path, output_path, search_pattern, interact)
    Log(command . "`r`n")
    RunWait, % command
    /*
    if (open_result) {
        Loop, 10 {
            if (!FileExist(output_path)) {
                Sleep, 500
            }
        }
        Run, % A_ComSpec . " /c " . output_path
    }
    */
Return

Stop:
Return

GetClientSize(hWnd, ByRef w := "", ByRef h := "") {
	VarSetCapacity(rect, 16)
	DllCall("GetClientRect", "ptr", hWnd, "ptr", &rect)
	w := NumGet(rect, 8, "int")
	h := NumGet(rect, 12, "int")
}

Log(str) {
    str .= "`r`n"
    if (hwndCtrl_Log) {
        AppendText(hwndCtrl_Log, &str)
    }
}

LogClear() {
	GuiControl,, Ctrl_Log,
}

AppendText(hEdit, ptrText) {
	SendMessage, 0x000E, 0, 0,, ahk_id %hEdit% ;WM_GETTEXTLENGTH
	SendMessage, 0x00B1, ErrorLevel, ErrorLevel,, ahk_id %hEdit% ;EM_SETSEL
	SendMessage, 0x00C2, False, ptrText,, ahk_id %hEdit% ;EM_REPLACESEL
}

GuiSize:
	Gui %hGui%:Default
	if !horzMargin
		return
	ctrlW := A_GuiWidth - horzMargin
	list = Title,Status,VisText,AllText,Freeze
	Loop, Parse, list, `,
		GuiControl, Move, Ctrl_%A_LoopField%, w%ctrlW%
Return

Restart:
	GoSub SaveSettings
	Reload
Return

LoadSettings:
    IniRead, image_dir, %IniFile%, Common, Image_Dir, % StartDir
    GuiControl,, ImageDir, % image_dir
    IniRead, tesseract_path, %IniFile%, Common, Tesseract Path, C:\Program Files\Tesseract-OCR\tesseract.exe
    GuiControl,, TesseractPath, % tesseract_path
    IniRead, user_pattern, %IniFile%, Common, UserPattern, %A_Space%
    GuiControl,, UserPattern, % user_pattern

    IniRead, pattern_group, %IniFile%, Common, PatternGroup, 1
    GoSub SetPattern%pattern_group%
    GuiControl,, PatternGroup%pattern_group%, 1
    IniRead, interact, %IniFile%, Common, Interact, 0
    GuiControl,, InteractOption, % interact
    IniRead, open_result, %IniFile%, Common, Execute Result, 1
    GuiControl,, ExecResult, % open_result

    IniRead, wX, % IniFile, Common, GUI_x, 600
    IniRead, wY, % IniFile, Common, GUI_y, 50
    WinMove, ahk_id %hGui%,, %wX%, %wY%
    
Return

SaveSettings:
	WinGetPos, wX, wY
	IniWrite, % wX, % IniFile, Common, GUI_x
	IniWrite, % wY, % IniFile, Common, GUI_y
    GuiControlGet, image_dir, , ImageDir
    IniWrite, %image_dir%, %IniFile%, Common, Image_Dir
    GuiControlGet, tesseract_path, , TesseractPath
    IniWrite, %tesseract_path%, %IniFile%, Common, Tesseract Path

    IniWrite, % App.PatternGroup, %IniFile%, Common, PatternGroup
    GuiControlGet, user_pattern, , UserPattern
    IniWrite, %user_pattern%, %IniFile%, Common, UserPattern
    GuiControlGet, interact, , InteractOption
    IniWrite, %interact%, %IniFile%, Common, Interact
    GuiControlGet, open_result, , ExecResult
    IniWrite, %open_result%, %IniFile%, Common, Execute Result
;	GoSub LoadSettings
Return

MainWindowGuiClose:
MainWindowGuiEscape:
	GoSub SaveSettings
ExitApp


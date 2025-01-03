; AutoHotkey v2 script to make a window stay always on top

; Use this to allow partial matching for the window title
SetTitleMatchMode(2)

; Check if a window with the title "Ping Results" exists
if WinExist("Pingoji")
{
    ; Make the window always on top by passing 1
    WinSetAlwaysOnTop(1)  ; 1 means "on", 0 would mean "off"

    ; Remove the title bar and borders
    ;WinSetStyle("-0xC00000", "ahk_class IEFrame")

}

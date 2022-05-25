#Persistent
#SingleInstance, Force
Menu, Tray, Icon, mail-queue-2.ico


SetTimer, OutlookProtect, 50

OutlookProtect:
	;outlook protection
	;   ahk_class rctrl_renwnd32
	;   ahk_exe OUTLOOK.EXE
	;       ClassNN:	_WwG1
	;       Text:	Message
	WinGet, winID, ID, A
	if (winID = oldwinID)
	{
		return
	}
	WinGetClass, winClass, A
	;msgbox % winClass
	if (winClass ="rctrl_renwnd32")
	{
      ControlGetText, control_text, Button1, ahk_id %winID%
      ;msgbox % winID
      ;msgbox % control_text
      if (control_text = "&Send")
      {
        Control, Disable,, Button1, ahk_id %winID%
      }
	}
    oldwinID := winID
 return 

:*?:==send::
enable_send_button()
{
  global
  WinGet, winID, ID, A
  WinGetClass, winClass, A
  if (winClass ="rctrl_renwnd32")
  {
    ControlGetText, control_text, Button1, ahk_id %winID%
    ;msgbox % winID
    ;msgbox % control_text
    if (control_text = "&Send")
    {
      SetTimer, OutlookProtect, Off
      Control, Enable,, Button1, ahk_id %winID%
      Sleep, 250 ; give message window time to respond/delete hotstring text.
      ControlFocus, Button1, ahk_id %winID%
      Sleep, 5000
      Control, Disable,, Button1, ahk_id %winID%
      ControlFocus, _WwG1, ahk_id %winID%
      SetTimer, OutlookProtect, 50
    }
  }
}

:*?:==draft::
save_draft_and_close()
{
  global
  WinGet, winID, ID, A
  WinGetClass, winClass, A
  if (winClass ="rctrl_renwnd32")
  {
    Send, ^s
	Sleep, 1000
	WinClose, ahk_id %winID%
  }
}
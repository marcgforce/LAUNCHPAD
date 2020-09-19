#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=..\..\Users\0469327\Pictures\Icones\launchpad.ico
#AutoIt3Wrapper_Outfile=launchpad.exe
#AutoIt3Wrapper_Res_Description=Launcher SPAFA
#AutoIt3Wrapper_Res_Fileversion=1.0.0.0
#AutoIt3Wrapper_Res_CompanyName=SPAFA MARSEILLE
#AutoIt3Wrapper_Res_LegalCopyright=Marc GRAZIANI
#AutoIt3Wrapper_Res_Language=1036
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
; ===============================================================================================================================
; Title .........: LaunchPad
; Description ...: button deck to be used as an applications launcher (and not only)
; Author(s) .....: Chimp (Gianni Addiego)
;                  credits to  @KaFu, @Danyfirex, @mikell (see comments for references)
; Modification...: Marcgforce (drag and drop add, passing ini to links)
; ===============================================================================================================================

#include <WindowsConstants.au3>
#include <GUIConstantsEx.au3>
#include <WinAPI.au3>;<WinAPISysWin.au3>
#include <SendMessage.au3>
#include <WinAPIFiles.au3>
#include <Array.au3>
#include <file.au3>
#include <TrayConstants.au3> ; Required for the $TRAY_ICONSTATE_SHOW constant.
#Include "DragDropEvent.au3"
#include <ButtonConstants.au3>

#include <ButtonConstants.au3>
#include <GDIPlus.au3>
#include <Misc.au3>
#include <ScreenCapture.au3>
#include <WinAPIShellEx.au3>
#include <WinAPIRes.au3>
#include <WinAPISysWin.au3>

;Turn off redirection for a 32-bit script on 64-bit system.
If @OSArch = "X64" And Not @AutoItX64 Then _WinAPI_Wow64EnableWow64FsRedirection(False)

; https://docs.microsoft.com/en-us/windows/win32/winmsg/wm-sizing
Global Const $WMSZ_LEFT = 1
Global Const $WMSZ_RIGHT = 2
Global Const $WMSZ_TOP = 3
Global Const $WMSZ_TOPLEFT = 4
Global Const $WMSZ_TOPRIGHT = 5
Global Const $WMSZ_BOTTOM = 6
Global Const $WMSZ_BOTTOMLEFT = 7
Global Const $WMSZ_BOTTOMRIGHT = 8
Global Enum $vButton_Tip = 0, $vButton_IconPath, $vButton_IconNumber, $vButton_Command
Global $version="V 0.1 alpha"
Global $dll_icones = @scriptdir & "\Assets\iconset.dll"
Global $array
Global $g_tStruct = DllStructCreate($tagPOINT) ; Create a structure that defines the point to be checked.
Dim $aPos[4]
Dim $idNewsubmenu[40]
Dim $idChangeIcon[40]
;listview notifications
local const $appdatauser = @AppDataDir & "\launchpad"
if not FileExists($appdatauser) then DirCreate($appdatauser)
if not FileExists(@scriptdir & "\Assets") then DirCreate(@scriptdir & "\Assets")

$search = FileFindFirstFile($dll_icones)
if $search = -1 Then
	FileInstall(".\Assets\iconset.dll", @ScriptDir & "\Assets\iconset.dll",1)
EndIf

#cs
    The following 2D array contains the settings that determine the behavior of each "Button"
    namely 4 parameters for each row (for each button);
    [n][0] the tooltip of the button
    [n][1] path of an icon or a file containing icons
    [n][2] the number of the icon (if the previous parameter is a collection)
    [n][3] AutoIt command(s) to be executed directly on button click (or also the name of a function)

#ce

Global const $aStartTools[][] = [ _ ; this arrays could be used as first links in the app
	['Settings', 'SHELL32.dll', 177, 'run("explorer.exe shell:::{D20EA4E1-3957-11d2-A40B-0C5020524153}")','bouton1'], _     ; 'Test()'], _ ; call a function 'Test()'
    ['Windows version', 'winver.exe', 1, 'run("explorer.exe shell:::{BB06C0E4-D293-4f75-8A90-CB05B6477EEE}")','bouton2'], _      ; or "Run('winver.exe')"
    ['This computer', 'netcenter.dll', 6, 'run("explorer.exe shell:::{20D04FE0-3AEA-1069-A2D8-08002B30309D}")','bouton3'], _
    ['Devices and Printers', 'SHELL32.dll', 272, 'run("explorer.exe shell:::{A8A91A66-3A7D-4424-8D24-04E180695C7A}")','bouton4'], _
    ['Folder options', 'SHELL32.dll', 210, 'run("explorer.exe  shell:::{6DFD7C5C-2451-11d3-A299-00C04F8EF6AF}")','bouton5'], _
    ['Command Prompt', @ComSpec, 1, 'Run(@ComSpec)','bouton6'], _
    ['Internet Explorer', @ProgramFilesDir & '\Internet Explorer\iexplore.exe', 1, "Run(@ProgramFilesDir & '\Internet Explorer\iexplore.exe')",'bouton7'], _
    ['Media Player', @ProgramFilesDir & '\Windows media player\wmplayer.exe', 1, "Run(@ProgramFilesDir & '\Windows media player\wmplayer.exe')",'bouton8'], _
    ['File browser', @WindowsDir & '\explorer.exe', 1, "Run(@WindowsDir & '\explorer.exe')",'bouton9'], _
    ['Notepad', @SystemDir & '\notepad.exe', 1, "Run(@SystemDir & '\notepad.exe')",'bouton10'], _
    ['Wordpad', @SystemDir & '\write.exe', 1, "Run(@SystemDir & '\write.exe')",'bouton11'], _
    ['Registry editor', @SystemDir & '\regedit.exe', 1, "ShellExecute('regedit.exe')",'bouton12'], _
    ['Connect to', 'netcenter.dll', 19, 'run("explorer.exe shell:::{38A98528-6CBF-4CA9-8DC0-B1E1D10F7B1B}")','bouton13'], _
    ['Calculator', @SystemDir & '\Calc.exe', 1, "Run(@SystemDir & '\calc.exe')",'bouton14'], _
    ['Control panel', 'control.exe', 1, 'run("explorer.exe shell:::{5399E694-6CE5-4D6C-8FCE-1D8870FDCBA0}")','bouton15'], _
    ['Users manager', @SystemDir & '\Netplwiz.exe', 1, "ShellExecute('Netplwiz.exe')",'bouton16'], _     ; {7A9D77BD-5403-11d2-8785-2E0420524153}
    ['Run', 'SHELL32.dll', 25, 'Run("explorer.exe Shell:::{2559a1f3-21d7-11d4-bdaf-00c04f60b9f0}")','bouton17'], _
    ['Search files', 'SHELL32.dll', 135, 'run("explorer.exe shell:::{9343812e-1c37-4a49-a12e-4b2d810d956b}")','bouton18'], _
    ['On screen Magnifier', @SystemDir & '\Magnify.exe', 1, "ShellExecute('Magnify.exe')",'bouton19'], _
    ['Paint', @SystemDir & '\mspaint.exe', 1, "Run(@SystemDir & '\mspaint.exe')",'bouton20'], _
    ['Remote desktop', @SystemDir & '\mstsc.exe', 1, " Run('mstsc.exe')",'bouton21'], _
    ['Resource monitoring', @SystemDir & '\resmon.exe', 1, "Run('resmon.exe')",'bouton22'], _
    ['Device manager', 'SHELL32.dll', 13, 'Run("explorer.exe Shell:::{74246bfc-4c96-11d0-abef-0020af6b0b7a}")','bouton23'], _
    ['Audio', 'SndVol.exe', 1, 'Run("explorer.exe Shell:::{F2DDFC82-8F12-4CDD-B7DC-D4FE1425AA4D}")','bouton24'], _     ; or 'run(@SystemDir & "\SndVol.exe")']
    ['Task view', 'SHELL32.dll', 133, 'Run("explorer.exe shell:::{3080F90E-D7AD-11D9-BD98-0000947B0257}")','bouton25'], _
    ['Task Manager', @SystemDir & '\taskmgr.exe', 1, 'Send("^+{ESC}")''bouton26'], _     ; "Run(@SystemDir & '\taskmgr.exe')",'bouton1'], _
    ['On Screen Keyboard', 'osk.exe', 1, 'ProcessExists("osc.exe") ? False : ShellExecute("osk.exe")','bouton27'] _     ; <-- ternary example
	]
	;#ce
	;_arraydisplay($aStartTools)
Global $fileConfig = $appdatauser & "\config.ini"
if not FileExists($fileConfig) Then
	$file = FileOpen($fileConfig,1)
	IniWriteSection($fileConfig,"Position","Left=" & @CR & "top=" & @cr & "lignes="& @cr & "colones=")
	IniWriteSection($fileConfig,"COLOR","color=")
	FileClose($file)
Else
	$apos[0] = IniRead($fileconfig,"Position","left",0)
	$apos[1] = iniread($fileconfig,"Position","top",0)
	$apos[2] = iniread($fileconfig,"Position","lignes",16)
	$apos[3] = iniread($fileconfig,"Position","colones",2)
EndIf
;-------------------------  recherche d'un fichier de lien-------------------------------
Global $filelink = $appdatauser & "\launchpad.link"
Global $aTools ; declaration of the array used by the application for all the links
;-------------------------  recherche d'un fichier de lien-------------------------------
local $nb_section
local $ini = $filelink ; Lecture du fichier ini qui contient les raccourcis
$sections = IniReadSectionNames($ini) ; lecture de toutes les sections du fichier ini (.link)
if @error <> 0 then $nb_section = 0
if $nb_section == 0 Then
	;consolewrite($nb_section & @CRLF)
	for $i = 1  to 40
		iniwritesection($ini, "bouton" & $i,"label=Libre" & @CR& "link=" & @CR & "icone=" & @crlf)
	Next
	;Sleep(1000)
	For $i = 0 to UBound($aStartTools) -1
		iniwrite($ini,$aStartTools[$i][4] ,"label", $aStartTools[$i][0])
		iniwrite($ini,$aStartTools[$i][4] ,"link", $aStartTools[$i][3])
		iniwrite($ini,$aStartTools[$i][4] ,"icone", $aStartTools[$i][1] & ","& $aStartTools[$i][2])
	Next
	$sections = IniReadSectionNames($ini)
EndIf
if IsArray($sections) then $nb_section = $sections[0]
if $nb_section < 40 Then
	fileopen($ini,1)
	for $i = $nb_section + 1  to 40
		iniwritesection($ini, "bouton" & $i,"label=Libre" & @CR& "link=" & @CR & "icone=" & @crlf)
	Next
	FileClose($ini)

	$sections = IniReadSectionNames($ini)
EndIf
$nb = $sections[0] ; tableau de toutes les sections
Local $res[$nb+1][5] ; création d'un tableau qui va contenir l'ensemble des liens
For $i =  1 to $nb  ; remplissage du tableau $res
  ;$res[$i][0] = $sections[$i]
  ;consolewrite (@CRLF & $i & @TAB & IniRead($ini,$sections[$i],"label","erreur"))
   $res[$i-1][0] = IniRead($ini,$sections[$i],"label","erreur") ; lecture du fichier et remplissage des ruches du tableau
   $fichier_icone = stringsplit(IniRead($ini,$sections[$i],"icone","erreur"),",")
   if $fichier_icone[0] > 1 Then
   ;_ArrayDisplay($fichier_icone)
		$res[$i-1][1] = $fichier_icone[1]
		$res[$i-1][2] = $fichier_icone[2]

   Else
		$res[$i-1][1] = IniRead($ini,$sections[$i],"icone","erreur")
   EndIf
   $res[$i-1][3] = IniRead($ini,$sections[$i],"link","erreur")
   $res[$i-1][4] = $sections[$i]
Next

$aTools = $res

; Show desktop       {3080F90D-D7AD-11D9-BD98-0000947B0257}
; Desktop Background {ED834ED6-4B5A-4bfe-8F11-A626DCB6A921}
; IE internet option {A3DD4F92-658A-410F-84FD-6FBBBEF2FFFE}
;       ['Notes', 'StikyNot.exe', 1, "ShellExecute('StikyNot')"], _
;    ['... if Notepad is running' & @CRLF & 'Send F5 to it', 'SHELL32.dll', 167, ' WinExists("[CLASS:Notepad]") ? ControlSend("[CLASS:Notepad]", "", "", "{F5}") : MsgBox(16, ":(", "Notepad not found", 2)'] _     ; Check if Notepad is currently running
;_ArrayDisplay($aTools)
Opt("TrayMenuMode", 3)

Global $iStep = 40 ; button size

Global $iNrPerLine
if $apos[3] = "" Then
	$iNrPerLine = 2
Else
	$iNrPerLine = $apos[3]
EndIf

Global $iNrOfLines
if $apos[2] = "" then
	$iNrOfLines = 20;Ceiling(UBound($aTools) / $iNrPerLine)
Else
	$iNrOfLines = $aPos[2]
EndIf

if $apos[0] = "" then $apos[0] = @DesktopWidth - $iStep * 3
if $apos[1] = "" then $apos[1] = @DesktopHeight / 20

Global $GUI = GUICreate('LaunchPad', 10, 10, $apos[0] , $apos[1] ,$WS_THICKFRAME + $WS_EX_ACCEPTFILES, BitOR($WS_EX_TOOLWINDOW, $WS_EX_TOPMOST))
GUICtrlSetBkColor(-1,0xFFFFFF)
$GuiContextMenu = GUICtrlCreateContextMenu()
$idSavePosition = GUICtrlCreateMenuItem("Sauver la position", $GuiContextMenu)
$idReinitializePosition = GUICtrlCreateMenuItem("Reinitialiser la position", $GuiContextMenu)
GUICtrlCreateMenuItem("", $GuiContextMenu)
$idReinstallLinkFile = GUICtrlCreateMenuItem("Reinitialisation complete", $GuiContextMenu)

Global $aMyMatrix = _GuiControlPanel("Button", $iNrPerLine, $iNrOfLines, $iStep, $iStep, BitOR(0x40, 0x1000), -1, 0, 0, 0, 0, 0, 0, False, "")
;_ArrayDisplay($idNewsubmenu)
Global $iPreviousX = ($aMyMatrix[0])[1], $iPreviousY = ($aMyMatrix[0])[2]

For $i = 1 To UBound($aMyMatrix)-1
  GUICtrlSetResizing($aMyMatrix[$i], $GUI_DOCKALL)   ; (2+32+256+512) so the control will not move during resizing
  If $i <= UBound($aTools) Then
	If $aTools[$i-1][$vButton_IconPath] = "" Then
		GUICtrlSetImage($aMyMatrix[$i], $dll_icones,80)
		GUICtrlSetTip(-1,"Glisser/déposer de (fichier/dossier/raccourcis) pour créer un nouveau lien")
	Else
		GUICtrlSetImage($aMyMatrix[$i], $aTools[$i-1 ][$vButton_IconPath], $aTools[$i-1 ][$vButton_IconNumber])
		GUICtrlSetTip($aMyMatrix[$i], $aTools[$i -1][$vButton_Command],$aTools[$i-1 ][$vButton_Tip] )
	EndIf
  EndIf
Next

_WinSetClientSize($GUI, ($aMyMatrix[0])[11], ($aMyMatrix[0])[12]) ; thanks to KaFu

DragDropEvent_Startup()
GUISetState(@SW_SHOW, $GUI)

;GUISetState(@SW_SHOW, $GuiIcon)

; https://devblogs.microsoft.com/oldnewthing/20110218-00/?p=11453
GUIRegisterMsg($WM_NCHITTEST, "WM_NCHITTEST")
GUIRegisterMsg($WM_SIZING, "WM_SIZING")
GUIRegisterMsg($WM_NOTIFY, WM_NOTIFY)

TrayCreateItem("LAUNCHPAD")
TrayCreateItem("") ; Create a separator line.
Global $idHide = TrayCreateItem("Hide")
Global $idShow = TrayCreateItem("Show")
TrayCreateItem("") ; Create a separator line.
Global $idExit = TrayCreateItem("Exit")
Global $idPos = TrayCreateItem("Save windows Position")
Global $idDfault = TrayCreateItem("Reinitialiser la position")
Global $idOpenFileLink = TrayCreateItem("ouvrir le dossier des liens")
TraySetState($TRAY_ICONSTATE_SHOW)

_MainLoop()

Func _MainLoop()
  Local $iDeltaX, $iDeltaY, $row, $col, $left, $top
  Global $hTimer = TimerInit()
  Local $aPos
  DragDropEvent_Register($GUI)

  GUIRegisterMsg($WM_DRAGENTER, "OnDragDrop")
  GUIRegisterMsg($WM_DRAGOVER, "OnDragDrop")
  GUIRegisterMsg($WM_DRAGLEAVE, "OnDragDrop")
  GUIRegisterMsg($WM_DROP, "OnDragDrop")

  While 1
    Sleep (10)

    If TimerDiff($hTimer) > 5000 Then
      GUISetState(@SW_HIDE, $GUI)
      $hTimer = 0
    EndIf
    $aPos = MouseGetPos ()
    If $aPos[0] = 0 Or $aPos[1] = 0 Or $aPos[0] = @DesktopWidth-1 Or $aPos[1] = @DesktopHeight-1 Then
      $hTimer = TimerInit ()
      GUISetState(@SW_SHOW, $GUI)
    EndIf
    $Msg = GUIGetMsg()
	for $i = 0 to ubound($idNewsubmenu)-1
		if $Msg = $idNewsubmenu[$i] and $aTools[$i][3] <> "" Then
			$sMsg = Msgbox(4,"","Etes vous sur de vouloir supprimer ce raccourcis ?")
			if $sMsg == 6 Then
				iniwrite($filelink,$aTools[$i][4],"label","Libre")
				iniwrite($filelink,$aTools[$i][4],"link","")
				iniwrite($filelink,$aTools[$i][4],"icone","")

				$atools[$i][0] = "Libre"
				$atools[$i][1] = ""
				$atools[$i][2] = ""
				$atools[$i][3] = ""
				GUICtrlSetImage($aMyMatrix[$i+1], $dll_icones,80)
				ToolTip("")
			EndIf
		EndIf
	Next

	For $i = 0 to ubound($idChangeIcon) - 1
		If $Msg = $idChangeIcon[$i] and $atools[$i][3] <> "" Then
			$aRet = _PickIconDlg($dll_icones)
			If Not @error Then
				$aTools[$i][1] = $aRet[0]
				$aTools[$i][2] = $aRet[1]
				IniWrite($filelink,$aTools[$i][4],"icone",$aRet[0] & "," & $aRet[1])
				GUICtrlSetImage($aMyMatrix[$i+1], $aTools[$i][1], $aTools[$i][2])
			EndIf
		EndIf
	Next
    Switch $Msg
		Case $GUI_EVENT_CLOSE
			GUISetState(@SW_HIDE)
			$hTimer = 0
		Case $aMyMatrix[1] To $aMyMatrix[40]
			$hTimer = TimerInit ()
			For $i = 1 To UBound($aMyMatrix) - 1
				If $Msg = $aMyMatrix[$i] Then
					If $i <= UBound($aTools) and $aTools[$i-1][3] <> "" Then
						if StringInStr($atools[$i-1][3],"run(") then
							$dummy = Execute($aTools[$i - 1][3])
						Else
							$dummy = ShellExecute($aTools[$i - 1][3])
						EndIf
					EndIf
				EndIf
			Next
		case $idSavePosition
			$apos = WinGetPos($gui)
			iniwrite($fileConfig,"Position","left",$apos[0])
			iniwrite($fileConfig,"Position","top",$apos[1])
			IniWrite($fileconfig,"Position","lignes",($aMyMatrix[0])[2])
			IniWrite($fileconfig,"Position","colones",($aMyMatrix[0])[1])
			msgbox(0,"","Position sauvegardée avec succès",3)

		Case $idReinitializePosition
			iniwrite($fileconfig,"position","left","")
			iniwrite($fileConfig,"Position","top","")
			IniWrite($fileconfig,"Position","lignes","")
			IniWrite($fileconfig,"Position","colones","")
			_RestartProgram()

		Case $idReinstallLinkFile
			local $sMsg  = Msgbox(4,"","Etes vous certain de vouloir reinitialiser tous les liens ?")
			if $sMsg == 6 Then
				FileDelete($filelink)
				_RestartProgram()
			EndIf

	EndSwitch
	Switch TrayGetMsg()
		Case $idShow
			$hTimer = TimerInit ()
			GUISetState(@SW_SHOW)
		Case $idHide
			GUISetState(@SW_HIDE)
			$hTimer = 0
        Case $idExit
            ExitLoop
		Case $idPos
			$apos = WinGetPos($gui)
			iniwrite($fileConfig,"Position","left",$apos[0])
			iniwrite($fileConfig,"Position","top",$apos[1])
			IniWrite($fileconfig,"Position","lignes",($aMyMatrix[0])[2])
			IniWrite($fileconfig,"Position","colones",($aMyMatrix[0])[1])
		case $idDfault
			iniwrite($fileconfig,"position","left","")
			iniwrite($fileConfig,"Position","top","")
			IniWrite($fileconfig,"Position","lignes","")
			IniWrite($fileconfig,"Position","colones","")
			_RestartProgram()
		Case $idOpenFileLink
			ShellExecute(@AppDataDir & "\Launchpad")

    EndSwitch

    ; check if any size has changed
    If $iPreviousX <> ($aMyMatrix[0])[1] Or $iPreviousY <> ($aMyMatrix[0])[2] Then
      ; calculate the variations
      $iDeltaX = Abs($iPreviousX - ($aMyMatrix[0])[1])
      $iDeltaY = Abs($iPreviousY - ($aMyMatrix[0])[2])

      ; if both dimensions changed at the same time, the largest variation prevails over the other
      If $iDeltaX >= $iDeltaY Then       ; keep the new number of columns
        ; calculate and set the correct number of lines accordingly
        _SubArraySet($aMyMatrix[0], 2, Ceiling((UBound($aMyMatrix) - 1) / ($aMyMatrix[0])[1]))
      Else       ; otherwise keep the new number of rows
        ; calculate and set the correct number of columns accordingly
        _SubArraySet($aMyMatrix[0], 1, Ceiling((UBound($aMyMatrix) - 1) / ($aMyMatrix[0])[2]))
      EndIf

      ; set client area new sizes
      _WinSetClientSize($GUI, ($aMyMatrix[0])[1] * $iStep, ($aMyMatrix[0])[2] * $iStep)

      ; remember the new panel settings
      $iPreviousX = ($aMyMatrix[0])[1]
      $iPreviousY = ($aMyMatrix[0])[2]

      ; rearrange the controls inside the panel
      For $i = 0 To UBound($aMyMatrix) - 2
        ; coordinates 1 based
        $col = Mod($i, $iPreviousX) + 1         ; Horizontal position within the grid (column)
        $row = Int($i / $iPreviousX) + 1         ; Vertical position within the grid (row number)
        $left = ($aMyMatrix[0])[5] + (((($aMyMatrix[0])[3] + ($aMyMatrix[0])[9]) * $col) - ($aMyMatrix[0])[9]) - ($aMyMatrix[0])[3] + ($aMyMatrix[0])[7]
        $top = ($aMyMatrix[0])[6] + (((($aMyMatrix[0])[4] + ($aMyMatrix[0])[10]) * $row) - ($aMyMatrix[0])[10]) - ($aMyMatrix[0])[4] + ($aMyMatrix[0])[8]
        GUICtrlSetPos($aMyMatrix[$i + 1], $left, $top)
	  Next
    EndIf
  WEnd
EndFunc   ;==>_MainLoop

Func WM_NOTIFY ($hwnd, $iMsg, $iwParam, $ilParam)
  If $hwnd = $GUI Then $hTimer = TimerInit ()
  Return $GUI_RUNDEFMSG
EndFunc

; Allow/Disallow specific borders resizing
; thanks to Danyfirex
;           ---------
; https://www.autoitscript.com/forum/topic/201464-partially-resizable-window-how-solved-by-danyfirex-%F0%9F%91%8D/?do=findComment&comment=1445748
Func WM_NCHITTEST($hwnd, $iMsg, $iwParam, $ilParam)
  If $hwnd = $GUI Then
    $hTimer = TimerInit()
    Local $iRet = _WinAPI_DefWindowProc($hwnd, $iMsg, $iwParam, $ilParam)
    ; https://docs.microsoft.com/en-us/windows/win32/inputdev/wm-nchittest
    If $iRet = $HTBOTTOM Or $iRet = $HTRIGHT Or $iRet = $HTBOTTOMRIGHT Or $iRet = $HTCAPTION Or $iRet = $HTCLOSE Then
      Return $iRet       ; default process of border resizing
    Else     ; resizing not allowed
      Return $HTCLIENT       ; do like if cursor is in the client area
    EndIf
  EndIf
  Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_NCHITTEST

; controls and process resizing operations in real time
; thanks to mikell
;           ------
; https://www.autoitscript.com/forum/topic/201464-partially-resizable-window-how-solved-by-danyfirex-%F0%9F%91%8D/?do=findComment&comment=1445754
Func WM_SIZING($hwnd, $iMsg, $wparam, $lparam)
  ; https://docs.microsoft.com/en-us/windows/win32/winmsg/wm-sizing
  Local $iCols = ($aMyMatrix[0])[1]
  Local $iRows = ($aMyMatrix[0])[2]
  Local $xClientSizeNew, $yClientSizeNew

  #cs $wparam
      The edge of the window that is being sized.

      $lparam
      A pointer to a RECT structure with the screen coordinates of the drag rectangle.
      To change the size or position of the drag rectangle, an application must change the members of this structure.

      Return value
      Type: LRESULT

  #ce $wparam

  $aPos = WinGetPos($GUI)
  #cs Success : a 4 - element array containing the following information :
      $aArray[0] = X position
      $aArray[1] = Y position
      $aArray[2] = Width
  #ce Success : a 4 - element array containing the following information :

  $aPos2 = WinGetClientSize($GUI)
  #cs Success: a 2-element array containing the following information:
      $aArray[0] = Width of window's client area
  #ce Success: a 2-element array containing the following information:

  ; https://docs.microsoft.com/en-us/previous-versions//dd162897(v=vs.85)
  Local $sRect = DllStructCreate("Int[4]", $lparam)   ; outer dimensions (includes borders)
  Local $left = DllStructGetData($sRect, 1, 1)
  Local $top = DllStructGetData($sRect, 1, 2)
  Local $Right = DllStructGetData($sRect, 1, 3)
  Local $bottom = DllStructGetData($sRect, 1, 4)

  ; border width
  Local $iEdgeWidth = ($aPos[2] - $aPos2[0]) / 2
  Local $iHeadHeigth = $aPos[3] - $aPos2[1] - $iEdgeWidth * 2

  Local $aEdges[2]
  $aEdges[0] = $aPos[2] - $aPos2[0]   ; x
  $aEdges[1] = $aPos[3] - $aPos2[1]   ; y

  $xClientSizeNew = $Right - $left - $aEdges[0]
  $xClientSizeNew = Round($xClientSizeNew / $iStep) * $iStep

  $yClientSizeNew = $bottom - $top - $aEdges[1]
  $yClientSizeNew = Round($yClientSizeNew / $iStep) * $iStep

  Switch $wparam
    Case $WMSZ_RIGHT
      ; calculate the new position of the right border
      DllStructSetData($sRect, 1, $left + $xClientSizeNew + $aEdges[0], 3)
    Case $WMSZ_BOTTOM
      ; calculate the new position of the bottom border
      DllStructSetData($sRect, 1, $top + $yClientSizeNew + $aEdges[1], 4)
    Case $WMSZ_BOTTOMRIGHT
      ; calculate the new position of both borders
      DllStructSetData($sRect, 1, $left + $xClientSizeNew + $aEdges[0], 3)
      DllStructSetData($sRect, 1, $top + $yClientSizeNew + $aEdges[1], 4)
  EndSwitch

  #cs If DllStructGetData($sRect, 1, 3) > @DesktopWidth Then ; $Right
      DllStructSetData($sRect, 1, DllStructGetData($sRect, 1, 3) - $iStep, 3)
      $xClientSizeNew -= $iStep
      EndIf

      If DllStructGetData($sRect, 1, 4) > @DesktopHeight Then ; $bottom
      DllStructSetData($sRect, 1, DllStructGetData($sRect, 1, 4), 4)
      $yClientSizeNew -= $iStep
  #ce If DllStructGetData($sRect, 1, 3) > @DesktopWidth Then ; $Right

  ; check if number of rows has changed
  If $iRows <> $yClientSizeNew / $iStep Then
    _SubArraySet($aMyMatrix[0], 2, $yClientSizeNew / $iStep)
  EndIf

  ; check if number of columns has changed
  If $iCols <> $xClientSizeNew / $iStep Then
    _SubArraySet($aMyMatrix[0], 1, $xClientSizeNew / $iStep)
  EndIf
  ;consolewrite(@crlf & "rows =" & @tab & $iRows & @tab & "Col =" & @TAB & $iCols & @tab  & "$xClientSizeNew = " & @tab & $xClientSizeNew & @tab & "$yClientSizeNew =" & @tab & $yClientSizeNew )
  Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_SIZING

; set client area new sizes
; thanks to KaFu
;           ----
; https://www.autoitscript.com/forum/topic/201524-guicreate-and-wingetclientsize-mismatch/?do=findComment&comment=1446141
Func _WinSetClientSize($hwnd, $iW, $iH)
  Local $aWinPos = WinGetPos($hwnd)
  Local $sRect = DllStructCreate("int;int;int;int;")
  DllStructSetData($sRect, 3, $iW)
  DllStructSetData($sRect, 4, $iH)
  _WinAPI_AdjustWindowRectEx($sRect, _WinAPI_GetWindowLong($hwnd, $GWL_STYLE), _WinAPI_GetWindowLong($hwnd, $GWL_EXSTYLE))
  WinMove($hwnd, "", $aWinPos[0], $aWinPos[1], $aWinPos[2] + (DllStructGetData($sRect, 3) - $aWinPos[2]) - DllStructGetData($sRect, 1), $aWinPos[3] + (DllStructGetData($sRect, 4) - $aWinPos[3]) - DllStructGetData($sRect, 2))
EndFunc   ;==>_WinSetClientSize

;
; #FUNCTION# ====================================================================================================================
; Name...........: _GuiControlPanel
; Description ...: Creates a rectangular panel with adequate size to contain the required amount of controls
;                  and then fills it with the same controls by placing them according to the parameters
; Syntax.........: _GuiControlPanel($ControlType, $nrPerLine, $nrOfLines, $ctrlWidth, $ctrlHeight, $style, $exStyle, $xPos = 0, $yPos = 0, $xBorder, $yBorder, $xSpace = 1, $ySpace = 1, $Group = false, , $sGrpTitle = "")
; Parameters ....: $ControlType  - Type of controls to be generated ("Button"; "Text"; .....
;                  $nrPerLine  - Nr. of controls per line in the matrix
;                  $nrOfLines - Nr. of lines in the matrix
;                  $ctrlWidth - Width of each control
;                  $ctrlHeight - Height of each control
;                  $Style - Defines the style of the control
;                  $exStyle - Defines the extended style of the control
;                  $xPanelPos - x Position of panel in GUI
;                  $yPanelPos - y Position of panel in GUI
;                  $xBorder - distance from lateral panel's borders to the matrix (width of left and right margin) default = 0
;                  $yBorder - distance from upper and lower panel's borders to the matrix (width of upper and lower margin) default = 0
;                  $xSpace - horizontal distance between the controls
;                  $ySpace - vertical distance between the controls
;                  $Group - if you want to group the controls (true or false)
;                  $sGrpTitle - title of the group (ignored if above is false)
; Return values .: an 1 based 1d array containing references to each control
;                  element [0] contains an 1d array containing various parameters about the panel
; Author ........: Gianni Addiego (Chimp)
; Modified.......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......:
; ===============================================================================================================================

Func _GuiControlPanel($ControlType, $nrPerLine, $nrOfLines, $ctrlWidth, $ctrlHeight, $style = -1, $exStyle = -1, $xPanelPos = 0, $yPanelPos = 0, $xBorder = 0, $yBorder = 0, $xSpace = 1, $ySpace = 1, $Group = False, $sGrpTitle = "")

  Local Static $sAllowedControls = "|Label|Input|Edit|Button|CheckBox|Radio|List|Combo|Pic|Icon|Graphic|"
  If Not StringInStr($sAllowedControls, '|' & $ControlType & '|') Then Return SetError(1, 0, "Unkown control")

  Local $PanelWidth = (($ctrlWidth + $xSpace) * $nrPerLine) - $xSpace + ($xBorder * 2)
  Local $PanelHeight = (($ctrlHeight + $ySpace) * $nrOfLines) - $ySpace + ($yBorder * 2)

  Local $hGroup

  If $Group Then
    If $sGrpTitle = "" Then
      $xPanelPos += 1
      $yPanelPos += 1
      $hGroup = GUICtrlCreateGroup("", $xPanelPos - 1, $yPanelPos - 7, $PanelWidth + 2, $PanelHeight + 8)

	 GUICtrlSetColor(-1, 0xFFFFFF)

    Else
      $xPanelPos += 1
      $yPanelPos += 15
      $hGroup = GUICtrlCreateGroup($sGrpTitle, $xPanelPos - 1, $yPanelPos - 15, $PanelWidth + 2, $PanelHeight + 16)
	  GUICtrlSetColor(-1, 0xFFFFFF)
    EndIf
  EndIf

  ; create the controls
  Local $aGuiGridCtrls[$nrPerLine * $nrOfLines + 1]
  Local $aPanelParams[14] = [ _
      $ControlType, $nrPerLine, $nrOfLines, $ctrlWidth, $ctrlHeight, _
      $xPanelPos, $yPanelPos, $xBorder, $yBorder, $xSpace, $ySpace, $PanelWidth, $PanelHeight, $hGroup]

  For $i = 0 To $nrPerLine * $nrOfLines - 1
    ; coordinates 1 based
    $col = Mod($i, $nrPerLine) + 1     ; Horizontal position within the grid (column)
    $row = Int($i / $nrPerLine) + 1     ;  Vertical position within the grid (row)
    $left = $xPanelPos + ((($ctrlWidth + $xSpace) * $col) - $xSpace) - $ctrlWidth + $xBorder
    $top = $yPanelPos + ((($ctrlHeight + $ySpace) * $row) - $ySpace) - $ctrlHeight + $yBorder
    $text = $i + 1     ; "*" ; "." ; "(*)"
    ; create the control(s)
	if $i >= ubound($aTools)-1 Then
		ExitLoop
	Else
		$aGuiGridCtrls[$i + 1] = Execute("GUICtrlCreate" & $ControlType & "($text, $left, $top, $ctrlWidth, $ctrlHeight, $style, $exStyle)")
		;Global $g_tile_notif = GUICtrlCreateDummy()
		$idContextmenu = GUICtrlCreateContextMenu($aGuiGridCtrls[$i + 1])
		$idNewsubmenu[$i] = GUICtrlCreateMenuItem("Supprimer", $idContextmenu)
		$idChangeIcon[$i] = GUICtrlCreateMenuItem("Changer Icone", $idContextmenu)
		if $aTools[$i][1] = "" then
			;GUICtrlSetColor(-1, 0xFFFFFF)
			GUICtrlSetBkColor(-1,$GUI_BKCOLOR_TRANSPARENT)
			;WinSetTrans(-1,Default,100)
		EndIf
	EndIf
  Next

  If $Group Then GUICtrlCreateGroup("", -99, -99, 1, 1)   ; close group
  $aGuiGridCtrls[0] = $aPanelParams
  Return $aGuiGridCtrls
EndFunc   ;==>_GuiControlPanel

; writes a value to an element of an array embedded in another array
Func _SubArraySet(ByRef $aSubArray, $iElement, $vValue)
  $aSubArray[$iElement] = $vValue
EndFunc   ;==>_SubArraySet

;Func _WinAPI_AdjustWindowRectEx(ByRef $tRECT, $iStyle, $iExStyle = 0, $bMenu = False)
;	Local $aRet = DllCall('user32.dll', 'bool', 'AdjustWindowRectEx', 'struct*', $tRECT, 'dword', $iStyle, 'bool', $bMenu, _
;			'dword', $iExStyle)
;	If @error Then Return SetError(@error, @extended, False)
;	; If Not $aRet[0] Then Return SetError(1000, 0, 0)
;
;	Return $aRet[0]
;EndFunc   ;==>_WinAPI_AdjustWindowRectEx

Func OnDragDrop($hWnd, $Msg, $wParam, $lParam)
	;consolewrite($hWnd & @tab & $Msg & @crlf)
    Static $DropAccept
    Switch $Msg
        Case $WM_DRAGENTER, $WM_DROP
            ToolTip("")
			Select
                Case DragDropEvent_IsFile($wParam)
                    If $Msg = $WM_DROP Then
						Position() ;  needed for _WinAPI_WindowFromPoint
						Local $mouseId = _WinAPI_WindowFromPoint($g_tStruct) ; Find de hwnd of the control because GUIGetCursorInfo doesn't work with text file
						$mouseId = _WinAPI_GetDlgCtrlID($mouseId) ; gets the Id of hwnd
                        Local $FileList = DragDropEvent_GetFile($wParam)
						consolewrite(@CRLF & $FileList)
						Local $section , $latools
						Local $aDetails = FileGetShortcut($FileList)
						if IsArray($aDetails) Then
							$Program = $aDetails[0]
						EndIf
						Local $ProposeLink = stringsplit($FileList,"\")
						$ProposeLink = StringRegExpReplace($ProposeLink[UBound($ProposeLink)-1], '(.*)\..*', "$1")
						For $i = 1 To UBound($aMyMatrix) - 1
							If $mouseId = $aMyMatrix[$i] and $aTools[$i - 1][3] = "" Then
								$section = $aTools[$i -1][4]
								$latools = $i - 1
								consolewrite(@crlf & $section & @TAB & $fileList)
								iniwrite($filelink,$section,"link",$FileList)
								ExitLoop
							Elseif $mouseId = $aMyMatrix[$i] and  $aTools[$i -1][3] <> "" Then
								$section = $aTools[$i -1][4]
								$latools = $i -1
								Local $question  = msgbox(4,"", "Etes vous certain de vouloir écraser le raccourcis existant ?")
								if $question == 6 then
									consolewrite(@crlf & $section & @TAB & $filelist)
									iniwrite($filelink,$section,"link",$FileList)
								else
									Return
								EndIf
								ExitLoop
							EndIf
						Next
						Local $reponse = InputBox("Nom de l'icone","Donnez un titre au raccourcis",$ProposeLink); quel label aura le raccourcis
						if @error == 1 or $reponse = "" Then
							iniwrite($filelink,$section,"link","")
							Iniwrite($filelink,$section,"label","")
							Return
						EndIf
						Iniwrite($filelink,$section,"label",$reponse); tout est ok on peut ecrire la valeur dans le fichier de config
						local $sExt = StringRegExpReplace($FileList, "^.*\.", "") ; extraction de son extension
						Switch $sExt
							case "doc" , "docx" , "odt"
								IniWrite($filelink,$section,"icone",$dll_icones &",436")
							case "xls" , "xlsx" , "ods"
								IniWrite($filelink,$section,"icone",$dll_icones &",441")
							case "pdf"
								IniWrite($filelink,$section,"icone",$dll_icones &",400")
							Case "ppt" , "pptx" , "odp"
								IniWrite($filelink,$section,"icone",$dll_icones &",431")
							Case "txt" , "rtf"
								IniWrite($filelink,$section,"icone",$dll_icones &",406")
							Case Else
								if $sExt = "lnk" then $FileList = $Program
								if _WinAPI_ExtractIconEx( $FileList,-1,0,0,0) > 0 Then ; permet de tester si le fichier possède une ou plusiers icone(s)
									Local $aIcon[3] = [64, 32, 16]
									For $i = 0 To UBound($aIcon) - 1
										$aIcon[$i] = _WinAPI_Create32BitHICON(_WinAPI_ShellExtractIcon($FileList,0, $aIcon[$i], $aIcon[$i]), 1)
									Next
									_WinAPI_SaveHICONToFile(@ScriptDir & "\Assets\" & $reponse & ".ico", $aIcon)
									For $i = 0 To UBound($aIcon) - 1
										_WinAPI_DestroyIcon($aIcon[$i])
									Next
									IniWrite($filelink,$section,"icone",@ScriptDir & "\Assets\" & $reponse & ".ico"); si oui écriture dans le fichier
								Else
									$aRet = _PickIconDlg($dll_icones)
									If Not @error Then
										IniWrite($filelink,$section,"icone",$aRet[0] & "," & $aRet[1])
									Else
										iniwrite($filelink,$section,"link","")
										Iniwrite($filelink,$section,"label","")
										Return
									EndIf
								EndIf
						EndSwitch
								$fichier_icone = stringsplit(IniRead($filelink,$aTools[$latools][4],"icone","erreur"),",")
								if $fichier_icone[0] > 1 Then
									$aTools[$latools][1] = $fichier_icone[1]
									$aTools[$latools][2] = $fichier_icone[2]
								Else
									$aTools[$latools][1] = IniRead($filelink,$aTools[$latools][4],"icone","erreur")
									$aTools[$latools][2] = "," & Number("0")
								EndIf
								$aTools[$latools][0] = $reponse
								$aTools[$latools][3] = IniRead($filelink,$section,"link","erreur")
								for $i = 0 to 4
									consolewrite (@CRLF & "colone " & $i & @TAB & $aTools[$latools][$i])
								Next
								;GUICtrlSetBkColor($mouseId,$GUI_BKCOLOR_TRANSPARENT)
								GUICtrlSetImage($mouseId, $aTools[$latools][1], $aTools[$latools][2])
								GUICtrlSetTip($mouseId, $aTools[$latools][3],$aTools[$latools][0])


					EndIf

                    $DropAccept = $DROPEFFECT_COPY

                Case DragDropEvent_IsText($wParam)
                    If $Msg = $WM_DROP Then
						Position() ;  needed for _WinAPI_WindowFromPoint to get the right hwnd
						Local $mouseId = _WinAPI_WindowFromPoint($g_tStruct) ; Find de hwnd of the control because GUIGetCursorInfo doesn't work
						$mouseId = _WinAPI_GetDlgCtrlID($mouseId) ; gets the Id of hwnd
						$hyperlink = DragDropEvent_GetText($wParam)
						Local $section , $latools
						For $i = 1 To UBound($aMyMatrix) - 1
							;consolewrite($i & " " )
							If $mouseId = $aMyMatrix[$i] and $aTools[$i-1][3] = "" Then
								$section = $aTools[$i-1][4]
								$latools = $i -1
								iniwrite($filelink,$section,"link",$hyperlink)
								ExitLoop
							Elseif $mouseId = $aMyMatrix[$i] and  $aTools[$i-1][3] <> "" Then
								$section = $aTools[$i-1][4]
								$latools = $i -1
								Local $question  = msgbox(4,"", "Etes vous certain de vouloir écraser le raccourcis existant ?")
								if $question == 6 then
									iniwrite($filelink,$section,"link",$hyperlink)
								else
									Return
								EndIf
								ExitLoop
							EndIf
						Next

						$reponse = InputBox("Nom du lien","Donnez un titre !",StringTrimLeft($hyperlink,7))
						if @error == 1 or $reponse = "" Then Return
						Iniwrite($filelink,$section,"label",$reponse)
						$aRet = _PickIconDlg($dll_icones)
						If Not @error Then
							IniWrite($filelink,$section,"icone",$aRet[0] & "," & $aRet[1])
						Else
							iniwrite($filelink,$section,"link","")
							Iniwrite($filelink,$section,"label","")
							Return
						EndIf
						$fichier_icone = stringsplit(IniRead($filelink,$aTools[$latools][4],"icone","erreur"),",")
						if $fichier_icone[0] > 1 Then
							$aTools[$latools][1] = $fichier_icone[1]
							$aTools[$latools][2] = $fichier_icone[2]
						Else
							$aTools[$latools][1] = IniRead($filelink,$aTools[$latools][4],"icone","erreur")
							$aTools[$latools][2] = "," & Number("0")
						EndIf
						$aTools[$latools][0] = $reponse
						$aTools[$latools][3] = $hyperlink
						for $i = 0 to 4
							consolewrite (@CRLF & "colone " & $i & @TAB & $aTools[$latools][$i])
						Next
						;GUICtrlSetBkColor($mouseId,$GUI_BKCOLOR_TRANSPARENT)
						GUICtrlSetImage($mouseId, $aTools[$latools][1], $aTools[$latools][2])
						GUICtrlSetTip($mouseId, $aTools[$latools][3],$aTools[$latools][0])
                    EndIf
                    $DropAccept = $DROPEFFECT_COPY

                Case Else
                    $DropAccept = $DROPEFFECT_NONE

            EndSelect
            Return $DropAccept

        Case $WM_DRAGOVER

            Return $DropAccept

        Case $WM_DRAGLEAVE
            ToolTip("")

    EndSwitch
EndFunc

Func Test()
  MsgBox(0, 0, ":)", 1)
EndFunc   ;==>Test


Func _Check_LabelForbidden($string)
	if $string = "Libre" then $string = 1
	Return  $string
EndFunc

Func _PickIconDlg($sFileName, $nIconIndex=0, $hWnd=0)
    Local $nRet, $aRetArr[2]

    $nRet = DllCall("shell32.dll", "int", "PickIconDlg", _
        "hwnd", $hWnd, _
        "wstr", $sFileName, "int", 1000, "int*", $nIconIndex)

    If Not $nRet[0] Then Return SetError(1, 0, -1)

    $aRetArr[0] = $nRet[2]
    $aRetArr[1] = $nRet[4] + 1

    Return $aRetArr
EndFunc

Func Position()
    DllStructSetData($g_tStruct, "x", MouseGetPos(0))
    DllStructSetData($g_tStruct, "y", MouseGetPos(1))
EndFunc   ;==>Position



#Region --- Restart Program ---
    Func _RestartProgram()
        If @Compiled = 1 Then
            Run(FileGetShortName(@ScriptFullPath))
        Else
            Run(FileGetShortName(@AutoItExe) & " " & FileGetShortName(@ScriptFullPath))
        EndIf
        Exit
    EndFunc; ==> _RestartProgram
#EndRegion --- Restart Program ---
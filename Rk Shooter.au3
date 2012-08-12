;~ #AutoIt3Wrapper_AU3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6
#AutoIt3Wrapper_UseX64=n
#AutoIt3Wrapper_Icon=Camera.ico
#AutoIt3Wrapper_Res_Description=Makes screenshot of any visible window, a selectable region on the desktop or from any web site, creates a video stream of your screen
#AutoIt3Wrapper_Res_Fileversion=0.0.2
#AutoIt3Wrapper_Res_Field=CompanyName|Chagolchana Production
#AutoIt3Wrapper_Res_Field=ProductName|Rk Screencapture
#AutoIt3Wrapper_Res_Field=ProductVersion|%AutoItVer%
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_Field=Coded by|Rabimba
#AutoIt3Wrapper_Res_Field=URL|http://www.rabimba.com
#AutoIt3Wrapper_Res_Field=Build|2011-11-19 Final
#AutoIt3Wrapper_Res_Field=Compile Date|%longdate% %time%
#AutoIt3Wrapper_Run_Obfuscator=y
#Obfuscator_Parameters=/sf /sv /om /cs=0 /cn=0
#AutoIt3Wrapper_UseUpx=n
#AutoIt3Wrapper_Run_After=del /f /q "%scriptdir%\%scriptfile%_Obfuscated.au3"
#AutoIt3Wrapper_Run_After=..\..\..\ResourceHacker\ResHacker.exe -delete %out%, %out%, ICON, 1,
#AutoIt3Wrapper_Run_After=..\..\..\ResourceHacker\ResHacker.exe -delete %out%, %out%, ICON, 2,
#AutoIt3Wrapper_Run_After=..\..\..\ResourceHacker\ResHacker.exe -delete %out%, %out%, ICON, 3,
#AutoIt3Wrapper_Run_After=..\..\..\ResourceHacker\ResHacker.exe -delete %out%, %out%, MENU, 166,
#AutoIt3Wrapper_Run_After=..\..\..\ResourceHacker\ResHacker.exe -delete %out%, %out%, DIALOG, 1000,
#AutoIt3Wrapper_Run_After=..\..\..\ResourceHacker\ResHacker.exe -delete %out%, %out%, ICONGROUP, 162,
#AutoIt3Wrapper_Run_After=..\..\..\ResourceHacker\ResHacker.exe -delete %out%, %out%, ICONGROUP, 164,
#AutoIt3Wrapper_Run_After=..\..\..\ResourceHacker\ResHacker.exe -delete %out%, %out%, ICONGROUP, 169,
#AutoIt3Wrapper_Run_After=upx.exe --best --lzma "%out%"
;~ #AutoIt3Wrapper_Run_After=upx.exe --ultra-brute --crp-ms=999999 --all-methods --all-filters "%out%"

Break(0)
#NoTrayIcon
#include <ButtonConstants.au3>
#include <Clipboard.au3>
#include <Constants.au3>
#include <Date.au3>
#include <File.au3>
#include <GUIButton.au3>
#include <GUIListBox.au3>
#include <GUIListView.au3>
#include <GUIMenu.au3>
#include <GUIScrollBars.au3>
#include <GUISlider.au3>
#include <INet.au3>
#include <Misc.au3>
#include <ScreenCapture.au3>
#include <ScrollBarConstants.au3>
#include <SliderConstants.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include "BinaryStrings.au3"
#include "GIFAnimation.au3"
#include "MemoryDll.au3"
#include "MPDF_UDF.au3"

Opt("MustDeclareVars", 1)
Opt("GUICoordMode", 1)
Opt("GUICloseOnESC", 0)
Opt("GUIOnEventMode", 1)
Opt("MouseCoordMode", 1)

Global Const $title = "Rk Screencapture Alpha "
Global Const $ver = "v0.02"


If @OSBuild < 2600 Then Exit _WinAPI_ShowError($title & @LF & @LF & "is not running on " & @OSVersion & " !!! :-(")

#OnAutoItStartRegister "OnAutoItStart"
Global $__Restart = False

Global Const $hDwmApiDll = DllOpen("dwmapi.dll")
Global $sChkAero = DllStructCreate("int;")
DllCall($hDwmApiDll, "int", "DwmIsCompositionEnabled", "ptr", DllStructGetPtr($sChkAero))
Global $aero = DllStructGetData($sChkAero, 1)

#region DEP check
Global $DEP_Status
Global Const $aDEP = DllCall("kernel32.dll", "uint", "GetSystemDEPPolicy") ;http://msdn.microsoft.com/en-us/library/windows/desktop/bb736298(v=vs.85).aspx
If Not @error Then
	$DEP_Status = $aDEP[0]
Else ;alternative check needed for WinXP with SP2 or less or Vista without any service pack
	Global $line, $DEP_out, $DEP_chk, $DEP_info
	Global $pid = Run(@ComSpec & " /c wmic OS Get DataExecutionPrevention_SupportPolicy", @SystemDir, @SW_HIDE, $STDOUT_CHILD)
	While 1
		$line = StdoutRead($pid)
		If @error Then ExitLoop
		If $line <> "" Then $DEP_out &= StringStripWS($line, 7)
	WEnd
	$DEP_chk = StringRight($DEP_out, 1)
	If Asc($DEP_chk) < 48 Or Asc($DEP_chk) > 57 Then
		$DEP_info = MsgBox(4 + 48 + 262144, "Warning", "Could not check whether Data Execution Prevention (DEP) is enabled on your system!" & @LF & @LF & _
				"Please check manually whether DEP is enabled for all programs and" & @LF & _
				"servers and add " & @ScriptName & " to the exclusion list manually!" & @LF & @LF & _
				"Otherwise " & @ScriptName & " may hard crash!!!" & @LF & @LF & @LF & _
				"Do you want more information about DEP?" & @LF & @LF & _
				"(http://windows.microsoft.com/en-US/windows-vista/Data-Execution-Prevention-frequently-asked-questions will be opened)")
		If $DEP_info = 6 Then ShellExecute("http://windows.microsoft.com/en-US/windows-vista/Data-Execution-Prevention-frequently-asked-questions")
	Else
		$DEP_Status = Number($DEP_chk)
	EndIf
EndIf

If $DEP_Status > 2 Then
	Global $DEP_RegKey_Prefix = "HKLM"
	If @OSArch = "X64" Then $DEP_RegKey_Prefix = "HKLM64"
	Global Const $DEP_RegKey = $DEP_RegKey_Prefix & "\SOFTWARE\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers"
	Global $rz = 1, $var, $DEP_Add2List, $DEP_ok = False
	While True
		$var = RegEnumVal($DEP_RegKey, $rz)
		If @error <> 0 Then ExitLoop
		$rz += 1
		If @ScriptFullPath = $var Then
			$DEP_ok = True
			ExitLoop
		EndIf
	WEnd

	If Not $DEP_ok Then
		$DEP_Add2List = MsgBox(4 + 48 + 262144, "Warning", _
				"Data Execution Prevention (DEP) is enabled for all processes!" & @LF & @LF & _
				"Some functions will CRASH when started if " & @ScriptName & " is not in the exception list!" & @LF & @LF & @LF & _
				"Should I add " & @ScriptName & " to the DEP exception list?" & @LF & @LF)
		If $DEP_Add2List = 6 Then
			Run('rundll32.exe sysdm.cpl, NoExecuteAddFileOptOutList "' & @ScriptFullPath & '"', @SystemDir, @SW_HIDE)
			If MsgBox(4 + 64 + 262144, "Information", @ScriptName & " properly added to exception list!" & @LF & @LF & _
					"Restart Rk Screencapture now?" & @LF & @LF, 30) = 6 Then _ScriptRestart()
			Exit
		EndIf
	EndIf
EndIf
#endregion DEP check

Global $icon_GUI = 196
Global $icon_redo = 238
Global $icon_jpg_qual = 302
Global $icon_timestamp = 239
If @OSBuild < 6000 Or Not $aero Then
	MsgBox(48, "Warning", "This application is designed for Vista+ OS with Aero feature enabled!" & @LF & @LF & _
			"Either AERO is disbaled or current operating system" & @LF & _
			"does not support some features e.g. to make screenshots" & @LF & _
			"of web pages!", 60)
	$icon_GUI = 196
	$icon_redo = 146
	$icon_jpg_qual = 161
	$icon_timestamp = 20
EndIf

;capture sound
Global Const $SND_SYNC = 0, $SND_ASYNC = 1, $SND_NODEFAULT = 2, $SND_MEMORY = 4, $SND_NOWAIT = 0x2000
Global $binCapture_snd = Capture_Sound()
Global $ByteStruct = DllStructCreate("byte[" & BinaryLen($binCapture_snd) & "]")
DllStructSetData($ByteStruct, 1, $binCapture_snd)
$binCapture_snd = 0
Global Const $BytePtr = DllStructGetPtr($ByteStruct)
Global Const $fuSound = BitOR($SND_MEMORY, $SND_ASYNC, $SND_NODEFAULT, $SND_NOWAIT)
Global $silent = False

#region GUI part
Global Const $AW_ACTIVATE = 0x00020000, $AW_BLEND = 0x00080000, $AW_HIDE = 0x00010000

Global Const $MSLLHOOKSTRUCT = $tagPOINT & ";dword mouseData;dword flags;dword time;ulong_ptr dwExtraInfo"
Global $hKey_Proc = DllCallbackRegister("_Mouse_Proc", "int", "int;ptr;ptr")
Global $hM_Module = DllCall("kernel32.dll", "hwnd", "GetModuleHandle", "ptr", 0)
Global $hM_Hook = DllCall("user32.dll", "hwnd", "SetWindowsHookEx", "int", $WH_MOUSE_LL, "ptr", DllCallbackGetPtr($hKey_Proc), "hwnd", $hM_Module[0], "dword", 0)

Global Const $gdip_x = 264
Global Const $gdip_y = 80
Global Const $gdip_w = 635
Global Const $gdip_h = 454

Global Const $dll = DllOpen("user32.dll")

Global $hGui_Slider, $Slider, $hSlider, $Slider_Label_Text, $Slider_Label_Value, $_msg, $aPos_SC, $Grab_x, $Grab_y, $Grab_w, $Grab_h

Global $_MDCodeBuffer, $_MDLoadOffset, $_MDGetOffset, $_MDFreeOffset, $_MDKernel32Dll, $F_DLL
Global $_MFHookPtr, $_MFHookBak, $_MFHookApi = "LocalCompact"

Global $cursor, $Avi_Handle, $Avi32_Dll, $rec_time, $compress_avi = True
Global Const $OF_CREATE = 0x00001000
Global Const $ICMF_CHOOSE_KEYFRAME = 1, $ICMF_CHOOSE_DATARATE = 2
Global Const $AVIIF_KEYFRAME = 0x00000010
Global Const $AVIERR_UNSUPPORTED = 0x80044065
Global Const $AVIERR_BADPARAM = 0x80044066
Global Const $AVIERR_MEMORY = 0x80044067
Global Const $AVIERR_NOCOMPRESSOR = 0x80044071
Global Const $AVIERR_CANTCOMPRESS = 0x80044075
Global Const $AVIERR_ERROR = 0x800440C7
Global Const $AVIERR_OK = 0

;http://msdn.microsoft.com/en-us/library/dd183374(v=vs.85).aspx
Global Const $BITMAPFILEHEADER = "WORD bfType;DWORD bfSize;WORD bfReserved1;WORD bfReserved2;DWORD bfOffBits;"
;~ Global Const $BITMAPFILEHEADER = "align 2;char magic[2];int size;short res1;short res2;ptr offset;"

;http://msdn.microsoft.com/en-us/library/dd183376(v=vs.85).aspx
Global Const $BITMAPINFOHEADER = _
		"dword biSize;long biWidth;long biHeight;short biPlanes;short biBitCount;dword biCompression;" & _
		"dword biSizeImage;long biXPelsPerMeter;long biYPelsPerMeter;dword biClrUsed;dword biClrImportant;"

;http://msdn.microsoft.com/en-us/library/dd318180(v=VS.85).aspx
Global Const $AVIMAINHEADER = _
		"FOURCC fcc;DWORD cb;DWORD dwMicroSecPerFrame;DWORD dwMaxBytesPerSec;DWORD dwPaddingGranularity;" & _
		"DWORD dwFlags;DWORD dwTotalFrames;DWORD dwInitialFrames;DWORD dwStreams;DWORD dwSuggestedBufferSize;" & _
		"DWORD dwWidth;DWORD dwHeight;"

;http://msdn.microsoft.com/en-us/library/ms899423.aspx
Global Const $AVISTREAMINFO = _
		"dword fccType;dword fccHandler;dword dwFlags;dword dwCaps;short wPriority;short wLanguage;dword dwScale;" & _
		"dword dwRate;dword dwStart;dword dwLength;dword dwInitialFrames;dword dwSuggestedBufferSize;dword dwQuality;" & _
		"dword dwSampleSize;int rleft;int rtop;int rright;int rbottom;dword dwEditCount;dword dwFormatChangeCount;wchar[64];"

;http://msdn.microsoft.com/en-us/library/dd756791(v=VS.85).aspx
Global Const $AVICOMPRESSOPTIONS = _
		"DWORD fccType;DWORD fccHandler;DWORD dwKeyFrameEvery;DWORD dwQuality;DWORD dwBytesPerSecond;" & _
		"DWORD dwFlags;PTR lpFormat;DWORD cbFormat;PTR lpParms;DWORD cbParms;DWORD dwInterleaveEvery;"

Global Const $zmin = 5
Global Const $zmax = 15
Global Const $margin = 6
Global $hGUI_Area, $hGUI_Dot, $w_area, $h_area
Global $z = 10
Global $aWnd, $URL_Input, $cont = 1, $i, $Grab_Button, $Dummy, $filename
Global $f, $bW, $BH, $scroll_x, $scroll_y
Global Const $scroll_speedx = 10, $scroll_speedy = 10
Global $hGfx, $hCtxt, $hGUI_About, $hHBITMAP, $hBmp_A, $hImg_A, $tRectF, $hBrush, $hFormat, $aText, $tLayout, $speed, $About_End
Global $vBitmap, $hBitmap_s, $undo, $undo_chk = False
Global Const $bubbles = 12, $max_speed = 3, $min_size = 50, $max_size = 70, $dh = 0
Global $aData[$bubbles][6] ;x,y,vx,vy,size,bmp

Global $hBackImage, $hImageContext, $hIA
Global $hBitmap2, $hClipboard_Bitmap, $hBmp, $hMemoryBMP
Global $hDC_Region, $hMemDC, $memBitmap, $aFullScreen, $hFullScreen
$hFullScreen = WinGetHandle("[TITLE:Program Manager;CLASS:Progman]")
$aFullScreen = WinGetPos($hFullScreen)
Global $hObj = _WinAPI_SelectObject($hMemDC, $memBitmap)
_GDIPlus_Startup()

Global Const $width = 907, $height = 642
Global $hGUI = GUICreate($title & $ver, $width, $height, -1, -1)
GUISetFont(9, 400, 0, "Times New Roman")
Global $bg_c = "BFCDDB"
GUISetBkColor("0x" & $bg_c, $hGUI)
Global $hBrush_Clear = _GDIPlus_BrushCreateSolid("0xFF" & $bg_c)

Global Enum $id_ChkUpd = 0x400, $id_VisitWeb, $id_Ruler, $id_OpenMailClient, $id_About, $id_Print, $id_Exit, $id_Low, $id_Medium, $id_High, $id_Reset, _
		$id_Grey, $id_BW, $id_Invert, $id_Undo, $id_Rotm90, $id_Rotp90, $id_ImgEditor

Global $hQMenu, $hQMenu_Sub, $hQMenu_Sub2, $hQMenu_Sub3

$hQMenu_Sub = _GUICtrlMenu_CreatePopup()
_GUICtrlMenu_InsertMenuItem($hQMenu_Sub, 0, "Low", $id_Low)
_GUICtrlMenu_InsertMenuItem($hQMenu_Sub, 1, "Medium", $id_Medium)
_GUICtrlMenu_InsertMenuItem($hQMenu_Sub, 2, "High", $id_High)

$hQMenu_Sub2 = _GUICtrlMenu_CreatePopup()
_GUICtrlMenu_InsertMenuItem($hQMenu_Sub2, 0, "Reset", $id_Reset)

$hQMenu_Sub3 = _GUICtrlMenu_CreatePopup()
_GUICtrlMenu_InsertMenuItem($hQMenu_Sub3, 0, "Grayscale", $id_Grey)
_GUICtrlMenu_InsertMenuItem($hQMenu_Sub3, 1, "Black & White", $id_BW)
_GUICtrlMenu_InsertMenuItem($hQMenu_Sub3, 2, "Invert", $id_Invert)
_GUICtrlMenu_InsertMenuItem($hQMenu_Sub3, 3, "Rotate 90° left", $id_Rotm90)
_GUICtrlMenu_InsertMenuItem($hQMenu_Sub3, 4, "Rotate 90° right", $id_Rotp90)
_GUICtrlMenu_InsertMenuItem($hQMenu_Sub3, 5, "", 0)
_GUICtrlMenu_InsertMenuItem($hQMenu_Sub3, 6, "Undo", $id_Undo)
_GUICtrlMenu_InsertMenuItem($hQMenu_Sub3, 7, "", 0)
_GUICtrlMenu_InsertMenuItem($hQMenu_Sub3, 8, "Image Editor", $id_ImgEditor)
_GUICtrlMenu_EnableMenuItem($hQMenu_Sub3, 6, 2)
_GUICtrlMenu_EnableMenuItem($hQMenu_Sub3, 8, 2)

$hQMenu = _GUICtrlMenu_CreatePopup()
_GUICtrlMenu_InsertMenuItem($hQMenu, 0, "Display Image Quality", 0, $hQMenu_Sub)
_GUICtrlMenu_InsertMenuItem($hQMenu, 1, "", 0)
_GUICtrlMenu_InsertMenuItem($hQMenu, 2, "Reset View", 0, $hQMenu_Sub2)
_GUICtrlMenu_InsertMenuItem($hQMenu, 3, "", 0)
_GUICtrlMenu_InsertMenuItem($hQMenu, 4, "Image Editing", 0, $hQMenu_Sub3)

Global $hBMP_Quality = Load_BMP_From_Mem(Quality_Icon(), True)
_GUICtrlMenu_SetItemBmp($hQMenu, 0, $hBMP_Quality)
Global $hBMP_Reset = Load_BMP_From_Mem(ResetView_Icon(), True)
_GUICtrlMenu_SetItemBmp($hQMenu, 2, $hBMP_Reset)
Global $hBMP_Reset2 = Load_BMP_From_Mem(Reset_Icon(), True)
_GUICtrlMenu_SetItemBmp($hQMenu_Sub2, 0, $hBMP_Reset2)

Global $hBMP_ImageEdit = Load_BMP_From_Mem(Editing_Icon(), True)
_GUICtrlMenu_SetItemBmp($hQMenu, 4, $hBMP_ImageEdit)

Global $hBMP_ImageEdit_Invert = Load_BMP_From_Mem(Invert_Icon(), True)
;~ Global $hDC_Source = _WinAPI_CreateCompatibleDC(0)
;~ Global $obj = _WinAPI_SelectObject($hDC_Source, $hBMP_ImageEdit_Invert ) ; invert icon
;~ _WinAPI_BitBlt($hDC_Source, 0, 0, 32, 32, $hDC_Source, 0, 0, $DSTINVERT)
;~ _WinAPI_DeleteDC($hDC_Source)
;~ _WinAPI_SelectObject($hDC_Source, $obj)
_GUICtrlMenu_SetItemBmp($hQMenu_Sub3, 2, $hBMP_ImageEdit_Invert)

Global $hBMP_ImageEdit_Gray = Load_BMP_From_Mem(Gray_Icon(), True)
_GUICtrlMenu_SetItemBmp($hQMenu_Sub3, 0, $hBMP_ImageEdit_Gray)
Global $hBMP_ImageEdit_BW = Load_BMP_From_Mem(BW_Icon(), True)
_GUICtrlMenu_SetItemBmp($hQMenu_Sub3, 1, $hBMP_ImageEdit_BW)
Global $hBMP_ImageEdit_Rotp90 = Load_BMP_From_Mem(Rotate_Icon(), True)
_GUICtrlMenu_SetItemBmp($hQMenu_Sub3, 3, $hBMP_ImageEdit_Rotp90)
Global $hBMP_ImageEdit_Rotm90 = Load_BMP_From_Mem(Rotate_Icon(), True)
_GUICtrlMenu_SetItemBmp($hQMenu_Sub3, 4, $hBMP_ImageEdit_Rotm90)
Global $hBMP_ImageEdit_Undo = Load_BMP_From_Mem(Undo_Icon(), True)
_GUICtrlMenu_SetItemBmp($hQMenu_Sub3, 6, $hBMP_ImageEdit_Undo)
Global $hBMP_ImageEdit_Editor = Load_BMP_From_Mem(Editor_Icon(), True)
_GUICtrlMenu_SetItemBmp($hQMenu_Sub3, 8, $hBMP_ImageEdit_Editor)

Global $hMenu = _GUICtrlMenu_GetSystemMenu($hGUI)
_GUICtrlMenu_AppendMenu($hMenu, $MF_SEPARATOR, 0, 0)
_GUICtrlMenu_AppendMenu($hMenu, $MF_STRING, $id_Ruler, "Ruler")
_GUICtrlMenu_AppendMenu($hMenu, $MF_STRING, $id_OpenMailClient, "Send current image via mail")
_GUICtrlMenu_AppendMenu($hMenu, $MF_STRING, $id_Print, "Print current image")
_GUICtrlMenu_AppendMenu($hMenu, $MF_STRING, $id_VisitWeb, "Visit Web Site")
_GUICtrlMenu_AppendMenu($hMenu, $MF_STRING, $id_About, "About")
_GUICtrlMenu_AppendMenu($hMenu, $MF_SEPARATOR, 0, 0)
_GUICtrlMenu_AppendMenu($hMenu, $MF_STRING, $id_Exit, "Exit")

Global $hBMP_Ruler = Load_BMP_From_Mem(Ruler_Image(), True)
_GUICtrlMenu_SetItemBmp($hMenu, 8, $hBMP_Ruler)
Global $hBMP_OpenMailClient = Load_BMP_From_Mem(Envelope_Icon(), True)
_GUICtrlMenu_SetItemBmp($hMenu, 9, $hBMP_OpenMailClient)
Global $hBMP_Print = _GUICtrlMenu_CreateBitmap_WinAPI(@SystemDir & "\Shell32.dll", 143)
_GUICtrlMenu_SetItemBmp($hMenu, 10, $hBMP_Print)
Global $hBMP_ChkUpd = _GUICtrlMenu_CreateBitmap_WinAPI(@SystemDir & "\Shell32.dll", 91)
_GUICtrlMenu_SetItemBmp($hMenu, 11, $hBMP_ChkUpd)
Global $hBMP_VisitWeb = _GUICtrlMenu_CreateBitmap_WinAPI(@SystemDir & "\Shell32.dll", 13)
_GUICtrlMenu_SetItemBmp($hMenu, 12, $hBMP_VisitWeb)
Global $hBMP_About = _GUICtrlMenu_CreateBitmap_WinAPI(@SystemDir & "\Shell32.dll", 130)
_GUICtrlMenu_SetItemBmp($hMenu, 13, $hBMP_About)
Global $hBMP_Exit = _GUICtrlMenu_CreateBitmap_WinAPI(@SystemDir & "\Shell32.dll", 27)
_GUICtrlMenu_SetItemBmp($hMenu, 14, $hBMP_Exit)

Global $List = GUICtrlCreateListView("", 4, 4, 250, 605, $LVS_SHOWSELALWAYS + $LVS_SINGLESEL, $LVS_EX_FULLROWSELECT + $LVS_EX_GRIDLINES + $WS_EX_CLIENTEDGE + $LVS_EX_DOUBLEBUFFER)
;~ Global $List = GUICtrlCreateListView("", 4, 4, 250, 605, BitOR($WS_EX_DLGMODALFRAME, $WS_EX_CLIENTEDGE), BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_SUBITEMIMAGES, $LVS_EX_GRIDLINES, $LVS_EX_CHECKBOXES, $LVS_EX_DOUBLEBUFFER))
GUICtrlSetTip(-1, "Displays only visible windows - minimized windows not possible!" & @LF & @LF & _
		"Double click with rmb to use alternative method", "", 0, 1)
GUICtrlSetBkColor(-1, 0xFFFFFF)

Global $hListView = GUICtrlGetHandle($List)
Global $Button_Refresh = GUICtrlCreateButton("&Refresh List", 3, 610, 252)
GUICtrlSetCursor(-1, 0)
GUICtrlSetTip(-1, "Click button or press F5 to refresh list", "", 0, 1)

Global $Button_Exit = GUICtrlCreateButton(" E&xit", 800, 555, 99, 82)
GUICtrlSetFont(-1, 13, 400, 0, "Times New Roman")
_GUICtrlButton_SetImage($Button_Exit, "shell32.dll", 27, True)
GUICtrlSetCursor(-1, 0)

Global $JPEG_Quality = 90
Global $Button_Save = GUICtrlCreateButton("&Save" & @LF & " Image", 693, 555, 99, 82, $BS_MULTILINE)
_GUICtrlButton_SetImage($Button_Save, "shell32.dll", 6, True)
GUICtrlSetFont(-1, 12, 400, 0, "Times New Roman")
GUICtrlSetCursor(-1, 0)
GUICtrlSetTip(-1, "Save image in its full resolution." & @LF & _
		"Enter the extension to save in appropriate format." & @LF & _
		"Press RMB to set JPG quality" & @LF & _
		"Default is 90" & @LF & _
		"or add time stamp to saved image.", "", 0, 1)

Global $Button_Menu = GUICtrlCreateContextMenu($Button_Save)
Global $Button_Menu_Item1 = GUICtrlCreateMenuItem("Set JPEG save quality", $Button_Menu)
GUICtrlCreateMenuItem("", $Button_Menu)
Global $Button_Menu_Item2 = GUICtrlCreateMenuItem("Add time stamp to saved image", $Button_Menu)

Global $hButton_Menu_Save = GUICtrlGetHandle($Button_Menu)
Global $hBMP_Menu_JPG_Qual = _GUICtrlMenu_CreateBitmap_WinAPI(@SystemDir & "\Shell32.dll", $icon_jpg_qual)
_GUICtrlMenu_SetItemBmp($hButton_Menu_Save, 0, $hBMP_Menu_JPG_Qual)

Global $hButton_Menu_TS = GUICtrlGetHandle($Button_Menu)
Global $hBMP_Menu_Timestamp = _GUICtrlMenu_CreateBitmap_WinAPI(@SystemDir & "\Shell32.dll", $icon_timestamp)
_GUICtrlMenu_SetItemBmp($hButton_Menu_TS, 2, $hBMP_Menu_Timestamp)

Global $Button_Clipboard = GUICtrlCreateButton("&Put to" & @LF & "Clipboard", 587, 555, 99, 82, $BS_MULTILINE)
_GUICtrlButton_SetImage($Button_Clipboard, "shell32.dll", 260, True)
GUICtrlSetFont(-1, 10, 400, 0, "Times New Roman")
GUICtrlSetCursor(-1, 0)
GUICtrlSetTip(-1, "Put image to clipboard", "", 0, 1)

Global $Button_WebGrab = GUICtrlCreateButton("Screenshot" & @LF & "a &Web Site", 372, 555, 99, 82, $BS_MULTILINE)
GUICtrlSetFont(-1, 9, 400, 0, "Times New Roman")
GUICtrlSetCursor(-1, 0)
GUICtrlSetTip(-1, "Screenshot a web site (Aero must be enabled!)" & @LF & "Press RMB to disable URL in grabbed image", "", 0, 1)
_GUICtrlButton_SetImage($Button_WebGrab, "shell32.dll", 13, True)

Global $Button_Grab2AVI = GUICtrlCreateButton("Grab Screen" & @LF & "to &AVI", 478, 555, 99, 82, $BS_MULTILINE)
GUICtrlSetFont(-1, 9, 400, 0, "Times New Roman")
GUICtrlSetCursor(-1, 0)
GUICtrlSetTip(-1, "Grab window under mouse to AVI." & @LF & @LF & _
		"Press RMB to modify AVI settings!", "", 0, 1)
_GUICtrlButton_SetImage($Button_Grab2AVI, "shell32.dll", 115, True)

Global $Button_Menu_AVI = GUICtrlCreateContextMenu($Button_Grab2AVI)
Global $Button_Menu_AVI_Sub1 = GUICtrlCreateMenu("AVI Record Time", $Button_Menu_AVI, 1)
GUICtrlCreateMenuItem("", $Button_Menu_AVI)
Global $aButton_Menu_AVI_Sub1_Item[5], $AVI_ReqTime[5] = [5, 10, 30, 60, 0], $AVI_FPS[5] = [1, 2, 5, 10, 0]

$aButton_Menu_AVI_Sub1_Item[0] = GUICtrlCreateMenuItem($AVI_ReqTime[0] & " sec", $Button_Menu_AVI_Sub1, 1)
$aButton_Menu_AVI_Sub1_Item[1] = GUICtrlCreateMenuItem($AVI_ReqTime[1] & " sec", $Button_Menu_AVI_Sub1, 2)
GUICtrlSetState(-1, $GUI_CHECKED)
$aButton_Menu_AVI_Sub1_Item[2] = GUICtrlCreateMenuItem($AVI_ReqTime[2] & " sec", $Button_Menu_AVI_Sub1, 3)
$aButton_Menu_AVI_Sub1_Item[3] = GUICtrlCreateMenuItem($AVI_ReqTime[3] & " sec", $Button_Menu_AVI_Sub1, 4)
GUICtrlCreateMenuItem("", $Button_Menu_AVI_Sub1)
$aButton_Menu_AVI_Sub1_Item[4] = GUICtrlCreateMenuItem("Custom", $Button_Menu_AVI_Sub1, 5)

Global $Button_Menu_AVI_Sub2 = GUICtrlCreateMenu("AVI FPS", $Button_Menu_AVI, 2)
Global $aButton_Menu_AVI_Sub2_Item[6]
$aButton_Menu_AVI_Sub2_Item[0] = GUICtrlCreateMenuItem($AVI_FPS[0] & " FPS", $Button_Menu_AVI_Sub2, 1)
GUICtrlSetState(-1, $GUI_CHECKED)
$aButton_Menu_AVI_Sub2_Item[1] = GUICtrlCreateMenuItem($AVI_FPS[1] & " FPS", $Button_Menu_AVI_Sub2, 2)
$aButton_Menu_AVI_Sub2_Item[2] = GUICtrlCreateMenuItem($AVI_FPS[2] & " FPS", $Button_Menu_AVI_Sub2, 3)
$aButton_Menu_AVI_Sub2_Item[3] = GUICtrlCreateMenuItem($AVI_FPS[3] & " FPS", $Button_Menu_AVI_Sub2, 4)
GUICtrlCreateMenuItem("", $Button_Menu_AVI_Sub2)
$aButton_Menu_AVI_Sub2_Item[4] = GUICtrlCreateMenuItem("Custom", $Button_Menu_AVI_Sub2, 5)

GUICtrlCreateMenuItem("", $Button_Menu_AVI)
Global Const $Button_Menu_AVI_Sub3 = GUICtrlCreateMenu("Add timestamp to frames", $Button_Menu_AVI, 4)
Global Const $mButton_Menu_AVI_Sub3_Item1 = GUICtrlCreateMenuItem("Enabled", $Button_Menu_AVI_Sub3, 1, 1)
Global Const $mButton_Menu_AVI_Sub3_Item2 = GUICtrlCreateMenuItem("Disabled", $Button_Menu_AVI_Sub3, 2, 1)
GUICtrlSetState($mButton_Menu_AVI_Sub3_Item2, $GUI_CHECKED)

GUICtrlCreateMenuItem("", $Button_Menu_AVI)
Global Const $Button_Menu_AVI_Sub4 = GUICtrlCreateMenu("Compress AVI stream", $Button_Menu_AVI, 6)
Global Const $mButton_Menu_AVI_Sub4_Item1 = GUICtrlCreateMenuItem("Enabled", $Button_Menu_AVI_Sub4, 1, 1)
Global Const $mButton_Menu_AVI_Sub4_Item2 = GUICtrlCreateMenuItem("Disabled", $Button_Menu_AVI_Sub4, 2, 1)
GUICtrlSetState($mButton_Menu_AVI_Sub4_Item1, $GUI_CHECKED)

GUICtrlCreateMenuItem("", $Button_Menu_AVI)
Global Const $Button_Menu_AVI_Sub5 = GUICtrlCreateMenu("Capture Mouse Cursor", $Button_Menu_AVI, 8)
Global Const $mButton_Menu_AVI_Sub5_Item1 = GUICtrlCreateMenuItem("Enabled", $Button_Menu_AVI_Sub5, 1, 1)
Global Const $mButton_Menu_AVI_Sub5_Item2 = GUICtrlCreateMenuItem("Disabled", $Button_Menu_AVI_Sub5, 2, 1)
GUICtrlSetState($mButton_Menu_AVI_Sub5_Item1, $GUI_CHECKED)

Global Const $Button_Menu_Web = GUICtrlCreateContextMenu($Button_WebGrab)
Global Const $Button_Menu_Web_Item = GUICtrlCreateMenuItem("Add URL to grabbed image", $Button_Menu_Web)
GUICtrlSetState(-1, $GUI_CHECKED)

If @OSBuild < 6000 Or Not $aero Then GUICtrlSetState($Button_WebGrab, $GUI_DISABLE)
$sChkAero = 0

Global Const $txt = "Grab region from desktop manually." & @LF & _
		"Use mousewheel to change zoom factor," & @LF & _
		"Hold STRG pressed to select control - " & @LF & _
		"STRG + SHIFT will grab the control." & @LF & _
		"SHIFT only to use freehand capturing!" & @LF & _
		"RMB to capture again on last position."
Global $Button_Menu_GrabScreen, $Button_Menu_GrabScreen_Item
If @OSBuild < 6000 Then
	Global $Pic_Grab = GUICtrlCreatePic("", 274, 555)
	GUICtrlSetCursor(-1, 0)
	Global $hHBITMAP_GrabScreen = Load_BMP_From_Mem(Grab_Image(), True)
	Global $hPic_Grab = GUICtrlGetHandle($Pic_Grab)
	Global $a_Ret = DllCall("User32.dll", "hwnd", "SendMessage", "hwnd", $hPic_Grab, "int", 0x0172, "int", 0, "int", $hHBITMAP_GrabScreen)
	If $a_Ret[0] <> 0 Then _WinAPI_DeleteObject($a_Ret[0])
	Global $label1 = GUICtrlCreateLabel("", 264, 557, 99, 78, Default, $WS_EX_STATICEDGE)
	GUICtrlSetBkColor(-1, -2)
	GUICtrlSetOnEvent($Pic_Grab, "Grab_Screen")
	GUICtrlSetTip($Pic_Grab, $txt, "", 0, 1)
	$Button_Menu_GrabScreen = GUICtrlCreateContextMenu($Pic_Grab)
	$Button_Menu_GrabScreen_Item = GUICtrlCreateMenuItem("Capture again on last position", $Button_Menu_GrabScreen)
Else
	Global $Button_GrabScreen = GUICtrlCreateButton("&Grab Screen", 264, 555, 99, 82, $BS_BITMAP)
	Global $hHBITMAP_GrabScreen = Load_BMP_From_Mem(Grab_Image(), True)
	Global Const $hButton_GrabScreen = GUICtrlGetHandle($Button_GrabScreen)
	_WinAPI_DeleteObject(_SendMessage($hButton_GrabScreen, $BM_SETIMAGE, 0, $hHBITMAP_GrabScreen))
	_WinAPI_UpdateWindow($hButton_GrabScreen)
	_WinAPI_DeleteObject($hHBITMAP_GrabScreen)
	GUICtrlSetCursor($Button_GrabScreen, 0)
	GUICtrlSetTip($Button_GrabScreen, $txt, "", 0, 1)
	GUICtrlSetOnEvent($Button_GrabScreen, "Grab_Screen")
	$Button_Menu_GrabScreen = GUICtrlCreateContextMenu($Button_GrabScreen)
	$Button_Menu_GrabScreen_Item = GUICtrlCreateMenuItem("Capture again on last position", $Button_Menu_GrabScreen)
EndIf
Global $hButton_Menu_GrabScreen = GUICtrlGetHandle($Button_Menu_GrabScreen)
Global $hBMP_Menu_GrabScreen_Redo = _GUICtrlMenu_CreateBitmap_WinAPI(@SystemDir & "\Shell32.dll", $icon_redo)
_GUICtrlMenu_SetItemBmp($hButton_Menu_GrabScreen, 0, $hBMP_Menu_GrabScreen_Redo)

Global $Group = GUICtrlCreateGroup("Rk 2011 ", 259, 0, $gdip_w + 9, 70, $GUI_SS_DEFAULT_GROUP + $BS_RIGHT)
GUICtrlSetFont(-1, 8, 200, 0, "Times New Roman")
GUICtrlCreateGroup("", -99, -99, 1, 1)

_3DText("Rk Screencapture", 277, -3)

Global $hGraphic = _GDIPlus_GraphicsCreateFromHWND($hGUI)
Global $hBuffer_Bmp = _GDIPlus_BitmapCreateFromGraphics($gdip_w, $gdip_h, $hGraphic)
Global $hContext = _GDIPlus_ImageGetGraphicsContext($hBuffer_Bmp)
DllCall($ghGDIPDll, "uint", "GdipSetPixelOffsetMode", "hwnd", $hContext, "int", 2)
_GDIPlus_GraphicsSetInterpolationMode($hContext, 0) ; set interpolation quality
_GUICtrlMenu_CheckRadioItem($hQMenu_Sub, 0, 2, 1)

Global $hMatrix = _GDIPlus_MatrixCreate()
_GDIPlus_MatrixTranslate($hMatrix, $gdip_w / 2, $gdip_h / 2)

Global $Gfx = GUICtrlCreateGraphic($gdip_x - 2, $gdip_y - 2, $gdip_w + 4, $gdip_h + 4)
If @OSBuild >= 6000 Then GUICtrlSetBkColor($Gfx, 0xFFFFFF)
GUICtrlSetColor($Gfx, 0x000060)

Global $label3 = GUICtrlCreateLabel("Image Dimension: ", $gdip_x + 2, 538, -1, 18)
GUICtrlSetBkColor(-1, -2)
Global $label4 = GUICtrlCreateLabel("", $gdip_x + 90, 538, 65, 14)
GUICtrlSetBkColor(-1, -2)

;~ Global $label5 = GUICtrlCreateLabel("", $gdip_x - 2, $gdip_y - 2, $gdip_w + 4, $gdip_h + 4, 0x0000000D) ;0x0000000D = $SS_OwnerDraw
;~ GUICtrlSetBkColor(-1, -2)
;~ GUICtrlSetTip(-1, "Test")

_GUICtrlListView_InsertColumn($List, 0, "Windows Name", 165)
_GUICtrlListView_InsertColumn($List, 1, "Handle", 80)
Refresh_Wnd_List()

GUICtrlSetOnEvent($Button_Refresh, "Refresh_Wnd_List")
GUICtrlSetOnEvent($Button_WebGrab, "WebGrab")
GUICtrlSetOnEvent($Button_Grab2AVI, "Grab2AVI")
GUICtrlSetOnEvent($Button_Clipboard, "Pic2Clipboard")
GUICtrlSetOnEvent($Button_Save, "Save_Bitmap")
GUICtrlSetOnEvent($Button_Exit, "_Exit")
#endregion GUI part

GUIRegisterMsg($WM_COMMAND, "WM_COMMAND")
GUIRegisterMsg($WM_CONTEXTMENU, "WM_CONTEXTMENU")

GUIRegisterMsg($WM_NOTIFY, "WM_NOTIFY")
GUIRegisterMsg($WM_SYSCOMMAND, "WM_SYSCOMMAND")

Global $NCPaint = False
If @OSBuild < 6000 Then
	GUIRegisterMsg($WM_PAINT, "Redraw_XP")
Else
	GUIRegisterMsg($WM_NCPAINT, "Redraw")
EndIf

;~ _WinAPI_AnimateWindow($hGUI, BitOR($AW_BLEND, $AW_ACTIVATE), 300)
GUISetState(@SW_SHOW, $hGUI)

GUISetIcon(@SystemDir & "\Shell32.dll", -$icon_GUI, $hGUI)

;activate random window
Refresh_Wnd_List()
Global $random = Random(0, _GUICtrlListView_GetItemCount($List) - 1, 1)
_GUICtrlListView_SetItemSelected($List, $random)
Capture_Window(_GUICtrlListView_GetItemText($List, $random, 1), $aWnd[$random][2], $aWnd[$random][3])
_GUICtrlListView_SetItemSelected($List, $random)

Global Const $oErrorHandler = ObjEvent("AutoIt.Error", "ObjErrorHandler")
Global $hGUI_Freehand, $hBmp_Freehand, $hGraphic_Freehand
Global $mc, $mc2, $mc3, $sx, $sy, $mpointer
Global $zx, $zy, $up, $zwm, $hGUI_Zoom, $hDC_Zoom, $hGUI_ZoomDC
Global $dirx = 1, $drx = 0, $diry = 1, $dry = 0, $lower_border, $right_border, $aLastPos[2]
Global Const $radius = 180
Global Const $zw = 256
Global Const $zh = 256
Global Const $zoom_level_default = 6
Global $zoom_level = $zoom_level_default
Global $zoomW = Int($zw / $zoom_level)
Global $zoomH = Int($zh / $zoom_level)
Global Const $zoom_min = 2, $zoom_max = 24
Global $red, $green, $blue

Global $B_DESCENDING[_GUICtrlListView_GetColumnCount($hListView)]

HotKeySet("^!{F12}", "Grab_Active_Window") ;Ctrl+Alt+F12 to take a screenshot of active window

GUISetOnEvent($GUI_EVENT_CLOSE, "_Exit")

While Sleep(50)
	If $NCPaint Then
		_GDIPlus_GraphicsDrawImageRect($hGraphic, $hBuffer_Bmp, $gdip_x, $gdip_y, $gdip_w, $gdip_h)
		$NCPaint = False
	EndIf
	If WinActive($hGUI) Then
		If _IsPressed("79", $dll) And $undo_chk Then ;F10
			Undo()
			Zoom(2)
			_GUICtrlMenu_EnableMenuItem($hQMenu_Sub3, 6, 2)
		EndIf
		If _IsPressed("74", $dll) Then Refresh_Wnd_List() ;F5
		If _IsPressed("7A", $dll) Then ;F11
			GUISetState(@SW_MINIMIZE, $hGUI)
			Sleep(250)
			Grab_Region($aFullScreen[0], $aFullScreen[1], $aFullScreen[2], $aFullScreen[3], True)
			sndPlaySound($BytePtr, $fuSound)
			GUISetState(@SW_RESTORE, $hGUI)
		EndIf
		If _IsPressed("7B", $dll) Then ;F12
			GUISetState(@SW_MINIMIZE, $hGUI)
			Sleep(250)
			Grab_Region($aFullScreen[0], $aFullScreen[1], $aFullScreen[2], $aFullScreen[3])
			sndPlaySound($BytePtr, $fuSound)
			GUISetState(@SW_RESTORE, $hGUI)
		EndIf
		If _IsPressed("6B", $dll) Then Zoom(1) ;Numpad +
		If _IsPressed("6D", $dll) Then Zoom(0) ;Numpad -
		If _IsPressed("68", $dll) Then ;Numpad 8
			$scroll_y += $scroll_speedy
			Draw2Graphic($hBmp)
			Zoom(2)
		EndIf
		If _IsPressed("62", $dll) Then ;Numpad 2
			$scroll_y -= $scroll_speedy
			Draw2Graphic($hBmp)
			Zoom(2)
		EndIf
		If _IsPressed("64", $dll) Then ;Numpad 4
			$scroll_x += $scroll_speedx
			Draw2Graphic($hBmp)
			Zoom(2)
		EndIf
		If _IsPressed("66", $dll) Then ;Numpad 6
			$scroll_x -= $scroll_speedx
			Draw2Graphic($hBmp)
			Zoom(2)
		EndIf
		$mc = GUIGetCursorInfo($hGUI)
		If Not @error Then
			If _IsPressed("01", $dll) And $mc[4] = $Gfx Then
				$mpointer = MouseGetCursor()
				GUICtrlSetCursor($Gfx, 9)
				$mc2 = MouseGetPos()
				While _IsPressed("01", $dll)
					$mc3 = MouseGetPos()
					If $mc2[0] <> $mc3[0] Or $mc2[1] <> $mc3[1] Then
						$scroll_x = $sx + $mc3[0] - $mc2[0]
						$scroll_y = $sy + $mc3[1] - $mc2[1]
						Draw2Graphic($hBmp)
						Zoom(2)
					EndIf
					Sleep(50)
				WEnd
				GUICtrlSetCursor($Gfx, $mpointer)
			EndIf
		EndIf
		$sx = $scroll_x
		$sy = $scroll_y
	EndIf
;~ 	If _IsPressed("11", $dll) And _IsPressed("77", $dll) Then GrabPrintScreenClipboard() ;Ctrl + F8 -> Printscreen
;~ 	If _IsPressed("11", $dll) And _IsPressed("78", $dll) Then GrabPrintScreenClipboard(True) ;Ctrl + F9 -> Alt+Printscreen
	If _IsPressed("2C", $dll) Then GrabPrintScreenClipboard() ;capture Printscreen key press
WEnd

#region Some parts for displaying preview GDI+ screen
Func Redraw($hWnd, $Msg, $wParam, $lParam)
	#forceref $hWnd, $Msg, $wParam, $lParam
	_WinAPI_RedrawWindow($hGUI, 0, 0, $RDW_UPDATENOW + $RDW_NOINTERNALPAINT)
	_GDIPlus_GraphicsDrawImageRect($hGraphic, $hBuffer_Bmp, $gdip_x, $gdip_y, $gdip_w, $gdip_h)
	$NCPaint = True
	Return "GUI_RUNDEFMSG"
EndFunc   ;==>Redraw

Func Redraw_XP($hWnd, $Msg, $wParam, $lParam)
	#forceref $hWnd, $Msg, $wParam, $lParam
	_WinAPI_RedrawWindow($hGUI, "", "", BitOR($RDW_INVALIDATE, $RDW_UPDATENOW))
	Draw2Graphic($hBmp)
	Zoom(2)
	$NCPaint = True
	Return "GUI_RUNDEFMSG"
EndFunc   ;==>Redraw_XP

Func Zoom($zoom_dir)
	Switch $zoom_dir
		Case -1
			_GDIPlus_MatrixDispose($hMatrix)
			$hMatrix = _GDIPlus_MatrixCreate()
			_GDIPlus_MatrixTranslate($hMatrix, $gdip_w / 2, $gdip_h / 2)
			_GDIPlus_MatrixScale($hMatrix, 1 / $f, 1 / $f)
			$scroll_x = 0
			$scroll_y = 0
		Case 0
			If $z > $zmin Then
				_GDIPlus_MatrixScale($hMatrix, 0.95, 0.95)
				$z -= 0.05
			EndIf
		Case 1
			If $z <= $zmax Then
				_GDIPlus_MatrixScale($hMatrix, 1.05, 1.05)
				$z += 0.05
			EndIf
		Case 2
			_GDIPlus_MatrixScale($hMatrix, 1, 1)
	EndSwitch
	_GDIPlus_GraphicsSetTransform($hContext, $hMatrix)
	_GDIPlus_GraphicsClear($hContext, "0xFF" & $bg_c)
	_GDIPlus_GraphicsDrawImage($hContext, $hBmp, -$bW / 2 + $scroll_x, -$BH / 2 + $scroll_y)
	_GDIPlus_GraphicsDrawImageRect($hGraphic, $hBuffer_Bmp, $gdip_x, $gdip_y, $gdip_w, $gdip_h)
EndFunc   ;==>Zoom

Func Draw2Graphic($hImage)
	Local $w, $h
	_GDIPlus_GraphicsClear($hContext, "0xFF" & $bg_c)
	If $bW <= $gdip_w And $BH <= $gdip_h Then
		$f = 1
		_GDIPlus_GraphicsDrawImageRect($hContext, $hImage, $gdip_w / 2 - $bW / 2 - $scroll_x, $gdip_h / 2 - $BH / 2 - $scroll_y, $bW, $BH)
	Else
		If $bW > $BH Then
			$f = $bW / $gdip_w
		Else
			$f = $BH / $gdip_h
		EndIf
		$w = Int($bW / $f)
		$h = Int($BH / $f)
		_GDIPlus_GraphicsDrawImageRect($hContext, $hImage, $gdip_w / 2 - $w / 2 - $scroll_x, $gdip_h / 2 - $h / 2 - $scroll_y, $w, $h)
	EndIf
EndFunc   ;==>Draw2Graphic

Func Undo()
	$hBmp = $undo
	$hClipboard_Bitmap = _WinAPI_CopyImage(_GDIPlus_BitmapCreateHBITMAPFromBitmap($hBmp), 0, 0, 0, $LR_COPYDELETEORG + $LR_COPYRETURNORG)
	$bW = _GDIPlus_ImageGetWidth($undo)
	$BH = _GDIPlus_ImageGetHeight($undo)
	Draw2Graphic($hBmp)
EndFunc   ;==>Undo

Func FlipImage($iRotateFlipType)
	Local $iWidth = _GDIPlus_ImageGetWidth($hBmp)
	Local $iHeight = _GDIPlus_ImageGetHeight($hBmp)
	$undo = _GDIPlus_BitmapCloneArea($hBmp, 0, 0, $iWidth, $iHeight)
	$undo_chk = True
	DllCall($ghGDIPDll, "uint", "GdipImageRotateFlip", "hwnd", $hBmp, "int", $iRotateFlipType)
	$hClipboard_Bitmap = _WinAPI_CopyImage(_GDIPlus_BitmapCreateHBITMAPFromBitmap($hBmp), 0, 0, 0, $LR_COPYDELETEORG + $LR_COPYRETURNORG)
	Local $tmp = $bW
	$bW = $BH
	$BH = $tmp
	Draw2Graphic($hBmp)
EndFunc   ;==>FlipImage

Func Load_BMP_From_Mem($bImage, $hHBITMAP = False)
	If Not IsBinary($bImage) Then Return SetError(1, 0, 0)
	Local $declared = True
	If Not $ghGDIPDll Then
		_GDIPlus_Startup()
		$declared = False
	EndIf
	Local $aResult
	Local Const $memBitmap = Binary($bImage) ;load image  saved in variable (memory) and convert it to binary
	Local Const $len = BinaryLen($memBitmap) ;get length of image
	Local Const $hData = _MemGlobalAlloc($len, $GMEM_MOVEABLE) ;allocates movable memory  ($GMEM_MOVEABLE = 0x0002)
	Local Const $pData = _MemGlobalLock($hData) ;translate the handle into a pointer
	Local $tMem = DllStructCreate("byte[" & $len & "]", $pData) ;create struct
	DllStructSetData($tMem, 1, $memBitmap) ;fill struct with image data
	_MemGlobalUnlock($hData) ;decrements the lock count  associated with a memory object that was allocated with GMEM_MOVEABLE
	$aResult = DllCall("ole32.dll", "int", "CreateStreamOnHGlobal", "handle", $pData, "int", True, "ptr*", 0) ;Creates a stream object that uses an HGLOBAL memory handle to store the stream contents
	If @error Then SetError(2, 0, 0)
	Local Const $hStream = $aResult[3]
	$aResult = DllCall($ghGDIPDll, "uint", "GdipCreateBitmapFromStream", "ptr", $hStream, "int*", 0) ;Creates a Bitmap object based on an IStream COM interface
	If @error Then SetError(3, 0, 0)
	Local Const $hBitmap = $aResult[2]
	Local $tVARIANT = DllStructCreate("word vt;word r1;word r2;word r3;ptr data; ptr")
	DllCall("oleaut32.dll", "long", "DispCallFunc", "ptr", $hStream, "dword", 8 + 8 * @AutoItX64, _
			"dword", 4, "dword", 23, "dword", 0, "ptr", 0, "ptr", 0, "ptr", DllStructGetPtr($tVARIANT)) ;release memory from $hStream to avoid memory leak
	$tMem = 0
	$tVARIANT = 0
	If $hHBITMAP Then
		Local Const $hHBmp = _GDIPlus_BitmapCreateHBITMAPFromBitmap($hBitmap)
		_GDIPlus_BitmapDispose($hBitmap)
		If Not $declared Then _GDIPlus_Shutdown()
		Return $hHBmp
	EndIf
	If Not $declared Then _GDIPlus_Shutdown()
	Return $hBitmap
EndFunc   ;==>Load_BMP_From_Mem
#endregion Some parts for displaying preview GDI+ screen
#region Grab2AVI
Func Grab2AVI()
	GUISetState(@SW_HIDE, $hGUI)
	Local Const $hGUI_AU3 = WinGetHandle(AutoItWinGetTitle())

	Local Const $hGUI_Cross_AVI = GUICreate("", 30, 30, 0, 0, $WS_POPUP, $WS_DISABLED + $WS_EX_TOOLWINDOW + $WS_EX_TOPMOST, $hGUI_AU3)
	WinSetTrans($hGUI_Cross_AVI, "", 1)
	GUISetState(@SW_HIDE, $hGUI_Cross_AVI)

	Local Const $hGUI_Grab2AVI = GUICreate("", 0, 0, 0, 0, $WS_POPUP, $WS_EX_TOPMOST + $WS_EX_TOOLWINDOW, $hGUI_AU3)
	GUISetBkColor(0xFF0000, $hGUI_Grab2AVI)
	GUISetState(@SW_SHOW, $hGUI_Grab2AVI)
	Local $aMPos, $hWin, $hWinAncestor, $hWnd, $aRet_prev, $aPos, $q, $fps, $hBmp_AVI
	Local Const $frame_size = 3
	Local $tPoint = DllStructCreate($tagPOINT)
	Local $esc = True, $rem1 = 0, $rem2 = 0
	$aLastPos[0] = -1
	$aLastPos[1] = -1

	Local $m_startx, $m_starty

	Local Const $hGUI_AVI_Mark = GUICreate("", 0, 0, $m_startx, $m_starty, $WS_POPUP, $WS_EX_TOPMOST, $hGUI_AU3)
	GUISetBkColor(0xA00000)
	WinSetTrans($hGUI_AVI_Mark, "", 0x40)
	GUISetState(@SW_HIDE, $hGUI_AVI_Mark)
	Local $w, $h, $size = 1
	Local $label, $label_text = ""
	$label = GUICtrlCreateLabel($label_text, $size, $size, 0, 0)
	GUICtrlSetFont($label, 6)
	GUICtrlSetBkColor($label, 0xFFD8C8)

	Local $iX1, $iY1

	While Not _IsPressed("1B", $dll) * Sleep(25)
		If Not $rem1 Then
			GUISetState(@SW_HIDE, $hGUI_AVI_Mark)
			GUISetState(@SW_HIDE, $hGUI_Cross_AVI)
			GUISetState(@SW_SHOW, $hGUI_Grab2AVI)
			$rem1 = 1
			$rem2 = 0
			$aLastPos[0] = -1
			$aLastPos[1] = -1
			$w = 0
			$h = 0
		EndIf
		$aMPos = MouseGetPos()
		DllStructSetData($tPoint, 1, $aMPos[0])
		DllStructSetData($tPoint, 2, $aMPos[1])
		$hWin = _WinAPI_WindowFromPoint($tPoint)

;~ 		$hWinAncestor = _WinAPI_GetAncestor($hWin, 2)
;~ 		$hWnd = HWnd($hWinAncestor)
		$hWnd = HWnd($hWin)
		$aRet_prev = -1
		$aPos = WinGetPos($hWnd)
		If $hWnd <> $hGUI_Grab2AVI And $hWnd <> $aRet_prev And $hWnd <> $hGUI_Cross_AVI And $hWnd <> $hGUI_AVI_Mark Then
			$aRet_prev = $hWnd
			If $aMPos[0] <> $aLastPos[0] Or $aMPos[1] <> $aLastPos[1] Then
				If _WinAPI_GetClassName($hWin) <> "tooltips_class32" Then
					WinMove($hGUI_Grab2AVI, "", $aPos[0], $aPos[1], $aPos[2], $aPos[3], 1)
					_GuiHole($hGUI_Grab2AVI, $frame_size, $frame_size, $aPos[2] - 2 * $frame_size, $aPos[3] - 2 * $frame_size, $aPos[2], $aPos[3])
					WinSetOnTop($hGUI_Grab2AVI, 0, 1)
				EndIf
				ToolTip(	"Press CTRL to start capturing of marked window!" & @LF & _
								"Or press SHIFT to switch to mark desktop area manually" & @LF & _
								"Windows Size: " & $aPos[2] & " x " & $aPos[3], $aMPos[0] + 20, $aMPos[1] + 20)
				$aLastPos = $aMPos
			EndIf
		EndIf

		While _IsPressed("10", $dll) * Sleep(25)
			If Not $rem2 Then
				GUISetState(@SW_HIDE, $hGUI_Grab2AVI)
				GUISetState(@SW_SHOW, $hGUI_AVI_Mark)
				GUISetState(@SW_SHOW, $hGUI_Cross_AVI)
				$rem2 = 1
				$rem1 = 0
				$aLastPos[0] = -1
				$aLastPos[1] = -1
			EndIf
			$aMPos = MouseGetPos()
			If $aMPos[0] <> $aLastPos[0] Or $aMPos[1] <> $aLastPos[1] Then
				ToolTip("Mark area on desktop by holding lmb down!" & @LF & "Release SHIFT key to switch back")
				WinMove($hGUI_Cross_AVI, "", $aMPos[0] - 15, $aMPos[1] - 15, 30, 30, 1)
				$aLastPos = $aMPos
			EndIf
			GUISetCursor(3, 1, $hGUI_Cross_AVI)
			$m_startx = $aMPos[0]
			$m_starty = $aMPos[1]
			While _IsPressed("01", $dll) * Sleep(10)
				GUISetCursor(3, 1, $hGUI_Cross_AVI)
				$aMPos = MouseGetPos()
				If $aMPos[0] <> $aLastPos[0] Or $aMPos[1] <> $aLastPos[1] Then
					$w = Abs($aMPos[0] - $m_startx) + 1
					$h = Abs($aMPos[1] - $m_starty) + 1
					If $aMPos[0] < $m_startx Then
						$iX1 = $aMPos[0]
					Else
						$iX1 = $m_startx
					EndIf
					If $aMPos[1] < $m_starty Then
						$iY1 = $aMPos[1]
					Else
						$iY1 = $m_starty
					EndIf
					WinMove($hGUI_Cross_AVI, "", $aMPos[0] - 15, $aMPos[1] - 15, 30, 30, 1)
					WinMove($hGUI_AVI_Mark, "", $iX1, $iY1, $w, $h)
					$aLastPos = $aMPos
					ToolTip("Size: " & $w & "x" & $h & @CRLF & "Position: " & $aMPos[0] & "," & $aMPos[1], $aMPos[0] + 10, $aMPos[1] + 10)
					GUICtrlSetPos($label, $size, $size, $w - $size * 2, $h - $size * 2)
				EndIf
			WEnd
			If $w > 1 And $h > 1 Then
				$esc = False
				$aPos = WinGetPos($hGUI_AVI_Mark)
				ExitLoop 2
			EndIf
		WEnd

		If _IsPressed("11", $dll) Then
			$esc = False
			ExitLoop
		EndIf
	WEnd
	ToolTip("")
	$tPoint = 0
	GUIDelete($hGUI_AVI_Mark)
	GUIDelete($hGUI_Grab2AVI)
	GUIDelete($hGUI_Cross_AVI)

	If Not $esc Then
		For $q = 0 To UBound($aButton_Menu_AVI_Sub1_Item) - 1
			If BitAND(GUICtrlRead($aButton_Menu_AVI_Sub1_Item[$q]), $GUI_CHECKED) = $GUI_CHECKED Then
				$rec_time = $AVI_ReqTime[$q]
				ExitLoop
			EndIf
		Next
		For $q = 0 To UBound($aButton_Menu_AVI_Sub2_Item) - 1
			If BitAND(GUICtrlRead($aButton_Menu_AVI_Sub2_Item[$q]), $GUI_CHECKED) = $GUI_CHECKED Then
				$fps = $AVI_FPS[$q]
				ExitLoop
			EndIf
		Next
	EndIf

	If $esc Then Return GUISetState(@SW_SHOW, $hGUI)

	Local $timestamp = @YEAR & @MON & @MDAY & "_" & @HOUR & @MIN & @SEC & "_"
	Local $WinTitle = StringRegExpReplace(WinGetTitle(_WinAPI_GetAncestor($hWnd, $GA_ROOT)), '(\?|\*|\<|\>|\"|\:|\\)', "")

	Local $add_ts = False
	If BitAND(GUICtrlRead($mButton_Menu_AVI_Sub3_Item1), $GUI_CHECKED) = $GUI_CHECKED Then $add_ts = True
	If BitAND(GUICtrlRead($mButton_Menu_AVI_Sub4_Item2), $GUI_CHECKED) = $GUI_CHECKED Then
		$compress_avi = False
	Else
		$compress_avi = True
	EndIf

	_StartAviLibrary()
	Local Const $AVI_Filename = @ScriptDir & "\" & $timestamp & $WinTitle & ".avi"
	Local $AVI_File = _CreateAvi($AVI_Filename, $fps, $aPos[2], $aPos[3])
	If @error Then
		_CloseAvi($AVI_File)
		_StopAviLibrary()
		FileDelete($AVI_Filename)
		GUISetState(@SW_SHOW, $hGUI)
		Return MsgBox(64 + 262144, "Information", "Grab Screen to AVI has been aborted!", 15, $hGUI)
	EndIf

	Local $hBitmap_AVI, $hBitmap_AVI_TS, $hBmp_AVI_TS, $OldBMP
	Local $iLeft, $iRight, $iTop, $iBottom
	Local $hIcon, $aIcon
	Local $tCursor, $tInfo, $iCursor, $aCursor[5], $aIcon[6]
	Local $total_FPS = $rec_time * $fps, $fps_c = 1
	Local Const $fc = 1000 / $fps

	Local Const $k32_dll = DllOpen("kernel32.dll")
	Local Const $g32_dll = DllOpen("gdi32.dll")
	Local Const $u32_dll = DllOpen("user32.dll")
	Local Const $DC = _WinAPI_GetDC(0)
	Local Const $hDC = _WinAPI_CreateCompatibleDC($DC)

	Local $bits = DllStructCreate("byte[" & DllStructGetData($AVI_File[3], "biSizeImage") & "]")
	Local Const $pBits = DllStructGetPtr($bits), $iLines = Abs(DllStructGetData($AVI_File[3], "biHeight"))
	Local Const $pHeader = DllStructGetPtr($AVI_File[3]), $iSize = DllStructGetSize($bits)

	$iLeft = $aPos[0]
	$iRight = $aPos[0] + $aPos[2]
	$iTop = $aPos[1]
	$iBottom = $aPos[1] + $aPos[3]
	If $iRight = -1 Then $iRight = _WinAPI_GetSystemMetrics($__SCREENCAPTURECONSTANT_SM_CXSCREEN)
	If $iBottom = -1 Then $iBottom = _WinAPI_GetSystemMetrics($__SCREENCAPTURECONSTANT_SM_CYSCREEN)
	Local $iW = $aPos[2]
	Local $iH = $aPos[3]
	$hBmp_AVI = _WinAPI_CreateCompatibleBitmap($DC, $iW, $iH)

	If BitAND(GUICtrlRead($mButton_Menu_AVI_Sub5_Item1), $GUI_CHECKED) = $GUI_CHECKED Then
		$cursor = True
	Else
		$cursor = False
	EndIf

	Local $t, $td, $rec_timer = TimerInit()
	If $add_ts Then
		Do
			$t = TimerInit()
			$OldBMP = DllCall($g32_dll, "handle", "SelectObject", "handle", $hDC, "handle", $hBmp_AVI) ;_WinAPI_SelectObject()
			DllCall($g32_dll, "bool", "BitBlt", "handle", $hDC, "int", 0, "int", 0, "int", $iW, "int", $iH, "handle", $DC, "int", $iLeft, "int", $iTop, "dword", $SRCCOPY) ;_WinAPI_BitBlt()

			If $cursor Then
				$tCursor = DllStructCreate($tagCURSORINFO)
				$iCursor = DllStructGetSize($tCursor)
				DllStructSetData($tCursor, "Size", $iCursor)
				DllCall($u32_dll, "bool", "GetCursorInfo", "ptr", DllStructGetPtr($tCursor))
				$aCursor[1] = DllStructGetData($tCursor, "Flags") <> 0
				$aCursor[2] = DllStructGetData($tCursor, "hCursor")
				$aCursor[3] = DllStructGetData($tCursor, "X")
				$aCursor[4] = DllStructGetData($tCursor, "Y")
				If $aCursor[1] Then
					$hIcon = DllCall($u32_dll, "handle", "CopyIcon", "handle", $aCursor[2])
					$tInfo = DllStructCreate($tagICONINFO)
					DllCall($u32_dll, "bool", "GetIconInfo", "handle", $hIcon[0], "ptr", DllStructGetPtr($tInfo))
					$aIcon[2] = DllStructGetData($tInfo, "XHotSpot")
					$aIcon[3] = DllStructGetData($tInfo, "YHotSpot")
					$aIcon[4] = DllStructGetData($tInfo, "hMask")
					DllCall($g32_dll, "bool", "DeleteObject", "handle", $aIcon[4])
					DllCall($u32_dll, "bool", "DrawIcon", "handle", $hDC, "int", $aCursor[3] - $aIcon[2] - $iLeft, "int", $aCursor[4] - $aIcon[3] - $iTop, "handle", $hIcon[0])
					DllCall($u32_dll, "bool", "DestroyIcon", "handle", $hIcon[0])
				EndIf
			EndIf

			$hBitmap_AVI = DllCall($ghGDIPDll, "int", "GdipCreateBitmapFromHBITMAP", "handle", $hBmp_AVI, "handle", 0, "ptr*", 0) ;_GDIPlus_BitmapCreateFromHBITMAP()
			$hBitmap_AVI_TS = WTOB($hBitmap_AVI[3], _Now())
			$hBmp_AVI_TS = DllCall($ghGDIPDll, "int", "GdipCreateHBITMAPFromBitmap", "handle", $hBitmap_AVI_TS, "ptr*", 0, "dword", 0xFF000000) ;_GDIPlus_BitmapCreateHBITMAPFromBitmap()

			DllCall($g32_dll, "int", "GetDIBits", "handle", $hDC, "handle", $hBmp_AVI_TS[2], "uint", 0, "uint", $iLines, "ptr", $pBits, "ptr", $pHeader, "uint", 0) ;_WinAPI_GetDIBits()
			DllCall($g32_dll, "handle", "SelectObject", "handle", $hDC, "handle", $OldBMP[0]) ;_WinAPI_SelectObject()
			DllCall($Avi32_Dll, "int", "AVIStreamWrite", "ptr", $AVI_File[1], "long", $AVI_File[2], "long", 1, "ptr", $pBits, "long", $iSize, "long", $AVIIF_KEYFRAME, "ptr*", 0, "ptr*", 0)
			$AVI_File[2] += 1

			DllCall($k32_dll, "none", "Sleep", "dword", $fc)

			DllCall($g32_dll, "bool", "DeleteObject", "handle", $hBmp_AVI_TS[2]) ;_WinAPI_DeleteObject()
			DllCall($ghGDIPDll, "int", "GdipDisposeImage", "handle", $hBitmap_AVI[3]) ;_GDIPlus_BitmapDispose()
			DllCall($ghGDIPDll, "int", "GdipDisposeImage", "handle", $hBitmap_AVI_TS) ;_GDIPlus_BitmapDispose()

			$fps_c += 1

			$td = $fc - TimerDiff($t)
			If $td > 0 Then DllCall($k32_dll, "none", "Sleep", "dword", $td)

		Until $fps_c > $total_FPS
	Else
		Do
			$t = TimerInit()
			;_ScreenCapture_Capture() reduced
			$OldBMP = DllCall($g32_dll, "handle", "SelectObject", "handle", $hDC, "handle", $hBmp_AVI) ;_WinAPI_SelectObject()
			DllCall($g32_dll, "bool", "BitBlt", "handle", $hDC, "int", 0, "int", 0, "int", $iW, "int", $iH, "handle", $DC, "int", $iLeft, "int", $iTop, "dword", $SRCCOPY) ;_WinAPI_BitBlt()
			If $cursor Then
				$tCursor = DllStructCreate($tagCURSORINFO)
				$iCursor = DllStructGetSize($tCursor)
				DllStructSetData($tCursor, "Size", $iCursor)
				DllCall($u32_dll, "bool", "GetCursorInfo", "ptr", DllStructGetPtr($tCursor))
				$aCursor[1] = DllStructGetData($tCursor, "Flags") <> 0
				$aCursor[2] = DllStructGetData($tCursor, "hCursor")
				$aCursor[3] = DllStructGetData($tCursor, "X")
				$aCursor[4] = DllStructGetData($tCursor, "Y")
				If $aCursor[1] Then
					$hIcon = DllCall($u32_dll, "handle", "CopyIcon", "handle", $aCursor[2])
					$tInfo = DllStructCreate($tagICONINFO)
					DllCall($u32_dll, "bool", "GetIconInfo", "handle", $hIcon[0], "ptr", DllStructGetPtr($tInfo))
					$aIcon[2] = DllStructGetData($tInfo, "XHotSpot")
					$aIcon[3] = DllStructGetData($tInfo, "YHotSpot")
					$aIcon[4] = DllStructGetData($tInfo, "hMask")
					DllCall($g32_dll, "bool", "DeleteObject", "handle", $aIcon[4])
					DllCall($u32_dll, "bool", "DrawIcon", "handle", $hDC, "int", $aCursor[3] - $aIcon[2] - $iLeft, "int", $aCursor[4] - $aIcon[3] - $iTop, "handle", $hIcon[0])
					DllCall($u32_dll, "bool", "DestroyIcon", "handle", $hIcon[0])
				EndIf
			EndIf
			DllCall($g32_dll, "int", "GetDIBits", "handle", $hDC, "handle", $hBmp_AVI, "uint", 0, "uint", $iLines, "ptr", $pBits, "ptr", $pHeader, "uint", 0) ;_WinAPI_GetDIBits()
			DllCall($g32_dll, "handle", "SelectObject", "handle", $hDC, "handle", $OldBMP[0]) ;_WinAPI_SelectObject()
			DllCall($Avi32_Dll, "int", "AVIStreamWrite", "ptr", $AVI_File[1], "long", $AVI_File[2], "long", 1, "ptr", $pBits, "long", $iSize, "long", $AVIIF_KEYFRAME, "ptr*", 0, "ptr*", 0)
			$AVI_File[2] += 1

			$fps_c += 1

			$td = $fc - TimerDiff($t)
			If $td > 0 Then DllCall($k32_dll, "none", "Sleep", "dword", $td)

		Until $fps_c > $total_FPS
	EndIf
	Local $rec_dur = Round(TimerDiff($rec_timer) / 1000, 2)
	_GDIPlus_BitmapDispose($hBmp)
	$hBmp = _GDIPlus_BitmapCreateFromHBITMAP($hBmp_AVI)
	Update_Resolution_Info($hBmp)
	Draw2Graphic($hBmp)
	$hClipboard_Bitmap = $hBitmap_s
	$z = 10
	Zoom(-1)
	_WinAPI_DeleteObject($hBmp_AVI)
	_WinAPI_SelectObject($DC, $OldBMP[0])
	_WinAPI_ReleaseDC(0, $DC)
	_WinAPI_DeleteDC($hDC)

	_CloseAvi($AVI_File)
	_StopAviLibrary()
	DllClose($k32_dll)
	DllClose($g32_dll)
	DllClose($u32_dll)
	$bits = 0
	$tInfo = 0
	$tCursor = 0
	MsgBox(64 + 262144, "Information", _
			"AVI video successfully created: " & @LF & @LF & _
			"Filename: " & $AVI_Filename & @LF & _
			"Filesize: " & Round(FileGetSize($AVI_Filename) / 1024 ^ 2, 2) & " MB" & @LF & _
			"Duration: " & $rec_time & " seconds" & @LF & _
			"Frames per second: " & $fps & @LF & _
			"Video resolution: " & $aPos[2] & " x " & $aPos[3] & @LF & _
			"Record duration: " & $rec_dur & " seconds", 30, $hGUI)
	GUISetState(@SW_SHOW, $hGUI)

	Local $show_avi = MsgBox(4 + 48 + 256, "Question", "Open AVI with default viewer?", 15, $hGUI)
	If $show_avi = 6 Then ShellExecute($AVI_Filename)
EndFunc   ;==>Grab2AVI
#endregion Grab2AVI
#region Screen Grab
Func Mark_Area($disable_aero = False)
	;http://msdn.microsoft.com/en-us/library/aa969510%28VS.85%29.aspx
	Local $hGUI_AU3 = WinGetHandle(AutoItWinGetTitle())
	Local Const $hDwmApiDll = DllOpen("dwmapi.dll")
	Local $sChkAero = DllStructCreate("int;")
	DllCall($hDwmApiDll, "int", "DwmIsCompositionEnabled", "ptr", DllStructGetPtr($sChkAero))
	Local $aero_on = DllStructGetData($sChkAero, 1)
	If $aero_on And $disable_aero Then DllCall($hDwmApiDll, "int", "DwmEnableComposition", "uint", False)
	$sChkAero = 0

	Local $mpos = MouseGetPos()
;~ 	Local $hGUI_Cross = GUICreate("", 30, 30, $mpos[0], $mpos[1], $WS_POPUP, $WS_DISABLED + $WS_EX_TOOLWINDOW + $WS_EX_TOPMOST + $WS_EX_LAYERED, WinGetHandle(AutoItWinGetTitle()))
	Local $hGUI_Cross = GUICreate("", 30, 30, $mpos[0], $mpos[1], $WS_POPUP, $WS_DISABLED + $WS_EX_TOOLWINDOW + $WS_EX_TOPMOST, $hGUI_AU3)
	$hGUI_Dot = GUICreate("", 1, 1, 1, 1, $WS_POPUP, $WS_EX_TOPMOST + $WS_EX_MDICHILD, $hGUI_Cross) ; WinGetHandle(AutoItWinGetTitle()))
	GUISetBkColor(0, $hGUI_Dot)
	GUISetBkColor(0, $hGUI_Cross)
	WinSetTrans($hGUI_Cross, "", 1)
	GUISetState(@SW_SHOW, $hGUI_Cross)
	GUISetState(@SW_SHOW, $hGUI_Dot)
;~ 	_WinAPI_SetLayeredWindowAttributes($hGUI_Cross, 0x00, 0xFF)

	Local $zl
	$hGUI_Zoom = GUICreate("", $zw, $zh, $zx, $zx, BitOR($WS_POPUP, $DS_MODALFRAME), BitOR($WS_EX_OVERLAPPEDWINDOW, $WS_EX_TOPMOST, $WS_EX_WINDOWEDGE), $hGUI_AU3)
	GUISetState(@SW_SHOW, $hGUI_Zoom)

	$hDC_Zoom = _WinAPI_GetDC(0)
	$hGUI_ZoomDC = _WinAPI_GetDC($hGUI_Zoom)
	$zx = 0
	$zy = 0
	$up = 1

	WinMove($hGUI_Zoom, "", $zx, $zy)

	Local $esc = False, $pixel_color, $v
	Local Const $dist_x = 24, $dist_y = 24

	Local $aRet, $aRet_prev, $area, $ctrl_handle, $ctrl_id, $Style
	Local $hGUI_Mark = GUICreate("", 0, 0, 0, 0, $WS_POPUP, $WS_EX_TOPMOST + $WS_EX_COMPOSITED + $WS_EX_TOOLWINDOW, $hGUI_AU3)
	GUISetBkColor(0xFF0000, $hGUI_Mark)
	GUISetState(@SW_HIDE, $hGUI_Mark)
	Local Const $frame_size = 3

	Local $freehand = False
	$hGUI_Freehand = GUICreate("", $aFullScreen[2], $aFullScreen[3], $aFullScreen[0], $aFullScreen[1], $WS_POPUP, $WS_EX_LAYERED + $WS_EX_TOPMOST, $hGUI_AU3)
	GUISetBkColor(0xABCDEF)
	GUISetState(@SW_HIDE, $hGUI_Freehand)
	_WinAPI_SetLayeredWindowAttributes($hGUI_Freehand, 0xABCDEF, 0xFF)

	GUIRegisterMsg(0x020A, "WM_MOUSEWHEEL")
	Local $tdiff = TimerInit()
	While Not _IsPressed("01", $dll) * Sleep(20)
		If _IsPressed("11", $dll) Then ;control capturing
			ToolTip("")
			If _WinAPI_IsWindowVisible($hGUI_Zoom) Then GUISetState(@SW_HIDE, $hGUI_Zoom)
			If _WinAPI_IsWindowVisible($hGUI_Cross) Then GUISetState(@SW_HIDE, $hGUI_Cross)
			If _WinAPI_IsWindowVisible($hGUI_Dot) Then GUISetState(@SW_HIDE, $hGUI_Dot)
			GUISetState(@SW_SHOW, $hGUI_Mark)
			$mpos = MouseGetPos()
			$aRet = DllCall($dll, "int", "WindowFromPoint", "long", $mpos[0], "long", $mpos[1])
			$aPos_SC = WinGetPos(HWnd($aRet[0]))
;~ 			$ctrl_id = _WinAPI_GetDlgCtrlID(HWnd($aRet[0]))
;~ 			$ctrl_handle = ControlGetHandle(HWnd($aRet[0]), "", $ctrl_id)
;~ 			$Style = _WinAPI_GetWindowLong(HWnd($aRet[0]), $GWL_STYLE)
;~ 			If BitAND($Style, $WS_VSCROLL) = $WS_VSCROLL Then ConsoleWrite(_WinAPI_GetClassName(HWnd($aRet[0])) & " Have a Vertical Scroll" & @CRLF)

;~ 			ConsoleWrite(_GUIScrollBars_GetScrollRange(HWnd($aRet[0]), $SB_VERT) & @CRLF)
			If $aRet[0] <> $hGUI_Mark And $aRet[0] <> 0 And $aRet[0] <> $aRet_prev And $aRet[0] <> $hGUI_Cross Then
				$aRet_prev = $aRet[0]
				WinMove($hGUI_Mark, "", $aPos_SC[0], $aPos_SC[1], $aPos_SC[2], $aPos_SC[3])
				_GuiHole($hGUI_Mark, $frame_size, $frame_size, $aPos_SC[2] - 2 * $frame_size, $aPos_SC[3] - 2 * $frame_size, $aPos_SC[2], $aPos_SC[3])
				WinSetOnTop($hGUI_Mark, 0, 1)
			EndIf
			$area = True
			If _IsPressed("10", $dll) Then ExitLoop ;shift to make the screenshot
			If _IsPressed("02", $dll) And TimerDiff($tdiff) > 500 Then
				MouseClick("left") ;rmb to send lmb click to activate a GUI menu which can be capture afterwards
				$tdiff = TimerInit()
			EndIf
		ElseIf _IsPressed("10", $dll) Then ;freehand screen capturing with shift key
			$zl = $zoom_level
			ToolTip("")
			$area = False
			$freehand = True
			$esc = True
			GUISetState(@SW_HIDE, $hGUI_Cross)
			GUISetState(@SW_HIDE, $hGUI_Dot)
			GUISetState(@SW_HIDE, $hGUI_Zoom)
			GUISetState(@SW_HIDE, $hGUI_Mark)

			Local $hHBITMAP_Freehand = _ScreenCapture_Capture("", $aFullScreen[0], $aFullScreen[1], $aFullScreen[2], $aFullScreen[3], False)
			$hBmp_Freehand = _GDIPlus_BitmapCreateFromHBITMAP($hHBITMAP_Freehand)
			_WinAPI_DeleteObject($hHBITMAP_Freehand)

			GUISetState(@SW_SHOW, $hGUI_Dot)
			GUISetState(@SW_SHOW, $hGUI_Freehand)
			GUISetState(@SW_SHOW, $hGUI_Zoom)

			$hGraphic_Freehand = _GDIPlus_GraphicsCreateFromHWND($hGUI_Freehand)
			_GDIPlus_GraphicsSetSmoothingMode($hGraphic_Freehand, 0)
			Local $hPen = _GDIPlus_PenCreate(0xFFFF0000, 1)

			Local $min_x = $aFullScreen[2]
			Local $min_y = $aFullScreen[3]
			Local $max_x = $aFullScreen[0]
			Local $max_y = $aFullScreen[1]
			Local $aPoints[10000][2], $mxo, $myo, $mx, $my
			Local $p = 1, $hide = False
			Local $mo = MouseGetCursor()
			GUISetCursor(14, 1, $hGUI_Freehand)
			GUISetCursor(14, 1, $hGUI_Dot)

			$zoom_level = 4
			If Not $aero Then WinSetTrans($hGUI_Freehand, "", 0x60)
			GUISetCursor(14, 1, $hGUI_Freehand)
			GUISetCursor(14, 1, $hGUI_Dot)
			While _IsPressed("10", $dll) * Sleep(30) ;shift key
				Draw_Zoom_Preview($mpos[0], $mpos[1])
				$mpos = MouseGetPos()
				$mxo = $mpos[0] + Abs($aFullScreen[0])
				$myo = $mpos[1] + Abs($aFullScreen[1])
				WinMove($hGUI_Dot, "", $mpos[0], $mpos[1], 1, 1, 1)
				While _IsPressed("01", $dll) * Sleep(10)
					$mpos = MouseGetPos()
					$mx = $mpos[0] + Abs($aFullScreen[0])
					$my = $mpos[1] + Abs($aFullScreen[1])
					If Not $hide Then
						GUISetState(@SW_HIDE, $hGUI_Dot)
						$hide = True
					EndIf
					_GDIPlus_GraphicsDrawLine($hGraphic_Freehand, $mx, $my, $mxo, $myo, $hPen)
					$mxo = $mx
					$myo = $my
					If $aPoints[$p - 1][0] <> $mx And $aPoints[$p - 1][1] <> $my Then
						$aPoints[$p][0] = $mx
						$aPoints[$p][1] = $my
						$p += 1
					EndIf
					$min_x = Min($min_x, $mx)
					$min_y = Min($min_y, $my)
					$max_x = Max($max_x, $mx)
					$max_y = Max($max_y, $my)
					If $p = UBound($aPoints) Then ExitLoop
					Draw_Zoom_Preview($mpos[0], $mpos[1])
				WEnd
			WEnd
			ReDim $aPoints[$p][2]
			$aPoints[0][0] = $p - 1
			GUISetCursor($mo, 1, $hGUI_Freehand)
			_GDIPlus_GraphicsDispose($hGraphic_Freehand)
			_GDIPlus_PenDispose($hPen)
			GUISetState(@SW_HIDE, $hGUI_Freehand)
			GUISetState(@SW_HIDE, $hGUI_Zoom)
			GUISetState(@SW_HIDE, $hGUI_Dot)
			GUISetState(@SW_HIDE, $hGUI_Cross)
			If $aPoints[0][0] > 2 Then
				Local $hTextureBrush = DllCall($ghGDIPDll, "uint", "GdipCreateTexture", "ptr", $hBmp_Freehand, "int", 0, "int*", 0)
				$hTextureBrush = $hTextureBrush[3]
				Local $hBitmap_Hidden = DllCall($ghGDIPDll, "uint", "GdipCreateBitmapFromScan0", "int", $aFullScreen[2], "int", $aFullScreen[3], "int", 0, "int", 0x0026200A, "ptr", 0, "int*", 0)
				$hBitmap_Hidden = $hBitmap_Hidden[6]
				Local $hContext_Freehand = _GDIPlus_ImageGetGraphicsContext($hBitmap_Hidden)
				_GDIPlus_GraphicsSetSmoothingMode($hContext_Freehand, 2)
				_GDIPlus_GraphicsClear($hContext_Freehand, 0xFFFFFFFF)
				_GDIPlus_GraphicsFillClosedCurve($hContext_Freehand, $aPoints, $hTextureBrush)
				Local $new_w = $max_x - $min_x
				Local $new_h = $max_y - $min_y
				$hBmp = _GDIPlus_BitmapCloneArea($hBitmap_Hidden, $min_x, $min_y, $new_w, $new_h)
				_GDIPlus_GraphicsDispose($hContext_Freehand)
				_GDIPlus_BitmapDispose($hBmp_Freehand)
				_GDIPlus_BitmapDispose($hBitmap_Hidden)
				_GDIPlus_BrushDispose($hTextureBrush)
				Update_Resolution_Info($hBmp)
				Draw2Graphic($hBmp)
				$z = 10
				Zoom(-1)
				sndPlaySound($BytePtr, $fuSound)
				$aPoints = 0
				$zoom_level = $zl
				ExitLoop
			EndIf
			$zoom_level = $zl
			_GDIPlus_BitmapDispose($hBmp_Freehand)
			Sleep(25)
		Else
			$area = False
			$freehand = False
			If Not _WinAPI_IsWindowVisible($hGUI_Cross) Then GUISetState(@SW_SHOW, $hGUI_Cross)
			If Not _WinAPI_IsWindowVisible($hGUI_Dot) Then GUISetState(@SW_SHOW, $hGUI_Dot)
			If Not _WinAPI_IsWindowVisible($hGUI_Zoom) Then GUISetState(@SW_SHOW, $hGUI_Zoom)
			If _WinAPI_IsWindowVisible($hGUI_Mark) Then GUISetState(@SW_HIDE, $hGUI_Mark)
			GUISetCursor(3, 1, $hGUI_Cross)
			GUISetCursor(3, 1, $hGUI_Dot)
			$red = Hex(0xFF * (Cos($v * 1.10) + 1) / 2, 2)
			$green = Hex(0xFF * (Sin($v * 1.00) + 1) / 2, 2)
			$blue = Hex(0xFF * (Sin($v * 1.20) + 1) / 2, 2)
			GUISetBkColor("0x" & $red & $green & $blue, $hGUI_Dot)
			$v += 0.15
;~ 		GUISetBkColor(Random(0x000000, 0xFFFFFF, 1), $hGUI_Dot)
			$mpos = MouseGetPos()
			If $mpos[0] <> $aLastPos[0] Or $mpos[1] <> $aLastPos[1] Then
				$pixel_color = Hex(PixelGetColor($mpos[0], $mpos[1]), 6)
				WinMove($hGUI_Cross, "", $mpos[0] - 15, $mpos[1] - 15, 30, 30, 1)
				WinMove($hGUI_Dot, "", $mpos[0], $mpos[1], 1, 1, 1)
				ToolTip("Mark area on desktop!" & @CRLF & _
						"Hold CTRL to select" & @CRLF & _
						"controls, CTRL+ALT to grab" & @CRLF & _
						"controls, SHIFT for freehand" & @CRLF & _
						"capturing! RMB to get pixel" & @CRLF & _
						"color. ESC to abort" & @CRLF & @CRLF & _
						"Position: " & $mpos[0] & "," & $mpos[1] & @CRLF & _
						"Pixel Color: 0x" & $pixel_color, $drx + $mpos[0] + $dirx * ($dist_x + $zoomW / 3), $dry + $mpos[1] + $diry * ($dist_y + $zoomH / 3))
				$aLastPos = $mpos
				$right_border = Pixel_Distance($mpos[0], $mpos[1], @DesktopWidth, $mpos[1])
				If $right_border < $radius Then
					$dirx = -1
					$drx = -$radius
				Else
					$dirx = 1
					$drx = 0
				EndIf
				$lower_border = Pixel_Distance($mpos[0], $mpos[1], $mpos[0], @DesktopHeight)
				If $lower_border < $radius Then
					$diry = -1
					$dry = -$radius
				Else
					$diry = 1
					$dry = 0
				EndIf
			EndIf
			Draw_Zoom_Preview($mpos[0], $mpos[1])
			If _IsPressed("1B", $dll) Then ;ESC to exit
				$esc = True
				ExitLoop
			EndIf
			If _IsPressed("02", $dll) Then ;copy pixel color to clipboard
				ToolTip("")
				If _ClipBoard_SetData($pixel_color) Then
					MsgBox(64, "Information", "RGB hex color value (" & $pixel_color & ") copied to clipboard!", 20, $hGUI)
				Else
					MsgBox(16, "ERROR", "Cannot copy color value to clipboard!", 20, $hGUI)
				EndIf
			EndIf
		EndIf
	WEnd
	ToolTip("")
	GUIDelete($hGUI_Mark)
	GUIRegisterMsg(0x020A, "")

	If Not $esc Then
		If $area Then
			$Grab_x = $aPos_SC[0]
			$Grab_y = $aPos_SC[1]
			$Grab_w = $aPos_SC[2]
			$Grab_h = $aPos_SC[3]
			Grab_Region($aPos_SC[0], $aPos_SC[1], $aPos_SC[2], $aPos_SC[3])
		ElseIf Not $freehand Then
			$mpos = MouseGetPos()
			Local $m_startx = $mpos[0]
			Local $m_starty = $mpos[1]

			$hGUI_Area = GUICreate("", 0, 0, $m_startx, $m_starty, $WS_POPUP + $WS_BORDER, $WS_EX_TOPMOST + $WS_EX_COMPOSITED);, $hGUI_AU3)
			GUISetBkColor(0xFF8080, $hGUI_Area)
			GUISetState(@SW_SHOW, $hGUI_Area)
			WinSetTrans($hGUI_Area, "", 0x40)

			GUISetState(@SW_HIDE, $hGUI_Cross)
			GUISetState(@SW_SHOW, $hGUI_Dot)
			Local $iX1, $iY1, $selected = False

			GUIRegisterMsg($WM_LBUTTONDOWN, "WM_LBUTTONDOWN")
			GUIRegisterMsg($WM_MOUSEMOVE, "SetCursor")
			GUIRegisterMsg($WM_SIZE, "WM_SIZE")

			Do
				$red = Hex(0xFF * (Cos($v * 1.20) + 1) / 2, 2)
				$green = Hex(0xFF * (Sin($v * 1.10) + 1) / 2, 2)
				$blue = Hex(0xFF * (Sin($v * 1.00) + 1) / 2, 2)
				GUISetBkColor("0x" & $red & $green & $blue, $hGUI_Dot)
				$v += 0.15
	;~ 		GUISetBkColor(Random(0x000000, 0xFFFFFF, 1), $hGUI_Dot)
				$mpos = MouseGetPos()
				If $mpos[0] <> $aLastPos[0] Or $mpos[1] <> $aLastPos[1] Then
					$pixel_color = Hex(PixelGetColor($mpos[0], $mpos[1]), 6)
					WinMove($hGUI_Dot, "", $mpos[0], $mpos[1], 1, 1, 1)
					If Not $selected Then
						$w_area = Abs($mpos[0] - $m_startx) + 1
						$h_area = Abs($mpos[1] - $m_starty) + 1
					EndIf
					If $mpos[0] < $m_startx Then
						$iX1 = $mpos[0]
					Else
						$iX1 = $m_startx
					EndIf
					If $mpos[1] < $m_starty Then
						$iY1 = $mpos[1]
					Else
						$iY1 = $m_starty
					EndIf

					$right_border = Pixel_Distance($mpos[0], $mpos[1], @DesktopWidth, $mpos[1])
					If $right_border < $radius Then
						$drx = -$radius
					Else
						$drx = $radius / 4
					EndIf
					$lower_border = Pixel_Distance($mpos[0], $mpos[1], $mpos[0], @DesktopHeight)
					If $lower_border < $radius Then
						$dry = -$radius
					Else
						$dry = $radius / 4
					EndIf

					If Not $selected Then WinMove($hGUI_Area, "", $iX1, $iY1, $w_area, $h_area)
					$aLastPos = $mpos
					ToolTip("Size: " & $w_area & "x" & $h_area & @CRLF & "Position: " & $mpos[0] & "," & $mpos[1] & @LF & "Press ENTER to grab", $drx + $mpos[0], $dry + $mpos[1])
				EndIf
				Draw_Zoom_Preview($mpos[0], $mpos[1])
				If Not _IsPressed("01", $dll) Then $selected = True
				Sleep(20)
				If _IsPressed("1B", $dll) Then
					$esc = True
					ExitLoop
				EndIf
			Until _IsPressed("0D", $dll) Or _IsPressed("02", $dll)
			GUIRegisterMsg($WM_SIZE, "")
			GUIRegisterMsg($WM_LBUTTONDOWN, "")
			GUIRegisterMsg($WM_MOUSEMOVE, "")
			Local $Win_Coord = WinGetPos($hGUI_Area)
			GUIDelete($hGUI_Area)
			If Not $esc And ($Win_Coord[2] > 0 And $Win_Coord[3] > 0) Then
				$Grab_x = $Win_Coord[0]
				$Grab_y = $Win_Coord[1]
				$Grab_w = $Win_Coord[2]
				$Grab_h = $Win_Coord[3]
				Grab_Region($Win_Coord[0], $Win_Coord[1], $Win_Coord[2], $Win_Coord[3])
			EndIf
		EndIf
	EndIf
	$zoom_level = $zoom_level_default
	_WinAPI_ReleaseDC($hGUI_Zoom, $hGUI_ZoomDC)
	_WinAPI_ReleaseDC(0, $hDC_Zoom)
;~ 	_WinAPI_DeleteDC($hGUI_ZoomDC)
;~ 	_WinAPI_DeleteDC ($hDC_Zoom)

	ToolTip("")
	GUIDelete($hGUI_Freehand)
	GUIDelete($hGUI_Dot)
	GUIDelete($hGUI_Zoom)
	GUIDelete($hGUI_Cross)
	GUIDelete($hGUI_ZoomDC)

	If $aero_on Then
		DllCall($hDwmApiDll, "int", "DwmEnableComposition", "uint", True)
	EndIf
	DllClose($hDwmApiDll)
	$undo_chk = False
EndFunc   ;==>Mark_Area

Func GetMousePosType() ;thanks to martin
    Local $cp = GUIGetCursorInfo($hGUI_Area)
    Local $wp = WinGetPos($hGUI_Area)
	If @error Then Return 0
    Local $side = 0
    Local $TopBot = 0
    Local $curs
    If $cp[0] < $margin Then $side = 1
    If $cp[0] > $wp[2] - $margin Then $side = 2
    If $cp[1] < $margin Then $TopBot = 3
    If $cp[1] > $wp[3] - $margin Then $TopBot = 6
    Return $side + $TopBot
EndFunc ;==>GetMousePosType

Func SetCursor() ;thanks to martin
	Local $curs
    Switch GetMousePosType()
        Case 0
            $curs = 2
        Case 1, 2
            $curs = 13
        Case 3, 6
            $curs = 11
        Case 5, 7
            $curs = 10
        Case 4, 8
            $curs = 12
    EndSwitch
    GUISetCursor($curs, 1, $hGUI_Area)
EndFunc ;==>SetCursor

Func WM_LBUTTONDOWN($hWnd, $iMsg, $StartWIndowPosaram, $lParam) ;thanks to martin
    Local $drag = GetMousePosType()
    If $drag > 0 Then DllCall("user32.dll", "long", "SendMessage", "hwnd", $hWnd, "int", $WM_SYSCOMMAND, "int", 0xF000 + $drag, "int", 0)
;F001 = LHS, F002 = RHS, F003 = top, F004 = TopLeft, F005 = TopRight, F006 = Bottom, F007 = BL, F008 = BR
;F009 = move gui, same as F011 F12  to F01F
;F010,  moves cursor to centre top of gui - no idea what that is useful for.
;F020 minimizes
;F030 maximizes
EndFunc ;==>WM_LBUTTONDOWN

Func WM_SIZE()
    Local $aGUI_size = WinGetClientSize($hGUI_Area)
    Local $mpos = MouseGetPos()
	WinMove($hGUI_Dot, "", $mpos[0], $mpos[1], 1, 1, 1)
	Draw_Zoom_Preview($mpos[0], $mpos[1])
	$w_area = $aGUI_size[0] + 2
	$h_area = $aGUI_size[1] + 2
    ToolTip("Size: " & $aGUI_size[0] + 2 & "x" & $aGUI_size[1] + 2 & @CRLF & "Position: " & $mpos[0] & "," & $mpos[1] & @LF & "Press ENTER to grab", $drx + $mpos[0], $dry + $mpos[1])
    Return "GUI_RUNDEFMSG"
EndFunc   ;==>WM_SIZE

Func Capture_Window($hWnd, $w, $h)
	_GDIPlus_BitmapDispose($hBmp) ;otherwise memory leak
	$hBmp = 0
	_WinAPI_DeleteObject($hBitmap_s)
	$undo_chk = False
	Local $hDC_Capture = _WinAPI_GetWindowDC($hWnd)
	Local $hMemDC = _WinAPI_CreateCompatibleDC($hDC_Capture)
	$hBitmap_s = _WinAPI_CreateCompatibleBitmap($hDC_Capture, $w, $h)
	_WinAPI_SelectObject($hMemDC, $hBitmap_s)
	Local $at = DllCall("User32.dll", "int", "PrintWindow", "hwnd", $hWnd, "handle", $hMemDC, "int", 0)
	_WinAPI_DeleteDC($hMemDC)
	_WinAPI_ReleaseDC($hWnd, $hDC_Capture)
	$hBmp = _GDIPlus_BitmapCreateFromHBITMAP($hBitmap_s)
	Update_Resolution_Info($hBmp)
	Draw2Graphic($hBmp)
	$hClipboard_Bitmap = $hBitmap_s
	$z = 10
	Zoom(-1)
EndFunc   ;==>Capture_Window

Func Capture_Window2($hWnd, $2front = True)
	_GDIPlus_BitmapDispose($hBmp) ;otherwise memory leak
	$hBmp = 0
	_WinAPI_DeleteObject($hBitmap_s)
	$undo_chk = False
	Local $mini = False
	Local $ws = WinGetState(HWnd($hWnd), "")
	If BitAND($ws, 0x10) Then
		_WinAPI_ShowWindow($hWnd, @SW_RESTORE)
		$ws = WinGetState(HWnd($hWnd), "")
		$mini = True
	Else
		WinActivate(HWnd($hWnd))
	EndIf
	Sleep(500)
	_WinAPI_DeleteObject($hBitmap_s)
	$hBitmap_s = 0
	$hBitmap_s = __ScreenCapture_CaptureWnd("", $hWnd, 0, 0, -1, -1, $ws)
	$hBmp = _GDIPlus_BitmapCreateFromHBITMAP($hBitmap_s)
	Update_Resolution_Info($hBmp)
	Draw2Graphic($hBmp)
	$hClipboard_Bitmap = $hBitmap_s
	$z = 10
	Zoom(-1)
	If $2front Then WinActivate(WinGetHandle($hGUI))
	If $mini Then _WinAPI_ShowWindow($hWnd, @SW_MINIMIZE)
EndFunc   ;==>Capture_Window2

Func __ScreenCapture_CaptureWnd($sFileName, $hWnd, $iLeft = 0, $iTop = 0, $iRight = -1, $iBottom = -1, $maximized = False)
	Local $dex = 4, $dey = 1
	Local $tRect = _WinAPI_GetWindowRect($hWnd)
	If BitAND($maximized, 32) Then
		$dex = 8
		$dey = 8
	EndIf
	$iLeft += DllStructGetData($tRect, "Left")
	$iTop += DllStructGetData($tRect, "Top")
	If $iRight = -1 Then $iRight = DllStructGetData($tRect, "Right") - DllStructGetData($tRect, "Left")
	If $iBottom = -1 Then $iBottom = DllStructGetData($tRect, "Bottom") - DllStructGetData($tRect, "Top")
	$iRight += DllStructGetData($tRect, "Left")
	$iBottom += DllStructGetData($tRect, "Top")
	If $iLeft > DllStructGetData($tRect, "Right") Then $iLeft = DllStructGetData($tRect, "Left")
	If $iTop > DllStructGetData($tRect, "Bottom") Then $iTop = DllStructGetData($tRect, "Top")
	If $iRight > DllStructGetData($tRect, "Right") Then $iRight = DllStructGetData($tRect, "Right")
	If $iBottom > DllStructGetData($tRect, "Bottom") Then $iBottom = DllStructGetData($tRect, "Bottom")
	Return _ScreenCapture_Capture($sFileName, $iLeft + $dex, $iTop + $dey, $iRight - $dex, $iBottom - $dey, False)
EndFunc   ;==>__ScreenCapture_CaptureWnd

Func Grab_Active_Window()
	Capture_Window2(WinGetHandle(WinGetTitle("")), False)
	Return 1
EndFunc   ;==>Grab_Active_Window

Func Grab_Screen()
	While _IsPressed("01", $dll)
	WEnd
	GUISetState(@SW_HIDE, $hGUI)
	Mark_Area()
	GUISetState(@SW_SHOW, $hGUI)
EndFunc   ;==>Grab_Screen

Func Grab_Region($x, $y, $w, $h, $fCursor = False, $capture_only = False)
	_GDIPlus_BitmapDispose($hBmp) ;otherwise memory leak
	_WinAPI_DeleteObject($memBitmap)
	$hDC_Region = _WinAPI_GetDC(0)
	$hMemDC = _WinAPI_CreateCompatibleDC($hDC_Region)
	$memBitmap = _WinAPI_CreateCompatibleBitmap($hDC_Region, $w, $h)
	$hObj = _WinAPI_SelectObject($hMemDC, $memBitmap)
	_WinAPI_BitBlt($hMemDC, 0, 0, $w, $h, $hDC_Region, $x, $y, $SRCCOPY)
	If $fCursor Then
		Local $aCursor = _WinAPI_GetCursorInfo()
		If $aCursor[1] Then
			Local $hIcon = _WinAPI_CopyIcon($aCursor[2])
			Local $aIcon = _WinAPI_GetIconInfo($hIcon)
			_WinAPI_DeleteObject($aIcon[4])
			_WinAPI_DrawIcon($hMemDC, $aCursor[3] - $aIcon[2], $aCursor[4] - $aIcon[3], $hIcon)
			_WinAPI_DestroyIcon($hIcon)
		EndIf
	EndIf
	If $capture_only Then Return $memBitmap
	$hBmp = _GDIPlus_BitmapCreateFromHBITMAP($memBitmap)
	If @error Then Return SetError(1, @extended, _WinAPI_ShowError("Unable to create HBITMAP!"))
	$hClipboard_Bitmap = $memBitmap
	Update_Resolution_Info($hBmp)
	Update_Resolution_Info($hBmp)
	Draw2Graphic($hBmp)
	$z = 10
	Zoom(-1)
	sndPlaySound($BytePtr, $fuSound)
EndFunc   ;==>Grab_Region

Func GrabPrintScreenClipboard($alt = False, $skip = True)
	If Not $skip Then
		If Not $alt Then
			Send("{PRINTSCREEN}")
		Else
			Send("!{PRINTSCREEN}")
		EndIf
	EndIf
	Sleep(125) ;wait for the bitmap to be copied to the clipboard
	_GDIPlus_BitmapDispose($hBmp) ;otherwise memory leak
	_WinAPI_DeleteObject($hMemoryBMP)
	If Not _ClipBoard_Open(0) Then Return _WinAPI_ShowError("_ClipBoard_Open failed")
	$hMemoryBMP = _ClipBoard_GetDataEx($CF_BITMAP)
	If Not $hMemoryBMP Then Return _WinAPI_ShowError("_ClipBoard_GetDataEx failed")
	_ClipBoard_Close()
	$hBmp = _GDIPlus_BitmapCreateFromHBITMAP($hMemoryBMP)
;~ 	$hClipboard_Bitmap = $hMemoryBMP
	Update_Resolution_Info($hBmp)
	$z = 10
	Update_Resolution_Info($hBmp)
	Draw2Graphic($hBmp)
	Zoom(-1)
	sndPlaySound($BytePtr, $fuSound)
EndFunc   ;==>GrabPrintScreenClipboard

Func Update_Resolution_Info($hImg)
	$bW = _GDIPlus_ImageGetWidth($hImg)
	$BH = _GDIPlus_ImageGetHeight($hImg)
	GUICtrlSetData($label4, $bW & " x " & $BH)
EndFunc   ;==>Update_Resolution_Info

Func Draw_Zoom_Preview($mposX, $mposY, $cl_p1 = 0.4)
	Local Const $lW = 40
	Local Const $lH = 40
	Local Const $dw = 10
	If CheckRectCollision($zx - $lW, $zy - $lH, $zx + $zw + $lW, $zy + $zh + $lH, $mposX, $mposY) Then
		If $up = -1 Then
			$zx = @DesktopWidth - $zw - 4
			$zy = 0
			WinMove($hGUI_Zoom, "", $zx, $zy)
			If Not $aero Then _GDIPlus_GraphicsDrawImageRectRect($hGraphic_Freehand, $hBmp_Freehand, 0, 0, $zw + $dw, $zh + $dw, 0, 0, $zw + $dw, $zh + $dw)
		Else
			$zy = 0
			$zx = 0
			WinMove($hGUI_Zoom, "", $zx, $zy)
			If Not $aero Then _GDIPlus_GraphicsDrawImageRectRect($hGraphic_Freehand, $hBmp_Freehand, @DesktopWidth - $zw - 4, 0, $zw + $dw, $zh + $dw, @DesktopWidth - $zw - 4, 0, $zw + $dw, $zh + $dw)
		EndIf
;~ 		If Not $aero Then _GDIPlus_GraphicsDrawImage($hGraphic_Freehand, $hBmp_Freehand, 0, 0)
		$up *= -1
	EndIf
	$zoomW = Int($zw / $zoom_level)
	$zoomH = Int($zh / $zoom_level)

	Local Const $cl_p2 = 1 - $cl_p1
	Local $hPen = _WinAPI_CreatePen($PS_SOLID, $zoom_level / 2, "0x" & $red & $green & $blue)
	Local $obj_orig = _WinAPI_SelectObject($hGUI_ZoomDC, $hPen)
	_WinAPI_StretchBlt($hGUI_ZoomDC, 0, 0, $zw, $zh, $hDC_Zoom, $mposX - $zoomW / 2, $mposY - $zoomH / 2, $zoomW, $zoomH, $SRCCOPY)
	_WinAPI_DrawLine($hGUI_ZoomDC, 0, $zh / 2 + $zoom_level / 2, $zw * $cl_p1, $zh / 2 + $zoom_level / 2)
	_WinAPI_DrawLine($hGUI_ZoomDC, $zw * $cl_p2, $zh / 2 + $zoom_level / 2, $zw, $zh / 2 + $zoom_level / 2)
	_WinAPI_DrawLine($hGUI_ZoomDC, $zw / 2 + $zoom_level / 2, 0, $zw / 2 + $zoom_level / 2, $zh * $cl_p1)
	_WinAPI_DrawLine($hGUI_ZoomDC, $zw / 2 + $zoom_level / 2, $zh * $cl_p2, $zw / 2 + $zoom_level / 2, $zh)
;~ 	_WinAPI_DrawLine($hGUI_ZoomDC, $zw / 2 + $zoom_level, $zh / 2, $zw / 2 + $zoom_level, $zh / 2)
	_WinAPI_SelectObject($hGUI_ZoomDC, $obj_orig)
	_WinAPI_DeleteObject($hPen)
EndFunc   ;==>Draw_Zoom_Preview

Func _WinAPI_StretchBlt($hDestDC, $iXDest, $iYDest, $iWidthDest, $iHeightDest, $hSrcDC, $iXSrc, $iYSrc, $iWidthSrc, $iHeightSrc, $iRop)
	Local $Ret = DllCall("gdi32.dll", "int", "StretchBlt", "hwnd", $hDestDC, "int", $iXDest, "int", $iYDest, "int", $iWidthDest, "int", $iHeightDest, "hwnd", $hSrcDC, "int", $iXSrc, "int", $iYSrc, "int", $iWidthSrc, "int", $iHeightSrc, "dword", $iRop)
	If (@error) Or (Not $Ret[0]) Then Return SetError(1, 0, 0)
	Return 1
EndFunc   ;==>_WinAPI_StretchBlt

Func CheckRectCollision($iLeft, $iTop, $iRight, $iBottom, $iX, $iY)
	Local $tagRECT = "int Left;int Top;int Right;int Bottom"
	Local $tRect = DllStructCreate($tagRECT)
	DllStructSetData($tRect, 1, $iLeft)
	DllStructSetData($tRect, 2, $iTop)
	DllStructSetData($tRect, 3, $iRight)
	DllStructSetData($tRect, 4, $iBottom)
	Local $aResult = DllCall("User32.dll", "int", "PtInRect", "ptr", DllStructGetPtr($tRect), "int", $iX, "int", $iY)
	$tRect = 0
	Return $aResult[0]
EndFunc   ;==>CheckRectCollision

Func _GuiHole($hWnd, $i_x, $i_y, $i_sizew, $i_sizeh, $width, $height)
	Local $outer_rgn, $inner_rgn, $combined_rgn
	$outer_rgn = _WinAPI_CreateRectRgn(0, 0, $width, $height)
	$inner_rgn = _WinAPI_CreateRectRgn($i_x, $i_y, $i_x + $i_sizew, $i_y + $i_sizeh)
	$combined_rgn = _WinAPI_CreateRectRgn(0, 0, 0, 0)
	_WinAPI_CombineRgn($combined_rgn, $outer_rgn, $inner_rgn, $RGN_DIFF)
	_WinAPI_DeleteObject($outer_rgn)
	_WinAPI_DeleteObject($inner_rgn)
	_WinAPI_SetWindowRgn($hWnd, $combined_rgn)
EndFunc   ;==>_GuiHole
#endregion Screen Grab
#region WebGrab
Func WebGrab()
	GUIRegisterMsg($WM_NCPAINT, "")
	Opt("GUIOnEventMode", 0)
	Local $break = False, $aWin, $mouse, $k, $e, $button_pos = $gdip_w - 75, $steps = Int(255 / 27)
	GUICtrlSetState($Button_WebGrab, $GUI_DISABLE)
	GUICtrlSetState($Button_Exit, $GUI_DISABLE)
	GUICtrlSetState($Button_Save, $GUI_DISABLE)
	GUICtrlSetState($Button_Clipboard, $GUI_DISABLE)
	GUICtrlSetState($Button_Refresh, $GUI_DISABLE)
	GUICtrlSetState($Button_Grab2AVI, $GUI_DISABLE)
	GUICtrlSetState($Button_GrabScreen, $GUI_DISABLE)
	GUICtrlSetState($List, $GUI_DISABLE)
;~ 	Local $hGUI_URL = GUICreate("",  526, 26, 264, $height - 114, $WS_POPUP,  $WS_EX_LAYERED + $WS_EX_MDICHILD, $hGUI)
	Local $hGUI_URL = GUICreate("", $gdip_w, 26, 264, $height - 114, $WS_POPUP + $WS_OVERLAPPED, $WS_EX_MDICHILD, $hGUI)
	WinSetTrans($hGUI_URL, "", 0)
	GUISetBkColor(0xBFCDDB, $hGUI_URL)
	$URL_Input = GUICtrlCreateInput("http://", 1, 27, $gdip_w - 77, 23)
	GUICtrlSetTip(-1, "Enter a valid URL. Press ESC to cancel!", "", 0, 1)
	GUICtrlSetFont(-1, 12, 400, 0, "Times New Roman")
	$Grab_Button = GUICtrlCreateButton("Grab &Web", $button_pos, 27, 75, 25)
	GUICtrlSetState(-1, $GUI_DISABLE)
	GUISetState(@SW_DISABLE, $hGUI)
	GUISetState(@SW_SHOW, $hGUI_URL)
;~ 	_WinAPI_SetLayeredWindowAttributes($hGUI_URL, 0xBFCDDB)
	$e = 1
	For $k = 27 To 1 Step -1
		Sleep(10)
		GUICtrlSetPos($URL_Input, 1, $k)
		GUICtrlSetPos($Grab_Button, $button_pos, $k - 1)
		WinSetTrans($hGUI_URL, "", $steps * $e)
		$e += 1
	Next
	ControlClick($hGUI_URL, "", $URL_Input)
	$Dummy = GUICtrlCreateDummy()
	GUIRegisterMsg($WM_COMMAND, "Check_URL_Input")
	While Sleep(50)
		Switch GUIGetMsg()
			Case $Grab_Button
				ExitLoop
			Case $Dummy
				ExitLoop
		EndSwitch
		If _IsPressed("1B", $dll) Then
			$break = True
			ExitLoop
		EndIf
	WEnd
	GUIRegisterMsg($WM_COMMAND, "")
	$Dummy = ""
	If Not $break Then
		GUICtrlSetState($Grab_Button, $GUI_DISABLE)
		GUICtrlSetState($URL_Input, $GUI_DISABLE)
		$aWin = WinGetPos($hGUI)
		SplashTextOn("Screenshooting Web Site", "Please wait while taking screenshot", 435, 40, $aWin[0] + $aWin[2] / 2 - 150, $aWin[1] + $aWin[3] / 2 - 20, 39, "Lucida Console", 12, 400)
		AdlibRegister("Wait", 250)
		$mouse = MouseGetCursor()
		GUISetCursor($hGUI, 15)
		Web_Screenshot(GUICtrlRead($URL_Input))
		GUISetCursor($hGUI, $mouse)
		AdlibUnRegister("Wait")
		SplashOff()
	EndIf
	For $k = 0 To 27
		Sleep(10)
		GUICtrlSetPos($URL_Input, 1, $k + 1)
		GUICtrlSetPos($Grab_Button, $button_pos, $k)
		WinSetTrans($hGUI_URL, "", 0xFF - ($steps * $k))
	Next
	GUISetState(@SW_ENABLE, $hGUI)
	GUIDelete($hGUI_URL)
	GUICtrlSetState($Button_WebGrab, $GUI_ENABLE)
	GUICtrlSetState($Button_Exit, $GUI_ENABLE)
	GUICtrlSetState($Button_Save, $GUI_ENABLE)
	GUICtrlSetState($Button_Clipboard, $GUI_ENABLE)
	GUICtrlSetState($Button_Refresh, $GUI_ENABLE)
	GUICtrlSetState($Button_Grab2AVI, $GUI_ENABLE)
	GUICtrlSetState($Button_GrabScreen, $GUI_ENABLE)
	GUICtrlSetState($List, $GUI_ENABLE)
	Opt("GUIOnEventMode", 1)
	GUIRegisterMsg($WM_NCPAINT, "Redraw")
	GUIRegisterMsg($WM_COMMAND, "WM_COMMAND")
	$undo_chk = False
;~ 	WinActivate($hGUI)
EndFunc   ;==>WebGrab

Func Check_URL_Input($hWnd, $Msg, $wParam, $lParam)
	#forceref $hWnd, $Msg, $wParam, $lParam
	Local $input, $chk, $chk2, $go = False
	If BitAND($wParam, 0x0000FFFF) = $URL_Input Then
		$input = GUICtrlRead($URL_Input)
		$chk2 = StringReplace($input, "://", "://")
		$chk2 = @extended
		$chk = StringRegExp($input, "\A(?i:http://.+|https://.+|ftp://.+)", 3)
		If Not @error And $chk2 = 1 Then
			GUICtrlSetState($Grab_Button, $GUI_ENABLE)
			$go = True
		Else
			GUICtrlSetState($Grab_Button, $GUI_DISABLE)
			$go = False
		EndIf
	EndIf
	If Not BitAND($wParam, 0xFFFF0000) And $go Then GUICtrlSendToDummy($Dummy)
EndFunc   ;==>Check_URL_Input

Func Web_Screenshot($url, $IEwidth = 1045) ;requires .NET v3.0
	Local $oIE, $GUIActiveX, $oDocument, $oBody, $BodyWidth, $BodyHeight, $oHtml
	Local $hGUI_WebGrab, $hWin, $aWin, $aMP
	Local Const $BrowserNavConstant = 2 + 128 + 256 + 512 + 4096 + 32768 ;BrowserNavConstant -> http://msdn.microsoft.com/en-us/library/dd565688(v=vs.85).aspx
	$oIE = ObjCreate("Shell.Explorer.2") ;http://msdn.microsoft.com/en-us/library/aa752084(v=vs.85).aspx

	$hWin = WinGetHandle("Program Manager")
	$aWin = WinGetPos($hWin)

	#region render web site to get height
	$hGUI_WebGrab = GUICreate("", 0, 0, $aWin[0] - $BodyWidth + 1, $aWin[1] - $BodyHeight + 1, BitOR($WS_CLIPSIBLINGS, $WS_CLIPCHILDREN, $WS_POPUP), Default, WinGetHandle(AutoItWinGetTitle()))
	$GUIActiveX = GUICtrlCreateObj($oIE, 0, 0, $IEwidth, 256)
	GUISetState(@SW_MINIMIZE, $hGUI_WebGrab)

	Local $timeout = 60 * 1000
	Local $timer = TimerInit()
	With $oIE
		.Silent = True
		.FullScreen = True
		.Resizable = False
		.Visible = False
		.StatusBar = False
		.AddressBar = False
		.Navigate($url, $BrowserNavConstant, "_top")
		Do
			If TimerDiff($timer) > $timeout Then ExitLoop
			Sleep(500)
		Until .ReadyState = 4
		.Stop
	EndWith
	Sleep(500)

	$oDocument = $oIE.document
	$oBody = $oDocument.body
	$oBody.scroll = "no"
	$oBody.style.borderStyle = "none"
	$BodyWidth = Min($oBody.scrollWidth, 8192)
	$BodyHeight = Min($oBody.scrollHeight, 8192)
	GUIDelete($hGUI_WebGrab)
	$GUIActiveX = 0
	#endregion render web site to get height

	Local $IE_Ver = RegRead("HKLM\SOFTWARE\Microsoft\Internet Explorer", "Version"), $height
	If $IE_Ver > 8 Then $BodyHeight = Min($BodyHeight, 8190)

	$hGUI_WebGrab = GUICreate("", $BodyWidth, $BodyHeight, $aWin[0] - $BodyWidth + 1, $aWin[1] - $BodyHeight + 1, BitOR($WS_CLIPSIBLINGS, $WS_CLIPCHILDREN, $WS_POPUP), Default, WinGetHandle(AutoItWinGetTitle()))
;~ 	$hGUI_WebGrab = GUICreate("",  $BodyWidth,  $BodyHeight, -1, -1, BitOR($WS_CLIPSIBLINGS, $WS_CLIPCHILDREN, $WS_POPUP), Default, WinGetHandle(AutoItWinGetTitle()))
	$GUIActiveX = GUICtrlCreateObj($oIE, 0, 0, $BodyWidth, $BodyHeight) ;with IE9 max. height can be only 8192 otherwise white screen only!?!?
;~ 	GUISetState(@SW_SHOWNA, $hGUI_WebGrab) ;Show GUI for some milli seconds to make a screenshot

	$timer = TimerInit()
	With $oIE
		.Silent = True
		.FullScreen = True
		.Resizable = False
		.Visible = False
		.StatusBar = False
		.AddressBar = False
		.Navigate($url, $BrowserNavConstant, "_top")
		Do
			If TimerDiff($timer) > $timeout Then ExitLoop
			Sleep(500)
		Until .ReadyState = 4
		.Stop
		Sleep(1000)
	EndWith

	$oDocument = $oIE.document
	$oBody = $oDocument.body
	$oHtml = $oDocument.documentElement
	$oBody.scroll = "no"
	$oBody.style.borderStyle = "none"
	$oBody.style.border = "0px"
	$oHtml.style.overflow = 'hidden'
	$aMP = MouseGetPos()
	GUISetState(@SW_SHOWNA, $hGUI_WebGrab) ;Show GUI for some milli seconds to make a screenshot
	Sleep(500)
	MouseMove($aMP[0], $aMP[1], 0)
	_GDIPlus_BitmapDispose($hBmp) ;otherwise memory leak
	Local $hDC_Web = _WinAPI_GetWindowDC($hGUI_WebGrab)
	Local $hDC_Dummy = _WinAPI_GetWindowDC(0)
	Local $hMemDC = _WinAPI_CreateCompatibleDC($hDC_Dummy)
	Local $hBitmap = _WinAPI_CreateCompatibleBitmap($hDC_Dummy, $BodyWidth, $BodyHeight)
	_WinAPI_SelectObject($hMemDC, $hBitmap)
	_WinAPI_BitBlt($hMemDC, 0, 0, $BodyWidth, $BodyHeight, $hDC_Web, 0, 0, $SRCCOPY)

	$hBmp = _GDIPlus_BitmapCreateFromHBITMAP($hBitmap)
	If BitAND(GUICtrlRead($Button_Menu_Web_Item), $GUI_CHECKED) = $GUI_CHECKED Then
		Local Const $hCtx_Web = _GDIPlus_ImageGetGraphicsContext($hBmp)
		Local Const $width_web = _GDIPlus_ImageGetWidth($hBmp)
		Local Const $height_web = _GDIPlus_ImageGetHeight($hBmp)
		Local Const $hBrushBG_Web = _GDIPlus_CreateLineBrush(0, 0, $width_web + 16, 16, 0xE0FFFFF0, 0x08FFFFF0)
		_GDIPlus_GraphicsFillRect($hCtx_Web, 2, $height_web - 18, $width_web - 4, 16, $hBrushBG_Web)
		_GDIPlus_GraphicsDrawString($hCtx_Web, "URL: " & $url, 3, $height_web - 18, "Times New Roman", 10)
		_GDIPlus_BrushDispose($hBrushBG_Web)
		_GDIPlus_GraphicsDispose($hCtx_Web)
	EndIf

	$GUIActiveX = 0
	$oIE = 0
	Update_Resolution_Info($hBmp)
	Draw2Graphic($hBmp)
	$z = 10
	Zoom(-1)
	$hClipboard_Bitmap = $hBitmap
;~ 	_WinAPI_DeleteObject($hBitmap) ;if activated clipboard is not working anymore because bitmap is deleted from mem
	_WinAPI_DeleteDC($hMemDC)
	_WinAPI_ReleaseDC($hGUI_WebGrab, $hDC_Web)
	_WinAPI_ReleaseDC(0, $hDC_Dummy)
	_WinAPI_ReleaseDC(0, $hDC_Web)
	GUIDelete($hGUI_WebGrab)
	sndPlaySound($BytePtr, $fuSound)
	_ReduceMemory(@AutoItPID)
EndFunc   ;==>Web_Screenshot

Func Min($a, $b)
	If $a < $b Then Return $a
	Return $b
EndFunc   ;==>Min

Func Max($a, $b)
	If $a > $b Then Return $a
	Return $b
EndFunc   ;==>Max

Func Wait()
	Switch Mod($i, 5)
		Case 0
			ControlSetText("Screenshooting Web Site", "", "Static1", "Please wait while taking screenshot " & ChrW(8250))
		Case 1
			ControlSetText("Screenshooting Web Site", "", "Static1", "Please wait while taking screenshot —" & ChrW(8250))
		Case 2
			ControlSetText("Screenshooting Web Site", "", "Static1", "Please wait while taking screenshot ——" & ChrW(8250))
		Case 3
			ControlSetText("Screenshooting Web Site", "", "Static1", "Please wait while taking screenshot ———" & ChrW(8250))
		Case 4
			ControlSetText("Screenshooting Web Site", "", "Static1", "Please wait while taking screenshot ————" & ChrW(8250))
	EndSwitch
	$i += 1
EndFunc   ;==>Wait

Func _ReduceMemory($ProcID = 0)
	If $ProcID = 0 Or ProcessExists($ProcID) = 0 Then ; No process id specified or process doesnt exist - use current process instead.
		Local $ai_GetCurrentProcess = DllCall('kernel32.dll', 'ptr', 'GetCurrentProcess')
		Local $ai_Return = DllCall("psapi.dll", 'int', 'EmptyWorkingSet', 'ptr', $ai_GetCurrentProcess[0])
		Return $ai_Return[0]
	EndIf
	Local $ai_Handle = DllCall("kernel32.dll", 'ptr', 'OpenProcess', 'int', 0x1f0fff, 'int', False, 'int', $ProcID)
	Local $ai_Return = DllCall("psapi.dll", 'int', 'EmptyWorkingSet', 'ptr', $ai_Handle[0])
	DllCall('kernel32.dll', 'int', 'CloseHandle', 'ptr', $ai_Handle[0])
	Return $ai_Return[0]
EndFunc   ;==>_ReduceMemory

Func ObjErrorHandler()
	ConsoleWrite("A COM Error has occured!" & @CRLF & @CRLF & _
			"err.description is: " & @TAB & $oErrorHandler.description & @CRLF & _
			"err.windescription:" & @TAB & $oErrorHandler & @CRLF & _
			"err.number is: " & @TAB & Hex($oErrorHandler.number, 8) & @CRLF & _
			"err.lastdllerror is: " & @TAB & $oErrorHandler.lastdllerror & @CRLF & _
			"err.scriptline is: " & @TAB & $oErrorHandler.scriptline & @CRLF & _
			"err.source is: " & @TAB & $oErrorHandler.source & @CRLF & _
			"err.helpfile is: " & @TAB & $oErrorHandler.helpfile & @CRLF & _
			"err.helpcontext is: " & @TAB & $oErrorHandler.helpcontext & @CRLF _
			)
EndFunc   ;==>ObjErrorHandler
#endregion WebGrab
#region Save Bitmap
Func Save_Bitmap()
	Local $m, $save, $ok = True
	$filename = FileSaveDialog("Save Image", @ScriptDir, "Images (*.jpg;*.png;*.bmp;*.gif;*.tif;*.pdf)", 18, "", $hGUI)
	If Not @error Then
		If BitAND(GUICtrlRead($Button_Menu_Item2), $GUI_CHECKED) = $GUI_CHECKED Then
			$hBmp = WTOB($hBmp, _Now())
			Zoom(0)
		EndIf
		If StringMid($filename, StringLen($filename) - 3, 1) <> "." Then
			$filename &= ".png"
			If FileExists($filename) Then
				$m = MsgBox(4 + 48, "Confirm save", $filename & " already exists." & @CRLF & "Do you want t replace it?")
				If $m = 7 Then $ok = False
			EndIf
		EndIf
		If $ok Then
			Switch StringRight($filename, 4)
				Case ".jpg"
					Local $sCLSID = _GDIPlus_EncodersGetCLSID("JPG")
					Local $tParams = _GDIPlus_ParamInit(1)
					Local $tData = DllStructCreate("int Quality")
					DllStructSetData($tData, "Quality", $JPEG_Quality)
					Local $pData = DllStructGetPtr($tData)
					_GDIPlus_ParamAdd($tParams, $GDIP_EPGQUALITY, 1, $GDIP_EPTLONG, $pData)
					Local $pParams = DllStructGetPtr($tParams)
					$save = _GDIPlus_ImageSaveToFileEx($hBmp, $filename, $sCLSID, $pParams)
					$tData = ""
					$tParams = ""
				Case ".pdf"
					$save = Save2PDF($filename)
				Case Else
					$save = _GDIPlus_ImageSaveToFile($hBmp, $filename)
			EndSwitch
		EndIf

		If $save Then
			Local $c = MsgBox(4 + 256, "Information", "Image saved properly to: " & @CRLF & @CRLF & $filename & @CRLF & @CRLF & @CRLF & _
					"Display saved image with default app?", 20)
			If $c = 6 Then ShellExecute($filename)
		Else
			MsgBox(16, "ERROR", "Image could not be saved!", 10)
		EndIf
	EndIf
EndFunc   ;==>Save_Bitmap

Func Save2PDF($filename)
	If BitAND(GUICtrlRead($Button_Menu_Item2), $GUI_CHECKED) = $GUI_CHECKED Then
		$hBmp = WTOB($hBmp, _Now())
		Zoom(0)
	EndIf
	_SetTitle("Image2PDF")
	_SetSubject("Converted by Rk Screencapture")
	_SetKeywords("pdf, Rk, Rk SCreencapture")
	_SetUnit($PDF_UNIT_PT) ;pixel
	_SetPaperSize("A4")
	_SetZoomMode($PDF_ZOOM_CUSTOM, 66)
	_SetOrientation($PDF_ORIENTATION_PORTRAIT)
	_SetLayoutMode($PDF_LAYOUT_SINGLE)
	_InitPDF($filename)
	_LoadFontTT("_Arial", $PDF_FONT_ARIAL)
	_LoadResImageMem("UEZ", Bitmap2String($hBmp), $bW, $BH)
	_BeginPage()
	Local $x, $y
	Local $pW = _GetPageWidth() ;595
	Local $pH = _GetPageHeight() ;841
	If $bW < $pW And $BH < $pH Then
		$x = ($pW - $bW) / 2
		$y = ($pH - $BH) / 2
		_InsertImage("UEZ", $x, $y, $bW, $BH)
	ElseIf $bW > $pW And $BH < $pH Then
		$x = 0
		$y = ($pH - ($pW / $bW) * $BH) / 2
		_InsertImage("UEZ", $x, $y, $pW, ($pW / $bW) * $BH)
	ElseIf $bW < $pW And $BH > $pH Then
		$x = ($pW - $bW) / 2
		$y = 0
		_InsertImage("UEZ", $x, $y, $bW, $pH)
	ElseIf $bW > $pW And $BH > $pH Then
		If $bW > $BH Then
			$x = 0
			$y = ($pH - ($pW / $bW) * $BH) / 2
			_InsertImage("UEZ", $x, $y, $pW, ($pW / $bW) * $BH)
		Else
			$x = ($pW - ($pH / $BH) * $bW) / 2
			$y = 0
			_InsertImage("UEZ", $x, $y, ($pH / $BH) * $bW, $pH)
		EndIf
	Else
		_InsertImage("UEZ", 0, 0, $pW, $pH)
	EndIf
	_SetTextRenderingMode(2)
	_SetColourFill(0x300000)
	_DrawText(7, 2, "Created by Rk Screencapture", "_Arial", 8, $PDF_ALIGN_LEFT, 90)
	_SetTextRenderingMode(0)
	_SetColourFill(0xE80000)
	_DrawText(7, 2, "Created by Rk Screencapture", "_Arial", 8, $PDF_ALIGN_LEFT, 90)
	_EndPage()
	_ClosePDFFile()
	Return True
EndFunc   ;==>Save2PDF

Func Bitmap2String($Bitmap)
	Local $CLSID_Encoder = _GDIPlus_EncodersGetCLSID("jpg")
	Local $tagGUID_Struct = _WinAPI_GUIDFromString($CLSID_Encoder)
	Local $pTagGUID = DllStructGetPtr($tagGUID_Struct)
	Local $Stream = DllCall("ole32.dll", "uint", "CreateStreamOnHGlobal", "ptr", 0, "bool", 1, "ptr*", 0)
	DllCall($ghGDIPDll, "uint", "GdipSaveImageToStream", "ptr", $Bitmap, "ptr", $Stream[3], "ptr", $pTagGUID, "ptr", 0)
	Local $Memory = DllCall("ole32.dll", "uint", "GetHGlobalFromStream", "ptr", $Stream[3], "ptr*", 0)
	Local $Mem_Size = _MemGlobalSize($Memory[2])
	Local $Mem_Ptr = _MemGlobalLock($Memory[2])
	Local $bData_Struct = DllStructCreate("byte[" & $Mem_Size & "]", $Mem_Ptr)
	Local $bData = DllStructGetData($bData_Struct, 1)
	Local $tVARIANT = DllStructCreate("word vt;word r1;word r2;word r3;ptr data;ptr")
	Local $aCall = DllCall("oleaut32.dll", "long", "DispCallFunc", "ptr", $Stream[3], "dword", 8 + 8 * @AutoItX64, "dword", 4, "dword", 23, "dword", 0, "ptr", 0, "ptr", 0, "ptr", DllStructGetPtr($tVARIANT))
	_MemGlobalFree($Memory[2])
	$bData_Struct = 0
	$tVARIANT = 0
	Return BinaryToString($bData)
EndFunc   ;==>Bitmap2String

Func WTOB($hImage, $text, $fName = "Arial", $fSize = -1, $x = -1, $y = -1, $alignX = 2, $alignY = 2, $fColor = 0xFFFFFFFF, $ebg = 1)
	Local $iW = _GDIPlus_ImageGetWidth($hImage)
	Local $iH = _GDIPlus_ImageGetHeight($hImage)

	Local $hBitmap = _GDIPlus_BitmapCreateFromScan0($iW, $iH)
	Local $hContext = _GDIPlus_ImageGetGraphicsContext($hBitmap)
	_GDIPlus_GraphicsSetSmoothingMode($hContext, 2)
	DllCall($ghGDIPDll, "int", "GdipSetInterpolationMode", "hwnd", $hContext, "int", 7)

	$fSize = Max(5, Int($iW * 0.010))

	Local $hPinsel = _GDIPlus_BrushCreateSolid($fColor)
	Local $hFormat = _GDIPlus_StringFormatCreate()
	Local $hFamily = _GDIPlus_FontFamilyCreate($fName)
	Local $hFont = _GDIPlus_FontCreate($hFamily, $fSize, 1)
	Local $tLayout = _GDIPlus_RectFCreate(0, 0, 0, 0)
	Local $aInfo = _GDIPlus_GraphicsMeasureString($hContext, $text, $hFont, $tLayout, $hFormat)

	_GDIPlus_GraphicsDrawImageRect($hContext, $hImage, 0, 0, $iW, $iH)

	Local $fWidth = DllStructGetData($aInfo[0], "Width")
	Local $fHeight = DllStructGetData($aInfo[0], "Height")

	If $x < 0 Then
		Switch $alignX
			Case 0 ;alignment center
				DllStructSetData($tLayout, "x", $iW / 2 - Round($fWidth / 2, 0))
			Case 1 ;alignment left
				DllStructSetData($tLayout, "x", 0)
			Case 2 ;alignment right
				DllStructSetData($tLayout, "x", $iW - $fWidth - 1)
		EndSwitch
	Else
		DllStructSetData($tLayout, "x", $x)
	EndIf

	If $y < 0 Then
		Switch $alignY
			Case 0 ;alignment center
				DllStructSetData($tLayout, "y", $iH / 2 - Floor($fHeight / 2))
			Case 1 ;alignment top
				DllStructSetData($tLayout, "y", 0)
			Case 2 ;alignment buttom
				DllStructSetData($tLayout, "y", $iH - $fHeight - 1)
		EndSwitch
	Else
		DllStructSetData($tLayout, "y", $y)
	EndIf

	Local $hBrush_back = _GDIPlus_CreateLineBrush(0, 0, 0, $fSize, 0xFF101010, 0xFF808080)
	Local $i, $fs = $fSize * 0.075
	Local $tLayout2 = _GDIPlus_RectFCreate(0, 0, 0, 0)
	DllStructSetData($tLayout2, "Width", $fWidth)
	DllStructSetData($tLayout2, "Height", $fHeight)
	If $ebg Then
		For $i = 0 To 3
			Switch $i
				Case 0
					DllStructSetData($tLayout2, "x", DllStructGetData($tLayout, "x"))
					DllStructSetData($tLayout2, "y", DllStructGetData($tLayout, "y") - $fs)
				Case 1
					DllStructSetData($tLayout2, "x", DllStructGetData($tLayout, "x") + $fs)
					DllStructSetData($tLayout2, "y", DllStructGetData($tLayout, "y"))
				Case 2
					DllStructSetData($tLayout2, "x", DllStructGetData($tLayout, "x"))
					DllStructSetData($tLayout2, "y", DllStructGetData($tLayout, "y") + $fs)
				Case 3
					DllStructSetData($tLayout2, "x", DllStructGetData($tLayout, "x") - $fs)
					DllStructSetData($tLayout2, "y", DllStructGetData($tLayout, "y"))
			EndSwitch
			_GDIPlus_GraphicsDrawStringEx($hContext, $text, $hFont, $tLayout2, $hFormat, $hBrush_back)
		Next
	EndIf
	_GDIPlus_GraphicsDrawStringEx($hContext, $text, $hFont, $tLayout, $hFormat, $hPinsel)

	_GDIPlus_ImageDispose($hImage)
	_GDIPlus_FontDispose($hFont)
	_GDIPlus_FontFamilyDispose($hFamily)
	_GDIPlus_StringFormatDispose($hFormat)
	_GDIPlus_BrushDispose($hPinsel)
	_GDIPlus_BrushDispose($hBrush_back)
	_GDIPlus_GraphicsDispose($hContext)
	$tLayout = 0
	$tLayout2 = 0
	Return $hBitmap
EndFunc   ;==>WTOB

Func Pic2Clipboard()
	Local $err, $err_txt
	If Not _ClipBoard_Open(0) Then
		$err = @error
		$err_txt = "_ClipBoard_Open failed!"
	EndIf
	If Not _ClipBoard_Empty() Then
		$err = @error
		$err_txt = "_ClipBoard_Empty failed!"
	EndIf
	If Not _ClipBoard_SetDataEx($hClipboard_Bitmap, $CF_BITMAP) Then
		$err = @error
		$err_txt = "_ClipBoard_SetDataEx failed!"
	EndIf
	_ClipBoard_Close()
	If Not $silent Then
		If Not $err Then
			MsgBox(0, "Information", "Image put to clipboard!", 10)
			Return 1
		Else
			MsgBox(16, "Error", "An error has occured: " & $err_txt, 10)
			Return 0
		EndIf
	EndIf
	If Not $err Then Return 1
	Return 0
EndFunc   ;==>Pic2Clipboard

Func Slider_Label_Color($value)
	Switch $value
		Case 0 To 40
			GUICtrlSetColor($Slider_Label_Value, 0xFF0000)
		Case 41 To 70
			GUICtrlSetColor($Slider_Label_Value, 0xFF8000)
		Case 71 To 89
			GUICtrlSetColor($Slider_Label_Value, 0x80F000)
		Case 90 To 100
			GUICtrlSetColor($Slider_Label_Value, 0x00F000)
	EndSwitch
EndFunc   ;==>Slider_Label_Color
#endregion Save Bitmap

#region Print current image
Func PrintImage() ;sends current image to printer
	Local $tempFile = _TempFile(@TempDir, "~", ".png", 8)
	If _GDIPlus_ImageSaveToFile($hBmp, $tempFile) Then
		GUISetState(@SW_DISABLE, $hGUI)
		RunWait('Rundll32.exe "' & @SystemDir & '\mshtml.dll",PrintHTML "' & $tempFile & '"', @SystemDir)
;~ 		ShellExecute("Rundll32.exe", @SystemDir & "\mshtml.dll,PrintHTML " & $tempFile)
;~ 		ConsoleWrite(@SystemDir & "\Rundll32.exe" & " " &  @SystemDir & "\mshtml.dll,PrintHTML " & $tempFile & @CRLF)
		FileDelete($tempFile)
		WinActivate($hGUI)
		GUISetState(@SW_ENABLE, $hGUI)
	Else
		MsgBox(16, "ERROR", "Unable to send current image to printer!", 10)
	EndIf
EndFunc   ;==>PrintImage
#endregion Print current image
#region Check for updates

#endregion Check for updates

#region Windows Message Codes
Func WM_COMMAND($hWnd, $Msg, $wParam, $lParam)
	#forceref $hWnd, $Msg, $wParam, $lParam
	Local $d, $dd, $iX, $iY
	Switch $wParam
		Case $id_Low
			_GUICtrlMenu_CheckRadioItem($hQMenu_Sub, 0, 2, 0)
			_GDIPlus_GraphicsSetInterpolationMode($hContext, 5)
			Zoom(2)
		Case $id_Medium
			_GUICtrlMenu_CheckRadioItem($hQMenu_Sub, 0, 2, 1)
			_GDIPlus_GraphicsSetInterpolationMode($hContext, 0)
			Zoom(2)
		Case $id_High
			_GUICtrlMenu_CheckRadioItem($hQMenu_Sub, 0, 2, 2)
			_GDIPlus_GraphicsSetInterpolationMode($hContext, 7)
			Zoom(2)
		Case $id_Reset
			$z = 10
			Zoom(-1)
		Case $id_Grey
			ASM_Bitmap_Grey_BnW($hBmp, 0)
			Zoom(2)
			_GUICtrlMenu_EnableMenuItem($hQMenu_Sub3, 6, 0)
		Case $id_BW
			ASM_Bitmap_Grey_BnW($hBmp, 1)
			Zoom(2)
			_GUICtrlMenu_EnableMenuItem($hQMenu_Sub3, 6, 0)
		Case $id_Invert
			ASM_Bitmap_Invert($hBmp)
			Zoom(2)
			_GUICtrlMenu_EnableMenuItem($hQMenu_Sub3, 6, 0)
		Case $id_Rotm90 ;rotate -90°
			FlipImage(3)
			Zoom(-1)
			_GUICtrlMenu_EnableMenuItem($hQMenu_Sub3, 6, 0)
		Case $id_Rotp90 ;rotate +90°
			FlipImage(1)
			Zoom(-1)
			_GUICtrlMenu_EnableMenuItem($hQMenu_Sub3, 6, 0)
		Case $id_Undo
			Undo()
			Zoom(2)
			_GUICtrlMenu_EnableMenuItem($hQMenu_Sub3, 6, 2)
		Case $Button_Menu_Item1
			Opt("GUICloseOnESC", 1)
			Opt("GUIOnEventMode", 0)
			Opt("MouseCoordMode", 2)
			GUISetOnEvent($GUI_EVENT_CLOSE, "")
			Local $mp = MouseGetPos()
			$hGui_Slider = GUICreate("JPEG Save Image Quality", 306, 150, $mp[0] - 153, $mp[1] - 100, $WS_SYSMENU, $WS_EX_MDICHILD, $hGUI)
			GUISetBkColor(0x6060A0)
			$Slider = GUICtrlCreateSlider(10, 10, 280, 50, BitOR($GUI_SS_DEFAULT_SLIDER, $TBS_BOTH, $TBS_ENABLESELRANGE, $WS_TABSTOP))
			GUICtrlSetLimit($Slider, 100, 0)
			GUICtrlSetData($Slider, $JPEG_Quality)
			GUICtrlSetBkColor($Slider, 0xA0A0F0)
			$hSlider = GUICtrlGetHandle($Slider)
			$Slider_Label_Text = GUICtrlCreateLabel("Quality: ", 10, 80, 85, 30)
			GUICtrlSetColor($Slider_Label_Text, 0xF0F0FF)
			GUICtrlSetFont($Slider_Label_Text, 18)
			$Slider_Label_Value = GUICtrlCreateLabel($JPEG_Quality, 105, 80, 45, 30, $SS_CENTER)
			GUICtrlSetFont($Slider_Label_Value, 18)
			Local $Slider_Button = GUICtrlCreateButton("OK", 225, 80, 60, 30)
			GUICtrlSetBkColor($Slider_Button, 0xF0F050)
			Slider_Label_Color($JPEG_Quality)
			GUISetState(@SW_SHOW, $hGui_Slider)
			GUIRegisterMsg($WM_HSCROLL, "WM_HSCROLL")
			Local $aRect = _GUICtrlSlider_GetThumbRect($hSlider) ;get slider button position
			MouseMove($aRect[2] + 4, $aRect[3], 10)
			While 1
				$_msg = GUIGetMsg()
				Switch $_msg
					Case $GUI_EVENT_CLOSE
						ExitLoop
					Case $Slider_Button
						$JPEG_Quality = _GUICtrlSlider_GetPos($hSlider)
						ExitLoop
				EndSwitch
			WEnd
			GUIDelete($hGui_Slider)
			GUIRegisterMsg($WM_HSCROLL, "")
			Opt("GUIOnEventMode", 1)
			Opt("GUICloseOnESC", 0)
			Opt("MouseCoordMode", 1)
			GUISetOnEvent($GUI_EVENT_CLOSE, "_Exit")
		Case $Button_Menu_Item2
			If BitAND(GUICtrlRead($Button_Menu_Item2), $GUI_CHECKED) = $GUI_CHECKED Then
				GUICtrlSetState($Button_Menu_Item2, $GUI_UNCHECKED)
			Else
				GUICtrlSetState($Button_Menu_Item2, $GUI_CHECKED)
			EndIf
		Case $aButton_Menu_AVI_Sub1_Item[0]
			$dd = 0
			For $d = 0 To UBound($aButton_Menu_AVI_Sub1_Item) - 1
				If $d = $dd Then
					GUICtrlSetState($aButton_Menu_AVI_Sub1_Item[$d], $GUI_CHECKED)
				Else
					GUICtrlSetState($aButton_Menu_AVI_Sub1_Item[$d], $GUI_UNCHECKED)
				EndIf
			Next
			GUICtrlSetData($aButton_Menu_AVI_Sub1_Item[UBound($aButton_Menu_AVI_Sub1_Item) - 1], "Custom")
		Case $aButton_Menu_AVI_Sub1_Item[1]
			$dd = 1
			For $d = 0 To UBound($aButton_Menu_AVI_Sub1_Item) - 1
				If $d = $dd Then
					GUICtrlSetState($aButton_Menu_AVI_Sub1_Item[$d], $GUI_CHECKED)
				Else
					GUICtrlSetState($aButton_Menu_AVI_Sub1_Item[$d], $GUI_UNCHECKED)
				EndIf
			Next
			GUICtrlSetData($aButton_Menu_AVI_Sub1_Item[UBound($aButton_Menu_AVI_Sub1_Item) - 1], "Custom")
		Case $aButton_Menu_AVI_Sub1_Item[2]
			$dd = 2
			For $d = 0 To UBound($aButton_Menu_AVI_Sub1_Item) - 1
				If $d = $dd Then
					GUICtrlSetState($aButton_Menu_AVI_Sub1_Item[$d], $GUI_CHECKED)
				Else
					GUICtrlSetState($aButton_Menu_AVI_Sub1_Item[$d], $GUI_UNCHECKED)
				EndIf
			Next
			GUICtrlSetData($aButton_Menu_AVI_Sub1_Item[UBound($aButton_Menu_AVI_Sub1_Item) - 1], "Custom")
		Case $aButton_Menu_AVI_Sub1_Item[3]
			$dd = 3
			For $d = 0 To UBound($aButton_Menu_AVI_Sub1_Item) - 1
				If $d = $dd Then
					GUICtrlSetState($aButton_Menu_AVI_Sub1_Item[$d], $GUI_CHECKED)
				Else
					GUICtrlSetState($aButton_Menu_AVI_Sub1_Item[$d], $GUI_UNCHECKED)
				EndIf
			Next
			GUICtrlSetData($aButton_Menu_AVI_Sub1_Item[UBound($aButton_Menu_AVI_Sub1_Item) - 1], "Custom")
		Case $aButton_Menu_AVI_Sub1_Item[UBound($aButton_Menu_AVI_Sub1_Item) - 1]
			$dd = 4
			Local $new_time = "10000"
			$iX = MouseGetPos(0) - 250
			$iY = MouseGetPos(1) - 150
			While Not StringIsInt($new_time) Or $new_time < 1 Or $new_time > 600
				Local $new_time = InputBox("New Record Time Value", "Please enter a custom record time value in seconds (1 - 600):", 30, "", -1, 140, $iX, $iY, 0, $hGUI)
				If @error Then ExitLoop
			WEnd
			If Not @error Then
				$AVI_ReqTime[$dd] = $new_time
				GUICtrlSetData($aButton_Menu_AVI_Sub1_Item[$dd], "Custom: " & $new_time & " sec")
				For $d = 0 To UBound($aButton_Menu_AVI_Sub1_Item) - 1
					If $d = $dd Then
						GUICtrlSetState($aButton_Menu_AVI_Sub1_Item[$d], $GUI_CHECKED)
					Else
						GUICtrlSetState($aButton_Menu_AVI_Sub1_Item[$d], $GUI_UNCHECKED)
					EndIf
				Next
			EndIf
		Case $aButton_Menu_AVI_Sub2_Item[0]
			$dd = 0
			For $d = 0 To UBound($aButton_Menu_AVI_Sub2_Item) - 1
				If $d = $dd Then
					GUICtrlSetState($aButton_Menu_AVI_Sub2_Item[$d], $GUI_CHECKED)
				Else
					GUICtrlSetState($aButton_Menu_AVI_Sub2_Item[$d], $GUI_UNCHECKED)
				EndIf
			Next
			GUICtrlSetData($aButton_Menu_AVI_Sub2_Item[4], "Custom")
		Case $aButton_Menu_AVI_Sub2_Item[1]
			$dd = 1
			For $d = 0 To UBound($aButton_Menu_AVI_Sub2_Item) - 1
				If $d = $dd Then
					GUICtrlSetState($aButton_Menu_AVI_Sub2_Item[$d], $GUI_CHECKED)
				Else
					GUICtrlSetState($aButton_Menu_AVI_Sub2_Item[$d], $GUI_UNCHECKED)
				EndIf
			Next
			GUICtrlSetData($aButton_Menu_AVI_Sub2_Item[4], "Custom")
		Case $aButton_Menu_AVI_Sub2_Item[2]
			$dd = 2
			For $d = 0 To UBound($aButton_Menu_AVI_Sub2_Item) - 1
				If $d = $dd Then
					GUICtrlSetState($aButton_Menu_AVI_Sub2_Item[$d], $GUI_CHECKED)
				Else
					GUICtrlSetState($aButton_Menu_AVI_Sub2_Item[$d], $GUI_UNCHECKED)
				EndIf
			Next
			GUICtrlSetData($aButton_Menu_AVI_Sub2_Item[4], "Custom")
		Case $aButton_Menu_AVI_Sub2_Item[3]
			$dd = 3
			For $d = 0 To UBound($aButton_Menu_AVI_Sub2_Item) - 1
				If $d = $dd Then
					GUICtrlSetState($aButton_Menu_AVI_Sub2_Item[$d], $GUI_CHECKED)
				Else
					GUICtrlSetState($aButton_Menu_AVI_Sub2_Item[$d], $GUI_UNCHECKED)
				EndIf
			Next
			GUICtrlSetData($aButton_Menu_AVI_Sub2_Item[4], "Custom")
		Case $aButton_Menu_AVI_Sub2_Item[4]
			$dd = 4
			Local $new_fps = "100"
			$iX = MouseGetPos(0) - 250
			$iY = MouseGetPos(1) - 150
			While Not StringIsInt($new_fps) Or $new_fps < 1 Or $new_fps > 30
				Local $new_fps = InputBox("New FPS Value", "Please enter a custom fps value (1 - 30):", 25, "", -1, 140, $iX, $iY, 0, $hGUI)
				If @error Then ExitLoop
			WEnd
			If Not @error Then
				$AVI_FPS[$dd] = $new_fps
				GUICtrlSetData($aButton_Menu_AVI_Sub2_Item[$dd], "Custom: " & $new_fps & " FPS")
				For $d = 0 To UBound($aButton_Menu_AVI_Sub2_Item) - 1
					If $d = $dd Then
						GUICtrlSetState($aButton_Menu_AVI_Sub2_Item[$d], $GUI_CHECKED)
					Else
						GUICtrlSetState($aButton_Menu_AVI_Sub2_Item[$d], $GUI_UNCHECKED)
					EndIf
				Next
			EndIf
		Case $Button_Menu_GrabScreen_Item
			If ($Grab_x + $Grab_y > 0) Or ($Grab_w + $Grab_h > 0) Then
				GUISetState(@SW_HIDE, $hGUI)
				Sleep(350)
				Grab_Region($Grab_x, $Grab_y, $Grab_w, $Grab_h)
				GUISetState(@SW_SHOW, $hGUI)
			EndIf
		Case $Button_Menu_Web_Item
			If BitAND(GUICtrlRead($Button_Menu_Web_Item), $GUI_CHECKED) = $GUI_CHECKED Then
				GUICtrlSetState($Button_Menu_Web_Item, $GUI_UNCHECKED)
			Else
				GUICtrlSetState($Button_Menu_Web_Item, $GUI_CHECKED)
			EndIf
	EndSwitch
	Return "GUI_RUNDEFMSG"
EndFunc   ;==>WM_COMMAND

Func WM_HSCROLL($hWnd, $Msg, $wParam, $lParam)
	#forceref $hWnd, $Msg, $wParam, $lParam
	Local $pos
	If $lParam = $hSlider Then
		$pos = _GUICtrlSlider_GetPos($hSlider)
		GUICtrlSetData($Slider_Label_Value, $pos)
		Slider_Label_Color($pos)
	EndIf
EndFunc   ;==>WM_HSCROLL

Func WM_CONTEXTMENU($hWnd, $Msg, $wParam, $lParam)
	#forceref $hWnd, $Msg, $wParam, $lParam
	Local $mi = GUIGetCursorInfo($hGUI)
	If Not @error Then
		If $mi[4] = $Gfx Then
			_GUICtrlMenu_TrackPopupMenu($hQMenu, $hWnd)
			Return True
		EndIf
	EndIf
	Return "GUI_RUNDEFMSG"
EndFunc   ;==>WM_CONTEXTMENU

Func WM_SYSCOMMAND($hWnd, $Msg, $wParam, $lParam)
	#forceref $hWnd, $Msg, $wParam, $lParam
	Local $idFrom
	$idFrom = BitAND($wParam, 0x0000FFFF)
	Switch $idFrom
		Case $id_Ruler
			Ruler()
		Case $id_OpenMailClient
			OpenMailClient()
		Case $id_Print
			PrintImage()
		Case $id_ChkUpd
			Check4Update()
		Case $id_VisitWeb
			ShellExecute("http://www.rabimba.com")
		Case $id_About
			About()
		Case $id_Exit
			_Exit()
	EndSwitch
	Return "GUI_RUNDEFMSG"
EndFunc   ;==>WM_SYSCOMMAND

Func WM_NOTIFY($hWnd, $MsgID, $wParam, $lParam)
	#forceref $hWnd, $MsgID, $wParam
	Local $hWndFrom, $iCode, $tNMHDR, $hWndListView, $tInfo, $index, $subitem
	$hWndListView = $hListView
	If Not IsHWnd($hListView) Then $hWndListView = GUICtrlGetHandle($List)
	$tNMHDR = DllStructCreate($tagNMHDR, $lParam)
	$hWndFrom = HWnd(DllStructGetData($tNMHDR, "hWndFrom"))
	$iCode = DllStructGetData($tNMHDR, "Code")
	Switch $hWndFrom
		Case $hWndListView
			Switch $iCode
				Case $NM_CLICK ; Sent by a list-view control when the user clicks an item with the left mouse button
					$tInfo = DllStructCreate($tagNMITEMACTIVATE, $lParam)
					$index = DllStructGetData($tInfo, "Index")
					If $index > -1 Then Capture_Window(_GUICtrlListView_GetItemText($List, $index, 1), $aWnd[$index][2], $aWnd[$index][3])
				Case $LVN_KEYDOWN
					$tInfo = DllStructCreate($tagNMLVKEYDOWN, $lParam)
					Switch DllStructGetData($tInfo, 'VKey')
						Case 38
							$index = _GUICtrlListView_GetSelectionMark($List) - 1
							If $index > -1 Then Capture_Window(_GUICtrlListView_GetItemText($List, $index, 1), $aWnd[$index][2], $aWnd[$index][3])
						Case 40
							$index = _GUICtrlListView_GetSelectionMark($List) + 1
							If $index < _GUICtrlListView_GetItemCount($hListView) Then Capture_Window(_GUICtrlListView_GetItemText($List, $index, 1), $aWnd[$index][2], $aWnd[$index][3])
					EndSwitch
				Case $NM_RDBLCLK
					$index = _GUICtrlListView_GetSelectionMark($List)
					If $index > -1 And $index < _GUICtrlListView_GetItemCount($List) And _GUICtrlListView_GetSelectedCount($List) > 0 Then Capture_Window2(_GUICtrlListView_GetItemText($List, $index, 1))
				Case $LVN_COLUMNCLICK
					$tInfo = DllStructCreate($tagNMLISTVIEW, $lParam)
					$subitem = DllStructGetData($tInfo, "SubItem")
					_GUICtrlListView_SimpleSort($hWndFrom, $B_DESCENDING, $subitem)
					Switch $subitem
						Case 0
							_ArraySort($aWnd, Not $B_DESCENDING[0], 0, 0, $subitem)
						Case 1
							_ArraySort($aWnd, Not $B_DESCENDING[1], 0, 0, $subitem)
					EndSwitch
			EndSwitch
	EndSwitch
	Return "GUI_RUNDEFMSG"
EndFunc   ;==>WM_NOTIFY

Func WM_MOUSEWHEEL($hWnd, $iMsg, $wParam, $lParam)
	Local $wheel_Dir = _WinAPI_HiWord($wParam)
	If $wheel_Dir > 0 Then
		If $zoom_level < $zoom_max Then
			$zoom_level += 1
		EndIf
	Else
		If $zoom_level > $zoom_min Then
			$zoom_level -= 1
		EndIf
	EndIf
	Return "GUI_RUNDEFMSG"
EndFunc   ;==>WM_MOUSEWHEEL
#endregion Windows Message Codes
#region Ruler
Func Ruler()
	GUISetState(@SW_HIDE, $hGUI)
	Local $zl = $zoom_level, $v, $esc = True
	$hGUI_Zoom = GUICreate("", $zw, $zh, $zx, $zx, BitOR($WS_POPUP, $DS_MODALFRAME), BitOR($WS_EX_OVERLAPPEDWINDOW, $WS_EX_TOPMOST, $WS_EX_WINDOWEDGE), $hGUI) ;WinGetHandle(AutoItWinGetTitle()))
	GUISetState(@SW_SHOW, $hGUI_Zoom)
	$hDC_Zoom = _WinAPI_GetDC(0)
	$hGUI_ZoomDC = _WinAPI_GetDC($hGUI_Zoom)
	$zx = @DesktopWidth - $zw - 4
	$zy = 0

	WinMove($hGUI_Zoom, "", $zx, $zy)
	Local $mpos, $cmpos, $aDPI_x, $aDPI_y
	Local $hGUI_Ruler = GUICreate("Ruler", $aFullScreen[2], $aFullScreen[3], $aFullScreen[0], $aFullScreen[1], $WS_POPUP, $WS_EX_LAYERED + $WS_EX_TOPMOST, $hGUI);, WinGetHandle(AutoItWinGetTitle()))
	GUISetBkColor(0xABCDEF)
	_WinAPI_SetLayeredWindowAttributes($hGUI_Ruler, 0xABCDEF, 0xFF)
	GUISetState()
	If Not $aero Then WinSetTrans($hGUI_Ruler, "", 0x80)
	WinSetOnTop($hGUI_Zoom, "", 1)
	Local $hGraphic_Ruler = _GDIPlus_GraphicsCreateFromHWND($hGUI_Ruler)
	_GDIPlus_GraphicsSetSmoothingMode($hGraphic_Ruler, 0)
	_GDIPlus_GraphicsClear($hGraphic_Ruler, 0xFFABCDEF)
	Local $hPen_Ruler = _GDIPlus_PenCreate(0xFFFF0000)
	$aDPI_x = DllCall($ghGDIPDll, "uint", "GdipGetDpiX", "handle", $hGraphic_Ruler, "float*", 0)
	$aDPI_y = DllCall($ghGDIPDll, "uint", "GdipGetDpiY", "handle", $hGraphic_Ruler, "float*", 0)

	GUIRegisterMsg(0x020A, "WM_MOUSEWHEEL")
	Local $mo = MouseGetCursor()
	Local $dy = 80
	$up = -1
	While Not _IsPressed("1B", $dll) * Sleep(50)
		GUISetCursor(3, 1, $hGUI_Freehand)
		$red = Hex(0xFF * (Cos($v * 1.10) + 1) / 2, 2)
		$green = Hex(0xFF * (Sin($v * 1.00) + 1) / 2, 2)
		$blue = Hex(0xFF * (Sin($v * 1.20) + 1) / 2, 2)
		Draw_Zoom_Preview(MouseGetPos(0), MouseGetPos(1), 0.5)
		ToolTip("Press LMB to select start point" & @LF & _
				"x: " & MouseGetPos(0) & ", y: " & MouseGetPos(1), MouseGetPos(0) + 10, MouseGetPos(1) + $dy, "", 0, 4)
		If MouseGetPos(1) > ($aFullScreen[3] - 100) Then $dy = -100
		If MouseGetPos(1) < ($aFullScreen[1] + 80) Then $dy = 80

		If _IsPressed("01", $dll) Then
			ToolTip("")
			$mpos = MouseGetPos()
			$dy = 80
			Sleep(150)
			Do
				ToolTip(	"Press LMB to select end point" & @LF & _
								"Start x: " & $mpos[0] & ", Start y: " & $mpos[1] & @LF & _
								"x: " & MouseGetPos(0) & ", y: " & MouseGetPos(1) & @LF & _
								"Distance: " & Round(Pixel_Distance($mpos[0], $mpos[1], MouseGetPos(0), MouseGetPos(1)), 0) & " pixel", MouseGetPos(0) + 10, MouseGetPos(1) + $dy, "", 0, 4)

				If MouseGetPos(1) > ($aFullScreen[3] - 110) Then $dy = -110
				If MouseGetPos(1) < ($aFullScreen[1] + 90) Then $dy = 90
				GUISetCursor(14, 1, $hGUI_Freehand)
				Draw_Zoom_Preview(MouseGetPos(0), MouseGetPos(1), 0.5)
				$red = Hex(0xFF * (Cos($v * 1.10) + 1) / 2, 2)
				$green = Hex(0xFF * (Sin($v * 1.00) + 1) / 2, 2)
				$blue = Hex(0xFF * (Sin($v * 1.20) + 1) / 2, 2)
				$v += 0.15
				_GDIPlus_GraphicsClear($hGraphic_Ruler, 0xFFABCDEF)
				$cmpos = MouseGetPos()
				_GDIPlus_GraphicsDrawLine($hGraphic_Ruler, $mpos[0] + Abs($aFullScreen[0]), $mpos[1] + Abs($aFullScreen[1]), $cmpos[0] + Abs($aFullScreen[0]), $cmpos[1] + Abs($aFullScreen[1]), $hPen_Ruler)
				Sleep(20)
			Until _IsPressed("01", $dll)
			$esc = False
			ExitLoop
		EndIf
		$v += 0.15
	WEnd
	ToolTip("")
	GUISetCursor($mo, 1, $hGUI_Freehand)
	_GDIPlus_PenDispose($hPen_Ruler)
	_GDIPlus_GraphicsDispose($hGraphic_Ruler)
	_WinAPI_ReleaseDC($hGUI_Zoom, $hGUI_ZoomDC)
	_WinAPI_ReleaseDC(0, $hDC_Zoom)
	GUIDelete($hGUI_Zoom)
	If Not $esc And IsArray($cmpos) And IsArray($mpos) Then
		WinSetOnTop($hGUI_Ruler, "", 0)
		Local $dist = Pixel_Distance($mpos[0], $mpos[1], $cmpos[0], $cmpos[1])
		MsgBox(0, "Rk Information Ruler", _
				"Start point x: " & $mpos[0] & @LF & _
				"Start point y: " & $mpos[1] & @LF & _
				"End point x: " & $cmpos[0] & @LF & _
				"End point y: " & $cmpos[1] & @LF & _
				"DPI width: " & $aDPI_x[2] & @LF & _
				"DPI height: " & $aDPI_y[2] & @LF & @LF & _
				"Distance:" & _
				@TAB & Round($dist, 0) & " pixel" & @LF & _
				@TAB & StringFormat("%05.2f", $dist * 2.54 / $aDPI_x[2]) & " centimeter" & @LF & _
				@TAB & StringFormat("%05.2f", $dist / $aDPI_x[2]) & " inch", 60, $hGUI)
	EndIf
	GUIDelete($hGUI_Ruler)
	GUIRegisterMsg(0x020A, "")
	$zoom_level = $zl
	GUISetState(@SW_SHOW, $hGUI)
EndFunc   ;==>Ruler
#endregion Ruler
#region Send mail
Func OpenMailClient()
	$silent = True
	If Not Pic2Clipboard() Then Return MsgBox(16, "ERROR", "Unable to put image to clipboard!", 30)
	$silent = False
	Local $s = _INetMail("", "Rk Screencapture", "<Paste your image from your clipboard here>")
	If @error Then Return MsgBox(16, "ERROR", "Unable to start default mail client!", 30)
EndFunc   ;==>OpenMailClient
#endregion Send mail
#region About Intro
Func About()
	GUIRegisterMsg($WM_NOTIFY, "")
	GUIRegisterMsg($WM_SYSCOMMAND, "")
	GUIRegisterMsg($WM_CONTEXTMENU, "")
	GUISetStyle(Default, $WS_EX_TOPMOST, $hGUI)
	Local $z
	Opt("GUICloseOnESC", 1)
	GUISetOnEvent($GUI_EVENT_CLOSE, "")
	#region binary string read
	Local $BassModDLL = BassModDLL()
	Local $_XM_Tune = ChipMuzik()
	$F_DLL = MemoryDllOpen($BassModDLL)
	Local $chip = Binary($_XM_Tune)
	Local $len = BinaryLen($chip)
	Local $tMem = DllStructCreate("byte[" & $len & "]")
	DllStructSetData($tMem, 1, $chip)

	MemoryDllCall($F_DLL, "int", "BASSMOD_Init", "int", -1, "dword", 44100, "dword", 0)
	MemoryDllCall($F_DLL, "int", "BASSMOD_MusicLoad", _
														"bool", True, _
														"ptr", DllStructGetPtr($tMem), _
														"dword", 0, _
														"dword", $len, _
														"dword", 4 + 1024)
	MemoryDllCall($F_DLL, "int", "BASSMOD_MusicPlay")
	#endregion binary string read

	Local $fade, $size
	If @OSBuild < 6000 Or Not $aero Then
		$fade = 1
		$size = 3
	Else
		$fade = 0.075
	EndIf

	$hGUI_About = GUICreate("", $width - $size, $height - $size, 0, 0, $WS_POPUPWINDOW + $WS_CLIPSIBLINGS, $WS_EX_MDICHILD, $hGUI)
	WinSetTrans($hGUI_About, "", 0x00)
	GUISetState(@SW_SHOW, $hGUI_About)

	#region GDI+ Init
	$hGfx = _GDIPlus_GraphicsCreateFromHWND($hGUI_About)
	$hBMP_About = _GDIPlus_BitmapCreateFromGraphics($width, $height, $hGfx)
	$hCtxt = _GDIPlus_ImageGetGraphicsContext($hBMP_About)
	DllCall($ghGDIPDll, "uint", "GdipSetPixelOffsetMode", "hwnd", $hCtxt, "int", 2)
	_GDIPlus_GraphicsSetSmoothingMode($hCtxt, 2)
	DllCall($ghGDIPDll, "uint", "GdipSetTextRenderingHint", "handle", $hCtxt, "int", 4) ;http://msdn.microsoft.com/en-us/library/ms535817(v=vs.85).aspx

	$hHBITMAP = _ScreenCapture_CaptureWnd("", $hGUI, 0, 0, -1, -1, False)
	$hBmp_A = _GDIPlus_BitmapCreateFromHBITMAP($hHBITMAP)
	_WinAPI_DeleteObject($hHBITMAP)
	$hImg_A = _Blur($hBmp_A, $width, $height)
	_GDIPlus_GraphicsDrawImage($hGfx, $hImg_A, 0, 0)
	#endregion GDI+ Init

	#region Dancing Animation
	Local $bAnim = GIF_Anim()
	Local $aGIFDim = _GIF_GetDimension($bAnim)
	Local $hGUI_Anim = GUICreate("", $aGIFDim[0], $aGIFDim[1], $width - $aGIFDim[0], $height - $aGIFDim[1], $WS_POPUP, BitOR($WS_EX_LAYERED, $WS_EX_MDICHILD), $hGUI_About)
	GUISetBkColor(0xABCDEF, $hGUI_Anim)
	WinSetTrans($hGUI_Anim, "", 0x00)
	Local $hAnim = _GUICtrlCreateGIF($bAnim, "", 0, 0, $aGIFDim[0], $aGIFDim[1])
	_WinAPI_SetLayeredWindowAttributes($hGUI_Anim, 0xABCDEF)
	GUISetState(@SW_SHOW, $hGUI_Anim)
	#endregion Dancing Animation

	Local $i
	For $i = 0 To $bubbles - 1
		$aData[$i][0] = Random(0, $width - $max_size, 1)
		$aData[$i][1] = Random(0, $height - $max_size, 1)
		$aData[$i][2] = _Random(-$max_speed, $max_speed, -1.5, 1.5) ;vx
		$aData[$i][3] = _Random(-$max_speed, $max_speed, -1.5, 1.5) ;vy
		$aData[$i][4] = Random($min_size, $max_size, 1) ;size
		$aData[$i][5] = _CreateBubbleBitmap($aData[$i][4]) ;handle to bitmap
	Next

	#region fade-in GUI
	Local Const $snd_vol = 0.75
	Local Const $zz = 100 / 0xFF
	MemoryDllCall($F_DLL, "int", "BASSMOD_SetVolume", "int",  0)
	For $z = 0 To 0xFF - $fade Step $fade
		WinSetTrans($hGUI_About, "", $z)
		MemoryDllCall($F_DLL, "int", "BASSMOD_SetVolume", "int",  ($z * $zz) * $snd_vol)
	Next
	#endregion fade-in GUI

	#region Init Scroller
	Local $bFont1 = Ethnocentric_Font()
	Local $bFont2 = UEZ_Logo_Font()
	Local $tFont_Mem1 = DllStructCreate('byte[' & BinaryLen($bFont1) & ']') ;Ethnocentric
	Local $tFont_Mem2 = DllStructCreate('byte[' & BinaryLen($bFont2) & ']') ;Ethnocentric
	DllStructSetData($tFont_Mem1, 1, $bFont1)
	DllStructSetData($tFont_Mem2, 1, $bFont2)
	Local $hCollection = _GDIPlus_NewPrivateFontCollection()
	_GDIPlus_PrivateAddMemoryFont($hCollection, DllStructGetPtr($tFont_Mem1), DllStructGetSize($tFont_Mem1))
	_GDIPlus_PrivateAddMemoryFont($hCollection, DllStructGetPtr($tFont_Mem2), DllStructGetSize($tFont_Mem2))

	$tRectF = _GDIPlus_RectFCreate(0, 0, $width / 5, $height / 5)
	$hBrush = _GDIPlus_LineBrushCreateFromRectWithAngle($tRectF, 0x60201010, 0xFFF04040, 45, True, 1)
	$hFormat = _GDIPlus_StringFormatCreate()
	#region scroller text
	;																text,						$hFamily, $hFont,	 font type,			font size
	Dim $aText[40][9] = [ _
											["CTS", 							   				   	1, 0,	"Ethnocentric",				56], _
											["Rk Screencapture", 					0, 0,	"Georgia",			48], _
											[$ver,															0, 0,	"Comic Sans MS",			24], _
											[" ", 																0, 0,	"Arial", 							50], _
											["Main Code", 											0, 0, "Verdana",						30], _
											["by", 															0, 0, "Times New Roman",	40], _
											["Rk", 														2, 0, "Vtks Revolt",		  	  180], _
											[" ", 																0, 0,	"Arial", 						  	20], _
						["Thanks to:",											0, 0, "Tunga", 						40], _
											[" ", 																0, 0,	"Arial", 						  	20], _
											["un4seen for BassMod.dll",					0, 0, "Palatino Linotype",		36], _
											[" ", 																0, 0,	"Arial", 						  120], _
											["Press ESC to quit",									0, 0, "Latha",							30], _
											[" ", 																0, 0,	"Arial", 						  120], _
											[ChrW(9996),												0, 0,	"Arial", 						  500]]
	#endregion
	$tLayout = _GDIPlus_RectFCreate(0, 0, 0, 0)
	Local $y = $height, $sy = 0, $aInfo
	Local Const $dy = 10

	For $z = 0 To UBound($aText) - 1
		If $aText[$z][1]  Then
			$aText[$z][1] = _GDIPlus_CreateFontFamilyFromName($aText[$z][3], $hCollection)
		Else
			$aText[$z][1] = _GDIPlus_FontFamilyCreate($aText[$z][3]) ;$hFamily
		EndIf
		$aText[$z][2] = _GDIPlus_FontCreate($aText[$z][1], $aText[$z][4]) ;$hFont
		$aInfo = _GDIPlus_GraphicsMeasureString($hGfx, $aText[$z][0], $aText[$z][2], $tLayout, $hFormat)
		$aText[$z][5] = Floor(DllStructGetData($aInfo[0], "Width"))
		$aText[$z][6] = Floor(DllStructGetData($aInfo[0], "Height"))
		$aText[$z][7] = Floor($width / 2 - ($aText[$z][5] / 2))
		$aText[$z][8] = $y + $dy + $sy
		$sy += $aText[$z][6]
	Next
	$speed = 1
	#endregion Init Scroller

	GUIRegisterMsg($WM_TIMER, "DrawGDIp") ;$WM_TIMER = 0x0113
	DllCall($dll, "int", "SetTimer", "hwnd", $hGUI_About, "int", 0, "int", 50, "int", 0)

	#region Sleep until ESC is pressed
	While Sleep(25) * Not _IsPressed("1B", $dll)
		If $About_End Then ExitLoop
	WEnd
	#endregion Sleep until ESC is pressed
	GUIRegisterMsg($WM_TIMER, "")

	#region Fade-out
	For $z = 0xFF - $fade To 0 Step -$fade
		WinSetTrans($hGUI_About, "", $z)
		MemoryDllCall($F_DLL, "int", "BASSMOD_SetVolume", "int", ($z * $zz) * $snd_vol)
	Next
	_GIF_DeleteGIF($hAnim)
	_GIF_ExitAnimation($hAnim)
	GUIDelete($hGUI_Anim)

	GUISetState(@SW_HIDE, $hGUI_About)
	_WinAPI_RedrawWindow($hGUI, 0, 0, $RDW_NOFRAME + $RDW_NOINTERNALPAINT + $RDW_NOERASE)
	GUIDelete($hGUI_About)
	_WinAPI_RedrawWindow($hGUI, 0, 0, $RDW_NOFRAME + $RDW_NOINTERNALPAINT + $RDW_NOERASE)
	#endregion Fade-out

	#region Release Resources and Exit

	GUIRegisterMsg($WM_NOTIFY, "WM_NOTIFY")
	GUIRegisterMsg($WM_SYSCOMMAND, "WM_SYSCOMMAND")
	GUIRegisterMsg($WM_CONTEXTMENU, "WM_CONTEXTMENU")

	For $z = 0 To UBound($aText) - 1
		_GDIPlus_FontDispose($aText[$z][2])
		_GDIPlus_FontFamilyDispose($aText[$z][1])
	Next
	_GDIPlus_DeletePrivateFontCollection($hCollection)
	_GDIPlus_StringFormatDispose($hFormat)
	_GDIPlus_BrushDispose($hBrush)
	_GDIPlus_BitmapDispose($hImg_A)
	_GDIPlus_BitmapDispose($hBmp_A)
	_GDIPlus_BitmapDispose($hBMP_About)
	_GDIPlus_GraphicsDispose($hCtxt)
	_GDIPlus_GraphicsDispose($hGfx)
	For $z = 0 To $bubbles - 1
		_GDIPlus_BitmapDispose($aData[$z][5])
	Next

	MemoryDllCall($F_DLL, "int", "BASSMOD_MusicStop")
	MemoryDllCall($F_DLL, "int", "BASSMOD_MusicFree")
	MemoryDllClose($F_DLL)

	$tFont_Mem1 = 0
	$tFont_Mem2 = 0
	$tLayout = 0
	$tMem = 0
	$BassModDLL = 0
	$_XM_Tune = 0
	$bAnim = 0
	$tRectF = 0
	Opt("GUICloseOnESC", 0)
	GUISetOnEvent($GUI_EVENT_CLOSE, "_Exit")
	_ReduceMemory(@AutoItPID)
	#endregion Release Resources and Exit
EndFunc   ;==>About

Func DrawGDIp()
	_GDIPlus_GraphicsDrawImage($hCtxt, $hImg_A, 0, 0)
	For $z = 0 To UBound($aText) - 1
		If $aText[$z][8] < $height And $aText[$z][8] > -$aText[$z][6] Then
			$tLayout = _GDIPlus_RectFCreate($aText[$z][7], $aText[$z][8], 0, 0)
			_GDIPlus_GraphicsDrawStringEx($hCtxt, $aText[$z][0], $aText[$z][2], $tLayout, $hFormat, $hBrush)
		EndIf
		$aText[$z][8] -= $speed
	Next
	Bubble_Anim()
	_GDIPlus_GraphicsDrawImageRect($hGfx, $hBMP_About, 0, 0, $width, $height)
	If $aText[UBound($aText) - 1][8] < -$aText[UBound($aText) - 1][6] Then
		$About_End = True
		GUIRegisterMsg($WM_TIMER, "")
	EndIf
EndFunc   ;==>DrawGDIp

Func Bubble_Anim()
	Local $j
	;draw bubbles
	For $j = 0 To $bubbles - 1
		_GDIPlus_GraphicsDrawImage($hCtxt, $aData[$j][5], $aData[$j][0], $aData[$j][1])
	Next
	;calculate new position incl. border collision check
	For $j = 0 To $bubbles - 1
		$aData[$j][0] += $aData[$j][2] ;increase x coordinate with appropriate slope
		$aData[$j][1] += $aData[$j][3] ;increase y coordinate with appropriate slope
		If $aData[$j][0] <= 0 Then ;border collision x left
			$aData[$j][0] = 1
			$aData[$j][2] *= -1
		ElseIf $aData[$j][0] >= $width - $aData[$j][4] Then ;border collision x right
			$aData[$j][0] = $width - ($aData[$j][4] + 1)
			$aData[$j][2] *= -1
		EndIf
		If $aData[$j][1] <= 0 Then ;border collision y top
			$aData[$j][1] = 1
			$aData[$j][3] *= -1
		ElseIf $aData[$j][1] >= $height - $aData[$j][4] - $dh Then ;border collision y bottom
			$aData[$j][1] = $height - ($aData[$j][4] + 1) - $dh
			$aData[$j][3] *= -1
		EndIf
	Next
	Collision_Check()
EndFunc   ;==>Bubble_Anim

Func Collision_Check() ;0:x, 1:y, 2:vx, 3:vy, 4:size, 5:bmp
	Local $Delta_X, $Delta_Y, $Distance, $Matrix11, $Matrix12, $Matrix21, $Matrix22, $Vp1, $Vp2, $Vs1, $Vs2
	For $i = 0 To $bubbles - 1
		For $j = $i + 1 To $bubbles - 1
			If Pixel_Distance($aData[$i][0], $aData[$i][1], $aData[$j][0], $aData[$j][1]) < ($aData[$i][4] + $aData[$j][4]) / 2 Then
				$Delta_X = $aData[$i][0] - $aData[$j][0]
				$Delta_Y = $aData[$i][1] - $aData[$j][1]
				$Distance = Sqrt($Delta_X * $Delta_X + $Delta_Y * $Delta_Y)

				$Matrix11 = $Delta_X / $Distance
				$Matrix12 = -$Delta_Y / $Distance
				$Matrix21 = $Delta_Y / $Distance
				$Matrix22 = $Delta_X / $Distance

				$Vp1 = $aData[$i][2] * $Matrix11 + $aData[$i][3] * - $Matrix12
				$Vs1 = $aData[$i][2] * - $Matrix21 + $aData[$i][3] * $Matrix22
				$Vp2 = $aData[$j][2] * $Matrix11 + $aData[$j][3] * - $Matrix12
				$Vs2 = $aData[$j][2] * - $Matrix21 + $aData[$j][3] * $Matrix22

				If $Vp1 - $Vp2 < 0 Then
					$aData[$i][2] = $Matrix11 + $Vs1 * $Matrix12
					$aData[$i][3] = $Matrix21 + $Vs1 * $Matrix22
					$aData[$j][2] = $Matrix11 + $Vs2 * $Matrix12
					$aData[$j][3] = $Matrix21 + $Vs2 * $Matrix22
				EndIf
			EndIf
		Next
	Next
EndFunc   ;==>Collision_Check

Func Pixel_Distance($x1, $y1, $x2, $y2) ;Pythagoras theorem
	Local $a, $b
	If $x2 = $x1 And $y2 = $y1 Then Return 0
	$a = $y2 - $y1
	$b = $x2 - $x1
	Return Sqrt($a * $a + $b * $b)
EndFunc   ;==>Pixel_Distance

Func _CreateBubbleBitmap($size = 75, $gradient_start = 0xA0F0C0C0, $gradient_end = 0xA0C0F0C0, $angle1 = 90, $angle2 = 45, $pen1 = 0x55AAAAAF, $pen2 = 0x50FFFFFF)
	Local $ps1 = 2, $ps2 = Ceiling($size / 12)
	Local $hBitmap = _GDIPlus_BitmapCreateFromScan0($size + 1, $size + 1)
	Local $hContext = _GDIPlus_ImageGetGraphicsContext($hBitmap)
	_GDIPlus_GraphicsSetSmoothingMode($hContext, 2)
	Local $Pen_Border = _GDIPlus_PenCreate($pen1, $ps1)
	Local $Pen_Reflection = _GDIPlus_PenCreate($pen2, $ps2)
	Local $Brush_Gradient = _GDIPlus_CreateLineBrush(0, 0, $size, $size, $gradient_start, $gradient_end)
	_GDIPlus_GraphicsFillEllipse($hContext, 0, 0, $size, $size, $Brush_Gradient)
	_GDIPlus_GraphicsDrawArc($hContext, $size / 3, $size / 5, $size / 2, $size / 2, 10, -$angle1, $Pen_Reflection)
	_GDIPlus_GraphicsDrawArc($hContext, $size / 6, $size / 3.5, $size / 2, $size / 2, -210, -$angle2, $Pen_Reflection)
	_GDIPlus_GraphicsDrawEllipse($hContext, 0, 0, $size - $ps1 / 2, $size - $ps1 / 2, $Pen_Border)
	_GDIPlus_PenDispose($Pen_Border)
	_GDIPlus_PenDispose($Pen_Reflection)
	_GDIPlus_BrushDispose($Brush_Gradient)
	_GDIPlus_GraphicsDispose($hContext)
	Return $hBitmap
EndFunc   ;==>_CreateBubbleBitmap

Func _Random($min, $max, $emin, $emax, $int = 0)
	Local $r1 = Random($min, $emin, $int)
	Local $r2 = Random($emax, $max, $int)
	If Random(0, 1, 1) Then Return $r1
	Return $r2
EndFunc   ;==>_Random
#endregion About Intro


#region Some GUI elements
Func Refresh_Wnd_List()
	_GUICtrlListView_DeleteAllItems(GUICtrlGetHandle($List))
	$aWnd = GetAllWindow()
	Local $i
	For $i = 0 To UBound($aWnd) - 1
		_GUICtrlListView_AddItem($List, $aWnd[$i][0])
		_GUICtrlListView_AddSubItem($List, $i, $aWnd[$i][1], 1)
	Next
	Dim $B_DESCENDING[_GUICtrlListView_GetColumnCount($hListView)]
;~ 	_GUICtrlListView_SimpleSort($hListView, $B_DESCENDING, 0)
;~ 	_ArraySort($aWnd, Not $B_DESCENDING[0], 0, 0, 0)
EndFunc   ;==>Refresh_Wnd_List

Func _Exit()
	GUIRegisterMsg($WM_COMMAND, "")
	GUIRegisterMsg($WM_CONTEXTMENU, "")
	GUIRegisterMsg($WM_NOTIFY, "")
	GUIRegisterMsg($WM_SYSCOMMAND, "")
	GUIRegisterMsg($WM_TIMER, "")
	GUIRegisterMsg($WM_HSCROLL, "")
	GUIRegisterMsg(0x020A, "")
	DllCall("user32.dll", "int", "UnhookWindowsHookEx", "hwnd", $hM_Hook[0])
	$hM_Hook[0] = 0
	DllCallbackFree($hKey_Proc)
	$hKey_Proc = 0
	GUISetOnEvent($GUI_EVENT_PRIMARYDOWN, "")
	DllClose($dll)
	DllClose($hGIFDLL__KERNEL32)
	DllClose($hGIFDLL__USER32)
	DllClose($hGIFDLL__GDI32)
	DllClose($hGIFDLL__COMCTL32)
	DllClose($hGIFDLL__OLE32)
	DllClose($hGIFDLL__GDIPLUS)
	DllClose($hDwmApiDll)
	_WinAPI_DeleteObject($hBMP_Quality)
	_WinAPI_DeleteObject($hBMP_Reset)
	_WinAPI_DeleteObject($hBMP_Reset2)
	_WinAPI_DeleteObject($hBMP_ImageEdit)
	_WinAPI_DeleteObject($hBMP_ImageEdit_Gray)
	_WinAPI_DeleteObject($hBMP_ImageEdit_BW)
	_WinAPI_DeleteObject($hBMP_ImageEdit_Invert)
	_WinAPI_DeleteObject($hBMP_ImageEdit_Rotm90)
	_WinAPI_DeleteObject($hBMP_ImageEdit_Rotp90)
	_WinAPI_DeleteObject($hBMP_ImageEdit_Undo)
	_WinAPI_DeleteObject($hBMP_ImageEdit_Editor)
	_WinAPI_DeleteObject($hBMP_ChkUpd)
	_WinAPI_DeleteObject($hBMP_VisitWeb)
	_WinAPI_DeleteObject($hBMP_Print)
	_WinAPI_DeleteObject($hBMP_About)
	_WinAPI_DeleteObject($hBMP_Exit)
	_WinAPI_DeleteObject($hBitmap_s)
	_WinAPI_DeleteObject($hBMP_Menu_GrabScreen_Redo)
	_WinAPI_DeleteObject($hBMP_Menu_Timestamp)
	_WinAPI_DeleteObject($hBMP_Menu_JPG_Qual)
	_WinAPI_DeleteDC($hMemDC)
	_WinAPI_DeleteObject($memBitmap)
	_WinAPI_ReleaseDC(0, $hDC_Region)
	_WinAPI_DeleteDC($hDC_Region)
	_GDIPlus_GraphicsDispose($hImageContext)
	_GDIPlus_BitmapDispose($hBackImage)
	_GDIPlus_BitmapDispose($vBitmap)
	_GDIPlus_BitmapDispose($undo)
	_GDIPlus_MatrixDispose($hMatrix)
	_GDIPlus_BrushDispose($hBrush_Clear)
	_GDIPlus_BitmapDispose($hBmp)
	_GDIPlus_GraphicsDispose($hContext)
	_GDIPlus_BitmapDispose($hBuffer_Bmp)
	_GDIPlus_GraphicsDispose($hGraphic)
	_GDIPlus_Shutdown()
	If @OSBuild > 5999 Then _WinAPI_AnimateWindow($hGUI, BitOR($AW_BLEND, $AW_HIDE), 750)
	GUIDelete($hGUI)
	Exit
EndFunc   ;==>_Exit
#endregion Some GUI elements

#region external functions
#region _ScriptRestart by Yashied
Func _ScriptRestart($fExit = 1)
	Local $Pid
	If Not $__Restart Then
		If @compiled Then
			$Pid = Run(@ScriptFullPath & ' ' & $CmdLineRaw, @ScriptDir, Default, 1)
		Else
			$Pid = Run(@AutoItExe & ' "' & @ScriptFullPath & '" ' & $CmdLineRaw, @ScriptDir, Default, 1)
		EndIf
		If @error Then
			Return SetError(@error, 0, 0)
		EndIf
		StdinWrite($Pid, @AutoItPID)
	EndIf
	$__Restart = 1
	If $fExit Then
		Sleep(50)
		Exit
	EndIf
	Return 1
EndFunc   ;==>_ScriptRestart
Func OnAutoItStart()
	Sleep(50)
	Local $Pid = ConsoleRead(1)
	If @extended Then
		While ProcessExists($Pid)
			Sleep(100)
		WEnd
	EndIf
EndFunc   ;==>OnAutoItStart
#endregion
#region extended function for MPDF_UDF.au3 by taietel
; #FUNCTION# ====================================================================================================================
; Name ..........: _LoadResImageMem()
; Description ...: Loads an image into the pdf (if you use it multiple times it decreases the size of the pdf)
; Syntax ........: _LoadResImageMem( $sImgAlias , $bImage  )
; Parameters ....: $sImgAlias           -  an alias to identify the image in the pdf (e.g. "Cheese").
;                  $bImage              -  binary image
; Return values .: Success      - True
;                  Failure      - False
; Author(s) .....: Mihai Iancu (taietel at yahoo dot com)
; Modified ......: UEZ
; Remarks .......: Image types accepted: BMP, GIF, TIF, TIFF, PNG, JPG, JPEG (those are tested)
; Related .......:
; Link ..........: http://www.autoitscript.com/forum/topic/118827-create-pdf-from-your-application/
; Example .......: No
; ===============================================================================================================================
Func _LoadResImageMem($sImgAlias, $bImage, $iW, $iH, $bInterpolate = "true") ;true must be in small letters!
	Local $iObj
	If $sImgAlias = "" Then __Error("You don't have an alias for the image", @ScriptLineNumber)
	If $bImage = "" Then
		__Error("You don't have any images to insert or binary string is invalid", @ScriptLineNumber)
	Else
		$iObj = __InitObj()
;~ 		__ToBuffer("<</Type /XObject /Subtype /Image /Name /" & $sImgAlias & " /Width " & $iW & " /Height " & $iH & " /Filter /DCTDecode /ColorSpace /DeviceRGB /BitsPerComponent 8" & " /Length " & $iObj + 1 & " 0 R" & ">>")
		__ToBuffer("<</Type /XObject /Subtype /Image /Name /" & $sImgAlias & " /Width " & $iW & " /Height " & $iH & " /Filter /DCTDecode /ColorSpace /DeviceRGB /BitsPerComponent 8 /Interpolate " & $bInterpolate & " /Length " & $iObj + 1 & " 0 R" & ">>")
		__ToBuffer("stream" & @CRLF & $bImage & @CRLF & "endstream")
		__EndObj()
		$_Image &= "/" & $sImgAlias & " " & $iObj & " 0 R " & @CRLF
		__InitObj()
		__ToBuffer(BinaryLen($bImage))
		__EndObj()
	EndIf
	Return $_Image
EndFunc   ;==>_LoadResImageMem
#endregion extended function for MPDF_UDF.au3 by taietel
;http://msdn.microsoft.com/en-us/library/aa910369.aspx WAVE
;http://msdn.microsoft.com/en-us/library/aa909803.aspx
#region play wave from memory by  wolf9228
Func sndPlaySound($lpszSound, $fuSound)
	Local $Type = "wstr", $aRes
	If IsPtr($lpszSound) Then $Type = "ptr"
	$aRes = DllCall("winmm.dll", "int", "sndPlaySound", $Type, $lpszSound, "dword", $fuSound)
	If @error Then Return SetError(1, 0, 0)
	Return $aRes[0]
EndFunc   ;==>sndPlaySound
#endregion play wave from memory by  wolf9228
#region Blur by eukalyptus
Func _Blur($hBitmap, $iW, $iH, $fScale = 0.175, $qual = 6); by eukalyptus
	Local $hGraphics = _GDIPlus_GraphicsCreateFromHWND(_WinAPI_GetDesktopWindow())
	Local $hBmpSmall = _GDIPlus_BitmapCreateFromGraphics($iW, $iH, $hGraphics)
	Local $hGfxSmall = _GDIPlus_ImageGetGraphicsContext($hBmpSmall)
	DllCall($ghGDIPDll, "uint", "GdipSetPixelOffsetMode", "hwnd", $hGfxSmall, "int", 2)
	Local $hBmpBig = _GDIPlus_BitmapCreateFromGraphics($iW, $iH, $hGraphics)
	Local $hGfxBig = _GDIPlus_ImageGetGraphicsContext($hBmpBig)
	DllCall($ghGDIPDll, "uint", "GdipSetPixelOffsetMode", "hwnd", $hGfxBig, "int", 2)
	_GDIPlus_GraphicsScaleTransform($hGfxSmall, $fScale, $fScale)
	_GDIPlus_GraphicsSetInterpolationMode($hGfxSmall, $qual)

	_GDIPlus_GraphicsScaleTransform($hGfxBig, 1 / $fScale, 1 / $fScale)
	_GDIPlus_GraphicsSetInterpolationMode($hGfxBig, $qual)

	_GDIPlus_GraphicsDrawImageRect($hGfxSmall, $hBitmap, 0, -5, $iW, $iH + 12)
	_GDIPlus_GraphicsDrawImageRect($hGfxBig, $hBmpSmall, 0, -3, $iW, $iH + 9)

	_GDIPlus_GraphicsDispose($hGraphics)
	_GDIPlus_BitmapDispose($hBmpSmall)
	_GDIPlus_GraphicsDispose($hGfxSmall)
	_GDIPlus_GraphicsDispose($hGfxBig)
	Return $hBmpBig
EndFunc   ;==>_Blur
#endregion Blur by eukalyptus
#region MemFont by eukalyptus
; #FUNCTION# ======================================================================================
; Name ..........: _WinAPI_RemoveFontMemResourceEx()
; Description ...: Removes the fonts added from a memory image file.
; Syntax ........: _WinAPI_RemoveFontMemResourceEx($hFont)
; Parameters ....: $hFont    - [in] A handle to the font-resource.
; Return values .: Success   - True
;                  Failure   - False
; Author ........: Eukalyptus
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: @@MsdnLink@@ RemoveFontMemResourceEx
; Example .......:
; =================================================================================================
Func _WinAPI_RemoveFontMemResourceEx($hFont)
	Local $aResult = DllCall('gdi32.dll', 'boolean', 'RemoveFontMemResourceEx', 'handle', $hFont)
	If @error Then Return SetError(@error, 0, False)
	Return $aResult[0]
EndFunc   ;==>_WinAPI_RemoveFontMemResourceEx

; #FUNCTION# ======================================================================================
; Name ..........: _WinAPI_AddFontMemResourceEx()
; Description ...: Adds the font resource from a memory image to the system.
; Syntax ........: _WinAPI_AddFontMemResourceEx($pbFont, $cbFont, $pdv, $pcFonts)
; Parameters ....: $pbFont   - [in] A pointer to a font resource.
;                  $cbFont   - [in] The number of bytes in the font resource that is pointed to by pbFont.
;                  $pdv      - [in] Reserved. Must be 0.
;                  $pcFonts  - [in] A pointer to a variable that specifies the number of fonts installed.
; Return values .: Success   - the return value specifies the handle to the font added.
;                  Failure   - 0
; Author ........: Eukalyptus
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: @@MsdnLink@@ AddFontMemResourceEx
; Example .......:
; =================================================================================================
Func _WinAPI_AddFontMemResourceEx($pbFont, $cbFont, $pdv=0, $pcFonts=0)
	Local $aResult = DllCall('gdi32.dll', 'handle', 'AddFontMemResourceEx', 'ptr', $pbFont, 'dword', $cbFont, 'ptr', $pdv, 'dword*', $pcFonts)
	If @error Then Return SetError(@error, 0, False)
	Return $aResult[0]
EndFunc   ;==>_WinAPI_AddFontMemResourceEx

; #FUNCTION# ======================================================================================
; Name ..........: _GDIPlus_CreateFontFamilyFromName()
; Description ...: Creates a FontFamily object based on a specified font family.
; Syntax ........: _GDIPlus_CreateFontFamilyFromName($sFontname[, $hCollection = 0])
; Parameters ....: $sFontname   - [in] Name of the font family. For example, Arial.ttf is the name of the Arial font family.
;                  $hCollection - [optional] [in] Pointer to the PrivateFontCollection object to delete. (default:0)
; Return values .: Success      - a pointer to the new FontFamily object.
;                  Failure      - 0
; Author ........: Prog@ndy, Yashied
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: @@MsdnLink@@ GdipCreateFontFamilyFromName
; Example .......:
; =================================================================================================
Func _GDIPlus_CreateFontFamilyFromName($sFontname, $hCollection = 0)
	Local $aResult = DllCall($ghGDIPDll, 'int', 'GdipCreateFontFamilyFromName', 'wstr', $sFontname, 'ptr', $hCollection, 'ptr*', 0)
	If @error Then Return SetError(1, 0, 0)
	Return SetError($aResult[0], 0, $aResult[3])
EndFunc   ;==>_GDIPlus_CreateFontFamilyFromName

; #FUNCTION# ======================================================================================
; Name ..........: _GDIPlus_DeletePrivateFontCollection()
; Description ...: Deletes the specified PrivateFontCollection object.
; Syntax ........: _GDIPlus_DeletePrivateFontCollection($hCollection)
; Parameters ....: $hCollection - [in] Pointer to the PrivateFontCollection object to delete.
; Return values .: Success      - True
;                  Failure      - False
; Author ........: Prog@ndy, Yashied
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: @@MsdnLink@@ GdipDeletePrivateFontCollection
; Example .......:
; =================================================================================================
Func _GDIPlus_DeletePrivateFontCollection($hCollection)
	Local $aResult = DllCall($ghGDIPDll, 'int', 'GdipDeletePrivateFontCollection', 'ptr*', $hCollection)
	If @error Then Return SetError(1, 0, False)
	Return SetError($aResult[0], 0, $aResult[0] = 0)
EndFunc   ;==>_GDIPlus_DeletePrivateFontCollection

; #FUNCTION# ======================================================================================
; Name ..........: _GDIPlus_NewPrivateFontCollection()
; Description ...: Creates an PrivateFontCollection object.
; Syntax ........: _GDIPlus_NewPrivateFontCollection()
; Parameters ....:
; Return values .: Success      - a pointer to the PrivateFontCollection object.
;                  Failure      - 0
; Author ........: Prog@ndy, Yashied
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: @@MsdnLink@@ GdipNewPrivateFontCollection
; Example .......:
; =================================================================================================
Func _GDIPlus_NewPrivateFontCollection()
	Local $aResult = DllCall($ghGDIPDll, 'int', 'GdipNewPrivateFontCollection', 'ptr*', 0)
	If @error Then Return SetError(1, 0, 0)
	Return SetError($aResult[0], 0, $aResult[1])
EndFunc   ;==>_GDIPlus_NewPrivateFontCollection

; #FUNCTION# ======================================================================================
; Name ..........: _GDIPlus_PrivateAddMemoryFont()
; Description ...: Adds a font file from memory to the private font collection.
; Syntax ........: _GDIPlus_PrivateAddMemoryFont($hCollection, $pMemory, $iLength)
; Parameters ....: $hCollection - [in] Pointer to the font collection object.
;                  $pMemory     - [in] A pointer to a font resource.
;                  $iLength     - [in] The number of bytes in the font resource that is pointed to by $pMemory.
; Return values .: Success      - True
;                  Failure      - False
; Author ........: Eukalyptus
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: @@MsdnLink@@ GdipPrivateAddMemoryFont
; Example .......:
; =================================================================================================
Func _GDIPlus_PrivateAddMemoryFont($hCollection, $pMemory, $iLength)
	Local $aResult = DllCall($ghGDIPDll, 'int', 'GdipPrivateAddMemoryFont', 'ptr', $hCollection, 'ptr', $pMemory, 'int', $iLength)
	If @error Then Return SetError(1, 0, False)
	Return SetError($aResult[0], 0, $aResult[0] = 0)
EndFunc   ;==>_GDIPlus_PrivateAddMemoryFont
#endregion MemFont by eukalyptus
#region Image Editing
Func ASM_Bitmap_Grey_BnW($vBmp, $iBlackAndWhite = 0, $iLight = 128) ;ASM code by AndyG
	Local $iWidth = _GDIPlus_ImageGetWidth($vBmp)
	Local $iHeight = _GDIPlus_ImageGetHeight($vBmp)
	$undo = _GDIPlus_BitmapCloneArea($vBmp, 0, 0, $iWidth, $iHeight)
	$undo_chk = True
	$vBitmap = _GDIPlus_BitmapCloneArea($vBmp, 0, 0, $iWidth, $iHeight)
	Local $hBitmapData = _GDIPlus_BitmapLockBits($vBitmap, 0, 0, $iWidth, $iHeight, BitOR($GDIP_ILMREAD, $GDIP_ILMWRITE), $GDIP_PXF32RGB)
	Local $Scan = DllStructGetData($hBitmapData, "Scan0")
	Local $Stride = DllStructGetData($hBitmapData, "Stride")
	Local $tPixelData = DllStructCreate("dword[" & (Abs($Stride * $iHeight)) & "]", $Scan)
	Local $bytecode = "0x8B7C24048B5424088B5C240CB900000000C1E202575352518B040FBA00000000BB00000000B90000000088C2C1E80888C3C1E80888C18B44240883F800772FB85555000001CB01D3F7E3C1E810BB00000000B3FFC1E30888C3C1E30888C3C1E30888C389D8595A5B5F89040FEB3B89C839C3720289D839C2720289D05089F839C3770289D839C2770289D05B01D8BBDC780000F7E3C1E810595A5B5F3B4424107213C7040FFFFFFF0083C10439D1730EE95FFFFFFFC7040F00000000EBEBC3"
	Local $tCodebuffer = DllStructCreate("byte[" & BinaryLen($bytecode) & "]")
	Local $Ret = DllStructSetData($tCodebuffer, 1, $bytecode)
	DllCall("kernel32.dll", "int", "VirtualProtect", "ptr", DllStructGetPtr($tCodebuffer), "ulong", BinaryLen($bytecode), "dword", $PAGE_EXECUTE_READWRITE, "dword*", 0) ;avoid crash when DEP is activates for all programs and services! Thanks to progandy
	DllCall("user32.dll", "int", "CallWindowProcW", "ptr", DllStructGetPtr($tCodebuffer), "ptr", DllStructGetPtr($tPixelData), "int", $iWidth * $iHeight, "int", $iBlackAndWhite, "int", $iLight);
	_GDIPlus_BitmapUnlockBits($vBitmap, $hBitmapData)
	$hBmp = $vBitmap
	$hClipboard_Bitmap = _WinAPI_CopyImage(_GDIPlus_BitmapCreateHBITMAPFromBitmap($hBmp), 0, 0, 0, $LR_COPYDELETEORG + $LR_COPYRETURNORG)
	Draw2Graphic($hBmp)
	$tPixelData = 0
	$tCodebuffer = 0
EndFunc   ;==>ASM_Bitmap_Grey_BnW

Func ASM_Bitmap_Invert($vBmp) ;ASM code by AndyG
	Local $iWidth = _GDIPlus_ImageGetWidth($vBmp)
	Local $iHeight = _GDIPlus_ImageGetHeight($vBmp)
	$undo = _GDIPlus_BitmapCloneArea($vBmp, 0, 0, $iWidth, $iHeight)
	$undo_chk = True
	$vBitmap = _GDIPlus_BitmapCloneArea($vBmp, 0, 0, $iWidth, $iHeight)
	Local $hBitmapData = _GDIPlus_BitmapLockBits($vBitmap, 0, 0, $iWidth, $iHeight, BitOR($GDIP_ILMREAD, $GDIP_ILMWRITE), $GDIP_PXF32RGB)
	Local $Scan = DllStructGetData($hBitmapData, "Scan0")
	Local $Stride = DllStructGetData($hBitmapData, "Stride")
	Local $tCodebuffer = DllStructCreate("byte[23]")
	DllStructSetData($tCodebuffer, 1, "0x8B7424048B4C24088136FFFFFF0083C60483E90177F2C3")
	DllCall("kernel32.dll", "int", "VirtualProtect", "ptr", DllStructGetPtr($tCodebuffer), "ulong", 23, "dword", $PAGE_EXECUTE_READWRITE, "dword*", 0) ;avoid crash when DEP is activates for all programs and services! Thanks to progandy
	DllCall("user32.dll", "ptr", "CallWindowProcW", "ptr", DllStructGetPtr($tCodebuffer), "ptr", $Scan, "int", $iWidth * $iHeight, "int", 0, "int", 0)
	_GDIPlus_BitmapUnlockBits($vBitmap, $hBitmapData)
	$hBmp = $vBitmap
	$hClipboard_Bitmap = _WinAPI_CopyImage(_GDIPlus_BitmapCreateHBITMAPFromBitmap($hBmp), 0, 0, 0, $LR_COPYDELETEORG + $LR_COPYRETURNORG)
	Draw2Graphic($hBmp)
	$tCodebuffer = 0
EndFunc   ;==>ASM_Bitmap_Invert
#endregion Image Editing
#region Copy Image  from _WinAPIEx.au3 by Yashied
Func _WinAPI_CopyImage($hImage, $iType = 0, $xDesired = 0, $yDesired = 0, $iFlags = 0); from _WinAPIEx by Yashied
	Local $Ret = DllCall('user32.dll', 'ptr', 'CopyImage', 'ptr', $hImage, 'int', $iType, 'int', $xDesired, 'int', $yDesired, 'int', $iFlags)
	If (@error) Or (Not $Ret[0]) Then Return SetError(1, 0, 0)
	Return $Ret[0]
EndFunc   ;==>_WinAPI_CopyImage
#endregion Copy Image  from _WinAPIEx.au3 by Yashied
#region Additional GDI+ Functions
Func _GDIPlus_GraphicsScaleTransform($hGraphics, $nScaleX, $nScaleY, $iOrder = 0)
	Local $aResult = DllCall($ghGDIPDll, "uint", "GdipScaleWorldTransform", "hwnd", $hGraphics, "float", $nScaleX, "float", $nScaleY, "int", $iOrder)
	If @error Then Return SetError(@error, @extended, False)
	Return $aResult[0] = 0
EndFunc   ;==>_GDIPlus_GraphicsScaleTransform

Func _GDIPlus_GraphicsSetInterpolationMode($hGraphics, $iInterpolationMode)
	Local $aResult = DllCall($ghGDIPDll, "uint", "GdipSetInterpolationMode", "hwnd", $hGraphics, "int", $iInterpolationMode)
	If @error Then Return SetError(@error, @extended, False)
	Return $aResult[0] = 0
EndFunc   ;==>_GDIPlus_GraphicsSetInterpolationMode

Func _GDIPlus_LineBrushCreateFromRectWithAngle($tRectF, $iARGBClr1, $iARGBClr2, $nAngle, $fIsAngleScalable = True, $iWrapMode = 0)
	Local $pRectF, $aResult
	$pRectF = DllStructGetPtr($tRectF)
	$aResult = DllCall($ghGDIPDll, "uint", "GdipCreateLineBrushFromRectWithAngle", "ptr", $pRectF, "uint", $iARGBClr1, "uint", $iARGBClr2, "float", $nAngle, "int", $fIsAngleScalable, "int", $iWrapMode, "int*", 0)
	If @error Then Return SetError(@error, @extended, 0)
	Return $aResult[7]
EndFunc   ;==>_GDIPlus_LineBrushCreateFromRectWithAngle

Func _GDIPlus_BitmapCreateFromScan0($iWidth, $iHeight, $iStride = 0, $iPixelFormat = 0x0026200A, $pScan0 = 0)
	Local $aResult = DllCall($ghGDIPDll, "uint", "GdipCreateBitmapFromScan0", "int", $iWidth, "int", $iHeight, "int", $iStride, "int", $iPixelFormat, "ptr", $pScan0, "int*", 0)
	If @error Then Return SetError(@error, @extended, 0)
	Return $aResult[6]
EndFunc   ;==>_GDIPlus_BitmapCreateFromScan0

Func _GDIPlus_CreateLineBrush($iPoint1X, $iPoint1Y, $iPoint2X, $iPoint2Y, $iArgb1 = 0xFF0000FF, $iArgb2 = 0xFFFF0000, $WrapMode = 0)
	Local $tPoint1, $pPoint1, $tPoint2, $pPoint2, $aRet
	If $iArgb1 = "" Then $iArgb1 = 0xFF0000FF
	If $iArgb2 = "" Then $iArgb2 = 0xFFFF0000
	If $WrapMode = -1 Then $WrapMode = 0
	$tPoint1 = DllStructCreate("float X;float Y")
	$pPoint1 = DllStructGetPtr($tPoint1)
	DllStructSetData($tPoint1, "X", $iPoint1X)
	DllStructSetData($tPoint1, "Y", $iPoint1Y)
	$tPoint2 = DllStructCreate("float X;float Y")
	$pPoint2 = DllStructGetPtr($tPoint2)
	DllStructSetData($tPoint2, "X", $iPoint2X)
	DllStructSetData($tPoint2, "Y", $iPoint2Y)
	$aRet = DllCall($ghGDIPDll, "int", "GdipCreateLineBrush", "ptr", $pPoint1, "ptr", $pPoint2, "int", $iArgb1, "int", $iArgb2, "int", $WrapMode, "int*", 0)
	Return $aRet[6]
EndFunc   ;==>_GDIPlus_CreateLineBrush
#endregion Additional GDI+ Functions

#region GUI Fadeout from _WinAPIEx.au3 by Yashied
Func _WinAPI_AnimateWindow($hWnd, $iFlags, $iDuration = 1000)
	Local $Ret = DllCall('user32.dll', 'int', 'AnimateWindow', 'hwnd', $hWnd, 'dword', $iDuration, 'dword', $iFlags)
	If (@error) Or (Not $Ret[0]) Then Return SetError(1, 0, 0)
	Return 1
EndFunc   ;==>_WinAPI_AnimateWindow
#endregion GUI Fadeout from _WinAPIEx.au3 by Yashied
#region 3D Text by taietel
Func _3DText($sText, $iX, $iY, $iW = -1, $iH = -1, $iFontSize = 40, $sWeight = "b", $sStyle = "n", $sFont = "Comic Sans Ms", $iFontColor = 0x28070A, $iFontQuality = 4) ;code by taietel
	Local $iWeight, $iStyle
	If $iW = -1 Or $iW = Default Then $iW = Int(StringLen($sText) * $iFontSize / 1.2)
	If $iH = -1 Or $iH = Default Then $iH = Int(1.75 * $iFontSize)
	If $iFontSize = -1 Or $iFontSize = Default Then $iFontSize = 14
	If $sWeight = -1 Or $sWeight = Default Then $sWeight = "b"
	If $sStyle = -1 Or $sStyle = Default Then $sStyle = "n"
	If $sFont = -1 Or $sFont = Default Then $sFont = "Arial"
;~ 	Local $f = 0x081A12
	Switch $sWeight
		Case "b"
			$iWeight = 800
		Case "n"
			$iWeight = 400
	EndSwitch
	Switch $sStyle
		Case "n"
			$iStyle = 0
		Case "i"
			$iStyle = 2
		Case "u"
			$iStyle = 4
	EndSwitch
	Local $iZ = $iFontSize / 8
	Local $s = 1
	For $i = 0 To $iZ Step $s
		GUICtrlCreateLabel($sText, $iX - $i, $iY + $i, $iW, $iH)
		GUICtrlSetColor(-1, $iFontColor * ($i + 1) / $s)
		GUICtrlSetFont(-1, $iFontSize, $iWeight, $iStyle, $sFont, $iFontQuality)
		GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
	Next
EndFunc   ;==>_3DText
#endregion 3D Text by taietel

#region List all window handles by Authenticity
Func GetAllWindow() ;code by Authenticity - modified by UEZ
	Local $aWin = WinList(), $aWindows[1][4]
	Local $iStyle, $iEx_Style, $iCounter = 0
	Local $i, $hWnd_state, $aWinPos

	For $i = 1 To $aWin[0][0]
		$iEx_Style = BitAND(_WinAPI_GetWindowLong($aWin[$i][1], $GWL_EXSTYLE), $WS_EX_TOOLWINDOW)
		$iStyle = BitAND(WinGetState($aWin[$i][1]), 2)
		If $iEx_Style <> -1 And Not $iEx_Style And $iStyle Then
			$aWinPos = WinGetPos($aWin[$i][1])
			If $aWinPos[2] > 1 And $aWinPos[3] > 1 Then
				$aWindows[$iCounter][0] = $aWin[$i][0]
				$aWindows[$iCounter][1] = $aWin[$i][1]
;~ 				$hWnd_state = WinGetState($aWin[$i][1])
;~ 				ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : Abs($aWinPos[0] = ' & Abs($aFullScreen[2]) & @crlf & '>Error code: ' & @error & @crlf) ;### Debug Console
;~ 				If BitAND($hWnd_state, 32) = 32 And (Abs($aWinPos[0]) < Abs($aFullScreen[2]) And Abs($aWinPos[1]) < Abs($aFullScreen[3])) Then
;~ 					$aWindows[$iCounter][2] = $aWinPos[2] - 2 * Abs($aWinPos[0])
;~ 					$aWindows[$iCounter][3] = $aWinPos[3] - 2 * Abs($aWinPos[1])
;~ 				Else
				$aWindows[$iCounter][2] = $aWinPos[2]
				$aWindows[$iCounter][3] = $aWinPos[3]
;~ 				EndIf
				$iCounter += 1
			EndIf
			ReDim $aWindows[$iCounter + 1][4]
		EndIf
	Next
	ReDim $aWindows[$iCounter][4]
	Return $aWindows
EndFunc   ;==>GetAllWindow
#endregion List all window handles by Authenticity
#region _Mouse_Proc by _Kurt
;http://www.autoitscript.com/forum/index.php?showtopic=81761
Func _Mouse_Proc($nCode, $wParam, $lParam) ;function called for mouse events.. Made by _Kurt
	;define local vars
	Local $info, $mouseData, $Ret, $mi
	If $nCode < 0 Then ;recommended, see http://msdn.microsoft.com/en-us/library/ms644986(VS.85).aspx
		$Ret = DllCall("user32.dll", "long", "CallNextHookEx", "hwnd", $hM_Hook[0], "int", $nCode, "ptr", $wParam, "ptr", $lParam) ;recommended
		Return $Ret[0]
	EndIf
	$info = DllStructCreate($MSLLHOOKSTRUCT, $lParam)
	$mouseData = DllStructGetData($info, 3)
	;Find which event happened
	Select
		Case $wParam = $WM_MOUSEWHEEL And WinActive($hGUI)
			$mi = GUIGetCursorInfo()
			If Not @error Then
				If $mi[4] <> $List Then ;activate zooming only when mouse is not hovering listview section
					If _WinAPI_HiWord($mouseData) > 0 Then
						Zoom(1)
					Else
						Zoom(0)
					EndIf
					Return 1
				EndIf
			EndIf
	EndSelect

	;This is recommended instead of Return 0
	$Ret = DllCall("user32.dll", "long", "CallNextHookEx", "hwnd", $hM_Hook[0], "int", $nCode, "ptr", $wParam, "ptr", $lParam)
	Return $Ret[0]
EndFunc   ;==>_Mouse_Proc
#endregion _Mouse_Proc by _Kurt
#region Extract icons from a DLL and display it in a gui control menu by Yashied
Func _GUICtrlMenu_CreateBitmap_WinAPI($file, $iIndex = 0, $iX = 16, $iY = 16) ;thanks to Yashied
	If FileExists($file) Then
		Local $aRet, $hIcon, $hBitmap
		Local $hDC, $hBackDC, $hBackSv

		$aRet = DllCall("shell32", "long", "ExtractAssociatedIcon", "int", 0, "str", $file, "word*", $iIndex)
		If @error Then Return SetError(@error, @extended, 0)
		$hIcon = $aRet[0]

		$hDC = _WinAPI_GetDC(0) ;thanks to Yashied
		$hBackDC = _WinAPI_CreateCompatibleDC($hDC)
		$hBitmap = _WinAPI_CreateSolidBitmap(0, _WinAPI_GetSysColor($COLOR_MENU), $iX, $iY)
		$hBackSv = _WinAPI_SelectObject($hBackDC, $hBitmap)
		_WinAPI_DrawIconEx($hBackDC, 0, 0, $hIcon, $iX, $iY, 0, 0, 3)
		_WinAPI_DestroyIcon($hIcon)

		_WinAPI_SelectObject($hBackDC, $hBackSv)
		_WinAPI_ReleaseDC(0, $hDC)
		_WinAPI_DeleteDC($hBackDC)
		Return $hBitmap
	Else
		Return SetError(1, 0, 0)
	EndIf
EndFunc   ;==>_GUICtrlMenu_CreateBitmap_WinAPI
#endregion Extract icons from a DLL and display it in a gui control menu by Yashied

#region AVIWriter by monoceres, progandy
;Adds a bitmap file to an already opened avi file.
;monoceres, Prog@ndy
Func _AddHBitmapToAvi(ByRef $Avi_Handle, $hBitmap)
	Local $DC = _WinAPI_GetDC(0)
	Local $hDC = _WinAPI_CreateCompatibleDC($DC)
	_WinAPI_ReleaseDC(0, $DC)

	Local $OldBMP = _WinAPI_SelectObject($hDC, $hBitmap)
	Local $bits = DllStructCreate("byte[" & DllStructGetData($Avi_Handle[3], "biSizeImage") & "]")
	_WinAPI_GetDIBits($hDC, $hBitmap, 0, Abs(DllStructGetData($Avi_Handle[3], "biHeight")), DllStructGetPtr($bits), DllStructGetPtr($Avi_Handle[3]), 0)
	_WinAPI_SelectObject($hDC, $OldBMP)
	_WinAPI_DeleteDC($hDC)

	DllCall($Avi32_Dll, "int", "AVIStreamWrite", "ptr", $Avi_Handle[1], "long", $Avi_Handle[2], "long", 1, "ptr", DllStructGetPtr($bits), _
			"long", DllStructGetSize($bits), "long", $AVIIF_KEYFRAME, "ptr*", 0, "ptr*", 0)
	$Avi_Handle[2] += 1
EndFunc   ;==>_AddHBitmapToAvi

; Init the avi library
Func _StartAviLibrary()
	$Avi32_Dll = DllOpen("Avifil32.dll")
	DllCall($Avi32_Dll, "none", "AVIFileInit")
EndFunc   ;==>_StartAviLibrary

; Release the library
Func _StopAviLibrary()
	DllCall($Avi32_Dll, "none", "AVIFileExit")
	DllClose($Avi32_Dll)
EndFunc   ;==>_StopAviLibrary

Func _CloseAvi($Avi_Handle)
	DllCall($Avi32_Dll, "int", "AVIStreamRelease", "ptr", $Avi_Handle[1])
	DllCall($Avi32_Dll, "int", "AVIStreamRelease", "ptr", $Avi_Handle[5])
	DllCall($Avi32_Dll, "int", "AVIFileRelease", "ptr", $Avi_Handle[0])
EndFunc   ;==>_CloseAvi

;monoceres, Prog@ndy, UEZ
Func _CreateAvi($sFileName, $FrameRate, $width, $height, $BitCount = 24, $mmioFOURCC = "DIB")
	Local $RetArr[6] ;avi file handle, compressed stream handle, bitmap count, bitmap info header, stride, stream handle

	Local $aRet, $pFile, $asi, $amh, $aco, $pStream, $psCompressed

	Local $Stride = BitAND(($width * ($BitCount / 8) + 3), BitNOT(3))

	Local $bi = DllStructCreate($BITMAPINFOHEADER)
	DllStructSetData($bi, "biSize", DllStructGetSize($bi))
	DllStructSetData($bi, "biWidth", $width)
	DllStructSetData($bi, "biHeight", $height)
	DllStructSetData($bi, "biPlanes", 1)
	DllStructSetData($bi, "biBitCount", $BitCount)
	DllStructSetData($bi, "biSizeImage", $Stride * $height)

	$asi = DllStructCreate($AVISTREAMINFO)
	DllStructSetData($asi, "fccType", _Create_mmioFOURCC("vids"))
	DllStructSetData($asi, "fccHandler", _Create_mmioFOURCC($mmioFOURCC))
	DllStructSetData($asi, "dwRate", $FrameRate)
	DllStructSetData($asi, "dwScale", 1)
	DllStructSetData($asi, "dwQuality", -1) ;Quality is represented as a number between 0 and 10,000. For compressed data, this typically represents the value of the quality parameter passed to the compression software. If set to &#8211;1, drivers use the default quality value.
	DllStructSetData($asi, "dwSuggestedBufferSize", $Stride * $height)
	DllStructSetData($asi, "rright", $width)
	DllStructSetData($asi, "rbottom", $height)

;~     $amh = DllStructCreate($AVIMAINHEADER)
;~     DllStructSetData($asi, "fccType", _Create_mmioFOURCC("vids"))
;~     DllStructSetData($asi, "dwMicroSecPerFrame", Int(1000000 / $FrameRate))
;~     DllStructSetData($asi, "dwMaxBytesPerSec", 3728000)
;~     DllStructSetData($asi, "dwTotalFrames", $rec_time * $FrameRate)
;~     DllStructSetData($asi, "dwStreams", 1)
;~     DllStructSetData($asi, "dwSuggestedBufferSize", $stride * $Height)
;~     DllStructSetData($asi, "dwWidth", $Width)
;~     DllStructSetData($asi, "dwHeight", $Height)
;~     DllStructSetData($asi, "dwReserved", 0)


	$aco = DllStructCreate($AVICOMPRESSOPTIONS)
	DllStructSetData($aco, "fccType", _Create_mmioFOURCC("vids"))
	DllStructSetData($aco, "fccHandler", _Create_mmioFOURCC($mmioFOURCC))
;~     DllStructSetData($aco, "dwKeyFrameEvery", 10)

	$aRet = DllCall($Avi32_Dll, "int", "AVIFileOpenW", "ptr*", 0, "wstr", $sFileName, "uint", $OF_CREATE, "ptr", 0)
	$pFile = $aRet[1]

	$aRet = DllCall($Avi32_Dll, "int", "AVIFileCreateStream", "ptr", $pFile, "ptr*", 0, "ptr", DllStructGetPtr($asi))
	$pStream = $aRet[2]

	If $compress_avi Then
		AdlibRegister("AviHwndOnTop")
		$aRet = DllCall($Avi32_Dll, "int_ptr", "AVISaveOptions", "hwnd", 0, "uint", BitOR($ICMF_CHOOSE_DATARATE, $ICMF_CHOOSE_KEYFRAME), "int", 1, "ptr*", $pStream, "ptr*", DllStructGetPtr($aco))
		If $aRet[0] <> 1 Then
			$RetArr[0] = $pFile
			$RetArr[1] = $pStream
			$RetArr[2] = 0
			$RetArr[3] = $bi
			$RetArr[4] = $Stride
			$RetArr[5] = $pStream
			Return SetError(1, 0, $RetArr)
		EndIf
	EndIf

	;http://msdn.microsoft.com/en-us/library/dd756811(v=VS.85).aspx
	$aRet = DllCall($Avi32_Dll, "int", "AVIMakeCompressedStream", "ptr*", 0, "ptr", $pStream, "ptr", DllStructGetPtr($aco), "ptr", 0)
	If $aRet[0] <> $AVIERR_OK Then
		$RetArr[0] = $pFile
		$RetArr[1] = $pStream
		$RetArr[2] = 0
		$RetArr[3] = $bi
		$RetArr[4] = $Stride
		$RetArr[5] = $pStream
		Return SetError(2, 0, $RetArr)
	EndIf
	$psCompressed = $aRet[1]

	;The format for the stream is the same as BITMAPINFOHEADER
	$aRet = DllCall($Avi32_Dll, "int", "AVIStreamSetFormat", "ptr", $psCompressed, "long", 0, "ptr", DllStructGetPtr($bi), "long", DllStructGetSize($bi))
	$RetArr[0] = $pFile
	$RetArr[1] = $psCompressed
	$RetArr[2] = 0
	$RetArr[3] = $bi
	$RetArr[4] = $Stride
	$RetArr[5] = $pStream
	Return $RetArr
EndFunc   ;==>_CreateAvi

Func AviHwndOnTop()
	Local $hWnd = WinGetHandle("[CLASS:#32770]")
	If Not @error Then
		If _WinAPI_GetWindowLong($hWnd, $GWL_STYLE) = 0x94C800C4 Then WinSetOnTop($hWnd, "", 1)
	EndIf
	AdlibUnRegister("AviHwndOnTop")
EndFunc   ;==>AviHwndOnTop

;http://www.fourcc.org/codecs.php
Func _Create_mmioFOURCC($FOURCC) ;coded by UEZ
	If StringLen($FOURCC) <> 4 Then Return SetError(1, 0, 0)
	Local $aFOURCC = StringSplit($FOURCC, "", 2)
	Return BitOR(Asc($aFOURCC[0]), BitShift(Asc($aFOURCC[1]), -8), BitShift(Asc($aFOURCC[2]), -16), BitShift(Asc($aFOURCC[3]), -24))
EndFunc   ;==>_Create_mmioFOURCC

Func _DecodeFOURCC($iFOURCC) ;coded by UEZ
	If Not IsInt($iFOURCC) Then Return SetError(1, 0, 0)
	Return Chr(BitAND($iFOURCC, 0xFF)) & Chr(BitShift(BitAND(0x0000FF00, $iFOURCC), 8)) & Chr(BitShift(BitAND(0x00FF0000, $iFOURCC), 16)) & Chr(BitShift($iFOURCC, 24))
EndFunc   ;==>_DecodeFOURCC
#endregion AVIWriter by monoceres, progandy
#endregion external functions
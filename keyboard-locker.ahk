#Persistent
#SingleInstance Ignore

#Include Ini.ahk
#Include Settings.ahk
#Include Lib\AutoHotInterception.ahk

FileInstall, unlocked.ico, unlocked.ico, 0
FileInstall, locked.ico, locked.ico, 0
FileInstall, install-interception.exe, install-interception.exe, 0

settings := new Settings()

;(DO NOT CHANGE) tracks whether or not the keyboard is currently locked
locked := false
mouseLocked := false
lockMode := ""
laptopKeyboardDeviceIds := []
ahi := ""
laptopKeyboardInterceptionId := 0
laptopKeyboardInterceptionHandle := ""

;ensure built-in keyboard devices are re-enabled if script exits while laptop-only lock is active
OnExit("CleanupOnExit")

;create the tray icon and do initial setup
initialize()

;set up the keyboard shortcut to lock the keyboard
Hotkey, % settings.Shortcut(), ShortcutTriggered

;end execution here - the rest of the file is functions and callbacks
return

initialize()
{
	;initialize the tray icon and menu
	Menu, Tray, Icon, %A_ScriptDir%\unlocked.ico
	Menu, Tray, NoStandard
	Menu, Tray, Tip, % "Press " . settings.ShortcutHint() . " to lock your keyboard"
	Menu, Tray, Add, Lock keyboard, ToggleKeyboard
	Menu, Tray, Add, Lock laptop keyboard only, LockLaptopKeyboardOnly
	Menu, Tray, Add, Install Interception driver, InstallInterceptionDriver
	Menu, Tray, Add, Show keyboard device IDs, ShowKeyboardDeviceIds
	Menu, Tray, Add, Open settings.ini, OpenSettingsFile
	Menu, Tray, Add, Open app folder, OpenAppFolder
	if (settings.HideTooltips()) {
		Menu, Tray, add, Show tray notifications, ToggleTray
	} else {
		Menu, Tray, add, Hide tray notifications, ToggleTray
	}
	Menu, Tray, Add, Exit, Exit

	if (settings.LockOnOpen()) {
		LockKeyboard(true)
	} else if (!settings.HideTooltips()) {
		TrayTip,,% "Press " . settings.ShortcutHint() . " to lock your keyboard",10,1
	}
}

;callback for when the keyboard shortcut is pressed
ShortcutTriggered:
    ;check if shortcut is disabled in settings
    if (settings.DisableShortcut())
    {
        return
    }

    ;if we're already locked, stop here
    if (locked)
    {
        return
    }

	;wait for each shortcut key to be released, so they don't get "stuck"
	for index, key in StrSplit(settings.ShortcutHint(), "+")
	{
		KeyWait, %key%
    }

	LockKeyboard(true)
return


;"Lock/Unlock keyboard" menu clicked
ToggleKeyboard()
{
	global locked
	global lockMode

	if (locked) {
		if (lockMode == "laptop-only" || lockMode == "laptop-only-ahi") {
			UnlockLaptopKeyboardOnly()
		} else {
			LockKeyboard(false)
		}
	} else {
		LockKeyboard(true)
	}
}

;"Lock laptop keyboard only" menu clicked
LockLaptopKeyboardOnly()
{
	global locked
	global lockMode
	global laptopKeyboardDeviceIds
	global laptopKeyboardInterceptionId
	global laptopKeyboardInterceptionHandle

	if (locked) {
		return
	}

	interceptionFailure := ""
	if (TryLockLaptopKeyboardWithInterception(interceptionFailure)) {
		return
	}

	if (!A_IsAdmin) {
		message := "Interception mode was not available, and fallback mode requires administrator privileges to disable built-in keyboard devices." 
		if (interceptionFailure != "") {
			message .= "`n`nInterception details:`n" . interceptionFailure
		}
		message .= "`n`nRun as administrator or install the Interception driver from tray menu."
		MsgBox, 48, Lock failed, % message
		return
	}

	deviceIds := GetLaptopKeyboardDeviceIds()
	if (deviceIds.Length() == 0) {
		summary := GetDetectedKeyboardSummary()
		configuredValue := settings.LaptopKeyboardDeviceIds()
		message := "Could not auto-detect your built-in keyboard.`n`nSet ""laptop-keyboard-device-ids"" in:`n" . GetSettingsPath() . "`nUse ""|"" to separate multiple IDs.`n`nCurrent value:`n" . configuredValue . "`n`nDetected keyboard devices:`n" . summary
		MsgBox, 48, Laptop keyboard not found, % message
		return
	}

	disabledIds := []
	failureDetails := ""
	for index, deviceId in deviceIds {
		failure := ""
		if (SetKeyboardDeviceState(deviceId, false, failure)) {
			disabledIds.Push(deviceId)
		} else {
			if (failureDetails != "") {
				failureDetails .= "`n`n"
			}
			failureDetails .= deviceId
			if (failure != "") {
				failureDetails .= "`n" . failure
			}
		}
	}

	if (disabledIds.Length() == 0) {
		summary := GetDetectedKeyboardSummary()
		message := "Could not disable the built-in keyboard device(s).`n`nIf this machine uses a non-disableable built-in keyboard driver, set laptop-keyboard-device-ids manually in settings.ini.`n`nDetected keyboard devices:`n" . summary
		if (interceptionFailure != "") {
			message .= "`n`nInterception details:`n" . interceptionFailure
		}
		if (failureDetails != "") {
			message .= "`n`nAttempt details:`n" . failureDetails
		}
		if (InStr(failureDetails, "Cannot disable critical system device")) {
			message .= "`n`nThis laptop keyboard is marked as a critical system device by Windows. Install Interception driver from tray menu to use device-level blocking instead of device disable."
		}
		MsgBox, 48, Lock failed, % message
		return
	}

	laptopKeyboardDeviceIds := disabledIds
	laptopKeyboardInterceptionId := 0
	laptopKeyboardInterceptionHandle := ""
	locked := true
	lockMode := "laptop-only"

	Menu, Tray, Icon, %A_ScriptDir%\locked.ico
	Menu, Tray, Tip, Laptop keyboard is locked. Use tray menu to unlock.
	Menu, Tray, Rename, Lock keyboard, Unlock keyboard
	Menu, Tray, Disable, Lock laptop keyboard only

	if (!settings.HideTooltips()) {
		TrayTip,, Laptop keyboard is now locked.`nUSB/Bluetooth keyboards stay active.`nUse tray menu "Unlock keyboard" to unlock.,10,1
	}
}

TryLockLaptopKeyboardWithInterception(ByRef failureDetail := "")
{
	global ahi
	global locked
	global lockMode
	global laptopKeyboardInterceptionId
	global laptopKeyboardInterceptionHandle

	failureDetail := ""

	driverNeedsReboot := false
	if (!IsInterceptionDriverInstalled(driverNeedsReboot)) {
		failureDetail := "Interception driver is not installed."
		return false
	}

	if (driverNeedsReboot) {
		failureDetail := "Interception driver is installed but not active yet. Please reboot Windows first."
		return false
	}

	if (!EnsureAhiLibraryFilesPresent()) {
		failureDetail := "AutoHotInterception library files are missing from the Lib folder."
		return false
	}

	if (!IsObject(ahi)) {
		ahi := new AutoHotInterception()
	}

	matchedHandle := ""
	idDetail := ""
	keyboardId := GetLaptopInterceptionKeyboardId(matchedHandle, idDetail)
	if (keyboardId <= 0) {
		failureDetail := "Could not map laptop keyboard to an Interception device id."
		if (idDetail != "") {
			failureDetail .= "`n" . idDetail
		}
		return false
	}

	try {
		ahi.SubscribeKeyboard(keyboardId, true, Func("LaptopKeyboardInterceptedKeyEvent"))
	} catch e {
		failureDetail := "Interception subscription failed: " . e.Message
		return false
	}

	laptopKeyboardInterceptionId := keyboardId
	laptopKeyboardInterceptionHandle := matchedHandle
	locked := true
	lockMode := "laptop-only-ahi"

	Menu, Tray, Icon, %A_ScriptDir%\locked.ico
	Menu, Tray, Tip, Laptop keyboard is locked (Interception mode). Use tray menu to unlock.
	Menu, Tray, Rename, Lock keyboard, Unlock keyboard
	Menu, Tray, Disable, Lock laptop keyboard only

	if (!settings.HideTooltips()) {
		TrayTip,, Laptop keyboard is now locked.`nUSB/Bluetooth keyboards stay active.`nUse tray menu "Unlock keyboard" to unlock.,10,1
	}

	return true
}

LaptopKeyboardInterceptedKeyEvent(code, state)
{
	return
}

InstallInterceptionDriver()
{
	installer := A_ScriptDir . "\install-interception.exe"
	if (!FileExist(installer)) {
		MsgBox, 48, Installer not found, Could not find install-interception.exe in script folder.
		return
	}

	if (!A_IsAdmin) {
		MsgBox, 48, Administrator required, Please run this app as administrator, then click this menu item again.
		return
	}

	command := ComSpec . " /c " . Chr(34) . installer . Chr(34) . " /install"
	result := ExecCommand(command)
	if (result.exitCode = 0 || InStr(toLower(result.output), "successfully installed")) {
		MsgBox, 64, Driver installed, Interception driver install completed.`nPlease reboot Windows before testing "Lock laptop keyboard only".
		return
	}

	message := "Interception driver install failed.`n`nExit code: " . result.exitCode
	if (result.output != "") {
		message .= "`n`nOutput:`n" . result.output
	}
	MsgBox, 48, Install failed, % message
}

EnsureAhiLibraryFilesPresent(showMessage := false)
{
	paths := [A_ScriptDir . "\Lib\AutoHotInterception.ahk", A_ScriptDir . "\Lib\CLR.ahk", A_ScriptDir . "\Lib\AutoHotInterception.dll", A_ScriptDir . "\Lib\x64\interception.dll", A_ScriptDir . "\Lib\x86\interception.dll"]
	missing := ""
	for index, path in paths {
		if (!FileExist(path)) {
			if (missing != "") {
				missing .= "`n"
			}
			missing .= path
		}
	}

	if (missing != "" && showMessage) {
		MsgBox, 48, AHI files missing, Missing required files:`n%missing%
	}

	return (missing == "")
}

IsInterceptionDriverInstalled(ByRef rebootRequired := false)
{
	return IsInterceptionDriverInstalledInternal(rebootRequired)
}

IsInterceptionDriverInstalledInternal(ByRef rebootRequired)
{
	rebootRequired := false

	result := ExecCommand(ComSpec . " /c sc query keyboard")
	if (result.exitCode = 0) {
		return true
	}

	hasKeyboardFile := FileExist(A_WinDir . "\System32\drivers\keyboard.sys")
	hasMouseFile := FileExist(A_WinDir . "\System32\drivers\mouse.sys")

	kFilters := ""
	mFilters := ""
	RegRead, kFilters, HKLM, SYSTEM\CurrentControlSet\Control\Class\{4D36E96B-E325-11CE-BFC1-08002BE10318}, UpperFilters
	RegRead, mFilters, HKLM, SYSTEM\CurrentControlSet\Control\Class\{4D36E96F-E325-11CE-BFC1-08002BE10318}, UpperFilters

	hasKeyboardFilter := InStr(toLower(kFilters), "keyboard")
	hasMouseFilter := InStr(toLower(mFilters), "mouse")

	if (hasKeyboardFile && hasMouseFile && hasKeyboardFilter && hasMouseFilter) {
		rebootRequired := true
		return true
	}

	return false
}

GetInterceptionDriverStatusText()
{
	rebootRequired := false
	if (!IsInterceptionDriverInstalledInternal(rebootRequired)) {
		return "No"
	}

	if (rebootRequired) {
		return "Installed (restart required)"
	}

	return "Yes"
}

GetLaptopInterceptionKeyboardId(ByRef matchedHandle := "", ByRef detail := "")
{
	global ahi

	matchedHandle := ""
	detail := ""

	if (!IsObject(ahi)) {
		detail := "AHI is not initialized."
		return 0
	}

	handles := GetLaptopKeyboardHandles()
	if (!IsObject(handles) || handles.Length() == 0) {
		detail := "No laptop keyboard handles were detected."
		return 0
	}

	for index, handle in handles {
		Loop, 5 {
			id := ahi.Instance.GetDeviceIdFromHandle(false, handle, A_Index)
			if (id >= 1 && id <= 10) {
				matchedHandle := handle
				return id
			}
		}
	}

	tried := ""
	for index, handle in handles {
		if (tried != "") {
			tried .= ", "
		}
		tried .= handle
	}

	detail := "Tried handles: " . tried . "`nInterception keyboard list:`n" . GetAhiKeyboardSummary()
	return 0
}

GetLaptopKeyboardHandles()
{
	configured := Trim(settings.LaptopKeyboardHandles())
	if (configured != "") {
		return ParseHandleList(configured)
	}

	handles := []
	for index, deviceId in GetLaptopKeyboardDeviceIds() {
		handle := DeviceIdToInterceptionHandle(deviceId)
		if (handle != "" && !inArray(handle, handles)) {
			handles.Push(handle)
		}
	}

	return handles
}

ParseHandleList(handleText)
{
	handles := []
	Loop, Parse, handleText, |
	{
		handle := Trim(A_LoopField)
		if (handle != "" && !inArray(handle, handles)) {
			handles.Push(handle)
		}
	}
	return handles
}

DeviceIdToInterceptionHandle(deviceId)
{
	parts := StrSplit(deviceId, "\")
	if (!IsObject(parts) || parts.Length() < 2) {
		return ""
	}

	return parts[1] . "\" . parts[2]
}

GetAhiKeyboardSummary()
{
	global ahi

	if (!IsObject(ahi)) {
		return "(AHI not initialized)"
	}

	summary := ""
	deviceList := ahi.GetDeviceList()
	for id, device in deviceList {
		if (device.IsMouse) {
			continue
		}

		line := "ID " . id . " -> " . device.Handle
		if (summary != "") {
			summary .= "`n"
		}
		summary .= line
	}

	if (summary == "") {
		return "(none found)"
	}

	return summary
}

ShowKeyboardDeviceIds()
{
	summary := GetDetectedKeyboardSummary()
	handles := GetLaptopKeyboardHandles()
	handleSummary := ""
	for index, handle in handles {
		if (handleSummary != "") {
			handleSummary .= "`n"
		}
		handleSummary .= handle
	}

	message := "Settings file: " . GetSettingsPath() . "`nConfigured laptop-keyboard-device-ids: " . settings.LaptopKeyboardDeviceIds() . "`nConfigured laptop-keyboard-handles: " . settings.LaptopKeyboardHandles() . "`n`nInterception driver: " . GetInterceptionDriverStatusText() . "`n`nDetected keyboard devices:`n" . summary
	if (handleSummary != "") {
		message .= "`n`nLaptop handle candidates:`n" . handleSummary
	}

	MsgBox, 64, Keyboard device IDs, %message%
}

UnlockLaptopKeyboardOnly(showTip := true, shouldClose := true)
{
	global locked
	global lockMode
	global laptopKeyboardDeviceIds
	global ahi
	global laptopKeyboardInterceptionId
	global laptopKeyboardInterceptionHandle

	if (lockMode != "laptop-only" && lockMode != "laptop-only-ahi") {
		return
	}

	if (lockMode == "laptop-only-ahi") {
		if (IsObject(ahi) && laptopKeyboardInterceptionId > 0) {
			ahi.UnsubscribeKeyboard(laptopKeyboardInterceptionId)
		}
		laptopKeyboardInterceptionId := 0
		laptopKeyboardInterceptionHandle := ""
		laptopKeyboardDeviceIds := []
	} else {
		for index, deviceId in laptopKeyboardDeviceIds {
			SetKeyboardDeviceState(deviceId, true)
		}
		laptopKeyboardDeviceIds := []
	}

	locked := false
	lockMode := ""

	Menu, Tray, Icon, %A_ScriptDir%\unlocked.ico
	Menu, Tray, Tip, % "Press " . settings.ShortcutHint() . " to lock your keyboard"
	Menu, Tray, Rename, Unlock keyboard, Lock keyboard
	Menu, Tray, Enable, Lock laptop keyboard only

	if (showTip && !settings.HideTooltips()) {
		TrayTip,, Laptop keyboard is now unlocked.,10,1
	}

	if (shouldClose && settings.CloseOnUnlock())
	{
		ExitApp
	}
}

GetLaptopKeyboardDeviceIds()
{
	configured := Trim(settings.LaptopKeyboardDeviceIds())
	if (configured != "") {
		return ParseDeviceIdList(configured)
	}

	return AutoDetectLaptopKeyboardDeviceIds()
}

ParseDeviceIdList(deviceIdText)
{
	ids := []
	Loop, Parse, deviceIdText, |
	{
		id := Trim(A_LoopField)
		if (id != "") {
			ids.Push(id)
		}
	}
	return ids
}

AutoDetectLaptopKeyboardDeviceIds()
{
	ids := []
	query := ComObjGet("winmgmts:").ExecQuery("SELECT PNPDeviceID, Name, Description, Status FROM Win32_PnPEntity WHERE PNPClass='Keyboard'")

	for keyboard in query {
		deviceId := Trim(keyboard.PNPDeviceID . "")
		if (deviceId == "") {
			continue
		}

		deviceName := Trim((keyboard.Name . " " . keyboard.Description) . "")
		if (IsLikelyLaptopKeyboard(deviceId, deviceName) && !inArray(deviceId, ids)) {
			ids.Push(deviceId)
		}
	}

	return ids
}

IsLikelyLaptopKeyboard(deviceId, deviceName)
{
	if (RegExMatch(deviceId, "i)^(USB|BTH|BTHENUM)\\")) {
		return false
	}

	if (RegExMatch(deviceId, "i)^ACPI\\")) {
		return true
	}

	if (RegExMatch(deviceName, "i)PS/2|AT Translated")) {
		return true
	}

	return false
}

GetDetectedKeyboardSummary()
{
	summary := ""
	query := ComObjGet("winmgmts:").ExecQuery("SELECT PNPDeviceID, Name, Description, Status FROM Win32_PnPEntity WHERE PNPClass='Keyboard'")

	for keyboard in query {
		deviceId := Trim(keyboard.PNPDeviceID . "")
		deviceName := Trim((keyboard.Name . " " . keyboard.Description) . "")
		status := Trim(keyboard.Status . "")
		if (deviceId == "") {
			continue
		}

		if (summary != "") {
			summary .= "`n"
		}
		summary .= deviceName . " [" . status . "] -> " . deviceId
	}

	if (summary == "") {
		return "(none found)"
	}

	return summary
}

GetSettingsPath()
{
	return A_ScriptDir . "\\settings.ini"
}

OpenSettingsFile()
{
	settingsPath := GetSettingsPath()
	if (!FileExist(settingsPath)) {
		MsgBox, 48, File not found, Could not find settings.ini at:`n%settingsPath%
		return
	}

	notepadPlusPlus := A_ProgramFiles . "\Notepad++\notepad++.exe"
	if (FileExist(notepadPlusPlus)) {
		Run, "%notepadPlusPlus%" "%settingsPath%"
		return
	}

	notepadPlusPlus86 := A_ProgramFiles . " (x86)\Notepad++\notepad++.exe"
	if (FileExist(notepadPlusPlus86)) {
		Run, "%notepadPlusPlus86%" "%settingsPath%"
		return
	}

	Run, notepad.exe "%settingsPath%"
}

OpenAppFolder()
{
	Run, %A_ScriptDir%
}

FilterPresentKeyboardDeviceIds(deviceIds)
{
	if (!IsObject(deviceIds) || deviceIds.Length() == 0) {
		return []
	}

	presentMap := GetPresentKeyboardDeviceMap()
	presentIds := []
	for index, deviceId in deviceIds {
		if (presentMap.HasKey(toLower(deviceId))) {
			presentIds.Push(deviceId)
		}
	}

	return presentIds
}

GetPresentKeyboardDeviceMap()
{
	map := {}
	query := ComObjGet("winmgmts:").ExecQuery("SELECT PNPDeviceID, Status FROM Win32_PnPEntity WHERE PNPClass='Keyboard'")
	for keyboard in query {
		deviceId := Trim(keyboard.PNPDeviceID . "")
		status := Trim(keyboard.Status . "")
		if (deviceId == "") {
			continue
		}

		if (status == "OK") {
			map[toLower(deviceId)] := true
		}
	}

	return map
}

toLower(str)
{
	StringLower, lowered, str
	return lowered
}

SetKeyboardDeviceState(deviceId, enabled, ByRef failureDetail := "")
{
	failureDetail := ""
	action := enabled ? "/enable-device" : "/disable-device"
	command := ComSpec . " /c pnputil " . action . " " . Chr(34) . deviceId . Chr(34)
	result := ExecCommand(command)
	if (result.exitCode = 0 || result.exitCode = 3010) {
		return true
	}

	;fallback for systems where Disable/Enable-PnpDevice works better than pnputil
	psAction := enabled ? "Enable-PnpDevice" : "Disable-PnpDevice"
	escapedDeviceId := StrReplace(deviceId, "'", "''")
	psCommand := "powershell -NoProfile -ExecutionPolicy Bypass -Command " . Chr(34) . "$id='" . escapedDeviceId . "'; " . psAction . " -InstanceId $id -Confirm:$false -ErrorAction Stop | Out-Null" . Chr(34)
	psResult := ExecCommand(psCommand)
	if (psResult.exitCode = 0) {
		return true
	}

	failureDetail := "pnputil exit " . result.exitCode
	if (result.output != "") {
		failureDetail .= ": " . result.output
	}
	failureDetail .= "`nPowerShell exit " . psResult.exitCode
	if (psResult.output != "") {
		failureDetail .= ": " . psResult.output
	}

	return false
}

ExecCommand(command)
{
	shell := ComObjCreate("WScript.Shell")
	exec := shell.Exec(command)
	while (exec.Status = 0) {
		Sleep, 50
	}

	output := Trim(exec.StdOut.ReadAll())
	errorOutput := Trim(exec.StdErr.ReadAll())
	if (errorOutput != "") {
		if (output != "") {
			output .= "`n"
		}
		output .= errorOutput
	}

	return {exitCode: exec.ExitCode, output: output}
}

CleanupOnExit(exitReason, exitCode)
{
	global lockMode

	if (lockMode == "laptop-only" || lockMode == "laptop-only-ahi") {
		UnlockLaptopKeyboardOnly(false, false)
	}
}

;"Hide/Show tray notifications" menu clicked
ToggleTray()
{
	if (settings.HideTooltips()) {
	    settings.SetHideTooltips(false)
		Menu, Tray, Rename, Show tray notifications, Hide tray notifications
	} else {
	    settings.SetHideTooltips(true)
		Menu, Tray, Rename, Hide tray notifications, Show tray notifications
	}
}

;"Exit" menu clicked
Exit()
{
	ExitApp
}

;Lock or unlock the keyboard
LockKeyboard(lock)
{
	global locked
	global mouseLocked
	global lockMode

	;handle pointing to the keyboard hook
	static hHook = 0

	;lock status already matches what we were asked to do, no action necessary
	if ((hHook != 0) = lock) {
		return
	}
 
	if (lock) {
	    ;check that we didn't leave ourselves without a way to unlock again
	    if(settings.DisablePassword() && settings.LockMouse())
	    {
	        MsgBox, You have disabled password unlocking and enabled mouse locking, which will prevent you from unlocking your system. Please re-enable one or the other.
	        return
	    }

	    ;change the tray icon to a lock
		Menu, Tray, Icon, %A_ScriptDir%\locked.ico

        ;hint at the unlock password
		Menu, Tray, Tip, % "Type """ . settings.Password() . """ to unlock your keyboard"

        ;update menu to unlock
		Menu, Tray, Rename, Lock keyboard, Unlock keyboard

        ;lock the keyboard
		hHook := DllCall("SetWindowsHookEx", "Ptr", WH_KEYBOARD_LL:=13, "Ptr", RegisterCallback("Hook_Keyboard","Fast"), "Uint", DllCall("GetModuleHandle", "Uint", 0, "Ptr"), "Uint", 0, "Ptr")
		locked := true
		lockMode := "all"
		Menu, Tray, Disable, Lock laptop keyboard only

		;also lock the mouse, if configured to do so
		if (settings.LockMouse()) {
			Hotkey, LButton, doNothing
			Hotkey, RButton, doNothing
			Hotkey, MButton, doNothing
			BlockInput, MouseMove
			mouseLocked := true
		} else {
			mouseLocked := false
		}

        ;remind user what the password is
		if (!settings.HideTooltips()) {
			TrayTip,,% "Your keyboard is now locked.`nType """ . settings.Password() . """ to unlock it.",10,1
		}
	} else {
        ;unlock the keyboard
		DllCall("UnhookWindowsHookEx", "Ptr", hHook)
		hHook := 0
		locked := false
		lockMode := ""
		Menu, Tray, Enable, Lock laptop keyboard only

        ;also unlock the mouse, if configured to do so
		if (mouseLocked) {
            Hotkey, LButton, Off
            Hotkey, MButton, Off
            Hotkey, RButton, Off
            BlockInput, MouseMoveOff
			mouseLocked := false
        }

	    ;change tray icon back to unlocked
		Menu, Tray, Icon, %A_ScriptDir%\unlocked.ico

        ;hint at the keyboard shortcut to lock again
		Menu, Tray, Tip, % "Press " . settings.ShortcutHint() . " to lock your keyboard"

        ;update menu to lock
		Menu, Tray, Rename, Unlock keyboard, Lock keyboard

        ;remind user what the keyboard shortcut to lock is
		if (!settings.HideTooltips()) {
			TrayTip,,% "Your keyboard is now unlocked.`nPress " . settings.ShortcutHint() . " to lock it again.",10,1
		}

		if(settings.CloseOnUnlock())
		{
		    ExitApp
		}
	}
}

;Catch and discard keypresses when the keyboard is locked, and monitor for password inputs
Hook_Keyboard(nCode, wParam, lParam)
{
    ;track our position while correctly typing the password
	static count = 0

    ;is this a keyUp event (or keyDown)
    isKeyUp := NumGet(lParam+0, 8, "UInt") & 0x80

    ;get the scan code of the key pressed/released
    gotScanCode := NumGet(lParam+0, 4, "UInt")

    ;track the left/right shift keys, to handle capitals and symbols in passwords, because getkeystate calls don't work with our method of locking the keyboard
    ;if you can figure out how to use a getkeystate call to check for shift, or you have a better way to handle upper case letters and symbols, let me know
	static shifted = 0
    if(gotScanCode = 0x2A || gotScanCode = 0x36) {
        if(isKeyUp) {
            shifted := 0
        } else {
            shifted := 1
        }
        return 1
    }

	;check password progress/completion
	if (!settings.DisablePassword() && !isKeyUp) {
	    expectedCharacter := SubStr(settings.Password(), count+1, 1)
        expectedScanCode := GetKeySC(expectedCharacter)
        requiresShift := requiresShift(expectedCharacter)

        ;did they type the correct next password letter?
	    if(expectedScanCode == gotScanCode && requiresShift == shifted) {
	        count := count + 1

	        ;password is complete!
	        if(count == StrLen(settings.Password())) {
                count = 0
                shifted = 0
                LockKeyboard(false)
            }
	    } else {
			count = 0
        }
    }

	return 1
}

;Determine if this character requires shift to be pressed (capital letter or symbol)
requiresShift(chr)
{
    ;upper case characters always require shift
    if(isUpperCase(chr)) {
        return true
    }

    ;symbols that require shift
    static symbols = ["~", "!", "@", "#", "$", "%", "^", "&", "*", "(", ")", "_", "+", "{", "}", "|", ":", """", "<", ">", "?"]
    if(inArray(chr, symbols)) {
        return true
    }

    ;anything else is false
    return false
}

;Is the string (or character) upper case
isUpperCase(str)
{
    if str is upper
        return true
    else
        return false
}

;Is the string (or character) lower case
isLowerCase(str)
{
    if str is lower
        return true
    else
        return false
}

;Check if the haystack array contains the needle
inArray(needle, haystack) {
    ;only accept objects and arrays
	if(!IsObject(haystack) || haystack.Length() == 0) {
	    return false
	}

	for index, value in haystack {
		if (value == needle) {
		    return index
		}
    }
	return false
}

;this is used to block mouse input
doNothing:
return

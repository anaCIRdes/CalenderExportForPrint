#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         Gerald Boehm

 Script Function:
	Test Notepad

n
TODOs:
	- openFileDialog fuer Auswahl des zu verwendeten Ordners machen und Pfad dann speichern
	x alle Fotos eines Ordners auslesen und nacheinander abarbeiten bevor Affinity geschlossen wird
	- JPEG-Export noch sicherstellen dass das richtige ausgewählt wird, wenn doch einmal etwas anderes eingestell wurde
	- Affinity maximiert öffnen ( falls es doch einmal nicht so wäre )
	x Affinity zum Schluss schliesen ( abfangen wenn gefragt wird ob Änderungen in Vorlage gespeichert werden sollen -> nein auswählen )



#ce ----------------------------------------------------------------------------

; Script Start

#include <MsgBoxConstants.au3>
#include <AutoItConstants.au3>
#include <Array.au3>
#include <File.au3>



$path_to_allPhotos = "C:\Test"
$filename = "2024_00.jpg"
$path_to_photo1 = "C:\Test\2024_00.jpg"		; TODO: um hier mehr zu automatisieren muessten alle vorhandenen Files ausgelesen werden, gemerkt werden und alle nacheinander abgearbeitet werden
$path_to_photo1_new = "C:\Test\2024_00_3mm-sRGB.jpg"
$iCountOfFiles = 0


HotKeySet("{Esc}", "ExitScript")
Func ExitScript()
    Exit
EndFunc


Local $aFileList = _FileListToArray($path_to_allPhotos, "*", 1)				; Filer gesetzt (1) damit nur Files angezeigt werden
;_ArrayDisplay($aFileList, "$aFileList")		; erstes File im Array

;$sFileString = $aFileList[1]
;MsgBox($MB_SYSTEMMODAL, "", $sFileString)	; Ausgabe eines Eintrags vom Array

$iCountOfFiles = $aFileList[0]  			; Anzahl der gelisteten Files im Array in Variable speichern
;MsgBox($MB_SYSTEMMODAL, "", "Anzahl der Files: " & $iCountOfFiles)


For $i2 = 1 to $iCountOfFiles Step 1				; Dateien fangen bei 1 an
	; Affinity starten
	; Run("C:\Program Files\Affinity\Photo\Photo.exe")

	;A4-Format:
	;old nicht verwenden; ShellExecute("C:\Program Files\Affinity\Photo\Photo.exe", FileGetShortName("D:\Daten\Programmierung\Programmieren + Tutorials\AutoIT\Scripte\myScripts\AffinityCTRL_V1\A4_mit_3mm_Rand_Vorlage.psd"))
	;ShellExecute("C:\Program Files\Affinity\Photo\Photo.exe", FileGetShortName("C:\Test2\A4_mit_3mm_Rand_Vorlage.psd"))


	; A3-Format:
	ShellExecute("C:\Program Files\Affinity\Photo\Photo.exe", FileGetShortName("C:\Test2\A3_mit_3mm_Rand_Vorlage_PSD.psd"))

	;Panes
	$placeWind = "[CLASS:DirectUIHWND; INSTANCE:2]"



	; Warten, bis Affinity gestartet und aktiv ist
	WinWaitActive("Affinity")

	$affinity = WinActivate("Affinity")
	Sleep(3000)


	$filename = $aFileList[$i2]
	$path_to_photo = $path_to_allPhotos & "\" & $filename
	;MsgBox($MB_SYSTEMMODAL, "", "$path_to_photo: " & $path_to_photo)


	; neuen Ordner erstellen fuer neue Fotos
	$path_to_allPhotos_3mmSRGB = $path_to_allPhotos & "\3mmSRGB"

	;MsgBox($MB_SYSTEMMODAL, "", "Title: " & $path_to_allPhotos_3mmSRGB)

	Send("!D")		; fuer Datei-Menue
	Send("{DOWN}")	; Pfeiltaste nach unten um Menue zu oeffnen_00.jpg

	Send("z")		; fuer Platzieren

	DirCreate($path_to_allPhotos_3mmSRGB)

	; Platzieren von ersten Bild
	Send($path_to_photo)
	Send("{ENTER}")

	Sleep(500)
	$mousePos = MouseGetPos()
	$mouseX = $mousePos[0]
	$mouseY = $mousePos[1]

	Sleep(200)
	MouseMove(1000, 500)
	MouseClick($MOUSE_CLICK_LEFT)


	; Bild zentrieren
	Send("!n")		; fuer Datei-Menue
	Send("{DOWN}")	; Pfeiltaste nach unten um Menue zu oeffnen

	For $i1 = 0 to 4 Step 1
		Send("{DOWN}")	; Pfeiltaste nach unten
	Next


	Sleep(500)
	Send("{ENTER}")

	; Bild in der Hoehe mittig ausrichten
	Send("!n")		; fuer Datei-Menue
	Send("{DOWN}")	; Pfeiltaste nach unten um Menue zu oeffnen


	For $i1 = 0 to 7 Step 1
		Send("{DOWN}")	; Pfeiltaste nach unten
	Next

	Sleep(500)
	Send("{ENTER}")

	Sleep(100)
	; Bild rechts oben nach aussen ziehen
	;MouseClickDrag($MOUSE_CLICK_LEFT, 1401, 184, 1425, 168)
	MouseClickDrag($MOUSE_CLICK_LEFT, 1388, 173, 1396, 164)

	;$mousePos = MouseGetPos()
	;$mouseX = $mousePos[0]
	;$mouseY = $mousePos[1]

	;MsgBox($MB_SYSTEMMODAL, "", "Mousepos: " & $mouseX & "," &$mouseY)
	;sleep(100000)
	; Bild links unten nach aussen ziehen
	MouseClickDrag($MOUSE_CLICK_LEFT, 288, 949, 278, 961)


	; Farbraum auf sRGB umstellen
	Send("!d")		; fuer Datei-Menue
	Send("d")
	Send("{DOWN}")	; Pfeiltaste nach unten um Menue zu oeffnen
	Send("{DOWN}")	; Pfeiltaste nach unten
	Send("{DOWN}")	; Pfeiltaste nach unten
	Send("{DOWN}")	; Pfeiltaste nach unten

	Sleep(500)
	Send("{ENTER}")

	For $i1 = 0 to 12 Step 1
		Send("{DOWN}")	; Pfeiltaste nach unten
	Next

	Sleep(500)
	Send("{ENTER}")	; Profil auswaehlen


	; exportieren: STRG + ALT + UMSCHALT + S
	Send("^!+s")		; vorsicht: es wird davon ausgegangen, dass JPEG vorab ausgewaehlt ist was Standard ist außer TIF könnte auch sein
	Send("{ENTER}")

	Send($path_to_allPhotos_3mmSRGB & "\" & $filename)
	Sleep(500)
	Send("{ENTER}")


	; eine kleine Pause einlegen um zu warten bis Export-Fenster fertig ist (solange der Bilderexport dauert)
	Sleep(3000)		; 3 sek

	; Fenster schließen
	WinClose("Affinity", "")


	; nicht speichern auswaehlen
	Send("n")

	; eine kleine Pause einlegen
	Sleep(1000)		; 1 sek

Next

; Skript beenden
Exit
#SingleInstance Force
Persistent

#Include blockedCodes.ahk
CurrentDate := FormatTime(, "dd.MM.yyyy")

if !xl := ComObjActive("Excel.Application") {
    TrayTip("Excel is not running")
    return
}

currentUser := IniRead("settings.ini", "Settings", "currentUser")

F1::findBlocked()

findBlocked() {
    
    destinationCountry := unset

    blockCount := 0

    findCountry()

    selectedRange := xl.Selection
    selectedRange.ClearFormats
    selectedRange.Offset(0,-3).ClearFormats
    selectedRange.Offset(0,1).ClearFormats
    selectedRange.Offset(0,4).ClearFormats

    cellLevel:
    for each, cell in selectedRange {
        codeLevel:
        for code, description in RU {
            if InStr(each.Value, code) {
                each.Style := "Bad"
                each.Offset(0,-3).Style := "Bad"
                each.Offset(0,4).Value := "Y"
                each.Offset(0,4).Style := "Bad"
                ; blockcount++
            }
        }
        if destinationCountry = "BY"
            for code, description in BY {
                if InStr(each.Value, code) {
                    each.Style := "Bad"
                    each.Offset(0,-3).Style := "Bad"
                    each.Offset(0,4).Value := "Y"
                    each.Offset(0,4).Style := "Bad"
                    ; blockcount++
                }
            }
        if InStr(each.Offset(0,1).Value, "US") {
            each.Offset(0,1).Style := "Bad"
            each.Offset(0,-3).Style := "Bad"
        }
    }

    for each, cell in selectedRange {
        if each.Offset(0,4).Value = "Y" {
            blockCount++
            blockedPositions .= each.Offset(0,-3).Text ", "
        }
        else
            each.Offset(0,4).Value := "N"
    }

    if blockCount >= 1 {
        if checkResult := xl.ActiveSheet.Range("A:AZ").Find("Equipmnet Check") or checkResult := xl.ActiveSheet.Range("A:AZ").Find("Equipment Check")
            checkResult.Offset(0,1).Value := "Blocked, " currentUser " " CurrentDate "; Remove positions: " RTrim(blockedPositions, ", ") "."
        TrayTip("Found " blockCount " blocks!", "Check complete!")
    }
    else
        TrayTip("Zero blocks so far.", "Check complete!")
        
    findCountry() {
    
        try Found := xl.ActiveSheet.Range("A:AZ").Find("Country:")
        catch Error as e {
            TrayTip("The `"Country:`" cell is not found!")
        }
    
        if RegExMatch(Found.Offset(0,1).Value, "RU|Russia|Russian Federation|RF")
            destinationCountry := "RU"
        else if RegExMatch(Found.Offset(0,1).Value, "BY|Belarus|Belorussia")
            destinationCountry := "BY"
        else {
            TrayTip("The country is not defined. Add two cells: `"Country:`" and `"Russia/Belarus`", then try again!")
            return
        }
    
        return destinationCountry
    
    }
}
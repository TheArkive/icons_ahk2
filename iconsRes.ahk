; ============================================================================
; EXAMPLES
; ============================================================================
ii := icons("shell32.dll")

;;;;;;;;;;;;;;;; Example #1
L := ii.ListAllIcons() ; list all icons and sizes for a single file
list := ""
For i, obj in L {
    list .= (list?"`n`n":"") "Name: " obj.name " / Lang: " obj.lang " / count: " obj.icons.Length "`n"
    list2 := ""
    For i, obj2 in obj.icons {
        list2 .= (list2?"`n":"") "- " obj2.width " x " obj2.height " / " obj2.type
    }
    list .= list2
}

A_Clipboard := list
msgbox "Check icons list from clipboard`n`nThe list is quite long" ; display the result (in clipboard)

;;;;;;;;;;;;;;;; Example #2
ii.SaveGroup(2, "test.ico")
Msgbox "Icon #2 from shell32.dll is now in the script dir.`n`nIt is called 'test.ico'."


;;;;;;;;;;;;;;;; Example #3
_icons := ii.ListIcons(16,true) ; extract array if icons and copy RAW data from RT_ICON_GROUP

_ico := 2                           ; specify index
hicon := ii.GetHICON(_icons,_ico)   ; get HICON
msg := "Loading icon:`n"
     . "Size: " _icons[_ico].width " x " _icons[_ico].height
g := Gui()
g.OnEvent("close",gui_close)
g.Add("Text",,msg)
g.Add("Picture",,"HICON:" hIcon)
g.Show("h320 w320")

gui_close(*) {
    ExitApp
}

msgbox "Check mem usage"

; ============================================================================
; PNG Structure Links:
;   https://www.w3.org/TR/PNG-Structure.html
;   https://www.w3.org/TR/PNG-Chunks.html#C.Additional-chunk-types
;       ... for reading PNG dimensions directly.
;       PNG dims seem to be reverse endian compared to what was expected, so
;       bytes need to be reversed to read the sizes properly in those fields.
; ============================================================================
;   icons class - for exporting and modifying icon resources
;
;   Properties:
;
;       obj.names
;
;           An array of RI_ICON_GROUP resource names and the associated
;           resource language.  Each element in the array is an object with
;           two properties (obj.name / obj.lang).
;
;       obj.file
;
;           The file name passed when the class obj was initially created.
;
;   Methods:
;
;       obj.SaveGroup(name, file_name)
;
;           Saves the specified RT_ICON_GROUP resource name into the
;           specified file.  The file name should end with ".ico"
;
;       obj.List()
;
;           Lists the icon resources into an array in the "names" property.
;           This method also returns the array.  See "obj.names" above for
;           the format of this property.  This method is normally only used
;           internally.
;
;       obj.ListIcons(name, copy := false, free := true)
;
;           Returns an array of icons from the specified RT_ICON_GROUP resource
;           data.  The format of the array is the same as "obj.icons" described
;           above (in the obj.ListAllIcons() method).
;
;           Properties of each array element:
;
;               - obj.w, obj.h, obj.width, obj.height
;                   w/h are the raw w/h, width/height is the "actual height".
;                   For PNG icons, w/h is zero.  But width/height will contain
;                   the actual width/height.
;               - obj.colorCount
;                   Usually 0, unless the image uses a palette.
;               - obj.planes
;                   Usually 1.
;               - obj.bits
;                   Color depth (1,2,4,8,16,32, etc)
;               - obj.size
;                   Not an actual data member.  This is used when writing ICO
;                   data to an ICO file.  This is the size of the icon data,
;                   and should be the same as obj.data.size (see below).
;               - obj.nID
;                   The resource name for an individual icon (type 3).
;               - obj.data
;                   The raw icon data in a buffer object.  If you do not
;                   specity [ copy := true ] when calling this method then this
;                   member is blank ("").
;               - obj.type
;                   This value will either be "ico" or "png".
;
;           The loaded library is automatically freed unless you
;           specify [ free := false ].
;
;       obj.ListAllIcons(copy := false)
;
;           An array of objects.  Each object has properties that pertain to
;           each individual icon in an RT_ICON_GROUP structure.
;
;           Each element properties:
;
;               obj.name (a string)
;               obj.lang (a number)
;               obj.icons (an array of objects)
;
;           The format of each element in obj.icons is the same as the return
;           value from the obj.ListIcons() method.
;
;           If you specify [ copy := true ] then obj.icons.data will contain
;           a buffer that contains the icon data.  In the case of PNGs, this
;           is actually a PNG file.  In the case of icons you will get the
;           icon data as it exists in EXE/DLL files.
;
;       obj.GetHICON(icon_data, index)
;
;           icon_data = This must be the output from [ obj.ListIcons() ].
;                       Please see above for how to use that method.
;
;           index     = A numeric index.  In this case you are not specifying
;                       a width/height (like when using LoadPicture()).
;
;           You can iterate through the icons in icon_data and check the
;           width and height properties prior to calling obj.GetHICON()
;
;           Please see [ obj.ListIcons() ] method above for more details.
;
;           NOTE: The coder is responsible for calling DestroyIcon to properly
;           dispose of the HICON handle.
;
;           This method is provided as a convenience and has no actual
;           advantge over simply calling LoadPicture() and specifying the
;           necessary options to load the desired icon.  See AHK docs.
;
;           This method is mostly useful if a single icon group has several
;           icons of the same size, likely different images, and thus
;           specifying the size with LoadPicture() is insufficient to load
;           the proper icon.  These types of icons are not common.
; ============================================================================
class icons {
    hUpdate := 0, hModule := 0, EnumCb := {}
    file := "", names := []
    
    __New(sFile, name:="") {
        this.file := sFile
        (name) ? this.ListIcons(name) : "" ; load icon group if specified
    }
    __Delete() {
        this._FreeLibrary()
    }
    EnumRes(hModule, sType, p*) { ; sName, [Lang,] lParam
        name := (((sName:=p[1])>>16)=0) ? sName : StrGet(sName)
        (p.Length = 2) ? this.names.Push(name) : this.ResList.Push({name:name, lang:p[2]})
        return true
    }
    List() { ; list all RT_ICON_GROUP resources
        If !(hModule := DllCall("LoadLibrary","Str",this.file, "UPtr"))
            return false
        
        this.EnumCb.fnc := ObjBindMethod(this,"EnumRes")
        this.ResList := []
        
        Loop (p:=2) {
            this.EnumCb.ptr := CallbackCreate(this.EnumCb.fnc,"F",p+2)
            
            If (A_Index=1) {
                r1 := DllCall("EnumResourceNames", "UPtr", hModule, "Ptr", 14, "UPtr", this.EnumCb.ptr, "UPtr", 0, "Int")
            } Else {
                For i, name in this.names
                    DllCall("EnumResourceLanguagesEx", "UPtr", hModule, "Ptr", 14, "Ptr", name, "UPtr", this.EnumCb.ptr
                                                     , "UPtr", 0, "UInt", 0x1, "UShort", 0)
            }
            p++, CallbackFree(this.EnumCb.ptr)
        }
        this.names := this.ResList, this.ResList := [], this.EnumCb.ptr := 0
        r1 := DllCall("FreeLibrary","UPtr",hModule)
        return this.names
    }
    GetHICON(data, idx) {
        hIcon := DllCall("User32\CreateIconFromResourceEx", "UPtr", (_ico := data[idx]).data.ptr
                                                          , "UInt", _ico.size
                                                          , "Int", 1
                                                          , "UInt", 0x30000
                                                          , "Int", _ico.width
                                                          , "Int", _ico.height
                                                          , "UInt", 0, "UPtr")
        return hIcon
    }
    ListAllIcons(copy:=false) {
        list := []
        For i, o in this.List()
            list.Push({name:o.name, lang:o.lang, icons:this.ListIcons(o.name,copy,false)})
        this._FreeLibrary()
        return list
    }
    ListIcons(name, copy:=false, free:=true) { ; load individual icons in RT_ICON_GROUP
        If !this.hModule
            this.hModule := DllCall("LoadLibraryEx", "Str", (this.file), "UPtr", 0, "UInt", 0x2, "UPtr")
        
        buf := this._GetResourceBuffer(this.hModule,name,14)
        icon_count := NumGet(buf,4,"UShort")
        
        off := 6, group := []
        Loop icon_count { ; get icon data, and copy bits
            obj := {type:"ico"}
            obj.w := obj.width  := NumGet(buf,off,"UChar")
            obj.h := obj.height := NumGet(buf,off+1,"UChar")
            obj.colorCount := NumGet(buf,off+2,"UChar")
            
            obj.planes := NumGet(buf,off+4,"UChar")
            obj.bits := NumGet(buf,off+6,"UChar")
            obj.size := NumGet(buf,off+8,"UInt")
            obj.nID := NumGet(buf,off+12,"UShort")
            
            buf2 := this._GetResourceBuffer(this.hModule,obj.nID,3)
            
            If (obj.w=0 || obj.h=0) { ; read png dimensions
                obj.width := (NumGet(buf2,16,"UChar")<<24) | (NumGet(buf2,17,"UChar")<<16) | (NumGet(buf2,18,"UChar")<<8) | NumGet(buf2,19,"UChar")
                obj.height := (NumGet(buf2,20,"UChar")<<24) | (NumGet(buf2,21,"UChar")<<16) | (NumGet(buf2,22,"UChar")<<8) | NumGet(buf2,23,"UChar")
                obj.type := "png"
            }
            
            (copy) ? obj.data := buf2 : obj.data := ""
            off += 14, group.Push(obj)
        }
        (free) ? this._FreeLibrary() : ""
        return group
    }
    SaveGroup(name, sFile) {
        group := this.ListIcons(name,true)
        
        hdr := 6 + (16 * group.Length)
        For i, obj in group
            hdr += obj.size
        
        f := FileOpen(sFile,"w")
        
        buf := Buffer(hdr,0)
        NumPut("UShort", 1, buf, 2)
        NumPut("UShort", group.Length, buf, 4)
        
        off := 6, off2 := hdr
        For i, obj in group {
            NumPut("UChar", obj.w, buf, off)
            NumPut("UChar", obj.h, buf, off+1)
            NumPut("UChar", obj.colorCount, buf, off+2)
            NumPut("UShort", obj.planes, buf, off+4)
            NumPut("UShort", obj.bits, buf, off+6)
            NumPut("UInt", obj.size, buf, off+8)
            NumPut("UInt", off2, buf, off+12)
            off += 16, off2 += obj.size
        }
        
        f.RawWrite(buf)
        
        off2 := hdr
        For i, obj in group
            f.RawWrite(obj.data), off2 += obj.size
        
        f.close()
    }
    _GetResourceBuffer(hModule, name, type) { ; private
        name := IsInteger(name) ? "#" name : name
        fRsc := DllCall("FindResource","UPtr",hModule,"Str",name,"Ptr",type,"UPtr")
        hRsc := DllCall("LoadResource","UPtr",hModule,"UPtr",fRsc,"UPtr")   ; resource handle
        pRsc := DllCall("LockResource","UPtr",hRsc,"UPtr")                  ; resource pointer
        buf := Buffer(DllCall("SizeofResource","UPtr",hModule,"UPtr",fRsc,"UInt"))
        DllCall("RtlCopyMemory", "UPtr", buf.ptr, "UPtr", pRsc, "UPtr", buf.size) ; copy icon group data
        return buf
    }
    _FreeLibrary() { ; private
        If this.hModule
            r1 := DllCall("FreeLibrary","UPtr",this.hModule,"UPtr")
    }
}


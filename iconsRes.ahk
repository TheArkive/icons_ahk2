; ============================================================================
; EXAMPLES - copy AutoHotkey.exe to the script dir for this demo.
; ============================================================================

sFile := ""
Loop Files "AutoHotkey*.exe"
{
    sFile := A_LoopFileName
    Break
}
If (!sFile) {
    Msgbox "Copy any variation of AutoHotkey.exe to the script folder first!"
    return
}

ii := icons(sFile)

;;;;;;;;;;;;;;;; Example #1
L := ii.List(3) ; list all icons and sizes for a single module
list := ""
For i, obj in L {
    list .= (list?"`n":"") "Name: " obj.name " / Lang: " obj.lang
    list2 := ""
    list .= list2
}

msgbox "Individual icons in AutoHotkey.exe (type 3):`n`n" list

;;;;;;;;;;;;;;;; Example #2
L := ii.List() ; list all icons and sizes for a single module
list := ""
For i, obj in L {
    list .= (list?"`n":"") "Name: " obj.name " / Lang: " obj.lang
    list2 := ""
    list .= list2
}

msgbox "Icon groups in AutoHotkey.exe (type 14):`n`n" list

;;;;;;;;;;;;;;;; Example #3
L := ii.ListAllIcons() ; list all icon groups and individual icons + sizes for a single module
list := ""
For i, obj in L {
    list .= (list?"`n`n":"") "Name: " obj.name " / Lang: " obj.lang " / count: " obj.icons.Length "`n"
    list2 := ""
    For i, obj2 in obj.icons {
        list2 .= (list2?"`n":"") "- " obj2.width " x " obj2.height " / " obj2.type " / nID: " obj2.nID
    }
    list .= list2
}

msgbox "All icon groups and individual icons in AutoHotkey.exe (type 3 and 14):`n`n" list

;;;;;;;;;;;;;;;; Example #4
ii := icons("shell32.dll",ico_res := 16) ; one of the computer monitor icons
ii.SaveGroup(ico_res, "test.ico")
Msgbox "Icon #" ico_res " from shell32.dll is now in the script dir.`n`nIt is called 'test.ico'."

;;;;;;;;;;;;;;;; Example #5
_icons := ii.ListIcons(16,true) ; extract array if icons and copy RAW data from RT_ICON_GROUP

_ico := 1                           ; specify index
hicon := ii.GetHICON(_icons,_ico)   ; get HICON
msg := "Loading icon:`n"
     . "Size: " _icons[_ico].width " x " _icons[_ico].height
g := Gui()
g.OnEvent("close",gui_close)
g.OnEvent("escape",gui_close)
g.Add("Text",,msg)
g.Add("Picture",,"HICON:" hIcon)
g.Add("Text",,"Close this window or press ESC to continue the example.")
g.Show("h320 w320")

gui_close(_g) {
    Global g
    g.Destroy()
    g := {hwnd:0}
}

While g.hwnd
    Sleep 100

;;;;;;;;;;;;;;;; Example #6

ii := icons(sFile)
ii.BeginUpdate()
ii.Apply(159,0x409,0x409,"test.ico")
ii.Save()
ii.EndUpdate()

If DirExist("icon_test")
    DirDelete "icon_test", 1
msgbox "Now an Explorer window will open with the new AutoHotkey.exe, and showing the new icon."

DirCreate "icon_test"
FileCopy sFile, "icon_test\" sFile
Run "explorer.exe " Chr(34) A_ScriptDir "\icon_test" Chr(34)

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
;   Usage:
;
;       obj := icons("file_name.exe or dll" [ , name := ""] )
;
;           Specify the module to load at a minimum.  Optionally you can also
;           specify the exact resource name to load from the module.
;
;   Properties:
;
;       obj.names
;
;           After using the obj.List() method (read more below) this member
;           contains an array of icon resource names and the associated language.
;           Each element in the array is an object with two properties (obj.name
;           and obj.lang).  
;
;       obj.file
;
;           The file name passed when the class obj was initially created.
;
;       obj.lang / obj.oldLang
;
;           Currently these properties don't do anything.  You must specify
;           the oldLang and/or lang (new lang) to use when performing resource
;           updates with the Apply() method (read more below).
;
;       obj.hUpdate
;
;           This contains the handle returned after calling the obj.BeginUpdate()
;           method.  This method is set to 0 after a successful call of the
;           obj.EndUpdate() method. (read more below).
;
;   Methods:
;
;       obj.SaveGroup(name, file_name, count := 0)
;
;           Saves the specified RT_ICON_GROUP resource name into the
;           specified file.  The file name should end with ".ico".
;
;           If you specify count, then only the first "count" icons will
;           be saved.  Set count to 0 or ignore this parameter to save
;           all icons of the specified resource name to an ico file.
;
;       obj.List(_type:=14)
;
;           Lists the icon resources into an array in the "names" property.
;           This method also returns the array.  See "obj.names" above for
;           the format of this property.  This method is normally only used
;           internally.
;
;           The default type of icons to list is type 14 (RT_ICON_GROUP).
;
;           If you want to list individual icons, specify type 3.
;
;       obj.ListIcons(name, copy := false, free := true, _type := "res")
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
;               - obj.nID  OR  obj.offset
;                   If (_type = "res") this is the resource name for an individual
;                   icon (type 3).  If (_type = "file") this is the offset from
;                   the beginning of the file where the individual icon raw data
;                   is located.
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
;           value from the obj.ListIcons() method (see above).
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
;
;       obj.Apply(name, oldLang, lang:="", ico_file:="")
;
;           Constructs the new icon resource and saves it to the obj.outObj
;           property.  The icon resource name to replace must be specified.
;
;           If you want to replace an icon resource, specify "oldLang" and "lang".
;           Both values must be the same.
;
;           If you just want to remove an icon resource, specify "oldLang" and
;           leave "lang" blank.
;
;           If you want to add an icon resource to a file that has none,
;           specify oldLang as "" and specify the "lang" to add.
;
;           If you want to remove a version resource, and add a new resource
;           as a different lang, then specify "oldLang" and "lang" accordingly.
;
;       obj.BeginUpdate()
;
;           Preps the file for update. You don't need to do this if it has
;           already been done, ie. if you are updating multiple resources.
;           The output is the handle to the file for the update operation.;
;           This handle is also stored in the obj.hUpdate property.
;
;       obj.Save(hUpdate:=0)
;
;           Performs the resource update.  Returns bool (True = success).
;           You can optionally specify an external hUpdate handle.  Otherwise
;           the internally recorded handle is used if it exists.
;
;       obj.EndUpdate(hUpdate:=0)
;
;           Commits the update changes to the file. Returns bool (True = success).
;           You can specify an external hUpdate handle.  Otherwise the internally
;           recorded handle is used if it exists.
;
;           After you use this method, you should normally destroy, ignore, or
;           overwrite the class instance.
; ============================================================================
class icons {
    hUpdate := 0, hModule := 0, EnumCb := {}
    file := "", names := [], name := ""
    lang := 0, oldLang := 0
    origObj := "", outObj := ""
    LastError := 0
    
    __New(sFile, name:="") {
        this.file := sFile
        (name) ? this.ListIcons(name) : "" ; load icon group if specified
    }
    __Delete() {
        this._FreeLibrary()
    }
    EnumRes(List, hModule, sType, p*) { ; sName, [Lang,] lParam
        name := (((sName:=p[1])>>16)=0) ? sName : StrGet(sName)
        (p.Length = 2) ? List.Push(name) : List.Push({name:name, lang:p[2]})
        return true
    }
    List(_type:=14) { ; list all RT_ICON_GROUP resources
        If !(hModule := DllCall("LoadLibrary","Str",this.file, "UPtr"))
            return false
        
        Loop (p:=2) {
            this.EnumCb.fnc := ObjBindMethod(this,"EnumRes",List:=[])
            this.EnumCb.ptr := CallbackCreate(this.EnumCb.fnc,"F",p+2)
            
            If (A_Index=1) {
                r1 := DllCall("EnumResourceNames", "UPtr", hModule, "Ptr", _type, "UPtr", this.EnumCb.ptr, "UPtr", 0, "Int")
                this.names := List
            } Else {
                For i, name in this.names
                    DllCall("EnumResourceLanguagesEx", "UPtr", hModule, "Ptr", _type, "Ptr", name, "UPtr", this.EnumCb.ptr
                                                     , "UPtr", 0, "UInt", 0x1, "UShort", 0)
                this.names := List
            }
            p++, CallbackFree(this.EnumCb.ptr)
        }
        this.EnumCb.ptr := 0
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
    DestroyIcon(HICON) {
        return DllCall("DestroyIcon","UPtr",HICON)
    }
    ListAllIcons(copy:=false) {
        list := []
        For i, o in this.List()
            list.Push({name:o.name, lang:o.lang, icons:this.ListIcons(o.name,copy,false)}) ; don't free module
        this._FreeLibrary() ; free module manually now
        return list
    }
    ListIcons(name, copy:=false, free:=true, _type:="res") { ; load individual icons in RT_ICON_GROUP
        If (_type="res") {
            If !this.hModule
                this.hModule := DllCall("LoadLibraryEx", "Str", (this.file), "UPtr", 0, "UInt", 0x2, "UPtr")
            buf := this._GetResourceBuffer(this.hModule,name,14)
        } Else {
            SplitPath name,,, &sExt
            If (sExt!="ico") || !FileExist(name)
                throw Error("Invalid file type, or file not found.  File must be an ICO file.")
            buf := FileRead(name,"RAW")
        }
        
        If (NumGet(buf,2,"UShort")!=1)
            throw Error("Invalid ICO file.")
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
            
            If (_type="res") {                          ; read resource data
                obj.nID := NumGet(buf,off+12,"UShort")
                buf2 := this._GetResourceBuffer(this.hModule,obj.nID,3)
            } Else {                                    ; read ICO file data
                obj.offset := NumGet(buf,off+12,"UInt")
                buf2 := Buffer(obj.size,0)
                DllCall("RtlCopyMemory", "UPtr", buf2.ptr, "UPtr", buf.ptr + obj.offset, "UPtr", obj.size)
            }
            
            If (obj.w=0 || obj.h=0) { ; read png dimensions directly
                obj.width := (NumGet(buf2,16,"UChar")<<24) | (NumGet(buf2,17,"UChar")<<16) | (NumGet(buf2,18,"UChar")<<8) | NumGet(buf2,19,"UChar")
                obj.height := (NumGet(buf2,20,"UChar")<<24) | (NumGet(buf2,21,"UChar")<<16) | (NumGet(buf2,22,"UChar")<<8) | NumGet(buf2,23,"UChar")
                obj.type := "png"
            }
            
            obj.data := (copy) ? buf2 : ""
            off += ((_type="res")?14:16), group.Push(obj)
        }
        (free) ? this._FreeLibrary() : ""
        return group
    }
    SaveGroup(name, sFile, count:=0) {
        group := this.ListIcons(name,true)
        count := (!count)?group.Length:count
        
        f := FileOpen(sFile,"w")
        f.RawWrite(this._MakeIconHeader(group,"file",count))
        For i, obj in group
            f.RawWrite(obj.data)
        
        f.close()
    }
    _MakeIconHeader(group, _type:="res", count:=0) { ; private
        hdr_factor := (_type="res") ? 14 : 16
        count := (!count)?group.Length:count
        
        hdr := 6 + (hdr_factor * count)
        For i, obj in group
            hdr += obj.size
        
        buf := Buffer(hdr,0)
        NumPut("UShort", 1, buf, 2)
        NumPut("UShort", count, buf, 4)
        
        off := 6, off2 := hdr
        Loop count {
            obj := group[A_Index]
            NumPut("UChar", obj.w, buf, off)
            NumPut("UChar", obj.h, buf, off+1)
            NumPut("UChar", obj.colorCount, buf, off+2)
            NumPut("UShort", obj.planes, buf, off+4)
            NumPut("UShort", obj.bits, buf, off+6)
            NumPut("UInt", obj.size, buf, off+8)
            If (_type="res")
                NumPut("UShort", obj.nID, buf, off+12)
            Else
                NumPut("UInt", off2, buf, off+12)
            off += hdr_factor, off2 += obj.size
        }
        return buf
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
    Apply(name, oldLang:="", lang:="", ico_file:="") { ; this can't do too much...
        this.name := name
        this.origObj := this.ListIcons(name,true)   ; load original resource
        this.oldLang := oldLang
        If !(this.lang := lang) ; in this case, we are removing and not replacing
            return
        
        If !FileExist(ico_file)
            throw Error("Specified ICO file cannot be found",,"File: " ico_file)
        this.outObj := this.ListIcons(ico_file,true,,"file")    ; prep icon data to replace/add
        
        largest := 0
        For i, obj in this.List(3)
            If (obj.name > largest)
                largest := obj.name
        
        For i, obj in this.outObj
            this.outObj[i].nID := (this.origObj.Has(i)) ? this.origObj[i].nID : ++largest
    }
    BeginUpdate() {
        return (this.hUpdate := DllCall("BeginUpdateResource", "Str", this.file, "UInt", 0, "UPtr"))
    }
    Save(hUpdate:=0) {
        If !(hUpdate := (hUpdate)?hUpdate:this.hUpdate)
            return false
        
        If (this.oldLang) {
            For i, obj in this.origObj { ; delete old icon resource
                If !(DllCall("UpdateResource", "Ptr", hUpdate, "Ptr", 3, "Ptr", obj.nID 
                                             , "UShort", this.oldLang, "UPtr", 0, "UInt", 0)) {
                    msgbox "failed to remove icon type 3 resource`nLastError: " A_LastError
                    return false
                }
            }
            
            If !(DllCall("UpdateResource", "Ptr", hUpdate, "Ptr", 14, "Ptr", this.name 
                                         , "UShort", this.oldLang, "UPtr", 0, "UInt", 0)) {
                msgbox "failed to remove icon type 14 resource`nLastError: " A_LastError
                return false
            }
        }
        
        If (this.lang) {
            For i, obj in this.outObj { ; write new resource icons (type 3)
                If !(DllCall("UpdateResource", "Ptr", hUpdate, "Ptr", 3, "Ptr", obj.nID
                                             , "UShort", this.lang, "Ptr", obj.data.ptr, "UInt", obj.data.size)) {
                    msgbox "failed to update icon type 3 resource`nLastError: " A_LastError
                    return false
                }
            }
            
            buf := this._MakeIconHeader(this.outObj,"res")
            If !(DllCall("UpdateResource", "Ptr", hUpdate, "Ptr", 14, "Ptr", this.name
                                         , "UShort", this.lang, "Ptr", buf.ptr, "UInt", buf.size)) {
                msgbox "failed to update icon type 14 resource`nLastError: " A_LastError
                return false
            }
        }
        return true
    }
    EndUpdate(hUpdate:=0) {
        If !(hUpdate := (hUpdate)?hUpdate:this.hUpdate)
            return false
        return DllCall("EndUpdateResource","UPtr",hUpdate,"Int",(this.hUpdate := 0))
    }
}


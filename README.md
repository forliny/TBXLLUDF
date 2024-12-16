# TBXLLUDF v1.0.1
twinBASIC XLL UDF Addin Example

![image](https://github.com/user-attachments/assets/b842901d-528a-4701-a467-8b1aad1a6df7)

This is a more useful followup to [my initial proof-of-concept](https://github.com/fafalone/HelloWorldXllTB) for creating an XLL Addin using [twinBASIC](https://github.com/twinbasic/documentation/wiki/twinBASIC-Frequently-Asked-Questions-(FAQs)), which is backwards compatible with the VB6/VBA7 language.

### Usage
To compile this demo, you need only open the file in twinBASIC, then File->Build. Make sure to build for the same bitness (32bit or 64bit) as your version of Office. twinBASIC has a dropdown in the toolbar that has `win32` and `win64` options. If you don't know which type of Office you have, in Excel, click on File, then Account, then About Excel:
![image](https://github.com/user-attachments/assets/355ea258-db17-4c02-89e1-0bf27a9b45ff)

This type of addin is loaded on 'Excel Add-ins', not 'Add-ins' or 'COM addins'.

### About
This demo shows how to create User-Defined Functions (UDFs), Excel functions capable of very high performance because they directly execute native compiled code, which is much faster than VBA P-code. Like the event handlers, UDFs are simply functions exported from our XLL, which is a renamed Standard DLL. twinBASIC supports Standard DLLs natively, and making something an export is as simple as adding `[DllExport]` above the procedure:

```vba
    [DllExport]
    Public Function TBXLLUDFRomanNumeral(pIn As XLOPER12) As LongPtr 'LPXLOPER12
        ...
    End Function
    [DllExport]
    Public Function TBXLLUDFNumberName(pIn As XLOPER12) As LongPtr 'LPXLOPER12
        ...
    End Function
```

Those are our two demo functions. `TBXLLUDFRomanNumeral` converts a whole number between 1 and MAXLONG to a Roman numeral, e.g. 9 to IX, and `TBXLLUDFNumberName` converts a whole number between 0 and MAXLONGLONG to its English name, e.g. 21 to Twenty One. The `xlfRegister` command is used to register these for use, and we need to supply 10 different strings describing it:

```vba
    Private Const FuncName0 As String = "TBXLLUDFNumberName" 'Procedure
    Private Const FuncName1 As String = "UU" 'type_text
    Private Const FuncName2 As String = "TBXLLUDFNumberName" 'function_text
    Private Const FuncName3 As String = "Number to name" 'argument_text
    Private Const FuncName4 As String = "1" 'macro_type
    Private Const FuncName5 As String = "tB XLL UDF Add-In" 'category
    Private Const FuncName6 As String = "" 'shortcut_text
    Private Const FuncName7 As String = "" 'help_topic
    Private Const FuncName8 As String = "Returns the text name of a number, e.g. 1 to One" 'function_help
    Private Const FuncName9 As String = "The number to name" 'argument_help1
```

The most important is the name of the procedure, and the type_type which described the return type and argument types. See MSDN for [more info](https://learn.microsoft.com/en-us/office/client-developer/excel/data-types-used-by-excel) on those, for here just know UU means it returns an `XLOPER12` struct, and takes a single one for the argument. Each of these needs to be converted into an `XLOPER12` type, and in the initial project you saw that seemed like quite a pain. But this time, I've added helper functions that create `XLOPER12`s for you, so it's much more approachable now, here's the entire `XlAutoOpen` function:

```vba
    [DllExport]
    Public Function xlAutoOpen() As Integer

        Dim oper(1) As XLOPER12
 
        oper(0) = GetXLString12("Welcome to the tB XLL UDF Demo!")
        oper(1) = GetXLInt12(2)

        Excel12v(xlcAlert, vbNullPtr, 2, oper)
        
        
        Dim xDLL(0) As XLOPER12
        Dim dummy(0) As XLOPER12
        
        Excel12v(xlGetName, xDLL(0), 0, dummy)
        
        
        Dim func1def(10) As XLOPER12
        func1def(0) = xDLL(0)
        func1def(1) = GetXLString12(FuncName0)
        func1def(2) = GetXLString12(FuncName1)
        func1def(3) = GetXLString12(FuncName2)
        func1def(4) = GetXLString12(FuncName3)
        func1def(5) = GetXLString12(FuncName4)
        func1def(6) = GetXLString12(FuncName5)
        func1def(7) = GetXLString12(FuncName6)
        func1def(8) = GetXLString12(FuncName7)
        func1def(9) = GetXLString12(FuncName8)
        func1def(10) = GetXLString12(FuncName9)
        
        Excel12v(xlfRegister, vbNullPtr, 11, func1def)
        
        Dim func2def(10) As XLOPER12
        func2def(0) = xDLL(0)
        func2def(1) = GetXLString12(FuncRoman0)
        func2def(2) = GetXLString12(FuncRoman1)
        func2def(3) = GetXLString12(FuncRoman2)
        func2def(4) = GetXLString12(FuncRoman3)
        func2def(5) = GetXLString12(FuncRoman4)
        func2def(6) = GetXLString12(FuncRoman5)
        func2def(7) = GetXLString12(FuncRoman6)
        func2def(8) = GetXLString12(FuncRoman7)
        func2def(9) = GetXLString12(FuncRoman8)
        func2def(10) = GetXLString12(FuncRoman9)
        
        Excel12v(xlfRegister, vbNullPtr, 11, func2def)
                
        Excel12v(xlFree, vbNullPtr, 1, xDLL)
        Return 1
    End Function

```

Well that covers all the basics for now, dig into the source for all the details. ExcelSDK.twin is portable and meant to be reused in new XLL projects, I'll probably turn it into a package when there's more wrappers to simplify things.

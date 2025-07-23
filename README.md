# 🌈 ColorFontLabel v1.4
Enhanced Label UserControl/ActiveX Control Supporting Color Fonts (i.e. color emojis)

![image](https://github.com/user-attachments/assets/20f47c40-fc97-4c82-9f68-e3c0e72cb4ed) ![image](https://github.com/user-attachments/assets/1b239fa8-4d44-44eb-9e56-cb1737d2091c)



---

(c) 2025 Jon Johnson
Creative Commons Attribution-NonCommercial 4.0 International license.
For commercial use, contact me at fafalone@gmail.com. 

**Updates**

- v1.4 (23 Jul 2025)\
  Thanks to GitHub user 11EJDE11 who created a fork and added these nice changes and
  fixes. Could have just submitted a PR :)

  - Added support for multiple font effect ranges. You can now use AddUnderlineRange,
     AddBoldRange, AddItalicRange, AddStrikethruRange, and AddStretchRange to add
     additional ranges. The original range methods are still available; calling these
     will reset the list of ranges for that effect and insert the new one as the first.
     ClearBoldRanges etc is also available to clear and disable effects.
     Note: If you've specified more than one range, the StdFont property for that effect
           will be ignored and the effects applied.
  - Bug fix: BorderStyle, PictureOffsetX/Y used Boolean instead of ControlBorderStyleConstants/Long.
             Note: May still not work depending on twinBASIC version.
  - Bug fix: SetCustomTransformMatrix arguments improperly translated.
    
- v1.3 (07 Jul 2025) 
   - Bug fix: Font was smaller than regular label with same font/size.
   - Bug fix: PictureStretch size wrong when DPI awareness enabled.

- v1.2 (07 Jul 2025) 
  - Added support for Picture property to set a background image. PictureStretch property sets whether it's stretched to fill the label. If not, you can set an x/y offset with PictureOffsetX/Y.

  - ![image](https://github.com/user-attachments/assets/7bd57627-c4f2-49c1-9184-defa09eddb8b)

- v1.1 (05 Jul 2025) 
    - Now using quicker/more stable DC render target. Thanks to Wayne Phillips for
        this contribution. This works better in the IDE and faster at runtime.
    - (Bug fix) Size/position badly broken when DPI awareness enabled.

There's some size issues on high DPI displays, but since there's a lot fewer problems otherwise and I have to go for now, I'm posting this and leaving the old version available as ucColorFontLabel10.twinproj,

### Features

- Displays color fonts, most commonly used for color emojis. 
- Font effects (bold, italic, underline, strikethru, condense/expand) can be applied only to a specific range.
- Angled text
- Color gradients (linear and radial)
- Locale can be specified
- Several options for word wrap mode
- Set line spacing
- Antialiasing options
- Can act as a drop target for DragDrop from other apps and displays the fancy icons like Explorer.
- Mouse events including MouseWheel
- Background picture property with stretch or offset options
- Normal Label properties like Alignment, RightToLeft, ForeColor/BackColor, etc.
- Comes with OCX version that works in VB6 and Office (both 32bit and 64bit). (OCX can be used in tB as well but the UserControl version from ucColorLabelTest.twinproj is strongly recommended.)

![image](https://github.com/user-attachments/assets/c15e3126-c791-489b-84e7-ee380040c27c)

![image](https://github.com/user-attachments/assets/4bd528d3-33a4-4acf-946e-a7af8ce1161c)

![image](https://github.com/user-attachments/assets/c3ad5dad-8afa-475e-9cd3-e54642a34d6e)

![image](https://github.com/user-attachments/assets/2aebfb7f-7b36-4fc3-a7b4-c9c458c545f4)

### Requirements
- Color Font support is only available on Windows 8 and above. This control should work on Windows 7, but everything will be black and white. Does not work on XP or earlier.
- This project is written in [twinBASIC](https://github.com/twinbasic/documentation/wiki/twinBASIC-Frequently-Asked-Questions-(FAQs)). You'll need [a recent version](https://github.com/twinbasic/twinbasic/releases) to compile it; absolutely minimum Beta 820. Note that if you're not a subscriber the 64bit version will have a splash screen added.
- To use as a UserControl in twinBASIC, the project must reference Windows Development Library for twinBASIC (WinDevLib) v9.1.566 or higher (References->Available packages).

> [!NOTE]
> The project file for building OCXs, ColorFontLabel.twinproj, is configured to register to `HKEY_LOCAL_MACHINE` so that VB6 can see the control. This requires running the IDE as Administrator to compile. If you don't need to build for VB6, in Project Settings, you can change the 'Register to `HKEY_LOCAL_MACHINE`' option, and it will no longer require admin and will register under `HKEY_CURRENT_USER`. twinBASIC, VBA, and Visual Studio will all see it there, as should any other modern host.

### Usage
- For twinBASIC, it's recommended you use this as a .tbcontrol- as source in your project. ucColorFontLabelTest.twinproj shows using this. For a new project, add a reference Windows Development Library for twinBASIC (WinDevLib) v9.1.566 or higher (References->Available packages), then import ucColorFontLabel.tbcontrol and ucColorFontLabel.twin.
- ColorFontLabel.twinproj is the project for building the ActiveX Control version (OCX).


### Known Issues

- In Excel VBA 64bit, and likely other VBA hosts, it does not render in Design Mode.
- Alignment besides left align may interact poorly with TextAngle values besides 0.

### Usage notes

- Limiting font effect ranges must be done at runtime, through the BoldRange, ItalicizeRange, UnderlineRange, StrikethruRange, and StretchRange. The bEnable argument for those methods indicates that **if active**, the  alternate range supplied should be used, e.g. the Font underline property must be set in addition to passing bEnabled=True. This is so it can be toggled  without changing the font every time.
- twinBASIC fully supports Unicode text and color fonts in the editor, so you can set the text through the Properties in design mode or at runtime just by using the string. If you use this as an ActiveX control in VB6 or VBA, note that you'll need to use an alternative ChrW implementation and add emojis with that:
```vba   
    Public Function ChrW2(ByVal CharCode As Long) As String
    Const POW10 As Long = 2 ^ 10
    If CharCode <= &HFFFF& Then ChrW2 = ChrW$(CharCode) Else _
                                ChrW2 = ChrW$(&HD800& + (CharCode And &HFFFF&) \ POW10) & _
                                        ChrW$(&HDC00& + (CharCode And (POW10 - 1)))
    End Function

    ColorFontLabel1.Text = ChrW2(&H1F308) & " ColorFontLabel"
```

- To specify a gradient, call TextLinearGradientSet or TextRadialGradientSet to configure and enable it. The first two arguments must be the first member of an array; see the test project for example. Colors are standard OLE_COLOR values. The positions represent percentages where it changes, e.g. for an evenly spaced 3-color gradient, you'd use 0.0, 0.5, 1.0. Call TextGradientClear to return to a solid color. The coordinate arguments directly track the Direct2D parameters, for more info:
  - https://learn.microsoft.com/en-us/windows/win32/api/d2d1/ns-d2d1-d2d1_linear_gradient_brush_properties
  - https://learn.microsoft.com/en-us/windows/win32/api/d2d1/ns-d2d1-d2d1_radial_gradient_brush_properties

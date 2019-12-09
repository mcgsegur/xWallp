Attribute VB_Name = "ModCorMenus"
Option Explicit
Private Const MIM_BACKGROUND As Long = &H2
Private Const MIM_APPLYTOSUBMENUS As Long = &H80000000

Private Type MENUINFO
    cbSize As Long
    fMask As Long
    dwStyle As Long
    cyMax As Long
    hbrBack As Long
    dwContextHelpID As Long
    dwMenuData As Long
End Type

Private Declare Function DrawMenuBar Lib "user32" _
    (ByVal hWnd As Long) As Long

Private Declare Function GetSubMenu Lib "user32" _
    (ByVal hMenu As Long, ByVal nPos As Long) As Long

Private Declare Function GetMenu Lib "user32" _
    (ByVal hWnd As Long) As Long

Private Declare Function SetMenuInfo Lib "user32" _
    (ByVal hMenu As Long, _
     mi As MENUINFO) As Long

Private Declare Function CreateSolidBrush Lib "gdi32" _
    (ByVal crColor As Long) As Long
    
'Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, _
'       ByVal un As Long, ByVal b As Long, lpMenuItemInfo As MENUITEMINFO) As Long
'Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, _
'       ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long
    
    
Public Sub preparaCorMenus(hWndFrm As Long)
  Dim mi As MENUINFO
  Dim hWndMenu As Long, hWndSubMenu As Long
   With mi
     .cbSize = Len(mi)
     .fMask = MIM_BACKGROUND
     .hbrBack = CreateSolidBrush(vbWhite)
     .fMask = MIM_BACKGROUND Or MIM_APPLYTOSUBMENUS
     hWndMenu = GetMenu(hWndFrm)
     .hbrBack = CreateSolidBrush(RGB(255, 255, 210))
     SetMenuInfo hWndMenu, mi  'main menu bar
'     hWndSubMenu = GetSubMenu(hWndMenu, 0)
'     SetMenuInfo hWndSubMenu, mi 'primeiro menu (item 0)
   End With
   DrawMenuBar hWndFrm
End Sub


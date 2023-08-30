Attribute VB_Name = "modFunctions"
' This is a macro-enabled global template containing the XML and VBA for the DISARM Word Plug-In.

' Copyright (C) 2023 DISARM Foundation

' This program is free software: you can redistribute it and/or modify it under the terms of the GNU 'General Public License
' as published by the Free Software Foundation, either version 3 of the License, or 'any later version.

' This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY;
' without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
' See the GNU General Public License at https://www.gnu.org/licenses/ for more details.

' To report bugs or suggestions for improvement please send email to info@disarm.foundation.
'___________________________________________________________________________________________________________________________

'
' This module contains functions used to handle multi-column list boxes. It does not include
' all the functions in this project, many of which are in modMain.
' Thanks to Greg Maxey for the code to handle multi-column list boxes including all the functions below.
' See https://gregmaxey.com/word_tip_pages/populate_list_combo_boxes_with_advanced_functions.html.
'
' Some of Greg's routines borrow from the work of Chip Pearson, www.cpearson.com, chip@cpearson.com
' and Graham Mayor http://www.gmayor.com/Index.htm.
'

Option Explicit

Private m_lngIndex As Long          'index counter for oListOrComboBox.List
Private m_lngSelIndex As Long       'index of selected items
Private m_lngCount As Long          'number of selected items
Private m_lngFirstSel As Long       'first selected item index
Private m_lngLastSel As Long        'last selected item index
Private m_lngSaveIndex As Long      'saved index to reselect items at end
Private m_strTemp() As String
Private m_lngColCount As Long

Public Function SortListBox(oListOrComboBox As ListBox, lngCol As Long, bAlphabet As Boolean, bAscending As Boolean)

'
' Sort items in a listbox by a specific column
'

Dim varItems As Variant
Dim lngItem As Long, lngItemNext As Long
Dim varTemp As Variant
 
  If gcHandleFuncErrors Then On Error GoTo FUNC_ERR
  PushCallStack "SortListBox"
  
  'Reindex lngCol.  List columns are indexed starting with 0
  lngCol = lngCol - 1
  'Put the items in a variant variable.
  varItems = oListOrComboBox.List
  If bAlphabet Then
    'Sort alphabetically.
    For lngItem = LBound(varItems, 1) To UBound(varItems, 1) - 1
      For lngItemNext = lngItem + 1 To UBound(varItems, 1)
        'Sort Ascending (1)
        If bAscending Then
          If varItems(lngItem, lngCol) > varItems(lngItemNext, lngCol) Then
            For m_lngColCount = 0 To oListOrComboBox.ColumnCount - 1
              varTemp = varItems(lngItem, m_lngColCount)
              varItems(lngItem, m_lngColCount) = varItems(lngItemNext, m_lngColCount)
              varItems(lngItemNext, m_lngColCount) = varTemp
            Next m_lngColCount
          End If
        'Sort Descending (2)
        Else
           If varItems(lngItem, lngCol) < varItems(lngItemNext, lngCol) Then
             For m_lngColCount = 0 To oListOrComboBox.ColumnCount - 1        'Allows sorting of multi-column ListBoxes
               varTemp = varItems(lngItem, m_lngColCount)
               varItems(lngItem, m_lngColCount) = varItems(lngItemNext, m_lngColCount)
               varItems(lngItemNext, m_lngColCount) = varTemp
             Next m_lngColCount
          End If
        End If
      Next lngItemNext
    Next lngItem
  Else
    'Sort the Array Numerically(2)
    '(Substitute CInt with another conversion type (CLng, CDec, etc.) depending on type of numbers in the column)
    For lngItem = LBound(varItems, 1) To UBound(varItems, 1) - 1
      For lngItemNext = lngItem + 1 To UBound(varItems, 1)
        'Sort Ascending (1)
        If bAscending Then
          If CInt(varItems(lngItem, lngCol)) > CInt(varItems(lngItemNext, lngCol)) Then
            For m_lngColCount = 0 To oListOrComboBox.ColumnCount - 1        'Allows sorting of multi-column ListBoxes
              varTemp = varItems(lngItem, m_lngColCount)
              varItems(lngItem, m_lngColCount) = varItems(lngItemNext, m_lngColCount)
              varItems(lngItemNext, m_lngColCount) = varTemp
            Next m_lngColCount
          End If
        'Sort Descending
        Else
          If CInt(varItems(lngItem, lngCol)) < CInt(varItems(lngItemNext, lngCol)) Then
            For m_lngColCount = 0 To oListOrComboBox.ColumnCount - 1        'Allows sorting of multi-column ListBoxes
              varTemp = varItems(lngItem, m_lngColCount)
              varItems(lngItem, m_lngColCount) = varItems(lngItemNext, m_lngColCount)
              varItems(lngItemNext, m_lngColCount) = varTemp
            Next m_lngColCount
          End If
        End If
      Next lngItemNext
    Next lngItem
  End If
  'Set the list to the array
  oListOrComboBox.List = varItems

FUNC_EXIT:
  PopCallStack
  Exit Function

FUNC_ERR:
  GlobalErrHandler
  Resume FUNC_EXIT
End Function

Sub fcnSelectionData(ByRef oListOrComboBox As MSForms.ListBox, lngCountSelected As Long, _
                     lngFirstSelectedItemIndex As Long, lngLastSelectedItemIndex As Long)
'Provides:
'1) Count Of Selected Items 2) Index Number Of First Selected Item 3) Index Number Of Last Selected Item
'

Dim lngCount As Long:   lngCount = 0

  If gcHandleProcErrors Then On Error GoTo PROC_ERR
  PushCallStack "fcnSelectionData"
  
  m_lngFirstSel = -1
  m_lngLastSel = -1
  If oListOrComboBox.ListCount = 0 Then Exit Sub
  With oListOrComboBox
    If .ListCount = 0 Then
      lngCountSelected = 0
      lngFirstSelectedItemIndex = -1
      lngLastSelectedItemIndex = -1
      Exit Sub
    End If
    If .ListIndex < 0 Then
      lngCountSelected = 0
      lngFirstSelectedItemIndex = -1
      lngLastSelectedItemIndex = -1
      Exit Sub
    End If
    For m_lngIndex = 0 To .ListCount - 1
      If .Selected(m_lngIndex) = True Then
        If m_lngFirstSel < 0 Then
          m_lngFirstSel = m_lngIndex
        End If
        lngCount = lngCount + 1
        m_lngLastSel = m_lngIndex
      End If
    Next m_lngIndex
  End With
  lngCountSelected = lngCount
  lngFirstSelectedItemIndex = m_lngFirstSel
  lngLastSelectedItemIndex = m_lngLastSel

PROC_EXIT:
  PopCallStack
  Exit Sub

PROC_ERR:
  GlobalErrHandler
  Resume PROC_EXIT
End Sub

Public Function IsArrayAllocated(arrStrings As Variant) As Boolean

'Returns TRUE if the array is allocated (either a static array or a dynamic array that has been sized with Redim)
'or
'FALSE if the array is not allocated (a dynamic that has not yet been sized with Redim, or a dynamic array that has been Erased).
'Static arrays are always allocated.

'The VBA IsArray function indicates whether a variable is an array, but it does not distinguish between allocated and unallocated arrays.
'It will return TRUE for both allocated and unallocated arrays. This function tests whether the array has actually been allocated.
'This function is just the reverse of IsArrayEmpty.

Dim lngIndex As Long

  If gcHandleFuncErrors Then On Error GoTo FUNC_ERR
  PushCallStack "IsArrayAllocated"
  
  On Error Resume Next
  'If arrStrings is not an array, return FALSE and get out.
  If IsArray(arrStrings) = False Then
    IsArrayAllocated = False
    GoTo FUNC_EXIT
  End If
  'Attempt to get the UBound of the array. If the array has not been allocated, an error will occur. Test Err.Number to see if an error occurred.
  lngIndex = UBound(arrStrings, 1)
  If (Err.Number = 0) Then
    'Under some circumstances, if an array is not allocated, Err.Number will be '0. To acccomodate this case, we test whether LBound <= Ubound. If this
    'is True, the array is allocated. Otherwise, the array is not allocated.
    If LBound(arrStrings) <= UBound(arrStrings) Then
      'No error. array has been allocated.
      IsArrayAllocated = True
    Else
      IsArrayAllocated = False
    End If
  Else
    'Error unallocated array
    IsArrayAllocated = False
  End If

FUNC_EXIT:
  PopCallStack
  Exit Function

FUNC_ERR:
  GlobalErrHandler
  Resume FUNC_EXIT
End Function

Function IsListEmpty(ByVal oListOrComboBox As ListBox) As Boolean
  
'
' Returns True if list box is empty
'

  If gcHandleFuncErrors Then On Error GoTo FUNC_ERR
  PushCallStack "IsListEmpty"
  
  IsListEmpty = True
  If oListOrComboBox.ListCount > 0 Then IsListEmpty = False

FUNC_EXIT:
  PopCallStack
  Exit Function

FUNC_ERR:
  GlobalErrHandler
  Resume FUNC_EXIT
End Function

Public Function fcnSelectedItems(ByRef oListOrComboBox As MSForms.ListBox) As String()

'
' This returns a two dimensional array containing data in selected listbox list items and columns.
'
Dim arrSelectedItems() As String
Dim lngArrIndex As Long

  If gcHandleFuncErrors Then On Error GoTo FUNC_ERR
  PushCallStack "fcnSelectedItems"
  
  If IsListEmpty(oListOrComboBox) Then GoTo FUNC_EXIT
  fcnSelectionData oListOrComboBox, m_lngCount, m_lngFirstSel, m_lngLastSel
  If m_lngCount = 0 Then GoTo FUNC_EXIT
  lngArrIndex = 0
  ReDim arrSelectedItems(m_lngCount - 1, oListOrComboBox.ColumnCount - 1)
  With oListOrComboBox
    For m_lngIndex = 0 To .ListCount - 1
      If .Selected(m_lngIndex) = True Then
        For m_lngColCount = 0 To oListOrComboBox.ColumnCount - 1
          arrSelectedItems(lngArrIndex, m_lngColCount) = .List(m_lngIndex, m_lngColCount)
        Next m_lngColCount
        lngArrIndex = lngArrIndex + 1
      End If
    Next m_lngIndex
  End With
  fcnSelectedItems = arrSelectedItems

FUNC_EXIT:
  PopCallStack
  Exit Function

FUNC_ERR:
  GlobalErrHandler
  Resume FUNC_EXIT
End Function

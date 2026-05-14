VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICompare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule UseMeaningfulName
'@Folder("Interface")
'@ModuleDescription("Interface for polymorphic sort and search helper.")

Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Author: Vincent van Geerestein
' E-mail: vincent@vangeerestein.com
' Description: ICompare interface definition for polymorphic sort and search
' Add-in: RubberDuck (https://rubberduckvba.com/)
' Version: 2025.01.01
'
' Methods
' Assign            Assigns a value to a variable
' Compare           Compares two values and returns -1, 0 or 1.
' Equal             Returns True if two values are equal or False otherwise
' Swap              Swaps the values of two variables

'------------------------------------------------------------------------------
' MIT License
'
' Copyright (c) 2025 Vincent van Geerestein
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.
'------------------------------------------------------------------------------

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Public Declarations
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Enum ECompare
    ecLess = -1
    ecEqual = 0
    ecGreater = 1
End Enum


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Public methods
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'@Description "Assigns a value or an object reference to a variable."
Public Sub Assign(ByRef variable As Variant, ByVal value As Variant)
End Sub


'@Description "Compares two values and returns -1, 0 or 1."
Public Function Compare(ByVal value1 As Variant, ByVal value2 As Variant) As ECompare
End Function


'@Description "Returns True if two values are equal or False otherwise."
Public Function Equal(ByVal value1 As Variant, ByVal value2 As Variant) As Boolean
End Function


'@Description "Swaps the values of two variables."
Public Sub Swap(ByRef variable1 As Variant, ByRef variable2 As Variant)
End Sub

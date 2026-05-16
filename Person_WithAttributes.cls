VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Person"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Person object with ICompare interface."

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

Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Author: Vincent van Geerestein
' E-mail: vincent@vangeerestein.com
' Description: ICompare implementation for a person class
' Dependency: ICompare interface definition module
' Add-in: RubberDuck (https://rubberduckvba.com/)
' Version: 2025.01.01
'
' Methods
' FullName([FullNameOrder]) Returns the full name of the person
'
' Properties
' Title                     Sets or returns the title of the person
' FirstName                 Sets or returns the first name of the person
' LastName                  Sets or returns the last name of the person
' Prefix                    Sets or returns the last name suffix of the person

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Public declarations
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Enum FullNameOrder
    fnoFirstLast = 1
    fnoLastFirst = 2
    fnoInitials = 4
End Enum


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private declarations
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Type TPRIVATE
    Title As String
    FirstName As String
    Prefix As String
    Lastname As String
End Type
Private this As TPRIVATE


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Interface methods
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Implements ICompare


Private Sub ICompare_Assign(variable As Variant, ByVal value As Variant)

    Set variable = value

End Sub


Private Function ICompare_Compare(ByVal value1 As Variant, ByVal value2 As Variant) As ECompare

    Dim Compare As ECompare

    Compare = VBA.StrComp(value1.Lastname, value2.Lastname, vbTextCompare)
    If Compare <> ecEqual Then
        ICompare_Compare = Compare
        Exit Function
    End If

    Compare = VBA.StrComp(value1.Prefix, value2.Prefix, vbTextCompare)
    If Compare <> ecEqual Then
        ICompare_Compare = Compare
        Exit Function
    End If

    ICompare_Compare = VBA.StrComp(value1.FirstName, value2.FirstName, vbTextCompare)

End Function


Private Function ICompare_Equal(ByVal value1 As Variant, ByVal value2 As Variant) As Boolean

    If VBA.StrComp(value1.Lastname, value2.Lastname, vbTextCompare) <> ecEqual Then Exit Function

    If VBA.StrComp(value1.Prefix, value2.Prefix, vbTextCompare) <> ecEqual Then Exit Function

    ICompare_Equal = VBA.StrComp(value1.FirstName, value2.FirstName, vbTextCompare) = ecEqual

End Function


Private Sub ICompare_Swap(variable1 As Variant, variable2 As Variant)

    Dim obj As Variant
    Set obj = variable1
    Set variable1 = variable2
    Set variable2 = obj

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Public methods
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Title() As String
Attribute Title.VB_Description = "Sets or returns the title of the person."

    Title = this.Title

End Property
Public Property Let Title(ByVal RHS As String)

    this.Title = Capitalize(RHS)

End Property


Public Property Get FirstName() As String
Attribute FirstName.VB_Description = "Sets or returns the first name of the person."

    FirstName = this.FirstName

End Property
Public Property Let FirstName(ByVal RHS As String)

    this.FirstName = Capitalize(RHS)

End Property


Public Property Get Prefix() As String
Attribute Prefix.VB_Description = "Sets or returns the last name suffix of the person."

    Prefix = this.Prefix

End Property
Public Property Let Prefix(ByVal RHS As String)

    this.Prefix = LCase$(RHS)

End Property


Public Property Get Lastname() As String
Attribute Lastname.VB_Description = "Sets or returns the last name of the person."

    Lastname = this.Lastname

End Property
Public Property Let Lastname(ByVal RHS As String)

    this.Lastname = Capitalize(RHS)

End Property


Public Function FullName(Optional ByVal Order As FullNameOrder = fnoFirstLast) As String
Attribute FullName.VB_Description = "Returns the full name of a person in a selected format."
Attribute FullName.VB_UserMemId = 0

    With this
        Dim Name As String
        If .Title <> vbNullString Then
            Name = .Title & " "
        End If

        If Order And fnoFirstLast Then
            If .Prefix = vbNullString Then
                If Order And fnoInitials Then
                    Name = Name & GetInitials(.FirstName) & " " & .Lastname
                Else
                    Name = Name & .FirstName & " " & .Lastname
                End If
            Else
                If Order And fnoInitials Then
                    Name = Name & GetInitials(.FirstName) & " " & .Prefix & " " & .Lastname
                Else
                    Name = Name & .FirstName & " " & .Prefix & " " & .Lastname
                End If
            End If
        Else
            If .Prefix = vbNullString Then
                If Order And fnoInitials Then
                    Name = Name & .Lastname & ", " & GetInitials(.FirstName)
                Else
                    Name = Name & .Lastname & ", " & .FirstName
                End If
            Else
                If Order And fnoInitials Then
                    Name = Name & .Prefix & " " & .Lastname & ", " & GetInitials(.FirstName)
                Else
                    Name = Name & .Prefix & " " & .Lastname & ", " & .FirstName
                End If
            End If
        End If

        FullName = Trim$(Name)
    End With

End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private methods
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function GetInitials(ByVal str As String) As String

    Dim Word As Variant, Result As String
    For Each Word In Split(UCase$(str))
        Result = Result + Left$(Word, 1) & "."
    Next

    GetInitials = Left$(Result, Len(Result) - 1)

End Function


Private Function Capitalize(ByVal str As String) As String

    Dim Word As Variant, Result As String
    For Each Word In Split(LCase$(str))
        Result = Result + UCase$(Left$(Word, 1)) & Mid$(Word, 2) & " "
    Next

    Capitalize = Left$(Result, Len(Result) - 1)

End Function

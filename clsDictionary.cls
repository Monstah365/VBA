VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDictionaryExt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ********************************************* '
' Constants and Private Variables
' ********************************************* '

#Const DictionaryAvailable = True

#If Mac Or Not DictionaryAvailable Then

Private mac_KeyValues As Collection
Private mac_Keys() As Variant
Private mac_Items() As Variant
Private mac_KeyObjects As Collection
Private mac_CompareMode As CompareMethod

#Else

Private dict_ As Object

#End If

' ********************************************* '
' Types
' ********************************************* '

Public Enum CompareMethod
    BinaryCompare = VBA.vbBinaryCompare
    TextCompare = VBA.vbTextCompare
    DatabaseCompare = VBA.vbDatabaseCompare
End Enum

' ********************************************* '
' Properties
' ********************************************* '

Public Property Get CompareMode() As CompareMethod
    #If Mac Or Not DictionaryAvailable Then
        CompareMode = mac_CompareMode
    #Else
        CompareMode = dict_.CompareMode
    #End If
End Property

Public Property Let CompareMode(Value As CompareMethod)
    #If Mac Or Not DictionaryAvailable Then
        If Me.Count > 0 Then
            err.Raise 5 ' Invalid procedure call or argument
        End If
    
        mac_CompareMode = Value
    #Else
        dict_.CompareMode = Value
    #End If
End Property

Public Property Get Count() As Long
    'Get the number of items in the "Dictionary"
    #If Mac Or Not DictionaryAvailable Then
        Count = mac_KeyValues.Count
    #Else
        Count = dict_.Count
    #End If
End Property

Public Property Get Item(Key As Variant) As Variant
    #If Mac Or Not DictionaryAvailable Then
        Dim KeyValue As Variant
        KeyValue = GetKeyValue(Key)
        
        If Not IsEmpty(KeyValue) Then
            If VBA.IsObject(KeyValue(2)) Then
                Set Item = KeyValue(2)
            Else
                Item = KeyValue(2)
            End If
        Else
            ' Not found, Return Empty
        End If
    #Else
        If VBA.IsObject(dict_.Item(Key)) Then
            Set Item = dict_.Item(Key)
        Else
            Item = dict_.Item(Key)
        End If
    #End If
End Property

Public Property Let Item(Key As Variant, Value As Variant)
    'Set Item value (Non Object)
    #If Mac Or Not DictionaryAvailable Then
        If Me.Exists(Key) Then
            ReplaceKeyValue GetKeyValue(Key), Key, Value
        Else
            AddKeyValue Key, Value
        End If
    #Else
        dict_.Item(Key) = Value
    #End If
End Property
Public Property Set Item(Key As Variant, Value As Variant)
    'Set Item if Object
    #If Mac Or Not DictionaryAvailable Then
        If Me.Exists(Key) Then
            ReplaceKeyValue GetKeyValue(Key), Key, Value
        Else
            AddKeyValue Key, Value
        End If
    #Else
        Set dict_.Item(Key) = Value
    #End If
End Property

Public Property Let Key(Previous As Variant, Updated As Variant)
    'Update Key
    #If Mac Or Not DictionaryAvailable Then
        Dim KeyValue As Variant
        KeyValue = GetKeyValue(Previous)
    
        If Not VBA.IsEmpty(KeyValue) Then
            ReplaceKeyValue KeyValue, Updated, KeyValue(2)
        End If
    #Else
        dict_.Key(Previous) = Updated
    #End If
End Property

' ********************************************* '
' Public Functions
' ********************************************* '

Public Sub Add(Key As Variant, Item As Variant)
    ' Add a new key and item to the dictionary
    #If Mac Or Not DictionaryAvailable Then
        If Not Me.Exists(Key) Then
            AddKeyValue Key, Item
        Else
            ' This key is already associated with an element of this collection
            err.Raise 457
        End If
    #Else
        dict_.Add Key, Item
    #End If
End Sub

Public Function Exists(Key As Variant) As Boolean
    ' Determine if a given key is in the dictionary."
    #If Mac Or Not DictionaryAvailable Then
        Exists = Not IsEmpty(GetKeyValue(Key))
    #Else
        Exists = dict_.Exists(Key)
    #End If
End Function


Public Function Items() As Variant
    ' Get an array containing all items in the dictionary
    #If Mac Or Not DictionaryAvailable Then
        If Me.Count > 0 Then
            Items = mac_Items
        Else
            Items = VBA.Split("")
        End If
    #Else
        Items = dict_.Items
    #End If
End Function

Public Function Keys() As Variant
   ' Get an array containing all keys in the dictionary
    #If Mac Or Not DictionaryAvailable Then
        If Me.Count > 0 Then
            Keys = mac_Keys
        Else
            Keys = VBA.Split("")
        End If
    #Else
        Keys = dict_.Keys
    #End If
End Function

Public Sub Remove(Key As Variant)
    ' Remove a given key from the dictionary."
    #If Mac Or Not DictionaryAvailable Then
        Dim KeyValue As Variant
        KeyValue = GetKeyValue(Key)
    
        If Not VBA.IsEmpty(KeyValue) Then
            RemoveKeyValue KeyValue
        Else
            ' Application-defined or object-defined error
            err.Raise 32811
        End If
    #Else
        dict_.Remove Key
    #End If
End Sub

Public Sub RemoveAll()
    ' Remove all information from the dictionary
    #If Mac Or Not DictionaryAvailable Then
        Set mac_KeyValues = New Collection
    
        Erase mac_Keys
        Erase mac_Items
    #Else
        dict_.RemoveAll
    #End If
End Sub

' ********************************************* '
' Private Functions
' ********************************************* '

#If Mac Or Not DictionaryAvailable Then

    ' ********************************************* '
    ' Private Functions for Mac Emulation
    ' ********************************************* '
    
    
    Private Function GetKeyValue(Key As Variant) As Variant
        On Error Resume Next
        GetKeyValue = mac_KeyValues(GetFormattedKey(Key))
        err.Clear
    End Function
    
    Private Sub AddKeyValue(Key As Variant, Value As Variant, Optional Index As Long = -1)
        If Me.Count = 0 Then
            ReDim mac_Keys(0 To 0)
            ReDim mac_Items(0 To 0)
        Else
            ReDim Preserve mac_Keys(0 To UBound(mac_Keys) + 1)
            ReDim Preserve mac_Items(0 To UBound(mac_Items) + 1)
        End If
    
        Dim FormattedKey As String
        FormattedKey = GetFormattedKey(Key)
    
        If Index >= 0 And Index < mac_KeyValues.Count Then
            ' Shift keys/items after adding, including index into empty last slot
            Dim i As Long
            For i = UBound(mac_Keys) To Index + 1 Step -1
                mac_Keys(i) = mac_Keys(i - 1)
                If VBA.IsObject(mac_Items(i - 1)) Then
                    Set mac_Items(i) = mac_Items(i - 1)
                Else
                    mac_Items(i) = mac_Items(i - 1)
                End If
            Next i
    
            ' Add key/item at index
            mac_Keys(Index) = Key
            If VBA.IsObject(Value) Then
                Set mac_Items(Index) = Value
            Else
                mac_Items(Index) = Value
            End If
    
            ' Add key-value at proper index
            mac_KeyValues.Add Array(FormattedKey, Key, Value), FormattedKey, Before:=Index + 1
        Else
            ' Add key-value as last item
            If VBA.IsObject(Key) Then
                Set mac_Keys(UBound(mac_Keys)) = Key
            Else
                mac_Keys(UBound(mac_Keys)) = Key
            End If
            If VBA.IsObject(Value) Then
                Set mac_Items(UBound(mac_Items)) = Value
            Else
                mac_Items(UBound(mac_Items)) = Value
            End If
    
            mac_KeyValues.Add Array(FormattedKey, Key, Value), FormattedKey
        End If
    End Sub
    
    Private Sub ReplaceKeyValue(KeyValue As Variant, Key As Variant, Value As Variant)
        Dim Index As Long
        Dim i As Integer
    
        Index = GetKeyIndex(KeyValue(1))
    
        ' Remove existing Value
        RemoveKeyValue KeyValue, Index
    
        ' Add new Key Value back
        AddKeyValue Key, Value, Index
    End Sub
    
    Private Sub RemoveKeyValue(KeyValue As Variant, Optional ByVal Index As Long = -1)
        Dim i As Long
        If Index = -1 Then
            Index = GetKeyIndex(KeyValue(1))
        End If
    
        If Index >= 0 And Index <= UBound(mac_Keys) Then
            ' Shift keys/items after index down
            For i = Index To UBound(mac_Keys) - 1
                mac_Keys(i) = mac_Keys(i + 1)
    
                If VBA.IsObject(mac_Items(i + 1)) Then
                    Set mac_Items(i) = mac_Items(i + 1)
                Else
                    mac_Items(i) = mac_Items(i + 1)
                End If
            Next i
    
            ' Resize keys/items to remove empty slot
            If UBound(mac_Keys) = 0 Then
                Erase mac_Keys
                Erase mac_Items
            Else
                ReDim Preserve mac_Keys(0 To UBound(mac_Keys) - 1)
                ReDim Preserve mac_Items(0 To UBound(mac_Items) - 1)
            End If
        End If
    
        mac_KeyValues.Remove KeyValue(0)
        RemoveObjectKey KeyValue(1)
    End Sub
    
    Private Function GetFormattedKey(Key As Variant) As String
        If VBA.IsObject(Key) Then
            GetFormattedKey = GetObjectKey(Key)
        ElseIf VarType(Key) = VBA.vbBoolean Then
            GetFormattedKey = IIf(Key, "-1__-1", "0__0")
        ElseIf VarType(Key) = VBA.vbString Then
            GetFormattedKey = Key
    
            If Me.CompareMode = CompareMethod.BinaryCompare Then
                ' Collection does not have method of setting key comparison
                ' So case-sensitive keys aren't supported by default
                ' Append lowercase characters to original key
                ' AbC -> AbC___b, abc -> abc__abc, ABC -> ABC_____
                Dim Lowercase As String
                Lowercase = ""
    
                Dim i As Integer
                Dim Char As String
                Dim Ascii As Integer
                For i = 1 To VBA.Len(GetFormattedKey)
                    Char = VBA.Mid$(GetFormattedKey, i, 1)
                    Ascii = VBA.Asc(Char)
                    If Ascii >= 97 And Ascii <= 122 Then
                        Lowercase = Lowercase & Char
                    Else
                        Lowercase = Lowercase & "_"
                    End If
                Next i
    
                If Lowercase <> "" Then
                    GetFormattedKey = GetFormattedKey & "__" & Lowercase
                End If
            End If
        Else
            ' For numbers, add duplicate to distinguish from strings
            ' ->  123  -> "123__123"
            '    "123" -> "123"
            GetFormattedKey = VBA.CStr(Key) & "__" & CStr(Key)
        End If
    End Function
    
    Private Function GetObjectKey(ObjKey As Variant) As String
        Dim i As Integer
        For i = 1 To mac_KeyObjects.Count
            If mac_KeyObjects.Item(i) Is ObjKey Then
                GetObjectKey = "__object__" & i
                Exit Function
            End If
        Next i
    
        mac_KeyObjects.Add ObjKey
        GetObjectKey = "__object__" & mac_KeyObjects.Count
    End Function
    
    Private Sub RemoveObjectKey(ObjKey As Variant)
        Dim i As Integer
        For i = 1 To mac_KeyObjects.Count
            If mac_KeyObjects.Item(i) Is ObjKey Then
                mac_KeyObjects.Remove i
                Exit Sub
            End If
        Next i
    End Sub
    
    Private Function GetKeyIndex(Key As Variant) As Long
        Dim i As Long
        For i = 0 To UBound(mac_Keys)
            If VBA.IsObject(mac_Keys(i)) And VBA.IsObject(Key) Then
                If mac_Keys(i) Is Key Then
                    GetKeyIndex = i
                    Exit For
                End If
            ElseIf VBA.IsObject(mac_Keys(i)) Or VBA.IsObject(Key) Then
                ' Both need to be objects, skip
            ElseIf mac_Keys(i) = Key Then
                GetKeyIndex = i
                Exit For
            End If
        Next i
    End Function

#End If

Private Sub Class_Initialize()
    #If Mac Or Not DictionaryAvailable Then
        Set mac_KeyValues = New Collection
        Erase mac_Keys
        Erase mac_Items
        Set mac_KeyObjects = New Collection
    #Else
        Set dict_ = CreateObject("Scripting.Dictionary")
    #End If
End Sub

Private Sub Class_Terminate()
    #If Mac Or Not DictionaryAvailable Then
        Set mac_KeyValues = Nothing
        Set mac_KeyObjects = Nothing
    #Else
        Set dict_ = Nothing
    #End If
End Sub

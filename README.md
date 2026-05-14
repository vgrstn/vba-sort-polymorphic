# vba-sort-polymorphic
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
![Platform](https://img.shields.io/badge/Platform-VBA%20(Excel%2C%20Access%2C%20Word%2C%20Outlook%2C%20PowerPoint)-blue)
![Architecture](https://img.shields.io/badge/Architecture-x86%20%7C%20x64-lightgrey)
![Rubberduck](https://img.shields.io/badge/Rubberduck-Ready-orange)

VBA standard module for polymorphic sorting and searching via `ICompare` interface — sort any array type, including object arrays, through a pluggable compare strategy.

Uses a **Dual-Pivot QuickSort + InsertionSort** algorithm (threshold 47). Built-in implementations handle numbers, strings (binary or text compare), and objects. Custom implementations cover any other type.

---

## 📦 Features

- **Polymorphic** — swap in any `ICompare` implementation; no changes to the sort module
- **In-place or by-index** — sort and search either modify the array directly or operate through a `Long()` index array, leaving the original untouched
- **Ascending or descending** — optional `asc` parameter reverses the sort order
- **Binary search** — O(log n) search on a sorted array; returns the lowest index for duplicate values
- **Consecutive search** — find the next occurrence of a value after a prior result
- **Auto-select** — numeric arrays use `CompareDefault`; string arrays respect `VbCompareMethod`; object arrays require `Set CompareCustom`
- Works with any VBA type that can be compared (numbers, strings, dates, objects); UDT arrays are not supported
- Pure VBA, zero dependencies, Rubberduck-friendly annotations

---

## 📁 Files

| File | Type | Description |
|---|---|---|
| `ICompare.cls` | Interface | Defines `ECompare`, `Assign`, `Compare`, `Equal`, `Swap` |
| `CompareDefault.cls` | Class | Comparison by VB operators (`<`, `=`) — numbers, strings, dates |
| `CompareText.cls` | Class | Case-insensitive string comparison (`vbTextCompare`) |
| `CompareBinary.cls` | Class | Case-sensitive string comparison (`vbBinaryCompare`) |
| `Person.cls` | Class | Example object Class implementing `ICompare` |
| `PolymorphicSort.bas` | Module | Sort, search, and IsSorted — the main entry point |

---

## ⚙️ Public Interface

### `PolymorphicSort` module

| Member | Description |
|---|---|
| `Sort arr [, idx [, method [, asc]]]` | Sorts `arr` in place, or by index if `idx` is provided. `asc = False` reverses the order. |
| `Search(arr, value [, idx [, method [, start]]])` | Binary search in a sorted array. Returns the lowest matching index, or `Null` if not found. Pass a prior result as `start` to find the next duplicate. |
| `IsSorted(arr [, idx [, method]])` | Returns `True` if `arr` is sorted (ascending or descending). |
| `CompareCustom` *(Get/Set)* | Sets or returns the active `ICompare` implementation. Set to `Nothing` to restore auto-selection. |

### `ICompare` interface

| Method | Description |
|---|---|
| `Assign(variable, value)` | Assigns a value or object reference to a variable (`Let` or `Set`) |
| `Compare(value1, value2)` | Returns `ecLess` (-1), `ecEqual` (0), or `ecGreater` (1) |
| `Equal(value1, value2)` | Returns `True` if the two values are equal |
| `Swap(variable1, variable2)` | Swaps the contents of two variables |

All methods require a non-empty one-dimensional array. `vbErrorTypeMismatch (13)` is raised for multi-dimensional or empty arrays.

---

## 🚀 Quick Start

```vb
' Sort numbers in place (auto-selects CompareDefault)
Dim a() As Variant
a = Array(5, 3, 8, 1, 4)
Sort a                           ' -> (1, 3, 4, 5, 8)

' Sort strings case-insensitive
Dim s() As Variant
s = Array("Banana", "apple", "Cherry")
Sort s, , vbTextCompare          ' -> ("apple", "Banana", "Cherry")

' Sort by index (original array unchanged)
Dim idx As Variant
Sort a, idx                      ' idx holds sorted positions

' Binary search
Sort a
Dim pos As Variant
pos = Search(a, 4)               ' -> index of 4, or Null

' Find all duplicates
a = Array(1, 2, 2, 2, 3)
Sort a
pos = Search(a, 2)               ' first occurrence
Do While Not IsNull(pos)
    Debug.Print pos              ' prints each index where a(i) = 2
    pos = Search(a, 2, , , pos)  ' next occurrence
Loop

' Check if sorted
Debug.Print IsSorted(a)          ' True or False
```

---

## 🔑 Object arrays with a custom `ICompare`

Implement `ICompare` in your Class and set it before sorting:

```vb
' Person implements ICompare — sorts by LastName, Prefix, FirstName
Dim people(1 To 3) As Object
Set people(1) = New Person: people(1).FirstName = "Charlie": people(1).Lastname = "Smith"
Set people(2) = New Person: people(2).FirstName = "Alice":   people(2).Lastname = "Jones"
Set people(3) = New Person: people(3).FirstName = "Bob":     people(3).Lastname = "Jones"

Set CompareCustom = New Person   ' Person implements ICompare
Sort people                      ' -> Jones Alice, Jones Bob, Smith Charlie

' Reset to auto-select for the next call
Set CompareCustom = Nothing
```

---

## 🔑 Index-based sorting

When `idx` is passed to `Sort`, the original array is unchanged and a `Long()` index array is returned. All other methods accept the same `idx` to operate in index space:

```vb
Dim names() As Variant
names = Array("Charlie", "Alice", "Bob")

Dim idx As Variant
Sort names, idx
' names is unchanged: ("Charlie", "Alice", "Bob")
' idx maps sorted order: names(idx(0)) = "Alice", etc.

Dim pos As Variant
pos = Search(names, "Bob", idx)
Debug.Print IsSorted(names, idx)  ' -> True
```

---

## 🧠 Algorithm

| Phase | Algorithm | Condition |
|---|---|---|
| Large partitions | Dual-Pivot QuickSort | partition size ≥ 47 |
| Small partitions | Insertion sort | partition size < 47 |

- **Recursive** — straightforward top-down recursion; intended as a readable showcase for interface-based polymorphism
- **Random pivots** — two pivot elements selected randomly from the left and right thirds; guards against worst-case O(n²) on sorted input
- **Three-way partition** — elements less than left pivot / between pivots / greater than right pivot
- **Equal-element optimisation** — when the centre partition is large (> 4/7 of array length), equal elements are clustered to avoid redundant comparisons
- **Direction-aware search** — binary search detects ascending vs. descending order from first and last elements

---

## 📄 License

MIT © 2025 Vincent van Geerestein

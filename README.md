
# VBA-iArray
VBA arrays for 21st century, based on Collections

>iArray is VBA Class Module which provides easy usage of arrays known from different programming languages.

## Methods
### Available methods
[AddAfter](#AddAfter), [AddArrayAfter](#AddArrayAfter), [AddArrayBefore](#AddArrayBefore), [AddBefore](#AddBefore), [Avg *(since v1.0)*](#Avg)  
[Clear](#Clear), [Clone](#Clone), [Contains](#Contains), [ContainsAll *(since v0.3)*](#ContainsAll), [ContainsOnlyNumeric *(since v1.0)*](#ContainsOnlyNumeric), [CountOccurrences](#CountOccurrences)  
[Dequeue](#Dequeue), [Difference](#Difference), [DropLeft](#DropLeft), [DropRight](#DropRight)  
[Enqueue](#Enqueue), [EnqueueArray](#EnqueueArray)  
[First](#First)  
[Head *(since v0.4)*](#Head)  
[Intersect *(since v1.0)*](#Intersect)  
[Join](#Join)  
[Last](#Last)  
[OccurrenceIndexes *(since v1.0)*](#OccurrenceIndexes)  
[Pop](#Pop), [Push](#Push), [PushArray](#PushArray)  
[RemoveDuplicates](#RemoveDuplicates), [Reverse](#Reverse)  
[Shift](#Shift), [Shuffle](#Shuffle), [Sum *(since v1.0)*](#Sum)  
[Tail *(since v0.4)*](#Tail), [ToString](#ToString)  
[Union *(since v1.0)*](#Union), [Unique *(since v0.4)*](#Unique), [Unshift](#Unshift), [UnshiftArray](#UnshiftArray)
### (Default Members)
All elements inside iArray are indexed (from 1 to count of elements) and are available directly by its index number
```vba
dim iArr as new iArray
iArr.PushArray ("a","b","c","d","e","f")
Debug.Print iArr(2) ' "b"
iArr(4) = "Fourth"
Debug.Print iArr.ToString ' {"a";"b";"c";"Fourth";"e";"f"}
```
### .AddAfter
Adds element after the given index. [Return to available methods](#Available-methods)  
- **Affects original iArray**
 - When *index* >= count of elements, inserts value at the end.
 - When *index* < count of elements, inserts value at the beginning

**@param Long index** Index after which will be added an element  
**@param Variant val** One element (String, number, ...) to add into iArray  
**@return Long** Count of elements inside iArray
```vba
Dim iArr As New iArray: iArr.PushArray Array(1, 2, 3, 4, 5)
iArr.AddAfter 4, "Hello"
Debug.Print iArr.ToString ' {1;2;3;4;"Hello";5}
```
### .AddArrayAfter  
Adds elements after the given index. [Return to available methods](#Available-methods)  
- **Affects original iArray**
- When *index* >= count of elements, inserts values at the end.  
- When *index* < count of elements, inserts values at the beginning  

**@param Long index** Index after which will be added elements  
**@param Variant val** Array() or iArray of elements to add into iArray  
**@return Long** Count of elements inside iArray
```vba
Dim iArr As New iArray: iArr.PushArray Array(1, 2, 3, 4, 5)
iArr.AddAfterArray 4, Array(True, False)
Debug.Print iArr.ToString ' {1;2;3;4;True;False;5}
```
### .AddArrayBefore
Adds elements before the given index. [Return to available methods](#Available-methods)  
- **Affects original iArray**
- When *index* > count of elements, inserts values at the end.  
- When *index* <= count of elements, inserts values at the beginning  

**@param Long index** Index before which will be added elements  
**@param Variant val** Array() or iArray of elements to add into iArray  
**@return Long** Count of elements inside iArray
```vba
Dim iArr As New iArray: iArr.PushArray Array(1, 2, 3, 4, 5)
iArr.AddArrayBefore 4, Array(True, False)
Debug.Print iArr.ToString ' {1;2;3;True;False;4;5}
```
### .AddBefore
Adds element before the given index. [Return to available methods](#Available-methods)  
- **Affects original iArray**
- When *index* > count of elements, inserts value at the end.  
- When *index* <= count of elements, inserts value at the beginning  

**@param Long index** Index before which will be added an element  
**@param Variant val** One element (String, number, ...) to add into iArray  
**@return Long** Count of elements inside iArray
```vba
Dim iArr As New iArray: iArr.PushArray Array(1, 2, 3, 4, 5)
iArr.AddBefore 4, "Hello"
Debug.Print iArr.ToString ' {1;2;3;"Hello";4;5}
```
### .Avg
Calculates the average of the numeric iArray. [Return to available methods](#Available-methods)
- **~~Affects original iArray~~**

**@return Variant** Average value / "NaN" if iArray contains non-numeric value(s)
```vba
Dim iArr As New iArray
iArr.PushArray Array("3", 4, 3.1415, 2, "1E-2")
Debug.Print iArr.Avg ' 2.4303
```
### .Clear
Empties iArray. [Return to available methods](#Available-methods)
- **Affects original iArray**
```vba
Dim iArr As New iArray: iArr.PushArray Array(1, 2, 3, 4, 5)
Debug.Print iArr.ToString ' {1;2;3;4;5}
iArr.Clear
Debug.Print iArr.ToString ' {}
```
### .Clone
Makes a hard copy of the iArray. [Return to available methods](#Available-methods) 
- **~~Affects original iArray~~**

**@return iArray** Return exact copy of itself
```vba
Dim iArr As New iArray
iArr.PushArray Array("3", 4, 1, 2, 3, 4, 5, "a", "b", "c", True)
Dim iArrCloned As New iArray
Set iArrCloned = arrToClone.Clone
iArr.Clear
Debug.Print iArrCloned.ToString ' {"3";4;1;2;3;4;5;"a";"b";"c";True}
```
### .Contains
Checks if given value is used inside iArray. [Return to available methods](#Available-methods)  
- **~~Affects original iArray~~**

**@param Variant val** An element (String, number, ...) to be checked for existence in iArray  
**@return Boolean** True = element exists, False = element doesn't exist in iArray
```vba
Dim iArr As New iArray: iArr.PushArray Array(1, 2, 2, 1, 3, 1, 2)
Debug.Print iArr.Contains(1) ' True
Debug.Print iArr.Contains(5) ' False
```
### .ContainsAll
Checks if all given values are used inside iArray. [Return to available methods](#Available-methods)
- **~~Affects original iArray~~**

**@param Array|iArray val** Array of values to by checked if it exists in iArray  
**@return Boolean** True = all exists, False = one or more values doesn't exist
```vba
Dim iArr As New iArray: iArr.PushArray Array(1, 2, 2, 1, 3, 1, 2)
Debug.Print iArr.Contains(Array(1, 2)) ' True
Debug.Print iArr.Contains(Array(1, 2, 5)) ' False
```
### .ContainsOnlyNumeric
Verify that all elements are numbers or convertable to numbers. [Return to available methods](#Available-methods)
- **~~Affects original iArray~~**

**@return Boolean** True = contains only numeric values, False = contains inconvertible values
```vba
Dim iArr As New iArray: iArr.PushArray Array("3", 4, "3.1415", 2, "1E-2")
Debug.Print iArr.ContainsOnlyNumeric ' True
iArr.Push("3a")
Debug.Print iArr.ContainsOnlyNumeric ' False
```
### .CountOccurrences
Checks how many times is given value used inside iArray. [Return to available methods](#Available-methods)
- **~~Affects original iArray~~**

**@param Variant val** One element (String, number, ...) to be checked for occurrence in iArray  
**@return Long** Count of matched occurrences
```vba
Dim iArr As New iArray: iArr.PushArray Array(1, 2, 2, 1, 3, 1, 2)
Debug.Print iArr.CountOccurrences(2) ' 3
Debug.Print iArr.CountOccurrences(4) ' 0
```
### .Dequeue
Removes an element from the beginning of the iArray. [Return to available methods](#Available-methods)  
- **Affects original iArray**

**@return Variant** Removed element or Empty, if iArray is empty
```vba
dim iArr as new iArray
iArr.PushArray Array("First element","Second element","Queued","Next queued")
Debug.Print iArr.Dequeue  ' "First element"
Debug.Print iArr.ToString ' {"Second element";"Queued";"Next queued"}
```
### .Difference
Checks for number of differences between two arrays, what was added/deleted or combination. [Return to available methods](#Available-methods)
- **~~Affects original iArray~~**

**@param iArray oldArray** Old iArray to be compared with current iArray  
**@param Optional String retType** "d" = deleted from old iArray, "a" = added in current iArray, "c" = combination of both (default)  
**@return iArray** iArray with differences found
```vba
Dim iArr1 As New iArray: iArr1.PushArray Array(1, 2, 3)
Dim iArr2 As New iArray: iArr2.PushArray Array(2, 3, 4)
Dim iArr3 As New iArray
Set iArr3 = iArr2.Difference(iArr1) ' = iArr2.Difference(iArr1, "c")
Debug.Print iArr3.ToString ' {1;4}
Set iArr3 = iArr2.Difference(iArr1, "d")
Debug.Print iArr3.ToString ' {1}
Set iArr3 = iArr2.Difference(iArr1, "a")
Debug.Print iArr3.ToString ' {4}
```
### .DropLeft
Remove *n* elements from the beginning of the iArray. If *n* > count of iArray elements, all elements are removed. [Return to available methods](#Available-methods)
- **Affects original iArray**

**@param Long n** Number of elements to be removed  
**@return iArray** iArray of the removed elements
```vba
Dim iArr As New iArray: iArr.PushArray Array(1, 2, 3, "a", "b", "c")
Debug.Print iArr.DropLeft(2).ToString ' {1;2}
Debug.Print iArr.ToString ' {3;"a";"b";"c"}
```
### .DropRight
Remove *n* elements from the end of the iArray. If *n* > count of iArray elements, all elements are removed. [Return to available methods](#Available-methods)
- **Affects original iArray**

**@param Long n** Number of elements to be removed  
**@return iArray** iArray of the removed elements
```vba
Dim iArr As New iArray: iArr.PushArray Array(1, 2, 3, "a", "b", "c")
Debug.Print iArr.DropRight(2).ToString ' {"b";"c"}
Debug.Print iArr.ToString ' {1;2;3;"a"}
```
### .Enqueue
Adds an element at the end of the iArray *(alias for [**.Push**](#Push))*. [Return to available methods](#Available-methods)
- **Affects original iArray**

**@param Variant val** One element (String, number, ...) to add into iArray  
**@return Long** Count of elements inside iArray
```vba
dim iArr as new iArray
iArr.PushArray Array("First element", "Second element")
iArr.Enqueue "Queued"
Debug.Print iArr.ToString ' {First element";"Second element";"Queued"}
```
### .EnqueueArray
Adds elements at the end of the iArray *(alias for [**.PushArray**](#PushArray))*. [Return to available methods](#Available-methods)
- **Affects original iArray**

**@param Variant val** Array() or iArray of elements to add into iArray  
**@return Long** Count of elements inside iArray
```vba
dim iArr as new iArray
iArr.PushArray Array("First element", "Second element")
iArr.EnqueueArray Array("Queued","Next queued")
Debug.Print iArr.ToString ' {"First element";"Second element";"Queued";"Next queued"}
```
### .First
Returns value of the first element of the iArray. [Return to available methods](#Available-methods)  
- **~~Affects original iArray~~**

**@return Variant** Value of the first element or Empty if iArray is Empty
```vba
dim iArr as New iArray
iArr.PushArray Array(1, 2, 3, 4, 5)
Debug.Print iArr.First ' 1
```
### .Head
Returns all elements of iArray, except the last one. [Return to available methods](#Available-methods)
- **~~Affects original iArray~~**
- If there is less than two elements inside original iArray, empty iArray is returned

**@return iArray** Copy of original array, without the last element
```vba
Set iArr = New iArray
iArr.PushArray Array("3", 4, 1, 2, 3, 4, 5, "a", "b", "c", True)
Dim headArr As New iArray
Set headArr = iArr.Head
Debug.Print headArr.ToString ' {"3";4;1;2;3;4;5;"a";"b";"c"}
```
### .Intersect
Search for element that exists in both iArrays. [Return to available methods](#Available-methods)
- **~~Affects original iArray~~**

**@param iArray intArray** iArray to be intersected with current iArray
**@return iArray** iArray with intersected values
```vba
Dim iArr1 As New iArray: iArr1.PushArray Array(1, 2, 3, "a", "b", "a")
Dim iArr2 As New iArray: iArr2.PushArray Array(3, 2, 6, "a", "a", "f")
Dim iArrUnion As iArray
Set iArrUnion = iArr1.Intersect(iArr2)
Debug.Print arrJoined.ToString ' {2;3;"a"}
```
### .Join
Joins two iArrays. [Return to available methods](#Available-methods)  
- **~~Affects original iArrays~~**

**@param iArray jArray** iArray to be joined with current iArray  
**@return iArray**  Joined iArray
```vba
Dim iArr1 As New iArray: iArr1.PushArray Array(1, 2, 3, "a", "b", "c")
Dim iArr2 As New iArray: iArr2.PushArray Array(4, 5, 6, "d", "e", "f")
Dim iArrJoined As iArray
Set iArrJoined = iArr1.Join(iArr2)
Debug.Print arrJoined.ToString ' {1;2;3;"a";"b";"c";4;5;6;"d";"e";"f"}
```
### .Last
Returns value of the last element of the iArray. [Return to available methods](#Available-methods) 
- **~~Affects original iArray~~**

**@return Variant** Value of the last element or Empty if iArray is Empty
```vba
dim iArr as New iArray
iArr.PushArray Array(1, 2, 3, 4, 5)
Debug.Print iArr.Last ' 5
```
### .OccurrenceIndexes
Returns all indexes of a value in iArray. [Return to available methods](#Available-methods)
- **~~Affects original iArray~~**

**@return iArray** iArray containing indexes of occurrence
```vba
dim iArr as New iArray
iArr.PushArray Array(1, 2, True, "Abc", 2, "1", 3, 1, 2)
Debug.Print iArr.OccurenceIndexes(1).ToString ' {1,8}
```
### .Pop
Removes an element from the end of the iArray. [Return to available methods](#Available-methods)  
- **Affects original iArray**

**@return Variant** Removed element or Empty, if iArray is empty
```vba
dim iArr as new iArray
iArr.PushArray Array("First element", "Second element")
Debug.Print iArr.Pop  ' "Second element"
Debug.Print iArr.ToString ' {"First element"}
```
### .Push
Adds an element at the end of the iArray. [Return to available methods](#Available-methods)  
- **Affects original iArray**

**@param Variant val** One element (String, number, ...) to add into iArray  
**@return Long** Count of elements inside iArray
```vba
dim iArr as new iArray
iArr.Push "First element"
Debug.Print iArr.ToString ' {"First element"}
```
### .PushArray
Adds elements at the end of the iArray. [Return to available methods](#Available-methods)  
- **Affects original iArray**

**@param Variant val** Array() or iArray of elements to add into iArray  
**@return Long** Count of elements inside iArray
```vba
dim iArr as new iArray
iArr.PushArray Array("First element", "Second element")
Debug.Print iArr.ToString ' {"First element";"Second element"}
```
### .RemoveDuplicates
Keeps only the first occurrence of the value. [Return to available methods](#Available-methods)
- **Affects original iArray**
- The method without interfering with the original iArray is called [**.Unique**](#Unique)

**@return Long** Count of the removed elements
```vba
Dim iArr As New iArray
iArr.PushArray Array(1, 2, "a", 2, 3, 2, 3.14, "b", True, 4, "a")
Debug.Print iArr.RemoveDuplicates ' 3
Debug.Print iArr.ToString ' {1;2;"a";3;3.14;"b";True;4}
```
### .Reverse
Reverses the content of the iArray. [Return to available methods](#Available-methods)  
- **~~Affects original iArray~~**

**@return iArray** Reversed iArray
```vba
Dim iArr As New iArray
iArr.PushArray Array("3", 4, 1, 2, 3, 4, 5, "a", "b", "c", True)
Dim iArrRev As New iArray
Set iArrRev = iArr.Reverse
Debug.Print iArrRev.ToString ' {True;"c";"b";"a";5;4;3;2;1;4;"3"}
```
### .Shift
Removes an element from the beginning of the iArray. [Return to available methods](#Available-methods)  
- **Affects original iArray**

**@return Variant** Removed element or Empty, if iArray is empty
```vba
dim iArr as new iArray
iArr.PushArray Array("First element", "Second element")
Debug.Print iArr.Shift  ' "First element"
Debug.Print iArr.ToString ' {"Second element"}
```
### .Shuffle
Randomly mixes content of the iArray. [Return to available methods](#Available-methods)  
- **~~Affects original iArray~~**

**@return iArray** Shuffled iArray
```vba
Dim iArr As New iArray
iArr.PushArray Array("3", 4, 1, 2, 3, 4, 5, "a", "b", "c", True)
Dim iArrShufled As New iArray
Set iArrShufled = iArr.Shuffle
Debug.Print iArrShufled.ToString ' e.g. {"3";"c";4;"a";5;3;"b";1;4;2;True}
```
### .Sum
Calculates the sum of the numeric iArray. [Return to available methods](#Available-methods)
- **~~Affects original iArray~~**

**@return Variant** Sum value / "NaN" if iArray contains non-numeric value(s)
```vba
Dim iArr As New iArray
iArr.PushArray Array("3", 4, 3.1415, 2, "1E-2")
Debug.Print iArr.Sum ' 12.1515
```
### .Tail
Returns all elements of iArray, except the first one. [Return to available methods](#Available-methods)
- **~~Affects original iArray~~**
- If there is less than two elements inside original iArray, empty iArray is returned

**@return iArray** Copy of original array, without the first element
```vba
Set iArr = New iArray
iArr.PushArray Array("3", 4, 1, 2, 3, 4, 5, "a", "b", "c", True)
Dim tailArr As New iArray
Set tailArr = iArr.Tail
Debug.Print tailArr.ToString ' {4;1;2;3;4;5;"a";"b";"c";True}
```
### .ToString
Creates string representation of the iArray. [Return to available methods](#Available-methods)  
- **~~Affects original iArray~~**

**@param Optional String delimiter** Optional character to separate the iArray's elements (default = ";")  
**@return String** Formatted representation of tha iArray
```vba
Dim iArr As New iArray
iArr.PushArray Array("a", 123456, Empty, "...", True)
Debug.Print iArr.ToString ' {"a";123456;;"...";True}
```
### .Union
Combines both iArrays and removes duplicities. [Return to available methods](#Available-methods)
- **~~Affects original iArray~~**

**@param iArray uArray** iArray to be unioned with current iArray
**@return iArray** iArray with combined values without duplicities
```vba
Dim iArr1 As New iArray: iArr1.PushArray Array(1, 2, 3, "a", "b", "a")
Dim iArr2 As New iArray: iArr2.PushArray Array(3, 2, 6, "a", "a", "f")
Dim iArrUnion As iArray
Set iArrUnion = iArr1.Union(iArr2)
Debug.Print arrJoined.ToString ' {1;2;3;"a";"b";6;"f"}
```
### .Unique
Returns copy of iArray without duplicated values. [Return to available methods](#Available-methods)
- **~~Affects original iArray~~**

**@return iArray** iArray with unique values
```vba
Dim iArr As New iArray
iArr.PushArray Array("3", 4, 1, 2, 3, 4, 5, "a", "b", "3", "c", "a", True)
Dim uniqueArr As New iArray
Set uniqueArr = iArr.Unique
Debug.Print iArr.Unique.ToString ' {"3";4;1;2;3;5;"a";"b";"c";True}
```
### .Unshift
Adds an element at the beginning of the iArray. [Return to available methods](#Available-methods)  
- **Affects original iArray**

**@param Variant val** One element (String, number, ...) to add into iArray  
**@return Long** Count of elements inside iArray
```vba
dim iArr as new iArray
iArr.PushArray Array("First element", "Second element")
iArr.Unshift "1st"
Debug.Print iArr.ToString ' {"1st";"First element";"Second element"}
```
### .UnshiftArray
Adds elements at the beginning of the iArray. [Return to available methods](#Available-methods) 
- **Affects original iArray**

**@param Variant val** Array() or iArray of elements to add into iArray  
**@return Long** Count of elements inside iArray
```vba
dim iArr as new iArray
iArr.PushArray Array("First element", "Second element")
iArr.UnshiftArray Array("1st","2nd")
Debug.Print iArr.ToString ' {"1st";"2nd";"First element";"Second element"}
```

## Installation
Just import **iArray.cls** into your project, and you can directly use it

## Contributing
Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

Please make sure to update tests as appropriate.

# VBA-iArray
VBA arrays for 21st century, based on Collections

>iArray is VBA Class Module which provides easy usage of arrays known from different programming languages.

## Methods

### (Default Members)
All items inside iArray are indexed (from 1 to count of items) and are available directly by its index number
```vba
dim iArr as new iArray
iArr.PushArray ("a","b","c","d","e","f")
Debug.Print iArr(2) ' "b"
iArr(4) = "Fourth"
Debug.Print iArr.ToString ' {"a","b","c","Fourth","e","f"}
```
### .AddAfter
Adds item after given index.  
When index >= count of items, pushes value at the end.  
When index < count of items, unshifts value at the begining  
**@param Long index** Index after which will be added an item  
**@param Variant val** One item (String, number, ...) to add into iArray  
**@return Long** Count of items inside iArray
```vba
Dim iArr As New iArray: iArr.PushArray Array(1, 2, 3, 4, 5)
iArr.AddAfter 4, "Hello"
Debug.Print iArr.ToString ' {1,2,3,4,"Hello",5}
```
### AddArrayAfter
Adds items after given index.  
When index >= count of items, pushes values at the end.  
When index < count of items, unshifts values at the begining  
**@param Long index** Index after which will be added items  
**@param Variant val** Array() or iArray of items to add into iArray  
**@return Long** Count of items inside iArray
```vba
Dim iArr As New iArray: iArr.PushArray Array(1, 2, 3, 4, 5)
iArr.AddAfterArray 4, Array(True, False)
Debug.Print iArr.ToString ' {1,2,3,4,True,False,5}
```
### AddArrayBefore
Adds items before given index.  
When index > count of items, pushes values at the end.  
When index <= count of items, unshifts values at the begining  
**@param Long index** Index before which will be added items  
**@param Variant val** Array() or iArray of items to add into iArray  
**@return Long** Count of items inside iArray
```vba
Dim iArr As New iArray: iArr.PushArray Array(1, 2, 3, 4, 5)
iArr.AddArrayBefore 4, Array(True, False)
Debug.Print iArr.ToString ' {1,2,3,True,False,4,5}
```
### .AddBefore
Adds item before given index.  
When index > count of items, pushes value at the end.  
When index <= count of items, unshifts value at the begining  
**@param Long index** Index before which will be added an item  
**@param Variant val** One item (String, number, ...) to add into iArray  
**@return Long** Count of items inside iArray
```vba
Dim iArr As New iArray: iArr.PushArray Array(1, 2, 3, 4, 5)
iArr.AddBefore 4, "Hello"
Debug.Print iArr.ToString ' {1,2,3,"Hello",4,5}
```
### .Clear
Empties iArray.
```vba
Dim iArr As New iArray: iArr.PushArray Array(1, 2, 3, 4, 5)
Debug.Print iArr.ToString ' {1,2,3,4,5}
iArr.Clear
Debug.Print iArr.ToString ' {}
```
### .Clone
Makes a hard copy of the iArray.  
**@return iArray** Return exact copy of itself
```vba
Dim iArr As New iArray
iArr.PushArray Array("3", 4, 1, 2, 3, 4, 5, "a", "b", "c", True)
Dim iArrCloned As New iArray
Set iArrCloned = arrToClone.Clone
iArr.Clear
Debug.Print iArrCloned.ToString ' {"3",4,1,2,3,4,5,"a","b","c",True}
```
### .Contains
Checks if given value is used inside iArray. 
**@param Variant val** One item (String, number, ...) to by checked if exists in iArray  
**@return Boolean** True = item exists, False = item doesn't exists in iArray
```vba
Dim iArr As New iArray: iArr.PushArray Array(1, 2, 2, 1, 3, 1, 2)
Debug.Print iArr.Contains(1) ' True
Debug.Print iArr.Contains(5) ' False
```
### .ContainsAll
Checks if all given values are used inside iArray.
**@param Array|iArray val** Array of values to by checked if it exists in iArray  
**@return Boolean** True = all exists, False = one ore more values doesn't exists
```vba
Dim iArr As New iArray: iArr.PushArray Array(1, 2, 2, 1, 3, 1, 2)
Debug.Print iArr.Contains(Array(1,2)) ' True
Debug.Print iArr.Contains(Array(1,2,5)) ' False
```
### .CountOccurences
Checks how many times is given value used inside iArray.  
**@param Variant val** One item (String, number, ...) to by checked  
**@return Long** Count of matched occurences
```vba
Dim iArr As New iArray: iArr.PushArray Array(1, 2, 2, 1, 3, 1, 2)
Debug.Print iArr.CountOccurences(2) ' 3
Debug.Print iArr.CountOccurences(4) ' 0
```
### .Dequeue
Removes an item from the begining of the iArray.  
**@return Variant** Removed item or Empty, if iArray is empty
```vba
dim iArr as new iArray
iArr.PushArray Array("First item","Second item","Queued","Next queued")
Debug.Print iArr.Dequeue  ' "First item"
Debug.Print iArr.ToString ' {"Second item","Queued","Next queued"}
```
### .Difference
Checks for number of differences between two arrays, what was added/deleted or combination.  
**@param iArray oldArray** Old iArray to be compared with current iArray  
**@param Optional String retType** "d" = deleted from old iArray, "a" = added in current iArray, "c" = combination of both (default)  
**@return iArray** iArray with differences found
```vba
Dim iArr1 As New iArray: iArr1.PushArray Array(1, 2, 3)
Dim iArr2 As New iArray: iArr2.PushArray Array(2, 3, 4)
Dim iArr3 As New iArray
Set iArr3 = iArr2.Difference(iArr1) ' = iArr2.Difference(iArr1, "c")
Debug.Print iArr3.ToString ' {1,4}
Set iArr3 = iArr2.Difference(iArr1, "d")
Debug.Print iArr3.ToString ' {1}
Set iArr3 = iArr2.Difference(iArr1, "a")
Debug.Print iArr3.ToString ' {4}
```
### .DropLeft
Remove n items from the beginning of the iArray. If n > count of iArray items, all items are removed.  
**@param Long n** Number of items to be removed  
**@return iArray** iArray of the removed items
```vba
Dim iArr As New iArray: iArr.PushArray Array(1, 2, 3, "a", "b", "c")
Debug.Print iArr.DropLeft(2).ToString ' {1, 2}
Debug.Print iArr.ToString ' {3,"a","b","c"}
```
### .DropRight
Remove n items from the end of the iArray. If n > count of iArray items, all items are removed.  
**@param Long n** Number of items to be removed  
**@return iArray** iArray of the removed items
```vba
Dim iArr As New iArray: iArr.PushArray Array(1, 2, 3, "a", "b", "c")
Debug.Print iArr.DropRight(2).ToString ' {"b","c"}
Debug.Print iArr.ToString ' {1,2,3,"a"}
```
### .Enqueue
Adds an item at the end of the iArray *(internally calls Push)*.  
**@param Variant val** One item (String, number, ...) to add into iArray  
**@return Long** Count of items inside iArray
```vba
dim iArr as new iArray
iArr.PushArray Array("First item", "Second item")
iArr.Enqueue "Queued"
Debug.Print iArr.ToString ' {First item","Second item", "Queued"}
```
### .EnqueueArray
Adds items at the end of the iArray *(internally calls PushArray)*.  
**@param Variant val** Array() or iArray of items to add into iArray  
**@return Long** Count of items inside iArray
```vba
dim iArr as new iArray
iArr.PushArray Array("First item", "Second item")
iArr.EnqueueArray Array("Queued","Next queued")
Debug.Print iArr.ToString ' {"First item","Second item","Queued","Next queued"}
```
### .First
Returns value of the first element of the iArray.  
**@return Variant** Value of the first element or Empty if iArray is Empty
```vba
dim iArr as New iArray
iArr.PushArray Array(1, 2, 3, 4, 5)
Debug.Print iArr.First ' 1
```
### .Join
Joins two iArrays.  
**@param iArray jArray** iArray to be joined with current iArray  
**@return iArray**  Joined iArray
```vba
Dim iArr1 As New iArray: iArr1.PushArray Array(1, 2, 3, "a", "b", "c")
Dim iArr2 As New iArray: iArr2.PushArray Array(4, 5, 6, "d", "e", "f")
Dim iArrJoined As iArray
Set iArrJoined = iArr1.Join(iArr2)
Debug.Print arrJoined.ToString ' {1,2,3,"a","b","c",4,5,6,"d","e","f"}
```
### .Last
Returns value of the last element of the iArray.  
**@return Variant** Value of the last element or Empty if iArray is Empty
```vba
dim iArr as New iArray
iArr.PushArray Array(1, 2, 3, 4, 5)
Debug.Print iArr.Last ' 5
```
### .Pop
Removes an item from the end of the iArray.  
**@return Variant** Removed item or Empty, if iArray is empty
```vba
dim iArr as new iArray
iArr.PushArray Array("First item", "Second item")
Debug.Print iArr.Pop  ' "Second item"
Debug.Print iArr.ToString ' {"First item"}
```
### .Push
Adds an item at the end of the iArray.  
**@param Variant val** One item (String, number, ...) to add into iArray  
**@return Long** Count of items inside iArray
```vba
dim iArr as new iArray
iArr.Push "First item"
Debug.Print iArr.ToString ' {"First item"}
```
### .PushArray
Adds items at the end of the iArray.  
**@param Variant val** Array() or iArray of items to add into iArray  
**@return Long** Count of items inside iArray
```vba
dim iArr as new iArray
iArr.PushArray Array("First item", "Second item")
Debug.Print iArr.ToString ' {"First item","Second item"}
```
### .RemoveDuplicates
Keeps only first occurences of the values.  
**@return Long** Count of the removed items
```vba
Dim iArr As New iArray
iArr.PushArray Array(1, 2, "a", 2, 3, 2, 3.14, "b", True, 4, "a")
Debug.Print iArr.RemoveDuplicates ' 3
Debug.Print iArr.ToString ' {1,2,"a",3,3.14,"b",True,4}
```
### .Reverse
Reverses the content of the iArray.  
**@return iArray** Reversed iArray
```vba
Dim iArr As New iArray
iArr.PushArray Array("3", 4, 1, 2, 3, 4, 5, "a", "b", "c", True)
Dim iArrRev As New iArray
Set iArrRev = iArr.Reverse
Debug.Print iArrRev.ToString ' {True,"c","b","a",5,4,3,2,1,4,"3"}
```
### .Shift
Removes an item from the begining of the iArray.  
**@return Variant** Removed item or Empty, if iArray is empty
```vba
dim iArr as new iArray
iArr.PushArray Array("First item", "Second item")
Debug.Print iArr.Shift  ' "First item"
Debug.Print iArr.ToString ' {"Second item"}
```
### .Shuffle
Randomly mixes content of the iArray.  
**@return iArray** Shuffled iArray
```vba
Dim iArr As New iArray
iArr.PushArray Array("3", 4, 1, 2, 3, 4, 5, "a", "b", "c", True)
Dim iArrShufled As New iArray
Set iArrShufled = iArr.Shuffle
Debug.Print iArrShufled.ToString ' e.g. {"3","c",4,"a",5,3,"b",1,4,2,True}
```
### .ToString
Creates string representation of the iArray.  
**@param Optional String delimiter** Optional character to separate the iArray's items (default = ",")  
**@return String** Formated representation of tha iArray
```vba
Dim iArr As New iArray
iArr.PushArray Array("a",123456, Empty,"...", True)
Debug.Print iArr.ToString ' {"a",123456,,"...",True}
```
### .Unshift
Adds an item at the begining of the iArray.  
**@param Variant val** One item (String, number, ...) to add into iArray  
**@return Long** Count of items inside iArray
```vba
dim iArr as new iArray
iArr.PushArray Array("First item", "Second item")
iArr.Unshift "1st"
Debug.Print iArr.ToString ' {"1st","First item","Second item"}
```
### .UnshiftArray
Adds items at the begining of the iArray.  
**@param Variant val** Array() or iArray of items to add into iArray  
**@return Long** Count of items inside iArray
```vba
dim iArr as new iArray
iArr.PushArray Array("First item", "Second item")
iArr.UnshiftArray Array("1st","2nd")
Debug.Print iArr.ToString ' {"1st", "2nd", First item","Second item"}
```
## To be done
- **.OccurenceIndexes**
Returns all indexes of a value in iArray. If nothing found returns empty iArray
- **.Intersect**
Returns iArray which contains only elements which are same in two given iArrays
- **.Union**
Returns iArray which contains only elements which are same in two given iArrays, without duplicates
- **.Unique**
Returns values of iArray, which are not used two or more times.
- **.IsNumericArray**
Returns True if all values are numbers
- **.Sum**
Returns sum of items in iArray. Only for numeric iArray
- **.Average**
Returns average value of items in iArray. Only for numeric iArray
- **.Tail**
Returns all elements of iArray, except first one. If there is only one element, it will be returned

## Installation
Just import **iArray.cls** into your project and you can directly use it

## Contributing
Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

Please make sure to update tests as appropriate.
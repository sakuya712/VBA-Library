

Contents
<!-- @import "[TOC]" {cmd="toc" depthFrom=1 depthTo=6 orderedList=false} -->
<!-- code_chunk_output -->

- [List class](#list-class)
  - [Methods](#methods)
    - [Add](#add)
    - [AddValue](#addvalue)
    - [Concat](#concat)
    - [Const](#const)
    - [Except](#except)
    - [Includes](#includes)
    - [IndexOf](#indexof)
    - [IsOverlap](#isoverlap)
    - [Item](#item)
    - [Join](#join)
    - [OverlapList](#overlaplist)
    - [Remove](#remove)
    - [RemoveOverlap](#removeoverlap)
    - [Slice](#slice)
    - [ToArray](#toarray)

<!-- /code_chunk_output -->

# List class

A class that reproduces the List type found in other languages.  
The contents are an extension of Collection type.

## Methods

### Add
Argument： Add element  
Return value： None

* Add element. Same as Collection type
* Arrays and objects are also stored as collections

### AddValue
Argument： Add element  
Return value： None

* Stores array values, Range values, and default member values ​​for objects
* If you pass an object that is not of type Collection, an error will be returned

### Concat
Argument： List to be combined (multiple possible)   
Return value： Combined List

* Create a new List by combining Lists
* Works the same as JS Concat

### Const
Argument： None  
Return value： element count

* Returns the number of elements. Same as Collection type

### Except
Argument： List of elements to remove  
Return value： List removed

* Create a List of difference sets.
* Works the same as SQL or C# Except

### Includes
Argument： String, [perfect match]   
Return value： Boolean value

* Returns a boolean value indicating whether there is an element that matches the argument in List
* Works the same as JS Includes
* You can set perfect match or partial match with the second argument. The default value is perfect match

### IndexOf
Argument： String, [perfect match]   
Return value： Index number
  
* Returns the index of the first element in the List that matches the argument
* Works the same as JS or C# IndexOf
* You can set perfect match or partial match with the second argument. The default value is perfect match
* Returns -1 if not found.

### IsOverlap
Argument： None  
Return value： Boolean value

* Check for duplicate elements. Returns True if there is.

### Item
Argument： None  
Return value： Element

* Returns the element. Same as Collection type
* Default member

### Join
Argument： Characters to insert when connecting  
Return value： Concatenated string

* Returns the elements in the List as a single concatenated string.
* Works the same as VBA Join in Array
* If the object cannot be converted to a string, an error will be returned

### OverlapList
Argument： None  
Return value： Duplicate List

* Create a List that fetches duplicate elements
* Duplicate elements do not overlap

### Remove
Argument： Index number   
Return value： None

* Delete the element. Same as Collection type

### RemoveOverlap
Argument： None  
Return value： Unique List

* Create a unique list.

### Slice
Argument： First index number, [Last index number]    
Return value： List type with elements cut off

* Returns the StartIndex th to EndIndex th elements of the index number
* Works the same as JS Slice
* If the second argument is omitted, up to the last index number will be specified

### ToArray
Argument： なし  
Return value： Primary array

* Convert List to primary array
* Index number starts from 1


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
    - [Map](#map)
    - [OverlapList](#overlaplist)
    - [Remove](#remove)
    - [Slice](#slice)
    - [ToArray](#toarray)
    - [ToList](#tolist)
    - [Unique](#unique)
    - [Where](#where)

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

### Map
Argument: Arithmetic operator, arithmetic target (List type, Collection type, primitive type supported)  
Return value： List type

* Create a new List by calculating and combining the contents of elements
* Works the same as C# Select
* Since the arrow operator cannot be used in VBA, it is represented by the following enum type

**ArithmeticOperatorsEnum**
|  Element name   |  Explanation  |
| --------------- | --------|
|  lsSum          |  Sum (+)|
|  lsDifference   |  Difference(-)|
|  lsMultiply     |  Multiply(*)|
|  lsDivide       |  Divide(/)|
|  lsMod          |  Modulo(mod)|
|  lsExponent     |  Exponent(^)|
|  lsConcatenate  |  Concatenate(&)|

### OverlapList
Argument： None  
Return value： Duplicate List

* Create a List that fetches duplicate elements
* Duplicate elements do not overlap

### Remove
Argument： Index number   
Return value： None

* Delete the element. Same as Collection type

### Slice
Argument： First index number, [Last index number]    
Return value： List type with elements cut off

* Returns the StartIndex th to EndIndex th elements of the index number
* Works the same as JS Slice
* If the second argument is omitted, up to the last index number will be specified

### ToArray
Argument： None  
Return value： Primary array

* Convert List to primary array
* Index number starts from 1

### ToList
Argument: The elements to store in the List   
Return value： List type

* Create a List with the same number of elements

### Unique
Argument： None  
Return value： Unique List

* Create a unique list.
* Works the same as C# Distinct

### Where
Argument: Comparison operator, comparison target (corresponding to List type, Collection type, primitive type)  
Return value： List type

* Create a new List, leaving only the elements that meet the conditions
* Works the same as C# Where
* Since the arrow operator cannot be used in VBA, it is represented by the following enum type

**ComparisonOperatorsEnum**
|  Element name     |  Explanation  |
| ------------------| --------|
|  lsEqual          |  Equal(=)|
|  lsNotEqual       |  NotEqual(<>)|
|  lsGreater        |  Greater(>)|
|  lsLess           |  Less(<)|
|  lsGreaterEqual   |  GreaterEqual(>=)|
|  lsLessEqual      |  LessEqual(<=)|
|  lsObjectEqual    |  Reference comparison(Is)|
|  lsLike           |  String comparison(Like)|
|  lsNotLike        |  String comparison(Not Like)|
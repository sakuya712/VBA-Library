

Contents
<!-- @import "[TOC]" {cmd="toc" depthFrom=1 depthTo=6 orderedList=false} -->
<!-- code_chunk_output -->

- [TimerLog class](#timerlog-class)
  - [Example](#example)
  - [Methods](#methods)
    - [Constructor](#constructor)
    - [FinishTime](#finishtime)
  - [Destructor](#destructor)

<!-- /code_chunk_output -->

# TimerLog class

Class for measuring the processing time in the immediate window   

## Example
Measurement is started by the Constructor method, measurement is terminated when this object is discarded (when this function exits), and the result is displayed in the immediate window.
```VB
Sub test()

    Dim log As New TimerLog: log.Constructor ("Example")
    
    'Processing
    Stop
        
End Sub
```
Immediate window
```VB
[Begin] Example
[Finish] Example , 1187[ms] '<=Processing time is displayed
```

## Methods

### Constructor
Argument： A string that is displayed with the time (Function name etc.)  
Return value： None

* Measurement is started

### FinishTime
Argument： None  
Return value： processing time [ms]

* Returns the measurement time result
* Not displayed in immediate window  
* Processing is done separately from the destructor

## Destructor

* The measurement result is displayed in the immediate window.



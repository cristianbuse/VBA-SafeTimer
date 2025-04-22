# VBA-SafeTimer
Reliable, no-crash timer for VBA. Code can be debugged and stopped safely.

```Windows``` notes:
- [```SetTimer```](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-settimer) is used for calling back VBA code at regular intervals
- VBA does all the heavy lifting
  - memory allocation - no need for side-allocations
  - break mode check - so that timer calls are ignored / skipped in break / debug mode
  - timer cleanup - timer is removed when the corresponding form is destroyed

```Mac``` notes:
- not implemented, yet

## Installation

Download the [latest release](https://github.com/cristianbuse/VBA-SafeTimer/releases/latest).

Just import the following code modules in your VBA Project:
* ```SafeTimer.cls```
* ```LibTimers.bas```
* ```TimerForm.frm``` - Alternatively, this can be easily recreated from scratch in 2 easy steps:
  - insert new form
  - rename it to ```TimerForm```

## Usage

Just add a ```SafeTimer``` instance ```WithEvents``` to any of your classes. The ```TimerCall``` event will be raised at the desired interval. See demo.

Optionally, you can pass a ```Variant``` when you call ```SafeTimer.StartTimer``` so that it is passed back in the event.

## Implementation

Windows:
- each instance of ```SafeTimer``` class calls [```SetTimer```](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-settimer) when its ```StartTimer``` method is called
- each callback call (at set interval) to ```SafeTimer.TimerProc``` will raise a ```TimerCall``` event to any *listening* classes
- a single call to [```VariantCopy```](https://learn.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-variantcopy) is made so that we can copy memory (no kernel.dll needed)
- a few [assmebly bytes](https://github.com/cristianbuse/VBA-SafeTimer/blob/bce8e221c3844e58256262c26ab5abe258b368f2/src/LibTimers.bas#L91-L115) (just 23 on x64 and just 20 on x32) are used to redirect the timer call to the correct ```SafeTimer``` instance

```Mac```:
- not implemented

## Demo

Import the following code modules:
* ```Demo.bas``` - run ```DemoMain```
* ```DemoTimer.cls```

There is also a Demo Workbook available for download.

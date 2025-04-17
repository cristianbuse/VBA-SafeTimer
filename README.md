# VBA-SafeTimer
Reliable, no-crash timer for VBA.

Windows notes:
- [```SetTimer```](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-settimer) is used for calling back VBA code at regular intervals
- a single call to [```VariantCopy```](https://learn.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-variantcopy) is made so that we can copy memory (no kernel.dll needed)
- VBA does all the memory allocation
- VBA checks if code is in break mode so that timer calls are ignored

To be continued...

## Installation

To be written...

## Implementation

Windows:
- each instance of ```SafeTimer``` class calls [```SetTimer```](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-settimer) when its ```StartTimer``` method is called
- each callback call (at set interval) to ```SafeTimer.TimerProc``` will raise a ```TimerCall``` event to any *listening* classes

To be continued...

## Demo

To be written...

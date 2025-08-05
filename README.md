# Trin\'s Super Mail Script
## Unlicense
> This is free and unencumbered software released into the public domain.
> 
> Anyone is free to copy, modify, publish, use, compile, sell, or distribute this software, either in source code form or as a compiled binary, for any purpose, commercial or non-commercial, and by any means.
> 
> In jurisdictions that recognize copyright laws, the author or authors of this software dedicate any and all copyright interest in the software to the public domain. We make this dedication for the benefit of the public at large and to the detriment of our heirs and successors. We intend this dedication to be an overt act of relinquishment in perpetuity of all present and future rights to this software under copyright law.
>
> THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
> 
> For more information, please refer to <http://unlicense.org/>

## Dependencies
+ Python 3
+ `pywin32`, `pyyaml` (Run `pip install pywin32 pyyaml` in PowerShell after installing Python)
+ [Classic Outlook](https://support.microsoft.com/en-us/office/install-or-reinstall-classic-outlook-on-a-windows-pc-5c94902b-31a5-4274-abb0-b07f4661edf5)

## Optional Dependencies
+ `click`, `markdown` (Run `pip install click markdown` in PowerShell after installing Python)

## Overview
This repository contains some sample files, but all you really need is `sendmail.py` and the above dependencies.

This is a python utility for sending mail via Outlook. Note that the recent web-app version won't work.
You'll need the classic version from https://support.microsoft.com/en-us/office/install-or-reinstall-classic-outlook-on-a-windows-pc-5c94902b-31a5-4274-abb0-b07f4661edf5.

Run with `--help` for more information.

## Examples (From inside PowerShell)
```ps
python sendmail.py --help
```

```ps
python sendmail.py .\draft.md -M .\data.csv -D ./Attachments
```

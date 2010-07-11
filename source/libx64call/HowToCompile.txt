x64call.asm was adapted from http://dyncall.org/

- Open a Windows 7.1 SDK command prompt
- setenv /Release /x64
- Navigate to this folder
- ml64 /c x64call.asm x64stub.asm
- lib /out:x64call.lib x64call.obj x64stub.obj

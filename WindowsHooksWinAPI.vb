'
' V.1
'
option explicit

public const WH_KEYBOARD_LL   = 13 ' Low level keyboard events (compare with WH_KEYBOARD)
public const HC_ACTION        =  0
public const INPUT_KEYBOARD   =  1
public const KEYEVENTF_KEYUP  =  2 ' Used for dwFlags in INPUT_

type KBDLLHOOKSTRUCT ' {
     vkCode      as long ' virtual key code in range 1 .. 254
     scanCode    as long ' hardware code
     flags       as long ' bit 4: if set -> alt key was pressed
                         ' bit 7: transition state, 0 -> the key is pressed, 1 -> key is being released.
     time        as long
     dwExtraInfo as long
end  type ' }


type INPUT_KI   '   typedef struct tagINPUT ' {
'
'    Used in conjunction with SendInput()
'
     dwType      as long
     wVK         as integer
     wScan       as integer
     dwFlags     as long
     dwTime      as long
     dwExtraInfo as long
     dwPadding   as currency        '   8 extra bytes, because of mouse events
end type ' }

declare function GetModuleHandle     lib "kernel32"     alias "GetModuleHandleA"  ( _
     byVal lpModuleName as string) as long

' CallNextHookEx {
declare function CallNextHookEx      lib "user32"                                 ( _
         byVal hHook        as long, _
         byVal nCode        as long, _
         byVal wParam       as long, _
               lParam       as any ) as long
' }

declare function SetWindowsHookEx    lib "user32"       alias "SetWindowsHookExA" ( _
     byVal idHook     as long   , _
     byVal lpfn       as longPtr, _
     byVal hmod       as long   , _
     byVal dwThreadId as long   ) as long

declare function UnhookWindowsHookEx  lib "user32"                                ( _
     byVal hHook      as long) as long

declare function SendInput           lib "user32"                                 ( _
      byVal nInputs as long, _
      byRef pInputs as any , _
      byVal cbSize  as long) as long

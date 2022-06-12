'
' V.1
'
option explicit

private keyboard_ev as IKeyboardEvent

private hookId as long

sub startLowLevelKeyboardHook(kb as IKeyboardEvent) ' {

    if hookId <> 0 then
       debug.print "Keyboard hook already started"
       exit sub
    end if

    set keyboard_ev = kb

    dim callBack as longPtr ' TODO: as long?
    callBack = getAddressOfCallback(addressOf LowLevelKeyboardProc)

    hookId = SetWindowsHookEx( _
         WH_KEYBOARD_LL                         , _
         callBack                               , _
         GetModuleHandle(vbNullString)          , _
         0 )

end sub ' }

sub endLowLevelKeyboardHook() ' {
    UnhookWindowsHookEx(hookId)
    hookId = 0
end sub ' }

function LowLevelKeyboardProc(byVal nCode as Long, byVal wParam as long, lParam as KBDLLHOOKSTRUCT) as long ' {
'
'   MSDN says (https://docs.microsoft.com/en-us/previous-versions/windows/desktop/legacy/ms644985(v=vs.85))
'
'     nCode
'        A code the hook procedure uses to determine how to process the
'        message.
'
'        If nCode is less than zero, the hook procedure must pass the
'        message to the CallNextHookEx function without further processing and
'        should return the value returned by CallNextHookEx.
'
'        If nCode == HC_ACTION, the wParam and lParam parameters contain
'        information about a keyboard message.
'
'     wParam
'        One of
'          -  WM_KEYDOWN,
'          -  WM_KEYUP
'          -  WM_SYSKEYDOWN
'          -  WM_SYSKEYUP.
'
'     lParam
'          A pointer to a KBDLLHOOKSTRUCT structure.
'
'     Return value
'          If the hook procedure processed the message, it may return a nonzero
'          value to prevent the system from passing the message to the rest of
'          the hook chain or the target window procedure.
'
'     -----------------------------------------------------------------------------
'
'     The hook procedure should process a message in less time than the data entry specified in the
'             LowLevelHooksTimeout
'     value in the following registry key under
'             HKEY_CURRENT_USER\Control Panel\Desktop
'

    dim upOrDown as string
'   dim altKey   as boolean
    dim char     as string

    dim keyEventString as string


    if nCode <> HC_ACTION then
       LowLevelKeyboardProc = CallNextHookEx(0, nCode, wParam, byVal lParam)
       exit function
    end if

    if not keyboard_ev.ev(                             _
          vk_keyCode :=     lParam.vkCode            , _
          pressed    := not lParam.flags      and 128, _
          alt        :=     lParam.flags      and  32, _
          scanCode   :=     lParam.scanCode          , _
          time       :=     lParam.time              ) then

     '
     ' Event was not processed, pass it on:
     '
       LowLevelKeyboardProc = CallNextHookEx(0, nCode, wParam, byVal lParam)
       exit function
    end if ' }

    LowLevelKeyboardProc = 1
    exit function

end function ' }

function getAddressOfCallback(addr as longPtr) as longPtr ' {
      '
      '  TODO: use long instead of longPtr?
      '

      '  See also
      '  https://renenyffenegger.ch/notes/development/languages/VBA/language/operators/addressOf
      '
         getAddressOfCallback = addr
end function ' }

/*******************************************************************************
 *
 *  Copyright (c) Shaun Ivory
 *
 *  Title:       MessageCrackers.h
 *
 *  Date:        10/30/2004
 *
 *  Description: Simple message crackers and dispatchers
 *
 *******************************************************************************/
#pragma once

#pragma warning(disable: 4244)

#define CLASS_INSTANCE_PROP_NAME L"WindowClassInstance"

#define BEGIN_MESSAGE_HANDLERS(DialogClassName)                                                \
    DialogClassName *_pDialog = NULL;                                                          \
    INT_PTR fResult = FALSE;                                                                   \
    if (WM_INITDIALOG == uMsg)                                                                 \
    {                                                                                          \
        _pDialog = new DialogClassName(hWnd);                                                  \
        if (_pDialog)                                                                          \
        {                                                                                      \
            SetPropW(hWnd, CLASS_INSTANCE_PROP_NAME, _pDialog);                                \
        }                                                                                      \
        else                                                                                   \
        {                                                                                      \
            EndDialog(hWnd, IDCANCEL);                                                         \
            return TRUE;                                                                       \
        }                                                                                      \
    }                                                                                          \
    _pDialog = reinterpret_cast<DialogClassName*>(GetPropW(hWnd, CLASS_INSTANCE_PROP_NAME));   \
    switch (uMsg)

#define WINDOWS_MESSAGE_HANDLER(Message, Handler)                                              \
    case (Message):                                                                            \
        {                                                                                      \
            if (_pDialog)                                                                      \
            {                                                                                  \
                fResult = SetDlgMsgResult(hWnd, (Message), _pDialog->Handler(wParam, lParam)); \
            }                                                                                  \
        }                                                                                      \
        break

#define BEGIN_COMMAND_HANDLERS()                                                               \
    case WM_COMMAND:                                                                           \
        {
        
#define COMMAND_HANDLER(Command, Handler)                                                      \
            if ((Command) == LOWORD(wParam))                                                   \
            {                                                                                  \
                if (_pDialog)                                                                  \
                {                                                                              \
                    _pDialog->Handler(wParam, lParam);                                         \
                    fResult = TRUE;                                                            \
                }                                                                              \
                goto ExitSwitch;                                                               \
            }

#define COMMAND_HANDLER_CODE(Command, Code, Handler)                                           \
            if ((Command) == LOWORD(wParam) && (Code) == HIWORD(wParam))                       \
            {                                                                                  \
                if (_pDialog)                                                                  \
                {                                                                              \
                    _pDialog->Handler(wParam, lParam);                                         \
                    fResult = TRUE;                                                            \
                }                                                                              \
                goto ExitSwitch;                                                               \
            }
            
#define END_COMMAND_HANDLERS()                                                                 \
        }                                                                                      \
        break

#define END_MESSAGE_HANDLERS()                                                                 \
    ExitSwitch:                                                                                \
    if (WM_DESTROY == uMsg)                                                                    \
    {                                                                                          \
        if (_pDialog)                                                                          \
        {                                                                                      \
            RemovePropW(hWnd, CLASS_INSTANCE_PROP_NAME);                                       \
            delete _pDialog;                                                                   \
        }                                                                                      \
    }                                                                                          \
    return fResult



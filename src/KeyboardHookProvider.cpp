// KeyboardHookProvider.cpp: Installs a keyboard hook and dispatches the messages to any registered handler

#include "stdafx.h"
#include "KeyboardHookProvider.h"


HHOOK hKeyboardHook = NULL;
#ifdef USE_STL
	std::vector<IKeyboardHookHandler*> keyboardHookHandlers;
#else
	CAtlArray<IKeyboardHookHandler*> keyboardHookHandlers;
#endif

KeyboardHookProvider::KeyboardHookProvider()
{
}

BOOL KeyboardHookProvider::InstallHook(IKeyboardHookHandler* pHandler)
{
	if(!hKeyboardHook) {
		hKeyboardHook = SetWindowsHookEx(WH_KEYBOARD, KeyboardHookProvider::KeyboardHookProc, NULL, GetCurrentThreadId());
		ATLASSERT(hKeyboardHook);
	}

	if(!hKeyboardHook) {
		return FALSE;
	}

	#ifdef USE_STL
		/*std::vector<IKeyboardHookHandler*>::iterator iter = std::find(keyboardHookHandlers.begin(), keyboardHookHandlers.end(), pHandler);
		if(iter == keyboardHookHandlers.end()) {*/
			keyboardHookHandlers.push_back(pHandler);
		//}
	#else
		/*BOOL found = FALSE;
		for(size_t i = 0; i < keyboardHookHandlers.GetCount() && !found; ++i) {
			if(keyboardHookHandlers[i] == pHandler) {
				found = TRUE;
			}
		}
		if(!found) {*/
			keyboardHookHandlers.Add(pHandler);
		//}
	#endif
	return TRUE;
}

void KeyboardHookProvider::RemoveHook(IKeyboardHookHandler* pHandler)
{
	#ifdef USE_STL
		std::vector<IKeyboardHookHandler*>::iterator iter = std::find(keyboardHookHandlers.begin(), keyboardHookHandlers.end(), pHandler);
		if(iter != keyboardHookHandlers.end()) {
			keyboardHookHandlers.erase(iter);
		}
		if(keyboardHookHandlers.size() == 0) {
	#else
		for(size_t i = 0; i < keyboardHookHandlers.GetCount(); ++i) {
			if(keyboardHookHandlers[i] == pHandler) {
				keyboardHookHandlers.RemoveAt(i);
				break;
			}
		}
		if(keyboardHookHandlers.GetCount() == 0) {
	#endif
		if(hKeyboardHook) {
			UnhookWindowsHookEx(hKeyboardHook);
			hKeyboardHook = NULL;
		}
	}
}


LRESULT CALLBACK KeyboardHookProvider::KeyboardHookProc(int code, WPARAM wParam, LPARAM lParam)
{
	LRESULT lr = 0;
	BOOL consumeMessage = FALSE;

	if(code >= 0) {
		#ifdef USE_STL
			for(std::vector<IKeyboardHookHandler*>::iterator iter = keyboardHookHandlers.begin(); iter != keyboardHookHandlers.end() && !consumeMessage; ++iter) {
				ATLASSERT(*iter);
				lr = (*iter)->HandleKeyboardMessage(code, wParam, lParam, consumeMessage);
			}
		#else
			for(size_t i = 0; i < keyboardHookHandlers.GetCount() && !consumeMessage; ++i) {
				ATLASSERT(keyboardHookHandlers[i]);
				lr = keyboardHookHandlers[i]->HandleKeyboardMessage(code, wParam, lParam, consumeMessage);
			}
		#endif
	}

	if(!consumeMessage) {
		lr = CallNextHookEx(hKeyboardHook, code, wParam, lParam);
	}
	return lr;
}
// MouseHookProvider.cpp: Installs a mouse hook and dispatches the messages to any registered handler

#include "stdafx.h"
#include "MouseHookProvider.h"


HHOOK hMouseHook = NULL;
#ifdef USE_STL
	std::vector<IMouseHookHandler*> mouseHookHandlers;
#else
	CAtlArray<IMouseHookHandler*> mouseHookHandlers;
#endif

MouseHookProvider::MouseHookProvider()
{
}

BOOL MouseHookProvider::InstallHook(IMouseHookHandler* pHandler)
{
	if(!hMouseHook) {
		hMouseHook = SetWindowsHookEx(WH_MOUSE, MouseHookProvider::MouseHookProc, NULL, GetCurrentThreadId());
		ATLASSERT(hMouseHook);
	}

	if(!hMouseHook) {
		return FALSE;
	}

	#ifdef USE_STL
		/*std::vector<IMouseHookHandler*>::iterator iter = std::find(mouseHookHandlers.begin(), mouseHookHandlers.end(), pHandler);
		if(iter == mouseHookHandlers.end()) {*/
			mouseHookHandlers.push_back(pHandler);
		//}
	#else
		/*BOOL found = FALSE;
		for(size_t i = 0; i < mouseHookHandlers.GetCount() && !found; ++i) {
			if(mouseHookHandlers[i] == pHandler) {
				found = TRUE;
			}
		}
		if(!found) {*/
			mouseHookHandlers.Add(pHandler);
		//}
	#endif
	return TRUE;
}

void MouseHookProvider::RemoveHook(IMouseHookHandler* pHandler)
{
	#ifdef USE_STL
		std::vector<IMouseHookHandler*>::iterator iter = std::find(mouseHookHandlers.begin(), mouseHookHandlers.end(), pHandler);
		if(iter != mouseHookHandlers.end()) {
			mouseHookHandlers.erase(iter);
		}
		if(mouseHookHandlers.size() == 0) {
	#else
		for(size_t i = 0; i < mouseHookHandlers.GetCount(); ++i) {
			if(mouseHookHandlers[i] == pHandler) {
				mouseHookHandlers.RemoveAt(i);
				break;
			}
		}
		if(mouseHookHandlers.GetCount() == 0) {
	#endif
		if(hMouseHook) {
			UnhookWindowsHookEx(hMouseHook);
			hMouseHook = NULL;
		}
	}
}


LRESULT CALLBACK MouseHookProvider::MouseHookProc(int code, WPARAM wParam, LPARAM lParam)
{
	LRESULT lr = 0;
	BOOL consumeMessage = FALSE;

	if(code >= 0) {
		#ifdef USE_STL
			for(std::vector<IMouseHookHandler*>::iterator iter = mouseHookHandlers.begin(); iter != mouseHookHandlers.end() && !consumeMessage; ++iter) {
				ATLASSERT(*iter);
				lr = (*iter)->HandleMouseMessage(code, wParam, lParam, consumeMessage);
			}
		#else
			for(size_t i = 0; i < mouseHookHandlers.GetCount() && !consumeMessage; ++i) {
				ATLASSERT(mouseHookHandlers[i]);
				lr = mouseHookHandlers[i]->HandleMouseMessage(code, wParam, lParam, consumeMessage);
			}
		#endif
	}

	if(!consumeMessage) {
		lr = CallNextHookEx(hMouseHook, code, wParam, lParam);
	}
	return lr;
}
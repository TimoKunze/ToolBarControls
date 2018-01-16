//////////////////////////////////////////////////////////////////////
/// \class KeyboardHookProvider
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Installs a keyboard hook and dispatches the messages to any registered handler</em>
///
/// This class installs a keyboard hook and dispatches the messages to any registered handler.
///
/// \sa IKeyboardHookHandler
//////////////////////////////////////////////////////////////////////

#pragma once

#include "IKeyboardHookHandler.h"


/// \brief <em>The handle of the keyboard hook</em>
extern HHOOK hKeyboardHook;
#ifdef USE_STL
	/// \brief <em>A list of all keyboard hook handlers</em>
	extern std::vector<IKeyboardHookHandler*> keyboardHookHandlers;
#else
	/// \brief <em>A list of all keyboard hook handlers</em>
	extern CAtlArray<IKeyboardHookHandler*> keyboardHookHandlers;
#endif


class KeyboardHookProvider
{
protected:
	/// \brief <em>The constructor of this class</em>
	///
	/// Used for initialization.
	KeyboardHookProvider();

public:
	/// \brief <em>Registers a keyboard hook handler</em>
	///
	/// Registers a keyboard hook handler. If it is the first handler to be registered, the keyboard hook is
	/// installed.
	///
	/// \param[in] pHandler The handler to register.
	///
	/// \return \c TRUE on success; otherwise \c FALSE.
	///
	/// \sa IKeyboardHookHandler
	static BOOL InstallHook(__in IKeyboardHookHandler* pHandler);
	/// \brief <em>Unregisters a keyboard hook handler</em>
	///
	/// Unregisters a keyboard hook handler. If it is the last handler to be unregistered, the keyboard hook
	/// is uninstalled.
	///
	/// \param[in] pHandler The handler to unregister.
	///
	/// \sa IKeyboardHookHandler
	static void RemoveHook(__in IKeyboardHookHandler* pHandler);

protected:
	/// \brief <em>A hooked keyboard message is to be processed</em>
	///
	/// This function is the callback function that we defined when we installed a keyboard hook to trap
	/// accelerator key presses for the \c ToolBar control.
	///
	/// \param[in] code A code the hook procedure uses to determine how to process the message.
	/// \param[in] wParam The virtual key code of the pressed key.
	/// \param[in] lParam Details about the key press.
	///
	/// \return The value returned by \c IKeyboardHookHandler::HandleKeyboardMessage.
	///
	/// \sa ToolBar::SendConfigurationMessages, IKeyboardHookHandler::HandleKeyboardMessage,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms644984.aspx">KeyboardProc</a>
	static LRESULT CALLBACK KeyboardHookProc(int code, WPARAM wParam, LPARAM lParam);
};     // KeyboardHookProvider
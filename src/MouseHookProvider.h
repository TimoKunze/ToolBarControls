//////////////////////////////////////////////////////////////////////
/// \class MouseHookProvider
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Installs a mouse hook and dispatches the messages to any registered handler</em>
///
/// This class installs a mouse hook and dispatches the messages to any registered handler.
///
/// \sa IMouseHookHandler
//////////////////////////////////////////////////////////////////////

#pragma once

#include "IMouseHookHandler.h"


/// \brief <em>The handle of the mouse hook</em>
extern HHOOK hMouseHook;
#ifdef USE_STL
	/// \brief <em>A list of all mouse hook handlers</em>
	extern std::vector<IMouseHookHandler*> mouseHookHandlers;
#else
	/// \brief <em>A list of all mouse hook handlers</em>
	extern CAtlArray<IMouseHookHandler*> mouseHookHandlers;
#endif


class MouseHookProvider
{
protected:
	/// \brief <em>The constructor of this class</em>
	///
	/// Used for initialization.
	MouseHookProvider();

public:
	/// \brief <em>Registers a mouse hook handler</em>
	///
	/// Registers a mouse hook handler. If it is the first handler to be registered, the mouse hook is
	/// installed.
	///
	/// \param[in] pHandler The handler to register.
	///
	/// \return \c TRUE on success; otherwise \c FALSE.
	///
	/// \sa IMouseHookHandler
	static BOOL InstallHook(__in IMouseHookHandler* pHandler);
	/// \brief <em>Unregisters a mouse hook handler</em>
	///
	/// Unregisters a mouse hook handler. If it is the last handler to be unregistered, the mouse hook is
	/// uninstalled.
	///
	/// \param[in] pHandler The handler to unregister.
	///
	/// \sa IMouseHookHandler
	static void RemoveHook(__in IMouseHookHandler* pHandler);

protected:
	/// \brief <em>A hooked mouse message is to be processed</em>
	///
	/// This function is the callback function that we defined when we installed a mouse hook to trap
	/// mouse messages that should dismiss the \c ToolBar control's chevron popup window.
	///
	/// \param[in] code A code the hook procedure uses to determine how to process the message.
	/// \param[in] wParam The identifier of the mouse message.
	/// \param[in] lParam Points to a \c MOUSEHOOKSTRUCT structure that contains more information.
	///
	/// \return The value returned by \c IMouseHookHandler::HandleMouseMessage.
	///
	/// \sa ToolBar::DisplayChevronPopupWindow, IMouseHookHandler::HandleMouseMessage,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms644968.aspx">MOUSEHOOKSTRUCT</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms644988.aspx">MouseProc</a>
	static LRESULT CALLBACK MouseHookProc(int code, WPARAM wParam, LPARAM lParam);
};     // MouseHookProvider
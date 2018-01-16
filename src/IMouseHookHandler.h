//////////////////////////////////////////////////////////////////////
/// \class IMouseHookHandler
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Communication between \c MouseHookProvider::MouseHookProc and the object that installed the mouse hook</em>
///
/// This interface allows \c MouseHookProvider::MouseHookProc to delegate the hooked mouse message to the
/// object that installed the mouse hook.
///
/// \sa ToolBar::DisplayChevronPopupWindow, MouseHookProvider::MouseHookProc
//////////////////////////////////////////////////////////////////////


#pragma once


class IMouseHookHandler
{
public:
	/// \brief <em>Processes a hooked mouse message</em>
	///
	/// This method is called by the callback function that we defined when we installed a mouse hook to trap
	/// mouse messages that should dismiss the \c ToolBar control's chevron popup window.
	///
	/// \param[in] code A code the hook procedure uses to determine how to process the message.
	/// \param[in] wParam The identifier of the mouse message.
	/// \param[in] lParam Points to a \c MOUSEHOOKSTRUCT structure that contains more information.
	/// \param[out] consumeMessage If set to \c TRUE, the message will not be passed to any further handlers
	///             and \c CallNextHookEx will not be called.
	///
	/// \return The value to return if consuming the message.
	///
	/// \sa ToolBar::HandleMouseMessage,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms644968.aspx">MOUSEHOOKSTRUCT</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms644988.aspx">MouseProc</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms644974.aspx">CallNextHookEx</a>
	virtual LRESULT HandleMouseMessage(int code, WPARAM wParam, LPARAM lParam, BOOL& consumeMessage) = 0;
};     // IMouseHookHandler
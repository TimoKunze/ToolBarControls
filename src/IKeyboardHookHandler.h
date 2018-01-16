//////////////////////////////////////////////////////////////////////
/// \class IKeyboardHookHandler
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Communication between \c KeyboardHookProvider::KeyboardHookProc and the object that installed the keyboard hook</em>
///
/// This interface allows \c KeyboardHookProvider::KeyboardHookProc to delegate the hooked keyboard message
/// to the object that installed the keyboard hook.
///
/// \sa ToolBar::SendConfigurationMessages, KeyboardHookProvider::KeyboardHookProc
//////////////////////////////////////////////////////////////////////


#pragma once


class IKeyboardHookHandler
{
public:
	/// \brief <em>Processes a hooked keyboard message</em>
	///
	/// This method is called by the callback function that we defined when we installed a keyboard hook to
	/// trap accelerator key presses for the \c ToolBar control.
	///
	/// \param[in] code A code the hook procedure uses to determine how to process the message.
	/// \param[in] wParam The virtual key code of the pressed key.
	/// \param[in] lParam Details about the key press.
	/// \param[out] consumeMessage If set to \c TRUE, the message will not be passed to any further handlers
	///             and \c CallNextHookEx will not be called.
	///
	/// \return The value to return if consuming the message.
	///
	/// \sa ToolBar::HandleKeyboardMessage,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms644984.aspx">KeyboardProc</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms644974.aspx">CallNextHookEx</a>
	virtual LRESULT HandleKeyboardMessage(int code, WPARAM wParam, LPARAM lParam, BOOL& consumeMessage) = 0;
};     // IKeyboardHookHandler
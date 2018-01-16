//////////////////////////////////////////////////////////////////////
/// \class APIWrapper
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Provides wrappers for API functions not available on all supported systems</em>
///
/// This class wraps calls to parts of the Win32 API that may be missing on the executing system.
//////////////////////////////////////////////////////////////////////


#pragma once


#ifndef DOXYGEN_SHOULD_SKIP_THIS
	typedef BOOL WINAPI CalculatePopupWindowPositionFn(__in const POINT *anchorPoint, __in const SIZE *windowSize, __in UINT flags, __in_opt RECT *excludeRect, __out RECT *popupWindowPosition);
	typedef HRESULT WINAPI SetWindowThemeFn(HWND hwnd, LPCWSTR pszSubAppName, LPCWSTR pszSubIdList);
	typedef HRESULT WINAPI SHGetImageListFn(int iImageList, REFIID riid, __out LPVOID* ppvObj);
#endif

class APIWrapper
{
private:
	/// \brief <em>The constructor of this class</em>
	///
	/// Used for initialization.
	APIWrapper(void);

public:
	/// \brief <em>Checks whether the executing system supports the \c CalculatePopupWindowPosition function</em>
	///
	/// \return \c TRUE if the function is supported; otherwise \c FALSE.
	///
	/// \sa CalculatePopupWindowPosition,
	///     <a href="https://msdn.microsoft.com/en-us/library/dd565861.aspx">CalculatePopupWindowPosition</a>
	static BOOL IsSupported_CalculatePopupWindowPosition(void);
	/// \brief <em>Checks whether the executing system supports the \c SetWindowTheme function</em>
	///
	/// \return \c TRUE if the function is supported; otherwise \c FALSE.
	///
	/// \sa SetWindowTheme,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb759827.aspx">SetWindowTheme</a>
	static BOOL IsSupported_SetWindowTheme(void);
	/// \brief <em>Checks whether the executing system supports the \c SHGetImageList function</em>
	///
	/// \return \c TRUE if the function is supported; otherwise \c FALSE.
	///
	/// \sa SHGetImageList,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb762185.aspx">SHGetImageList</a>
	static BOOL IsSupported_SHGetImageList(void);

	/// \brief <em>Calls the \c CalculatePopupWindowPosition function if it is available on the executing system</em>
	///
	/// \param[in] pAnchorPoint The \c anchorPoint parameter of the wrapped function.
	/// \param[in] pWindowSize The \c windowSize parameter of the wrapped function.
	/// \param[in] flags The \c flags parameter of the wrapped function.
	/// \param[in] pExcludeRect The \c excludeRect parameter of the wrapped function.
	/// \param[in] pPopupWindowPosition The \c popupWindowPosition parameter of the wrapped function.
	/// \param[out] pReturnValue Receives the wrapped function's return value.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa IsSupported_CalculatePopupWindowPosition,
	///     <a href="https://msdn.microsoft.com/en-us/library/dd565861.aspx">CalculatePopupWindowPosition</a>
	static HRESULT CalculatePopupWindowPosition(__in const LPPOINT pAnchorPoint, __in const LPSIZE pWindowSize, __in UINT flags, __in_opt LPRECT pExcludeRect, __out LPRECT pPopupWindowPosition, __deref_out_opt BOOL* pReturnValue);
	/// \brief <em>Calls the \c SetWindowTheme function if it is available on the executing system</em>
	///
	/// \param[in] hWnd The \c hwnd parameter of the wrapped function.
	/// \param[in] pSubAppName The \c pszSubAppName parameter of the wrapped function.
	/// \param[in] pSubIdList The \c pszSubIdList parameter of the wrapped function.
	/// \param[out] pReturnValue Receives the wrapped function's return value.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa IsSupported_SetWindowTheme,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb759827.aspx">SetWindowTheme</a>
	static HRESULT SetWindowTheme(HWND hWnd, LPCWSTR pSubAppName, LPCWSTR pSubIdList, __deref_out_opt HRESULT* pReturnValue);
	/// \brief <em>Calls the \c SHGetImageList function if it is available on the executing system</em>
	///
	/// \param[in] imageList The \c iImageList parameter of the wrapped function.
	/// \param[in] requiredInterface The \c riid parameter of the wrapped function.
	/// \param[in,out] ppInterfaceImpl The \c ppvObj parameter of the wrapped function.
	/// \param[out] pReturnValue Receives the wrapped function's return value.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa IsSupported_SHGetImageList,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb762185.aspx">SHGetImageList</a>
	static HRESULT SHGetImageList(int imageList, REFIID requiredInterface, LPVOID* ppInterfaceImpl, __deref_out_opt HRESULT* pReturnValue);

protected:
	/// \brief <em>Stores whether support for the \c CalculatePopupWindowPosition API function has been checked</em>
	static BOOL checkedSupport_CalculatePopupWindowPosition;
	/// \brief <em>Stores whether support for the \c SetWindowTheme API function has been checked</em>
	static BOOL checkedSupport_SetWindowTheme;
	/// \brief <em>Stores whether support for the \c SHGetImageList API function has been checked</em>
	static BOOL checkedSupport_SHGetImageList;
	/// \brief <em>Caches the pointer to the \c CalculatePopupWindowPosition API function</em>
	static CalculatePopupWindowPositionFn* pfnCalculatePopupWindowPosition;
	/// \brief <em>Caches the pointer to the \c SetWindowTheme API function</em>
	static SetWindowThemeFn* pfnSetWindowTheme;
	/// \brief <em>Caches the pointer to the \c SHGetImageList API function</em>
	static SHGetImageListFn* pfnSHGetImageList;
};
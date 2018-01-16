//////////////////////////////////////////////////////////////////////
/// \class ClassFactory
/// \author Timo "TimoSoft" Kunze
/// \brief <em>A helper class for creating special objects</em>
///
/// This class provides methods to create objects and initialize them with given values.
///
/// \todo Improve documentation.
///
/// \sa ReBar, ToolBar
//////////////////////////////////////////////////////////////////////


#pragma once

#include "ReBarBand.h"
#include "ReBarBands.h"
#include "ToolBarButton.h"
#include "ToolBarButtons.h"
#include "TargetOLEDataObject.h"


class ClassFactory
{
public:
	/// \brief <em>Creates an \c OLEDataObject object</em>
	///
	/// Creates an \c OLEDataObject object that wraps an object implementing \c IDataObject and returns
	/// the created object's \c IOLEDataObject implementation as a smart pointer.
	///
	/// \param[in] pDataObject The \c IDataObject implementation the \c OLEDataObject object will work
	///            on.
	///
	/// \return A smart pointer to the created object's \c IOLEDataObject implementation.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms688421.aspx">IDataObject</a>
	static CComPtr<IOLEDataObject> InitOLEDataObject(IDataObject* pDataObject)
	{
		CComPtr<IOLEDataObject> pOLEDataObj = NULL;
		InitOLEDataObject(pDataObject, IID_IOLEDataObject, reinterpret_cast<LPUNKNOWN*>(&pOLEDataObj));
		return pOLEDataObj;
	};

	/// \brief <em>Creates an \c OLEDataObject object</em>
	///
	/// \overload
	///
	/// Creates an \c OLEDataObject object that wraps an object implementing \c IDataObject and returns
	/// the created object's implementation of a given interface as a raw pointer.
	///
	/// \param[in] pDataObject The \c IDataObject implementation the \c OLEDataObject object will work
	///            on.
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppDataObject Receives the object's implementation of the interface identified by
	///             \c requiredInterface.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms688421.aspx">IDataObject</a>
	static void InitOLEDataObject(IDataObject* pDataObject, REFIID requiredInterface, LPUNKNOWN* ppDataObject)
	{
		ATLASSERT_POINTER(ppDataObject, LPUNKNOWN);
		ATLASSUME(ppDataObject);

		*ppDataObject = NULL;

		// create an OLEDataObject object and initialize it with the given parameters
		CComObject<TargetOLEDataObject>* pOLEDataObj = NULL;
		CComObject<TargetOLEDataObject>::CreateInstance(&pOLEDataObj);
		pOLEDataObj->AddRef();
		pOLEDataObj->Attach(pDataObject);
		pOLEDataObj->QueryInterface(requiredInterface, reinterpret_cast<LPVOID*>(ppDataObject));
		pOLEDataObj->Release();
	};

	/// \brief <em>Creates a \c ReBarBand object</em>
	///
	/// Creates a \c ReBarBand object that represents a given rebar band and returns its
	/// \c IReBarBand implementation as a smart pointer.
	///
	/// \param[in] bandIndex The index of the band for which to create the object.
	/// \param[in] pReBar The \c ReBar object the \c ReBarBand object will work on.
	/// \param[in] validateBandIndex If \c TRUE, the method will check whether the \c bandIndex parameter
	///            identifies an existing band; otherwise not.
	///
	/// \return A smart pointer to the created object's \c IReBarBand implementation.
	static inline CComPtr<IReBarBand> InitReBarBand(int bandIndex, ReBar* pReBar, BOOL validateBandIndex = TRUE)
	{
		CComPtr<IReBarBand> pBand = NULL;
		InitReBarBand(bandIndex, pReBar, IID_IReBarBand, reinterpret_cast<LPUNKNOWN*>(&pBand), validateBandIndex);
		return pBand;
	};

	/// \brief <em>Creates a \c ReBarBand object</em>
	///
	/// \overload
	///
	/// Creates a \c ReBarBand object that represents a given rebar band and returns its
	/// implementation of a given interface as a raw pointer.
	///
	/// \param[in] bandIndex The index of the band for which to create the object.
	/// \param[in] pReBar The \c ReBar object the \c ReBarBand object will work on.
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppBand Receives the object's implementation of the interface identified by
	///             \c requiredInterface.
	/// \param[in] validateBandIndex If \c TRUE, the method will check whether the \c bandIndex parameter
	///            identifies an existing band; otherwise not.
	static inline void InitReBarBand(int bandIndex, ReBar* pReBar, REFIID requiredInterface, LPUNKNOWN* ppBand, BOOL validateBandIndex = TRUE)
	{
		ATLASSERT_POINTER(ppBand, LPUNKNOWN);
		ATLASSUME(ppBand);

		*ppBand = NULL;
		if(validateBandIndex && !IsValidBandIndex(bandIndex, pReBar)) {
			// there's nothing useful we could return
			return;
		}

		// create a ReBarBand object and initialize it with the given parameters
		CComObject<ReBarBand>* pReBarBandObj = NULL;
		CComObject<ReBarBand>::CreateInstance(&pReBarBandObj);
		pReBarBandObj->AddRef();
		pReBarBandObj->SetOwner(pReBar);
		pReBarBandObj->Attach(bandIndex);
		pReBarBandObj->QueryInterface(requiredInterface, reinterpret_cast<LPVOID*>(ppBand));
		pReBarBandObj->Release();
	};

	/// \brief <em>Creates a \c ReBarBands object</em>
	///
	/// Creates a \c ReBarBands object that represents a collection of rebar bands and returns its
	/// \c IReBarBands implementation as a smart pointer.
	///
	/// \param[in] pReBar The \c ReBar object the \c ReBarBands object will work on.
	///
	/// \return A smart pointer to the created object's \c IReBarBands implementation.
	static inline CComPtr<IReBarBands> InitReBarBands(ReBar* pReBar)
	{
		CComPtr<IReBarBands> pBands = NULL;
		InitReBarBands(pReBar, IID_IReBarBands, reinterpret_cast<LPUNKNOWN*>(&pBands));
		return pBands;
	};

	/// \brief <em>Creates a \c ReBarBands object</em>
	///
	/// \overload
	///
	/// Creates a \c ReBarBands object that represents a collection of rebar bands and returns its
	/// implementation of a given interface as a raw pointer.
	///
	/// \param[in] pReBar The \c ReBar object the \c ReBarBands object will work on.
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppBands Receives the object's implementation of the interface identified by
	///             \c requiredInterface.
	static inline void InitReBarBands(ReBar* pReBar, REFIID requiredInterface, LPUNKNOWN* ppBands)
	{
		ATLASSERT_POINTER(ppBands, LPUNKNOWN);
		ATLASSUME(ppBands);

		*ppBands = NULL;
		// create a ReBarBands object and initialize it with the given parameters
		CComObject<ReBarBands>* pReBarBandsObj = NULL;
		CComObject<ReBarBands>::CreateInstance(&pReBarBandsObj);
		pReBarBandsObj->AddRef();

		pReBarBandsObj->SetOwner(pReBar);

		pReBarBandsObj->QueryInterface(requiredInterface, reinterpret_cast<LPVOID*>(ppBands));
		pReBarBandsObj->Release();
	};

	/// \brief <em>Creates a \c ToolBarButton object</em>
	///
	/// Creates a \c ToolBarButton object that represents a given tool bar button and returns its
	/// \c IToolBarButton implementation as a smart pointer.
	///
	/// \param[in] buttonIndex The index of the button for which to create the object.
	/// \param[in] partOfChevronPopup If \c TRUE, the button is part of the chevron popup tool bar control.
	///            In this case all methods of the created object will work on the chevron popup tool bar
	///            instead of the main tool bar.
	/// \param[in] pToolBar The \c ToolBar object the \c ToolBarButton object will work on.
	/// \param[in] validateButtonIndex If \c TRUE, the method will check whether the \c buttonIndex parameter
	///            identifies an existing button; otherwise not.
	///
	/// \return A smart pointer to the created object's \c IToolBarButton implementation.
	static inline CComPtr<IToolBarButton> InitToolBarButton(int buttonIndex, BOOL partOfChevronPopup, ToolBar* pToolBar, BOOL validateButtonIndex = TRUE)
	{
		CComPtr<IToolBarButton> pButton = NULL;
		InitToolBarButton(buttonIndex, partOfChevronPopup, pToolBar, IID_IToolBarButton, reinterpret_cast<LPUNKNOWN*>(&pButton), validateButtonIndex);
		return pButton;
	};

	/// \brief <em>Creates a \c ToolBarButton object</em>
	///
	/// \overload
	///
	/// Creates a \c ToolBarButton object that represents a given tool bar button and returns its
	/// implementation of a given interface as a raw pointer.
	///
	/// \param[in] buttonIndex The index of the button for which to create the object.
	/// \param[in] partOfChevronPopup If \c TRUE, the button is part of the chevron popup tool bar control.
	///            In this case all methods of the created object will work on the chevron popup tool bar
	///            instead of the main tool bar.
	/// \param[in] pToolBar The \c ToolBar object the \c ToolBarButton object will work on.
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppButton Receives the object's implementation of the interface identified by
	///             \c requiredInterface.
	/// \param[in] validateButtonIndex If \c TRUE, the method will check whether the \c buttonIndex parameter
	///            identifies an existing button; otherwise not.
	static inline void InitToolBarButton(int buttonIndex, BOOL partOfChevronPopup, ToolBar* pToolBar, REFIID requiredInterface, LPUNKNOWN* ppButton, BOOL validateButtonIndex = TRUE)
	{
		ATLASSERT_POINTER(ppButton, LPUNKNOWN);
		ATLASSUME(ppButton);

		*ppButton = NULL;
		if(validateButtonIndex && !IsValidButtonIndex(buttonIndex, pToolBar)) {
			// there's nothing useful we could return
			return;
		}

		// create a ToolBarButton object and initialize it with the given parameters
		CComObject<ToolBarButton>* pToolBarButtonObj = NULL;
		CComObject<ToolBarButton>::CreateInstance(&pToolBarButtonObj);
		pToolBarButtonObj->AddRef();
		pToolBarButtonObj->SetOwner(pToolBar);
		pToolBarButtonObj->Attach(buttonIndex, partOfChevronPopup);
		pToolBarButtonObj->QueryInterface(requiredInterface, reinterpret_cast<LPVOID*>(ppButton));
		pToolBarButtonObj->Release();
	};

	/// \brief <em>Creates a \c ToolBarButtons object</em>
	///
	/// Creates a \c ToolBarButtons object that represents a collection of tool bar buttons and returns its
	/// \c IToolBarButtons implementation as a smart pointer.
	///
	/// \param[in] pToolBar The \c ToolBar object the \c ToolBarButtons object will work on.
	///
	/// \return A smart pointer to the created object's \c IToolBarButtons implementation.
	static inline CComPtr<IToolBarButtons> InitToolBarButtons(ToolBar* pToolBar)
	{
		CComPtr<IToolBarButtons> pButtons = NULL;
		InitToolBarButtons(pToolBar, IID_IToolBarButtons, reinterpret_cast<LPUNKNOWN*>(&pButtons));
		return pButtons;
	};

	/// \brief <em>Creates a \c ToolBarButtons object</em>
	///
	/// \overload
	///
	/// Creates a \c ToolBarButtons object that represents a collection of tool bar buttons and returns its
	/// implementation of a given interface as a raw pointer.
	///
	/// \param[in] pToolBar The \c ToolBar object the \c ToolBarButtons object will work on.
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppButtons Receives the object's implementation of the interface identified by
	///             \c requiredInterface.
	static inline void InitToolBarButtons(ToolBar* pToolBar, REFIID requiredInterface, LPUNKNOWN* ppButtons)
	{
		ATLASSERT_POINTER(ppButtons, LPUNKNOWN);
		ATLASSUME(ppButtons);

		*ppButtons = NULL;
		// create a ToolBarButtons object and initialize it with the given parameters
		CComObject<ToolBarButtons>* pToolBarButtonsObj = NULL;
		CComObject<ToolBarButtons>::CreateInstance(&pToolBarButtonsObj);
		pToolBarButtonsObj->AddRef();

		pToolBarButtonsObj->SetOwner(pToolBar);

		pToolBarButtonsObj->QueryInterface(requiredInterface, reinterpret_cast<LPVOID*>(ppButtons));
		pToolBarButtonsObj->Release();
	};
};     // ClassFactory
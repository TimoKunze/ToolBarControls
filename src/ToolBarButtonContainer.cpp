// ToolBarButtonContainer.cpp: Manages a collection of ToolBarButton objects

#include "stdafx.h"
#include "ToolBarButtonContainer.h"
#include "ClassFactory.h"


DWORD ToolBarButtonContainer::nextID = 0;


ToolBarButtonContainer::ToolBarButtonContainer()
{
	containerID = ++nextID;
}

ToolBarButtonContainer::~ToolBarButtonContainer()
{
	properties.pOwnerTBar->DeregisterButtonContainer(containerID);
}


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP ToolBarButtonContainer::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IToolBarButtonContainer, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


//////////////////////////////////////////////////////////////////////
// implementation of IEnumVARIANT
STDMETHODIMP ToolBarButtonContainer::Clone(IEnumVARIANT** ppEnumerator)
{
	ATLASSERT_POINTER(ppEnumerator, LPENUMVARIANT);
	if(!ppEnumerator) {
		return E_POINTER;
	}

	*ppEnumerator = NULL;
	CComObject<ToolBarButtonContainer>* pTBButtonContainerObj = NULL;
	CComObject<ToolBarButtonContainer>::CreateInstance(&pTBButtonContainerObj);
	pTBButtonContainerObj->AddRef();

	// clone all settings
	pTBButtonContainerObj->properties = properties;
	if(pTBButtonContainerObj->properties.pOwnerTBar) {
		pTBButtonContainerObj->properties.pOwnerTBar->AddRef();
	}
	#ifdef USE_STL
		pTBButtonContainerObj->properties.buttons.resize(properties.buttons.size());
		std::copy(properties.buttons.begin(), properties.buttons.end(), pTBButtonContainerObj->properties.buttons.begin());
	#else
		//pTBButtonContainerObj->properties.buttons.Copy(properties.buttons);
		pTBButtonContainerObj->buttons.Copy(buttons);
	#endif

	pTBButtonContainerObj->QueryInterface(IID_IEnumVARIANT, reinterpret_cast<LPVOID*>(ppEnumerator));
	pTBButtonContainerObj->Release();

	if(*ppEnumerator) {
		properties.pOwnerTBar->RegisterButtonContainer(static_cast<IButtonContainer*>(pTBButtonContainerObj));
	}
	return S_OK;
}

STDMETHODIMP ToolBarButtonContainer::Next(ULONG numberOfMaxButtons, VARIANT* pButtons, ULONG* pNumberOfButtonsReturned)
{
	ATLASSERT_POINTER(pButtons, VARIANT);
	if(!pButtons) {
		return E_POINTER;
	}

	ULONG i = 0;
	for(i = 0; i < numberOfMaxButtons; ++i) {
		VariantInit(&pButtons[i]);

		#ifdef USE_STL
			if(properties.nextEnumeratedButton >= static_cast<int>(properties.buttons.size())) {
				properties.nextEnumeratedButton = 0;
				// there's nothing more to iterate
				break;
			}
			int buttonIndex = properties.pOwnerTBar->IDToButtonIndex(properties.buttons[properties.nextEnumeratedButton]);
		#else
			//if(properties.nextEnumeratedButton >= static_cast<int>(properties.buttons.GetCount())) {
			if(properties.nextEnumeratedButton >= static_cast<int>(buttons.GetCount())) {
				properties.nextEnumeratedButton = 0;
				// there's nothing more to iterate
				break;
			}
			//int buttonIndex = properties.pOwnerTBar->IDToButtonIndex(properties.buttons[properties.nextEnumeratedButton]);
			int buttonIndex = properties.pOwnerTBar->IDToButtonIndex(buttons[properties.nextEnumeratedButton]);
		#endif

		ClassFactory::InitToolBarButton(buttonIndex, FALSE, properties.pOwnerTBar, IID_IDispatch, reinterpret_cast<LPUNKNOWN*>(&pButtons[i].pdispVal));
		pButtons[i].vt = VT_DISPATCH;
		++properties.nextEnumeratedButton;
	}
	if(pNumberOfButtonsReturned) {
		*pNumberOfButtonsReturned = i;
	}

	return (i == numberOfMaxButtons ? S_OK : S_FALSE);
}

STDMETHODIMP ToolBarButtonContainer::Reset(void)
{
	properties.nextEnumeratedButton = 0;
	return S_OK;
}

STDMETHODIMP ToolBarButtonContainer::Skip(ULONG numberOfButtonsToSkip)
{
	properties.nextEnumeratedButton += numberOfButtonsToSkip;
	return S_OK;
}
// implementation of IEnumVARIANT
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IButtonContainer
void ToolBarButtonContainer::RemovedButton(LONG buttonID)
{
	if(buttonID == -1) {
		RemoveAll();
	} else {
		Remove(buttonID);
	}
}

DWORD ToolBarButtonContainer::GetID(void)
{
	return containerID;
}
// implementation of IButtonContainer
//////////////////////////////////////////////////////////////////////


ToolBarButtonContainer::Properties::~Properties()
{
	if(pOwnerTBar) {
		pOwnerTBar->Release();
	}
}

HWND ToolBarButtonContainer::Properties::GetTBarHWnd(void)
{
	ATLASSUME(pOwnerTBar);

	OLE_HANDLE handle = NULL;
	pOwnerTBar->get_hWnd(&handle);
	return static_cast<HWND>(LongToHandle(handle));
}


void ToolBarButtonContainer::SetOwner(ToolBar* pOwner)
{
	if(properties.pOwnerTBar) {
		properties.pOwnerTBar->Release();
	}
	properties.pOwnerTBar = pOwner;
	if(properties.pOwnerTBar) {
		properties.pOwnerTBar->AddRef();
	}
}


STDMETHODIMP ToolBarButtonContainer::get_Item(LONG buttonID, IToolBarButton** ppButton)
{
	ATLASSERT_POINTER(ppButton, IToolBarButton*);
	if(!ppButton) {
		return E_POINTER;
	}

	#ifdef USE_STL
		std::vector<LONG>::iterator iter = std::find(properties.buttons.begin(), properties.buttons.end(), buttonID);
		if(iter != properties.buttons.end()) {
			int buttonIndex = properties.pOwnerTBar->IDToButtonIndex(buttonID);
			if(buttonIndex != -1) {
				ClassFactory::InitToolBarButton(buttonIndex, FALSE, properties.pOwnerTBar, IID_IToolBarButton, reinterpret_cast<LPUNKNOWN*>(ppButton));
				return S_OK;
			}
		}
	#else
		//for(size_t i = 0; i < properties.buttons.GetCount(); ++i) {
		for(size_t i = 0; i < buttons.GetCount(); ++i) {
			//if(properties.buttons[i] == buttonID) {
			if(buttons[i] == buttonID) {
				int buttonIndex = properties.pOwnerTBar->IDToButtonIndex(buttonID);
				if(buttonIndex != -1) {
					ClassFactory::InitToolBarButton(buttonIndex, FALSE, properties.pOwnerTBar, IID_IToolBarButton, reinterpret_cast<LPUNKNOWN*>(ppButton));
					return S_OK;
				}
				break;
			}
		}
	#endif

	return E_INVALIDARG;
}

STDMETHODIMP ToolBarButtonContainer::get__NewEnum(IUnknown** ppEnumerator)
{
	ATLASSERT_POINTER(ppEnumerator, LPUNKNOWN);
	if(!ppEnumerator) {
		return E_POINTER;
	}

	Reset();
	return QueryInterface(IID_IUnknown, reinterpret_cast<LPVOID*>(ppEnumerator));
}


STDMETHODIMP ToolBarButtonContainer::Add(VARIANT buttons)
{
	HRESULT hr = E_FAIL;
	LONG id = 0;
	switch(buttons.vt) {
		case VT_DISPATCH:
			if(buttons.pdispVal) {
				CComQIPtr<IToolBarButton, &IID_IToolBarButton> pTBButton(buttons.pdispVal);
				if(pTBButton) {
					// add a single ToolBarButton object
					hr = pTBButton->get_ID(&id);
				} else {
					CComQIPtr<IToolBarButtons, &IID_IToolBarButtons> pTBButtons(buttons.pdispVal);
					if(pTBButtons) {
						// add a ToolBarButtons collection
						CComQIPtr<IEnumVARIANT, &IID_IEnumVARIANT> pEnumerator(pTBButtons);
						if(pEnumerator) {
							hr = S_OK;
							VARIANT button;
							VariantInit(&button);
							ULONG dummy = 0;
							while(pEnumerator->Next(1, &button, &dummy) == S_OK) {
								if((button.vt == VT_DISPATCH) && button.pdispVal) {
									pTBButton = button.pdispVal;
									if(pTBButton) {
										id = 0;
										pTBButton->get_ID(&id);
										#ifdef USE_STL
											std::vector<LONG>::iterator iter = std::find(properties.buttons.begin(), properties.buttons.end(), id);
											if(iter == properties.buttons.end()) {
												properties.buttons.push_back(id);
											}
										#else
											BOOL alreadyThere = FALSE;
											//for(size_t i = 0; i < properties.buttons.GetCount(); ++i) {
											for(size_t i = 0; i < this->buttons.GetCount(); ++i) {
												//if(properties.buttons[i] == id) {
												if(this->buttons[i] == id) {
													alreadyThere = TRUE;
													break;
												}
											}
											if(!alreadyThere) {
												//properties.buttons.Add(id);
												this->buttons.Add(id);
											}
										#endif
									}
								}
								VariantClear(&button);
							}
							return S_OK;
						}
					}
				}
			}
			break;
		default:
			VARIANT v;
			VariantInit(&v);
			hr = VariantChangeType(&v, &buttons, 0, VT_UI4);
			id = v.ulVal;
			break;
	}
	if(FAILED(hr)) {
		// invalid arg - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	#ifdef USE_STL
		std::vector<LONG>::iterator iter = std::find(properties.buttons.begin(), properties.buttons.end(), id);
		if(iter == properties.buttons.end()) {
			properties.buttons.push_back(id);
		}
	#else
		BOOL alreadyThere = FALSE;
		//for(size_t i = 0; i < properties.buttons.GetCount(); ++i) {
		for(size_t i = 0; i < this->buttons.GetCount(); ++i) {
			//if(properties.buttons[i] == id) {
			if(this->buttons[i] == id) {
				alreadyThere = TRUE;
				break;
			}
		}
		if(!alreadyThere) {
			//properties.buttons.Add(id);
			this->buttons.Add(id);
		}
	#endif
	return S_OK;
}

STDMETHODIMP ToolBarButtonContainer::Clone(IToolBarButtonContainer** ppClone)
{
	ATLASSERT_POINTER(ppClone, IToolBarButtonContainer*);
	if(!ppClone) {
		return E_POINTER;
	}

	*ppClone = NULL;
	CComObject<ToolBarButtonContainer>* pTBButtonContainerObj = NULL;
	CComObject<ToolBarButtonContainer>::CreateInstance(&pTBButtonContainerObj);
	pTBButtonContainerObj->AddRef();

	// clone all settings
	pTBButtonContainerObj->properties = properties;
	if(pTBButtonContainerObj->properties.pOwnerTBar) {
		pTBButtonContainerObj->properties.pOwnerTBar->AddRef();
	}
	#ifdef USE_STL
		pTBButtonContainerObj->properties.buttons.resize(properties.buttons.size());
		std::copy(properties.buttons.begin(), properties.buttons.end(), pTBButtonContainerObj->properties.buttons.begin());
	#else
		//pTBButtonContainerObj->properties.buttons.Copy(properties.buttons);
		pTBButtonContainerObj->buttons.Copy(buttons);
	#endif

	pTBButtonContainerObj->QueryInterface(__uuidof(IToolBarButtonContainer), reinterpret_cast<LPVOID*>(ppClone));
	pTBButtonContainerObj->Release();

	if(*ppClone) {
		properties.pOwnerTBar->RegisterButtonContainer(static_cast<IButtonContainer*>(pTBButtonContainerObj));
	}
	return S_OK;
}

STDMETHODIMP ToolBarButtonContainer::Count(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	#ifdef USE_STL
		*pValue = static_cast<LONG>(properties.buttons.size());
	#else
		//*pValue = static_cast<LONG>(properties.buttons.GetCount());
		*pValue = static_cast<LONG>(buttons.GetCount());
	#endif
	return S_OK;
}

STDMETHODIMP ToolBarButtonContainer::CreateDragImage(OLE_XPOS_PIXELS* pXUpperLeft/* = NULL*/, OLE_YPOS_PIXELS* pYUpperLeft/* = NULL*/, OLE_HANDLE* phImageList/* = NULL*/)
{
	ATLASSERT_POINTER(phImageList, OLE_HANDLE);
	if(!phImageList) {
		return E_POINTER;
	}

	*phImageList = NULL;
	#ifdef USE_STL
		switch(properties.buttons.size()) {
	#else
		//switch(properties.buttons.GetCount()) {
		switch(buttons.GetCount()) {
	#endif
		case 0:
			return S_OK;
			break;
		case 1: {
			ATLASSUME(properties.pOwnerTBar);
			//int buttonIndex = properties.pOwnerTBar->IDToButtonIndex(properties.buttons[0]);
			#ifdef USE_STL
				int buttonIndex = properties.pOwnerTBar->IDToButtonIndex(properties.buttons[0]);
			#else
				int buttonIndex = properties.pOwnerTBar->IDToButtonIndex(buttons[0]);
			#endif
			if(buttonIndex != -1) {
				POINT upperLeftPoint = {0};
				*phImageList = HandleToLong(properties.pOwnerTBar->CreateLegacyDragImage(buttonIndex, &upperLeftPoint, NULL));
				if(*phImageList) {
					if(pXUpperLeft) {
						*pXUpperLeft = upperLeftPoint.x;
					}
					if(pYUpperLeft) {
						*pYUpperLeft = upperLeftPoint.y;
					}
					return S_OK;
				}
			}
			break;
		}
		default: {
			// create a large drag image out of small drag images
			ATLASSUME(properties.pOwnerTBar);

			BOOL use32BPPImage = RunTimeHelper::IsCommCtrl6();

			// calculate the bitmap's required size and collect each button's imagelist
			#ifdef USE_STL
				std::vector<HIMAGELIST> imageLists;
				std::vector<RECT> buttonBoundingRects;
			#else
				CAtlArray<HIMAGELIST> imageLists;
				CAtlArray<RECT> buttonBoundingRects;
			#endif
			POINT upperLeftPoint = {0};
			WTL::CRect boundingRect;
			#ifdef USE_STL
				for(std::vector<LONG>::iterator iter = properties.buttons.begin(); iter != properties.buttons.end(); ++iter) {
					int buttonIndex = properties.pOwnerTBar->IDToButtonIndex(*iter);
			#else
				//for(size_t i = 0; i < properties.buttons.GetCount(); ++i) {
				for(size_t i = 0; i < buttons.GetCount(); ++i) {
					//int buttonIndex = properties.pOwnerTBar->IDToButtonIndex(properties.buttons[i]);
					int buttonIndex = properties.pOwnerTBar->IDToButtonIndex(buttons[i]);
			#endif
				if(buttonIndex != -1) {
					// NOTE: Windows skips items outside the client area to improve performance. We don't.
					POINT pt = {0};
					RECT buttonBoundingRect = {0};
					HIMAGELIST hImageList = properties.pOwnerTBar->CreateLegacyDragImage(buttonIndex, &pt, &buttonBoundingRect);
					boundingRect.UnionRect(&boundingRect, &buttonBoundingRect);

					#ifdef USE_STL
						if(imageLists.size() == 0) {
					#else
						if(imageLists.GetCount() == 0) {
					#endif
						upperLeftPoint = pt;
					} else {
						upperLeftPoint.x = min(upperLeftPoint.x, pt.x);
						upperLeftPoint.y = min(upperLeftPoint.y, pt.y);
					}
					#ifdef USE_STL
						imageLists.push_back(hImageList);
						buttonBoundingRects.push_back(boundingRect);
					#else
						imageLists.Add(hImageList);
						buttonBoundingRects.Add(boundingRect);
					#endif
				}
			}
			WTL::CRect dragImageRect(0, 0, boundingRect.Width(), boundingRect.Height());

			// setup the DCs we'll draw into
			HDC hCompatibleDC = GetDC(HWND_DESKTOP);
			CDC memoryDC;
			memoryDC.CreateCompatibleDC(hCompatibleDC);
			CDC maskMemoryDC;
			maskMemoryDC.CreateCompatibleDC(hCompatibleDC);

			// create the bitmap and its mask
			CBitmap dragImage;
			LPRGBQUAD pDragImageBits = NULL;
			if(use32BPPImage) {
				BITMAPINFO bitmapInfo = {0};
				bitmapInfo.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
				bitmapInfo.bmiHeader.biWidth = dragImageRect.Width();
				bitmapInfo.bmiHeader.biHeight = -dragImageRect.Height();
				bitmapInfo.bmiHeader.biPlanes = 1;
				bitmapInfo.bmiHeader.biBitCount = 32;
				bitmapInfo.bmiHeader.biCompression = BI_RGB;

				dragImage.CreateDIBSection(NULL, &bitmapInfo, DIB_RGB_COLORS, reinterpret_cast<LPVOID*>(&pDragImageBits), NULL, 0);
			} else {
				dragImage.CreateCompatibleBitmap(hCompatibleDC, dragImageRect.Width(), dragImageRect.Height());
			}
			HBITMAP hPreviousBitmap = memoryDC.SelectBitmap(dragImage);
			memoryDC.FillRect(&dragImageRect, static_cast<HBRUSH>(GetStockObject(WHITE_BRUSH)));
			CBitmap dragImageMask;
			dragImageMask.CreateBitmap(dragImageRect.Width(), dragImageRect.Height(), 1, 1, NULL);
			HBITMAP hPreviousBitmapMask = maskMemoryDC.SelectBitmap(dragImageMask);
			maskMemoryDC.FillRect(&dragImageRect, static_cast<HBRUSH>(GetStockObject(WHITE_BRUSH)));

			// draw each single drag image into our bitmap
			BOOL rightToLeft = FALSE;
			HWND hWndTBar = properties.GetTBarHWnd();
			if(IsWindow(hWndTBar)) {
				rightToLeft = ((CWindow(hWndTBar).GetExStyle() & WS_EX_LAYOUTRTL) == WS_EX_LAYOUTRTL);
			}
			#ifdef USE_STL
				for(size_t i = 0; i < imageLists.size(); ++i) {
			#else
				for(size_t i = 0; i < imageLists.GetCount(); ++i) {
			#endif
				if(rightToLeft) {
					ImageList_Draw(imageLists[i], 0, memoryDC, boundingRect.right - buttonBoundingRects[i].right, buttonBoundingRects[i].top - boundingRect.top, ILD_NORMAL);
					ImageList_Draw(imageLists[i], 0, maskMemoryDC, boundingRect.right - buttonBoundingRects[i].right, buttonBoundingRects[i].top - boundingRect.top, ILD_MASK);
				} else {
					ImageList_Draw(imageLists[i], 0, memoryDC, buttonBoundingRects[i].left - boundingRect.left, buttonBoundingRects[i].top - boundingRect.top, ILD_NORMAL);
					ImageList_Draw(imageLists[i], 0, maskMemoryDC, buttonBoundingRects[i].left - boundingRect.left, buttonBoundingRects[i].top - boundingRect.top, ILD_MASK);
				}

				ImageList_Destroy(imageLists[i]);
			}

			// clean up
			#ifdef USE_STL
				imageLists.clear();
				buttonBoundingRects.clear();
			#else
				imageLists.RemoveAll();
				buttonBoundingRects.RemoveAll();
			#endif
			memoryDC.SelectBitmap(hPreviousBitmap);
			maskMemoryDC.SelectBitmap(hPreviousBitmapMask);
			ReleaseDC(HWND_DESKTOP, hCompatibleDC);

			// create the imagelist
			HIMAGELIST hImageList = ImageList_Create(dragImageRect.Width(), dragImageRect.Height(), (RunTimeHelper::IsCommCtrl6() ? ILC_COLOR32 : ILC_COLOR24) | ILC_MASK, 1, 0);
			ImageList_SetBkColor(hImageList, CLR_NONE);
			ImageList_Add(hImageList, dragImage, dragImageMask);
			*phImageList = HandleToLong(hImageList);

			if(*phImageList) {
				if(pXUpperLeft) {
					*pXUpperLeft = upperLeftPoint.x;
				}
				if(pYUpperLeft) {
					*pYUpperLeft = upperLeftPoint.y;
				}
				return S_OK;
			}
			break;
		}
	}

	return E_FAIL;
}

STDMETHODIMP ToolBarButtonContainer::Remove(LONG buttonID, VARIANT_BOOL removePhysically/* = VARIANT_FALSE*/)
{
	#ifdef USE_STL
		for(size_t i = 0; i < properties.buttons.size(); ++i) {
			if(properties.buttons[i] == buttonID) {
				if(removePhysically == VARIANT_FALSE) {
					properties.buttons.erase(std::find(properties.buttons.begin(), properties.buttons.end(), buttonID));
					if(i < static_cast<size_t>(properties.nextEnumeratedButton)) {
						--properties.nextEnumeratedButton;
					}
				} else {
					HWND hWndTBar = properties.GetTBarHWnd();
					if(IsWindow(hWndTBar)) {
						int buttonIndex = properties.pOwnerTBar->IDToButtonIndex(buttonID);
						SendMessage(hWndTBar, TB_DELETEBUTTON, buttonIndex, 0);
					}

					// our owner will notify us about the deletion, so we don't need to remove the button explicitly
				}

				return S_OK;
			}
		}
	#else
		//for(size_t i = 0; i < properties.buttons.GetCount(); ++i) {
		for(size_t i = 0; i < buttons.GetCount(); ++i) {
			//if(properties.buttons[i] == buttonID) {
			if(buttons[i] == buttonID) {
				if(removePhysically == VARIANT_FALSE) {
					//properties.buttons.RemoveAt(i);
					buttons.RemoveAt(i);
					if(i < static_cast<size_t>(properties.nextEnumeratedButton)) {
						--properties.nextEnumeratedButton;
					}
				} else {
					HWND hWndTBar = properties.GetTBarHWnd();
					if(IsWindow(hWndTBar)) {
						int buttonIndex = properties.pOwnerTBar->IDToButtonIndex(buttonID);
						SendMessage(hWndTBar, TB_DELETEBUTTON, buttonIndex, 0);
					}

					// our owner will notify us about the deletion, so we don't need to remove the button explicitly
				}

				return S_OK;
			}
		}
	#endif
	return E_FAIL;
}

STDMETHODIMP ToolBarButtonContainer::RemoveAll(VARIANT_BOOL removePhysically/* = VARIANT_FALSE*/)
{
	if(removePhysically != VARIANT_FALSE) {
		HWND hWndTBar = properties.GetTBarHWnd();
		if(IsWindow(hWndTBar)) {
			#ifdef USE_STL
				while(properties.buttons.size() > 0) {
					int buttonIndex = properties.pOwnerTBar->IDToButtonIndex(*properties.buttons.begin());
					ATLASSERT(buttonIndex != -1);
					SendMessage(hWndTBar, TB_DELETEBUTTON, buttonIndex, 0);
				}
			#else
				//while(properties.buttons.GetCount() > 0) {
				while(buttons.GetCount() > 0) {
					//int buttonIndex = properties.pOwnerTBar->IDToButtonIndex(properties.buttons[0]);
					int buttonIndex = properties.pOwnerTBar->IDToButtonIndex(buttons[0]);
					ATLASSERT(buttonIndex != -1);
					SendMessage(hWndTBar, TB_DELETEBUTTON, buttonIndex, 0);
				}
			#endif
		}
	} else {
		#ifdef USE_STL
			properties.buttons.clear();
		#else
			//properties.buttons.RemoveAll();
			buttons.RemoveAll();
		#endif
	}
	Reset();
	return S_OK;
}
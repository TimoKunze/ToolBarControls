// CommonProperties.cpp: The controls' "Common" property page

#include "stdafx.h"
#include "CommonProperties.h"


CommonProperties::CommonProperties()
{
	m_dwTitleID = IDS_TITLECOMMONPROPERTIES;
	m_dwDocStringID = IDS_DOCSTRINGCOMMONPROPERTIES;
}


//////////////////////////////////////////////////////////////////////
// implementation of IPropertyPage
STDMETHODIMP CommonProperties::Apply(void)
{
	ApplySettings();
	return S_OK;
}

STDMETHODIMP CommonProperties::Activate(HWND hWndParent, LPCRECT pRect, BOOL modal)
{
	IPropertyPage2Impl<CommonProperties>::Activate(hWndParent, pRect, modal);

	// attach to the controls
	controls.disabledEventsList.SubclassWindow(GetDlgItem(IDC_DISABLEDEVENTSBOX));
	HIMAGELIST hStateImageList = SetupStateImageList(controls.disabledEventsList.GetImageList(LVSIL_STATE));
	controls.disabledEventsList.SetImageList(hStateImageList, LVSIL_STATE);
	controls.disabledEventsList.SetExtendedListViewStyle(LVS_EX_DOUBLEBUFFER | LVS_EX_INFOTIP, LVS_EX_DOUBLEBUFFER | LVS_EX_INFOTIP | LVS_EX_FULLROWSELECT);
	controls.disabledEventsList.AddColumn(TEXT(""), 0);
	controls.disabledEventsList.GetToolTips().SetTitle(TTI_INFO, TEXT("Affected events"));

	// setup the toolbar
	WTL::CRect toolbarRect;
	GetClientRect(&toolbarRect);
	toolbarRect.OffsetRect(0, 2);
	toolbarRect.left += toolbarRect.right - 46;
	toolbarRect.bottom = toolbarRect.top + 22;
	controls.toolbar.Create(*this, toolbarRect, NULL, WS_CHILDWINDOW | WS_VISIBLE | TBSTYLE_TRANSPARENT | TBSTYLE_FLAT | TBSTYLE_TOOLTIPS | CCS_NODIVIDER | CCS_NOPARENTALIGN | CCS_NORESIZE, 0);
	controls.toolbar.SetButtonStructSize();
	controls.imagelistEnabled.CreateFromImage(IDB_TOOLBARENABLED, 16, 0, RGB(255, 0, 255), IMAGE_BITMAP, LR_CREATEDIBSECTION);
	controls.toolbar.SetImageList(controls.imagelistEnabled);
	controls.imagelistDisabled.CreateFromImage(IDB_TOOLBARDISABLED, 16, 0, RGB(255, 0, 255), IMAGE_BITMAP, LR_CREATEDIBSECTION);
	controls.toolbar.SetDisabledImageList(controls.imagelistDisabled);

	// insert the buttons
	TBBUTTON buttons[2];
	ZeroMemory(buttons, sizeof(buttons));
	buttons[0].iBitmap = 0;
	buttons[0].idCommand = ID_LOADSETTINGS;
	buttons[0].fsState = TBSTATE_ENABLED;
	buttons[0].fsStyle = BTNS_BUTTON;
	buttons[1].iBitmap = 1;
	buttons[1].idCommand = ID_SAVESETTINGS;
	buttons[1].fsStyle = BTNS_BUTTON;
	buttons[1].fsState = TBSTATE_ENABLED;
	controls.toolbar.AddButtons(2, buttons);

	LoadSettings();
	return S_OK;
}

STDMETHODIMP CommonProperties::SetObjects(ULONG objects, IUnknown** ppControls)
{
	if(m_bDirty) {
		Apply();
	}
	IPropertyPage2Impl<CommonProperties>::SetObjects(objects, ppControls);
	LoadSettings();
	return S_OK;
}
// implementation of IPropertyPage
//////////////////////////////////////////////////////////////////////


LRESULT CommonProperties::OnListViewGetInfoTipNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMLVGETINFOTIP pDetails = reinterpret_cast<LPNMLVGETINFOTIP>(pNotificationDetails);
	LPTSTR pBuffer = new TCHAR[pDetails->cchTextMax + 1];

	if(pNotificationDetails->hwndFrom == controls.disabledEventsList) {
		int numberOfReBars = 0;
		int numberOfToolBars = 0;
		for(UINT object = 0; object < m_nObjects; ++object) {
			LPUNKNOWN pControl = NULL;
			if(m_ppUnk[object]->QueryInterface(IID_IReBar, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
				++numberOfReBars;
				pControl->Release();
			} else if(m_ppUnk[object]->QueryInterface(IID_IToolBar, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
				++numberOfToolBars;
				pControl->Release();
			}
		}

		if(pDetails->iItem == properties.disabledEventsItemIndices.deMouseEvents) {
			if(numberOfReBars > 0 && numberOfToolBars > 0) {
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("MouseDown, MouseUp, MouseEnter, MouseHover, MouseLeave, BandMouseEnter, BandMouseLeave, ButtonMouseEnter, ButtonMouseLeave, MouseMove, NonClientHitTest"))));
			} else if(numberOfToolBars == 0) {
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("MouseDown, MouseUp, MouseEnter, MouseHover, MouseLeave, BandMouseEnter, BandMouseLeave, MouseMove, NonClientHitTest"))));
			} else {
				ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("MouseDown, MouseUp, MouseEnter, MouseHover, MouseLeave, ButtonMouseEnter, ButtonMouseLeave, MouseMove"))));
			}
		} else if(pDetails->iItem == properties.disabledEventsItemIndices.deClickEvents) {
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("Click, DblClick, MClick, MDblClick, RClick, RDblClick, XClick, XDblClick"))));
		} else if(pDetails->iItem == properties.disabledEventsItemIndices.deKeyboardEvents) {
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("KeyDown, KeyUp, KeyPress"))));
		} else if(pDetails->iItem == properties.disabledEventsItemIndices.deBandInsertionEvents) {
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("InsertingBand, InsertedBand"))));
		} else if(pDetails->iItem == properties.disabledEventsItemIndices.deButtonInsertionEvents) {
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("InsertingButton, InsertedButton"))));
		} else if(pDetails->iItem == properties.disabledEventsItemIndices.deBandDeletionEvents) {
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("RemovingBand, RemovedBand"))));
		} else if(pDetails->iItem == properties.disabledEventsItemIndices.deButtonDeletionEvents) {
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("RemovingButton, RemovedButton"))));
		} else if(pDetails->iItem == properties.disabledEventsItemIndices.deFreeBandData) {
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("FreeBandData"))));
		} else if(pDetails->iItem == properties.disabledEventsItemIndices.deFreeButtonData) {
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("FreeButtonData"))));
		} else if(pDetails->iItem == properties.disabledEventsItemIndices.deCustomDraw) {
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("CustomDraw"))));
		} else if(pDetails->iItem == properties.disabledEventsItemIndices.deHotButtonChangeEvents) {
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("HotButtonChangeWrapping, HotButtonChanging"))));
		} else if(pDetails->iItem == properties.disabledEventsItemIndices.deAcceleratorEvents) {
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("IsDuplicateAccelerator, MapAccelerator"))));
		} else if(pDetails->iItem == properties.disabledEventsItemIndices.deRawMenuMessage) {
			ATLVERIFY(SUCCEEDED(StringCchCopy(pBuffer, pDetails->cchTextMax + 1, TEXT("RawMenuMessage"))));
		}
		ATLVERIFY(SUCCEEDED(StringCchCopy(pDetails->pszText, pDetails->cchTextMax, pBuffer)));
	}

	delete[] pBuffer;
	return 0;
}

LRESULT CommonProperties::OnListViewItemChangedNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMLISTVIEW pDetails = reinterpret_cast<LPNMLISTVIEW>(pNotificationDetails);
	if(pDetails->uChanged & LVIF_STATE) {
		if((pDetails->uNewState & LVIS_STATEIMAGEMASK) != (pDetails->uOldState & LVIS_STATEIMAGEMASK)) {
			if((pDetails->uNewState & LVIS_STATEIMAGEMASK) >> 12 == 3) {
				if(pNotificationDetails->hwndFrom != properties.hWndCheckMarksAreSetFor) {
					LVITEM item = {0};
					item.state = INDEXTOSTATEIMAGEMASK(1);
					item.stateMask = LVIS_STATEIMAGEMASK;
					::SendMessage(pNotificationDetails->hwndFrom, LVM_SETITEMSTATE, pDetails->iItem, reinterpret_cast<LPARAM>(&item));
				}
			}
			SetDirty(TRUE);
		}
	}
	return 0;
}

LRESULT CommonProperties::OnToolTipGetDispInfoNotificationA(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMTTDISPINFOA pDetails = reinterpret_cast<LPNMTTDISPINFOA>(pNotificationDetails);
	pDetails->hinst = ModuleHelper::GetResourceInstance();
	switch(pDetails->hdr.idFrom) {
		case ID_LOADSETTINGS:
			pDetails->lpszText = MAKEINTRESOURCEA(IDS_LOADSETTINGS);
			break;
		case ID_SAVESETTINGS:
			pDetails->lpszText = MAKEINTRESOURCEA(IDS_SAVESETTINGS);
			break;
	}
	return 0;
}

LRESULT CommonProperties::OnToolTipGetDispInfoNotificationW(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMTTDISPINFOW pDetails = reinterpret_cast<LPNMTTDISPINFOW>(pNotificationDetails);
	pDetails->hinst = ModuleHelper::GetResourceInstance();
	switch(pDetails->hdr.idFrom) {
		case ID_LOADSETTINGS:
			pDetails->lpszText = MAKEINTRESOURCEW(IDS_LOADSETTINGS);
			break;
		case ID_SAVESETTINGS:
			pDetails->lpszText = MAKEINTRESOURCEW(IDS_SAVESETTINGS);
			break;
	}
	return 0;
}

LRESULT CommonProperties::OnLoadSettingsFromFile(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& /*wasHandled*/)
{
	ATLASSERT(m_nObjects == 1);

	IUnknown* pControl = NULL;
	if(m_ppUnk[0]->QueryInterface(IID_IReBar, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
		CFileDialog dlg(TRUE, NULL, NULL, OFN_ENABLESIZING | OFN_EXPLORER | OFN_FILEMUSTEXIST | OFN_HIDEREADONLY | OFN_LONGNAMES | OFN_PATHMUSTEXIST, TEXT("All files\0*.*\0\0"), *this);
		if(dlg.DoModal() == IDOK) {
			CComBSTR file = dlg.m_szFileName;

			VARIANT_BOOL b = VARIANT_FALSE;
			reinterpret_cast<IReBar*>(pControl)->LoadSettingsFromFile(file, &b);
			if(b == VARIANT_FALSE) {
				MessageBox(TEXT("The specified file could not be loaded."), TEXT("Error!"), MB_ICONERROR | MB_OK);
			}
		}
		pControl->Release();

	} else if(m_ppUnk[0]->QueryInterface(IID_IToolBar, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
		CFileDialog dlg(TRUE, NULL, NULL, OFN_ENABLESIZING | OFN_EXPLORER | OFN_FILEMUSTEXIST | OFN_HIDEREADONLY | OFN_LONGNAMES | OFN_PATHMUSTEXIST, TEXT("All files\0*.*\0\0"), *this);
		if(dlg.DoModal() == IDOK) {
			CComBSTR file = dlg.m_szFileName;

			VARIANT_BOOL b = VARIANT_FALSE;
			reinterpret_cast<IToolBar*>(pControl)->LoadSettingsFromFile(file, &b);
			if(b == VARIANT_FALSE) {
				MessageBox(TEXT("The specified file could not be loaded."), TEXT("Error!"), MB_ICONERROR | MB_OK);
			}
		}
		pControl->Release();
	}
	return 0;
}

LRESULT CommonProperties::OnSaveSettingsToFile(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& /*wasHandled*/)
{
	ATLASSERT(m_nObjects == 1);

	IUnknown* pControl = NULL;
	if(m_ppUnk[0]->QueryInterface(IID_IReBar, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
		CFileDialog dlg(FALSE, NULL, TEXT("ReBar Settings.dat"), OFN_ENABLESIZING | OFN_EXPLORER | OFN_CREATEPROMPT | OFN_HIDEREADONLY | OFN_LONGNAMES | OFN_PATHMUSTEXIST | OFN_OVERWRITEPROMPT, TEXT("All files\0*.*\0\0"), *this);
		if(dlg.DoModal() == IDOK) {
			CComBSTR file = dlg.m_szFileName;

			VARIANT_BOOL b = VARIANT_FALSE;
			reinterpret_cast<IReBar*>(pControl)->SaveSettingsToFile(file, &b);
			if(b == VARIANT_FALSE) {
				MessageBox(TEXT("The specified file could not be written."), TEXT("Error!"), MB_ICONERROR | MB_OK);
			}
		}
		pControl->Release();

	} else if(m_ppUnk[0]->QueryInterface(IID_IToolBar, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
		CFileDialog dlg(FALSE, NULL, TEXT("ToolBar Settings.dat"), OFN_ENABLESIZING | OFN_EXPLORER | OFN_CREATEPROMPT | OFN_HIDEREADONLY | OFN_LONGNAMES | OFN_PATHMUSTEXIST | OFN_OVERWRITEPROMPT, TEXT("All files\0*.*\0\0"), *this);
		if(dlg.DoModal() == IDOK) {
			CComBSTR file = dlg.m_szFileName;

			VARIANT_BOOL b = VARIANT_FALSE;
			reinterpret_cast<IToolBar*>(pControl)->SaveSettingsToFile(file, &b);
			if(b == VARIANT_FALSE) {
				MessageBox(TEXT("The specified file could not be written."), TEXT("Error!"), MB_ICONERROR | MB_OK);
			}
		}
		pControl->Release();
	}
	return 0;
}


void CommonProperties::ApplySettings(void)
{
	for(UINT object = 0; object < m_nObjects; ++object) {
		LPUNKNOWN pControl = NULL;
		if(m_ppUnk[object]->QueryInterface(IID_IReBar, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
			DisabledEventsConstants disabledEvents = static_cast<DisabledEventsConstants>(0);
			reinterpret_cast<IReBar*>(pControl)->get_DisabledEvents(&disabledEvents);
			LONG l = static_cast<LONG>(disabledEvents);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deMouseEvents, LVIS_STATEIMAGEMASK), l, deMouseEvents);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deClickEvents, LVIS_STATEIMAGEMASK), l, deClickEvents);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deKeyboardEvents, LVIS_STATEIMAGEMASK), l, deKeyboardEvents);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deBandInsertionEvents, LVIS_STATEIMAGEMASK), l, deBandInsertionEvents);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deBandDeletionEvents, LVIS_STATEIMAGEMASK), l, deBandDeletionEvents);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deFreeBandData, LVIS_STATEIMAGEMASK), l, deFreeBandData);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deCustomDraw, LVIS_STATEIMAGEMASK), l, deCustomDraw);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deRawMenuMessage, LVIS_STATEIMAGEMASK), l, deRawMenuMessage);
			reinterpret_cast<IReBar*>(pControl)->put_DisabledEvents(static_cast<DisabledEventsConstants>(l));

		} else if(m_ppUnk[object]->QueryInterface(IID_IToolBar, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
			DisabledEventsConstants disabledEvents = static_cast<DisabledEventsConstants>(0);
			reinterpret_cast<IToolBar*>(pControl)->get_DisabledEvents(&disabledEvents);
			LONG l = static_cast<LONG>(disabledEvents);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deMouseEvents, LVIS_STATEIMAGEMASK), l, deMouseEvents);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deClickEvents, LVIS_STATEIMAGEMASK), l, deClickEvents);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deKeyboardEvents, LVIS_STATEIMAGEMASK), l, deKeyboardEvents);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deButtonInsertionEvents, LVIS_STATEIMAGEMASK), l, deButtonInsertionEvents);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deButtonDeletionEvents, LVIS_STATEIMAGEMASK), l, deButtonDeletionEvents);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deFreeButtonData, LVIS_STATEIMAGEMASK), l, deFreeButtonData);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deCustomDraw, LVIS_STATEIMAGEMASK), l, deCustomDraw);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deHotButtonChangeEvents, LVIS_STATEIMAGEMASK), l, deHotButtonChangeEvents);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deAcceleratorEvents, LVIS_STATEIMAGEMASK), l, deAcceleratorEvents);
			SetBit(controls.disabledEventsList.GetItemState(properties.disabledEventsItemIndices.deRawMenuMessage, LVIS_STATEIMAGEMASK), l, deRawMenuMessage);
			reinterpret_cast<IToolBar*>(pControl)->put_DisabledEvents(static_cast<DisabledEventsConstants>(l));
			pControl->Release();
		}
	}

	SetDirty(FALSE);
}

void CommonProperties::LoadSettings(void)
{
	if(!controls.toolbar.IsWindow()) {
		// this will happen in Visual Studio's dialog editor if settings are loaded from a file
		return;
	}

	controls.toolbar.EnableButton(ID_LOADSETTINGS, (m_nObjects == 1));
	controls.toolbar.EnableButton(ID_SAVESETTINGS, (m_nObjects == 1));

	// get the settings
	int numberOfReBars = 0;
	int numberOfToolBars = 0;
	DisabledEventsConstants* pDisabledEvents = static_cast<DisabledEventsConstants*>(HeapAlloc(GetProcessHeap(), 0, m_nObjects * sizeof(DisabledEventsConstants)));
	if(pDisabledEvents) {
		ZeroMemory(pDisabledEvents, m_nObjects * sizeof(DisabledEventsConstants));
		for(UINT object = 0; object < m_nObjects; ++object) {
			LPUNKNOWN pControl = NULL;
			if(m_ppUnk[object]->QueryInterface(IID_IReBar, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
				++numberOfReBars;
				reinterpret_cast<IReBar*>(pControl)->get_DisabledEvents(&pDisabledEvents[object]);
				pControl->Release();

			} else if(m_ppUnk[object]->QueryInterface(IID_IToolBar, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
				++numberOfToolBars;
				reinterpret_cast<IToolBar*>(pControl)->get_DisabledEvents(&pDisabledEvents[object]);
				pControl->Release();
			}
		}

		// fill the listboxes
		properties.disabledEventsItemIndices.Reset();
		LONG* pl = reinterpret_cast<LONG*>(pDisabledEvents);
		properties.hWndCheckMarksAreSetFor = controls.disabledEventsList;
		controls.disabledEventsList.DeleteAllItems();
		properties.disabledEventsItemIndices.deMouseEvents = controls.disabledEventsList.AddItem(0, 0, TEXT("Mouse events"));
		controls.disabledEventsList.SetItemState(properties.disabledEventsItemIndices.deMouseEvents, CalculateStateImageMask(m_nObjects, pl, deMouseEvents), LVIS_STATEIMAGEMASK);
		properties.disabledEventsItemIndices.deClickEvents = controls.disabledEventsList.AddItem(1, 0, TEXT("Click events"));
		controls.disabledEventsList.SetItemState(properties.disabledEventsItemIndices.deClickEvents, CalculateStateImageMask(m_nObjects, pl, deClickEvents), LVIS_STATEIMAGEMASK);
		properties.disabledEventsItemIndices.deKeyboardEvents = controls.disabledEventsList.AddItem(2, 0, TEXT("Keyboard events"));
		controls.disabledEventsList.SetItemState(properties.disabledEventsItemIndices.deKeyboardEvents, CalculateStateImageMask(m_nObjects, pl, deKeyboardEvents), LVIS_STATEIMAGEMASK);
		if(numberOfReBars > 0) {
			properties.disabledEventsItemIndices.deBandInsertionEvents = controls.disabledEventsList.AddItem(3, 0, TEXT("Band insertion events"));
			controls.disabledEventsList.SetItemState(properties.disabledEventsItemIndices.deBandInsertionEvents, CalculateStateImageMask(m_nObjects, pl, deBandInsertionEvents), LVIS_STATEIMAGEMASK);
		}
		if(numberOfToolBars > 0) {
			properties.disabledEventsItemIndices.deButtonInsertionEvents = controls.disabledEventsList.AddItem(4, 0, TEXT("Button insertion events"));
			controls.disabledEventsList.SetItemState(properties.disabledEventsItemIndices.deButtonInsertionEvents, CalculateStateImageMask(m_nObjects, pl, deButtonInsertionEvents), LVIS_STATEIMAGEMASK);
		}
		if(numberOfReBars > 0) {
			properties.disabledEventsItemIndices.deBandDeletionEvents = controls.disabledEventsList.AddItem(5, 0, TEXT("Band deletion events"));
			controls.disabledEventsList.SetItemState(properties.disabledEventsItemIndices.deBandDeletionEvents, CalculateStateImageMask(m_nObjects, pl, deBandDeletionEvents), LVIS_STATEIMAGEMASK);
		}
		if(numberOfToolBars > 0) {
			properties.disabledEventsItemIndices.deButtonDeletionEvents = controls.disabledEventsList.AddItem(6, 0, TEXT("Button deletion events"));
			controls.disabledEventsList.SetItemState(properties.disabledEventsItemIndices.deButtonDeletionEvents, CalculateStateImageMask(m_nObjects, pl, deButtonDeletionEvents), LVIS_STATEIMAGEMASK);
		}
		if(numberOfReBars > 0) {
			properties.disabledEventsItemIndices.deFreeBandData = controls.disabledEventsList.AddItem(7, 0, TEXT("FreeBandData event"));
			controls.disabledEventsList.SetItemState(properties.disabledEventsItemIndices.deFreeBandData, CalculateStateImageMask(m_nObjects, pl, deFreeBandData), LVIS_STATEIMAGEMASK);
		}
		if(numberOfToolBars > 0) {
			properties.disabledEventsItemIndices.deFreeButtonData = controls.disabledEventsList.AddItem(8, 0, TEXT("FreeButtonData event"));
			controls.disabledEventsList.SetItemState(properties.disabledEventsItemIndices.deFreeButtonData, CalculateStateImageMask(m_nObjects, pl, deFreeButtonData), LVIS_STATEIMAGEMASK);
		}
		properties.disabledEventsItemIndices.deCustomDraw = controls.disabledEventsList.AddItem(9, 0, TEXT("CustomDraw events"));
		controls.disabledEventsList.SetItemState(properties.disabledEventsItemIndices.deCustomDraw, CalculateStateImageMask(m_nObjects, pl, deCustomDraw), LVIS_STATEIMAGEMASK);
		if(numberOfToolBars > 0) {
			properties.disabledEventsItemIndices.deHotButtonChangeEvents = controls.disabledEventsList.AddItem(10, 0, TEXT("Hot button change events"));
			controls.disabledEventsList.SetItemState(properties.disabledEventsItemIndices.deHotButtonChangeEvents, CalculateStateImageMask(m_nObjects, pl, deHotButtonChangeEvents), LVIS_STATEIMAGEMASK);
			properties.disabledEventsItemIndices.deAcceleratorEvents = controls.disabledEventsList.AddItem(11, 0, TEXT("Keyboard accelerator events"));
			controls.disabledEventsList.SetItemState(properties.disabledEventsItemIndices.deAcceleratorEvents, CalculateStateImageMask(m_nObjects, pl, deAcceleratorEvents), LVIS_STATEIMAGEMASK);
		}
		properties.disabledEventsItemIndices.deRawMenuMessage = controls.disabledEventsList.AddItem(12, 0, TEXT("RawMenuMessage event"));
		controls.disabledEventsList.SetItemState(properties.disabledEventsItemIndices.deRawMenuMessage, CalculateStateImageMask(m_nObjects, pl, deRawMenuMessage), LVIS_STATEIMAGEMASK);
		controls.disabledEventsList.SetColumnWidth(0, LVSCW_AUTOSIZE);

		properties.hWndCheckMarksAreSetFor = NULL;
		HeapFree(GetProcessHeap(), 0, pDisabledEvents);
	}

	SetDirty(FALSE);
}

int CommonProperties::CalculateStateImageMask(UINT arraysize, LONG* pArray, LONG bitsToCheckFor)
{
	int stateImageIndex = 1;
	for(UINT object = 0; object < arraysize; ++object) {
		if(pArray[object] & bitsToCheckFor) {
			if(stateImageIndex == 1) {
				stateImageIndex = (object == 0 ? 2 : 3);
			}
		} else {
			if(stateImageIndex == 2) {
				stateImageIndex = (object == 0 ? 1 : 3);
			}
		}
	}

	return INDEXTOSTATEIMAGEMASK(stateImageIndex);
}

void CommonProperties::SetBit(int stateImageMask, LONG& value, LONG bitToSet)
{
	stateImageMask >>= 12;
	switch(stateImageMask) {
		case 1:
			value &= ~bitToSet;
			break;
		case 2:
			value |= bitToSet;
			break;
	}
}
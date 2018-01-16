// ToolBarButtons.cpp: Manages a collection of ToolBarButton objects

#include "stdafx.h"
#include "ToolBarButtons.h"
#include "ClassFactory.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP ToolBarButtons::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IToolBarButtons, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


//////////////////////////////////////////////////////////////////////
// implementation of IEnumVARIANT
STDMETHODIMP ToolBarButtons::Clone(IEnumVARIANT** ppEnumerator)
{
	ATLASSERT_POINTER(ppEnumerator, LPENUMVARIANT);
	if(!ppEnumerator) {
		return E_POINTER;
	}

	*ppEnumerator = NULL;
	CComObject<ToolBarButtons>* pTBarButtonsObj = NULL;
	CComObject<ToolBarButtons>::CreateInstance(&pTBarButtonsObj);
	pTBarButtonsObj->AddRef();

	// clone all settings
	properties.CopyTo(&pTBarButtonsObj->properties);

	pTBarButtonsObj->QueryInterface(IID_IEnumVARIANT, reinterpret_cast<LPVOID*>(ppEnumerator));
	pTBarButtonsObj->Release();
	return S_OK;
}

STDMETHODIMP ToolBarButtons::Next(ULONG numberOfMaxButtons, VARIANT* pButtons, ULONG* pNumberOfButtonsReturned)
{
	ATLASSERT_POINTER(pButtons, VARIANT);
	if(!pButtons) {
		return E_POINTER;
	}

	HWND hWndTBar = properties.GetTBarHWnd();
	ATLASSERT(IsWindow(hWndTBar));

	// check each button in the tool bar
	int numberOfButtons = static_cast<int>(SendMessage(hWndTBar, TB_BUTTONCOUNT, 0, 0));
	ULONG i = 0;
	for(i = 0; i < numberOfMaxButtons; ++i) {
		VariantInit(&pButtons[i]);

		do {
			if(properties.lastEnumeratedButton >= 0) {
				if(properties.lastEnumeratedButton < numberOfButtons) {
					properties.lastEnumeratedButton = GetNextButtonToProcess(properties.lastEnumeratedButton, numberOfButtons);
				}
			} else {
				properties.lastEnumeratedButton = GetFirstButtonToProcess(numberOfButtons);
			}
			if(properties.lastEnumeratedButton >= numberOfButtons) {
				properties.lastEnumeratedButton = -1;
			}
		} while((properties.lastEnumeratedButton != -1) && (!IsPartOfCollection(properties.lastEnumeratedButton, hWndTBar)));

		if(properties.lastEnumeratedButton != -1) {
			ClassFactory::InitToolBarButton(properties.lastEnumeratedButton, FALSE, properties.pOwnerTBar, IID_IDispatch, reinterpret_cast<LPUNKNOWN*>(&pButtons[i].pdispVal));
			pButtons[i].vt = VT_DISPATCH;
		} else {
			// there's nothing more to iterate
			break;
		}
	}
	if(pNumberOfButtonsReturned) {
		*pNumberOfButtonsReturned = i;
	}

	return (i == numberOfMaxButtons ? S_OK : S_FALSE);
}

STDMETHODIMP ToolBarButtons::Reset(void)
{
	properties.lastEnumeratedButton = -1;
	return S_OK;
}

STDMETHODIMP ToolBarButtons::Skip(ULONG numberOfButtonsToSkip)
{
	VARIANT dummy;
	ULONG numButtonsReturned = 0;
	// we could skip all buttons at once, but it's easier to skip them one after the other
	for(ULONG i = 1; i <= numberOfButtonsToSkip; ++i) {
		HRESULT hr = Next(1, &dummy, &numButtonsReturned);
		VariantClear(&dummy);
		if(hr != S_OK || numButtonsReturned == 0) {
			// there're no more buttons to skip, so don't try it anymore
			break;
		}
	}
	return S_OK;
}
// implementation of IEnumVARIANT
//////////////////////////////////////////////////////////////////////


BOOL ToolBarButtons::IsValidBooleanFilter(VARIANT& filter)
{
	BOOL isValid = TRUE;

	if((filter.vt == (VT_ARRAY | VT_VARIANT)) && filter.parray) {
		LONG lBound = 0;
		LONG uBound = 0;

		if((SafeArrayGetLBound(filter.parray, 1, &lBound) == S_OK) && (SafeArrayGetUBound(filter.parray, 1, &uBound) == S_OK)) {
			// now that we have the bounds, iterate the array
			VARIANT element;
			if(lBound > uBound) {
				isValid = FALSE;
			}
			for(LONG i = lBound; i <= uBound && isValid; ++i) {
				if(SafeArrayGetElement(filter.parray, &i, &element) == S_OK) {
					isValid = (element.vt == VT_BOOL);
					VariantClear(&element);
				} else {
					isValid = FALSE;
				}
			}
		} else {
			isValid = FALSE;
		}
	} else {
		isValid = FALSE;
	}

	return isValid;
}

BOOL ToolBarButtons::IsValidIntegerFilter(VARIANT& filter, int minValue, int maxValue)
{
	BOOL isValid = TRUE;

	if((filter.vt == (VT_ARRAY | VT_VARIANT)) && filter.parray) {
		LONG lBound = 0;
		LONG uBound = 0;

		if((SafeArrayGetLBound(filter.parray, 1, &lBound) == S_OK) && (SafeArrayGetUBound(filter.parray, 1, &uBound) == S_OK)) {
			// now that we have the bounds, iterate the array
			VARIANT element;
			if(lBound > uBound) {
				isValid = FALSE;
			}
			for(LONG i = lBound; i <= uBound && isValid; ++i) {
				if(SafeArrayGetElement(filter.parray, &i, &element) == S_OK) {
					isValid = SUCCEEDED(VariantChangeType(&element, &element, 0, VT_INT)) && element.intVal >= minValue && element.intVal <= maxValue;
					VariantClear(&element);
				} else {
					isValid = FALSE;
				}
			}
		} else {
			isValid = FALSE;
		}
	} else {
		isValid = FALSE;
	}

	return isValid;
}

BOOL ToolBarButtons::IsValidIntegerFilter(VARIANT& filter, int minValue)
{
	BOOL isValid = TRUE;

	if((filter.vt == (VT_ARRAY | VT_VARIANT)) && filter.parray) {
		LONG lBound = 0;
		LONG uBound = 0;

		if((SafeArrayGetLBound(filter.parray, 1, &lBound) == S_OK) && (SafeArrayGetUBound(filter.parray, 1, &uBound) == S_OK)) {
			// now that we have the bounds, iterate the array
			VARIANT element;
			if(lBound > uBound) {
				isValid = FALSE;
			}
			for(LONG i = lBound; i <= uBound && isValid; ++i) {
				if(SafeArrayGetElement(filter.parray, &i, &element) == S_OK) {
					isValid = SUCCEEDED(VariantChangeType(&element, &element, 0, VT_INT)) && element.intVal >= minValue;
					VariantClear(&element);
				} else {
					isValid = FALSE;
				}
			}
		} else {
			isValid = FALSE;
		}
	} else {
		isValid = FALSE;
	}

	return isValid;
}

BOOL ToolBarButtons::IsValidIntegerFilter(VARIANT& filter)
{
	BOOL isValid = TRUE;

	if((filter.vt == (VT_ARRAY | VT_VARIANT)) && filter.parray) {
		LONG lBound = 0;
		LONG uBound = 0;

		if((SafeArrayGetLBound(filter.parray, 1, &lBound) == S_OK) && (SafeArrayGetUBound(filter.parray, 1, &uBound) == S_OK)) {
			// now that we have the bounds, iterate the array
			VARIANT element;
			if(lBound > uBound) {
				isValid = FALSE;
			}
			for(LONG i = lBound; i <= uBound && isValid; ++i) {
				if(SafeArrayGetElement(filter.parray, &i, &element) == S_OK) {
					isValid = SUCCEEDED(VariantChangeType(&element, &element, 0, VT_UI4));
					VariantClear(&element);
				} else {
					isValid = FALSE;
				}
			}
		} else {
			isValid = FALSE;
		}
	} else {
		isValid = FALSE;
	}

	return isValid;
}

BOOL ToolBarButtons::IsValidStringFilter(VARIANT& filter)
{
	BOOL isValid = TRUE;

	if((filter.vt == (VT_ARRAY | VT_VARIANT)) && filter.parray) {
		LONG lBound = 0;
		LONG uBound = 0;

		if((SafeArrayGetLBound(filter.parray, 1, &lBound) == S_OK) && (SafeArrayGetUBound(filter.parray, 1, &uBound) == S_OK)) {
			// now that we have the bounds, iterate the array
			VARIANT element;
			if(lBound > uBound) {
				isValid = FALSE;
			}
			for(LONG i = lBound; i <= uBound && isValid; ++i) {
				if(SafeArrayGetElement(filter.parray, &i, &element) == S_OK) {
					isValid = (element.vt == VT_BSTR);
					VariantClear(&element);
				} else {
					isValid = FALSE;
				}
			}
		} else {
			isValid = FALSE;
		}
	} else {
		isValid = FALSE;
	}

	return isValid;
}

int ToolBarButtons::GetFirstButtonToProcess(int numberOfButtons)
{
	if(numberOfButtons == 0) {
		return -1;
	}
	return 0;
}

int ToolBarButtons::GetNextButtonToProcess(int previousButton, int numberOfButtons)
{
	if(previousButton < numberOfButtons - 1) {
		return previousButton + 1;
	}
	return -1;
}

BOOL ToolBarButtons::IsBooleanInSafeArray(LPSAFEARRAY pSafeArray, VARIANT_BOOL value, LPVOID pComparisonFunction)
{
	LONG lBound = 0;
	LONG uBound = 0;
	SafeArrayGetLBound(pSafeArray, 1, &lBound);
	SafeArrayGetUBound(pSafeArray, 1, &uBound);

	VARIANT element;
	BOOL foundMatch = FALSE;
	for(LONG i = lBound; i <= uBound && !foundMatch; ++i) {
		SafeArrayGetElement(pSafeArray, &i, &element);
		if(pComparisonFunction) {
			typedef BOOL WINAPI ComparisonFn(VARIANT_BOOL, VARIANT_BOOL);
			ComparisonFn* pComparisonFn = reinterpret_cast<ComparisonFn*>(pComparisonFunction);
			if(pComparisonFn(value, element.boolVal)) {
				foundMatch = TRUE;
			}
		} else {
			if(element.boolVal == value) {
				foundMatch = TRUE;
			}
		}
		VariantClear(&element);
	}

	return foundMatch;
}

BOOL ToolBarButtons::IsIntegerInSafeArray(LPSAFEARRAY pSafeArray, int value, LPVOID pComparisonFunction)
{
	LONG lBound = 0;
	LONG uBound = 0;
	SafeArrayGetLBound(pSafeArray, 1, &lBound);
	SafeArrayGetUBound(pSafeArray, 1, &uBound);

	VARIANT element;
	BOOL foundMatch = FALSE;
	for(LONG i = lBound; i <= uBound && !foundMatch; ++i) {
		SafeArrayGetElement(pSafeArray, &i, &element);
		int v = 0;
		if(SUCCEEDED(VariantChangeType(&element, &element, 0, VT_INT))) {
			v = element.intVal;
		}
		if(pComparisonFunction) {
			typedef BOOL WINAPI ComparisonFn(LONG, LONG);
			ComparisonFn* pComparisonFn = reinterpret_cast<ComparisonFn*>(pComparisonFunction);
			if(pComparisonFn(value, v)) {
				foundMatch = TRUE;
			}
		} else {
			if(v == value) {
				foundMatch = TRUE;
			}
		}
		VariantClear(&element);
	}

	return foundMatch;
}

BOOL ToolBarButtons::IsStringInSafeArray(LPSAFEARRAY pSafeArray, BSTR value, LPVOID pComparisonFunction)
{
	LONG lBound = 0;
	LONG uBound = 0;
	SafeArrayGetLBound(pSafeArray, 1, &lBound);
	SafeArrayGetUBound(pSafeArray, 1, &uBound);

	VARIANT element;
	BOOL foundMatch = FALSE;
	for(LONG i = lBound; i <= uBound && !foundMatch; ++i) {
		SafeArrayGetElement(pSafeArray, &i, &element);
		if(pComparisonFunction) {
			typedef BOOL WINAPI ComparisonFn(BSTR, BSTR);
			ComparisonFn* pComparisonFn = reinterpret_cast<ComparisonFn*>(pComparisonFunction);
			if(pComparisonFn(value, element.bstrVal)) {
				foundMatch = TRUE;
			}
		} else {
			if(properties.caseSensitiveFilters) {
				if(lstrcmpW(OLE2W(element.bstrVal), OLE2W(value)) == 0) {
					foundMatch = TRUE;
				}
			} else {
				if(lstrcmpiW(OLE2W(element.bstrVal), OLE2W(value)) == 0) {
					foundMatch = TRUE;
				}
			}
		}
		VariantClear(&element);
	}

	return foundMatch;
}

BOOL ToolBarButtons::IsPartOfCollection(int buttonIndex, HWND hWndTBar/* = NULL*/)
{
	if(!hWndTBar) {
		hWndTBar = properties.GetTBarHWnd();
	}
	ATLASSERT(IsWindow(hWndTBar));

	if(!IsValidButtonIndex(buttonIndex, hWndTBar)) {
		return FALSE;
	}

	BOOL isPart = FALSE;
	// we declare this one here to avoid compiler warnings
	TBBUTTONINFO button = {0};
	button.cbSize = sizeof(button);
	button.dwMask = TBIF_BYINDEX;

	if(properties.filterType[fpIndex] != ftDeactivated) {
		if(IsIntegerInSafeArray(properties.filter[fpIndex].parray, buttonIndex, properties.comparisonFunction[fpIndex])) {
			if(properties.filterType[fpIndex] == ftExcluding) {
				goto Exit;
			}
		} else {
			if(properties.filterType[fpIndex] == ftIncluding) {
				goto Exit;
			}
		}
	}

	BOOL mustRetrieveButtonData = FALSE;
	if(properties.filterType[fpIconIndex] != ftDeactivated || properties.filterType[fpImageListIndex] != ftDeactivated) {
		button.dwMask |= TBIF_IMAGE | TBIF_STYLE;
		mustRetrieveButtonData = TRUE;
	}
	if(properties.filterType[fpButtonData] != ftDeactivated) {
		button.dwMask |= TBIF_LPARAM;
		mustRetrieveButtonData = TRUE;
	}
	if(properties.filterType[fpText] != ftDeactivated) {
		button.dwMask |= TBIF_COMMAND;
		mustRetrieveButtonData = TRUE;
	}
	if(properties.filterType[fpAutoSize] != ftDeactivated || properties.filterType[fpButtonType] != ftDeactivated || properties.filterType[fpDisplayText] != ftDeactivated || properties.filterType[fpDropDownStyle] != ftDeactivated || properties.filterType[fpPartOfGroup] != ftDeactivated || properties.filterType[fpUseMnemonic] != ftDeactivated) {
		button.dwMask |= TBIF_STYLE;
		mustRetrieveButtonData = TRUE;
	}
	if(properties.filterType[fpEnabled] != ftDeactivated || properties.filterType[fpFollowedByLineBreak] != ftDeactivated || properties.filterType[fpMarked] != ftDeactivated || properties.filterType[fpPushed] != ftDeactivated || properties.filterType[fpSelectionState] != ftDeactivated || properties.filterType[fpShowingEllipsis] != ftDeactivated || properties.filterType[fpVisible] != ftDeactivated) {
		button.dwMask |= TBIF_STATE;
		mustRetrieveButtonData = TRUE;
	}
	if(properties.filterType[fpWidth] != ftDeactivated) {
		button.dwMask |= TBIF_SIZE;
		mustRetrieveButtonData = TRUE;
	}
	if(properties.filterType[fpID] != ftDeactivated || properties.filterType[fpButtonType] != ftDeactivated) {
		button.dwMask |= TBIF_COMMAND;
		mustRetrieveButtonData = TRUE;
	}

	if(mustRetrieveButtonData) {
		if(!SendMessage(hWndTBar, TB_GETBUTTONINFO, buttonIndex, reinterpret_cast<LPARAM>(&button))) {
			goto Exit;
		}

		// apply the filters

		if(properties.filterType[fpID] != ftDeactivated) {
			if(IsIntegerInSafeArray(properties.filter[fpID].parray, static_cast<int>(button.idCommand), properties.comparisonFunction[fpID])) {
				if(properties.filterType[fpID] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpID] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpButtonData] != ftDeactivated) {
			if(IsIntegerInSafeArray(properties.filter[fpButtonData].parray, static_cast<int>(button.lParam), properties.comparisonFunction[fpButtonData])) {
				if(properties.filterType[fpButtonData] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpButtonData] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpText] != ftDeactivated) {
			CComBSTR text = _T("");
			int textLength = static_cast<int>(SendMessage(hWndTBar, TB_GETBUTTONTEXT, button.idCommand, NULL));
			if(textLength > -1) {
				LPTSTR pBuffer = static_cast<LPTSTR>(HeapAlloc(GetProcessHeap(), 0, (textLength + 1) * sizeof(TCHAR)));
				if(pBuffer) {
					ZeroMemory(pBuffer, (textLength + 1) * sizeof(TCHAR));
					if(static_cast<int>(SendMessage(hWndTBar, TB_GETBUTTONTEXT, button.idCommand, reinterpret_cast<LPARAM>(pBuffer))) > -1) {
						text = _bstr_t(pBuffer).Detach();
					}
					HeapFree(GetProcessHeap(), 0, pBuffer);
				}
			}
			if(IsStringInSafeArray(properties.filter[fpText].parray, text, properties.comparisonFunction[fpText])) {
				if(properties.filterType[fpText] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpText] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpIconIndex] != ftDeactivated) {
			int iconIndex = (button.iImage == I_IMAGECALLBACK ? -1 : (button.fsStyle & BTNS_SEP ? I_IMAGENONE : LOWORD(button.iImage)));
			if(IsIntegerInSafeArray(properties.filter[fpIconIndex].parray, iconIndex, properties.comparisonFunction[fpIconIndex])) {
				if(properties.filterType[fpIconIndex] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpIconIndex] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpImageListIndex] != ftDeactivated) {
			int iconIndex = (button.iImage == I_IMAGECALLBACK ? -1 : (button.fsStyle & BTNS_SEP ? 0 : HIWORD(button.iImage)));
			if(IsIntegerInSafeArray(properties.filter[fpImageListIndex].parray, iconIndex, properties.comparisonFunction[fpImageListIndex])) {
				if(properties.filterType[fpImageListIndex] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpImageListIndex] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpEnabled] != ftDeactivated) {
			if(IsBooleanInSafeArray(properties.filter[fpEnabled].parray, BOOL2VARIANTBOOL(button.fsState & TBSTATE_ENABLED), properties.comparisonFunction[fpEnabled])) {
				if(properties.filterType[fpEnabled] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpEnabled] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpVisible] != ftDeactivated) {
			if(IsBooleanInSafeArray(properties.filter[fpVisible].parray, BOOL2VARIANTBOOL(!(button.fsState & TBSTATE_HIDDEN)), properties.comparisonFunction[fpVisible])) {
				if(properties.filterType[fpVisible] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpVisible] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpMarked] != ftDeactivated) {
			if(IsBooleanInSafeArray(properties.filter[fpMarked].parray, BOOL2VARIANTBOOL(button.fsState & TBSTATE_MARKED), properties.comparisonFunction[fpMarked])) {
				if(properties.filterType[fpMarked] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpMarked] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpButtonType] != ftDeactivated) {
			ButtonTypeConstants buttonType = btyCommandButton;
			if(button.fsStyle & BTNS_SEP) {
				buttonType = btySeparator;
			} else if(button.fsStyle & BTNS_CHECK) {
				buttonType = btyCheckButton;
			} else if(button.fsStyle == BTNS_BUTTON) {
				if(properties.pOwnerTBar->IsPlaceholderButton(button.idCommand)) {
					buttonType = btyPlaceholder;
				} else {
					buttonType = btyCommandButton;
				}
			}
			if(IsIntegerInSafeArray(properties.filter[fpButtonType].parray, buttonType, properties.comparisonFunction[fpButtonType])) {
				if(properties.filterType[fpButtonType] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpButtonType] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpDropDownStyle] != ftDeactivated) {
			DropDownStyleConstants dropDownStyle = ddstNoDropDown;
			if(button.fsStyle & BTNS_WHOLEDROPDOWN) {
				dropDownStyle = ddstAlwaysWholeButton;
			} else if(button.fsStyle & BTNS_DROPDOWN) {
				dropDownStyle = ddstNormal;
			}
			if(IsIntegerInSafeArray(properties.filter[fpDropDownStyle].parray, dropDownStyle, properties.comparisonFunction[fpDropDownStyle])) {
				if(properties.filterType[fpDropDownStyle] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpDropDownStyle] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpSelectionState] != ftDeactivated) {
			SelectionStateConstants selectionState = ssUnchecked;
			switch(button.fsState & (TBSTATE_CHECKED | TBSTATE_INDETERMINATE)) {
				case 0:
					selectionState = ssUnchecked;
					break;
				case TBSTATE_CHECKED:
					selectionState = ssChecked;
					break;
				case TBSTATE_INDETERMINATE:
					selectionState = ssIndeterminate;
					break;
			}
			if(IsIntegerInSafeArray(properties.filter[fpSelectionState].parray, selectionState, properties.comparisonFunction[fpSelectionState])) {
				if(properties.filterType[fpSelectionState] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpSelectionState] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpPartOfGroup] != ftDeactivated) {
			if(IsBooleanInSafeArray(properties.filter[fpPartOfGroup].parray, BOOL2VARIANTBOOL(button.fsStyle & BTNS_GROUP), properties.comparisonFunction[fpPartOfGroup])) {
				if(properties.filterType[fpPartOfGroup] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpPartOfGroup] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpPushed] != ftDeactivated) {
			if(IsBooleanInSafeArray(properties.filter[fpPushed].parray, BOOL2VARIANTBOOL(button.fsState & TBSTATE_PRESSED), properties.comparisonFunction[fpPushed])) {
				if(properties.filterType[fpPushed] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpPushed] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpDisplayText] != ftDeactivated) {
			if(IsBooleanInSafeArray(properties.filter[fpDisplayText].parray, BOOL2VARIANTBOOL(button.fsStyle & BTNS_SHOWTEXT), properties.comparisonFunction[fpDisplayText])) {
				if(properties.filterType[fpDisplayText] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpDisplayText] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpShowingEllipsis] != ftDeactivated) {
			if(IsBooleanInSafeArray(properties.filter[fpShowingEllipsis].parray, BOOL2VARIANTBOOL(button.fsState & TBSTATE_ELLIPSES), properties.comparisonFunction[fpShowingEllipsis])) {
				if(properties.filterType[fpShowingEllipsis] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpShowingEllipsis] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpAutoSize] != ftDeactivated) {
			if(IsBooleanInSafeArray(properties.filter[fpAutoSize].parray, BOOL2VARIANTBOOL(button.fsStyle & BTNS_AUTOSIZE), properties.comparisonFunction[fpAutoSize])) {
				if(properties.filterType[fpAutoSize] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpAutoSize] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpWidth] != ftDeactivated) {
			if(IsIntegerInSafeArray(properties.filter[fpWidth].parray, button.cx, properties.comparisonFunction[fpWidth])) {
				if(properties.filterType[fpWidth] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpWidth] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpFollowedByLineBreak] != ftDeactivated) {
			if(IsBooleanInSafeArray(properties.filter[fpFollowedByLineBreak].parray, BOOL2VARIANTBOOL(button.fsState & TBSTATE_WRAP), properties.comparisonFunction[fpFollowedByLineBreak])) {
				if(properties.filterType[fpFollowedByLineBreak] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpFollowedByLineBreak] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpUseMnemonic] != ftDeactivated) {
			if(IsBooleanInSafeArray(properties.filter[fpUseMnemonic].parray, BOOL2VARIANTBOOL(!(button.fsStyle & BTNS_NOPREFIX)), properties.comparisonFunction[fpUseMnemonic])) {
				if(properties.filterType[fpUseMnemonic] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpUseMnemonic] == ftIncluding) {
					goto Exit;
				}
			}
		}
	}
	isPart = TRUE;

Exit:
	if(button.pszText) {
		SECUREFREE(button.pszText);
	}
	return isPart;
}

void ToolBarButtons::OptimizeFilter(FilteredPropertyConstants filteredProperty)
{
	if(filteredProperty != fpAutoSize && filteredProperty != fpDisplayText && filteredProperty != fpEnabled && filteredProperty != fpFollowedByLineBreak && filteredProperty != fpMarked && filteredProperty != fpPartOfGroup && filteredProperty != fpPushed && filteredProperty != fpShowingEllipsis && filteredProperty != fpUseMnemonic && filteredProperty != fpVisible) {
		// currently we optimize boolean filters only
		return;
	}

	LONG lBound = 0;
	LONG uBound = 0;
	SafeArrayGetLBound(properties.filter[filteredProperty].parray, 1, &lBound);
	SafeArrayGetUBound(properties.filter[filteredProperty].parray, 1, &uBound);

	// now that we have the bounds, iterate the array
	VARIANT element;
	UINT numberOfTrues = 0;
	UINT numberOfFalses = 0;
	for(LONG i = lBound; i <= uBound; ++i) {
		SafeArrayGetElement(properties.filter[filteredProperty].parray, &i, &element);
		if(element.boolVal == VARIANT_FALSE) {
			++numberOfFalses;
		} else {
			++numberOfTrues;
		}

		VariantClear(&element);
	}

	if(numberOfTrues > 0 && numberOfFalses > 0) {
		// we've something like True Or False Or True - we can deactivate this filter
		properties.filterType[filteredProperty] = ftDeactivated;
		VariantClear(&properties.filter[filteredProperty]);
	} else if(numberOfTrues == 0 && numberOfFalses > 1) {
		// False Or False Or False... is still False, so we need just one False
		VariantClear(&properties.filter[filteredProperty]);
		properties.filter[filteredProperty].vt = VT_ARRAY | VT_VARIANT;
		properties.filter[filteredProperty].parray = SafeArrayCreateVectorEx(VT_VARIANT, 1, 1, NULL);

		VARIANT newElement;
		VariantInit(&newElement);
		newElement.vt = VT_BOOL;
		newElement.boolVal = VARIANT_FALSE;
		LONG index = 1;
		SafeArrayPutElement(properties.filter[filteredProperty].parray, &index, &newElement);
		VariantClear(&newElement);
	} else if(numberOfFalses == 0 && numberOfTrues > 1) {
		// True Or True Or True... is still True, so we need just one True
		VariantClear(&properties.filter[filteredProperty]);
		properties.filter[filteredProperty].vt = VT_ARRAY | VT_VARIANT;
		properties.filter[filteredProperty].parray = SafeArrayCreateVectorEx(VT_VARIANT, 1, 1, NULL);

		VARIANT newElement;
		VariantInit(&newElement);
		newElement.vt = VT_BOOL;
		newElement.boolVal = VARIANT_TRUE;
		LONG index = 1;
		SafeArrayPutElement(properties.filter[filteredProperty].parray, &index, &newElement);
		VariantClear(&newElement);
	}
}

#ifdef USE_STL
	HRESULT ToolBarButtons::RemoveButtons(std::list<int>& buttonsToRemove, HWND hWndTBar)
#else
	HRESULT ToolBarButtons::RemoveButtons(CAtlList<int>& buttonsToRemove, HWND hWndTBar)
#endif
{
	ATLASSERT(IsWindow(hWndTBar));

	CWindowEx(hWndTBar).InternalSetRedraw(FALSE);
	// sort in reverse order
	#ifdef USE_STL
		buttonsToRemove.sort(std::greater<int>());
		for(std::list<int>::const_iterator iter = buttonsToRemove.begin(); iter != buttonsToRemove.end(); ++iter) {
			SendMessage(hWndTBar, TB_DELETEBUTTON, *iter, 0);
		}
	#else
		// perform a crude bubble sort
		for(size_t j = 0; j < buttonsToRemove.GetCount(); ++j) {
			for(size_t i = 0; i < buttonsToRemove.GetCount() - 1; ++i) {
				if(buttonsToRemove.GetAt(buttonsToRemove.FindIndex(i)) < buttonsToRemove.GetAt(buttonsToRemove.FindIndex(i + 1))) {
					buttonsToRemove.SwapElements(buttonsToRemove.FindIndex(i), buttonsToRemove.FindIndex(i + 1));
				}
			}
		}

		for(size_t i = 0; i < buttonsToRemove.GetCount(); ++i) {
			SendMessage(hWndTBar, TB_DELETEBUTTON, buttonsToRemove.GetAt(buttonsToRemove.FindIndex(i)), 0);
		}
	#endif
	CWindowEx(hWndTBar).InternalSetRedraw(TRUE);

	return S_OK;
}


ToolBarButtons::Properties::~Properties()
{
	if(pOwnerTBar) {
		pOwnerTBar->Release();
	}
}

void ToolBarButtons::Properties::CopyTo(ToolBarButtons::Properties* pTarget)
{
	ATLASSERT_POINTER(pTarget, Properties);
	if(pTarget) {
		pTarget->pOwnerTBar = this->pOwnerTBar;
		if(pTarget->pOwnerTBar) {
			pTarget->pOwnerTBar->AddRef();
		}
		pTarget->lastEnumeratedButton = this->lastEnumeratedButton;
		pTarget->caseSensitiveFilters = this->caseSensitiveFilters;

		for(int i = 0; i < NUMBEROFFILTERS_TB; ++i) {
			VariantCopy(&pTarget->filter[i], &this->filter[i]);
			pTarget->filterType[i] = this->filterType[i];
			pTarget->comparisonFunction[i] = this->comparisonFunction[i];
		}
		pTarget->usingFilters = this->usingFilters;
	}
}

HWND ToolBarButtons::Properties::GetTBarHWnd(BOOL getChevronPopup/* = FALSE*/)
{
	ATLASSUME(pOwnerTBar);

	OLE_HANDLE handle = NULL;
	if(getChevronPopup) {
		pOwnerTBar->get_hWndChevronToolBar(&handle);
	} else {
		pOwnerTBar->get_hWnd(&handle);
	}
	return static_cast<HWND>(LongToHandle(handle));
}


void ToolBarButtons::SetOwner(ToolBar* pOwner)
{
	if(properties.pOwnerTBar) {
		properties.pOwnerTBar->Release();
	}
	properties.pOwnerTBar = pOwner;
	if(properties.pOwnerTBar) {
		properties.pOwnerTBar->AddRef();
	}
}


STDMETHODIMP ToolBarButtons::get_CaseSensitiveFilters(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.caseSensitiveFilters);
	return S_OK;
}

STDMETHODIMP ToolBarButtons::put_CaseSensitiveFilters(VARIANT_BOOL newValue)
{
	properties.caseSensitiveFilters = VARIANTBOOL2BOOL(newValue);
	return S_OK;
}

STDMETHODIMP ToolBarButtons::get_ComparisonFunction(FilteredPropertyConstants filteredProperty, LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	switch(filteredProperty) {
		case fpPartOfGroup:
		case fpIconIndex:
		case fpIndex:
		case fpButtonData:
		case fpSelectionState:
		case fpText:
		case fpWidth:
		case fpID:
		case fpFollowedByLineBreak:
		case fpDisplayText:
			*pValue = static_cast<LONG>(reinterpret_cast<LONG_PTR>(properties.comparisonFunction[filteredProperty]));
			return S_OK;
			break;
		default:
			if(filteredProperty >= fpVisible && filteredProperty <= fpUseMnemonic) {
				*pValue = static_cast<LONG>(reinterpret_cast<LONG_PTR>(properties.comparisonFunction[filteredProperty]));
				return S_OK;
			}
			break;
	}
	return E_INVALIDARG;
}

STDMETHODIMP ToolBarButtons::put_ComparisonFunction(FilteredPropertyConstants filteredProperty, LONG newValue)
{
	switch(filteredProperty) {
		case fpPartOfGroup:
		case fpIconIndex:
		case fpIndex:
		case fpButtonData:
		case fpSelectionState:
		case fpText:
		case fpWidth:
		case fpID:
		case fpFollowedByLineBreak:
		case fpDisplayText:
			properties.comparisonFunction[filteredProperty] = reinterpret_cast<LPVOID>(static_cast<LONG_PTR>(newValue));
			return S_OK;
			break;
		default:
			if(filteredProperty >= fpVisible && filteredProperty <= fpUseMnemonic) {
				properties.comparisonFunction[filteredProperty] = reinterpret_cast<LPVOID>(static_cast<LONG_PTR>(newValue));
				return S_OK;
			}
			break;
	}
	return E_INVALIDARG;
}

STDMETHODIMP ToolBarButtons::get_Filter(FilteredPropertyConstants filteredProperty, VARIANT* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT);
	if(!pValue) {
		return E_POINTER;
	}

	switch(filteredProperty) {
		case fpPartOfGroup:
		case fpIconIndex:
		case fpIndex:
		case fpButtonData:
		case fpSelectionState:
		case fpText:
		case fpWidth:
		case fpID:
		case fpFollowedByLineBreak:
		case fpDisplayText:
			VariantClear(pValue);
			VariantCopy(pValue, &properties.filter[filteredProperty]);
			return S_OK;
			break;
		default:
			if(filteredProperty >= fpVisible && filteredProperty <= fpUseMnemonic) {
				VariantClear(pValue);
				VariantCopy(pValue, &properties.filter[filteredProperty]);
				return S_OK;
			}
			break;
	}
	return E_INVALIDARG;
}

STDMETHODIMP ToolBarButtons::put_Filter(FilteredPropertyConstants filteredProperty, VARIANT newValue)
{
	// check 'newValue'
	switch(filteredProperty) {
		case fpIconIndex:
			if(!IsValidIntegerFilter(newValue, -2)) {
				// invalid value - raise VB runtime error 380
				return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
			}
			break;
		case fpImageListIndex:
			if(!IsValidIntegerFilter(newValue, -1)) {
				// invalid value - raise VB runtime error 380
				return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
			}
			break;
		case fpIndex:
		case fpWidth:
			if(!IsValidIntegerFilter(newValue, 0)) {
				// invalid value - raise VB runtime error 380
				return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
			}
			break;
		case fpDropDownStyle:
		case fpSelectionState:
			if(!IsValidIntegerFilter(newValue, 0, 2)) {
				// invalid value - raise VB runtime error 380
				return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
			}
			break;
		case fpButtonType:
			if(!IsValidIntegerFilter(newValue, 0, 3)) {
				// invalid value - raise VB runtime error 380
				return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
			}
			break;
		case fpButtonData:
		case fpID:
			if(!IsValidIntegerFilter(newValue)) {
				// invalid value - raise VB runtime error 380
				return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
			}
			break;
		case fpText:
			if(!IsValidStringFilter(newValue)) {
				// invalid value - raise VB runtime error 380
				return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
			}
			break;
		case fpAutoSize:
		case fpDisplayText:
		case fpEnabled:
		case fpFollowedByLineBreak:
		case fpMarked:
		case fpPartOfGroup:
		case fpPushed:
		case fpShowingEllipsis:
		case fpUseMnemonic:
		case fpVisible:
			if(!IsValidBooleanFilter(newValue)) {
				// invalid value - raise VB runtime error 380
				return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
			}
			break;
		default:
			return E_INVALIDARG;
			break;
	}

	VariantClear(&properties.filter[filteredProperty]);
	VariantCopy(&properties.filter[filteredProperty], &newValue);
	OptimizeFilter(filteredProperty);
	return S_OK;
}

STDMETHODIMP ToolBarButtons::get_FilterType(FilteredPropertyConstants filteredProperty, FilterTypeConstants* pValue)
{
	ATLASSERT_POINTER(pValue, FilterTypeConstants);
	if(!pValue) {
		return E_POINTER;
	}

	switch(filteredProperty) {
		case fpPartOfGroup:
		case fpIconIndex:
		case fpIndex:
		case fpButtonData:
		case fpSelectionState:
		case fpText:
		case fpWidth:
		case fpID:
		case fpFollowedByLineBreak:
		case fpDisplayText:
			*pValue = properties.filterType[filteredProperty];
			return S_OK;
			break;
		default:
			if(filteredProperty >= fpVisible && filteredProperty <= fpUseMnemonic) {
				*pValue = properties.filterType[filteredProperty];
				return S_OK;
			}
			break;
	}
	return E_INVALIDARG;
}

STDMETHODIMP ToolBarButtons::put_FilterType(FilteredPropertyConstants filteredProperty, FilterTypeConstants newValue)
{
	if(newValue < 0 || newValue > 2) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	BOOL isOkay = FALSE;
	switch(filteredProperty) {
		case fpPartOfGroup:
		case fpIconIndex:
		case fpIndex:
		case fpButtonData:
		case fpSelectionState:
		case fpText:
		case fpWidth:
		case fpID:
		case fpFollowedByLineBreak:
		case fpDisplayText:
			isOkay = TRUE;
			break;
		default:
			isOkay = (filteredProperty >= fpVisible && filteredProperty <= fpUseMnemonic);
			break;
	}
	if(isOkay) {
		properties.filterType[filteredProperty] = newValue;
		if(newValue != ftDeactivated) {
			properties.usingFilters = TRUE;
		} else {
			properties.usingFilters = FALSE;
			for(int i = 0; i < NUMBEROFFILTERS_TB; ++i) {
				if(properties.filterType[i] != ftDeactivated) {
					properties.usingFilters = TRUE;
					break;
				}
			}
		}
		return S_OK;
	}
	return E_INVALIDARG;
}

STDMETHODIMP ToolBarButtons::get_Item(LONG buttonIdentifier, ButtonIdentifierTypeConstants buttonIdentifierType/* = btitIndex*/, VARIANT_BOOL getChevronToolBarButton/* = VARIANT_FALSE*/, IToolBarButton** ppButton/* = NULL*/)
{
	ATLASSERT_POINTER(ppButton, IToolBarButton*);
	if(!ppButton) {
		return E_POINTER;
	}
	HWND hWndToUse = properties.GetTBarHWnd(VARIANTBOOL2BOOL(getChevronToolBarButton));
	if(!IsWindow(hWndToUse)) {
		return E_FAIL;
	}

	// retrieve the button's index
	int buttonIndex = -1;
	switch(buttonIdentifierType) {
		case btitID:
			if(properties.pOwnerTBar) {
				buttonIndex = properties.pOwnerTBar->IDToButtonIndex(buttonIdentifier, hWndToUse);
			}
			break;
		case btitIndex:
			buttonIndex = buttonIdentifier;
			break;
	}

	if(buttonIndex != -1) {
		if(IsPartOfCollection(buttonIndex, hWndToUse)) {
			ClassFactory::InitToolBarButton(buttonIndex, VARIANTBOOL2BOOL(getChevronToolBarButton), properties.pOwnerTBar, IID_IToolBarButton, reinterpret_cast<LPUNKNOWN*>(ppButton));
			if(*ppButton) {
				return S_OK;
			}
		}
	}

	// button not found
	if(buttonIdentifierType == btitIndex) {
		return DISP_E_BADINDEX;
	} else {
		return E_INVALIDARG;
	}
}

STDMETHODIMP ToolBarButtons::get__NewEnum(IUnknown** ppEnumerator)
{
	ATLASSERT_POINTER(ppEnumerator, LPUNKNOWN);
	if(!ppEnumerator) {
		return E_POINTER;
	}

	Reset();
	return QueryInterface(IID_IUnknown, reinterpret_cast<LPVOID*>(ppEnumerator));
}


STDMETHODIMP ToolBarButtons::Add(LONG id, LONG insertAt/* = -1*/, LONG iconIndex/* = -2*/, LONG imageListIndex/* = 0*/, BSTR text/* = L""*/, VARIANT_BOOL displayText/* = VARIANT_FALSE*/, VARIANT_BOOL useMnemonic/* = VARIANT_TRUE*/, ButtonTypeConstants buttonType/* = btyCommandButton*/, VARIANT_BOOL partOfGroup/* = VARIANT_FALSE*/, SelectionStateConstants selectionState/* = ssUnchecked*/, DropDownStyleConstants dropDownStyle/* = ddstNoDropDown*/, LONG buttonData/* = 0*/, VARIANT_BOOL visible/* = VARIANT_TRUE*/, VARIANT_BOOL enabled/* = VARIANT_TRUE*/, VARIANT_BOOL autoSize/* = VARIANT_TRUE*/, OLE_XSIZE_PIXELS width/* = -1*/, VARIANT_BOOL followedByLineBreak/* = VARIANT_FALSE*/, VARIANT_BOOL showingEllipsis/* = VARIANT_FALSE*/, VARIANT_BOOL marked/* = VARIANT_FALSE*/, IToolBarButton** ppAddedButton/* = NULL*/)
{
	ATLASSERT_POINTER(ppAddedButton, IToolBarButton*);
	if(!ppAddedButton) {
		return E_POINTER;
	}

	if(insertAt < -1 || iconIndex < -2 || imageListIndex < -1 || width < -1) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	HWND hWndTBar = properties.GetTBarHWnd();
	ATLASSERT(IsWindow(hWndTBar));

	HRESULT hr = E_FAIL;
	#ifndef UNICODE
		COLE2T converter(text);
	#endif

	TBBUTTON insertionData = {0};
	insertionData.idCommand = id;
	insertionData.dwData = buttonData;
	if(buttonType == btyPlaceholder) {
		// manually sized, disabled command button without text and icon
		if(properties.pOwnerTBar->IsPlaceholderButton(insertionData.idCommand)) {
			// duplicate id - raise VB runtime error 457
			return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 457);
		}
		insertionData.iBitmap = I_IMAGENONE;
		if(width >= 0) {
			insertionData.bReserved[0] = HIBYTE(LOWORD(width));
			insertionData.bReserved[1] = LOBYTE(LOWORD(width));
		}
		if(partOfGroup != VARIANT_FALSE) {
			insertionData.fsStyle |= BTNS_GROUP;
		}
		insertionData.fsStyle |= BTNS_BUTTON;
		if(visible == VARIANT_FALSE) {
			insertionData.fsState |= TBSTATE_HIDDEN;
		}
		if(followedByLineBreak != VARIANT_FALSE) {
			insertionData.fsState |= TBSTATE_WRAP;
		}
	} else {
		if(buttonType != btySeparator) {
			if(iconIndex == -2) {
				insertionData.iBitmap = I_IMAGENONE;
			} else if(iconIndex == -1 || imageListIndex == -1) {
				insertionData.iBitmap = I_IMAGECALLBACK;
			} else {
				insertionData.iBitmap = MAKELONG(LOWORD(iconIndex), LOWORD(imageListIndex));
			}
		}
		// TODO: Is LPSTR_TEXTCALLBACK really valid? No!
		LPTSTR pText = NULL;
		if(text) {
			if(SysStringLen(text) > 0) {     // This is important! Adding a button fails if iString points to an empty string.
				#ifdef UNICODE
					pText = OLE2W(text);
				#else
					pText = converter;
				#endif
			}
		} else {
			pText = LPSTR_TEXTCALLBACK;
		}
		insertionData.iString = reinterpret_cast<INT_PTR>(pText);
		if(width >= 0) {
			if(buttonType == btySeparator) {
				insertionData.iBitmap = width;
			} else {
				insertionData.bReserved[0] = HIBYTE(LOWORD(width));
				insertionData.bReserved[1] = LOBYTE(LOWORD(width));
			}
		}
		if(displayText != VARIANT_FALSE) {
			insertionData.fsStyle |= BTNS_SHOWTEXT;
		}
		if(useMnemonic == VARIANT_FALSE) {
			insertionData.fsStyle |= BTNS_NOPREFIX;
		}
		if(partOfGroup != VARIANT_FALSE) {
			insertionData.fsStyle |= BTNS_GROUP;
		}
		if(autoSize != VARIANT_FALSE) {
			insertionData.fsStyle |= BTNS_AUTOSIZE;
		}
		switch(buttonType) {
			case btyCommandButton:
				insertionData.fsStyle |= BTNS_BUTTON;
				break;
			case btyCheckButton:
				insertionData.fsStyle |= BTNS_CHECK;
				break;
			case btySeparator:
				insertionData.fsStyle |= BTNS_SEP;
				break;
		}
		switch(dropDownStyle) {
			case ddstNormal:
				insertionData.fsStyle |= BTNS_DROPDOWN;
				break;
			case ddstAlwaysWholeButton:
				insertionData.fsStyle |= BTNS_WHOLEDROPDOWN;
				break;
		}
		if(visible == VARIANT_FALSE) {
			insertionData.fsState |= TBSTATE_HIDDEN;
		}
		if(enabled != VARIANT_FALSE) {
			insertionData.fsState |= TBSTATE_ENABLED;
		}
		if(followedByLineBreak != VARIANT_FALSE) {
			insertionData.fsState |= TBSTATE_WRAP;
		}
		if(showingEllipsis != VARIANT_FALSE) {
			insertionData.fsState |= TBSTATE_ELLIPSES;
		}
		if(marked != VARIANT_FALSE) {
			insertionData.fsState |= TBSTATE_MARKED;
		}
		switch(selectionState) {
			case ssChecked:
				insertionData.fsState |= TBSTATE_CHECKED;
				break;
			case ssIndeterminate:
				insertionData.fsState |= TBSTATE_INDETERMINATE;
				break;
		}
	}
	if(buttonType == btyPlaceholder) {
		properties.pOwnerTBar->RegisterPlaceholderButton(insertionData.idCommand);
	}

	int insertedButton = -1;
	BOOL succeeded;
	if(insertAt == -1) {
		succeeded = SendMessage(hWndTBar, TB_ADDBUTTONS, 1, reinterpret_cast<LPARAM>(&insertionData));
		if(succeeded) {
			insertAt = SendMessage(hWndTBar, TB_BUTTONCOUNT, 0, 0) - 1;
		}
	} else {
		succeeded = SendMessage(hWndTBar, TB_INSERTBUTTON, insertAt, reinterpret_cast<LPARAM>(&insertionData));
	}
	if(succeeded) {
		insertedButton = insertAt;

		if(width >= 0) {
			// seems like the width has to be set separatly
			TBBUTTONINFO button = {0};
			button.cbSize = sizeof(button);
			// NOTE: TBIF_BYINDEX requires comctl32.dll 5.80.
			button.dwMask = TBIF_BYINDEX | TBIF_SIZE;
			button.cx = static_cast<WORD>(width);
			SendMessage(hWndTBar, TB_SETBUTTONINFO, insertedButton, reinterpret_cast<LPARAM>(&button));
		}

		ClassFactory::InitToolBarButton(insertedButton, FALSE, properties.pOwnerTBar, IID_IToolBarButton, reinterpret_cast<LPUNKNOWN*>(ppAddedButton));
		hr = S_OK;
	} else if(buttonType == btyPlaceholder) {
		properties.pOwnerTBar->DeregisterPlaceholderButton(insertionData.idCommand);
	}
	return hr;
}

STDMETHODIMP ToolBarButtons::Contains(LONG buttonIdentifier, ButtonIdentifierTypeConstants buttonIdentifierType/* = btitIndex*/, VARIANT_BOOL checkChevronToolBarButton/* = VARIANT_FALSE*/, VARIANT_BOOL* pValue/* = NULL*/)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}
	HWND hWndToUse = properties.GetTBarHWnd(VARIANTBOOL2BOOL(checkChevronToolBarButton));
	if(!IsWindow(hWndToUse)) {
		return E_FAIL;
	}

	// retrieve the button's index
	int buttonIndex = -1;
	switch(buttonIdentifierType) {
		case btitID:
			if(properties.pOwnerTBar) {
				buttonIndex = properties.pOwnerTBar->IDToButtonIndex(buttonIdentifier, hWndToUse);
			}
			break;
		case btitIndex:
			buttonIndex = buttonIdentifier;
			break;
	}

	*pValue = BOOL2VARIANTBOOL(IsPartOfCollection(buttonIndex, hWndToUse));
	return S_OK;
}

STDMETHODIMP ToolBarButtons::Count(VARIANT_BOOL countChevronToolBarButtons/* = VARIANT_FALSE*/, LONG* pValue/* = NULL*/)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndTBar = properties.GetTBarHWnd(VARIANTBOOL2BOOL(countChevronToolBarButtons));
	if(!IsWindow(hWndTBar)) {
		return E_FAIL;
	}

	if(!properties.usingFilters) {
		*pValue = SendMessage(hWndTBar, TB_BUTTONCOUNT, 0, 0);
		return S_OK;
	}

	// count the buttons manually
	*pValue = 0;
	int numberOfButtons = SendMessage(hWndTBar, TB_BUTTONCOUNT, 0, 0);
	int buttonIndex = GetFirstButtonToProcess(numberOfButtons);
	while(buttonIndex != -1) {
		if(IsPartOfCollection(buttonIndex, hWndTBar)) {
			++(*pValue);
		}
		buttonIndex = GetNextButtonToProcess(buttonIndex, numberOfButtons);
	}
	return S_OK;
}

STDMETHODIMP ToolBarButtons::Remove(LONG buttonIdentifier, ButtonIdentifierTypeConstants buttonIdentifierType/* = btitIndex*/)
{
	HWND hWndTBar = properties.GetTBarHWnd();
	ATLASSERT(IsWindow(hWndTBar));

	// retrieve the button's index
	int buttonIndex = -1;
	switch(buttonIdentifierType) {
		case btitID:
			if(properties.pOwnerTBar) {
				buttonIndex = properties.pOwnerTBar->IDToButtonIndex(buttonIdentifier);
			}
			break;
		case btitIndex:
			buttonIndex = buttonIdentifier;
			break;
	}

	if(buttonIndex != -1) {
		if(IsPartOfCollection(buttonIndex)) {
			if(SendMessage(hWndTBar, TB_DELETEBUTTON, buttonIndex, 0)) {
				return S_OK;
			}
		}
	} else {
		// button not found
		if(buttonIdentifierType == btitIndex) {
			return DISP_E_BADINDEX;
		} else {
			return E_INVALIDARG;
		}
	}

	return E_FAIL;
}

STDMETHODIMP ToolBarButtons::RemoveAll(void)
{
	HWND hWndTBar = properties.GetTBarHWnd();
	ATLASSERT(IsWindow(hWndTBar));

	int numberOfButtons = static_cast<int>(SendMessage(hWndTBar, TB_BUTTONCOUNT, 0, 0));
	if(numberOfButtons == -1) {
		return E_FAIL;
	}

	if(!properties.usingFilters) {
		for(int i = numberOfButtons - 1; i >= 0; --i) {
			if(!SendMessage(hWndTBar, TB_DELETEBUTTON, i, 0)) {
				return E_FAIL;
			}
		}
	}

	// find the buttons to remove manually
	#ifdef USE_STL
		std::list<int> buttonsToRemove;
	#else
		CAtlList<int> buttonsToRemove;
	#endif
	int buttonIndex = GetFirstButtonToProcess(numberOfButtons);
	while(buttonIndex != -1) {
		if(IsPartOfCollection(buttonIndex, hWndTBar)) {
			#ifdef USE_STL
				buttonsToRemove.push_back(buttonIndex);
			#else
				buttonsToRemove.AddTail(buttonIndex);
			#endif
		}
		buttonIndex = GetNextButtonToProcess(buttonIndex, numberOfButtons);
	}
	return RemoveButtons(buttonsToRemove, hWndTBar);
}
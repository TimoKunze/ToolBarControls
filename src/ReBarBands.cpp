// ReBarBands.cpp: Manages a collection of ReBarBand objects

#include "stdafx.h"
#include "ReBarBands.h"
#include "ClassFactory.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP ReBarBands::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IReBarBands, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


//////////////////////////////////////////////////////////////////////
// implementation of IEnumVARIANT
STDMETHODIMP ReBarBands::Clone(IEnumVARIANT** ppEnumerator)
{
	ATLASSERT_POINTER(ppEnumerator, LPENUMVARIANT);
	if(!ppEnumerator) {
		return E_POINTER;
	}

	*ppEnumerator = NULL;
	CComObject<ReBarBands>* pRBarBandsObj = NULL;
	CComObject<ReBarBands>::CreateInstance(&pRBarBandsObj);
	pRBarBandsObj->AddRef();

	// clone all settings
	properties.CopyTo(&pRBarBandsObj->properties);

	pRBarBandsObj->QueryInterface(IID_IEnumVARIANT, reinterpret_cast<LPVOID*>(ppEnumerator));
	pRBarBandsObj->Release();
	return S_OK;
}

STDMETHODIMP ReBarBands::Next(ULONG numberOfMaxBands, VARIANT* pBands, ULONG* pNumberOfBandsReturned)
{
	ATLASSERT_POINTER(pBands, VARIANT);
	if(!pBands) {
		return E_POINTER;
	}

	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	// check each band in the rebar
	int numberOfBands = static_cast<int>(SendMessage(hWndRBar, RB_GETBANDCOUNT, 0, 0));
	ULONG i = 0;
	for(i = 0; i < numberOfMaxBands; ++i) {
		VariantInit(&pBands[i]);

		do {
			if(properties.lastEnumeratedBand >= 0) {
				if(properties.lastEnumeratedBand < numberOfBands) {
					properties.lastEnumeratedBand = GetNextBandToProcess(properties.lastEnumeratedBand, numberOfBands);
				}
			} else {
				properties.lastEnumeratedBand = GetFirstBandToProcess(numberOfBands);
			}
			if(properties.lastEnumeratedBand >= numberOfBands) {
				properties.lastEnumeratedBand = -1;
			}
		} while((properties.lastEnumeratedBand != -1) && (!IsPartOfCollection(properties.lastEnumeratedBand, hWndRBar)));

		if(properties.lastEnumeratedBand != -1) {
			ClassFactory::InitReBarBand(properties.lastEnumeratedBand, properties.pOwnerRBar, IID_IDispatch, reinterpret_cast<LPUNKNOWN*>(&pBands[i].pdispVal));
			pBands[i].vt = VT_DISPATCH;
		} else {
			// there's nothing more to iterate
			break;
		}
	}
	if(pNumberOfBandsReturned) {
		*pNumberOfBandsReturned = i;
	}

	return (i == numberOfMaxBands ? S_OK : S_FALSE);
}

STDMETHODIMP ReBarBands::Reset(void)
{
	properties.lastEnumeratedBand = -1;
	return S_OK;
}

STDMETHODIMP ReBarBands::Skip(ULONG numberOfBandsToSkip)
{
	VARIANT dummy;
	ULONG numBandsReturned = 0;
	// we could skip all bands at once, but it's easier to skip them one after the other
	for(ULONG i = 1; i <= numberOfBandsToSkip; ++i) {
		HRESULT hr = Next(1, &dummy, &numBandsReturned);
		VariantClear(&dummy);
		if(hr != S_OK || numBandsReturned == 0) {
			// there're no more bands to skip, so don't try it anymore
			break;
		}
	}
	return S_OK;
}
// implementation of IEnumVARIANT
//////////////////////////////////////////////////////////////////////


BOOL ReBarBands::IsValidBooleanFilter(VARIANT& filter)
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

BOOL ReBarBands::IsValidIntegerFilter(VARIANT& filter, int minValue, int maxValue)
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

BOOL ReBarBands::IsValidIntegerFilter(VARIANT& filter, int minValue)
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

BOOL ReBarBands::IsValidIntegerFilter(VARIANT& filter)
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

BOOL ReBarBands::IsValidStringFilter(VARIANT& filter)
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

int ReBarBands::GetFirstBandToProcess(int numberOfBands)
{
	if(numberOfBands == 0) {
		return -1;
	}
	return 0;
}

int ReBarBands::GetNextBandToProcess(int previousBand, int numberOfBands)
{
	if(previousBand < numberOfBands - 1) {
		return previousBand + 1;
	}
	return -1;
}

BOOL ReBarBands::IsBooleanInSafeArray(LPSAFEARRAY pSafeArray, VARIANT_BOOL value, LPVOID pComparisonFunction)
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

BOOL ReBarBands::IsIntegerInSafeArray(LPSAFEARRAY pSafeArray, int value, LPVOID pComparisonFunction)
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

BOOL ReBarBands::IsStringInSafeArray(LPSAFEARRAY pSafeArray, BSTR value, LPVOID pComparisonFunction)
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

BOOL ReBarBands::IsPartOfCollection(int bandIndex, HWND hWndRBar/* = NULL*/)
{
	if(!hWndRBar) {
		hWndRBar = properties.GetRBarHWnd();
	}
	ATLASSERT(IsWindow(hWndRBar));

	if(!IsValidBandIndex(bandIndex, hWndRBar)) {
		return FALSE;
	}

	BOOL isPart = FALSE;
	// we declare this one here to avoid compiler warnings
	REBARBANDINFO band = {0};
	band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();

	if(properties.filterType[fpIndex] != ftDeactivated) {
		if(IsIntegerInSafeArray(properties.filter[fpIndex].parray, bandIndex, properties.comparisonFunction[fpIndex])) {
			if(properties.filterType[fpIndex] == ftExcluding) {
				goto Exit;
			}
		} else {
			if(properties.filterType[fpIndex] == ftIncluding) {
				goto Exit;
			}
		}
	}

	BOOL mustRetrieveBandData = FALSE;
	if(properties.filterType[fpIconIndex] != ftDeactivated) {
		band.fMask |= RBBIM_IMAGE;
		mustRetrieveBandData = TRUE;
	}
	if(properties.filterType[fpBandData] != ftDeactivated) {
		band.fMask |= RBBIM_LPARAM;
		mustRetrieveBandData = TRUE;
	}
	if(properties.filterType[fpText] != ftDeactivated) {
		band.fMask |= RBBIM_TEXT;
		band.cch = MAX_BANDTEXTLENGTH;
		band.lpText = reinterpret_cast<LPTSTR>(malloc((band.cch + 1) * sizeof(TCHAR)));
		mustRetrieveBandData = TRUE;
	}
	if(properties.filterType[fpAddMarginsAroundChild] != ftDeactivated || properties.filterType[fpFixedBackgroundBitmapOrigin] != ftDeactivated || properties.filterType[fpHideIfVertical] != ftDeactivated || properties.filterType[fpKeepInFirstRow] != ftDeactivated || properties.filterType[fpNewLine] != ftDeactivated || properties.filterType[fpResizable] != ftDeactivated || properties.filterType[fpShowTitle] != ftDeactivated || properties.filterType[fpSizingGripVisibility] != ftDeactivated || properties.filterType[fpUseChevron] != ftDeactivated || properties.filterType[fpVariableHeight] != ftDeactivated || properties.filterType[fpVisible] != ftDeactivated) {
		band.fMask |= RBBIM_STYLE;
		mustRetrieveBandData = TRUE;
	}
	if(properties.filterType[fpBackColor] != ftDeactivated || properties.filterType[fpForeColor] != ftDeactivated) {
		band.fMask |= RBBIM_COLORS;
		mustRetrieveBandData = TRUE;
	}
	if(properties.filterType[fpCurrentHeight] != ftDeactivated || properties.filterType[fpHeightChangeStepSize] != ftDeactivated || properties.filterType[fpMaximumHeight] != ftDeactivated || properties.filterType[fpMinimumHeight] != ftDeactivated || properties.filterType[fpMinimumWidth] != ftDeactivated) {
		band.fMask |= RBBIM_CHILDSIZE;
		mustRetrieveBandData = TRUE;
	}
	if(properties.filterType[fpCurrentWidth] != ftDeactivated) {
		band.fMask |= RBBIM_SIZE;
		mustRetrieveBandData = TRUE;
	}
	if(properties.filterType[fpHBackgroundBitmap] != ftDeactivated) {
		band.fMask |= RBBIM_BACKGROUND;
		mustRetrieveBandData = TRUE;
	}
	if(properties.filterType[fpHContainedWindow] != ftDeactivated) {
		band.fMask |= RBBIM_CHILD;
		mustRetrieveBandData = TRUE;
	}
	if(properties.filterType[fpID] != ftDeactivated) {
		band.fMask |= RBBIM_ID;
		mustRetrieveBandData = TRUE;
	}
	if(properties.filterType[fpIdealWidth] != ftDeactivated) {
		band.fMask |= RBBIM_IDEALSIZE;
		mustRetrieveBandData = TRUE;
	}
	if(properties.filterType[fpTitleWidth] != ftDeactivated) {
		band.fMask |= RBBIM_HEADERSIZE;
		mustRetrieveBandData = TRUE;
	}

	if(mustRetrieveBandData) {
		if(!SendMessage(hWndRBar, RB_GETBANDINFO, bandIndex, reinterpret_cast<LPARAM>(&band))) {
			goto Exit;
		}

		// apply the filters

		if(properties.filterType[fpID] != ftDeactivated) {
			if(IsIntegerInSafeArray(properties.filter[fpID].parray, static_cast<int>(band.wID), properties.comparisonFunction[fpID])) {
				if(properties.filterType[fpID] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpID] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpBandData] != ftDeactivated) {
			if(IsIntegerInSafeArray(properties.filter[fpBandData].parray, static_cast<int>(band.lParam), properties.comparisonFunction[fpBandData])) {
				if(properties.filterType[fpBandData] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpBandData] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpIdealWidth] != ftDeactivated) {
			if(IsIntegerInSafeArray(properties.filter[fpIdealWidth].parray, static_cast<int>(band.cxIdeal), properties.comparisonFunction[fpIdealWidth])) {
				if(properties.filterType[fpIdealWidth] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpIdealWidth] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpText] != ftDeactivated) {
			if(IsStringInSafeArray(properties.filter[fpText].parray, CComBSTR(band.lpText), properties.comparisonFunction[fpText])) {
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
			if(IsIntegerInSafeArray(properties.filter[fpIconIndex].parray, (band.iImage == -1 ? -2 : band.iImage), properties.comparisonFunction[fpIconIndex])) {
				if(properties.filterType[fpIconIndex] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpIconIndex] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpVisible] != ftDeactivated) {
			if(IsBooleanInSafeArray(properties.filter[fpVisible].parray, BOOL2VARIANTBOOL(!(band.fStyle & RBBS_HIDDEN)), properties.comparisonFunction[fpVisible])) {
				if(properties.filterType[fpVisible] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpVisible] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpHideIfVertical] != ftDeactivated) {
			if(IsBooleanInSafeArray(properties.filter[fpHideIfVertical].parray, BOOL2VARIANTBOOL((band.fStyle & RBBS_NOVERT) == RBBS_NOVERT), properties.comparisonFunction[fpHideIfVertical])) {
				if(properties.filterType[fpHideIfVertical] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpHideIfVertical] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpShowTitle] != ftDeactivated) {
			if(IsBooleanInSafeArray(properties.filter[fpShowTitle].parray, BOOL2VARIANTBOOL(!(band.fStyle & RBBS_HIDETITLE)), properties.comparisonFunction[fpShowTitle])) {
				if(properties.filterType[fpShowTitle] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpShowTitle] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpUseChevron] != ftDeactivated) {
			if(IsBooleanInSafeArray(properties.filter[fpUseChevron].parray, BOOL2VARIANTBOOL((band.fStyle & RBBS_USECHEVRON) == RBBS_USECHEVRON), properties.comparisonFunction[fpUseChevron])) {
				if(properties.filterType[fpUseChevron] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpUseChevron] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpNewLine] != ftDeactivated) {
			if(IsBooleanInSafeArray(properties.filter[fpNewLine].parray, BOOL2VARIANTBOOL((band.fStyle & RBBS_BREAK) == RBBS_BREAK), properties.comparisonFunction[fpNewLine])) {
				if(properties.filterType[fpNewLine] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpNewLine] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpSizingGripVisibility] != ftDeactivated) {
			SizingGripVisibilityConstants sgv = sgvAutomatic;
			if(band.fStyle & RBBS_NOGRIPPER) {
				sgv = sgvNever;
			} else if(band.fStyle & RBBS_GRIPPERALWAYS) {
				sgv = sgvAlways;
			}
			if(IsIntegerInSafeArray(properties.filter[fpSizingGripVisibility].parray, sgv, properties.comparisonFunction[fpSizingGripVisibility])) {
				if(properties.filterType[fpSizingGripVisibility] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpSizingGripVisibility] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpHContainedWindow] != ftDeactivated) {
			if(IsIntegerInSafeArray(properties.filter[fpHContainedWindow].parray, HandleToLong(band.hwndChild), properties.comparisonFunction[fpHContainedWindow])) {
				if(properties.filterType[fpHContainedWindow] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpHContainedWindow] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpCurrentHeight] != ftDeactivated) {
			if(IsIntegerInSafeArray(properties.filter[fpCurrentHeight].parray, band.cyChild, properties.comparisonFunction[fpCurrentHeight])) {
				if(properties.filterType[fpCurrentHeight] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpCurrentHeight] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpCurrentWidth] != ftDeactivated) {
			if(IsIntegerInSafeArray(properties.filter[fpCurrentWidth].parray, band.cx, properties.comparisonFunction[fpCurrentWidth])) {
				if(properties.filterType[fpCurrentWidth] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpCurrentWidth] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpMaximumHeight] != ftDeactivated) {
			if(IsIntegerInSafeArray(properties.filter[fpMaximumHeight].parray, band.cyMaxChild, properties.comparisonFunction[fpMaximumHeight])) {
				if(properties.filterType[fpMaximumHeight] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpMaximumHeight] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpMinimumHeight] != ftDeactivated) {
			if(IsIntegerInSafeArray(properties.filter[fpMinimumHeight].parray, band.cyMinChild, properties.comparisonFunction[fpMinimumHeight])) {
				if(properties.filterType[fpMinimumHeight] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpMinimumHeight] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpMinimumWidth] != ftDeactivated) {
			if(IsIntegerInSafeArray(properties.filter[fpMinimumWidth].parray, band.cxMinChild, properties.comparisonFunction[fpMinimumWidth])) {
				if(properties.filterType[fpMinimumWidth] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpMinimumWidth] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpTitleWidth] != ftDeactivated) {
			if(IsIntegerInSafeArray(properties.filter[fpTitleWidth].parray, band.cxHeader, properties.comparisonFunction[fpTitleWidth])) {
				if(properties.filterType[fpTitleWidth] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpTitleWidth] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpHBackgroundBitmap] != ftDeactivated) {
			if(IsIntegerInSafeArray(properties.filter[fpHBackgroundBitmap].parray, HandleToLong(band.hbmBack), properties.comparisonFunction[fpHBackgroundBitmap])) {
				if(properties.filterType[fpHBackgroundBitmap] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpHBackgroundBitmap] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpHeightChangeStepSize] != ftDeactivated) {
			if(IsIntegerInSafeArray(properties.filter[fpHeightChangeStepSize].parray, band.cyIntegral, properties.comparisonFunction[fpHeightChangeStepSize])) {
				if(properties.filterType[fpHeightChangeStepSize] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpHeightChangeStepSize] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpKeepInFirstRow] != ftDeactivated) {
			if(IsBooleanInSafeArray(properties.filter[fpKeepInFirstRow].parray, BOOL2VARIANTBOOL((band.fStyle & RBBS_TOPALIGN) == RBBS_TOPALIGN), properties.comparisonFunction[fpKeepInFirstRow])) {
				if(properties.filterType[fpKeepInFirstRow] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpKeepInFirstRow] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpResizable] != ftDeactivated) {
			if(IsBooleanInSafeArray(properties.filter[fpResizable].parray, BOOL2VARIANTBOOL(!(band.fStyle & RBBS_FIXEDSIZE)), properties.comparisonFunction[fpResizable])) {
				if(properties.filterType[fpResizable] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpResizable] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpVariableHeight] != ftDeactivated) {
			if(IsBooleanInSafeArray(properties.filter[fpVariableHeight].parray, BOOL2VARIANTBOOL((band.fStyle & RBBS_VARIABLEHEIGHT) == RBBS_VARIABLEHEIGHT), properties.comparisonFunction[fpVariableHeight])) {
				if(properties.filterType[fpVariableHeight] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpVariableHeight] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpFixedBackgroundBitmapOrigin] != ftDeactivated) {
			if(IsBooleanInSafeArray(properties.filter[fpFixedBackgroundBitmapOrigin].parray, BOOL2VARIANTBOOL((band.fStyle & RBBS_FIXEDBMP) == RBBS_FIXEDBMP), properties.comparisonFunction[fpFixedBackgroundBitmapOrigin])) {
				if(properties.filterType[fpFixedBackgroundBitmapOrigin] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpFixedBackgroundBitmapOrigin] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpAddMarginsAroundChild] != ftDeactivated) {
			if(IsBooleanInSafeArray(properties.filter[fpAddMarginsAroundChild].parray, BOOL2VARIANTBOOL((band.fStyle & RBBS_CHILDEDGE) == RBBS_CHILDEDGE), properties.comparisonFunction[fpAddMarginsAroundChild])) {
				if(properties.filterType[fpAddMarginsAroundChild] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpAddMarginsAroundChild] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpBackColor] != ftDeactivated) {
			if(IsIntegerInSafeArray(properties.filter[fpBackColor].parray, (band.clrBack == CLR_DEFAULT ? static_cast<COLORREF>(SendMessage(hWndRBar, RB_GETBKCOLOR, 0, 0)) : band.clrBack), properties.comparisonFunction[fpBackColor])) {
				if(properties.filterType[fpBackColor] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpBackColor] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpForeColor] != ftDeactivated) {
			if(IsIntegerInSafeArray(properties.filter[fpForeColor].parray, (band.clrFore == CLR_DEFAULT ? static_cast<COLORREF>(SendMessage(hWndRBar, RB_GETTEXTCOLOR, 0, 0)) : band.clrFore), properties.comparisonFunction[fpForeColor])) {
				if(properties.filterType[fpForeColor] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpForeColor] == ftIncluding) {
					goto Exit;
				}
			}
		}
	}
	isPart = TRUE;

Exit:
	if(band.lpText) {
		SECUREFREE(band.lpText);
	}
	return isPart;
}

void ReBarBands::OptimizeFilter(FilteredPropertyConstants filteredProperty)
{
	if(filteredProperty != fpAddMarginsAroundChild && filteredProperty != fpFixedBackgroundBitmapOrigin && filteredProperty != fpHideIfVertical && filteredProperty != fpKeepInFirstRow && filteredProperty != fpNewLine && filteredProperty != fpResizable && filteredProperty != fpShowTitle && filteredProperty != fpUseChevron && filteredProperty != fpVariableHeight && filteredProperty != fpVisible) {
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
	HRESULT ReBarBands::RemoveBands(std::list<int>& bandsToRemove, HWND hWndRBar)
#else
	HRESULT ReBarBands::RemoveBands(CAtlList<int>& bandsToRemove, HWND hWndRBar)
#endif
{
	ATLASSERT(IsWindow(hWndRBar));

	CWindowEx2(hWndRBar).InternalSetRedraw(FALSE);
	// sort in reverse order
	#ifdef USE_STL
		bandsToRemove.sort(std::greater<int>());
		for(std::list<int>::const_iterator iter = bandsToRemove.begin(); iter != bandsToRemove.end(); ++iter) {
			SendMessage(hWndRBar, RB_DELETEBAND, *iter, 0);
		}
	#else
		// perform a crude bubble sort
		for(size_t j = 0; j < bandsToRemove.GetCount(); ++j) {
			for(size_t i = 0; i < bandsToRemove.GetCount() - 1; ++i) {
				if(bandsToRemove.GetAt(bandsToRemove.FindIndex(i)) < bandsToRemove.GetAt(bandsToRemove.FindIndex(i + 1))) {
					bandsToRemove.SwapElements(bandsToRemove.FindIndex(i), bandsToRemove.FindIndex(i + 1));
				}
			}
		}

		for(size_t i = 0; i < bandsToRemove.GetCount(); ++i) {
			SendMessage(hWndRBar, RB_DELETEBAND, bandsToRemove.GetAt(bandsToRemove.FindIndex(i)), 0);
		}
	#endif
	CWindowEx2(hWndRBar).InternalSetRedraw(TRUE);

	return S_OK;
}


ReBarBands::Properties::~Properties()
{
	for(int i = 0; i < NUMBEROFFILTERS_RB; ++i) {
		VariantClear(&filter[i]);
	}
	if(pOwnerRBar) {
		pOwnerRBar->Release();
	}
}

void ReBarBands::Properties::CopyTo(ReBarBands::Properties* pTarget)
{
	ATLASSERT_POINTER(pTarget, Properties);
	if(pTarget) {
		pTarget->pOwnerRBar = this->pOwnerRBar;
		if(pTarget->pOwnerRBar) {
			pTarget->pOwnerRBar->AddRef();
		}
		pTarget->lastEnumeratedBand = this->lastEnumeratedBand;
		pTarget->caseSensitiveFilters = this->caseSensitiveFilters;

		for(int i = 0; i < NUMBEROFFILTERS_RB; ++i) {
			VariantCopy(&pTarget->filter[i], &this->filter[i]);
			pTarget->filterType[i] = this->filterType[i];
			pTarget->comparisonFunction[i] = this->comparisonFunction[i];
		}
		pTarget->usingFilters = this->usingFilters;
	}
}

HWND ReBarBands::Properties::GetRBarHWnd(void)
{
	ATLASSUME(pOwnerRBar);

	OLE_HANDLE handle = NULL;
	pOwnerRBar->get_hWnd(&handle);
	return static_cast<HWND>(LongToHandle(handle));
}


void ReBarBands::SetOwner(ReBar* pOwner)
{
	if(properties.pOwnerRBar) {
		properties.pOwnerRBar->Release();
	}
	properties.pOwnerRBar = pOwner;
	if(properties.pOwnerRBar) {
		properties.pOwnerRBar->AddRef();
	}
}


STDMETHODIMP ReBarBands::get_CaseSensitiveFilters(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.caseSensitiveFilters);
	return S_OK;
}

STDMETHODIMP ReBarBands::put_CaseSensitiveFilters(VARIANT_BOOL newValue)
{
	properties.caseSensitiveFilters = VARIANTBOOL2BOOL(newValue);
	return S_OK;
}

STDMETHODIMP ReBarBands::get_ComparisonFunction(FilteredPropertyConstants filteredProperty, LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	switch(filteredProperty) {
		case fpIconIndex:
		case fpIndex:
		case fpBandData:
		case fpText:
			*pValue = static_cast<LONG>(reinterpret_cast<LONG_PTR>(properties.comparisonFunction[filteredProperty]));
			return S_OK;
			break;
		default:
			if(filteredProperty >= fpAddMarginsAroundChild && filteredProperty <= fpVisible) {
				*pValue = static_cast<LONG>(reinterpret_cast<LONG_PTR>(properties.comparisonFunction[filteredProperty]));
				return S_OK;
			}
			break;
	}
	return E_INVALIDARG;
}

STDMETHODIMP ReBarBands::put_ComparisonFunction(FilteredPropertyConstants filteredProperty, LONG newValue)
{
	switch(filteredProperty) {
		case fpIconIndex:
		case fpIndex:
		case fpBandData:
		case fpText:
			properties.comparisonFunction[filteredProperty] = reinterpret_cast<LPVOID>(static_cast<LONG_PTR>(newValue));
			return S_OK;
			break;
		default:
			if(filteredProperty >= fpAddMarginsAroundChild && filteredProperty <= fpVisible) {
				properties.comparisonFunction[filteredProperty] = reinterpret_cast<LPVOID>(static_cast<LONG_PTR>(newValue));
				return S_OK;
			}
			break;
	}
	return E_INVALIDARG;
}

STDMETHODIMP ReBarBands::get_Filter(FilteredPropertyConstants filteredProperty, VARIANT* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT);
	if(!pValue) {
		return E_POINTER;
	}

	switch(filteredProperty) {
		case fpIconIndex:
		case fpIndex:
		case fpBandData:
		case fpText:
			VariantClear(pValue);
			VariantCopy(pValue, &properties.filter[filteredProperty]);
			return S_OK;
			break;
		default:
			if(filteredProperty >= fpAddMarginsAroundChild && filteredProperty <= fpVisible) {
				VariantClear(pValue);
				VariantCopy(pValue, &properties.filter[filteredProperty]);
				return S_OK;
			}
			break;
	}
	return E_INVALIDARG;
}

STDMETHODIMP ReBarBands::put_Filter(FilteredPropertyConstants filteredProperty, VARIANT newValue)
{
	// check 'newValue'
	switch(filteredProperty) {
		case fpIconIndex:
			if(!IsValidIntegerFilter(newValue, -2)) {
				// invalid value - raise VB runtime error 380
				return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
			}
			break;
		case fpIndex:
		case fpCurrentHeight:
		case fpCurrentWidth:
		case fpHeightChangeStepSize:
		case fpIdealWidth:
		case fpMaximumHeight:
		case fpMinimumHeight:
		case fpMinimumWidth:
		case fpTitleWidth:
			if(!IsValidIntegerFilter(newValue, 0)) {
				// invalid value - raise VB runtime error 380
				return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
			}
			break;
		case fpSizingGripVisibility:
			if(!IsValidIntegerFilter(newValue, 0, 2)) {
				// invalid value - raise VB runtime error 380
				return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
			}
			break;
		case fpBandData:
		case fpBackColor:
		case fpForeColor:
		case fpHBackgroundBitmap:
		case fpHContainedWindow:
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
		case fpAddMarginsAroundChild:
		case fpFixedBackgroundBitmapOrigin:
		case fpHideIfVertical:
		case fpKeepInFirstRow:
		case fpNewLine:
		case fpResizable:
		case fpShowTitle:
		case fpUseChevron:
		case fpVariableHeight:
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

STDMETHODIMP ReBarBands::get_FilterType(FilteredPropertyConstants filteredProperty, FilterTypeConstants* pValue)
{
	ATLASSERT_POINTER(pValue, FilterTypeConstants);
	if(!pValue) {
		return E_POINTER;
	}

	switch(filteredProperty) {
		case fpIconIndex:
		case fpIndex:
		case fpBandData:
		case fpText:
			*pValue = properties.filterType[filteredProperty];
			return S_OK;
			break;
		default:
			if(filteredProperty >= fpAddMarginsAroundChild && filteredProperty <= fpVisible) {
				*pValue = properties.filterType[filteredProperty];
				return S_OK;
			}
			break;
	}
	return E_INVALIDARG;
}

STDMETHODIMP ReBarBands::put_FilterType(FilteredPropertyConstants filteredProperty, FilterTypeConstants newValue)
{
	if(newValue < 0 || newValue > 2) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	BOOL isOkay = FALSE;
	switch(filteredProperty) {
		case fpIconIndex:
		case fpIndex:
		case fpBandData:
		case fpText:
			isOkay = TRUE;
			break;
		default:
			isOkay = (filteredProperty >= fpAddMarginsAroundChild && filteredProperty <= fpVisible);
			break;
	}
	if(isOkay) {
		properties.filterType[filteredProperty] = newValue;
		if(newValue != ftDeactivated) {
			properties.usingFilters = TRUE;
		} else {
			properties.usingFilters = FALSE;
			for(int i = 0; i < NUMBEROFFILTERS_RB; ++i) {
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

STDMETHODIMP ReBarBands::get_Item(LONG bandIdentifier, BandIdentifierTypeConstants bandIdentifierType/* = bitIndex*/, IReBarBand** ppBand/* = NULL*/)
{
	ATLASSERT_POINTER(ppBand, IReBarBand*);
	if(!ppBand) {
		return E_POINTER;
	}

	// retrieve the band's index
	int bandIndex = -1;
	switch(bandIdentifierType) {
		case bitID:
			if(properties.pOwnerRBar) {
				bandIndex = properties.pOwnerRBar->IDToBandIndex(bandIdentifier);
			}
			break;
		case bitIndex:
			bandIndex = bandIdentifier;
			break;
	}

	if(bandIndex != -1) {
		if(IsPartOfCollection(bandIndex)) {
			ClassFactory::InitReBarBand(bandIndex, properties.pOwnerRBar, IID_IReBarBand, reinterpret_cast<LPUNKNOWN*>(ppBand));
			if(*ppBand) {
				return S_OK;
			}
		}
	}

	// band not found
	if(bandIdentifierType == bitIndex) {
		return DISP_E_BADINDEX;
	} else {
		return E_INVALIDARG;
	}
}

STDMETHODIMP ReBarBands::get__NewEnum(IUnknown** ppEnumerator)
{
	ATLASSERT_POINTER(ppEnumerator, LPUNKNOWN);
	if(!ppEnumerator) {
		return E_POINTER;
	}

	Reset();
	return QueryInterface(IID_IUnknown, reinterpret_cast<LPVOID*>(ppEnumerator));
}


STDMETHODIMP ReBarBands::Add(LONG insertAt/* = -1*/, OLE_HANDLE hContainedWindow/* = NULL*/, VARIANT_BOOL newLine/* = VARIANT_FALSE*/, SizingGripVisibilityConstants sizingGripVisibility/* = sgvAutomatic*/, VARIANT_BOOL resizable/* = VARIANT_TRUE*/, VARIANT_BOOL keepInFirstRow/* = VARIANT_FALSE*/, VARIANT_BOOL visible/* = VARIANT_TRUE*/, OLE_XSIZE_PIXELS idealWidth/* = 0*/, OLE_XSIZE_PIXELS minimumWidth/* = -1*/, OLE_YSIZE_PIXELS minimumHeight/* = -1*/, OLE_YSIZE_PIXELS maximumHeight/* = -1*/, OLE_YSIZE_PIXELS heightChangeStepSize/* = -1*/, VARIANT_BOOL variableHeight/* = VARIANT_FALSE*/, VARIANT_BOOL showTitle/* = VARIANT_TRUE*/, BSTR text/* = L""*/, LONG iconIndex/* = -2*/, LONG bandData/* = 0*/, OLE_XSIZE_PIXELS titleWidth/* = -1*/, VARIANT_BOOL hideIfVertical/* = VARIANT_FALSE*/, VARIANT_BOOL addMarginsAroundChild/* = VARIANT_FALSE*/, IReBarBand** ppAddedBand/* = NULL*/)
{
	ATLASSERT_POINTER(ppAddedBand, IReBarBand*);
	if(!ppAddedBand) {
		return E_POINTER;
	}

	if(insertAt < -1 || idealWidth < 0 || minimumWidth < -1 || minimumHeight < -1 || maximumHeight < -1 || heightChangeStepSize < -1 || titleWidth < -1) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	HRESULT hr = E_FAIL;

	REBARBANDINFO insertionData = {0};
	insertionData.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
	insertionData.cxMinChild = static_cast<UINT>(-1);
	insertionData.cyChild = static_cast<UINT>(-1);
	insertionData.cyIntegral = static_cast<UINT>(-1);
	insertionData.cyMaxChild = static_cast<UINT>(-1);
	insertionData.cyMinChild = static_cast<UINT>(-1);
	insertionData.fMask = RBBIM_IDEALSIZE | RBBIM_LPARAM | RBBIM_STYLE | RBBIM_TEXT | RBBIM_SETID;
	insertionData.cxIdeal = idealWidth;
	insertionData.lParam = bandData;
	#ifdef UNICODE
		insertionData.lpText = OLE2W(text);
	#else
		COLE2T converter(text);
		insertionData.lpText = converter;
	#endif
	if(newLine != VARIANT_FALSE) {
		insertionData.fStyle |= RBBS_BREAK;
	}
	switch(sizingGripVisibility) {
		case sgvNever:
			insertionData.fStyle |= RBBS_NOGRIPPER;
			break;
		case sgvAlways:
			insertionData.fStyle |= RBBS_GRIPPERALWAYS;
			break;
	}
	if(resizable == VARIANT_FALSE) {
		insertionData.fStyle |= RBBS_FIXEDSIZE;
	}
	if(keepInFirstRow != VARIANT_FALSE) {
		insertionData.fStyle |= RBBS_TOPALIGN;
	}
	if(visible == VARIANT_FALSE) {
		insertionData.fStyle |= RBBS_HIDDEN;
	}
	if(variableHeight != VARIANT_FALSE) {
		insertionData.fStyle |= RBBS_VARIABLEHEIGHT;
	}
	if(showTitle == VARIANT_FALSE) {
		insertionData.fStyle |= RBBS_HIDETITLE;
	}
	if(hideIfVertical != VARIANT_FALSE) {
		insertionData.fStyle |= RBBS_NOVERT;
	}
	if(addMarginsAroundChild != VARIANT_FALSE) {
		insertionData.fStyle |= RBBS_CHILDEDGE;
	}
	if(hContainedWindow) {
		insertionData.fMask |= RBBIM_CHILD;
		insertionData.hwndChild = static_cast<HWND>(LongToHandle(hContainedWindow));
	}
	if(minimumWidth > -1) {
		insertionData.fMask |= RBBIM_CHILDSIZE;
		insertionData.cxMinChild = minimumWidth;
	}
	if(minimumHeight > -1) {
		insertionData.fMask |= RBBIM_CHILDSIZE;
		insertionData.cyMinChild = minimumHeight;
		insertionData.cyChild = minimumHeight;
	}
	if(maximumHeight > -1) {
		insertionData.fMask |= RBBIM_CHILDSIZE;
		insertionData.cyMaxChild = maximumHeight;
	}
	if(heightChangeStepSize > -1) {
		insertionData.fMask |= RBBIM_CHILDSIZE;
		insertionData.cyIntegral = heightChangeStepSize;
	}
	if(iconIndex > -1) {
		insertionData.fMask |= RBBIM_IMAGE;
		insertionData.iImage = iconIndex;
	}
	if(titleWidth > -1) {
		insertionData.fMask |= RBBIM_HEADERSIZE;
		insertionData.cxHeader = titleWidth;
	}

	int insertedBand = -1;
	if(SendMessage(hWndRBar, RB_INSERTBAND, insertAt, reinterpret_cast<LPARAM>(&insertionData))) {
		ATLASSERT(insertionData.fMask & RBBIM_ID);
		insertedBand = properties.pOwnerRBar->IDToBandIndex(insertionData.wID);
		ClassFactory::InitReBarBand(insertedBand, properties.pOwnerRBar, IID_IReBarBand, reinterpret_cast<LPUNKNOWN*>(ppAddedBand));
		hr = S_OK;
	}
	return hr;
}

STDMETHODIMP ReBarBands::Contains(LONG bandIdentifier, BandIdentifierTypeConstants bandIdentifierType/* = bitIndex*/, VARIANT_BOOL* pValue/* = NULL*/)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	// retrieve the band's index
	int bandIndex = -1;
	switch(bandIdentifierType) {
		case bitID:
			if(properties.pOwnerRBar) {
				bandIndex = properties.pOwnerRBar->IDToBandIndex(bandIdentifier);
			}
			break;
		case bitIndex:
			bandIndex = bandIdentifier;
			break;
	}

	*pValue = BOOL2VARIANTBOOL(IsPartOfCollection(bandIndex));
	return S_OK;
}

STDMETHODIMP ReBarBands::Count(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	if(!properties.usingFilters) {
		*pValue = SendMessage(hWndRBar, RB_GETBANDCOUNT, 0, 0);
		return S_OK;
	}

	// count the bands manually
	*pValue = 0;
	int numberOfBands = SendMessage(hWndRBar, RB_GETBANDCOUNT, 0, 0);
	int bandIndex = GetFirstBandToProcess(numberOfBands);
	while(bandIndex != -1) {
		if(IsPartOfCollection(bandIndex, hWndRBar)) {
			++(*pValue);
		}
		bandIndex = GetNextBandToProcess(bandIndex, numberOfBands);
	}
	return S_OK;
}

STDMETHODIMP ReBarBands::Remove(LONG bandIdentifier, BandIdentifierTypeConstants bandIdentifierType/* = bitIndex*/)
{
	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	// retrieve the band's index
	int bandIndex = -1;
	switch(bandIdentifierType) {
		case bitID:
			if(properties.pOwnerRBar) {
				bandIndex = properties.pOwnerRBar->IDToBandIndex(bandIdentifier);
			}
			break;
		case bitIndex:
			bandIndex = bandIdentifier;
			break;
	}

	if(bandIndex != -1) {
		if(IsPartOfCollection(bandIndex)) {
			if(SendMessage(hWndRBar, RB_DELETEBAND, bandIndex, 0)) {
				return S_OK;
			}
		}
	} else {
		// band not found
		if(bandIdentifierType == bitIndex) {
			return DISP_E_BADINDEX;
		} else {
			return E_INVALIDARG;
		}
	}

	return E_FAIL;
}

STDMETHODIMP ReBarBands::RemoveAll(void)
{
	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	int numberOfBands = static_cast<int>(SendMessage(hWndRBar, RB_GETBANDCOUNT, 0, 0));
	if(numberOfBands == -1) {
		return E_FAIL;
	}

	if(!properties.usingFilters) {
		for(int i = numberOfBands - 1; i >= 0; --i) {
			if(!SendMessage(hWndRBar, RB_DELETEBAND, i, 0)) {
				return E_FAIL;
			}
		}
	}

	// find the bands to remove manually
	#ifdef USE_STL
		std::list<int> bandsToRemove;
	#else
		CAtlList<int> bandsToRemove;
	#endif
	int bandIndex = GetFirstBandToProcess(numberOfBands);
	while(bandIndex != -1) {
		if(IsPartOfCollection(bandIndex, hWndRBar)) {
			#ifdef USE_STL
				bandsToRemove.push_back(bandIndex);
			#else
				bandsToRemove.AddTail(bandIndex);
			#endif
		}
		bandIndex = GetNextBandToProcess(bandIndex, numberOfBands);
	}
	return RemoveBands(bandsToRemove, hWndRBar);
}
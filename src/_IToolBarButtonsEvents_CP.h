//////////////////////////////////////////////////////////////////////
/// \class Proxy_IToolBarButtonsEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IToolBarButtonsEvents interface</em>
///
/// \if UNICODE
///   \sa ToolBarButtons, TBarCtlsLibU::_IToolBarButtonsEvents
/// \else
///   \sa ToolBarButtons, TBarCtlsLibA::_IToolBarButtonsEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IToolBarButtonsEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IToolBarButtonsEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IToolBarButtonsEvents
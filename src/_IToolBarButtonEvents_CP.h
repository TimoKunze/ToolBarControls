//////////////////////////////////////////////////////////////////////
/// \class Proxy_IToolBarButtonEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IToolBarButtonEvents interface</em>
///
/// \if UNICODE
///   \sa ToolBarButton, TBarCtlsLibU::_IToolBarButtonEvents
/// \else
///   \sa ToolBarButton, TBarCtlsLibA::_IToolBarButtonEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IToolBarButtonEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IToolBarButtonEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IToolBarButtonEvents
//////////////////////////////////////////////////////////////////////
/// \class Proxy_IVirtualToolBarButtonEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IVirtualToolBarButtonEvents interface</em>
///
/// \if UNICODE
///   \sa VirtualToolBarButton, TBarCtlsLibU::_IVirtualToolBarButtonEvents
/// \else
///   \sa VirtualToolBarButton, TBarCtlsLibA::_IVirtualToolBarButtonEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IVirtualToolBarButtonEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IVirtualToolBarButtonEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IVirtualToolBarButtonEvents
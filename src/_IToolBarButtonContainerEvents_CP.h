//////////////////////////////////////////////////////////////////////
/// \class Proxy_IToolBarButtonContainerEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IToolBarButtonContainerEvents interface</em>
///
/// \if UNICODE
///   \sa ToolBarButtonContainer, TBarCtlsLibU::_IToolBarButtonContainerEvents
/// \else
///   \sa ToolBarButtonContainer, TBarCtlsLibA::_IToolBarButtonContainerEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IToolBarButtonContainerEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IToolBarButtonContainerEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IToolBarButtonContainerEvents
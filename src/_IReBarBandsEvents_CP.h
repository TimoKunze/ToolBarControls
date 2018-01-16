//////////////////////////////////////////////////////////////////////
/// \class Proxy_IReBarBandsEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IReBarBandsEvents interface</em>
///
/// \if UNICODE
///   \sa ReBarBands, TBarCtlsLibU::_IReBarBandsEvents
/// \else
///   \sa ReBarBands, TBarCtlsLibA::_IReBarBandsEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IReBarBandsEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IReBarBandsEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IReBarBandsEvents
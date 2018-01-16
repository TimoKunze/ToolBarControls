//////////////////////////////////////////////////////////////////////
/// \class Proxy_IVirtualReBarBandEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IVirtualReBarBandEvents interface</em>
///
/// \if UNICODE
///   \sa VirtualReBarBand, TBarCtlsLibU::_IVirtualReBarBandEvents
/// \else
///   \sa VirtualReBarBand, TBarCtlsLibA::_IVirtualReBarBandEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IVirtualReBarBandEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IVirtualReBarBandEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IVirtualReBarBandEvents
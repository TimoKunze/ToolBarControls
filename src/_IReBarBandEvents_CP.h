//////////////////////////////////////////////////////////////////////
/// \class Proxy_IReBarBandEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IReBarBandEvents interface</em>
///
/// \if UNICODE
///   \sa IReBarBand, TBarCtlsLibU::_IReBarBandEvents
/// \else
///   \sa IReBarBand, TBarCtlsLibA::_IReBarBandEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IReBarBandEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IReBarBandEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IReBarBandEvents
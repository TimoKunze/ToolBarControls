//////////////////////////////////////////////////////////////////////
/// \class CSimpleStack
/// \author Timo "TimoSoft" Kunze
/// \brief <em>A derivative of \c ATL::CSimpleArray that can be used as a stack</em>
///
/// \sa ToolBar
//////////////////////////////////////////////////////////////////////


#pragma once

#include "stdafx.h"


template <class T>
class CSimpleStack :
    #ifdef USE_STL
      public std::vector<T>
    #else
      public CSimpleArray<T>
    #endif
{
public:
	/// \brief <em>Pushes a value onto the stack</em>
	///
	/// \param[in] t The value to add.
	///
	/// \return \c TRUE on success; otherwise \c FALSE.
	///
	/// \sa Pop, GetCurrent
	BOOL Push(T t)
	{
    #ifdef USE_STL
			push_back(t);
			return TRUE;
		#else
			return Add(t);
		#endif
	}

	/// \brief <em>Pops a value from the stack</em>
	///
	/// \return The popped value.
	///
	/// \sa Push, GetCurrent
	T Pop(void)
	{
    #ifdef USE_STL
			if(size() == 0) {
		#else
			int last = GetSize() - 1;
			if(last < 0) {
		#endif
			return NULL;
		}
    #ifdef USE_STL
			std::vector<T>::iterator last = begin() + (size() - 1);
			T t = *last;
			erase(last);
		#else
			T t = m_aT[last];
			if(!RemoveAt(last)) {
				return NULL;
			}
		#endif
		return t;
	}

	/// \brief <em>Retrieves the current top of the stack</em>
	///
	/// \return The current top of the stack.
	///
	/// \sa Push, Pop
	T GetCurrent(void)
	{
    #ifdef USE_STL
			std::vector<T>::reverse_iterator last = rbegin();
			if(last == rend()) {
				return NULL;
			}
			return *last;
		#else
			int last = GetSize() - 1;
			if(last < 0) {
				return NULL;
			}
			return m_aT[last];
		#endif
	}

	#ifdef USE_STL
		/// \brief <em>Returns the number of elements currently stored on the stack</em>
		///
		/// \return The number of elements.
		///
		/// \sa Push, Pop
		int GetSize(void)
		{
			return size();
		}
	#endif
};     // CSimpleStack
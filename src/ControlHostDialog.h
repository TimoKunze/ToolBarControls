//////////////////////////////////////////////////////////////////////
/// \class ControlHostDialog
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Implements a dialog for hosting a free-floating control</em>
///
/// This class implements the dialog used by \c ControlHostWindow to provide a host window for
/// free-floating controls.
///
/// \sa ControlHostWindow
//////////////////////////////////////////////////////////////////////


#pragma once
#include "atlwin.h"
#include "res/resource.h"
#ifdef UNICODE
	#include "TBarCtlsU.h"
#else
	#include "TBarCtlsA.h"
#endif
#include "helpers.h"


class ControlHostDialog :
	public CDialogImpl<ControlHostDialog>
{
public:

	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		enum { IDD = IDD_CONTROLHOSTWINDOW };

		BEGIN_MSG_MAP(ControlHostDialog)
			REFLECT_NOTIFICATIONS_EX()
		END_MSG_MAP()
	#endif
};
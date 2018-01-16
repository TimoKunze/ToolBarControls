//////////////////////////////////////////////////////////////////////
/// \class IButtonContainer
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Communication between a \c ToolBarButtonContainer object and its creator object</em>
///
/// This interface allows an \c ToolBar object to inform a \c ToolBarButtonContainer object
/// about button deletions.
///
/// \sa ToolBar::RegisterButtonContainer, ToolBarButtonContainer
//////////////////////////////////////////////////////////////////////


#pragma once


class IButtonContainer
{
public:
	/// \brief <em>A button was removed and the item container should check its content</em>
	///
	/// \param[in] buttonID The unique ID of the removed item. If 0, all items were removed.
	///
	/// \sa ToolBar::RemoveButtonFromButtonContainers
	virtual void RemovedButton(LONG buttonID) = 0;
	/// \brief <em>Retrieves the container's ID used to identify it</em>
	///
	/// \return The container's ID.
	///
	/// \sa ToolBar::DeregisterButtonContainer, containerID
	virtual DWORD GetID(void) = 0;

	/// \brief <em>Holds the container's ID</em>
	DWORD containerID;
};     // IButtonContainer
// DispIDs.h: Defines a DISPID for each COM property and method we're providing.

// IReBar
// properties
#define DISPID_RB_ALLOWBANDREORDERING													1
#define DISPID_RB_APPEARANCE																	2
#define DISPID_RB_AUTOUPDATELAYOUT														3
#define DISPID_RB_BACKCOLOR																		4
#define DISPID_RB_BANDS																				5
#define DISPID_RB_BORDERSTYLE																	6
#define DISPID_RB_CONTROLHEIGHT																7
#define DISPID_RB_DISABLEDEVENTS															8
#define DISPID_RB_DISPLAYBANDSEPARATORS												9
#define DISPID_RB_DISPLAYSPLITTER															10
#define DISPID_RB_DONTREDRAW																	11
#define DISPID_RB_ENABLED																			DISPID_ENABLED
#define DISPID_RB_FIXEDBANDHEIGHT															12
#define DISPID_RB_FONT																				13
#define DISPID_RB_FORECOLOR																		14
#define DISPID_RB_HIGHLIGHTCOLOR															15
#define DISPID_RB_HIMAGELIST																	16
#define DISPID_RB_HPALETTE																		17
#define DISPID_RB_HOVERTIME																		18
#define DISPID_RB_HWND																				DISPID_VALUE
#define DISPID_RB_HWNDNOTIFICATIONRECEIVER										19
//#define DISPID_RB_HWNDTOOLTIP																	20
#define DISPID_RB_MDIFRAMEMENUBAND														21
#define DISPID_RB_MOUSEICON																		22
#define DISPID_RB_MOUSEPOINTER																23
#define DISPID_RB_NATIVEDROPTARGET														24
#define DISPID_RB_NUMBEROFROWS																25
#define DISPID_RB_ORIENTATION																	26
#define DISPID_RB_REGISTERFOROLEDRAGDROP											27
#define DISPID_RB_REPLACEMDIFRAMEMENU													28
#define DISPID_RB_RIGHTTOLEFT																	29
#define DISPID_RB_SHADOWCOLOR																	30
#define DISPID_RB_SUPPORTOLEDRAGIMAGES												31
#define DISPID_RB_TOGGLEONDOUBLECLICK													32
#define DISPID_RB_USESYSTEMFONT																33
#define DISPID_RB_VERSION																			34
#define DISPID_RB_VERTICALSIZINGGRIPSONVERTICALORIENTATION		35
// fingerprint
#define DISPID_RB_APPID																				500
#define DISPID_RB_APPNAME																			501
#define DISPID_RB_APPSHORTNAME																502
#define DISPID_RB_BUILD																				503
#define DISPID_RB_CHARSET																			504
#define DISPID_RB_ISRELEASE																		505
#define DISPID_RB_PROGRAMMER																	506
#define DISPID_RB_TESTER																			507
// methods
#define DISPID_RB_ABOUT																				DISPID_ABOUTBOX
#define DISPID_RB_BEGINDRAGBAND																36
#define DISPID_RB_DRAGMOVEBAND																37
#define DISPID_RB_ENDDRAGBAND																	38
#define DISPID_RB_GETMARGINS																	39
#define DISPID_RB_HITTEST																			40
#define DISPID_RB_LOADSETTINGSFROMFILE												41
#define DISPID_RB_REFRESH																			DISPID_REFRESH
#define DISPID_RB_SAVESETTINGSTOFILE													42
#define DISPID_RB_SIZETORECTANGLE															43
#define DISPID_RB_FINISHOLEDRAGDROP														44


// _IReBarEvents
// methods
#define DISPID_RBE_AUTOBREAKINGBAND														1
#define DISPID_RBE_AUTOSIZED																	2
#define DISPID_RBE_BANDBEGINDRAG															3
#define DISPID_RBE_BANDBEGINRDRAG															4
#define DISPID_RBE_BANDENDDRAG																5
#define DISPID_RBE_BANDMOUSEENTER															6
#define DISPID_RBE_BANDMOUSELEAVE															7
#define DISPID_RBE_BEFOREDISPLAYMDICHILDSYSTEMMENU						8
#define DISPID_RBE_CHEVRONCLICK																9
#define DISPID_RBE_CLEANUPMDICHILDSYSTEMMENU									10
#define DISPID_RBE_CLICK																			11
#define DISPID_RBE_CONTEXTMENU																12
#define DISPID_RBE_CUSTOMDRAW																	13
#define DISPID_RBE_DBLCLICK																		14
#define DISPID_RBE_DESTROYEDCONTROLWINDOW											15
#define DISPID_RBE_DRAGGINGSPLITTER														16
#define DISPID_RBE_FREEBANDDATA																17
#define DISPID_RBE_HEIGHTCHANGED															18
#define DISPID_RBE_INSERTEDBAND																19
#define DISPID_RBE_INSERTINGBAND															20
#define DISPID_RBE_LAYOUTCHANGED															21
#define DISPID_RBE_MCLICK																			22
#define DISPID_RBE_MDBLCLICK																	23
#define DISPID_RBE_MOUSEDOWN																	24
#define DISPID_RBE_MOUSEENTER																	25
#define DISPID_RBE_MOUSEHOVER																	26
#define DISPID_RBE_MOUSELEAVE																	27
#define DISPID_RBE_MOUSEMOVE																	28
#define DISPID_RBE_MOUSEUP																		29
#define DISPID_RBE_NONCLIENTHITTEST														30
#define DISPID_RBE_OLEDRAGDROP																31
#define DISPID_RBE_OLEDRAGENTER																32
#define DISPID_RBE_OLEDRAGLEAVE																33
#define DISPID_RBE_OLEDRAGMOUSEMOVE														34
#define DISPID_RBE_RAWMENUMESSAGE															35
#define DISPID_RBE_RCLICK																			36
#define DISPID_RBE_RDBLCLICK																	37
#define DISPID_RBE_RECREATEDCONTROLWINDOW											38
#define DISPID_RBE_RELEASEDMOUSECAPTURE												39
#define DISPID_RBE_REMOVEDBAND																40
#define DISPID_RBE_REMOVINGBAND																41
#define DISPID_RBE_RESIZEDCONTROLWINDOW												42
#define DISPID_RBE_RESIZINGCONTAINEDWINDOW										43
#define DISPID_RBE_SELECTEDMENUITEM														44
#define DISPID_RBE_TOGGLINGBAND																45
#define DISPID_RBE_XCLICK																			46
#define DISPID_RBE_XDBLCLICK																	47


// IReBarBand
// properties
#define DISPID_RBB_ADDMARGINSAROUNDCHILD											1
#define DISPID_RBB_BACKCOLOR																	2
#define DISPID_RBB_BANDDATA																		3
#define DISPID_RBB_CHEVRONBUTTONOBJECTSTATE										4
#define DISPID_RBB_CURRENTHEIGHT															5
#define DISPID_RBB_CURRENTWIDTH																6
#define DISPID_RBB_FIXEDBACKGROUNDBITMAPORIGIN								7
#define DISPID_RBB_FORECOLOR																	8
#define DISPID_RBB_HBACKGROUNDBITMAP													9
#define DISPID_RBB_HCONTAINEDWINDOW														DISPID_VALUE
#define DISPID_RBB_HEIGHTCHANGESTEPSIZE												10
#define DISPID_RBB_HIDEIFVERTICAL															11
#define DISPID_RBB_ICONINDEX																	12
#define DISPID_RBB_ID																					13
#define DISPID_RBB_IDEALWIDTH																	14
#define DISPID_RBB_INDEX																			15
#define DISPID_RBB_KEEPINFIRSTROW															16
#define DISPID_RBB_MAXIMUMHEIGHT															17
#define DISPID_RBB_MINIMUMHEIGHT															18
#define DISPID_RBB_MINIMUMWIDTH																19
#define DISPID_RBB_NEWLINE																		20
#define DISPID_RBB_RESIZABLE																	21
#define DISPID_RBB_ROWHEIGHT																	22
#define DISPID_RBB_SHOWTITLE																	23
#define DISPID_RBB_SIZINGGRIPVISIBILITY												24
#define DISPID_RBB_TEXT																				25
#define DISPID_RBB_TITLEWIDTH																	26
#define DISPID_RBB_USECHEVRON																	27
#define DISPID_RBB_VARIABLEHEIGHT															28
#define DISPID_RBB_VISIBLE																		29
#define DISPID_RBB_CHEVRONHOT																	38
#define DISPID_RBB_CHEVRONPUSHED															39
#define DISPID_RBB_CHEVRONVISIBLE															40
// methods
#define DISPID_RBB_CLICKCHEVRON																30
#define DISPID_RBB_GETBORDERSIZES															31
#define DISPID_RBB_GETCHEVRONRECTANGLE												32
#define DISPID_RBB_GETRECTANGLE																33
#define DISPID_RBB_HIDE																				34
#define DISPID_RBB_MAXIMIZE																		35
#define DISPID_RBB_MINIMIZE																		36
#define DISPID_RBB_SHOW																				37


// IReBarBands
// properties
#define DISPID_RBBS_CASESENSITIVEFILTERS											1
#define DISPID_RBBS_COMPARISONFUNCTION												2
#define DISPID_RBBS_FILTER																		3
#define DISPID_RBBS_FILTERTYPE																4
#define DISPID_RBBS_ITEM																			DISPID_VALUE
#define DISPID_RBBS__NEWENUM																	DISPID_NEWENUM
// methods
#define DISPID_RBBS_ADD																				5
#define DISPID_RBBS_CONTAINS																	6
#define DISPID_RBBS_COUNT																			7
#define DISPID_RBBS_REMOVE																		8
#define DISPID_RBBS_REMOVEALL																	9


// IVirtualReBarBand
// properties
#define DISPID_VRBB_ADDMARGINSAROUNDCHILD											1
#define DISPID_VRBB_BACKCOLOR																	2
#define DISPID_VRBB_BANDDATA																	3
#define DISPID_VRBB_CHEVRONBUTTONOBJECTSTATE									4
#define DISPID_VRBB_CURRENTHEIGHT															5
#define DISPID_VRBB_CURRENTWIDTH															6
#define DISPID_VRBB_FIXEDBACKGROUNDBITMAPORIGIN								7
#define DISPID_VRBB_FORECOLOR																	8
#define DISPID_VRBB_HBACKGROUNDBITMAP													9
#define DISPID_VRBB_HCONTAINEDWINDOW													DISPID_VALUE
#define DISPID_VRBB_HEIGHTCHANGESTEPSIZE											10
#define DISPID_VRBB_HIDEIFVERTICAL														11
#define DISPID_VRBB_ICONINDEX																	12
#define DISPID_VRBB_ID																				13
#define DISPID_VRBB_IDEALWIDTH																14
#define DISPID_VRBB_INDEX																			15
#define DISPID_VRBB_KEEPINFIRSTROW														16
#define DISPID_VRBB_MAXIMUMHEIGHT															17
#define DISPID_VRBB_MINIMUMHEIGHT															18
#define DISPID_VRBB_MINIMUMWIDTH															19
#define DISPID_VRBB_NEWLINE																		20
#define DISPID_VRBB_RESIZABLE																	21
#define DISPID_VRBB_SHOWTITLE																	22
#define DISPID_VRBB_SIZINGGRIPVISIBILITY											23
#define DISPID_VRBB_TEXT																			24
#define DISPID_VRBB_TITLEWIDTH																25
#define DISPID_VRBB_USECHEVRON																26
#define DISPID_VRBB_VARIABLEHEIGHT														27
#define DISPID_VRBB_VISIBLE																		28
// methods
#define DISPID_VRBB_GETCHEVRONRECTANGLE												29


// IToolBar
// properties
#define DISPID_TB_ALLOWCUSTOMIZATION													1
#define DISPID_TB_ALWAYSDISPLAYBUTTONTEXT											2
#define DISPID_TB_ANCHORHIGHLIGHTING													3
#define DISPID_TB_APPEARANCE																	4
#define DISPID_TB_BACKSTYLE																		5
#define DISPID_TB_BORDERSTYLE																	6
#define DISPID_TB_BUTTONHEIGHT																7
#define DISPID_TB_BUTTONS																			8
#define DISPID_TB_BUTTONROWCOUNT															9
#define DISPID_TB_BUTTONSTYLE																	10
#define DISPID_TB_BUTTONTEXTPOSITION													11
#define DISPID_TB_BUTTONWIDTH																	12
#define DISPID_TB_COMMANDENABLED															13
#define DISPID_TB_DISABLEDEVENTS															14
#define DISPID_TB_DISPLAYMENUDIVIDER													15
#define DISPID_TB_DISPLAYPARTIALLYCLIPPEDBUTTONS							16
#define DISPID_TB_DONTREDRAW																	17
#define DISPID_TB_DRAGCLICKTIME																18
#define DISPID_TB_DRAGDROPCUSTOMIZATIONMODIFIERKEY						19
#define DISPID_TB_DRAGGEDBUTTONS															20
#define DISPID_TB_DROPDOWNGAP																	21
#define DISPID_TB_ENABLED																			DISPID_ENABLED
#define DISPID_TB_FIRSTBUTTONINDENTATION											22
#define DISPID_TB_FOCUSONCLICK																23
#define DISPID_TB_FONT																				24
#define DISPID_TB_HCHEVRONMENU																25
//#define DISPID_TB_HDRAGIMAGELIST															26
#define DISPID_TB_HIGHLIGHTCOLOR															27
#define DISPID_TB_HIMAGELIST																	28
#define DISPID_TB_HORIZONTALBUTTONPADDING											29
#define DISPID_TB_HORIZONTALBUTTONSPACING											30
#define DISPID_TB_HORIZONTALICONCAPTIONGAP										31
#define DISPID_TB_HORIZONTALTEXTALIGNMENT											32
#define DISPID_TB_HOTBUTTON																		33
#define DISPID_TB_HOVERTIME																		34
#define DISPID_TB_HWND																				35
#define DISPID_TB_HWNDCHEVRONTOOLBAR													36
#define DISPID_TB_HWNDTOOLTIP																	37
#define DISPID_TB_IDEALHEIGHT																	38
#define DISPID_TB_IDEALWIDTH																	39
#define DISPID_TB_IMAGELISTCOUNT															40
#define DISPID_TB_INSERTMARKCOLOR															41
#define DISPID_TB_MAXIMUMBUTTONWIDTH													42
#define DISPID_TB_MAXIMUMHEIGHT																43
#define DISPID_TB_MAXIMUMTEXTROWS															44
#define DISPID_TB_MAXIMUMWIDTH																45
#define DISPID_TB_MENUBARTHEME																46
#define DISPID_TB_MENUMODE																		47
#define DISPID_TB_MINIMUMBUTTONWIDTH													48
#define DISPID_TB_MOUSEICON																		49
#define DISPID_TB_MOUSEPOINTER																50
#define DISPID_TB_MULTICOLUMN																	51
#define DISPID_TB_NATIVEDROPTARGET														52
#define DISPID_TB_NORMALDROPDOWNBUTTONSTYLE										53
#define DISPID_TB_OLEDRAGIMAGESTYLE														54
#define DISPID_TB_ORIENTATION																	55
#define DISPID_TB_PROCESSCONTEXTMENUKEYS											56
#define DISPID_TB_RAISECUSTOMDRAWEVENTONERASEBACKGROUND				57
#define DISPID_TB_REGISTERFOROLEDRAGDROP											58
#define DISPID_TB_RIGHTTOLEFT																	59
#define DISPID_TB_SHADOWCOLOR																	60
#define DISPID_TB_SHOWDRAGIMAGE																61
#define DISPID_TB_SHOWTOOLTIPS																62
#define DISPID_TB_SUGGESTEDICONSIZE														63
#define DISPID_TB_SUPPORTOLEDRAGIMAGES												64
#define DISPID_TB_USEMNEMONICS																65
#define DISPID_TB_USESYSTEMFONT																66
#define DISPID_TB_VERSION																			67
#define DISPID_TB_VERTICALBUTTONPADDING												68
#define DISPID_TB_VERTICALBUTTONSPACING												69
#define DISPID_TB_VERTICALTEXTALIGNMENT												70
#define DISPID_TB_WRAPBUTTONS																	71
// fingerprint
#define DISPID_TB_APPID																				500
#define DISPID_TB_APPNAME																			501
#define DISPID_TB_APPSHORTNAME																502
#define DISPID_TB_BUILD																				503
#define DISPID_TB_CHARSET																			504
#define DISPID_TB_ISRELEASE																		505
#define DISPID_TB_PROGRAMMER																	506
#define DISPID_TB_TESTER																			507
// methods
#define DISPID_TB_ABOUT																				DISPID_ABOUTBOX
#define DISPID_TB_AUTOSIZE																		72
#define DISPID_TB_COUNTACCELERATOROCCURRENCES									73
#define DISPID_TB_CREATEBUTTONCONTAINER												74
#define DISPID_TB_CUSTOMIZE																		75
#define DISPID_TB_DISPLAYCHEVRONPOPUPWINDOW										76
#define DISPID_TB_GETCLOSESTINSERTMARKPOSITION								77
#define DISPID_TB_GETCOMMANDFORHOTKEY													78
#define DISPID_TB_GETINSERTMARKPOSITION												79
#define DISPID_TB_HITTEST																			80
#define DISPID_TB_HITTESTCHEVRONTOOLBAR												81
#define DISPID_TB_LOADDEFAULTIMAGES														82
#define DISPID_TB_LOADSETTINGSFROMFILE												83
#define DISPID_TB_LOADTOOLBARSTATEFROMREGISTRY								84
#define DISPID_TB_OLEDRAG																			85
#define DISPID_TB_REFRESH																			DISPID_REFRESH
#define DISPID_TB_REGISTERHOTKEY															86
#define DISPID_TB_SAVESETTINGSTOFILE													87
#define DISPID_TB_SAVETOOLBARSTATETOREGISTRY									88
#define DISPID_TB_SETBUTTONROWCOUNT														89
#define DISPID_TB_SETHOTBUTTON																90
#define DISPID_TB_SETINSERTMARKPOSITION												91
#define DISPID_TB_UNREGISTERHOTKEY														92
#define DISPID_TB_FINISHOLEDRAGDROP														93


// _IToolBarEvents
// methods
#define DISPID_TBE_BEFOREDISPLAYCHEVRONPOPUP									1
#define DISPID_TBE_BEGINCUSTOMIZATION													2
#define DISPID_TBE_BUTTONBEGINDRAG														3
#define DISPID_TBE_BUTTONBEGINRDRAG														4
#define DISPID_TBE_BUTTONGETDISPLAYINFO												5
#define DISPID_TBE_BUTTONGETDROPDOWNMENU											6
#define DISPID_TBE_BUTTONGETINFOTIPTEXT												7
#define DISPID_TBE_BUTTONMOUSEENTER														8
#define DISPID_TBE_BUTTONMOUSELEAVE														9
#define DISPID_TBE_BUTTONSELECTIONSTATECHANGED								10
#define DISPID_TBE_CLICK																			11
#define DISPID_TBE_CONTEXTMENU																12
#define DISPID_TBE_CUSTOMDRAW																	13
#define DISPID_TBE_CUSTOMIZEDCONTROL													14
#define DISPID_TBE_CUSTOMIZEDIALOGREMOVINGBUTTON							15
#define DISPID_TBE_DBLCLICK																		16
#define DISPID_TBE_DESTROYEDCONTROLWINDOW											17
#define DISPID_TBE_DESTROYINGCHEVRONPOPUP											18
#define DISPID_TBE_DISPLAYCUSTOMIZATIONHELP										19
#define DISPID_TBE_DROPDOWN																		20
#define DISPID_TBE_ENDCUSTOMIZATION														21
#define DISPID_TBE_EXECUTECOMMAND															22
#define DISPID_TBE_FREEBUTTONDATA															23
#define DISPID_TBE_GETAVAILABLEBUTTON													24
#define DISPID_TBE_HOTBUTTONCHANGEWRAPPING										25
#define DISPID_TBE_HOTBUTTONCHANGING													26
#define DISPID_TBE_INITIALIZECUSTOMIZATIONDIALOG							27
#define DISPID_TBE_INITIALIZETOOLBARSTATEREGISTRYRESTORAGE		28
#define DISPID_TBE_INITIALIZETOOLBARSTATEREGISTRYSTORAGE			29
#define DISPID_TBE_INSERTEDBUTTON															30
#define DISPID_TBE_INSERTINGBUTTON														31
#define DISPID_TBE_ISDUPLICATEACCELERATOR											32
#define DISPID_TBE_KEYDOWN																		33
#define DISPID_TBE_KEYPRESS																		34
#define DISPID_TBE_KEYUP																			35
#define DISPID_TBE_MAPACCELERATOR															36
#define DISPID_TBE_MCLICK																			37
#define DISPID_TBE_MDBLCLICK																	38
#define DISPID_TBE_MOUSEDOWN																	39
#define DISPID_TBE_MOUSEENTER																	40
#define DISPID_TBE_MOUSEHOVER																	41
#define DISPID_TBE_MOUSELEAVE																	42
#define DISPID_TBE_MOUSEMOVE																	43
#define DISPID_TBE_MOUSEUP																		44
#define DISPID_TBE_OLECOMPLETEDRAG														45
#define DISPID_TBE_OLEDRAGDROP																46
#define DISPID_TBE_OLEDRAGENTER																47
#define DISPID_TBE_OLEDRAGENTERPOTENTIALTARGET								48
#define DISPID_TBE_OLEDRAGLEAVE																49
#define DISPID_TBE_OLEDRAGLEAVEPOTENTIALTARGET								50
#define DISPID_TBE_OLEDRAGMOUSEMOVE														51
#define DISPID_TBE_OLEGIVEFEEDBACK														52
#define DISPID_TBE_OLEQUERYCONTINUEDRAG												53
#define DISPID_TBE_OLERECEIVEDNEWDATA													54
#define DISPID_TBE_OLESETDATA																	55
#define DISPID_TBE_OLESTARTDRAG																56
#define DISPID_TBE_OPENPOPUPMENU															57
#define DISPID_TBE_QUERYINSERTBUTTON													58
#define DISPID_TBE_QUERYREMOVEBUTTON													59
#define DISPID_TBE_RAWMENUMESSAGE															60
#define DISPID_TBE_RCLICK																			61
#define DISPID_TBE_RDBLCLICK																	62
#define DISPID_TBE_RECREATEDCONTROLWINDOW											63
#define DISPID_TBE_REMOVEDBUTTON															64
#define DISPID_TBE_REMOVINGBUTTON															65
#define DISPID_TBE_RESETCUSTOMIZATIONS												66
#define DISPID_TBE_RESIZEDCONTROLWINDOW												67
#define DISPID_TBE_RESTOREBUTTONFROMREGISTRYSTREAM						68
#define DISPID_TBE_SAVEBUTTONTOREGISTRYSTREAM									69
#define DISPID_TBE_SELECTEDMENUITEM														70
#define DISPID_TBE_XCLICK																			71
#define DISPID_TBE_XDBLCLICK																	72


// ToolBarButton
// properties
#define DISPID_TBB_AUTOSIZE																		1
#define DISPID_TBB_BUTTONDATA																	2
#define DISPID_TBB_BUTTONTYPE																	3
#define DISPID_TBB_DISPLAYTEXT																4
#define DISPID_TBB_DROPDOWNSTYLE															5
#define DISPID_TBB_DROPPEDDOWN																6
#define DISPID_TBB_ENABLED																		7
#define DISPID_TBB_FOLLOWEDBYLINEBREAK												8
#define DISPID_TBB_HOT																				9
#define DISPID_TBB_ICONINDEX																	10
#define DISPID_TBB_ID																					DISPID_VALUE
#define DISPID_TBB_IMAGELISTINDEX															11
#define DISPID_TBB_INDEX																			12
#define DISPID_TBB_MARKED																			13
#define DISPID_TBB_PARTOFCHEVRONTOOLBAR												14
#define DISPID_TBB_PARTOFGROUP																15
#define DISPID_TBB_PUSHED																			16
#define DISPID_TBB_SELECTIONSTATE															17
#define DISPID_TBB_SHOULDBEDISPLAYEDINCHEVRONPOPUP						18
#define DISPID_TBB_SHOWINGELLIPSIS														19
#define DISPID_TBB_TEXT																				20
#define DISPID_TBB_USEMNEMONIC																21
#define DISPID_TBB_VISIBLE																		22
#define DISPID_TBB_WIDTH																			23
// methods
#define DISPID_TBB_GETCONTAINEDWINDOW													24
#define DISPID_TBB_GETRECTANGLE																25
#define DISPID_TBB_SETCONTAINEDWINDOW													26


// IToolBarButtons
// properties
#define DISPID_TBBS_CASESENSITIVEFILTERS											1
#define DISPID_TBBS_COMPARISONFUNCTION												2
#define DISPID_TBBS_FILTER																		3
#define DISPID_TBBS_FILTERTYPE																4
#define DISPID_TBBS_ITEM																			DISPID_VALUE
#define DISPID_TBBS__NEWENUM																	DISPID_NEWENUM
// methods
#define DISPID_TBBS_ADD																				5
#define DISPID_TBBS_CONTAINS																	6
#define DISPID_TBBS_COUNT																			7
#define DISPID_TBBS_REMOVE																		8
#define DISPID_TBBS_REMOVEALL																	9


// IToolBarButtonContainer
// properties
#define DISPID_TBBC_ITEM																			DISPID_VALUE
#define DISPID_TBBC__NEWENUM																	DISPID_NEWENUM
// methods
#define DISPID_TBBC_ADD																				1
#define DISPID_TBBC_CLONE																			2
#define DISPID_TBBC_COUNT																			3
#define DISPID_TBBC_CREATEDRAGIMAGE														4
#define DISPID_TBBC_REMOVE																		5
#define DISPID_TBBC_REMOVEALL																	6


// IVirtualToolBarButton
// properties
#define DISPID_VTBB_AUTOSIZE																	1
#define DISPID_VTBB_BUTTONDATA																2
#define DISPID_VTBB_BUTTONTYPE																3
#define DISPID_VTBB_DISPLAYTEXT																4
#define DISPID_VTBB_DROPDOWNSTYLE															5
#define DISPID_VTBB_ENABLED																		6
#define DISPID_VTBB_FOLLOWEDBYLINEBREAK												7
#define DISPID_VTBB_ICONINDEX																	8
#define DISPID_VTBB_ID																				DISPID_VALUE
#define DISPID_VTBB_IMAGELISTINDEX														9
#define DISPID_VTBB_INDEX																			10
#define DISPID_VTBB_MARKED																		11
#define DISPID_VTBB_PARTOFGROUP																12
#define DISPID_VTBB_PUSHED																		13
#define DISPID_VTBB_SELECTIONSTATE														14
#define DISPID_VTBB_SHOWINGELLIPSIS														15
#define DISPID_VTBB_TEXT																			16
#define DISPID_VTBB_USEMNEMONIC																17
#define DISPID_VTBB_VISIBLE																		18
#define DISPID_VTBB_WIDTH																			19
// methods


// IOLEDataObject
// properties
// methods
#define DISPID_ODO_CLEAR																			1
#define DISPID_ODO_GETCANONICALFORMAT													2
#define DISPID_ODO_GETDATA																		3
#define DISPID_ODO_GETFORMAT																	4
#define DISPID_ODO_SETDATA																		5
#define DISPID_ODO_GETDROPDESCRIPTION													6
#define DISPID_ODO_SETDROPDESCRIPTION													7


// IControlHostWindow
// properties
#define DISPID_CHW_HWND																				DISPID_VALUE
#define DISPID_CHW_HWNDCHILD																	1
#define DISPID_CHW_RESIZEABLE																	2
#define DISPID_CHW_VISIBLE																		3
#define DISPID_CHW_WINDOWTITLE																4
// methods
#define DISPID_CHW_CALCULATEWINDOWSIZE												5
#define DISPID_CHW_CREATE																			6
#define DISPID_CHW_DESTROY																		7
#define DISPID_CHW_GETCLIENTRECTANGLE													8
#define DISPID_CHW_GETRECTANGLE																9
#define DISPID_CHW_MOVE																				10


// _IControlHostWindowEvents
// methods
#define DISPID_CHWE_ACTIVATE																	1
#define DISPID_CHWE_CLOSING																		2
#define DISPID_CHWE_CREATED																		3
#define DISPID_CHWE_DEACTIVATE																4
#define DISPID_CHWE_DESTROYED																	5
#define DISPID_CHWE_MOVED																			6
#define DISPID_CHWE_MOVING																		7
#define DISPID_CHWE_RESIZED																		8
#define DISPID_CHWE_RESIZING																	9
#define DISPID_CHWE_TITLEDBLCLICK															10
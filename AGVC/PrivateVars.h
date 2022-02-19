// ----------------------------------------------------------------------------------------
//                              COPYRIGHT NOTICE
// ----------------------------------------------------------------------------------------
//
// The Source Code Store LLC
// ACTIVEGANTT SCHEDULER COMPONENT FOR C++ - ActiveGanttVC
// ActiveX Control
// Copyright (c) 2002-2017 The Source Code Store LLC
//
// All Rights Reserved. No parts of this file may be reproduced, modified or transmitted 
// in any form or by any means without the written permission of the author.
//
// ----------------------------------------------------------------------------------------
#pragma once

typedef enum E_MOUSEKEYBOARDEVENTS
{
	MouseHover = 0,
    MouseDown = 1,
    MouseMove = 2,
    MouseUp = 3,
    MouseClick = 4,
    MouseDblClick = 5,
    MouseWheel = 6,
    KeyPress = 7,
    KeyDown = 8,
    KeyUp = 9,
}E_MOUSEKEYBOARDEVENTS;

typedef enum E_FOCUSTYPE
{
    FCT_NORMAL = 0,
	FCT_KEEPLEFTRIGHTBOUNDS = 1,
	FCT_ADD = 2,
	FCT_VERTICALSPLITTER = 3,
}E_FOCUSTYPE;

typedef enum E_AREA
{
	EA_NONE = 0,
    EA_LEFT = 1,
    EA_TOP = 2,
    EA_RIGHT = 3,
    EA_BOTTOM = 4,
    EA_CENTER = 5,
}E_AREA;


typedef enum E_PROPOPTYPE
{
	POT_INITPROPERTIES = 0,
	POT_READPROPERTIES = 1,
	POT_WRITEPROPERTIES = 2
}E_PROPOPTYPE;


typedef enum E_SCROLLSTATE
{
	SS_CANTDISPLAY = 0,
	SS_NOTNEEDED = 1,
	SS_NEEDED = 2,
	SS_SHOWN = 3,
	SS_HIDDEN = 4
}E_SCROLLSTATE;

typedef enum E_CURSORTYPE
{
    CT_NORMAL = 0,
    CT_SIZETASK = 1,
    CT_MOVETASK = 2,
    CT_MOVEMILESTONE = 3,
    CT_CLIENTAREA = 4,
    CT_MOVESPLITTER = 5,
    CT_ROWHEIGHT = 6,
    CT_COLUMNWIDTH = 7,
    CT_MOVEROW = 8,
    CT_MOVECOLUMN = 11,
    CT_SCROLLTIMELINE = 9,
    CT_NODROP = 10,
    CT_PERCENTAGE = 12,
	CT_PREDECESSOR = 13,
	CT_IBEAM = 14,
}E_CURSORTYPE;

typedef enum GRE_BUTTONSTATE
{
	BTNS_INACTIVE = 0x100,
	BTNS_FLAT = 0x4000,
	BTNS_NORMAL = 0,
}GRE_BUTTONSTATE;

typedef enum GRE_SCROLLBUTTONS
{
    SCRB_SCROLLDOWN = 0x1,
    SCRB_SCROLLLEFT = 0x2,
    SCRB_SCROLLRIGHT = 0x3,
    SCRB_SCROLLUP = 0x0,
}GRE_SCROLLBUTTONS;

typedef enum E_CLIENTAREAVISIBILITY
{
	VS_ABOVEVISIBLEAREA = 0,
	VS_INSIDEVISIBLEAREA = 1,
	VS_BELOWVISIBLEAREA = 2,
	VS_LEFTOFVISIBLEAREA = 3,
	VS_RIGHTOFVISIBLEAREA = 4,
}E_CLIENTAREAVISIBILITY;

typedef enum E_SCROLLBUTTON
{
	SBC_UP = 1,
	SBC_DOWN = 2,
	SBC_LEFT = 3,
	SBC_RIGHT = 4
}E_SCROLLBUTTON;

typedef enum E_SCROLLBUTTONSTATE
{
	BS_NORMAL = 1,
	BS_PUSHED = 2,
	BS_INACTIVE = 3
}E_SCROLLBUTTONSTATE;
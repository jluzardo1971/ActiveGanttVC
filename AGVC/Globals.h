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

typedef struct tagRECTD
{
   LONG Left;
   LONG Top;
   LONG Width;
   LONG Height;
} RECTD;

typedef enum GRE_IMAGEFORMAT
{
	IMF_BMP = 0,
	IMF_JPEG = 1,
	IMF_GIF = 2,
	IMF_TIFF = 3,
	IMF_PNG = 4
}GRE_IMAGEFORMAT;

typedef enum GRE_ARROWDIRECTION
{
	AWD_UP = 0,
	AWD_DOWN = 1,
	AWD_LEFT = 2,
	AWD_RIGHT = 3
}GRE_ARROWDIRECTION;

typedef enum GRE_LINEDRAWSTYLE
{
	LDS_SOLID = 0,
	LDS_DASH = 1,
	LDS_DOT = 2,
	LDS_DASHDOT = 3,
	LDS_DASHDOTDOT = 4
}GRE_LINEDRAWSTYLE;

typedef enum GRE_EDGETYPE
{
	ET_SUNKEN = 1,
	ET_RAISED = 2
}GRE_EDGETYPE;

typedef enum GRE_BUTTONSTYLE
{
	BT_NORMALWINDOWS = 0,
	BT_LIGHTWEIGHT = 1
}GRE_BUTTONSTYLE;

typedef enum GRE_COLORS
{
	CLR_ACTIVEBORDER = -2147483638,
	CLR_ACTIVECAPTION = -2147483646,
	CLR_ACTIVECAPTIONTEXT = -2147483639,
	CLR_APPWORKSPACE = -2147483636,
	CLR_CONTROL = -2147483633,
	CLR_CONTROLDARK = -2147483632,
	CLR_CONTROLDARKDARK = -2147483627,
	CLR_CONTROLLIGHT = -2147483626,
	CLR_CONTROLLIGHTLIGHT = -2147483628,
	CLR_CONTROLTEXT = -2147483630,
	CLR_DESKTOP = -2147483647,
	CLR_GRAYTEXT = -2147483631,
	CLR_HIGHLIGHT = -2147483635,
	CLR_HIGHLIGHTTEXT = -2147483634,
	CLR_HOTTRACK = -2147483635,
	CLR_INACTIVEBORDER = -2147483637,
	CLR_INACTIVECAPTION = -2147483645,
	CLR_INACTIVECAPTIONTEXT = -2147483629,
	CLR_INFO = -2147483624,
	CLR_INFOTEXT = -2147483625,
	CLR_MENU = -2147483644,
	CLR_MENUTEXT = -2147483641,
	CLR_SCROLLBAR = -214748364,
	CLR_WINDOW = -2147483643,
	CLR_WINDOWFRAME = -2147483642,
	CLR_WINDOWTEXT = -2147483640,
	CLR_TRANSPARENT = 16777215,
	CLR_ALICEBLUE = 16775408,
	CLR_ANTIQUEWHITE = 14150650,
	CLR_AQUA = 16776960,
	CLR_AQUAMARINE = 13959039,
	CLR_AZURE = 16777200,
	CLR_BEIGE = 14480885,
	CLR_BISQUE = 12903679,
	CLR_BLACK = 0,
	CLR_BLANCHEDALMOND = 13495295,
	CLR_BLUE = 16711680,
	CLR_BLUEVIOLET = 14822282,
	CLR_BROWN = 2763429,
	CLR_BURLYWOOD = 8894686,
	CLR_CADETBLUE = 10526303,
	CLR_CHARTREUSE = 65407,
	CLR_CHOCOLATE = 1993170,
	CLR_CORAL = 5275647,
	CLR_CORNFLOWERBLUE = 15570276,
	CLR_CORNSILK = 14481663,
	CLR_CRIMSON = 3937500,
	CLR_CYAN = 16776960,
	CLR_DARKBLUE = 9109504,
	CLR_DARKCYAN = 9145088,
	CLR_DARKGOLDENROD = 755384,
	CLR_DARKGRAY = 11119017,
	CLR_DARKGREEN = 25600,
	CLR_DARKKHAKI = 7059389,
	CLR_DARKMAGENTA = 9109643,
	CLR_DARKOLIVEGREEN = 3107669,
	CLR_DARKORANGE = 36095,
	CLR_DARKORCHID = 13382297,
	CLR_DARKRED = 139,
	CLR_DARKSALMON = 8034025,
	CLR_DARKSEAGREEN = 9157775,
	CLR_DARKSLATEBLUE = 9125192,
	CLR_DARKSLATEGRAY = 5197615,
	CLR_DARKTURQUOISE = 13749760,
	CLR_DARKVIOLET = 13828244,
	CLR_DEEPPINK = 9639167,
	CLR_DEEPSKYBLUE = 16760576,
	CLR_DIMGRAY = 6908265,
	CLR_DODGERBLUE = 16748574,
	CLR_FIREBRICK = 2237106,
	CLR_FLORALWHITE = 15792895,
	CLR_FORESTGREEN = 2263842,
	CLR_FUCHSIA = 16711935,
	CLR_GAINSBORO = 14474460,
	CLR_GHOSTWHITE = 16775416,
	CLR_GOLD = 55295,
	CLR_GOLDENROD = 2139610,
	CLR_GRAY = 8421504,
	CLR_GREEN = 32768,
	CLR_GREENYELLOW = 3145645,
	CLR_HONEYDEW = 15794160,
	CLR_HOTPINK = 11823615,
	CLR_INDIANRED = 6053069,
	CLR_INDIGO = 8519755,
	CLR_IVORY = 15794175,
	CLR_KHAKI = 9234160,
	CLR_LAVENDER = 16443110,
	CLR_LAVENDERBLUSH = 16118015,
	CLR_LAWNGREEN = 64636,
	CLR_LEMONCHIFFON = 13499135,
	CLR_LIGHTBLUE = 15128749,
	CLR_LIGHTCORAL = 8421616,
	CLR_LIGHTCYAN = 16777184,
	CLR_LIGHTGOLDENRODYELLOW = 13826810,
	CLR_LIGHTGRAY = 13882323,
	CLR_LIGHTGREEN = 9498256,
	CLR_LIGHTPINK = 12695295,
	CLR_LIGHTSALMON = 8036607,
	CLR_LIGHTSEAGREEN = 11186720,
	CLR_LIGHTSKYBLUE = 16436871,
	CLR_LIGHTSLATEGRAY = 10061943,
	CLR_LIGHTSTEELBLUE = 14599344,
	CLR_LIGHTYELLOW = 14745599,
	CLR_LIME = 65280,
	CLR_LIMEGREEN = 3329330,
	CLR_LINEN = 15134970,
	CLR_MAGENTA = 16711935,
	CLR_MAROON = 128,
	CLR_MEDIUMAQUAMARINE = 11193702,
	CLR_MEDIUMBLUE = 13434880,
	CLR_MEDIUMORCHID = 13850042,
	CLR_MEDIUMPURPLE = 14381203,
	CLR_MEDIUMSEAGREEN = 7451452,
	CLR_MEDIUMSLATEBLUE = 15624315,
	CLR_MEDIUMSPRINGGREEN = 10156544,
	CLR_MEDIUMTURQUOISE = 13422920,
	CLR_MEDIUMVIOLETRED = 8721863,
	CLR_MIDNIGHTBLUE = 7346457,
	CLR_MINTCREAM = 16449525,
	CLR_MISTYROSE = 14804223,
	CLR_MOCCASIN = 11920639,
	CLR_NAVAJOWHITE = 11394815,
	CLR_NAVY = 8388608,
	CLR_OLDLACE = 15136253,
	CLR_OLIVE = 32896,
	CLR_OLIVEDRAB = 2330219,
	CLR_ORANGE = 42495,
	CLR_ORANGERED = 17919,
	CLR_ORCHID = 14053594,
	CLR_PALEGOLDENROD = 11200750,
	CLR_PALEGREEN = 10025880,
	CLR_PALETURQUOISE = 15658671,
	CLR_PALEVIOLETRED = 9662683,
	CLR_PAPAYAWHIP = 14020607,
	CLR_PEACHPUFF = 12180223,
	CLR_PERU = 4163021,
	CLR_PINK = 13353215,
	CLR_PLUM = 14524637,
	CLR_POWDERBLUE = 15130800,
	CLR_PURPLE = 8388736,
	CLR_RED = 255,
	CLR_ROSYBROWN = 9408444,
	CLR_ROYALBLUE = 14772545,
	CLR_SADDLEBROWN = 1262987,
	CLR_SALMON = 7504122,
	CLR_SANDYBROWN = 6333684,
	CLR_SEAGREEN = 5737262,
	CLR_SEASHELL = 15660543,
	CLR_SIENNA = 2970272,
	CLR_SILVER = 12632256,
	CLR_SKYBLUE = 15453831,
	CLR_SLATEBLUE = 13458026,
	CLR_SLATEGRAY = 9470064,
	CLR_SNOW = 16448255,
	CLR_SPRINGGREEN = 8388352,
	CLR_STEELBLUE = 11829830,
	CLR_TAN = 9221330,
	CLR_TEAL = 8421376,
	CLR_THISTLE = 14204888,
	CLR_TOMATO = 4678655,
	CLR_TURQUOISE = 13688896,
	CLR_VIOLET = 15631086,
	CLR_WHEAT = 11788021,
	CLR_WHITE = 16777215,
	CLR_WHITESMOKE = 16119285,
	CLR_YELLOW = 65535,
	CLR_YELLOWGREEN = 3329434,
	CLR_DARKGREY = 8421504,
	CLR_VERYDARKGREY = 4210752,
	CLR_VERYLIGHTGREY = 15461355,
	CLR_ALMOSTBLACK = 4342082,
	CLR_BUTTONFACE = 13160660
}GRE_COLORS;

typedef enum GRE_LINETYPE
{
	LT_NORMAL = 0,
	LT_BORDER = 1,
	LT_FILLED = 2
}GRE_LINETYPE;

typedef enum GRE_BACKGROUNDMODE
{
	FP_SOLID = 0,
	FP_TRANSPARENT = 1,
	FP_GRADIENT = 2,
	FP_PATTERN = 3,
	FP_HATCH = 4
}GRE_BACKGROUNDMODE;

typedef enum GRE_PATTERN
{
	FP_HORIZONTALLINE = 2,
	FP_VERTICALLINE = 3,
	FP_UPWARDDIAGONAL = 4,
	FP_DOWNWARDDIAGONAL = 5,
	FP_CROSS = 6,
	FP_DIAGONALCROSS = 7,
	FP_LIGHT = 8,
	FP_MEDIUM = 9,
	FP_DARK = 10
}GRE_PATTERN;

typedef enum GRE_FIGURETYPE
{
	FT_NONE = 0,
    FT_PROJECTUP = 1,
    FT_PROJECTDOWN = 2,
    FT_DIAMOND = 3,
    FT_CIRCLEDIAMOND = 4,
    FT_TRIANGLEUP = 5,
    FT_TRIANGLEDOWN = 6,
    FT_TRIANGLERIGHT = 7,
    FT_TRIANGLELEFT = 8,
    FT_CIRCLETRIANGLEUP = 9,
    FT_CIRCLETRIANGLEDOWN = 10,
    FT_ARROWUP = 11,
    FT_ARROWDOWN = 12,
    FT_CIRCLEARROWUP = 13,
    FT_CIRCLEARROWDOWN = 14,
    FT_SMALLPROJECTUP = 15,
    FT_SMALLPROJECTDOWN = 16,
    FT_RECTANGLE = 17,
    FT_SQUARE = 18,
    FT_CIRCLE = 19
}GRE_FIGURETYPE;

typedef enum GRE_DRAWINGOBJECT
{
	DO_GENERAL = 0,
	DO_MILESTONE = 1
}GRE_DRAWINGOBJECT;

typedef enum GRE_PAPERTYPE
{
	PK_CUSTOM = 0,
	PK_LETTER = 1,
	PK_LETTERSMALL = 2,
	PK_TABLOID = 3,
	PK_LEDGER = 4,
	PK_LEGAL = 5,
	PK_STATEMENT = 6,
	PK_EXECUTIVE = 7,
	PK_A3 = 8,
	PK_A4 = 9,
	PK_A4SMALL = 10,
	PK_A5 = 11,
	PK_B4 = 12,
	PK_B5 = 13,
	PK_FOLIO = 14,
	PK_QUARTO = 15,
	PK_STANDARD10X14 = 16,
	PK_STANDARD11X17 = 17,
	PK_NOTE = 18,
	PK_NUMBER9ENVELOPE = 19,
	PK_NUMBER10ENVELOPE = 20,
	PK_NUMBER11ENVELOPE = 21,
	PK_NUMBER12ENVELOPE = 22,
	PK_NUMBER14ENVELOPE = 23,
	PK_CSHEET = 24,
	PK_DSHEET = 25,
	PK_ESHEET = 26,
	PK_DLENVELOPE = 27,
	PK_C5ENVELOPE = 28,
	PK_C3ENVELOPE = 29,
	PK_C4ENVELOPE = 30,
	PK_C6ENVELOPE = 31,
	PK_C65ENVELOPE = 32,
	PK_B4ENVELOPE = 33,
	PK_B5ENVELOPE = 34,
	PK_B6ENVELOPE = 35,
	PK_ITALYENVELOPE = 36,
	PK_MONARCHENVELOPE = 37,
	PK_PERSONALENVELOPE = 38,
	PK_USSTANDARDFANFOLD = 39,
	PK_GERMANSTANDARDFANFOLD = 40,
	PK_GERMANLEGALFANFOLD = 41,
	PK_ISOB4 = 42,
	PK_JAPANESEPOSTCARD = 43,
	PK_STANDARD9X11 = 44,
	PK_STANDARD10X11 = 45,
	PK_STANDARD15X11 = 46,
	PK_INVITEENVELOPE = 47,
	PK_48 = 48,
	PK_49 = 49,
	PK_LETTEREXTRA = 50,
	PK_LEGALEXTRA = 51,
	PK_TABLOIDEXTRA = 52,
	PK_A4EXTRA = 53,
	PK_LETTERTRANSVERSE = 54,
	PK_A4TRANSVERSE = 55,
	PK_LETTEREXTRATRANSVERSE = 56,
	PK_APLUS = 57,
	PK_BPLUS = 58,
	PK_LETTERPLUS = 59,
	PK_A4PLUS = 60,
	PK_A5TRANSVERSE = 61,
	PK_B5TRANSVERSE = 62,
	PK_A3EXTRA = 63,
	PK_A5EXTRA = 64,
	PK_B5EXTRA = 65,
	PK_A2 = 66,
	PK_A3TRANSVERSE = 67,
	PK_A3EXTRATRANSVERSE = 68,
	PK_JAPANESEDOUBLEPOSTCARD = 69,
	PK_A6 = 70,
	PK_JAPANESEENVELOPEKAKUNUMBER2 = 71,
	PK_JAPANESEENVELOPEKAKUNUMBER3 = 72,
	PK_JAPANESEENVELOPECHOUNUMBER3 = 73,
	PK_JAPANESEENVELOPECHOUNUMBER4 = 74,
	PK_LETTERROTATED = 75,
	PK_A3ROTATED = 76,
	PK_A4ROTATED = 77,
	PK_A5ROTATED = 78,
	PK_B4JISROTATED = 79,
	PK_B5JISROTATED = 80,
	PK_JAPANESEPOSTCARDROTATED = 81,
	PK_JAPANESEDOUBLEPOSTCARDROTATED = 82,
	PK_A6ROTATED = 83,
	PK_JAPANESEENVELOPEKAKUNUMBER2ROTATED = 84,
	PK_JAPANESEENVELOPEKAKUNUMBER3ROTATED = 85,
	PK_JAPANESEENVELOPECHOUNUMBER3ROTATED = 86,
	PK_JAPANESEENVELOPECHOUNUMBER4ROTATED = 87,
	PK_B6JIS = 88,
	PK_B6JISROTATED = 89,
	PK_STANDARD12X11 = 90,
	PK_JAPANESEENVELOPEYOUNUMBER4 = 91,
	PK_JAPANESEENVELOPEYOUNUMBER4ROTATED = 92,
	PK_PRC16K = 93,
	PK_PRC32K = 94,
	PK_PRC32KBIG = 95,
	PK_PRCENVELOPENUMBER1 = 96,
	PK_PRCENVELOPENUMBER2 = 97,
	PK_PRCENVELOPENUMBER3 = 98,
	PK_PRCENVELOPENUMBER4 = 99,
	PK_PRCENVELOPENUMBER5 = 100,
	PK_PRCENVELOPENUMBER6 = 101,
	PK_PRCENVELOPENUMBER7 = 102,
	PK_PRCENVELOPENUMBER8 = 103,
	PK_PRCENVELOPENUMBER9 = 104,
	PK_PRCENVELOPENUMBER10 = 105,
	PK_PRC16KROTATED = 106,
	PK_PRC32KROTATED = 107,
	PK_PRC32KBIGROTATED = 108,
	PK_PRCENVELOPENUMBER1ROTATED = 109,
	PK_PRCENVELOPENUMBER2ROTATED = 110,
	PK_PRCENVELOPENUMBER3ROTATED = 111,
	PK_PRCENVELOPENUMBER4ROTATED = 112,
	PK_PRCENVELOPENUMBER5ROTATED = 113,
	PK_PRCENVELOPENUMBER6ROTATED = 114,
	PK_PRCENVELOPENUMBER7ROTATED = 115,
	PK_PRCENVELOPENUMBER8ROTATED = 116,
	PK_PRCENVELOPENUMBER9ROTATED = 117,
	PK_PRCENVELOPENUMBER10ROTATED = 118
}GRE_PAPERTYPE;

typedef enum GRE_ORIENTATION
{
	OT_PORTRAIT = 0,
	OT_LANDSCAPE = 1
}GRE_ORIENTATION;

typedef enum GRE_PRINTERRESOLUTION
{
	PR_HIGH = 0,
	PR_MEDIUM = 1,
	PR_LOW = 2,
	PR_DRAFT = 3,
	PR_CUSTOM = 4
}GRE_PRINTERRESOLUTION;

//GRE_DRAWTEXTCONSTANTS

typedef enum GRE_VERTICALALIGNMENT 
{
	VAL_TOP = 1,
	VAL_CENTER = 2,
	VAL_BOTTOM = 3
}GRE_VERTICALALIGNMENT;

typedef enum GRE_BORDERSTYLE
{
	SBR_NONE = 0,
	SBR_SINGLE = 1,
	SBR_CUSTOM = 2
}GRE_BORDERSTYLE;

typedef enum GRE_HORIZONTALALIGNMENT
{
	HAL_LEFT = 1,
	HAL_CENTER = 2,
	HAL_RIGHT = 3
}GRE_HORIZONTALALIGNMENT;

typedef enum GRE_GRADIENTFILLMODE
{
	GDT_HORIZONTAL = 0,
	GDT_VERTICAL = 1
}GRE_GRADIENTFILLMODE;

typedef enum GRE_FILLMODE
{
	FM_COMPLETELYFILLED = 0,
	FM_UPPERHALFFILLED = 1,
	FM_LOWERHALFFILLED = 2
}GRE_FILLMODE;

// Public Enumerations //////////////////////////////////////////////////////////////

typedef enum E_TIERPOSITION
{
	SP_UPPER = 0,
	SP_LOWER = 1,
	SP_MIDDLE = 2
}E_TIERPOSITION;

typedef enum E_SORTTYPE 
{
	ES_STRING = 0,
	ES_NUMERIC = 1,
	ES_DATE = 2
}E_SORTTYPE;

typedef enum E_ADDMODE 
{
	AT_TASKADD = 0,
	AT_MILESTONEADD = 1,
	AT_BOTH = 2,
	AT_DURATION_TASKADD = 3,
    AT_DURATION_MILESTONEADD = 4,
    AT_DURATION_BOTH = 5
}E_ADDMODE;

typedef enum E_OBJECTTYPE
{
	OT_TASK = 0,
	OT_MILESTONE = 1
}E_OBJECTTYPE;

typedef enum E_PLACEMENT
{
	PLC_ROWEXTENTSPLACEMENT = 0,
	PLC_OFFSETPLACEMENT = 1
}E_PLACEMENT;

typedef enum E_STYLEAPPEARANCE
{
	SA_RAISED = 0,
	SA_SUNKEN = 1,
	SA_FLAT = 2,
	SA_GRAPHICAL = 3
}E_STYLEAPPEARANCE;

typedef enum E_SCROLLBEHAVIOUR
{
	SB_DISABLE = 0,
	SB_HIDE = 1
}E_SCROLLBEHAVIOUR;

typedef enum E_REPORTERRORS 
{
	RE_MSGBOX = 0,
	RE_RAISE = 1,
	RE_RAISEEVENT = 2,
	RE_HIDE = 3
}E_REPORTERRORS;

typedef enum E_BORDERSTYLE
{
	TLB_NONE = 0,
	TLB_SINGLE = 1,
	TLB_3D = 2
}E_BORDERSTYLE;

typedef enum E_TIERTYPE
{
	ST_NOT_VISIBLE = 0,
	ST_DAYOFWEEK = 1,
	ST_MONTH = 2,
	ST_QUARTER = 3,
	ST_YEAR = 4,
	ST_WEEK = 5,
	ST_CUSTOM = 6,
	ST_DAY = 7,
	ST_DAYOFYEAR = 8,
	ST_HOUR = 9,
	ST_MINUTE = 10,
	ST_SECOND = 11,
	ST_MILLISECOND = 12,
	ST_MICROSECOND = 13
}E_TIERTYPE;

typedef enum E_TIMEBLOCKBEHAVIOUR
{
	TBB_ROWEXTENTS = 0,
	TBB_CONTROLEXTENTS = 1
}E_TIMEBLOCKBEHAVIOUR;

typedef enum E_MOVEMENTTYPE
{
	MT_UNRESTRICTED = 0,
	MT_RESTRICTEDTOROW = 1,
	MT_MOVEMENTDISABLED = 2
}E_MOVEMENTTYPE;

typedef enum E_TICKMARKTYPES
{
	TLT_BIG = 0,
	TLT_MEDIUM = 1,
	TLT_SMALL = 2
}E_TICKMARKTYPES;


typedef enum E_SCROLLBAR
{
	SCR_VERTICAL = 0,
	SCR_HORIZONTAL1 = 1,
	SCR_HORIZONTAL2 = 2
}E_SCROLLBAR;

typedef enum E_EVENTTARGET
{
    EVT_NONE = 0,
    EVT_TASK = 1,
    EVT_PERCENTAGE = 2,
    EVT_MILESTONE = 3,
    EVT_ROW = 4,
    EVT_CELL = 5,
    EVT_COLUMN = 6,
    EVT_CLIENTAREA = 7,
    EVT_EMPTYAREA = 8,
    EVT_GRID = 9,
    EVT_TIMELINE = 10,
    EVT_TIMEBLOCK = 11,
    EVT_HSCROLLBAR = 12,
    EVT_TIMELINESCROLLBAR = 13,
    EVT_VSCROLLBAR = 14,
    EVT_TREEVIEW = 15,
    EVT_SPLITTER = 16,
    EVT_NODE = 17,
    EVT_TREEVIEWCHECKBOX = 19,
    EVT_TREEVIEWSIGN = 20,
    EVT_SELECTEDCOLUMN = 50,
    EVT_SELECTEDROW = 51,
    EVT_SELECTEDTASK = 52,
    EVT_SELECTEDPERCENTAGE = 54,
	EVT_PREDECESSOR = 56,
	EVT_SELECTEDPREDECESSOR = 57
}E_EVENTTARGET;



typedef enum E_CONTROLMODE
{
    CM_GRID = 0,
    CM_TREEVIEW = 1,
}E_CONTROLMODE;

typedef enum GRE_HATCHSTYLE
{
	HSE_HORIZONTAL = 0,
	HSE_VERTICAL = 1,
	HS_FORWARDDIAGONAL = 2,
	HS_BACKWARDDIAGONAL = 3,
	HS_LARGEGRID = 4,
	HS_DIAGONALCROSS = 5,
	HS_PERCENT05 = 6,
	HS_PERCENT10 = 7,
	HS_PERCENT20 = 8,
	HS_PERCENT25 = 9,
	HS_PERCENT30 = 10,
	HS_PERCENT40 = 11,
	HS_PERCENT50 = 12,
	HS_PERCENT60 = 13,
	HS_PERCENT70 = 14,
	HS_PERCENT75 = 15,
	HS_PERCENT80 = 16,
	HS_PERCENT90 = 17,
	HS_LIGHTDOWNWARDDIAGONAL = 18,
	HS_LIGHTUPWARDDIAGONAL = 19,
	HS_DARKDOWNWARDDIAGONAL = 20,
	HS_DARKUPWARDDIAGONAL = 21,
	HS_WIDEDOWNWARDDIAGONAL = 22,
	HS_WIDEUPWARDDIAGONAL = 23,
	HS_LIGHTVERTICAL = 24,
	HS_LIGHTHORIZONTAL = 25,
	HS_NARROWVERTICAL = 26,
	HS_NARROWHORIZONTAL = 27,
	HS_DARKVERTICAL = 28,
	HS_DARKHORIZONTAL = 29,
	HS_DASHEDDOWNWARDDIAGONAL = 30,
	HS_DASHEDUPWARDDIAGONAL = 31,
	HS_DASHEDHORIZONTAL = 32,
	HS_DASHEDVERTICAL = 33,
	HS_SMALLCONFETTI = 34,
	HS_LARGECONFETTI = 35,
	HS_ZIGZAG = 36,
	HS_WAVE = 37,
	HS_DIAGONALBRICK = 38,
	HS_HORIZONTALBRICK = 39,
	HS_WEAVE = 40,
	HS_PLAID = 41,
	HS_DIVOT = 42,
	HS_DOTTEDGRID = 43,
	HS_DOTTEDDIAMOND = 44,
	HS_SHINGLE = 45,
	HS_TRELLIS = 46,
	HS_SPHERE = 47,
	HS_SMALLGRID = 48,
	HS_SMALLCHECKERBOARD = 49,
	HS_LARGECHECKERBOARD = 50,
	HS_OUTLINEDDIAMOND = 51,
	HS_SOLIDDIAMOND = 52
}GRE_HATCHSTYLE;

typedef enum E_TEXTOBJECTTYPE
{
    TOT_TASK = 1,
    TOT_ROW = 2,
    TOT_CELL = 3,
    TOT_COLUMN = 4,
    TOT_NODE = 5,
}E_TEXTOBJECTTYPE;

typedef enum E_PREDECESSORMODE
{
    PM_FORCE = 0,
    PM_CREATEWARNINGFLAG = 1,
    PM_RAISEEVENT = 2
}E_PREDECESSORMODE;

typedef enum E_TASKTYPE
{
    TT_START_END = 0,
    TT_DURATION = 1,
    TT_UNITS_DURATION_WORK = 2
}E_TASKTYPE;

typedef enum E_TBINTERVALTYPE
{
    TBIT_AUTOMATIC = 0,
    TBIT_MANUAL = 1
}E_TBINTERVALTYPE;

typedef enum E_TREEVIEWMODE
{
    TM_NONE = 0,
    TM_PLUSMINUSSIGNS = 1,
    TM_EXPANDCOLLAPSEICONS = 2,
}E_TREEVIEWMODE;

typedef enum E_CHANGETYPE
{
    CT_MOVE = 0,
    CT_SIZE = 1
}E_CHANGETYPE;

typedef enum SYS_ERRORS
{
	C_ROIV_UNK_TYPE = 50001,
	C_ROBV_UNK_TYPE = 50002,
	C_ROSV_UNK_TYPE = 50003,
	C_ROSV_DISP_TYPE = 50004,
	C_ROLV_UNK_TYPE = 50005,
	C_INVALID_PROP_NAME_1 = 50006,
	C_INVALID_PROP_NAME_2 = 50007,
	ERR_RETARRELEMKEY_G = 51132,
	ERR_COLLREMWHERE_1_G = 51133,
	ERR_COLLREMWHERE_2_G = 51134,
	ERR_COLLREMWHERE_3_G = 51135,
	ERR_COLLREMWHERE_4_G = 51136,
	ERR_COLLREMWHERENOT_1_G = 51137,
	ERR_COLLREMWHERENOT_2_G = 51138,
	ERR_COLLREMWHERENOT_3_G = 51139,
	ERR_COLLREMWHERENOT_4_G = 51140,
	ERR_ADDMODE_G = 51141,
	TASKS_ITEM_1 = 51142,
	TASKS_ITEM_2 = 51143,
	TASKS_ITEM_3 = 51144,
	TASKS_ITEM_4 = 51145,
	TASKS_ADD_1 = 51146,
	TASKS_ADD_2 = 51147,
	TASKS_ADD_3 = 51148,
	TASKS_REMOVE_1 = 51149,
	TASKS_REMOVE_2 = 51150,
	TASKS_REMOVE_3 = 51151,
	TASKS_REMOVE_4 = 51152,
	TASKS_SET_KEY = 51153,
	ROWS_ITEM_1 = 51155,
	ROWS_ITEM_2 = 51156,
	ROWS_ITEM_3 = 51157,
	ROWS_ITEM_4 = 51158,
	ROWS_ADD_1 = 51159,
	ROWS_ADD_2 = 51160,
	ROWS_ADD_3 = 51161,
	ROWS_REMOVE_1 = 51162,
	ROWS_REMOVE_2 = 51163,
	ROWS_REMOVE_3 = 51164,
	ROWS_REMOVE_4 = 51165,
	ROWS_SET_KEY = 51166,
	COLUMNS_ITEM_1 = 51168,
	COLUMNS_ITEM_2 = 51169,
	COLUMNS_ITEM_3 = 51170,
	COLUMNS_ITEM_4 = 51171,
	COLUMNS_ADD_1 = 51172,
	COLUMNS_ADD_2 = 51173,
	COLUMNS_ADD_3 = 51174,
	COLUMNS_REMOVE_1 = 51175,
	COLUMNS_REMOVE_2 = 51176,
	COLUMNS_REMOVE_3 = 51177,
	COLUMNS_REMOVE_4 = 51178,
	COLUMNS_SET_KEY = 51179,
	CELLS_ITEM_1 = 51181,
	CELLS_ITEM_2 = 51182,
	CELLS_ITEM_3 = 51183,
	CELLS_ITEM_4 = 51184,
	CELLS_ADD_1 = 51185,
	CELLS_ADD_2 = 51186,
	CELLS_ADD_3 = 51187,
	CELLS_REMOVE_1 = 51188,
	CELLS_REMOVE_2 = 51189,
	CELLS_REMOVE_3 = 51190,
	CELLS_REMOVE_4 = 51191,
	CELLS_SET_KEY = 51192,
	PREDECESSORS_ITEM_1 = 51194,
	PREDECESSORS_ITEM_2 = 51195,
	PREDECESSORS_ITEM_3 = 51196,
	PREDECESSORS_ITEM_4 = 51197,
	PREDECESSORS_ADD_1 = 51198,
	PREDECESSORS_ADD_2 = 51199,
	PREDECESSORS_ADD_3 = 51200,
	PREDECESSORS_REMOVE_1 = 51201,
	PREDECESSORS_REMOVE_2 = 51202,
	PREDECESSORS_REMOVE_3 = 51203,
	PREDECESSORS_REMOVE_4 = 51204,
	PREDECESSORS_SET_KEY = 51205,
	TIMEBLOCKS_ITEM_1 = 51207,
	TIMEBLOCKS_ITEM_2 = 51208,
	TIMEBLOCKS_ITEM_3 = 51209,
	TIMEBLOCKS_ITEM_4 = 51210,
	TIMEBLOCKS_ADD_1 = 51211,
	TIMEBLOCKS_ADD_2 = 51212,
	TIMEBLOCKS_ADD_3 = 51213,
	TIMEBLOCKS_REMOVE_1 = 51214,
	TIMEBLOCKS_REMOVE_2 = 51215,
	TIMEBLOCKS_REMOVE_3 = 51216,
	TIMEBLOCKS_REMOVE_4 = 51217,
	TIMEBLOCKS_SET_KEY = 51218,
	LAYERS_ITEM_1 = 51220,
	LAYERS_ITEM_2 = 51221,
	LAYERS_ITEM_3 = 51222,
	LAYERS_ITEM_4 = 51223,
	LAYERS_ADD_1 = 51224,
	LAYERS_ADD_2 = 51225,
	LAYERS_ADD_3 = 51226,
	LAYERS_REMOVE_1 = 51227,
	LAYERS_REMOVE_2 = 51228,
	LAYERS_REMOVE_3 = 51229,
	LAYERS_REMOVE_4 = 51230,
	LAYERS_SET_KEY = 51231,
	STYLES_ITEM_1 = 51233,
	STYLES_ITEM_2 = 51234,
	STYLES_ITEM_3 = 51235,
	STYLES_ITEM_4 = 51236,
	STYLES_ADD_1 = 51237,
	STYLES_ADD_2 = 51238,
	STYLES_ADD_3 = 51239,
	STYLES_REMOVE_1 = 51240,
	STYLES_REMOVE_2 = 51241,
	STYLES_REMOVE_3 = 51242,
	STYLES_REMOVE_4 = 51243,
	STYLES_SET_KEY = 51244,
	PERCENTAGES_ITEM_1 = 51246,
	PERCENTAGES_ITEM_2 = 51247,
	PERCENTAGES_ITEM_3 = 51248,
	PERCENTAGES_ITEM_4 = 51249,
	PERCENTAGES_ADD_1 = 51250,
	PERCENTAGES_ADD_2 = 51251,
	PERCENTAGES_ADD_3 = 51252,
	PERCENTAGES_REMOVE_1 = 51253,
	PERCENTAGES_REMOVE_2 = 51254,
	PERCENTAGES_REMOVE_3 = 51255,
	PERCENTAGES_REMOVE_4 = 51256,
	PERCENTAGES_SET_KEY = 51257,
	VIEWS_ITEM_1 = 51272,
	VIEWS_ITEM_2 = 51273,
	VIEWS_ITEM_3 = 51274,
	VIEWS_ITEM_4 = 51275,
	VIEWS_ADD_1 = 51276,
	VIEWS_ADD_2 = 51277,
	VIEWS_ADD_3 = 51278,
	VIEWS_REMOVE_1 = 51279,
	VIEWS_REMOVE_2 = 51280,
	VIEWS_REMOVE_3 = 51281,
	VIEWS_REMOVE_4 = 51282,
	VIEWS_SET_KEY = 51283,
	TIERCOLORS_ITEM_1 = 51285,
	TIERCOLORS_ITEM_2 = 51286,
	TIERCOLORS_ITEM_3 = 51287,
	TIERCOLORS_ITEM_4 = 51288,
	TIERCOLORS_ADD_1 = 51289,
	TIERCOLORS_ADD_2 = 51290,
	TIERCOLORS_ADD_3 = 51291,
	TIERCOLORS_REMOVE_1 = 51292,
	TIERCOLORS_REMOVE_2 = 51293,
	TIERCOLORS_REMOVE_3 = 51294,
	TIERCOLORS_REMOVE_4 = 51295,
	TIERCOLORS_SET_KEY = 51296,
	TICKMARKS_ITEM_1 = 51298,
	TICKMARKS_ITEM_2 = 51299,
	TICKMARKS_ITEM_3 = 51300,
	TICKMARKS_ITEM_4 = 51301,
	TICKMARKS_ADD_1 = 51302,
	TICKMARKS_ADD_2 = 51303,
	TICKMARKS_ADD_3 = 51304,
	TICKMARKS_REMOVE_1 = 51305,
	TICKMARKS_REMOVE_2 = 51306,
	TICKMARKS_REMOVE_3 = 51307,
	TICKMARKS_REMOVE_4 = 51308,
	TICKMARKS_SET_KEY = 51309,
	INVALID_INTERVAL = 51310,
	INVALID_LAYER_INDEX = 51311,
	GETINDEXANDKEY_ITEM1 = 51312,
	GETINDEXANDKEY_ITEM2 = 51313,
	GETINDEXANDKEY_ITEM3 = 51314,
	GETINDEXANDKEY_ITEM4 = 51315,
	INVALID_ROW_KEY = 51316,
	STYLE_INVALID_INDEX = 51317,
	STYLE_INVALID_KEY = 51318,
	MP_REMOVE_1 = 51596,
	MP_REMOVE_2 = 51597,
	MP_REMOVE_3 = 51598,
	MP_REMOVE_4 = 51599,
	MP_ITEM_1 = 51600,
	MP_ITEM_2 = 51601,
	MP_ITEM_3 = 51602,
	MP_ITEM_4 = 51603,
	MP_ADD_1 = 51604,
	MP_ADD_2 = 51605,
	MP_ADD_3 = 51606,
	MP_SET_KEY = 51607,
	SPLITTER_INVALIDOP = 51608,
	SPLITTER_INVALID_INDEX = 51609,
	SPLITTER_INVALID_WIDTH = 51610,
    STYLE_NULL = 51611,
    INVALID_DURATION_INTERVAL = 51612,
    CHECK_DURATION_ERROR = 51613,
	ERR_DURATION_INCONSISTENT = 51614,
	PROGRESSLINE_INVALIDOP = 51615,
	PROGRESSLINE_INVALID_INDEX = 51616,
	PROGRESSLINE_INVALID_WIDTH = 51617,
	PRINTER_INVALID_DATE_RANGE = 51618,
	PRINTER_NO_ROWS = 51619,
	PRINTER_INVALID_ROW_RANGE = 51620,
	PRINTER_INVALID_COLUMN = 51621,
	PRINTER_INVALID_ROW = 51622,
	PRINTER_INVALID_PAGE = 51623,
    PRINTER_INVALID_PAGE_RANGE = 51624,
	PRINTER_INVALID_SCALE = 51625,
    PRINTER_INVALID_SPECS_MARGINS = 51626,
    PRINTER_INVALID_SPECS_HEIGHT = 51627,
    PRINTER_INVALID_SPECS_COLUMNSINEVERYPAGE = 51628,
    PRINTER_INVALID_SPECS_COLUMNS = 51629,
    PRINTER_INVALID_SPECS_INTERVAL = 51630,
	OBJECT_COUNT_ERROR = 51631,
	ERR_RETARRELEMINDEX_G = 51632,
	NODE_INVALID_DEPTH = 51633,
	ROWS_INVALID_TREE_STRUCTURE = 51634,
}SYS_ERRORS;

typedef enum E_OPERATION
{
    EO_NONE = 0,

    EO_TASKADDITION = 1,
    EO_TASKMOVEMENT = 2,
    EO_TASKSTRETCHLEFT = 3,
    EO_TASKSTRETCHRIGHT = 4,
    EO_TASKSELECTION = 5,

    EO_MILESTONEADDITION = 6,


    EO_ROWSIZING = 9,
    EO_ROWMOVEMENT = 10,
    EO_ROWSELECTION = 11,

    EO_COLUMNSIZING = 12,
    EO_COLUMNMOVEMENT = 13,
    EO_COLUMNSELECTION = 14,



    EO_TIMELINEMOVEMENT = 16,


    EO_SPLITTERMOVEMENT = 18,

    EO_HORIZONTALSCROLLBAR = 19,
    EO_VERTICALSCROLLBAR = 20,
    EO_TIMELINESCROLLBAR = 21,
    EO_TREEVIEW = 22,

    EO_PERCENTAGESELECTION = 23,
    EO_PERCENTAGESIZING = 24,

    //Treeview
    EO_CHECKBOXCLICK = 27,
    EO_SIGNCLICK = 28,

	EO_PREDECESSORADDITION = 29,
	EO_PREDECESSORSELECTION = 30
		
}E_OPERATION;

typedef enum E_MOUSEBUTTONS
{
	BTN_NONE = 0,
	BTN_LEFT = 1048576,
	BTN_RIGHT = 2097152,
	BTN_MIDDLE = 4194304,
	BTN_XBUTTON1 = 8388608,
	BTN_XBUTTON2 = 16777216
}E_MOUSEBUTTONS;

typedef enum E_CONSTRAINTTYPE
{
	PCT_END_TO_START = 0,
	PCT_START_TO_START = 1,
	PCT_END_TO_END = 2,
	PCT_START_TO_END = 3
}E_CONSTRAINTTYPE;

typedef enum E_TOOLTIPTYPE
{
	TPT_HOVER = 0,
	TPT_MOVEMENT = 1
}E_TOOLTIPTYPE;

typedef enum E_TIMEBLOCKTYPE
{
	TBT_SINGLE_OCURRENCE = 0,
	TBT_RECURRING = 1
}E_TIMEBLOCKTYPE;

typedef enum E_RECURRINGTYPE
{
    RCT_DAY = 3,
    RCT_WEEK = 4,
    RCT_MONTH = 5,
    RCT_YEAR = 7
}E_RECURRINGTYPE;

typedef enum E_WEEKDAY
{
	WD_SUNDAY = 0,
	WD_MONDAY = 1,
	WD_TUESDAY = 2,
	WD_WEDNESDAY = 3,
	WD_THURSDAY = 4,
	WD_FRIDAY = 5,
	WD_SATURDAY = 6
}E_WEEKDAY;

typedef enum E_CALENDARWEEKRULES
{
    CWR_FIRSTDAY = 0,
    CWR_FIRSTFULLWEEK = 1,
    CWR_FIRSTFOURDAYWEEK = 2
}E_CALENDARWEEKRULES;

typedef enum E_TEXTPLACEMENT
{
	SCP_OBJECTEXTENTSPLACEMENT = 0,
	SCP_OFFSETPLACEMENT = 1,
	SCP_EXTERIORPLACEMENT = 2
}E_TEXTPLACEMENT;

typedef enum E_PROGRESSLINELENGTH
{
    TLMA_TICKMARKAREA = 0,
    TLMA_CLIENTAREA = 1,
    TLMA_TIMELINE = 2,
    TLMA_CA_TIMELINE = 3,
    TLMA_CA_TICKMARKAREA = 4,
    TLMA_NONE = 5
}E_PROGRESSLINELENGTH;

typedef enum E_PROGRESSLINESTYLE
{
    PLT_STANDARD = 1,
    PLT_USERDEFINED = 2,
    PLT_STYLE = 3
}E_PROGRESSLINESTYLE;

typedef enum E_PROGRESSLINETYPE
{
	TLMT_SYSTEMTIME = 0,
	TLMT_USER = 1
}E_PROGRESSLINETYPE;

typedef enum E_LAYEROBJECTENABLE
{
	EC_INCURRENTLAYERONLY = 0,
	EC_INALLLAYERS = 1
}E_LAYEROBJECTENABLE;

typedef enum GRE_FONTSTYLE
{
    FS_FONTSTYLEREGULAR = 0,
    FS_FONTSTYLEBOLD = 1,
    FS_FONTSTYLEITALIC = 2,
    FS_FONTSTYLEBOLDITALIC = 3,
    FS_FONTSTYLEUNDERLINE = 4,
    FS_FONTSTYLESTRIKEOUT = 8
}GRE_FONTSTYLE;

typedef enum GRE_UNIT
{
    UT_UNITWORLD = 0,
    UT_UNITDISPLAY = 1,
    UT_UNITPIXEL = 2,
    UT_UNITPOINT = 3,
    UT_UNITINCH = 4,
    UT_UNITDOCUMENT = 5,
    UT_UNITMILLIMETER = 6
}GRE_UNIT;

typedef enum E_INTERVAL
{
    IL_SECOND = 3,
    IL_MINUTE = 4,
    IL_HOUR = 5,
    IL_DAY = 6,
    IL_WEEK = 7,
    IL_MONTH = 8,
    IL_QUARTER = 9,
    IL_YEAR = 10
}E_INTERVAL;

typedef enum E_SPLITTERTYPE
{
	SA_APPEARANCE = 1,
	SA_USERDEFINED = 2,
	SA_STYLE = 3
}E_SPLITTERTYPE;

typedef enum E_TIERBACKGROUNDMODE
{
	ETBGM_TIERAPPEARANCE = 0,
	ETBGM_STYLE = 2
}E_TIERBACKGROUNDMODE;

typedef enum E_OBJECTSCOPE
{
    OS_CONTROL = 0,
    OS_VIEW = 1
}E_OBJECTSCOPE;


typedef enum E_OBJECTTEMPLATE
{
    STO_NONE = 0,
    STO_BW_HATCH = 1,
    STO_COLOR_HATCH = 2,
    STO_GREY_SOLID = 3,
    STO_COLORS_SOLID = 4,
    STO_DEFAULT = 5,
    STO_VARIATION_1 = 6
}E_OBJECTTEMPLATE;

typedef enum E_CONTROLTEMPLATE
{
    STC_NONE = 0,
    STC_CH_SOLID_WHITE = 1,
    STC_CH_SOLID_DARK_BLUE = 2,
    STC_CH_SOLID_VIOLET = 3,
    STC_CH_SOLID_GREEN = 4,
    STC_CH_SOLID_RED = 5,
    STC_CH_SOLID_LIGHT_BLUE = 6,
    STC_CH_SOLID_GREY = 7,
    STC_CH_SOLID_LIGHT_STEEL_BLUE = 8,
    STC_CH_VGRAD_YELLOW = 9,
    STC_CH_VGRAD_ORANGE = 10,
    STC_CH_VGRAD_RED = 11,
    STC_CH_VGRAD_CRIMSON = 12,
    STC_CH_VGRAD_MAGENTA = 13,
    STC_CH_VGRAD_MULBERRY = 14,
    STC_CH_VGRAD_BLUE_VIOLET = 15,
    STC_CH_VGRAD_ANAKIWA_BLUE = 16,
    STC_CH_VGRAD_BLUE_BELL = 17,
    STC_CH_VGRAD_BLUE = 18,
    STC_CH_VGRAD_AERO = 19,
    STC_CH_HGRAD_YELLOW = 20,
    STC_CH_HGRAD_ORANGE = 21,
    STC_CH_HGRAD_RED = 22,
    STC_CH_HGRAD_CRIMSON = 23,
    STC_CH_HGRAD_MAGENTA = 24,
    STC_CH_HGRAD_MULBERRY = 25,
    STC_CH_HGRAD_BLUE_VIOLET = 26,
    STC_CH_HGRAD_ANAKIWA_BLUE = 27,
    STC_CH_HGRAD_BLUE_BELL = 28,
    STC_CH_HGRAD_BLUE = 29,
    STC_CH_HGRAD_AERO = 30
}E_CONTROLTEMPLATE;

typedef enum E_SELECTIONRECTANGLEMODE
{
    SRM_DOTTED = 0,
    SRM_COLOR = 1
}E_SELECTIONRECTANGLEMODE;

typedef enum E_KEYS
{
	KEYS_NONE = 0,
	KEYS_LBUTTON = 1,
	KEYS_RBUTTON = 2,
	KEYS_CANCEL = 3,
	KEYS_MBUTTON = 4,
	KEYS_XBUTTON1 = 5,
	KEYS_XBUTTON2 = 6,
	KEYS_LBUTTON_XBUTTON2 = 7,
	KEYS_BACK = 8,
	KEYS_TAB = 9,
	KEYS_LINEFEED = 10,
	KEYS_LBUTTON_LINEFEED = 11,
	KEYS_CLEAR = 12,
	KEYS_ENTER = 13,
	KEYS_RBUTTON_CLEAR = 14,
	KEYS_RBUTTON_ENTER = 15,
	KEYS_SHIFTKEY = 16,
	KEYS_CONTROLKEY = 17,
	KEYS_MENU = 18,
	KEYS_PAUSE = 19,
	KEYS_CAPITAL = 20,
	KEYS_HANGULMODE = 21,
	KEYS_RBUTTON_CAPITAL = 22,
	KEYS_JUNJAMODE = 23,
	KEYS_FINALMODE = 24,
	KEYS_HANJAMODE = 25,
	KEYS_RBUTTON_FINALMODE = 26,
	KEYS_ESCAPE = 27,
	KEYS_IMECONVERT = 28,
	KEYS_IMENONCONVERT = 29,
	KEYS_IMEACEEPT = 30,
	KEYS_IMEMODECHANGE = 31,
	KEYS_SPACE = 32,
	KEYS_PRIOR = 33,
	KEYS_NEXT = 34,
	KEYS_END = 35,
	KEYS_HOME = 36,
	KEYS_LEFT = 37,
	KEYS_UP = 38,
	KEYS_RIGHT = 39,
	KEYS_DOWN = 40,
	KEYS_SELECT = 41,
	KEYS_PRINT = 42,
	KEYS_EXECUTE = 43,
	KEYS_SNAPSHOT = 44,
	KEYS_INSERT = 45,
	KEYS_DELETE = 46,
	KEYS_HELP = 47,
	KEYS_D0 = 48,
	KEYS_D1 = 49,
	KEYS_D2 = 50,
	KEYS_D3 = 51,
	KEYS_D4 = 52,
	KEYS_D5 = 53,
	KEYS_D6 = 54,
	KEYS_D7 = 55,
	KEYS_D8 = 56,
	KEYS_D9 = 57,
	KEYS_RBUTTON_D8 = 58,
	KEYS_RBUTTON_D9 = 59,
	KEYS_MBUTTON_D8 = 60,
	KEYS_MBUTTON_D9 = 61,
	KEYS_XBUTTON2_D8 = 62,
	KEYS_XBUTTON2_D9 = 63,
	KEYS_64 = 64,
	KEYS_A = 65,
	KEYS_B = 66,
	KEYS_C = 67,
	KEYS_D = 68,
	KEYS_E = 69,
	KEYS_F = 70,
	KEYS_G = 71,
	KEYS_H = 72,
	KEYS_I = 73,
	KEYS_J = 74,
	KEYS_K = 75,
	KEYS_L = 76,
	KEYS_M = 77,
	KEYS_N = 78,
	KEYS_O = 79,
	KEYS_P = 80,
	KEYS_Q = 81,
	KEYS_R = 82,
	KEYS_S = 83,
	KEYS_T = 84,
	KEYS_U = 85,
	KEYS_V = 86,
	KEYS_W = 87,
	KEYS_X = 88,
	KEYS_Y = 89,
	KEYS_Z = 90,
	KEYS_LWIN = 91,
	KEYS_RWIN = 92,
	KEYS_APPS = 93,
	KEYS_RBUTTON_RWIN = 94,
	KEYS_RBUTTON_APPS = 95,
	KEYS_NUMPAD0 = 96,
	KEYS_NUMPAD1 = 97,
	KEYS_NUMPAD2 = 98,
	KEYS_NUMPAD3 = 99,
	KEYS_NUMPAD4 = 100,
	KEYS_NUMPAD5 = 101,
	KEYS_NUMPAD6 = 102,
	KEYS_NUMPAD7 = 103,
	KEYS_NUMPAD8 = 104,
	KEYS_NUMPAD9 = 105,
	KEYS_MULTIPLY = 106,
	KEYS_ADD = 107,
	KEYS_SEPARATOR = 108,
	KEYS_SUBTRACT = 109,
	KEYS_DECIMAL = 110,
	KEYS_DIVIDE = 111,
	KEYS_F1 = 112,
	KEYS_F2 = 113,
	KEYS_F3 = 114,
	KEYS_F4 = 115,
	KEYS_F5 = 116,
	KEYS_F6 = 117,
	KEYS_F7 = 118,
	KEYS_F8 = 119,
	KEYS_F9 = 120,
	KEYS_F10 = 121,
	KEYS_F11 = 122,
	KEYS_F12 = 123,
	KEYS_F13 = 124,
	KEYS_F14 = 125,
	KEYS_F15 = 126,
	KEYS_F16 = 127,
	KEYS_F17 = 128,
	KEYS_F18 = 129,
	KEYS_F19 = 130,
	KEYS_F20 = 131,
	KEYS_F21 = 132,
	KEYS_F22 = 133,
	KEYS_F23 = 134,
	KEYS_F24 = 135,
	KEYS_BACK_F17 = 136,
	KEYS_BACK_F18 = 137,
	KEYS_BACK_F19 = 138,
	KEYS_BACK_F20 = 139,
	KEYS_BACK_F21 = 140,
	KEYS_BACK_F22 = 141,
	KEYS_BACK_F23 = 142,
	KEYS_BACK_F24 = 143,
	KEYS_NUMLOCK = 144,
	KEYS_SCROLL = 145,
	KEYS_RBUTTON_NUMLOCK = 146,
	KEYS_RBUTTON_SCROLL = 147,
	KEYS_MBUTTON_NUMLOCK = 148,
	KEYS_MBUTTON_SCROLL = 149,
	KEYS_XBUTTON2_NUMLOCK = 150,
	KEYS_XBUTTON2_SCROLL = 151,
	KEYS_BACK_NUMLOCK = 152,
	KEYS_BACK_SCROLL = 153,
	KEYS_LINEFEED_NUMLOCK = 154,
	KEYS_LINEFEED_SCROLL = 155,
	KEYS_CLEAR_NUMLOCK = 156,
	KEYS_CLEAR_SCROLL = 157,
	KEYS_RBUTTON_CLEAR_NUMLOCK = 158,
	KEYS_RBUTTON_CLEAR_SCROLL = 159,
	KEYS_LSHIFTKEY = 160,
	KEYS_RSHIFTKEY = 161,
	KEYS_LCONTROLKEY = 162,
	KEYS_RCONTROLKEY = 163,
	KEYS_LMENU = 164,
	KEYS_RMENU = 165,
	KEYS_BROWSERBACK = 166,
	KEYS_BROWSERFORWARD = 167,
	KEYS_BROWSERREFRESH = 168,
	KEYS_BROWSERSTOP = 169,
	KEYS_BROWSERSEARCH = 170,
	KEYS_BROWSERFAVORITES = 171,
	KEYS_BROWSERHOME = 172,
	KEYS_VOLUMEMUTE = 173,
	KEYS_VOLUMEDOWN = 174,
	KEYS_VOLUMEUP = 175,
	KEYS_MEDIANEXTTRACK = 176,
	KEYS_MEDIAPREVIOUSTRACK = 177,
	KEYS_MEDIASTOP = 178,
	KEYS_MEDIAPLAYPAUSE = 179,
	KEYS_LAUNCHMAIL = 180,
	KEYS_SELECTMEDIA = 181,
	KEYS_LAUNCHAPPLICATION1 = 182,
	KEYS_LAUNCHAPPLICATION2 = 183,
	KEYS_BACK_MEDIANEXTTRACK = 184,
	KEYS_BACK_MEDIAPREVIOUSTRACK = 185,
	KEYS_OEMSEMICOLON = 186,
	KEYS_OEMPLUS = 187,
	KEYS_OEMCOMMA = 188,
	KEYS_OEMMINUS = 189,
	KEYS_OEMPERIOD = 190,
	KEYS_OEMQUESTION = 191,
	KEYS_OEMTILDE = 192,
	KEYS_LBUTTON_OEMTILDE = 193,
	KEYS_RBUTTON_OEMTILDE = 194,
	KEYS_CANCEL_OEMTILDE = 195,
	KEYS_MBUTTON_OEMTILDE = 196,
	KEYS_XBUTTON1_OEMTILDE = 197,
	KEYS_XBUTTON2_OEMTILDE = 198,
	KEYS_LBUTTON_XBUTTON2_OEMTILDE = 199,
	KEYS_BACK_OEMTILDE = 200,
	KEYS_TAB_OEMTILDE = 201,
	KEYS_LINEFEED_OEMTILDE = 202,
	KEYS_LBUTTON_LINEFEED_OEMTILDE = 203,
	KEYS_CLEAR_OEMTILDE = 204,
	KEYS_ENTER_OEMTILDE = 205,
	KEYS_RBUTTON_CLEAR_OEMTILDE = 206,
	KEYS_RBUTTON_ENTER_OEMTILDE = 207,
	KEYS_SHIFTKEY_OEMTILDE = 208,
	KEYS_CONTROLKEY_OEMTILDE = 209,
	KEYS_MENU_OEMTILDE = 210,
	KEYS_PAUSE_OEMTILDE = 211,
	KEYS_CAPITAL_OEMTILDE = 212,
	KEYS_HANGULMODE_OEMTILDE = 213,
	KEYS_RBUTTON_CAPITAL_OEMTILDE = 214,
	KEYS_JUNJAMODE_OEMTILDE = 215,
	KEYS_FINALMODE_OEMTILDE = 216,
	KEYS_HANJAMODE_OEMTILDE = 217,
	KEYS_RBUTTON_FINALMODE_OEMTILDE = 218,
	KEYS_OEMOPENBRACKETS = 219,
	KEYS_OEMPIPE = 220,
	KEYS_OEMCLOSEBRACKETS = 221,
	KEYS_OEMQUOTES = 222,
	KEYS_OEM8 = 223,
	KEYS_SPACE_OEMTILDE = 224,
	KEYS_PRIOR_OEMTILDE = 225,
	KEYS_OEMBACKSLASH = 226,
	KEYS_LBUTTON_OEMBACKSLASH = 227,
	KEYS_HOME_OEMTILDE = 228,
	KEYS_PROCESSKEY = 229,
	KEYS_MBUTTON_OEMBACKSLASH = 230,
	KEYS_RBUTTON_PROCESSKEY = 231,
	KEYS_DOWN_OEMTILDE = 232,
	KEYS_SELECT_OEMTILDE = 233,
	KEYS_BACK_OEMBACKSLASH = 234,
	KEYS_TAB_OEMBACKSLASH = 235,
	KEYS_SNAPSHOT_OEMTILDE = 236,
	KEYS_BACK_PROCESSKEY = 237,
	KEYS_CLEAR_OEMBACKSLASH = 238,
	KEYS_LINEFEED_PROCESSKEY = 239,
	KEYS_D0_OEMTILDE = 240,
	KEYS_D1_OEMTILDE = 241,
	KEYS_SHIFTKEY_OEMBACKSLASH = 242,
	KEYS_CONTROLKEY_OEMBACKSLASH = 243,
	KEYS_D4_OEMTILDE = 244,
	KEYS_SHIFTKEY_PROCESSKEY = 245,
	KEYS_ATTN = 246,
	KEYS_CRSEL = 247,
	KEYS_EXSEL = 248,
	KEYS_ERASEEOF = 249,
	KEYS_PLAY = 250,
	KEYS_ZOOM = 251,
	KEYS_NONAME = 252,
	KEYS_PA1 = 253,
	KEYS_OEMCLEAR = 254,
	KEYS_LBUTTON_OEMCLEAR = 255,
}E_KEYS;
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

typedef enum ParamType
{
    PT_STRING = 0,
    PT_STRING_EMPTYISNULL = 1,
    PT_NUMERIC = 2,
    PT_NUMERIC_ZEROISNULL = 3,
    PT_BOOL = 4,
    PT_DATE = 5,
    PT_NULLVALUE = 6,
    PT_UNIQUEIDENTIFIER = 7,
    PT_DATE_EMPTYISNULL = 8
}ParamType;

typedef enum E_DATASOURCETYPE
{
	DST_NONE = 0,
	DST_ACCESS = 1,
	DST_XML = 2
}E_DATASOURCETYPE;

typedef enum HPE_ADDMODE
{
    AM_RESERVATION = 0,
    AM_RENTAL = 1,
    AM_MAINTENANCE = 2
}HPE_ADDMODE;

typedef enum PRG_DIALOGMODE
{
    DM_ADD = 0,
    DM_EDIT = 1,
}PRG_DIALOGMODE;

struct S_tb_CR_Rows
{
	LONG lRowID;
    LONG lDepth;
    LONG lOrder;
    CString sLicensePlates;
    LONG lCarTypeID;
    CString sNotes;
    DOUBLE cRate;
    CString sACRISSCode;
    CString sCity;
    CString sBranchName;
    CString sState;
    CString sPhone;
    CString sManagerName;
    CString sManagerMobile;
    CString sAddress;
    CString sZIP;
};

typedef CList<S_tb_CR_Rows, S_tb_CR_Rows> T_tb_CR_Rows;

struct S_tb_CR_Rentals
{
	LONG lTaskID;
    LONG lRowID;
    HPE_ADDMODE yMode;
    CString sName;
    CString sAddress;
    CString sCity;
    CString sState;
    CString sZIP;
    CString sPhone;
    CString sMobile;
	COleDateTime dtPickUp;
	COleDateTime dtReturn;
    DOUBLE cRate;
    DOUBLE cALI;
    DOUBLE dCRF;
    DOUBLE cERF;
    DOUBLE cGPS;
    DOUBLE cLDW;
    DOUBLE cPAI;
    DOUBLE cPEP;
    DOUBLE cRCFC;
    DOUBLE cVLF;
    DOUBLE cWTB;
    DOUBLE dTax;
    DOUBLE cEstimatedTotal;
    BOOL bGPS;
    BOOL bFSO;
    BOOL bLDW;
    BOOL bPAI;
    BOOL bPEP;
    BOOL bALI;
};

typedef CList<S_tb_CR_Rentals, S_tb_CR_Rentals> T_tb_CR_Rentals;

struct S_tb_CR_Car_Types
{
	LONG lCarTypeID;
    CString sDescription;
    CString sACRISSCode;
    DOUBLE cStdRate;
};

typedef CList<S_tb_CR_Car_Types, S_tb_CR_Car_Types> T_tb_CR_Car_Types;

struct S_tb_CR_US_States
{
	CString ID;
    CString Name;
    DOUBLE dCarRentalTax;
};

typedef CList<S_tb_CR_US_States, S_tb_CR_US_States> T_tb_CR_US_States;

struct S_tb_CR_ACRISS_Codes
{
	LONG ID;
    CString Letter;
    CString Description;
    LONG Position;
};

typedef CList<S_tb_CR_ACRISS_Codes, S_tb_CR_ACRISS_Codes> T_tb_CR_ACRISS_Codes;

struct S_tb_CR_Taxes_Surcharges_Options
{
	CString sID;
    CString sDescription;
    DOUBLE sRate;
};

typedef CList<S_tb_CR_Taxes_Surcharges_Options, S_tb_CR_Taxes_Surcharges_Options> T_tb_CR_Taxes_Surcharges_Options;


struct S_tb_US_Cities
{
	LONG ID;
	CString Name;
	CString State;
};

typedef CList<S_tb_US_Cities, S_tb_US_Cities> T_tb_US_Cities;

struct S_tb_US_Streets
{
	LONG ID;
	CString Name;
	CString USPS;
};

typedef CList<S_tb_US_Streets, S_tb_US_Streets> T_tb_US_Streets;

struct S_tb_First_Names
{
	LONG ID;
	CString Name;
	CString sSex;
};

typedef CList<S_tb_First_Names, S_tb_First_Names> T_tb_First_Names;

struct S_tb_Last_Names
{
	LONG ID;
	CString LastName;
};

typedef CList<S_tb_Last_Names, S_tb_Last_Names> T_tb_Last_Names;
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

class TickMarkTextDrawEventArgs : public CCmdTarget
{
	DECLARE_DYNCREATE(TickMarkTextDrawEventArgs)

	public:
		TickMarkTextDrawEventArgs(void);
		virtual ~TickMarkTextDrawEventArgs();
		virtual void OnFinalRelease();

	// Member Variables
	public:
		CString mp_sText;
		BOOL mp_bCustomDraw;
		COleDateTime mp_dtDate;


	//Internal Properties (Extensions to ODL)
	public:

		CString GetText(void);
		void SetText(CString Value);
		BOOL GetCustomDraw(void);
		void SetCustomDraw(BOOL Value);


	//Internal Properties
	public:

	//Internal Methods
	public:
		void Clear(void);

	protected:
		DECLARE_DISPATCH_MAP()
		DECLARE_MESSAGE_MAP()
		DECLARE_OLECREATE(TickMarkTextDrawEventArgs)
		DECLARE_INTERFACE_MAP()

	BSTR odl_GetText(void)
	{
		return GetText().AllocSysString();
	}

	void odl_SetText(LPCTSTR Value)
	{
		SetText(Value);
	}

	BOOL odl_GetCustomDraw(void)
	{
		return GetCustomDraw();
	}

	void odl_SetCustomDraw(BOOL Value)
	{	
		SetCustomDraw(Value);
	}

	DATE odl_GetdtDate(void)
	{
		return mp_dtDate;
	}

};
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


#include <gdiplus.h>



class xAuxInitGDIPlus 
{
private:
	HANDLE							m_hMap;
	bool							m_bInited, m_bInitCtorDtor;
	ULONG_PTR						m_gdiplusToken;
	ULONG_PTR						m_gdiplusBGThreadToken;
	Gdiplus::GdiplusStartupInput	m_gdiplusStartupInput;
	Gdiplus::GdiplusStartupOutput   m_gdiplusStartupOutput;
	long							m_initcount;

public:
	xAuxInitGDIPlus(bool bInitCtorDtor = false) : m_bInitCtorDtor(bInitCtorDtor), m_bInited(false), m_hMap(NULL), m_gdiplusToken(NULL), m_gdiplusStartupInput(NULL), m_initcount(0)  
	{
		if (m_bInitCtorDtor) 
		{
			Initialize();
		}
	}

	virtual ~xAuxInitGDIPlus()  
	{
		if (m_bInitCtorDtor) 
		{
			Deinitialize();
		}
	}

	void Initialize() 
	{
		if (!m_bInited) 
		{
			TCHAR buffer[1024];
			m_hMap = CreateFileMapping((HANDLE) INVALID_HANDLE_VALUE, NULL, PAGE_READWRITE | SEC_COMMIT, 0, sizeof(long), buffer);
			if (m_hMap != NULL) 
			{
				if (GetLastError() == ERROR_ALREADY_EXISTS) 
				{ 
					CloseHandle(m_hMap); 
				} 
				else 
				{
					m_bInited = true;
					m_gdiplusStartupInput.SuppressBackgroundThread = TRUE;
					Gdiplus::GdiplusStartup(&m_gdiplusToken, &m_gdiplusStartupInput, &m_gdiplusStartupOutput);
					Gdiplus::Status status = m_gdiplusStartupOutput.NotificationHook(&m_gdiplusBGThreadToken);
				}
			}
		}
		m_initcount++;
	}

	void Deinitialize()
	{
		m_initcount--;
		if (m_bInited && m_initcount == 0) 
		{
			m_gdiplusStartupOutput.NotificationUnhook(m_gdiplusBGThreadToken);
			Gdiplus::GdiplusShutdown(m_gdiplusToken);
			CloseHandle(m_hMap);
			m_bInited = false;
		}
	}
};
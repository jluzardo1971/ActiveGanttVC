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
#include "afxdb.h"

CString g_DST_ACCESS_GetConnectionString();


COleDateTime GetDateField(CRecordset &oRecordset, CString sFieldName);
BOOL GetBoolField(CRecordset &oRecordset, CString sFieldName);
float GetFloatField(CRecordset &oRecordset, CString sFieldName);
CString GetStrField(CRecordset &oRecordset, CString sFieldName);
int GetIntField(CRecordset &oRecordset, CString sFieldName);
short GetShortField(CRecordset &oRecordset, CString sFieldName);
double GetDblField(CRecordset &oRecordset, CString sFieldName);

CString g_DST_ACCESS_ConvertDate(COleDateTime dtDate);

void g_DST_ACCESS_FillComboBox(CComboBox* oControl, CString sSQL, CString sValueMemberField, CString sDisplayMemberField, CDatabase &oConn);
void g_DST_ACCESS_FillComboBoxString(CComboBox* oControl, CString sSQL, CString sValueMemberField, CString sDisplayMemberField, CDatabase &oConn);
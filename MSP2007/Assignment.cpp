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
#include "stdafx.h"
#include "clsXML.h"
#include "Assignment.h"

IMPLEMENT_DYNCREATE(Assignment, CCmdTarget)

//{A77FAFB6-DCF6-4152-9367-B7B26C367C95}
static const IID IID_IAssignment = { 0xA77FAFB6, 0xDCF6, 0x4152, { 0x93, 0x67, 0xB7, 0xB2, 0x6C, 0x36, 0x7C, 0x95} };

//{DC18266E-47F6-499D-B438-812C16428C73}
IMPLEMENT_OLECREATE_FLAGS(Assignment, "MSP2007.Assignment", afxRegApartmentThreading, 0xDC18266E, 0x47F6, 0x499D, 0xB4, 0x38, 0x81, 0x2C, 0x16, 0x42, 0x8C, 0x73)

BEGIN_DISPATCH_MAP(Assignment, CCmdTarget)
DISP_PROPERTY_EX_ID(Assignment, "lUID", 1, odl_GetUID, odl_SetUID, VT_I4)
DISP_PROPERTY_EX_ID(Assignment, "lTaskUID", 2, odl_GetTaskUID, odl_SetTaskUID, VT_I4)
DISP_PROPERTY_EX_ID(Assignment, "lResourceUID", 3, odl_GetResourceUID, odl_SetResourceUID, VT_I4)
DISP_PROPERTY_EX_ID(Assignment, "lPercentWorkComplete", 4, odl_GetPercentWorkComplete, odl_SetPercentWorkComplete, VT_I4)
DISP_PROPERTY_EX_ID(Assignment, "cActualCost", 5, odl_GetActualCost, odl_SetActualCost, VT_R8)
DISP_PROPERTY_EX_ID(Assignment, "dtActualFinish", 6, odl_GetActualFinish, odl_SetActualFinish, VT_DATE)
DISP_PROPERTY_EX_ID(Assignment, "cActualOvertimeCost", 7, odl_GetActualOvertimeCost, odl_SetActualOvertimeCost, VT_R8)
DISP_PROPERTY_EX_ID(Assignment, "oActualOvertimeWork", 8, odl_GetActualOvertimeWork, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(Assignment, "dtActualStart", 9, odl_GetActualStart, odl_SetActualStart, VT_DATE)
DISP_PROPERTY_EX_ID(Assignment, "oActualWork", 10, odl_GetActualWork, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(Assignment, "fACWP", 11, odl_GetACWP, odl_SetACWP, VT_R4)
DISP_PROPERTY_EX_ID(Assignment, "bConfirmed", 12, odl_GetConfirmed, odl_SetConfirmed, VT_BOOL)
DISP_PROPERTY_EX_ID(Assignment, "cCost", 13, odl_GetCost, odl_SetCost, VT_R8)
DISP_PROPERTY_EX_ID(Assignment, "yCostRateTable", 14, odl_GetCostRateTable, odl_SetCostRateTable, VT_I4)
DISP_PROPERTY_EX_ID(Assignment, "fCostVariance", 15, odl_GetCostVariance, odl_SetCostVariance, VT_R4)
DISP_PROPERTY_EX_ID(Assignment, "fCV", 16, odl_GetCV, odl_SetCV, VT_R4)
DISP_PROPERTY_EX_ID(Assignment, "lDelay", 17, odl_GetDelay, odl_SetDelay, VT_I4)
DISP_PROPERTY_EX_ID(Assignment, "dtFinish", 18, odl_GetFinish, odl_SetFinish, VT_DATE)
DISP_PROPERTY_EX_ID(Assignment, "lFinishVariance", 19, odl_GetFinishVariance, odl_SetFinishVariance, VT_I4)
DISP_PROPERTY_EX_ID(Assignment, "sHyperlink", 20, odl_GetHyperlink, odl_SetHyperlink, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sHyperlinkAddress", 21, odl_GetHyperlinkAddress, odl_SetHyperlinkAddress, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sHyperlinkSubAddress", 22, odl_GetHyperlinkSubAddress, odl_SetHyperlinkSubAddress, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "fWorkVariance", 23, odl_GetWorkVariance, odl_SetWorkVariance, VT_R4)
DISP_PROPERTY_EX_ID(Assignment, "bHasFixedRateUnits", 24, odl_GetHasFixedRateUnits, odl_SetHasFixedRateUnits, VT_BOOL)
DISP_PROPERTY_EX_ID(Assignment, "bFixedMaterial", 25, odl_GetFixedMaterial, odl_SetFixedMaterial, VT_BOOL)
DISP_PROPERTY_EX_ID(Assignment, "lLevelingDelay", 26, odl_GetLevelingDelay, odl_SetLevelingDelay, VT_I4)
DISP_PROPERTY_EX_ID(Assignment, "yLevelingDelayFormat", 27, odl_GetLevelingDelayFormat, odl_SetLevelingDelayFormat, VT_I4)
DISP_PROPERTY_EX_ID(Assignment, "bLinkedFields", 28, odl_GetLinkedFields, odl_SetLinkedFields, VT_BOOL)
DISP_PROPERTY_EX_ID(Assignment, "bMilestone", 29, odl_GetMilestone, odl_SetMilestone, VT_BOOL)
DISP_PROPERTY_EX_ID(Assignment, "sNotes", 30, odl_GetNotes, odl_SetNotes, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "bOverallocated", 31, odl_GetOverallocated, odl_SetOverallocated, VT_BOOL)
DISP_PROPERTY_EX_ID(Assignment, "cOvertimeCost", 32, odl_GetOvertimeCost, odl_SetOvertimeCost, VT_R8)
DISP_PROPERTY_EX_ID(Assignment, "oOvertimeWork", 33, odl_GetOvertimeWork, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(Assignment, "fPeakUnits", 34, odl_GetPeakUnits, odl_SetPeakUnits, VT_R4)
DISP_PROPERTY_EX_ID(Assignment, "oRegularWork", 35, odl_GetRegularWork, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(Assignment, "cRemainingCost", 36, odl_GetRemainingCost, odl_SetRemainingCost, VT_R8)
DISP_PROPERTY_EX_ID(Assignment, "cRemainingOvertimeCost", 37, odl_GetRemainingOvertimeCost, odl_SetRemainingOvertimeCost, VT_R8)
DISP_PROPERTY_EX_ID(Assignment, "oRemainingOvertimeWork", 38, odl_GetRemainingOvertimeWork, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(Assignment, "oRemainingWork", 39, odl_GetRemainingWork, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(Assignment, "bResponsePending", 40, odl_GetResponsePending, odl_SetResponsePending, VT_BOOL)
DISP_PROPERTY_EX_ID(Assignment, "dtStart", 41, odl_GetStart, odl_SetStart, VT_DATE)
DISP_PROPERTY_EX_ID(Assignment, "dtStop", 42, odl_GetStop, odl_SetStop, VT_DATE)
DISP_PROPERTY_EX_ID(Assignment, "dtResume", 43, odl_GetResume, odl_SetResume, VT_DATE)
DISP_PROPERTY_EX_ID(Assignment, "lStartVariance", 44, odl_GetStartVariance, odl_SetStartVariance, VT_I4)
DISP_PROPERTY_EX_ID(Assignment, "bSummary", 45, odl_GetSummary, odl_SetSummary, VT_BOOL)
DISP_PROPERTY_EX_ID(Assignment, "fSV", 46, odl_GetSV, odl_SetSV, VT_R4)
DISP_PROPERTY_EX_ID(Assignment, "fUnits", 47, odl_GetUnits, odl_SetUnits, VT_R4)
DISP_PROPERTY_EX_ID(Assignment, "bUpdateNeeded", 48, odl_GetUpdateNeeded, odl_SetUpdateNeeded, VT_BOOL)
DISP_PROPERTY_EX_ID(Assignment, "fVAC", 49, odl_GetVAC, odl_SetVAC, VT_R4)
DISP_PROPERTY_EX_ID(Assignment, "oWork", 50, odl_GetWork, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(Assignment, "yWorkContour", 51, odl_GetWorkContour, odl_SetWorkContour, VT_I4)
DISP_PROPERTY_EX_ID(Assignment, "fBCWS", 52, odl_GetBCWS, odl_SetBCWS, VT_R4)
DISP_PROPERTY_EX_ID(Assignment, "fBCWP", 53, odl_GetBCWP, odl_SetBCWP, VT_R4)
DISP_PROPERTY_EX_ID(Assignment, "yBookingType", 54, odl_GetBookingType, odl_SetBookingType, VT_I4)
DISP_PROPERTY_EX_ID(Assignment, "oActualWorkProtected", 55, odl_GetActualWorkProtected, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(Assignment, "oActualOvertimeWorkProtected", 56, odl_GetActualOvertimeWorkProtected, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(Assignment, "dtCreationDate", 57, odl_GetCreationDate, odl_SetCreationDate, VT_DATE)
DISP_PROPERTY_EX_ID(Assignment, "sAssnOwner", 58, odl_GetAssnOwner, odl_SetAssnOwner, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sAssnOwnerGuid", 59, odl_GetAssnOwnerGuid, odl_SetAssnOwnerGuid, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "cBudgetCost", 60, odl_GetBudgetCost, odl_SetBudgetCost, VT_R8)
DISP_PROPERTY_EX_ID(Assignment, "oBudgetWork", 61, odl_GetBudgetWork, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(Assignment, "oExtendedAttribute", 62, odl_GetExtendedAttribute, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(Assignment, "oBaseline", 63, odl_GetBaseline, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(Assignment, "sf404000", 64, odl_Getf404000, odl_Setf404000, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404001", 65, odl_Getf404001, odl_Setf404001, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404002", 66, odl_Getf404002, odl_Setf404002, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404003", 67, odl_Getf404003, odl_Setf404003, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404004", 68, odl_Getf404004, odl_Setf404004, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404005", 69, odl_Getf404005, odl_Setf404005, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404006", 70, odl_Getf404006, odl_Setf404006, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404007", 71, odl_Getf404007, odl_Setf404007, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404008", 72, odl_Getf404008, odl_Setf404008, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404009", 73, odl_Getf404009, odl_Setf404009, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf40400a", 74, odl_Getf40400a, odl_Setf40400a, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf40400b", 75, odl_Getf40400b, odl_Setf40400b, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf40400c", 76, odl_Getf40400c, odl_Setf40400c, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf40400d", 77, odl_Getf40400d, odl_Setf40400d, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf40400e", 78, odl_Getf40400e, odl_Setf40400e, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf40400f", 79, odl_Getf40400f, odl_Setf40400f, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404010", 80, odl_Getf404010, odl_Setf404010, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404011", 81, odl_Getf404011, odl_Setf404011, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404012", 82, odl_Getf404012, odl_Setf404012, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404013", 83, odl_Getf404013, odl_Setf404013, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404014", 84, odl_Getf404014, odl_Setf404014, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404015", 85, odl_Getf404015, odl_Setf404015, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404016", 86, odl_Getf404016, odl_Setf404016, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404017", 87, odl_Getf404017, odl_Setf404017, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404018", 88, odl_Getf404018, odl_Setf404018, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404019", 89, odl_Getf404019, odl_Setf404019, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf40401a", 90, odl_Getf40401a, odl_Setf40401a, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf40401b", 91, odl_Getf40401b, odl_Setf40401b, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf40401c", 92, odl_Getf40401c, odl_Setf40401c, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf40401d", 93, odl_Getf40401d, odl_Setf40401d, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf40401e", 94, odl_Getf40401e, odl_Setf40401e, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf40401f", 95, odl_Getf40401f, odl_Setf40401f, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404020", 96, odl_Getf404020, odl_Setf404020, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404021", 97, odl_Getf404021, odl_Setf404021, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404022", 98, odl_Getf404022, odl_Setf404022, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404023", 99, odl_Getf404023, odl_Setf404023, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404024", 100, odl_Getf404024, odl_Setf404024, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404025", 101, odl_Getf404025, odl_Setf404025, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404026", 102, odl_Getf404026, odl_Setf404026, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404027", 103, odl_Getf404027, odl_Setf404027, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404028", 104, odl_Getf404028, odl_Setf404028, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404029", 105, odl_Getf404029, odl_Setf404029, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf40402a", 106, odl_Getf40402a, odl_Setf40402a, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf40402b", 107, odl_Getf40402b, odl_Setf40402b, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf40402c", 108, odl_Getf40402c, odl_Setf40402c, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf40402d", 109, odl_Getf40402d, odl_Setf40402d, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf40402e", 110, odl_Getf40402e, odl_Setf40402e, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf40402f", 111, odl_Getf40402f, odl_Setf40402f, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404030", 112, odl_Getf404030, odl_Setf404030, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404031", 113, odl_Getf404031, odl_Setf404031, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404032", 114, odl_Getf404032, odl_Setf404032, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404033", 115, odl_Getf404033, odl_Setf404033, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404034", 116, odl_Getf404034, odl_Setf404034, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404035", 117, odl_Getf404035, odl_Setf404035, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404036", 118, odl_Getf404036, odl_Setf404036, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404037", 119, odl_Getf404037, odl_Setf404037, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404038", 120, odl_Getf404038, odl_Setf404038, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404039", 121, odl_Getf404039, odl_Setf404039, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf40403a", 122, odl_Getf40403a, odl_Setf40403a, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf40403b", 123, odl_Getf40403b, odl_Setf40403b, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf40403c", 124, odl_Getf40403c, odl_Setf40403c, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf40403d", 125, odl_Getf40403d, odl_Setf40403d, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf40403e", 126, odl_Getf40403e, odl_Setf40403e, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf40403f", 127, odl_Getf40403f, odl_Setf40403f, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404040", 128, odl_Getf404040, odl_Setf404040, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404041", 129, odl_Getf404041, odl_Setf404041, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404042", 130, odl_Getf404042, odl_Setf404042, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404043", 131, odl_Getf404043, odl_Setf404043, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404044", 132, odl_Getf404044, odl_Setf404044, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404045", 133, odl_Getf404045, odl_Setf404045, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404046", 134, odl_Getf404046, odl_Setf404046, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404047", 135, odl_Getf404047, odl_Setf404047, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404048", 136, odl_Getf404048, odl_Setf404048, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404049", 137, odl_Getf404049, odl_Setf404049, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf40404a", 138, odl_Getf40404a, odl_Setf40404a, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf40404b", 139, odl_Getf40404b, odl_Setf40404b, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf40404c", 140, odl_Getf40404c, odl_Setf40404c, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf40404d", 141, odl_Getf40404d, odl_Setf40404d, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf40404e", 142, odl_Getf40404e, odl_Setf40404e, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf40404f", 143, odl_Getf40404f, odl_Setf40404f, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404050", 144, odl_Getf404050, odl_Setf404050, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404051", 145, odl_Getf404051, odl_Setf404051, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404052", 146, odl_Getf404052, odl_Setf404052, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404053", 147, odl_Getf404053, odl_Setf404053, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404054", 148, odl_Getf404054, odl_Setf404054, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404055", 149, odl_Getf404055, odl_Setf404055, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404056", 150, odl_Getf404056, odl_Setf404056, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404057", 151, odl_Getf404057, odl_Setf404057, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404058", 152, odl_Getf404058, odl_Setf404058, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404059", 153, odl_Getf404059, odl_Setf404059, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf40405a", 154, odl_Getf40405a, odl_Setf40405a, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf40405b", 155, odl_Getf40405b, odl_Setf40405b, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf40405c", 156, odl_Getf40405c, odl_Setf40405c, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf40405d", 157, odl_Getf40405d, odl_Setf40405d, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf40405e", 158, odl_Getf40405e, odl_Setf40405e, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf40405f", 159, odl_Getf40405f, odl_Setf40405f, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404060", 160, odl_Getf404060, odl_Setf404060, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404061", 161, odl_Getf404061, odl_Setf404061, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404062", 162, odl_Getf404062, odl_Setf404062, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404063", 163, odl_Getf404063, odl_Setf404063, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404064", 164, odl_Getf404064, odl_Setf404064, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404065", 165, odl_Getf404065, odl_Setf404065, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404066", 166, odl_Getf404066, odl_Setf404066, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404067", 167, odl_Getf404067, odl_Setf404067, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404068", 168, odl_Getf404068, odl_Setf404068, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404069", 169, odl_Getf404069, odl_Setf404069, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf40406a", 170, odl_Getf40406a, odl_Setf40406a, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf40406b", 171, odl_Getf40406b, odl_Setf40406b, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf40406c", 172, odl_Getf40406c, odl_Setf40406c, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf40406d", 173, odl_Getf40406d, odl_Setf40406d, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf40406e", 174, odl_Getf40406e, odl_Setf40406e, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf40406f", 175, odl_Getf40406f, odl_Setf40406f, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404070", 176, odl_Getf404070, odl_Setf404070, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404071", 177, odl_Getf404071, odl_Setf404071, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404072", 178, odl_Getf404072, odl_Setf404072, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404073", 179, odl_Getf404073, odl_Setf404073, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404074", 180, odl_Getf404074, odl_Setf404074, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404075", 181, odl_Getf404075, odl_Setf404075, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404076", 182, odl_Getf404076, odl_Setf404076, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404077", 183, odl_Getf404077, odl_Setf404077, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404078", 184, odl_Getf404078, odl_Setf404078, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404079", 185, odl_Getf404079, odl_Setf404079, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf40407a", 186, odl_Getf40407a, odl_Setf40407a, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf40407b", 187, odl_Getf40407b, odl_Setf40407b, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf40407c", 188, odl_Getf40407c, odl_Setf40407c, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf40407d", 189, odl_Getf40407d, odl_Setf40407d, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf40407e", 190, odl_Getf40407e, odl_Setf40407e, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf40407f", 191, odl_Getf40407f, odl_Setf40407f, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404080", 192, odl_Getf404080, odl_Setf404080, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404081", 193, odl_Getf404081, odl_Setf404081, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404082", 194, odl_Getf404082, odl_Setf404082, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404083", 195, odl_Getf404083, odl_Setf404083, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404084", 196, odl_Getf404084, odl_Setf404084, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404085", 197, odl_Getf404085, odl_Setf404085, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404086", 198, odl_Getf404086, odl_Setf404086, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404087", 199, odl_Getf404087, odl_Setf404087, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404088", 200, odl_Getf404088, odl_Setf404088, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404089", 201, odl_Getf404089, odl_Setf404089, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf40408a", 202, odl_Getf40408a, odl_Setf40408a, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf40408b", 203, odl_Getf40408b, odl_Setf40408b, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf40408c", 204, odl_Getf40408c, odl_Setf40408c, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf40408d", 205, odl_Getf40408d, odl_Setf40408d, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf40408e", 206, odl_Getf40408e, odl_Setf40408e, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf40408f", 207, odl_Getf40408f, odl_Setf40408f, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404090", 208, odl_Getf404090, odl_Setf404090, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404091", 209, odl_Getf404091, odl_Setf404091, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404092", 210, odl_Getf404092, odl_Setf404092, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404093", 211, odl_Getf404093, odl_Setf404093, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404094", 212, odl_Getf404094, odl_Setf404094, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404095", 213, odl_Getf404095, odl_Setf404095, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404096", 214, odl_Getf404096, odl_Setf404096, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404097", 215, odl_Getf404097, odl_Setf404097, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404098", 216, odl_Getf404098, odl_Setf404098, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf404099", 217, odl_Getf404099, odl_Setf404099, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf40409a", 218, odl_Getf40409a, odl_Setf40409a, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf40409b", 219, odl_Getf40409b, odl_Setf40409b, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf40409c", 220, odl_Getf40409c, odl_Setf40409c, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf40409d", 221, odl_Getf40409d, odl_Setf40409d, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf40409e", 222, odl_Getf40409e, odl_Setf40409e, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf40409f", 223, odl_Getf40409f, odl_Setf40409f, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf4040a0", 224, odl_Getf4040a0, odl_Setf4040a0, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf4040a1", 225, odl_Getf4040a1, odl_Setf4040a1, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf4040a2", 226, odl_Getf4040a2, odl_Setf4040a2, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf4040a3", 227, odl_Getf4040a3, odl_Setf4040a3, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf4040a4", 228, odl_Getf4040a4, odl_Setf4040a4, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf4040a5", 229, odl_Getf4040a5, odl_Setf4040a5, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf4040a6", 230, odl_Getf4040a6, odl_Setf4040a6, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf4040a7", 231, odl_Getf4040a7, odl_Setf4040a7, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf4040a8", 232, odl_Getf4040a8, odl_Setf4040a8, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf4040a9", 233, odl_Getf4040a9, odl_Setf4040a9, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf4040aa", 234, odl_Getf4040aa, odl_Setf4040aa, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf4040ab", 235, odl_Getf4040ab, odl_Setf4040ab, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf4040ac", 236, odl_Getf4040ac, odl_Setf4040ac, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf4040ad", 237, odl_Getf4040ad, odl_Setf4040ad, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf4040ae", 238, odl_Getf4040ae, odl_Setf4040ae, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf4040af", 239, odl_Getf4040af, odl_Setf4040af, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf4040b0", 240, odl_Getf4040b0, odl_Setf4040b0, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf4040b1", 241, odl_Getf4040b1, odl_Setf4040b1, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf4040b2", 242, odl_Getf4040b2, odl_Setf4040b2, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf4040b3", 243, odl_Getf4040b3, odl_Setf4040b3, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf4040b4", 244, odl_Getf4040b4, odl_Setf4040b4, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf4040b5", 245, odl_Getf4040b5, odl_Setf4040b5, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf4040b6", 246, odl_Getf4040b6, odl_Setf4040b6, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf4040b7", 247, odl_Getf4040b7, odl_Setf4040b7, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf4040b8", 248, odl_Getf4040b8, odl_Setf4040b8, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf4040b9", 249, odl_Getf4040b9, odl_Setf4040b9, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf4040ba", 250, odl_Getf4040ba, odl_Setf4040ba, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf4040bb", 251, odl_Getf4040bb, odl_Setf4040bb, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf4040bc", 252, odl_Getf4040bc, odl_Setf4040bc, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf4040bd", 253, odl_Getf4040bd, odl_Setf4040bd, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf4040be", 254, odl_Getf4040be, odl_Setf4040be, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf4040bf", 255, odl_Getf4040bf, odl_Setf4040bf, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf4040c0", 256, odl_Getf4040c0, odl_Setf4040c0, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf4040c1", 257, odl_Getf4040c1, odl_Setf4040c1, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf4040c2", 258, odl_Getf4040c2, odl_Setf4040c2, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf4040c3", 259, odl_Getf4040c3, odl_Setf4040c3, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf4040c4", 260, odl_Getf4040c4, odl_Setf4040c4, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf4040c5", 261, odl_Getf4040c5, odl_Setf4040c5, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf4040c6", 262, odl_Getf4040c6, odl_Setf4040c6, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf4040c7", 263, odl_Getf4040c7, odl_Setf4040c7, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "sf4040c8", 264, odl_Getf4040c8, odl_Setf4040c8, VT_BSTR)
DISP_PROPERTY_EX_ID(Assignment, "oTimephasedData", 265, odl_GetTimephasedData, SetNotSupported, VT_DISPATCH)
DISP_PROPERTY_EX_ID(Assignment, "Key", 266, odl_GetKey, odl_SetKey, VT_BSTR)
DISP_FUNCTION_ID(Assignment, "GetXML", 267, odl_GetXML, VT_BSTR, VTS_NONE)
DISP_FUNCTION_ID(Assignment, "SetXML", 268, odl_SetXML, VT_EMPTY, VTS_BSTR)
DISP_FUNCTION_ID(Assignment, "IsNull", 269, IsNull, VT_BOOL, VTS_NONE)
DISP_FUNCTION_ID(Assignment, "Initialize", 270, Initialize, VT_EMPTY, VTS_NONE)
END_DISPATCH_MAP()

BEGIN_INTERFACE_MAP(Assignment, CCmdTarget)
	INTERFACE_PART(Assignment, IID_IAssignment, Dispatch)
END_INTERFACE_MAP()

BEGIN_MESSAGE_MAP(Assignment, CCmdTarget)
END_MESSAGE_MAP()


void Assignment::Initialize(void)
{
	InitVars();
}

Assignment::Assignment()
{
	EnableAutomation();
	AfxOleLockApp();
	InitVars();
}

void Assignment::InitVars(void)
{
	mp_lUID = 0;
	mp_lTaskUID = 0;
	mp_lResourceUID = 0;
	mp_lPercentWorkComplete = 0;
	mp_cActualCost = 0;
	mp_dtActualFinish = (DATE) 0;
	mp_cActualOvertimeCost = 0;
	mp_oActualOvertimeWork = new Duration();
	mp_dtActualStart = (DATE) 0;
	mp_oActualWork = new Duration();
	mp_fACWP = 0;
	mp_bConfirmed = FALSE;
	mp_cCost = 0;
	mp_yCostRateTable = CRT_COST_RATE_TABLE_0;
	mp_fCostVariance = 0;
	mp_fCV = 0;
	mp_lDelay = 0;
	mp_dtFinish = (DATE) 0;
	mp_lFinishVariance = 0;
	mp_sHyperlink = "";
	mp_sHyperlinkAddress = "";
	mp_sHyperlinkSubAddress = "";
	mp_fWorkVariance = 0;
	mp_bHasFixedRateUnits = FALSE;
	mp_bFixedMaterial = FALSE;
	mp_lLevelingDelay = 0;
	mp_yLevelingDelayFormat = LDF_M;
	mp_bLinkedFields = FALSE;
	mp_bMilestone = FALSE;
	mp_sNotes = "";
	mp_bOverallocated = FALSE;
	mp_cOvertimeCost = 0;
	mp_oOvertimeWork = new Duration();
	mp_fPeakUnits = 0;
	mp_oRegularWork = new Duration();
	mp_cRemainingCost = 0;
	mp_cRemainingOvertimeCost = 0;
	mp_oRemainingOvertimeWork = new Duration();
	mp_oRemainingWork = new Duration();
	mp_bResponsePending = FALSE;
	mp_dtStart = (DATE) 0;
	mp_dtStop = (DATE) 0;
	mp_dtResume = (DATE) 0;
	mp_lStartVariance = 0;
	mp_bSummary = FALSE;
	mp_fSV = 0;
	mp_fUnits = 0;
	mp_bUpdateNeeded = FALSE;
	mp_fVAC = 0;
	mp_oWork = new Duration();
	mp_yWorkContour = WC_FLAT;
	mp_fBCWS = 0;
	mp_fBCWP = 0;
	mp_yBookingType = BT_COMMITED;
	mp_oActualWorkProtected = new Duration();
	mp_oActualOvertimeWorkProtected = new Duration();
	mp_dtCreationDate = (DATE) 0;
	mp_sAssnOwner = "";
	mp_sAssnOwnerGuid = "";
	mp_cBudgetCost = 0;
	mp_oBudgetWork = new Duration();
	mp_oExtendedAttribute_C = new AssignmentExtendedAttribute_C();
	mp_oBaseline_C = new AssignmentBaseline_C();
	mp_sf404000 = "";
	mp_sf404001 = "";
	mp_sf404002 = "";
	mp_sf404003 = "";
	mp_sf404004 = "";
	mp_sf404005 = "";
	mp_sf404006 = "";
	mp_sf404007 = "";
	mp_sf404008 = "";
	mp_sf404009 = "";
	mp_sf40400a = "";
	mp_sf40400b = "";
	mp_sf40400c = "";
	mp_sf40400d = "";
	mp_sf40400e = "";
	mp_sf40400f = "";
	mp_sf404010 = "";
	mp_sf404011 = "";
	mp_sf404012 = "";
	mp_sf404013 = "";
	mp_sf404014 = "";
	mp_sf404015 = "";
	mp_sf404016 = "";
	mp_sf404017 = "";
	mp_sf404018 = "";
	mp_sf404019 = "";
	mp_sf40401a = "";
	mp_sf40401b = "";
	mp_sf40401c = "";
	mp_sf40401d = "";
	mp_sf40401e = "";
	mp_sf40401f = "";
	mp_sf404020 = "";
	mp_sf404021 = "";
	mp_sf404022 = "";
	mp_sf404023 = "";
	mp_sf404024 = "";
	mp_sf404025 = "";
	mp_sf404026 = "";
	mp_sf404027 = "";
	mp_sf404028 = "";
	mp_sf404029 = "";
	mp_sf40402a = "";
	mp_sf40402b = "";
	mp_sf40402c = "";
	mp_sf40402d = "";
	mp_sf40402e = "";
	mp_sf40402f = "";
	mp_sf404030 = "";
	mp_sf404031 = "";
	mp_sf404032 = "";
	mp_sf404033 = "";
	mp_sf404034 = "";
	mp_sf404035 = "";
	mp_sf404036 = "";
	mp_sf404037 = "";
	mp_sf404038 = "";
	mp_sf404039 = "";
	mp_sf40403a = "";
	mp_sf40403b = "";
	mp_sf40403c = "";
	mp_sf40403d = "";
	mp_sf40403e = "";
	mp_sf40403f = "";
	mp_sf404040 = "";
	mp_sf404041 = "";
	mp_sf404042 = "";
	mp_sf404043 = "";
	mp_sf404044 = "";
	mp_sf404045 = "";
	mp_sf404046 = "";
	mp_sf404047 = "";
	mp_sf404048 = "";
	mp_sf404049 = "";
	mp_sf40404a = "";
	mp_sf40404b = "";
	mp_sf40404c = "";
	mp_sf40404d = "";
	mp_sf40404e = "";
	mp_sf40404f = "";
	mp_sf404050 = "";
	mp_sf404051 = "";
	mp_sf404052 = "";
	mp_sf404053 = "";
	mp_sf404054 = "";
	mp_sf404055 = "";
	mp_sf404056 = "";
	mp_sf404057 = "";
	mp_sf404058 = "";
	mp_sf404059 = "";
	mp_sf40405a = "";
	mp_sf40405b = "";
	mp_sf40405c = "";
	mp_sf40405d = "";
	mp_sf40405e = "";
	mp_sf40405f = "";
	mp_sf404060 = "";
	mp_sf404061 = "";
	mp_sf404062 = "";
	mp_sf404063 = "";
	mp_sf404064 = "";
	mp_sf404065 = "";
	mp_sf404066 = "";
	mp_sf404067 = "";
	mp_sf404068 = "";
	mp_sf404069 = "";
	mp_sf40406a = "";
	mp_sf40406b = "";
	mp_sf40406c = "";
	mp_sf40406d = "";
	mp_sf40406e = "";
	mp_sf40406f = "";
	mp_sf404070 = "";
	mp_sf404071 = "";
	mp_sf404072 = "";
	mp_sf404073 = "";
	mp_sf404074 = "";
	mp_sf404075 = "";
	mp_sf404076 = "";
	mp_sf404077 = "";
	mp_sf404078 = "";
	mp_sf404079 = "";
	mp_sf40407a = "";
	mp_sf40407b = "";
	mp_sf40407c = "";
	mp_sf40407d = "";
	mp_sf40407e = "";
	mp_sf40407f = "";
	mp_sf404080 = "";
	mp_sf404081 = "";
	mp_sf404082 = "";
	mp_sf404083 = "";
	mp_sf404084 = "";
	mp_sf404085 = "";
	mp_sf404086 = "";
	mp_sf404087 = "";
	mp_sf404088 = "";
	mp_sf404089 = "";
	mp_sf40408a = "";
	mp_sf40408b = "";
	mp_sf40408c = "";
	mp_sf40408d = "";
	mp_sf40408e = "";
	mp_sf40408f = "";
	mp_sf404090 = "";
	mp_sf404091 = "";
	mp_sf404092 = "";
	mp_sf404093 = "";
	mp_sf404094 = "";
	mp_sf404095 = "";
	mp_sf404096 = "";
	mp_sf404097 = "";
	mp_sf404098 = "";
	mp_sf404099 = "";
	mp_sf40409a = "";
	mp_sf40409b = "";
	mp_sf40409c = "";
	mp_sf40409d = "";
	mp_sf40409e = "";
	mp_sf40409f = "";
	mp_sf4040a0 = "";
	mp_sf4040a1 = "";
	mp_sf4040a2 = "";
	mp_sf4040a3 = "";
	mp_sf4040a4 = "";
	mp_sf4040a5 = "";
	mp_sf4040a6 = "";
	mp_sf4040a7 = "";
	mp_sf4040a8 = "";
	mp_sf4040a9 = "";
	mp_sf4040aa = "";
	mp_sf4040ab = "";
	mp_sf4040ac = "";
	mp_sf4040ad = "";
	mp_sf4040ae = "";
	mp_sf4040af = "";
	mp_sf4040b0 = "";
	mp_sf4040b1 = "";
	mp_sf4040b2 = "";
	mp_sf4040b3 = "";
	mp_sf4040b4 = "";
	mp_sf4040b5 = "";
	mp_sf4040b6 = "";
	mp_sf4040b7 = "";
	mp_sf4040b8 = "";
	mp_sf4040b9 = "";
	mp_sf4040ba = "";
	mp_sf4040bb = "";
	mp_sf4040bc = "";
	mp_sf4040bd = "";
	mp_sf4040be = "";
	mp_sf4040bf = "";
	mp_sf4040c0 = "";
	mp_sf4040c1 = "";
	mp_sf4040c2 = "";
	mp_sf4040c3 = "";
	mp_sf4040c4 = "";
	mp_sf4040c5 = "";
	mp_sf4040c6 = "";
	mp_sf4040c7 = "";
	mp_sf4040c8 = "";
	mp_oTimephasedData_C = new TimephasedData_C();
}

Assignment::~Assignment()
{
	delete mp_oActualOvertimeWork;
	delete mp_oActualWork;
	delete mp_oOvertimeWork;
	delete mp_oRegularWork;
	delete mp_oRemainingOvertimeWork;
	delete mp_oRemainingWork;
	delete mp_oWork;
	delete mp_oActualWorkProtected;
	delete mp_oActualOvertimeWorkProtected;
	delete mp_oBudgetWork;
	delete mp_oExtendedAttribute_C;
	delete mp_oBaseline_C;
	delete mp_oTimephasedData_C;
	AfxOleUnlockApp();
}

void Assignment::OnFinalRelease()
{
	clsItemBase::OnFinalRelease();
}

LONG Assignment::odl_GetUID(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetUID();
}

LONG Assignment::GetUID(void)
{
	return (LONG) mp_lUID;
}

void Assignment::odl_SetUID(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetUID((LONG)newVal);
}

void Assignment::SetUID(LONG newVal)
{
	mp_lUID = newVal;
}

LONG Assignment::odl_GetTaskUID(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetTaskUID();
}

LONG Assignment::GetTaskUID(void)
{
	return (LONG) mp_lTaskUID;
}

void Assignment::odl_SetTaskUID(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetTaskUID((LONG)newVal);
}

void Assignment::SetTaskUID(LONG newVal)
{
	mp_lTaskUID = newVal;
}

LONG Assignment::odl_GetResourceUID(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetResourceUID();
}

LONG Assignment::GetResourceUID(void)
{
	return (LONG) mp_lResourceUID;
}

void Assignment::odl_SetResourceUID(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetResourceUID((LONG)newVal);
}

void Assignment::SetResourceUID(LONG newVal)
{
	mp_lResourceUID = newVal;
}

LONG Assignment::odl_GetPercentWorkComplete(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetPercentWorkComplete();
}

LONG Assignment::GetPercentWorkComplete(void)
{
	return (LONG) mp_lPercentWorkComplete;
}

void Assignment::odl_SetPercentWorkComplete(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetPercentWorkComplete((LONG)newVal);
}

void Assignment::SetPercentWorkComplete(LONG newVal)
{
	mp_lPercentWorkComplete = newVal;
}

DOUBLE Assignment::odl_GetActualCost(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetActualCost();
}

DOUBLE Assignment::GetActualCost(void)
{
	return (DOUBLE) mp_cActualCost;
}

void Assignment::odl_SetActualCost(DOUBLE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetActualCost((DOUBLE)newVal);
}

void Assignment::SetActualCost(DOUBLE newVal)
{
	mp_cActualCost = newVal;
}

DATE Assignment::odl_GetActualFinish(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetActualFinish();
}

COleDateTime Assignment::GetActualFinish(void)
{
	return (COleDateTime) mp_dtActualFinish;
}

void Assignment::odl_SetActualFinish(DATE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetActualFinish((COleDateTime)newVal);
}

void Assignment::SetActualFinish(COleDateTime newVal)
{
	mp_dtActualFinish = newVal;
}

DOUBLE Assignment::odl_GetActualOvertimeCost(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetActualOvertimeCost();
}

DOUBLE Assignment::GetActualOvertimeCost(void)
{
	return (DOUBLE) mp_cActualOvertimeCost;
}

void Assignment::odl_SetActualOvertimeCost(DOUBLE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetActualOvertimeCost((DOUBLE)newVal);
}

void Assignment::SetActualOvertimeCost(DOUBLE newVal)
{
	mp_cActualOvertimeCost = newVal;
}

IDispatch* Assignment::odl_GetActualOvertimeWork(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetActualOvertimeWork()->GetIDispatch(TRUE);
}

Duration* Assignment::GetActualOvertimeWork(void)
{
	return mp_oActualOvertimeWork;
}

DATE Assignment::odl_GetActualStart(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetActualStart();
}

COleDateTime Assignment::GetActualStart(void)
{
	return (COleDateTime) mp_dtActualStart;
}

void Assignment::odl_SetActualStart(DATE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetActualStart((COleDateTime)newVal);
}

void Assignment::SetActualStart(COleDateTime newVal)
{
	mp_dtActualStart = newVal;
}

IDispatch* Assignment::odl_GetActualWork(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetActualWork()->GetIDispatch(TRUE);
}

Duration* Assignment::GetActualWork(void)
{
	return mp_oActualWork;
}

FLOAT Assignment::odl_GetACWP(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetACWP();
}

FLOAT Assignment::GetACWP(void)
{
	return (FLOAT) mp_fACWP;
}

void Assignment::odl_SetACWP(FLOAT newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetACWP((FLOAT)newVal);
}

void Assignment::SetACWP(FLOAT newVal)
{
	mp_fACWP = newVal;
}

BOOL Assignment::odl_GetConfirmed(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetConfirmed();
}

BOOL Assignment::GetConfirmed(void)
{
	return (BOOL) mp_bConfirmed;
}

void Assignment::odl_SetConfirmed(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetConfirmed((BOOL)newVal);
}

void Assignment::SetConfirmed(BOOL newVal)
{
	mp_bConfirmed = newVal;
}

DOUBLE Assignment::odl_GetCost(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCost();
}

DOUBLE Assignment::GetCost(void)
{
	return (DOUBLE) mp_cCost;
}

void Assignment::odl_SetCost(DOUBLE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetCost((DOUBLE)newVal);
}

void Assignment::SetCost(DOUBLE newVal)
{
	mp_cCost = newVal;
}

LONG Assignment::odl_GetCostRateTable(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCostRateTable();
}

E_COSTRATETABLE Assignment::GetCostRateTable(void)
{
	return (E_COSTRATETABLE) mp_yCostRateTable;
}

void Assignment::odl_SetCostRateTable(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetCostRateTable((E_COSTRATETABLE)newVal);
}

void Assignment::SetCostRateTable(E_COSTRATETABLE newVal)
{
	mp_yCostRateTable = newVal;
}

FLOAT Assignment::odl_GetCostVariance(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCostVariance();
}

FLOAT Assignment::GetCostVariance(void)
{
	return (FLOAT) mp_fCostVariance;
}

void Assignment::odl_SetCostVariance(FLOAT newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetCostVariance((FLOAT)newVal);
}

void Assignment::SetCostVariance(FLOAT newVal)
{
	mp_fCostVariance = newVal;
}

FLOAT Assignment::odl_GetCV(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCV();
}

FLOAT Assignment::GetCV(void)
{
	return (FLOAT) mp_fCV;
}

void Assignment::odl_SetCV(FLOAT newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetCV((FLOAT)newVal);
}

void Assignment::SetCV(FLOAT newVal)
{
	mp_fCV = newVal;
}

LONG Assignment::odl_GetDelay(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetDelay();
}

LONG Assignment::GetDelay(void)
{
	return (LONG) mp_lDelay;
}

void Assignment::odl_SetDelay(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetDelay((LONG)newVal);
}

void Assignment::SetDelay(LONG newVal)
{
	mp_lDelay = newVal;
}

DATE Assignment::odl_GetFinish(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetFinish();
}

COleDateTime Assignment::GetFinish(void)
{
	return (COleDateTime) mp_dtFinish;
}

void Assignment::odl_SetFinish(DATE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetFinish((COleDateTime)newVal);
}

void Assignment::SetFinish(COleDateTime newVal)
{
	mp_dtFinish = newVal;
}

LONG Assignment::odl_GetFinishVariance(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetFinishVariance();
}

LONG Assignment::GetFinishVariance(void)
{
	return (LONG) mp_lFinishVariance;
}

void Assignment::odl_SetFinishVariance(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetFinishVariance((LONG)newVal);
}

void Assignment::SetFinishVariance(LONG newVal)
{
	mp_lFinishVariance = newVal;
}

BSTR Assignment::odl_GetHyperlink(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetHyperlink().AllocSysString();
}

CString Assignment::GetHyperlink(void)
{
	return (CString) mp_sHyperlink;
}

void Assignment::odl_SetHyperlink(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetHyperlink(newVal);
}

void Assignment::SetHyperlink(CString newVal)
{
	if (newVal.GetLength() > 512)
	{
		newVal = newVal.Mid(0, 512);
	}
	mp_sHyperlink = newVal;
}

BSTR Assignment::odl_GetHyperlinkAddress(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetHyperlinkAddress().AllocSysString();
}

CString Assignment::GetHyperlinkAddress(void)
{
	return (CString) mp_sHyperlinkAddress;
}

void Assignment::odl_SetHyperlinkAddress(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetHyperlinkAddress(newVal);
}

void Assignment::SetHyperlinkAddress(CString newVal)
{
	if (newVal.GetLength() > 512)
	{
		newVal = newVal.Mid(0, 512);
	}
	mp_sHyperlinkAddress = newVal;
}

BSTR Assignment::odl_GetHyperlinkSubAddress(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetHyperlinkSubAddress().AllocSysString();
}

CString Assignment::GetHyperlinkSubAddress(void)
{
	return (CString) mp_sHyperlinkSubAddress;
}

void Assignment::odl_SetHyperlinkSubAddress(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetHyperlinkSubAddress(newVal);
}

void Assignment::SetHyperlinkSubAddress(CString newVal)
{
	if (newVal.GetLength() > 512)
	{
		newVal = newVal.Mid(0, 512);
	}
	mp_sHyperlinkSubAddress = newVal;
}

FLOAT Assignment::odl_GetWorkVariance(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetWorkVariance();
}

FLOAT Assignment::GetWorkVariance(void)
{
	return (FLOAT) mp_fWorkVariance;
}

void Assignment::odl_SetWorkVariance(FLOAT newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetWorkVariance((FLOAT)newVal);
}

void Assignment::SetWorkVariance(FLOAT newVal)
{
	mp_fWorkVariance = newVal;
}

BOOL Assignment::odl_GetHasFixedRateUnits(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetHasFixedRateUnits();
}

BOOL Assignment::GetHasFixedRateUnits(void)
{
	return (BOOL) mp_bHasFixedRateUnits;
}

void Assignment::odl_SetHasFixedRateUnits(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetHasFixedRateUnits((BOOL)newVal);
}

void Assignment::SetHasFixedRateUnits(BOOL newVal)
{
	mp_bHasFixedRateUnits = newVal;
}

BOOL Assignment::odl_GetFixedMaterial(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetFixedMaterial();
}

BOOL Assignment::GetFixedMaterial(void)
{
	return (BOOL) mp_bFixedMaterial;
}

void Assignment::odl_SetFixedMaterial(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetFixedMaterial((BOOL)newVal);
}

void Assignment::SetFixedMaterial(BOOL newVal)
{
	mp_bFixedMaterial = newVal;
}

LONG Assignment::odl_GetLevelingDelay(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetLevelingDelay();
}

LONG Assignment::GetLevelingDelay(void)
{
	return (LONG) mp_lLevelingDelay;
}

void Assignment::odl_SetLevelingDelay(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetLevelingDelay((LONG)newVal);
}

void Assignment::SetLevelingDelay(LONG newVal)
{
	mp_lLevelingDelay = newVal;
}

LONG Assignment::odl_GetLevelingDelayFormat(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetLevelingDelayFormat();
}

E_LEVELINGDELAYFORMAT Assignment::GetLevelingDelayFormat(void)
{
	return (E_LEVELINGDELAYFORMAT) mp_yLevelingDelayFormat;
}

void Assignment::odl_SetLevelingDelayFormat(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetLevelingDelayFormat((E_LEVELINGDELAYFORMAT)newVal);
}

void Assignment::SetLevelingDelayFormat(E_LEVELINGDELAYFORMAT newVal)
{
	mp_yLevelingDelayFormat = newVal;
}

BOOL Assignment::odl_GetLinkedFields(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetLinkedFields();
}

BOOL Assignment::GetLinkedFields(void)
{
	return (BOOL) mp_bLinkedFields;
}

void Assignment::odl_SetLinkedFields(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetLinkedFields((BOOL)newVal);
}

void Assignment::SetLinkedFields(BOOL newVal)
{
	mp_bLinkedFields = newVal;
}

BOOL Assignment::odl_GetMilestone(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetMilestone();
}

BOOL Assignment::GetMilestone(void)
{
	return (BOOL) mp_bMilestone;
}

void Assignment::odl_SetMilestone(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetMilestone((BOOL)newVal);
}

void Assignment::SetMilestone(BOOL newVal)
{
	mp_bMilestone = newVal;
}

BSTR Assignment::odl_GetNotes(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetNotes().AllocSysString();
}

CString Assignment::GetNotes(void)
{
	return (CString) mp_sNotes;
}

void Assignment::odl_SetNotes(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetNotes(newVal);
}

void Assignment::SetNotes(CString newVal)
{
	mp_sNotes = newVal;
}

BOOL Assignment::odl_GetOverallocated(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetOverallocated();
}

BOOL Assignment::GetOverallocated(void)
{
	return (BOOL) mp_bOverallocated;
}

void Assignment::odl_SetOverallocated(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetOverallocated((BOOL)newVal);
}

void Assignment::SetOverallocated(BOOL newVal)
{
	mp_bOverallocated = newVal;
}

DOUBLE Assignment::odl_GetOvertimeCost(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetOvertimeCost();
}

DOUBLE Assignment::GetOvertimeCost(void)
{
	return (DOUBLE) mp_cOvertimeCost;
}

void Assignment::odl_SetOvertimeCost(DOUBLE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetOvertimeCost((DOUBLE)newVal);
}

void Assignment::SetOvertimeCost(DOUBLE newVal)
{
	mp_cOvertimeCost = newVal;
}

IDispatch* Assignment::odl_GetOvertimeWork(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetOvertimeWork()->GetIDispatch(TRUE);
}

Duration* Assignment::GetOvertimeWork(void)
{
	return mp_oOvertimeWork;
}

FLOAT Assignment::odl_GetPeakUnits(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetPeakUnits();
}

FLOAT Assignment::GetPeakUnits(void)
{
	return (FLOAT) mp_fPeakUnits;
}

void Assignment::odl_SetPeakUnits(FLOAT newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetPeakUnits((FLOAT)newVal);
}

void Assignment::SetPeakUnits(FLOAT newVal)
{
	mp_fPeakUnits = newVal;
}

IDispatch* Assignment::odl_GetRegularWork(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetRegularWork()->GetIDispatch(TRUE);
}

Duration* Assignment::GetRegularWork(void)
{
	return mp_oRegularWork;
}

DOUBLE Assignment::odl_GetRemainingCost(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetRemainingCost();
}

DOUBLE Assignment::GetRemainingCost(void)
{
	return (DOUBLE) mp_cRemainingCost;
}

void Assignment::odl_SetRemainingCost(DOUBLE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetRemainingCost((DOUBLE)newVal);
}

void Assignment::SetRemainingCost(DOUBLE newVal)
{
	mp_cRemainingCost = newVal;
}

DOUBLE Assignment::odl_GetRemainingOvertimeCost(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetRemainingOvertimeCost();
}

DOUBLE Assignment::GetRemainingOvertimeCost(void)
{
	return (DOUBLE) mp_cRemainingOvertimeCost;
}

void Assignment::odl_SetRemainingOvertimeCost(DOUBLE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetRemainingOvertimeCost((DOUBLE)newVal);
}

void Assignment::SetRemainingOvertimeCost(DOUBLE newVal)
{
	mp_cRemainingOvertimeCost = newVal;
}

IDispatch* Assignment::odl_GetRemainingOvertimeWork(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetRemainingOvertimeWork()->GetIDispatch(TRUE);
}

Duration* Assignment::GetRemainingOvertimeWork(void)
{
	return mp_oRemainingOvertimeWork;
}

IDispatch* Assignment::odl_GetRemainingWork(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetRemainingWork()->GetIDispatch(TRUE);
}

Duration* Assignment::GetRemainingWork(void)
{
	return mp_oRemainingWork;
}

BOOL Assignment::odl_GetResponsePending(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetResponsePending();
}

BOOL Assignment::GetResponsePending(void)
{
	return (BOOL) mp_bResponsePending;
}

void Assignment::odl_SetResponsePending(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetResponsePending((BOOL)newVal);
}

void Assignment::SetResponsePending(BOOL newVal)
{
	mp_bResponsePending = newVal;
}

DATE Assignment::odl_GetStart(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetStart();
}

COleDateTime Assignment::GetStart(void)
{
	return (COleDateTime) mp_dtStart;
}

void Assignment::odl_SetStart(DATE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetStart((COleDateTime)newVal);
}

void Assignment::SetStart(COleDateTime newVal)
{
	mp_dtStart = newVal;
}

DATE Assignment::odl_GetStop(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetStop();
}

COleDateTime Assignment::GetStop(void)
{
	return (COleDateTime) mp_dtStop;
}

void Assignment::odl_SetStop(DATE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetStop((COleDateTime)newVal);
}

void Assignment::SetStop(COleDateTime newVal)
{
	mp_dtStop = newVal;
}

DATE Assignment::odl_GetResume(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetResume();
}

COleDateTime Assignment::GetResume(void)
{
	return (COleDateTime) mp_dtResume;
}

void Assignment::odl_SetResume(DATE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetResume((COleDateTime)newVal);
}

void Assignment::SetResume(COleDateTime newVal)
{
	mp_dtResume = newVal;
}

LONG Assignment::odl_GetStartVariance(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetStartVariance();
}

LONG Assignment::GetStartVariance(void)
{
	return (LONG) mp_lStartVariance;
}

void Assignment::odl_SetStartVariance(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetStartVariance((LONG)newVal);
}

void Assignment::SetStartVariance(LONG newVal)
{
	mp_lStartVariance = newVal;
}

BOOL Assignment::odl_GetSummary(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetSummary();
}

BOOL Assignment::GetSummary(void)
{
	return (BOOL) mp_bSummary;
}

void Assignment::odl_SetSummary(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetSummary((BOOL)newVal);
}

void Assignment::SetSummary(BOOL newVal)
{
	mp_bSummary = newVal;
}

FLOAT Assignment::odl_GetSV(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetSV();
}

FLOAT Assignment::GetSV(void)
{
	return (FLOAT) mp_fSV;
}

void Assignment::odl_SetSV(FLOAT newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetSV((FLOAT)newVal);
}

void Assignment::SetSV(FLOAT newVal)
{
	mp_fSV = newVal;
}

FLOAT Assignment::odl_GetUnits(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetUnits();
}

FLOAT Assignment::GetUnits(void)
{
	return (FLOAT) mp_fUnits;
}

void Assignment::odl_SetUnits(FLOAT newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetUnits((FLOAT)newVal);
}

void Assignment::SetUnits(FLOAT newVal)
{
	mp_fUnits = newVal;
}

BOOL Assignment::odl_GetUpdateNeeded(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetUpdateNeeded();
}

BOOL Assignment::GetUpdateNeeded(void)
{
	return (BOOL) mp_bUpdateNeeded;
}

void Assignment::odl_SetUpdateNeeded(BOOL newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetUpdateNeeded((BOOL)newVal);
}

void Assignment::SetUpdateNeeded(BOOL newVal)
{
	mp_bUpdateNeeded = newVal;
}

FLOAT Assignment::odl_GetVAC(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetVAC();
}

FLOAT Assignment::GetVAC(void)
{
	return (FLOAT) mp_fVAC;
}

void Assignment::odl_SetVAC(FLOAT newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetVAC((FLOAT)newVal);
}

void Assignment::SetVAC(FLOAT newVal)
{
	mp_fVAC = newVal;
}

IDispatch* Assignment::odl_GetWork(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetWork()->GetIDispatch(TRUE);
}

Duration* Assignment::GetWork(void)
{
	return mp_oWork;
}

LONG Assignment::odl_GetWorkContour(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetWorkContour();
}

E_WORKCONTOUR Assignment::GetWorkContour(void)
{
	return (E_WORKCONTOUR) mp_yWorkContour;
}

void Assignment::odl_SetWorkContour(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetWorkContour((E_WORKCONTOUR)newVal);
}

void Assignment::SetWorkContour(E_WORKCONTOUR newVal)
{
	mp_yWorkContour = newVal;
}

FLOAT Assignment::odl_GetBCWS(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetBCWS();
}

FLOAT Assignment::GetBCWS(void)
{
	return (FLOAT) mp_fBCWS;
}

void Assignment::odl_SetBCWS(FLOAT newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetBCWS((FLOAT)newVal);
}

void Assignment::SetBCWS(FLOAT newVal)
{
	mp_fBCWS = newVal;
}

FLOAT Assignment::odl_GetBCWP(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetBCWP();
}

FLOAT Assignment::GetBCWP(void)
{
	return (FLOAT) mp_fBCWP;
}

void Assignment::odl_SetBCWP(FLOAT newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetBCWP((FLOAT)newVal);
}

void Assignment::SetBCWP(FLOAT newVal)
{
	mp_fBCWP = newVal;
}

LONG Assignment::odl_GetBookingType(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetBookingType();
}

E_BOOKINGTYPE Assignment::GetBookingType(void)
{
	return (E_BOOKINGTYPE) mp_yBookingType;
}

void Assignment::odl_SetBookingType(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetBookingType((E_BOOKINGTYPE)newVal);
}

void Assignment::SetBookingType(E_BOOKINGTYPE newVal)
{
	mp_yBookingType = newVal;
}

IDispatch* Assignment::odl_GetActualWorkProtected(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetActualWorkProtected()->GetIDispatch(TRUE);
}

Duration* Assignment::GetActualWorkProtected(void)
{
	return mp_oActualWorkProtected;
}

IDispatch* Assignment::odl_GetActualOvertimeWorkProtected(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetActualOvertimeWorkProtected()->GetIDispatch(TRUE);
}

Duration* Assignment::GetActualOvertimeWorkProtected(void)
{
	return mp_oActualOvertimeWorkProtected;
}

DATE Assignment::odl_GetCreationDate(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetCreationDate();
}

COleDateTime Assignment::GetCreationDate(void)
{
	return (COleDateTime) mp_dtCreationDate;
}

void Assignment::odl_SetCreationDate(DATE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetCreationDate((COleDateTime)newVal);
}

void Assignment::SetCreationDate(COleDateTime newVal)
{
	mp_dtCreationDate = newVal;
}

BSTR Assignment::odl_GetAssnOwner(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetAssnOwner().AllocSysString();
}

CString Assignment::GetAssnOwner(void)
{
	return (CString) mp_sAssnOwner;
}

void Assignment::odl_SetAssnOwner(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetAssnOwner(newVal);
}

void Assignment::SetAssnOwner(CString newVal)
{
	mp_sAssnOwner = newVal;
}

BSTR Assignment::odl_GetAssnOwnerGuid(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetAssnOwnerGuid().AllocSysString();
}

CString Assignment::GetAssnOwnerGuid(void)
{
	return (CString) mp_sAssnOwnerGuid;
}

void Assignment::odl_SetAssnOwnerGuid(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetAssnOwnerGuid(newVal);
}

void Assignment::SetAssnOwnerGuid(CString newVal)
{
	mp_sAssnOwnerGuid = newVal;
}

DOUBLE Assignment::odl_GetBudgetCost(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetBudgetCost();
}

DOUBLE Assignment::GetBudgetCost(void)
{
	return (DOUBLE) mp_cBudgetCost;
}

void Assignment::odl_SetBudgetCost(DOUBLE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetBudgetCost((DOUBLE)newVal);
}

void Assignment::SetBudgetCost(DOUBLE newVal)
{
	mp_cBudgetCost = newVal;
}

IDispatch* Assignment::odl_GetBudgetWork(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetBudgetWork()->GetIDispatch(TRUE);
}

Duration* Assignment::GetBudgetWork(void)
{
	return mp_oBudgetWork;
}

IDispatch* Assignment::odl_GetExtendedAttribute(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetExtendedAttribute()->GetIDispatch(TRUE);
}

AssignmentExtendedAttribute_C* Assignment::GetExtendedAttribute(void)
{
	return mp_oExtendedAttribute_C;
}

IDispatch* Assignment::odl_GetBaseline(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetBaseline()->GetIDispatch(TRUE);
}

AssignmentBaseline_C* Assignment::GetBaseline(void)
{
	return mp_oBaseline_C;
}

BSTR Assignment::odl_Getf404000(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404000().AllocSysString();
}

CString Assignment::Getf404000(void)
{
	return (CString) mp_sf404000;
}

void Assignment::odl_Setf404000(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404000(newVal);
}

void Assignment::Setf404000(CString newVal)
{
	mp_sf404000 = newVal;
}

BSTR Assignment::odl_Getf404001(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404001().AllocSysString();
}

CString Assignment::Getf404001(void)
{
	return (CString) mp_sf404001;
}

void Assignment::odl_Setf404001(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404001(newVal);
}

void Assignment::Setf404001(CString newVal)
{
	mp_sf404001 = newVal;
}

BSTR Assignment::odl_Getf404002(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404002().AllocSysString();
}

CString Assignment::Getf404002(void)
{
	return (CString) mp_sf404002;
}

void Assignment::odl_Setf404002(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404002(newVal);
}

void Assignment::Setf404002(CString newVal)
{
	mp_sf404002 = newVal;
}

BSTR Assignment::odl_Getf404003(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404003().AllocSysString();
}

CString Assignment::Getf404003(void)
{
	return (CString) mp_sf404003;
}

void Assignment::odl_Setf404003(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404003(newVal);
}

void Assignment::Setf404003(CString newVal)
{
	mp_sf404003 = newVal;
}

BSTR Assignment::odl_Getf404004(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404004().AllocSysString();
}

CString Assignment::Getf404004(void)
{
	return (CString) mp_sf404004;
}

void Assignment::odl_Setf404004(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404004(newVal);
}

void Assignment::Setf404004(CString newVal)
{
	mp_sf404004 = newVal;
}

BSTR Assignment::odl_Getf404005(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404005().AllocSysString();
}

CString Assignment::Getf404005(void)
{
	return (CString) mp_sf404005;
}

void Assignment::odl_Setf404005(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404005(newVal);
}

void Assignment::Setf404005(CString newVal)
{
	mp_sf404005 = newVal;
}

BSTR Assignment::odl_Getf404006(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404006().AllocSysString();
}

CString Assignment::Getf404006(void)
{
	return (CString) mp_sf404006;
}

void Assignment::odl_Setf404006(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404006(newVal);
}

void Assignment::Setf404006(CString newVal)
{
	mp_sf404006 = newVal;
}

BSTR Assignment::odl_Getf404007(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404007().AllocSysString();
}

CString Assignment::Getf404007(void)
{
	return (CString) mp_sf404007;
}

void Assignment::odl_Setf404007(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404007(newVal);
}

void Assignment::Setf404007(CString newVal)
{
	mp_sf404007 = newVal;
}

BSTR Assignment::odl_Getf404008(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404008().AllocSysString();
}

CString Assignment::Getf404008(void)
{
	return (CString) mp_sf404008;
}

void Assignment::odl_Setf404008(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404008(newVal);
}

void Assignment::Setf404008(CString newVal)
{
	mp_sf404008 = newVal;
}

BSTR Assignment::odl_Getf404009(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404009().AllocSysString();
}

CString Assignment::Getf404009(void)
{
	return (CString) mp_sf404009;
}

void Assignment::odl_Setf404009(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404009(newVal);
}

void Assignment::Setf404009(CString newVal)
{
	mp_sf404009 = newVal;
}

BSTR Assignment::odl_Getf40400a(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf40400a().AllocSysString();
}

CString Assignment::Getf40400a(void)
{
	return (CString) mp_sf40400a;
}

void Assignment::odl_Setf40400a(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf40400a(newVal);
}

void Assignment::Setf40400a(CString newVal)
{
	mp_sf40400a = newVal;
}

BSTR Assignment::odl_Getf40400b(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf40400b().AllocSysString();
}

CString Assignment::Getf40400b(void)
{
	return (CString) mp_sf40400b;
}

void Assignment::odl_Setf40400b(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf40400b(newVal);
}

void Assignment::Setf40400b(CString newVal)
{
	mp_sf40400b = newVal;
}

BSTR Assignment::odl_Getf40400c(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf40400c().AllocSysString();
}

CString Assignment::Getf40400c(void)
{
	return (CString) mp_sf40400c;
}

void Assignment::odl_Setf40400c(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf40400c(newVal);
}

void Assignment::Setf40400c(CString newVal)
{
	mp_sf40400c = newVal;
}

BSTR Assignment::odl_Getf40400d(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf40400d().AllocSysString();
}

CString Assignment::Getf40400d(void)
{
	return (CString) mp_sf40400d;
}

void Assignment::odl_Setf40400d(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf40400d(newVal);
}

void Assignment::Setf40400d(CString newVal)
{
	mp_sf40400d = newVal;
}

BSTR Assignment::odl_Getf40400e(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf40400e().AllocSysString();
}

CString Assignment::Getf40400e(void)
{
	return (CString) mp_sf40400e;
}

void Assignment::odl_Setf40400e(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf40400e(newVal);
}

void Assignment::Setf40400e(CString newVal)
{
	mp_sf40400e = newVal;
}

BSTR Assignment::odl_Getf40400f(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf40400f().AllocSysString();
}

CString Assignment::Getf40400f(void)
{
	return (CString) mp_sf40400f;
}

void Assignment::odl_Setf40400f(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf40400f(newVal);
}

void Assignment::Setf40400f(CString newVal)
{
	mp_sf40400f = newVal;
}

BSTR Assignment::odl_Getf404010(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404010().AllocSysString();
}

CString Assignment::Getf404010(void)
{
	return (CString) mp_sf404010;
}

void Assignment::odl_Setf404010(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404010(newVal);
}

void Assignment::Setf404010(CString newVal)
{
	mp_sf404010 = newVal;
}

BSTR Assignment::odl_Getf404011(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404011().AllocSysString();
}

CString Assignment::Getf404011(void)
{
	return (CString) mp_sf404011;
}

void Assignment::odl_Setf404011(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404011(newVal);
}

void Assignment::Setf404011(CString newVal)
{
	mp_sf404011 = newVal;
}

BSTR Assignment::odl_Getf404012(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404012().AllocSysString();
}

CString Assignment::Getf404012(void)
{
	return (CString) mp_sf404012;
}

void Assignment::odl_Setf404012(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404012(newVal);
}

void Assignment::Setf404012(CString newVal)
{
	mp_sf404012 = newVal;
}

BSTR Assignment::odl_Getf404013(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404013().AllocSysString();
}

CString Assignment::Getf404013(void)
{
	return (CString) mp_sf404013;
}

void Assignment::odl_Setf404013(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404013(newVal);
}

void Assignment::Setf404013(CString newVal)
{
	mp_sf404013 = newVal;
}

BSTR Assignment::odl_Getf404014(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404014().AllocSysString();
}

CString Assignment::Getf404014(void)
{
	return (CString) mp_sf404014;
}

void Assignment::odl_Setf404014(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404014(newVal);
}

void Assignment::Setf404014(CString newVal)
{
	mp_sf404014 = newVal;
}

BSTR Assignment::odl_Getf404015(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404015().AllocSysString();
}

CString Assignment::Getf404015(void)
{
	return (CString) mp_sf404015;
}

void Assignment::odl_Setf404015(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404015(newVal);
}

void Assignment::Setf404015(CString newVal)
{
	mp_sf404015 = newVal;
}

BSTR Assignment::odl_Getf404016(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404016().AllocSysString();
}

CString Assignment::Getf404016(void)
{
	return (CString) mp_sf404016;
}

void Assignment::odl_Setf404016(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404016(newVal);
}

void Assignment::Setf404016(CString newVal)
{
	mp_sf404016 = newVal;
}

BSTR Assignment::odl_Getf404017(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404017().AllocSysString();
}

CString Assignment::Getf404017(void)
{
	return (CString) mp_sf404017;
}

void Assignment::odl_Setf404017(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404017(newVal);
}

void Assignment::Setf404017(CString newVal)
{
	mp_sf404017 = newVal;
}

BSTR Assignment::odl_Getf404018(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404018().AllocSysString();
}

CString Assignment::Getf404018(void)
{
	return (CString) mp_sf404018;
}

void Assignment::odl_Setf404018(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404018(newVal);
}

void Assignment::Setf404018(CString newVal)
{
	mp_sf404018 = newVal;
}

BSTR Assignment::odl_Getf404019(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404019().AllocSysString();
}

CString Assignment::Getf404019(void)
{
	return (CString) mp_sf404019;
}

void Assignment::odl_Setf404019(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404019(newVal);
}

void Assignment::Setf404019(CString newVal)
{
	mp_sf404019 = newVal;
}

BSTR Assignment::odl_Getf40401a(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf40401a().AllocSysString();
}

CString Assignment::Getf40401a(void)
{
	return (CString) mp_sf40401a;
}

void Assignment::odl_Setf40401a(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf40401a(newVal);
}

void Assignment::Setf40401a(CString newVal)
{
	mp_sf40401a = newVal;
}

BSTR Assignment::odl_Getf40401b(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf40401b().AllocSysString();
}

CString Assignment::Getf40401b(void)
{
	return (CString) mp_sf40401b;
}

void Assignment::odl_Setf40401b(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf40401b(newVal);
}

void Assignment::Setf40401b(CString newVal)
{
	mp_sf40401b = newVal;
}

BSTR Assignment::odl_Getf40401c(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf40401c().AllocSysString();
}

CString Assignment::Getf40401c(void)
{
	return (CString) mp_sf40401c;
}

void Assignment::odl_Setf40401c(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf40401c(newVal);
}

void Assignment::Setf40401c(CString newVal)
{
	mp_sf40401c = newVal;
}

BSTR Assignment::odl_Getf40401d(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf40401d().AllocSysString();
}

CString Assignment::Getf40401d(void)
{
	return (CString) mp_sf40401d;
}

void Assignment::odl_Setf40401d(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf40401d(newVal);
}

void Assignment::Setf40401d(CString newVal)
{
	mp_sf40401d = newVal;
}

BSTR Assignment::odl_Getf40401e(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf40401e().AllocSysString();
}

CString Assignment::Getf40401e(void)
{
	return (CString) mp_sf40401e;
}

void Assignment::odl_Setf40401e(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf40401e(newVal);
}

void Assignment::Setf40401e(CString newVal)
{
	mp_sf40401e = newVal;
}

BSTR Assignment::odl_Getf40401f(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf40401f().AllocSysString();
}

CString Assignment::Getf40401f(void)
{
	return (CString) mp_sf40401f;
}

void Assignment::odl_Setf40401f(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf40401f(newVal);
}

void Assignment::Setf40401f(CString newVal)
{
	mp_sf40401f = newVal;
}

BSTR Assignment::odl_Getf404020(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404020().AllocSysString();
}

CString Assignment::Getf404020(void)
{
	return (CString) mp_sf404020;
}

void Assignment::odl_Setf404020(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404020(newVal);
}

void Assignment::Setf404020(CString newVal)
{
	mp_sf404020 = newVal;
}

BSTR Assignment::odl_Getf404021(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404021().AllocSysString();
}

CString Assignment::Getf404021(void)
{
	return (CString) mp_sf404021;
}

void Assignment::odl_Setf404021(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404021(newVal);
}

void Assignment::Setf404021(CString newVal)
{
	mp_sf404021 = newVal;
}

BSTR Assignment::odl_Getf404022(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404022().AllocSysString();
}

CString Assignment::Getf404022(void)
{
	return (CString) mp_sf404022;
}

void Assignment::odl_Setf404022(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404022(newVal);
}

void Assignment::Setf404022(CString newVal)
{
	mp_sf404022 = newVal;
}

BSTR Assignment::odl_Getf404023(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404023().AllocSysString();
}

CString Assignment::Getf404023(void)
{
	return (CString) mp_sf404023;
}

void Assignment::odl_Setf404023(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404023(newVal);
}

void Assignment::Setf404023(CString newVal)
{
	mp_sf404023 = newVal;
}

BSTR Assignment::odl_Getf404024(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404024().AllocSysString();
}

CString Assignment::Getf404024(void)
{
	return (CString) mp_sf404024;
}

void Assignment::odl_Setf404024(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404024(newVal);
}

void Assignment::Setf404024(CString newVal)
{
	mp_sf404024 = newVal;
}

BSTR Assignment::odl_Getf404025(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404025().AllocSysString();
}

CString Assignment::Getf404025(void)
{
	return (CString) mp_sf404025;
}

void Assignment::odl_Setf404025(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404025(newVal);
}

void Assignment::Setf404025(CString newVal)
{
	mp_sf404025 = newVal;
}

BSTR Assignment::odl_Getf404026(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404026().AllocSysString();
}

CString Assignment::Getf404026(void)
{
	return (CString) mp_sf404026;
}

void Assignment::odl_Setf404026(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404026(newVal);
}

void Assignment::Setf404026(CString newVal)
{
	mp_sf404026 = newVal;
}

BSTR Assignment::odl_Getf404027(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404027().AllocSysString();
}

CString Assignment::Getf404027(void)
{
	return (CString) mp_sf404027;
}

void Assignment::odl_Setf404027(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404027(newVal);
}

void Assignment::Setf404027(CString newVal)
{
	mp_sf404027 = newVal;
}

BSTR Assignment::odl_Getf404028(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404028().AllocSysString();
}

CString Assignment::Getf404028(void)
{
	return (CString) mp_sf404028;
}

void Assignment::odl_Setf404028(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404028(newVal);
}

void Assignment::Setf404028(CString newVal)
{
	mp_sf404028 = newVal;
}

BSTR Assignment::odl_Getf404029(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404029().AllocSysString();
}

CString Assignment::Getf404029(void)
{
	return (CString) mp_sf404029;
}

void Assignment::odl_Setf404029(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404029(newVal);
}

void Assignment::Setf404029(CString newVal)
{
	mp_sf404029 = newVal;
}

BSTR Assignment::odl_Getf40402a(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf40402a().AllocSysString();
}

CString Assignment::Getf40402a(void)
{
	return (CString) mp_sf40402a;
}

void Assignment::odl_Setf40402a(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf40402a(newVal);
}

void Assignment::Setf40402a(CString newVal)
{
	mp_sf40402a = newVal;
}

BSTR Assignment::odl_Getf40402b(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf40402b().AllocSysString();
}

CString Assignment::Getf40402b(void)
{
	return (CString) mp_sf40402b;
}

void Assignment::odl_Setf40402b(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf40402b(newVal);
}

void Assignment::Setf40402b(CString newVal)
{
	mp_sf40402b = newVal;
}

BSTR Assignment::odl_Getf40402c(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf40402c().AllocSysString();
}

CString Assignment::Getf40402c(void)
{
	return (CString) mp_sf40402c;
}

void Assignment::odl_Setf40402c(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf40402c(newVal);
}

void Assignment::Setf40402c(CString newVal)
{
	mp_sf40402c = newVal;
}

BSTR Assignment::odl_Getf40402d(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf40402d().AllocSysString();
}

CString Assignment::Getf40402d(void)
{
	return (CString) mp_sf40402d;
}

void Assignment::odl_Setf40402d(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf40402d(newVal);
}

void Assignment::Setf40402d(CString newVal)
{
	mp_sf40402d = newVal;
}

BSTR Assignment::odl_Getf40402e(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf40402e().AllocSysString();
}

CString Assignment::Getf40402e(void)
{
	return (CString) mp_sf40402e;
}

void Assignment::odl_Setf40402e(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf40402e(newVal);
}

void Assignment::Setf40402e(CString newVal)
{
	mp_sf40402e = newVal;
}

BSTR Assignment::odl_Getf40402f(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf40402f().AllocSysString();
}

CString Assignment::Getf40402f(void)
{
	return (CString) mp_sf40402f;
}

void Assignment::odl_Setf40402f(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf40402f(newVal);
}

void Assignment::Setf40402f(CString newVal)
{
	mp_sf40402f = newVal;
}

BSTR Assignment::odl_Getf404030(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404030().AllocSysString();
}

CString Assignment::Getf404030(void)
{
	return (CString) mp_sf404030;
}

void Assignment::odl_Setf404030(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404030(newVal);
}

void Assignment::Setf404030(CString newVal)
{
	mp_sf404030 = newVal;
}

BSTR Assignment::odl_Getf404031(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404031().AllocSysString();
}

CString Assignment::Getf404031(void)
{
	return (CString) mp_sf404031;
}

void Assignment::odl_Setf404031(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404031(newVal);
}

void Assignment::Setf404031(CString newVal)
{
	mp_sf404031 = newVal;
}

BSTR Assignment::odl_Getf404032(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404032().AllocSysString();
}

CString Assignment::Getf404032(void)
{
	return (CString) mp_sf404032;
}

void Assignment::odl_Setf404032(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404032(newVal);
}

void Assignment::Setf404032(CString newVal)
{
	mp_sf404032 = newVal;
}

BSTR Assignment::odl_Getf404033(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404033().AllocSysString();
}

CString Assignment::Getf404033(void)
{
	return (CString) mp_sf404033;
}

void Assignment::odl_Setf404033(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404033(newVal);
}

void Assignment::Setf404033(CString newVal)
{
	mp_sf404033 = newVal;
}

BSTR Assignment::odl_Getf404034(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404034().AllocSysString();
}

CString Assignment::Getf404034(void)
{
	return (CString) mp_sf404034;
}

void Assignment::odl_Setf404034(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404034(newVal);
}

void Assignment::Setf404034(CString newVal)
{
	mp_sf404034 = newVal;
}

BSTR Assignment::odl_Getf404035(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404035().AllocSysString();
}

CString Assignment::Getf404035(void)
{
	return (CString) mp_sf404035;
}

void Assignment::odl_Setf404035(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404035(newVal);
}

void Assignment::Setf404035(CString newVal)
{
	mp_sf404035 = newVal;
}

BSTR Assignment::odl_Getf404036(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404036().AllocSysString();
}

CString Assignment::Getf404036(void)
{
	return (CString) mp_sf404036;
}

void Assignment::odl_Setf404036(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404036(newVal);
}

void Assignment::Setf404036(CString newVal)
{
	mp_sf404036 = newVal;
}

BSTR Assignment::odl_Getf404037(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404037().AllocSysString();
}

CString Assignment::Getf404037(void)
{
	return (CString) mp_sf404037;
}

void Assignment::odl_Setf404037(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404037(newVal);
}

void Assignment::Setf404037(CString newVal)
{
	mp_sf404037 = newVal;
}

BSTR Assignment::odl_Getf404038(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404038().AllocSysString();
}

CString Assignment::Getf404038(void)
{
	return (CString) mp_sf404038;
}

void Assignment::odl_Setf404038(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404038(newVal);
}

void Assignment::Setf404038(CString newVal)
{
	mp_sf404038 = newVal;
}

BSTR Assignment::odl_Getf404039(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404039().AllocSysString();
}

CString Assignment::Getf404039(void)
{
	return (CString) mp_sf404039;
}

void Assignment::odl_Setf404039(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404039(newVal);
}

void Assignment::Setf404039(CString newVal)
{
	mp_sf404039 = newVal;
}

BSTR Assignment::odl_Getf40403a(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf40403a().AllocSysString();
}

CString Assignment::Getf40403a(void)
{
	return (CString) mp_sf40403a;
}

void Assignment::odl_Setf40403a(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf40403a(newVal);
}

void Assignment::Setf40403a(CString newVal)
{
	mp_sf40403a = newVal;
}

BSTR Assignment::odl_Getf40403b(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf40403b().AllocSysString();
}

CString Assignment::Getf40403b(void)
{
	return (CString) mp_sf40403b;
}

void Assignment::odl_Setf40403b(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf40403b(newVal);
}

void Assignment::Setf40403b(CString newVal)
{
	mp_sf40403b = newVal;
}

BSTR Assignment::odl_Getf40403c(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf40403c().AllocSysString();
}

CString Assignment::Getf40403c(void)
{
	return (CString) mp_sf40403c;
}

void Assignment::odl_Setf40403c(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf40403c(newVal);
}

void Assignment::Setf40403c(CString newVal)
{
	mp_sf40403c = newVal;
}

BSTR Assignment::odl_Getf40403d(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf40403d().AllocSysString();
}

CString Assignment::Getf40403d(void)
{
	return (CString) mp_sf40403d;
}

void Assignment::odl_Setf40403d(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf40403d(newVal);
}

void Assignment::Setf40403d(CString newVal)
{
	mp_sf40403d = newVal;
}

BSTR Assignment::odl_Getf40403e(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf40403e().AllocSysString();
}

CString Assignment::Getf40403e(void)
{
	return (CString) mp_sf40403e;
}

void Assignment::odl_Setf40403e(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf40403e(newVal);
}

void Assignment::Setf40403e(CString newVal)
{
	mp_sf40403e = newVal;
}

BSTR Assignment::odl_Getf40403f(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf40403f().AllocSysString();
}

CString Assignment::Getf40403f(void)
{
	return (CString) mp_sf40403f;
}

void Assignment::odl_Setf40403f(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf40403f(newVal);
}

void Assignment::Setf40403f(CString newVal)
{
	mp_sf40403f = newVal;
}

BSTR Assignment::odl_Getf404040(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404040().AllocSysString();
}

CString Assignment::Getf404040(void)
{
	return (CString) mp_sf404040;
}

void Assignment::odl_Setf404040(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404040(newVal);
}

void Assignment::Setf404040(CString newVal)
{
	mp_sf404040 = newVal;
}

BSTR Assignment::odl_Getf404041(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404041().AllocSysString();
}

CString Assignment::Getf404041(void)
{
	return (CString) mp_sf404041;
}

void Assignment::odl_Setf404041(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404041(newVal);
}

void Assignment::Setf404041(CString newVal)
{
	mp_sf404041 = newVal;
}

BSTR Assignment::odl_Getf404042(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404042().AllocSysString();
}

CString Assignment::Getf404042(void)
{
	return (CString) mp_sf404042;
}

void Assignment::odl_Setf404042(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404042(newVal);
}

void Assignment::Setf404042(CString newVal)
{
	mp_sf404042 = newVal;
}

BSTR Assignment::odl_Getf404043(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404043().AllocSysString();
}

CString Assignment::Getf404043(void)
{
	return (CString) mp_sf404043;
}

void Assignment::odl_Setf404043(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404043(newVal);
}

void Assignment::Setf404043(CString newVal)
{
	mp_sf404043 = newVal;
}

BSTR Assignment::odl_Getf404044(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404044().AllocSysString();
}

CString Assignment::Getf404044(void)
{
	return (CString) mp_sf404044;
}

void Assignment::odl_Setf404044(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404044(newVal);
}

void Assignment::Setf404044(CString newVal)
{
	mp_sf404044 = newVal;
}

BSTR Assignment::odl_Getf404045(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404045().AllocSysString();
}

CString Assignment::Getf404045(void)
{
	return (CString) mp_sf404045;
}

void Assignment::odl_Setf404045(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404045(newVal);
}

void Assignment::Setf404045(CString newVal)
{
	mp_sf404045 = newVal;
}

BSTR Assignment::odl_Getf404046(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404046().AllocSysString();
}

CString Assignment::Getf404046(void)
{
	return (CString) mp_sf404046;
}

void Assignment::odl_Setf404046(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404046(newVal);
}

void Assignment::Setf404046(CString newVal)
{
	mp_sf404046 = newVal;
}

BSTR Assignment::odl_Getf404047(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404047().AllocSysString();
}

CString Assignment::Getf404047(void)
{
	return (CString) mp_sf404047;
}

void Assignment::odl_Setf404047(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404047(newVal);
}

void Assignment::Setf404047(CString newVal)
{
	mp_sf404047 = newVal;
}

BSTR Assignment::odl_Getf404048(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404048().AllocSysString();
}

CString Assignment::Getf404048(void)
{
	return (CString) mp_sf404048;
}

void Assignment::odl_Setf404048(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404048(newVal);
}

void Assignment::Setf404048(CString newVal)
{
	mp_sf404048 = newVal;
}

BSTR Assignment::odl_Getf404049(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404049().AllocSysString();
}

CString Assignment::Getf404049(void)
{
	return (CString) mp_sf404049;
}

void Assignment::odl_Setf404049(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404049(newVal);
}

void Assignment::Setf404049(CString newVal)
{
	mp_sf404049 = newVal;
}

BSTR Assignment::odl_Getf40404a(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf40404a().AllocSysString();
}

CString Assignment::Getf40404a(void)
{
	return (CString) mp_sf40404a;
}

void Assignment::odl_Setf40404a(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf40404a(newVal);
}

void Assignment::Setf40404a(CString newVal)
{
	mp_sf40404a = newVal;
}

BSTR Assignment::odl_Getf40404b(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf40404b().AllocSysString();
}

CString Assignment::Getf40404b(void)
{
	return (CString) mp_sf40404b;
}

void Assignment::odl_Setf40404b(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf40404b(newVal);
}

void Assignment::Setf40404b(CString newVal)
{
	mp_sf40404b = newVal;
}

BSTR Assignment::odl_Getf40404c(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf40404c().AllocSysString();
}

CString Assignment::Getf40404c(void)
{
	return (CString) mp_sf40404c;
}

void Assignment::odl_Setf40404c(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf40404c(newVal);
}

void Assignment::Setf40404c(CString newVal)
{
	mp_sf40404c = newVal;
}

BSTR Assignment::odl_Getf40404d(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf40404d().AllocSysString();
}

CString Assignment::Getf40404d(void)
{
	return (CString) mp_sf40404d;
}

void Assignment::odl_Setf40404d(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf40404d(newVal);
}

void Assignment::Setf40404d(CString newVal)
{
	mp_sf40404d = newVal;
}

BSTR Assignment::odl_Getf40404e(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf40404e().AllocSysString();
}

CString Assignment::Getf40404e(void)
{
	return (CString) mp_sf40404e;
}

void Assignment::odl_Setf40404e(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf40404e(newVal);
}

void Assignment::Setf40404e(CString newVal)
{
	mp_sf40404e = newVal;
}

BSTR Assignment::odl_Getf40404f(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf40404f().AllocSysString();
}

CString Assignment::Getf40404f(void)
{
	return (CString) mp_sf40404f;
}

void Assignment::odl_Setf40404f(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf40404f(newVal);
}

void Assignment::Setf40404f(CString newVal)
{
	mp_sf40404f = newVal;
}

BSTR Assignment::odl_Getf404050(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404050().AllocSysString();
}

CString Assignment::Getf404050(void)
{
	return (CString) mp_sf404050;
}

void Assignment::odl_Setf404050(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404050(newVal);
}

void Assignment::Setf404050(CString newVal)
{
	mp_sf404050 = newVal;
}

BSTR Assignment::odl_Getf404051(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404051().AllocSysString();
}

CString Assignment::Getf404051(void)
{
	return (CString) mp_sf404051;
}

void Assignment::odl_Setf404051(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404051(newVal);
}

void Assignment::Setf404051(CString newVal)
{
	mp_sf404051 = newVal;
}

BSTR Assignment::odl_Getf404052(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404052().AllocSysString();
}

CString Assignment::Getf404052(void)
{
	return (CString) mp_sf404052;
}

void Assignment::odl_Setf404052(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404052(newVal);
}

void Assignment::Setf404052(CString newVal)
{
	mp_sf404052 = newVal;
}

BSTR Assignment::odl_Getf404053(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404053().AllocSysString();
}

CString Assignment::Getf404053(void)
{
	return (CString) mp_sf404053;
}

void Assignment::odl_Setf404053(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404053(newVal);
}

void Assignment::Setf404053(CString newVal)
{
	mp_sf404053 = newVal;
}

BSTR Assignment::odl_Getf404054(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404054().AllocSysString();
}

CString Assignment::Getf404054(void)
{
	return (CString) mp_sf404054;
}

void Assignment::odl_Setf404054(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404054(newVal);
}

void Assignment::Setf404054(CString newVal)
{
	mp_sf404054 = newVal;
}

BSTR Assignment::odl_Getf404055(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404055().AllocSysString();
}

CString Assignment::Getf404055(void)
{
	return (CString) mp_sf404055;
}

void Assignment::odl_Setf404055(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404055(newVal);
}

void Assignment::Setf404055(CString newVal)
{
	mp_sf404055 = newVal;
}

BSTR Assignment::odl_Getf404056(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404056().AllocSysString();
}

CString Assignment::Getf404056(void)
{
	return (CString) mp_sf404056;
}

void Assignment::odl_Setf404056(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404056(newVal);
}

void Assignment::Setf404056(CString newVal)
{
	mp_sf404056 = newVal;
}

BSTR Assignment::odl_Getf404057(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404057().AllocSysString();
}

CString Assignment::Getf404057(void)
{
	return (CString) mp_sf404057;
}

void Assignment::odl_Setf404057(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404057(newVal);
}

void Assignment::Setf404057(CString newVal)
{
	mp_sf404057 = newVal;
}

BSTR Assignment::odl_Getf404058(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404058().AllocSysString();
}

CString Assignment::Getf404058(void)
{
	return (CString) mp_sf404058;
}

void Assignment::odl_Setf404058(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404058(newVal);
}

void Assignment::Setf404058(CString newVal)
{
	mp_sf404058 = newVal;
}

BSTR Assignment::odl_Getf404059(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404059().AllocSysString();
}

CString Assignment::Getf404059(void)
{
	return (CString) mp_sf404059;
}

void Assignment::odl_Setf404059(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404059(newVal);
}

void Assignment::Setf404059(CString newVal)
{
	mp_sf404059 = newVal;
}

BSTR Assignment::odl_Getf40405a(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf40405a().AllocSysString();
}

CString Assignment::Getf40405a(void)
{
	return (CString) mp_sf40405a;
}

void Assignment::odl_Setf40405a(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf40405a(newVal);
}

void Assignment::Setf40405a(CString newVal)
{
	mp_sf40405a = newVal;
}

BSTR Assignment::odl_Getf40405b(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf40405b().AllocSysString();
}

CString Assignment::Getf40405b(void)
{
	return (CString) mp_sf40405b;
}

void Assignment::odl_Setf40405b(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf40405b(newVal);
}

void Assignment::Setf40405b(CString newVal)
{
	mp_sf40405b = newVal;
}

BSTR Assignment::odl_Getf40405c(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf40405c().AllocSysString();
}

CString Assignment::Getf40405c(void)
{
	return (CString) mp_sf40405c;
}

void Assignment::odl_Setf40405c(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf40405c(newVal);
}

void Assignment::Setf40405c(CString newVal)
{
	mp_sf40405c = newVal;
}

BSTR Assignment::odl_Getf40405d(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf40405d().AllocSysString();
}

CString Assignment::Getf40405d(void)
{
	return (CString) mp_sf40405d;
}

void Assignment::odl_Setf40405d(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf40405d(newVal);
}

void Assignment::Setf40405d(CString newVal)
{
	mp_sf40405d = newVal;
}

BSTR Assignment::odl_Getf40405e(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf40405e().AllocSysString();
}

CString Assignment::Getf40405e(void)
{
	return (CString) mp_sf40405e;
}

void Assignment::odl_Setf40405e(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf40405e(newVal);
}

void Assignment::Setf40405e(CString newVal)
{
	mp_sf40405e = newVal;
}

BSTR Assignment::odl_Getf40405f(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf40405f().AllocSysString();
}

CString Assignment::Getf40405f(void)
{
	return (CString) mp_sf40405f;
}

void Assignment::odl_Setf40405f(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf40405f(newVal);
}

void Assignment::Setf40405f(CString newVal)
{
	mp_sf40405f = newVal;
}

BSTR Assignment::odl_Getf404060(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404060().AllocSysString();
}

CString Assignment::Getf404060(void)
{
	return (CString) mp_sf404060;
}

void Assignment::odl_Setf404060(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404060(newVal);
}

void Assignment::Setf404060(CString newVal)
{
	mp_sf404060 = newVal;
}

BSTR Assignment::odl_Getf404061(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404061().AllocSysString();
}

CString Assignment::Getf404061(void)
{
	return (CString) mp_sf404061;
}

void Assignment::odl_Setf404061(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404061(newVal);
}

void Assignment::Setf404061(CString newVal)
{
	mp_sf404061 = newVal;
}

BSTR Assignment::odl_Getf404062(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404062().AllocSysString();
}

CString Assignment::Getf404062(void)
{
	return (CString) mp_sf404062;
}

void Assignment::odl_Setf404062(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404062(newVal);
}

void Assignment::Setf404062(CString newVal)
{
	mp_sf404062 = newVal;
}

BSTR Assignment::odl_Getf404063(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404063().AllocSysString();
}

CString Assignment::Getf404063(void)
{
	return (CString) mp_sf404063;
}

void Assignment::odl_Setf404063(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404063(newVal);
}

void Assignment::Setf404063(CString newVal)
{
	mp_sf404063 = newVal;
}

BSTR Assignment::odl_Getf404064(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404064().AllocSysString();
}

CString Assignment::Getf404064(void)
{
	return (CString) mp_sf404064;
}

void Assignment::odl_Setf404064(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404064(newVal);
}

void Assignment::Setf404064(CString newVal)
{
	mp_sf404064 = newVal;
}

BSTR Assignment::odl_Getf404065(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404065().AllocSysString();
}

CString Assignment::Getf404065(void)
{
	return (CString) mp_sf404065;
}

void Assignment::odl_Setf404065(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404065(newVal);
}

void Assignment::Setf404065(CString newVal)
{
	mp_sf404065 = newVal;
}

BSTR Assignment::odl_Getf404066(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404066().AllocSysString();
}

CString Assignment::Getf404066(void)
{
	return (CString) mp_sf404066;
}

void Assignment::odl_Setf404066(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404066(newVal);
}

void Assignment::Setf404066(CString newVal)
{
	mp_sf404066 = newVal;
}

BSTR Assignment::odl_Getf404067(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404067().AllocSysString();
}

CString Assignment::Getf404067(void)
{
	return (CString) mp_sf404067;
}

void Assignment::odl_Setf404067(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404067(newVal);
}

void Assignment::Setf404067(CString newVal)
{
	mp_sf404067 = newVal;
}

BSTR Assignment::odl_Getf404068(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404068().AllocSysString();
}

CString Assignment::Getf404068(void)
{
	return (CString) mp_sf404068;
}

void Assignment::odl_Setf404068(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404068(newVal);
}

void Assignment::Setf404068(CString newVal)
{
	mp_sf404068 = newVal;
}

BSTR Assignment::odl_Getf404069(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404069().AllocSysString();
}

CString Assignment::Getf404069(void)
{
	return (CString) mp_sf404069;
}

void Assignment::odl_Setf404069(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404069(newVal);
}

void Assignment::Setf404069(CString newVal)
{
	mp_sf404069 = newVal;
}

BSTR Assignment::odl_Getf40406a(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf40406a().AllocSysString();
}

CString Assignment::Getf40406a(void)
{
	return (CString) mp_sf40406a;
}

void Assignment::odl_Setf40406a(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf40406a(newVal);
}

void Assignment::Setf40406a(CString newVal)
{
	mp_sf40406a = newVal;
}

BSTR Assignment::odl_Getf40406b(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf40406b().AllocSysString();
}

CString Assignment::Getf40406b(void)
{
	return (CString) mp_sf40406b;
}

void Assignment::odl_Setf40406b(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf40406b(newVal);
}

void Assignment::Setf40406b(CString newVal)
{
	mp_sf40406b = newVal;
}

BSTR Assignment::odl_Getf40406c(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf40406c().AllocSysString();
}

CString Assignment::Getf40406c(void)
{
	return (CString) mp_sf40406c;
}

void Assignment::odl_Setf40406c(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf40406c(newVal);
}

void Assignment::Setf40406c(CString newVal)
{
	mp_sf40406c = newVal;
}

BSTR Assignment::odl_Getf40406d(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf40406d().AllocSysString();
}

CString Assignment::Getf40406d(void)
{
	return (CString) mp_sf40406d;
}

void Assignment::odl_Setf40406d(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf40406d(newVal);
}

void Assignment::Setf40406d(CString newVal)
{
	mp_sf40406d = newVal;
}

BSTR Assignment::odl_Getf40406e(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf40406e().AllocSysString();
}

CString Assignment::Getf40406e(void)
{
	return (CString) mp_sf40406e;
}

void Assignment::odl_Setf40406e(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf40406e(newVal);
}

void Assignment::Setf40406e(CString newVal)
{
	mp_sf40406e = newVal;
}

BSTR Assignment::odl_Getf40406f(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf40406f().AllocSysString();
}

CString Assignment::Getf40406f(void)
{
	return (CString) mp_sf40406f;
}

void Assignment::odl_Setf40406f(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf40406f(newVal);
}

void Assignment::Setf40406f(CString newVal)
{
	mp_sf40406f = newVal;
}

BSTR Assignment::odl_Getf404070(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404070().AllocSysString();
}

CString Assignment::Getf404070(void)
{
	return (CString) mp_sf404070;
}

void Assignment::odl_Setf404070(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404070(newVal);
}

void Assignment::Setf404070(CString newVal)
{
	mp_sf404070 = newVal;
}

BSTR Assignment::odl_Getf404071(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404071().AllocSysString();
}

CString Assignment::Getf404071(void)
{
	return (CString) mp_sf404071;
}

void Assignment::odl_Setf404071(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404071(newVal);
}

void Assignment::Setf404071(CString newVal)
{
	mp_sf404071 = newVal;
}

BSTR Assignment::odl_Getf404072(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404072().AllocSysString();
}

CString Assignment::Getf404072(void)
{
	return (CString) mp_sf404072;
}

void Assignment::odl_Setf404072(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404072(newVal);
}

void Assignment::Setf404072(CString newVal)
{
	mp_sf404072 = newVal;
}

BSTR Assignment::odl_Getf404073(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404073().AllocSysString();
}

CString Assignment::Getf404073(void)
{
	return (CString) mp_sf404073;
}

void Assignment::odl_Setf404073(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404073(newVal);
}

void Assignment::Setf404073(CString newVal)
{
	mp_sf404073 = newVal;
}

BSTR Assignment::odl_Getf404074(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404074().AllocSysString();
}

CString Assignment::Getf404074(void)
{
	return (CString) mp_sf404074;
}

void Assignment::odl_Setf404074(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404074(newVal);
}

void Assignment::Setf404074(CString newVal)
{
	mp_sf404074 = newVal;
}

BSTR Assignment::odl_Getf404075(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404075().AllocSysString();
}

CString Assignment::Getf404075(void)
{
	return (CString) mp_sf404075;
}

void Assignment::odl_Setf404075(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404075(newVal);
}

void Assignment::Setf404075(CString newVal)
{
	mp_sf404075 = newVal;
}

BSTR Assignment::odl_Getf404076(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404076().AllocSysString();
}

CString Assignment::Getf404076(void)
{
	return (CString) mp_sf404076;
}

void Assignment::odl_Setf404076(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404076(newVal);
}

void Assignment::Setf404076(CString newVal)
{
	mp_sf404076 = newVal;
}

BSTR Assignment::odl_Getf404077(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404077().AllocSysString();
}

CString Assignment::Getf404077(void)
{
	return (CString) mp_sf404077;
}

void Assignment::odl_Setf404077(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404077(newVal);
}

void Assignment::Setf404077(CString newVal)
{
	mp_sf404077 = newVal;
}

BSTR Assignment::odl_Getf404078(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404078().AllocSysString();
}

CString Assignment::Getf404078(void)
{
	return (CString) mp_sf404078;
}

void Assignment::odl_Setf404078(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404078(newVal);
}

void Assignment::Setf404078(CString newVal)
{
	mp_sf404078 = newVal;
}

BSTR Assignment::odl_Getf404079(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404079().AllocSysString();
}

CString Assignment::Getf404079(void)
{
	return (CString) mp_sf404079;
}

void Assignment::odl_Setf404079(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404079(newVal);
}

void Assignment::Setf404079(CString newVal)
{
	mp_sf404079 = newVal;
}

BSTR Assignment::odl_Getf40407a(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf40407a().AllocSysString();
}

CString Assignment::Getf40407a(void)
{
	return (CString) mp_sf40407a;
}

void Assignment::odl_Setf40407a(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf40407a(newVal);
}

void Assignment::Setf40407a(CString newVal)
{
	mp_sf40407a = newVal;
}

BSTR Assignment::odl_Getf40407b(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf40407b().AllocSysString();
}

CString Assignment::Getf40407b(void)
{
	return (CString) mp_sf40407b;
}

void Assignment::odl_Setf40407b(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf40407b(newVal);
}

void Assignment::Setf40407b(CString newVal)
{
	mp_sf40407b = newVal;
}

BSTR Assignment::odl_Getf40407c(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf40407c().AllocSysString();
}

CString Assignment::Getf40407c(void)
{
	return (CString) mp_sf40407c;
}

void Assignment::odl_Setf40407c(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf40407c(newVal);
}

void Assignment::Setf40407c(CString newVal)
{
	mp_sf40407c = newVal;
}

BSTR Assignment::odl_Getf40407d(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf40407d().AllocSysString();
}

CString Assignment::Getf40407d(void)
{
	return (CString) mp_sf40407d;
}

void Assignment::odl_Setf40407d(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf40407d(newVal);
}

void Assignment::Setf40407d(CString newVal)
{
	mp_sf40407d = newVal;
}

BSTR Assignment::odl_Getf40407e(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf40407e().AllocSysString();
}

CString Assignment::Getf40407e(void)
{
	return (CString) mp_sf40407e;
}

void Assignment::odl_Setf40407e(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf40407e(newVal);
}

void Assignment::Setf40407e(CString newVal)
{
	mp_sf40407e = newVal;
}

BSTR Assignment::odl_Getf40407f(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf40407f().AllocSysString();
}

CString Assignment::Getf40407f(void)
{
	return (CString) mp_sf40407f;
}

void Assignment::odl_Setf40407f(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf40407f(newVal);
}

void Assignment::Setf40407f(CString newVal)
{
	mp_sf40407f = newVal;
}

BSTR Assignment::odl_Getf404080(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404080().AllocSysString();
}

CString Assignment::Getf404080(void)
{
	return (CString) mp_sf404080;
}

void Assignment::odl_Setf404080(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404080(newVal);
}

void Assignment::Setf404080(CString newVal)
{
	mp_sf404080 = newVal;
}

BSTR Assignment::odl_Getf404081(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404081().AllocSysString();
}

CString Assignment::Getf404081(void)
{
	return (CString) mp_sf404081;
}

void Assignment::odl_Setf404081(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404081(newVal);
}

void Assignment::Setf404081(CString newVal)
{
	mp_sf404081 = newVal;
}

BSTR Assignment::odl_Getf404082(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404082().AllocSysString();
}

CString Assignment::Getf404082(void)
{
	return (CString) mp_sf404082;
}

void Assignment::odl_Setf404082(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404082(newVal);
}

void Assignment::Setf404082(CString newVal)
{
	mp_sf404082 = newVal;
}

BSTR Assignment::odl_Getf404083(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404083().AllocSysString();
}

CString Assignment::Getf404083(void)
{
	return (CString) mp_sf404083;
}

void Assignment::odl_Setf404083(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404083(newVal);
}

void Assignment::Setf404083(CString newVal)
{
	mp_sf404083 = newVal;
}

BSTR Assignment::odl_Getf404084(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404084().AllocSysString();
}

CString Assignment::Getf404084(void)
{
	return (CString) mp_sf404084;
}

void Assignment::odl_Setf404084(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404084(newVal);
}

void Assignment::Setf404084(CString newVal)
{
	mp_sf404084 = newVal;
}

BSTR Assignment::odl_Getf404085(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404085().AllocSysString();
}

CString Assignment::Getf404085(void)
{
	return (CString) mp_sf404085;
}

void Assignment::odl_Setf404085(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404085(newVal);
}

void Assignment::Setf404085(CString newVal)
{
	mp_sf404085 = newVal;
}

BSTR Assignment::odl_Getf404086(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404086().AllocSysString();
}

CString Assignment::Getf404086(void)
{
	return (CString) mp_sf404086;
}

void Assignment::odl_Setf404086(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404086(newVal);
}

void Assignment::Setf404086(CString newVal)
{
	mp_sf404086 = newVal;
}

BSTR Assignment::odl_Getf404087(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404087().AllocSysString();
}

CString Assignment::Getf404087(void)
{
	return (CString) mp_sf404087;
}

void Assignment::odl_Setf404087(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404087(newVal);
}

void Assignment::Setf404087(CString newVal)
{
	mp_sf404087 = newVal;
}

BSTR Assignment::odl_Getf404088(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404088().AllocSysString();
}

CString Assignment::Getf404088(void)
{
	return (CString) mp_sf404088;
}

void Assignment::odl_Setf404088(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404088(newVal);
}

void Assignment::Setf404088(CString newVal)
{
	mp_sf404088 = newVal;
}

BSTR Assignment::odl_Getf404089(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404089().AllocSysString();
}

CString Assignment::Getf404089(void)
{
	return (CString) mp_sf404089;
}

void Assignment::odl_Setf404089(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404089(newVal);
}

void Assignment::Setf404089(CString newVal)
{
	mp_sf404089 = newVal;
}

BSTR Assignment::odl_Getf40408a(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf40408a().AllocSysString();
}

CString Assignment::Getf40408a(void)
{
	return (CString) mp_sf40408a;
}

void Assignment::odl_Setf40408a(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf40408a(newVal);
}

void Assignment::Setf40408a(CString newVal)
{
	mp_sf40408a = newVal;
}

BSTR Assignment::odl_Getf40408b(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf40408b().AllocSysString();
}

CString Assignment::Getf40408b(void)
{
	return (CString) mp_sf40408b;
}

void Assignment::odl_Setf40408b(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf40408b(newVal);
}

void Assignment::Setf40408b(CString newVal)
{
	mp_sf40408b = newVal;
}

BSTR Assignment::odl_Getf40408c(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf40408c().AllocSysString();
}

CString Assignment::Getf40408c(void)
{
	return (CString) mp_sf40408c;
}

void Assignment::odl_Setf40408c(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf40408c(newVal);
}

void Assignment::Setf40408c(CString newVal)
{
	mp_sf40408c = newVal;
}

BSTR Assignment::odl_Getf40408d(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf40408d().AllocSysString();
}

CString Assignment::Getf40408d(void)
{
	return (CString) mp_sf40408d;
}

void Assignment::odl_Setf40408d(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf40408d(newVal);
}

void Assignment::Setf40408d(CString newVal)
{
	mp_sf40408d = newVal;
}

BSTR Assignment::odl_Getf40408e(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf40408e().AllocSysString();
}

CString Assignment::Getf40408e(void)
{
	return (CString) mp_sf40408e;
}

void Assignment::odl_Setf40408e(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf40408e(newVal);
}

void Assignment::Setf40408e(CString newVal)
{
	mp_sf40408e = newVal;
}

BSTR Assignment::odl_Getf40408f(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf40408f().AllocSysString();
}

CString Assignment::Getf40408f(void)
{
	return (CString) mp_sf40408f;
}

void Assignment::odl_Setf40408f(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf40408f(newVal);
}

void Assignment::Setf40408f(CString newVal)
{
	mp_sf40408f = newVal;
}

BSTR Assignment::odl_Getf404090(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404090().AllocSysString();
}

CString Assignment::Getf404090(void)
{
	return (CString) mp_sf404090;
}

void Assignment::odl_Setf404090(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404090(newVal);
}

void Assignment::Setf404090(CString newVal)
{
	mp_sf404090 = newVal;
}

BSTR Assignment::odl_Getf404091(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404091().AllocSysString();
}

CString Assignment::Getf404091(void)
{
	return (CString) mp_sf404091;
}

void Assignment::odl_Setf404091(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404091(newVal);
}

void Assignment::Setf404091(CString newVal)
{
	mp_sf404091 = newVal;
}

BSTR Assignment::odl_Getf404092(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404092().AllocSysString();
}

CString Assignment::Getf404092(void)
{
	return (CString) mp_sf404092;
}

void Assignment::odl_Setf404092(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404092(newVal);
}

void Assignment::Setf404092(CString newVal)
{
	mp_sf404092 = newVal;
}

BSTR Assignment::odl_Getf404093(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404093().AllocSysString();
}

CString Assignment::Getf404093(void)
{
	return (CString) mp_sf404093;
}

void Assignment::odl_Setf404093(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404093(newVal);
}

void Assignment::Setf404093(CString newVal)
{
	mp_sf404093 = newVal;
}

BSTR Assignment::odl_Getf404094(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404094().AllocSysString();
}

CString Assignment::Getf404094(void)
{
	return (CString) mp_sf404094;
}

void Assignment::odl_Setf404094(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404094(newVal);
}

void Assignment::Setf404094(CString newVal)
{
	mp_sf404094 = newVal;
}

BSTR Assignment::odl_Getf404095(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404095().AllocSysString();
}

CString Assignment::Getf404095(void)
{
	return (CString) mp_sf404095;
}

void Assignment::odl_Setf404095(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404095(newVal);
}

void Assignment::Setf404095(CString newVal)
{
	mp_sf404095 = newVal;
}

BSTR Assignment::odl_Getf404096(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404096().AllocSysString();
}

CString Assignment::Getf404096(void)
{
	return (CString) mp_sf404096;
}

void Assignment::odl_Setf404096(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404096(newVal);
}

void Assignment::Setf404096(CString newVal)
{
	mp_sf404096 = newVal;
}

BSTR Assignment::odl_Getf404097(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404097().AllocSysString();
}

CString Assignment::Getf404097(void)
{
	return (CString) mp_sf404097;
}

void Assignment::odl_Setf404097(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404097(newVal);
}

void Assignment::Setf404097(CString newVal)
{
	mp_sf404097 = newVal;
}

BSTR Assignment::odl_Getf404098(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404098().AllocSysString();
}

CString Assignment::Getf404098(void)
{
	return (CString) mp_sf404098;
}

void Assignment::odl_Setf404098(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404098(newVal);
}

void Assignment::Setf404098(CString newVal)
{
	mp_sf404098 = newVal;
}

BSTR Assignment::odl_Getf404099(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf404099().AllocSysString();
}

CString Assignment::Getf404099(void)
{
	return (CString) mp_sf404099;
}

void Assignment::odl_Setf404099(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf404099(newVal);
}

void Assignment::Setf404099(CString newVal)
{
	mp_sf404099 = newVal;
}

BSTR Assignment::odl_Getf40409a(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf40409a().AllocSysString();
}

CString Assignment::Getf40409a(void)
{
	return (CString) mp_sf40409a;
}

void Assignment::odl_Setf40409a(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf40409a(newVal);
}

void Assignment::Setf40409a(CString newVal)
{
	mp_sf40409a = newVal;
}

BSTR Assignment::odl_Getf40409b(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf40409b().AllocSysString();
}

CString Assignment::Getf40409b(void)
{
	return (CString) mp_sf40409b;
}

void Assignment::odl_Setf40409b(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf40409b(newVal);
}

void Assignment::Setf40409b(CString newVal)
{
	mp_sf40409b = newVal;
}

BSTR Assignment::odl_Getf40409c(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf40409c().AllocSysString();
}

CString Assignment::Getf40409c(void)
{
	return (CString) mp_sf40409c;
}

void Assignment::odl_Setf40409c(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf40409c(newVal);
}

void Assignment::Setf40409c(CString newVal)
{
	mp_sf40409c = newVal;
}

BSTR Assignment::odl_Getf40409d(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf40409d().AllocSysString();
}

CString Assignment::Getf40409d(void)
{
	return (CString) mp_sf40409d;
}

void Assignment::odl_Setf40409d(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf40409d(newVal);
}

void Assignment::Setf40409d(CString newVal)
{
	mp_sf40409d = newVal;
}

BSTR Assignment::odl_Getf40409e(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf40409e().AllocSysString();
}

CString Assignment::Getf40409e(void)
{
	return (CString) mp_sf40409e;
}

void Assignment::odl_Setf40409e(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf40409e(newVal);
}

void Assignment::Setf40409e(CString newVal)
{
	mp_sf40409e = newVal;
}

BSTR Assignment::odl_Getf40409f(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf40409f().AllocSysString();
}

CString Assignment::Getf40409f(void)
{
	return (CString) mp_sf40409f;
}

void Assignment::odl_Setf40409f(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf40409f(newVal);
}

void Assignment::Setf40409f(CString newVal)
{
	mp_sf40409f = newVal;
}

BSTR Assignment::odl_Getf4040a0(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf4040a0().AllocSysString();
}

CString Assignment::Getf4040a0(void)
{
	return (CString) mp_sf4040a0;
}

void Assignment::odl_Setf4040a0(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf4040a0(newVal);
}

void Assignment::Setf4040a0(CString newVal)
{
	mp_sf4040a0 = newVal;
}

BSTR Assignment::odl_Getf4040a1(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf4040a1().AllocSysString();
}

CString Assignment::Getf4040a1(void)
{
	return (CString) mp_sf4040a1;
}

void Assignment::odl_Setf4040a1(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf4040a1(newVal);
}

void Assignment::Setf4040a1(CString newVal)
{
	mp_sf4040a1 = newVal;
}

BSTR Assignment::odl_Getf4040a2(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf4040a2().AllocSysString();
}

CString Assignment::Getf4040a2(void)
{
	return (CString) mp_sf4040a2;
}

void Assignment::odl_Setf4040a2(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf4040a2(newVal);
}

void Assignment::Setf4040a2(CString newVal)
{
	mp_sf4040a2 = newVal;
}

BSTR Assignment::odl_Getf4040a3(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf4040a3().AllocSysString();
}

CString Assignment::Getf4040a3(void)
{
	return (CString) mp_sf4040a3;
}

void Assignment::odl_Setf4040a3(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf4040a3(newVal);
}

void Assignment::Setf4040a3(CString newVal)
{
	mp_sf4040a3 = newVal;
}

BSTR Assignment::odl_Getf4040a4(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf4040a4().AllocSysString();
}

CString Assignment::Getf4040a4(void)
{
	return (CString) mp_sf4040a4;
}

void Assignment::odl_Setf4040a4(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf4040a4(newVal);
}

void Assignment::Setf4040a4(CString newVal)
{
	mp_sf4040a4 = newVal;
}

BSTR Assignment::odl_Getf4040a5(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf4040a5().AllocSysString();
}

CString Assignment::Getf4040a5(void)
{
	return (CString) mp_sf4040a5;
}

void Assignment::odl_Setf4040a5(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf4040a5(newVal);
}

void Assignment::Setf4040a5(CString newVal)
{
	mp_sf4040a5 = newVal;
}

BSTR Assignment::odl_Getf4040a6(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf4040a6().AllocSysString();
}

CString Assignment::Getf4040a6(void)
{
	return (CString) mp_sf4040a6;
}

void Assignment::odl_Setf4040a6(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf4040a6(newVal);
}

void Assignment::Setf4040a6(CString newVal)
{
	mp_sf4040a6 = newVal;
}

BSTR Assignment::odl_Getf4040a7(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf4040a7().AllocSysString();
}

CString Assignment::Getf4040a7(void)
{
	return (CString) mp_sf4040a7;
}

void Assignment::odl_Setf4040a7(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf4040a7(newVal);
}

void Assignment::Setf4040a7(CString newVal)
{
	mp_sf4040a7 = newVal;
}

BSTR Assignment::odl_Getf4040a8(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf4040a8().AllocSysString();
}

CString Assignment::Getf4040a8(void)
{
	return (CString) mp_sf4040a8;
}

void Assignment::odl_Setf4040a8(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf4040a8(newVal);
}

void Assignment::Setf4040a8(CString newVal)
{
	mp_sf4040a8 = newVal;
}

BSTR Assignment::odl_Getf4040a9(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf4040a9().AllocSysString();
}

CString Assignment::Getf4040a9(void)
{
	return (CString) mp_sf4040a9;
}

void Assignment::odl_Setf4040a9(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf4040a9(newVal);
}

void Assignment::Setf4040a9(CString newVal)
{
	mp_sf4040a9 = newVal;
}

BSTR Assignment::odl_Getf4040aa(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf4040aa().AllocSysString();
}

CString Assignment::Getf4040aa(void)
{
	return (CString) mp_sf4040aa;
}

void Assignment::odl_Setf4040aa(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf4040aa(newVal);
}

void Assignment::Setf4040aa(CString newVal)
{
	mp_sf4040aa = newVal;
}

BSTR Assignment::odl_Getf4040ab(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf4040ab().AllocSysString();
}

CString Assignment::Getf4040ab(void)
{
	return (CString) mp_sf4040ab;
}

void Assignment::odl_Setf4040ab(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf4040ab(newVal);
}

void Assignment::Setf4040ab(CString newVal)
{
	mp_sf4040ab = newVal;
}

BSTR Assignment::odl_Getf4040ac(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf4040ac().AllocSysString();
}

CString Assignment::Getf4040ac(void)
{
	return (CString) mp_sf4040ac;
}

void Assignment::odl_Setf4040ac(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf4040ac(newVal);
}

void Assignment::Setf4040ac(CString newVal)
{
	mp_sf4040ac = newVal;
}

BSTR Assignment::odl_Getf4040ad(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf4040ad().AllocSysString();
}

CString Assignment::Getf4040ad(void)
{
	return (CString) mp_sf4040ad;
}

void Assignment::odl_Setf4040ad(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf4040ad(newVal);
}

void Assignment::Setf4040ad(CString newVal)
{
	mp_sf4040ad = newVal;
}

BSTR Assignment::odl_Getf4040ae(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf4040ae().AllocSysString();
}

CString Assignment::Getf4040ae(void)
{
	return (CString) mp_sf4040ae;
}

void Assignment::odl_Setf4040ae(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf4040ae(newVal);
}

void Assignment::Setf4040ae(CString newVal)
{
	mp_sf4040ae = newVal;
}

BSTR Assignment::odl_Getf4040af(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf4040af().AllocSysString();
}

CString Assignment::Getf4040af(void)
{
	return (CString) mp_sf4040af;
}

void Assignment::odl_Setf4040af(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf4040af(newVal);
}

void Assignment::Setf4040af(CString newVal)
{
	mp_sf4040af = newVal;
}

BSTR Assignment::odl_Getf4040b0(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf4040b0().AllocSysString();
}

CString Assignment::Getf4040b0(void)
{
	return (CString) mp_sf4040b0;
}

void Assignment::odl_Setf4040b0(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf4040b0(newVal);
}

void Assignment::Setf4040b0(CString newVal)
{
	mp_sf4040b0 = newVal;
}

BSTR Assignment::odl_Getf4040b1(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf4040b1().AllocSysString();
}

CString Assignment::Getf4040b1(void)
{
	return (CString) mp_sf4040b1;
}

void Assignment::odl_Setf4040b1(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf4040b1(newVal);
}

void Assignment::Setf4040b1(CString newVal)
{
	mp_sf4040b1 = newVal;
}

BSTR Assignment::odl_Getf4040b2(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf4040b2().AllocSysString();
}

CString Assignment::Getf4040b2(void)
{
	return (CString) mp_sf4040b2;
}

void Assignment::odl_Setf4040b2(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf4040b2(newVal);
}

void Assignment::Setf4040b2(CString newVal)
{
	mp_sf4040b2 = newVal;
}

BSTR Assignment::odl_Getf4040b3(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf4040b3().AllocSysString();
}

CString Assignment::Getf4040b3(void)
{
	return (CString) mp_sf4040b3;
}

void Assignment::odl_Setf4040b3(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf4040b3(newVal);
}

void Assignment::Setf4040b3(CString newVal)
{
	mp_sf4040b3 = newVal;
}

BSTR Assignment::odl_Getf4040b4(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf4040b4().AllocSysString();
}

CString Assignment::Getf4040b4(void)
{
	return (CString) mp_sf4040b4;
}

void Assignment::odl_Setf4040b4(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf4040b4(newVal);
}

void Assignment::Setf4040b4(CString newVal)
{
	mp_sf4040b4 = newVal;
}

BSTR Assignment::odl_Getf4040b5(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf4040b5().AllocSysString();
}

CString Assignment::Getf4040b5(void)
{
	return (CString) mp_sf4040b5;
}

void Assignment::odl_Setf4040b5(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf4040b5(newVal);
}

void Assignment::Setf4040b5(CString newVal)
{
	mp_sf4040b5 = newVal;
}

BSTR Assignment::odl_Getf4040b6(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf4040b6().AllocSysString();
}

CString Assignment::Getf4040b6(void)
{
	return (CString) mp_sf4040b6;
}

void Assignment::odl_Setf4040b6(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf4040b6(newVal);
}

void Assignment::Setf4040b6(CString newVal)
{
	mp_sf4040b6 = newVal;
}

BSTR Assignment::odl_Getf4040b7(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf4040b7().AllocSysString();
}

CString Assignment::Getf4040b7(void)
{
	return (CString) mp_sf4040b7;
}

void Assignment::odl_Setf4040b7(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf4040b7(newVal);
}

void Assignment::Setf4040b7(CString newVal)
{
	mp_sf4040b7 = newVal;
}

BSTR Assignment::odl_Getf4040b8(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf4040b8().AllocSysString();
}

CString Assignment::Getf4040b8(void)
{
	return (CString) mp_sf4040b8;
}

void Assignment::odl_Setf4040b8(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf4040b8(newVal);
}

void Assignment::Setf4040b8(CString newVal)
{
	mp_sf4040b8 = newVal;
}

BSTR Assignment::odl_Getf4040b9(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf4040b9().AllocSysString();
}

CString Assignment::Getf4040b9(void)
{
	return (CString) mp_sf4040b9;
}

void Assignment::odl_Setf4040b9(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf4040b9(newVal);
}

void Assignment::Setf4040b9(CString newVal)
{
	mp_sf4040b9 = newVal;
}

BSTR Assignment::odl_Getf4040ba(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf4040ba().AllocSysString();
}

CString Assignment::Getf4040ba(void)
{
	return (CString) mp_sf4040ba;
}

void Assignment::odl_Setf4040ba(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf4040ba(newVal);
}

void Assignment::Setf4040ba(CString newVal)
{
	mp_sf4040ba = newVal;
}

BSTR Assignment::odl_Getf4040bb(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf4040bb().AllocSysString();
}

CString Assignment::Getf4040bb(void)
{
	return (CString) mp_sf4040bb;
}

void Assignment::odl_Setf4040bb(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf4040bb(newVal);
}

void Assignment::Setf4040bb(CString newVal)
{
	mp_sf4040bb = newVal;
}

BSTR Assignment::odl_Getf4040bc(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf4040bc().AllocSysString();
}

CString Assignment::Getf4040bc(void)
{
	return (CString) mp_sf4040bc;
}

void Assignment::odl_Setf4040bc(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf4040bc(newVal);
}

void Assignment::Setf4040bc(CString newVal)
{
	mp_sf4040bc = newVal;
}

BSTR Assignment::odl_Getf4040bd(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf4040bd().AllocSysString();
}

CString Assignment::Getf4040bd(void)
{
	return (CString) mp_sf4040bd;
}

void Assignment::odl_Setf4040bd(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf4040bd(newVal);
}

void Assignment::Setf4040bd(CString newVal)
{
	mp_sf4040bd = newVal;
}

BSTR Assignment::odl_Getf4040be(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf4040be().AllocSysString();
}

CString Assignment::Getf4040be(void)
{
	return (CString) mp_sf4040be;
}

void Assignment::odl_Setf4040be(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf4040be(newVal);
}

void Assignment::Setf4040be(CString newVal)
{
	mp_sf4040be = newVal;
}

BSTR Assignment::odl_Getf4040bf(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf4040bf().AllocSysString();
}

CString Assignment::Getf4040bf(void)
{
	return (CString) mp_sf4040bf;
}

void Assignment::odl_Setf4040bf(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf4040bf(newVal);
}

void Assignment::Setf4040bf(CString newVal)
{
	mp_sf4040bf = newVal;
}

BSTR Assignment::odl_Getf4040c0(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf4040c0().AllocSysString();
}

CString Assignment::Getf4040c0(void)
{
	return (CString) mp_sf4040c0;
}

void Assignment::odl_Setf4040c0(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf4040c0(newVal);
}

void Assignment::Setf4040c0(CString newVal)
{
	mp_sf4040c0 = newVal;
}

BSTR Assignment::odl_Getf4040c1(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf4040c1().AllocSysString();
}

CString Assignment::Getf4040c1(void)
{
	return (CString) mp_sf4040c1;
}

void Assignment::odl_Setf4040c1(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf4040c1(newVal);
}

void Assignment::Setf4040c1(CString newVal)
{
	mp_sf4040c1 = newVal;
}

BSTR Assignment::odl_Getf4040c2(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf4040c2().AllocSysString();
}

CString Assignment::Getf4040c2(void)
{
	return (CString) mp_sf4040c2;
}

void Assignment::odl_Setf4040c2(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf4040c2(newVal);
}

void Assignment::Setf4040c2(CString newVal)
{
	mp_sf4040c2 = newVal;
}

BSTR Assignment::odl_Getf4040c3(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf4040c3().AllocSysString();
}

CString Assignment::Getf4040c3(void)
{
	return (CString) mp_sf4040c3;
}

void Assignment::odl_Setf4040c3(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf4040c3(newVal);
}

void Assignment::Setf4040c3(CString newVal)
{
	mp_sf4040c3 = newVal;
}

BSTR Assignment::odl_Getf4040c4(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf4040c4().AllocSysString();
}

CString Assignment::Getf4040c4(void)
{
	return (CString) mp_sf4040c4;
}

void Assignment::odl_Setf4040c4(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf4040c4(newVal);
}

void Assignment::Setf4040c4(CString newVal)
{
	mp_sf4040c4 = newVal;
}

BSTR Assignment::odl_Getf4040c5(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf4040c5().AllocSysString();
}

CString Assignment::Getf4040c5(void)
{
	return (CString) mp_sf4040c5;
}

void Assignment::odl_Setf4040c5(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf4040c5(newVal);
}

void Assignment::Setf4040c5(CString newVal)
{
	mp_sf4040c5 = newVal;
}

BSTR Assignment::odl_Getf4040c6(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf4040c6().AllocSysString();
}

CString Assignment::Getf4040c6(void)
{
	return (CString) mp_sf4040c6;
}

void Assignment::odl_Setf4040c6(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf4040c6(newVal);
}

void Assignment::Setf4040c6(CString newVal)
{
	mp_sf4040c6 = newVal;
}

BSTR Assignment::odl_Getf4040c7(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf4040c7().AllocSysString();
}

CString Assignment::Getf4040c7(void)
{
	return (CString) mp_sf4040c7;
}

void Assignment::odl_Setf4040c7(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf4040c7(newVal);
}

void Assignment::Setf4040c7(CString newVal)
{
	mp_sf4040c7 = newVal;
}

BSTR Assignment::odl_Getf4040c8(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return Getf4040c8().AllocSysString();
}

CString Assignment::Getf4040c8(void)
{
	return (CString) mp_sf4040c8;
}

void Assignment::odl_Setf4040c8(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Setf4040c8(newVal);
}

void Assignment::Setf4040c8(CString newVal)
{
	mp_sf4040c8 = newVal;
}

IDispatch* Assignment::odl_GetTimephasedData(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetTimephasedData()->GetIDispatch(TRUE);
}

TimephasedData_C* Assignment::GetTimephasedData(void)
{
	return mp_oTimephasedData_C;
}

BSTR Assignment::odl_GetKey(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetKey().AllocSysString();
}

void Assignment::odl_SetKey(LPCTSTR newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	USES_CONVERSION;
	SetKey(F_TOUCSTR(T2A(newVal)));
}

CString Assignment::GetKey(void)
{
	return mp_sKey;
}

void Assignment::SetKey(CString newVal)
{
	mp_oCollection->mp_SetKey(&mp_sKey, newVal, MP_SET_KEY);
}

BSTR Assignment::odl_GetXML(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return GetXML().AllocSysString();
}

void Assignment::odl_SetXML(LPCTSTR sXML)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	SetXML(sXML);
}

BOOL Assignment::IsNull(void)
{
	BOOL bReturn = TRUE;
	if (mp_lUID != 0)
	{
		bReturn = FALSE;
	}
	if (mp_lTaskUID != 0)
	{
		bReturn = FALSE;
	}
	if (mp_lResourceUID != 0)
	{
		bReturn = FALSE;
	}
	if (mp_lPercentWorkComplete != 0)
	{
		bReturn = FALSE;
	}
	if (mp_cActualCost != 0)
	{
		bReturn = FALSE;
	}
	if (mp_dtActualFinish.m_dt != 36494)
	{
		bReturn = FALSE;
	}
	if (mp_cActualOvertimeCost != 0)
	{
		bReturn = FALSE;
	}
	if (mp_oActualOvertimeWork->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_dtActualStart.m_dt != 36494)
	{
		bReturn = FALSE;
	}
	if (mp_oActualWork->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_fACWP != 0)
	{
		bReturn = FALSE;
	}
	if (mp_bConfirmed != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_cCost != 0)
	{
		bReturn = FALSE;
	}
	if (mp_yCostRateTable != CRT_COST_RATE_TABLE_0)
	{
		bReturn = FALSE;
	}
	if (mp_fCostVariance != 0)
	{
		bReturn = FALSE;
	}
	if (mp_fCV != 0)
	{
		bReturn = FALSE;
	}
	if (mp_lDelay != 0)
	{
		bReturn = FALSE;
	}
	if (mp_dtFinish.m_dt != 36494)
	{
		bReturn = FALSE;
	}
	if (mp_lFinishVariance != 0)
	{
		bReturn = FALSE;
	}
	if (mp_sHyperlink != "")
	{
		bReturn = FALSE;
	}
	if (mp_sHyperlinkAddress != "")
	{
		bReturn = FALSE;
	}
	if (mp_sHyperlinkSubAddress != "")
	{
		bReturn = FALSE;
	}
	if (mp_fWorkVariance != 0)
	{
		bReturn = FALSE;
	}
	if (mp_bHasFixedRateUnits != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_bFixedMaterial != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_lLevelingDelay != 0)
	{
		bReturn = FALSE;
	}
	if (mp_yLevelingDelayFormat != LDF_M)
	{
		bReturn = FALSE;
	}
	if (mp_bLinkedFields != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_bMilestone != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_sNotes != "")
	{
		bReturn = FALSE;
	}
	if (mp_bOverallocated != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_cOvertimeCost != 0)
	{
		bReturn = FALSE;
	}
	if (mp_oOvertimeWork->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_fPeakUnits != 0)
	{
		bReturn = FALSE;
	}
	if (mp_oRegularWork->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_cRemainingCost != 0)
	{
		bReturn = FALSE;
	}
	if (mp_cRemainingOvertimeCost != 0)
	{
		bReturn = FALSE;
	}
	if (mp_oRemainingOvertimeWork->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_oRemainingWork->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_bResponsePending != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_dtStart.m_dt != 36494)
	{
		bReturn = FALSE;
	}
	if (mp_dtStop.m_dt != 36494)
	{
		bReturn = FALSE;
	}
	if (mp_dtResume.m_dt != 36494)
	{
		bReturn = FALSE;
	}
	if (mp_lStartVariance != 0)
	{
		bReturn = FALSE;
	}
	if (mp_bSummary != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_fSV != 0)
	{
		bReturn = FALSE;
	}
	if (mp_fUnits != 0)
	{
		bReturn = FALSE;
	}
	if (mp_bUpdateNeeded != FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_fVAC != 0)
	{
		bReturn = FALSE;
	}
	if (mp_oWork->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_yWorkContour != WC_FLAT)
	{
		bReturn = FALSE;
	}
	if (mp_fBCWS != 0)
	{
		bReturn = FALSE;
	}
	if (mp_fBCWP != 0)
	{
		bReturn = FALSE;
	}
	if (mp_yBookingType != BT_COMMITED)
	{
		bReturn = FALSE;
	}
	if (mp_oActualWorkProtected->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_oActualOvertimeWorkProtected->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_dtCreationDate.m_dt != 36494)
	{
		bReturn = FALSE;
	}
	if (mp_sAssnOwner != "")
	{
		bReturn = FALSE;
	}
	if (mp_sAssnOwnerGuid != "")
	{
		bReturn = FALSE;
	}
	if (mp_cBudgetCost != 0)
	{
		bReturn = FALSE;
	}
	if (mp_oBudgetWork->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_oExtendedAttribute_C->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_oBaseline_C->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	if (mp_sf404000 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404001 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404002 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404003 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404004 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404005 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404006 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404007 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404008 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404009 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf40400a != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf40400b != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf40400c != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf40400d != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf40400e != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf40400f != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404010 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404011 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404012 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404013 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404014 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404015 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404016 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404017 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404018 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404019 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf40401a != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf40401b != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf40401c != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf40401d != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf40401e != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf40401f != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404020 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404021 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404022 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404023 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404024 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404025 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404026 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404027 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404028 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404029 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf40402a != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf40402b != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf40402c != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf40402d != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf40402e != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf40402f != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404030 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404031 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404032 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404033 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404034 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404035 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404036 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404037 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404038 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404039 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf40403a != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf40403b != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf40403c != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf40403d != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf40403e != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf40403f != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404040 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404041 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404042 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404043 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404044 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404045 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404046 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404047 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404048 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404049 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf40404a != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf40404b != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf40404c != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf40404d != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf40404e != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf40404f != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404050 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404051 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404052 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404053 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404054 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404055 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404056 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404057 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404058 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404059 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf40405a != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf40405b != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf40405c != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf40405d != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf40405e != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf40405f != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404060 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404061 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404062 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404063 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404064 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404065 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404066 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404067 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404068 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404069 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf40406a != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf40406b != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf40406c != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf40406d != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf40406e != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf40406f != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404070 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404071 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404072 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404073 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404074 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404075 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404076 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404077 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404078 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404079 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf40407a != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf40407b != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf40407c != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf40407d != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf40407e != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf40407f != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404080 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404081 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404082 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404083 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404084 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404085 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404086 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404087 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404088 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404089 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf40408a != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf40408b != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf40408c != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf40408d != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf40408e != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf40408f != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404090 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404091 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404092 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404093 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404094 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404095 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404096 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404097 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404098 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf404099 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf40409a != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf40409b != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf40409c != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf40409d != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf40409e != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf40409f != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf4040a0 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf4040a1 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf4040a2 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf4040a3 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf4040a4 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf4040a5 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf4040a6 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf4040a7 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf4040a8 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf4040a9 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf4040aa != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf4040ab != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf4040ac != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf4040ad != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf4040ae != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf4040af != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf4040b0 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf4040b1 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf4040b2 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf4040b3 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf4040b4 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf4040b5 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf4040b6 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf4040b7 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf4040b8 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf4040b9 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf4040ba != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf4040bb != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf4040bc != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf4040bd != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf4040be != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf4040bf != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf4040c0 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf4040c1 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf4040c2 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf4040c3 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf4040c4 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf4040c5 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf4040c6 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf4040c7 != "")
	{
		bReturn = FALSE;
	}
	if (mp_sf4040c8 != "")
	{
		bReturn = FALSE;
	}
	if (mp_oTimephasedData_C->IsNull() == FALSE)
	{
		bReturn = FALSE;
	}
	return bReturn;
}

CString Assignment::GetXML(void)
{
	if (IsNull() == TRUE)
	{
		return "<Assignment/>";
	}
	clsXML oXML("Assignment");
	oXML.InitializeWriter();
	oXML.SetSupportOptional(TRUE);
	oXML.SetBoolsAreNumeric(TRUE);
	oXML.WriteProperty("UID", mp_lUID);
	oXML.WriteProperty("TaskUID", mp_lTaskUID);
	oXML.WriteProperty("ResourceUID", mp_lResourceUID);
	oXML.WriteProperty("PercentWorkComplete", mp_lPercentWorkComplete);
	oXML.WriteProperty("ActualCost", mp_cActualCost);
	if (mp_dtActualFinish.m_dt != 36494)
	{
		oXML.WriteProperty("ActualFinish", mp_dtActualFinish);
	}
	oXML.WriteProperty("ActualOvertimeCost", mp_cActualOvertimeCost);
	oXML.WriteProperty("ActualOvertimeWork", mp_oActualOvertimeWork);
	if (mp_dtActualStart.m_dt != 36494)
	{
		oXML.WriteProperty("ActualStart", mp_dtActualStart);
	}
	oXML.WriteProperty("ActualWork", mp_oActualWork);
	oXML.WriteProperty("ACWP", mp_fACWP);
	oXML.WriteProperty("Confirmed", mp_bConfirmed);
	oXML.WriteProperty("Cost", mp_cCost);
	oXML.WriteProperty("CostRateTable", mp_yCostRateTable);
	oXML.WriteProperty("CostVariance", mp_fCostVariance);
	oXML.WriteProperty("CV", mp_fCV);
	oXML.WriteProperty("Delay", mp_lDelay);
	if (mp_dtFinish.m_dt != 36494)
	{
		oXML.WriteProperty("Finish", mp_dtFinish);
	}
	oXML.WriteProperty("FinishVariance", mp_lFinishVariance);
	if (mp_sHyperlink != "")
	{
		oXML.WriteProperty("Hyperlink", mp_sHyperlink);
	}
	if (mp_sHyperlinkAddress != "")
	{
		oXML.WriteProperty("HyperlinkAddress", mp_sHyperlinkAddress);
	}
	if (mp_sHyperlinkSubAddress != "")
	{
		oXML.WriteProperty("HyperlinkSubAddress", mp_sHyperlinkSubAddress);
	}
	oXML.WriteProperty("WorkVariance", mp_fWorkVariance);
	oXML.WriteProperty("HasFixedRateUnits", mp_bHasFixedRateUnits);
	oXML.WriteProperty("FixedMaterial", mp_bFixedMaterial);
	oXML.WriteProperty("LevelingDelay", mp_lLevelingDelay);
	oXML.WriteProperty("LevelingDelayFormat", mp_yLevelingDelayFormat);
	oXML.WriteProperty("LinkedFields", mp_bLinkedFields);
	oXML.WriteProperty("Milestone", mp_bMilestone);
	if (mp_sNotes != "")
	{
		oXML.WriteProperty("Notes", mp_sNotes);
	}
	oXML.WriteProperty("Overallocated", mp_bOverallocated);
	oXML.WriteProperty("OvertimeCost", mp_cOvertimeCost);
	oXML.WriteProperty("OvertimeWork", mp_oOvertimeWork);
	oXML.WriteProperty("PeakUnits", mp_fPeakUnits);
	oXML.WriteProperty("RegularWork", mp_oRegularWork);
	oXML.WriteProperty("RemainingCost", mp_cRemainingCost);
	oXML.WriteProperty("RemainingOvertimeCost", mp_cRemainingOvertimeCost);
	oXML.WriteProperty("RemainingOvertimeWork", mp_oRemainingOvertimeWork);
	oXML.WriteProperty("RemainingWork", mp_oRemainingWork);
	oXML.WriteProperty("ResponsePending", mp_bResponsePending);
	if (mp_dtStart.m_dt != 36494)
	{
		oXML.WriteProperty("Start", mp_dtStart);
	}
	if (mp_dtStop.m_dt != 36494)
	{
		oXML.WriteProperty("Stop", mp_dtStop);
	}
	if (mp_dtResume.m_dt != 36494)
	{
		oXML.WriteProperty("Resume", mp_dtResume);
	}
	oXML.WriteProperty("StartVariance", mp_lStartVariance);
	oXML.WriteProperty("Summary", mp_bSummary);
	oXML.WriteProperty("SV", mp_fSV);
	oXML.WriteProperty("Units", mp_fUnits);
	oXML.WriteProperty("UpdateNeeded", mp_bUpdateNeeded);
	oXML.WriteProperty("VAC", mp_fVAC);
	oXML.WriteProperty("Work", mp_oWork);
	oXML.WriteProperty("WorkContour", mp_yWorkContour);
	oXML.WriteProperty("BCWS", mp_fBCWS);
	oXML.WriteProperty("BCWP", mp_fBCWP);
	oXML.WriteProperty("BookingType", mp_yBookingType);
	if (mp_oActualWorkProtected->IsNull() == FALSE)
	{
		oXML.WriteProperty("ActualWorkProtected", mp_oActualWorkProtected);
	}
	if (mp_oActualOvertimeWorkProtected->IsNull() == FALSE)
	{
		oXML.WriteProperty("ActualOvertimeWorkProtected", mp_oActualOvertimeWorkProtected);
	}
	if (mp_dtCreationDate.m_dt != 36494)
	{
		oXML.WriteProperty("CreationDate", mp_dtCreationDate);
	}
	if (mp_sAssnOwner != "")
	{
		oXML.WriteProperty("AssnOwner", mp_sAssnOwner);
	}
	if (mp_sAssnOwnerGuid != "")
	{
		oXML.WriteProperty("AssnOwnerGuid", mp_sAssnOwnerGuid);
	}
	oXML.WriteProperty("BudgetCost", mp_cBudgetCost);
	if (mp_oBudgetWork->IsNull() == FALSE)
	{
		oXML.WriteProperty("BudgetWork", mp_oBudgetWork);
	}
	if (mp_oExtendedAttribute_C->IsNull() == FALSE)
	{
		mp_oExtendedAttribute_C->WriteObjectProtected(oXML);
	}
	if (mp_oBaseline_C->IsNull() == FALSE)
	{
		mp_oBaseline_C->WriteObjectProtected(oXML);
	}
	if (mp_sf404000 != "")
	{
		oXML.WriteProperty("f404000", mp_sf404000);
	}
	if (mp_sf404001 != "")
	{
		oXML.WriteProperty("f404001", mp_sf404001);
	}
	if (mp_sf404002 != "")
	{
		oXML.WriteProperty("f404002", mp_sf404002);
	}
	if (mp_sf404003 != "")
	{
		oXML.WriteProperty("f404003", mp_sf404003);
	}
	if (mp_sf404004 != "")
	{
		oXML.WriteProperty("f404004", mp_sf404004);
	}
	if (mp_sf404005 != "")
	{
		oXML.WriteProperty("f404005", mp_sf404005);
	}
	if (mp_sf404006 != "")
	{
		oXML.WriteProperty("f404006", mp_sf404006);
	}
	if (mp_sf404007 != "")
	{
		oXML.WriteProperty("f404007", mp_sf404007);
	}
	if (mp_sf404008 != "")
	{
		oXML.WriteProperty("f404008", mp_sf404008);
	}
	if (mp_sf404009 != "")
	{
		oXML.WriteProperty("f404009", mp_sf404009);
	}
	if (mp_sf40400a != "")
	{
		oXML.WriteProperty("f40400a", mp_sf40400a);
	}
	if (mp_sf40400b != "")
	{
		oXML.WriteProperty("f40400b", mp_sf40400b);
	}
	if (mp_sf40400c != "")
	{
		oXML.WriteProperty("f40400c", mp_sf40400c);
	}
	if (mp_sf40400d != "")
	{
		oXML.WriteProperty("f40400d", mp_sf40400d);
	}
	if (mp_sf40400e != "")
	{
		oXML.WriteProperty("f40400e", mp_sf40400e);
	}
	if (mp_sf40400f != "")
	{
		oXML.WriteProperty("f40400f", mp_sf40400f);
	}
	if (mp_sf404010 != "")
	{
		oXML.WriteProperty("f404010", mp_sf404010);
	}
	if (mp_sf404011 != "")
	{
		oXML.WriteProperty("f404011", mp_sf404011);
	}
	if (mp_sf404012 != "")
	{
		oXML.WriteProperty("f404012", mp_sf404012);
	}
	if (mp_sf404013 != "")
	{
		oXML.WriteProperty("f404013", mp_sf404013);
	}
	if (mp_sf404014 != "")
	{
		oXML.WriteProperty("f404014", mp_sf404014);
	}
	if (mp_sf404015 != "")
	{
		oXML.WriteProperty("f404015", mp_sf404015);
	}
	if (mp_sf404016 != "")
	{
		oXML.WriteProperty("f404016", mp_sf404016);
	}
	if (mp_sf404017 != "")
	{
		oXML.WriteProperty("f404017", mp_sf404017);
	}
	if (mp_sf404018 != "")
	{
		oXML.WriteProperty("f404018", mp_sf404018);
	}
	if (mp_sf404019 != "")
	{
		oXML.WriteProperty("f404019", mp_sf404019);
	}
	if (mp_sf40401a != "")
	{
		oXML.WriteProperty("f40401a", mp_sf40401a);
	}
	if (mp_sf40401b != "")
	{
		oXML.WriteProperty("f40401b", mp_sf40401b);
	}
	if (mp_sf40401c != "")
	{
		oXML.WriteProperty("f40401c", mp_sf40401c);
	}
	if (mp_sf40401d != "")
	{
		oXML.WriteProperty("f40401d", mp_sf40401d);
	}
	if (mp_sf40401e != "")
	{
		oXML.WriteProperty("f40401e", mp_sf40401e);
	}
	if (mp_sf40401f != "")
	{
		oXML.WriteProperty("f40401f", mp_sf40401f);
	}
	if (mp_sf404020 != "")
	{
		oXML.WriteProperty("f404020", mp_sf404020);
	}
	if (mp_sf404021 != "")
	{
		oXML.WriteProperty("f404021", mp_sf404021);
	}
	if (mp_sf404022 != "")
	{
		oXML.WriteProperty("f404022", mp_sf404022);
	}
	if (mp_sf404023 != "")
	{
		oXML.WriteProperty("f404023", mp_sf404023);
	}
	if (mp_sf404024 != "")
	{
		oXML.WriteProperty("f404024", mp_sf404024);
	}
	if (mp_sf404025 != "")
	{
		oXML.WriteProperty("f404025", mp_sf404025);
	}
	if (mp_sf404026 != "")
	{
		oXML.WriteProperty("f404026", mp_sf404026);
	}
	if (mp_sf404027 != "")
	{
		oXML.WriteProperty("f404027", mp_sf404027);
	}
	if (mp_sf404028 != "")
	{
		oXML.WriteProperty("f404028", mp_sf404028);
	}
	if (mp_sf404029 != "")
	{
		oXML.WriteProperty("f404029", mp_sf404029);
	}
	if (mp_sf40402a != "")
	{
		oXML.WriteProperty("f40402a", mp_sf40402a);
	}
	if (mp_sf40402b != "")
	{
		oXML.WriteProperty("f40402b", mp_sf40402b);
	}
	if (mp_sf40402c != "")
	{
		oXML.WriteProperty("f40402c", mp_sf40402c);
	}
	if (mp_sf40402d != "")
	{
		oXML.WriteProperty("f40402d", mp_sf40402d);
	}
	if (mp_sf40402e != "")
	{
		oXML.WriteProperty("f40402e", mp_sf40402e);
	}
	if (mp_sf40402f != "")
	{
		oXML.WriteProperty("f40402f", mp_sf40402f);
	}
	if (mp_sf404030 != "")
	{
		oXML.WriteProperty("f404030", mp_sf404030);
	}
	if (mp_sf404031 != "")
	{
		oXML.WriteProperty("f404031", mp_sf404031);
	}
	if (mp_sf404032 != "")
	{
		oXML.WriteProperty("f404032", mp_sf404032);
	}
	if (mp_sf404033 != "")
	{
		oXML.WriteProperty("f404033", mp_sf404033);
	}
	if (mp_sf404034 != "")
	{
		oXML.WriteProperty("f404034", mp_sf404034);
	}
	if (mp_sf404035 != "")
	{
		oXML.WriteProperty("f404035", mp_sf404035);
	}
	if (mp_sf404036 != "")
	{
		oXML.WriteProperty("f404036", mp_sf404036);
	}
	if (mp_sf404037 != "")
	{
		oXML.WriteProperty("f404037", mp_sf404037);
	}
	if (mp_sf404038 != "")
	{
		oXML.WriteProperty("f404038", mp_sf404038);
	}
	if (mp_sf404039 != "")
	{
		oXML.WriteProperty("f404039", mp_sf404039);
	}
	if (mp_sf40403a != "")
	{
		oXML.WriteProperty("f40403a", mp_sf40403a);
	}
	if (mp_sf40403b != "")
	{
		oXML.WriteProperty("f40403b", mp_sf40403b);
	}
	if (mp_sf40403c != "")
	{
		oXML.WriteProperty("f40403c", mp_sf40403c);
	}
	if (mp_sf40403d != "")
	{
		oXML.WriteProperty("f40403d", mp_sf40403d);
	}
	if (mp_sf40403e != "")
	{
		oXML.WriteProperty("f40403e", mp_sf40403e);
	}
	if (mp_sf40403f != "")
	{
		oXML.WriteProperty("f40403f", mp_sf40403f);
	}
	if (mp_sf404040 != "")
	{
		oXML.WriteProperty("f404040", mp_sf404040);
	}
	if (mp_sf404041 != "")
	{
		oXML.WriteProperty("f404041", mp_sf404041);
	}
	if (mp_sf404042 != "")
	{
		oXML.WriteProperty("f404042", mp_sf404042);
	}
	if (mp_sf404043 != "")
	{
		oXML.WriteProperty("f404043", mp_sf404043);
	}
	if (mp_sf404044 != "")
	{
		oXML.WriteProperty("f404044", mp_sf404044);
	}
	if (mp_sf404045 != "")
	{
		oXML.WriteProperty("f404045", mp_sf404045);
	}
	if (mp_sf404046 != "")
	{
		oXML.WriteProperty("f404046", mp_sf404046);
	}
	if (mp_sf404047 != "")
	{
		oXML.WriteProperty("f404047", mp_sf404047);
	}
	if (mp_sf404048 != "")
	{
		oXML.WriteProperty("f404048", mp_sf404048);
	}
	if (mp_sf404049 != "")
	{
		oXML.WriteProperty("f404049", mp_sf404049);
	}
	if (mp_sf40404a != "")
	{
		oXML.WriteProperty("f40404a", mp_sf40404a);
	}
	if (mp_sf40404b != "")
	{
		oXML.WriteProperty("f40404b", mp_sf40404b);
	}
	if (mp_sf40404c != "")
	{
		oXML.WriteProperty("f40404c", mp_sf40404c);
	}
	if (mp_sf40404d != "")
	{
		oXML.WriteProperty("f40404d", mp_sf40404d);
	}
	if (mp_sf40404e != "")
	{
		oXML.WriteProperty("f40404e", mp_sf40404e);
	}
	if (mp_sf40404f != "")
	{
		oXML.WriteProperty("f40404f", mp_sf40404f);
	}
	if (mp_sf404050 != "")
	{
		oXML.WriteProperty("f404050", mp_sf404050);
	}
	if (mp_sf404051 != "")
	{
		oXML.WriteProperty("f404051", mp_sf404051);
	}
	if (mp_sf404052 != "")
	{
		oXML.WriteProperty("f404052", mp_sf404052);
	}
	if (mp_sf404053 != "")
	{
		oXML.WriteProperty("f404053", mp_sf404053);
	}
	if (mp_sf404054 != "")
	{
		oXML.WriteProperty("f404054", mp_sf404054);
	}
	if (mp_sf404055 != "")
	{
		oXML.WriteProperty("f404055", mp_sf404055);
	}
	if (mp_sf404056 != "")
	{
		oXML.WriteProperty("f404056", mp_sf404056);
	}
	if (mp_sf404057 != "")
	{
		oXML.WriteProperty("f404057", mp_sf404057);
	}
	if (mp_sf404058 != "")
	{
		oXML.WriteProperty("f404058", mp_sf404058);
	}
	if (mp_sf404059 != "")
	{
		oXML.WriteProperty("f404059", mp_sf404059);
	}
	if (mp_sf40405a != "")
	{
		oXML.WriteProperty("f40405a", mp_sf40405a);
	}
	if (mp_sf40405b != "")
	{
		oXML.WriteProperty("f40405b", mp_sf40405b);
	}
	if (mp_sf40405c != "")
	{
		oXML.WriteProperty("f40405c", mp_sf40405c);
	}
	if (mp_sf40405d != "")
	{
		oXML.WriteProperty("f40405d", mp_sf40405d);
	}
	if (mp_sf40405e != "")
	{
		oXML.WriteProperty("f40405e", mp_sf40405e);
	}
	if (mp_sf40405f != "")
	{
		oXML.WriteProperty("f40405f", mp_sf40405f);
	}
	if (mp_sf404060 != "")
	{
		oXML.WriteProperty("f404060", mp_sf404060);
	}
	if (mp_sf404061 != "")
	{
		oXML.WriteProperty("f404061", mp_sf404061);
	}
	if (mp_sf404062 != "")
	{
		oXML.WriteProperty("f404062", mp_sf404062);
	}
	if (mp_sf404063 != "")
	{
		oXML.WriteProperty("f404063", mp_sf404063);
	}
	if (mp_sf404064 != "")
	{
		oXML.WriteProperty("f404064", mp_sf404064);
	}
	if (mp_sf404065 != "")
	{
		oXML.WriteProperty("f404065", mp_sf404065);
	}
	if (mp_sf404066 != "")
	{
		oXML.WriteProperty("f404066", mp_sf404066);
	}
	if (mp_sf404067 != "")
	{
		oXML.WriteProperty("f404067", mp_sf404067);
	}
	if (mp_sf404068 != "")
	{
		oXML.WriteProperty("f404068", mp_sf404068);
	}
	if (mp_sf404069 != "")
	{
		oXML.WriteProperty("f404069", mp_sf404069);
	}
	if (mp_sf40406a != "")
	{
		oXML.WriteProperty("f40406a", mp_sf40406a);
	}
	if (mp_sf40406b != "")
	{
		oXML.WriteProperty("f40406b", mp_sf40406b);
	}
	if (mp_sf40406c != "")
	{
		oXML.WriteProperty("f40406c", mp_sf40406c);
	}
	if (mp_sf40406d != "")
	{
		oXML.WriteProperty("f40406d", mp_sf40406d);
	}
	if (mp_sf40406e != "")
	{
		oXML.WriteProperty("f40406e", mp_sf40406e);
	}
	if (mp_sf40406f != "")
	{
		oXML.WriteProperty("f40406f", mp_sf40406f);
	}
	if (mp_sf404070 != "")
	{
		oXML.WriteProperty("f404070", mp_sf404070);
	}
	if (mp_sf404071 != "")
	{
		oXML.WriteProperty("f404071", mp_sf404071);
	}
	if (mp_sf404072 != "")
	{
		oXML.WriteProperty("f404072", mp_sf404072);
	}
	if (mp_sf404073 != "")
	{
		oXML.WriteProperty("f404073", mp_sf404073);
	}
	if (mp_sf404074 != "")
	{
		oXML.WriteProperty("f404074", mp_sf404074);
	}
	if (mp_sf404075 != "")
	{
		oXML.WriteProperty("f404075", mp_sf404075);
	}
	if (mp_sf404076 != "")
	{
		oXML.WriteProperty("f404076", mp_sf404076);
	}
	if (mp_sf404077 != "")
	{
		oXML.WriteProperty("f404077", mp_sf404077);
	}
	if (mp_sf404078 != "")
	{
		oXML.WriteProperty("f404078", mp_sf404078);
	}
	if (mp_sf404079 != "")
	{
		oXML.WriteProperty("f404079", mp_sf404079);
	}
	if (mp_sf40407a != "")
	{
		oXML.WriteProperty("f40407a", mp_sf40407a);
	}
	if (mp_sf40407b != "")
	{
		oXML.WriteProperty("f40407b", mp_sf40407b);
	}
	if (mp_sf40407c != "")
	{
		oXML.WriteProperty("f40407c", mp_sf40407c);
	}
	if (mp_sf40407d != "")
	{
		oXML.WriteProperty("f40407d", mp_sf40407d);
	}
	if (mp_sf40407e != "")
	{
		oXML.WriteProperty("f40407e", mp_sf40407e);
	}
	if (mp_sf40407f != "")
	{
		oXML.WriteProperty("f40407f", mp_sf40407f);
	}
	if (mp_sf404080 != "")
	{
		oXML.WriteProperty("f404080", mp_sf404080);
	}
	if (mp_sf404081 != "")
	{
		oXML.WriteProperty("f404081", mp_sf404081);
	}
	if (mp_sf404082 != "")
	{
		oXML.WriteProperty("f404082", mp_sf404082);
	}
	if (mp_sf404083 != "")
	{
		oXML.WriteProperty("f404083", mp_sf404083);
	}
	if (mp_sf404084 != "")
	{
		oXML.WriteProperty("f404084", mp_sf404084);
	}
	if (mp_sf404085 != "")
	{
		oXML.WriteProperty("f404085", mp_sf404085);
	}
	if (mp_sf404086 != "")
	{
		oXML.WriteProperty("f404086", mp_sf404086);
	}
	if (mp_sf404087 != "")
	{
		oXML.WriteProperty("f404087", mp_sf404087);
	}
	if (mp_sf404088 != "")
	{
		oXML.WriteProperty("f404088", mp_sf404088);
	}
	if (mp_sf404089 != "")
	{
		oXML.WriteProperty("f404089", mp_sf404089);
	}
	if (mp_sf40408a != "")
	{
		oXML.WriteProperty("f40408a", mp_sf40408a);
	}
	if (mp_sf40408b != "")
	{
		oXML.WriteProperty("f40408b", mp_sf40408b);
	}
	if (mp_sf40408c != "")
	{
		oXML.WriteProperty("f40408c", mp_sf40408c);
	}
	if (mp_sf40408d != "")
	{
		oXML.WriteProperty("f40408d", mp_sf40408d);
	}
	if (mp_sf40408e != "")
	{
		oXML.WriteProperty("f40408e", mp_sf40408e);
	}
	if (mp_sf40408f != "")
	{
		oXML.WriteProperty("f40408f", mp_sf40408f);
	}
	if (mp_sf404090 != "")
	{
		oXML.WriteProperty("f404090", mp_sf404090);
	}
	if (mp_sf404091 != "")
	{
		oXML.WriteProperty("f404091", mp_sf404091);
	}
	if (mp_sf404092 != "")
	{
		oXML.WriteProperty("f404092", mp_sf404092);
	}
	if (mp_sf404093 != "")
	{
		oXML.WriteProperty("f404093", mp_sf404093);
	}
	if (mp_sf404094 != "")
	{
		oXML.WriteProperty("f404094", mp_sf404094);
	}
	if (mp_sf404095 != "")
	{
		oXML.WriteProperty("f404095", mp_sf404095);
	}
	if (mp_sf404096 != "")
	{
		oXML.WriteProperty("f404096", mp_sf404096);
	}
	if (mp_sf404097 != "")
	{
		oXML.WriteProperty("f404097", mp_sf404097);
	}
	if (mp_sf404098 != "")
	{
		oXML.WriteProperty("f404098", mp_sf404098);
	}
	if (mp_sf404099 != "")
	{
		oXML.WriteProperty("f404099", mp_sf404099);
	}
	if (mp_sf40409a != "")
	{
		oXML.WriteProperty("f40409a", mp_sf40409a);
	}
	if (mp_sf40409b != "")
	{
		oXML.WriteProperty("f40409b", mp_sf40409b);
	}
	if (mp_sf40409c != "")
	{
		oXML.WriteProperty("f40409c", mp_sf40409c);
	}
	if (mp_sf40409d != "")
	{
		oXML.WriteProperty("f40409d", mp_sf40409d);
	}
	if (mp_sf40409e != "")
	{
		oXML.WriteProperty("f40409e", mp_sf40409e);
	}
	if (mp_sf40409f != "")
	{
		oXML.WriteProperty("f40409f", mp_sf40409f);
	}
	if (mp_sf4040a0 != "")
	{
		oXML.WriteProperty("f4040a0", mp_sf4040a0);
	}
	if (mp_sf4040a1 != "")
	{
		oXML.WriteProperty("f4040a1", mp_sf4040a1);
	}
	if (mp_sf4040a2 != "")
	{
		oXML.WriteProperty("f4040a2", mp_sf4040a2);
	}
	if (mp_sf4040a3 != "")
	{
		oXML.WriteProperty("f4040a3", mp_sf4040a3);
	}
	if (mp_sf4040a4 != "")
	{
		oXML.WriteProperty("f4040a4", mp_sf4040a4);
	}
	if (mp_sf4040a5 != "")
	{
		oXML.WriteProperty("f4040a5", mp_sf4040a5);
	}
	if (mp_sf4040a6 != "")
	{
		oXML.WriteProperty("f4040a6", mp_sf4040a6);
	}
	if (mp_sf4040a7 != "")
	{
		oXML.WriteProperty("f4040a7", mp_sf4040a7);
	}
	if (mp_sf4040a8 != "")
	{
		oXML.WriteProperty("f4040a8", mp_sf4040a8);
	}
	if (mp_sf4040a9 != "")
	{
		oXML.WriteProperty("f4040a9", mp_sf4040a9);
	}
	if (mp_sf4040aa != "")
	{
		oXML.WriteProperty("f4040aa", mp_sf4040aa);
	}
	if (mp_sf4040ab != "")
	{
		oXML.WriteProperty("f4040ab", mp_sf4040ab);
	}
	if (mp_sf4040ac != "")
	{
		oXML.WriteProperty("f4040ac", mp_sf4040ac);
	}
	if (mp_sf4040ad != "")
	{
		oXML.WriteProperty("f4040ad", mp_sf4040ad);
	}
	if (mp_sf4040ae != "")
	{
		oXML.WriteProperty("f4040ae", mp_sf4040ae);
	}
	if (mp_sf4040af != "")
	{
		oXML.WriteProperty("f4040af", mp_sf4040af);
	}
	if (mp_sf4040b0 != "")
	{
		oXML.WriteProperty("f4040b0", mp_sf4040b0);
	}
	if (mp_sf4040b1 != "")
	{
		oXML.WriteProperty("f4040b1", mp_sf4040b1);
	}
	if (mp_sf4040b2 != "")
	{
		oXML.WriteProperty("f4040b2", mp_sf4040b2);
	}
	if (mp_sf4040b3 != "")
	{
		oXML.WriteProperty("f4040b3", mp_sf4040b3);
	}
	if (mp_sf4040b4 != "")
	{
		oXML.WriteProperty("f4040b4", mp_sf4040b4);
	}
	if (mp_sf4040b5 != "")
	{
		oXML.WriteProperty("f4040b5", mp_sf4040b5);
	}
	if (mp_sf4040b6 != "")
	{
		oXML.WriteProperty("f4040b6", mp_sf4040b6);
	}
	if (mp_sf4040b7 != "")
	{
		oXML.WriteProperty("f4040b7", mp_sf4040b7);
	}
	if (mp_sf4040b8 != "")
	{
		oXML.WriteProperty("f4040b8", mp_sf4040b8);
	}
	if (mp_sf4040b9 != "")
	{
		oXML.WriteProperty("f4040b9", mp_sf4040b9);
	}
	if (mp_sf4040ba != "")
	{
		oXML.WriteProperty("f4040ba", mp_sf4040ba);
	}
	if (mp_sf4040bb != "")
	{
		oXML.WriteProperty("f4040bb", mp_sf4040bb);
	}
	if (mp_sf4040bc != "")
	{
		oXML.WriteProperty("f4040bc", mp_sf4040bc);
	}
	if (mp_sf4040bd != "")
	{
		oXML.WriteProperty("f4040bd", mp_sf4040bd);
	}
	if (mp_sf4040be != "")
	{
		oXML.WriteProperty("f4040be", mp_sf4040be);
	}
	if (mp_sf4040bf != "")
	{
		oXML.WriteProperty("f4040bf", mp_sf4040bf);
	}
	if (mp_sf4040c0 != "")
	{
		oXML.WriteProperty("f4040c0", mp_sf4040c0);
	}
	if (mp_sf4040c1 != "")
	{
		oXML.WriteProperty("f4040c1", mp_sf4040c1);
	}
	if (mp_sf4040c2 != "")
	{
		oXML.WriteProperty("f4040c2", mp_sf4040c2);
	}
	if (mp_sf4040c3 != "")
	{
		oXML.WriteProperty("f4040c3", mp_sf4040c3);
	}
	if (mp_sf4040c4 != "")
	{
		oXML.WriteProperty("f4040c4", mp_sf4040c4);
	}
	if (mp_sf4040c5 != "")
	{
		oXML.WriteProperty("f4040c5", mp_sf4040c5);
	}
	if (mp_sf4040c6 != "")
	{
		oXML.WriteProperty("f4040c6", mp_sf4040c6);
	}
	if (mp_sf4040c7 != "")
	{
		oXML.WriteProperty("f4040c7", mp_sf4040c7);
	}
	if (mp_sf4040c8 != "")
	{
		oXML.WriteProperty("f4040c8", mp_sf4040c8);
	}
	if (mp_oTimephasedData_C->IsNull() == FALSE)
	{
		mp_oTimephasedData_C->WriteObjectProtected(oXML);
	}
	return oXML.GetXML();
}

void Assignment::SetXML(CString sXML)
{
	clsXML oXML("Assignment");
	oXML.SetSupportOptional(TRUE);
	oXML.SetXML(sXML);
	oXML.InitializeReader();
	oXML.ReadProperty("UID", mp_lUID);
	oXML.ReadProperty("TaskUID", mp_lTaskUID);
	oXML.ReadProperty("ResourceUID", mp_lResourceUID);
	oXML.ReadProperty("PercentWorkComplete", mp_lPercentWorkComplete);
	oXML.ReadProperty("ActualCost", mp_cActualCost);
	oXML.ReadProperty("ActualFinish", mp_dtActualFinish);
	oXML.ReadProperty("ActualOvertimeCost", mp_cActualOvertimeCost);
	oXML.ReadProperty("ActualOvertimeWork", mp_oActualOvertimeWork);
	oXML.ReadProperty("ActualStart", mp_dtActualStart);
	oXML.ReadProperty("ActualWork", mp_oActualWork);
	oXML.ReadProperty("ACWP", mp_fACWP);
	oXML.ReadProperty("Confirmed", mp_bConfirmed);
	oXML.ReadProperty("Cost", mp_cCost);
	oXML.ReadProperty("CostRateTable", mp_yCostRateTable);
	oXML.ReadProperty("CostVariance", mp_fCostVariance);
	oXML.ReadProperty("CV", mp_fCV);
	oXML.ReadProperty("Delay", mp_lDelay);
	oXML.ReadProperty("Finish", mp_dtFinish);
	oXML.ReadProperty("FinishVariance", mp_lFinishVariance);
	oXML.ReadProperty("Hyperlink", mp_sHyperlink);
	if (mp_sHyperlink.GetLength() > 512)
	{
		mp_sHyperlink = mp_sHyperlink.Mid(0, 512);
	}
	oXML.ReadProperty("HyperlinkAddress", mp_sHyperlinkAddress);
	if (mp_sHyperlinkAddress.GetLength() > 512)
	{
		mp_sHyperlinkAddress = mp_sHyperlinkAddress.Mid(0, 512);
	}
	oXML.ReadProperty("HyperlinkSubAddress", mp_sHyperlinkSubAddress);
	if (mp_sHyperlinkSubAddress.GetLength() > 512)
	{
		mp_sHyperlinkSubAddress = mp_sHyperlinkSubAddress.Mid(0, 512);
	}
	oXML.ReadProperty("WorkVariance", mp_fWorkVariance);
	oXML.ReadProperty("HasFixedRateUnits", mp_bHasFixedRateUnits);
	oXML.ReadProperty("FixedMaterial", mp_bFixedMaterial);
	oXML.ReadProperty("LevelingDelay", mp_lLevelingDelay);
	oXML.ReadProperty("LevelingDelayFormat", mp_yLevelingDelayFormat);
	oXML.ReadProperty("LinkedFields", mp_bLinkedFields);
	oXML.ReadProperty("Milestone", mp_bMilestone);
	oXML.ReadProperty("Notes", mp_sNotes);
	oXML.ReadProperty("Overallocated", mp_bOverallocated);
	oXML.ReadProperty("OvertimeCost", mp_cOvertimeCost);
	oXML.ReadProperty("OvertimeWork", mp_oOvertimeWork);
	oXML.ReadProperty("PeakUnits", mp_fPeakUnits);
	oXML.ReadProperty("RegularWork", mp_oRegularWork);
	oXML.ReadProperty("RemainingCost", mp_cRemainingCost);
	oXML.ReadProperty("RemainingOvertimeCost", mp_cRemainingOvertimeCost);
	oXML.ReadProperty("RemainingOvertimeWork", mp_oRemainingOvertimeWork);
	oXML.ReadProperty("RemainingWork", mp_oRemainingWork);
	oXML.ReadProperty("ResponsePending", mp_bResponsePending);
	oXML.ReadProperty("Start", mp_dtStart);
	oXML.ReadProperty("Stop", mp_dtStop);
	oXML.ReadProperty("Resume", mp_dtResume);
	oXML.ReadProperty("StartVariance", mp_lStartVariance);
	oXML.ReadProperty("Summary", mp_bSummary);
	oXML.ReadProperty("SV", mp_fSV);
	oXML.ReadProperty("Units", mp_fUnits);
	oXML.ReadProperty("UpdateNeeded", mp_bUpdateNeeded);
	oXML.ReadProperty("VAC", mp_fVAC);
	oXML.ReadProperty("Work", mp_oWork);
	oXML.ReadProperty("WorkContour", mp_yWorkContour);
	oXML.ReadProperty("BCWS", mp_fBCWS);
	oXML.ReadProperty("BCWP", mp_fBCWP);
	oXML.ReadProperty("BookingType", mp_yBookingType);
	oXML.ReadProperty("ActualWorkProtected", mp_oActualWorkProtected);
	oXML.ReadProperty("ActualOvertimeWorkProtected", mp_oActualOvertimeWorkProtected);
	oXML.ReadProperty("CreationDate", mp_dtCreationDate);
	oXML.ReadProperty("AssnOwner", mp_sAssnOwner);
	oXML.ReadProperty("AssnOwnerGuid", mp_sAssnOwnerGuid);
	oXML.ReadProperty("BudgetCost", mp_cBudgetCost);
	oXML.ReadProperty("BudgetWork", mp_oBudgetWork);
	mp_oExtendedAttribute_C->ReadObjectProtected(oXML);
	mp_oBaseline_C->ReadObjectProtected(oXML);
	oXML.ReadProperty("f404000", mp_sf404000);
	oXML.ReadProperty("f404001", mp_sf404001);
	oXML.ReadProperty("f404002", mp_sf404002);
	oXML.ReadProperty("f404003", mp_sf404003);
	oXML.ReadProperty("f404004", mp_sf404004);
	oXML.ReadProperty("f404005", mp_sf404005);
	oXML.ReadProperty("f404006", mp_sf404006);
	oXML.ReadProperty("f404007", mp_sf404007);
	oXML.ReadProperty("f404008", mp_sf404008);
	oXML.ReadProperty("f404009", mp_sf404009);
	oXML.ReadProperty("f40400a", mp_sf40400a);
	oXML.ReadProperty("f40400b", mp_sf40400b);
	oXML.ReadProperty("f40400c", mp_sf40400c);
	oXML.ReadProperty("f40400d", mp_sf40400d);
	oXML.ReadProperty("f40400e", mp_sf40400e);
	oXML.ReadProperty("f40400f", mp_sf40400f);
	oXML.ReadProperty("f404010", mp_sf404010);
	oXML.ReadProperty("f404011", mp_sf404011);
	oXML.ReadProperty("f404012", mp_sf404012);
	oXML.ReadProperty("f404013", mp_sf404013);
	oXML.ReadProperty("f404014", mp_sf404014);
	oXML.ReadProperty("f404015", mp_sf404015);
	oXML.ReadProperty("f404016", mp_sf404016);
	oXML.ReadProperty("f404017", mp_sf404017);
	oXML.ReadProperty("f404018", mp_sf404018);
	oXML.ReadProperty("f404019", mp_sf404019);
	oXML.ReadProperty("f40401a", mp_sf40401a);
	oXML.ReadProperty("f40401b", mp_sf40401b);
	oXML.ReadProperty("f40401c", mp_sf40401c);
	oXML.ReadProperty("f40401d", mp_sf40401d);
	oXML.ReadProperty("f40401e", mp_sf40401e);
	oXML.ReadProperty("f40401f", mp_sf40401f);
	oXML.ReadProperty("f404020", mp_sf404020);
	oXML.ReadProperty("f404021", mp_sf404021);
	oXML.ReadProperty("f404022", mp_sf404022);
	oXML.ReadProperty("f404023", mp_sf404023);
	oXML.ReadProperty("f404024", mp_sf404024);
	oXML.ReadProperty("f404025", mp_sf404025);
	oXML.ReadProperty("f404026", mp_sf404026);
	oXML.ReadProperty("f404027", mp_sf404027);
	oXML.ReadProperty("f404028", mp_sf404028);
	oXML.ReadProperty("f404029", mp_sf404029);
	oXML.ReadProperty("f40402a", mp_sf40402a);
	oXML.ReadProperty("f40402b", mp_sf40402b);
	oXML.ReadProperty("f40402c", mp_sf40402c);
	oXML.ReadProperty("f40402d", mp_sf40402d);
	oXML.ReadProperty("f40402e", mp_sf40402e);
	oXML.ReadProperty("f40402f", mp_sf40402f);
	oXML.ReadProperty("f404030", mp_sf404030);
	oXML.ReadProperty("f404031", mp_sf404031);
	oXML.ReadProperty("f404032", mp_sf404032);
	oXML.ReadProperty("f404033", mp_sf404033);
	oXML.ReadProperty("f404034", mp_sf404034);
	oXML.ReadProperty("f404035", mp_sf404035);
	oXML.ReadProperty("f404036", mp_sf404036);
	oXML.ReadProperty("f404037", mp_sf404037);
	oXML.ReadProperty("f404038", mp_sf404038);
	oXML.ReadProperty("f404039", mp_sf404039);
	oXML.ReadProperty("f40403a", mp_sf40403a);
	oXML.ReadProperty("f40403b", mp_sf40403b);
	oXML.ReadProperty("f40403c", mp_sf40403c);
	oXML.ReadProperty("f40403d", mp_sf40403d);
	oXML.ReadProperty("f40403e", mp_sf40403e);
	oXML.ReadProperty("f40403f", mp_sf40403f);
	oXML.ReadProperty("f404040", mp_sf404040);
	oXML.ReadProperty("f404041", mp_sf404041);
	oXML.ReadProperty("f404042", mp_sf404042);
	oXML.ReadProperty("f404043", mp_sf404043);
	oXML.ReadProperty("f404044", mp_sf404044);
	oXML.ReadProperty("f404045", mp_sf404045);
	oXML.ReadProperty("f404046", mp_sf404046);
	oXML.ReadProperty("f404047", mp_sf404047);
	oXML.ReadProperty("f404048", mp_sf404048);
	oXML.ReadProperty("f404049", mp_sf404049);
	oXML.ReadProperty("f40404a", mp_sf40404a);
	oXML.ReadProperty("f40404b", mp_sf40404b);
	oXML.ReadProperty("f40404c", mp_sf40404c);
	oXML.ReadProperty("f40404d", mp_sf40404d);
	oXML.ReadProperty("f40404e", mp_sf40404e);
	oXML.ReadProperty("f40404f", mp_sf40404f);
	oXML.ReadProperty("f404050", mp_sf404050);
	oXML.ReadProperty("f404051", mp_sf404051);
	oXML.ReadProperty("f404052", mp_sf404052);
	oXML.ReadProperty("f404053", mp_sf404053);
	oXML.ReadProperty("f404054", mp_sf404054);
	oXML.ReadProperty("f404055", mp_sf404055);
	oXML.ReadProperty("f404056", mp_sf404056);
	oXML.ReadProperty("f404057", mp_sf404057);
	oXML.ReadProperty("f404058", mp_sf404058);
	oXML.ReadProperty("f404059", mp_sf404059);
	oXML.ReadProperty("f40405a", mp_sf40405a);
	oXML.ReadProperty("f40405b", mp_sf40405b);
	oXML.ReadProperty("f40405c", mp_sf40405c);
	oXML.ReadProperty("f40405d", mp_sf40405d);
	oXML.ReadProperty("f40405e", mp_sf40405e);
	oXML.ReadProperty("f40405f", mp_sf40405f);
	oXML.ReadProperty("f404060", mp_sf404060);
	oXML.ReadProperty("f404061", mp_sf404061);
	oXML.ReadProperty("f404062", mp_sf404062);
	oXML.ReadProperty("f404063", mp_sf404063);
	oXML.ReadProperty("f404064", mp_sf404064);
	oXML.ReadProperty("f404065", mp_sf404065);
	oXML.ReadProperty("f404066", mp_sf404066);
	oXML.ReadProperty("f404067", mp_sf404067);
	oXML.ReadProperty("f404068", mp_sf404068);
	oXML.ReadProperty("f404069", mp_sf404069);
	oXML.ReadProperty("f40406a", mp_sf40406a);
	oXML.ReadProperty("f40406b", mp_sf40406b);
	oXML.ReadProperty("f40406c", mp_sf40406c);
	oXML.ReadProperty("f40406d", mp_sf40406d);
	oXML.ReadProperty("f40406e", mp_sf40406e);
	oXML.ReadProperty("f40406f", mp_sf40406f);
	oXML.ReadProperty("f404070", mp_sf404070);
	oXML.ReadProperty("f404071", mp_sf404071);
	oXML.ReadProperty("f404072", mp_sf404072);
	oXML.ReadProperty("f404073", mp_sf404073);
	oXML.ReadProperty("f404074", mp_sf404074);
	oXML.ReadProperty("f404075", mp_sf404075);
	oXML.ReadProperty("f404076", mp_sf404076);
	oXML.ReadProperty("f404077", mp_sf404077);
	oXML.ReadProperty("f404078", mp_sf404078);
	oXML.ReadProperty("f404079", mp_sf404079);
	oXML.ReadProperty("f40407a", mp_sf40407a);
	oXML.ReadProperty("f40407b", mp_sf40407b);
	oXML.ReadProperty("f40407c", mp_sf40407c);
	oXML.ReadProperty("f40407d", mp_sf40407d);
	oXML.ReadProperty("f40407e", mp_sf40407e);
	oXML.ReadProperty("f40407f", mp_sf40407f);
	oXML.ReadProperty("f404080", mp_sf404080);
	oXML.ReadProperty("f404081", mp_sf404081);
	oXML.ReadProperty("f404082", mp_sf404082);
	oXML.ReadProperty("f404083", mp_sf404083);
	oXML.ReadProperty("f404084", mp_sf404084);
	oXML.ReadProperty("f404085", mp_sf404085);
	oXML.ReadProperty("f404086", mp_sf404086);
	oXML.ReadProperty("f404087", mp_sf404087);
	oXML.ReadProperty("f404088", mp_sf404088);
	oXML.ReadProperty("f404089", mp_sf404089);
	oXML.ReadProperty("f40408a", mp_sf40408a);
	oXML.ReadProperty("f40408b", mp_sf40408b);
	oXML.ReadProperty("f40408c", mp_sf40408c);
	oXML.ReadProperty("f40408d", mp_sf40408d);
	oXML.ReadProperty("f40408e", mp_sf40408e);
	oXML.ReadProperty("f40408f", mp_sf40408f);
	oXML.ReadProperty("f404090", mp_sf404090);
	oXML.ReadProperty("f404091", mp_sf404091);
	oXML.ReadProperty("f404092", mp_sf404092);
	oXML.ReadProperty("f404093", mp_sf404093);
	oXML.ReadProperty("f404094", mp_sf404094);
	oXML.ReadProperty("f404095", mp_sf404095);
	oXML.ReadProperty("f404096", mp_sf404096);
	oXML.ReadProperty("f404097", mp_sf404097);
	oXML.ReadProperty("f404098", mp_sf404098);
	oXML.ReadProperty("f404099", mp_sf404099);
	oXML.ReadProperty("f40409a", mp_sf40409a);
	oXML.ReadProperty("f40409b", mp_sf40409b);
	oXML.ReadProperty("f40409c", mp_sf40409c);
	oXML.ReadProperty("f40409d", mp_sf40409d);
	oXML.ReadProperty("f40409e", mp_sf40409e);
	oXML.ReadProperty("f40409f", mp_sf40409f);
	oXML.ReadProperty("f4040a0", mp_sf4040a0);
	oXML.ReadProperty("f4040a1", mp_sf4040a1);
	oXML.ReadProperty("f4040a2", mp_sf4040a2);
	oXML.ReadProperty("f4040a3", mp_sf4040a3);
	oXML.ReadProperty("f4040a4", mp_sf4040a4);
	oXML.ReadProperty("f4040a5", mp_sf4040a5);
	oXML.ReadProperty("f4040a6", mp_sf4040a6);
	oXML.ReadProperty("f4040a7", mp_sf4040a7);
	oXML.ReadProperty("f4040a8", mp_sf4040a8);
	oXML.ReadProperty("f4040a9", mp_sf4040a9);
	oXML.ReadProperty("f4040aa", mp_sf4040aa);
	oXML.ReadProperty("f4040ab", mp_sf4040ab);
	oXML.ReadProperty("f4040ac", mp_sf4040ac);
	oXML.ReadProperty("f4040ad", mp_sf4040ad);
	oXML.ReadProperty("f4040ae", mp_sf4040ae);
	oXML.ReadProperty("f4040af", mp_sf4040af);
	oXML.ReadProperty("f4040b0", mp_sf4040b0);
	oXML.ReadProperty("f4040b1", mp_sf4040b1);
	oXML.ReadProperty("f4040b2", mp_sf4040b2);
	oXML.ReadProperty("f4040b3", mp_sf4040b3);
	oXML.ReadProperty("f4040b4", mp_sf4040b4);
	oXML.ReadProperty("f4040b5", mp_sf4040b5);
	oXML.ReadProperty("f4040b6", mp_sf4040b6);
	oXML.ReadProperty("f4040b7", mp_sf4040b7);
	oXML.ReadProperty("f4040b8", mp_sf4040b8);
	oXML.ReadProperty("f4040b9", mp_sf4040b9);
	oXML.ReadProperty("f4040ba", mp_sf4040ba);
	oXML.ReadProperty("f4040bb", mp_sf4040bb);
	oXML.ReadProperty("f4040bc", mp_sf4040bc);
	oXML.ReadProperty("f4040bd", mp_sf4040bd);
	oXML.ReadProperty("f4040be", mp_sf4040be);
	oXML.ReadProperty("f4040bf", mp_sf4040bf);
	oXML.ReadProperty("f4040c0", mp_sf4040c0);
	oXML.ReadProperty("f4040c1", mp_sf4040c1);
	oXML.ReadProperty("f4040c2", mp_sf4040c2);
	oXML.ReadProperty("f4040c3", mp_sf4040c3);
	oXML.ReadProperty("f4040c4", mp_sf4040c4);
	oXML.ReadProperty("f4040c5", mp_sf4040c5);
	oXML.ReadProperty("f4040c6", mp_sf4040c6);
	oXML.ReadProperty("f4040c7", mp_sf4040c7);
	oXML.ReadProperty("f4040c8", mp_sf4040c8);
	mp_oTimephasedData_C->ReadObjectProtected(oXML);
}
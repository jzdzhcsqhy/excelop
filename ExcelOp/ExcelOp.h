
// ExcelOp.h : PROJECT_NAME Ӧ�ó������ͷ�ļ�
//

#pragma once

#ifndef __AFXWIN_H__
	#error "�ڰ������ļ�֮ǰ������stdafx.h�������� PCH �ļ�"
#endif

#include "resource.h"		// ������


// CExcelOpApp:
// �йش����ʵ�֣������ ExcelOp.cpp
//

class CExcelOpApp : public CWinApp
{
public:
	CExcelOpApp();

// ��д
public:
	virtual BOOL InitInstance();

// ʵ��

	DECLARE_MESSAGE_MAP()
};

extern CExcelOpApp theApp;
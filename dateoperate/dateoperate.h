
// dateoperate.h : PROJECT_NAME Ӧ�ó������ͷ�ļ�
//

#pragma once

#ifndef __AFXWIN_H__
	#error "�ڰ������ļ�֮ǰ������stdafx.h�������� PCH �ļ�"
#endif

#include "resource.h"		// ������


// CdateoperateApp:
// �йش����ʵ�֣������ dateoperate.cpp
//

class CdateoperateApp : public CWinApp
{
public:
	CdateoperateApp();

// ��д
public:
	virtual BOOL InitInstance();

// ʵ��

	DECLARE_MESSAGE_MAP()
};

extern CdateoperateApp theApp;
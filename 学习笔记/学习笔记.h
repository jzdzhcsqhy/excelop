
// ѧϰ�ʼ�.h : PROJECT_NAME Ӧ�ó������ͷ�ļ�
//

#pragma once

#ifndef __AFXWIN_H__
	#error "�ڰ������ļ�֮ǰ������stdafx.h�������� PCH �ļ�"
#endif

#include "resource.h"		// ������


// CNoteApp:
// �йش����ʵ�֣������ ѧϰ�ʼ�.cpp
//

class CNoteApp : public CWinApp
{
public:
	CNoteApp();

// ��д
public:
	virtual BOOL InitInstance();

// ʵ��

	DECLARE_MESSAGE_MAP()
};

extern CNoteApp theApp;
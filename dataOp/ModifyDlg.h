#pragma once
#include "afxwin.h"


// CModifyDlg �Ի���

class CModifyDlg : public CDialogEx
{
	DECLARE_DYNAMIC(CModifyDlg)

public:
	CModifyDlg(CWnd* pParent = NULL);   // ��׼���캯��
	CModifyDlg(CString str, CWnd* pParen = NULL);
	virtual ~CModifyDlg();

// �Ի�������
	enum { IDD = IDD_MODIFY };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��
	virtual BOOL OnInitDialog();
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedButton1();
	afx_msg void OnEnChangeInput();

public: 
	CString m_strData;

	afx_msg void OnBnClickedButton2();
};

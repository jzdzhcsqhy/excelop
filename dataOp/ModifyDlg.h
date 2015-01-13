#pragma once
#include "afxwin.h"


// CModifyDlg 对话框

class CModifyDlg : public CDialogEx
{
	DECLARE_DYNAMIC(CModifyDlg)

public:
	CModifyDlg(CWnd* pParent = NULL);   // 标准构造函数
	CModifyDlg(CString str, CWnd* pParen = NULL);
	virtual ~CModifyDlg();

// 对话框数据
	enum { IDD = IDD_MODIFY };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持
	virtual BOOL OnInitDialog();
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedButton1();
	afx_msg void OnEnChangeInput();

public: 
	CString m_strData;

	afx_msg void OnBnClickedButton2();
};

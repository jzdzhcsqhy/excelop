#pragma once


// CModifyDlg 对话框

class CModifyDlg : public CDialogEx
{
	DECLARE_DYNAMIC(CModifyDlg)

public:
	CModifyDlg(CWnd* pParent = NULL);   // 标准构造函数
	virtual ~CModifyDlg();

// 对话框数据
	enum { IDD = IDD_MODIFY };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedButton1();
};

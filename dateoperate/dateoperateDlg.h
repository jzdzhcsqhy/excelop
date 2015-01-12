
// dateoperateDlg.h : 头文件
//

#pragma once

#include <queue>
#include <vector>

using namespace std;

// CdateoperateDlg 对话框
class CdateoperateDlg : public CDialogEx
{
// 构造
public:
	CdateoperateDlg(CWnd* pParent = NULL);	// 标准构造函数

// 对话框数据
	enum { IDD = IDD_DATEOPERATE_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支持


// 实现
protected:
	HICON m_hIcon;

	// 生成的消息映射函数
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnEnChangeSelectfile();

	//add by lzb

private	:
	CString m_strPath;
	vector<CString> m_vFileList;
public:
	afx_msg void OnSize(UINT nType, int cx, int cy);
};

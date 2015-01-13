
// dataOpDlg.h : 头文件
//

#pragma once

#include <vector>
#include <queue>

#include "resource.h"
#include <fstream>
using namespace std;

// CdataOpDlg 对话框
class CdataOpDlg : public CDialogEx
{
// 构造
public:
	CdataOpDlg(CWnd* pParent = NULL);	// 标准构造函数
	~CdataOpDlg();

// 对话框数据
	enum { IDD = IDD_DATAOP_DIALOG };

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

private:
	vector<CString> m_vFileList;
	CString m_strPath;
	CWinThread* m_pthMainProcess;
	CApplication m_ExcelApp;
	CWorkbooks m_books;
	bool		m_bIsExcel;

	CString		m_strCurBook;
	CString		m_strCurSheet;

public:
	void refreshListBox();
	void DisPlay( vector<double> vd, CString sheetname);
	void DisPlay( vector<double> vd);
	void ResetOutput( void );
	void CdataOpDlg::saveAs( vector<double> &vd );
	static UINT MainProcess( LPVOID lParam );
	void MainProcess(void );
	static void dealWith( const CString &filename, CdataOpDlg* p);
	static CString VariantToCString( VARIANT var );
	static double GetNumber(CString strNumber, CString strSplit, int *pos);
public:
	afx_msg void OnEnChangeSelectfile();
//	afx_msg void OnLbnSelchangeFilelist();
	afx_msg void OnLbnDblclkFilelist();
	afx_msg void OnBnClickedButton1();
	afx_msg void OnBnClickedButton2();
	afx_msg void OnBnClickedOk();
	afx_msg void OnBnClickedCancel();
	afx_msg void OnNMClickOutput(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnNMDblclkOutput(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnNMCustomdrawOutput(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnBnClickedExport();
};

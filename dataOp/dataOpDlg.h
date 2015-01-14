
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


#ifndef _XLFILEFORMAT_
#define _XLFILEFORMAT_
enum XlFileFormat
{
	xlAddIn = 18,
	xlCSV = 6,
	xlCSVMac = 22,
	xlCSVMSDOS = 24,
	xlCSVWindows = 23,
	xlDBF2 = 7,
	xlDBF3 = 8,
	xlDBF4 = 11,
	xlDIF = 9,
	xlExcel2 = 16,
	xlExcel2FarEast = 27,
	xlExcel3 = 29,
	xlExcel4 = 33,
	xlExcel5 = 39,
	xlExcel7 = 39,
	xlExcel9795 = 43,
	xlExcel4Workbook = 35,
	xlIntlAddIn = 26,
	xlIntlMacro = 25,
	xlWorkbookNormal = -4143,
	xlSYLK = 2,
	xlTemplate = 17,
	xlCurrentPlatformText = -4158,
	xlTextMac = 19,
	xlTextMSDOS = 21,
	xlTextPrinter = 36,
	xlTextWindows = 20,
	xlWJ2WD1 = 14,
	xlWK1 = 5,
	xlWK1ALL = 31,
	xlWK1FMT = 30,
	xlWK3 = 15,
	xlWK4 = 38,
	xlWK3FM3 = 32,
	xlWKS = 4,
	xlWorks2FarEast = 28,
	xlWQ1 = 34,
	xlWJ3 = 40,
	xlWJ3FJ3 = 41,
	xlUnicodeText = 42,
	xlHtml = 44,
	xlWebArchive = 45,
	xlXMLSpreadsheet = 46,
	xlExcel12 = 50,
	xlOpenXMLWorkbook = 51,
	xlOpenXMLWorkbookMacroEnabled = 52,
	xlOpenXMLTemplateMacroEnabled = 53,
	xlTemplate8 = 17,
	xlOpenXMLTemplate = 54,
	xlAddIn8 = 18,
	xlOpenXMLAddIn = 55,
	xlExcel8 = 56,
	xlOpenDocumentSpreadsheet = 60,
	xlWorkbookDefault = 51
};

enum XlSaveConflictResolution
{
	xlLocalSessionChanges = 2,
	xlOtherSessionChanges = 3,
	xlUserResolution = 1
};

#endif
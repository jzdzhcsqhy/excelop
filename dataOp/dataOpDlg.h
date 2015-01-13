
// dataOpDlg.h : ͷ�ļ�
//

#pragma once

#include <vector>
#include <queue>

#include "resource.h"
#include <fstream>
using namespace std;

// CdataOpDlg �Ի���
class CdataOpDlg : public CDialogEx
{
// ����
public:
	CdataOpDlg(CWnd* pParent = NULL);	// ��׼���캯��
	~CdataOpDlg();

// �Ի�������
	enum { IDD = IDD_DATAOP_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV ֧��


// ʵ��
protected:
	HICON m_hIcon;

	// ���ɵ���Ϣӳ�亯��
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

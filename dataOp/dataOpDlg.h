
// dataOpDlg.h : ͷ�ļ�
//

#pragma once

#include <vector>
#include <queue>

#include "resource.h"
using namespace std;

// CdataOpDlg �Ի���
class CdataOpDlg : public CDialogEx
{
// ����
public:
	CdataOpDlg(CWnd* pParent = NULL);	// ��׼���캯��

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

public:
	void refreshListBox();

	static UINT MainProcess( LPVOID lParam );
	static void dealWith( const CString &filename, CdataOpDlg* p);

public:
	afx_msg void OnEnChangeSelectfile();
//	afx_msg void OnLbnSelchangeFilelist();
	afx_msg void OnLbnDblclkFilelist();
	afx_msg void OnBnClickedButton1();
	afx_msg void OnBnClickedButton2();
	afx_msg void OnBnClickedOk();
	afx_msg void OnBnClickedCancel();
};

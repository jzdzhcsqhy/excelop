
// dateoperateDlg.h : ͷ�ļ�
//

#pragma once

#include <queue>
#include <vector>

using namespace std;

// CdateoperateDlg �Ի���
class CdateoperateDlg : public CDialogEx
{
// ����
public:
	CdateoperateDlg(CWnd* pParent = NULL);	// ��׼���캯��

// �Ի�������
	enum { IDD = IDD_DATEOPERATE_DIALOG };

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
public:
	afx_msg void OnEnChangeSelectfile();

	//add by lzb

private	:
	CString m_strPath;
	vector<CString> m_vFileList;
public:
	afx_msg void OnSize(UINT nType, int cx, int cy);
};

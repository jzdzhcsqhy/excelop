// ModifyDlg.cpp : 实现文件
//

#include "stdafx.h"
#include "dataOp.h"
#include "ModifyDlg.h"
#include "afxdialogex.h"


// CModifyDlg 对话框

IMPLEMENT_DYNAMIC(CModifyDlg, CDialogEx)

CModifyDlg::CModifyDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(CModifyDlg::IDD, pParent)
{
	
}

CModifyDlg::CModifyDlg(CString str ,CWnd* pParent )
	: CDialogEx(CModifyDlg::IDD, pParent)
{
	this->m_strData = str;
}

BOOL CModifyDlg::OnInitDialog()
{
	this->SetDlgItemTextW(IDC_INPUT, this->m_strData );
	return TRUE;
}


CModifyDlg::~CModifyDlg()
{
}

void CModifyDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	//  DDX_Control(pDX, IDC_INPUT, m_cEdit);
}


BEGIN_MESSAGE_MAP(CModifyDlg, CDialogEx)
	ON_BN_CLICKED(IDC_BUTTON1, &CModifyDlg::OnBnClickedButton1)
	ON_EN_CHANGE(IDC_INPUT, &CModifyDlg::OnEnChangeInput)
	ON_BN_CLICKED(IDC_BUTTON2, &CModifyDlg::OnBnClickedButton2)
END_MESSAGE_MAP()


// CModifyDlg 消息处理程序


void CModifyDlg::OnBnClickedButton1()
{
	// TODO: ÔÚ´ËÌí¼Ó¿Ø¼þÍ¨Öª´¦Àí³ÌÐò´úÂë
	this->GetDlgItemTextW(IDC_INPUT, this->m_strData);
	CDialogEx::OnOK();
}


void CModifyDlg::OnEnChangeInput()
{
	// TODO:  Èç¹û¸Ã¿Ø¼þÊÇ RICHEDIT ¿Ø¼þ£¬Ëü½«²»
	// ·¢ËÍ´ËÍ¨Öª£¬³ý·ÇÖØÐ´ CDialogEx::OnInitDialog()
	// º¯Êý²¢µ÷ÓÃ CRichEditCtrl().SetEventMask()£¬
	// Í¬Ê±½« ENM_CHANGE ±êÖ¾¡°»ò¡±ÔËËãµ½ÑÚÂëÖÐ¡£

	// TODO:  ÔÚ´ËÌí¼Ó¿Ø¼þÍ¨Öª´¦Àí³ÌÐò´úÂë
}


void CModifyDlg::OnBnClickedButton2()
{
	// TODO: ÔÚ´ËÌí¼Ó¿Ø¼þÍ¨Öª´¦Àí³ÌÐò´úÂë
	CDialogEx::OnCancel();
}

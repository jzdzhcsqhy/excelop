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

CModifyDlg::~CModifyDlg()
{
}

void CModifyDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}


BEGIN_MESSAGE_MAP(CModifyDlg, CDialogEx)
	ON_BN_CLICKED(IDC_BUTTON1, &CModifyDlg::OnBnClickedButton1)
END_MESSAGE_MAP()


// CModifyDlg 消息处理程序


void CModifyDlg::OnBnClickedButton1()
{
	// TODO: ÔÚ´ËÌí¼Ó¿Ø¼þÍ¨Öª´¦Àí³ÌÐò´úÂë
	CDialogEx::OnOK();
}

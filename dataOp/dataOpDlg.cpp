﻿
// dataOpDlg.cpp : 实现文件
//

#include "stdafx.h"
#include "dataOp.h"
#include "dataOpDlg.h"
#include "afxdialogex.h"
#include "excel9.h"
#include <atlbase.h>
#include <atlstr.h>


#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// 用于应用程序“关于”菜单项的 CAboutDlg 对话框

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// 对话框数据
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

// 实现
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialogEx(CAboutDlg::IDD)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// CdataOpDlg 对话框




CdataOpDlg::CdataOpDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(CdataOpDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CdataOpDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CdataOpDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_EN_CHANGE(IDC_SELECTFILE, &CdataOpDlg::OnEnChangeSelectfile)
//	ON_LBN_SELCHANGE(IDC_FILELIST, &CdataOpDlg::OnLbnSelchangeFilelist)
	ON_LBN_DBLCLK(IDC_FILELIST, &CdataOpDlg::OnLbnDblclkFilelist)
	ON_BN_CLICKED(IDC_BUTTON1, &CdataOpDlg::OnBnClickedButton1)
	ON_BN_CLICKED(IDC_BUTTON2, &CdataOpDlg::OnBnClickedButton2)
	ON_BN_CLICKED(IDOK, &CdataOpDlg::OnBnClickedOk)
	ON_BN_CLICKED(IDCANCEL, &CdataOpDlg::OnBnClickedCancel)
	ON_BN_CLICKED(IDC_EXPORT, &CdataOpDlg::OnBnClickedExport)
END_MESSAGE_MAP()


CdataOpDlg::~CdataOpDlg()
{
	this->m_books.ReleaseDispatch();
	this->m_ExcelApp.Quit();
	this->m_ExcelApp.ReleaseDispatch();
	
}


// CdataOpDlg 消息处理程序

BOOL CdataOpDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// 将“关于...”菜单项添加到系统菜单中。

	// IDM_ABOUTBOX 必须在系统命令范围内。
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// 设置此对话框的图标。当应用程序主窗口不是对话框时，框架将自动
	//  执行此操作
	SetIcon(m_hIcon, TRUE);			// 设置大图标
	SetIcon(m_hIcon, FALSE);		// 设置小图标

	// TODO: 在此添加额外的初始化代码

	if( ! AfxOleInit() )
	{
		AfxMessageBox(_T("启动OLE失败"));
		return FALSE;
	}

	//创建Excel 服务器(启动Excel)
	if(!this->m_ExcelApp.CreateDispatch(_T("Excel.Application")) )
	{
		AfxMessageBox(_T("启动Excel服务器失败!"));
		CDialogEx::OnCancel();
		return FALSE;
	}

	

	/*判断当前Excel的版本*/
	CString strExcelVersion = this->m_ExcelApp.get_Version();
	int iStart = 0;
	strExcelVersion = strExcelVersion.Tokenize(_T("."), iStart);
	if (_T("11") == strExcelVersion)
	{
		AfxMessageBox(_T("当前Excel的版本是2003。"));
	}
	else if (_T("12") == strExcelVersion)
	{
		AfxMessageBox(_T("当前Excel的版本是2007。"));
	}
	else if (_T("14") == strExcelVersion)
	{
		AfxMessageBox(_T("当前Excel的版本是2010。"));
	}
	else
	{
		AfxMessageBox(_T("当前Excel的版本是其他版本。"));
	}

	this->m_ExcelApp.put_UserControl(FALSE);
	/*得到工作簿容器*/
	this->m_books.AttachDispatch(this->m_ExcelApp.get_Workbooks());

	this->m_bIsExcel = true;

	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
}

void CdataOpDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialogEx::OnSysCommand(nID, lParam);
	}
}

// 如果向对话框添加最小化按钮，则需要下面的代码
//  来绘制该图标。对于使用文档/视图模型的 MFC 应用程序，
//  这将由框架自动完成。

void CdataOpDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 用于绘制的设备上下文

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// 使图标在工作区矩形中居中
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// 绘制图标
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

//当用户拖动最小化窗口时系统调用此函数取得光标
//显示。
HCURSOR CdataOpDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}



void CdataOpDlg::OnEnChangeSelectfile()
{
	CFileFind fileFinder;
	CString path;

	this->m_vFileList.clear();
	this->GetDlgItemTextA(IDC_SELECTFILE, this->m_strPath);

	path = this->m_strPath + _T("\\*.xls");

	bool isExist = fileFinder.FindFile(path);
	while( isExist )
	{
		isExist = fileFinder.FindNextFileA();
		this->m_vFileList.push_back(fileFinder.GetFileName());
		refreshListBox();
	}
}

void CdataOpDlg::refreshListBox()
{
	CListBox* pList= (CListBox *)this->GetDlgItem(IDC_FILELIST);
	pList->ResetContent();
	
	int i;
	for(i=0; i<this->m_vFileList.size(); i++ )
	{
		pList->AddString(this->m_vFileList[i]);
	}
}


void CdataOpDlg::OnLbnDblclkFilelist()
{
	CListBox* pList = (CListBox *)this->GetDlgItem(IDC_FILELIST);
	CArray<int, int> listContent;
	int iCnt = pList->GetSelCount();
	listContent.SetSize(iCnt);
	
	pList->GetSelItems(iCnt, listContent.GetData());

	
	if( IDYES == MessageBox(_T("确认不处理这些文件吗？"),_T("提示"),MB_YESNO))
	{
		for(int i=listContent.GetSize()-1; i>=0; i-- )
		{
			pList->DeleteString(listContent.GetAt(i));
		}
	}
}


void CdataOpDlg::OnBnClickedButton1()
{
	OnLbnDblclkFilelist();
}


void CdataOpDlg::OnBnClickedButton2()
{
	refreshListBox();
}


void CdataOpDlg::OnBnClickedOk()
{
	CListBox* pList = (CListBox *)this->GetDlgItem(IDC_FILELIST);
	int iCnt = pList->GetCount();

	if( 0 == iCnt )
	{
		AfxMessageBox(_T("请选择至少一个文件!"));
		return ;
	}

	CString str ;
	str.Format(_T("您一共选择了 %d 个文件，是否开始处理？"),iCnt);
	if( IDYES == MessageBox(str, _T("提示"), MB_YESNO) )
	{
		MainProcess();
		//this->m_pthMainProcess = AfxBeginThread(CdataOpDlg::MainProcess,this);
		AfxMessageBox(_T("处理完成"));
	}
	else
	{

	}
	//CDialogEx::OnOK();
}


void CdataOpDlg::OnBnClickedCancel()
{
	this->m_books.ReleaseDispatch();
	this->m_books.Close();
	this->m_ExcelApp.Quit();
	this->m_ExcelApp.ReleaseDispatch();
	CDialogEx::OnCancel();
}


UINT CdataOpDlg::MainProcess( LPVOID lParam )
{
	CdataOpDlg* pThis = (CdataOpDlg *)lParam;

	CWnd* pStatus = pThis->GetDlgItem(IDC_STATUS);
	CListBox* pFileList =(CListBox* ) pThis->GetDlgItem(IDC_FILELIST);

	int iCnt = pFileList->GetCount();
	for(int i=0; i<iCnt; i++ )
	{
		CString str;
		pFileList->GetText(i,str);
		pThis->SetDlgItemTextA(IDC_STATUS,_T("正在处理文件 ") + str + _T("...") );
		CdataOpDlg::dealWith(str, pThis);
	}


	return 0;
}

void CdataOpDlg::MainProcess(void )
{
	CdataOpDlg* pThis = this;

	CWnd* pStatus = pThis->GetDlgItem(IDC_STATUS);
	CListBox* pFileList =(CListBox* ) pThis->GetDlgItem(IDC_FILELIST);

	int iCnt = pFileList->GetCount();
	for(int i=0; i<iCnt; i++ )
	{
		CString str;
		pFileList->GetText(i,str);
		pThis->SetDlgItemTextA(IDC_STATUS,_T("正在处理文件 ") + str + _T("...") );
		CdataOpDlg::dealWith(str, pThis);

	}


	return ;
}

void CdataOpDlg::ResetOutput()
{
}
// 
// void CdataOpDlg::DisPlay( vector<double> vd)
// {
// 	CListCtrl* pList=(CListCtrl*) this->GetDlgItem(IDC_OUTPUT);
// 	CString str;
// 
// 	int row = vd.size() /20 -1;
// 	if( vd.size() %20 )
// 	{
// 		row += 1;
// 	}
// 	str.Format(_T(FORMAT_STRING ), vd[row*20] );
// 	int nRow= pList->InsertItem(row,str );
// 	for( int i=1; i<20 && row*20+i < vd.size(); i++ )
// 	{
// 		LV_ITEM lvitem = {0};
// 		lvitem.mask = LVIF_TEXT;
// 		lvitem.iItem = nRow;
// 		lvitem.iSubItem = i;
// 
// 		str.Format(_T(FORMAT_STRING ), vd[row*20 +i] );
// 		lvitem.pszText = str.GetBuffer();
// 		pList->SetItem(&lvitem);
// 		//pList->SetItemText(nRow, i+1, str);
// 	}
// }

void CdataOpDlg::DisPlay( vector<double> vd)
{
	if( !this->m_bIsExcel )
	{
		CString name =  this->m_strPath+"\\"+this->m_strCurBook+"_"+this->m_strCurSheet+".txt";
		this->DisPlay(vd, name);
		return ;
	}

	this->saveAs(vd);


}

void CdataOpDlg::DisPlay( vector<double> vd, CString sheetname)
{
	FILE* fp;

	if( NULL == ( fp = fopen(CT2A(sheetname), "a+")))
	{
		CString str;
		str.Format(_T("%S %d"), CT2A(sheetname),GetLastError() );
		AfxMessageBox(_T("打开文件失败")+sheetname+" " +str);
		GetLastError();
		return ;
	}


	int row = vd.size() /20 -1;
	if( vd.size() %20 )
	{
		row += 1;
	}
	for( int i=0; i<20 && row*20+i < vd.size(); i++ )
	{
		fprintf(fp,FORMAT_STRING, vd[row*20 +i ]);
		if( i != 19 && row *20 +i != vd.size() -1 )
		{
			fprintf(fp, " ");
		}
	}
	fprintf(fp, "\n");
	fclose(fp);
}

void CdataOpDlg::OnNMClickOutput(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMITEMACTIVATE pNMItemActivate = reinterpret_cast<LPNMITEMACTIVATE>(pNMHDR);
	*pResult = 0;
}


void CdataOpDlg::OnNMDblclkOutput(NMHDR *pNMHDR, LRESULT *pResult)
{
	
}


void CdataOpDlg::OnNMCustomdrawOutput(NMHDR *pNMHDR, LRESULT *pResult)
{
	
}


void CdataOpDlg::OnBnClickedExport()
{
	if(this->m_bIsExcel )
	{
		this->m_bIsExcel = !this->m_bIsExcel;
		this->SetDlgItemTextA(IDC_EXPORT,_T("导出TXT"));
	}
	else
	{
		this->SetDlgItemTextA(IDC_EXPORT,_T("导出EXCEL"));
		this->m_bIsExcel = !this->m_bIsExcel;
	}
}


void CdataOpDlg::saveAs( vector<double> &vd )
{
	//CString filename = this->m_strPath+"\\rs_"+this->m_strCurBook;
	CString filename = this->m_strPath+"\\"+this->m_strCurBook+"___"+this->m_strCurSheet+".xls";
// 
// 	CString shtnm = this->m_strCurSheet;
// 
// 	for( int j=0; j<shtnm.GetLength(); j++ )
// 	{
// 		if(L' ' == shtnm.GetAt(j) )
// 		{
// 			shtnm.SetAt(j,L'0');
// 		}
// 	}
	CSpreadSheet newxls(filename,"sheet1");
	
	int i;
	CStringArray row;
	row.RemoveAll();

	for(int i =0; i<20; i++ )
	{
		CString str;
		str.Format("第%d列", i+1);
		row.Add(str);
	}

	newxls.AddHeaders(row);
	newxls.BeginTransaction();
	row.RemoveAll();
	int iCnt =2;

	for( i=0; i<vd.size(); i++ )
	{
		CString str;
		str.Format(FORMAT_STRING, vd[i] );
		row.Add(str);
		if( 20 == row.GetSize() )
		{
			
			newxls.AddRow(row,iCnt++, true);
			row.RemoveAll();
			
		}
	}
	if( 0 != row.GetSize() )
	{
		newxls.AddRow(row, iCnt++, true);
	}
	newxls.Commit();
}
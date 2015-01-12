#include "stdafx.h"
#include "dataOpDlg.h"
#include "excel9.h"

void CdataOpDlg::dealWith(const CString &filename, CdataOpDlg* pThis)
{
	
	CWorkbooks books;
	CWorkbook book;
	CWorksheets sheets;
	CWorksheet sheet;
	CRange range;
	LPDISPATCH lpDisp = NULL;

	//创建Excel 服务器(启动Excel)
	

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
	else
	{
		AfxMessageBox(_T("当前Excel的版本是其他版本。"));
	}

	this->m_ExcelApp.put_Visible(TRUE);
	this->m_ExcelApp.put_UserControl(FALSE);

	/*得到工作簿容器*/
	books.AttachDispatch(this->m_ExcelApp.get_Workbooks());

	/*打开一个工作簿，如不存在，则新增一个工作簿*/
	CString strBookPath = _T("C:\\tmp.xls");
	try
	{
		/*打开一个工作簿*/
		lpDisp = books.Open(strBookPath, 
			vtMissing, vtMissing, vtMissing, vtMissing, vtMissing,
			vtMissing, vtMissing, vtMissing, vtMissing, vtMissing, 
			vtMissing, vtMissing, vtMissing, vtMissing);
		book.AttachDispatch(lpDisp);
	}
	catch(...)
	{
		/*增加一个新的工作簿*/
		lpDisp = books.Add(vtMissing);
		book.AttachDispatch(lpDisp);
	}


	/*得到工作簿中的Sheet的容器*/
	sheets.AttachDispatch(book.get_Sheets());

	/*打开一个Sheet，如不存在，就新增一个Sheet*/
	CString strSheetName = _T("NewSheet");
	try
	{
		/*打开一个已有的Sheet*/
		lpDisp = sheets.get_Item(_variant_t(strSheetName));
		sheet.AttachDispatch(lpDisp);
	}
	catch(...)
	{
		/*创建一个新的Sheet*/
		lpDisp = sheets.Add(vtMissing, vtMissing, _variant_t((long)1), vtMissing);
		sheet.AttachDispatch(lpDisp);
		sheet.put_Name(strSheetName);
	}

	system("pause");

	/*向Sheet中写入多个单元格,规模为10*10 */
	lpDisp = sheet.get_Range(_variant_t("A1"), _variant_t("J10"));
	range.AttachDispatch(lpDisp);

	VARTYPE vt = VT_I4; /*数组元素的类型，long*/
	SAFEARRAYBOUND sabWrite[2]; /*用于定义数组的维数和下标的起始值*/
	sabWrite[0].cElements = 10;
	sabWrite[0].lLbound = 0;
	sabWrite[1].cElements = 10;
	sabWrite[1].lLbound = 0;

	COleSafeArray olesaWrite;
	olesaWrite.Create(vt, sizeof(sabWrite)/sizeof(SAFEARRAYBOUND), sabWrite);

	/*通过指向数组的指针来对二维数组的元素进行间接赋值*/
	long (*pArray)[2] = NULL;
	olesaWrite.AccessData((void **)&pArray);
	memset(pArray, 0, sabWrite[0].cElements * sabWrite[1].cElements * sizeof(long));

	/*释放指向数组的指针*/
	olesaWrite.UnaccessData();
	pArray = NULL;

	/*对二维数组的元素进行逐个赋值*/
	long index[2] = {0, 0};
	long lFirstLBound = 0;
	long lFirstUBound = 0;
	long lSecondLBound = 0;
	long lSecondUBound = 0;
	olesaWrite.GetLBound(1, &lFirstLBound);
	olesaWrite.GetUBound(1, &lFirstUBound);
	olesaWrite.GetLBound(2, &lSecondLBound);
	olesaWrite.GetUBound(2, &lSecondUBound);
	for (long i = lFirstLBound; i <= lFirstUBound; i++)
	{
		index[0] = i;
		for (long j = lSecondLBound; j <= lSecondUBound; j++)
		{
			index[1] = j;
			long lElement = i * sabWrite[1].cElements + j; 
			olesaWrite.PutElement(index, &lElement);
		}
	}

	/*把ColesaWritefeArray变量转换为VARIANT,并写入到Excel表格中*/
	VARIANT varWrite = (VARIANT)olesaWrite;
	range.put_Value2(varWrite);

	system("pause");



	/*读取Excel表中的多个单元格的值，在listctrl中显示*/
	VARIANT varRead = range.get_Value2();
	COleSafeArray olesaRead(varRead);

	VARIANT varItem;
	CString strItem;
	lFirstLBound = 0;
	lFirstUBound = 0;
	lSecondLBound = 0;
	lSecondUBound = 0;
	olesaRead.GetLBound(1, &lFirstLBound);
	olesaRead.GetUBound(1, &lFirstUBound);
	olesaRead.GetLBound(2, &lSecondLBound);
	olesaRead.GetUBound(2, &lSecondUBound);
	memset(index, 0, 2 * sizeof(long));


	/*释放资源*/
	sheet.ReleaseDispatch();
	sheets.ReleaseDispatch();
	book.ReleaseDispatch();
	books.ReleaseDispatch();
	this->m_ExcelApp.Quit();
	this->m_ExcelApp.ReleaseDispatch();
}
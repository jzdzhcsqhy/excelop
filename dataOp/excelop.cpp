#include "stdafx.h"
#include "dataOpDlg.h"
#include "excel9.h"


double CdataOpDlg::GetNumber(CString strNumber, CString strSplit, int *pos)
{
	TCHAR szNumber[20];
	memset(szNumber, 0, 20*sizeof(TCHAR));
	int end ;
	int poss = *pos;
	CString strTemp = strNumber;
	strTemp.Delete(0,*pos);
	end = strTemp.Find(strSplit);
	_tcsncpy(szNumber, strNumber.GetBuffer()+*pos,end );
	poss += end;
	strTemp.Delete(0,end);
	while( strTemp.GetLength() >= 0 && strTemp[0] == ' ' )
	{
		strTemp.Delete(0,1);
		poss ++;
	}
	*pos = poss;
// 	CString str;
// 	str.Format(_T("%s\n"), szNumber);
// 	AfxMessageBox(str);
	return _tstof(szNumber);
}

CString CdataOpDlg::VariantToCString(VARIANT var)
{
	CString strValue;
	_variant_t var_t;
	_bstr_t bst_t;
	time_t cur_time;
	CTime time_value;
	COleCurrency var_currency;
	switch(var.vt)
	{
	case VT_EMPTY:
		strValue=_T("");
		break;
	case VT_UI1:
		strValue.Format(_T("%d"),var.bVal);
		break;
	case VT_I2:
		strValue.Format(_T("%d"),var.iVal);
		break;
	case VT_I4:
		strValue.Format(_T("%d"),var.lVal);
		break;
	case VT_R4:
		strValue.Format(_T("%.2f"),var.fltVal);
		break;
	case VT_R8:
		strValue.Format(_T("%.2f"),var.dblVal);
		break;
	case VT_CY:
		var_currency=var;
		strValue=var_currency.Format(0);
		break;
	case VT_BSTR:
		strValue = var.bstrVal;
		break;
	case VT_NULL:
		strValue=_T("");
		break;
	case VT_DATE:
		cur_time = (long)var.date;
		time_value=cur_time;
		strValue=time_value.Format("%A,%B%d,%Y");
		break;
	case VT_BOOL:
		strValue.Format(_T("%d"),var.boolVal );
		break;
	default: 
		strValue=_T("");
		break;
	}
	return strValue;
}


void CdataOpDlg::dealWith(const CString &filename, CdataOpDlg* pThis)
{
	
	
	CWorkbook book;
	CWorksheets sheets;
	CWorksheet sheet;
	CRange range;
	LPDISPATCH lpDisp = NULL;
	vector<double> vDq;
	vDq.clear();
	
	CListBox* pList = (CListBox*) pThis->GetDlgItem(IDC_RS);
	CString rs;
	

	/*��һ�����������粻���ڣ�������һ��������*/
	CString strBookPath = pThis->m_strPath + "\\" + filename;
	CString strOutBookPath = pThis->m_strPath +"\\rs_" +filename; 
	rs = filename + _T(" ����ʼ����");
	pList->AddString(rs);
	try
	{
		/*��һ��������*/
		pThis->m_strCurBook = filename;
		lpDisp = pThis->m_books.Open(strBookPath, 
			vtMissing, vtMissing, vtMissing, vtMissing, vtMissing,
			vtMissing, vtMissing, vtMissing, vtMissing, vtMissing, 
			vtMissing, vtMissing, vtMissing, vtMissing);
		book.AttachDispatch(lpDisp);
		rs = pThis->m_strCurBook + _T(" �򿪳ɹ�");
		pList->AddString(rs);
	}
	catch(...)
	{
	
		rs = pThis->m_strCurBook + _T(" �򿪴�������");
		pList->AddString(rs);
		return;
	}


	/*�õ��������е�Sheet������*/
	sheets.AttachDispatch(book.get_Sheets());

	int n = sheets.get_Count();
	for(int sheet_cnt =1; sheet_cnt<=n; sheet_cnt ++)
	{
// 		CString shtcntstr;
// 		shtcntstr.Format(_T("%d"), sheet_cnt);
// 		AfxMessageBox(shtcntstr);

		/*��һ��Sheet���粻���ڣ�������һ��Sheet*/
		vDq.clear();
		try
		{
			/*��һ�����е�Sheet*/
			lpDisp = sheets.get_Item(_variant_t(sheet_cnt));
			
			sheet.AttachDispatch(lpDisp);
			/*AfxMessageBox(sheet.get_Name());*/
			pThis->m_strCurSheet = sheet.get_Name();
			//pThis->m_strCurSheet.Format(_T("%d"), sheet_cnt);
			rs = pThis->m_strCurSheet + _T(" �򿪳ɹ�");
			pList->AddString(rs);
		}
		catch(...)
		{
			/*����һ���µ�Sheet*/
			rs = pThis->m_strCurSheet + _T(" �򿪴�������");
			pList->AddString(rs);
		}

		/*��Sheet��д������Ԫ��,��ģΪ10*10 */
		lpDisp = sheet.get_UsedRange();
		range.AttachDispatch(lpDisp);

		range.put_NumberFormat(COleVariant(L"@"));
		/*��ȡExcel���еĶ����Ԫ���ֵ����listctrl����ʾ*/
		VARIANT varRead = range.get_Value2();
		COleSafeArray olesaRead(varRead);
	
		VARIANT varItem;
		CString strItem;
		long index[2] = {0, 0};
		long lFirstLBound = 0;
		long lFirstUBound = 0;
		long lSecondLBound = 0;
		long lSecondUBound = 0;
		olesaRead.GetLBound(1, &lFirstLBound);
		olesaRead.GetUBound(1, &lFirstUBound);
		olesaRead.GetLBound(2, &lSecondLBound);
		olesaRead.GetUBound(2, &lSecondUBound);
		memset(index, 0, 2 * sizeof(long));

	//  	CString sCount;
	//  	sCount.Format(_T("һ��%d %d %d %d "), 
	//  			lFirstLBound, 
	//  			lFirstUBound,
	//  			lSecondLBound,
	//  			lSecondUBound);
	//  	AfxMessageBox(sCount);

		int i,j;
		for(i=lFirstLBound; i<lFirstUBound; i++)
		{
			index[0] = i;
			for( j=8; j<=11; j++ )
			{
			
				index[1] = j;
				CString str;
			
				olesaRead.GetElement(index, &varItem);
				str.Format(_T("%d"),varItem.vt);
				str = VariantToCString(varItem);
				
				for(int strclri=0; strclri<str.GetLength(); strclri++ )
				{
					if( (str[strclri] < L'0' || str[strclri] > L'9') && str[strclri] != '.')
					{
						str.SetAt(strclri,' ');
					}
				}

				str.TrimLeft();
				str.TrimRight();
				/*AfxMessageBox(str);*/
				if( str != "" && str.GetLength() >0 )
				{
					
					int pos = 0;
					str += " ";
					
		
					while( pos < str.GetLength() )
					{
						vDq.push_back( GetNumber(str, _T(" "),&pos) );
						if( vDq.size() % 20 == 0 )
						{
							pThis->DisPlay( vDq );
						}
					}
					
				}	
			}
		}
		pThis->DisPlay( vDq );
		rs = pThis->m_strCurSheet + _T(" �������");
		pList->AddString(rs);
	}

	rs = pThis->m_strCurBook + _T(" �������");
	pList->AddString(rs);
// 	CString str ;
// 	str.Format(_T("%d"), vDq.size());
// 	AfxMessageBox(str);

	/*�ͷ���Դ*/
	sheet.ReleaseDispatch();
	sheets.ReleaseDispatch();
	book.ReleaseDispatch();
}


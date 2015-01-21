
#include "StdAfx.h"
#include "ReadTableToExcel.h"


bool sortDVTextByDx(tDBTextInfor text1, tDBTextInfor text2)
{
	return text1.dX>text2.dX;
}

//////////////////////////////////////////////////////////////////////////

CReadTableToExcel::CReadTableToExcel()
{

}

void CReadTableToExcel::Run()
{
	vector<vector<CString> > vecTable;
	readTable(_T("\nѡ��ͼֽ�ϵı���Զ����浽excel��"), vecTable);
	CreateDirectory(_T("D:\\��ȡDWG���\\"), NULL);
	acedInitGet(NULL, NULL);
	ACHAR chFileName[MAX_PATH] = {0};
	acedGetString(0, _T("\n���뱣����ļ����ƣ����ô���չ����"), chFileName);
	CString strFullFileName = _T("D:\\��ȡDWG���\\");
	strFullFileName = strFullFileName + chFileName ;
	strFullFileName.TrimRight(_T(".XLS"));
	strFullFileName += _T(".XLS");
	printToFile(strFullFileName, vecTable);
	acutPrintf(_T("\n������ϣ�"));
	
}

void CReadTableToExcel::printToFile(CString strPath, const vector<vector<CString> > vec2Table)
{
	if (vec2Table.size()==0)
	{
		return;
	}
	//д�ļ�
	AcCStdioFile  writeFile;
	BOOL bFile = writeFile.Open(strPath, CFile::modeWrite|CFile::modeCreate);
	if (!bFile)
	{
		return ;
	}

	CString strLine;
	if (bReverse)
	{
		int nRow = vec2Table.size();
		for (int i=0; i<nRow; i++)
		{
			strLine = _T("");
			for (int j=vec2Table.at(0).size()-1; j>=0; j--)
			{
				CString strCell = vec2Table[i][j];
				if (strCell.IsEmpty())
				{
					strCell = _T(" ");
				} 

				strLine += strCell + _T("\t");
			}

			strLine += _T("\n");
			writeFile.WriteString(strLine);
		} 
	}
	else
	{
		int nRow = vec2Table.size();
		for (int i=nRow-1; i>=0; i--)
		{
			strLine = _T("");
			for (int j=0; j<vec2Table.at(0).size(); j++)
			{
				CString strCell = vec2Table[i][j];
				if (strCell.IsEmpty())
				{
					strCell = _T(" ");
				} 

				strLine += strCell + _T("\t");
			}

			strLine += _T("\n");
			writeFile.WriteString(strLine);
		}
	}

	writeFile.Close();
}

bool CReadTableToExcel::readTable(ACHAR * strPrompt, vector<vector<CString> > &vec2Table)
{
	double dTxtHight = 0.0;//���ָ߶�
	double dTxtWidth = 0.0;//���ֿ��
	Acad::ErrorStatus error = Acad::eOk;
	ads_name ssName={0};
	long nLength = 0;
	vector<tDBTextInfor> vecTmpTable;
	vector<vector< vector<tDBTextInfor> > > vec3SortTable;//���յı��
	ACHAR * prompts[2] = {strPrompt, _T("")};
	int rt = acedSSGet(_T(":$:L"), prompts, NULL, NULL, ssName);
	if(RTCAN == rt)
	{
		return false;
	}
	else if(RTNORM == rt)
	{
		rt = acedSSLength(ssName, &nLength);

		//��ӵ�ͼ��������
		for (int i=0;i<nLength;i++)
		{
			ads_name entName;
			int rt = acedSSName(ssName, i, entName);
			if (RTNORM == rt)
			{
				AcDbObjectId id;
				error = acdbGetObjectId(id,entName);
				AcDbEntity *pEnt = NULL;
				if (Acad::eOk == acdbOpenObject(pEnt, id, AcDb::kForRead))
				{
					if (pEnt->isKindOf(AcDbText::desc()))
					{
						AcDbText *pTtextEnt = AcDbText::cast(pEnt);
						dTxtHight = pTtextEnt->height();
						dTxtWidth = pTtextEnt->widthFactor();
						tDBTextInfor textElement;
						textElement.strDBText = pTtextEnt->textString();
						textElement.dX = pTtextEnt->position().x;
						textElement.dY = pTtextEnt->position().y;
						double dRotation = pTtextEnt->rotation();
						if (cos(dRotation) == -1.0)
						{
							bReverse = true;
						}
						else
						{
							bReverse = false;
						}

						vecTmpTable.push_back(textElement);
						pTtextEnt->close();
					}
					else if (pEnt->isKindOf(AcDbLine::desc()))
					{
						AcDbLine *pLine = AcDbLine::cast(pEnt);
						AcGePoint3d startPt = pLine->startPoint();
						AcGePoint3d endPt = pLine->endPoint();
						if (fabs(startPt.x-endPt.x)< 1.0)
						{//����
							m_setCol.insert(startPt.x);
						}
						else if (fabs(startPt.y-endPt.y)<1.0)
						{//����
							m_setRow.insert(startPt.y);
						}
						pLine->close();
					}
					else if (pEnt->isKindOf(AcDbPolyline::desc()))
					{//���
						AcDbPolyline *pPolyLine = AcDbPolyline::cast(pEnt);
						int nCount = pPolyLine->numVerts();
						assert(nCount>1);
						//AcGePoint3d startPt ;
						//AcGePoint3d endPt;
						//pPolyLine->getPointAt(0,startPt);
						//pPolyLine->getPointAt(nCount-1, endPt);
						//double dLength = startPt.distanceTo(endPt);
						//if (dLength<50.0)
						//{
						//	pPolyLine->close();
						//	continue;
						//}

						for (int i=1; i<nCount; i++)
						{
							AcGePoint3d point;
							pPolyLine->getPointAt(i, point);
							AcGePoint3d point1;
							pPolyLine->getPointAt(i-1, point1);
							double dLength = point1.distanceTo(point);
							if (dLength<50.0)
							{
								continue;
							}
							else
							{
								if (i==1)
								{
									m_setCol.insert(point1.x);
									m_setRow.insert(point1.y);
								}
								m_setCol.insert(point.x);
								m_setRow.insert(point.y);
							}

						}

						pPolyLine->close();
					}

					pEnt->close();
				} 

			}
		}

		acedSSFree(ssName);//�ͷ�ѡ��

	}

	sortByTxtCoordinate(vecTmpTable, dTxtHight);

	vector<double> vecXPt;
	vector<double> vecYPt;

	set<double>::iterator iter = m_setCol.begin();
	for (;iter!=m_setCol.end(); iter++)
	{
		if (vecXPt.size()==0)
		{
			vecXPt.push_back(*iter);
			continue;
		}

		double dValue = vecXPt.at(vecXPt.size()-1);
		if (fabs(dValue - *iter)>5.0)//�о���С��5.0��Ϊͬһ��
		{
			vecXPt.push_back(*iter);
		}

	}
	iter = m_setRow.begin();
	for (;iter!= m_setRow.end();iter++)
	{
		vecYPt.push_back(*iter);
	}

	int nCol = vecXPt.size()-1;
	int nRow = vecYPt.size()-1;
	if (nCol<=0 || nRow<=0)
	{
		acutPrintf(_T("\n��ȡ�����л�����С��0��"));
		return false;
	}

	vec2Table.resize(nRow);
	vec3SortTable.resize(nRow);
	for (int i=0;i<nRow;i++)
	{//���ö�ά���������
		vec2Table.at(i).resize(nCol);
		vec3SortTable.at(i).resize(nCol);
		for (int j=0;j<nCol;j++)
		{
			vec2Table[i][j] = _T("");
		}
	}

	int nCount = vecTmpTable.size();
	for (int i=0; i<nCount; i++)
	{
		tDBTextInfor curTextInfor = vecTmpTable.at(i);
		int nFind = 0;
		int m=0;
		for (; m<nCol; m++)
		{//��
			if (curTextInfor.dX>vecXPt[m] && curTextInfor.dX < vecXPt[m+1])
			{
				nFind++;
				break;
			}
		}

		int n=0;
		for (; n<nRow; n++)
		{//��
			if (curTextInfor.dY > vecYPt[n] && curTextInfor.dY < vecYPt[n+1])
			{
				nFind++;
				break;
			}

		}
		if (nFind != 2 )
		{
			continue;
		}

		//vec2Table[n][m] += vecTmpTable.at(i).strDBText;

		//if (!vec2SortTable[n][m].strDBText.IsEmpty())
		//{
		//	vec2SortTable[n][m].dX = vecTmpTable.at(i).dX;
		//	vec2SortTable[n][m].dY = vecTmpTable.at(i).dY;
		//	vec2SortTable[n][m].strDBText = vecTmpTable.at(i).strDBText;
		//	if ()
		//	{
		//	} 
		//	else
		//	{
		//	}
		//} 
		//else
		//{
		//	vec2SortTable[n][m] = vecTmpTable.at(i);
		//}

		vec3SortTable[n][m].push_back(vecTmpTable.at(i));
	}

	nRow = vec3SortTable.size();//��
	for (int i=0; i<nRow; i++)
	{//���ö�ά���������
		nCol = vec3SortTable.at(i).size();//��
		for (int j=0; j<nCol; j++)
		{
			vector<tDBTextInfor> curUnit;
			curUnit = vec3SortTable[i][j];
			sort(curUnit.begin(), curUnit.end(), sortDVTextByDx);
			if (!bReverse)
			{
				std::reverse(curUnit.begin(), curUnit.end());
			}
			CString strUnit;
			for (int n=0; n<curUnit.size(); n++)
			{
				strUnit += curUnit.at(n).strDBText + _T(" ");
			}

			vec2Table[i][j] = strUnit;
		}
	}

	return true;
}

//��ȡ�����ı���������򣬰�xy����ֵ�Ĵ�С����
void CReadTableToExcel::sortByTxtCoordinate(vector<tDBTextInfor> vecTmpTable, double dHight)
{
	double dTolerance = dHight;//���е��ݲ
	int nLength = vecTmpTable.size();
	if (bReverse)
	{//���ߵ�
		//��Y��С��������
		for (int i=0; i<nLength; i++)
		{
			for (int j=i+1;j<nLength;j++)
			{
				if (vecTmpTable.at(i).dY < vecTmpTable.at(j).dY)
				{
					tDBTextInfor tmpTextInfor;
					tmpTextInfor.dX		= vecTmpTable.at(j).dX;
					tmpTextInfor.dY		= vecTmpTable.at(j).dY;
					tmpTextInfor.strDBText = vecTmpTable.at(j).strDBText;

					vecTmpTable.at(j).dX		= vecTmpTable.at(i).dX;
					vecTmpTable.at(j).dY		= vecTmpTable.at(i).dY;
					vecTmpTable.at(j).strDBText = vecTmpTable.at(i).strDBText;

					vecTmpTable.at(i).dX = tmpTextInfor.dX;
					vecTmpTable.at(i).dY = tmpTextInfor.dY;
					vecTmpTable.at(i).strDBText = tmpTextInfor.strDBText;

				} 

			}
		}

		//��X�Ӵ�С����,����������
		for (int i=0; i<nLength; i++)
		{
			for (int j=i+1;j<nLength;j++)
			{
				if (fabs(vecTmpTable.at(i).dY - vecTmpTable.at(j).dY) <= dTolerance)
				{//����Ԫ�ؽ���λ��
					if ((vecTmpTable.at(i).dX > vecTmpTable.at(j).dX))
					{
						tDBTextInfor tmpTextInfor;
						tmpTextInfor.dX		= vecTmpTable.at(j).dX;
						tmpTextInfor.dY		= vecTmpTable.at(j).dY;
						tmpTextInfor.strDBText = vecTmpTable.at(j).strDBText;
						vecTmpTable.at(j).dX		= vecTmpTable.at(i).dX;
						vecTmpTable.at(j).dY		= vecTmpTable.at(i).dY;
						vecTmpTable.at(j).strDBText = vecTmpTable.at(i).strDBText;
						vecTmpTable.at(i).dX = tmpTextInfor.dX;
						vecTmpTable.at(i).dY = tmpTextInfor.dY;
						vecTmpTable.at(i).strDBText = tmpTextInfor.strDBText;
					}
				} 
				else
				{
					break;
				}
			}
		}
	} 
	else
	{//���û�еߵ�
		//��Y�Ӵ�С����
		for (int i=0; i<nLength; i++)
		{
			for (int j=i+1;j<nLength;j++)
			{
				if (vecTmpTable.at(i).dY > vecTmpTable.at(j).dY)
				{
					tDBTextInfor tmpTextInfor;
					tmpTextInfor.dX		= vecTmpTable.at(j).dX;
					tmpTextInfor.dY		= vecTmpTable.at(j).dY;
					tmpTextInfor.strDBText = vecTmpTable.at(j).strDBText;

					vecTmpTable.at(j).dX		= vecTmpTable.at(i).dX;
					vecTmpTable.at(j).dY		= vecTmpTable.at(i).dY;
					vecTmpTable.at(j).strDBText = vecTmpTable.at(i).strDBText;

					vecTmpTable.at(i).dX = tmpTextInfor.dX;
					vecTmpTable.at(i).dY = tmpTextInfor.dY;
					vecTmpTable.at(i).strDBText = tmpTextInfor.strDBText;

				} 

			}
		}

		//��X��С��������,����������
		for (int i=0; i<nLength; i++)
		{
			for (int j=i+1;j<nLength;j++)
			{
				if (fabs(vecTmpTable.at(i).dY - vecTmpTable.at(j).dY) <= dTolerance)
				{//����Ԫ�ؽ���λ��
					if ((vecTmpTable.at(i).dX < vecTmpTable.at(j).dX))
					{
						tDBTextInfor tmpTextInfor;
						tmpTextInfor.dX		= vecTmpTable.at(j).dX;
						tmpTextInfor.dY		= vecTmpTable.at(j).dY;
						tmpTextInfor.strDBText = vecTmpTable.at(j).strDBText;
						vecTmpTable.at(j).dX		= vecTmpTable.at(i).dX;
						vecTmpTable.at(j).dY		= vecTmpTable.at(i).dY;
						vecTmpTable.at(j).strDBText = vecTmpTable.at(i).strDBText;
						vecTmpTable.at(i).dX = tmpTextInfor.dX;
						vecTmpTable.at(i).dY = tmpTextInfor.dY;
						vecTmpTable.at(i).strDBText = tmpTextInfor.strDBText;
					}
				} 
				else
				{
					break;
				}
			}
		}
	}

}
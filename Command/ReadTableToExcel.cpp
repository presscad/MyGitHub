
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
	readTable(_T("\n选择图纸上的表格，自动保存到excel！"), vecTable);
	CreateDirectory(_T("D:\\读取DWG表格\\"), NULL);
	acedInitGet(NULL, NULL);
	ACHAR chFileName[MAX_PATH] = {0};
	acedGetString(0, _T("\n输入保存的文件名称（不用带扩展名）"), chFileName);
	CString strFullFileName = _T("D:\\读取DWG表格\\");
	strFullFileName = strFullFileName + chFileName ;
	strFullFileName.TrimRight(_T(".XLS"));
	strFullFileName += _T(".XLS");
	printToFile(strFullFileName, vecTable);
	acutPrintf(_T("\n保存完毕！"));
	
}

void CReadTableToExcel::printToFile(CString strPath, const vector<vector<CString> > vec2Table)
{
	if (vec2Table.size()==0)
	{
		return;
	}
	//写文件
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
	double dTxtHight = 0.0;//文字高度
	double dTxtWidth = 0.0;//文字宽度
	Acad::ErrorStatus error = Acad::eOk;
	ads_name ssName={0};
	long nLength = 0;
	vector<tDBTextInfor> vecTmpTable;
	vector<vector< vector<tDBTextInfor> > > vec3SortTable;//最终的表格
	ACHAR * prompts[2] = {strPrompt, _T("")};
	int rt = acedSSGet(_T(":$:L"), prompts, NULL, NULL, ssName);
	if(RTCAN == rt)
	{
		return false;
	}
	else if(RTNORM == rt)
	{
		rt = acedSSLength(ssName, &nLength);

		//添加到图层向量中
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
						{//竖线
							m_setCol.insert(startPt.x);
						}
						else if (fabs(startPt.y-endPt.y)<1.0)
						{//横线
							m_setRow.insert(startPt.y);
						}
						pLine->close();
					}
					else if (pEnt->isKindOf(AcDbPolyline::desc()))
					{//外框
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

		acedSSFree(ssName);//释放选择集

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
		if (fabs(dValue - *iter)>5.0)//列距离小于5.0视为同一行
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
		acutPrintf(_T("\n读取表格的行或列数小于0！"));
		return false;
	}

	vec2Table.resize(nRow);
	vec3SortTable.resize(nRow);
	for (int i=0;i<nRow;i++)
	{//设置二维数组的行列
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
		{//列
			if (curTextInfor.dX>vecXPt[m] && curTextInfor.dX < vecXPt[m+1])
			{
				nFind++;
				break;
			}
		}

		int n=0;
		for (; n<nRow; n++)
		{//行
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

	nRow = vec3SortTable.size();//行
	for (int i=0; i<nRow; i++)
	{//设置二维数组的行列
		nCol = vec3SortTable.at(i).size();//列
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

//读取出来的表格内容排序，按xy坐标值的大小排序
void CReadTableToExcel::sortByTxtCoordinate(vector<tDBTextInfor> vecTmpTable, double dHight)
{
	double dTolerance = dHight;//运行的容差；
	int nLength = vecTmpTable.size();
	if (bReverse)
	{//表格颠倒
		//按Y从小到大排列
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

		//按X从大到小排列,及行内排序
		for (int i=0; i<nLength; i++)
		{
			for (int j=i+1;j<nLength;j++)
			{
				if (fabs(vecTmpTable.at(i).dY - vecTmpTable.at(j).dY) <= dTolerance)
				{//两个元素交换位置
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
	{//表格没有颠倒
		//按Y从大到小排列
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

		//按X从小到大排列,及行内排序
		for (int i=0; i<nLength; i++)
		{
			for (int j=i+1;j<nLength;j++)
			{
				if (fabs(vecTmpTable.at(i).dY - vecTmpTable.at(j).dY) <= dTolerance)
				{//两个元素交换位置
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
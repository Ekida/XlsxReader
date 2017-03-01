using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.Text.RegularExpressions;

namespace UnP.myPackage
{
	
	public class XlsxSheet
	{
		XmlNodeList shareStrList;
		XmlDocument sheetDoc= new XmlDocument();
		XmlNodeList xRowList;

		protected Dictionary<string, XlsxCell> activeXlsxCell;

		public delegate void UseRowValue(string[] strs);

		public XlsxSheet(XmlNodeList shareStrList, string sheetXmlDocPath)
		{
			this.shareStrList = shareStrList;
			sheetDoc.Load(sheetXmlDocPath);
			XmlElement root = sheetDoc.DocumentElement;
			xRowList = root.GetElementsByTagName("row");

			activeXlsxCell = new Dictionary<string, XlsxCell>();
		}


		/// <summary>
		/// 获取行中对应所示字符串
		/// </summary>
		/// <param name="index">xml下标</param>
		/// <param name="row">xml行</param>
		/// <returns></returns>
		protected string CellStr(int index, XmlNodeList row) //获取xRow里index下标所示字符串
		{
			if (index >= row.Count || index < 0)
				return "";
			string type = ((XmlElement)row[index]).GetAttribute("t");
			string value = ((XmlElement)row[index]).GetElementsByTagName("v")[0].InnerText;
			if (type == "s")
				return shareStrList[Convert.ToInt32(value)].InnerText;
			else
				return value;
		}

		/// <summary>
		/// 获取Tag为str对应的参数
		/// </summary>
		/// <param name="str">tag</param>
		/// <param name="node">xml节点</param>
		/// <returns></returns>
		protected string GetAttribute(string str, XmlNode node)
		{
			return ((XmlElement)node).GetAttribute(str);
		}
		/// <summary>
		/// 获取列代表字母位于xml文档中的真实下标位置，可为负数
		/// </summary>
		/// <param name="colChar">列代表字母</param>
		/// <param name="row">xml行</param>
		/// <returns></returns>
		protected int GetFirstCol(char colChar, XmlNodeList row)
		{
			string spans = ((XmlElement)row[0]).GetAttribute("r");
			return Convert.ToInt32(colChar - spans[0]);
		}

		/// <summary>
		/// 快速读取xlsx文件
		/// </summary>
		/// <param name="colNums">开始的行数</param>
		/// <param name="userInf">对每一行数据的处理的委托</param>
		public void easyRead(int[] colNums, UseRowValue useValue)
		{
			EasyRead(colNums, 1, xRowList.Count, useValue);
		}

		/// <summary>
		/// 快速读取xlsx文件
		/// </summary>
		/// <param name="colNums">需要读的列数</param>
		/// <param name="beginRow">开始的行数</param>
		/// <param name="endRow">结束的行数</param>
		/// <param name="userInf">对每一行数据的处理的委托</param>
		public void EasyRead(int[] colNums, int beginRow, int endRow, UseRowValue useValue)
		{
			int beginCol;
			int readCellNum = colNums.Length;
			string[] strs = new string[readCellNum];
			for (int i = beginRow - 1; i < endRow; i++)
			{
				XmlNodeList xRow = ((XmlElement)xRowList[i]).GetElementsByTagName("c");
				beginCol = GetFirstCol('A', xRow);
				for (int resultIndex = 0; resultIndex < readCellNum; resultIndex++)
				{
					strs[resultIndex] = CellStr(colNums[resultIndex] - beginCol - 1, xRow);
				}
				useValue(strs);
			}
		}

		/// <summary>
		/// 获取相应下标的单元格字符串
		/// </summary>
		/// <param name="rowNum"></param>
		/// <param name="colNum"></param>
		/// <returns></returns>
		public string getCellStr(int rowNum, int colNum)
		{
			XmlNodeList xRow = ((XmlElement)xRowList[rowNum - 1]).GetElementsByTagName("c");
			int beginCol = GetFirstCol('A', xRow);
			return CellStr(colNum - 1,xRow);
		}
		public string getCellStr(int rowNum, string colStr)
		{
			return getCellStr(rowNum, colStr[0] - 'A' + 1);
		}
		public string getCellStr(string indexStr)
		{
			int rowNum = Convert.ToInt32(Regex.Replace(indexStr, @"[^0-9]+", ""));
			string colChar = Regex.Replace(indexStr, @"[^A-Z]+", "");
			return getCellStr(rowNum, colChar);
		}
		public string this[int rowNum, int colNum]
		{
			get
			{
				return getCellStr(rowNum, colNum);
			}
		}
		public string this[string indexStr]
		{
			get
			{
				return getCellStr(indexStr);
			}
		}

		/// <summary>
		/// 获取单元格
		/// </summary>
		/// <param name="rowNum"></param>
		/// <param name="colStr"></param>
		/// <returns></returns>
		public XlsxCell getCell(int rowNum, string colStr)
		{
			var key = colStr + rowNum.ToString();
			XlsxCell result;
			if (!activeXlsxCell.ContainsKey(key))
			{
				result = new XlsxCell();
				activeXlsxCell.Add(colStr + rowNum.ToString(), result);
			}
			else
			{
				result = activeXlsxCell[key];
			}
			return new XlsxCell();
		}
	}
}

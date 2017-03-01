using ICSharpCode.SharpZipLib.Zip;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;
using UnP.myPackage;

namespace UnP
{


	public class XlsxDoctment
	{

		protected XmlNodeList shareStrList;
		protected string filePath;
		protected string tempDirector;
		protected bool isOpen = false;

	
		/// <summary>
		/// 构造函数，选择需要读取的文件
		/// </summary>
		/// <param name="filePath">文件路径</param>
		public XlsxDoctment(string path)
		{
			if(!File.Exists(path))
				throw new FileNotFoundException("File not Found", path);
			filePath = path;
			tempDirector = Path.GetDirectoryName(filePath) + "\\~$" + Path.GetFileName(filePath);
		}

		/// <summary>
		/// 打开文件
		/// </summary>
		public void Open()
		{
			if (isOpen)
				return;	
			if (!Directory.Exists(tempDirector))
				Directory.CreateDirectory(tempDirector);
			(new FastZip()).ExtractZip(filePath, tempDirector, "");
			XmlDocument doc = new XmlDocument();
			doc.Load(tempDirector + @"\xl\sharedStrings.xml");
			XmlElement root = doc.DocumentElement;
			shareStrList = root.GetElementsByTagName("si");

			isOpen = true;
		}
		/// <summary>
		/// 关闭文件
		/// </summary>
		public void Close()
		{

			if (Directory.Exists(tempDirector) && isOpen)
				Directory.Delete(tempDirector, true);
			shareStrList = null;
			isOpen = false;
		}
		/// <summary>
		/// 获取工作表
		/// </summary>
		/// <param name="sheetNum">工作表编号</param>
		/// <returns></returns>
		public XlsxSheet GetSheet(int sheetNum)
		{
			if (!File.Exists(tempDirector + "\\xl\\worksheets\\sheet" + sheetNum + ".xml"))
				return null;
			return new XlsxSheet(shareStrList, tempDirector + "\\xl\\worksheets\\sheet" + sheetNum + ".xml");
		}



	}
}

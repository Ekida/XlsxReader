using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace UnP
{
	class Program
	{
		static void Main(string[] args)
		{
			XlsxDoctment doc = new XlsxDoctment(args[0]);
			doc.Open();
			var sheet = doc.GetSheet(1);
			int[] index = { 2, 4 };
			sheet.easyRead(index, (str) =>
			 {
				 Console.WriteLine(str[0] + " and " + str[1]);
			 }
			);
			doc.Close();
			Console.ReadKey();	
		}
	}
}

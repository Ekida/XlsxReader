using System;
using System.Collections.Generic;
using System.Text;

namespace UnP.myPackage
{
	public class XlsxCell
	{
		public enum cellType
		{

		}

		protected cellType cType;
		protected string value;
		protected string func;
		protected string shareString;

		public bool IsChange { get; protected set; }
		public string Value
		{
			get
			{
				if (shareString == "")
					return value;
				else
					return shareString;
			}
			set
			{
				this.value = value;
				IsChange = true;
			}
		}
		public string Func
		{
			get
			{
				return func;
			}
			set
			{
				this.func = value;
				IsChange = true; 
			}
		}
		public XlsxCell()
		{
			IsChange = false;


		}

	}

}

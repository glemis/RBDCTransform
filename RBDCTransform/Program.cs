using System;

namespace RBDCTransform
{
	class Program
	{
		[STAThread]
		static void Main(string[] args)
		{
			TransformFunctions handler = new TransformFunctions();
			handler.CreateMembershipImport();
			//handler.CreateFinancialReport();
		}
	}
}


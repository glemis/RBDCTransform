using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace RBDCTransform
{
	class TransformFunctions
	{

		public void CreateMembershipImport()
		{
			string path = "";
			OpenFileDialog file = new OpenFileDialog();
			if (file.ShowDialog() == DialogResult.OK)
			{
				path = file.FileName;
			}
			var reader = new StreamReader(File.OpenRead(path));
			List<string> Headings = new List<string>();
			List<List<string>> Rows = new List<List<string>>();

			bool isHeadings = true;
			while (!reader.EndOfStream)
			{
				var line = reader.ReadLine();
				var values = line.Split(',');

				if (isHeadings)
				{
					Headings.AddRange(values);
					isHeadings = false;
				}
				else
				{
					Rows.Add(new List<string>(values));
				}
			}

			List<List<string>> RowsToRemove = new List<List<string>>();
			var RefundedRows = Rows.Where(r => r[2] == "\"refunded\"");

			for (int k = 0; k < RefundedRows.Count(); k++)
			{
				RowsToRemove.AddRange(Rows.Where(r => RefundedRows.ElementAt(k)[0] == r[0]));
			}

			for (int i = 0; i < RowsToRemove.Count(); i++)
			{
				Rows.IndexOf(RowsToRemove.ElementAt(i));
				Rows.Remove(RowsToRemove.ElementAt(i));
			}

			//Columns to remove    2,3,4,5,6,7,8,9,10,11,12,13
			//var columnsToRemove = new List<int>();
			//columnsToRemove.Add(Headings.IndexOf("\"Product Form: New\""));
			//columnsToRemove.Add(Headings.IndexOf("\"Checkout Form: Note / Additional Info\""));
			//columnsToRemove.Add(Headings.IndexOf("\"Product Form: Liability Waiver\""));
			//columnsToRemove.Add(Headings.IndexOf("\"Lineitem requires shipping\""));

			Headings.RemoveAt(57);
			Headings.RemoveAt(56);
			Headings.RemoveAt(55);
			Headings.RemoveRange(21, 20);
			Headings.RemoveRange(16, 4);
			Headings.RemoveAt(17);
			Headings.RemoveRange(2, 13);
			Headings.RemoveAt(1);
			for (var i = 0; i < Headings.Count(); i++)
			{
				Headings[i] = Headings[i].Replace("Product Form: ", "");
				Headings[i] = Headings[i].Replace("Postal", "PSTL");
				Headings[i] = Headings[i].Replace("Postal", "PSTL");
				Headings[i] = Headings[i].Replace("New Member", "New");
				Headings[i] = Headings[i].Replace("Email", "E-Mail");
				Headings[i] = Headings[i].Replace("Created at", "Created On");
				Headings[i] = Headings[i].Replace("Phone", "Home");
				Headings[i] = Headings[i].Replace("Lineitem variant", "SEMESTER");

				Console.WriteLine(i + " " + Headings[i]);
			}
			Headings.Add("MEMBER TYPE");


			//Rows at 20   "\"Fall and Winter/Non-Student\""
			for (var i = 0; i < Rows.Count(); i++)
			{
				Rows[i].RemoveAt(57);
				Rows[i].RemoveAt(56);
				Rows[i].RemoveAt(55);
				Rows[i].RemoveRange(21, 20);
				Rows[i].RemoveRange(16, 4);
				Rows[i].RemoveAt(17);
				Rows[i].RemoveRange(2, 13);
				Rows[i].RemoveAt(1);
				Rows[i][7] = Rows[i][7].Replace("\"Male\"", "M");
				Rows[i][7] = Rows[i][7].Replace("\"Female\"", "F");
				var Varient = Rows[i][2].Split('/')[0];
				var Membership = Rows[i][2].Split('/')[1];
				if (Varient.Contains("and"))
				{
					Rows[i][2] = "\"Year\"";
				}
				else if (Varient.Contains("Fall"))
				{
					Rows[i][2] = "\"Fall\"";
				}
				else
				{
					Rows[i][2] = "\"Winter\"";
				}

				if (Membership.Contains("Non-Student") || Membership.Contains("Standard Price"))
				{
					Rows[i].Add("\"Member\"");
				}
				else
				{
					Rows[i].Add("\"Student\"");
				}
			}


			string delimter = ",";
			var savePath = path.TrimEnd(path.Split('\\').Last().ToCharArray()) + "Squarespace_Orders_" + DateTime.Now.ToString("yyyy_MM_dd") + ".csv";
			if (!File.Exists(savePath))
			{
				File.Create(savePath).Close();
			}

			List<string[]> output = new List<string[]>();
			output.Add(Headings.ToArray());
			for (var i = 0; i < Rows.Count(); i++)
			{
				output.Add(Rows[i].ToArray());
			}
			int length = output.Count;

			using (System.IO.TextWriter writer = File.CreateText(savePath))
			{
				for (int index = 0; index < length; index++)
				{
					writer.WriteLine(string.Join(delimter, output[index]));
				}
			}
		}

		public void CreateFinancialReport()
		{
			string path = "";
			OpenFileDialog file = new OpenFileDialog();
			if (file.ShowDialog() == DialogResult.OK)
			{
				path = file.FileName;
			}
			var reader = new StreamReader(File.OpenRead(path));
			List<string> Headings = new List<string>();
			List<List<string>> Rows = new List<List<string>>();

			bool isHeadings = true;
			while (!reader.EndOfStream)
			{
				var line = reader.ReadLine();
				var values = line.Split(',');

				if (isHeadings)
				{
					Headings.AddRange(values);
					isHeadings = false;
				}
				else
				{
					Rows.Add(new List<string>(values));
				}
			}
			//Columns to remove    2,3,4,5,6,7,8,9,10,11,12,13
			Headings.RemoveAt(41);
			Headings.RemoveAt(40);
			Headings.RemoveRange(32, 7);
			Headings.RemoveAt(22);
			Headings.RemoveAt(21);
			Headings.RemoveAt(14);
			Headings.RemoveAt(8);
			Headings.RemoveAt(7);
			Headings.RemoveAt(5);
			Headings.RemoveAt(4);

			for (var i = 0; i < Headings.Count(); i++)
			{
				//Headings[i] = Headings[i].Replace("Product Form: ", "");

				Console.WriteLine(i + " " + Headings[i]);
			}


			//Rows at 20   "\"Fall and Winter/Non-Student\""
			for (var i = 0; i < Rows.Count(); i++)
			{
				Rows[i].RemoveAt(41);
				Rows[i].RemoveAt(40);
				Rows[i].RemoveRange(32, 7);
				Rows[i].RemoveAt(22);
				Rows[i].RemoveAt(21);
				Rows[i].RemoveAt(14);
				Rows[i].RemoveAt(8);
				Rows[i].RemoveAt(7);
				Rows[i].RemoveAt(5);
				Rows[i].RemoveAt(4);
			}

			List<string> ReorderedHeadings = new List<string>();
			List<List<string>> ReorderedRows = new List<List<string>>();

			int[] order = { 0, 2, 3, 4, 5, 14, 12, 15, 11, 13, 16, 8, 9, 7, 6, 10, 1, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27 };

			for (int i = 0; i < order.Length; i++)
			{
				ReorderedHeadings.Add(Headings[order[i]]);
			}

			for (var k = 0; k < Rows.Count(); k++)
			{
				List<string> reorderedRow = new List<string>();

				for (int i = 0; i < order.Length; i++)
				{
					reorderedRow.Add(Rows[k][order[i]]);
				}

				ReorderedRows.Add(reorderedRow);
			}
			ReorderedHeadings.Insert(14, "Fee Percent");
			ReorderedHeadings.Insert(15, "Fee Charge");
			ReorderedHeadings.Insert(16, "Net Profit");

			ReorderedHeadings.Insert(10, "Fee Percent");
			ReorderedHeadings.Insert(11, "Line item Fee Charge");
			ReorderedHeadings.Insert(12, "Line item Net Profit");
			ReorderedHeadings.RemoveAt(4);

			for (var k = 0; k < ReorderedRows.Count(); k++)
			{
				if (ReorderedRows[k][13] == "\"\"" || ReorderedRows[k][13] == "")
				{
					ReorderedRows[k].Insert(14, "3.2%");
					ReorderedRows[k].Insert(15, "$0.00");
					ReorderedRows[k].Insert(16, "$0.00");
				}
				else
				{
					Double Total = Double.Parse(ReorderedRows[k][13].Trim('"'));
					Double Fee = Total * 0.032;
					Double Net = Total - Fee;

					ReorderedRows[k].Insert(14, "3.2%");
					ReorderedRows[k].Insert(15, Fee.ToString("$0.00"));
					ReorderedRows[k].Insert(16, Net.ToString("$0.00"));
				}

				if (ReorderedRows[k][9] == "\"\"" || ReorderedRows[k][9] == "")
				{
					ReorderedRows[k].Insert(10, "3.2%");
					ReorderedRows[k].Insert(11, "$0.00");
					ReorderedRows[k].Insert(12, "$0.00");
					ReorderedRows[k].RemoveAt(4);
				}
				else
				{
					Double LineTotal = Double.Parse(ReorderedRows[k][9].Trim('"'));
					Double LineQuantity = Double.Parse(ReorderedRows[k][8].Trim('"'));
					Double LineFee = (LineTotal * LineQuantity) * 0.032;
					Double LineNet = LineTotal - LineFee;

					ReorderedRows[k].Insert(10, "3.2%");
					ReorderedRows[k].Insert(11, LineFee.ToString("$0.00"));
					ReorderedRows[k].Insert(12, LineNet.ToString("$0.00"));
					
				}

				ReorderedRows[k].RemoveAt(4);
			}

			string delimter = ",";
			var savePath = path.TrimEnd(path.Split('\\').Last().ToCharArray()) + "Squarespace_Financial_Report_" + DateTime.Now.ToString("yyyy_MM_dd") + ".csv";
			if (!File.Exists(savePath))
			{
				File.Create(savePath).Close();
			}

			List<string[]> output = new List<string[]>();
			output.Add(ReorderedHeadings.ToArray());
			for (var i = 0; i < ReorderedRows.Count(); i++)
			{
				output.Add(ReorderedRows[i].ToArray());
			}
			int length = output.Count;

			using (System.IO.TextWriter writer = File.CreateText(savePath))
			{
				for (int index = 0; index < length; index++)
				{
					writer.WriteLine(string.Join(delimter, output[index]));
				}
			}
		}
	}
}

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using LinqToExcel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using Cell = DocumentFormat.OpenXml.Spreadsheet.Cell;
using Row = DocumentFormat.OpenXml.Spreadsheet.Row;

namespace HomeSchoolNetwork {

	class Program {

		static void Main(string[] args) {

			DisplayInstructions();

			FileInfo csvFileInfo = null;
			var directory = @"W:\HSN\";
			var d = new DirectoryInfo(directory);
			var files = d.GetFiles("*.csv"); 
			foreach (var file in files) {
				csvFileInfo = file;
				Console.WriteLine(file.Name);
			}
			if (csvFileInfo == null) return;

			var sheetName = "";
			var excelFile = new ExcelQueryFactory(csvFileInfo.FullName);
			AddMapping(excelFile);

			var registrations = excelFile.Worksheet<Registration>(sheetName).ToList().OrderBy(x => x.ParentLastName).ThenBy(x => x.UserId).ThenBy(x => x.BirthDate).ToList();

			var filePath = $"{directory}HSNMergeFormat.xlsx";
			using (var package = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook)) {
				var workbookPart1 = package.AddWorkbookPart();
				var workbook1 = new Workbook();
				workbook1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

				var sheets1 = new Sheets();
				var sheet1 = new Sheet { Name = "Sheet1", SheetId = 1U, Id = "rId1" };
				sheets1.Append(sheet1);

				workbook1.Append(sheets1);
				workbookPart1.Workbook = workbook1;

				var worksheetPart = workbookPart1.AddNewPart<WorksheetPart>("rId1");
				var worksheet = new Worksheet();
				var sheetData = new SheetData();

				sheetData.Append(CreateHeader());

				var total = registrations.Count;
				var parentCount = 1;
				for (var i = 0; i < total; i++) {
					var reg = registrations[i];
					var userId = reg.UserId;
					parentCount++;
					var parent = AddParent(parentCount, reg);
					AddChild(parent, 1, parentCount, reg);
					if (i + 1 < total && userId == registrations[i + 1].UserId) {
						i++;
						AddChild(parent, 2, parentCount, registrations[i]); // Child 2
						if (i + 1 < total && userId == registrations[i + 1].UserId) {
							i++;
							AddChild(parent, 3, parentCount, registrations[i]); // Child 3
							if (i + 1 < total && userId == registrations[i + 1].UserId) {
								i++;
								AddChild(parent, 4, parentCount, registrations[i]); // Child 4
								if (i + 1 < total && userId == registrations[i + 1].UserId) {
									i++;
									AddChild(parent, 5, parentCount, registrations[i]); // Child 5
									if (i + 1 < total && userId == registrations[i + 1].UserId) {
										i++;
										AddChild(parent, 6, parentCount, registrations[i]); // Child 6
										if (i + 1 < total && userId == registrations[i + 1].UserId) {
											i++;
											AddChild(parent, 7, parentCount, registrations[i]); // Child 7
											if (i + 1 < total && userId == registrations[i + 1].UserId) {
												i++;
												AddChild(parent, 8, parentCount, registrations[i]); // Child 8
											}
										}
									}
								}
							}
						}
					}
					sheetData.Append(parent);
				}
				worksheet.Append(sheetData);
				worksheetPart.Worksheet = worksheet;
			}
			Console.WriteLine("Press any key to finish...");
			Console.ReadKey();
		}

		private static void DisplayInstructions() {
			Console.WriteLine("Ensure downloaded CSV file is in the HSN folder then press ENTER...");
			Console.ReadKey();
		}

		private static Row AddParent(int rowCount, Registration reg) {
			Console.WriteLine($"{reg.UserId} - {reg.ParentLastName} - {reg.ParentFirstName}");
			var row = new Row();
			row.Append(CreateCell("A", rowCount, reg.ParentLastName));
			row.Append(CreateCell("B", rowCount, reg.ParentFirstName));
			row.Append(CreateCell("C", rowCount, reg.Address1));
			row.Append(CreateCell("D", rowCount, reg.City));
			row.Append(CreateCell("E", rowCount, reg.State));
			row.Append(CreateCell("F", rowCount, reg.Email));
			row.Append(CreateCell("G", rowCount, reg.Zip));
			row.Append(CreateCell("H", rowCount, reg.Phone));
			row.Append(CreateCell("I", rowCount, reg.Cell));
			row.Append(CreateCell("J", rowCount, reg.SpouseFirstName));
			row.Append(CreateCell("K", rowCount, reg.SpouseLastName));
			row.Append(CreateCell("L", rowCount, reg.AgreeToFundraising));
			row.Append(CreateCell("M", rowCount, reg.AgreeToFundraisingHighOnly));
			row.Append(CreateCell("N", rowCount, reg.AgreeToFundraisingCovered));
			row.Append(CreateCell("O", rowCount, reg.AgreeToFundraisingSubstitute));
			row.Append(CreateCell("P", rowCount, reg.AgreeToNameBadgesWorn));
			row.Append(CreateCell("Q", rowCount, reg.AgreeToSupervision));
			row.Append(CreateCell("R", rowCount, reg.AgreeToIAmResponsible));
			return row;
		}

		private static void AddChild(Row parent, int child, int row, Registration reg) {
			Console.WriteLine($"    {reg.UserId} - {reg.StudentLastName} - {reg.StudentFirstName} - {reg.BirthDate}");
			var col1 = "";
			var col2 = "";
			var col3 = "";
			var col4 = "";
			var col5 = "";
			var col6 = "";
			switch (child) {
				default:
					col1 = "S";
					col2 = "T";
					col3 = "U";
					col4 = "V";
					col5 = "W";
					col6 = "X";
					break;
				case 2:
					col1 = "Y";
					col2 = "Z";
					col3 = "AA";
					col4 = "AB";
					col5 = "AC";
					col6 = "AD";
					break;
				case 3:
					col1 = "AE";
					col2 = "AF";
					col3 = "AG";
					col4 = "AH";
					col5 = "AI";
					col6 = "AJ";
					break;
				case 4:
					col1 = "AK";
					col2 = "AL";
					col3 = "AM";
					col4 = "AN";
					col5 = "AO";
					col6 = "AP";
					break;
				case 5:
					col1 = "AQ";
					col2 = "AR";
					col3 = "AS";
					col4 = "AT";
					col5 = "AU";
					col6 = "AV";
					break;
				case 6:
					col1 = "AW";
					col2 = "AX";
					col3 = "AY";
					col4 = "AZ";
					col5 = "BA";
					col6 = "BB";
					break;
				case 7:
					col1 = "BC";
					col2 = "BD";
					col3 = "BE";
					col4 = "BF";
					col5 = "BG";
					col6 = "BH";
					break;
				case 8:
					col1 = "BI";
					col2 = "BJ";
					col3 = "BK";
					col4 = "BL";
					col5 = "BM";
					col6 = "BN";
					break;
			}
			parent.Append(CreateCell(col1, row, reg.StudentFirstName));
			parent.Append(CreateCell(col2, row, reg.StudentLastName));
			parent.Append(CreateCell(col3, row, reg.BirthDate.ToShortDateString()));
			parent.Append(CreateCell(col4, row, reg.SpecialNeeds));
			parent.Append(CreateCell(col5, row, reg.HoursAttending));
			parent.Append(CreateCell(col6, row, reg.CurrentGrade));
		}

		private static Row CreateHeader() {
			var row = new Row();
			row.Append(CreateCell("A", 1, "Parent Name"));
			row.Append(CreateCell("B", 1, "Last"));
			row.Append(CreateCell("C", 1, "Address"));
			row.Append(CreateCell("D", 1, "City"));
			row.Append(CreateCell("E", 1, "State"));
			row.Append(CreateCell("F", 1, "Email"));
			row.Append(CreateCell("G", 1, "Zip Code"));
			row.Append(CreateCell("H", 1, "Phone"));
			row.Append(CreateCell("I", 1, "Mobile Phone"));
			row.Append(CreateCell("J", 1, "Spouse Name"));
			row.Append(CreateCell("K", 1, "Last"));
			row.Append(CreateCell("L", 1, "Fundraising"));
			row.Append(CreateCell("M", 1, "Fundraising for High Schooler"));
			row.Append(CreateCell("N", 1, "If I am unable to be present to fill my teacher/helper position, I will make arrangements to be covered and notify my area coordinator in advance."));
			row.Append(CreateCell("O", 1, "If I am unable to be present for my cleaning/fundraising dates I will find someone who will trade slots with me and notify the coordinator in advance."));
			row.Append(CreateCell("P", 1, "Name badges are to be on at all times. (If you forget yours get a stick-on one from the administration) This applies to BOTH adults and students. If they need to be replaced, there will be an additional fee."));
			row.Append(CreateCell("Q", 1, "I understand that I am to remain on the premises and inside the building while my child/children are attending HSN activities. If I need to leave for an emergency, I will notify and make arrangements with the Leadership Team.  If I need to leave for a non-emergency reason I will notify the Leadership Team and take my children with me."));
			row.Append(CreateCell("R", 1, "I understand when my child is not in a supervised classroom I am responsible for them. This includes in the morning before classes, at lunch and after classes."));

			// Student 1
			row.Append(CreateCell("S", 1, "Student Name"));
			row.Append(CreateCell("T", 1, "Last"));
			row.Append(CreateCell("U", 1, "Date of Birth"));
			row.Append(CreateCell("V", 1, "Allergies / Special Needs"));
			row.Append(CreateCell("W", 1, "Hours Attending"));
			row.Append(CreateCell("X", 1, "Current Grade"));

			// Student 2
			row.Append(CreateCell("Y", 1, "Student Name"));
			row.Append(CreateCell("Z", 1, "Last"));
			row.Append(CreateCell("AA", 1, "Date of Birth"));
			row.Append(CreateCell("AB", 1, "Allergies / Special Needs"));
			row.Append(CreateCell("AC", 1, "Hours Attending"));
			row.Append(CreateCell("AD", 1, "Current Grade"));

			// Student 3
			row.Append(CreateCell("AE", 1, "Student Name"));
			row.Append(CreateCell("AF", 1, "Last"));
			row.Append(CreateCell("AG", 1, "Date of Birth"));
			row.Append(CreateCell("AH", 1, "Allergies / Special Needs"));
			row.Append(CreateCell("AI", 1, "Hours Attending"));
			row.Append(CreateCell("AJ", 1, "Current Grade"));

			// Student 4
			row.Append(CreateCell("AK", 1, "Student Name"));
			row.Append(CreateCell("AL", 1, "Last"));
			row.Append(CreateCell("AM", 1, "Date of Birth"));
			row.Append(CreateCell("AN", 1, "Allergies / Special Needs"));
			row.Append(CreateCell("AO", 1, "Hours Attending"));
			row.Append(CreateCell("AP", 1, "Current Grade"));

			// Student 5
			row.Append(CreateCell("AQ", 1, "Student Name"));
			row.Append(CreateCell("AR", 1, "Last"));
			row.Append(CreateCell("AS", 1, "Date of Birth"));
			row.Append(CreateCell("AT", 1, "Allergies / Special Needs"));
			row.Append(CreateCell("AU", 1, "Hours Attending"));
			row.Append(CreateCell("AV", 1, "Current Grade"));

			// Student 6
			row.Append(CreateCell("AW", 1, "Student Name"));
			row.Append(CreateCell("AX", 1, "Last"));
			row.Append(CreateCell("AY", 1, "Date of Birth"));
			row.Append(CreateCell("AZ", 1, "Allergies / Special Needs"));
			row.Append(CreateCell("BA", 1, "Hours Attending"));
			row.Append(CreateCell("BB", 1, "Current Grade"));

			// Student 7
			row.Append(CreateCell("BC", 1, "Student Name"));
			row.Append(CreateCell("BD", 1, "Last"));
			row.Append(CreateCell("BE", 1, "Date of Birth"));
			row.Append(CreateCell("BF", 1, "Allergies / Special Needs"));
			row.Append(CreateCell("BG", 1, "Hours Attending"));
			row.Append(CreateCell("BH", 1, "Current Grade"));

			// Student 8
			row.Append(CreateCell("BI", 1, "Student Name"));
			row.Append(CreateCell("BJ", 1, "Last"));
			row.Append(CreateCell("BK", 1, "Date of Birth"));
			row.Append(CreateCell("BL", 1, "Allergies / Special Needs"));
			row.Append(CreateCell("BM", 1, "Hours Attending"));
			row.Append(CreateCell("BN", 1, "Current Grade"));

			return row;
		}

		private static void AddMapping(ExcelQueryFactory excelFile) {
			excelFile.AddMapping("UserId", "User ID");
			excelFile.AddMapping("StudentFirstName", "Student Name");
			excelFile.AddMapping("StudentLastName", "Student Last");
			excelFile.AddMapping("BirthDate", "Date of Birth");
			excelFile.AddMapping("SpecialNeeds", "Allergies / Special Needs");
			excelFile.AddMapping("HoursAttending", "Hours Attending");
			excelFile.AddMapping("CurrentGrade", "Current Grade, they will be entering");
			excelFile.AddMapping("ClassChoice9am", "9:00 am class choice");
			excelFile.AddMapping("ClassChoice9amOther", "Other 9am choice");
			excelFile.AddMapping("ClassChoice10am", "10:00 am class choice");
			excelFile.AddMapping("ClassChoice10amOther", "Other 10am choice");
			excelFile.AddMapping("ClassChoice11am", "11:00 am class choice");
			excelFile.AddMapping("ClassChoice11amOther", "Other 11am choice");
			excelFile.AddMapping("ClassChoice1pm", "1:00 pm class choice");
			excelFile.AddMapping("ClassChoice1pmOther", "Other 1pm choice");
			excelFile.AddMapping("ClassChoice2pm", "2:00 pm class choice");
			excelFile.AddMapping("ClassChoice2pmOther", "Other 2pm choice");
			excelFile.AddMapping("ParentFirstName", "Parent Name");
			excelFile.AddMapping("ParentLastName", "Parent Last");
			excelFile.AddMapping("Address1", "Address");
			excelFile.AddMapping("Address2", "Address Line 2");
			excelFile.AddMapping("City", "City");
			excelFile.AddMapping("State", "State");
			excelFile.AddMapping("Email", "Email");
			excelFile.AddMapping("Zip", "Zip Code");
			excelFile.AddMapping("Phone", "Phone");
			excelFile.AddMapping("Cell", "Mobile Phone");
			excelFile.AddMapping("SpouseFirstName", "Spouse Name");
			excelFile.AddMapping("SpouseLastName", "Spouse Last");
			excelFile.AddMapping("AgreeToFundraising", "Fundraising");
			excelFile.AddMapping("AgreeToFundraisingHighOnly", "Fundraising for Jr High and High School Students Only");
			excelFile.AddMapping("AgreeToFundraisingCovered", "Make Arrangements");
			excelFile.AddMapping("AgreeToFundraisingSubstitute", "Fundraising Replacement");
			excelFile.AddMapping("AgreeToNameBadgesWorn", "Name Badges");
			excelFile.AddMapping("AgreeToSupervision", "Remain on the Premises");
			excelFile.AddMapping("AgreeToIAmResponsible", "Child Supervision");
			excelFile.AddMapping("Date", "Today's Date");
			excelFile.AddMapping("TimeStamp", "Timestamp");
			excelFile.AddMapping("LastUpdated", "Last Updated");
			excelFile.AddMapping("CreatedBy", "Created By");
			excelFile.AddMapping("UpdatedBy", "Updated By");
			excelFile.AddMapping("Draft", "Draft");
			excelFile.AddMapping("Ip", "IP");
			excelFile.AddMapping("Id", "ID");
			excelFile.AddMapping("Key", "Key");
		}

		private static Cell CreateCell(string col, int row, string value) {
			var cell = new Cell { CellReference = $"{col}{row}", DataType = CellValues.InlineString };
			var inlineString1 = new InlineString();
			inlineString1.Append(new Text { Text = value });
			cell.Append(inlineString1);
			return cell;
		}


	}

}

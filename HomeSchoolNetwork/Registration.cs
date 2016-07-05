using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HomeSchoolNetwork {

	public class Registration {

		private string parentLastName;
		private string parentFirstName;
		private string address1;
		private string address2;
		private string city;
		private string state;
		private string phone;
		private string cell;
		private string spouseFirstName;
		private string spouseLastName;
		private string studentFirstName;
		private string studentLastName;

		public string UserId { get; set; }

		public string StudentFirstName {
			get { return studentFirstName.ToTitleCase(); }
			set { studentFirstName = value; }
		}

		public string StudentLastName {
			get { return studentLastName.ToTitleCase(); }
			set { studentLastName = value; }
		}

		public DateTime BirthDate { get; set; }

		public string SpecialNeeds { get; set; }

		public string HoursAttending { get; set; }

		public string CurrentGrade { get; set; }

		public string ClassChoice9am { get; set; }

		public string ClassChoice9amOther { get; set; }

		public string ClassChoice10am { get; set; }

		public string ClassChoice10amOther { get; set; }

		public string ClassChoice11am { get; set; }

		public string ClassChoice11amOther { get; set; }

		public string ClassChoice1pm { get; set; }

		public string ClassChoice1pmOther { get; set; }

		public string ClassChoice2pm { get; set; }

		public string ClassChoice2pmOther { get; set; }

		public string ParentFirstName {
			get { return parentFirstName.ToTitleCase(); }
			set { parentFirstName = value; }
		}

		public string ParentLastName {
			get { return parentLastName.ToTitleCase(); }
			set { parentLastName = value; }
		}

		public string Address1 {
			get { return address1.ToTitleCase(); }
			set { address1 = value; }
		}

		public string Address2 {
			get { return address2.ToTitleCase(); }
			set { address2 = value; }
		}

		public string City {
			get { return city.ToTitleCase(); }
			set { city = value; }
		}

		public string State {
			get {
				if (string.IsNullOrEmpty(state)) state = "MO";
				state = state.ToUpper().Replace(".", "");
				if (state.StartsWith("M")) state = "MO";
				if (state.StartsWith("K")) state = "KS";
				if (state.StartsWith("O")) state = "OK";
				return state;
			}
			set { state = value; }
		}

		public string Email { get; set; }

		public string Zip { get; set; }

		public string Phone {
			get {
				if (string.IsNullOrEmpty(phone)) return phone;
				phone = phone.Replace("(", "").Replace(")", "").Replace("-", "").Replace(" ", "");
				if (phone.Length == 10) phone = phone.Substring(0, 3) + "-" + phone.Substring(3, 3) + "-" + phone.Substring(6);
				return phone;
			}
			set { phone = value; }
		}

		public string Cell {
			get {
				if (string.IsNullOrEmpty(cell)) return cell;
				cell = cell.Replace("(", "").Replace(")", "").Replace("-", "").Replace(" ", "");
				if (cell.Length == 10) cell = cell.Substring(0, 3) + "-" + cell.Substring(3, 3) + "-" + cell.Substring(6);
				return cell;
			}
			set { cell = value; }
		}

		public string SpouseFirstName {
			get { return spouseFirstName.ToTitleCase(); }
			set { spouseFirstName = value; }
		}

		public string SpouseLastName {
			get { return spouseLastName.ToTitleCase(); }
			set { spouseLastName = value; }
		}

		public string AgreeToFundraising { get; set; }

		public string AgreeToFundraisingHighOnly { get; set; }

		public string AgreeToFundraisingCovered { get; set; }

		public string AgreeToFundraisingSubstitute { get; set; }

		public string AgreeToNameBadgesWorn { get; set; }

		public string AgreeToSupervision { get; set; }

		public string AgreeToIAmResponsible { get; set; }

		public string Date { get; set; }

		public string TimeStamp { get; set; }

		public string LastUpdated { get; set; }

		public string CreatedBy { get; set; }

		public string UpdatedBy { get; set; }

		public string Draft { get; set; }

		public string Ip { get; set; }

		public string Id { get; set; }

		public string Key { get; set; }

	}

	public static class RegistrationExtensions {

		public static string ToTitleCase(this string myString) {
			if (!string.IsNullOrEmpty(myString)) {
				myString = System.Threading.Thread.CurrentThread.CurrentCulture.TextInfo.ToTitleCase(myString.ToLower());
				if (myString.StartsWith("Mc")) {
					var x = myString[2].ToString().ToUpper();
					var xx = myString.Substring(3);
					myString = $"Mc{x}{xx}";
				}
			}
			return myString;
		}

	}

}

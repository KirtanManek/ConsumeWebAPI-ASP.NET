using ClosedXML.Excel;
using ConsumeWebAPI.DAL;
using ConsumeWebAPI.Models;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;

namespace ConsumeWebAPI.Controllers
{
	[CheckAccess]
	public class PersonController : Controller
	{
		readonly HttpClient _client;
		readonly Uri baseAddress = new("http://localhost:24254/api");

		public PersonController()
		{
			_client = new HttpClient
			{
				BaseAddress = baseAddress
			};
		}

		[HttpGet]
		public List<PersonModel> GetPerson()
		{
			List<PersonModel> personList = [];
			HttpResponseMessage response = _client.GetAsync($"{_client.BaseAddress}/Person/Get").Result;
			if (response.IsSuccessStatusCode)
			{
				string data = response.Content.ReadAsStringAsync().Result;
				dynamic jsonobject = JsonConvert.DeserializeObject(data);
				var dataOfObject = jsonobject.data;
				var extractedDataJson = JsonConvert.SerializeObject(dataOfObject, Formatting.Indented);
				personList = JsonConvert.DeserializeObject<List<PersonModel>>(extractedDataJson);
			}
			return personList;
		}

		[HttpGet]
		public IActionResult GetAllPerson()
		{
			//List<PersonModel> personList = GetPerson();
			return View("GetAllPerson");
		}
		
		public IActionResult ReturnJson()
		{
			var persons = GetPerson();
			return Json(persons);
		}

        [HttpGet]
        public IActionResult GetAllPersonMulti()
        {
            List<PersonModel> personList1 = GetPerson();
            return View("GetAllPersonMulti", personList1);
        }

        [HttpDelete]
		public IActionResult Delete(int PersonID)
		{
			HttpResponseMessage response = _client.DeleteAsync($"{_client.BaseAddress}/Person/DeleteByPersonID/{PersonID}").Result;
			if (response.IsSuccessStatusCode)
			{
				TempData["Message"] = "Person Deleted Successfully";
			}
			return View("GetAllPerson");
		}

		#region EditPerson
		[HttpGet]
		public IActionResult EditPerson(int PersonID)
		{
			PersonModel person = new();
			HttpResponseMessage response = _client.GetAsync($"{_client.BaseAddress}/Person/Get/{PersonID}").Result;
			if (response.IsSuccessStatusCode)
			{
				string data = response.Content.ReadAsStringAsync().Result;
				dynamic jsonobject = JsonConvert.DeserializeObject(data);
				var dataOfObject = jsonobject.data;
				var extractedDataJson = JsonConvert.SerializeObject(dataOfObject, Formatting.Indented);
				person = JsonConvert.DeserializeObject<PersonModel>(extractedDataJson);
			}
			return View("PersonAddEdit", person);
		}
		#endregion

		#region Save
		[HttpPost]
		public IActionResult Save(PersonModel model)
		{
			try
			{
				MultipartFormDataContent formData = new()
				{
					{ new StringContent(model.Name), "Name" },
					{ new StringContent(model.Contact), "Contact" },
					{ new StringContent(model.Email), "Email" }
				};

				if (model.PersonID == 0)
				{
					HttpResponseMessage response = _client.PostAsync($"{_client.BaseAddress}/Person/InsertPerson", formData).Result;
					if (response.IsSuccessStatusCode)
					{
						TempData["Message"] = "Person Saved Successfully";
					}
				}
				else
				{
					HttpResponseMessage response = _client.PutAsync($"{_client.BaseAddress}/Person/UpdatePerson/{model.PersonID}", formData).Result;
					if (response.IsSuccessStatusCode)
					{
						TempData["Message"] = "Person Updated Successfully";
					}
				}
				return RedirectToAction("GetAllPerson");
			}
			catch (Exception ex)
			{
				TempData["Error"] = "Error occured : " + ex.Message;
			}
			return RedirectToAction("GetAllPerson");
		}
		#endregion

		#region ExportToExcel
		[Route("Person/ExportExcel")]
		public ActionResult ExportExcel()
		{
			List<PersonModel> persons = GetPerson();
			using var workbook = new XLWorkbook();
			var worksheet = workbook.Worksheets.Add("Persons");
			worksheet.Cell(1, 1).Value = "Name";
			worksheet.Cell(1, 2).Value = "Contact";
			worksheet.Cell(1, 3).Value = "Email";
			int currentRow = 2;
			foreach (var person in persons)
			{
				worksheet.Cell(currentRow, 1).Value = person.Name;
				worksheet.Cell(currentRow, 2).Value = person.Contact;
				worksheet.Cell(currentRow, 3).Value = person.Email;
				currentRow++;
			}
			worksheet.RangeUsed().Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
			worksheet.RangeUsed().Style.Border.SetInsideBorder(XLBorderStyleValues.Thin);
			worksheet.RangeUsed().Style.Font.SetFontName("Calibri");
			worksheet.RangeUsed().Style.Font.SetFontSize(12);
			worksheet.RangeUsed().Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
			worksheet.RangeUsed().Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
			worksheet.Columns().AdjustToContents();
			using var stream = new MemoryStream();
			workbook.SaveAs(stream);
			var content = stream.ToArray();
			return File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", $"{DateTime.Now}_Persons.xlsx");
		}
		#endregion

		#region DeleteSelected
		[HttpPost]
		[Route("Person/DeleteSelected")]
		public IActionResult DeleteSelected(List<int> selectedPersons)
		{
			if (selectedPersons != null && selectedPersons.Count != 0)
			{
				foreach (var personID in selectedPersons)
				{
					_ = _client.DeleteAsync($"{_client.BaseAddress}/Person/DeleteByPersonID/{personID}").Result;
				}
				TempData["SuccessMessage"] = "Selected persons deleted successfully.";
			}
			else
			{
				TempData["ErrorMessage"] = "Please select at least one person for deletion.";
			}
			return RedirectToAction("GetAllPersonMulti");
		}
		#endregion
	}
}

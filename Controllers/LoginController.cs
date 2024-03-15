using Microsoft.AspNetCore.Mvc;

namespace ConsumeWebAPI.Controllers
{
    public class LoginController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }
    }
}

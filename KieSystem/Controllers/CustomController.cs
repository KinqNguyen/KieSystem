using KieSystem.DTOs;
using Microsoft.AspNetCore.Mvc;
using System.Collections.Generic;

namespace KieSystem.Controllers
{

    [Obsolete("This controller is deprecated and should not be used.")]
    public class ObsoleteControllerBase : ControllerBase
    {
    }

    [ApiController]
    [Route("api/[controller]")]
    public class CustomController : ObsoleteControllerBase
    {


        [HttpGet(Name = "get")]
        public IEnumerable<BlogClass> Get() { 
        
            IEnumerable<BlogClass> list = autogenerate(100);
            return list;
        }

        private List<BlogClass> autogenerate(int number) {
            List<BlogClass> list = new List<BlogClass>();
            for (int i = 0; i < number; i++) {
                list.Add(new BlogClass
                {
                    Id = i,
                    Title = "Title " + i.ToString(),
                    Body = "BOdy " + i.ToString(),
                    UserId = 1,
                });
            }
            return list;
        }
    }


}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace EdenredSolution_WebApi.Controllers
{
    [Route("api/[controller]/[action]")]
    [ApiController]
    //[Authorize]
    public class CovidController : ControllerBase
    {
        [HttpGet]
        public async Task<IActionResult> GetCountryList() {
            return Ok(await CallExternalAPI("https://covid-193.p.rapidapi.com/countries"));
        }

        [HttpGet]
        public async Task<IActionResult> GetstatisticsList()
        {
            return Ok(await CallExternalAPI("https://covid-193.p.rapidapi.com/statistics"));
        }

        [HttpGet]
        public async Task<IActionResult> GetstatisticsByCountryName([FromQuery] string countryName)
        {
            return Ok(await CallExternalAPI("https://covid-193.p.rapidapi.com/statistics?country=" + countryName));
        }


        [HttpGet]
        public async Task<IActionResult> GetHistoryByCountryName([FromQuery] string countryName)
        {
            return Ok(await CallExternalAPI("https://covid-193.p.rapidapi.com/history?country=" + countryName));
        }

        public async Task<String> CallExternalAPI(string url)
        {
            using (var client = new HttpClient())
            {
                var request = new HttpRequestMessage
                {
                    Method = HttpMethod.Get,
                    RequestUri = new Uri(url),
                    Headers =
                    {
                        { "x-rapidapi-key", "122e1b4971msh53633742bf33c9ap18a51ejsn7de0ad387d2b" },
                        { "x-rapidapi-host", "covid-193.p.rapidapi.com" },
                    },
                };
                var response = await client.SendAsync(request);
                response.EnsureSuccessStatusCode();
                return await response.Content.ReadAsStringAsync();
            }
        }

    }
}

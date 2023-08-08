using AspNetCore.Reporting;
using AspNetCore.Reporting.ReportExecutionService;
using Microsoft.AspNetCore.Mvc;
using System.Data;

// For more information on enabling Web API for empty projects, visit https://go.microsoft.com/fwlink/?LinkID=397860

namespace ReportAPI.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ReportController : ControllerBase
    {

        public ReportController()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        }   
        
        // GET: api/<ReportController>
        [HttpGet]
        public IActionResult Get()
        {
            int extension = 1;
            var _reportPath = Path.Combine(Environment.CurrentDirectory, @"Reports\testeReport.rdlc");

            LocalReport localReport = new(_reportPath);


            //Dados
            System.Data.DataTable dt = new();
            dt.Clear();
            dt.Columns.Add("Id", typeof(int));
            dt.Columns.Add("Name", typeof(string));
            dt.Rows.Add(1, "Paulo");
            dt.Rows.Add(2, "Jose");
            localReport.AddDataSource("DataSet1", dt);


            //Parametros do relatório
            var reportParams = new Dictionary<string, string>();
            //reportParams.Add("Key1", "value1");
            //reportParams.Add("Key2", "value2");
            if (reportParams != null && reportParams.Count > 0)// if you use parameter in report
            {
                List<ReportParameter> reportparameter = new();
                foreach (var record in reportParams)
                {
                    reportparameter.Add(new ReportParameter());
                }

            }

            //Geração do arquivo
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            var result = localReport.Execute(RenderType.Pdf, extension, parameters: reportParams);
            byte[] file = result.MainStream;

            Stream stream = new MemoryStream(file);
            return File(stream, "application/pdf", "testeReport.pdf");

        }


        // GET: api/<ReportController>
        [HttpGet("/api/Booking")]
        public IActionResult GetBookingReport()
        {
            int extension = 1;
            var _reportPath = Path.Combine(Environment.CurrentDirectory, @"Reports\BookingReport.rdlc");

            LocalReport localReport = new(_reportPath);


            //Dados
            DataTable dt = new();
            dt.Clear();
            dt.Columns.Add("Id", typeof(int));
            dt.Columns.Add("Booking Id", typeof(string));
            dt.Columns.Add("Travel Date", typeof(DateTime));
            dt.Columns.Add("Airline", typeof(string));
            dt.Columns.Add("Route", typeof(string));
            dt.Columns.Add("Ticket Number", typeof(string));
            dt.Columns.Add("Passenger Name", typeof(string));
            dt.Columns.Add("Mobile Number", typeof(string));
            dt.Columns.Add("Fare", typeof(string));

            Random random = new();

            var randomData = Enumerable.Range(1, 50).Select(id => new
            {
                Id = id,
                BookingId = "Booking" + id.ToString(),
                TravelDate = DateTime.Now.AddDays(random.Next(1, 365)),
                Airline = "Airline" + random.Next(1, 11).ToString(),
                Route = "Route" + random.Next(1, 6).ToString(),
                TicketNumber = "Ticket" + id.ToString(),
                PassengerName = "Passenger" + id.ToString(),
                MobileNumber = "Mobile" + id.ToString(),
                Fare = random.Next(100, 1000).ToString()
            });

            foreach (var data in randomData)
            {
                dt.Rows.Add(data.Id, data.BookingId, data.TravelDate, data.Airline, data.Route, data.TicketNumber, data.PassengerName, data.MobileNumber, data.Fare);
            }


            localReport.AddDataSource("OTABooking", dt);


            //Parametros do relatório
            var reportParams = new Dictionary<string, string>();
            if (reportParams != null && reportParams.Count > 0)
            {
                List<ReportParameter> reportparameter = new();
                foreach (var record in reportParams)
                {
                    reportparameter.Add(new ReportParameter());
                }

            }

            //Geração do arquivo
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            var result = localReport.Execute(RenderType.Excel, extension, parameters: reportParams);
            byte[] file = result.MainStream;
            return File(result.MainStream, "application/xls", "BookingReport.xls");

        }


    }
}

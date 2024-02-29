using System.Data;
using System.Net;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Logging;
using Azure.Storage.Blobs;
using ExcelDataReader;
using Microsoft.Data.SqlClient;
using System.Text.Encodings;
using System.Diagnostics;
using System.Text;


namespace FunctionApp1
{
    public class Function1
    {
        private readonly ILogger _logger;

        public Function1(ILoggerFactory loggerFactory)
        {
            _logger = loggerFactory.CreateLogger<Function1>();
        }

        [Function("Function1")]
        public HttpResponseData Run([HttpTrigger(AuthorizationLevel.Function, "get", "post")] HttpRequestData req)
        {
            _logger.LogInformation("C# HTTP trigger function processed a request.");

            var response = req.CreateResponse(HttpStatusCode.OK);
            response.Headers.Add("Content-Type", "text/plain; charset=utf-8");

            response.WriteString("Welcome to Azure Functions!");

            // Create a new Stopwatch instance
            Stopwatch stopwatch = new Stopwatch();

            // Start the Stopwatch
            stopwatch.Start();
            response.WriteString("Started");

            //read an excel file from a blob store and save it to sql database

            // Create a BlobServiceClient object which will be used to create a container client
            BlobServiceClient blobServiceClient = new BlobServiceClient("Your blob endpoint");

            // Create the container and return a container client object
            BlobContainerClient containerClient = blobServiceClient.GetBlobContainerClient("files");

            // Get a reference to a blob
            BlobClient blobClient = containerClient.GetBlobClient("sample4.xlsx");

            // Add this at the start of your function
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            //read a file from blob storage as a filestream to pass to exceldatareader
            var fileStream = new MemoryStream();
            blobClient.DownloadTo(fileStream);
            fileStream.Position = 0; // Reset stream position

            DataTable dt = new DataTable();

            try
            {

                using (var reader = ExcelReaderFactory.CreateReader(fileStream, new ExcelReaderConfiguration()
                {
                    // Gets or sets the encoding to use when the input XLS lacks a CodePage
                    // record, or when the input CSV lacks a BOM and does not parse as UTF8. 
                    // Default: cp1252 (XLS BIFF2-5 and CSV only)
                    FallbackEncoding = Encoding.GetEncoding(1252),

                    // Gets or sets the password used to open password protected workbooks.
                    Password = "password",
                }))
                {
                    var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                        {
                            UseHeaderRow = true
                        }
                    });

                    dt = result.Tables[0];
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error reading Excel file from Blob Storage");
            }

            try
            {
                using (SqlConnection dbConnection = new SqlConnection("Your connectionstring"))
                {
                    dbConnection.Open();

                    using (SqlBulkCopy s = new SqlBulkCopy(dbConnection))
                    {
                        s.DestinationTableName = "test";
                        s.WriteToServer(dt);
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error saving data to SQL Database");
            }

            // Stop the Stopwatch
            stopwatch.Stop();

            // Get the elapsed time as a TimeSpan value
            TimeSpan ts = stopwatch.Elapsed;

            // Format and display the TimeSpan value
            string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}",
                ts.Hours, ts.Minutes, ts.Seconds,
                ts.Milliseconds / 10);
            Console.WriteLine("RunTime " + elapsedTime);
            response.WriteString("Running time" + elapsedTime);

            return response;
        }
    }
}

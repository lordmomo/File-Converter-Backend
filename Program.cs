
using Aspose.Cells.Charts;
using FileConversion.Controllers;
using FileConversion.Service.ServiceImplementation;
using FileConversion.Service.ServiceInterface;

namespace FileConversion
{
    public class Program
    {
        public static void Main(string[] args)
        {

            //ExampleStart.HelloWorld();
            var builder = WebApplication.CreateBuilder(args);

            // Add services to the container.

            builder.Services.AddControllers();
            // Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
            builder.Services.AddEndpointsApiExplorer();
            builder.Services.AddSwaggerGen();

            
            // For Hosting in AWS Lambda
            builder.Services.AddAWSLambdaHosting(LambdaEventSource.HttpApi);

            builder.Services.AddScoped<ReportGeneratorInterface, ReportGeneratorImplementation>();

            builder.Services.AddSingleton<StudentReportInterfae, StudentReportImplementation>();
            builder.Services.AddScoped<DocumentConverterInterface, DocumentConverterImplementation>();

            var app = builder.Build();

            app.UseCors(options =>
            {
                options.AllowAnyOrigin().AllowAnyMethod().AllowAnyHeader();
            });

            // Configure the HTTP request pipeline.
            if (app.Environment.IsDevelopment())
            {
                app.UseSwagger();
                app.UseSwaggerUI();
            }

            app.UseAuthorization();



            app.MapControllers();

            app.Run();
        }
    }
}

using Microsoft.EntityFrameworkCore;
using API_INDER_INFORMES.Data;
using API_INDER_INFORMES.Services;
using OfficeOpenXml;


ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

var builder = WebApplication.CreateBuilder(args);


builder.Services.AddControllers();


builder.Services.AddSingleton<EmailService>();


// Configurar la conexión a la base de datos INDER_DATABASE
builder.Services.AddDbContext<ApplicationDbContext>(options =>
    options.UseSqlServer(builder.Configuration.GetConnectionString("DefaultConnection")));

// Configurar la conexión a la base de datos DASHBOARD_PRODUCCION
builder.Services.AddDbContext<DashboardDbContext>(options =>
    options.UseSqlServer(builder.Configuration.GetConnectionString("DashboardConnection")));

// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

var app = builder.Build();


if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseHttpsRedirection();

app.UseAuthorization();

app.MapControllers();

app.Run();

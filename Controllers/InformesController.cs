using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using API_INDER_INFORMES.Data;
using API_INDER_INFORMES.Models;
using API_INDER_INFORMES.Services;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;

namespace API_INDER_INFORMES.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class InformesController : ControllerBase
    {
        private readonly ApplicationDbContext _context;
        private readonly IWebHostEnvironment _environment;
        private readonly EmailService _emailService;

        public InformesController(
            ApplicationDbContext context, 
            IWebHostEnvironment environment,
            EmailService emailService)
        {
            _context = context;
            _environment = environment;
            _emailService = emailService;
        }

        [HttpGet("generar-informe")]
        public async Task<IActionResult> GenerarInforme([FromQuery] DateTime fecha)
        {
            var registros = await _context.FormularioInder
                .Join(_context.DisponibilidadHorariosInder,
                    f => f.Id,
                    d => d.IdDis,
                    (f, d) => new
                    {
                        f.Nombres,
                        f.Apellidos,
                        f.Correo,
                        f.Direccion,
                        f.FechaNacimiento,
                        f.TipoDocumento,
                        f.NumeroDocumento,
                        f.Genero,
                        f.Celular,
                        f.Edad,
                        Lugar = d.Lugar,
                        FechaRegistro = f.FechaRegistro
                    })
                .Where(r => r.FechaRegistro.Date == fecha.Date)
                .OrderBy(r => r.Apellidos)
                .ThenBy(r => r.Nombres)
                .ToListAsync();

            var fileName = $"Informe_INDER_{fecha:yyyyMMdd}.xlsx";
            var wwwrootPath = Path.Combine(_environment.WebRootPath, "Informes");
            
            if (!Directory.Exists(wwwrootPath))
            {
                Directory.CreateDirectory(wwwrootPath);
            }

            var filePath = Path.Combine(wwwrootPath, fileName);

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Registros");
                
            
                worksheet.Cells[1, 1, 1, 12].Merge = true;
                worksheet.Cells[1, 1].Value = $"INFORME DIARIO INDER - {fecha:dd/MM/yyyy}";
                var titleStyle = worksheet.Cells[1, 1, 1, 12].Style;
                titleStyle.Font.Size = 16;
                titleStyle.Font.Bold = true;
                titleStyle.Fill.PatternType = ExcelFillStyle.Solid;
                titleStyle.Fill.BackgroundColor.SetColor(Color.FromArgb(0, 112, 192)); 
                titleStyle.Font.Color.SetColor(Color.White);
                titleStyle.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                titleStyle.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Row(1).Height = 30; 

             
                var headerStyle = worksheet.Cells[2, 1, 2, 12].Style;
                headerStyle.Fill.PatternType = ExcelFillStyle.Solid;
                headerStyle.Fill.BackgroundColor.SetColor(Color.LightBlue);
                headerStyle.Font.Bold = true;
                headerStyle.Font.Size = 12;
                headerStyle.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                headerStyle.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Row(2).Height = 25; 

              
                worksheet.Cells[2, 1].Value = "NOMBRES";
                worksheet.Cells[2, 2].Value = "APELLIDOS";
                worksheet.Cells[2, 3].Value = "CORREO";
                worksheet.Cells[2, 4].Value = "DIRECCIÓN";
                worksheet.Cells[2, 5].Value = "FECHA NACIMIENTO";
                worksheet.Cells[2, 6].Value = "TIPO DOCUMENTO";
                worksheet.Cells[2, 7].Value = "NÚMERO DOCUMENTO";
                worksheet.Cells[2, 8].Value = "GÉNERO";
                worksheet.Cells[2, 9].Value = "CELULAR";
                worksheet.Cells[2, 10].Value = "EDAD";
                worksheet.Cells[2, 11].Value = "LUGAR";
                worksheet.Cells[2, 12].Value = "FECHA REGISTRO";

              
                int row = 3;
                foreach (var registro in registros)
                {
                    worksheet.Cells[row, 1].Value = registro.Nombres;
                    worksheet.Cells[row, 2].Value = registro.Apellidos;
                    worksheet.Cells[row, 3].Value = registro.Correo;
                    worksheet.Cells[row, 4].Value = registro.Direccion;
                    worksheet.Cells[row, 5].Value = registro.FechaNacimiento;
                    worksheet.Cells[row, 6].Value = registro.TipoDocumento;
                    worksheet.Cells[row, 7].Value = registro.NumeroDocumento;
                    worksheet.Cells[row, 8].Value = registro.Genero;
                    worksheet.Cells[row, 9].Value = registro.Celular;
                    worksheet.Cells[row, 10].Value = registro.Edad;
                    worksheet.Cells[row, 11].Value = registro.Lugar;
               
                    worksheet.Cells[row, 12].Value = registro.FechaRegistro.ToString("dd/MM/yyyy HH:mm:ss");
                    row++;
                }
                
        
                var dataRange = worksheet.Cells[2, 1, row - 1, 12];
                dataRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                dataRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                dataRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                dataRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

   
                worksheet.Cells.AutoFitColumns();

       
                worksheet.Column(5).Style.Numberformat.Format = "dd/mm/yyyy";
                
              
                worksheet.Cells[row + 2, 1, row + 2, 3].Merge = true;
                worksheet.Cells[row + 2, 1].Value = "INFORMACIÓN DEL INFORME";
                worksheet.Cells[row + 2, 1].Style.Font.Bold = true;
                worksheet.Cells[row + 2, 1].Style.Font.Size = 12;
                worksheet.Cells[row + 2, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[row + 2, 1].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                worksheet.Cells[row + 2, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                
       
                worksheet.Cells[row + 3, 1].Value = "INFORME GENERADO EL:";
                worksheet.Cells[row + 3, 1].Style.Font.Bold = true;
                worksheet.Cells[row + 3, 2].Value = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
                

                worksheet.Cells[row + 4, 1].Value = "TOTAL DE REGISTROS:";
                worksheet.Cells[row + 4, 1].Style.Font.Bold = true;
                worksheet.Cells[row + 4, 2].Value = registros.Count;
                
              
                worksheet.Cells[row + 6, 1, row + 6, 12].Merge = true;
                worksheet.Cells[row + 6, 1].Value = "© INDER 2025 - TODOS LOS DERECHOS RESERVADOS";
                worksheet.Cells[row + 6, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[row + 6, 1].Style.Font.Italic = true;

            
                await package.SaveAsAsync(new FileInfo(filePath));
            }

           
            var emailBody = $@"
                <h2>Informe Diario INDER</h2>
                <p>Se adjunta el informe del día {fecha:dd/MM/yyyy}</p>
                <p>Total de registros: {registros.Count}</p>
                <p>Este es un correo automático, por favor no responder.</p>";

            await _emailService.SendEmailAsync(
                "wilyd2@hotmail.com",
                $"Informe INDER - {fecha:dd/MM/yyyy}",
                emailBody,
                filePath
            );

          
            var bytes = await System.IO.File.ReadAllBytesAsync(filePath);
            return File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }
    }
} 
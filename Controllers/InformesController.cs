using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using API_INDER_INFORMES.Data;
using API_INDER_INFORMES.Models;
using API_INDER_INFORMES.Services;
using System.Text;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using System.IO;
using System.Globalization;
using System.Linq;
using System.Collections.Generic;
using System.Threading.Tasks;
using System;

namespace API_INDER_INFORMES.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class InformesController : ControllerBase
    {
        private readonly ApplicationDbContext _context;
        private readonly DashboardDbContext _dashboardContext;
        private readonly IWebHostEnvironment _environment;
        private readonly EmailService _emailService;

        public InformesController(
            ApplicationDbContext context,
            DashboardDbContext dashboardContext,
            IWebHostEnvironment environment,
            EmailService emailService)
        {
            _context = context;
            _dashboardContext = dashboardContext;
            _environment = environment;
            _emailService = emailService;
        }

        [HttpGet("generar-informe")]
        public async Task<IActionResult> GenerarInforme([FromQuery] DateTime fecha)
        {
            // Obtener los registros de INDER_DATABASE
            var registrosInder = await _context.FormularioInder
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
                
            // Obtener las transacciones de DASHBOARD_PRODUCCION directamente sin proyección
            var transacciones = await _dashboardContext.Transactions
                .AsNoTracking() // Para mejorar el rendimiento
                .Where(t => t.DATE_CREATED.Date == fecha.Date)
                .ToListAsync();
                
            // Imprimir información de diagnóstico para verificar los datos
            System.Diagnostics.Debug.WriteLine($"Total de transacciones recuperadas: {transacciones.Count}");
            foreach (var t in transacciones.Take(5)) // Imprimir las primeras 5 para diagnóstico
            {
                System.Diagnostics.Debug.WriteLine($"ID: {t.ID}, Doc: {t.DOCUMENT}, Ref: {t.REFERENCE}, Monto: {t.TOTAL_AMOUNT}");
            }
                
            // Obtener los detalles de transacciones
            var detallesTransacciones = await _dashboardContext.TransactionDetails
                .ToListAsync();
                
            // Normalizar los números de documento para mejorar la coincidencia
            foreach (var transaccion in transacciones)
            {
                if (transaccion.DOCUMENT != null)
                {
                    // Eliminar espacios, puntos y guiones
                    transaccion.DOCUMENT = transaccion.DOCUMENT.Replace(" ", "").Replace(".", "").Replace("-", "");
                }
            }
            
            // Obtener los registros de INDER con sus transacciones
            var registros = registrosInder
                .Select(r => new
                {
                    r.Nombres,
                    r.Apellidos,
                    r.Correo,
                    r.Direccion,
                    r.FechaNacimiento,
                    r.TipoDocumento,
                    r.NumeroDocumento,
                    r.Genero,
                    r.Celular,
                    r.Lugar,
                    r.FechaRegistro,
                    // Buscar las transacciones que coinciden con este número de documento
                    Transacciones = transacciones
                        .Where(t => t.DOCUMENT != null && 
                               r.NumeroDocumento != null && 
                               t.DOCUMENT.Trim().Equals(r.NumeroDocumento.Trim(), StringComparison.OrdinalIgnoreCase))
                        .ToList()
                })
                .OrderBy(r => r.Apellidos)
                .ThenBy(r => r.Nombres)
                .ToList();

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
                // Agregar título principal
                worksheet.Cells[1, 1, 1, 14].Merge = true;
                worksheet.Cells[1, 1].Value = $"INFORME DIARIO INDER - {fecha:dd/MM/yyyy}";
                var titleStyle = worksheet.Cells[1, 1, 1, 14].Style;
                titleStyle.Font.Size = 16;
                titleStyle.Font.Bold = true;
                titleStyle.Fill.PatternType = ExcelFillStyle.Solid;
                titleStyle.Fill.BackgroundColor.SetColor(Color.FromArgb(0, 112, 192)); // Azul INDER
                titleStyle.Font.Color.SetColor(Color.White);
                titleStyle.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                titleStyle.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Row(1).Height = 30; // Altura de la fila del título
                
                // Estilo para el encabezado de columnas (ahora en fila 2)
                var headerStyle = worksheet.Cells[2, 1, 2, 14].Style;
                headerStyle.Fill.PatternType = ExcelFillStyle.Solid;
                headerStyle.Fill.BackgroundColor.SetColor(Color.LightBlue);
                headerStyle.Font.Bold = true;
                headerStyle.Font.Size = 12;
                headerStyle.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                headerStyle.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Row(2).Height = 25; // Altura para la fila de encabezados

                // Agregar encabezados en fila 2
                worksheet.Cells[2, 1].Value = "NOMBRES";
                worksheet.Cells[2, 2].Value = "APELLIDOS";
                worksheet.Cells[2, 3].Value = "CORREO";
                worksheet.Cells[2, 4].Value = "DIRECCIÓN";
                worksheet.Cells[2, 5].Value = "FECHA NACIMIENTO";
                worksheet.Cells[2, 6].Value = "TIPO DOCUMENTO";
                worksheet.Cells[2, 7].Value = "NÚMERO DOCUMENTO";
                worksheet.Cells[2, 8].Value = "GÉNERO";
                worksheet.Cells[2, 9].Value = "CELULAR";
                worksheet.Cells[2, 10].Value = "LUGAR";
                worksheet.Cells[2, 11].Value = "FECHA REGISTRO";
                worksheet.Cells[2, 12].Value = "REFERENCIA";
                worksheet.Cells[2, 13].Value = "TOTAL MONTO";
                worksheet.Cells[2, 14].Value = "FECHA TRANSACCIÓN";

                // Agregar datos empezando en la fila 3
                int row = 3;
                foreach (var registro in registros)
                {
                    bool primeraFila = true;
                    
                    // Solo procesar registros que tienen transacciones asociadas
                    if (registro.Transacciones != null && registro.Transacciones.Count > 0 && 
                        !string.IsNullOrEmpty(registro.NumeroDocumento))
                    {
                        foreach (var transaccion in registro.Transacciones)
                        {
                            // Solo en la primera fila de cada persona, mostrar sus datos personales
                            if (primeraFila)
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
                                worksheet.Cells[row, 10].Value = registro.Lugar;
                                worksheet.Cells[row, 11].Value = registro.FechaRegistro.ToString("dd/MM/yyyy HH:mm:ss");
                                primeraFila = false;
                            }
                            else
                            {
                                // Para filas adicionales, repetir solo el número de documento para referencia
                                worksheet.Cells[row, 7].Value = registro.NumeroDocumento;
                            }
                            
                            try
                            {
                                // Imprimir datos de diagnóstico para cada transacción
                                System.Diagnostics.Debug.WriteLine($"Procesando transacción: ID={transaccion.ID}, Doc={transaccion.DOCUMENT}, Ref={transaccion.REFERENCE}, Monto={transaccion.TOTAL_AMOUNT}");
                                
                                // Manejar valores nulos con más detalle
                                if (transaccion.REFERENCE != null)
                                {
                                    worksheet.Cells[row, 12].Value = transaccion.REFERENCE;
                                }
                                else
                                {
                                    worksheet.Cells[row, 12].Value = "N/A";
                                }
                                
                                if (transaccion.TOTAL_AMOUNT.HasValue)
                                {
                                    worksheet.Cells[row, 13].Value = transaccion.TOTAL_AMOUNT.Value;
                                    // Aplicar formato de moneda a la columna de monto
                                    worksheet.Cells[row, 13].Style.Numberformat.Format = "$#,##0.00";
                                }
                                else
                                {
                                    worksheet.Cells[row, 13].Value = 0.0;
                                    worksheet.Cells[row, 13].Style.Numberformat.Format = "$#,##0.00";
                                }
                                
                                if (transaccion.DATE_CREATED != DateTime.MinValue)
                                {
                                    worksheet.Cells[row, 14].Value = transaccion.DATE_CREATED.ToString("dd/MM/yyyy HH:mm:ss");
                                }
                                else
                                {
                                    worksheet.Cells[row, 14].Value = "N/A";
                                }
                            }
                            catch (Exception ex)
                            {
                                // Si hay algún error al procesar la transacción, registrarlo y continuar
                                System.Diagnostics.Debug.WriteLine($"Error al procesar transacción: {ex.Message}");
                                worksheet.Cells[row, 12].Value = "Error: " + ex.Message;
                                worksheet.Cells[row, 13].Value = 0;
                                worksheet.Cells[row, 14].Value = "Error";
                            }
                            
                            row++;
                        }
                    }
                    // No mostramos registros sin transacciones
                }
                
                var dataRange = worksheet.Cells[2, 1, row - 1, 14];
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
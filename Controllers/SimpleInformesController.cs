using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using API_INDER_INFORMES.Data;
using API_INDER_INFORMES.Models;
using System.Collections.Generic;
using System.Threading.Tasks;
using System;
using System.Linq;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using Microsoft.AspNetCore.Hosting;
using System.Text;
using API_INDER_INFORMES.Services;

namespace API_INDER_INFORMES.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class SimpleInformesController : ControllerBase
    {
        private readonly ApplicationDbContext _inderContext;
        private readonly DashboardDbContext _dashboardContext;
        private readonly IWebHostEnvironment _environment;
        private readonly EmailService _emailService;

        public SimpleInformesController(
            ApplicationDbContext inderContext,
            DashboardDbContext dashboardContext,
            IWebHostEnvironment environment,
            EmailService emailService)
        {
            _inderContext = inderContext;
            _dashboardContext = dashboardContext;
            _environment = environment;
            _emailService = emailService;
        }

        [HttpGet("generar-informe-simple")]
        public async Task<IActionResult> GenerarInformeSimple([FromQuery] DateTime fecha)
        {
            try
            {
                // 1. Obtener transacciones de la fecha especificada con estado Aprobada
        
                var estadoAprobada = await _dashboardContext.StateTransactions
                    .FirstOrDefaultAsync(s => s.STATE == "Aprobada");
                
                int idEstadoAprobada = estadoAprobada?.ID ?? 0;
                
                var transacciones = await _dashboardContext.Transactions
                 .Where(t => t.DATE_CREATED.Date == fecha.Date && t.ID_STATE_TRANSACTION == idEstadoAprobada)
                 .Join(
                     _dashboardContext.PayPads,
                     t => t.ID_PAYPAD,
                     p => p.ID,
                     (t, p) => new
                     {
                         ID = t.ID,
                         NumeroDocumento = t.DOCUMENT,
                         Referencia = t.REFERENCE,
                         Producto = t.PRODUCT,
                         Monto = t.TOTAL_AMOUNT,
                         IdPaypad = t.ID_PAYPAD,
                         IdEstado = t.ID_STATE_TRANSACTION,
                         Descripcion = p.DESCRIPTION,
                         FechaTransaccion = t.DATE_CREATED
                     })
                    .ToListAsync();

                // 2. Obtener todos los usuarios de INDER para buscar coincidencias
                var usuariosInder = await _inderContext.FormularioInder
                    .Select(u => new
                    {
                        u.Id,
                        u.Nombres,
                        u.Apellidos,
                        u.Correo,
                        u.Direccion,
                        u.FechaNacimiento,
                        u.TipoDocumento,
                        u.NumeroDocumento,
                        u.Genero,
                        u.Celular,
                        u.FechaRegistro
                    })
                    .ToListAsync();

                // 3. Crear un diccionario para buscar usuarios por número de documento
                var usuariosPorDocumento = new Dictionary<string, dynamic>(StringComparer.OrdinalIgnoreCase);
                foreach (var usuario in usuariosInder)
                {
                    if (!string.IsNullOrEmpty(usuario.NumeroDocumento))
                    {
                      
                        var docNormalizado = usuario.NumeroDocumento.Replace(" ", "").Replace(".", "").Replace("-", "");
                        usuariosPorDocumento[docNormalizado] = usuario;
                    }
                }

                // 4. Crear el resultado combinando transacciones con datos de usuario
                var resultado = new List<object>();
                foreach (var transaccion in transacciones)
                {
                    
                object infoTransaccion;
                
          
                var datosUsuarioDefault = new
                {
                    Nombres = "Sin registro",
                    Apellidos = "Sin registro",
                    Correo = "Sin registro",
                    Direccion = "Sin registro",
                    FechaNacimiento = "Sin registro",
                    TipoDocumento = "Sin registro",
                    Genero = "Sin registro",
                    Celular = "Sin registro",
                    UsuarioEncontrado = false
                };
                
                infoTransaccion = new
                {
                    transaccion.ID,
                    transaccion.NumeroDocumento,
                    transaccion.Referencia,
                    transaccion.Producto,
                    transaccion.Monto,
                    transaccion.IdPaypad,
                    transaccion.IdEstado,
                    transaccion.Descripcion,
                    transaccion.FechaTransaccion,
                    DatosUsuario = datosUsuarioDefault
                };

              
                    if (!string.IsNullOrEmpty(transaccion.NumeroDocumento))
                    {
                        var docNormalizado = transaccion.NumeroDocumento.Replace(" ", "").Replace(".", "").Replace("-", "");
                        if (usuariosPorDocumento.TryGetValue(docNormalizado, out var usuario))
                        {
                    
                            var datosUsuarioEncontrado = new
                            {
                                Nombres = usuario.Nombres ?? "Sin registro",
                                Apellidos = usuario.Apellidos ?? "Sin registro",
                                Correo = usuario.Correo ?? "Sin registro",
                                Direccion = usuario.Direccion ?? "Sin registro",
                                FechaNacimiento = usuario.FechaNacimiento,
                                TipoDocumento = usuario.TipoDocumento ?? "Sin registro",
                                Genero = usuario.Genero ?? "Sin registro",
                                Celular = usuario.Celular ?? "Sin registro",
                                UsuarioEncontrado = true
                            };
                            
                          
                          infoTransaccion = new
                            {
                                transaccion.ID,
                                transaccion.NumeroDocumento,
                                transaccion.Referencia,
                                transaccion.Producto,
                                transaccion.Monto,
                                transaccion.IdPaypad,
                                transaccion.IdEstado,
                                transaccion.Descripcion,
                                transaccion.FechaTransaccion,
                                DatosUsuario = datosUsuarioEncontrado
                            };
                        }
                    }

            
                    if (!string.IsNullOrEmpty(transaccion.NumeroDocumento) && 
                        ((dynamic)infoTransaccion).DatosUsuario.UsuarioEncontrado)
                    {
                        resultado.Add(infoTransaccion);
                    }
                }

                return Ok(new
                {
                    Fecha = fecha.ToString("yyyy-MM-dd"),
                    TotalTransacciones = resultado.Count,
                    Transacciones = resultado
                });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { error = ex.Message, stackTrace = ex.StackTrace });
            }
        }

        [HttpGet("descargar-excel")]
        public async Task<IActionResult> DescargarExcel([FromQuery] DateTime fecha)
        {
            try
            {
                // 1. Obtener transacciones de la fecha especificada con estado Aprobada
            
                var estadoAprobada = await _dashboardContext.StateTransactions
                    .FirstOrDefaultAsync(s => s.STATE == "Aprobada");
                
                int idEstadoAprobada = estadoAprobada?.ID ?? 0;
                
                var transacciones = await _dashboardContext.Transactions
                   .Where(t => t.DATE_CREATED.Date == fecha.Date && t.ID_STATE_TRANSACTION == idEstadoAprobada)
                   .Join(
                       _dashboardContext.PayPads,
                       t => t.ID_PAYPAD,
                       p => p.ID,
                       (t, p) => new
                       {
                           ID = t.ID,
                           NumeroDocumento = t.DOCUMENT,
                           Referencia = t.REFERENCE,
                           Producto = t.PRODUCT,
                           Monto = t.TOTAL_AMOUNT,
                           IdPaypad = t.ID_PAYPAD,
                           IdEstado = t.ID_STATE_TRANSACTION,
                           Descripcion = p.DESCRIPTION,
                           FechaTransaccion = t.DATE_CREATED
                       })
                    .ToListAsync();

                // 2. Obtener todos los usuarios de INDER para buscar coincidencias
                var usuariosInder = await _inderContext.FormularioInder
                    .Select(u => new
                    {
                        u.Id,
                        u.Nombres,
                        u.Apellidos,
                        u.Correo,
                        u.Direccion,
                        u.FechaNacimiento,
                        u.TipoDocumento,
                        u.NumeroDocumento,
                        u.Genero,
                        u.Celular,
                        u.FechaRegistro
                    })
                    .ToListAsync();

                // 3. Crear un diccionario para buscar usuarios por número de documento
                var usuariosPorDocumento = new Dictionary<string, dynamic>(StringComparer.OrdinalIgnoreCase);
                foreach (var usuario in usuariosInder)
                {
                    if (!string.IsNullOrEmpty(usuario.NumeroDocumento))
                    {
                       
                        var docNormalizado = usuario.NumeroDocumento.Replace(" ", "").Replace(".", "").Replace("-", "");
                        usuariosPorDocumento[docNormalizado] = usuario;
                    }
                }

                // 4. Crear el resultado combinando transacciones con datos de usuario
                var resultado = new List<object>();
                foreach (var transaccion in transacciones)
                {
                  
                    object infoTransaccion;
                
                  
                    var datosUsuarioDefault = new
                    {
                        Nombres = "Sin registro",
                        Apellidos = "Sin registro",
                        Correo = "Sin registro",
                        Direccion = "Sin registro",
                        FechaNacimiento = "Sin registro",
                        TipoDocumento = "Sin registro",
                        Genero = "Sin registro",
                        Celular = "Sin registro",
                        UsuarioEncontrado = false
                    };
                
                    infoTransaccion = new
                    {
                        transaccion.ID,
                        transaccion.NumeroDocumento,
                        transaccion.Referencia,
                        transaccion.Producto,
                        transaccion.Monto,
                        transaccion.IdPaypad,
                        transaccion.IdEstado,
                        transaccion.Descripcion,
                        transaccion.FechaTransaccion,
                        DatosUsuario = datosUsuarioDefault
                    };

                  
                    if (!string.IsNullOrEmpty(transaccion.NumeroDocumento))
                    {
                        var docNormalizado = transaccion.NumeroDocumento.Replace(" ", "").Replace(".", "").Replace("-", "");
                        if (usuariosPorDocumento.TryGetValue(docNormalizado, out var usuario))
                        {
                            
                            var datosUsuarioEncontrado = new
                            {
                                Nombres = usuario.Nombres ?? "Sin registro",
                                Apellidos = usuario.Apellidos ?? "Sin registro",
                                Correo = usuario.Correo ?? "Sin registro",
                                Direccion = usuario.Direccion ?? "Sin registro",
                                FechaNacimiento = usuario.FechaNacimiento,
                                TipoDocumento = usuario.TipoDocumento ?? "Sin registro",
                                Genero = usuario.Genero ?? "Sin registro",
                                Celular = usuario.Celular ?? "Sin registro",
                                UsuarioEncontrado = true
                            };
                            
                   
                            infoTransaccion = new
                            {
                                transaccion.ID,
                                transaccion.NumeroDocumento,
                                transaccion.Referencia,
                                transaccion.Producto,
                                transaccion.Monto,
                                transaccion.IdPaypad,
                                transaccion.IdEstado,
                                transaccion.Descripcion,
                                transaccion.FechaTransaccion,
                                DatosUsuario = datosUsuarioEncontrado,
                                DescripcionPaypad = transaccion.Descripcion
                            };
                        }
                    }

                   
                    if (!string.IsNullOrEmpty(transaccion.NumeroDocumento) && 
                        ((dynamic)infoTransaccion).DatosUsuario.UsuarioEncontrado)
                    {
                        resultado.Add(infoTransaccion);
                    }
                }

                // 5. Generar archivo Excel
                var fileName = $"Informe_Transacciones_{fecha:yyyy-MM-dd}.xlsx";
                var filePath = Path.Combine(_environment.ContentRootPath, "Temp", fileName);

                
                Directory.CreateDirectory(Path.Combine(_environment.ContentRootPath, "Temp"));
                
               
                if (System.IO.File.Exists(filePath))
                {
                    System.IO.File.Delete(filePath);
                }

              
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    var worksheet = package.Workbook.Worksheets.Add("Informe");

                    
                    worksheet.Cells[1, 1].Value = "Nombres";
                    worksheet.Cells[1, 2].Value = "Apellidos";
                    worksheet.Cells[1, 3].Value = "Correo";
                    worksheet.Cells[1, 4].Value = "Dirección";
                    worksheet.Cells[1, 5].Value = "Fecha Nacimiento";
                    worksheet.Cells[1, 6].Value = "Tipo Documento";
                    worksheet.Cells[1, 7].Value = "Número Documento";
                    worksheet.Cells[1, 8].Value = "Género";
                    worksheet.Cells[1, 9].Value = "Celular";
                    worksheet.Cells[1, 10].Value = "ID Transacción";
                    worksheet.Cells[1, 11].Value = "Referencia";
                    worksheet.Cells[1, 12].Value = "Producto";
                    worksheet.Cells[1, 13].Value = "Monto";
                    worksheet.Cells[1, 14].Value = "Fecha Transacción";
                    worksheet.Cells[1, 15].Value = "Descripción Paypad";

                 
                    var headerRange = worksheet.Cells[1, 1, 1, 15];
                    headerRange.Style.Font.Bold = true;
                    headerRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    headerRange.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                    headerRange.Style.Font.Color.SetColor(Color.Black);

                   
                    int row = 2;
                    foreach (dynamic item in resultado)
                    {
                        worksheet.Cells[row, 1].Value = item.DatosUsuario.Nombres;
                        worksheet.Cells[row, 2].Value = item.DatosUsuario.Apellidos;
                        worksheet.Cells[row, 3].Value = item.DatosUsuario.Correo;
                        worksheet.Cells[row, 4].Value = item.DatosUsuario.Direccion;
                        worksheet.Cells[row, 5].Value = item.DatosUsuario.FechaNacimiento;
                        worksheet.Cells[row, 6].Value = item.DatosUsuario.TipoDocumento;
                        worksheet.Cells[row, 7].Value = item.NumeroDocumento;
                        worksheet.Cells[row, 8].Value = item.DatosUsuario.Genero;
                        worksheet.Cells[row, 9].Value = item.DatosUsuario.Celular;
                        worksheet.Cells[row, 10].Value = item.ID;
                        worksheet.Cells[row, 11].Value = item.Referencia;
                        worksheet.Cells[row, 12].Value = item.Producto;
                        worksheet.Cells[row, 13].Value = item.Monto;
                        worksheet.Cells[row, 14].Value = item.FechaTransaccion.ToString("dd/MM/yyyy HH:mm:ss");
                        worksheet.Cells[row, 15].Value = item.DescripcionPaypad;
                        row++;
                    }

                 
                    worksheet.Cells.AutoFitColumns();

                   
                    package.Save();
                }

                
                await EnviarCorreoConExcel(filePath, fileName, fecha);
                
           
                byte[] fileBytes = System.IO.File.ReadAllBytes(filePath);
                return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { error = ex.Message, stackTrace = ex.StackTrace });
            }
        }

        private async Task EnviarCorreoConExcel(string filePath, string fileName, DateTime fecha)
        {
            try
            {
                string subject = $"Informe de Transacciones INDER - {fecha:yyyy-MM-dd}";
                string body = $"<html><body><p>Adjunto encontrará el informe de transacciones del {fecha:dd/MM/yyyy}.</p><p>Este es un correo automático, por favor no responda a este mensaje.</p></body></html>";
                
               
                await _emailService.SendEmailAsync("wruiz@e-city.co", subject, body, filePath);
                
               
                await _emailService.SendEmailAsync("contabilidad.inder@bello.gov.co", subject, body, filePath);
                await _emailService.SendEmailAsync("tesoreria.inder@bello.edu.co", subject, body, filePath);
            }
            catch (Exception ex)
            {
                
                Console.WriteLine($"Error al enviar correo: {ex.Message}");
            }
        }
    }
}

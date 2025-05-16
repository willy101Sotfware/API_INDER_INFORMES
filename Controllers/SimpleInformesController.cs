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

namespace API_INDER_INFORMES.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class SimpleInformesController : ControllerBase
    {
        private readonly ApplicationDbContext _inderContext;
        private readonly DashboardDbContext _dashboardContext;
        private readonly IWebHostEnvironment _environment;

        public SimpleInformesController(
            ApplicationDbContext inderContext,
            DashboardDbContext dashboardContext,
            IWebHostEnvironment environment)
        {
            _inderContext = inderContext;
            _dashboardContext = dashboardContext;
            _environment = environment;
        }

        [HttpGet("generar-informe-simple")]
        public async Task<IActionResult> GenerarInformeSimple([FromQuery] DateTime fecha)
        {
            try
            {
                // 1. Obtener transacciones de la fecha especificada
                var transacciones = await _dashboardContext.Transactions
                  .Where(t => t.DATE_CREATED.Date == fecha.Date && t.INCOME_AMOUNT != 0)
                    .Select(t => new
                    {
                        ID = t.ID,
                        NumeroDocumento = t.DOCUMENT,
                        Referencia = t.REFERENCE,
                        Producto = t.PRODUCT,
                        Monto = t.TOTAL_AMOUNT,
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
                        // Normalizar el documento para mejorar coincidencias
                        var docNormalizado = usuario.NumeroDocumento.Replace(" ", "").Replace(".", "").Replace("-", "");
                        usuariosPorDocumento[docNormalizado] = usuario;
                    }
                }

                // 4. Crear el resultado combinando transacciones con datos de usuario
                var resultado = new List<object>();
                foreach (var transaccion in transacciones)
                {
                    // Crear el objeto con la información de la transacción
                object infoTransaccion;
                
                // Valores por defecto para usuario no encontrado
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
                    transaccion.FechaTransaccion,
                    DatosUsuario = datosUsuarioDefault
                };

                    // Si hay un número de documento, buscar si hay un usuario asociado
                    if (!string.IsNullOrEmpty(transaccion.NumeroDocumento))
                    {
                        var docNormalizado = transaccion.NumeroDocumento.Replace(" ", "").Replace(".", "").Replace("-", "");
                        if (usuariosPorDocumento.TryGetValue(docNormalizado, out var usuario))
                        {
                            // Si se encuentra el usuario, crear un nuevo objeto con los datos del usuario
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
                            
                            // Actualizar infoTransaccion con los datos del usuario encontrado
                            infoTransaccion = new
                            {
                                transaccion.ID,
                                transaccion.NumeroDocumento,
                                transaccion.Referencia,
                                transaccion.Producto,
                                transaccion.Monto,
                                transaccion.FechaTransaccion,
                                DatosUsuario = datosUsuarioEncontrado
                            };
                        }
                    }

                    // Solo agregar al resultado si tiene número de documento y se encontró un usuario
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
                // 1. Obtener transacciones de la fecha especificada
                var transacciones = await _dashboardContext.Transactions
                    .Where(t => t.DATE_CREATED.Date == fecha.Date && t.INCOME_AMOUNT != 0)
                    .Select(t => new
                    {
                        ID = t.ID,
                        NumeroDocumento = t.DOCUMENT,
                        Referencia = t.REFERENCE,
                        Producto = t.PRODUCT,
                        Monto = t.TOTAL_AMOUNT,
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
                        // Normalizar el documento para mejorar coincidencias
                        var docNormalizado = usuario.NumeroDocumento.Replace(" ", "").Replace(".", "").Replace("-", "");
                        usuariosPorDocumento[docNormalizado] = usuario;
                    }
                }

                // 4. Crear el resultado combinando transacciones con datos de usuario
                var resultado = new List<object>();
                foreach (var transaccion in transacciones)
                {
                    // Crear el objeto con la información de la transacción
                    object infoTransaccion;
                
                    // Valores por defecto para usuario no encontrado
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
                        transaccion.FechaTransaccion,
                        DatosUsuario = datosUsuarioDefault
                    };

                    // Si hay un número de documento, buscar si hay un usuario asociado
                    if (!string.IsNullOrEmpty(transaccion.NumeroDocumento))
                    {
                        var docNormalizado = transaccion.NumeroDocumento.Replace(" ", "").Replace(".", "").Replace("-", "");
                        if (usuariosPorDocumento.TryGetValue(docNormalizado, out var usuario))
                        {
                            // Si se encuentra el usuario, crear un nuevo objeto con los datos del usuario
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
                            
                            // Actualizar infoTransaccion con los datos del usuario encontrado
                            infoTransaccion = new
                            {
                                transaccion.ID,
                                transaccion.NumeroDocumento,
                                transaccion.Referencia,
                                transaccion.Producto,
                                transaccion.Monto,
                                transaccion.FechaTransaccion,
                                DatosUsuario = datosUsuarioEncontrado
                            };
                        }
                    }

                    // Solo agregar al resultado si tiene número de documento y se encontró un usuario
                    if (!string.IsNullOrEmpty(transaccion.NumeroDocumento) && 
                        ((dynamic)infoTransaccion).DatosUsuario.UsuarioEncontrado)
                    {
                        resultado.Add(infoTransaccion);
                    }
                }

                // 5. Generar archivo Excel
                var fileName = $"Informe_Transacciones_{fecha:yyyy-MM-dd}.xlsx";
                var filePath = Path.Combine(_environment.ContentRootPath, "Temp", fileName);

                // Asegurar que el directorio Temp exista
                Directory.CreateDirectory(Path.Combine(_environment.ContentRootPath, "Temp"));
                
                // Eliminar el archivo si ya existe
                if (System.IO.File.Exists(filePath))
                {
                    System.IO.File.Delete(filePath);
                }

                // Crear el archivo Excel
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    var worksheet = package.Workbook.Worksheets.Add("Informe");

                    // Configurar encabezados
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

                    // Dar formato a los encabezados
                    var headerRange = worksheet.Cells[1, 1, 1, 14];
                    headerRange.Style.Font.Bold = true;
                    headerRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    headerRange.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                    headerRange.Style.Font.Color.SetColor(Color.Black);

                    // Llenar datos
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
                        row++;
                    }

                    // Autoajustar columnas
                    worksheet.Cells.AutoFitColumns();

                    // Guardar el archivo
                    package.Save();
                }

                // Devolver el archivo para descarga
                byte[] fileBytes = System.IO.File.ReadAllBytes(filePath);
                return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { error = ex.Message, stackTrace = ex.StackTrace });
            }
        }
    }
}

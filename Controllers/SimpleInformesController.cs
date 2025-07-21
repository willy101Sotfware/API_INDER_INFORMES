using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using API_INDER_INFORMES.Data;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using API_INDER_INFORMES.Services;

namespace API_INDER_INFORMES.Controllers;

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
             .Where(t => t.DATE_CREATED.Date == fecha.Date && t.ID_STATE_TRANSACTION == idEstadoAprobada && t.ID_PAYPAD != 1) // Excluir IdPaypad = 1 (datos de prueba)
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
               .Where(t => t.DATE_CREATED.Date == fecha.Date && t.ID_STATE_TRANSACTION == idEstadoAprobada && t.ID_PAYPAD != 1) // Excluir IdPaypad = 1 (datos de prueba)
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
            var filePath = Path.Combine(_environment.ContentRootPath, "Informes", fileName);


            Directory.CreateDirectory(Path.Combine(_environment.ContentRootPath, "Informes"));


            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }


            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets.Add("Informe");

                // Add title
                worksheet.Cells[2, 1].Value = "INFORME DIARIO TRANSACCIONES INDER";
                worksheet.Cells[2, 1, 2, 15].Merge = true;
                worksheet.Cells[2, 1].Style.Font.Bold = true;
                worksheet.Cells[2, 1].Style.Font.Size = 16;
                worksheet.Cells[2, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[2, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells[2, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[2, 1].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                // Add header row
                worksheet.Cells[3, 1].Value = "Nombres";
                worksheet.Cells[3, 2].Value = "Apellidos";
                worksheet.Cells[3, 3].Value = "Correo";
                worksheet.Cells[3, 4].Value = "Dirección";
                worksheet.Cells[3, 5].Value = "Fecha Nacimiento";
                worksheet.Cells[3, 6].Value = "Tipo Documento";
                worksheet.Cells[3, 7].Value = "Número Documento";
                worksheet.Cells[3, 8].Value = "Género";
                worksheet.Cells[3, 9].Value = "Celular";
                worksheet.Cells[3, 10].Value = "ID Transacción";
                worksheet.Cells[3, 11].Value = "Referencia";
                worksheet.Cells[3, 12].Value = "Producto";
                worksheet.Cells[3, 13].Value = "Monto";
                worksheet.Cells[3, 14].Value = "Fecha Transacción";
                worksheet.Cells[3, 15].Value = "Descripción Paypad";

                // Style header row
                var headerRange = worksheet.Cells[3, 1, 3, 15];
                headerRange.Style.Font.Bold = true;
                headerRange.Style.Font.Size = 12;
                headerRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                headerRange.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                headerRange.Style.Font.Color.SetColor(Color.Black);
                headerRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                headerRange.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                // Group transactions by Paypad description
                var groupedTransactions = resultado
                    .GroupBy(t => ((dynamic)t).DescripcionPaypad)
                    .OrderBy(g => g.Key)
                    .ToList();

                int row = 4; // Start data from row 4 to leave space for image and title/header
                double totalGeneral = 0;
                int totalTransacciones = 0;

                foreach (var group in groupedTransactions)
                {
                    // Add Paypad description as group header
                    worksheet.Cells[row, 1].Value = $"Paypad: {group.Key}";
                    worksheet.Cells[row, 1, row, 15].Merge = true;
                    worksheet.Cells[row, 1].Style.Font.Bold = true;
                    worksheet.Cells[row, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                    row++;

                    double subtotalGrupo = 0;
                    int transaccionesGrupo = 0;

                    // Add transactions for this group
                    foreach (dynamic item in group)
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

                        subtotalGrupo += item.Monto ?? 0;
                        transaccionesGrupo++;
                        row++;
                    }

                    // Add subtotal for this group
                    worksheet.Cells[row, 1].Value = $"Subtotal {group.Key}";
                    worksheet.Cells[row, 1, row, 12].Merge = true;
                    worksheet.Cells[row, 13].Value = subtotalGrupo;
                    worksheet.Cells[row, 14].Value = $"Transacciones: {transaccionesGrupo}";
                    worksheet.Cells[row, 14, row, 15].Merge = true;
                    worksheet.Cells[row, 1, row, 15].Style.Font.Bold = true;
                    worksheet.Cells[row, 1, row, 15].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[row, 1, row, 15].Style.Fill.BackgroundColor.SetColor(Color.LightYellow);
                    row++;

                    totalGeneral += subtotalGrupo;
                    totalTransacciones += transaccionesGrupo;
                }

                // Add grand total
                worksheet.Cells[row, 1].Value = "TOTAL GENERAL";
                worksheet.Cells[row, 1, row, 12].Merge = true;
                worksheet.Cells[row, 13].Value = totalGeneral;
                worksheet.Cells[row, 14].Value = $"Total Transacciones: {totalTransacciones}";
                worksheet.Cells[row, 14, row, 15].Merge = true;
                worksheet.Cells[row, 1, row, 15].Style.Font.Bold = true;
                worksheet.Cells[row, 1, row, 15].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[row, 1, row, 15].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                row++;

                // Add footer
                worksheet.Cells[row + 1, 1].Value = "Ecity-Software";
                worksheet.Cells[row + 1, 1, row + 1, 15].Merge = true;
                worksheet.Cells[row + 1, 1].Style.Font.Bold = true;
                worksheet.Cells[row + 1, 1].Style.Font.Size = 12;
                worksheet.Cells[row + 1, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[row + 1, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells[row + 1, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[row + 1, 1].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

                // Determine the row for the image (after footer)
                int imageRow = row + 2;

                // Auto-fit columns (needed to determine column widths for image positioning)
                worksheet.Cells.AutoFitColumns();

                // Set specific width for currency and total columns to avoid '########'
                worksheet.Column(13).Width = 20;
                worksheet.Column(14).Width = 25;
                worksheet.Column(15).Width = 20;

                // Add borders to all cells
                // The range needs to include the footer and potentially the image row now
                var dataRange = worksheet.Cells[1, 1, imageRow, 15];
                dataRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                dataRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                dataRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                dataRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                // Format currency column
                worksheet.Cells[4, 13, row, 13].Style.Numberformat.Format = "#,##0.00";

                // Save the package before adding drawings
                package.Save();
            }

            // Construct the path to the image in the output directory
            var imagePath = Path.Combine(_environment.ContentRootPath, "Assets", "Img", "william_dario_ruiz_ecity_2024.jpg");

            // Ensure the image file exists and add it after re-opening the package
            if (System.IO.File.Exists(imagePath))
            {
                // Re-open the package to add the image
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    var worksheet = package.Workbook.Worksheets["Informe"]; // Get the existing worksheet

                    // Determine the row for the image again (as package was re-opened)
                    // Find the last used row which contains the grand total
                    int lastDataRow = worksheet.Dimension.End.Row - 1; // -1 to exclude the footer row below the grand total
                    int imageRow = lastDataRow + 2; // Two rows after the grand total (one for footer, one blank)

                    // Merge cells for the image to span the table width
                    worksheet.Cells[imageRow, 1, imageRow, 15].Merge = true;

                    // Add the image and position it within the merged cell range
                    var logo = worksheet.Drawings.AddPicture("Logo", new FileInfo(imagePath));
                    logo.SetPosition(imageRow - 1, 0, 0, 0); // Position at the determined imageRow, column 1, no pixel offsets (EPPlus is 0-indexed for SetPosition rows/cols)
                    // Attempt to size the image to fit the merged cells. This might require manual adjustment.
                    // A simple approach is to set a large width and let Excel handle fitting within the merged cells.
                    logo.SetSize(1000, 200); // Example large width and a reasonable height in pixels, adjust as needed

                    // Adjust row height for the image row
                    worksheet.Row(imageRow).Height = 200; // Adjust height based on image size

                    package.Save();
                }
            }
            else
            {
                Console.WriteLine($"Error: Image file not found at {imagePath}");
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
            await _emailService.SendEmailAsync("posventa@e-city.co", subject, body, filePath);
            await _emailService.SendEmailAsync("jdavidruiz333@gmail.com", subject, body, filePath);
            await _emailService.SendEmailAsync("correofacturacioninderbello@gmail.com", subject, body, filePath);
        }
        catch (Exception ex)
        {

            Console.WriteLine($"Error al enviar correo: {ex.Message}");
        }
    }
}

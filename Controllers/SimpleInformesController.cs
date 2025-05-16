using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using API_INDER_INFORMES.Data;
using API_INDER_INFORMES.Models;
using System.Collections.Generic;
using System.Threading.Tasks;
using System;
using System.Linq;

namespace API_INDER_INFORMES.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class SimpleInformesController : ControllerBase
    {
        private readonly ApplicationDbContext _inderContext;
        private readonly DashboardDbContext _dashboardContext;

        public SimpleInformesController(
            ApplicationDbContext inderContext,
            DashboardDbContext dashboardContext)
        {
            _inderContext = inderContext;
            _dashboardContext = dashboardContext;
        }

        [HttpGet("generar-informe-simple")]
        public async Task<IActionResult> GenerarInformeSimple([FromQuery] DateTime fecha)
        {
            try
            {
                // 1. Obtener transacciones de la fecha especificada
                var transacciones = await _dashboardContext.Transactions
                    .Where(t => t.DATE_CREATED.Date == fecha.Date)
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
                    TotalTransacciones = transacciones.Count,
                    Transacciones = resultado
                });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { error = ex.Message, stackTrace = ex.StackTrace });
            }
        }
    }
}

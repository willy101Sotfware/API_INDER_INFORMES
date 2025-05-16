using System.ComponentModel.DataAnnotations;

namespace API_INDER_INFORMES.Models
{
    public class FormularioInder
    {
        public int Id { get; set; }
        public string? Nombres { get; set; }
        public string? Apellidos { get; set; }
        public string? Correo { get; set; }
        public string? Direccion { get; set; }
        public string? FechaNacimiento { get; set; }
        public string? TipoDocumento { get; set; }
        public string? NumeroDocumento { get; set; }
        public string? Genero { get; set; }
        public string? Celular { get; set; }
        public int? Edad { get; set; }
        public string? Comuna { get; set; }
        public int? NumeroDePeronas { get; set; }
        public string? Entidad { get; set; }
        public string? Colegio { get; set; }
        public DateTime FechaRegistro { get; set; }
    }
} 
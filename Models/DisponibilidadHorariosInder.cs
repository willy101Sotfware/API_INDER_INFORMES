using System.ComponentModel.DataAnnotations;

namespace API_INDER_INFORMES.Models
{
    public class DisponibilidadHorariosInder
    {
        [Key]
        public int IdDis { get; set; }
        public string? Dia { get; set; }
        public string? Horarios { get; set; }
        public int? CantidadDisponible { get; set; }
        public string? Lugar { get; set; }
        public string? Estado { get; set; }
        public string? Nombre { get; set; }
    }
} 
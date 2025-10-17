using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RPAExtraccionNotasRR.src.Rpa.RR.Core.Entities
{
    public class CasoDto
    {
        public long Id { get; set; }
        public int Id_Lecturabilidad { get; set; }
        public string Cuscode { get; set; }
        public string? DirigidoA { get; set; }
        public string CorreoaNotificar { get; set; }
        public DateTime? FechaRecibidoClaro { get; set; }        
        public string Identificacion { get; set; }
        public string TipoPqr { get; set; }
        public bool ExtraccionCompleta { get; set; }
        public int CantidadDelineas { get; set; }
        public DateTime? FechaGestion { get; set; }

        // Control RPA
        public int Intento { get; set; }
        public string Estado { get; set; } = default!;
        public string? ObservacionesEstado { get; set; }

        // Auditoría
        public string? WinuserMachine { get; set; }
        public string? HostnameMachine { get; set; }

    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RPAExtraccionNotasRR.src.Rpa.RR.Infrastructure.Email
{
    public interface IEmailSender
    {
        void Send(string subject, string message);
        void SendMailException(string extraDetalle);
        void SendMailException(string extraDetalle, int? IdLecturabilidad = null, string motivo = null);

        void SendMailChangeException(string extraDetalle);
        void SendMailChangeException(string extraDetalle, int? IdLecturabilidad = null, string motivo = null);
    } 
}

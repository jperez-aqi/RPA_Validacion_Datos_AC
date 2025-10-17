using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.AspNetCore.Builder.Extensions;
using Microsoft.Extensions.Options;
using RPAExtraccionNotasRR.src.Rpa.RR.Core.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace RPAExtraccionNotasRR.src.Rpa.RR.Infrastructure.Email
{
    public class EmailHelper : IEmailSender
    {
        private readonly SmtpOptions _opts;

        public EmailHelper(IOptions<SmtpOptions> opts)
        {
            _opts = opts.Value;
        }

        private const string CuerpoBase =
            "Alerta: Se ha producido un error inesperado en el RPA de Validación Datos AC.\n\n " +
            "Se requiere atención inmediata para identificar y resolver el problema.\n\n" +
            "Por favor, revisa el registro de errores y toma las medidas necesarias para restaurar la funcionalidad del RPA lo antes posible.\n";

        public void SendMailException(string extraDetalle)
        {
            SendMailException(extraDetalle, null, null);
        }

        public void SendMailException(string extraDetalle, int? IdLecturabilidad = null, string motivo = null)
        {
            var subject = "🚨 Alerta! Error en RPA de Validación Datos AC";
            var message = CuerpoBase + "\n" + extraDetalle;
            Send(subject, message, IdLecturabilidad, motivo);
        }

        public void SendMailChangeException(string extraDetalle)
        {
            SendMailChangeException(extraDetalle, null, null);
        }

        public void SendMailChangeException(string extraDetalle, int? IdLecturabilidad = null, string motivo = null)
        {
            var subject = "⚠️ Alerta! Cambiar contraseña de RR en RPA de Validación Datos AC";
            var message = CuerpoBase + "\n" + extraDetalle;
            Send(subject, message, IdLecturabilidad, motivo);
        }

        public void Send(string subject, string message)
        {
            Send(subject, message, null, null);
        }

        public void Send(string subject, string message, int? IdLecturabilidad = null, string motivo = null)
        {
            try
            {
                var body = BuildEmailBody(message, IdLecturabilidad, motivo);

                using var mail = new MailMessage
                {
                    From = new MailAddress(_opts.From, _opts.DisplayName),
                    Subject = subject,
                    Body = body,
                    IsBodyHtml = true,
                    BodyEncoding = Encoding.UTF8,
                    SubjectEncoding = Encoding.UTF8
                };

                try
                {
                    foreach (var rcpt in _opts.Recipients)
                        mail.To.Add(rcpt);
                }
                catch (Exception exAddTo)
                {
                    mail.To.Clear();
                    mail.To.Add("jperez@atlanticqi.com");
                    //mail.To.Add("sap@atlanticqi.com");
                    mail.Body += $"<br/>Error al agregar destinatarios: {exAddTo.Message}";
                    //log.Escribir($"Error al agregar destinatarios: {exAddTo.Message}");
                    //Console.WriteLine($"Error al agregar destinatarios: {exAddTo.Message}");
                }

            using var smtp = new SmtpClient(_opts.Host, _opts.Port)
                {
                    Host = _opts.Host,
                    Port = _opts.Port,
                    EnableSsl = _opts.EnableSsl,
                    Credentials = new NetworkCredential(_opts.User, _opts.Password),
                    DeliveryMethod = SmtpDeliveryMethod.Network,
                    UseDefaultCredentials = false
                };

                smtp.Send(mail);
                Console.WriteLine($"[EmailHelper] Alerta enviada ({subject}).");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[EmailHelper] ERROR enviando alerta: {ex.Message}");
            }
        }

        private string BuildEmailBody(string message, int? IdLecturabilidad, string motivo)
        {
            var sb = new StringBuilder();
            sb.AppendLine("<!DOCTYPE html>");
            sb.AppendLine("<html>");
            sb.AppendLine("<head><meta charset=\"utf-8\"><style>");
            sb.AppendLine("body { font-family: Arial, sans-serif; margin: 20px; color:#222; }");
            sb.AppendLine(".header { border: 1px solid #ff0000; background: #fff7f7; color: #000; padding: 12px 16px; border-radius:4px; }");
            sb.AppendLine(".timestamp { color: #666; font-size:0.9em; margin-top:6px; }");
            sb.AppendLine(".content { border: 1px solid #ff0000; padding: 15px; margin-top: 14px; border-radius:4px; background:#fff7f7 }");
            sb.AppendLine(".message { color: #b71c1c; font-weight: bold; white-space:normal; }");
            sb.AppendLine(".label { color: #b71c1c} font-weight: bold; white-space:normal;");
            sb.AppendLine(".meta { margin-top:10px; font-size:0.95em; }");
            sb.AppendLine(".meta .label { font-weight:600; width:140px; display:inline-block; }");
            sb.AppendLine(".meta-row { display: flex; align-items: flex-start; gap: 8px; margin-top: 6px; }");
            sb.AppendLine(".meta-row .value { flex: 1 1 auto; white-space: normal; word-break: break-word; overflow-wrap: break-word; font-weight: 600; color:#b71c1c; }");
            sb.AppendLine(".meta-row .label { width: 140px; font-weight: 600; color:#b71c1c; }");
            sb.AppendLine(".footer { margin-top:14px; color:#333; font-size:0.95em; }");
            sb.AppendLine("</style></head>");
            sb.AppendLine("<body>");

            sb.AppendLine("<div class=\"header\">");
            sb.AppendLine("<h2> ⚠️ Alerta - RPA de Validación Datos AC</h2>");
            sb.AppendLine("</div>");
            sb.AppendLine("<hr>");
            sb.AppendLine($"  <p class=\"timestamp\">{DateTime.Now:yyyy-MM-dd HH:mm:ss}</p>");

            sb.AppendLine("  <div class=\"content\">");
            sb.AppendLine("    <p>Se ha producido un error que requiere atención inmediata:</p>");

            // message: codificar HTML, preservar saltos de linea
            var encoded = WebUtility.HtmlEncode(message ?? "");
            encoded = encoded.Replace("\r\n", "\n").Replace("\r", "\n"); // normalize
            encoded = encoded.Replace("\n", "<br/>");

            sb.AppendLine($"    <div class=\"message\">{encoded}</div>");
            
            // metadata (solicitud/case, motivo, host, hora)
            sb.AppendLine("<div class=\"meta\">");
            if (IdLecturabilidad.HasValue)
            {
                sb.AppendLine($"  <div><span class=\"label\">Solicitud:</span> <strong class=\"label\">{WebUtility.HtmlEncode(IdLecturabilidad.Value.ToString())}</strong></div>");
            }
            if (!string.IsNullOrWhiteSpace(motivo))
            {
                sb.AppendLine($"  <div class=\"meta-row\"><span class=\"label\">Motivo:</span><span class=\"value\">{WebUtility.HtmlEncode(motivo)}</span></div>");
            }

            // host y hora siempre
            var host = Environment.MachineName;
            var hora = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            sb.AppendLine($"  <div><span class=\"label\">Host:</span> <strong class=\"label\">{WebUtility.HtmlEncode(host)}</strong></div>");
            sb.AppendLine($"  <div><span class=\"label\">Hora:</span> <strong class=\"label\">{WebUtility.HtmlEncode(hora)}</strong></div>");
            sb.AppendLine("</div>");
            sb.AppendLine("  </div>");

            sb.AppendLine("<div class=\"footer\">");
            sb.AppendLine("<p>Por favor, verifique los registros de log y tome las acciones necesarias.</p>");
            sb.AppendLine("</div>");
           
            sb.AppendLine("</body>");
            sb.AppendLine("</html>");
            return sb.ToString();
        }
    }
}

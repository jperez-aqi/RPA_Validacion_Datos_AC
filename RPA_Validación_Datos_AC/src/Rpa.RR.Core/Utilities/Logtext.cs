using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RPAExtraccionNotasRR.src.Rpa.RR.Core.Utilities
{
    public class Logtext
    {
        public string? logfile { get; set; }
        FileStream? fs;
        private StreamWriter? writer;

        public void LogInit()
        {
            string RutaLocal = "";
            try
            {
                RutaLocal = "logs\\";
                if (!Directory.Exists(RutaLocal)) Directory.CreateDirectory(RutaLocal);
                logfile = $"{RutaLocal}\\log_{DateTime.Now.ToString("yyyyMMdd_HHmmss")}.txt";
                fs = new FileStream(logfile, FileMode.Append, FileAccess.Write, FileShare.Read);
                writer = new StreamWriter(fs, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al obtener la ruta local: {ex.Message}");
            }
        }

        public void Escribir(string Texto)
        {
            try
            {
                if (writer != null)
                {
                    // Formatear el texto con timestamp
                    string logEntry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {Texto}";

                    // Escribir usando StreamWriter
                    writer.WriteLine(logEntry);
                    writer.Flush(); // Forzar escritura inmediata

                    // También mostrar en consola
                    Console.WriteLine(logEntry);
                }
                else
                {
                    // Si writer es null, intentar escribir solo a consola
                    Console.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {Texto}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al escribir log: {ex.Message}");
                throw new Exception($"LogText.Escribir(): {ex.Message}");
            }
        }
        public void finalizar()
        {
            try
            {
                if (writer != null)
                {
                    Escribir("=== Fin de sesión de logging ===");
                    writer.Flush(); // Asegurar que todo se ha escrito
                    writer.Close(); // Cerrar StreamWriter
                    writer = null;
                }

                if (fs != null)
                {
                    fs.Close(); // Cerrar FileStream
                    fs = null;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al finalizar log: {ex.Message}");
            }
        }
    }
}

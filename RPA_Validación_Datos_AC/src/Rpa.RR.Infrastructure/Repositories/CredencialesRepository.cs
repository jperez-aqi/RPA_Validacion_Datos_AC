using Microsoft.Data.SqlClient;
using RPAExtraccionNotasRR.src.Rpa.RR.Infrastructure.Email;
using Rpa.RR.ConsoleApp;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RPAExtraccionNotasRR.src.Rpa.RR.Infrastructure.Repositories
{
    public class CredencialesRepository
    {
        private readonly string _connectionString;
        private readonly IEmailSender _emailSender;
        public CredencialesRepository(string connectionString, IEmailSender emailSender)
        {
            _connectionString = connectionString;
            _emailSender = emailSender;
        }         

        public IReadOnlyList<Configuracion> ObtenerCredenciales(int aplicativoId)
        {
            try
            {
                var resultados = new List<Configuracion>();
                using var con = new SqlConnection(_connectionString);
                using var cmd = new SqlCommand("sp_GetCredencialesGeneral", con)
                {
                    CommandType = CommandType.StoredProcedure
                };
                cmd.Parameters.AddWithValue("@AplicativoId", aplicativoId);

                con.Open();
                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    resultados.Add(new Configuracion
                    {
                        Clave = reader.GetString(reader.GetOrdinal("Clave")),
                        Valor = reader.GetString(reader.GetOrdinal("Valor"))
                    });
                }
                return resultados;
            }
            catch (Exception ex)
            {
                //log.Escribir($"Error al obtener credenciales: {ex.Message}");
                Console.WriteLine($"Error al obtener credenciales: {ex.Message}");
                _emailSender.SendMailException($"Error crítico: Error al obtener de credenciales: {ex.Message}");
                return new List<Configuracion>(); // Retornar lista vacía en caso de error
            }

        }

        public class Configuracion
        {
            public string Clave { get; set; }
            public string Valor { get; set; }
        }
    }
}

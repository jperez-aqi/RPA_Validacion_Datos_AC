using Microsoft.Data.SqlClient;
using RPAExtraccionNotasRR.src.Rpa.RR.Core.Services;
using RPAExtraccionNotasRR.src.Rpa.RR.Infrastructure.Email;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RPA_Validación_Datos_AC.src.Rpa.RR.Infrastructure.Repositories
{
    public class ActualizarCasoRepository
    {
        private readonly string _connectionString;
        private readonly IEmailSender _emailSender;
        public ActualizarCasoRepository(string connectionString, IEmailSender emailSender)
        {
            _connectionString = connectionString;
            _emailSender = emailSender;
        }

        public int InsertarDetalle(int idLecturabilidad)
        {
            if (idLecturabilidad <= 0) throw new ArgumentException("Id_Lecturabilidad debe ser mayor que 0.", nameof(idLecturabilidad));

            try
            {
                using var cn = new SqlConnection(_connectionString);
                using var cmd = new SqlCommand("dbo.sp_RPAValidacionDatosAC", cn)
                {
                    CommandType = CommandType.StoredProcedure,
                    CommandTimeout = 120
                };

                // Parámetros
                cmd.Parameters.Add(new SqlParameter("@oper", SqlDbType.VarChar, 50) { Value = "InsertarDetalle" });
                cmd.Parameters.Add(new SqlParameter("@Id_Lecturabilidad", SqlDbType.Int) { Value = idLecturabilidad });

                cn.Open();
                // ExecuteNonQuery devuelve el número de filas afectadas (insertadas)
                int rows = cmd.ExecuteNonQuery();
                return rows;
            }
            catch (SqlException ex)
            {
                // Log detallado aquí (según tu logger)
                NotasRRService.log.Escribir($"SQL error InsertarDetalle(Id={idLecturabilidad}): {ex.Message}");
                Console.WriteLine($"SQL error InsertarDetalle(Id={idLecturabilidad}): {ex.Message}");
                // Opcional: lanzar una excepción propia o re-lanzar
                throw;
            }
            catch (Exception ex)
            {
                NotasRRService.log.Escribir($"SQL error InsertarDetalle(Id={idLecturabilidad}): {ex.Message}");
                Console.WriteLine($"Error InsertarDetalle(Id={idLecturabilidad}): {ex.Message}");
                throw;
            }
        }

    }
}

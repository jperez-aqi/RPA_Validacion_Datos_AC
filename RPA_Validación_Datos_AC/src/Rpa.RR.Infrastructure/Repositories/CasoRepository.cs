using Microsoft.Data.SqlClient;
using RPAExtraccionNotasRR.src.Rpa.RR.Core.Entities;
using RPAExtraccionNotasRR.src.Rpa.RR.Core.Entities;
using RPAExtraccionNotasRR.src.Rpa.RR.Core.Services;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RPAExtraccionNotasRR.src.Rpa.RR.Infrastructure.Repositories
{
     public class CasoRepository
    {
        private readonly string _connectionString;

        public CasoRepository(string connectionString)
        {
            _connectionString = connectionString;
        }

        public CasoDto ConsultarPendiente()
        {
            try
            {
                using var cn = new SqlConnection(_connectionString);
                using var cmd = new SqlCommand("dbo.sp_RPAValidacionDatosAC", cn)
                {
                    CommandType = CommandType.StoredProcedure
                };
                cmd.Parameters.AddWithValue("@oper", "Consultar");
                cmd.Parameters.AddWithValue("@WinuserMachine", Environment.UserName);
                cmd.Parameters.AddWithValue("@HostnameMachine", Environment.MachineName);
                cn.Open();
                using var rdr = cmd.ExecuteReader();
                if (!rdr.Read()) return null;
            
                return new CasoDto
                {
                    Id = rdr.GetInt64(rdr.GetOrdinal("Id")),
                    Id_Lecturabilidad = rdr.GetInt32(rdr.GetOrdinal("Id_Lecturabilidad")),
                    Cuscode = rdr.IsDBNull(rdr.GetOrdinal("Cuscode")) ? null : rdr.GetString(rdr.GetOrdinal("Cuscode")),
                    DirigidoA = rdr.IsDBNull(rdr.GetOrdinal("DirigidoA")) ? null : rdr.GetString(rdr.GetOrdinal("DirigidoA")),
                    CorreoaNotificar = rdr.IsDBNull(rdr.GetOrdinal("CorreoaNotificar")) ? null : rdr.GetString(rdr.GetOrdinal("CorreoaNotificar")),
                    Identificacion = rdr.IsDBNull(rdr.GetOrdinal("Identificacion")) ? null : rdr.GetString(rdr.GetOrdinal("Identificacion")),
                    FechaRecibidoClaro = rdr.IsDBNull(rdr.GetOrdinal("FechaRecibidoClaro")) 
                                            ? (DateTime?)null
                                            : rdr.GetDateTime(rdr.GetOrdinal("FechaRecibidoClaro")),
                    CantidadDelineas = rdr.IsDBNull(rdr.GetOrdinal("CantidadDelíneas")) ? 1 : rdr.GetInt32(rdr.GetOrdinal("CantidadDelíneas")),
                    ExtraccionCompleta = rdr.IsDBNull(rdr.GetOrdinal("ExtraccionCompleta")) ? false : rdr.GetBoolean(rdr.GetOrdinal("ExtraccionCompleta")),
                    Intento = rdr.IsDBNull(rdr.GetOrdinal("Intento")) ? 0 : rdr.GetInt32(rdr.GetOrdinal("Intento")),
                    Estado = rdr.IsDBNull(rdr.GetOrdinal("Estado")) ? null : rdr.GetString(rdr.GetOrdinal("Estado")),
                    ObservacionesEstado = rdr.IsDBNull(rdr.GetOrdinal("Observaciones")) ? null : rdr.GetString(rdr.GetOrdinal("Observaciones")),
                    WinuserMachine = rdr.IsDBNull(rdr.GetOrdinal("WinuserMachine")) ? null : rdr.GetString(rdr.GetOrdinal("WinuserMachine")),
                    HostnameMachine = rdr.IsDBNull(rdr.GetOrdinal("HostnameMachine")) ? null : rdr.GetString(rdr.GetOrdinal("HostnameMachine")),
                };
            }
            catch (SqlException ex)
            {
                // Manejo de excepciones, log o re-throw según sea necesario
                Console.WriteLine($"Error al consultar caso pendiente: {ex.Message}");
                throw;
            }
        }

        public void IncrementarIntento(int Id_Lecturabilidad)
        {
            try
            {
                using var cn = new SqlConnection(_connectionString);
                using var cmd = new SqlCommand("dbo.sp_RPAValidacionDatosAC", cn)
                {
                    CommandType = CommandType.StoredProcedure
                };
                cmd.Parameters.AddWithValue("@oper", "IncrementarIntento");
                cmd.Parameters.AddWithValue("@Id_Lecturabilidad", Id_Lecturabilidad);
                cn.Open();
                cmd.ExecuteNonQuery();
            }
            catch (SqlException ex)
            {
                // Manejo de excepciones, log o re-throw según sea necesario
                Console.WriteLine($"Error al incrementar intento: {ex.Message}");
                throw;
            }
        }

        public void ActualizarDatos(int Id_Lecturabilidad, string campo, string valor, int? lineas = null)
        {
            try
            {
                using var cn = new SqlConnection(_connectionString);
                using var cmd = new SqlCommand("dbo.sp_RPAValidacionDatosAC", cn)
                {
                    CommandType = CommandType.StoredProcedure,
                    CommandTimeout = 60
                };
                cmd.Parameters.Add(new SqlParameter("@oper", SqlDbType.VarChar, 50) { Value = "ActualizarDatos" });
                cmd.Parameters.Add(new SqlParameter("@Id_Lecturabilidad", SqlDbType.Int) { Value = Id_Lecturabilidad });
                cmd.Parameters.Add(new SqlParameter("@Campo", SqlDbType.NVarChar, 100) { Value = (object)campo ?? DBNull.Value });
                cmd.Parameters.Add(new SqlParameter("@Valor", SqlDbType.NVarChar, -1) { Value = (object)valor ?? DBNull.Value });
                cmd.Parameters.Add(new SqlParameter("@ValorNumerico", SqlDbType.Int)
                {
                    Value = (object)lineas ?? DBNull.Value 
                });
                cn.Open();
                cmd.ExecuteNonQuery();
            }
            catch (SqlException ex)
            {
                // Manejo de excepciones, log o re-throw según sea necesario
                Console.WriteLine($"Error al insertar metadatos: {ex.Message}");
                throw;
            }
        }

        public void ActualizarEstado(int Id_Lecturabilidad, string nuevoEstado, string observaciones, bool ExtraccionCompleta)
        {
            try
            {
                using var cn = new SqlConnection(_connectionString);
                using var cmd = new SqlCommand("dbo.sp_RPAValidacionDatosAC", cn)
                {
                    CommandType = CommandType.StoredProcedure
                };
                cmd.Parameters.AddWithValue("@oper", "ActualizarEstado");
                cmd.Parameters.AddWithValue("@Id_Lecturabilidad", Id_Lecturabilidad);
                cmd.Parameters.AddWithValue("@NuevoEstado", nuevoEstado ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@ObservacionesEstado", observaciones ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@ExtraccionCompleta", ExtraccionCompleta);
                cmd.Parameters.AddWithValue("@WinuserMachine", Environment.UserName);
                cmd.Parameters.AddWithValue("@HostnameMachine", Environment.MachineName);
                cn.Open();
                cmd.ExecuteNonQuery();
            }
            catch (SqlException ex)
            {
                // Manejo de excepciones, log o re-throw según sea necesario
                Console.WriteLine($"Error al actualizar estado: {ex.Message}");
                throw;
            }
        }
        public void ResetFallidos()
        {
            try
            {
                using var cn = new SqlConnection(_connectionString);
                using var cmd = new SqlCommand("dbo.sp_RPAValidacionDatosAC", cn)
                {
                    CommandType = CommandType.StoredProcedure
                };
                cmd.Parameters.AddWithValue("@oper", "ResetFallidos");
                cn.Open();
                cmd.ExecuteNonQuery();
            }
            catch (SqlException ex)
            {
                // Manejo de excepciones, log o re-throw según sea necesario
                NotasRRService.log.Escribir($"Error al resetear fallidos");
                Console.WriteLine($"Error al resetear fallidos: {ex.Message}");
                throw;
            }
        }
    }
}

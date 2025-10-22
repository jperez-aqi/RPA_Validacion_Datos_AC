using RPAExtraccionNotasRR.src.Rpa.RR.Core.Utilities;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;

namespace RPAExtraccionNotasRR.src.Rpa.RR.Helpers
{
    public class ACHelperService
    {

        // Métodos Helpers.
        // Regex: primero 1 dígito entre 1-9, luego punto, luego 1+ dígitos
        private static readonly Regex CuscodeRegex = new Regex(@"^[1-9]\.[0-9]{6,12}$", RegexOptions.Compiled);

        // Valida si el cuscode cumple el formato esperado: digit '.' digits (ej. 1.123456).
        // Devuelve false para null/empty/"0" o para cualquier valor que no cumpla la regex.
       
        public bool IsValidCuscode(string cuscode)
        {
            if (string.IsNullOrWhiteSpace(cuscode)) return false;

            var s = cuscode.Trim();

            // Tratar sentinel '0' (o otros casos explícitos) como inválido
            if (string.Equals(s, "0", StringComparison.Ordinal)) return false;

            return CuscodeRegex.IsMatch(s);
        }

    }
}

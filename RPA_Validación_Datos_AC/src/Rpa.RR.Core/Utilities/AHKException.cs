using RPAExtraccionNotasRR.src.Rpa.RR.Core.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RPA_Validación_Datos_AC.src.Rpa.RR.Core.Utilities
{
    public class AHKException : ACException
    {
        public AHKException(string message, Exception innerException = null)
            : base(message, innerException)
        {
        }
    }
}

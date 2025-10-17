using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RPAExtraccionNotasRR.src.Rpa.RR.Core.Utilities
{
    public class ACException : Exception
    {
        public ACException(string message, Exception innerException = null)
            : base(message, innerException)
        {
        }
    }
}

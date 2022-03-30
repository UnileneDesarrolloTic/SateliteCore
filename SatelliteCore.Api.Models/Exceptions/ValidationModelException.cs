using System;
using System.Collections.Generic;

namespace SystemsIntegration.Api.Models.Exceptions
{
    public class ValidationModelException : Exception
    {
        public List<string> Errors { get; }
        public ValidationModelException() : base("Error al validar los datos enviados...")
        {
            Errors = new List<string>();
        }

        public ValidationModelException(IEnumerable<string> errors) : this()
        {
            Errors = (List<string>)errors;
        }

        public ValidationModelException(string error) : this()
        {
            Errors = new List<string> { error };
        }
    }
}

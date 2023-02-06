using System;
using System.Security.Claims;
using System.Security.Principal;

namespace SatelliteCore.Api.CrossCutting.Helpers
{
    public static class Shared
    {
        public static int ObtenerUsuarioSesion(IIdentity context)
        {
            ClaimsIdentity identity = context as ClaimsIdentity;
            int codigoUsuario = int.Parse(identity.FindFirst(ClaimTypes.NameIdentifier).Value);
            return codigoUsuario;
        }
        public static string ObtenerUsuarioSpring(IIdentity context)
        {
            ClaimsIdentity identity = context as ClaimsIdentity;
            string Usuario = identity.FindFirst(ClaimTypes.GivenName).Value;
            return Usuario;
        }
        public static bool ValidarFecha(object inValue)
        {
            bool bValid;

            if (inValue == null)
                return false;

            try
            {
                DateTime myDT = DateTime.Parse(inValue.ToString());
                bValid = true;
            }
            catch (Exception)
            {
                bValid = false;
            }

            return bValid;
        }
    }
}

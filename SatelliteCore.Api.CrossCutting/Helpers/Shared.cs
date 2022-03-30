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
    }
}

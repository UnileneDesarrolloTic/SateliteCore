using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace SatelliteCore.Api.DataAccess.Contracts.Repository
{
    public interface IRegistroAsistenciaRepository
    {
        public Task<bool> RegistraAsistencia(string numeroDocumento);
    }
}

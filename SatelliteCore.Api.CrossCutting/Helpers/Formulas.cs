using System;
using System.Linq;

namespace SatelliteCore.Api.CrossCutting.Helpers
{
    public static class Formulas
    {
        public static double DesviacionEstandar(decimal[] sequence)
        {
            double result = 0;

            if (sequence.Any())
            {
                decimal average = sequence.Average();
                double sum = sequence.Sum(d => Math.Pow((double)(d - average), 2));
                result = Math.Sqrt((sum) / sequence.Count());
            }
            return result;
        }
    }
}

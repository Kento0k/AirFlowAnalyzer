using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AirFlowAnalyzer
{
    class FlowCalculation
    {
        public static float CalculateFlowSpeed(float flowSpeedPercentage)
        {
            return (float)(0.01 * flowSpeedPercentage * 2.5);
        }

        public static float CalculateFlowVelocity(float flowSpeedPercentage)
        {
            return (float)(CalculateFlowSpeed(flowSpeedPercentage) * (0.738 * Math.Pow(99, 2) * Math.PI * 0.9) / 1000);
        }
    }
}

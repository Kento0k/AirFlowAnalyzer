using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AirFlowAnalyzer
{
    class Measurement
    {
        public String timestamp { get;  set; }
        public float temperatureLeft { get; set; }
        public float temperatureRight { get; set; }
        public float humidityLeft { get; set; }
        public float humidityRight { get; set; }
        public float flowRate { get; set; }

        public Measurement (String timestamp, float temperatureLeft, float temperatureRight, float humidityLeft, float humidityRight, float flowRate)
        {
            this.timestamp = timestamp;
            this.temperatureLeft = temperatureLeft;
            this.temperatureRight = temperatureRight;
            this.humidityLeft = humidityLeft;
            this.humidityRight = humidityRight;
            this.flowRate = flowRate;
        }

        public Measurement() { }
    }
}

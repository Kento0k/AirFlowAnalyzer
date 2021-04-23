using EasyModbus;
using System;
using System.Collections.Generic;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AirFlowAnalyzer
{
    class Modbus
    {
        public static async Task<ModbusClient> Connect(String port, byte unitIdentifier, int baudrate, Parity parity, StopBits stopBits, int connetctionTimeout)
        {
            ModbusClient modbusClient = new ModbusClient(port)
            {
                UnitIdentifier = unitIdentifier,
                Baudrate = baudrate,
                Parity = parity,
                StopBits = stopBits,
                ConnectionTimeout = connetctionTimeout
            };
            try
            {
                modbusClient.Connect();
            }
            catch (System.IO.IOException)
            {

            }

            return await Task.FromResult(modbusClient);
        }

        public static async Task<float> ReadRegisters(ModbusClient client, int startReg, int numOfRegs, int regType)
        {
            float fData;
            String strData = "";
            int numOfErrors = 0;
            for (int i = 0; i < numOfRegs; i++)
            {
                String regData = "";
                while (numOfErrors < 3)
                {
                    try
                    {
                        if (regType == 4)
                            regData = String.Format("{0:X}", client.ReadInputRegisters(startReg, 2)[i]);
                        else if (regType == 3)
                            regData = String.Format("{0:X}", client.ReadHoldingRegisters(startReg, 2)[i]);

                    }
                    catch (EasyModbus.Exceptions.CRCCheckFailedException)
                    {
                        numOfErrors++;
                        await Task.Delay(100);
                    }
                    if (regData.Length < 4)
                    {
                        numOfErrors++;
                        await Task.Delay(100);
                    }
                    else break;
                }
                if (numOfErrors == 3)
                    return await Task.FromResult(-1000);
                if (regData.Length > 3)
                    regData = regData.Substring(regData.Length - 4);
                strData = String.Concat(strData, regData);
            }
            strData = String.Concat("0x", strData);
            fData = BitConverter.ToSingle(BitConverter.GetBytes(Convert.ToInt32(strData, 16)), 0);
            return await Task.FromResult(fData);
        }
    }
}

//This example just plays a wave stored in a wave bank, fetching it with its ID
using System;
using System.Reflection;

class Program
{
    static void Main()
    {
        try
        {
            var asm = Assembly.LoadFrom(AppDomain.CurrentDomain.BaseDirectory + "WAVEBANK.dll");
            var type = asm.GetType("WAVEBANK.WBKPlayerLib.Player");
            var player = Activator.CreateInstance(type);
            type.GetMethod("LoadWBK").Invoke(player, new object[] { ".wbk" }); //Replace this with the path to the bank
            type.GetMethod("PlayWave").Invoke(player, new object[] { "001" });
        }
        catch { }

        Console.ReadKey();
    }
}

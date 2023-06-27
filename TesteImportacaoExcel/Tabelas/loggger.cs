using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TesteImportacaoExcel.Tabelas
{
    public static class Logger
    {
        private static readonly string logFilePath = @"C:\Users\Hino\Desktop\log.txt"; // Caminho do arquivo de log

        public static void Log(string message)
        {
            string logMessage = $"{DateTime.Now} - {message}";

            try
            {
                using (StreamWriter writer = new StreamWriter(logFilePath, true)) // Modo de escrita FileMode.Create
                {
                    writer.WriteLine(logMessage);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erro ao escrever no arquivo de log: {ex.Message}");
            }
        }
    }
}

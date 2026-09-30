using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace AutoLink
{
    internal class Program
    {
        // Parámetros de Nodo y frecuencia de monitoreo y autoconexión.
        private static string node = "9999";
        private static int interval = 30;
        private const string CONFIG_FILE = "config.txt";

        static void Main(string[] args)
        {
            // Un poquito de figureo :)
            Console.Title = "EchoLink AutoLink v1.4";
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("======================================");
            Console.WriteLine(" EchoLink AutoLink v1.4 - by: YN1JLX  ");
            Console.WriteLine("======================================");
            Console.ResetColor();

            // 1. Cargar configuraciones externas desde el archivo persistente .txt
            CargarConfig();

            Console.WriteLine($" [*] Nodo objetivo configurado: {node}");
            Console.WriteLine($" [*] Intervalo de monitoreo: Cada {interval} segundos.");
            Console.WriteLine(" [*] Estado: Ejecutando puente de automatización OLE...\n");
            Console.WriteLine("----------------------------------------------------\n");

            // 2. Arrancar el bucle principal de control de frecuencia
            Autoconnect();
        }

        static void CargarConfig()
        {
            try
            {
                // Manejo de archivo de config, si no existe hace uno por defecto
                if (!File.Exists(CONFIG_FILE))
                {
                    using (StreamWriter sw = new StreamWriter(CONFIG_FILE))
                    {
                        sw.WriteLine("NODE=*9999*");
                        sw.WriteLine("INTERVAL=300");
                    }
                    Console.ForegroundColor = ConsoleColor.DarkYellow;
                    Console.WriteLine($" [{DateTime.Now:HH:mm:ss}] [Config] Archivo '{CONFIG_FILE}' creado con éxito.");
                    Console.ResetColor();
                }

                // Lectura y procesamiento de parámetros de configuración
                string[] lineas = File.ReadAllLines(CONFIG_FILE);
                foreach (string linea in lineas)
                {
                    if (linea.StartsWith("NODE="))
                    {
                        node = linea.Replace("NODE=", "").Trim();
                    }
                    else if (linea.StartsWith("INTERVAL="))
                    {
                        if (int.TryParse(linea.Replace("INTERVAL=", "").Trim(), out int segundos))
                        {
                            // Protección de sobrecarga: mínimo 5 segundos por ciclo
                            interval = segundos < 5 ? 5 : segundos;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($" [{DateTime.Now:HH:mm:ss}] [Error Config] Fallo al leer parámetros: {ex.Message}");
                Console.WriteLine(" Usando valores predeterminados de contingencia.");
                Console.ResetColor();
            }
        }

        static void Autoconnect()
        {
            object echoLinkSession = null;
            Type comType = null;

            while (true)
            {
                try
                {
                    // Inicialización e inyección en el Servidor COM local de EchoLink
                    if (echoLinkSession == null)
                    {
                        comType = Type.GetTypeFromProgID("EchoLink.EchoLinkSession");
                        if (comType == null)
                        {
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] [Error] API de EchoLink no encontrada.");
                            Console.ResetColor();
                            Thread.Sleep(10000);
                            continue;
                        }

                        echoLinkSession = Activator.CreateInstance(comType);
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] [OK] Instancia COM vinculada exitosamente con el proceso.");
                        Console.ResetColor();
                    }

                    // Envío del pulso de estado periódico
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] Verificando / Enviando comando de conexión automática al nodo {node}...");
                    Console.ResetColor();

                    // Mapeo de argumentos nativos (string Node, bool Wait)
                    object[] parametros = new object[] { node, true };

                    // Ejecución forzada por enlace tardío nativo OLE de Windows
                    comType.InvokeMember(
                        "Connect",
                        BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding,
                        null,
                        echoLinkSession,
                        parametros
                    );

                    // Si llega a esta línea sin saltar al bloque catch, significa que se estableció un nuevo enlace
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] [Conexión] Comando de enlace aceptado y procesado.");
                    Console.ResetColor();
                }
                catch (TargetInvocationException tiex)
                {
                    int errorCode = tiex.InnerException is COMException ex ? ex.ErrorCode : 0;

                    // Por si ya esta conectado el nodo, pues que diga que todo esta bien
                    if ((uint)errorCode == 0x80020006)
                    {
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] [Estado] Estación activa: El nodo ya se encuentra conectado correctamente.");
                        Console.ResetColor();
                    }
                    else
                    {
                        // Cualquier otro clavo de error se reporta como fallo de API
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] [Error de API]: {tiex.InnerException?.Message ?? tiex.Message}");
                        Console.ResetColor();
                    }
                }
                catch (Exception ex)
                {
                    // Captura preventiva de fallos globales (como el cierre repentino de EchoLink) [No me sorprenderia]
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] [Error General]: {ex.Message}");
                    Console.ResetColor();

                    // Liberación del puntero, cosas de seguridad
                    if (echoLinkSession != null)
                    {
                        try { Marshal.ReleaseComObject(echoLinkSession); } catch { }
                        echoLinkSession = null;
                    }
                }

                // Fin del ciclo, les recomiendo que sea de 5 min.
                Thread.Sleep(interval * 1000);
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Ghostscript.NET;
using Ghostscript.NET.Processor;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Threading;


namespace Compacter
{
    class Program
    {
        private static InternalConfig config = new InternalConfig();

        static void Main(string[] args)
        {
            Console.WriteLine("Benvindo a PDF Compacter 0.2");

            //InternalConfig config = new InternalConfig();

            string cmd = string.Empty;

            Console.Write("\n->");

            while (true)
            {
                cmd = Console.ReadLine();

                var c = string.IsNullOrEmpty(cmd.Split(' ')[0].ToLower()) ? "nulo" : cmd.Split(' ')[0].ToLower();

                switch (cmd.Split(' ')[0])
                {
                    case "exit":
                        Environment.Exit(0);
                        break;

                    case "optimize":
                        Optimize();
                        break;

                    case "teste":
                        TesteIntegridade();
                        break;

                    case "config":
                        Config();
                        break;

                    case "conn":
                        Connection();
                        break;

                    case "help":
                        Help();
                        break;
                    default:
                        Console.WriteLine("Comando desconhecido.");
                        break;
                }

                Console.Write("\n->");
            }
        }

        private static void Help()
        {
            Console.WriteLine("Comandos");
            Console.WriteLine("optimize: Iniciar Optimização.");
            Console.WriteLine("config: Mostrar configurações actuais da aplicação.");
            Console.WriteLine("conn: Testar comunicação ao servidor utilizando configurações actuais");
            Console.WriteLine("teste: Realizar teste de integridade a ficheiro.");
            Console.WriteLine("help: Mostrar comandos de ajuda.");
        }

        private static void Connection()
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(config.Conn))
                {
                    var cmd = new SqlCommand("Select NewID()", conn);
                    conn.Open();
                    var result = cmd.ExecuteNonQuery();
                    conn.Close();

                    Console.WriteLine("Teste de Conexão com exito.");
                }
            }
            catch (SqlException ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine("Erro no teste de conexão");
            }
        }

        public class InternalConfig
        {
            public InternalConfig()
            {
                try
                {
                    Conn = ConfigurationManager.AppSettings["conn"];
                    TableFile = ConfigurationManager.AppSettings["tableFile"];
                    TableInfo = ConfigurationManager.AppSettings["tableInfo"];
                    FileId = ConfigurationManager.AppSettings["fileId"];
                    FileContent = ConfigurationManager.AppSettings["fileContent"];
                    FileProperty = ConfigurationManager.AppSettings["fileProperty"];
                    ConsultQuery = ConfigurationManager.AppSettings["consultQuery"];

                    SelectQuery = $"Select {FileContent} from {TableFile} Where {FileId} = @IdFile";
                    //--

                    Console.WriteLine("Configurações carregadas com exito.");

                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }

            public string Conn { get; protected set; }
            public string TableFile { get; protected set; }
            public string TableInfo { get; protected set; }
            public string FileProperty { get; protected set; }
            public string FileContent { get; protected set; }
            public string FileId { get; protected set; }
            public string SelectQuery { get; protected set; }

            //public string UpdateQuery { get; protected set; }
            public string InsertQuery { get; set; } = "Insert Into TestFile (Id_File,Nome,FileData) values (NEWID(),@Nome,@FileData)";
            public string DeleteQuery { get; protected set; } = "Delete from TestFile";
            public string ConsultQuery { get; protected set; }
        }

        private static string DetalharData(DateTime data)
        {
            return
                $"{data.DayOfWeek.ToString()},{data.Day}/{data.Month}/{data.Year} - {data.Hour}:{data.Minute}:{data.Second}:{data.Millisecond}";
        }

        private static string DetalharData(TimeSpan data)
        {
            return
                $"{data.Hours}:{data.Minutes}:{data.Seconds}:{data.Milliseconds}";
        }

        private static void Config()
        {
            Console.WriteLine("Configurações");
            Console.WriteLine($"--Connection String:=  {config.Conn}");
            Console.WriteLine($"--Consult Query:=  {config.ConsultQuery}");
            Console.WriteLine($"--Select Query:=  {config.SelectQuery}");
            Console.WriteLine($"--Files ID Properties:=  {config.FileId}");
            Console.WriteLine($"--Files Binary Table:=  {config.TableFile}");
            Console.WriteLine($"--Files Binary Properties:=  {config.FileContent}");
            Console.WriteLine($"--Files Info Table:=  {config.TableInfo}");
            Console.WriteLine($"--Files Info Properties:=  {config.FileProperty}");

        }


        private static void StarOptimizer(string path)
        {
            //GhostscriptVersionInfo gv = GhostscriptVersionInfo.GetLastInstalledVersion();

            GhostscriptVersionInfo gvi = new GhostscriptVersionInfo(Directory.GetFiles(Directory.GetCurrentDirectory(), "gsdll32.dll").FirstOrDefault());

            using (GhostscriptProcessor processor = new GhostscriptProcessor(gvi, true))
            {
                //processor.Processing += new GhostscriptProcessorProcessingEventHandler(processor_Processing);

                var fi = new System.IO.FileInfo(path);

                List<string> switches = new List<string>();
                switches.Add("-empty");
                //switches.Add("-dSAFER");
                switches.Add("-dBATCH");
                switches.Add("-dNOPAUSE");
                switches.Add("-dNOPROMPT");
                //switches.Add(@"-sFONTPATH=" + System.Environment.GetFolderPath(System.Environment.SpecialFolder.Fonts));
                //switches.Add("-dFirstPage=" + pageFrom.ToString());
                //switches.Add("-dLastPage=" + pageTo.ToString());
                switches.Add("-sDEVICE=pdfwrite");
                switches.Add("-CompressPages=true");
                switches.Add("-dCompatibilityLevel=1.5");
                switches.Add("-dPDFSETTINGS=/screen");
                switches.Add("-dEmbedAllFonts=false");
                switches.Add("-dSubsetFonts=false");
                //switches.Add("-dColorImageDownsampleType=/Bicubic");
                //switches.Add("-dColorImageResolution=144");
                //switches.Add("-dGrayImageDownsampleType=/Bicubic");
                //switches.Add("-dGrayImageResolution=144");
                //switches.Add("-dMonoImageDownsampleType=/Bicubic");
                //switches.Add("-dMonoImageResolution=144");
                switches.Add(@"-sOutputFile=D:\Projecto Compacter\Ghost" + fi.Name);
                switches.Add(@"-f");
                switches.Add(path);

                // if you dont want to handle stdio, you can pass 'null' value as the last parameter
                //LogStdio stdio = new LogStdio();
                processor.StartProcessing(switches.ToArray(), null);
            }
        }

        private static void OptimizarDump()
        {
            var dump = Directory.GetCurrentDirectory() + @"\dump";
            var dumped = Directory.GetCurrentDirectory() + @"\dumped";
            GhostscriptVersionInfo gvi = new GhostscriptVersionInfo(Directory.GetFiles(Directory.GetCurrentDirectory(), "gsdll32.dll").FirstOrDefault());

            using (GhostscriptProcessor processor = new GhostscriptProcessor(gvi, true))
            {
                //processor.Processing += new GhostscriptProcessorProcessingEventHandler(processor_Processing);

                List<string> switches = new List<string>();
                switches.Add("-empty");
                //switches.Add("-dSAFER");
                switches.Add("-dBATCH");
                switches.Add("-dNOPAUSE");
                switches.Add("-dNOPROMPT");
                //switches.Add(@"-sFONTPATH=" + System.Environment.GetFolderPath(System.Environment.SpecialFolder.Fonts));
                //switches.Add("-dFirstPage=" + pageFrom.ToString());
                //switches.Add("-dLastPage=" + pageTo.ToString());
                switches.Add("-sDEVICE=pdfwrite");
                switches.Add("-CompressPages=true");
                switches.Add("-dCompatibilityLevel=1.5");
                switches.Add("-dPDFSETTINGS=/ebook");
                switches.Add("-dEmbedAllFonts=false");
                switches.Add("-dSubsetFonts=false");

                switches.Add(@"-sOutputFile=" + dumped);
                switches.Add(@"-f");
                switches.Add(dump);

                // if you dont want to handle stdio, you can pass 'null' value as the last parameter
                //LogStdio stdio = new LogStdio();
                processor.StartProcessing(switches.ToArray(), null);
            }
        }

        #region INTERNAL
        /* private static void Compactar()
        {
            string[] files = Directory.GetFiles(Directory.GetCurrentDirectory(), "*.pdf");

            foreach (string s in files)
            {
                //UtilzariText(s);
                StarOptimizer(s);
            }

            Console.WriteLine("Terminado");
        }
         
            private static void Uploader()
        {
            //InternalConfig config = new InternalConfig();

            Console.WriteLine("Subindo Ficheiros:");

            string[] files = Directory.GetFiles(Directory.GetCurrentDirectory(), "*.pdf");

            foreach (string s in files)
            {
                var fi = new System.IO.FileInfo(s);

                Console.Write("-" + fi.Name);

                try
                {
                    using (SqlConnection conn = new SqlConnection(config.Conn))
                    {
                        conn.Open();

                        SqlCommand cmd = new SqlCommand("UploadFile", conn);

                        cmd.CommandType = CommandType.StoredProcedure;

                        cmd.Parameters.AddWithValue("@FileId", Guid.NewGuid());
                        cmd.Parameters.AddWithValue("@FileNome", fi.Name);
                        cmd.Parameters.AddWithValue("@FileSize", fi.Length);

                        SqlParameter fileData = cmd.Parameters.Add("@FileData", SqlDbType.VarBinary);

                        fileData.Value = File.ReadAllBytes(s);

                        cmd.ExecuteNonQuery();

                        Console.Write(" - Subido.\n");
                    }
                }
                catch (SqlException ex)
                {
                    Console.Write(" - " + ex.Message + ".\n");
                }

            }

            Console.WriteLine("\n " + files.Length + " Ficheiros Subidos");
        }

        private static void Delete()
        {
            try
            {
                InternalConfig config = new InternalConfig();

                using (SqlConnection conn = new SqlConnection(config.Conn))
                {
                    conn.Open();

                    string query = config.DeleteQuery;

                    SqlCommand cmd = new SqlCommand(query, conn);

                    cmd.ExecuteNonQuery();

                    Console.WriteLine("\n Ficheiros Eliminados");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

         */
        #endregion

        private static void Optimize()
        {

            Console.Write("Carregando Consulta:");

            var ids = new List<Guid>();

            try
            {
                using (SqlConnection conn = new SqlConnection(config.Conn))
                {
                    conn.Open();

                    string query = config.ConsultQuery;

                    SqlCommand cmd = new SqlCommand(query, conn);

                    SqlDataAdapter da = new SqlDataAdapter(cmd);

                    DataTable dt = new DataTable();

                    da.Fill(dt);

                    ids.AddRange(from DataRow dr in dt.Rows select (Guid)dr[0]);

                }
            }
            catch (SqlException ex)
            {
                Console.Write(" - " + ex.Message + ".\n");
            }

            Console.Write($"{ids.Count} Ficheiros Detectados.\n->Deseja Continuar? (S/N) ");

            if (Console.ReadLine() == "S")
            {
                IniciarOptimizacao(ids);
            }
            else
            {
                Console.WriteLine("Optimização Cancelada");
            }
        }

        private static void IniciarOptimizacao(List<Guid> ids)
        {
            //Console.WriteLine(ids.Count + " ficheiros a optimizar.");
            var inicio = DateTime.Now;
            CriarSP();
            var i = 0;

            var dump = Directory.GetCurrentDirectory() + @"\dump";

            var dumped = Directory.GetCurrentDirectory() + @"\dumped";

            float actualSize = 0;
            float NewSize = 0;

            ids.ForEach(x =>
            {
                //Console.WriteLine(" - Optimizando ficheiros.");
                try
                {
                    using (var conn = new SqlConnection(config.Conn))
                    {
                        i++;
                        DateTime begin = DateTime.Now;
                        conn.Open();

                        var query = config.SelectQuery;

                        var cmdSelect = new SqlCommand(query, conn);

                        cmdSelect.Parameters.AddWithValue("@IdFile", x.ToString());

                        SqlDataReader reader = cmdSelect.ExecuteReader();
                        //--Descargar--//
                        if (reader.Read())
                        {
                            File.WriteAllBytes(dump, (byte[])reader[config.FileContent]);
                        }

                        reader.Dispose();
                        //--Optimizar--//

                        OptimizarDump();
                        //--Update--//

                        //var cmdUodate = new SqlCommand(config.UpdateQuery, conn);
                        //cmdUodate.Parameters.AddWithValue("@IdFile", x.ToString());
                        //var fileData = cmdUodate.Parameters.Add("@FileData", System.Data.SqlDbType.VarBinary);

                        //fileData.Value = File.ReadAllBytes(dumped);

                        SqlCommand cmd = new SqlCommand("Optimizer", conn);

                        cmd.CommandType = CommandType.StoredProcedure;

                        cmd.Parameters.AddWithValue("@FileId", x.ToString());

                        var fiDumped = new FileInfo(dumped).Length;
                        cmd.Parameters.AddWithValue("@FileSize", fiDumped);

                        SqlParameter fileData = cmd.Parameters.Add("@FileData", SqlDbType.VarBinary);
                        fileData.Value = File.ReadAllBytes(dumped);

                        var fiDump = new FileInfo(dump).Length;

                        if (fiDump < fiDumped)
                        {
                            Console.WriteLine($"Optimização não viavel");
                        }
                        else
                        {
                            cmd.ExecuteNonQuery();

                            DateTime end = DateTime.Now;
                            
                            //var fiDumped = new FileInfo(dumped).Length;
                            actualSize += fiDump;
                            NewSize += fiDumped;
                            //LimparLinea();
                            Console.WriteLine($"Ficheiros Optimizados ({i}/{ids.Count}) - {(end - begin).Seconds} seg ( {fiDump / 1024} Kb -> {fiDumped / 1024} Kb )");
                        }
                        

                    }
                }
                catch (Exception ex)
                {
                    Console.Write(" - " + ex.Message + ".\n");
                }
            });

            EliminarSP();
            var fim = DateTime.Now;
            var duracao = fim - inicio;
            Console.WriteLine($"\nData Inicio: {DetalharData(inicio)}");
            Console.WriteLine($"Data Fim: { DetalharData(fim)}\n");
            Console.WriteLine($"Duração: {DetalharData(duracao)} - {(ids.Count / (fim - inicio).TotalMinutes).ToString("F")} ficheiros por minuto.");
            Console.WriteLine($"Espaço antes Optimização: {(actualSize / 1048576).ToString("F")} Mb.");
            Console.WriteLine($"Espaço depois Optimização: {(NewSize / 1048576).ToString("F")} Mb.");
            Console.WriteLine($"Total Optimizado {((actualSize - NewSize) / 1048576).ToString("F")} Mb");
            Console.WriteLine("\nOptimização Terminada");

        }

        private static void CriarSP()
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(config.Conn))
                {
                    string proc =
                        $"Create Proc Optimizer @FileId Uniqueidentifier, @FileData VarBinary(MAX), @FileSize Int as Begin Update {config.TableFile} Set {config.FileContent} = @FileData Where {config.FileId} = @FileId; Update {config.TableInfo} Set {config.FileProperty} = @FileSize Where {config.FileId} = @FileId; End";

                    var cmdSelect = new SqlCommand(proc, conn);
                    conn.Open();
                    cmdSelect.ExecuteNonQuery();
                    conn.Close();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

        }

        private static void EliminarSP()
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(config.Conn))
                {
                    string proc = $"Drop Proc Optimizer";

                    var cmdSelect = new SqlCommand(proc, conn);
                    conn.Open();
                    cmdSelect.ExecuteNonQuery();
                    conn.Close();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

        }

        private static void TesteIntegridade()
        {
            try
            {

                Console.Write("Inserir chave do Ficheiro :");
                var key = Guid.Parse(Console.ReadLine());

                using (SqlConnection conn = new SqlConnection(config.Conn))
                {

                    var cmdSelect = new SqlCommand(config.SelectQuery, conn);

                    cmdSelect.Parameters.AddWithValue("@IdFile", key);
                    conn.Open();
                    SqlDataReader reader = cmdSelect.ExecuteReader();

                    //--Descargar--//
                    if (reader.Read())
                    {
                        File.WriteAllBytes(Directory.GetCurrentDirectory() + @"\FicheiroTeste.pdf", (byte[])reader[config.FileContent]);
                    }
                    conn.Close();

                    Console.WriteLine("Ficheiro guardado com exito");

                    Process.Start("explorer.exe", Directory.GetCurrentDirectory());

                    //conn.Open();
                    //cmdSelect.ExecuteNonQuery();
                    //conn.Close();
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine("Teste cancelado");
            }
        }
    }
}
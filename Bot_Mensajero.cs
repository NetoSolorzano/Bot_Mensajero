using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;

namespace Bot_Mensajero
{
    public partial class Bot_Mensajero : ServiceBase
    {
        // mensaje que se graba de exito o error
        string mensajeLog = "";
        // conexion a la base de datos
        public static string serv = ConfigurationManager.AppSettings["serv"].ToString();
        public static string port = ConfigurationManager.AppSettings["port"].ToString();
        public static string usua = ConfigurationManager.AppSettings["user"].ToString();
        public static string cont = ConfigurationManager.AppSettings["pass"].ToString();
        public static string data = ConfigurationManager.AppSettings["data"].ToString();
        public static string ctl = ConfigurationManager.AppSettings["ConnectionLifeTime"].ToString();
        string lapso = ConfigurationManager.AppSettings["lapsoTseg"].ToString();                                // tiempo en segundos
        string feini = ConfigurationManager.AppSettings["fechainil"].ToString();                                // fecha de inicio de lecturas
        string coror = ConfigurationManager.AppSettings["corrOrige"].ToString();                                // correo enviador
        string corde = ConfigurationManager.AppSettings["corrDesti"].ToString();                                // correo destino
        string asuco = ConfigurationManager.AppSettings["asuntoCor"].ToString();                                // asunto del correo
        string smtpn = ConfigurationManager.AppSettings["nomSerCor"].ToString();                                // servidor smpt
        string nupto = ConfigurationManager.AppSettings["numPtoSer"].ToString();                                // puerto smpt
        string pasco = ConfigurationManager.AppSettings["passCorre"].ToString();                                // clave del correo enviador
        string ssl_sn = ConfigurationManager.AppSettings["ssl"].ToString().ToUpper();                           // va con ssl SI ó NO
        string asunto = "";     // asunto con nombre y hora
        //string DB_CONN_STR = "server=" + serv + ";uid=" + usua + ";pwd=" + cont + ";database=" + data + ";";  // PARA MYSQL debian/ubuntu
        string DB_CONN_STR = "Server=" + serv + ";Database=" + data + ";Uid=" + usua + ";Pwd=" + cont + ";";    // PARA MySQL CentOs
        Timer timer1 = new Timer();
        string ultid = "";  // ultimo id enviado correo
        public Bot_Mensajero()
        {
            InitializeComponent();
            eventoSistema = new EventLog();
            if (!EventLog.SourceExists("Bot_Mensajero"))
            {
                EventLog.CreateEventSource("Bot_Mensajero", "Application");
            }
            eventoSistema.Source = "Bot_Mensajero";
            eventoSistema.Log = "Application";
        }

        protected override void OnStart(string[] args)
        {
            // escribe en el log de eventos del sistema - inicio del mensajero
            eventoSistema.WriteEntry("Inicio del Bot_Mensajero");
            timer1.Enabled = true;
            timer1.Interval = int.Parse(lapso) * 1000;          // en xml es segundos, en c# es milisegundos, por eso multiplicamos por 1000
            timer1.Elapsed += timer1_Tick;
            timer1.AutoReset = true;
            timer1.Start();
        }

        protected override void OnStop()
        {
            // escribe en el log de eventos del sistema - detención del mensajero
            eventoSistema.WriteEntry("Detención del Bot_Mensajero");
            timer1.Stop();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            using (MySqlConnection conn = new MySqlConnection(DB_CONN_STR))
            {
                try
                {
                    conn.Open();
                    if (conn.State == System.Data.ConnectionState.Open)
                    {
                        DataTable dt = new DataTable();
                        // lee registros nuevos
                        //if (plan_lector(conn, dt) == true)
                        {
                            string jala = "SELECT a.emp_code,DATE_FORMAT(a.punch_time,'%H:%i:%s %d/%m/%Y'),a.area_alias,b.first_name,ifnull(b.last_name,'') AS last_name,a.id " +
                                "FROM iclock_transaction a LEFT JOIN personnel_employee b ON a.emp_code = b.emp_code " +
                                "WHERE date(a.punch_time)>@feini AND a.marca = 0";
                            using (MySqlCommand micon = new MySqlCommand(jala, conn))
                            {
                                micon.Parameters.AddWithValue("@feini", feini);
                                using (MySqlDataAdapter da = new MySqlDataAdapter(micon))
                                {
                                    dt.Rows.Clear();
                                    da.Fill(dt);
                                    if (dt.Rows.Count > 0) //retorna = true;
                                    {
                                        foreach (DataRow row in dt.Rows)
                                        {
                                            asunto = "";
                                            string cuerpo = getHtml(row.ItemArray[3].ToString() + " " +
                                                row.ItemArray[4].ToString(),        // nombre del fulano
                                                row.ItemArray[1].ToString(),        // fecha/hora
                                                row.ItemArray[0].ToString(),        // codigo empleado
                                                row.ItemArray[2].ToString());       // area o tienda
                                            asunto = asuco + row.ItemArray[3].ToString() + " " + row.ItemArray[4].ToString() + " - " + row.ItemArray[1].ToString();
                                            if (ultid != row.ItemArray[5].ToString())
                                            {
                                                try
                                                {
                                                    MailMessage message = new MailMessage();
                                                    SmtpClient smtp = new SmtpClient();
                                                    //message.From = new MailAddress(coror);      // correo quien envía
                                                    message.From = new MailAddress(coror, "Bot Mensajero");
                                                    message.To.Add(new MailAddress(corde));     // correo destinatario
                                                    message.Subject = asunto;                   // texto general + nombre y fecha/hora
                                                    message.IsBodyHtml = true;                  // correo en html?
                                                    message.Body = cuerpo;   // htmlString;                  // cuerpo del correo
                                                    smtp.Port = int.Parse(nupto);               // 26;  // 465;    // 587;
                                                    smtp.Host = smtpn;                          // "smtp.gmail.com";
                                                    if (ssl_sn == "NO")
                                                    {
                                                        smtp.EnableSsl = false;                     // correo con certificado de seguridad NO
                                                        smtp.UseDefaultCredentials = false;
                                                    }
                                                    else
                                                    {
                                                        smtp.EnableSsl = true;                     // correo con certificado de seguridad SI
                                                        smtp.UseDefaultCredentials = false;
                                                    }
                                                    smtp.Credentials = new NetworkCredential(coror, pasco);     // correoElectronico, contraseña
                                                    smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                                                    smtp.Send(message);
                                                    ultid = row.ItemArray[5].ToString();
                                                    mensajeLog = "Enviado correo del registro " + row.ItemArray[5].ToString() + " - " + row.ItemArray[3].ToString() + " " +
                                                        row.ItemArray[4].ToString() + " - " + row.ItemArray[1].ToString();
                                                    escribirLineaFichero();
                                                    //
                                                    string actua = "UPDATE iclock_transaction SET marca=1 WHERE id=@nreg";
                                                    using (MySqlCommand miupd = new MySqlCommand(actua, conn))
                                                    {
                                                        miupd.Parameters.AddWithValue("@nreg", row.ItemArray[5].ToString());
                                                        miupd.ExecuteNonQuery();
                                                    }
                                                }
                                                catch (Exception ex)
                                                {
                                                    mensajeLog = ex.Message.ToString();
                                                    escribirLineaFichero();
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        mensajeLog = "No hay datos nuevos para envío";
                                        escribirLineaFichero();
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        // mensaje de no hay conexión a la b.d.
                        mensajeLog = "No se puede conectar con la base de datos";
                        escribirLineaFichero();
                    }
                }
                catch (MySqlException ex)
                {
                    mensajeLog = ex.ToString();
                    escribirLineaFichero();
                }
            }
        }
        public static string getHtml(string nomb, string feho, string corr, string otro)
        {
            try
            {
                string messageBody = "<font>Horario de marcación entrada/salida: </font><br><br>";
                if (nomb.Length < 1) return messageBody;
                string htmlTableStart = "<table style=\"border-collapse:collapse; text-align:center;\" >";
                string htmlTableEnd = "</table>";
                string htmlHeaderRowStart = "<tr style=\"background-color:#6FA1D2; color:#ffffff;\">";
                string htmlHeaderRowEnd = "</tr>";
                string htmlTrStart = "<tr style=\"color:#555555;\">";
                string htmlTrEnd = "</tr>";
                string htmlTdStart = "<td style=\" border-color:#5c87b2; border-style:solid; border-width:thin; padding: 5px;\">";
                string htmlTdEnd = "</td>";
                messageBody += htmlTableStart;
                messageBody += htmlHeaderRowStart;
                messageBody += htmlTdStart + "Nombre trabajador" + htmlTdEnd;
                messageBody += htmlTdStart + "Fecha/Hora" + htmlTdEnd;
                messageBody += htmlTdStart + "Código" + htmlTdEnd;
                messageBody += htmlTdStart + "Area/Tda" + htmlTdEnd;
                messageBody += htmlHeaderRowEnd;
                //Loop all the rows from grid vew and added to html td  
                {
                    messageBody = messageBody + htmlTrStart;
                    messageBody = messageBody + htmlTdStart + nomb + htmlTdEnd; //adding nombre
                    messageBody = messageBody + htmlTdStart + feho + htmlTdEnd; //adding fecha hora
                    messageBody = messageBody + htmlTdStart + corr + htmlTdEnd; //adding correo
                    messageBody = messageBody + htmlTdStart + otro + htmlTdEnd; //adding celular  
                    messageBody = messageBody + htmlTrEnd;
                }
                messageBody = messageBody + htmlTableEnd;
                return messageBody; // return HTML Table as string from this function  
            }
            catch (Exception ex)
            {
                return null;
            }
        }
        // envío del mensaje
        public void escribirLineaFichero()
        {
            try
            {
                FileStream fs = new FileStream(@AppDomain.CurrentDomain.BaseDirectory +
                    "estado.log", FileMode.OpenOrCreate, FileAccess.Write);
                StreamWriter m_streamWriter = new StreamWriter(fs);
                m_streamWriter.BaseStream.Seek(0, SeekOrigin.End);
                //Quitar posibles saltos de línea del mensaje
                mensajeLog = mensajeLog.Replace(Environment.NewLine, " | ");
                mensajeLog = mensajeLog.Replace("\r\n", " | ").Replace("\n", " | ").Replace("\r", " | ");
                m_streamWriter.WriteLine(DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss") + " " + mensajeLog);
                m_streamWriter.Flush();
                m_streamWriter.Close();
            }
            catch
            {
                //Silenciosa
            }
        }


    }
}

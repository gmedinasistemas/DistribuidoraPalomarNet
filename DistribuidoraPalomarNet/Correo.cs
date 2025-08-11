using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;

namespace DistribuidoraPalomarNet
{
    public static class Correo
    {
        // Variables de configuración (equivalentes a las ini_* de VB)
        public static string HostEmail { get; set; } = "smtp.live.com";
        public static int PortEmail { get; set; } = 587;
        public static bool EnableSSL { get; set; } = true;

        public static void EnviarCorreo(string emisor, string password, string mensaje,
                                      string asunto, string destinatario, string ruta)
        {
            MailMessage correos = new MailMessage();
            SmtpClient envios = new SmtpClient();

            try
            {
                // Validate email addresses first
                try {
                    var fromAddress = new MailAddress(emisor);
                    
                    // Manejar múltiples destinatarios separados por comas
                    correos.To.Clear();
                    var destinatarios = destinatario.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                    foreach (var email in destinatarios)
                    {
                        var emailTrimmed = email.Trim();
                        // Validar cada dirección de correo
                        var toAddress = new MailAddress(emailTrimmed);
                        correos.To.Add(emailTrimmed);
                    }
                } catch (FormatException fex) {
                    throw new Exception($"Formato de correo inválido: {fex.Message}");
                }

                correos.Body = mensaje;
                correos.Subject = asunto;
                correos.IsBodyHtml = true;

                if (!string.IsNullOrEmpty(ruta))
                {
                    Attachment archivo = new Attachment(ruta);
                    correos.Attachments.Add(archivo);
                }

                correos.From = new MailAddress(emisor);

                // Configuración del cliente SMTP con timeout extendido
                envios.Host = Module.ini_host_email;
                envios.Port = Convert.ToInt32(Module.ini_port_email);
                envios.EnableSsl = Module.ini_enable_ssl;
                envios.UseDefaultCredentials = false;
                envios.DeliveryMethod = SmtpDeliveryMethod.Network;
                envios.Timeout = 20000; // 20 segundos de timeout

                // Configuración de credenciales
                var credenciales = new NetworkCredential(emisor, password);
                envios.Credentials = credenciales;

                // Logging de configuración
                System.Diagnostics.Debug.WriteLine($"Configuración SMTP:\nHost: {envios.Host}\nPuerto: {envios.Port}\nSSL: {envios.EnableSsl}\nUsuario: {emisor}");

                try {
                    // Prueba de conexión SMTP
                    using (var testClient = new SmtpClient(envios.Host, envios.Port))
                    {
                        testClient.Credentials = credenciales;
                        testClient.EnableSsl = envios.EnableSsl;
                        testClient.Timeout = 10000;
                        System.Diagnostics.Debug.WriteLine("Probando conexión SMTP...");
                        testClient.ServicePoint.MaxIdleTime = 2000;
                        testClient.ServicePoint.ConnectionLimit = 1;
                    }
                } catch (Exception ex) {
                    throw new Exception($"Error al probar conexión SMTP: {ex.Message}");
                }

                // Configuración de seguridad para certificados SSL
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                ServicePointManager.ServerCertificateValidationCallback =
                    (object s, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors) => true;

                try {
                    envios.Send(correos);
                    System.Windows.Forms.MessageBox.Show("El mensaje fue enviado correctamente.",
                                                       "Mensaje",
                                                       System.Windows.Forms.MessageBoxButtons.OK,
                                                       System.Windows.Forms.MessageBoxIcon.Information);
                }
                catch (SmtpFailedRecipientsException sfr)
                {
                    var failedAddresses = string.Join(", ", sfr.InnerExceptions.Select(x => x.FailedRecipient));
                    throw new Exception($"No se pudo enviar a los siguientes destinatarios: {failedAddresses}. Error: {sfr.Message}");
                }
                catch (SmtpException se)
                {
                    throw new Exception($"Error SMTP: {se.Message}. StatusCode: {se.StatusCode}, Host: {envios.Host}, Puerto: {envios.Port}");
                }
            }
            catch (Exception ex)
            {
                if (correos != null) correos.Dispose();
                throw new Exception($"Error en 'EnviarCorreo': {ex.Message}");
            }
        }
    }
}

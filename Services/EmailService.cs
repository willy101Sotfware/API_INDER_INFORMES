using System;
using System.IO;
using System.Threading.Tasks;
using MailKit.Net.Smtp;
using MailKit.Security;
using MimeKit;

namespace API_INDER_INFORMES.Services
{
    public class EmailService
    {
        private readonly IConfiguration _configuration;
        private readonly string _smtpServer;
        private readonly int _smtpPort;
        private readonly string _smtpUsername;
        private readonly string _smtpPassword;
        private readonly string _fromEmail;
        private readonly string _fromName;
        private readonly string _defaultRecipient;
        private readonly ILogger<EmailService> _logger;

        public EmailService(IConfiguration configuration, ILogger<EmailService> logger = null)
        {
            _configuration = configuration;
            _logger = logger;
            _smtpServer = _configuration["EmailSettings:SmtpServer"] ?? "mail.1cero1.com";
            _smtpPort = int.Parse(_configuration["EmailSettings:SmtpPort"] ?? "465");
            _smtpUsername = _configuration["EmailSettings:Username"] ?? "";
            _smtpPassword = _configuration["EmailSettings:Password"] ?? "";
            _fromEmail = _configuration["EmailSettings:FromEmail"] ?? "";
            _fromName = _configuration["EmailSettings:FromName"] ?? "API INDER Informes";
            _defaultRecipient = _configuration["EmailSettings:DefaultRecipient"] ?? "wilyd2@hotmail.com";
        }

        public async Task SendEmailAsync(string toEmail, string subject, string body, string attachmentPath = null)
        {
            try
            {
             
                if (string.IsNullOrEmpty(toEmail))
                {
                    toEmail = _defaultRecipient;
                }

                LogInfo($"Iniciando envío de correo a {toEmail} con asunto: {subject}");
                LogInfo($"Usando configuración: Servidor={_smtpServer}, Puerto={_smtpPort}, Usuario={_smtpUsername}");

         
                var email = new MimeMessage();
                email.From.Add(new MailboxAddress(_fromName, _fromEmail));
                email.To.Add(new MailboxAddress("Destinatario", toEmail));
                email.Subject = subject;

        
                var builder = new BodyBuilder
                {
                    HtmlBody = body
                };

          
                if (!string.IsNullOrEmpty(attachmentPath) && File.Exists(attachmentPath))
                {
                    LogInfo($"Agregando adjunto: {attachmentPath}");
                    builder.Attachments.Add(attachmentPath);
                }

                email.Body = builder.ToMessageBody();

                LogInfo("Mensaje creado, intentando conectar al servidor SMTP...");

         
                using (var smtp = new SmtpClient())
                {
                   
                    smtp.Timeout = 15000; 

                  
                    smtp.ServerCertificateValidationCallback = (s, c, h, e) => true;

                  
                    LogInfo($"Conectando a {_smtpServer}:{_smtpPort}...");

                    try
                    {
                        
                        if (_smtpPort == 465)
                        {
                            LogInfo("Usando SSL para la conexión...");
                            await smtp.ConnectAsync(_smtpServer, _smtpPort, SecureSocketOptions.SslOnConnect);
                        }
                        else
                        {
                        
                            LogInfo("Usando StartTls para la conexión...");
                            await smtp.ConnectAsync(_smtpServer, _smtpPort, SecureSocketOptions.StartTls);
                        }
                    }
                    catch (Exception ex)
                    {
                        LogInfo($"Error al conectar con el método principal: {ex.Message}. Intentando con Auto...");
                        try
                        {
                           
                            await smtp.ConnectAsync(_smtpServer, _smtpPort, SecureSocketOptions.Auto);
                        }
                        catch (Exception ex2)
                        {
                            LogInfo($"Error al conectar con Auto: {ex2.Message}. Intentando sin cifrado...");
                           
                            await smtp.ConnectAsync(_smtpServer, _smtpPort, SecureSocketOptions.None);
                        }
                    }

                 
                    LogInfo($"Autenticando con usuario {_smtpUsername}...");
                    await smtp.AuthenticateAsync(_smtpUsername, _smtpPassword);

                  
                    LogInfo("Enviando mensaje...");
                    await smtp.SendAsync(email);

               
                    await smtp.DisconnectAsync(true);
                    LogInfo("Mensaje enviado correctamente");
                }
            }
            catch (Exception ex)
            {
                LogError($"Error al enviar correo a {toEmail}: {ex.Message}");
                if (ex.InnerException != null)
                {
                    LogError($"Detalle del error: {ex.InnerException.Message}");
                }
                throw new Exception($"Error al enviar el correo: {ex.Message}", ex);
            }
        }

        private void LogInfo(string message)
        {
            _logger?.LogInformation(message);
           
            RegistrarLog(message);
        }

        private void LogError(string message)
        {
            _logger?.LogError(message);
         
            RegistrarLog($"ERROR: {message}");
        }

        private void RegistrarLog(string mensaje)
        {
            try
            {
                string directorio = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Logs");
                if (!Directory.Exists(directorio))
                {
                    Directory.CreateDirectory(directorio);
                }

                string fechaArchivo = DateTime.Now.ToString("yyyyMMdd");
                string archivoLog = Path.Combine(directorio, $"email_log_{fechaArchivo}.txt");

              
                string fechaHora = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
                string mensajeFormateado = $"[{fechaHora}] {mensaje}";

              
                File.AppendAllText(archivoLog, mensajeFormateado + Environment.NewLine);
            }
            catch
            {
              
            }
        }
    }
}
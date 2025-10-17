using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using RPAExtraccionNotasRR.src.Rpa.RR.Core.Entities;
using RPAExtraccionNotasRR.src.Rpa.RR.Core.Services;
using RPAExtraccionNotasRR.src.Rpa.RR.Infrastructure.Email;
using RPAExtraccionNotasRR.src.Rpa.RR.Infrastructure.Repositories;
using System;
using System.Threading.Tasks;

namespace Rpa.RR.ConsoleApp
{
    public class Program
    {
        public static async Task Main(string[] args)
        {
            try
            {
                using var host = Host.CreateDefaultBuilder(args)
                    .ConfigureAppConfiguration(cfg =>
                    {
                        cfg.AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);
                    })
                    .ConfigureServices((ctx, services) =>
                    {
                        // Resolve IEmailSender 
                        services.AddSingleton<IEmailSender, EmailHelper>();
                        services.Configure<SmtpOptions>(ctx.Configuration.GetSection("Smtp"));

                        services.AddSingleton(sp =>
                        new CredencialesRepository(
                            ctx.Configuration.GetConnectionString("ConnectionStringClaro"),
                            sp.GetRequiredService<IEmailSender>())); // Pass IEmailSender instance

                        services.AddSingleton(sp =>
                        new CasoRepository(ctx.Configuration.GetConnectionString("DefaultConnection")));

                        // Core Orquestador
                        services.AddSingleton<INotasRRService, NotasRRService>();
                    })
                    .Build();

                // Ejecuta el flujo principal
                var svc = host.Services.GetRequiredService<INotasRRService>();
                await svc.Run();

                Console.WriteLine("Proceso completado.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[Main] Error: {ex.Message}");
            }
        }
    }
}

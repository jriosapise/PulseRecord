using PulseRecord;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using System;
using Topshelf;

//IHost host = Host.CreateDefaultBuilder(args)
//    .ConfigureServices(services =>
//    {
//        services.AddHostedService<Worker>();
//    })
//    .Build();

//host.Run();


namespace PulseRecord
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // Configuramos el servicio de dependencias manualmente
            var serviceCollection = new ServiceCollection();
            ConfigureServices(serviceCollection);

            // Construimos el proveedor de servicios (serviceProvider)
            var serviceProvider = serviceCollection.BuildServiceProvider();

            // Configuramos el servicio de Topshelf
            HostFactory.Run(x =>
            {

                x.Service<Worker>(s =>
                {
                    // Inyectamos las dependencias usando el proveedor de servicios (serviceProvider)
                    s.ConstructUsing(() => serviceProvider.GetService<Worker>()!);

                    // Ejecutamos StartAsync de forma síncrona
                    s.WhenStarted(service => Task.Run(() => service.StartAsync(CancellationToken.None)).Wait());

                    // Ejecutamos StopAsync de forma síncrona
                    s.WhenStopped(service => Task.Run(() => service.StopAsync(CancellationToken.None)).Wait());
                });

                x.RunAsLocalSystem();
                x.SetServiceName("PulseRecord");
                x.SetDisplayName("Moving Files");
                x.SetDescription("Servicio para mover archivos de prueba periodicamente.");
                x.StartAutomatically();
            });
        }

        private static void ConfigureServices(IServiceCollection services)
        {
            // Configuramos el archivo appsettings.json
            var configuration = new ConfigurationBuilder()
                .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true)
                .Build();

            // Registramos las dependencias
            services.AddSingleton<IConfiguration>(configuration);
            services.AddLogging(configure => configure.AddConsole());
            services.AddSingleton<Worker>();  // Registramos el Worker
        }
    }
}
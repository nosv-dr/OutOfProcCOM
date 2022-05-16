using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;

namespace OutOfProcCOM
{
    [ComVisible(true)]
    [Guid(Contract.Constants.ServerClass)]
    [ComDefaultInterface(typeof(IServer))]
    public sealed class DllServer : IServer
    {
        private readonly Lazy<Proxy> lazyProxy = new Lazy<Proxy>(GetProxy);

        private static Proxy GetProxy()
        {
            AppDomain.CurrentDomain.AssemblyResolve += (object sender, ResolveEventArgs args) => Assembly.Load(args.Name);

            var assembly = Assembly.GetExecutingAssembly();

            var path = Path.GetDirectoryName(assembly.Location);

            AppDomain proxyDomain = AppDomain.CreateDomain("proxyDomain_" + Guid.NewGuid().ToString("N"),
                new System.Security.Policy.Evidence(AppDomain.CurrentDomain.Evidence),
                new AppDomainSetup
                {
                    ApplicationBase = path,
                    ConfigurationFile = Path.Combine(path, "DLLHost.exe.config"),
                    LoaderOptimization = LoaderOptimization.MultiDomainHost,
                    PrivateBinPath = path,
                    PrivateBinPathProbe = path,
                });

            var proxy = proxyDomain.CreateInstanceAndUnwrap(typeof(Proxy).Assembly.FullName, typeof(Proxy).FullName);

            return (Proxy)proxy;
        }

        double IServer.ComputePi()
        {
            Trace.WriteLine($"Running {nameof(DllServer)}.{nameof(IServer.ComputePi)}");
            
            return lazyProxy.Value.ComputePi();
        }

#if EMBEDDED_TYPE_LIBRARY
        private static readonly string tlbPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), $"{nameof(DllServer)}.comhost.dll");
#else
        private static readonly string tlbPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), Contract.Constants.TypeLibraryName);
#endif

        [ComRegisterFunction]
        internal static void RegisterFunction(Type t)
        {
            if (t != typeof(DllServer))
                return;

            // Register DLL surrogate and type library
            COMRegistration.DllSurrogate.Register(Contract.Constants.ServerClassGuid, Contract.Constants.ServerAppName, tlbPath: tlbPath);
        }

        [ComUnregisterFunction]
        internal static void UnregisterFunction(Type t)
        {
            if (t != typeof(DllServer))
                return;

            // Unregister DLL surrogate and type library
            COMRegistration.DllSurrogate.Unregister(Contract.Constants.ServerClassGuid, tlbPath);
        }
    }
}

using Microsoft.Management.Infrastructure;
using Microsoft.Management.Infrastructure.Options;
using System;
using System.Linq;
using System.Management;

namespace WMIAttempt
{
    internal class Program
    {
        static void CimVariantAdd(string hostname)
        {
            string currentNamespace = @"root\Microsoft\Windows\Defender";

            DComSessionOptions dcomOptions = new DComSessionOptions()
            {
                Impersonation = ImpersonationType.Impersonate,
                Timeout = TimeSpan.FromSeconds(30)
            };


            CimSession mySession = CimSession.Create(hostname, dcomOptions);

            CimInstance cimDefenderPreferences = mySession.GetInstance(currentNamespace, new CimInstance("MSFT_MpPreference"));

            CimMethodParametersCollection exclusions = new CimMethodParametersCollection
            {
                CimMethodParameter.Create("ExclusionPath", new string[] { $@"C:\Users\{System.Security.Principal.WindowsIdentity.GetCurrent().Name.Split( '\\' )[1]}\" }, Microsoft.Management.Infrastructure.CimType.StringArray, CimFlags.In)
            };

            mySession.InvokeMethod(cimDefenderPreferences, "Add", exclusions);
        }

        static void CimVariantAdd(string hostname, string[] directories)
        {
            string currentNamespace = @"root\Microsoft\Windows\Defender";

            DComSessionOptions dcomOptions = new DComSessionOptions()
            {
                Impersonation = ImpersonationType.Impersonate,
                Timeout = TimeSpan.FromSeconds(30)
            };

            CimSession mySession = CimSession.Create(hostname, dcomOptions);

            CimInstance cimDefenderPreferences = mySession.GetInstance(currentNamespace, new CimInstance("MSFT_MpPreference"));

            CimMethodParametersCollection exclusions = new CimMethodParametersCollection
            {
                CimMethodParameter.Create("ExclusionPath", directories, Microsoft.Management.Infrastructure.CimType.StringArray, CimFlags.In)
            };

            mySession.InvokeMethod(cimDefenderPreferences, "Add", exclusions);
        }

        //
        static void CimVariantRemove(string hostname)
        {
            string currentNamespace = @"root\Microsoft\Windows\Defender";

            DComSessionOptions dcomOptions = new DComSessionOptions()
            {
                Impersonation = ImpersonationType.Impersonate,
                Timeout = TimeSpan.FromSeconds(30)
            };


            CimSession mySession = CimSession.Create(hostname, dcomOptions);

            CimInstance cimDefenderPreferences = mySession.GetInstance(currentNamespace, new CimInstance("MSFT_MpPreference"));

            CimMethodParametersCollection exclusions = new CimMethodParametersCollection
            {
                CimMethodParameter.Create("ExclusionPath", new string[] { $@"C:\Users\{System.Security.Principal.WindowsIdentity.GetCurrent().Name.Split( '\\' )[1]}\" }, Microsoft.Management.Infrastructure.CimType.StringArray, CimFlags.In)
            };

            mySession.InvokeMethod(cimDefenderPreferences, "Remove", exclusions);
        }

        static void CimVariantRemove(string hostname, string[] directories)
        {
            string currentNamespace = @"root\Microsoft\Windows\Defender";

            DComSessionOptions dcomOptions = new DComSessionOptions()
            {
                Impersonation = ImpersonationType.Impersonate,
                Timeout = TimeSpan.FromSeconds(30)
            };

            CimSession mySession = CimSession.Create(hostname, dcomOptions);

            CimInstance cimDefenderPreferences = mySession.GetInstance(currentNamespace, new CimInstance("MSFT_MpPreference"));

            CimMethodParametersCollection exclusions = new CimMethodParametersCollection
            {
                CimMethodParameter.Create("ExclusionPath", directories, Microsoft.Management.Infrastructure.CimType.StringArray, CimFlags.In)
            };

            mySession.InvokeMethod(cimDefenderPreferences, "Remove", exclusions);
        }

        //

        static void WMIVariantAdd(string hostname)
        {
            ManagementScope managementScope = new ManagementScope($@"\\{hostname}\root\Microsoft\Windows\Defender", options: new ConnectionOptions() { });
            managementScope.Connect();

            ManagementPath mgmtPath = new ManagementPath("MSFT_MpPreference");

            ManagementClass obj = new ManagementClass(managementScope, mgmtPath, (ObjectGetOptions)null);
            obj.Get();

            obj.Scope.Options.EnablePrivileges = true;

            ManagementBaseObject information = obj.GetMethodParameters("Add");

            information.SetPropertyValue("ExclusionPath", new string[] { $@"C:\Users\{System.Security.Principal.WindowsIdentity.GetCurrent().Name.Split('\\')[1]}\" });
            information.SetPropertyValue("Force", true);

            obj.InvokeMethod("Add", information, null);

            return;
        }

        static void WMIVariantAdd(string hostname, string[] directories)
        {
            ManagementScope managementScope = new ManagementScope($@"\\{hostname}\root\Microsoft\Windows\Defender", options: new ConnectionOptions() { });
            managementScope.Connect();

            ManagementPath mgmtPath = new ManagementPath("MSFT_MpPreference");

            ManagementClass obj = new ManagementClass(managementScope, mgmtPath, (ObjectGetOptions)null);
            obj.Get();

            obj.Scope.Options.EnablePrivileges = true;

            ManagementBaseObject information = obj.GetMethodParameters("Add");

            information.SetPropertyValue("ExclusionPath", directories);
            information.SetPropertyValue("Force", true);

            obj.InvokeMethod("Add", information, null);

            return;
        }

        static void WMIVariantRemove(string hostname)
        {
            ManagementScope managementScope = new ManagementScope($@"\\{hostname}\root\Microsoft\Windows\Defender", options: new ConnectionOptions() { });
            managementScope.Connect();

            ManagementPath mgmtPath = new ManagementPath("MSFT_MpPreference");

            ManagementClass obj = new ManagementClass(managementScope, mgmtPath, (ObjectGetOptions)null);
            obj.Get();

            obj.Scope.Options.EnablePrivileges = true;

            ManagementBaseObject information = obj.GetMethodParameters("Remove");

            information.SetPropertyValue("ExclusionPath", new string[] { $@"C:\Users\{System.Security.Principal.WindowsIdentity.GetCurrent().Name.Split('\\')[1]}\" });
            information.SetPropertyValue("Force", true);

            obj.InvokeMethod("Remove", information, null);

            return;
        }

        static void WMIVariantRemove(string hostname, string[] directories)
        {
            ManagementScope managementScope = new ManagementScope($@"\\{hostname}\root\Microsoft\Windows\Defender", options: new ConnectionOptions() { });
            managementScope.Connect();

            ManagementPath mgmtPath = new ManagementPath("MSFT_MpPreference");

            ManagementClass obj = new ManagementClass(managementScope, mgmtPath, (ObjectGetOptions)null);
            obj.Get();

            obj.Scope.Options.EnablePrivileges = true;

            ManagementBaseObject information = obj.GetMethodParameters("Remove");

            information.SetPropertyValue("ExclusionPath", directories);
            information.SetPropertyValue("Force", true);

            obj.InvokeMethod("Remove", information, null);

            return;
        }

        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine("Send a computer name as your first argument.");

                return;
            }

            switch (args[1])
            {
                case "cim_a":
                    switch (args.Length)
                    {
                        case 2:
                            CimVariantAdd(args[0]);
                            break;

                        default:
                            CimVariantAdd(args[0], args.Skip(2).ToArray<string>());
                            break;
                    }
                    
                    break;
                
                case "wmi_a":
                    switch (args.Length)
                    {
                        case 2:
                            WMIVariantAdd(args[0]);
                            break;

                        default:
                            WMIVariantAdd(args[0], args.Skip(2).ToArray<string>());
                            break;
                    }
                    break;
                
                case "cim_r":
                    switch (args.Length)
                    {
                        case 2:
                            CimVariantRemove(args[0]);
                            break;

                        default:
                            CimVariantRemove(args[0], args.Skip(2).ToArray<string>());
                            break;
                    }
                    break;
                
                case "wmi_r":
                    switch (args.Length)
                    {
                        case 2:
                            WMIVariantRemove(args[0]);
                            break;

                        default:
                            WMIVariantRemove(args[0], args.Skip(2).ToArray<string>());
                            break;
                    }
                    break;
                
                default:
                    Console.WriteLine("Only wmi_a or cim_a with optional arguments.");
                    break;
            }

            return;
        }
    }
}
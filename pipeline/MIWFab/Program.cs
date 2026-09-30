// MIWFab.exe
// A thin bridge between the Tekla PowerFab Open API (.NET only) and everything else at MIW.
//
// It does one job: take an XML request document, execute it against PowerFab, write the
// XML response document. It knows nothing about tons, jobs, saws or plants, so it should
// never need changing again. All the business logic lives in Python on top of it.
//
// The API surface, confirmed by reflection against Tekla.PowerFab.API 1.0.1 on 2026-08-28.
// The whole assembly is two types and three methods:
//
//     string ExecuteDirect(string request)
//     string ExecuteRemote(string ipAddress, long port, string request)
//     void   SetPrgmFilesDirectory(string prgmFilesDirectory)
//
// Both are EXPLICIT interface implementations, so they do not exist on the concrete
// TeklaPowerFabAPI class. You must hold the object as ITeklaPowerFabAPI to call them.
// Note the port is a long, not an int.
//
// Usage:
//   MIWFab.exe --dump-api
//       Prints the whole reflected API surface. Kept because it is how the above was found
//       and it is the first thing to re-run after a PowerFab or package upgrade.
//
//   MIWFab.exe --mode remote --server 192.168.1.15 --port 9154 \
//              --user "MIW Reports" --pass secret --in request.xml --out response.xml
//
//   MIWFab.exe --mode direct --server 192.168.1.15 --port 3306 \
//              --user forest --pass secret --in request.xml --out response.xml
//
//   Optional on either mode:
//     --prgm "C:\Program Files\Tekla\Tekla PowerFab"
//         Passed to SetPrgmFilesDirectory. Auto-detected if not supplied. Direct mode
//         appears to shell out to a helper process over a named pipe, so it may need this;
//         remote mode probably does not.
//
// The input request.xml holds ONLY the commands you want to run. This program adds the
// connect and the connection plumbing.
//
// Build:  dotnet build -c Release
// Requires NuGet package Tekla.PowerFab.API

using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Xml.Linq;
using TeklaPowerFab.TeklaPowerFabAPI;

namespace MIWFab
{
    internal static class Program
    {
        private static int Main(string[] rawArgs)
        {
            try
            {
                var args = ParseArgs(rawArgs);

                // Held as the INTERFACE on purpose. The methods are explicit
                // implementations and are invisible on the concrete class.
                ITeklaPowerFabAPI api = new TeklaPowerFabAPI();

                if (args.ContainsKey("dump-api"))
                {
                    DumpApi(api);
                    return 0;
                }

                string mode    = Get(args, "mode", "remote");
                string server  = Require(args, "server");
                long   port    = long.Parse(Get(args, "port", mode == "direct" ? "3306" : "9154"));
                string user    = Require(args, "user");
                string pass    = Require(args, "pass");
                string inPath  = Require(args, "in");
                string outPath = Require(args, "out");

                string prgm = Get(args, "prgm", FindPowerFabDirectory());
                if (!string.IsNullOrEmpty(prgm))
                {
                    try { api.SetPrgmFilesDirectory(prgm); }
                    catch (Exception ex)
                    {
                        Console.Error.WriteLine("Note: SetPrgmFilesDirectory failed, continuing. "
                                                + ex.Message);
                    }
                }

                string payload = File.ReadAllText(inPath);
                string responseXml;

                if (mode == "direct")
                {
                    // Direct mode: every command in one document runs in order, so the
                    // Connect is simply the first child. No session state to manage, and
                    // ExecuteDirect takes nothing but the request.
                    string doc = InjectConnect(payload, server, port, user, pass);
                    responseXml = api.ExecuteDirect(doc);
                }
                else
                {
                    // Remote mode: ConnectRemote returns a ConnectionGUID that must be set
                    // on the root of every subsequent request. Address and port are method
                    // arguments, not part of the XML.
                    string connectDoc =
                        "<FabSuiteXMLRequest><ConnectRemote>"
                        + "<Username>" + Esc(user) + "</Username>"
                        + "<Password>" + Esc(pass) + "</Password>"
                        + "</ConnectRemote></FabSuiteXMLRequest>";

                    string connectResp = api.ExecuteRemote(server, port, connectDoc);
                    Check(connectResp, "ConnectRemote");

                    string guid = ValueOf(connectResp, "ConnectionGUID");
                    if (string.IsNullOrEmpty(guid))
                        throw new Exception("ConnectRemote succeeded but returned no ConnectionGUID.");

                    string doc = SetConnectionGuid(payload, guid);
                    responseXml = api.ExecuteRemote(server, port, doc);

                    // Best effort close. A leaked connection is not fatal but it is untidy.
                    try
                    {
                        api.ExecuteRemote(server, port,
                            "<FabSuiteXMLRequest><ConnectionGUID>" + guid
                            + "</ConnectionGUID><CloseRemote/></FabSuiteXMLRequest>");
                    }
                    catch { }
                }

                File.WriteAllText(outPath, responseXml, new UTF8Encoding(false));

                // Report problems on stderr with a non zero exit code so Python can react.
                return Report(responseXml);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("MIWFab error: " + ex.Message);
                return 2;
            }
        }

        private static string FindPowerFabDirectory()
        {
            // MIW runs the 64 bit install, confirmed 2026-08-28. The others are fallbacks
            // for older or differently sited installs.
            string[] guesses =
            {
                @"C:\Program Files\Tekla\Tekla PowerFab",
                @"C:\Program Files (x86)\Tekla\Tekla PowerFab",
                @"C:\Program Files\FabSuite",
                @"C:\Program Files (x86)\FabSuite",
            };
            return guesses.FirstOrDefault(Directory.Exists) ?? "";
        }

        // ---------------------------------------------------------------- reflection dump

        private static void DumpApi(object api)
        {
            var t = api.GetType();
            Console.WriteLine("Type:     " + t.FullName);
            Console.WriteLine("Assembly: " + t.Assembly.FullName);
            Console.WriteLine();

            Console.WriteLine("=== Interfaces implemented ===");
            foreach (var i in t.GetInterfaces()) Console.WriteLine("  " + i.FullName);
            Console.WriteLine();

            DumpMethods("Public instance methods on " + t.Name,
                        t, BindingFlags.Public | BindingFlags.Instance);
            DumpMethods("Non-public instance methods on " + t.Name
                        + "  (explicit interface implementations land here)",
                        t, BindingFlags.NonPublic | BindingFlags.Instance);
            DumpMethods("Static methods on " + t.Name,
                        t, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Static);

            foreach (var i in t.GetInterfaces())
                DumpMethods("Methods declared on interface " + i.Name,
                            i, BindingFlags.Public | BindingFlags.Instance);

            Console.WriteLine("=== Public instance properties ===");
            var props = t.GetProperties(BindingFlags.Public | BindingFlags.Instance);
            if (props.Length == 0) Console.WriteLine("  (none)");
            foreach (var p in props) Console.WriteLine("  " + p.PropertyType.Name + " " + p.Name);
            Console.WriteLine();

            Console.WriteLine("=== Public types in the assembly ===");
            foreach (var ty in t.Assembly.GetExportedTypes().OrderBy(x => x.FullName))
                Console.WriteLine("  " + ty.FullName);
        }

        private static void DumpMethods(string header, Type owner, BindingFlags flags)
        {
            Console.WriteLine("=== " + header + " ===");
            int n = 0;
            foreach (var m in owner.GetMethods(flags).OrderBy(m => m.Name))
            {
                if (m.DeclaringType == typeof(object)) continue;
                if (m.IsSpecialName) continue;
                string ps = string.Join(", ", m.GetParameters()
                    .Select(p => p.ParameterType.Name + " " + p.Name));
                Console.WriteLine("  " + m.ReturnType.Name + " " + m.Name + "(" + ps + ")");
                n++;
            }
            if (n == 0) Console.WriteLine("  (none)");
            Console.WriteLine();
        }

        // ---------------------------------------------------------------- xml helpers

        private static string InjectConnect(string payload, string server, long port,
                                            string user, string pass)
        {
            var doc = XDocument.Parse(payload);
            var ns = doc.Root.Name.Namespace;
            var connect = new XElement(ns + "Connect",
                new XElement(ns + "IPAddress", server),
                new XElement(ns + "PortNumber", port.ToString()),
                new XElement(ns + "Username", user),
                new XElement(ns + "Password", pass));
            doc.Root.AddFirst(connect);
            return doc.ToString();
        }

        // ConnectionGUID is an ELEMENT on the request root, not an attribute. It is a
        // documented property of fsreqFabSuiteXMLRequest alongside Lang, Version and the
        // Trimble identity tokens. Sending it as an attribute produces an XMLWarning,
        // PowerFab drops it, and the request runs with no session. Verified live 2026-08-28.
        private static string SetConnectionGuid(string payload, string guid)
        {
            var doc = XDocument.Parse(payload);
            var ns = doc.Root.Name.Namespace;
            doc.Root.Attribute("ConnectionGUID")?.Remove();
            doc.Root.Elements(ns + "ConnectionGUID").Remove();
            doc.Root.AddFirst(new XElement(ns + "ConnectionGUID", guid));
            return doc.ToString();
        }

        private static string ValueOf(string xml, string localName)
        {
            var doc = XDocument.Parse(xml);
            var el = doc.Descendants().FirstOrDefault(e => e.Name.LocalName == localName);
            return el == null ? null : el.Value.Trim();
        }

        private static void Check(string xml, string command)
        {
            var doc = XDocument.Parse(xml);

            var err = doc.Descendants().FirstOrDefault(e => e.Name.LocalName == "XMLError");
            if (err != null)
                throw new Exception("XMLError (fatal, nothing ran): " + err.Value);

            var cmd = doc.Descendants().FirstOrDefault(e => e.Name.LocalName == command);
            if (cmd != null)
            {
                var succ = cmd.Elements().FirstOrDefault(e => e.Name.LocalName == "Successful");
                if (succ != null && succ.Value.Trim() != "1")
                {
                    var msg = cmd.Elements().FirstOrDefault(e => e.Name.LocalName == "ErrorMessage");
                    throw new Exception(command + " failed: "
                                        + (msg == null ? "(no message)" : msg.Value));
                }
            }
        }

        // XMLWarning is non fatal to PowerFab, which silently drops the bad element and runs
        // the command anyway. That is exactly how you end up with an unfiltered dataset that
        // looks plausible. At MIW it is treated as a failure.
        private static int Report(string xml)
        {
            var doc = XDocument.Parse(xml);
            int rc = 0;

            foreach (var e in doc.Descendants().Where(e => e.Name.LocalName == "XMLError"))
            {
                Console.Error.WriteLine("XMLError: " + e.Value);
                rc = 2;
            }
            foreach (var e in doc.Descendants().Where(e => e.Name.LocalName == "XMLWarning"))
            {
                Console.Error.WriteLine(
                    "XMLWarning (element silently dropped, result is suspect): " + e.Value);
                if (rc == 0) rc = 1;
            }
            foreach (var s in doc.Descendants().Where(e => e.Name.LocalName == "Successful"))
            {
                if (s.Value.Trim() != "1")
                {
                    var parent = s.Parent;
                    var msg = parent.Elements()
                                    .FirstOrDefault(e => e.Name.LocalName == "ErrorMessage");
                    Console.Error.WriteLine(parent.Name.LocalName + " failed: "
                                            + (msg == null ? "(no message)" : msg.Value));
                    rc = 2;
                }
            }
            return rc;
        }

        // ---------------------------------------------------------------- arg helpers

        private static System.Collections.Generic.Dictionary<string, string> ParseArgs(string[] a)
        {
            var d = new System.Collections.Generic.Dictionary<string, string>();
            for (int i = 0; i < a.Length; i++)
            {
                if (!a[i].StartsWith("--")) continue;
                string k = a[i].Substring(2);
                if (i + 1 < a.Length && !a[i + 1].StartsWith("--")) d[k] = a[++i];
                else d[k] = "true";
            }
            return d;
        }

        private static string Get(System.Collections.Generic.Dictionary<string, string> d,
                                  string k, string def)
            => d.ContainsKey(k) ? d[k] : def;

        private static string Require(System.Collections.Generic.Dictionary<string, string> d,
                                      string k)
        {
            if (!d.ContainsKey(k)) throw new Exception("Missing required argument --" + k);
            return d[k];
        }

        private static string Esc(string s)
            => s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;");
    }
}

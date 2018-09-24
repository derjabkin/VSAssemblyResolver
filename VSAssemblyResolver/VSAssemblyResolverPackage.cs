﻿using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Design;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Toolbox;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Design;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;
using System.Runtime.Versioning;
using System.Threading;
using System.Threading.Tasks;
using System.Threading.Tasks.Dataflow;
using VSLangProj;

namespace SergejDerjabkin.VSAssemblyResolver
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    ///
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the 
    /// IVsPackage interface and uses the registration attributes defined in the framework to 
    /// register itself and its components with the shell.
    /// </summary>
    // This attribute tells the PkgDef creation utility (CreatePkgDef.exe) that this class is
    // a package.
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    // This attribute is used to register the information needed to show this package
    // in the Help/About dialog of Visual Studio.
    [InstalledProductRegistration("#110", "#112", "0.5", IconResourceID = 400)]
    // This attribute is needed to let the shell know that this package exposes some menus.
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionHasSingleProject_string)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionHasMultipleProjects_string)]
    [Guid(GuidList.guidVSAssemblyResolverPkgString)]
    public sealed class VSAssemblyResolverPackage : AsyncPackage
    {
        private Guid outputPaneGuid = new Guid("5540917A-D013-4C31-8F80-94F3B783FE88");
        private const string outputPaneTitle = "VSAssemblyResolver";
        private IVsOutputWindowPane outputPane;
        private readonly HashSet<string> missingAssembliesCache = new HashSet<string>();
        private bool solutionOpened;
        private BufferBlock<string> outputBuffer;



        private IVsOutputWindowPane OutputPane
        {
            get
            {
                if (outputPane == null)
                {
                    if (GetGlobalService(typeof(SVsOutputWindow)) is IVsOutputWindow window)
                    {
                        window.CreatePane(ref outputPaneGuid, outputPaneTitle, 1, 1);
                        window.GetPane(ref outputPaneGuid, out outputPane);
                    }
                }
                return outputPane;
            }
        }

        private void WriteOutput(string format, params object[] args)
        {
            outputBuffer.Post(string.Format(CultureInfo.CurrentCulture, format, args));
        }

        private async System.Threading.Tasks.Task WriteOutputAsync(ISourceBlock<string> source)
        {
            while (await source.OutputAvailableAsync())
            {
                string data = await source.ReceiveAsync();
                if (OutputPane != null)
                {
                    OutputPane.OutputStringThreadSafe(data);
                    OutputPane.OutputStringThreadSafe("\r\n");
                }
            }
        }

        /////////////////////////////////////////////////////////////////////////////
        // Overridden Package Implementation
        #region Package Members

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        /// 
        protected override async System.Threading.Tasks.Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {

            Debug.WriteLine(string.Format(CultureInfo.CurrentCulture, "Entering Initialize() of: {0}", this.ToString()));
            await base.InitializeAsync(cancellationToken, progress);

            // Create output buffer ans start consumer.
            outputBuffer = new BufferBlock<string>();
            WriteOutputAsync(outputBuffer);

            // Add our command handlers for menu (commands must exist in the .vsct file)
            if (await GetServiceAsync(typeof(IMenuCommandService)) is OleMenuCommandService mcs)
            {
                // Create the command for the menu item.
                CommandID menuCommandID = new CommandID(GuidList.guidVSAssemblyResolverCmdSet, (int)PkgCmdIDList.cmdidPopulateToolbox);
                MenuCommand menuItem = new MenuCommand(MenuItemCallback, menuCommandID);
                mcs.AddCommand(menuItem);
            }

            var dte = await GetDteAsync();
            dte.Events.SolutionEvents.Opened += () =>
            {
                foreach (Project project in dte.Solution.Projects)
                    WireProjectEvents(project);

                InvalidateCache();
                solutionOpened = true;
            };
            dte.Events.SolutionEvents.ProjectAdded += WireProjectEvents;
            await RegisterSolutionServices();

            AppDomain.CurrentDomain.AssemblyResolve += CurrentDomain_AssemblyResolve;
            AppDomain.CurrentDomain.ReflectionOnlyAssemblyResolve += CurrentDomain_ReflectionOnlyAssemblyResolve;

        }

        private async System.Threading.Tasks.Task RegisterSolutionServices()
        {
            string packageDirectory = GetPackagesDirectory(await GetDteAsync());
            if (!Directory.Exists(packageDirectory)) return;

            var servicesDirectories = Directory.GetDirectories(packageDirectory, "*VSARService*");
            string[] files = servicesDirectories.SelectMany(sd => Directory.GetFiles(sd, "*.dll", SearchOption.AllDirectories)).ToArray();

            foreach (var fileName in files)
            {
                Assembly assembly = null;

                try
                {
                    assembly = Assembly.ReflectionOnlyLoadFrom(fileName);
                }
                catch (BadImageFormatException) { }
                catch (TypeLoadException) { }
                catch (ReflectionTypeLoadException) { }

                if (assembly != null && GetRegistratorType(assembly) != null)
                {
                    assembly = Assembly.LoadFile(fileName);
                    Type registratorType = GetRegistratorType(assembly);
                    var method = registratorType.GetMethod("RegisterServices", BindingFlags.Static | BindingFlags.Public);
                    if (method != null)
                    {
                        try
                        {
                            method.Invoke(null, new object[] { this });
                        }
                        catch (Exception ex)
                        {
                            WriteOutput("ERROR in {0}.RegisterServices: \r\n{1}", registratorType.FullName, ex);
                        }
                    }
                }
            }
        }

        private static Type GetRegistratorType(Assembly assembly)
        {
            try
            {
                return assembly.GetTypes().FirstOrDefault(a => a.Name == "ServiceRegistrator");
            }
            catch (ReflectionTypeLoadException)
            {
                return null;
            }
            catch (TypeLoadException)
            {
                return null;
            }
        }
        private void WireProjectEvents(Project project)
        {

            if (project.Object is VSProject vsProject)
            {
                vsProject.Events.ReferencesEvents.ReferenceAdded += r => InvalidateCache();
                vsProject.Events.ReferencesEvents.ReferenceChanged += r => InvalidateCache();
            }
        }


        private void InvalidateCache()
        {
            lock (missingAssembliesCache)
            {
                missingAssembliesCache.Clear();
            }
            System.Threading.Tasks.Task.Run(async () => { referenceDirectories = await GetReferenceDirectoriesCoreAsync(); });

        }

        private async System.Threading.Tasks.Task PopulateReferenceDirectoriesAsync()
        {
        }

        Assembly CurrentDomain_ReflectionOnlyAssemblyResolve(object sender, ResolveEventArgs args)
        {
            string path = GetAssemblyPath(args.Name);
            if (!string.IsNullOrWhiteSpace(path))
                return Assembly.ReflectionOnlyLoadFrom(path);

            AddMissingAssembly(args);
            return null;
        }


        private async Task<string[]> GetReferenceDirectoriesEnum()
        {
            if (!solutionOpened)
                return new string[0];

            DTE dte = await GetDteAsync();

            return dte.Solution.Projects
                .OfType<Project>()
                .Select(p => p.Object)
                .OfType<VSProject>()
                .SelectMany(vsp => vsp.References.OfType<Reference>())
                .Where(r => !string.IsNullOrWhiteSpace(r.Path))
                .Select(r => Path.GetDirectoryName(r.Path))
                .Distinct()
                .ToArray();
        }



        private DTE dteCache;
        private async Task<DTE> GetDteAsync()
        {
            return dteCache ?? (dteCache = (DTE)await GetServiceAsync(typeof(DTE)));
        }


        private async Task<string[]> GetReferenceDirectoriesCoreAsync()
        {
            DTE dte = await GetDteAsync();
            if (dte == null || dte.Solution == null || string.IsNullOrEmpty(dte.Solution.FileName))
                return new string[0];

            var dirs =
                new string[] { GetPackagesDirectory(dte) }
                    .Concat(await GetReferenceDirectoriesEnum())
                    .Select(Path.GetFullPath)
                    .OrderBy(s => s.Length)
                    .ToList();


            for (int i = 1; i < dirs.Count; i++)
            {
                if (dirs.Where(d => dirs[i].StartsWith(d, StringComparison.OrdinalIgnoreCase), 0, i - 1).Any())
                    dirs.RemoveAt(i--);
            }

            return dirs.ToArray();
        }


        private string[] referenceDirectories;
        private string[] GetReferenceDirectories()
        {
            if (referenceDirectories == null)
            {
                var task = System.Threading.Tasks.Task.Run(() => GetReferenceDirectoriesCoreAsync());

                //This is a very dirty workaround for a deadlock that occurs when accessing solution object
                if (task.Wait(10000))
                    referenceDirectories = task.Result;
                else
                    WriteOutput("GetReferenceDirectories: Timeout expired.");
            }
            return referenceDirectories ?? new string[0];
        }

        private static string GetPackagesDirectory(DTE dte)
        {
            if (dte == null)
                throw new ArgumentNullException(nameof(dte));

            if (dte.Solution == null)
                return null;

            if (string.IsNullOrWhiteSpace(dte.Solution.FileName))
                return null;

            return Path.Combine(Path.GetDirectoryName(dte.Solution.FileName), "packages");
        }


        private string GetAssemblyPath(string name)
        {
            var dirs = GetReferenceDirectories();
            AssemblyName asmName = new AssemblyName(name);

            foreach (var rootDir in dirs)
            {


                if (Directory.Exists(rootDir))
                {
                    WriteOutput("Looking for {0} in {1}", name, rootDir);
                    foreach (
                        var dir in
                            rootDir.ToEnumerable()
                                   .Concat(Directory.GetDirectories(rootDir, "*.*", SearchOption.AllDirectories)))
                    {
                        string path = Path.Combine(dir, asmName.Name + ".dll");
                        if (File.Exists(path))
                        {
                            var foundName = AssemblyName.GetAssemblyName(path);

                            if (IsNameCompatible(foundName, asmName))
                            {
                                return path;
                            }
                        }
                    }
                }
            }

            return null;
        }

        private static bool IsNameCompatible(AssemblyName fullName, AssemblyName partialName)
        {
            return partialName.Name == fullName.Name &&
                   ((partialName.Version == null) || (partialName.Version == fullName.Version)) &&
                   ((partialName.GetPublicKeyToken() == null ||
                    partialName.GetPublicKeyToken().SequenceEqual(fullName.GetPublicKeyToken())));
        }




        private bool IsMissingAssembly(string name)
        {
            if (missingAssembliesCache.Contains(name))
            {
                WriteOutput("[CACHE] Missing assembly {0}", name);
                return true;
            }
            else
                return false;
        }
        private static bool resolving;


        private DynamicTypeService typeResolver;
        private DynamicTypeService TypeResolver
        {
            get
            {
                if (typeResolver == null)
                {
                    typeResolver = (DynamicTypeService)GetService(typeof(DynamicTypeService));
                }

                return typeResolver;
            }
        }

        Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {

            if (IsMissingAssembly(args.Name))
                return null;

            WriteOutput("Resolving assembly {0}", args.Name);

            if (resolving) return null;

            resolving = true;
            try
            {

                AppDomain domain = (AppDomain)sender;
                AssemblyName asmName = new AssemblyName(args.Name);
                try
                {
                    return domain.Load(asmName);
                }
                catch (FileLoadException)
                {
                }

                catch (FileNotFoundException)
                {
                }

                string path = GetAssemblyPath(args.Name);
                if (!string.IsNullOrWhiteSpace(path))
                    return TypeResolver.CreateDynamicAssembly(path);


                AddMissingAssembly(args);
                return null;
            }
            finally
            {
                resolving = false;
            }
        }


        private Assembly LoadReflectionAssemblySafe(string fileName)
        {
            try
            {
                return Assembly.ReflectionOnlyLoadFrom(fileName);
            }
            catch (BadImageFormatException) { }
            catch (FileLoadException) { }
            catch (FileNotFoundException) { }

            return null;
        }
        private void AddMissingAssembly(ResolveEventArgs args)
        {
            lock (missingAssembliesCache)
            {
                if (!missingAssembliesCache.Contains(args.Name))
                    missingAssembliesCache.Add(args.Name);
            }
        }
        #endregion


        private bool IsToolboxItem(Type type)
        {

            var dxToolboxAttribute = type.GetCustomAttributesData().FirstOrDefault(d =>
                d.AttributeType.Name == "DXToolboxItemAttribute" || d.AttributeType.Name == "ToolboxItemAttribute");
            if (dxToolboxAttribute != null)
            {

                var b = dxToolboxAttribute.ConstructorArguments[0].Value as bool?;
                return b == null || b.Value;
            }

            return false;
        }
        /// <summary>
        /// This function is the callback used to execute a command when the a menu item is clicked.
        /// See the Initialize method to see how the menu item is associated to this function using
        /// the OleMenuCommandService service and the MenuCommand class.
        /// </summary>
        private void MenuItemCallback(object sender, EventArgs e)
        {
            DynamicTypeService typeResolverService = (DynamicTypeService)GetService(typeof(DynamicTypeService));


            if (GetService(typeof(IToolboxService)) is IVsToolboxService2 toolBoxService)
            {
                try
                {
                    var items = new List<Tuple<string, ToolboxItem, IEnumerable<FrameworkName>, Guid>>();
                    HashSet<string> processedAssemblies = new HashSet<string>();
                    foreach (var directory in GetReferenceDirectories())
                    {
                        foreach (string fileName in Directory.GetFiles(directory, "*.dll", SearchOption.AllDirectories))
                        {
                            var assembly = typeResolverService.CreateDynamicAssembly(fileName);
                            if (assembly != null && !processedAssemblies.Contains(assembly.FullName))
                            {
                                processedAssemblies.Add(assembly.FullName);
                                try
                                {
                                    foreach (Type type in assembly.GetTypes())
                                    {

                                        if (!type.IsAbstract && !type.IsGenericTypeDefinition && IsToolboxItem(type))
                                        {
                                            if (typeof(IComponent).IsAssignableFrom(type))
                                            {
                                                items.Add(new Tuple<string, ToolboxItem, IEnumerable<FrameworkName>, Guid>(
                                                    "VSAssemblyResolver", new ResolverToolboxItem(type), null, Guid.Empty));
                                            }
                                        }
                                    }
                                }
                                catch (ReflectionTypeLoadException) { }
                                catch (TypeLoadException) { }
                                catch (FileNotFoundException) { }
                            }
                        }
                    }
                    toolBoxService.AddToolboxItems(items, null);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine(ex.ToString());
                }
            }
            // Show a Message Box to prove we were here
            IVsUIShell uiShell = (IVsUIShell)GetService(typeof(SVsUIShell));
            Guid clsid = Guid.Empty;
            ErrorHandler.ThrowOnFailure(uiShell.ShowMessageBox(
                       0,
                       ref clsid,
                       "VSAssemblyResolver",
                       string.Format(CultureInfo.CurrentCulture, "Inside {0}.MenuItemCallback()", this),
                       string.Empty,
                       0,
                       OLEMSGBUTTON.OLEMSGBUTTON_OK,
                       OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST,
                       OLEMSGICON.OLEMSGICON_INFO,
                       0,        // false
                       out int result));
        }


    }

    [Serializable]
    public class ResolverToolboxItem : ToolboxItem
    {
        public ResolverToolboxItem(Type toolType)
            : base(toolType)
        {
            this.Bitmap = GetImage(toolType);
        }

        private static Bitmap GetImage(Type toolType)
        {
            var tb = (ToolboxBitmapAttribute)toolType.GetCustomAttributes(typeof(ToolboxBitmapAttribute), false).FirstOrDefault();
            if (tb != null)
            {
                return (Bitmap)tb.GetImage(toolType);

            }

            return null;
        }

        protected ResolverToolboxItem(SerializationInfo info, StreamingContext context)
        {
            Deserialize(info, context);
        }
    }
}

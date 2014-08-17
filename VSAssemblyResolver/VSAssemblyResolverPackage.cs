using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.ComponentModel.Design;
using Microsoft.VisualStudio.Shell.Design;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Shell;
using EnvDTE;
using VSLangProj;
using Microsoft.VisualStudio.Toolbox;
using System.Drawing.Design;
using System.Runtime.Serialization;
using System.ComponentModel;
using System.Runtime.Versioning;

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
    [PackageRegistration(UseManagedResourcesOnly = true)]
    // This attribute is used to register the information needed to show this package
    // in the Help/About dialog of Visual Studio.
    [InstalledProductRegistration("#110", "#112", "0.5", IconResourceID = 400)]
    // This attribute is needed to let the shell know that this package exposes some menus.
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionHasSingleProject_string)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionHasMultipleProjects_string)]
    [Guid(GuidList.guidVSAssemblyResolverPkgString)]
    public sealed class VSAssemblyResolverPackage : Package
    {
        private Guid outputPaneGuid = new Guid("5540917A-D013-4C31-8F80-94F3B783FE88");
        private const string outputPaneTitle = "VSAssemblyResolver";
        private IVsOutputWindowPane outputPane;
        private readonly HashSet<string> missingAssembliesCache = new HashSet<string>();
        /// <summary>
        /// Default constructor of the package.
        /// Inside this method you can place any initialization code that does not require 
        /// any Visual Studio service because at this point the package object is created but 
        /// not sited yet inside Visual Studio environment. The place to do all the other 
        /// initialization is the Initialize method.
        /// </summary>
        public VSAssemblyResolverPackage()
        {
            Debug.WriteLine(string.Format(CultureInfo.CurrentCulture, "Entering constructor for: {0}", this.ToString()));
        }




        private IVsOutputWindowPane OutputPane
        {
            get
            {
                if (outputPane == null)
                {
                    IVsOutputWindow window = GetGlobalService(typeof(SVsOutputWindow)) as IVsOutputWindow;
                    if (window != null)
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
            if (OutputPane != null)
            {
                OutputPane.OutputString(string.Format(CultureInfo.CurrentCulture, format, args));
                OutputPane.OutputString("\r\n");
            }
        }
        /////////////////////////////////////////////////////////////////////////////
        // Overridden Package Implementation
        #region Package Members

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override void Initialize()
        {
            AppDomain.CurrentDomain.AssemblyResolve += CurrentDomain_AssemblyResolve;
            AppDomain.CurrentDomain.ReflectionOnlyAssemblyResolve += CurrentDomain_ReflectionOnlyAssemblyResolve;

            Debug.WriteLine(string.Format(CultureInfo.CurrentCulture, "Entering Initialize() of: {0}", this.ToString()));
            base.Initialize();

            // Add our command handlers for menu (commands must exist in the .vsct file)
            OleMenuCommandService mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (null != mcs)
            {
                // Create the command for the menu item.
                CommandID menuCommandID = new CommandID(GuidList.guidVSAssemblyResolverCmdSet, (int)PkgCmdIDList.cmdidPopulateToolbox);
                MenuCommand menuItem = new MenuCommand(MenuItemCallback, menuCommandID);
                mcs.AddCommand(menuItem);
            }

            var dte = GetDte();
            dte.Events.SolutionEvents.Opened += () =>
            {
                foreach (Project project in dte.Solution.Projects)
                    WireProjectEvents(project);

                InvalidateCache();
            };
            dte.Events.SolutionEvents.ProjectAdded += WireProjectEvents;
        }

        private void WireProjectEvents(Project project)
        {
            VSProject vsProject = project.Object as VSProject;

            if (vsProject != null)
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
        }

        Assembly CurrentDomain_ReflectionOnlyAssemblyResolve(object sender, ResolveEventArgs args)
        {
            string path = GetAssemblyPath(args.Name);
            if (!string.IsNullOrWhiteSpace(path))
                return Assembly.ReflectionOnlyLoadFrom(path);

            AddMissingAssembly(args);
            return null;
        }


        private IEnumerable<string> GetReferenceDirectoriesEnum()
        {
            DTE dte = GetDte();
            foreach (Project project in dte.Solution.Projects)
            {
                var vsp = project.Object as VSProject;
                if (vsp != null)
                {
                    foreach (Reference reference in vsp.References)
                    {
                        if (!string.IsNullOrWhiteSpace(reference.Path))
                        {
                            yield return Path.GetDirectoryName(reference.Path);
                        }
                    }
                }
            }
        }

        private DTE GetDte()
        {
            DTE dte = (DTE)GetService(typeof(DTE));
            return dte;
        }

        private string[] GetReferenceDirectories()
        {
            DTE dte = (DTE)GetService(typeof(DTE));

            if (string.IsNullOrEmpty(dte.Solution.FileName))
                return new string[0];

            var dirs =
                new string[] { Path.Combine(Path.GetDirectoryName(dte.Solution.FileName), "packages") }
                    .Concat(GetReferenceDirectoriesEnum())
                    .Select(d => Path.GetFullPath(d))
                    .OrderBy(s => s.Length)
                    .ToList();


            for (int i = 1; i < dirs.Count; i++)
            {
                if (dirs.Where(d => dirs[i].StartsWith(d, StringComparison.OrdinalIgnoreCase), 0, i - 1).Any())
                    dirs.RemoveAt(i--);
            }

            return dirs.ToArray();
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

                DynamicTypeService typeResolver = (DynamicTypeService)GetService(typeof(DynamicTypeService));
                string path = GetAssemblyPath(args.Name);
                if (!string.IsNullOrWhiteSpace(path))
                    return typeResolver.CreateDynamicAssembly(path);


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
            DynamicTypeService typeResolver = (DynamicTypeService)GetService(typeof(DynamicTypeService));
            IVsToolboxService2 toolBoxService = GetService(typeof(IToolboxService)) as IVsToolboxService2;

            if (toolBoxService != null)
            {
                List<Tuple<string, ToolboxItem, IEnumerable<FrameworkName>, Guid>> items = new List<Tuple<string, ToolboxItem, IEnumerable<FrameworkName>, Guid>>();
                HashSet<string> processedAssemblies = new HashSet<string>();
                foreach (var directory in GetReferenceDirectories())
                {
                    foreach (string fileName in Directory.GetFiles(directory, "*.dll", SearchOption.AllDirectories))
                    {
                        var assembly = typeResolver.CreateDynamicAssembly(fileName);
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
                        }
                    }
                }
                toolBoxService.AddToolboxItems(items, null);
            }
            // Show a Message Box to prove we were here
            IVsUIShell uiShell = (IVsUIShell)GetService(typeof(SVsUIShell));
            Guid clsid = Guid.Empty;
            int result;
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
                       out result));
        }

    }

    public class ResolverToolboxItem : ToolboxItem
    {
        public ResolverToolboxItem(Type toolType)
            : base(toolType)
        {
            
        }

        protected ResolverToolboxItem(SerializationInfo info, StreamingContext context)
        {
            Deserialize(info, context);
        }
    }
}

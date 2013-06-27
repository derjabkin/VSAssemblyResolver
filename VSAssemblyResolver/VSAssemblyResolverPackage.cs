﻿using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.ComponentModel.Design;
using Microsoft.VisualStudio.Shell.Design;
using Microsoft.Win32;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using EnvDTE;
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
    [PackageRegistration(UseManagedResourcesOnly = true)]
    // This attribute is used to register the information needed to show this package
    // in the Help/About dialog of Visual Studio.
    [InstalledProductRegistration("#110", "#112", "0.0.3", IconResourceID = 400)]
    // This attribute is needed to let the shell know that this package exposes some menus.
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionHasSingleProject_string)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionHasMultipleProjects_string)]
    [Guid(GuidList.guidVSAssemblyResolverPkgString)]
    public sealed class VSAssemblyResolverPackage : Package
    {
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
        }

        Assembly CurrentDomain_ReflectionOnlyAssemblyResolve(object sender, ResolveEventArgs args)
        {
            string path = GetAssemblyPath(args.Name);
            if (!string.IsNullOrWhiteSpace(path))
                return Assembly.ReflectionOnlyLoadFrom(path);

            return null;
        }


        private IEnumerable<string> GetReferenceDirectoriesEnum()
        {
            DTE dte = (DTE)GetService(typeof(DTE));
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

        private string[] GetReferenceDirectories()
        {
            return GetReferenceDirectoriesEnum().Distinct().ToArray();
        }


        private string GetAssemblyPath(string name)
        {
            DynamicTypeService typeResolver = (DynamicTypeService)GetService(typeof(DynamicTypeService));
            var dirs = GetReferenceDirectories();
            AssemblyName asmName = new AssemblyName(name);
            
            foreach (var dir in dirs)
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

            return null;
        }

        private static bool IsNameCompatible(AssemblyName fullName, AssemblyName partialName)
        {
            return partialName.Name == fullName.Name &&
                   ((partialName.Version == null) || (partialName.Version == fullName.Version)) &&
                   ((partialName.GetPublicKeyToken() == null ||
                    partialName.GetPublicKeyToken().SequenceEqual(fullName.GetPublicKeyToken())));
        }


        private Assembly FindLoadedAssembly(AssemblyName asmName)
        {
            
            
            return AppDomain.CurrentDomain.GetAssemblies().FirstOrDefault(a => IsNameCompatible(a.GetName(), asmName));
        }


        private static bool resolving;

        System.Reflection.Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            if (resolving)
                return null;

            resolving = true;
            try
            {

                IVsSolution solution = (IVsSolution) GetService(typeof (IVsSolution));
                DTE dte = (DTE) GetService(typeof (DTE));
                var project = ((IEnumerable<object>) dte.ActiveSolutionProjects).OfType<Project>().FirstOrDefault();

                IVsHierarchy hierarchy = null;

                if (project != null)
                    solution.GetProjectOfUniqueName(project.FullName, out hierarchy);

                AssemblyName asmName = new AssemblyName(args.Name);

                Assembly loaded = FindLoadedAssembly(asmName);
                if (loaded != null) return loaded;

                DynamicTypeService typeResolver = (DynamicTypeService) GetService(typeof (DynamicTypeService));
                if (hierarchy != null)
                {
                    var trs = typeResolver.GetTypeResolutionService(hierarchy);
                    var trsAssembly = trs.GetAssembly(asmName);
                    if (trsAssembly != null)
                        return trsAssembly;
                }
                string path = GetAssemblyPath(args.Name);
                if (!string.IsNullOrWhiteSpace(path))
                    return typeResolver.CreateDynamicAssembly(path);


                return null;
            }
            finally
            {
                resolving = false;
            }
        }
        #endregion

        /// <summary>
        /// This function is the callback used to execute a command when the a menu item is clicked.
        /// See the Initialize method to see how the menu item is associated to this function using
        /// the OleMenuCommandService service and the MenuCommand class.
        /// </summary>
        private void MenuItemCallback(object sender, EventArgs e)
        {
            // Show a Message Box to prove we were here
            IVsUIShell uiShell = (IVsUIShell)GetService(typeof(SVsUIShell));
            Guid clsid = Guid.Empty;
            int result;
            Microsoft.VisualStudio.ErrorHandler.ThrowOnFailure(uiShell.ShowMessageBox(
                       0,
                       ref clsid,
                       "VSAssemblyResolver",
                       string.Format(CultureInfo.CurrentCulture, "Inside {0}.MenuItemCallback()", this.ToString()),
                       string.Empty,
                       0,
                       OLEMSGBUTTON.OLEMSGBUTTON_OK,
                       OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST,
                       OLEMSGICON.OLEMSGICON_INFO,
                       0,        // false
                       out result));
        }

    }
}
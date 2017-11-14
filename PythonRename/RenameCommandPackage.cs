using Microsoft.VisualStudio.TextManager.Interop;
using Microsoft.VisualStudio.Threading;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Windows.Controls;
using System.Windows;
using System.Windows.Media;
using System;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.Win32;
using System.Linq;
using System.Collections.Generic;
using System.Windows.Forms.VisualStyles;
using EnvDTE;

namespace Gekka.VisualStudio.Extension.PythonRename
{
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)] // Info on this package for Help/About
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(RenameCommandPackage.PackageGuidString)]
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
    [ProvideAutoLoad(UIContextGuids.SolutionExists)]
    [ProvideAutoLoad(UIContextGuids.NoSolution)]
    public sealed class RenameCommandPackage : Package
    {
        public const string PackageGuidString = "a49131ee-08d2-4c01-8aa8-3410ab73f957";

        //private System.Guid KIND_PYTHON = new Guid("6bb5f8ee-4483-11d3-8bcf-00c04f8ec28c");

        #region Package Members
        protected override void Initialize()
        {
            GLOBAL.Initialize(this);

            HookLoaded(typeof(Microsoft.Internal.VisualStudio.PlatformUI.EditableItemTextBox)
                , (sender, e) =>
                {
                    var items = SolutionExplorerTool.GetSolutionSelectedProjectItems(".py").ToArray();
                    if (items.Length != 1)
                    {
                        return;
                    }
                    var txb = (Microsoft.Internal.VisualStudio.PlatformUI.EditableItemTextBox)sender;
                    if (!SolutionExplorerTool.IsSolutionExploreTextBox(txb))
                    {
                        return;
                    }

                    ProjectItem item = items[0];
                    string org = txb.Text;
                    string newName = org;
                    Action reload = () =>
                       {
                           string newText = txb.Text;
                           txb.Text = org;
                           if (org != newText)
                           {
                               if (SolutionExplorerTool.IsValidNewFileName(item, newText))
                               {
                                   string[] keys = { "BuildAction", "ItemType", "Publish", "Node" };
                                   Dictionary<string, object> dic = new Dictionary<string, object>();
                                   foreach (string key in keys)
                                   {
                                       try
                                       {
                                           dic[key] = item.Properties.Item(key).Value;
                                       }
                                       catch (ArgumentException)
                                       {
                                       }
                                   }

                                   bool isOpened = item.IsOpen[EnvDTE.Constants.vsViewKindCode];
                                   var newItem = SolutionExplorerTool.ReloadFile(item, newText);
                                   foreach (string key in keys)
                                   {
                                       if (dic.ContainsKey(key))
                                       {
                                           try
                                           {
                                               item.Properties.Item(key).Value = dic[key];
                                           }
                                           catch (ArgumentException)
                                           {
                                           }
                                           catch (InvalidOperationException)
                                           {
                                           }
                                       }
                                   }
                                   org = newText;
                               }
                           }
                       };
                    txb.PreviewKeyDown += (s2, e2) =>
                    {
                        if (e2.Key == System.Windows.Input.Key.Return)
                        {
                            e2.Handled = true;
                            reload();
                        }
                    };
                    txb.PreviewLostKeyboardFocus += (s2, e2) =>
                    {
                        reload();
                    };
                });

            base.Initialize();
        }

        public static void HookLoaded(Type t, System.Windows.RoutedEventHandler handler)
        {
            System.Windows.Style style = new System.Windows.Style(t);
            System.Windows.EventSetter es = new System.Windows.EventSetter(System.Windows.FrameworkElement.LoadedEvent, handler);
            style.Setters.Add(es);
            System.Windows.Application.Current.Resources.Add(t, style);
        }
        #endregion
    }

    class GLOBAL
    {
        public static void Initialize(System.IServiceProvider sp)
        {
            _DTE = sp.GetService(typeof(EnvDTE.DTE)) as EnvDTE80.DTE2;
        }
        public static EnvDTE80.DTE2 DTE
        {
            get
            {
                return _DTE;
            }
        }
        private static EnvDTE80.DTE2 _DTE;

        public static EnvDTE.CommandEvents CommandEvents
        {
            get
            {
                if (_CommandEvents == null && _DTE != null)
                {
                    _CommandEvents = _DTE.Events.CommandEvents;
                }
                return _CommandEvents;
            }
        }
        private static EnvDTE.CommandEvents _CommandEvents;
    }

    class SolutionExplorerTool
    {
        public static IEnumerable<EnvDTE.ProjectItem> GetSolutionSelectedProjectItems()
        {
            var selectedItems = GLOBAL.DTE.ToolWindows.SolutionExplorer.SelectedItems as System.Collections.IList;
            if (selectedItems != null)
            {
                return selectedItems.OfType<EnvDTE.UIHierarchyItem>().Select(_ => _.Object).OfType<EnvDTE.ProjectItem>().ToArray();
            }
            else
            {
                return new EnvDTE.ProjectItem[0];
            }
        }

        public static IEnumerable<EnvDTE.ProjectItem> GetSolutionSelectedProjectItems(string ext)
        {
            foreach (EnvDTE.ProjectItem projItem in GetSolutionSelectedProjectItems())
            {
                foreach (EnvDTE.Property p in projItem.Properties)
                {
                    if (string.Equals(p.Name, "Extension", StringComparison.InvariantCultureIgnoreCase))
                    {
                        if (p.Value == null && string.IsNullOrEmpty(ext))
                        {
                            yield return projItem;
                        }
                        else if (p.Value != null && p.Value.ToString() == ext)
                        {
                            yield return projItem;
                        }
                    }
                }
            }
        }

        public static bool HasExtensionInSelectedProjectItems(string ext)
        {
            return GetSolutionSelectedProjectItems(ext).Any();
        }

        public static EnvDTE.ProjectItem ReloadFile(EnvDTE.ProjectItem item, string newFileName)
        {
            if (item.FileCount != 1)
            {
                throw new InvalidOperationException();
            }
            string pathOrg = item.FileNames[0];
            string pathNew = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(pathOrg), newFileName);
            var proj = item.ContainingProject;
            if (proj != null)
            {
                bool isOpened = item.IsOpen;
                try
                {
                    item.Remove();
                    System.IO.File.Move(pathOrg, pathNew);
                    proj.ProjectItems.AddFromFile(pathNew);

                    foreach (EnvDTE.ProjectItem i in proj.ProjectItems)
                    {
                        if (i.FileCount == 1 && i.FileNames[0] == pathNew)
                        {
                            return i;
                        }
                    }
                }
                catch (Exception ex)
                {
                    if (System.IO.File.Exists(pathOrg))
                    {
                        proj.ProjectItems.AddFromFile(pathOrg);
                    }
                    System.Windows.Forms.MessageBox.Show(ex.Message, "PyRename", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                }
            }
            return null;
        }

        public static bool IsValidNewFileName(string newFileName)
        {
            var item = GetSolutionSelectedProjectItems().FirstOrDefault();
            return IsValidNewFileName(item, newFileName);
        }

        public static bool IsValidNewFileName(EnvDTE.ProjectItem item, string newFileName)
        {
            if (item == null)
            {
                return false;
            }
            string path = item.FileNames[0];
            string path2 = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(path), newFileName);
            try
            {
                var info = new System.IO.FileInfo(path2);
                return !System.IO.File.Exists(path2);
            }
            catch
            {
                return false;
            }
        }

        public static bool IsSolutionExploreTextBox(DependencyObject d)
        {
            return GetSolutionExploreControl(d) != null;
        }

        public static System.Windows.DependencyObject GetSolutionExploreControl(DependencyObject d)
        {
            while (d != null)
            {
                System.Diagnostics.Debug.WriteLine(d.GetType().ToString());
                if (d.GetType().FullName == "Microsoft.VisualStudio.PlatformUI.SolutionPivotTreeView")
                {
                    return d as System.Windows.Controls.ListBox;
                }
                d = VisualTreeHelper.GetParent(d);
            }
            return null;
        }
    }
}
;
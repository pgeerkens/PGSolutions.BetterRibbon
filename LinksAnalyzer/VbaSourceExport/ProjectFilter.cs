////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.IO;
using System.Text;

using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Vbe.Interop;

namespace PGSolutions.RibbonUtilities.VbaSourceExport {
    [CLSCompliant(false)]
    public abstract class ProjectFilter : IProjectFilter {
        public ProjectFilter(IApplication application, string description, string extensions) {
            Application = application;
            Description = description;
            Extensions  = extensions;
        }

        /// <inheritdoc/>
        public string Description { get; }

        /// <inheritdoc/>
        public string Extensions  { get; }

        /// <inheritdoc/>
        public IApplication Application { get; }

        /// <inheritdoc/>
        public abstract void ExtractProjects(FileDialogSelectedItems Items, bool destIsSrc);

        /// <summary>Exports modules from specified EXCEL workbook to an eponymous subdirectory.</summary>
        public virtual void ExtractOpenProject(bool destIsSrc)
        => ExtractOpenProject(Application.ActiveWorkbook,destIsSrc);

        /// <inheritdoc/>
        protected virtual void ExtractOpenProject(Workbook wkbk, bool destIsSrc)
        => ExtractProjectModules(wkbk.VBProject, CreateDirectory(wkbk.FullName, destIsSrc));

        protected void ExtractProjectModules(VBProject project, string path) {
            try {
                foreach (VBComponent component in project.VBComponents) {
                    SetStatusBarText(project.Name, component.Name);
                    component.Export(Path.ChangeExtension(Path.Combine(path, component.Name),
                            TypeExtension((VbExt_ct)component.Type)));
                }

                File.WriteAllText(Path.Combine(path, "VBAProject.xml"), GetProjectDefinitionXml(project));
            //} catch (COMException ex) when (ex.HResult == unchecked((int)0x800AC372)) {
            //    $"Directory conflict occurred. Please retry.".ShowMsgString();
            } finally {
                Application.StatusBar = false;
            }
        }

        protected void SetStatusBarText(string projectName, string componentName)
        => Application.StatusBar = $"Exporting {projectName}.{componentName} ...";

        /// <summary>Prepares this exporter by providing a directory as destination for exports.</summary>
        /// <param name="path">Full (absolute) path-name for the project being exported.</param>
        /// <param name="destIsSrc">True if the destination folder is to be named 'src' (rather than being eponymous with the project).</param>
        protected static string CreateDirectory(string path, bool destIsSrc) {
            var basePath = destIsSrc ? Path.Combine(Path.GetDirectoryName(path), "src")
                                     : Path.Combine(Path.GetDirectoryName(path),Path.GetFileNameWithoutExtension(path) + "VBA");

            if (Directory.Exists(basePath)) Directory.Delete(basePath,true);

            return Directory.CreateDirectory(basePath).FullName;
        }

        private enum VbExt_ct {
            vbext_ct_StdModule      = 1,
            vbext_ct_ClassModule    = 2,
            vbext_ct_MSForm         = 3,
            vbext_ct_Document       = 100
        }

        private static string GetProjectDefinitionXml(VBProject project) {
            var sb = new StringBuilder()
                    .AppendLine($"<Project")
                    .AppendLine($"  Name='{project.Name}'")
                    .AppendLine($"  FileName='{project.FileName}'")
                    .AppendLine($"  HelpContextID='{project.HelpContextID}'")
                    .AppendLine($"  HelpFile='{project.HelpFile}'")
                    .AppendLine($"  Protection='{project.Protection}'")
                    .AppendLine($"  Type='{project.Type}'")
                    .AppendLine($">");
            foreach (Reference r in project.References) {
                  sb.AppendLine($"   <References")
                    .AppendLine($"      Description='{r.Description}'")
                    .AppendLine($"      FullPath='{r.FullPath}'")
                    .AppendLine($"      Guid='{r.Guid}'")
                    .AppendLine($"      Major='{r.Major}'")
                    .AppendLine($"      Minor='{r.Minor}'")
                    .AppendLine($"      Name='{r.Name}'")
                    .AppendLine($"      Type='{r.Type}'")
                    .AppendLine($"   />");
            }

            return sb.AppendLine("</Project>").ToString();
        }

        /// <summary>Returns an appropriate file extension (prefixed with '.') for the supplied moduleType ordinal.</summary>
        private static string TypeExtension(VbExt_ct moduleType) =>
               moduleType == VbExt_ct.vbext_ct_StdModule ? "vba"
            :  moduleType == VbExt_ct.vbext_ct_MSForm    ? "frm"
            : (moduleType == VbExt_ct.vbext_ct_ClassModule
            || moduleType == VbExt_ct.vbext_ct_Document) ? "cls"
                                                         : "unk";
    }
}

﻿////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;

namespace PGSolutions.RibbonUtilities.VbaSourceExport {
    [CLSCompliant(false)]
    public abstract class ProjectFilter : IProjectFilter {
        protected ProjectFilter(string description, string extensions) {
            Description = description;
            Extensions  = extensions;
        }

        internal static event EventHandler<EventArgs<string>> StatusAvailable;

        protected static void OnStatusAvailable(object sender, string message)
        => StatusAvailable?.Invoke(sender, new EventArgs<string>(message));

        /// <inheritdoc/>
        public string Description { get; }

        /// <inheritdoc/>
        public string Extensions  { get; }

        /// <inheritdoc/>
        public abstract void ExtractProjects(FileDialogSelectedItems items, bool destIsSrc);

        [SuppressMessage("Microsoft.Usage", "CA2202:Do not dispose objects multiple times",
                Justification = "Appears to be a spurious, or perhaps irreconcilable, warning.")]
        protected static void ExtractProjectModules(VBProject project, string path) {
            if (project == null ) throw new ArgumentNullException(nameof(project));

            var retries = 3;
            while (retries > 0) {
                try {
                    foreach (VBComponent component in project.VBComponents) {
                        component.Export(Path.ChangeExtension(Path.Combine(path, component.Name),
                                TypeExtension((VbExt_ct)component.Type)));
                    }

                    File.WriteAllText(Path.Combine(path, "VBAProject.xml"), GetProjectDefinitionXml(project));
                    break;
                }
                catch (COMException ex) when (ex.HResult == unchecked((int)0x800AC372)
                                          ||  ex.HResult == unchecked((int)0x800AC373)
                                          ||  ex.HResult == unchecked((int)0x800AC35C)) {
                    if (retries-- > 0) continue;

                    throw new IOException(
                        $"A file or directory conflict occurred. Please retry.\n\nPath:\n{path}", ex);
                }
            }
        }

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
               moduleType == VbExt_ct.vbext_ct_StdModule ? "vb"
            :  moduleType == VbExt_ct.vbext_ct_MSForm    ? "frm"
            : (moduleType == VbExt_ct.vbext_ct_ClassModule
            || moduleType == VbExt_ct.vbext_ct_Document) ? "cls"
                                                         : "unk";
    }
}

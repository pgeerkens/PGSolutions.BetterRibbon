////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2018 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.IO;
using System.Reflection;
using System.Resources;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using PGSolutions.RibbonDispatcher.AbstractCOM;
using PGSolutions.RibbonDispatcher.ConcreteCOM;

namespace PGSolutions.SampleRibbon {
    /// <summary>The publicly available entry points to the library.</summary>
    [Serializable]
    [ComVisible(true)]
    [CLSCompliant(false)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IRibbonLoader))]
    public class RibbonViewModel : AbstractRibbonViewModel, IRibbonLoader { //, IRibbonExtensibility {
        public RibbonViewModel() { }

        const string _assemblyName = "PGSolutions.ExampleRibbon";

    //    public string GetCustomUI(string RibbonID) => GetResourceText($"{_assemblyName}.SampleRibbon.xml");

        public void InitializeRibbon(IRibbonUI ribbonUI) {
            Initialize(RibbonUI = ribbonUI);
            RibbonModel = Globals.ThisWorkbook.Application.Run("RibbonLoader.NewRibbonModel");
        }

        /// <inheritdoc/>
        public void ReinitializeRibbon() {
            Initialize(RibbonUI);
            RibbonModel = Globals.ThisWorkbook.Application.Run("RibbonLoader.NewRibbonModel");
        }

        IRibbonViewModel IRibbonLoader.RibbonViewModel => this;

        private IRibbonUI RibbonUI { get; set; }
        public IRibbonModel RibbonModel { get; private set; }

        #region Helpers
        private static string GetResourceText(string resourceName) {
            var asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i) {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0) {
                    using (var resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i]))) {
                        if (resourceReader != null) {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }
        #endregion

        protected override Lazy<ResourceManager> ResourceManager => new Lazy<ResourceManager>(
            () => new ResourceManager($"{_assemblyName}.Properties.Resources", Assembly.GetExecutingAssembly())
        );
    }
}

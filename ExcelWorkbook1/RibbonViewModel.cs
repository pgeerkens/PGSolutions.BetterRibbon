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
    [ComDefaultInterface(typeof(IRibbonExtensibility))]
   // [ProgId("PGSolutions.ExampleRibbon")]
    [ProgId("ExampleRibbon")]
    public class RibbonViewModel : AbstractRibbonViewModel, IRibbonExtensibility {
        public RibbonViewModel() { }

        const string _AssemblyName = "ExampleRibbon";

        public string GetCustomUI(string RibbonID) => GetResourceText("PGSolutions.ExampleRibbon.SampleRibbon.xml");

        public void OnRibbonLoad(IRibbonUI ribbonUI) {
            Initialize(ribbonUI, this);
            ReinitializeRibbon();
        }

        /// <inheritdoc/>
        public void ReinitializeRibbon() =>
            RibbonModel = Globals.ThisWorkbook.Application.Run("RibbonLoader.NewRibbonModel");

        public  IRibbonModel RibbonModel { get; private set; }

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
            () => new ResourceManager($"{_AssemblyName}.Properties.Resources", Assembly.GetExecutingAssembly())
        );
    }
}

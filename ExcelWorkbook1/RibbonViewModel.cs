////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2018 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ComClasses;
using PGSolutions.RibbonDispatcher.Utilities;

namespace PGSolutions.ExampleRibbon {
    /// <summary>The publicly available entry points to the library.</summary>
    [Serializable]
    [ComVisible(true)]
    [CLSCompliant(false)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IRibbonViewModel))]
    public class RibbonViewModel : AbstractRibbonViewModel, IRibbonExtensibility {
        public RibbonViewModel(IRibbonUI ribbonUI) : this() => OnRibbonLoad(ribbonUI);
        public RibbonViewModel() : base(new LocalResourceManager(_assemblyName)) { }

        const string _assemblyName = "PGSolutions.ExampleRibbon";

        public string GetCustomUI(string RibbonID) => GetResourceText($"{_assemblyName}.SampleRibbon.xml");

        public override void OnRibbonLoad(IRibbonUI ribbonUI) {
            base.OnRibbonLoad(ribbonUI);
            InitializeModel();
        }

        /// <inheritdoc/>
        public void InitializeModel() =>
            RibbonModel = Globals.ThisWorkbook.Application.Run("RibbonLoader.NewRibbonModel");

        internal  IRibbonModel RibbonModel { get; private set; }

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
    }
}

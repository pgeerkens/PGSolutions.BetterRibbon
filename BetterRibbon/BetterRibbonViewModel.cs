////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

using Microsoft.Office.Core;

using PGSolutions.RibbonDispatcher.ComClasses;

using BetterRibbon.Properties;

namespace PGSolutions.BetterRibbon {
    /// <summary>The (top-level) ViewModel for the ribbon interface.</summary>
    /// <remarks>
    /// <a href=" https://go.microsoft.com/fwlink/?LinkID=271226"> For more information about adding callback methods.</a>
    /// 
    /// Take care renaming this class, or its namespace; and coordinate any such with the content
    /// of the (hidden) ThisAddIn.Designer.xml file. Commit frequently. Excel is very tempermental
    /// on the naming of ribbon objects and provides poor, and very minimal, diagnostic information.
    /// 
    /// This class MUST be ComVisible for the ribbon to launch properly.
    /// </remarks>
    [Description("The (top-level) ViewModel for the ribbon interface.")]
    [ComVisible(true)]
    [CLSCompliant(false)]
    [SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable",
            Justification = "Public, Non-Creatable, class with exported Events.")]
    //[Guid("A8ED8DFB-C422-4F03-93BF-FB5453D8F213")]
    public sealed class BetterRibbonViewModel : AbstractRibbonViewModel, IRibbonExtensibility {
        const string _assemblyName = "BetterRibbon";

        internal BetterRibbonViewModel(string controlId)
        : base(new LocalResourceManager(_assemblyName))
        => Id = controlId;

        #region IRibbonExtensibility implementation
        /// <inheritdoc/>
        public override string GetCustomUI(string RibbonID) => Resources.Ribbon;
        #endregion

        /// <summary>.</summary>
        protected override string Id { get; }

        /// <summary>Creates, registers, and returns a new <see cref="RibbonGroupViewModel"/> of the specified name.</summary>
        public override RibbonGroupViewModel AddGroupViewModel(string groupName)
        => AddGroupViewModel(RibbonFactory.NewRibbonGroup(groupName));

        /// <summary>Registers and returns the <see cref="RibbonGroupViewModel"/> created by the supplied delegate.</summary>
        public override RibbonGroupViewModel AddGroupViewModel(
                        Func<IRibbonFactory, RibbonGroupViewModel> func)
        => AddGroupViewModel(func?.Invoke(RibbonFactory));

        /// <summary>Registers and returns the supplied <see cref="RibbonGroupViewModel"/></summary>
        private RibbonGroupViewModel AddGroupViewModel(RibbonGroupViewModel viewModel) {
            GroupViewModels.Add(viewModel);
            return viewModel;
        }

        private IList<RibbonGroupViewModel> GroupViewModels { get; }
                                        = new List<RibbonGroupViewModel>();
    }
}

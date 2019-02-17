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
    using GroupViewModelFactory = Func<IRibbonFactory,RibbonGroupViewModel>;

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

        public override RibbonGroupViewModel AddGroupViewModel(string groupName) {
            var viewModel = RibbonFactory.NewRibbonGroup(groupName);
            GroupViewModels.Add(viewModel);
            return viewModel;
        }

        public override RibbonGroupViewModel AddGroupViewModel(GroupViewModelFactory func) {
            if (func == null) throw new ArgumentNullException(nameof(func));

            var viewModel = func(RibbonFactory);
            GroupViewModels.Add(viewModel);
            return viewModel;
        }

        private IList<RibbonGroupViewModel> GroupViewModels { get; }
                                        = new List<RibbonGroupViewModel>();
    }
}

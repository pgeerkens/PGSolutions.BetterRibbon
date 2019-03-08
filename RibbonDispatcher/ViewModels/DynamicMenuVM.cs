////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.Xml.Linq;

using Microsoft.Office.Core;

namespace PGSolutions.RibbonDispatcher.ViewModels {
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Naming","CA1710:IdentifiersShouldHaveCorrectSuffix")]
    public class DynamicMenuVM : AbstractContainerVM<IDynamicMenuSource,IDynamicMenuVM>, IDynamicMenuVM,
            IActivatable<IDynamicMenuSource,IDynamicMenuVM>{
        internal DynamicMenuVM(ViewModelFactory factory, string itemId) : base(itemId)
        => Factory = factory;

        #region IActivatable implementation
        /// <summary>Attaches this control-model to the specified ribbon-control as data source and event sink.</summary>
        public override IDynamicMenuVM Attach(IDynamicMenuSource source) => Attach<DynamicMenuVM>(source);

        public override void Detach() { CheckSum = 0; base.Detach(); }

        public new bool ShowInactive => true;
        #endregion

        #region DynamicContent implementation
        public event ContentEventHandler GetContent;

        public event ClickedEventHandler ContentLoaded;

        public void OnGetContent(IRibbonControl control, out string content) {
            content = EmptyMenu;
            GetContent?.Invoke(control, ref content);

            var checkSum = GetHash(content);
            if (checkSum != CheckSum) {
                PurgeChildren();
                Controls = XDocument.Parse(content).Root.ParseXmlMenu(Factory);
                ContentLoaded?.Invoke(control);
                CheckSum = checkSum;
            }

            Invalidate(c => c.SetShowInactive(true));
        }

        private ViewModelFactory    Factory  { get; }

        private ulong               CheckSum { get; set; }

        private static ulong GetHash(string content) {
            var ba = new byte[8];
            for (int i=0, j=0; i < content.Length; i++, j++) {
                if (j==8) j = 0;
                ba[j] ^= (byte)content[i];
            }
            ulong result = 0;
            for (var j=0; j < 7; j++) result = (result + ba[j]) << 8;
            return result + ba[7];
        }
        #endregion

        private static string EmptyMenu =>
            @"<menu xmlns='http://schemas.microsoft.com/office/2009/07/customui'></menu>";
    }
}

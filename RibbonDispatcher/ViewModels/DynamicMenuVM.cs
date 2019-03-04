////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.Xml.Linq;

using Microsoft.Office.Core;

namespace PGSolutions.RibbonDispatcher.ViewModels {
    public class DynamicMenuVM : AbstractContainerVM<IDynamicMenuSource,IDynamicMenuVM>, IDynamicMenuVM,
            IActivatable<IDynamicMenuSource,IDynamicMenuVM>{
        internal DynamicMenuVM(ViewModelFactory factory, string itemId) : base(itemId)
        => Factory = factory;

        #region IActivatable implementation
        /// <summary>Attaches this control-model to the specified ribbon-control as data source and event sink.</summary>
        public override IDynamicMenuVM Attach(IDynamicMenuSource source) => Attach<DynamicMenuVM>(source);
        #endregion

        #region DynamicContent implementation
        public event ContentEventHandler GetContent;

        public event ClickedEventHandler ContentLoaded;

        public void OnGetContent(IRibbonControl control, out string content) {
            content = TestMenuContent;
            GetContent?.Invoke(control, ref content);

            var checkSum = GetHash(content);
            if (checkSum != CheckSum) {
                PurgeChildren();
                Controls = XDocument.Parse(TestMenuContent).Root.ParseXmlMenu(Factory);
                ContentLoaded?.Invoke(control);
                CheckSum = checkSum;
            }

            Invalidate(c => c.SetShowInactive(true));
        }

        private ViewModelFactory    Factory  { get; }

        private ulong               CheckSum { get; set; }

        public new bool ShowInactive => true;

        private ulong GetHash(string content) {
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

        private static string TestMenuContent =>
@"<menu xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
        <menuSeparator id='DynamicMenu2Sep1' title='A Press Button'/>

        <!-- Only one of getSelectedItemID and getSelectedIndex can be specified, though both are implemented. -->
        <!-- Attributes size and getSize are undefined for children of a menu control. -->
        <button id='Custom2VbaButton1'
            getImage='getImage' getShowImage='getShowImage' getShowLabel='getShowLabel'
            getLabel='getLabel' getScreentip='getScreentip' getSupertip='getSupertip' getKeytip='getKeytip'
            getVisible='getVisible' getEnabled='getEnabled'
            onAction='onAction'
        />
        <menuSeparator id='DynamicMenu2Sep2' title='A ToggleButton'/>

        <!--Only one of getSelectedItemID and getSelectedIndex can be specified, though both are implemented. -->
        <!-- Attributes size and getSize are undefined for children of a menu control. -->
        <toggleButton id='Custom2VbaToggleButton2'
            getImage='getImage' getShowImage='getShowImage' getShowLabel='getShowLabel'
            getLabel='getLabel' getScreentip='getScreentip' getSupertip='getSupertip' getKeytip='getKeytip'
            getVisible='getVisible' getEnabled='getEnabled' 
            onAction='onActionToggle' getPressed='getPressed'
        />
        <gallery id='Custom2VbaStaticGallery2'
            getLabel='getLabel' getScreentip='getScreentip' getSupertip='getSupertip' getKeytip='getKeytip'
            getEnabled='getEnabled' getVisible='getVisible' getDescription='getDescription'
            getShowImage='getShowImage' getShowLabel='getShowLabel' getImage='getImage'
            getItemCount='getItemCount' getItemID='getItemId'
            getItemHeight='getItemHeight' getItemWidth='getItemWidth'
            getItemLabel='getItemLabel' getItemScreentip='getItemScreentip' getItemSupertip='getItemSupertip'
            getItemImage='getItemImage' getSelectedItemIndex='getSelectedItemIndex'
            onAction='onActionSelected' columns='7' rows='5' invalidateContentOnDrop='false'
            sizeString='WW' showItemImage='false' showItemLabel='true'
            >
            <item id='G2item11' label='27'/>
            <item id='G2item12' label='28'/>
            <item id='G2item13' label='29'/>
            <item id='G2item14' label='30'/>
            <item id='G2item15' label='31'/>
            <item id='G2item16' label=' 1'/>
            <item id='G2item17' label=' 2'/>
                                
            <item id='G2item21' label=' 3'/>
            <item id='G2item22' label=' 4'/>
            <item id='G2item23' label=' 5'/>
            <item id='G2item24' label=' 6'/>
            <item id='G2item25' label=' 7'/>
            <item id='G2item26' label=' 8'/>
            <item id='G2item27' label=' 9'/>
                                
            <item id='G2item31' label='10'/>
            <item id='G2item32' label='11'/>
            <item id='G2item33' label='12'/>
            <item id='G2item34' label='13'/>
            <item id='G2item35' label='14'/>
            <item id='G2item36' label='15'/>
            <item id='G2item37' label='16'/>

            <item id='G2item41' label='17'/>
            <item id='G2item42' label='18'/>
            <item id='G2item43' label='19'/>
            <item id='G2item44' label='20'/>
            <item id='G2item45' label='21'/>
            <item id='G2item46' label='22'/>
            <item id='G2item47' label='23'/>

            <item id='G2item51' label='24'/>
            <item id='G2item52' label='25'/>
            <item id='G2item53' label='26'/>
            <item id='G2item54' label='27'/>
            <item id='G2item55' label='28'/>
            <item id='G2item56' label=' 1'/>
            <item id='G2item57' label=' 2'/>
        </gallery>
</menu>";
    }
}

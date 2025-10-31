using SAPbouiCOM;
using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Xml;
using Application = SAPbouiCOM.Framework.Application;

namespace CapgeminiConcur
{
    [FormAttribute("CapgeminiConcur.Form1", "Form1.b1f")]
    class Form1 : UserFormBase
    {
        private SAPbouiCOM.Form oForm;
        private SAPbouiCOM.Matrix oMatrixAvansi;
        private SAPbouiCOM.DBDataSource oDBDSAvansi;
        public Form1()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("Item_1").Specific));
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("Item_2").Specific));
            this.EditText2 = ((SAPbouiCOM.EditText)(this.GetItem("Item_3").Specific));
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("Item_4").Specific));
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("Item_5").Specific));
            this.EditText3 = ((SAPbouiCOM.EditText)(this.GetItem("Item_8").Specific));
            this.EditText8 = ((SAPbouiCOM.EditText)(this.GetItem("Item_13").Specific));
            this.EditText10 = ((SAPbouiCOM.EditText)(this.GetItem("Item_15").Specific));
            this.EditText11 = ((SAPbouiCOM.EditText)(this.GetItem("Item_16").Specific));
            this.EditText12 = ((SAPbouiCOM.EditText)(this.GetItem("Item_17").Specific));
            this.Folder0 = ((SAPbouiCOM.Folder)(this.GetItem("Item_20").Specific));
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_6").Specific));
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_7").Specific));
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_23").Specific));
            this.StaticText6 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_27").Specific));
            this.StaticText8 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_29").Specific));
            this.StaticText9 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_30").Specific));
            this.StaticText10 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_31").Specific));
            this.CheckBox0 = ((SAPbouiCOM.CheckBox)(this.GetItem("Item_32").Specific));
            this.StaticText11 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_33").Specific));
            this.StaticText12 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_34").Specific));
            this.StaticText13 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_35").Specific));
            this.EditText13 = ((SAPbouiCOM.EditText)(this.GetItem("Item_36").Specific));
            this.EditText14 = ((SAPbouiCOM.EditText)(this.GetItem("Item_37").Specific));
            this.EditText15 = ((SAPbouiCOM.EditText)(this.GetItem("Item_38").Specific));
            this.EditText16 = ((SAPbouiCOM.EditText)(this.GetItem("Item_39").Specific));
            this.StaticText14 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_41").Specific));
            this.Matrix0 = ((SAPbouiCOM.Matrix)(this.GetItem("Item_18").Specific));
            this.Matrix1 = ((SAPbouiCOM.Matrix)(this.GetItem("Item_43").Specific));
            this.Folder1 = ((SAPbouiCOM.Folder)(this.GetItem("Item_0").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
        }

        private void OnCustomInitialize()
        {
            
        }


        private SAPbouiCOM.EditText EditText0;
        private SAPbouiCOM.EditText EditText1;
        private SAPbouiCOM.EditText EditText2;
        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.Button Button1;
        private SAPbouiCOM.EditText EditText3;
        private SAPbouiCOM.EditText EditText8;
        private SAPbouiCOM.EditText EditText10;
        private SAPbouiCOM.EditText EditText11;
        private SAPbouiCOM.EditText EditText12;
        private SAPbouiCOM.Folder Folder0;
        private SAPbouiCOM.StaticText StaticText0;
        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.StaticText StaticText2;
        private SAPbouiCOM.StaticText StaticText6;
        private SAPbouiCOM.StaticText StaticText8;
        private SAPbouiCOM.StaticText StaticText9;
        private SAPbouiCOM.StaticText StaticText10;
        private SAPbouiCOM.CheckBox CheckBox0;
        private SAPbouiCOM.StaticText StaticText11;
        private SAPbouiCOM.StaticText StaticText12;
        private SAPbouiCOM.StaticText StaticText13;
        private SAPbouiCOM.EditText EditText13;
        private SAPbouiCOM.EditText EditText14;
        private SAPbouiCOM.EditText EditText15;
        private SAPbouiCOM.EditText EditText16;
        private SAPbouiCOM.StaticText StaticText14;
        private SAPbouiCOM.Matrix Matrix0;
        private SAPbouiCOM.Matrix Matrix1;
        private Folder Folder1;

    }
}
using SAPbouiCOM;
using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using Application = SAPbouiCOM.Framework.Application;

namespace CapgeminiConcur
{
    [FormAttribute("CapgeminiConcur.Form1", "Form1.b1f")]
    class Form1 : UserFormBase
    {
        private Form oForm;
        private Matrix oMatrixAvansi;
        private Matrix oMatrixTroskovi;
        private DBDataSource oDBDSAvansi;
        private DBDataSource oDBDSTroskovi;

        private SAPbobsCOM.Company oCompany;
        private bool _recalcGuard = false;
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
            this.Button0.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button0_ClickBefore);
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("Item_5").Specific));
            this.Button1.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button1_ClickBefore);
            //  this.Button1.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button1_ClickAfter);
            this.EditText3 = ((SAPbouiCOM.EditText)(this.GetItem("Item_8").Specific));
            this.EditText8 = ((SAPbouiCOM.EditText)(this.GetItem("Item_13").Specific));
            this.EditText10 = ((SAPbouiCOM.EditText)(this.GetItem("Item_15").Specific));
            this.EditText11 = ((SAPbouiCOM.EditText)(this.GetItem("Item_16").Specific));
            this.Folder0 = ((SAPbouiCOM.Folder)(this.GetItem("Item_20").Specific));
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_6").Specific));
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_7").Specific));
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_23").Specific));
            this.StaticText6 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_27").Specific));
            this.StaticText8 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_29").Specific));
            this.StaticText11 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_33").Specific));
            this.StaticText12 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_34").Specific));
            this.StaticText13 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_35").Specific));
            this.EditText13 = ((SAPbouiCOM.EditText)(this.GetItem("Item_36").Specific));
            this.EditText14 = ((SAPbouiCOM.EditText)(this.GetItem("Item_37").Specific));
            this.EditText15 = ((SAPbouiCOM.EditText)(this.GetItem("Item_38").Specific));
            this.Matrix0 = ((SAPbouiCOM.Matrix)(this.GetItem("Item_18").Specific));
            this.Matrix1 = ((SAPbouiCOM.Matrix)(this.GetItem("Item_43").Specific));
            this.Folder1 = ((SAPbouiCOM.Folder)(this.GetItem("Item_0").Specific));
            this.Button2 = ((SAPbouiCOM.Button)(this.GetItem("Item_9").Specific));
            this.Button2.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button2_ClickBefore);
            this.Button3 = ((SAPbouiCOM.Button)(this.GetItem("Item_10").Specific));
            this.Button3.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button3_ClickBefore);
            //    this.Button3.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button3_ClickAfter);
            this.Button4 = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.ComboBox0 = ((SAPbouiCOM.ComboBox)(this.GetItem("Item_50").Specific));
            this.StaticText3 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_12").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            Application.SBO_Application.ItemEvent += SBO_Application_ItemEvent;
        }

        private void OnCustomInitialize()
        {
            oForm = (Form)this.UIAPIRawForm;

            Application.SBO_Application.StatusBar.SetText("Forma učitana: " + oForm.UniqueID, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_None);

            try
            {
                // === Inicijalizacija matrica i data source-a ===
                oMatrixAvansi = (Matrix)oForm.Items.Item("Item_18").Specific;
                oDBDSAvansi = oForm.DataSources.DBDataSources.Item("@BO_CONCUR_CLAIM_OP");

                oMatrixTroskovi = (Matrix)oForm.Items.Item("Item_43").Specific;
                oDBDSTroskovi = oForm.DataSources.DBDataSources.Item("@BO_CONCUR_CLAIM_EXP");

                // B1f auto objekti:
                Matrix0 = oMatrixAvansi;
                Matrix1 = oMatrixTroskovi;

                oMatrixAvansi.Columns.Item("BrDok").Editable = true;
                oMatrixAvansi.Columns.Item("Opis").Editable = true;
                oMatrixAvansi.Columns.Item("Iznos").Editable = true;

                try
                {
                    AddCFL_OP();
                }
                catch (Exception ex)
                {
                    Application.SBO_Application.StatusBar.SetText(
                        $"Greška pri kreiranju CFL-a: {ex.Message}",
                        BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }

                // Nakon učitavanja forme i svih datasource-a:
                try
                {
                    IzracunajSume();   // odmah proračunaj sume po učitavanju forme
                }
                catch (Exception calcEx)
                {
                    Application.SBO_Application.StatusBar.SetText(
                        $"Sume nisu mogle da se izračunaju pri učitavanju forme: {calcEx.Message}",
                        BoMessageTime.bmt_Short,
                        BoStatusBarMessageType.smt_Warning
                    );
                }

            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText($"Greška pri inicijalizaciji forme: {ex.Message}",
                    BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        private void AddCFL_OP()
        {
            try
            {
                ChooseFromListCreationParams p =
                    (ChooseFromListCreationParams)Application.SBO_Application
                        .CreateObject(BoCreatableObjectType.cot_ChooseFromListCreationParams);

                p.ObjectType = "46"; // OT 46 → Outgoing Payments
                p.UniqueID = "CFL_OP";
                p.MultiSelection = false;

                ChooseFromList oCFL = oForm.ChooseFromLists.Add(p);

                Matrix mtx = (Matrix)oForm.Items.Item("Item_18").Specific;
                Column col = mtx.Columns.Item("BrDok");

                col.ChooseFromListUID = "CFL_OP";
                col.ChooseFromListAlias = "DocNum";
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    $"Greška pri kreiranju OP CFL: {ex.Message}",
                    BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        private void SBO_Application_ItemEvent(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            // Forma postoji?
            SAPbouiCOM.Form testForm = null;
            try { testForm = Application.SBO_Application.Forms.Item(FormUID); }
            catch { BubbleEvent = false; return; }

            if (pVal.FormTypeEx != "CapgeminiConcur.Form1")
                return;

            try
            {
                oForm = Application.SBO_Application.Forms.Item(FormUID);

                // ============================================================
                // === CFL – BrDok kolona (Custom Query CFL)                 ===
                // ============================================================
                if (pVal.ItemUID == "Item_18" &&
                    pVal.ColUID == "BrDok" &&
                    pVal.EventType == BoEventTypes.et_CHOOSE_FROM_LIST)
                {
                    if (oMatrixAvansi == null)
                        oMatrixAvansi = (Matrix)oForm.Items.Item("Item_18").Specific;

                    // === BEFORE ACTION ======================================
                    if (pVal.BeforeAction)
                    {
                        string employeeCode = ((EditText)oForm.Items.Item("Item_3").Specific).Value?.Trim();
                        if (string.IsNullOrEmpty(employeeCode))
                            return;

                        oCompany = Parametri.m_GetConnected_COMPANY();
                        var rs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                        // 1️⃣ Preuzmi empID
                        rs.DoQuery($@"SELECT ""empID"" FROM ""OHEM"" WHERE ""Code"" = '{employeeCode}'");
                        if (rs.EoF) return;

                        int empID = Convert.ToInt32(rs.Fields.Item("empID").Value);

                        // 2️⃣ Preuzmi sve OP za zaposlenog
                        rs.DoQuery($@"
                            SELECT DISTINCT T0.""DocEntry"", T0.""DocNum""
                            FROM ""OVPM"" T0
                            INNER JOIN ""VPM4"" T1 ON T0.""DocEntry"" = T1.""DocNum""
                            WHERE T0.""DocType"" = 'A'
                            AND T1.""U_EMPLOYEE"" = {empID}
                        ");

                        if (rs.RecordCount == 0)
                        {
                            Application.SBO_Application.MessageBox("Nema OP dokumenata za ovog zaposlenog.");
                            BubbleEvent = false;
                            return;
                        }

                        // ==============================================
                        // 🔥 3️⃣ Uklanjamo OP-ove koji su već u matrici
                        // ==============================================
                        List<string> allowed = new List<string>();

                        while (!rs.EoF)
                        {
                            string entry = rs.Fields.Item("DocEntry").Value.ToString();
                            string number = rs.Fields.Item("DocNum").Value.ToString();

                            bool alreadyUsed = false;

                            for (int i = 1; i <= oMatrixAvansi.RowCount; i++)
                            {
                                string existing =
                                    ((EditText)oMatrixAvansi.Columns.Item("BrDok")
                                        .Cells.Item(i).Specific).Value?.Trim();

                                if (existing == number)
                                {
                                    alreadyUsed = true;
                                    break;
                                }
                            }

                            if (!alreadyUsed)
                                allowed.Add(entry);

                            rs.MoveNext();
                        }

                        // 4️⃣ Ako nema više OP-ova — nema CFL
                        if (allowed.Count == 0)
                        {
                            Application.SBO_Application.StatusBar.SetText(
                                "Svi OP dokumenti zaposlenog su već iskorišćeni.",
                                BoMessageTime.bmt_Short,
                                BoStatusBarMessageType.smt_Warning
                            );

                            BubbleEvent = false;
                            return;
                        }

                        // 5️⃣ Postavi Conditions samo za preostale OP-ove
                        ChooseFromList oCFL = oForm.ChooseFromLists.Item("CFL_OP");
                        Conditions conds = (Conditions)Application.SBO_Application.CreateObject(BoCreatableObjectType.cot_Conditions);

                        for (int i = 0; i < allowed.Count; i++)
                        {
                            Condition c = conds.Add();
                            c.Alias = "DocEntry";
                            c.Operation = BoConditionOperation.co_EQUAL;
                            c.CondVal = allowed[i];

                            if (i < allowed.Count - 1)
                                c.Relationship = BoConditionRelationship.cr_OR;
                        }

                        oCFL.SetConditions(conds);
                    }

                    else // AFTER ACTION
                    {
                        try
                        {
                            IChooseFromListEvent ev = (IChooseFromListEvent)pVal;
                            DataTable dt = ev.SelectedObjects;

                            if (dt == null || dt.Rows.Count == 0)
                                return;

                            int row = pVal.Row;
                            oMatrixAvansi.FlushToDataSource();

                            string docNum = dt.GetValue("DocNum", 0).ToString();

                            // === 1️⃣ Uzmi trenutni ReportID ===
                            string reportID = ((EditText)oForm.Items.Item("Item_1").Specific).Value?.Trim();

                            // === 2️⃣ Proveri OP u OVPM ===
                            SAPbobsCOM.Recordset rsCheck =
                                (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                            rsCheck.DoQuery($@"
                                SELECT ""DocEntry"", ""U_REPORTID""
                                FROM ""OVPM""
                                WHERE ""DocNum"" = '{docNum}'
                            ");

                            if (rsCheck.EoF)
                            {
                                Application.SBO_Application.StatusBar.SetText(
                                    $"OP dokument {docNum} nije pronađen.",
                                    BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                return;
                            }

                            int docEntry = Convert.ToInt32(rsCheck.Fields.Item("DocEntry").Value);
                            string existingRepID = rsCheck.Fields.Item("U_REPORTID").Value.ToString().Trim();

                            // === 3️⃣ Ako je OP već korišćen u drugom obračunu → stop ===
                            if (!string.IsNullOrEmpty(existingRepID) && existingRepID != reportID)
                            {
                                Application.SBO_Application.MessageBox(
                                    $"Avans (OP {docNum}) je već iskorišćen!\n" +
                                    $"ReportID: {existingRepID}"
                                );
                                return;
                            }

                            // === 4️⃣ Ako nije korišćen – upiši trenutni ReportID ===
                            if (string.IsNullOrEmpty(existingRepID))
                            {
                                SAPbobsCOM.Payments pay =
                                    (SAPbobsCOM.Payments)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments);

                                if (!pay.GetByKey(docEntry))
                                {
                                    pay = (SAPbobsCOM.Payments)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments);
                                    pay.GetByKey(docEntry);
                                }

                                pay.UserFields.Fields.Item("U_REPORTID").Value = reportID;
                                int ret = pay.Update();

                                if (ret != 0)
                                {
                                    oCompany.GetLastError(out int errCode, out string errMsg);
                                    Application.SBO_Application.StatusBar.SetText(
                                        $"Greška pri upisu ReportID u OP {docNum}: {errMsg}",
                                        BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error
                                    );
                                    return;
                                }
                            }

                            // === 5️⃣ Nastavi normalno: upis memo + total u UDO ===

                            SAPbobsCOM.Recordset rs =
                                (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                            rs.DoQuery($@"
                                SELECT 
                                    T0.""Comments"",
                                    T0.""DocTotal""
                                FROM ""OVPM"" T0
                                WHERE T0.""DocNum"" = '{docNum}'
                            ");

                            string memo = rs.Fields.Item("Comments").Value.ToString();
                            string total = rs.Fields.Item("DocTotal").Value.ToString();

                            // Duplicate check
                            for (int i = 1; i <= oMatrixAvansi.RowCount; i++)
                            {
                                string existing =
                                    ((EditText)oMatrixAvansi.Columns.Item("BrDok").Cells.Item(i).Specific).Value;

                                if (existing == docNum && i != row)
                                {
                                    Application.SBO_Application.StatusBar.SetText(
                                        $"Dokument {docNum} je već dodat u redu {i}.",
                                        BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                                    return;
                                }
                            }

                            // Upis u UDO
                            oDBDSAvansi.SetValue("U_BrojDokumenta", row - 1, docNum);
                            oDBDSAvansi.SetValue("U_Opis", row - 1, memo);
                            oDBDSAvansi.SetValue("U_Iznos", row - 1, total);

                            oMatrixAvansi.LoadFromDataSource();
                            IzracunajSume();
                        }
                        catch (Exception ex)
                        {
                            Application.SBO_Application.StatusBar.SetText(
                                $"Greška u CFL AfterAction: {ex.Message}",
                                BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                        }
                    }

                }

                // ============================================================
                // === 2️⃣ ZAMENA za FORM_DATA_LOAD → DETEKCIJA SAP DUGMADI ===
                // ============================================================
                // Ovo radi ZA SVAKI LOAD, NEXT, PREV, FIND, UPDATE, ADD
                if (
                    (
                        pVal.ItemUID == "1" ||   // OK / Add / Update / Find
                        pVal.ItemUID == "1281" ||   // Find
                        pVal.ItemUID == "1282" ||   // Add
                        pVal.ItemUID == "1288" ||   // First
                        pVal.ItemUID == "1289" ||   // Last
                        pVal.ItemUID == "1290" ||   // Next
                        pVal.ItemUID == "1291"       // Previous
                    )
                    &&
                    (
                        pVal.EventType == BoEventTypes.et_ITEM_PRESSED ||
                        pVal.EventType == BoEventTypes.et_CLICK
                    )
                    &&
                    !pVal.BeforeAction
                    )
                {
                    try
                    {
                        DBDataSource dsHeader = oForm.DataSources.DBDataSources.Item("@BO_CONCUR_CLAIM_H");
                        string zatvoreno = dsHeader.GetValue("U_Zatvoreno", 0).Trim();

                        // Sync ComboBox
                        if (zatvoreno == "Y")
                            ComboBox0.Select("Y", BoSearchKey.psk_ByValue);
                        else
                            ComboBox0.Select("N", BoSearchKey.psk_ByValue);

                        // Zaključaj formu ako treba
                        if (zatvoreno == "Y")
                            DisableEditing();
                        else
                            EnableEditing();

                        // Reload matrica
                        oMatrixAvansi.LoadFromDataSource();
                        oMatrixTroskovi.LoadFromDataSource();

                        IzracunajSume();
                    }
                    catch (Exception ex)
                    {
                        Application.SBO_Application.StatusBar.SetText(
                            "Greška pri reloadu (SAP dugmad): " + ex.Message,
                            BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    }
                }

                // ============================================================
                // === 3️⃣ Recalc suma                                         ===
                // ============================================================
                if (!pVal.BeforeAction &&
                    (pVal.EventType == BoEventTypes.et_VALIDATE ||
                     pVal.EventType == BoEventTypes.et_LOST_FOCUS))
                {
                    if ((pVal.ItemUID == "Item_18" && pVal.ColUID == "Iznos") ||
                        (pVal.ItemUID == "Item_43" && pVal.ColUID == "Iznos1"))
                    {
                        IzracunajSume();
                    }
                }
                // === Univerzalni detektor promene UDO rekorda ===
                // Radi i za Next, Previous, Find, OK i sve ostalo
                if (pVal.EventType == BoEventTypes.et_FORM_ACTIVATE && !pVal.BeforeAction)
                {
                    try
                    {
                        // Kada pređeš na novi zapis — refresuj sve
                        oMatrixAvansi.LoadFromDataSource();
                        oMatrixTroskovi.LoadFromDataSource();
                        IzracunajSume();
                    }
                    catch { }
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    $"Greška u ItemEvent: {ex.Message}",
                    BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        private void DisableEditing()
        {
            try
            {
                oForm.Freeze(true);

                foreach (Item item in oForm.Items)
                {
                    string uid = item.UniqueID;

                    // --- SAP sistemska dugmad — ne diramo ---
                    bool isSystemButton =
                        uid == "1" ||   // OK / Update / Add / Find
                        uid == "2" ||   // Cancel SAP
                        uid == "1281" || // Find
                        uid == "1282" || // Add
                        uid == "1288" || uid == "1289" || // First / Last
                        uid == "1290" || uid == "1291";   // Next / Previous

                    // --- Tabovi moraju ostati uključeni ---
                    if (item.Type == BoFormItemTypes.it_FOLDER)
                        continue;

                    // --- Tvoje Cancel dugme (Item_5) ostaje aktivno ---
                    if (uid == "Item_5")
                        continue;

                    // --- ComboBox za zaključavanje (Item_50) mora ostati aktivan ---
                    if (uid == "Item_50")
                        continue;

                    // --- Sve ostalo se disabluje ---
                    if (!isSystemButton)
                        item.Enabled = false;
                }

                // Matrice — kolone ne smeju biti editable
                oMatrixAvansi.Columns.Item("BrDok").Editable = false;
                oMatrixAvansi.Columns.Item("Opis").Editable = false;
                oMatrixAvansi.Columns.Item("Iznos").Editable = false;

                oMatrixTroskovi.Columns.Item("Opis1").Editable = false;
                oMatrixTroskovi.Columns.Item("Iznos1").Editable = false;

                Application.SBO_Application.StatusBar.SetText(
                    "Dokument je zaključan za izmene.",
                    BoMessageTime.bmt_Short,
                    BoStatusBarMessageType.smt_Warning);
            }
            catch { }
            finally
            {
                try { oForm.Freeze(false); } catch { }
            }
        }

        private void EnableEditing()
        {
            try
            {
                oForm.Freeze(true);

                foreach (Item item in oForm.Items)
                {
                    string uid = item.UniqueID;

                    // Cancel dugme ostaje enabled
                    if (uid == "Item_5")
                        continue;

                    // Tabovi su uvek allowed
                    if (item.Type == BoFormItemTypes.it_FOLDER)
                        continue;

                    // Sve ostalo se ponovo uključuje
                    item.Enabled = true;
                }

                // Matrice — opet editable
                oMatrixAvansi.Columns.Item("BrDok").Editable = true;
                oMatrixAvansi.Columns.Item("Opis").Editable = true;
                oMatrixAvansi.Columns.Item("Iznos").Editable = true;

                oMatrixTroskovi.Columns.Item("Opis1").Editable = true;
                oMatrixTroskovi.Columns.Item("Iznos1").Editable = true;
            }
            catch { }
            finally
            {
                try { oForm.Freeze(false); } catch { }
            }
        }

        private void IzracunajSume()
        {
            if (_recalcGuard) return;
            _recalcGuard = true;

            try
            {
                oForm.Freeze(true);

                double sumaAvansa = 0.0;
                double sumaTroskova = 0.0;

                // === 1️⃣ Provera da li matrix avansa ima ijednu validnu stavku ===
                bool imaValidnihAvansa = false;

                for (int i = 1; i <= oMatrixAvansi.RowCount; i++)
                {
                    var cellBrDok = (SAPbouiCOM.EditText)oMatrixAvansi.Columns.Item("BrDok").Cells.Item(i).Specific;
                    string br = cellBrDok?.Value?.Trim() ?? "";

                    if (!string.IsNullOrEmpty(br))      // ako postoji ijedan popunjen avans
                    {
                        imaValidnihAvansa = true;
                        break;
                    }
                }

                // === 2️⃣ Ako nema validnih avansa → sumaAvansa ostaje 0.00 ===
                if (imaValidnihAvansa)
                {
                    // Saberi samo iznose iz popunjenih avansa
                    for (int i = 1; i <= oMatrixAvansi.RowCount; i++)
                    {
                        var cellIznos = (SAPbouiCOM.EditText)oMatrixAvansi.Columns.Item("Iznos").Cells.Item(i).Specific;
                        string val = cellIznos?.Value?.Trim() ?? "";

                        if (double.TryParse(val, System.Globalization.NumberStyles.Any,
                            System.Globalization.CultureInfo.InvariantCulture, out double iznos)
                            || double.TryParse(val, System.Globalization.NumberStyles.Any,
                            System.Globalization.CultureInfo.CurrentCulture, out iznos))
                        {
                            sumaAvansa += iznos;
                        }
                    }
                }
                else
                {
                    sumaAvansa = 0.00;
                }

                // === 3️⃣ Matrix za TROŠKOVE (uvek se sabiraju odmah) ===
                for (int i = 1; i <= oMatrixTroskovi.RowCount; i++)
                {
                    var cell = (SAPbouiCOM.EditText)oMatrixTroskovi.Columns.Item("Iznos1").Cells.Item(i).Specific;
                    string val = cell?.Value?.Trim() ?? "";

                    if (double.TryParse(val, System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture, out double iznos)
                        || double.TryParse(val, System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.CurrentCulture, out iznos))
                    {
                        sumaTroskova += iznos;
                    }
                }

                // === 4️⃣ Izračunaj razliku ===
                double razlika = sumaAvansa - sumaTroskova;

                // === 5️⃣ Upis u polja (bez DB update-a) ===
                EditText13.Value = sumaAvansa.ToString("F2", System.Globalization.CultureInfo.InvariantCulture); // suma avansa
                EditText14.Value = sumaTroskova.ToString("F2", System.Globalization.CultureInfo.InvariantCulture); // suma troskova
                EditText15.Value = razlika.ToString("F2", System.Globalization.CultureInfo.InvariantCulture); // razlika
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    $"Greška pri računanju suma: {ex.Message}",
                    BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                try { oForm.Freeze(false); } catch { }
                _recalcGuard = false;
            }
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
        private Button Button2;
        private Button Button3;

        private void Button2_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = false; // sprečava SAP da automatski doda svoj red

            /*try
            {
                oForm.Freeze(true);

                oMatrixAvansi.FlushToDataSource();

                // Dodaj ručno novi red
                oDBDSAvansi.InsertRecord(oDBDSAvansi.Size);
                oMatrixAvansi.LoadFromDataSource();
                oMatrixAvansi.SelectRow(oMatrixAvansi.RowCount, true, false);

                Application.SBO_Application.StatusBar.SetText(
                    "Novi red uspešno dodat.",
                    BoMessageTime.bmt_Short,
                    BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    $"Greška pri dodavanju reda: {ex.Message}",
                    BoMessageTime.bmt_Short,
                    BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }*/

            string zatvoreno = oForm.DataSources.DBDataSources
            .Item("@BO_CONCUR_CLAIM_H")
            .GetValue("U_Zatvoreno", 0).Trim();

            if (zatvoreno == "Y")
            {
                Application.SBO_Application.MessageBox("Dokument je zaključan. Nije dozvoljeno menjanje redova.");
                BubbleEvent = false;
                return;
            }

            BubbleEvent = true; 
            oMatrixAvansi.AddRow();
            oMatrixAvansi.ClearRowData(oMatrixAvansi.VisualRowCount);
            oMatrixAvansi.FlushToDataSource();
            oMatrixAvansi.LoadFromDataSource();
        }

        private void Button2_ClickAfter(object sboObject, SBOItemEventArg pVal)
        {
            try
            {
                oForm.Freeze(true);

                // 1️⃣ Sinhronizuj postojeće vrednosti iz matrice u data source
                oMatrixAvansi.FlushToDataSource();

                // 2️⃣ Proveri da li poslednji red uopšte ima neki unos
                bool lastRowEmpty = false;

                if (oMatrixAvansi.RowCount > 0)
                {
                    string lastDoc =
                        ((EditText)oMatrixAvansi.Columns.Item("BrDok").Cells.Item(oMatrixAvansi.RowCount).Specific).Value?.Trim();

                    // Ako je poslednji red prazan, ne dodaj novi
                    if (string.IsNullOrEmpty(lastDoc))
                        lastRowEmpty = true;
                }

                // 3️⃣ Samo ako poslednji red NIJE prazan — dodaj novi u data source
                if (!lastRowEmpty)
                {
                    oDBDSAvansi.InsertRecord(oDBDSAvansi.Size);
                }

                // 4️⃣ Osvježi prikaz matrice
                oMatrixAvansi.LoadFromDataSource();

                // 5️⃣ Selektuj novododati red (ako postoji)
                if (oMatrixAvansi.RowCount > 0)
                    oMatrixAvansi.SelectRow(oMatrixAvansi.RowCount, true, false);

                Application.SBO_Application.StatusBar.SetText(
                    "Novi red uspešno dodat.",
                    BoMessageTime.bmt_Short,
                    BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    $"Greška pri dodavanju reda: {ex.Message}",
                    BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        private void Button3_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                oForm.Freeze(true);

                string zatvoreno = oForm.DataSources.DBDataSources
                .Item("@BO_CONCUR_CLAIM_H")
                .GetValue("U_Zatvoreno", 0).Trim();

                if (zatvoreno == "Y")
                {
                    Application.SBO_Application.MessageBox("Dokument je zaključan. Nije dozvoljeno menjanje redova.");
                    BubbleEvent = false;
                    return;
                }

                CellPosition cell = oMatrixAvansi.GetCellFocus();
                if (cell.rowIndex <= 0)
                {
                    Application.SBO_Application.StatusBar.SetText(
                        "Nije selektovan red za brisanje.",
                        BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                    return;
                }

                // 1️⃣ Sinhronizuj pre brisanja
                oMatrixAvansi.FlushToDataSource();

                // 2️⃣ Ukloni red iz data source-a
                oDBDSAvansi.RemoveRecord(cell.rowIndex - 1);

                // 3️⃣ Osvježi matricu
                oMatrixAvansi.LoadFromDataSource();

                Application.SBO_Application.StatusBar.SetText(
                    $"Red {cell.rowIndex} uspešno obrisan.",
                    BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    $"Greška pri brisanju reda: {ex.Message}",
                    BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        private void Button3_ClickAfter(object sboObject, SBOItemEventArg pVal)
        {
            try
            {
                oForm.Freeze(true);

                if (oMatrixAvansi.RowCount == 0)
                {
                    Application.SBO_Application.StatusBar.SetText(
                        "Nema redova za brisanje.",
                        BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                    return;
                }

                // Dohvati selektovani red
                CellPosition cell = oMatrixAvansi.GetCellFocus();
                int rowToDelete = cell.rowIndex;

                if (rowToDelete <= 0)
                {
                    Application.SBO_Application.StatusBar.SetText(
                        "Nije selektovan red za brisanje.",
                        BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                    return;
                }

                // 1️⃣ Sinhronizuj matricu sa data source-om
                oMatrixAvansi.FlushToDataSource();

                // 2️⃣ Ukloni zapis
                if (rowToDelete <= oDBDSAvansi.Size)
                    oDBDSAvansi.RemoveRecord(rowToDelete - 1);

                // 3️⃣ Osvježi matricu
                oMatrixAvansi.LoadFromDataSource();

                // 4️⃣ Ako nema više redova, dodaj prazan da forma ne padne
                if (oMatrixAvansi.RowCount == 0)
                {
                    oDBDSAvansi.InsertRecord(oDBDSAvansi.Size);
                    oMatrixAvansi.LoadFromDataSource();
                }

                Application.SBO_Application.StatusBar.SetText(
                    $"Red {rowToDelete} uspešno obrisan.",
                    BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    $"Greška pri brisanju reda: {ex.Message}",
                    BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }


        private void Button1_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = false;      // SPREČAVA SAP da uradi svoje otkazivanje

            try
            {
                oForm.Close();        // Ručno zatvaranje forme
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    $"Greška: {ex.Message}",
                    BoMessageTime.bmt_Short,
                    BoStatusBarMessageType.smt_Error);
            }
        }

        private void Button0_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                IzracunajSume();

                double razlika = double.Parse(EditText15.Value); // ✅ Item_38 – razlika
                if (Math.Abs(razlika) < 0.01)
                {
                    Application.SBO_Application.StatusBar.SetText("Nema razlike za plaćanje.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                    return;
                }

                Application.SBO_Application.StatusBar.SetText("Pokrećem kreiranje dokumenta za plaćanje...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_None);

                oCompany = Parametri.m_GetConnected_COMPANY();

                string errMsg;
                int errCode;
                CreatePayDIAPI(ref oCompany, ref oForm, out errMsg, out errCode);

                if (errCode == 0)
                    Application.SBO_Application.StatusBar.SetText("Dokument za plaćanje uspešno kreiran.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                else
                    Application.SBO_Application.MessageBox("Greška pri kreiranju dokumenta: " + errMsg);
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText($"Greška pri obračunu: {ex.Message}", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }


        /*public void CreatePayDIAPI(ref SAPbobsCOM.Company oCompany, ref SAPbouiCOM.Form oForm, out string sErr, out int iErr)
        {
            iErr = 0;
            sErr = "";

            try
            {
                // 🔹 1️⃣ Učitavanje podataka sa forme
                double sumaAvansa = double.Parse(((EditText)oForm.Items.Item("Item_36").Specific).Value);
                double sumaTroskova = double.Parse(((EditText)oForm.Items.Item("Item_37").Specific).Value);
                double razlika = double.Parse(((EditText)oForm.Items.Item("Item_38").Specific).Value);
                string empCode = ((EditText)oForm.Items.Item("Item_3").Specific).Value; // Code iz OHEM
                string repId = ((EditText)oForm.Items.Item("Item_1").Specific).Value; //RepID iz Header-a
                string valuta = ((EditText)oForm.Items.Item("Item_15").Specific).Value;
                DateTime datum = DateTime.Now;

                if (Math.Abs(razlika) < 0.01)
                {
                    Application.SBO_Application.StatusBar.SetText("Nema razlike za plaćanje.",
                        BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                    return;
                }

                // 🔹 2️⃣ Prevod empCode → empID i ime zaposlenog
                int empID = 0;
                string empName = "";

                if (!string.IsNullOrEmpty(empCode))
                {
                    SAPbobsCOM.Recordset oRec =
                        (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                    oRec.DoQuery($@"SELECT ""empID"", ""firstName"" || ' ' || ""lastName"" AS ""FullName""
                            FROM ""OHEM""
                            WHERE ""Code"" = '{empCode}'");

                    if (!oRec.EoF)
                    {
                        empID = Convert.ToInt32(oRec.Fields.Item("empID").Value);
                        empName = oRec.Fields.Item("FullName").Value.ToString();
                    }
                }

                // 🔹 3️⃣ Određivanje tipa dokumenta
                bool isIncoming = razlika > 0; // ako je avans > trošak → uplata firmi
                string cashAccount = isIncoming ? "243000" : "243000"; // prilagodi konta

                SAPbobsCOM.Payments oPayment = (SAPbobsCOM.Payments)
                    oCompany.GetBusinessObject(isIncoming
                        ? SAPbobsCOM.BoObjectTypes.oIncomingPayments
                        : SAPbobsCOM.BoObjectTypes.oVendorPayments);

                oPayment.DocType = SAPbobsCOM.BoRcptTypes.rAccount;
                oPayment.DocCurrency = valuta;
                oPayment.DocDate = datum;
                oPayment.TaxDate = datum;
                oPayment.DueDate = datum;
                oPayment.CashAccount = cashAccount;
                oPayment.CashSum = Math.Abs(razlika);
                oPayment.Remarks = isIncoming
                    ? "Uplata zaposlenog po obračunu troškova (Concur Claim)"
                    : "Isplata zaposlenom po obračunu troškova (Concur Claim)";


                // 🔹 4️⃣ Dodavanje ReportID u zaglavlje OVPM dokumenta
                if (!string.IsNullOrEmpty(repId))
                    oPayment.UserFields.Fields.Item("U_REPORTID").Value = repId;

                // 🔹 4️⃣ Dodavanje reda (AccountPayments)
                oPayment.AccountPayments.AccountCode = "221101"; // konto troškova
                oPayment.AccountPayments.SumPaid = Math.Abs(razlika);
                oPayment.AccountPayments.Decription = "Refundacija troškova zaposlenog";

                if (empID > 0)
                {
                    // postavi UDF polja u AccountPayments redu
                    oPayment.AccountPayments.UserFields.Fields.Item("U_EMPLOYEE").Value = empID.ToString();
                    oPayment.AccountPayments.UserFields.Fields.Item("U_EPLOYEENAME").Value = empName;
                }

                oPayment.AccountPayments.Add();

                // 🔹 5️⃣ Dodavanje dokumenta
                int lRetCode = oPayment.Add();
                if (lRetCode != 0)
                {
                    oCompany.GetLastError(out iErr, out sErr);
                    Application.SBO_Application.StatusBar.SetText($"Greška pri kreiranju plaćanja: {sErr}",
                        BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    string newDocEntryStr;
                    oCompany.GetNewObjectCode(out newDocEntryStr);

                    Application.SBO_Application.StatusBar.SetText(
                        $"{(isIncoming ? "Uplata" : "Isplata")} uspešno kreirana (DocEntry: {newDocEntryStr})",
                        BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);

                    // 🔹 6️⃣ Upis nazad u UDO header
                    try
                    {
                        DBDataSource dsHeader = oForm.DataSources.DBDataSources.Item("@BO_CONCUR_CLAIM_H");
                        dsHeader.SetValue("U_PaymentEntry", 0, newDocEntryStr);
                        oForm.Update();
                    }
                    catch (Exception exInner)
                    {
                        Application.SBO_Application.StatusBar.SetText(
                            "Nije upisan PaymentEntry u UDO. " + exInner.Message,
                            BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                    }
                }
            }
            catch (Exception ex)
            {
                sErr = ex.Message;
                iErr = -1;
                Application.SBO_Application.StatusBar.SetText("Greška u CreatePayDIAPI: " + ex.Message,
                    BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }
        */
        //private CheckBox CheckBox1;

        public void CreatePayDIAPI(ref SAPbobsCOM.Company oCompany, ref SAPbouiCOM.Form oForm, out string sErr, out int iErr)
        {
            iErr = 0;
            sErr = "";

            try
            {
                // 🔹 1️⃣ Učitavanje podataka sa forme
                double sumaAvansa = double.Parse(((EditText)oForm.Items.Item("Item_36").Specific).Value);
                double sumaTroskova = double.Parse(((EditText)oForm.Items.Item("Item_37").Specific).Value);
                double razlika = double.Parse(((EditText)oForm.Items.Item("Item_38").Specific).Value);
                string empCode = ((EditText)oForm.Items.Item("Item_3").Specific).Value;
                string repId = ((EditText)oForm.Items.Item("Item_1").Specific).Value;  // 🔥 Report ID
                string valuta = ((EditText)oForm.Items.Item("Item_15").Specific).Value;
                DateTime datum = DateTime.Now;

                if (Math.Abs(razlika) < 0.01)
                {
                    Application.SBO_Application.StatusBar.SetText("Nema razlike za plaćanje.",
                        BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                    return;
                }

                // 🔹 2️⃣ Prevod empCode → empID i ime zaposlenog
                int empID = 0;
                string empName = "";

                if (!string.IsNullOrEmpty(empCode))
                {
                    SAPbobsCOM.Recordset oRec =
                        (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                    oRec.DoQuery($@"SELECT ""empID"", ""firstName"" || ' ' || ""lastName"" AS ""FullName""
                            FROM ""OHEM""
                            WHERE ""Code"" = '{empCode}'");

                    if (!oRec.EoF)
                    {
                        empID = Convert.ToInt32(oRec.Fields.Item("empID").Value);
                        empName = oRec.Fields.Item("FullName").Value.ToString();
                    }
                }

                // 🔹 3️⃣ Određivanje tipa dokumenta
                bool isIncoming = razlika > 0;
                string cashAccount = "243000"; // prilagodi banci

                SAPbobsCOM.Payments oPayment = (SAPbobsCOM.Payments)
                    oCompany.GetBusinessObject(isIncoming
                        ? SAPbobsCOM.BoObjectTypes.oIncomingPayments
                        : SAPbobsCOM.BoObjectTypes.oVendorPayments);

                oPayment.DocType = SAPbobsCOM.BoRcptTypes.rAccount;
                oPayment.DocCurrency = valuta;
                oPayment.DocDate = datum;
                oPayment.TaxDate = datum;
                oPayment.DueDate = datum;
                oPayment.CashAccount = cashAccount;
                oPayment.CashSum = Math.Abs(razlika);
                oPayment.Remarks = isIncoming
                    ? "Uplata zaposlenog po obračunu troškova (Concur Claim)"
                    : "Isplata zaposlenom po obračunu troškova (Concur Claim)";

                // 🔹 4️⃣ Dodavanje ReportID u zaglavlje OVPM dokumenta
                if (!string.IsNullOrEmpty(repId))
                    oPayment.UserFields.Fields.Item("U_REPORTID").Value = repId;

                // 🔹 5️⃣ Dodavanje reda AccountPayments
                oPayment.AccountPayments.AccountCode = "221101";
                oPayment.AccountPayments.SumPaid = Math.Abs(razlika);
                oPayment.AccountPayments.Decription = "Refundacija troškova zaposlenog";

                if (empID > 0)
                {
                    oPayment.AccountPayments.UserFields.Fields.Item("U_EMPLOYEE").Value = empID.ToString();
                    oPayment.AccountPayments.UserFields.Fields.Item("U_EPLOYEENAME").Value = empName;
                }

                oPayment.AccountPayments.Add();

                // 🔹 6️⃣ Dodavanje dokumenta
                int lRetCode = oPayment.Add();
                if (lRetCode != 0)
                {
                    oCompany.GetLastError(out iErr, out sErr);
                    Application.SBO_Application.StatusBar.SetText($"Greška pri kreiranju plaćanja: {sErr}",
                        BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    string newDocEntryStr;
                    oCompany.GetNewObjectCode(out newDocEntryStr);

                    Application.SBO_Application.StatusBar.SetText(
                        $"{(isIncoming ? "Uplata" : "Isplata")} uspešno kreirana (DocEntry: {newDocEntryStr})",
                        BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);

                    try
                    {
                        // Preuzmi novokreirani OVPM (DocEntry → DocNum, Total, Comments)
                        SAPbobsCOM.Recordset rsPay =
                            (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                        rsPay.DoQuery($@"
            SELECT 
                ""DocNum"", 
                ""DocTotal"", 
                ""Comments"",
                ""U_REPORTID""
            FROM ""OVPM""
            WHERE ""DocEntry"" = {newDocEntryStr}
        ");

                        if (!rsPay.EoF)
                        {
                            string newDocNum = rsPay.Fields.Item("DocNum").Value.ToString();
                            string newTotal = rsPay.Fields.Item("DocTotal").Value.ToString();
                            string newComment = rsPay.Fields.Item("Comments").Value.ToString();
                            string existingRepID = rsPay.Fields.Item("U_REPORTID").Value.ToString().Trim();

                            // 🔥 Uzmemo ReportID iz UDO headera
                            string repIdUDO = ((EditText)oForm.Items.Item("Item_1").Specific).Value?.Trim();

                            // ============================================================
                            // 1️⃣ Provera da OP već ima ReportID (koristi se negde drugde)
                            // ============================================================
                            if (!string.IsNullOrEmpty(existingRepID) && existingRepID != repIdUDO)
                            {
                                Application.SBO_Application.MessageBox(
                                    $"Avans (OP {newDocNum}) je već korišćen u drugom obračunu!\n" +
                                    $"ReportID: {existingRepID}"
                                );
                                return;
                            }

                            // ============================================================
                            // 2️⃣ Upisujemo ReportID u OVPM header — U_REPORTID
                            // ============================================================
                            SAPbobsCOM.Payments updPay =
                                (SAPbobsCOM.Payments)oCompany.GetBusinessObject(
                                    isIncoming ? SAPbobsCOM.BoObjectTypes.oIncomingPayments
                                               : SAPbobsCOM.BoObjectTypes.oVendorPayments
                                );

                            if (updPay.GetByKey(int.Parse(newDocEntryStr)))
                            {
                                updPay.UserFields.Fields.Item("U_REPORTID").Value = repIdUDO;

                                int ret = updPay.Update();
                                if (ret != 0)
                                {
                                    oCompany.GetLastError(out int errC, out string errT);
                                    Application.SBO_Application.StatusBar.SetText(
                                        "Greška pri upisu ReportID u Payment: " + errT,
                                        BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                }
                            }

                            // ============================================================
                            // 3️⃣ Ubacivanje novog OP u UDO child tabelu avansa
                            // ============================================================
                            DBDataSource dsOP = oForm.DataSources.DBDataSources.Item("@BO_CONCUR_CLAIM_OP");

                            int newRow = dsOP.Size;
                            dsOP.InsertRecord(newRow);

                            dsOP.SetValue("U_BrojDokumenta", newRow, newDocNum);
                            dsOP.SetValue("U_Opis", newRow, newComment);
                            dsOP.SetValue("U_Iznos", newRow, newTotal);

                            oMatrixAvansi.LoadFromDataSource();
                            IzracunajSume();
                        }
                    }
                    catch (Exception exPay)
                    {
                        Application.SBO_Application.StatusBar.SetText(
                            "Greška pri dodavanju dokumenta u Avanse: " + exPay.Message,
                            BoMessageTime.bmt_Short,
                            BoStatusBarMessageType.smt_Error);
                    }


                    // ============================================================
                    // 🔥 4️⃣ Upis nazad u UDO header (PaymentEntry + Zatvoreno)
                    // ============================================================
                    try
                    {
                        DBDataSource dsHeader = oForm.DataSources.DBDataSources.Item("@BO_CONCUR_CLAIM_H");

                        // Payment Entry
                        dsHeader.SetValue("U_PaymentEntry", 0, newDocEntryStr);

                        // Zatvaranje
                        string zatvorenoVal = ComboBox0.Selected == null ? "N" : ComboBox0.Selected.Value;
                        dsHeader.SetValue("U_Zatvoreno", 0, zatvorenoVal);

                        oForm.Update();

                        if (zatvorenoVal == "Y")
                        {
                            Application.SBO_Application.StatusBar.SetText(
                                "Dokument je automatski zaključen.",
                                BoMessageTime.bmt_Short,
                                BoStatusBarMessageType.smt_Warning
                            );

                            DisableEditing();
                        }
                    }
                    catch (Exception exInner)
                    {
                        Application.SBO_Application.StatusBar.SetText(
                            "Nije upisano PaymentEntry/Zatvoreno u UDO: " + exInner.Message,
                            BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                    }
                }

            }
            catch (Exception ex)
            {
                sErr = ex.Message;
                iErr = -1;
                Application.SBO_Application.StatusBar.SetText("Greška u CreatePayDIAPI: " + ex.Message,
                    BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        private Button Button4;
        private ComboBox ComboBox0;
        private StaticText StaticText3;
    }
}
using SAPbouiCOM;
using SAPbouiCOM.Framework;
using System;
using Application = SAPbouiCOM.Framework.Application;

namespace CapgeminiConcur
{
    [FormAttribute("CapgeminiConcur.Form1", "Form1.b1f")]
    class Form1 : UserFormBase
    {
        private Form oForm;
        private Matrix oMatrixAvansi;
        private DBDataSource oDBDSAvansi;
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
            // U OnInitializeComponent:
            this.Button1.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button1_ClickAfter);
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
            //this.Button3.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button3_ClickAfter);
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

                Matrix0 = (Matrix)oForm.Items.Item("Item_18").Specific;
                Matrix1 = (Matrix)oForm.Items.Item("Item_43").Specific;

                oMatrixAvansi.Columns.Item("BrDok").Editable = true;
                oMatrixAvansi.Columns.Item("Opis").Editable = true;
                oMatrixAvansi.Columns.Item("Iznos").Editable = true;

                try
                {
                    AddCFLForOutgoingPayments();
                }
                catch (Exception ex)
                {
                    Application.SBO_Application.StatusBar.SetText(
                        $"Greška pri kreiranju CFL-a: {ex.Message}",
                        BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText($"Greška pri inicijalizaciji forme: {ex.Message}",
                    BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// Kreira CFL za Outgoing Payments (OVPM) i vezuje ga za kolonu BrDok
        private void AddCFLForOutgoingPayments()
        {
            try
            {
                const string cflUID = "CFL_OP";
                ChooseFromList oCFL = null;

                // 🔹 Provera da li već postoji CFL
                try { oCFL = oForm.ChooseFromLists.Item(cflUID); }
                catch { oCFL = null; }

                if (oCFL == null)
                {
                    ChooseFromListCreationParams oCFLCreationParams =
                        (ChooseFromListCreationParams)Application.SBO_Application.CreateObject(
                            BoCreatableObjectType.cot_ChooseFromListCreationParams);

                    oCFLCreationParams.MultiSelection = false;
                    oCFLCreationParams.ObjectType = "46"; // Outgoing Payments
                    oCFLCreationParams.UniqueID = cflUID;

                    oCFL = oForm.ChooseFromLists.Add(oCFLCreationParams);

                    /*Application.SBO_Application.StatusBar.SetText(
                        "CFL_OP uspešno kreiran.",
                        BoMessageTime.bmt_Short,
                        BoStatusBarMessageType.smt_Success);*/
                }

                // 🔹 Vezivanje CFL-a na kolonu BrDok
                Column colBrDok = ((Matrix)oForm.Items.Item("Item_18").Specific).Columns.Item("BrDok");
                colBrDok.ChooseFromListUID = cflUID;
                colBrDok.ChooseFromListAlias = "DocNum";

                // 🔸 Napomena:
                // SAP B1 ne dozvoljava direktnu kontrolu prikaza kolona CFL prozora kroz SDK,
                // pa kolone koje vidiš u “List of Outgoing Payments” određuje sam sistem
                // na osnovu objekta (OVPM). Ako želiš drugačiji prikaz, mora se kreirati
                // zaseban custom CFL (user query CFL).
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    $"Greška pri kreiranju CFL-a: {ex.Message}",
                    BoMessageTime.bmt_Short,
                    BoStatusBarMessageType.smt_Error);
            }
        }

        private void SBO_Application_ItemEvent(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            // 🔹 Provera da li forma postoji
            SAPbouiCOM.Form testForm = null;
            try
            {
                testForm = Application.SBO_Application.Forms.Item(FormUID);
            }
            catch
            {
                // Forma više ne postoji
                BubbleEvent = false;
                return;
            }

            if (pVal.FormTypeEx != "CapgeminiConcur.Form1") return;

            try
            {
                oForm = Application.SBO_Application.Forms.Item(FormUID);
                oDBDSAvansi = oForm.DataSources.DBDataSources.Item("@BO_CONCUR_CLAIM_OP");

                // === 1️⃣ CFL logika za kolonu BrDok ===
                if (pVal.EventType == BoEventTypes.et_CHOOSE_FROM_LIST &&
                pVal.ItemUID == "Item_18" && pVal.ColUID == "BrDok")
                {
                    if (pVal.BeforeAction)
                    {
                        try
                        {
                            string employeeCode = ((EditText)oForm.Items.Item("Item_3").Specific).Value;
                            if (string.IsNullOrEmpty(employeeCode))
                            {
                                Application.SBO_Application.StatusBar.SetText(
                                    "EmployeeCode nije unet – ne može se filtrirati lista.",
                                    BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                                return;
                            }

                            oCompany = Parametri.m_GetConnected_COMPANY();
                            SAPbobsCOM.Recordset oRec = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            SAPbobsCOM.Recordset oDocs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                            try
                            {
                                // 1️⃣ Nađi empID
                                oRec.DoQuery($@"SELECT ""empID"" FROM ""OHEM"" WHERE ""Code"" = '{employeeCode}'");
                                if (oRec.EoF)
                                {
                                    Application.SBO_Application.StatusBar.SetText($"Nepostojeći zaposleni sa Code = {employeeCode}",
                                        BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                                    return;
                                }

                                int employeeID = Convert.ToInt32(oRec.Fields.Item("empID").Value);

                                // 2️⃣ Nađi sve OP dokumente tipa Account vezane za tog zaposlenog
                                string query = $@"
                                SELECT DISTINCT T0.""DocNum""
                                FROM ""OVPM"" T0
                                INNER JOIN ""VPM4"" T1 ON T0.""DocEntry"" = T1.""DocNum""
                                WHERE T0.""DocType"" = 'A' 
                                AND T1.""U_EMPLOYEE"" = {employeeID}";                

                                oDocs.DoQuery(query);

                                if (oDocs.EoF)
                                {
                                    Application.SBO_Application.StatusBar.SetText(
                                        "Nema Account plaćanja za ovog zaposlenog.",
                                        BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                                    return;
                                }

                                // 3️⃣ Kreiraj uslov za sve pronađene DocNum vrednosti
                                ChooseFromList oCFL = oForm.ChooseFromLists.Item("CFL_OP");
                                Conditions oConds = new Conditions();

                                bool first = true;
                                while (!oDocs.EoF)
                                {
                                    string docNum = oDocs.Fields.Item("DocNum").Value.ToString();
                                    Condition cond = oConds.Add();
                                    cond.Alias = "DocNum";
                                    cond.Operation = BoConditionOperation.co_EQUAL;
                                    cond.CondVal = docNum;

                                    if (!first)
                                        cond.Relationship = BoConditionRelationship.cr_OR;
                                    first = false;

                                    oDocs.MoveNext();
                                }

                                oCFL.SetConditions(oConds);

                                Application.SBO_Application.StatusBar.SetText(
                                    $"CFL filtriran – prikazuje samo OP tipa Account za empID={employeeID}",
                                    BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                            }
                            finally
                            {
                                try { System.Runtime.InteropServices.Marshal.ReleaseComObject(oRec); } catch { }
                                try { System.Runtime.InteropServices.Marshal.ReleaseComObject(oDocs); } catch { }
                            }
                        }
                        catch (Exception ex)
                        {
                            Application.SBO_Application.StatusBar.SetText(
                                $"Greška pri filtriranju CFL-a: {ex.Message}",
                                BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                        }
                    }
                    else
                    {
                        try
                        {
                            IChooseFromListEvent cflEvent = (IChooseFromListEvent)pVal;
                            DataTable dt = cflEvent.SelectedObjects;

                            if (dt != null && dt.Rows.Count > 0)
                            {
                                int row = pVal.Row;
                                oMatrixAvansi.FlushToDataSource();

                                // 🔹 Čitanje vrednosti iz izabranog reda
                                string docNum = dt.GetValue("DocNum", 0).ToString();

                                // 🔍 Proveri da li je isti avans već izabran u drugom redu
                                for (int i = 1; i <= oMatrixAvansi.RowCount; i++)
                                {
                                    string existing = ((EditText)oMatrixAvansi.Columns.Item("BrDok").Cells.Item(i).Specific).Value;
                                    if (!string.IsNullOrEmpty(existing) && existing == docNum && i != row)
                                    {
                                        Application.SBO_Application.StatusBar.SetText(
                                            $"Avans {docNum} je već odabran u redu {i}.",
                                            BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                                        return;
                                    }
                                }

                                string opis = dt.GetValue("JrnlMemo", 0).ToString();
                                string amount = dt.GetValue("DocTotal", 0).ToString();

                                // 🔹 Upis u data source (UDO liniju)
                                oDBDSAvansi.SetValue("U_BrojDokumenta", row - 1, docNum);
                                oDBDSAvansi.SetValue("U_Opis", row - 1, opis);
                                oDBDSAvansi.SetValue("U_Iznos", row - 1, amount);

                                oMatrixAvansi.LoadFromDataSource();

                                // 🔹 Ponovno izračunavanje suma
                                IzracunajSume();

                                Application.SBO_Application.StatusBar.SetText(
                                    $"Odabran dokument: {docNum} | {opis} | {amount}",
                                    BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                            }
                        }
                        catch (Exception ex)
                        {
                            Application.SBO_Application.StatusBar.SetText(
                                $"Greška pri obradi izbora iz CFL-a: {ex.Message}",
                                BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                        }
                    }


                }


                // === 2 Kad se učitaju podaci iz UDO-a ===
                if (pVal.EventType == BoEventTypes.et_FORM_DATA_LOAD && !pVal.BeforeAction)
                {
                    try
                    {
                        // Dohvati ime i prezime iz glavnog UDO objekta
                        DBDataSource dsHeader = oForm.DataSources.DBDataSources.Item("@BO_CONCUR_CLAIM");
                        string ime = dsHeader.GetValue("U_EmployeeName", 0).Trim();
                        string prezime = dsHeader.GetValue("U_EmployeeLastName", 0).Trim();

                        if (!string.IsNullOrEmpty(ime))
                        {
                            EditText imePrezime = (EditText)oForm.Items.Item("Item_3").Specific;
                            imePrezime.Value = ime;

                            if (!string.IsNullOrEmpty(prezime) && !imePrezime.Value.Contains(prezime))
                                imePrezime.Value += " " + prezime;
                        }

                        // Računanje suma odmah nakon učitavanja
                        IzracunajSume();
                    }
                    catch (Exception ex)
                    {
                        Application.SBO_Application.StatusBar.SetText(
                            $"Greška pri učitavanju imena/prezimena ili računanju suma: {ex.Message}",
                            BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    }
                }

                // === 3 Kad se promeni iznos u bilo kojoj matrici (AFTER) ===
                if (!pVal.BeforeAction &&
                    (pVal.EventType == BoEventTypes.et_VALIDATE || pVal.EventType == BoEventTypes.et_LOST_FOCUS))
                {
                    bool kolonaIznosAvans = pVal.ItemUID == "Item_18" && pVal.ColUID == "Iznos";
                    bool kolonaIznosTrosak = pVal.ItemUID == "Item_43" && pVal.ColUID == "Iznos1";

                    if (kolonaIznosAvans || kolonaIznosTrosak)
                    {
                        IzracunajSume();
                    }
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    $"Greška u ItemEvent: {ex.Message}",
                    BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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

                // --- Matrix za Avanse (Item_18, kolona "Iznos") ---
                for (int i = 1; i <= Matrix0.RowCount; i++)
                {
                    var cell = (SAPbouiCOM.EditText)Matrix0.Columns.Item("Iznos").Cells.Item(i).Specific;
                    string val = cell?.Value?.Trim() ?? "";
                    if (double.TryParse(val, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double iznos) ||
                        double.TryParse(val, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.CurrentCulture, out iznos))
                    {
                        sumaAvansa += iznos;
                    }
                }

                // --- Matrix za Troškove (Item_43, kolona "Iznos1") ---
                for (int i = 1; i <= Matrix1.RowCount; i++)
                {
                    var cell = (SAPbouiCOM.EditText)Matrix1.Columns.Item("Iznos1").Cells.Item(i).Specific;
                    string val = cell?.Value?.Trim() ?? "";
                    if (double.TryParse(val, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double iznos) ||
                        double.TryParse(val, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.CurrentCulture, out iznos))
                    {
                        sumaTroskova += iznos;
                    }
                }

                double razlika = sumaAvansa - sumaTroskova;

                // --- Upis vrednosti u polja na formi (Item_36, Item_37, Item_38) ---
                // Ovo je bezbedno jer su to EditText polja (Amount), a ne forsira se DB update forme.
                EditText13.Value = sumaAvansa.ToString("F2", System.Globalization.CultureInfo.InvariantCulture); // Item_36 – Suma avansa
                EditText14.Value = sumaTroskova.ToString("F2", System.Globalization.CultureInfo.InvariantCulture); // Item_37 – Suma troškova
                EditText15.Value = razlika.ToString("F2", System.Globalization.CultureInfo.InvariantCulture); // Item_38 – Razlika

                // --- (VAŽNO) NE raditi oForm.Update() ovde ---
                // Ako želiš da sume budu i u UDO headeru, odradi to na Save/OK događaju ili eksplicitnom dugmetu.
                // U suprotnom rizikuješ pad SAP-a tokom edit Validate faze.
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

            BubbleEvent = true; oMatrixAvansi.AddRow();
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

            /*BubbleEvent = true;
            CellPosition cell = oMatrixAvansi.GetCellFocus();
            if (cell.rowIndex > 0)
            {
                oMatrixAvansi.DeleteRow(cell.rowIndex);
                oMatrixAvansi.FlushToDataSource();
                oMatrixAvansi.LoadFromDataSource();
            }*/
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
            BubbleEvent = true; // Set BubbleEvent to true to allow the event to continue

            // Close the current form
            Application.SBO_Application.Forms.ActiveForm.Close();
        }

        private void Button1_ClickAfter(object sboObject, SBOItemEventArg pVal)
        {
            try
            {
                if (Application.SBO_Application.Forms.ActiveForm != null)
                {
                    Application.SBO_Application.Forms.ActiveForm.Close();
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    $"Greška pri zatvaranju forme: {ex.Message}",
                    BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                string empCode = ((EditText)oForm.Items.Item("Item_3").Specific).Value; // Code iz OHEM
                string valuta = "RSD";
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



    }
}
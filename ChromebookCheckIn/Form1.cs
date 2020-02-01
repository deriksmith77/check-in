using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using System.IO;
using Google.Apis.Util.Store;
using System.Threading;
using System.Data.OleDb;
using System.Data;
using System.Diagnostics;

namespace ChromebookCheckIn
{
    public partial class Form1 : Form
    {

        // Boolean checkDevice = true;
        // Boolean checkStats = true;

        public Form1()
        {
            InitializeComponent();
            TextBoxOrder.Add(txtInput, txtChromebook);
            txtInput.Tag = 1;
            txtChromebook.Tag = 2;
            txtCharger.Tag = 3;
            txtChromebook.KeyDown += BarcodeInputKeyDown;
            txtCharger.KeyDown += BarcodeInputKeyDown;
            txtInput.KeyDown += BarcodeInputKeyDown;
            txtChromebook.Leave += BarcodeInputLeave;
            txtCharger.Leave += BarcodeInputLeave;
            txtInput.Leave += BarcodeInputLeave;
        }

        #region ### Global Variables ###
        static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "CheckItIn";
        UserCredential credential;
        int index = 1;
        string ssCol = "";
        int chromebookUser = 0;
        string addnote = "";
        private Dictionary<Bunifu.Framework.UI.BunifuMetroTextbox, Bunifu.Framework.UI.BunifuMetroTextbox> TextBoxOrder = new Dictionary<Bunifu.Framework.UI.BunifuMetroTextbox, Bunifu.Framework.UI.BunifuMetroTextbox>();
        Boolean ChromebookID = false;
        Boolean ChargerID = false;
        String missing = "";
        int m = 0;
        Boolean missingMaster = false;
        Boolean missingChromebook = false;
        Boolean missingCharger = false;



        #endregion

        private void BarcodeInputKeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter && ActiveControl.GetType() == typeof(Bunifu.Framework.UI.BunifuMetroTextbox))
            {
                Bunifu.Framework.UI.BunifuMetroTextbox nextTextBox;
                if (TextBoxOrder.TryGetValue((Bunifu.Framework.UI.BunifuMetroTextbox)ActiveControl, out nextTextBox))
                {
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                    nextTextBox.Focus();
                }
            }
        }

        private void BarcodeInputLeave(object sender, EventArgs e)
        {
            if (sender.GetType() == typeof(Bunifu.Framework.UI.BunifuMetroTextbox))
            {
                Bunifu.Framework.UI.BunifuMetroTextbox textBox = (Bunifu.Framework.UI.BunifuMetroTextbox)sender;
                if (textBox.Tag.GetType() == typeof(int))
                {
                    BarcodeScanned(textBox.Text, (int)textBox.Tag);
                }
            }
        }

        private void BarcodeScanned(string barcode, int order)
        {
            if (txtInput.Text.Length < 15)
            {
                if (String.IsNullOrEmpty(txtChromebookHomeroom.Text))
                {
                    txtChromebook.Text = txtInput.Text;
                }
                
                txtInput.Text = "";
                txtInput.Focus();
            }
            else if (txtInput.Text.Length > 18)
            {
                if (String.IsNullOrEmpty(txtChargerHomeroom.Text))
                {
                    txtCharger.Text = txtInput.Text;
                }
                txtInput.Text = "";
                txtInput.Focus();
            }
            if (txtChromebook.Text.Length > 6 && txtChromebook.Text.Length < 16 && string.IsNullOrEmpty(txtChromebookName.Text))
            {
                ChromebookID = true;
                ReadAndSearch();
            }
            else
            {
                ChromebookID = false;
            }

            if (txtCharger.Text.Length > 18 && string.IsNullOrEmpty(txtChargerName.Text))
            {
                ChargerID = true;
                ReadAndSearch();
            }
            else
            {
                ChargerID = false;
            }
        }

        public void ReadAndSearch()
        {

            using (var stream =
                new FileStream("client_secret_675941396623-spedsb2u5hrsplch6c6gp1t8v1tqeltk.apps.googleusercontent.com.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;

            }



            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
            // read the data
            String spreadsheetId = "1tql5eJCaCZfaMGoFo01M4n-p1k6BT7LSrqyOuHdDOwY";
            String range = "Sheet3!A2:I";
            SpreadsheetsResource.ValuesResource.GetRequest request =
                    service.Spreadsheets.Values.Get(spreadsheetId, range);

            // Read the Google Sheets Database and compare Bar Code to EVERY STUDENT DEVICE
            // Once you find a match, run WriteIt(), log that the device was turned in
            ValueRange response = request.Execute();
            IList<IList<Object>> values = response.Values;



            #region ### Read the Chromebook ID ###
            if (values != null && values.Count > 0 && ChromebookID == true)
            {
                foreach (var row in values)
                {
                    index++; // count to current index (name) / ignores rows headings
                    chromebookUser = index; // gets the chromebook user for the unmatching log
                    try
                    {
                        // Row 2 is Chromebook ID in Database
                        if (String.Equals(row[4].ToString(), txtChromebook.Text, StringComparison.OrdinalIgnoreCase))
                        {

                            ssCol = "G";
                            WriteIt();
                            txtChromebookName.Text = row[1].ToString();
                            txtChromebookHomeroom.Text = (row[0].ToString());
                            CheckMatchingNames();
                            break;
                        }
                    }
                    catch
                    {
                        // nothing
                    }
                }
                ChromebookID = false;
            }

            ValueRange response2 = request.Execute();
            IList<IList<Object>> values2 = response.Values;
            #endregion

            #region ### Read the Charger ID ###
            if (values2 != null && values2.Count > 0 && ChargerID == true)
            {
                // Read the Chromebook ID
                foreach (var row in values2)
                {
                    index++; // count to current index (name) / ignores rows headings
                    try
                    {
                        // Row 2 is Chromebook ID in Database
                        if (String.Equals(row[5].ToString(), txtCharger.Text, StringComparison.OrdinalIgnoreCase))
                        {
                            ssCol = "H";
                            WriteIt();
                            txtChargerName.Text = row[1].ToString();
                            txtChargerHomeroom.Text = (row[0].ToString());
                            CheckMatchingNames();
                            break;
                        }
                    }
                    catch
                    {
                        // nothing
                    }
                }
                ChargerID = false;
            }
            #endregion
        }
        public void WriteIt()
        {
            // write the data (f5)
            String spreadsheetId2 = "1tql5eJCaCZfaMGoFo01M4n-p1k6BT7LSrqyOuHdDOwY";
            //String range2 = "Sheet2!" + ssCol + index;  // update cell F5 
            String range2 = "Sheet3!" + ssCol + index;  // update cell F5 
            ValueRange valueRange = new ValueRange();
            valueRange.MajorDimension = "COLUMNS";//"ROWS";//COLUMNS


            var oblist = new List<object>() { "Turned In - " + DateTime.Now };
            valueRange.Values = new List<IList<object>> { oblist };

            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            SpreadsheetsResource.ValuesResource.UpdateRequest update = service.Spreadsheets.Values.Update(valueRange, spreadsheetId2, range2);
            update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            UpdateValuesResponse result2 = update.Execute();

            index = 1;
        }
        public void WriteUnMatchLog()
        {
            String spreadsheetId2 = "1tql5eJCaCZfaMGoFo01M4n-p1k6BT7LSrqyOuHdDOwY";
            String range2 = "Sheet3!I" + chromebookUser;
            ValueRange valueRange = new ValueRange();
            valueRange.MajorDimension = "COLUMNS";//"ROWS";//COLUMNS


            var oblist = new List<object>() { "Turned in Other Kids's charger" };
            valueRange.Values = new List<IList<object>> { oblist };

            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            SpreadsheetsResource.ValuesResource.UpdateRequest update = service.Spreadsheets.Values.Update(valueRange, spreadsheetId2, range2);
            update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            UpdateValuesResponse result2 = update.Execute();
            index = 1;
        }
        public void AddDamageNote()
        {
            // write the data (f5)
            String spreadsheetId2 = "1tql5eJCaCZfaMGoFo01M4n-p1k6BT7LSrqyOuHdDOwY";
            //String range2 = "Sheet2!G" + chromebookUser;
            String range2 = "Sheet3!I" + chromebookUser;
            ValueRange valueRange = new ValueRange();
            valueRange.MajorDimension = "COLUMNS";//"ROWS";//COLUMNS


            var oblist = new List<object>() { addnote };
            valueRange.Values = new List<IList<object>> { oblist };

            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            SpreadsheetsResource.ValuesResource.UpdateRequest update = service.Spreadsheets.Values.Update(valueRange, spreadsheetId2, range2);
            update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            UpdateValuesResponse result2 = update.Execute();
            index = 1;
        }

        // reset on ESC Key Press
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Escape)
            {
                txtCharger.Text = "";
                txtChromebook.Text = "";
                txtChromebookName.Text = "";
                txtChromebookHomeroom.Text = "";
                txtChargerHomeroom.Text = "";
                txtChargerName.Text = "";
                txtInput.Focus();
                return true;    // indicate that you handled this keystroke
            }

            // Call the base class
            return base.ProcessCmdKey(ref msg, keyData);
        }

        private void CheckMatchingNames()
        {
            if (txtChromebookName.Text != "" && txtChargerName.Text != "")
            {
                if (txtChromebookName.Text != txtChargerName.Text)
                {
                    WriteUnMatchLog();
                }
            }
        }

        private void BtnNote_Click(object sender, EventArgs e)
        {
            //AddDamageNote();
            //addnote = "";
        }

        private void TxtInput_TextChanged(object sender, EventArgs e)
        {
            if (txtInput.Text != "")
            {
                addnote = txtInput.Text;
            }

        }



        private void BunifuImageButton2_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void BtnMnu_Click(object sender, EventArgs e)
        {
            if (sideMenu.Width == 50)
            {
                sideMenu.Visible = false;
                sideMenu.Width = 259;
                PanelAnimator.ShowSync(sideMenu);
                LogoAnimator.ShowSync(logo);
            }
            else
            {
                LogoAnimator.Hide(logo);
                sideMenu.Visible = true;
                sideMenu.Width = 50;
                PanelAnimator.ShowSync(sideMenu);
            }
        }

        private void BtnExit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void BtnCheckDevices_Click(object sender, EventArgs e)
        {
            panCheckStats.Visible = false;
            panCheckStats.SendToBack();
        }

        private void BtnCheckStats_Click(object sender, EventArgs e)
        {
            panCheckStats.Visible = true;
            panCheckStats.BringToFront();
            panCheckStatus.Visible = false;
            panCheckStatus.SendToBack();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            panCheckStats.Visible = false;
            panCheckStats.SendToBack();
            txtInput.Tag = 1;
            txtInput.TabIndex = 1;
            txtInput.Focus();
            txtInput.Select();
        }

        private void BtnCheckStatus_Click(object sender, EventArgs e)
        {
            panCheckStats.Visible = true;
            panCheckStats.BringToFront();
            panCheckStatus.Visible = true;
            panCheckStatus.BringToFront();
        }

        private void BtnCheckIDStatus_Click(object sender, EventArgs e)
        {
            // txtChromebookCheck.Text = "Missing";
            // txtChargerCheck.Text = "Turned In";
        }

        private void BtnInactive_Click(object sender, EventArgs e)
        {
            missingMaster = true;
            missing = "First Name,Last Name,Student ID,Homeroom Teacher,Chromebook,Charger" + Environment.NewLine;
            using (var stream =
                 new FileStream("client_secret_675941396623-spedsb2u5hrsplch6c6gp1t8v1tqeltk.apps.googleusercontent.com.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;

            }

            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
            // read the data
            String spreadsheetId3 = "1tql5eJCaCZfaMGoFo01M4n-p1k6BT7LSrqyOuHdDOwY";
            String range3 = "Sheet3!A2:J";
            SpreadsheetsResource.ValuesResource.GetRequest request =
                    service.Spreadsheets.Values.Get(spreadsheetId3, range3);

            ValueRange response3 = request.Execute();
            IList<IList<Object>> values = response3.Values;

            foreach (var row in values)
            {
                try
                {
                    m++;
                    if ((row[6].ToString() == null || row[6].ToString() == "") && (row[7].ToString() == null || row[7].ToString() == ""))
                    {
                        missing += row[1] + "," + row[2] + "," + row[0] + ",m" + ",m" + Environment.NewLine;
                    }
                    else if (row[6].ToString() == null || row[6].ToString() == "")
                    {
                        missing += row[1] + "," + row[2] + "," + row[0] + ",m" + Environment.NewLine;
                    }
                    else if (row[7].ToString() == null || row[7].ToString() == "")
                    {
                        missing += row[1] + "," + row[2] + "," + row[0] + "," + ",m" + Environment.NewLine;
                    }

                }
                catch
                {
                    // nothing
                }
            }

            TransferToDataGrid();
        }

        public void TransferToDataGrid()
        {
            String path = "";
            if (missingMaster == true) { path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "/MissingDevices.csv"; }
            else if (missingChromebook == true) { path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "/MissingChromebooks.csv"; }
            else if (missingCharger == true) { path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "/MissingChargers.csv"; }
            try
            {
                File.WriteAllText(path, missing);
            }
            catch
            {
                MessageBox.Show("You cannot create a new missing device list while the current Excel file is open.\nPlease close Excel before generating a new list.", "Close Excel", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            
            Process.Start(path);
            missingMaster = false;
            missingChromebook = false;
            missingCharger = false;
        }

        private void BtnMissingChromebooks_Click(object sender, EventArgs e)
        {
            missingChromebook = true;
            missing = "First Name,Last Name,Student ID,Homeroom Teacher,Chromebook" + Environment.NewLine;
            using (var stream =
                 new FileStream("client_secret_675941396623-spedsb2u5hrsplch6c6gp1t8v1tqeltk.apps.googleusercontent.com.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;

            }

            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
            // read the data
            String spreadsheetId3 = "1tql5eJCaCZfaMGoFo01M4n-p1k6BT7LSrqyOuHdDOwY";
            String range3 = "Sheet3!A2:J";
            SpreadsheetsResource.ValuesResource.GetRequest request =
                    service.Spreadsheets.Values.Get(spreadsheetId3, range3);

            ValueRange response3 = request.Execute();
            IList<IList<Object>> values = response3.Values;

            foreach (var row in values)
            {
                try
                {
                    m++;
                    if (row[6].ToString() == null || row[6].ToString() == "")
                    {
                        missing += row[1] + "," + row[2] + "," + row[0] + ",m" + Environment.NewLine;
                    }
                }
                catch
                {
                    // nothing
                }
            }

            TransferToDataGrid();
        }

        private void BtnMissingChargers_Click(object sender, EventArgs e)
        {
            missingCharger = true;
            missing = "First Name,Last Name,Student ID,Homeroom Teacher,Charger" + Environment.NewLine;
            using (var stream =
                 new FileStream("client_secret_675941396623-spedsb2u5hrsplch6c6gp1t8v1tqeltk.apps.googleusercontent.com.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;

            }

            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
            // read the data
            String spreadsheetId3 = "1tql5eJCaCZfaMGoFo01M4n-p1k6BT7LSrqyOuHdDOwY";
            String range3 = "Sheet3!A2:J";
            SpreadsheetsResource.ValuesResource.GetRequest request =
                    service.Spreadsheets.Values.Get(spreadsheetId3, range3);

            ValueRange response3 = request.Execute();
            IList<IList<Object>> values = response3.Values;

            foreach (var row in values)
            {
                try
                {
                    m++;
                    if (row[7].ToString() == null || row[7].ToString() == "")
                    {
                        missing += row[1] + "," + row[2] + "," + row[0] + ",m" + Environment.NewLine;
                    }

                }
                catch
                {
                    // nothing
                }
            }

            TransferToDataGrid();
        }
    }
}
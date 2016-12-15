using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using MetroFramework;
using System.Data.SqlClient;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Configuration;
using System.IO;
using System.Globalization;
using System.Xml;

namespace WindowsFormsApplication1
{
    public partial class MainForm : MetroFramework.Forms.MetroForm
    {
        //Define SQL Connection
        SqlConnection myConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["Mi_LAW-KCI"].ConnectionString);
        SqlConnection linksql2oracle = new SqlConnection(ConfigurationManager.ConnectionStrings["Link2Oracle"].ConnectionString);

        public MainForm()
        {
            InitializeComponent();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            this.ActiveControl = WO_Input;

            string xmlNode = "printers.xml";

            //Retrieve control's text from Resource
            metroButton1.Text = MiLAWKCI.Strings.es_MX.Resource1.Main_metroButton1;
            metroButton2.Text = MiLAWKCI.Strings.es_MX.Resource1.Main_metroButton2;
            Item_Label.Text = MiLAWKCI.Strings.es_MX.Resource1.Main_Item_Label;
            LabelQty_Label.Text = MiLAWKCI.Strings.es_MX.Resource1.Main_LabelQty_Label;
            PrinterLabel.Text = MiLAWKCI.Strings.es_MX.Resource1.Main_PrinterLabel;
            LS_Label.Text = MiLAWKCI.Strings.es_MX.Resource1.Main_LS_Label;
            BoxQty_Label.Text = MiLAWKCI.Strings.es_MX.Resource1.Main_BQ_Label;

            XmlReader xReader = XmlReader.Create(new StreamReader(xmlNode));
            while (xReader.Read())
            {
                switch (xReader.NodeType)
                {
                    case XmlNodeType.Text:
                        metroComboBox1.Items.Add(xReader.Value);
                        break;
                }
            }
        }

        private void metroButton1_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void metroButton2_Click(object sender, EventArgs e)
        {
            var MessageTitle = MiLAWKCI.Strings.es_MX.Resource1.Main_MessageBox_Title;
            var MessageInfo = MiLAWKCI.Strings.es_MX.Resource1.Main_MessageBox_Info;
            var AWScanned = LSTextBox.Text.ToString().ToUpper();

            //Validate if any field is null, if TRUE then show message box asking to fill out all the information
            if (WO_Input.Text == "" || LSTextBox.Text == "" || LabelQty_Input.Text == "" || metroComboBox1.SelectedItem == null)
            {
                MetroMessageBox.Show(this, Environment.NewLine + MessageInfo, MessageTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            else
            {
                string LWPRINTER = metroComboBox1.SelectedItem.ToString();
                string LWDROP = ConfigurationManager.AppSettings["LWDROP"];
                string printDate = string.Format("{0:yyyyMMdd-HHmmss}", DateTime.Now);
                string MfgDate = string.Format("{0:yyyy-MM-dd}", DateTime.Now);
                string csvfilename = LWDROP + "KCI-" + printDate + ".csv";

                linksql2oracle.Open();
                SqlCommand toBaan1 = new SqlCommand("DECLARE @sqlQuery VARCHAR(8000) DECLARE @finalQuery VARCHAR(8000) SET @sqlQuery = 'SELECT T1.T$CVAL, TRIM(T2.T$MITM) from BAAN.TFXSFC013034 T1, BAAN.TTISFC001034 T2 where T1.T$PDNO = T2.T$PDNO and T1.T$ATTR = 10 and T1.T$PDNO = ' + '''' + '''' + @varWO + '''' + '''' SET @finalQuery = 'SELECT * from OPENQUERY(AM3P1,'+''''+@sqlQuery+''''+')' EXEC (@finalQuery)", linksql2oracle);
                toBaan1.Parameters.AddWithValue("@varWO", WO_Input.Text.ToString().ToUpper());
                SqlDataReader toBaanReader = toBaan1.ExecuteReader();

                if ((toBaanReader.Read() == true))
                {
                    var resul = new Resul();
                    resul.attrib = toBaanReader[0].ToString().Trim();
                    resul.mitm = toBaanReader[1].ToString().Trim();
                    linksql2oracle.Close();

                    myConnection.Open();

                    //SQL Query. All columns where Item equals Item on DB
                    SqlCommand cmd = new SqlCommand("Select * from dbo.tblKCIPartNumbers where Item=@Item", myConnection);
                    cmd.Parameters.AddWithValue("@Item", resul.mitm.ToString().ToUpper().Trim());
                    SqlDataReader dr = cmd.ExecuteReader();

                    //If Query is successfull
                    if ((dr.Read() == true))
                    {
                        //Array on results for each column on the DB
                        var rSKID = resul.attrib.ToString().ToUpper().Trim();
                        var rMITM = resul.mitm.ToString().ToUpper().Trim();
                        var results = new Results();

                        results.Item = dr[1].ToString().Trim();
                        results.ShelfLife = dr[2].ToString().Trim();
                        results.PouchGTIN = dr[3].ToString().Trim();
                        results.CartonGTIN = dr[4].ToString().Trim();
                        results.PouchAW = dr[5].ToString().Trim();
                        results.CartonAW = dr[6].ToString().Trim();
                        results.ShipperAW = dr[7].ToString().Trim();
                        results.L1AW = dr[8].ToString().Trim();
                        results.L2AW = dr[9].ToString().Trim();
                        results.PouchLWL = dr[10].ToString().Trim();
                        results.CartonLWL = dr[11].ToString().Trim();
                        results.ShipperLWL = dr[12].ToString().Trim();
                        results.L1LWL = dr[13].ToString().Trim();
                        results.L2LWL = dr[14].ToString().Trim();
                        results.Origin = dr[15].ToString().Trim();
                        results.LblQty = LabelQty_Input.Text.ToString().Trim();
                        results.BoxQty = BoxQty_Input.Text.ToString().Trim();

                        myConnection.Close();

                        //Set Expiration Date to Now + Shelf Life. Set Expiration Day to last day of Calculated Month.
                        int dt;
                        int.TryParse(results.ShelfLife, out dt);
                        DateTime futureExp = DateTime.Now.AddMonths(dt);
                        int anio = futureExp.Year;
                        int mes = futureExp.Month;
                        var expLastDay = DateTime.DaysInMonth(anio, mes);
                        var ExpDate = futureExp.ToString("{0:yyyy-MM-dd}");
                        DateTime rEXPDATE = DateTime.ParseExact(ExpDate, "{0:yyyy-MM-dd}", CultureInfo.InvariantCulture);
                        var sEXPDATE = String.Format("{0:yyyy-MM-}" + expLastDay, rEXPDATE);

                        MetroMessageBox.Show(this, Environment.NewLine + results.PouchAW.ToString().ToUpper().Trim() + Environment.NewLine + "Favor de contactar a tu Supervisor y/o al equipo de IT Labeling.", "Error - Etiqueta no encontrada en DB", MessageBoxButtons.OK, MessageBoxIcon.Stop);

                        //Validate AW Input against PouchAW on DB
                        if (AWScanned == results.PouchAW.ToString().ToUpper().Trim())
                        {
                            if (results.PouchLWL == "" || results.PouchLWL.Substring(results.PouchLWL.Length - 4) != ".lwl")
                            {
                                MetroMessageBox.Show(this, Environment.NewLine + "No se encontro etiqueta dada de alta en Base de Datos de Mi-LAW." + Environment.NewLine + "Favor de contactar a tu Supervisor y/o al equipo de IT Labeling.", "Error - Etiqueta no encontrada en DB", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                            }

                            else
                            {
                                myConnection.Close();

                                TextWriter csvfile = new StreamWriter(csvfilename, true);
                                csvfile.WriteLine("PRTNUM,Format,LLMQTY,LOT_NUM,UMFG_DATE,UEXP_DATE,GTIN,ORIGEN");
                                csvfile.WriteLine('"' + LWPRINTER + '"' + "," + '"' + results.PouchLWL + '"' + "," + '"' + results.LblQty + '"' + "," + '"' + rSKID + '"' + ","
                                    + '"' + MfgDate + '"' + "," + '"' + sEXPDATE + '"' + "," + '"' + results.PouchGTIN + '"' + "," + '"' + results.Origin + '"');
                                csvfile.Close();

                                MetroMessageBox.Show(this, Environment.NewLine + "Archivo creado y depositado en Loftware. Las etiquetas deberan empezar a imprimirse y Mi-LAW cerrara automaticamente.", "Datos Validos", MessageBoxButtons.OK, MessageBoxIcon.Information);

                                myConnection.Open();
                                SqlCommand cmd3 = new SqlCommand("INSERT INTO dbo.tblKCIHistory (WorkOrder, Skid, Item, MfgDate, ExpDate, LblQty, BoxQty, LblAW, LblPath, LblGTIN, LblOrigin, Timestamp, iCodUser_ID, SysUser, MachineID, PrinterID) VALUES(@WO, @SKID, @Item, @MfgDate, @ExpDate, @LblQty, @BoxQty, @LblAW, @LblPath, @LblGTIN, @LblOrigin, @TimeStamp, @User, @SysUser, @Machine, @Printer)", myConnection);
                                cmd3.Parameters.AddWithValue("@WO", WO_Input.Text.ToString().ToUpper());
                                cmd3.Parameters.AddWithValue("@SKID", rSKID);
                                cmd3.Parameters.AddWithValue("@Item", resul.mitm);
                                cmd3.Parameters.AddWithValue("@MfgDate", MfgDate);
                                cmd3.Parameters.AddWithValue("@ExpDate", sEXPDATE);
                                cmd3.Parameters.AddWithValue("@LblQty", results.LblQty.ToString().Trim());
                                cmd3.Parameters.AddWithValue("@BoxQty", "");
                                cmd3.Parameters.AddWithValue("@LblAW", AWScanned);
                                cmd3.Parameters.AddWithValue("@LblPath", results.PouchLWL.ToString().Trim());
                                cmd3.Parameters.AddWithValue("@LblGTIN", results.PouchGTIN.ToString().Trim());
                                cmd3.Parameters.AddWithValue("@LblOrigin", results.Origin);
                                cmd3.Parameters.AddWithValue("@TimeStamp", printDate.ToString());
                                cmd3.Parameters.AddWithValue("@User", Login.UserN.ToString());
                                cmd3.Parameters.AddWithValue("@SysUser", System.Environment.UserName.ToString());
                                cmd3.Parameters.AddWithValue("@Machine", System.Environment.MachineName.ToString());
                                cmd3.Parameters.AddWithValue("@Printer", LWPRINTER);
                                cmd3.ExecuteNonQuery();
                                myConnection.Close();

                                Application.Exit();
                            }
                        }

                        //Validate AW Input against CartonAW on DB
                        else if (AWScanned == results.CartonAW.ToString().ToUpper().Trim())
                        {
                            if (results.CartonLWL == "" || results.CartonLWL.Substring(results.CartonLWL.Length - 4) != ".lwl")
                            {
                                MetroMessageBox.Show(this, Environment.NewLine + "No se encontro etiqueta dada de alta en Base de Datos de Mi-LAW." + Environment.NewLine + "Favor de contactar a tu Supervisor y/o al equipo de IT Labeling.", "Error - Etiqueta no encontrada en DB", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                            }

                            else
                            {
                                myConnection.Close();

                                TextWriter csvfile = new StreamWriter(csvfilename, true);
                                csvfile.WriteLine("PRTNUM,Format,LLMQTY,LOT_NUM,UMFG_DATE,UEXP_DATE,GTIN,ORIGEN");
                                csvfile.WriteLine('"' + LWPRINTER + '"' + "," + '"' + results.CartonLWL + '"' + "," + '"' + results.LblQty + '"' + "," + '"' + rSKID + '"' + ","
                                    + '"' + MfgDate + '"' + "," + '"' + sEXPDATE + '"' + "," + '"' + results.CartonGTIN + '"' + "," + '"' + results.Origin + '"');
                                csvfile.Close();

                                MetroMessageBox.Show(this, Environment.NewLine + "Archivo creado y depositado en Loftware. Las etiquetas deberan empezar a imprimirse y Mi-LAW cerrara automaticamente.", "Datos Validos", MessageBoxButtons.OK, MessageBoxIcon.Information);

                                myConnection.Open();
                                SqlCommand cmd3 = new SqlCommand("INSERT INTO dbo.tblKCIHistory (WorkOrder, Skid, Item, MfgDate, ExpDate, LblQty, BoxQty, LblAW, LblPath, LblGTIN, LblOrigin, Timestamp, iCodUser_ID, SysUser, MachineID, PrinterID) VALUES(@WO, @SKID, @Item, @MfgDate, @ExpDate, @LblQty, @BoxQty, @LblAW, @LblPath, @LblGTIN, @LblOrigin, @TimeStamp, @User, @SysUser, @Machine, @Printer)", myConnection);
                                cmd3.Parameters.AddWithValue("@WO", WO_Input.Text.ToString().ToUpper());
                                cmd3.Parameters.AddWithValue("@SKID", rSKID);
                                cmd3.Parameters.AddWithValue("@Item", resul.mitm);
                                cmd3.Parameters.AddWithValue("@MfgDate", MfgDate);
                                cmd3.Parameters.AddWithValue("@ExpDate", sEXPDATE);
                                cmd3.Parameters.AddWithValue("@LblQty", results.LblQty.ToString().Trim());
                                cmd3.Parameters.AddWithValue("@BoxQty", "");
                                cmd3.Parameters.AddWithValue("@LblAW", AWScanned);
                                cmd3.Parameters.AddWithValue("@LblPath", results.CartonLWL.ToString().Trim());
                                cmd3.Parameters.AddWithValue("@LblGTIN", results.CartonGTIN.ToString().Trim());
                                cmd3.Parameters.AddWithValue("@LblOrigin", results.Origin);
                                cmd3.Parameters.AddWithValue("@TimeStamp", printDate.ToString());
                                cmd3.Parameters.AddWithValue("@User", Login.UserN.ToString());
                                cmd3.Parameters.AddWithValue("@SysUser", System.Environment.UserName.ToString());
                                cmd3.Parameters.AddWithValue("@Machine", System.Environment.MachineName.ToString());
                                cmd3.Parameters.AddWithValue("@Printer", LWPRINTER);
                                cmd3.ExecuteNonQuery();
                                myConnection.Close();

                                Application.Exit();
                            }
                        }

                        //Validate AW Input against ShipperAW on DB
                        else if (AWScanned == results.ShipperAW.ToString().ToUpper().Trim())
                        {
                            if (results.ShipperLWL == "" || results.ShipperLWL.Substring(results.ShipperLWL.Length - 4) != ".lwl")
                            {
                                MetroMessageBox.Show(this, Environment.NewLine + "No se encontro etiqueta dada de alta en Base de Datos de Mi-LAW." + Environment.NewLine + "Favor de contactar a tu Supervisor y/o al equipo de IT Labeling.", "Error - Etiqueta no encontrada en DB", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                            }

                            else
                            {
                                myConnection.Close();

                                TextWriter csvfile = new StreamWriter(csvfilename, true);
                                csvfile.WriteLine("PRTNUM,Format,LLMQTY,BOXQTY,LOT_NUM,UMFG_DATE,UEXP_DATE");
                                csvfile.WriteLine('"' + LWPRINTER + '"' + "," + '"' + results.ShipperLWL + '"' + "," + '"' + results.LblQty + '"' + "," 
                                    + '"' + results.BoxQty + '"' + "," + '"' + rSKID + '"' + "," + '"' + MfgDate + '"' + "," + '"' + sEXPDATE + '"');
                                csvfile.Close();

                                MetroMessageBox.Show(this, Environment.NewLine + "Archivo creado y depositado en Loftware. Las etiquetas deberan empezar a imprimirse y Mi-LAW cerrara automaticamente.", "Datos Validos", MessageBoxButtons.OK, MessageBoxIcon.Information);

                                myConnection.Open();
                                SqlCommand cmd3 = new SqlCommand("INSERT INTO dbo.tblKCIHistory (WorkOrder, Skid, Item, MfgDate, ExpDate, LblQty, BoxQty, LblAW, LblPath, LblGTIN, LblOrigin, Timestamp, iCodUser_ID, SysUser, MachineID, PrinterID) VALUES(@WO, @SKID, @Item, @MfgDate, @ExpDate, @LblQty, @BoxQty, @LblAW, @LblPath, @LblGTIN, @LblOrigin, @TimeStamp, @User, @SysUser, @Machine, @Printer)", myConnection);
                                cmd3.Parameters.AddWithValue("@WO", WO_Input.Text.ToString().ToUpper());
                                cmd3.Parameters.AddWithValue("@SKID", rSKID);
                                cmd3.Parameters.AddWithValue("@Item", resul.mitm);
                                cmd3.Parameters.AddWithValue("@MfgDate", MfgDate);
                                cmd3.Parameters.AddWithValue("@ExpDate", sEXPDATE);
                                cmd3.Parameters.AddWithValue("@LblQty", results.LblQty.ToString().Trim());
                                cmd3.Parameters.AddWithValue("@BoxQty", results.BoxQty.ToString().Trim());
                                cmd3.Parameters.AddWithValue("@LblAW", AWScanned);
                                cmd3.Parameters.AddWithValue("@LblPath", results.ShipperLWL.ToString().Trim());
                                cmd3.Parameters.AddWithValue("@LblGTIN", "");
                                cmd3.Parameters.AddWithValue("@LblOrigin", "");
                                cmd3.Parameters.AddWithValue("@TimeStamp", printDate.ToString());
                                cmd3.Parameters.AddWithValue("@User", Login.UserN.ToString());
                                cmd3.Parameters.AddWithValue("@SysUser", System.Environment.UserName.ToString());
                                cmd3.Parameters.AddWithValue("@Machine", System.Environment.MachineName.ToString());
                                cmd3.Parameters.AddWithValue("@Printer", LWPRINTER);
                                cmd3.ExecuteNonQuery();
                                myConnection.Close();

                                Application.Exit();
                            }
                        }

                        //Validate AW Input against L1AW on DB
                        else if (AWScanned == results.L1AW.ToString().ToUpper().Trim())
                        {
                            if (results.L1LWL == "" || results.L1LWL.Substring(results.L1LWL.Length - 4) != ".lwl")
                            {
                                MetroMessageBox.Show(this, Environment.NewLine + "No se encontro etiqueta dada de alta en Base de Datos de Mi-LAW." + Environment.NewLine + "Favor de contactar a tu Supervisor y/o al equipo de IT Labeling.", "Error - Etiqueta no encontrada en DB", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                            }

                            else
                            {
                                myConnection.Close();

                                TextWriter csvfile = new StreamWriter(csvfilename, true);
                                csvfile.WriteLine("PRTNUM,Format");
                                csvfile.WriteLine('"' + LWPRINTER + '"' + "," + '"' + results.L1LWL + '"');
                                csvfile.Close();

                                MetroMessageBox.Show(this, Environment.NewLine + "Archivo creado y depositado en Loftware. Las etiquetas deberan empezar a imprimirse y Mi-LAW cerrara automaticamente.", "Datos Validos", MessageBoxButtons.OK, MessageBoxIcon.Information);

                                myConnection.Open();
                                SqlCommand cmd3 = new SqlCommand("INSERT INTO dbo.tblKCIHistory (WorkOrder, Skid, Item, MfgDate, ExpDate, LblQty, BoxQty, LblAW, LblPath, LblGTIN, LblOrigin, Timestamp, iCodUser_ID, SysUser, MachineID, PrinterID) VALUES(@WO, @SKID, @Item, @MfgDate, @ExpDate, @LblQty, @BoxQty, @LblAW, @LblPath, @LblGTIN, @LblOrigin, @TimeStamp, @User, @SysUser, @Machine, @Printer)", myConnection);
                                cmd3.Parameters.AddWithValue("@WO", WO_Input.Text.ToString().ToUpper());
                                cmd3.Parameters.AddWithValue("@SKID", rSKID);
                                cmd3.Parameters.AddWithValue("@Item", resul.mitm);
                                cmd3.Parameters.AddWithValue("@MfgDate", MfgDate);
                                cmd3.Parameters.AddWithValue("@ExpDate", sEXPDATE);
                                cmd3.Parameters.AddWithValue("@LblQty", results.LblQty.ToString().Trim());
                                cmd3.Parameters.AddWithValue("@BoxQty", "");
                                cmd3.Parameters.AddWithValue("@LblAW", AWScanned);
                                cmd3.Parameters.AddWithValue("@LblPath", results.PouchLWL.ToString().Trim());
                                cmd3.Parameters.AddWithValue("@LblGTIN", "");
                                cmd3.Parameters.AddWithValue("@LblOrigin", "");
                                cmd3.Parameters.AddWithValue("@TimeStamp", printDate.ToString());
                                cmd3.Parameters.AddWithValue("@User", Login.UserN.ToString());
                                cmd3.Parameters.AddWithValue("@SysUser", System.Environment.UserName.ToString());
                                cmd3.Parameters.AddWithValue("@Machine", System.Environment.MachineName.ToString());
                                cmd3.Parameters.AddWithValue("@Printer", LWPRINTER);
                                cmd3.ExecuteNonQuery();
                                myConnection.Close();

                                Application.Exit();
                            }
                        }

                        //Validate AW Input against L2AW on DB
                        else if (AWScanned == results.L2AW.ToString().ToUpper().Trim())
                        {
                            if (results.L2LWL == "" || results.L2LWL.Substring(results.L2LWL.Length - 4) != ".lwl")
                            {
                                MetroMessageBox.Show(this, Environment.NewLine + "No se encontro etiqueta dada de alta en Base de Datos de Mi-LAW." + Environment.NewLine + "Favor de contactar a tu Supervisor y/o al equipo de IT Labeling.", "Error - Etiqueta no encontrada en DB", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                            }

                            else
                            {
                                myConnection.Close();

                                TextWriter csvfile = new StreamWriter(csvfilename, true);
                                csvfile.WriteLine("PRTNUM,Format");
                                csvfile.WriteLine('"' + LWPRINTER + '"' + "," + '"' + results.L2LWL + '"');
                                csvfile.Close();

                                MetroMessageBox.Show(this, Environment.NewLine + "Archivo creado y depositado en Loftware. Las etiquetas deberan empezar a imprimirse y Mi-LAW cerrara automaticamente.", "Datos Validos", MessageBoxButtons.OK, MessageBoxIcon.Information);

                                myConnection.Open();
                                SqlCommand cmd3 = new SqlCommand("INSERT INTO dbo.tblKCIHistory (WorkOrder, Skid, Item, MfgDate, ExpDate, LblQty, BoxQty, LblAW, LblPath, LblGTIN, LblOrigin, Timestamp, iCodUser_ID, SysUser, MachineID, PrinterID) VALUES(@WO, @SKID, @Item, @MfgDate, @ExpDate, @LblQty, @BoxQty, @LblAW, @LblPath, @LblGTIN, @LblOrigin, @TimeStamp, @User, @SysUser, @Machine, @Printer)", myConnection);
                                cmd3.Parameters.AddWithValue("@WO", WO_Input.Text.ToString().ToUpper());
                                cmd3.Parameters.AddWithValue("@SKID", rSKID);
                                cmd3.Parameters.AddWithValue("@Item", resul.mitm);
                                cmd3.Parameters.AddWithValue("@MfgDate", MfgDate);
                                cmd3.Parameters.AddWithValue("@ExpDate", sEXPDATE);
                                cmd3.Parameters.AddWithValue("@LblQty", results.LblQty.ToString().Trim());
                                cmd3.Parameters.AddWithValue("@BoxQty", "");
                                cmd3.Parameters.AddWithValue("@LblAW", AWScanned);
                                cmd3.Parameters.AddWithValue("@LblPath", results.PouchLWL.ToString().Trim());
                                cmd3.Parameters.AddWithValue("@LblGTIN", "");
                                cmd3.Parameters.AddWithValue("@LblOrigin", "");
                                cmd3.Parameters.AddWithValue("@TimeStamp", printDate.ToString());
                                cmd3.Parameters.AddWithValue("@User", Login.UserN.ToString());
                                cmd3.Parameters.AddWithValue("@SysUser", System.Environment.UserName.ToString());
                                cmd3.Parameters.AddWithValue("@Machine", System.Environment.MachineName.ToString());
                                cmd3.Parameters.AddWithValue("@Printer", LWPRINTER);
                                cmd3.ExecuteNonQuery();
                                myConnection.Close();

                                Application.Exit();
                            }
                        }

                        else
                        {
                            myConnection.Close();
                            MetroMessageBox.Show(this, Environment.NewLine + "Error de numero de ArtWork. El Numero escaneado es Incorrecto o no corresponde al Numero de Parte, Favor de contactar a tu Supervisor.", "Error en Numero de ArtWork", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }

                    }
                    else
                    {
                        myConnection.Close();
                        MetroMessageBox.Show(this, Environment.NewLine + "El numero de parte de la WO que escaneaste no existe en la base de datos de Mi-LAW" + Environment.NewLine + "Favor de contactar a tu Supervisor/Ingeniero de Manufactura y/o al equipo de IT Labeling", "Error en Item Finish Good de la Work Order", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    }
                }
                else
                {
                    linksql2oracle.Close();
                    MetroMessageBox.Show(this, Environment.NewLine + "La Work Order que escaneaste/capturaste no existe en BAAN." + Environment.NewLine + "Favor de contactar a tu Supervisor.", "Error en Numero de Work Order", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                }
            }
        }

        public class Results
        {
            public Results() { }
            public string Item { get; set; }
            public string ShelfLife { get; set; }
            public string PouchGTIN { get; set; }
            public string CartonGTIN { get; set; }
            public string PouchAW { get; set; }
            public string CartonAW { get; set; }
            public string ShipperAW { get; set; }
            public string L1AW { get; set; }
            public string L2AW { get; set; }
            public string PouchLWL { get; set; }
            public string CartonLWL { get; set; }
            public string ShipperLWL { get; set; }
            public string L1LWL { get; set; }
            public string L2LWL { get; set; }
            public string Origin { get; set; }
            public string LblQty { get; set; }
            public string BoxQty { get; set; }
        }

        public class Resul
        {
            public Resul() { }
            public string attrib { get; set; }
            public string mitm { get; set; }
        }

        public class FormProvider
        {
            public static MainForm MainMenu
            {
                get
                {
                    if (_mainMenu == null)
                    {
                        _mainMenu = new MainForm();
                    }
                    return _mainMenu;
                }
            }
            private static MainForm _mainMenu;

        }
    }
}
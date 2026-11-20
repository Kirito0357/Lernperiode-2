using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Windows.Forms;

namespace Elden_Ring
{
    public partial class Form1 : Form
    {
        private readonly string excelDateiPfad =
            Path.Combine(
                AppContext.BaseDirectory,
                "Waffen_Datenbank Offiziell.xlsx"
            );

        private readonly Dictionary<string, string> gewaehlteRuestung =
            new Dictionary<string, string>();

        private const string KeineRuestung = "— Keine —";

        private FavoritenForm? favoritenFenster;

        public Form1()
        {
            InitializeComponent();

            // --------------------------------------------------
            // Grundkonfiguration
            // --------------------------------------------------

            button1.Visible = true;

            comboBox1.Visible = true;
            comboBox2.Visible = false;

            comboBox3.Visible = true;
            comboBox4.Visible = false;
            comboBox5.Visible = false;

            comboBox6.Visible = true;
            comboBox7.Visible = true;
            comboBox8.Visible = true;
            comboBox9.Visible = true;

            listBox1.Visible = false;
            listBox2.Visible = false;

            listBox3.Visible = true;
            listBox5.Visible = true;
            listBox6.Visible = true;
            listBox7.Visible = true;

            listBox1.HorizontalScrollbar = true;
            listBox1.SelectionMode = SelectionMode.None;

            listBox2.HorizontalScrollbar = true;
            listBox2.SelectionMode = SelectionMode.None;

            listBox3.HorizontalScrollbar = true;
            listBox3.SelectionMode = SelectionMode.None;

            listBox4.HorizontalScrollbar = true;
            listBox4.SelectionMode = SelectionMode.None;

            listBox5.HorizontalScrollbar = true;
            listBox5.SelectionMode = SelectionMode.None;

            listBox6.HorizontalScrollbar = true;
            listBox6.SelectionMode = SelectionMode.None;

            listBox7.HorizontalScrollbar = true;
            listBox7.SelectionMode = SelectionMode.None;

            this.BackColor = Design.Hintergrund;
            this.DoubleBuffered = true;
            this.ResizeRedraw = true;
            this.Paint += Form1_Paint;

            label12.Font = Design.Titel;
            label12.ForeColor = Design.GoldHell;
            label12.BackColor = Color.Transparent;

            linie1.BackColor = Design.GoldDunkel;
            linie2.BackColor = Design.GoldDunkel;
            linie3.BackColor = Design.GoldDunkel;
            linie4.BackColor = Design.GoldDunkel;
            linie5.BackColor = Design.GoldDunkel;
            linie6.BackColor = Design.GoldDunkel;

            label1.Font = Design.Ueberschrift;
            label1.ForeColor = Design.GoldHell;
            label1.BackColor = Color.Transparent;
            label2.Font = Design.Ueberschrift;
            label2.ForeColor = Design.GoldHell;
            label2.BackColor = Color.Transparent;
            label3.Font = Design.Ueberschrift;
            label3.ForeColor = Design.GoldHell;
            label3.BackColor = Color.Transparent;
            label4.Font = Design.Ueberschrift;
            label4.ForeColor = Design.GoldHell;
            label4.BackColor = Color.Transparent;
            label5.Font = Design.Ueberschrift;
            label5.ForeColor = Design.GoldHell;
            label5.BackColor = Color.Transparent;

            comboBox1.Font = Design.Text;
            comboBox1.ForeColor = Design.Gold;
            comboBox1.BackColor = Design.Feld;
            comboBox2.Font = Design.Text;
            comboBox2.ForeColor = Design.Gold;
            comboBox2.BackColor = Design.Feld;
            comboBox3.Font = Design.Text;
            comboBox3.ForeColor = Design.Gold;
            comboBox3.BackColor = Design.Feld;
            comboBox4.Font = Design.Text;
            comboBox4.ForeColor = Design.Gold;
            comboBox4.BackColor = Design.Feld;
            comboBox5.Font = Design.Text;
            comboBox5.ForeColor = Design.Gold;
            comboBox5.BackColor = Design.Feld;
            comboBox6.Font = Design.Text;
            comboBox6.ForeColor = Design.Gold;
            comboBox6.BackColor = Design.Feld;
            comboBox7.Font = Design.Text;
            comboBox7.ForeColor = Design.Gold;
            comboBox7.BackColor = Design.Feld;
            comboBox8.Font = Design.Text;
            comboBox8.ForeColor = Design.Gold;
            comboBox8.BackColor = Design.Feld;
            comboBox9.Font = Design.Text;
            comboBox9.ForeColor = Design.Gold;
            comboBox9.BackColor = Design.Feld;

            label13.Font = Design.Text;
            label13.ForeColor = Design.GoldHell;
            label13.BackColor = Color.Transparent;
            label14.Font = Design.Text;
            label14.ForeColor = Design.GoldHell;
            label14.BackColor = Color.Transparent;
            label15.Font = Design.Text;
            label15.ForeColor = Design.GoldHell;
            label15.BackColor = Color.Transparent;
            label16.Font = Design.Text;
            label16.ForeColor = Design.GoldHell;
            label16.BackColor = Color.Transparent;

            listBox1.Font = Design.Text;
            listBox1.ForeColor = Design.Gold;
            listBox1.BackColor = Design.Feld;
            listBox2.Font = Design.Text;
            listBox2.ForeColor = Design.Gold;
            listBox2.BackColor = Design.Feld;
            listBox3.Font = Design.Text;
            listBox3.ForeColor = Design.Gold;
            listBox3.BackColor = Design.Feld;
            listBox4.Font = Design.Text;
            listBox4.ForeColor = Design.Gold;
            listBox4.BackColor = Design.Feld;
            listBox5.Font = Design.Text;
            listBox5.ForeColor = Design.Gold;
            listBox5.BackColor = Design.Feld;
            listBox6.Font = Design.Text;
            listBox6.ForeColor = Design.Gold;
            listBox6.BackColor = Design.Feld;
            listBox7.Font = Design.Text;
            listBox7.ForeColor = Design.Gold;
            listBox7.BackColor = Design.Feld;

            button1.Font = Design.Text;
            button1.ForeColor = Design.Gold;
            button1.BackColor = Design.Feld;
            button1.FlatAppearance.BorderColor = Design.GoldDunkel;
            button1.FlatAppearance.MouseOverBackColor = Design.KnopfMaus;
            button1.FlatAppearance.MouseDownBackColor = Design.KnopfGedrueckt;

            button2.Font = Design.Text;
            button2.ForeColor = Design.GoldHell;
            button2.BackColor = Design.Feld;
            button2.FlatAppearance.BorderColor = Design.GoldDunkel;
            button2.FlatAppearance.MouseOverBackColor = Design.KnopfMaus;
            button2.FlatAppearance.MouseDownBackColor = Design.KnopfGedrueckt;

            button3.Font = Design.Text;
            button3.ForeColor = Design.GoldHell;
            button3.BackColor = Design.Feld;
            button3.FlatAppearance.BorderColor = Design.GoldDunkel;
            button3.FlatAppearance.MouseOverBackColor = Design.KnopfMaus;
            button3.FlatAppearance.MouseDownBackColor = Design.KnopfGedrueckt;

            button4.Font = Design.Text;
            button4.ForeColor = Design.GoldHell;
            button4.BackColor = Design.Feld;
            button4.FlatAppearance.BorderColor = Design.GoldDunkel;
            button4.FlatAppearance.MouseOverBackColor = Design.KnopfMaus;
            button4.FlatAppearance.MouseDownBackColor = Design.KnopfGedrueckt;

            comboBox5.SelectedIndexChanged += comboBox5_SelectedIndexChanged;
            comboBox7.SelectedIndexChanged += comboBox7_SelectedIndexChanged;
            comboBox8.SelectedIndexChanged += comboBox8_SelectedIndexChanged;
            comboBox9.SelectedIndexChanged += comboBox9_SelectedIndexChanged;

            // Stats: Farben
            label6.Font = Design.Text;
            label6.ForeColor = Design.Gold;
            label6.BackColor = Color.Transparent;
            label7.Font = Design.Text;
            label7.ForeColor = Design.Gold;
            label7.BackColor = Color.Transparent;
            label8.Font = Design.Text;
            label8.ForeColor = Design.Gold;
            label8.BackColor = Color.Transparent;
            label9.Font = Design.Text;
            label9.ForeColor = Design.Gold;
            label9.BackColor = Color.Transparent;
            label10.Font = Design.Text;
            label10.ForeColor = Design.Gold;
            label10.BackColor = Color.Transparent;

            numericUpDown1.Font = Design.Text;
            numericUpDown1.ForeColor = Design.Gold;
            numericUpDown1.BackColor = Design.Feld;
            numericUpDown2.Font = Design.Text;
            numericUpDown2.ForeColor = Design.Gold;
            numericUpDown2.BackColor = Design.Feld;
            numericUpDown3.Font = Design.Text;
            numericUpDown3.ForeColor = Design.Gold;
            numericUpDown3.BackColor = Design.Feld;
            numericUpDown4.Font = Design.Text;
            numericUpDown4.ForeColor = Design.Gold;
            numericUpDown4.BackColor = Design.Feld;
            numericUpDown5.Font = Design.Text;
            numericUpDown5.ForeColor = Design.Gold;
            numericUpDown5.BackColor = Design.Feld;

            label11.Font = Design.Text;
            label11.ForeColor = Design.GoldHell;
            label11.BackColor = Color.Transparent;
            numericUpDown6.Font = Design.Text;
            numericUpDown6.ForeColor = Design.GoldHell;
            numericUpDown6.BackColor = Design.Feld;

            numericUpDown1.ValueChanged += StatsGeaendert;
            numericUpDown2.ValueChanged += StatsGeaendert;
            numericUpDown3.ValueChanged += StatsGeaendert;
            numericUpDown4.ValueChanged += StatsGeaendert;
            numericUpDown5.ValueChanged += StatsGeaendert;

            numericUpDown1.TextChanged += StatEingabePruefen;
            numericUpDown2.TextChanged += StatEingabePruefen;
            numericUpDown3.TextChanged += StatEingabePruefen;
            numericUpDown4.TextChanged += StatEingabePruefen;
            numericUpDown5.TextChanged += StatEingabePruefen;
            numericUpDown6.ValueChanged += AusdauerGeaendert;
            numericUpDown6.TextChanged += StatEingabePruefen;

            this.FormClosing += Form1_FormClosing;

            // --------------------------------------------------
            // Excel-Datei prüfen
            // --------------------------------------------------

            if (!File.Exists(excelDateiPfad))
            {
                MessageBox.Show(
                    "Die Excel-Datei wurde nicht gefunden.\n\n" +
                    "Gesuchter Pfad:\n" +
                    excelDateiPfad +
                    "\n\n" +
                    "Stelle in Visual Studio bei der Excel-Datei " +
                    "\"Kopieren, wenn neuer\" ein.",
                    "Excel-Datei fehlt",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }

            // --------------------------------------------------
            // Kategorien laden
            // --------------------------------------------------

            comboBox1.Items.AddRange(new string[]
            {
                "Dolche",
                "Gerade Schwerter",
                "Grossschwerter",
                "Riesenschwerter",
                "Stossschwerter",
                "Schwere Stossschwerter",
                "Krummschwerter",
                "Krumme Grossschwerter",
                "Katanas",
                "Doppelklingen",
                "Äxte",
                "Grossäxte",
                "Hämmer",
                "Flegel",
                "Kriegshämmer",
                "Gigantische Waffen",
                "Speere",
                "Grosse Speere",
                "Hellebarden",
                "Schnitter",
                "Peitschen",
                "Fäuste",
                "Krallen"
            });

            comboBox3.Items.AddRange(new string[]
            {
                "Fernkampf",
                "Katalysatoren"
            });
            LadeRuestungsListe(comboBox6, "Kopfbedeckungen");
            LadeRuestungsListe(comboBox7, "Torsos");
            LadeRuestungsListe(comboBox8, "Armschützer");
            LadeRuestungsListe(comboBox9, "Beinschützer");

            LadeStats();

            // --------------------------------------------------
            // Startanzeige
            // --------------------------------------------------
            ZeigeLeereRuestung("Kopfbedeckungen");
            ZeigeLeereRuestung("Torsos");
            ZeigeLeereRuestung("Armschützer");
            ZeigeLeereRuestung("Beinschützer");

            AktualisiereGesamtgewicht();
        }

        // ============================================================
        // NAHKAMPF
        // ============================================================

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox2.Items.Clear();
            comboBox2.SelectedIndex = -1;
            comboBox2.Text = "";

            listBox1.Items.Clear();
            AktualisiereGesamtgewicht();

            if (comboBox1.SelectedIndex == -1)
            {
                comboBox2.Visible = false;
                return;
            }

            comboBox2.Visible = true;

            switch (comboBox1.Text)
            {
                case "Dolche":
                    comboBox2.Items.AddRange(new string[]
                    {
                        "Dolch", "Parierdolch", "Misericordia", "Grosses Messer",
                        "Blutbefleckter Dolch", "Erdenstahldolch", "Wakizashi",
                        "Zelebrantensichel", "Elfenbeinsichel", "Kristallmesser",
                        "Skorpionstachel", "Cinqueda", "Schimmersteinkris",
                        "Reduvia", "Klinge des Rufs", "Schwarzes Messer"
                    });
                    break;

                case "Gerade Schwerter":
                    comboBox2.Items.AddRange(new string[]
                    {
                        "Kurzschwert", "Langschwert", "Breitschwert", "Altes Schwert",
                        "Schwert der Getreuen", "Feines Adligenschwert", "Stockschwert",
                        "Kriegsfalkenkralle", "Lazuli-Schimmersteinschwert",
                        "Carianisches Ritterschwert", "Kristallschwert",
                        "Fäulniskristallschwert", "Miquella-Ritterschwert",
                        "Ornamentales Schwert", "Goldenes Epitaphium",
                        "Schwert von St. Trina", "Eochaids Zierde",
                        "Schwert der geheimen Worte", "Schwert der Nacht und der flamme"
                    });
                    break;

                case "Grossschwerter":
                    comboBox2.Items.AddRange(new string[]
                    {
                        "Bastardschwert", "Claymore", "Eisengrossschwert",
                        "Grossschwert der Getreuen", "Rittergrossschwert",
                        "Grossschwert der Exilritter", "Gegabeltes Grossschwert",
                        "Flamberge", "Gargoyle-Grossschwert", "Gargoyle-Schwarzschwert",
                        "Unzertrennliches Schwert", "Milos-Schwert",
                        "Henkerschwert der Marais", "Ordovis'Grossschwert",
                        "Alabasterfürstenschwert", "Schüreisen des Todes",
                        "Helphens Spitzturm", "Blasphemische Klinge",
                        "Grossschwert der Goldenen Ordnung", "Dunkelmondgrossschwert",
                        "Heiliges Reliktschwert"
                    });
                    break;

                case "Riesenschwerter":
                    comboBox2.Items.AddRange(new string[]
                    {
                        "Zweihänder", "Grossschwert", "Grossschwert des Bewachers",
                        "Goldenes Trollschwert", "Trollritterschwert",
                        "Königliches Grossschwert", "Grossschwert der Verpflanzung",
                        "Trümmergrossschwert", "Grossschwerter der Sternengeissel",
                        "Grossschwert der Gottestöterin", "Malikeths schwarze Klinge"
                    });
                    break;

                case "Stossschwerter":
                    comboBox2.Items.AddRange(new string[]
                    {
                        "Rapier", "Panzerbrecher", "Adeligenpanzerbrecher",
                        "Schwert des Reinfäule-Ritters", "Rogiers Rapier",
                        "Ameisenstachel-Rapier", "Eisnadel"
                    });
                    break;

                case "Schwere Stossschwerter":
                    comboBox2.Items.AddRange(new string[]
                    {
                        "Grosser Degen", "Götterskalpnadel", "Blutige Helix",
                        "Felsklinge des Drachenkönigs"
                    });
                    break;

                case "Krummschwerter":
                    comboBox2.Items.AddRange(new string[]
                    {
                        "Krummsäbel", "Falchion", "Shamshir", "Grossmesser",
                        "Banditenkrummschwert", "Shotel", "Plündererkrummschwert",
                        "Mantisklinge", "Bestienkrummschwert", "Fliessendes Krummschwert",
                        "Schlangengottkrummschwert", "Magmaklinge",
                        "Fliessendes Noxschwert", "Schwinge von Astel", "Shotel der Eklipse"
                    });
                    break;

                case "Krumme Grossschwerter":
                    comboBox2.Items.AddRange(new string[]
                    {
                        "Reitergeissel", "Malbeil", "Mönchsflammenschwert",
                        "Beil des Bestienmenschen", "Bluthundreisszahn",
                        "Onyxfürstengrossschwert", "Zamorkrummschwert",
                        "Magmalindwurm-Schuppenschwert", "Morgotts verfluchtes Schwert"
                    });
                    break;

                case "Katanas":
                    comboBox2.Items.AddRange(new string[]
                    {
                        "Uchigatana", "Nagakiba", "Schlangenknochenklinge",
                        "Meteorklinge", "Mondschleier", "Blutige Ströme",
                        "Drachenschuppenklinge", "Malenias Hand"
                    });
                    break;

                case "Doppelklingen":
                    comboBox2.Items.AddRange(new string[]
                    {
                        "Doppelklinge", "Ritter-Zwillingsschwerter", "Götterskalphäuter",
                        "Gargoyle-Doppelklinge", "Gargoyle-Doppelschwarzklinge",
                        "Eleonoras Stangenwaffe"
                    });
                    break;

                case "Äxte":
                    comboBox2.Items.AddRange(new string[]
                    {
                        "Handaxt", "Gespaltene Sichel", "Streitaxt", "Verformte Axt",
                        "Kieferknochen-Axt", "Eisenbeil", "Hochlandaxt",
                        "Zelebrantenbeil", "Opferpriesteraxt", "Eisiges Beil",
                        "Wellenklinge", "Sturmfalkenaxt", "Rosus'Axt"
                    });
                    break;

                case "Grossäxte":
                    comboBox2.Items.AddRange(new string[]
                    {
                        "Grossaxt", "Halbmondaxt", "Langstielbeil", "Henkergrossaxt",
                        "Maltöter-Grossbeil", "Rostiger Anker", "Fleischermesser",
                        "Gargoyle-Grossaxt", "Gargoyle-Schwarzaxt",
                        "Schwingengrosshorn", "Godricks Axt"
                    });
                    break;

                case "Hämmer":
                    comboBox2.Items.AddRange(new string[]
                    {
                        "Keule", "Geschwungene Keule", "Stachelkeule", "Steinkeule",
                        "Streitkolben", "Morgenstern", "Spaltaxt", "Hammer",
                        "Mönchsflammenstreitkolben", "Varrés Blumenstrauss",
                        "Horn des Gesandten", "Fliessender Noxhammer",
                        "Beringter Finger", "Zepter des Allwissenden"
                    });
                    break;

                case "Flegel":
                    comboBox2.Items.AddRange(new string[]
                    {
                        "Bastardsterne", "Flegel", "Flegel der Dunkelkavallerie",
                        "Kettenflegel", "Köpfe der Familie"
                    });
                    break;

                case "Kriegshämmer":
                    comboBox2.Items.AddRange(new string[]
                    {
                        "Grosse Keule", "Geschwungene Grosskeule", "Grossstreitkolben",
                        "Spitzhacke", "Ziegelhammer", "Streithammer",
                        "Fäulnisstreithammer", "Zelebrantenschädel", "Grosssterne",
                        "Grosshornhammer", "Langhorn des Gesandten",
                        "Schädelkerzenständer", "Bestienklauengrosshammer",
                        "Verschlingerzepter"
                    });
                    break;

                case "Gigantische Waffen":
                    comboBox2.Items.AddRange(new string[]
                    {
                        "Duellantengrossaxt", "Fäulnisgrossaxt", "Golemhellebarde",
                        "Riesenbrecher", "Infernalischer Prälatenstab", "Grosskeule",
                        "Trollhammer", "Grosse Drachenkralle", "Stab des Bewachers",
                        "Stab des Avatars", "Fäulnisstab", "Grosshorn des Gesandten",
                        "Ghizas Rad", "Sternenfallbestienkiefer", "Godfreys Axt"
                    });
                    break;

                case "Speere":
                    comboBox2.Items.AddRange(new string[]
                    {
                        "Kurzspeer", "Eisenspeer", "Speer", "Partisane", "Pike",
                        "Stachelspeer", "Kreuz-Naginata", "Tonmenschenharpune",
                        "Zelebrantenrippenharke", "Stabfackel",
                        "Girandole des Inquisitors", "Kristallspeer",
                        "Fäulniskristallspeer", "Reinfäule-Speer",
                        "Todesritualspeer", "Blitz von Gransax"
                    });
                    break;

                case "Grosse Speere":
                    comboBox2.Items.AddRange(new string[]
                    {
                        "Lanze", "Baumspeer", "Schlangenjäger", "Silurias Baum",
                        "Vykes Kriegsspeer", "Mohgwyns heiliger Speer"
                    });
                    break;

                case "Hellebarden":
                    comboBox2.Items.AddRange(new string[]
                    {
                        "Hellebarde", "Hellebarde der Exilritter", "Falkenschnabel",
                        "Gleve", "Shotel der Niederen Soldaten",
                        "Säge der Niederen Soldaten", "Wächter-Schwertspeer",
                        "Gargoyle-Hellebarde", "Gargoyle-Schwarzhellebarde",
                        "Gleve der Dunkelkavallerie", "Seuchengleve",
                        "Wellenhellebarde", "Goldene Hellebarde", "Drachenhellebarde",
                        "Lorettas Kriegssichel", "Kommandantenstandarte"
                    });
                    break;

                case "Schnitter":
                    comboBox2.Items.AddRange(new string[]
                    {
                        "Sense", "Grabsense", "Aureolensense", "Schwingensense"
                    });
                    break;

                case "Peitschen":
                    comboBox2.Items.AddRange(new string[]
                    {
                        "Peitsche", "Dornenpeitsche", "Urumi", "Hoslows Blütenpeitsche",
                        "Magmapeitschenleuchter", "Roter Riesenhaarzopf"
                    });
                    break;

                case "Fäuste":
                    comboBox2.Items.AddRange(new string[]
                    {
                        "Caestus", "Stachelcaestus", "Katar", "Eisenkugel",
                        "Sternfaust", "Greifende Knochenhand", "Veteranenprothese",
                        "Verschlüsseltes Pata", "Verpflanzter Drache"
                    });
                    break;

                case "Krallen":
                    comboBox2.Items.AddRange(new string[]
                    {
                        "Hakenkrallen", "Bluthundkrallen", "Schlangenzahn",
                        "Raptorkrallen"
                    });
                    break;
            }
        }

        // ============================================================
        // WAFFE AUS EXCEL LADEN
        // ============================================================

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            listBox1.Items.Clear();

            if (comboBox2.SelectedIndex == -1)
            {
                AktualisiereGesamtgewicht();
                return;
            }

            string waffe = comboBox2.SelectedItem?.ToString() ?? "";

            LadeWaffeAusExcel(waffe, listBox1);
            listBox1.Visible = true;

            AktualisiereGesamtgewicht();
        }

        private void LadeWaffeAusExcel(string waffe, ListBox zielListBox)
        {
            zielListBox.Items.Clear();

            if (!File.Exists(excelDateiPfad))
            {
                zielListBox.Items.Add("Excel-Datei wurde nicht gefunden.");
                zielListBox.Items.Add("");
                zielListBox.Items.Add("Gesuchter Pfad:");
                zielListBox.Items.Add(excelDateiPfad);
                return;
            }

            try
            {
                using (XLWorkbook workbook = OeffneExcel())
                {
                    foreach (var blatt in workbook.Worksheets)
                    {
                        string kopf = blatt.Cell(3, 1).GetString().Trim();

                        if (!string.Equals(kopf, "Waffe", StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }

                        int letzteZeile = blatt.LastRowUsed()?.RowNumber() ?? 3;

                        for (int zeile = 4; zeile <= letzteZeile; zeile++)
                        {
                            string name = blatt.Cell(zeile, 1).GetString().Trim();

                            if (NormalisiereName(name) != NormalisiereName(waffe))
                            {
                                continue;
                            }

                            int letzteSpalte = blatt.Row(3).LastCellUsed()?.Address.ColumnNumber ?? 1;

                            zielListBox.Items.Add("——————————  ·  ——————————");
                            zielListBox.Items.Add("Waffe: " + name);
                            zielListBox.Items.Add("——————————  ·  ——————————");

                            for (int spalte = 2; spalte <= letzteSpalte; spalte++)
                            {
                                if (blatt.Cell(3, spalte).GetString().Trim() == "Kriegsasche")
                                {
                                    string asche = blatt.Cell(zeile, spalte).GetString().Trim();

                                    if (string.IsNullOrWhiteSpace(asche))
                                        asche = "Keine Angabe";

                                    zielListBox.Items.Add("Kriegsasche: " + asche);
                                    zielListBox.Items.Add("--------------------------------");
                                    break;
                                }
                            }

                            for (int spalte = 2; spalte <= letzteSpalte; spalte++)
                            {
                                string bezeichnung = blatt.Cell(3, spalte).GetString().Trim();
                                string wert = blatt.Cell(zeile, spalte).GetFormattedString();

                                if (string.IsNullOrWhiteSpace(bezeichnung))
                                    continue;

                                if (bezeichnung == "Kriegsasche")
                                    continue;

                                if (string.IsNullOrWhiteSpace(wert))
                                    wert = "Keine Angabe";

                                zielListBox.Items.Add(
                                    bezeichnung + ": " + wert
                                );
                            }

                            ZeigeWaffeMitStats(blatt, zeile, zielListBox);

                            return;
                        }
                    }
                }

                zielListBox.Items.Add("Keine Daten gefunden für: " + waffe);
            }
            catch (Exception ex)
            {
                zielListBox.Items.Clear();
                zielListBox.Items.Add("Fehler beim Lesen der Excel-Datei:");
                zielListBox.Items.Add("");
                zielListBox.Items.Add(ex.Message);
            }
        }

        // ============================================================
        // FERNKAMPF / KATALYSATOREN
        // ============================================================

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox4.Items.Clear();
            comboBox5.Items.Clear();

            comboBox4.SelectedIndex = -1;
            comboBox5.SelectedIndex = -1;

            comboBox4.Text = "";
            comboBox5.Text = "";

            comboBox5.Visible = false;
            listBox2.Items.Clear();
            AktualisiereGesamtgewicht();

            if (comboBox3.SelectedIndex == -1)
            {
                comboBox4.Visible = false;
                return;
            }

            comboBox4.Visible = true;

            if (comboBox3.Text == "Fernkampf")
            {
                comboBox4.Items.AddRange(new string[]
                {
                    "Leichte Bögen",
                    "Bögen",
                    "Grossbögen",
                    "Armbrüste",
                    "Ballisten"
                });
            }
            else if (comboBox3.Text == "Katalysatoren")
            {
                comboBox4.Items.AddRange(new string[]
                {
                    "Schimmersteinstäbe",
                    "Heilige Siegel"
                });
            }
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox5.Items.Clear();
            comboBox5.SelectedIndex = -1;
            comboBox5.Text = "";
            listBox2.Items.Clear();
            AktualisiereGesamtgewicht();

            if (comboBox4.SelectedIndex == -1)
            {
                comboBox5.Visible = false;
                return;
            }

            comboBox5.Visible = true;

            switch (comboBox4.Text)
            {
                case "Leichte Bögen":
                    comboBox5.Items.AddRange(new string[]
                    {
                        "Kurzbogen", "Kompositbogen", "Rotholzkurzbogen",
                        "Scheusalkurzbogen", "Harfenbogen"
                    });
                    break;

                case "Bögen":
                    comboBox5.Items.AddRange(new string[]
                    {
                        "Langbogen", "Bogen der Albinaurics", "Schwarzer Bogen",
                        "Flaschenzugbogen", "Hornbogen", "Schlangenbogen",
                        "Erdenbaum-Bogen"
                    });
                    break;

                case "Grossbögen":
                    comboBox5.Items.AddRange(new string[]
                    {
                        "Grossbogen", "Golemgrossbogen", "Erdenbaum-Grossbogen",
                        "Löwengrossbogen"
                    });
                    break;

                case "Armbrüste":
                    comboBox5.Items.AddRange(new string[]
                    {
                        "Soldatenarmbrust", "Leichte Armbrust", "Schwere Armbrust",
                        "Arbalest", "Crepus' Schwarzschlüsselarmbrust",
                        "Flaschenzugarmbrust", "Vollmondarmbrust"
                    });
                    break;

                case "Ballisten":
                    comboBox5.Items.AddRange(new string[]
                    {
                        "Handballiste", "Gefässkanone"
                    });
                    break;

                case "Schimmersteinstäbe":
                    comboBox5.Items.AddRange(new string[]
                    {
                        "Astrologenstab", "Schimmersteinstab",
                        "Akademie-Schimmersteinstab", "Gräberstab",
                        "Stab der Halbmenschenkönigin", "Azurner Schimmersteinstab",
                        "Lusats Schimmersteinstab", "Carianischer Schimmersteinstab",
                        "Carianischer Schimmerklingenstab",
                        "Carianisches Königliches Zepter", "Albinauric-Stab",
                        "Stab des Verlusts", "Schimmersteinstab vom Gelmir",
                        "Kristallstab", "Fäulniskristallstab", "Meteoritenstab",
                        "Stab der Schuldigen", "Stab des Todesprinzen"
                    });
                    break;

                case "Heilige Siegel":
                    comboBox5.Items.AddRange(new string[]
                    {
                        "Finger-Siegel", "Siegel des Erdenbaums",
                        "Siegel der Goldenen Ordnung", "Schottersteinsiegel",
                        "Siegel des Riesen", "Siegel des Gottestöters",
                        "Klauenzeichen-Siegel", "Siegel der Rasenden Flamme",
                        "Siegel der Drachenkommunion"
                    });
                    break;
            }
        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            listBox2.Items.Clear();

            if (comboBox5.SelectedIndex == -1)
            {
                AktualisiereGesamtgewicht();
                return;
            }

            string waffe = comboBox5.SelectedItem?.ToString() ?? "";

            LadeWaffeAusExcel(waffe, listBox2);
            listBox2.Visible = true;

            AktualisiereGesamtgewicht();
        }
        private void LadeRuestungsListe(ComboBox box, string blattName)
        {
            box.DropDownStyle = ComboBoxStyle.DropDownList;
            box.Items.Clear();

            box.Items.Add(KeineRuestung);

            if (!File.Exists(excelDateiPfad))
                return;

            try
            {
                using (XLWorkbook workbook = OeffneExcel())
                {
                    if (!workbook.TryGetWorksheet(blattName, out IXLWorksheet blatt))
                        return;

                    int letzteZeile = blatt.LastRowUsed()?.RowNumber() ?? 3;

                    for (int zeile = 4; zeile <= letzteZeile; zeile++)
                    {
                        string name = blatt.Cell(zeile, 1).GetString().Trim();

                        if (!string.IsNullOrWhiteSpace(name))
                            box.Items.Add(name);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Fehler beim Lesen der Excel-Datei:\n\n" + ex.Message,
                    "Fehler",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }

        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {
            RuestungGewaehlt(comboBox6, "Kopfbedeckungen");
        }

        private void comboBox7_SelectedIndexChanged(object sender, EventArgs e)
        {
            RuestungGewaehlt(comboBox7, "Torsos");
        }

        private void comboBox8_SelectedIndexChanged(object? sender, EventArgs e)
        {
            RuestungGewaehlt(comboBox8, "Armschützer");
        }

        private void comboBox9_SelectedIndexChanged(object? sender, EventArgs e)
        {
            RuestungGewaehlt(comboBox9, "Beinschützer");
        }
        private void RuestungGewaehlt(ComboBox box, string blattName)
        {
            if (box.SelectedIndex == -1)
                return;

            string ruestung = box.SelectedItem?.ToString() ?? "";

            ListBox? ziel = ListBoxFuerRuestung(blattName);

            if (ziel == null)
                return;

            if (ruestung == KeineRuestung)
            {
                gewaehlteRuestung.Remove(blattName);
                ZeigeLeereRuestung(blattName);
            }
            else
            {
                LadeRuestungAusExcel(ruestung, blattName, ziel);
                gewaehlteRuestung[blattName] = ruestung;
            }

            AktualisiereGesamtgewicht();
        }

        private void LadeRuestungAusExcel(string ruestung, string blattName, ListBox zielListBox)
        {
            zielListBox.Items.Clear();

            if (!File.Exists(excelDateiPfad))
            {
                zielListBox.Items.Add("Excel-Datei wurde nicht gefunden.");
                zielListBox.Items.Add("");
                zielListBox.Items.Add("Gesuchter Pfad:");
                zielListBox.Items.Add(excelDateiPfad);
                return;
            }

            try
            {
                using (XLWorkbook workbook = OeffneExcel())
                {
                    if (!workbook.TryGetWorksheet(blattName, out IXLWorksheet blatt))
                    {
                        zielListBox.Items.Add("Blatt nicht gefunden: " + blattName);
                        return;
                    }

                    int letzteZeile = blatt.LastRowUsed()?.RowNumber() ?? 3;
                    int letzteSpalte = blatt.Row(3).LastCellUsed()?.Address.ColumnNumber ?? 15;

                    for (int zeile = 4; zeile <= letzteZeile; zeile++)
                    {
                        string name = blatt.Cell(zeile, 1).GetString().Trim();

                        if (NormalisiereName(name) != NormalisiereName(ruestung))
                        {
                            continue;
                        }

                        zielListBox.Items.Add("——————————  ·  ——————————");
                        zielListBox.Items.Add(blattName + ": " + name);
                        zielListBox.Items.Add("——————————  ·  ——————————");

                        for (int spalte = 2; spalte <= letzteSpalte; spalte++)
                        {
                            string bezeichnung = blatt.Cell(3, spalte).GetString().Trim();
                            string wert = blatt.Cell(zeile, spalte).GetFormattedString();

                            if (string.IsNullOrWhiteSpace(bezeichnung))
                                continue;

                            if (string.IsNullOrWhiteSpace(wert))
                                wert = "Keine Angabe";

                            zielListBox.Items.Add(
                                bezeichnung + ": " + wert
                            );
                        }

                        return;
                    }
                }

                zielListBox.Items.Add("Keine Daten gefunden für: " + ruestung);
            }
            catch (Exception ex)
            {
                zielListBox.Items.Clear();
                zielListBox.Items.Add("Fehler beim Lesen der Excel-Datei:");
                zielListBox.Items.Add("");
                zielListBox.Items.Add(ex.Message);
            }
        }
        private ListBox? ListBoxFuerRuestung(string blattName)
        {
            switch (blattName)
            {
                case "Kopfbedeckungen": return listBox3;
                case "Torsos": return listBox5;
                case "Armschützer": return listBox6;
                case "Beinschützer": return listBox7;
                default: return null;
            }
        }

        private void ZeigeLeereRuestung(string blattName)
        {
            ListBox? ziel = ListBoxFuerRuestung(blattName);

            if (ziel == null)
                return;

            ziel.Items.Clear();
            ziel.Items.Add("——————————  ·  ——————————");
            ziel.Items.Add(blattName);
            ziel.Items.Add("——————————  ·  ——————————");
            ziel.Items.Add("Noch nichts ausgewählt.");
        }

        private void AktualisiereGesamtgewicht()
        {
            listBox4.Items.Clear();
            listBox4.Visible = true;

            var auswahl = new List<(string Anzeige, string Name, string? Blatt)>();

            if (comboBox2.SelectedIndex != -1)
                auswahl.Add(("Nahkampfwaffe", comboBox2.SelectedItem?.ToString() ?? "", null));

            if (comboBox5.SelectedIndex != -1)
                auswahl.Add(("Fernkampf / Katalysator", comboBox5.SelectedItem?.ToString() ?? "", null));

            string[,] ruestungsTeile =
            {
                { "Kopfbedeckungen", "Kopfbedeckung" },
                { "Torsos", "Brustschutz" },
                { "Armschützer", "Armpanzerung" },
                { "Beinschützer", "Beinschutz" }
            };

            for (int i = 0; i < ruestungsTeile.GetLength(0); i++)
            {
                string blatt = ruestungsTeile[i, 0];

                if (gewaehlteRuestung.TryGetValue(blatt, out string? teil))
                    auswahl.Add((ruestungsTeile[i, 1], teil, blatt));
            }

            listBox4.Items.Add("——————————  ·  ——————————");
            listBox4.Items.Add("Gewicht der Ausrüstung");
            listBox4.Items.Add("——————————  ·  ——————————");

            if (auswahl.Count == 0)
            {
                listBox4.Items.Add("Noch keine Ausrüstung ausgewählt.");
                return;
            }

            if (!File.Exists(excelDateiPfad))
            {
                listBox4.Items.Add("Excel-Datei wurde nicht gefunden.");
                return;
            }

            try
            {
                double gesamt = 0;

                using (XLWorkbook workbook = OeffneExcel())
                {
                    foreach (var eintrag in auswahl)
                    {
                        double? gewicht = HoleGewicht(workbook, eintrag.Name, eintrag.Blatt);

                        if (gewicht.HasValue)
                        {
                            gesamt += gewicht.Value;
                            listBox4.Items.Add(
                                eintrag.Anzeige + ": " + eintrag.Name +
                                " (" + gewicht.Value.ToString("0.0") + ")"
                            );
                        }
                        else
                        {
                            listBox4.Items.Add(
                                eintrag.Anzeige + ": " + eintrag.Name +
                                " (Gewicht nicht gefunden)"
                            );
                        }
                    }
                }

                listBox4.Items.Add("--------------------------------");
                listBox4.Items.Add("Gesamtgewicht: " + gesamt.ToString("0.0"));

                ZeigeGewichtsklasse(gesamt);
            }
            catch (Exception ex)
            {
                listBox4.Items.Add("Fehler beim Lesen der Excel-Datei:");
                listBox4.Items.Add(ex.Message);
            }
        }
        private static double? HoleGewicht(XLWorkbook workbook, string name, string? blattName)
        {
            foreach (var blatt in workbook.Worksheets)
            {
                if (blattName != null)
                {
                    if (blatt.Name != blattName)
                        continue;
                }
                else
                {
                    string kopf = blatt.Cell(3, 1).GetString().Trim();

                    if (!string.Equals(kopf, "Waffe", StringComparison.OrdinalIgnoreCase))
                        continue;
                }
                int letzteSpalte = blatt.Row(3).LastCellUsed()?.Address.ColumnNumber ?? 0;
                int gewichtSpalte = -1;

                for (int spalte = 1; spalte <= letzteSpalte; spalte++)
                {
                    if (blatt.Cell(3, spalte).GetString().Trim() == "Gewicht")
                    {
                        gewichtSpalte = spalte;
                        break;
                    }
                }

                if (gewichtSpalte == -1)
                    continue;

                int letzteZeile = blatt.LastRowUsed()?.RowNumber() ?? 3;

                for (int zeile = 4; zeile <= letzteZeile; zeile++)
                {
                    string zellName = blatt.Cell(zeile, 1).GetString().Trim();

                    if (NormalisiereName(zellName) != NormalisiereName(name))
                        continue;

                    if (blatt.Cell(zeile, gewichtSpalte).TryGetValue(out double gewicht))
                        return gewicht;

                    return null;
                }
            }

            return null;
        }
        private XLWorkbook OeffneExcel()
        {
            using (FileStream stream = new FileStream(excelDateiPfad, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                return new XLWorkbook(stream);
            }
        }
        private void StatEingabePruefen(object? sender, EventArgs e)
        {
            NumericUpDown? feld = sender as NumericUpDown;

            if (feld == null)
                return;

            int zahl;

            if (int.TryParse(feld.Text, out zahl))
            {
                if (zahl > 99)
                {
                    feld.Value = 99;
                    feld.Text = "99";

                    feld.Select(feld.Text.Length, 0);
                }
            }
        }

        private void StatsGeaendert(object? sender, EventArgs e)
        {
            if (comboBox2.SelectedIndex != -1)
            {
                string waffe = comboBox2.SelectedItem?.ToString() ?? "";
                LadeWaffeAusExcel(waffe, listBox1);
            }

            if (comboBox5.SelectedIndex != -1)
            {
                string waffe = comboBox5.SelectedItem?.ToString() ?? "";
                LadeWaffeAusExcel(waffe, listBox2);
            }

            if (favoritenFenster != null && !favoritenFenster.IsDisposed)
            {
                favoritenFenster.AktualisiereAnzeige();
            }
        }
        private void ZeigeWaffeMitStats(IXLWorksheet blatt, int zeile, ListBox zielListBox)
        {
            int staerke = (int)numericUpDown1.Value;
            int geschick = (int)numericUpDown2.Value;
            int intelligenz = (int)numericUpDown3.Value;
            int glaube = (int)numericUpDown4.Value;
            int arkan = (int)numericUpDown5.Value;

            double phys = HoleZahl(blatt, zeile, "Phys.");
            double nebenSchaden = HoleZahl(blatt, zeile, "Neben-Schaden");

            string skalStaerke = HoleText(blatt, zeile, "Skal. STR");
            string skalGeschick = HoleText(blatt, zeile, "Skal. DEX");
            string skalIntelligenz = HoleText(blatt, zeile, "Skal. INT");
            string skalGlaube = HoleText(blatt, zeile, "Skal. GLA");
            string skalArkan = HoleText(blatt, zeile, "Skal. ARK");

            int anfStaerke = (int)HoleZahl(blatt, zeile, "Anf. STR");
            int anfGeschick = (int)HoleZahl(blatt, zeile, "Anf. DEX");
            int anfIntelligenz = (int)HoleZahl(blatt, zeile, "Anf. INT");
            int anfGlaube = (int)HoleZahl(blatt, zeile, "Anf. GLA");
            int anfArkan = (int)HoleZahl(blatt, zeile, "Anf. ARK");

            zielListBox.Items.Add("--------------------------------");
            zielListBox.Items.Add("Mit deinen Stats (geschätzt):");

            bool alleErfuellt = true;

            if (staerke < anfStaerke)
            {
                zielListBox.Items.Add("STR zu tief: " + staerke + " von " + anfStaerke);
                alleErfuellt = false;
            }

            if (geschick < anfGeschick)
            {
                zielListBox.Items.Add("DEX zu tief: " + geschick + " von " + anfGeschick);
                alleErfuellt = false;
            }

            if (intelligenz < anfIntelligenz)
            {
                zielListBox.Items.Add("INT zu tief: " + intelligenz + " von " + anfIntelligenz);
                alleErfuellt = false;
            }

            if (glaube < anfGlaube)
            {
                zielListBox.Items.Add("GLA zu tief: " + glaube + " von " + anfGlaube);
                alleErfuellt = false;
            }

            if (arkan < anfArkan)
            {
                zielListBox.Items.Add("ARK zu tief: " + arkan + " von " + anfArkan);
                alleErfuellt = false;
            }

            if (alleErfuellt)
            {
                zielListBox.Items.Add("Anforderungen: alle erfüllt");
            }

            double bonusPhys = 0;
            bonusPhys = bonusPhys + phys * Skalierungsfaktor(skalStaerke) * StatKurve(staerke);
            bonusPhys = bonusPhys + phys * Skalierungsfaktor(skalGeschick) * StatKurve(geschick);

            double bonusNeben = 0;
            bonusNeben = bonusNeben + nebenSchaden * Skalierungsfaktor(skalIntelligenz) * StatKurve(intelligenz);
            bonusNeben = bonusNeben + nebenSchaden * Skalierungsfaktor(skalGlaube) * StatKurve(glaube);
            bonusNeben = bonusNeben + nebenSchaden * Skalierungsfaktor(skalArkan) * StatKurve(arkan);

            double gesamtPhys = phys + bonusPhys;
            double gesamtNeben = nebenSchaden + bonusNeben;
            double gesamt = gesamtPhys + gesamtNeben;

            if (!alleErfuellt)
            {
                gesamt = gesamt * 0.6;
            }

            zielListBox.Items.Add("Phys.: " + phys.ToString("0") + " + " + bonusPhys.ToString("0") + " = " + gesamtPhys.ToString("0"));

            if (nebenSchaden > 0)
            {
                zielListBox.Items.Add("Neben-Schaden: " + nebenSchaden.ToString("0") + " + " + bonusNeben.ToString("0") + " = " + gesamtNeben.ToString("0"));
            }

            if (!alleErfuellt)
            {
                zielListBox.Items.Add("Abzug, weil Anforderungen fehlen: -40%");
            }

            zielListBox.Items.Add("Angriffskraft gesamt: " + gesamt.ToString("0"));
        }

        private static double Skalierungsfaktor(string buchstabe)
        {
            switch (buchstabe.Trim().ToUpper())
            {
                case "S": return 1.8;
                case "A": return 1.5;
                case "B": return 1.1;
                case "C": return 0.75;
                case "D": return 0.4;
                case "E": return 0.15;
                default: return 0;
            }
        }

        private static double StatKurve(int stat)
        {
            double prozent;

            if (stat <= 1)
            {
                prozent = 0;
            }
            else if (stat <= 18)
            {
                prozent = 25.0 * (stat - 1) / 17;
            }
            else if (stat <= 60)
            {
                prozent = 25 + 50.0 * (stat - 18) / 42;
            }
            else if (stat <= 80)
            {
                prozent = 75 + 15.0 * (stat - 60) / 20;
            }
            else
            {
                prozent = 90 + 10.0 * (stat - 80) / 19;
            }

            return prozent / 100;
        }

        private static string HoleText(IXLWorksheet blatt, int zeile, string ueberschrift)
        {
            int letzteSpalte = blatt.Row(3).LastCellUsed()?.Address.ColumnNumber ?? 1;

            for (int spalte = 1; spalte <= letzteSpalte; spalte++)
            {
                if (blatt.Cell(3, spalte).GetString().Trim() == ueberschrift)
                {
                    return blatt.Cell(zeile, spalte).GetString().Trim();
                }
            }

            return "";
        }
        private static double HoleZahl(IXLWorksheet blatt, int zeile, string ueberschrift)
        {
            int letzteSpalte = blatt.Row(3).LastCellUsed()?.Address.ColumnNumber ?? 1;

            for (int spalte = 1; spalte <= letzteSpalte; spalte++)
            {
                if (blatt.Cell(3, spalte).GetString().Trim() == ueberschrift)
                {
                    double zahl;

                    if (blatt.Cell(zeile, spalte).TryGetValue(out zahl))
                    {
                        return zahl;
                    }

                    return 0;
                }
            }

            return 0;
        }

        private void AusdauerGeaendert(object? sender, EventArgs e)
        {
            AktualisiereGesamtgewicht();
        }

        private static double MaximalesGewicht(int ausdauer)
        {
            double maximum;

            if (ausdauer <= 8)
            {
                maximum = 45;
            }
            else if (ausdauer <= 25)
            {
                maximum = 45 + 27.0 * (ausdauer - 8) / 17;
            }
            else if (ausdauer <= 60)
            {
                double anteil = (ausdauer - 25) / 35.0;
                maximum = 72 + 48 * (1 - Math.Pow(1 - anteil, 1.2));
            }
            else
            {
                maximum = 120 + 40.0 * (ausdauer - 60) / 39;
            }

            return maximum;
        }

        private void ZeigeGewichtsklasse(double gesamt)
        {
            int ausdauer = (int)numericUpDown6.Value;
            double maximum = MaximalesGewicht(ausdauer);

            double prozent = gesamt / maximum * 100;

            string klasse;

            if (prozent < 30)
            {
                klasse = "Niedriges Gewicht (schnelle Ausweichrolle)";
            }
            else if (prozent < 70)
            {
                klasse = "Mittleres Gewicht (normale Ausweichrolle)";
            }
            else if (prozent <= 100)
            {
                klasse = "Hohes Gewicht (langsame Ausweichrolle)";
            }
            else
            {
                klasse = "Überladen (kein Ausweichen, nur gehen)";
            }

            listBox4.Items.Add("Max. Gewicht (Ausdauer " + ausdauer + "): " + maximum.ToString("0.0"));
            listBox4.Items.Add("Belastung: " + prozent.ToString("0.0") + " %");
            listBox4.Items.Add("--------------------------------");
            listBox4.Items.Add(klasse);
        }

        private static string NormalisiereName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return "";

            return name
                .Trim()
                .ToLowerInvariant()
                .Replace("ß", "ss")
                .Replace(" ", "")
                .Replace("-", "");
        }

        private void Form1_Paint(object? sender, PaintEventArgs e)
        {
            Design.MaleHintergrund(e.Graphics, this.ClientRectangle);
        }

        private void label1_Click(object sender, EventArgs e)
        {
        }

        private void label2_Click(object sender, EventArgs e)
        {
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
        }

        // ============================================================
        // STATS SPEICHERN
        // ============================================================

        private readonly string statsDateiPfad =
            Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "Elden Ring Datenbank",
                "stats.txt"
            );

        private void Form1_FormClosing(object? sender, FormClosingEventArgs e)
        {
            SpeichereStats();
        }

        private void SpeichereStats()
        {
            string[] zeilen =
            {
                numericUpDown1.Value.ToString(),
                numericUpDown2.Value.ToString(),
                numericUpDown3.Value.ToString(),
                numericUpDown4.Value.ToString(),
                numericUpDown5.Value.ToString(),
                numericUpDown6.Value.ToString()
            };

            try
            {
                string? ordner = Path.GetDirectoryName(statsDateiPfad);

                if (ordner != null)
                {
                    Directory.CreateDirectory(ordner);
                }

                File.WriteAllLines(statsDateiPfad, zeilen);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Stats konnten nicht gespeichert werden:\n\n" + ex.Message);
            }
        }

        private void LadeStats()
        {
            if (!File.Exists(statsDateiPfad))
                return;

            try
            {
                string[] zeilen = File.ReadAllLines(statsDateiPfad);

                if (zeilen.Length < 6)
                    return;

                numericUpDown1.Value = StatAusText(zeilen[0]);
                numericUpDown2.Value = StatAusText(zeilen[1]);
                numericUpDown3.Value = StatAusText(zeilen[2]);
                numericUpDown4.Value = StatAusText(zeilen[3]);
                numericUpDown5.Value = StatAusText(zeilen[4]);
                numericUpDown6.Value = StatAusText(zeilen[5]);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Stats konnten nicht geladen werden:\n\n" + ex.Message);
            }
        }        private static decimal StatAusText(string text)
        {
            int zahl;

            if (!int.TryParse(text.Trim(), out zahl))
                return 10;

            if (zahl < 1)
                zahl = 1;

            if (zahl > 99)
                zahl = 99;

            return zahl;
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            
            this.Close();
        }

        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (comboBox2.SelectedIndex == -1)
            {
                MessageBox.Show("Bitte zuerst eine Nahkampfwaffe auswählen.");
                return;
            }

            string kategorie = comboBox1.Text;
            string waffe = comboBox2.SelectedItem?.ToString() ?? "";

            SpeichereFavorit("Nahkampf", kategorie, waffe);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (comboBox5.SelectedIndex == -1)
            {
                MessageBox.Show("Bitte zuerst eine Fernkampfwaffe oder einen Katalysator auswählen.");
                return;
            }

            string art = comboBox3.Text;       
            string kategorie = comboBox4.Text; 
            string waffe = comboBox5.SelectedItem?.ToString() ?? "";

            SpeichereFavorit(art, kategorie, waffe);
        }

        private void SpeichereFavorit(string art, string kategorie, string waffe)
        {
            try
            {
                bool neu = Favoriten.Hinzufuegen(art, kategorie, waffe);

                if (neu)
                {
                    MessageBox.Show(waffe + " wurde zu den Favoriten hinzugefügt.", "Favoriten");
                }
                else
                {
                    MessageBox.Show(waffe + " ist schon in den Favoriten.", "Favoriten");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Favorit konnte nicht gespeichert werden:\n\n" + ex.Message);
            }
        }
        private void button4_Click(object sender, EventArgs e)
        {
            if (favoritenFenster != null && !favoritenFenster.IsDisposed)
            {
                favoritenFenster.Activate();
                return;
            }

            favoritenFenster = new FavoritenForm(this);
            favoritenFenster.Show(this);
        }
        public void ZeigeWaffe(string waffe, ListBox zielListBox)
        {
            LadeWaffeAusExcel(waffe, zielListBox);
        }

        private void listBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
        }
    }
}

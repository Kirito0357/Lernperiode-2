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

        public Form1()
        {
            InitializeComponent();

            // --------------------------------------------------
            // Grundkonfiguration
            // --------------------------------------------------

            button1.Visible = false;

            comboBox1.Visible = true;
            comboBox2.Visible = false;

            comboBox3.Visible = true;
            comboBox4.Visible = false;
            comboBox5.Visible = false;

            comboBox6.Visible = true;
            comboBox7.Visible = false;

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

            this.BackColor = Color.FromArgb(25, 25, 25);

            label1.ForeColor = Color.Goldenrod;
            label2.ForeColor = Color.Goldenrod;
            label3.ForeColor = Color.Goldenrod;
            label4.ForeColor = Color.Goldenrod;

            comboBox1.ForeColor = Color.Goldenrod;
            comboBox1.BackColor = Color.FromArgb(40, 40, 40);
            comboBox2.ForeColor = Color.Goldenrod;
            comboBox2.BackColor = Color.FromArgb(40, 40, 40);
            comboBox3.ForeColor = Color.Goldenrod;
            comboBox3.BackColor = Color.FromArgb(40, 40, 40);
            comboBox4.ForeColor = Color.Goldenrod;
            comboBox4.BackColor = Color.FromArgb(40, 40, 40);
            comboBox5.ForeColor = Color.Goldenrod;
            comboBox5.BackColor = Color.FromArgb(40, 40, 40);
            comboBox6.ForeColor = Color.Goldenrod;
            comboBox6.BackColor = Color.FromArgb(40, 40, 40);
            comboBox7.ForeColor = Color.Goldenrod;
            comboBox7.BackColor = Color.FromArgb(40, 40, 40);

            listBox1.ForeColor = Color.Goldenrod;
            listBox1.BackColor = Color.FromArgb(40, 40, 40);
            listBox2.ForeColor = Color.Goldenrod;
            listBox2.BackColor = Color.FromArgb(40, 40, 40);
            listBox3.ForeColor = Color.Goldenrod;
            listBox3.BackColor = Color.FromArgb(40, 40, 40);
            listBox4.ForeColor = Color.Goldenrod;
            listBox4.BackColor = Color.FromArgb(40, 40, 40);
            listBox5.ForeColor = Color.Goldenrod;
            listBox5.BackColor = Color.FromArgb(40, 40, 40);
            listBox6.ForeColor = Color.Goldenrod;
            listBox6.BackColor = Color.FromArgb(40, 40, 40);
            listBox7.ForeColor = Color.Goldenrod;
            listBox7.BackColor = Color.FromArgb(40, 40, 40);

            button1.ForeColor = Color.Goldenrod;
            button1.BackColor = Color.FromArgb(40, 40, 40);

            comboBox5.SelectedIndexChanged += comboBox5_SelectedIndexChanged;
            comboBox7.SelectedIndexChanged += comboBox7_SelectedIndexChanged;

            // Stats: Farben
            label5.ForeColor = Color.Goldenrod;
            label6.ForeColor = Color.Goldenrod;
            label7.ForeColor = Color.Goldenrod;
            label8.ForeColor = Color.Goldenrod;
            label9.ForeColor = Color.Goldenrod;
            label10.ForeColor = Color.Goldenrod;

            numericUpDown1.ForeColor = Color.Goldenrod;
            numericUpDown1.BackColor = Color.FromArgb(40, 40, 40);
            numericUpDown2.ForeColor = Color.Goldenrod;
            numericUpDown2.BackColor = Color.FromArgb(40, 40, 40);
            numericUpDown3.ForeColor = Color.Goldenrod;
            numericUpDown3.BackColor = Color.FromArgb(40, 40, 40);
            numericUpDown4.ForeColor = Color.Goldenrod;
            numericUpDown4.BackColor = Color.FromArgb(40, 40, 40);
            numericUpDown5.ForeColor = Color.Goldenrod;
            numericUpDown5.BackColor = Color.FromArgb(40, 40, 40);

            // Stats: Wenn ein Wert geändert wird, die Waffen neu berechnen
            numericUpDown1.ValueChanged += StatsGeaendert;
            numericUpDown2.ValueChanged += StatsGeaendert;
            numericUpDown3.ValueChanged += StatsGeaendert;
            numericUpDown4.ValueChanged += StatsGeaendert;
            numericUpDown5.ValueChanged += StatsGeaendert;

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
            comboBox6.Items.AddRange(new string[]
           {
                "Kopfbedeckung",
                "Brustschutz",
                "Armpanzerung",
                "Beinschutz"
           });

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

                            zielListBox.Items.Add("================================");
                            zielListBox.Items.Add("Waffe: " + name);
                            zielListBox.Items.Add("================================");

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

                                // Kriegsasche wurde oben schon ausgegeben
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

        // ============================================================
        // RÜSTUNGEN
        // ============================================================
        private static string RuestungsBlatt(string kategorie)
        {
            switch (kategorie)
            {
                case "Kopfbedeckung": return "Kopfbedeckungen";
                case "Brustschutz": return "Torsos";
                case "Armpanzerung": return "Armschützer";
                case "Beinschutz": return "Beinschützer";
                default: return "";
            }
        }

        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox7.Items.Clear();
            comboBox7.SelectedIndex = -1;
            comboBox7.Text = "";

            if (comboBox6.SelectedIndex == -1)
            {
                comboBox7.Visible = false;
                return;
            }

            comboBox7.Visible = true;

            string blattName = RuestungsBlatt(comboBox6.Text);

            if (blattName == "")
                return;

            if (!File.Exists(excelDateiPfad))
            {
                MessageBox.Show(
                    "Die Excel-Datei wurde nicht gefunden.\n\nGesuchter Pfad:\n" + excelDateiPfad,
                    "Excel-Datei fehlt",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
                return;
            }

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
                            comboBox7.Items.Add(name);
                    }
                }

                if (gewaehlteRuestung.TryGetValue(blattName, out string? gemerkt) &&
                    comboBox7.Items.Contains(gemerkt))
                {
                    comboBox7.SelectedItem = gemerkt;
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

        private void comboBox7_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox7.SelectedIndex == -1)
                return;

            string ruestung = comboBox7.SelectedItem?.ToString() ?? "";
            string blattName = RuestungsBlatt(comboBox6.Text);

            ListBox? ziel = ListBoxFuerRuestung(blattName);

            if (ziel == null)
                return;

            LadeRuestungAusExcel(ruestung, blattName, ziel);

            if (blattName != "")
                gewaehlteRuestung[blattName] = ruestung;

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

                        zielListBox.Items.Add("================================");
                        zielListBox.Items.Add(blattName + ": " + name);
                        zielListBox.Items.Add("================================");

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
            ziel.Items.Add("================================");
            ziel.Items.Add(blattName);
            ziel.Items.Add("================================");
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

            listBox4.Items.Add("================================");
            listBox4.Items.Add("Gewicht der Ausrüstung");
            listBox4.Items.Add("================================");

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

        // ============================================================
        // STATS
        // ============================================================

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

        private void label1_Click(object sender, EventArgs e)
        {
        }

        private void label2_Click(object sender, EventArgs e)
        {
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
        }

        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
        }

        private void listBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
        }
    }
}

using ClosedXML.Excel;
using System;
using System.IO;
using System.Windows.Forms;

namespace Elden_Ring
{
    public partial class Form1 : Form
    {
        private readonly string excelDateiPfad =
            Path.Combine(
                AppContext.BaseDirectory,
                "Elden_Ring_Waffen_Ruestungen_Datenbank.xlsx"
            );

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
            listBox3.Visible = false;

            listBox1.HorizontalScrollbar = true;
            listBox1.SelectionMode = SelectionMode.None;

            listBox2.HorizontalScrollbar = true;
            listBox2.SelectionMode = SelectionMode.None;

            listBox3.HorizontalScrollbar = true;
            listBox3.SelectionMode = SelectionMode.None;

            // Eventhandler
            comboBox1.SelectedIndexChanged += comboBox1_SelectedIndexChanged;
            comboBox2.SelectedIndexChanged += comboBox2_SelectedIndexChanged;
            comboBox3.SelectedIndexChanged += comboBox3_SelectedIndexChanged;
            comboBox4.SelectedIndexChanged += comboBox4_SelectedIndexChanged;
            comboBox5.SelectedIndexChanged += comboBox5_SelectedIndexChanged;

            // Zuerst entfernen, falls der Designer den Handler schon verbunden hat,
            // damit er nicht doppelt ausgelöst wird.
            comboBox6.SelectedIndexChanged -= comboBox6_SelectedIndexChanged;
            comboBox6.SelectedIndexChanged += comboBox6_SelectedIndexChanged;
            comboBox7.SelectedIndexChanged -= comboBox7_SelectedIndexChanged;
            comboBox7.SelectedIndexChanged += comboBox7_SelectedIndexChanged;

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
                        "Blutbeflekter Dolch", "Erdenstahldolch", "Wakizashi",
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
                        "Kriegsfallenkralle", "Lazuli-Schimmersteinschwert",
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
                        "Flamberge", "Gargoile-Grossschwert", "Gargoile-Schwarzschwert",
                        "Unzertrennliches Schwert", "Milos-Schwert",
                        "Henkerschwert der Marais", "Ordovis'Grossschwert",
                        "Alabasterfürstenschwert", "Schülereisen des Todes",
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
                        "Trümmergrossschwert", "Grossschwert der Sternengeissel",
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
                        "Banditenkrummschwert", "Shotel", "Plünderungskrummschwert",
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
                        "Gargoyle-Grossaxt", "Gargoyle-Schwarzakt",
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
                        "Reisenbrecher", "Infernaler Prälatenstab", "Grosskeule",
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
                return;

            string waffe = comboBox2.SelectedItem?.ToString() ?? "";

            // Nahkampf-Waffen werden in listBox1 ausgegeben
            LadeWaffeAusExcel(waffe, listBox1);
            listBox1.Visible = true;
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
                using (var workbook = new XLWorkbook(excelDateiPfad))
                {
                    foreach (var blatt in workbook.Worksheets)
                    {
                        // Prüfen, ob es sich um ein Waffenblatt handelt.
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

                            // Waffenname
                            zielListBox.Items.Add("================================");
                            zielListBox.Items.Add("Waffe: " + name);
                            zielListBox.Items.Add("================================");


                            for (int spalte = 2; spalte <= 17; spalte++)
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
                        "Stab der Schuldigen", "Stab des Topesprinzen"
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

        // ============================================================
        // FERNKAMPF-WAFFE AUS EXCEL
        // ============================================================

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            listBox2.Items.Clear();

            if (comboBox5.SelectedIndex == -1)
                return;

            string waffe = comboBox5.SelectedItem?.ToString() ?? "";

            // Fernkampf/Katalysatoren werden in listBox2 ausgegeben
            LadeWaffeAusExcel(waffe, listBox2);
            listBox2.Visible = true;
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

        // ============================================================
        // RÜSTUNGEN
        // ============================================================

        // Ordnet den Einträgen in comboBox6 das passende Blatt in der Excel-Datei zu.
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

            listBox3.Items.Clear();

            if (comboBox6.SelectedIndex == -1)
            {
                comboBox7.Visible = false;
                return;
            }

            comboBox7.Visible = true;

            // Die Rüstungsteile werden direkt aus dem passenden Excel-Blatt gelesen,
            // so muss man die rund 570 Namen nicht von Hand eintippen.
            string blattName = RuestungsBlatt(comboBox6.Text);

            if (!File.Exists(excelDateiPfad) || blattName == "")
                return;

            try
            {
                using (var workbook = new XLWorkbook(excelDateiPfad))
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
            listBox3.Items.Clear();

            if (comboBox7.SelectedIndex == -1)
                return;

            string ruestung = comboBox7.SelectedItem?.ToString() ?? "";

            // Rüstungen werden in listBox3 ausgegeben
            LadeRuestungAusExcel(ruestung, RuestungsBlatt(comboBox6.Text), listBox3);
            listBox3.Visible = true;
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
                using (var workbook = new XLWorkbook(excelDateiPfad))
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

                        // Name des Rüstungsteils
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

        private void listBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
        }
    }
}

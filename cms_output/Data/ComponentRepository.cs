using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using CMS.Models;
using ClosedXML.Excel;

namespace CMS.Data
{
    public class CategoryField
    {
        public string Key   { get; set; } = "";
        public string Label { get; set; } = "";
    }

    public static class ComponentRepository
    {
        private static readonly string DataFile = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "CMS", "components.json");

        private static List<Component> _components = new();
        private static int _nextId = 1;

        public static Dictionary<string, List<(string Key, string Label)>> CategoryFields = new()
        {
            ["RES"]     = new() { ("Resistance(Ohm)", "Resistance"), ("Power(W)", "Power"), ("Tolerance(%)", "Tolerance") },
            ["CAP"]     = new() { ("Capacitance(F)", "Capacitance"), ("Voltage(V)", "Voltage"), ("Tolerance(%)", "Tolerance"), ("Temperature Coefficient", "Temp Coeff"), ("Failure Mode", "Failure Mode") },
            ["IND"]     = new() { ("Inductance(H)", "Inductance"), ("Irms Current(A)", "Irms Current"), ("Tolerance(%)", "Tolerance"), ("Max DCR(Ohm)", "Max DCR") },
            ["FB"]      = new() { ("Impedance(Ohm)@Frequency(Hz)", "Impedance@Freq"), ("Max Rated Current(A)", "Max Current"), ("Max DCR(Ohm)", "Max DCR") },
            ["CMC"]     = new() { ("Inductance(H)@Frequency(Hz)", "Inductance@Freq"), ("Impedance(Ohm)@Frequency(Hz)", "Impedance@Freq"), ("Max Rated Current(A)", "Max Current"), ("Max DCR(Ohm)", "Max DCR") },
            ["DIO-REC"] = new() { ("Max Peak Reverse Voltage(V)", "Reverse Voltage"), ("Max Forward Voltage(V)", "Forward Voltage"), ("Max Forward Current(A)", "Forward Current") },
            ["DIO-SW"]  = new() { ("Max Peak Reverse Voltage(V)", "Reverse Voltage"), ("Max Forward Voltage(V)", "Forward Voltage"), ("Max Forward Current(A)", "Forward Current") },
            ["DIO-ZEN"] = new() { ("Nom Zener Voltage(V)", "Zener Voltage"), ("Tolerance(%)", "Tolerance") },
            ["DIO-SCH"] = new() { ("Max Reverse Voltage(V)", "Reverse Voltage"), ("Max Forward Voltage(V)", "Forward Voltage"), ("Avg Forward Current(V)", "Avg Current") },
            ["DIO-TVS"] = new() { ("Reverse Standoff Voltage(V)", "Standoff Voltage"), ("Peak Pulse Power(W)", "Peak Power"), ("Max Peak Pulse Current(A)", "Peak Current"), ("Polarization", "Polarization") },
            ["NTC"]     = new() { ("Resistance@25°C(Ohm)", "R@25°C"), ("Resistance Tolerance(%)", "Tolerance"), ("Beta(K)", "Beta"), ("Beta Tolerance(%)", "Beta Tol.") },
            ["PTC"]     = new() { ("Resistance@25°C(Ohm)", "R@25°C"), ("Percentage R. Tolerance(%)", "Tolerance") },
            ["LED"]     = new() { ("Nom Forward Voltage(V)", "Forward Voltage"), ("Nom Forward Current(A)", "Forward Current"), ("BIN", "BIN") },
            ["TRAN"]    = new() { ("Type", "Type (NPN/PNP)"), ("Max Collector-Emitter Voltage(V)", "Vce Max"), ("Max Collector Current(I)", "Ic Max"), ("Max Base-Emitter Voltage(V)", "Vbe Max"), ("Power(W)", "Power") },
            ["MOS"]     = new() { ("Type", "Type (N-Ch/P-Ch)"), ("Max Drain-Source Voltage(V)", "Vds Max"), ("Max Drain Current(I)", "Id Max"), ("Max Gate-Source Voltage(V)", "Vgs Max"), ("Power(W)", "Power") },
            ["DC-DC"]   = new() { ("Type", "Converter Type"), ("Max Output Voltage(V)", "Output Voltage"), ("Max Output Current(I)", "Output Current") },
            ["LDO"]     = new() { ("Output Type", "Output Type"), ("Max Output Voltage(V)", "Output Voltage"), ("Max Output Current(I)", "Output Current") },
            ["LED-DRV"] = new() { ("Type", "Driver Type"), ("Total Channel Count", "Channels"), ("Output Current/Channel(A)", "Current/Ch") },
            ["MCU"]     = new() { ("Core Processor", "Core"), ("Core Size(Bit)", "Core Size"), ("Clock Speed(Hz)", "Clock Speed"), ("Program Memory Size(B)", "Flash Size"), ("Program Memory Type", "Memory Type"), ("ASIL Category", "ASIL") },
            ["OTHER-IC"]= new() { },
            ["CON"]     = new() { ("Gender", "Gender"), ("Pin Count", "Pin Count"), ("Pitch(mm)", "Pitch"), ("Color", "Color") },
            ["TERM"]    = new() { ("Gender", "Gender"), ("Base Material", "Base Material"), ("Plating Material", "Plating") },
            ["PAD"]     = new() { ("PAD X-Size(mm)", "PAD X"), ("PAD Y-Size(mm)", "PAD Y"), ("Hole X-Size(mm)", "Hole X"), ("Hole Y-Size(mm)", "Hole Y"), ("PAD Shape", "PAD Shape"), ("Hole Shape", "Hole Shape") },
        };

private static readonly Dictionary<string, List<(string Key, string Label)>> _builtInCategories =
    new(CategoryFields.ToDictionary(k => k.Key, k => new List<(string, string)>(k.Value)));
    // ── Custom Columns ─────────────────────────────────────────────────────
public static List<string> CustomColumns { get; private set; } = new();

public static void AddCustomColumn(string name)
{
    if (!string.IsNullOrWhiteSpace(name) && !CustomColumns.Contains(name))
        CustomColumns.Add(name);
}

public static void RemoveCustomColumn(string name)
{
    CustomColumns.Remove(name);
    foreach (var c in _components)
        c.ExtraFields.Remove(name);
    Save();
}
        public static List<string> AllCategories => CategoryFields.Keys.OrderBy(k => k).ToList();

        public static readonly List<string> AllManufacturers = new()
        {
            "Analog Devices", "Bourns", "Bourns Inc.", "Diodes", "Diotec",
            "Diotec Semiconductor", "Eaton", "Elmos", "Infineon", "KOA",
            "Kemet", "Kyocera", "Littelfuse", "Lumissil", "MCC", "Microchip",
            "Molex", "Murata", "NXP", "Nexperia", "Nichia", "Onsemi", "Osram",
            "Panasonic", "Panjit International Inc.", "Pulse", "Rohm", "STM",
            "STMicroelectronics", "Samsung", "Stackpole", "Stocko", "TDK",
            "Taiwan Semiconductor", "Taiyo Yuden", "Texas Instruments", "Vishay",
            "Würth", "Yageo", "Yangjie"
        };

        public static readonly List<string> AllFootprints = new()
        {
            "4 PC Pad(ACT45B)", "5023520400-Footprint-1", "CAP-0402", "CAP-0603", "CAP-0805",
            "CAP-1206", "CAP-1210", "D2PAK-3_Ftprnt", "DFN1010D-3 (SOT1215)",
            "DO-214AA(SMB)", "DO-214AB(SMC)", "DO-214AC(SMA)", "DO-219AB(SMF)",
            "ETSSOP-16", "FB-0402", "FB-0603", "FB-0805", "FB-1206", "FB-1206(4-Array)",
            "FCL_Cable_Pad_1X4", "FCL_Terminal_1X4", "HTSSOP-28",
            "IND-12.5x12.5x6(Shielded)", "IND-7.3x6.6x2.8(Shielded)", "IND-7.4x7x4.8(Shielded)",
            "KY DDLN31.FY LED", "LA G6SP.02 LED", "LA.A67F", "LQFP-48",
            "LS E67F_Fprt", "LW.G6SP LED", "LY.G6SP LED", "MFMP 7238 - 202 - W60",
            "NTC-0805", "PG-TSDSO-14", "PowerDI3333-8",
            "RES-0402", "RES-0603", "RES-0805", "RES-1206",
            "Radial, Can", "SC-70", "SMA-FL", "SO-8EP",
            "SOD-123", "SOD-123F", "SOD-123FL", "SOD-123W",
            "SOD-323", "SOD-323EP", "SOD-323F", "SOD-523", "SOD523",
            "SOIC-8", "SOIC16N-EP", "SOT-23", "SOT-26", "SOT-563", "SOT-89",
            "SOT223", "SOT23-3", "SOT403-1", "SOT89", "SOT−363",
            "TFDSO-16", "TQFN-EP(24-SW)", "TSSOP-16",
            "U-DFN2020-6 (Type F)", "WDFN-8", "eTSSOP-16_Fprt"
        };

        public static readonly List<string> AllLibraryRefs = new()
        {
            "AP64200Q_Symbol", "BCR430UW6_Symbol", "CMC_Symbol", "CON Male 2Pin 2.5mm",
            "CapacitorSymbol", "DRV8825PWP_Symbol",
            "Diode_Common_Cathode_Symbol_3pin", "Diode_Schottky_Symbol",
            "Diode_Schottky_Symbol_CathodeLeft", "Diode_Single_Symbol_3pin",
            "Diode_Symbol", "Diode_Symbol_CathodeLeft", "Diode_TVS_Bidirectional_Symbol",
            "Diode_TVS_Unidirectional_Symbol_CathodeLeft",
            "Diode_Zener_Common_Anode_Symbol_3pin", "Diode_Zener_Single_2pin",
            "Diode_Zener_Single_Symbol_3pin", "E522.80_Symbol",
            "FCL_Cable_Pad_1X4", "FCL_Terminal_1X4",
            "FerriteBead(4-Array)Symbol", "FerriteBeadSymbol",
            "IS32LT3144A_Symbol", "IS32LT3152_Symbol", "InductorSymbol",
            "LED_Symbol", "Molex-4Pin", "NCV317BD2TR4G_Symboll",
            "NPN_Transistor_Symbol", "PIC16F15213-E/SN_Symbol",
            "PNP_Transistor_Symbol", "ResistorSymbol",
            "STM32F103C8T6_Symbol", "TLD2142-1EP_Symbol",
            "TLD2382-3ET_Symbol", "TMP112_Symbol"
        };

        static ComponentRepository() { }

        public static Task InitializeAsync() => Task.Run(() => Load());

        public static IReadOnlyList<Component> GetAll() => _components.AsReadOnly();
        public static IReadOnlyList<Component> GetCustom() => _components.Where(c => c.IsCustom).ToList().AsReadOnly();

        public static void Add(Component c) { c.Id = _nextId++; _components.Add(c); Save(); }
        public static void Update(Component c) { Save(); }
        public static void Delete(int id) { _components.RemoveAll(c => c.Id == id); Save(); }

        public static void Approve(int id)
        {
            var c = _components.FirstOrDefault(x => x.Id == id);
            if (c != null) { c.Status = ComponentStatus.Approved; Save(); }
        }

        public static List<Component> Search(string query, string category,
            string rating, string process, string manufacturer)
        {
            var q = query?.ToLower() ?? "";
            return _components.Where(c =>
            {
                if (!string.IsNullOrEmpty(category) && c.Category != category) return false;
                if (!string.IsNullOrEmpty(rating) && c.Rating.ToString() != rating) return false;
                if (!string.IsNullOrEmpty(process) && c.Process.ToString() != process) return false;
                if (!string.IsNullOrEmpty(manufacturer) && c.Manufacturer != manufacturer) return false;
                if (!string.IsNullOrEmpty(q))
                {
                    var text = $"{c.Description} {c.PartNumber} {c.Manufacturer} {c.Category} {c.Footprint}".ToLower();
                    if (!text.Contains(q)) return false;
                }
                return true;
            }).ToList();
        }

        public static Dictionary<string, int> GetCategoryCounts() =>
            _components.GroupBy(c => c.Category).ToDictionary(g => g.Key, g => g.Count());

        public static Dictionary<string, int> GetManufacturerCounts() =>
            _components.Where(c => !string.IsNullOrEmpty(c.Manufacturer))
                       .GroupBy(c => c.Manufacturer)
                       .OrderByDescending(g => g.Count())
                       .Take(10)
                       .ToDictionary(g => g.Key, g => g.Count());

        public static List<string> GetManufacturers() =>
            _components.Select(c => c.Manufacturer)
                       .Where(m => !string.IsNullOrEmpty(m))
                       .Distinct().OrderBy(m => m).ToList();

        public static List<string> GetFootprints() =>
            _components.Select(c => c.Footprint)
                       .Where(f => !string.IsNullOrEmpty(f))
                       .Distinct().OrderBy(f => f).ToList();

        public static List<string> GetLibraryRefs() =>
            _components.Select(c => c.LibraryRef)
                       .Where(r => !string.IsNullOrEmpty(r))
                       .Distinct().OrderBy(r => r).ToList();

        public static List<string> GetFootprintRefs() =>
            _components.Select(c => c.FootprintRef)
                       .Where(r => !string.IsNullOrEmpty(r))
                       .Distinct().OrderBy(r => r).ToList();

        public static List<string> GetPartNumbers() =>
            _components.Select(c => c.PartNumber)
                       .Where(p => !string.IsNullOrEmpty(p))
                       .Distinct().OrderBy(p => p).ToList();

        public static List<string> GetFootprintsByCategory(string category) =>
            _components.Where(c => c.Category == category)
                       .Select(c => c.Footprint)
                       .Where(f => !string.IsNullOrEmpty(f))
                       .Distinct().OrderBy(f => f).ToList();

        public static List<string> GetManufacturersByCategory(string category) =>
            _components.Where(c => c.Category == category)
                       .Select(c => c.Manufacturer)
                       .Where(m => !string.IsNullOrEmpty(m))
                       .Distinct().OrderBy(m => m).ToList();

        public static List<string> GetLibraryRefsByCategory(string category) =>
            _components.Where(c => c.Category == category)
                       .Select(c => c.LibraryRef)
                       .Where(r => !string.IsNullOrEmpty(r))
                       .Distinct().OrderBy(r => r).ToList();

        public static List<string> GetFootprintRefsByCategory(string category) =>
            _components.Where(c => c.Category == category)
                       .Select(c => c.FootprintRef)
                       .Where(r => !string.IsNullOrEmpty(r))
                       .Distinct().OrderBy(r => r).ToList();

        

        public static void AddCategory(string categoryName)
        {
            if (string.IsNullOrWhiteSpace(categoryName)) return;
            if (!CategoryFields.ContainsKey(categoryName))
            {
                CategoryFields[categoryName] = new List<(string Key, string Label)>();
                Save();
            }
        }

        public static void DeleteCategory(string categoryName)
        {
            if (CategoryFields.ContainsKey(categoryName))
            {
                CategoryFields.Remove(categoryName);
                Save();
            }
        }

        public static void AddFieldToCategory(string categoryName, string key, string label)
        {
            if (string.IsNullOrWhiteSpace(categoryName) || string.IsNullOrWhiteSpace(key)) return;
            if (!CategoryFields.ContainsKey(categoryName))
                CategoryFields[categoryName] = new();

            if (!CategoryFields[categoryName].Any(f => f.Key == key))
            {
                CategoryFields[categoryName].Add((key, label));
                Save();
            }
        }

        public static void RemoveFieldFromCategory(string categoryName, string key)
        {
            if (CategoryFields.ContainsKey(categoryName))
            {
                CategoryFields[categoryName].RemoveAll(f => f.Key == key);
                Save();
            }
        }

        public static void UpdateFieldInCategory(string categoryName, string oldKey, string newKey, string newLabel)
        {
            if (!CategoryFields.ContainsKey(categoryName)) return;
            var fields = CategoryFields[categoryName];
            int idx = fields.FindIndex(f => f.Key == oldKey);
            if (idx >= 0)
            {
                fields[idx] = (newKey, newLabel);
                Save();
            }
        }


        private static void Load()
        {
            try
            {
                if (File.Exists(DataFile))
                {
                    var json = File.ReadAllText(DataFile);
                    var data = JsonSerializer.Deserialize<SaveData>(json);
                    if (data != null && data.Components != null && data.Components.Count > 0)
                    {
                        _components = data.Components;
                        
                        _nextId = data.NextId;

                        if (data.CustomCategories != null)
                        {
                            foreach (var kvp in data.CustomCategories)
                            {
                                CategoryFields[kvp.Key] = kvp.Value
                                    .Select(f => (f.Key, f.Label))
                                    .ToList();
                            }
                        }
                        foreach (var kvp in _builtInCategories)
{
    if (!CategoryFields.ContainsKey(kvp.Key))
        CategoryFields[kvp.Key] = new List<(string, string)>(kvp.Value);
}
                        return;
                    }
                    File.Delete(DataFile);
                }
            }
            catch { _components = new(); }

            SeedBuiltIn();
        }

        private static void SeedBuiltIn()
        {
            _components.Clear();
            _nextId = 1;

            void Add(string fekaNo, string cat, string desc, string partNo, string mfr,
                     string footprint, string libRef, string footprintRef,
                     ComponentRating rating, ComponentProcess proc,
                     Dictionary<string, string>? extra = null)
            {
                _components.Add(new Component
                {
                    Id = _nextId++,
                    FekaNo = fekaNo,
                    Category = cat,
                    Description = desc,
                    PartNumber = partNo,
                    Manufacturer = mfr,
                    Footprint = footprint,
                    LibraryRef = libRef,
                    FootprintRef = footprintRef,
                    Rating = rating,
                    Process = proc,
                    Status = ComponentStatus.Approved,
                    IsCustom = false,
                    ExtraFields = extra ?? new()
                });
            }

            var SMD  = ComponentProcess.SMD;
            var THT  = ComponentProcess.THT;
            var AECQ = ComponentRating.AECQ;
            var STD  = ComponentRating.STD;

            // RESISTORS
            Add("4E12000001","RES","RES 100Ohm 0.5W 1% 1206 AECQ","AC1206FR-7W100RL","Yageo","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="100Ohm", ["Power(W)"]="0.5W", ["Tolerance(%)"]="1%"});
            Add("4E12000001","RES","RES 100Ohm 0.5W 1% 1206 AECQ","HP06W2F1000T5E","Rohm","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="100Ohm", ["Power(W)"]="0.5W", ["Tolerance(%)"]="1%"});
            Add("4E12000002","RES","RES 110Ohm 0.5W 1% 1206 STD","HP06W2F1100T5E","Rohm","RES-1206","ResistorSymbol","RES-1206",STD,SMD,new(){["Resistance(Ohm)"]="110Ohm", ["Power(W)"]="0.5W", ["Tolerance(%)"]="1%"});
            Add("4E12000003","RES","RES 120Ohm 0.5W 1% 1206 AECQ","HP06W2F1200T5E","Rohm","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="120Ohm", ["Power(W)"]="0.5W", ["Tolerance(%)"]="1%"});
            Add("4E12000004","RES","RES 130Ohm 0.5W 1% 1206 AECQ","AC1206FR-7W130RL","Yageo","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="130Ohm", ["Power(W)"]="0.5W", ["Tolerance(%)"]="1%"});
            Add("4E12000005","RES","RES 150Ohm 0.5W 1% 1206 AECQ","AC1206FR-7W150RL","Yageo","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="150Ohm", ["Power(W)"]="0.5W", ["Tolerance(%)"]="1%"});
            Add("4E12000005","RES","RES 150Ohm 0.5W 1% 1206 AECQ","HP06W2F1500T5E","Rohm","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="150Ohm", ["Power(W)"]="0.5W", ["Tolerance(%)"]="1%"});
            Add("4E12000006","RES","RES 160Ohm 0.5W 1% 1206 AECQ","AC1206FR-7W160RL","Yageo","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="160Ohm", ["Power(W)"]="0.5W", ["Tolerance(%)"]="1%"});
            Add("4E12000006","RES","RES 160Ohm 0.5W 1% 1206 AECQ","HP06W2F1600T5E","Rohm","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="160Ohm", ["Power(W)"]="0.5W", ["Tolerance(%)"]="1%"});
            Add("4E12000006","RES","RES 160Ohm 0.5W 1% 1206 AECQ","RMCP1206FT160R","Stackpole","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="160Ohm", ["Power(W)"]="0.5W", ["Tolerance(%)"]="1%"});
            Add("4E12000007","RES","RES 180Ohm 0.25W 1% 1206 AECQ","AC1206FR-07180RL","Yageo","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="180Ohm", ["Power(W)"]="0.25W", ["Tolerance(%)"]="1%"});
            Add("4E12000007","RES","RES 180Ohm 0.5W 1% 1206 AECQ","HP06W2F1800T5E","Rohm","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="180Ohm", ["Power(W)"]="0.5W", ["Tolerance(%)"]="1%"});
            Add("4E12000008","RES","RES 200Ohm 0.25W 1% 1206 AECQ","AC1206FR-07200RL","Yageo","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="200Ohm", ["Power(W)"]="0.25W", ["Tolerance(%)"]="1%"});
            Add("4E12000008","RES","RES 200Ohm 0.5W 1% 1206 AECQ","HP06W2F2000T5E","Rohm","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="200Ohm", ["Power(W)"]="0.5W", ["Tolerance(%)"]="1%"});
            Add("4E12000009","RES","RES 220Ohm 0.25W 1% 1206 AECQ","AC1206FR-07220RL","Yageo","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="220Ohm", ["Power(W)"]="0.25W", ["Tolerance(%)"]="1%"});
            Add("4E12000009","RES","RES 220Ohm 0.5W 1% 1206 AECQ","HP06W2F2200T5E","Rohm","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="220Ohm", ["Power(W)"]="0.5W", ["Tolerance(%)"]="1%"});
            Add("4E12000009","RES","RES 220Ohm 0.5W 1% 1206 AECQ","RCS1206220RFKEA","Vishay","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="220Ohm", ["Power(W)"]="0.5W", ["Tolerance(%)"]="1%"});
            Add("4E12000010","RES","RES 120Ohm 0.25W 1% 1206 AECQ","AC1206FR-3W12RL","Yageo","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="120Ohm", ["Power(W)"]="0.25W", ["Tolerance(%)"]="1%"});
            Add("4E12000010","RES","RES 120Ohm 0.25W 1% 1206 AECQ","AC1206FR-0712RL","Yageo","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="120Ohm", ["Power(W)"]="0.25W", ["Tolerance(%)"]="1%"});
            Add("4E12000010","RES","RES 12Ohm 0.75W 1% 1206 AECQ","ESR18EZPF12R0","Rohm","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="12Ohm", ["Power(W)"]="0.75W", ["Tolerance(%)"]="1%"});
            Add("4E12000010","RES","RES 12Ohm 0.5W 1% 1206 AECQ","RCS120612R0FKEA","Vishay","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="12Ohm", ["Power(W)"]="0.5W", ["Tolerance(%)"]="1%"});
            Add("4E12000011","RES","RES 15Ohm 0.5W 1% 1206 AECQ","RCS120615R0FKEA","Vishay","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="15Ohm", ["Power(W)"]="0.5W", ["Tolerance(%)"]="1%"});
            Add("4E12000012","RES","RES 18Ohm 0.5W 1% 1206 AECQ","RCS120618R0FKEA","Vishay","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="18Ohm", ["Power(W)"]="0.5W", ["Tolerance(%)"]="1%"});
            Add("4E12000013","RES","RES 10Ohm 0.1W 1% 0603 AECQ","AC0603FR-0710RL","Yageo","RES-0603","ResistorSymbol","RES-0603",AECQ,SMD,new(){["Resistance(Ohm)"]="10Ohm", ["Power(W)"]="0.1W", ["Tolerance(%)"]="1%"});
            Add("4E12000014","RES","RES 2KOhm 0.1W 1% 0603 AECQ","AC0603FR-072KL","Yageo","RES-0603","ResistorSymbol","RES-0603",AECQ,SMD,new(){["Resistance(Ohm)"]="2KOhm", ["Power(W)"]="0.1W", ["Tolerance(%)"]="1%"});
            Add("4E12000015","RES","RES 3.9KOhm 0.125W 1% 0805 AECQ","AC0805FR-073K9L","Yageo","RES-0805","ResistorSymbol","RES-0805",AECQ,SMD,new(){["Resistance(Ohm)"]="3.9KOhm", ["Power(W)"]="0.125W", ["Tolerance(%)"]="1%"});
            Add("4E12000016","RES","RES 33Ohm 0.1W 1% 0603 AECQ","AC0603FR-0733RL","Yageo","RES-0603","ResistorSymbol","RES-0603",AECQ,SMD,new(){["Resistance(Ohm)"]="33Ohm", ["Power(W)"]="0.1W", ["Tolerance(%)"]="1%"});
            Add("4E12000016","RES","RES 33Ohm 0.1W 1% 0603 AECQ","CQ03SAF330JT5E","Rohm","RES-0603","ResistorSymbol","RES-0603",AECQ,SMD,new(){["Resistance(Ohm)"]="33Ohm", ["Power(W)"]="0.1W", ["Tolerance(%)"]="1%"});
            Add("4E12000017","RES","RES 3KOhm 0.25W 1% 1206 AECQ","AC1206FR-073KL","Yageo","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="3KOhm", ["Power(W)"]="0.25W", ["Tolerance(%)"]="1%"});
            Add("4E12000018","RES","RES 4.7Ohm 0.125W 1% 0805 AECQ","AC0805FR-074R7L","Yageo","RES-0805","ResistorSymbol","RES-0805",AECQ,SMD,new(){["Resistance(Ohm)"]="4.7Ohm", ["Power(W)"]="0.125W", ["Tolerance(%)"]="1%"});
            Add("4E12000019","RES","RES 430Ohm 0.125W 1% 0805 AECQ","AC0805FR-07430RL","Yageo","RES-0805","ResistorSymbol","RES-0805",AECQ,SMD,new(){["Resistance(Ohm)"]="430Ohm", ["Power(W)"]="0.125W", ["Tolerance(%)"]="1%"});
            Add("4E12000020","RES","RES 8.2KOhm 0.125W 1% 0805 AECQ","AC0805FR-078K2L","Yageo","RES-0805","ResistorSymbol","RES-0805",AECQ,SMD,new(){["Resistance(Ohm)"]="8.2KOhm", ["Power(W)"]="0.125W", ["Tolerance(%)"]="1%"});
            Add("4E12000021","RES","RES 10KOhm 0.125W 1% 0805 AECQ","AC0805FR-0710KL","Yageo","RES-0805","ResistorSymbol","RES-0805",AECQ,SMD,new(){["Resistance(Ohm)"]="10KOhm", ["Power(W)"]="0.125W", ["Tolerance(%)"]="1%"});
            Add("4E12000021","RES","RES 10KOhm 0.125W 1% 0805 AECQ","CRCW080510K0FKEA","Vishay","RES-0805","ResistorSymbol","RES-0805",AECQ,SMD,new(){["Resistance(Ohm)"]="10KOhm", ["Power(W)"]="0.125W", ["Tolerance(%)"]="1%"});
            Add("4E12000021","RES","RES 10KOhm 0.125W 5% 0805 AECQ","CRCW080510K0JNEC","Vishay","RES-0805","ResistorSymbol","RES-0805",AECQ,SMD,new(){["Resistance(Ohm)"]="10KOhm", ["Power(W)"]="0.125W", ["Tolerance(%)"]="5%"});
            Add("4E12000021","RES","RES 10KOhm 0.125W 1% 0805 AECQ","CQ05S8F1002T5E","Rohm","RES-0805","ResistorSymbol","RES-0805",AECQ,SMD,new(){["Resistance(Ohm)"]="10KOhm", ["Power(W)"]="0.125W", ["Tolerance(%)"]="1%"});
            Add("4E12000021","RES","RES 10KOhm 0.125W 5% 0805 AECQ","AC0805JR-0710KL","Yageo","RES-0805","ResistorSymbol","RES-0805",AECQ,SMD,new(){["Resistance(Ohm)"]="10KOhm", ["Power(W)"]="0.125W", ["Tolerance(%)"]="5%"});
            Add("4E12000022","RES","RES 10KOhm 0.1W 1% 0603 AECQ","AC0603FR-0710KL","Yageo","RES-0603","ResistorSymbol","RES-0603",AECQ,SMD,new(){["Resistance(Ohm)"]="10KOhm", ["Power(W)"]="0.1W", ["Tolerance(%)"]="1%"});
            Add("4E12000022","RES","RES 10KOhm 0.1W 1% 0603 AECQ","AC0603FR-1310KL","Yageo","RES-0603","ResistorSymbol","RES-0603",AECQ,SMD,new(){["Resistance(Ohm)"]="10KOhm", ["Power(W)"]="0.1W", ["Tolerance(%)"]="1%"});
            Add("4E12000022","RES","RES 10KOhm 0.1W 1% 0603 AECQ","CRCW060310K0FKEA","Vishay","RES-0603","ResistorSymbol","RES-0603",AECQ,SMD,new(){["Resistance(Ohm)"]="10KOhm", ["Power(W)"]="0.1W", ["Tolerance(%)"]="1%"});
            Add("4E12000022","RES","RES 10KOhm 0.1W 1% 0603 AECQ","RMCF0603FT10K0","Stackpole","RES-0603","ResistorSymbol","RES-0603",AECQ,SMD,new(){["Resistance(Ohm)"]="10KOhm", ["Power(W)"]="0.1W", ["Tolerance(%)"]="1%"});
            Add("4E12000022","RES","RES 10KOhm 0.1W 1% 0603 AECQ","CQ03SAF1002T5E","Rohm","RES-0603","ResistorSymbol","RES-0603",AECQ,SMD,new(){["Resistance(Ohm)"]="10KOhm", ["Power(W)"]="0.1W", ["Tolerance(%)"]="1%"});
            Add("4E12000022","RES","RES 10KOhm","MEA10K3CG1421","","","ResistorSymbol","",STD,SMD,new(){["Resistance(Ohm)"]="10KOhm"});
            Add("4E12000023","RES","RES 22KOhm 0.1W 1% 0603 AECQ","AC0603FR-0722KL","Yageo","RES-0603","ResistorSymbol","RES-0603",AECQ,SMD,new(){["Resistance(Ohm)"]="22KOhm", ["Power(W)"]="0.1W", ["Tolerance(%)"]="1%"});
            Add("4E12000024","RES","RES 100Ohm 0.25W 1% 1206 AECQ","AC1206FR-07100RL","Yageo","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="100Ohm", ["Power(W)"]="0.25W", ["Tolerance(%)"]="1%"});
            Add("4E12000024","RES","RES 100Ohm 0.75W 1% 1206 AECQ","CRCW1206100RFKEAHP","Vishay","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="100Ohm", ["Power(W)"]="0.75W", ["Tolerance(%)"]="1%"});
            Add("4E12000025","RES","RES 20KOhm 0.1W 1% 0603 AECQ","AC0603FR-0720KL","Yageo","RES-0603","ResistorSymbol","RES-0603",AECQ,SMD,new(){["Resistance(Ohm)"]="20KOhm", ["Power(W)"]="0.1W", ["Tolerance(%)"]="1%"});
            Add("4E12000025","RES","RES 20KOhm 0.1W 1% 0603 AECQ","CQ03SAF2002T5E","Rohm","RES-0603","ResistorSymbol","RES-0603",AECQ,SMD,new(){["Resistance(Ohm)"]="20KOhm", ["Power(W)"]="0.1W", ["Tolerance(%)"]="1%"});
            Add("4E12000025","RES","RES 20KOhm 0.1W 1% 0603 AECQ","CRCW060320K0FKEA","Vishay","RES-0603","ResistorSymbol","RES-0603",AECQ,SMD,new(){["Resistance(Ohm)"]="20KOhm", ["Power(W)"]="0.1W", ["Tolerance(%)"]="1%"});
            Add("4E12000026","RES","RES 1KOhm 0.1W 1% 0603 AECQ","AC0603FR-071KL","Yageo","RES-0603","ResistorSymbol","RES-0603",AECQ,SMD,new(){["Resistance(Ohm)"]="1KOhm", ["Power(W)"]="0.1W", ["Tolerance(%)"]="1%"});
            Add("4E12000026","RES","RES 1KOhm 0.1W 1% 0603 AECQ","CQ03SAF1001T5E","Rohm","RES-0603","ResistorSymbol","RES-0603",AECQ,SMD,new(){["Resistance(Ohm)"]="1KOhm", ["Power(W)"]="0.1W", ["Tolerance(%)"]="1%"});
            Add("4E12000027","RES","RES 0Ohm 0.125W 5% 0805 AECQ","AC0805JR-070RL","Yageo","RES-0805","ResistorSymbol","RES-0805",AECQ,SMD,new(){["Resistance(Ohm)"]="0Ohm", ["Power(W)"]="0.125W", ["Tolerance(%)"]="5%"});
            Add("4E12000027","RES","RES 0Ohm 0.125W Jumper 0805 AECQ","CRCW08050000Z0EA","Vishay","RES-0805","ResistorSymbol","RES-0805",AECQ,SMD,new(){["Resistance(Ohm)"]="0Ohm", ["Power(W)"]="0.125W", ["Tolerance(%)"]="Jumper"});
            Add("4E12000027","RES","RES 0Ohm 0.125W 5% 0805 AECQ","CQ05S8J0000T5E","Rohm","RES-0805","ResistorSymbol","RES-0805",AECQ,SMD,new(){["Resistance(Ohm)"]="0Ohm", ["Power(W)"]="0.125W", ["Tolerance(%)"]="5%"});
            Add("4E12000027","RES","RES 0Ohm 0.5W Jumper 0805 AECQ","CRCW08050000Z0EAHP","Vishay","RES-0805","ResistorSymbol","RES-0805",AECQ,SMD,new(){["Resistance(Ohm)"]="0Ohm", ["Power(W)"]="0.5W", ["Tolerance(%)"]="Jumper"});
            Add("4E12000028","RES","RES 100KOhm 0.1W 1% 0603 AECQ","AC0603FR-07100KL","Yageo","RES-0603","ResistorSymbol","RES-0603",AECQ,SMD,new(){["Resistance(Ohm)"]="100KOhm", ["Power(W)"]="0.1W", ["Tolerance(%)"]="1%"});
            Add("4E12000028","RES","RES 100KOhm 0.1W 1% 0603 AECQ","CRCW0603100KFKEA","Vishay","RES-0603","ResistorSymbol","RES-0603",AECQ,SMD,new(){["Resistance(Ohm)"]="100KOhm", ["Power(W)"]="0.1W", ["Tolerance(%)"]="1%"});
            Add("4E12000028","RES","RES 100KOhm 0.1W 1% 0603 AECQ","CQ03SAF1003T5E","Rohm","RES-0603","ResistorSymbol","RES-0603",AECQ,SMD,new(){["Resistance(Ohm)"]="100KOhm", ["Power(W)"]="0.1W", ["Tolerance(%)"]="1%"});
            Add("4E12000029","RES","RES 130Ohm 0.25W 1% 1206 AECQ","AC1206FR-07130RL","Yageo","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="130Ohm", ["Power(W)"]="0.25W", ["Tolerance(%)"]="1%"});
            Add("4E12000030","RES","RES 160Ohm 0.25W 1% 1206 AECQ","AC1206FR-07160RL","Yageo","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="160Ohm", ["Power(W)"]="0.25W", ["Tolerance(%)"]="1%"});
            Add("4E12000031","RES","RES","EPS197181","","","ResistorSymbol","",STD,SMD,null);
            Add("4E12000032","RES","RES","EPS197201","","","ResistorSymbol","",STD,SMD,null);
            Add("4E12000033","RES","RES","EPS196221","","","ResistorSymbol","",STD,SMD,null);
            Add("4E12000033","RES","RES","EPS198223","","","ResistorSymbol","",STD,SMD,null);
            Add("4E12000034","RES","RES 240Ohm 0.25W 1% 1206 AECQ","AC1206FR-07240RL","Yageo","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="240Ohm", ["Power(W)"]="0.25W", ["Tolerance(%)"]="1%"});
            Add("4E12000034","RES","RES 240Ohm 0.25W 1% 1206 AECQ","CRCW1206240RFKEA","Vishay","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="240Ohm", ["Power(W)"]="0.25W", ["Tolerance(%)"]="1%"});
            Add("4E12000034","RES","RES","EPS197241","","","ResistorSymbol","",STD,SMD,null);
            Add("4E12000036","RES","RES 2KOhm 0.25W 5% 1206 AECQ","AC1206JR-072KL","Yageo","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="2KOhm", ["Power(W)"]="0.25W", ["Tolerance(%)"]="5%"});
            Add("4E12000036","RES","RES 2KOhm 0.25W 5% 1206 AECQ","RMCF1206JT2K00","Stackpole","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="2KOhm", ["Power(W)"]="0.25W", ["Tolerance(%)"]="5%"});
            Add("4E12000036","RES","RES 2KOhm 0.25W 5% 1206 AECQ","CQ06S4J0202T5E","Rohm","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="2KOhm", ["Power(W)"]="0.25W", ["Tolerance(%)"]="5%"});
            Add("4E12000038","RES","RES 10KOhm 0.25W 1% 1206 AECQ","CQ06S4F1002T5E","Rohm","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="10KOhm", ["Power(W)"]="0.25W", ["Tolerance(%)"]="1%"});
            Add("4E12000038","RES","RES 10KOhm 0.25W 5% 1206 AECQ","CQ06S4J0103T5E","Rohm","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="10KOhm", ["Power(W)"]="0.25W", ["Tolerance(%)"]="5%"});
            Add("4E12000038","RES","RES 10KOhm 0.25W 1% 1206 AECQ","RK73H2BTTD1002F","KOA","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="10KOhm", ["Power(W)"]="0.25W", ["Tolerance(%)"]="1%"});
            Add("4E12000038","RES","RES 10KOhm 0.25W 1% 1206 AECQ","AC1206FR-0710KL","Yageo","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="10KOhm", ["Power(W)"]="0.25W", ["Tolerance(%)"]="1%"});
            Add("4E12000039","RES","RES 383Ohm 0.25W 1% 1206 AECQ","AC1206FR-07383RL","Yageo","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="383Ohm", ["Power(W)"]="0.25W", ["Tolerance(%)"]="1%"});
            Add("4E12000039","RES","RES 383Ohm 0.25W 1% 1206 AECQ","CQ06W4F3830T5E","Rohm","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="383Ohm", ["Power(W)"]="0.25W", ["Tolerance(%)"]="1%"});
            Add("4E12000040","RES","RES 390Ohm 0.25W 1% 1206 AECQ","AC1206FR-07390RL","Yageo","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="390Ohm", ["Power(W)"]="0.25W", ["Tolerance(%)"]="1%"});
            Add("4E12000040","RES","RES 390Ohm 0.25W 1% 1206 AECQ","CQ06W4F3900T5E","Rohm","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="390Ohm", ["Power(W)"]="0.25W", ["Tolerance(%)"]="1%"});
            Add("4E12000041","RES","RES 430Ohm 0.25W 1% 1206 AECQ","AC1206FR-07430RL","Yageo","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="430Ohm", ["Power(W)"]="0.25W", ["Tolerance(%)"]="1%"});
            Add("4E12000041","RES","RES 430Ohm 0.25W 1% 1206 AECQ","CQ06W4F4300T5E","Rohm","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="430Ohm", ["Power(W)"]="0.25W", ["Tolerance(%)"]="1%"});
            Add("4E12000042","RES","RES 442Ohm 0.25W 1% 1206 AECQ","AC1206FR-07442RL","Yageo","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="442Ohm", ["Power(W)"]="0.25W", ["Tolerance(%)"]="1%"});
            Add("4E12000043","RES","RES 499Ohm 0.25W 1% 1206 AECQ","AC1206FR-07499RL","Yageo","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="499Ohm", ["Power(W)"]="0.25W", ["Tolerance(%)"]="1%"});
            Add("4E12000043","RES","RES 499Ohm 0.25W 1% 1206 AECQ","CQ06W4F4990T5E","Rohm","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="499Ohm", ["Power(W)"]="0.25W", ["Tolerance(%)"]="1%"});
            Add("4E12000044","RES","RES 499Ohm 0.25W 1% 1206 AECQ","CRCW1206499RFKEA","Vishay","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="499Ohm", ["Power(W)"]="0.25W", ["Tolerance(%)"]="1%"});
            Add("4E12000045","RES","RES 560Ohm 0.25W 1% 1206 AECQ","AC1206FR-07560RL","Yageo","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="560Ohm", ["Power(W)"]="0.25W", ["Tolerance(%)"]="1%"});
            Add("4E12000045","RES","RES 560Ohm 0.25W 1% 1206 AECQ","CQ06W4F5600T5E","Rohm","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="560Ohm", ["Power(W)"]="0.25W", ["Tolerance(%)"]="1%"});
            Add("4E12000046","RES","RES 580Ohm  1% 1206","AC1206FR-07580RL","Yageo","RES-1206","ResistorSymbol","RES-1206",STD,SMD,new(){["Resistance(Ohm)"]="580Ohm", ["Tolerance(%)"]="1%"});
            Add("4E12000046","RES","RES 590Ohm 0.25W 1% 1206 AECQ","AC1206FR-07590RL","Yageo","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="590Ohm", ["Power(W)"]="0.25W", ["Tolerance(%)"]="1%"});
            Add("4E12000046","RES","RES 576Ohm 0.25W 1% 1206 AECQ","AC1206FR-07576R","Yageo","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="576Ohm", ["Power(W)"]="0.25W", ["Tolerance(%)"]="1%"});
            Add("4E12000046","RES","RES 590Ohm 0.25W 1% 1206 AECQ","CQ06W4F5900T5E","Rohm","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="590Ohm", ["Power(W)"]="0.25W", ["Tolerance(%)"]="1%"});
            Add("4E12000047","RES","RES 649Ohm 0.25W 1% 1206 AECQ","AC1206FR-07649RL","Yageo","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="649Ohm", ["Power(W)"]="0.25W", ["Tolerance(%)"]="1%"});
            Add("4E12000047","RES","RES 649Ohm 0.25W 1% 1206 AECQ","CQ06W4F6490T5E","Rohm","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="649Ohm", ["Power(W)"]="0.25W", ["Tolerance(%)"]="1%"});
            Add("4E12000048","RES","RES 665Ohm 0.25W 1% 1206 AECQ","AC1206FR-07665RL","Yageo","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="665Ohm", ["Power(W)"]="0.25W", ["Tolerance(%)"]="1%"});
            Add("4E12000048","RES","RES 665Ohm 0.25W 1% 1206 AECQ","CQ06W4F6650T5E","Rohm","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="665Ohm", ["Power(W)"]="0.25W", ["Tolerance(%)"]="1%"});
            Add("4E12000049","RES","RES  0.25W 1% 1206 AECQ","AC1206FR-07845RL","Yageo","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Power(W)"]="0.25W", ["Tolerance(%)"]="1%"});
            Add("4E12000049","RES","RES   1% 1206 AECQ","CQ06W4F8450T5E","Rohm","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Tolerance(%)"]="1%"});
            Add("4E12000167","RES","RES 10KOhm 0.125W 1% 0805 AECQ","ERJ-6ENF1002V","Panasonic","RES-0805","ResistorSymbol","RES-0805",AECQ,SMD,new(){["Resistance(Ohm)"]="10KOhm", ["Power(W)"]="0.125W", ["Tolerance(%)"]="1%"});
            Add("4E12000192","RES","RES 100KOhm 0.1W 1% 0603 AECQ","RK73H1JTTD1003F","KOA","RES-0603","ResistorSymbol","RES-0603",AECQ,SMD,new(){["Resistance(Ohm)"]="100KOhm", ["Power(W)"]="0.1W", ["Tolerance(%)"]="1%"});
            Add("4E12000192","RES","RES 100KOhm 0.125W 1% 0603 AECQ","-CRCW0603100KFKEA-","Vishay","RES-0603","ResistorSymbol","RES-0603",AECQ,SMD,new(){["Resistance(Ohm)"]="100KOhm", ["Power(W)"]="0.125W", ["Tolerance(%)"]="1%"});
            Add("4E12000305","RES","RES 3.6KOhm 0.125W 5% 0805 AECQ","CR-05JA7---3K6","Viking","RES-0805","ResistorSymbol","RES-0805",AECQ,SMD,new(){["Resistance(Ohm)"]="3.6KOhm", ["Power(W)"]="0.125W", ["Tolerance(%)"]="5%"});
            Add("4E12000305","RES","RES 3.6KOhm 0.125W 5% 0805 AECQ","RMCF0805JT3K60","Stackpole","RES-0805","ResistorSymbol","RES-0805",AECQ,SMD,new(){["Resistance(Ohm)"]="3.6KOhm", ["Power(W)"]="0.125W", ["Tolerance(%)"]="5%"});
            Add("4E12000305","RES","RES 3.6KOhm 0.125W 5% 0805 AECQ","RMCF0805JG3K60","Stackpole","RES-0805","ResistorSymbol","RES-0805",AECQ,SMD,new(){["Resistance(Ohm)"]="3.6KOhm", ["Power(W)"]="0.125W", ["Tolerance(%)"]="5%"});
            Add("4E12000305","RES","RES 3.6KOhm 0.125W 5% 0805 AECQ","CRCW08053K60JNEB","Vishay","RES-0805","ResistorSymbol","RES-0805",AECQ,SMD,new(){["Resistance(Ohm)"]="3.6KOhm", ["Power(W)"]="0.125W", ["Tolerance(%)"]="5%"});
            Add("4E12000305","RES","RES 3.6KOhm 0.125W 1% 0805 AECQ","CRCW08053K60FKEA","Vishay","RES-0805","ResistorSymbol","RES-0805",AECQ,SMD,new(){["Resistance(Ohm)"]="3.6KOhm", ["Power(W)"]="0.125W", ["Tolerance(%)"]="1%"});
            Add("4E12000290","RES","RES 4.53KOhm 0.125W 1% 0805 AECQ","RMCF0805FT4K53","Stackpole","RES-0805","ResistorSymbol","RES-0805",AECQ,SMD,new(){["Resistance(Ohm)"]="4.53KOhm", ["Power(W)"]="0.125W", ["Tolerance(%)"]="1%"});
            Add("4E12000291","RES","RES 38.3KOhm 0.125W 1% 0805 AECQ","RMCF0805FT38K3","Stackpole","RES-0805","ResistorSymbol","RES-0805",AECQ,SMD,new(){["Resistance(Ohm)"]="38.3KOhm", ["Power(W)"]="0.125W", ["Tolerance(%)"]="1%"});
            Add("4E12000238","RES","RES 22KOhm 0.125W 1% 0805 AECQ","AC0805FR-0722KL","Yageo","RES-0805","ResistorSymbol","RES-0805",AECQ,SMD,new(){["Resistance(Ohm)"]="22KOhm", ["Power(W)"]="0.125W", ["Tolerance(%)"]="1%"});
            Add("4E12000239","RES","RES 47KOhm 0.125W 1% 0805 AECQ","AC0805FR-0747KL","Yageo","RES-0805","ResistorSymbol","RES-0805",AECQ,SMD,new(){["Resistance(Ohm)"]="47KOhm", ["Power(W)"]="0.125W", ["Tolerance(%)"]="1%"});
            Add("4E12000414","RES","RES 330KOhm 0.125W 1% 0805 AECQ","CQ05S8F3303T5E","Rohm","RES-0805","ResistorSymbol","RES-0805",AECQ,SMD,new(){["Resistance(Ohm)"]="330KOhm", ["Power(W)"]="0.125W", ["Tolerance(%)"]="1%"});
            Add("4E12000414","RES","RES 330KOhm 0.125W 1% 0805 AECQ","AC0805FR-07330KL","Yageo","RES-0805","ResistorSymbol","RES-0805",AECQ,SMD,new(){["Resistance(Ohm)"]="330KOhm", ["Power(W)"]="0.125W", ["Tolerance(%)"]="1%"});
            Add("4E12000163","RES","RES 100KOhm 0.125W 1% 0805 AECQ","AC0805FR-07100KL","Yageo","RES-0805","ResistorSymbol","RES-0805",AECQ,SMD,new(){["Resistance(Ohm)"]="100KOhm", ["Power(W)"]="0.125W", ["Tolerance(%)"]="1%"});
            Add("4E12000163","RES","RES 100KOhm 0.125W 1% 0805 AECQ","ERJ-6ENF1003V","Panasonic","RES-0805","ResistorSymbol","RES-0805",AECQ,SMD,new(){["Resistance(Ohm)"]="100KOhm", ["Power(W)"]="0.125W", ["Tolerance(%)"]="1%"});
            Add("4E12000163","RES","RES 100KOhm 0.125W 1% 0805 AECQ","AC0805FR-13100KL","Yageo","RES-0805","ResistorSymbol","RES-0805",AECQ,SMD,new(){["Resistance(Ohm)"]="100KOhm", ["Power(W)"]="0.125W", ["Tolerance(%)"]="1%"});
            Add("4E12000163","RES","RES 100KOhm 0.125W 1% 0805 AECQ","CQ05S8F1003T5E","Rohm","RES-0805","ResistorSymbol","RES-0805",AECQ,SMD,new(){["Resistance(Ohm)"]="100KOhm", ["Power(W)"]="0.125W", ["Tolerance(%)"]="1%"});
            Add("4E12000304","RES","RES 2KOhm 0.125W 5% 0805 AECQ","ERJ-6GEYJ202V","Panasonic","RES-0805","ResistorSymbol","RES-0805",AECQ,SMD,new(){["Resistance(Ohm)"]="2KOhm", ["Power(W)"]="0.125W", ["Tolerance(%)"]="5%"});
            Add("4E12000304","RES","RES 2KOhm 0.125W 5% 0805 AECQ","RMCF0805JT2K00","Stackpole","RES-0805","ResistorSymbol","RES-0805",AECQ,SMD,new(){["Resistance(Ohm)"]="2KOhm", ["Power(W)"]="0.125W", ["Tolerance(%)"]="5%"});
            Add("4E12000293","RES","RES 5.6KOhm 0.125W 1% 0805 AECQ","RMCF0805FT5K60","Stackpole","RES-0805","ResistorSymbol","RES-0805",AECQ,SMD,new(){["Resistance(Ohm)"]="5.6KOhm", ["Power(W)"]="0.125W", ["Tolerance(%)"]="1%"});
            Add("4E12000394","RES","RES 6.34KOhm 0.125W 1% 0805 AECQ","RMCF0805FT6K34","Stackpole","RES-0805","ResistorSymbol","RES-0805",AECQ,SMD,new(){["Resistance(Ohm)"]="6.34KOhm", ["Power(W)"]="0.125W", ["Tolerance(%)"]="1%"});
            Add("4E12000386","RES","RES 7.32KOhm 0.125W 1% 0805 AECQ","RMCF0805FT7K32","Stackpole","RES-0805","ResistorSymbol","RES-0805",AECQ,SMD,new(){["Resistance(Ohm)"]="7.32KOhm", ["Power(W)"]="0.125W", ["Tolerance(%)"]="1%"});
            Add("4E12000169","RES","RES 10KOhm 0.1W 1% 0603 AECQ","-AC0603FR-0710KL-","Yageo","RES-0603","ResistorSymbol","RES-0603",AECQ,SMD,new(){["Resistance(Ohm)"]="10KOhm", ["Power(W)"]="0.1W", ["Tolerance(%)"]="1%"});
            Add("4E12000169","RES","RES 10KOhm 0.1W 1% 0603 AECQ","CRCW060310K0FKEI","Vishay","RES-0603","ResistorSymbol","RES-0603",AECQ,SMD,new(){["Resistance(Ohm)"]="10KOhm", ["Power(W)"]="0.1W", ["Tolerance(%)"]="1%"});
            Add("4E12000169","RES","RES 1.5KOhm 0.1W 1% 0603 AECQ","CQ03WAF1501T5E","Rohm","RES-0603","ResistorSymbol","RES-0603",AECQ,SMD,new(){["Resistance(Ohm)"]="1.5KOhm", ["Power(W)"]="0.1W", ["Tolerance(%)"]="1%"});
            Add("4E12000301","RES","RES 24Ohm 0.5W 0,01 1206 AECQ","SR1206FR-7W24RL","Yageo","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="24Ohm", ["Power(W)"]="0.5W", ["Tolerance(%)"]="0.01"});
            Add("4E12000301","RES","RES 24Ohm 0.5W 0,01 1206 AECQ","RC1206FR-7W24RL","Yageo","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="24Ohm", ["Power(W)"]="0.5W", ["Tolerance(%)"]="0.01"});
            Add("4E12000301","RES","RES 24Ohm 0.33W 0,01 1206 AECQ","SG73P2BTTD24R0F","KOA","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="24Ohm", ["Power(W)"]="0.33W", ["Tolerance(%)"]="0.01"});
            Add("4E12000301","RES","RES 24Ohm 0.5W 0,01 1206 AECQ","RCS120624R0FKEA","Vishay","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="24Ohm", ["Power(W)"]="0.5W", ["Tolerance(%)"]="0.01"});
            Add("4E12000301","RES","RES 24Ohm 0.75W 0,01 1206 AECQ","CRCW120624R0FKEAHP","Vishay","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="24Ohm", ["Power(W)"]="0.75W", ["Tolerance(%)"]="0.01"});
            Add("4E12000357","RES","RES 9.53KOhm 0.125W 1% 0805 AECQ","RMCF0805FT9K53","Stackpole","RES-0805","ResistorSymbol","RES-0805",AECQ,SMD,new(){["Resistance(Ohm)"]="9.53KOhm", ["Power(W)"]="0.125W", ["Tolerance(%)"]="1%"});
            Add("4E12000358","RES","RES 14.3KOhm 0.125W 1% 0805 AECQ","RMCF0805FT14K3","Stackpole","RES-0805","ResistorSymbol","RES-0805",AECQ,SMD,new(){["Resistance(Ohm)"]="14.3KOhm", ["Power(W)"]="0.125W", ["Tolerance(%)"]="1%"});
            Add("4E12000364","RES","RES 300Ohm 0.75W 1% 1206 AECQ","CRCW1206300RFKEAHP","Vishay","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="300Ohm", ["Power(W)"]="0.75W", ["Tolerance(%)"]="1%"});
            Add("4E12000093","RES","RES 0Ohm Jumper 5% 0603 AECQ","AC0603JR-070RL","Yageo","RES-0603","ResistorSymbol","RES-0603",AECQ,SMD,new(){["Resistance(Ohm)"]="0Ohm", ["Power(W)"]="Jumper", ["Tolerance(%)"]="5%"});
            Add("4E12000093","RES","RES 0Ohm Jumper 1% 0603 AECQ","AC0603FR-070RL","Yageo","RES-0603","ResistorSymbol","RES-0603",AECQ,SMD,new(){["Resistance(Ohm)"]="0Ohm", ["Power(W)"]="Jumper", ["Tolerance(%)"]="1%"});
            Add("4E12000369","RES","RES 12KOhm 0.5W 1% 0805 AECQ","CRCW080512K0FKEAHP","Vishay","RES-0805","ResistorSymbol","RES-0805",AECQ,SMD,new(){["Resistance(Ohm)"]="12KOhm", ["Power(W)"]="0.5W", ["Tolerance(%)"]="1%"});
            Add("4E12000193","RES","RES 47KOhm 0.1W 1% 0603 AECQ","ERJ3EKF4702V","Panasonic","RES-0603","ResistorSymbol","RES-0603",AECQ,SMD,new(){["Resistance(Ohm)"]="47KOhm", ["Power(W)"]="0.1W", ["Tolerance(%)"]="1%"});
            Add("4E12000193","RES","RES 47KOhm 0.125W 1% 0603 AECQ","CRCW060347K0FKEA","Vishay","RES-0603","ResistorSymbol","RES-0603",AECQ,SMD,new(){["Resistance(Ohm)"]="47KOhm", ["Power(W)"]="0.125W", ["Tolerance(%)"]="1%"});
            Add("4E12000193","RES","RES 47KOhm 0.1W 1% 0603 AECQ","AC0603FR-1347KL","Yageo","RES-0603","ResistorSymbol","RES-0603",AECQ,SMD,new(){["Resistance(Ohm)"]="47KOhm", ["Power(W)"]="0.1W", ["Tolerance(%)"]="1%"});
            Add("4E12000193","RES","RES 47KOhm 0.1W 1% 0603 AECQ","AC0603FR-0747KL","Yageo","RES-0603","ResistorSymbol","RES-0603",AECQ,SMD,new(){["Resistance(Ohm)"]="47KOhm", ["Power(W)"]="0.1W", ["Tolerance(%)"]="1%"});
            Add("4E12000193","RES","RES 47KOhm 0.1W 1% 0603 AECQ","RMCF0603FT47K0","Stackpole","RES-0603","ResistorSymbol","RES-0603",AECQ,SMD,new(){["Resistance(Ohm)"]="47KOhm", ["Power(W)"]="0.1W", ["Tolerance(%)"]="1%"});
            Add("4E12000181","RES","RES 33KOhm 0.1W 1% 0603 AECQ","RMCF0603FT33K0","Stackpole","RES-0603","ResistorSymbol","RES-0603",AECQ,SMD,new(){["Resistance(Ohm)"]="33KOhm", ["Power(W)"]="0.1W", ["Tolerance(%)"]="1%"});
            Add("4E12000434","RES","RES 9.09KOhm 0.1W 1% 0603 AECQ","RMCF0603FT9K09","Stackpole","RES-0603","ResistorSymbol","RES-0603",AECQ,SMD,new(){["Resistance(Ohm)"]="9.09KOhm", ["Power(W)"]="0.1W", ["Tolerance(%)"]="1%"});
            Add("4E12000434","RES","RES 9.09KOhm 0.1W 1% 0603 AECQ","ERJ-3EKF9091V","Panasonic","RES-0603","ResistorSymbol","RES-0603",AECQ,SMD,new(){["Resistance(Ohm)"]="9.09KOhm", ["Power(W)"]="0.1W", ["Tolerance(%)"]="1%"});
            Add("4E12000434","RES","RES 9.09KOhm 0.1W 1% 0603 AECQ","RK73H1JTTD9091F","KOA","RES-0603","ResistorSymbol","RES-0603",AECQ,SMD,new(){["Resistance(Ohm)"]="9.09KOhm", ["Power(W)"]="0.1W", ["Tolerance(%)"]="1%"});
            Add("4E12000177","RES","RES 82KOhm 0.1W 1% 0603 AECQ","RK73G1JTTD8202F","KOA","RES-0603","ResistorSymbol","RES-0603",AECQ,SMD,new(){["Resistance(Ohm)"]="82KOhm", ["Power(W)"]="0.1W", ["Tolerance(%)"]="1%"});
            Add("4E12000366","RES","RES 3.3KOhm 0.75W 1% 1206 AECQ","CRCW12063K30FKEAHP","Vishay","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="3.3KOhm", ["Power(W)"]="0.75W", ["Tolerance(%)"]="1%"});
            Add("4E12000367","RES","RES 1KOhm 0.5W 1% 0805 AECQ","CRCW08051K00FKEAHP","Vishay","RES-0805","ResistorSymbol","RES-0805",AECQ,SMD,new(){["Resistance(Ohm)"]="1KOhm", ["Power(W)"]="0.5W", ["Tolerance(%)"]="1%"});
            Add("4E12000368","RES","RES 3.3KOhm 0.5W 1% 0805 AECQ","CRCW08053K30FKEAHP","Vishay","RES-0805","ResistorSymbol","RES-0805",AECQ,SMD,new(){["Resistance(Ohm)"]="3.3KOhm", ["Power(W)"]="0.5W", ["Tolerance(%)"]="1%"});
            Add("4E12000370","RES","RES 33KOhm 0.5W 1% 0805 AECQ","RCS080533K0FKEA","Vishay","RES-0805","ResistorSymbol","RES-0805",AECQ,SMD,new(){["Resistance(Ohm)"]="33KOhm", ["Power(W)"]="0.5W", ["Tolerance(%)"]="1%"});
            Add("4E12000053","RES","RES","4E12000027","","","ResistorSymbol","",STD,SMD,null);
            Add("4E12000374","RES","RES 4.7KOhm 0.5W 1% 0805 AECQ","CRCW08054K70FKEAHP","Vishay","RES-0805","ResistorSymbol","RES-0805",AECQ,SMD,new(){["Resistance(Ohm)"]="4.7KOhm", ["Power(W)"]="0.5W", ["Tolerance(%)"]="1%"});
            Add("4E12000371","RES","RES 68KOhm 0.5W 1% 0805 AECQ","CRCW080568K0FKEAHP","Vishay","RES-0805","ResistorSymbol","RES-0805",AECQ,SMD,new(){["Resistance(Ohm)"]="68KOhm", ["Power(W)"]="0.5W", ["Tolerance(%)"]="1%"});
            Add("4E12000062","RES","RES 33KOhm 0.1W 1% 0603 AECQ","AC0603FR-0733KL","Yageo","RES-0603","ResistorSymbol","RES-0603",AECQ,SMD,new(){["Resistance(Ohm)"]="33KOhm", ["Power(W)"]="0.1W", ["Tolerance(%)"]="1%"});
            Add("4E12000261","RES","RES 330Ohm 0.1W 1% 0603 AECQ","AC0603FR-07330RL","Yageo","RES-0603","ResistorSymbol","RES-0603",AECQ,SMD,new(){["Resistance(Ohm)"]="330Ohm", ["Power(W)"]="0.1W", ["Tolerance(%)"]="1%"});
            Add("4E12000372","RES","RES 470Ohm 0.1W 1% 0603 AECQ","AC0603FR-07470RL","Yageo","RES-0603","ResistorSymbol","RES-0603",AECQ,SMD,new(){["Resistance(Ohm)"]="470Ohm", ["Power(W)"]="0.1W", ["Tolerance(%)"]="1%"});
            Add("4E12000373","RES","RES 0Ohm 0.75W Jumper 1206 AECQ","CRCW12060000Z0EAHP","Vishay","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="0Ohm", ["Power(W)"]="0.75W", ["Tolerance(%)"]="Jumper"});
            Add("4E12000429","RES","RES 27.7KOhm 0.1W 0.5% 0603 AECQ","RN73R1JTTD2772D25","KOA","RES-0603","ResistorSymbol","RES-0603",AECQ,SMD,new(){["Resistance(Ohm)"]="27.7KOhm", ["Power(W)"]="0.1W", ["Tolerance(%)"]="0.5%"});
            Add("4E12000418","RES","RES 34.4KOhm 0.1W 0.5% 0603 AECQ","RN73R1JTTD3442D25","KOA","RES-0603","ResistorSymbol","RES-0603",AECQ,SMD,new(){["Resistance(Ohm)"]="34.4KOhm", ["Power(W)"]="0.1W", ["Tolerance(%)"]="0.5%"});
            Add("4E12000419","RES","RES 37.9KOhm 0.1W 0.5% 0603 AECQ","RN73R1JTTD3792D25","KOA","RES-0603","ResistorSymbol","RES-0603",AECQ,SMD,new(){["Resistance(Ohm)"]="37.9KOhm", ["Power(W)"]="0.1W", ["Tolerance(%)"]="0.5%"});
            Add("4E12000441","RES","RES 7.68KOhm 0.1W 1% 0603 AECQ","AC0603FR-077K68L","Yageo","RES-0603","ResistorSymbol","RES-0603",AECQ,SMD,new(){["Resistance(Ohm)"]="7.68KOhm", ["Power(W)"]="0.1W", ["Tolerance(%)"]="1%"});
            Add("4E12000442","RES","RES 6.49KOhm 0.1W 1% 0603 AECQ","RMCF0603FT6K49","Stackpole","RES-0603","ResistorSymbol","RES-0603",AECQ,SMD,new(){["Resistance(Ohm)"]="6.49KOhm", ["Power(W)"]="0.1W", ["Tolerance(%)"]="1%"});
            Add("4E12000402","RES","RES 2.2Ohm 0.25W 5% 0805 AECQ","SG73P2ATTD2R2J","KOA","RES-0805","ResistorSymbol","RES-0805",AECQ,SMD,new(){["Resistance(Ohm)"]="2.2Ohm", ["Power(W)"]="0.25W", ["Tolerance(%)"]="5%"});
            Add("4E12000050","RES","RES 850Ohm 0.25W 1% 1206 AECQ","CQ06W4F8500T5E","Royalohm","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="850Ohm", ["Power(W)"]="0.25W", ["Tolerance(%)"]="1%"});
            Add("4E12000051","RES","RES 38.3Ohm 0.5W 1% 1210 AECQ","AC1210FR-0738R3L","Yageo","RES-1210","ResistorSymbol","RES-1210",AECQ,SMD,new(){["Resistance(Ohm)"]="38.3Ohm", ["Power(W)"]="0.5W", ["Tolerance(%)"]="1%"});
            Add("4E12000051","RES","RES 38.3Ohm 0.5W 1% 1210 AECQ","CQ07U2F383JT5E","Royalohm","RES-1210","ResistorSymbol","RES-1210",AECQ,SMD,new(){["Resistance(Ohm)"]="38.3Ohm", ["Power(W)"]="0.5W", ["Tolerance(%)"]="1%"});
            Add("4E12000052","RES","RES 24.9Ohm 0.5W 1% 1210 AECQ","AC1210FR-0724R9L","Yageo","RES-1210","ResistorSymbol","RES-1210",AECQ,SMD,new(){["Resistance(Ohm)"]="24.9Ohm", ["Power(W)"]="0.5W", ["Tolerance(%)"]="1%"});
            Add("4E12000052","RES","RES 24.9Ohm 0.5W 1% 1210 AECQ","CQ07U2F249JT5E","Royalohm","RES-1210","ResistorSymbol","RES-1210",AECQ,SMD,new(){["Resistance(Ohm)"]="24.9Ohm", ["Power(W)"]="0.5W", ["Tolerance(%)"]="1%"});
            Add("4E12000053","RES","RES 0Ohm 0.125W Jumper 0805 AECQ","CRCW08050000Z0EA","Vishay","RES-0805","ResistorSymbol","RES-0805",AECQ,SMD,new(){["Resistance(Ohm)"]="0Ohm", ["Power(W)"]="0.125W", ["Tolerance(%)"]="Jumper"});
            Add("4E12000054","RES","RES 0Ohm 0.25W 1% 1206 AECQ","AC1206FR-070RL","Yageo","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="0Ohm", ["Power(W)"]="0.25W", ["Tolerance(%)"]="1%"});
            Add("4E12000054","RES","RES 0Ohm 0.25W 5% 1206 AECQ","AC1206JR-070RL","Yageo","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="0Ohm", ["Power(W)"]="0.25W", ["Tolerance(%)"]="5%"});
            Add("4E12000054","RES","RES 0Ohm 0.5W Jumper 1206 AECQ","RCS12060000Z0EA","Vishay","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="0Ohm", ["Power(W)"]="0.5W", ["Tolerance(%)"]="Jumper"});
            Add("4E12000054","RES","RES 0Ohm 0.5W 1% 1206 AECQ","HP06W2F0000T5E","Royalohm","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="0Ohm", ["Power(W)"]="0.5W", ["Tolerance(%)"]="1%"});
            Add("4E12000055","RES","RES 30Ohm 0.25W 5% 1206 AECQ","AC1206JR-0730RL","Yageo","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="30Ohm", ["Power(W)"]="0.25W", ["Tolerance(%)"]="5%"});
            Add("4E12000055","RES","RES 30Ohm 0.25W 1% 1206 AECQ","AC1206FR-7W30RL","Yageo","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="30Ohm", ["Power(W)"]="0.25W", ["Tolerance(%)"]="1%"});
            Add("4E12000055","RES","RES 30Ohm 0.5W 1% 1206 AECQ","HP06W2F300JT5E","Royalohm","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="30Ohm", ["Power(W)"]="0.5W", ["Tolerance(%)"]="1%"});
            Add("4E12000055","RES","RES 30Ohm 0.5W 1% 1206 AECQ","SR1206FR-7W30RL","Yageo","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="30Ohm", ["Power(W)"]="0.5W", ["Tolerance(%)"]="1%"});
            Add("4E12000055","RES","RES 30Ohm 0.25W 1% 1206 AECQ","MR12X30R0FTL","Walsin","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="30Ohm", ["Power(W)"]="0.25W", ["Tolerance(%)"]="1%"});
            Add("4E12000056","RES","RES 10Ohm 0.1W 1% 0603 AECQ","CQ03SAF100JT5E","Royalohm","RES-0603","ResistorSymbol","RES-0603",AECQ,SMD,new(){["Resistance(Ohm)"]="10Ohm", ["Power(W)"]="0.1W", ["Tolerance(%)"]="1%"});
            Add("4E12000057","RES","RES 24.9Ohm 0.1W 1% 0603 AECQ","AC0603FR-0724R9L","Yageo","RES-0603","ResistorSymbol","RES-0603",AECQ,SMD,new(){["Resistance(Ohm)"]="24.9Ohm", ["Power(W)"]="0.1W", ["Tolerance(%)"]="1%"});
            Add("4E12000057","RES","RES 24.9Ohm 0.1W 1% 0603 AECQ","CQ03SAF249JT5E","Yageo","RES-0603","ResistorSymbol","RES-0603",AECQ,SMD,new(){["Resistance(Ohm)"]="24.9Ohm", ["Power(W)"]="0.1W", ["Tolerance(%)"]="1%"});
            Add("4E12000058","RES","RES 60.4Ohm 0.1W 1% 0603 AECQ","AC0603FR-0760R4L","Yageo","RES-0603","ResistorSymbol","RES-0603",AECQ,SMD,new(){["Resistance(Ohm)"]="60.4Ohm", ["Power(W)"]="0.1W", ["Tolerance(%)"]="1%"});
            Add("4E12000058","RES","RES 59Ohm 0.1W 1% 0603 AECQ","CQ03SAF590JT5E","Royalohm","RES-0603","ResistorSymbol","RES-0603",AECQ,SMD,new(){["Resistance(Ohm)"]="59Ohm", ["Power(W)"]="0.1W", ["Tolerance(%)"]="1%"});
            Add("4E12000058","RES","RES 60.4Ohm 0.125W 1% 0603 AECQ","CRCW060360R4FKEA","Vishay","RES-0603","ResistorSymbol","RES-0603",AECQ,SMD,new(){["Resistance(Ohm)"]="60.4Ohm", ["Power(W)"]="0.125W", ["Tolerance(%)"]="1%"});
            Add("4E12000059","RES","RES 69.8Ohm 0.1W 1% 0603 AECQ","CQ03SAF698JT5E","Royalohm","RES-0603","ResistorSymbol","RES-0603",AECQ,SMD,new(){["Resistance(Ohm)"]="69.8Ohm", ["Power(W)"]="0.1W", ["Tolerance(%)"]="1%"});
            Add("4E12000060","RES","RES 120Ohm 0.1W 1% 0603 AECQ","AC0603FR-07120RL","Yageo","RES-0603","ResistorSymbol","RES-0603",AECQ,SMD,new(){["Resistance(Ohm)"]="120Ohm", ["Power(W)"]="0.1W", ["Tolerance(%)"]="1%"});
            Add("4E12000060","RES","RES 120Ohm 0.1W 1% 0603 AECQ","CQ03SAF1200T5E","Royalohm","RES-0603","ResistorSymbol","RES-0603",AECQ,SMD,new(){["Resistance(Ohm)"]="120Ohm", ["Power(W)"]="0.1W", ["Tolerance(%)"]="1%"});
            Add("4E12000060","RES","RES 2.2KOhm 0.0625W 1% 0402 STD","0402WGF2201TCE","Royalohm","RES-0402","ResistorSymbol","RES-0402",STD,SMD,new(){["Resistance(Ohm)"]="2.2KOhm", ["Power(W)"]="0.0625W", ["Tolerance(%)"]="1%"});
            Add("4E12000061","RES","RES 390Ohm 0.1W 1% 0603 AECQ","AC0603FR-07390RL","Yageo","RES-0603","ResistorSymbol","RES-0603",AECQ,SMD,new(){["Resistance(Ohm)"]="390Ohm", ["Power(W)"]="0.1W", ["Tolerance(%)"]="1%"});
            Add("4E12000061","RES","RES 390Ohm 0.1W 1% 0603 AECQ","CQ03SAF3900T5E","Royalohm","RES-0603","ResistorSymbol","RES-0603",AECQ,SMD,new(){["Resistance(Ohm)"]="390Ohm", ["Power(W)"]="0.1W", ["Tolerance(%)"]="1%"});
            Add("4E12000061","RES","RES 390Ohm 0.1W 1% 0603 AECQ","ERJ3EKF3900V","Panasonic","RES-0603","ResistorSymbol","RES-0603",AECQ,SMD,new(){["Resistance(Ohm)"]="390Ohm", ["Power(W)"]="0.1W", ["Tolerance(%)"]="1%"});
            Add("4E12000062","RES","RES 33KOhm 0.1W 1% 0603 AECQ","AC0603FR-0733KL","Yageo","RES-0603","ResistorSymbol","RES-0603",AECQ,SMD,new(){["Resistance(Ohm)"]="33KOhm", ["Power(W)"]="0.1W", ["Tolerance(%)"]="1%"});
            Add("4E12000062","RES","RES","CRCW060333K0FKEA","","","ResistorSymbol","",STD,SMD,null);
            Add("4E12000062","RES","RES","WRC53233","","","ResistorSymbol","",STD,SMD,null);
            Add("4E12000062","RES","RES 33KOhm 0.1W 1% 0603 AECQ","AC0603FR-1033KL","Yageo","RES-0603","ResistorSymbol","RES-0603",AECQ,SMD,new(){["Resistance(Ohm)"]="33KOhm", ["Power(W)"]="0.1W", ["Tolerance(%)"]="1%"});
            Add("4E12000063","RES","RES 220Ohm 0.25W 5% 1206 AECQ","AC1206JR-07220RL","Yageo","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="220Ohm", ["Power(W)"]="0.25W", ["Tolerance(%)"]="5%"});
            Add("4E12000063","RES","RES 220Ohm 0.25W 1% 1206 AECQ","AC1206FR-07220RL","Yageo","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="220Ohm", ["Power(W)"]="0.25W", ["Tolerance(%)"]="1%"});
            Add("4E12000063","RES","RES 220Ohm 0.5W 1% 1206 AECQ","RCS1206220RFKEA","Vishay","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="220Ohm", ["Power(W)"]="0.5W", ["Tolerance(%)"]="1%"});
            Add("4E12000063","RES","RES 220Ohm 0.5W 1% 1206 AECQ","WHPC1206-220RFT5","Welwyn","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="220Ohm", ["Power(W)"]="0.5W", ["Tolerance(%)"]="1%"});
            Add("4E12000063","RES","RES 220Ohm 0.5W 1% 1206 AECQ","HP06W2F2200T5E","Royalohm","RES-1206","ResistorSymbol","RES-1206",AECQ,SMD,new(){["Resistance(Ohm)"]="220Ohm", ["Power(W)"]="0.5W", ["Tolerance(%)"]="1%"});
            Add("4E12000102","RES","RES 2.2KOhm 0.2W 1% 0603 AECQ","AC0603FR-7W2K2L","Yageo","RES-0603","ResistorSymbol","RES-0603",AECQ,SMD,new(){["Resistance(Ohm)"]="2.2KOhm", ["Power(W)"]="0.2W", ["Tolerance(%)"]="1%"});
            Add("4E12000102","RES","RES 2.2KOhm 0.1W 1% 0603 AECQ","RMCF0603FT2K20","Stackpole","RES-0603","ResistorSymbol","RES-0603",AECQ,SMD,new(){["Resistance(Ohm)"]="2.2KOhm", ["Power(W)"]="0.1W", ["Tolerance(%)"]="1%"});
            Add("4E12000336","RES","RES 3.6KOhm 0.1W 1% 0603 AECQ","AC0603FR-073K6L","Yageo","RES-0603","ResistorSymbol","RES-0603",AECQ,SMD,new(){["Resistance(Ohm)"]="3.6KOhm", ["Power(W)"]="0.1W", ["Tolerance(%)"]="1%"});
            Add("4E12000278","RES","RES 38.3KOhm 0.1W 1% 0603 AECQ","AC0603FR-0738K3L","Yageo","RES-0603","ResistorSymbol","RES-0603",AECQ,SMD,new(){["Resistance(Ohm)"]="38.3KOhm", ["Power(W)"]="0.1W", ["Tolerance(%)"]="1%"});
            Add("4E12000278","RES","RES 38.3KOhm 0.1W 1% 0603 STD","CR0603FX-3832ELF","Yageo","RES-0603","ResistorSymbol","RES-0603",STD,SMD,new(){["Resistance(Ohm)"]="38.3KOhm", ["Power(W)"]="0.1W", ["Tolerance(%)"]="1%"});
            Add("4E12000147","RES","RES 11KOhm 0.1W 1% 0603 AECQ","AC0603FR-0711KL","Yageo","RES-0603","ResistorSymbol","RES-0603",AECQ,SMD,new(){["Resistance(Ohm)"]="11KOhm", ["Power(W)"]="0.1W", ["Tolerance(%)"]="1%"});
            Add("4E12000147","RES","RES 11KOhm 0.1W 1% 0603 AECQ","AC0603FR-1311KL","Yageo","RES-0603","ResistorSymbol","RES-0603",AECQ,SMD,new(){["Resistance(Ohm)"]="11KOhm", ["Power(W)"]="0.1W", ["Tolerance(%)"]="1%"});
            Add("4E12000188","RES","RES 30KOhm 0.1W 1% 0603 AECQ","RMCF0603FT30K0","Stackpole","RES-0603","ResistorSymbol","RES-0603",AECQ,SMD,new(){["Resistance(Ohm)"]="30KOhm", ["Power(W)"]="0.1W", ["Tolerance(%)"]="1%"});
            Add("4E12000359","RES","RES 6.2KOhm 0.1W 1% 0603 AECQ","AC0603FR-076K2L","Yageo","RES-0603","ResistorSymbol","RES-0603",AECQ,SMD,new(){["Resistance(Ohm)"]="6.2KOhm", ["Power(W)"]="0.1W", ["Tolerance(%)"]="1%"});
            Add("4E12000228","RES","RES 150Ohm 0.5W 1% 1210 AECQ","AC1210FR-07150RL","Yageo","RES-1210","ResistorSymbol","RES-1210",AECQ,SMD,new(){["Resistance(Ohm)"]="150Ohm", ["Power(W)"]="0.5W", ["Tolerance(%)"]="1%"});
            Add("4E12000299","RES","RES 200Ohm 0.5W 1% 1210 AECQ","AC1210FR-07200RL","Yageo","RES-1210","ResistorSymbol","RES-1210",AECQ,SMD,new(){["Resistance(Ohm)"]="200Ohm", ["Power(W)"]="0.5W", ["Tolerance(%)"]="1%"});
            //  CAPACITORS 
            Add("4E22000001","CAP","CAP 1nF 50V 10% X7R NFM 0603 AECQ","AC0603KPX7R9BB102","Yageo","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="1nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000001","CAP","CAP 1nF 50V 10% X7R OM 0603 AECQ","AC0603KRX7R9BB102","Yageo","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="1nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="OM"});
            Add("4E22000001","CAP","CAP 1nF 50V 10% X7R OM 0603 AECQ","CL10B102KB8WPNC","Samsung","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="1nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="OM"});
            Add("4E22000001","CAP","CAP 1nF 50V 10% X7R ST 0603 AECQ","GCE188R71H102KA01D","Murata","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="1nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="ST"});
            Add("4E22000002","CAP","CAP 100nF 50V 10% X7R OM 0805 AECQ","C0805F104K5RACAUTO","Kemet","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="100nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="OM"});
            Add("4E22000002","CAP","CAP 100nF 50V 10% X7R ST 0805 AECQ","CGA4J2X7R1H104K125AE","TDK","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="100nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="ST"});
            Add("4E22000002","CAP","CAP 100nF 50V 10% X7R OM 0805 AECQ","CGA4J2X7R1H104K125AM","TDK","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="100nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="OM"});
            Add("4E22000002","CAP","CAP 100nF 50V 10% X7R OM 0805 AECQ","CL21B104KBFWPNE","Samsung","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="100nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="OM"});
            Add("4E22000002","CAP","CAP 100nF 50V 10% X7R ST 0805 AECQ","GCE21BR71H104KA01L","Murata","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="100nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="ST"});
            Add("4E22000002","CAP","CAP 100nF 50V 10% X7R ST 0805 AECQ","KAF21BR71H104KT","Kyocera","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="100nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="ST"});
            Add("4E22000003","CAP","CAP 1uF 50V 10% X7R NFM 0805 AECQ","CL21B105KBFVPJE","Samsung","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="1uF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000003","CAP","CAP 1uF 50V 10% X7R NFM 0805 AECQ","CL21B105KBFVPNE","Samsung","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="1uF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000003","CAP","CAP 1uF 50V 10% X7R NFM 0805 AECQ","CL21Y105KCFVPJE","Samsung","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="1uF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000003","CAP","CAP 1uF 50V 10% X7R ST 0805 AECQ","GCJ21BR71H105KA01L","Murata","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="1uF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="ST"});
            Add("4E22000003","CAP","CAP 1uF 50V 10% X7R NFM 0805 AECQ","GCM21BR71H105KA03L","Murata","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="1uF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000004","CAP","CAP 100nF 100V 10% X7S ST 0603 AECQ","CGA3E3X7S2A104K080AE","TDK","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="100nF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7S", ["Failure Mode"]="ST"});
            Add("4E22000004","CAP","CAP 100nF 100V 10% X7R ST 0805 AECQ","CGA4J2X7R2A104K125","TDK","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="100nF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="ST"});
            Add("4E22000004","CAP","CAP 100nF 100V 10% X7R OM 0805 AECQ","CL21B104KCFWPNE","Samsung","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="100nF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="OM"});
            Add("4E22000004","CAP","CAP 100nF 100V 10% X7R ST 0805 AECQ","GCJ21BR72A104KA01L","Murata","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="100nF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="ST"});
            Add("4E22000005","CAP","CAP 100pF 50V 5% NPO NFM 0603 AECQ","AC0603JRNPO9BN101","Yageo","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="100pF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="5%", ["Temperature Coefficient"]="NPO", ["Failure Mode"]="NFM"});
            Add("4E22000005","CAP","CAP 100pF 50V 5% C0G NFM 0603 AECQ","CL10C101JB81PNC","Samsung","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="100pF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="5%", ["Temperature Coefficient"]="C0G", ["Failure Mode"]="NFM"});
            Add("4E22000006","CAP","CAP 100nF 50V 10% X7R NFM 0603 AECQ","AC0603KRX7R9BB104","Yageo","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="100nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000006","CAP","CAP 100nF 50V 10% X7R OM 0603 AECQ","CL10B104KB85PNC","Samsung","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="100nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="OM"});
            Add("4E22000006","CAP","CAP 100nF 50V 10% X7R OM 0603 AECQ","CL10B104KB8WPNC","Samsung","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="100nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="OM"});
            Add("4E22000006","CAP","CAP 100nF 50V 10% X7R NFM 0603 AECQ","GCM188R71H104KA57D","Murata","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="100nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000007","CAP","CAP 47nF 100V 10% X7R OM 0603 AECQ","CL10B473KC8WPNC","Samsung","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="47nF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="OM"});
            Add("4E22000007","CAP","CAP 47nF 100V 10% X7R ST 0805 AECQ","GCJ21BR72A473KA01L","Murata","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="47nF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="ST"});
            Add("4E22000008","CAP","CAP 680pF 50V 5% NPO NFM 0805 AECQ","AC0805JRNPO9BN681","Yageo","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="680pF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="5%", ["Temperature Coefficient"]="NPO", ["Failure Mode"]="NFM"});
            Add("4E22000008","CAP","CAP 680pF 50V 5% C0G NFM 1206 STD","CL31C681JBCNNNC","Samsung","CAP-1206","CapacitorSymbol","CAP-1206",STD,SMD,new(){["Capacitance(F)"]="680pF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="5%", ["Temperature Coefficient"]="C0G", ["Failure Mode"]="NFM"});
            Add("4E22000008","CAP","CAP 680pF 50V 5% C0G NFM 0805 AECQ","GCM2165C1H681JA16D","Murata","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="680pF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="5%", ["Temperature Coefficient"]="C0G", ["Failure Mode"]="NFM"});
            Add("4E22000009","CAP","CAP 33pF 50V 5% NPO NFM 0603 AECQ","AC0603JRNPO9BN330","Yageo","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="33pF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="5%", ["Temperature Coefficient"]="NPO", ["Failure Mode"]="NFM"});
            Add("4E22000009","CAP","CAP 33pF 50V 5% C0G NFM 0603 AECQ","GCM1885C1H330JA16D","Murata","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="33pF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="5%", ["Temperature Coefficient"]="C0G", ["Failure Mode"]="NFM"});
            Add("4E22000010","CAP","CAP 10nF 50V 10% X7R ST 0603 AECQ","06035C103K4Z4A","Kyocera","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="10nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="ST"});
            Add("4E22000010","CAP","CAP 10nF 50V 10% X7R NFM 0603 AECQ","AC0603KRX7R9BB103","Yageo","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="10nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000010","CAP","CAP 10nF 50V 10% X7R NFM 0603 AECQ","CGA3E2X7R1H103K080AA","TDK","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="10nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000010","CAP","CAP 10nF 50V 10% X7R ST 0603 AECQ","CGA3E2X7R1H103K080AE","TDK","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="10nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="ST"});
            Add("4E22000010","CAP","CAP 10nF 50V 10% X7R OM 0603 AECQ","CL10B103KB8WPNC","Samsung","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="10nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="OM"});
            Add("4E22000010","CAP","CAP 10nF 50V 10% X7R NFM 0603 AECQ","GCM188R71H103KA37D","Murata","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="10nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000011","CAP","CAP 10nF 100V 10% X7R NFM 0805 AECQ","AC0805KRX7R0BB103","Yageo","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="10nF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000011","CAP","CAP 10nF 100V 10% X7R NFM 0805 AECQ","C0805C103K1RACAUTO","Kemet","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="10nF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000011","CAP","CAP 10nF 100V 10% X7R OM 0805 AECQ","CL21B103KC6WPNC","Samsung","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="10nF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="OM"});
            Add("4E22000012","CAP","CAP 47nF 50V 10% X7R NFM 0603 AECQ","AC0603KRX7R9BB473","Yageo","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="47nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000013","CAP","CAP 100nF 100V ? X7R NFM 0805 AECQ","CGA4J2X7R2A104xxxxxx","TDK","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="100nF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="?", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000013","CAP","CAP 2.2nF 50V 10% X7R ST 0805 AECQ","GCJ216R71H222KA","Murata","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="2.2nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="ST"});
            Add("4E22000013","CAP","CAP 100nF 100V ? X7R ST 0805 AECQ","GCJ21BR72A104xxxx","Murata","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="100nF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="?", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="ST"});
            Add("4E22000013","CAP","CAP 100nF 100V 10% X7R ST 0805 AECQ","GCJ21BR72A104KA01L","Murata","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="100nF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="ST"});
            Add("4E22000013","CAP","CAP 100nF 100V 10% X7R NFM 0805 AECQ","GCM21BR72A104KA37L","Murata","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="100nF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000014","CAP","CAP 1uF 100V ? X7S ? 0805 AECQ","CGA4J3X7S2A105xxxxxx","TDK","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="1uF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="?", ["Temperature Coefficient"]="X7S", ["Failure Mode"]="?"});
            Add("4E22000014","CAP","CAP 10nF 50V 10% X7R ST 0603 AECQ","GCE188R71H103K","Murata","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="10nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="ST"});
            Add("4E22000014","CAP","CAP 1uF 100V ? X7S ST 0805 AECQ","GCJ21BC72A105xxxx","Murata","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="1uF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="?", ["Temperature Coefficient"]="X7S", ["Failure Mode"]="ST"});
            Add("4E22000014","CAP","CAP 1uF 100V 10% X7S NFM 0805 AECQ","GCM21BC72A105KE36L","Murata","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="1uF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7S", ["Failure Mode"]="NFM"});
            Add("4E22000015","CAP","CAP 10nF 50V 10% X7R NFM 0805 AECQ","AC0805KRX7R9BB103","Yageo","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="10nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000015","CAP","CAP 10nF 50V 10% X7R NFM 0805 AECQ","C0805C103K5RACAUTO","Kemet","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="10nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000015","CAP","CAP 10nF 50V 10% X7R ST 0805 AECQ","CL21B103KB6WPNC","Samsung","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="10nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="ST"});
            Add("4E22000015","CAP","CAP 10nF 50V 10% X7R NFM 0805 AECQ","GCM216R71H103KA37D","Murata","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="10nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000016","CAP","CAP 220pF 50V 10% X7R NFM 0805 AECQ","C0805C221K5RACAUTO","Kemet","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="220pF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000017","CAP","CAP 1nF 50V 10% X7R OM 0805 AECQ","CL21B102KB6WPNC","Samsung","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="1nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="OM"});
            Add("4E22000017","CAP","CAP 1nF 50V 10% X7R NFM 0805 AECQ","GCM216R71H102KA37D","Murata","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="1nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000018","CAP","CAP 1nF 100V 10% X7R OM 0805 AECQ","CL21B102KC65PNC","Samsung","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="1nF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="OM"});
            Add("4E22000018","CAP","CAP 1nF 100V 10% X7R OM 0805 AECQ","CL21B102KC6WPNC","Samsung","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="1nF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="OM"});
            Add("4E22000019","CAP","CAP 100nF 100V 10% X7R NFM 0805 AECQ","GCM21BR72A104KA37L","Murata","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="100nF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000020","CAP","CAP 470pF 50V 10% X7R NFM 0805 AECQ","08055C471K4T2A","Samsung","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="470pF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000020","CAP","CAP 470pF 50V 10% X7R NFM 0805 AECQ","C0805C471K5RACAUTO","Kemet","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="470pF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000020","CAP","CAP 470pF 50V 10% X7R NFM 0805 STD","CC0805KRX7R9BB471","Yageo","CAP-0805","CapacitorSymbol","CAP-0805",STD,SMD,new(){["Capacitance(F)"]="470pF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000020","CAP","CAP 470pF 50V 5% C0G NFM 0805 AECQ","CL21C471JBC1PNC","Samsung","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="470pF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="5%", ["Temperature Coefficient"]="C0G", ["Failure Mode"]="NFM"});
            Add("4E22000021","CAP","CAP 1uF 100V 10% X7R ST 1206 AECQ","GCJ31CR72A105KA01L","Murata","CAP-1206","CapacitorSymbol","CAP-1206",AECQ,SMD,new(){["Capacitance(F)"]="1uF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="ST"});
            Add("4E22000022","CAP","CAP 4.7uF 100V 10% X7S NFM 1210 AECQ","GCM32DC72A475KE02L","Murata","CAP-1210","CapacitorSymbol","CAP-1210",AECQ,SMD,new(){["Capacitance(F)"]="4.7uF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7S", ["Failure Mode"]="NFM"});
            Add("4E22000023","CAP","CAP 10uF 50V 20% X7R NFM 1210 STD","C3225X7R1H106K250AC","TDK","CAP-1210","CapacitorSymbol","CAP-1210",STD,SMD,new(){["Capacitance(F)"]="10uF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="20%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000023","CAP","CAP 10uF 50V 10% X7S ST 1210 AECQ","GCJ32EC71H106KA01L","Murata","CAP-1210","CapacitorSymbol","CAP-1210",AECQ,SMD,new(){["Capacitance(F)"]="10uF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7S", ["Failure Mode"]="ST"});
            Add("4E22000023","CAP","CAP 10uF 50V 10% X7S NFM 1210 AECQ","GCM32EC71H106KA03L","Murata","CAP-1210","CapacitorSymbol","CAP-1210",AECQ,SMD,new(){["Capacitance(F)"]="10uF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7S", ["Failure Mode"]="NFM"});
            Add("4E22000024","CAP","CAP 2.2nF 50V 10% X7R ST 0805 AECQ","GCJ216R71H222KA01D","Murata","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="2.2nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="ST"});
            Add("4E22000024","CAP","CAP","NRS6045T470MMGKV","Taiyo Yuden","","","",STD,SMD,null);
            Add("4E22000025","CAP","CAP 2.2uF 50V 10% X7R ST 1206 AECQ","GCJ31CR71H225KA12L","Murata","CAP-1206","CapacitorSymbol","CAP-1206",AECQ,SMD,new(){["Capacitance(F)"]="2.2uF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="ST"});
            Add("4E22000026","CAP","CAP 100nF 100V 10% X7R ST 0603 AECQ","AS0603KRX7R0BB104","Yageo","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="100nF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="ST"});
            Add("4E22000026","CAP","CAP 100nF 100V 10% X7R ST 0603 AECQ","GCJ188R72A104KA01D","Murata","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="100nF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="ST"});
            Add("4E22000027","CAP","CAP 4.7nF 100V 10% X7R OM 0603 AECQ","CL10B472KC8WPNC","Samsung","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="4.7nF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="OM"});
            Add("4E22000027","CAP","CAP 4.7nF 100V 10% X7R ST 0603 AECQ","GCJ188R72A472KA01D","Murata","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="4.7nF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="ST"});
            Add("4E22000027","CAP","CAP 4.7nF 100V 10% X7R NFM 0603 AECQ","GCM188R72A472KA37D","Murata","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="4.7nF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000028","CAP","CAP 2.2uF 50V 10% X7R ST 1206 AECQ","GCJ31CR71H225KA12L","Murata","CAP-1206","CapacitorSymbol","CAP-1206",AECQ,SMD,new(){["Capacitance(F)"]="2.2uF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="ST"});
            Add("4E22000028","CAP","CAP 2.2uF 50V 10% X7R NFM 1206 AECQ","GCM31CR71H225KA55L","Murata","CAP-1206","CapacitorSymbol","CAP-1206",AECQ,SMD,new(){["Capacitance(F)"]="2.2uF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000029","CAP","CAP 100nF 50V 10% X7R NFM 0805 AECQ","GCM21BR71H104KA37L","Murata","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="100nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000030","CAP","CAP 4.7nF 50V 10% X7R OM 0805 AECQ","CL21B472KB65PNC","Samsung","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="4.7nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="OM"});
            Add("4E22000030","CAP","CAP 4.7nF 50V 5% X7R NFM 0805 AECQ","GCM216R71H472JA37D","Murata","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="4.7nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="5%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000031","CAP","CAP 100pF 50V 5% C0G NFM 0805 AECQ","CL21C101JB61PNC","Samsung","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="100pF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="5%", ["Temperature Coefficient"]="C0G", ["Failure Mode"]="NFM"});
            Add("4E22000031","CAP","CAP 100pF 50V 5% C0G NFM 0805 AECQ","GCM2165C1H101JA16D","Murata","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="100pF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="5%", ["Temperature Coefficient"]="C0G", ["Failure Mode"]="NFM"});
            Add("4E22000032","CAP","CAP 22uF 50V 20% ? ? Radial, Can AECQ","EEE1HA220WAP","Panasonic","Radial, Can","CapacitorSymbol","Radial, Can",AECQ,SMD,new(){["Capacitance(F)"]="22uF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="20%", ["Temperature Coefficient"]="?", ["Failure Mode"]="?"});
            Add("4E22000033","CAP","CAP 220nF 50V 10% X7R NFM 0805 AECQ","C0805C224K5RACAUTO","Kemet","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="220nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000033","CAP","CAP 220nF 50V 10% X7R NFM 0805 AECQ","CL21B224KBFVPNE","Samsung","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="220nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000034","CAP","CAP 10uF 50V 20% ? ? Radial, Can AECQ","EEE1HA100AP","Panasonic","Radial, Can","CapacitorSymbol","Radial, Can",AECQ,SMD,new(){["Capacitance(F)"]="10uF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="20%", ["Temperature Coefficient"]="?", ["Failure Mode"]="?"});
            Add("4E22000035","CAP","CAP 10uF 10V 10% X7R NFM 0805 AECQ","GCM21BR71A106KE22L","Murata","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="10uF", ["Voltage(V)"]="10V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000037","CAP","CAP 1uF 100V 10% X7R NFM 1206 AECQ","CGA5L2X7R2A105K160AA","TDK","CAP-1206","CapacitorSymbol","CAP-1206",AECQ,SMD,new(){["Capacitance(F)"]="1uF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000037","CAP","CAP 1uF 100V 10% X7R ST 1206 AECQ","GCJ31CR72A105KA01K","Murata","CAP-1206","CapacitorSymbol","CAP-1206",AECQ,SMD,new(){["Capacitance(F)"]="1uF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="ST"});
            Add("4E22000037","CAP","CAP 1uF 100V 10% X7R ST 1206 AECQ","GCJ31CR72A105KA01L","Murata","CAP-1206","CapacitorSymbol","CAP-1206",AECQ,SMD,new(){["Capacitance(F)"]="1uF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="ST"});
            Add("4E22000037","CAP","CAP 1uF 50V 20% X7R NFM 1206 AECQ","CGA5L3X7R1H105M160AB","TDK","CAP-1206","CapacitorSymbol","CAP-1206",AECQ,SMD,new(){["Capacitance(F)"]="1uF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="20%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000037","CAP","CAP 1uF 100V 10% X7R ST 1206 AECQ","12061C105K4Z4A","Kyocera","CAP-1206","CapacitorSymbol","CAP-1206",AECQ,SMD,new(){["Capacitance(F)"]="1uF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="ST"});
            Add("4E22000038","CAP","CAP 1nF 50V 10% X7R NFM 0603 AECQ","AC0603KPX7R9BB102","Yageo","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="1nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000038","CAP","CAP 1nF 50V 10% X7R NFM 0603 AECQ","AC0603KRX7R9BB102","Yageo","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="1nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000038","CAP","CAP 1nF 50V 10% X7R OM 0603 AECQ","CL10B102KB8WPNC","Samsung","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="1nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="OM"});
            Add("4E22000039","CAP","CAP 220nF 16V 10% X7R NFM 0603 STD","CL10B224KO8NNNC- NOT AC","Samsung","CAP-0603","CapacitorSymbol","CAP-0603",STD,SMD,new(){["Capacitance(F)"]="220nF", ["Voltage(V)"]="16V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000039","CAP","CAP 220nF 16V 10% X7R NFM 0603 AECQ","C0603C224K4RACAUTO7411","Kemet","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="220nF", ["Voltage(V)"]="16V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000039","CAP","CAP 220nF 16V 10% X7R NFM 0603 AECQ","CL10B224KO8VPNC","Samsung","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="220nF", ["Voltage(V)"]="16V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000039","CAP","CAP 220nF 16V 10% X7R ST 0603 AECQ","GCJ188R71C224KA01D","Murata","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="220nF", ["Voltage(V)"]="16V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="ST"});
            Add("4E22000040","CAP","CAP 1uF 50V 10% X5R NFM 0603 STD","CL10A105KB8NNNC-NOT AC","Samsung","CAP-0603","CapacitorSymbol","CAP-0603",STD,SMD,new(){["Capacitance(F)"]="1uF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X5R", ["Failure Mode"]="NFM"});
            Add("4E22000040","CAP","CAP 1uF 50V 10% X7R NFM 0603 STD","06035C105KAT2A","Kyocera","CAP-0603","CapacitorSymbol","CAP-0603",STD,SMD,new(){["Capacitance(F)"]="1uF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000040","CAP","CAP 1uF 50V 10% X5R NFM 0603 AECQ","GRT188R61H105KE13D","Murata","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="1uF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X5R", ["Failure Mode"]="NFM"});
            Add("4E22000041","CAP","CAP 10uF 16V 10% X7S NFM 0805 AECQ","CL21Y106KOBVPNE","Samsung","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="10uF", ["Voltage(V)"]="16V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7S", ["Failure Mode"]="NFM"});
            Add("4E22000041","CAP","CAP 10uF 16V 10% X7S NFM 0805 AECQ","CGA4J1X7S1C106K125AC","TDK","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="10uF", ["Voltage(V)"]="16V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7S", ["Failure Mode"]="NFM"});
            Add("4E22000041","CAP","CAP 10uF 16V 10% X7S NFM 0805 STD","GRM21BC71C106KE11L","Murata","CAP-0805","CapacitorSymbol","CAP-0805",STD,SMD,new(){["Capacitance(F)"]="10uF", ["Voltage(V)"]="16V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7S", ["Failure Mode"]="NFM"});
            Add("4E22000042","CAP","CAP 470pF 50V 10% X7R NFM 0603 AECQ","AC0603KRX7R9BB471","Yageo","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="470pF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000044","CAP","CAP 100nF 50V 10% X7R ST 0805 AECQ","08055C104K4Z2A","Kyocera","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="100nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="ST"});
            Add("4E22000044","CAP","CAP 100nF 50V 10% X7R ST 0805 AECQ","KAF21BR71H104KT","Kyocera","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="100nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="ST"});
            Add("4E22000044","CAP","CAP 100nF 50V 10% X7R ST 0805 AECQ","GCE21BR71H104KA01L","Murata","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="100nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="ST"});
            Add("4E22000044","CAP","CAP 100nF 50V 10% X7R OM 0805 AECQ","C0805F104K5RACAUTO","Kemet","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="100nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="OM"});
            Add("4E22000044","CAP","CAP 100nF 50V 10% X7R OM 0805 AECQ","CL21B104KBFWPNE","Samsung","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="100nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="OM"});
            Add("4E22000045","CAP","CAP 1nF 50V 5% C0G NFM 0805 AECQ","CL21C102JBC1PNC","Samsung","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="1nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="5%", ["Temperature Coefficient"]="C0G", ["Failure Mode"]="NFM"});
            Add("4E22000045","CAP","CAP 1nF 50V 10% X7R OM 0805 AECQ","CL21B102KB6WPNC","Samsung","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="1nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="OM"});
            Add("4E22000045","CAP","CAP 1nF 50V 10% X7R OM 0603 AECQ","CL10B102KB8WPNC","Samsung","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="1nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="OM"});
            Add("4E22000045","CAP","CAP 1nF 50V 10% X7R NFM 0805 AECQ","AC0805KRX7R9BB102","Samsung","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="1nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000046","CAP","CAP 1uF 100V 10% X7S NFM 0805 AECQ","CL21Y105KCFVPJE","Samsung","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="1uF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7S", ["Failure Mode"]="NFM"});
            Add("4E22000046","CAP","CAP 1uF 50V 10% X7R NFM 0805 AECQ","GCM21BR71H105KA03L","Murata","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="1uF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000047","CAP","CAP 1nF 50V 10% X7R NFM 0603 AECQ","AC0603KRX7R9BB102","Yageo","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="1nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000049","CAP","CAP 1uF 50V 10% X7R NFM 0805 AECQ","GCM21BR71H105KA03L","Murata","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="1uF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000050","CAP","CAP 1uF 50V 10% X7R ST 0805 AECQ","GCJ21BR71H105KA01L","Murata","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="1uF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="ST"});
            Add("4E22000051","CAP","CAP 2.2uF 50V 10% X7R NFM 1206 AECQ","GCM31CR71H225KA55L","Murata","CAP-1206","CapacitorSymbol","CAP-1206",AECQ,SMD,new(){["Capacitance(F)"]="2.2uF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000051","CAP","CAP 2.2uF 50V 10% X7R ST 1206 AECQ","GCJ31CR71H225KA12L","Murata","CAP-1206","CapacitorSymbol","CAP-1206",AECQ,SMD,new(){["Capacitance(F)"]="2.2uF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="ST"});
            Add("4E22000052","CAP","CAP 4.7nF 50V 10% X7R OM 0603 AECQ","CL10B472KB8WPNC","Samsung","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="4.7nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="OM"});
            Add("4E22000052","CAP","CAP 4.7nF 50V 10% X7R NFM 0603 AECQ","AC0603KRX7R9BB472","Yageo","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="4.7nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000052","CAP","CAP 4.7nF 50V 10% X7R NFM 0603 AECQ","C0603C472K5RACAUTO","Kemet","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="4.7nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000052","CAP","CAP 4.7nF 50V 10% X7R OM 0603 AECQ","CL10B472KB85PNC","Samsung","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="4.7nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="OM"});
            Add("4E22000053","CAP","CAP 2.2uF 50V 10% X7R ST 1206 AECQ","GCJ31CR71H225KA12L","Murata","CAP-1206","CapacitorSymbol","CAP-1206",AECQ,SMD,new(){["Capacitance(F)"]="2.2uF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="ST"});
            Add("4E22000054","CAP","CAP 100nF 50V 10% X7R NFM 0805 AECQ","AC0805KRX7R9BB104","Yageo","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="100nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000055","CAP","CAP 1uF 100V 10% X7S NFM 0805 AECQ","GCM21BC72A105KE36L","Murata","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="1uF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7S", ["Failure Mode"]="NFM"});
            Add("4E22000055","CAP","CAP 1uF 50V 20% X7R ST 0805 AECQ","CGA4J3X7R1H105M125AE","TDK","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="1uF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="20%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="ST"});
            Add("4E22000056","CAP","CAP 1nF 100V 10% X7R OM 0603 AECQ","CL10B102KC8WPNC","Samsung","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="1nF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="OM"});
            Add("4E22000056","CAP","CAP 1nF 100V 10% X7R ST 0603 AECQ","GCJ188R72A102KA01J","Murata","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="1nF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="ST"});
            Add("4E22000056","CAP","CAP 1nF 100V 10% X7R ST 0603 AECQ","CGA3E2X7R2A102K080AE","TDK","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="1nF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="ST"});
            Add("4E22000056","CAP","CAP 1nF 100V 10% X7R ST 0603 AECQ","GCJ188R72A102KA01D","Murata","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="1nF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="ST"});
            Add("4E22000057","CAP","CAP 100nF 100V 10% X7R ST 0805 AECQ","GCJ21BR72A104KA01L","Murata","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="100nF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="ST"});
            Add("4E22000057","CAP","CAP 100nF 100V 10% X7R OM 0805 AECQ","CL21B104KCFWPNE","Samsung","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="100nF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="OM"});
            Add("4E22000057","CAP","CAP 100nF 100V 10% X7R ST 0805 AECQ","GCJ21BR72A104KA01K","Murata","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="100nF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="ST"});
            Add("4E22000058","CAP","CAP 1uF 100V 10% X7S ST 0805 AECQ","GCJ21BC72A105KE02L","Murata","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="1uF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7S", ["Failure Mode"]="ST"});
            Add("4E22000058","CAP","CAP 1uF 100V 20% X7S ST 0805 AECQ","CGA4J3X7S2A105M125AE","TDK","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="1uF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="20%", ["Temperature Coefficient"]="X7S", ["Failure Mode"]="ST"});
            Add("4E22000058","CAP","CAP 1uF 100V 10% X7S ST 0805 STD","GRJ21BC72A105KE11L(ALTERNATIF)","Murata","CAP-0805","CapacitorSymbol","CAP-0805",STD,SMD,new(){["Capacitance(F)"]="1uF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7S", ["Failure Mode"]="ST"});
            Add("4E22000059","CAP","CAP 1uF 100V 10% X7S ST 0805 AECQ","CGA4J3X7S2A105K125AB","TDK","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="1uF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7S", ["Failure Mode"]="ST"});
            Add("4E22000059","CAP","CAP 1uF 100V 10% X7S ST 0805 AECQ","GCJ21BC72A105KE02L","Murata","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="1uF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7S", ["Failure Mode"]="ST"});
            Add("4E22000059","CAP","CAP 1uF 100V 10% X7S ST 0805 STD","GRJ21BC72A105KE11L","Murata","CAP-0805","CapacitorSymbol","CAP-0805",STD,SMD,new(){["Capacitance(F)"]="1uF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7S", ["Failure Mode"]="ST"});
            Add("4E22000060","CAP","CAP 100pF 100V 5% C0G NFM 0603 AECQ","CL10C101JC81PNC","Samsung","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="100pF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="5%", ["Temperature Coefficient"]="C0G", ["Failure Mode"]="NFM"});
            Add("4E22000060","CAP","CAP 100pF 100V 5% NPO NFM 0603 AECQ","AC0603JRNPO0BN101","Yageo","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="100pF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="5%", ["Temperature Coefficient"]="NPO", ["Failure Mode"]="NFM"});
            Add("4E22000061","CAP","CAP 10nF 100V 10% X7R NFM 0603 AECQ","MURGCM188R72A103KA37D","Murata","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="10nF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000061","CAP","CAP 10nF 100V 10% X7R NFM 0603 AECQ","GCM188R72A103KA37D","Murata","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="10nF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000062","CAP","CAP 1uF 100V 10% X7R OM 1210 AECQ","C1210F105K1RACAUTO","Kemet","CAP-1210","CapacitorSymbol","CAP-1210",AECQ,SMD,new(){["Capacitance(F)"]="1uF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="OM"});
            Add("4E22000063","CAP","CAP 100nF 100V 10% X7R OM 1210 AECQ","C1210F104K1RACAUTO","Kemet","CAP-1210","CapacitorSymbol","CAP-1210",AECQ,SMD,new(){["Capacitance(F)"]="100nF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="OM"});
            Add("4E22000063","CAP","CAP 100nF 100V 10% X7R NFM 1210 AECQ","HMF325B7104KNHT","Taiyo Yuden","CAP-1210","CapacitorSymbol","CAP-1210",AECQ,SMD,new(){["Capacitance(F)"]="100nF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000064","CAP","CAP 100nF 100V 10% X7R OM 0805 AECQ","CL21B104KCFWPNE","Samsung","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="100nF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="OM"});
            Add("4E22000064","CAP","CAP 100nF 50V 10% X7R NFM 0805 AECQ","GCM21BR72A104K","Murata","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="100nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000065","CAP","CAP 10nF 50V 10% X7R NFM 0603 AECQ","AC0603KRX7R9BB103","Yageo","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="10nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000065","CAP","CAP 10nF 50V 10% X7R OM 0603 AECQ","CL10B103KB8WPNC","Samsung","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="10nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="OM"});
            Add("4E22000065","CAP","CAP 10nF 50V 10% X7R NFM 0603 AECQ","CGA3E2X7R1H103K080AA","TDK","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="10nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000065","CAP","CAP 10nF 50V 10% X7R ST 0603 AECQ","GCE188R71H103KA01D","Murata","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="10nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="ST"});
            Add("4E22000065","CAP","CAP 10nF 50V 20% X7R FT 0603 AECQ","C0603X103M5RAC3121","Kemet","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="10nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="20%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="FT"});
            Add("4E22000066","CAP","CAP","YFA1N5819 TR","","","CapacitorSymbol","",STD,SMD,null);
            Add("4E22000066","CAP","CAP 1nF 50V 10% X7R NFM 0603 AECQ","C0603C102K5RACAUTO7411","Kemet","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="1nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000067","CAP","CAP 10nF 50V 10% X7R NFM 0805 AECQ","C0805C103K5RACAUTO","Kemet","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="10nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000067","CAP","CAP 10nF 50V 10% X7R NFM 0805 AECQ","GCM216R71H103KA37D","Murata","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="10nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000067","CAP","CAP 10nF 50V 10% X7R NFM 0805 AECQ","CL21B103KB6WPNC","Samsung","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="10nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000068","CAP","CAP 10nF 50V 10% X7R NFM 0603 AECQ","AC0603KRX7R9BB103","Yageo","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="10nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000069","CAP","CAP 100nF 100V 10% X7S NFM 0603 AECQ","CGA3E3X7S2A104K080AB","TDK","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="100nF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7S", ["Failure Mode"]="NFM"});
            Add("4E22000069","CAP","CAP 100nF 100V 10% X7R NFM 0805 AECQ","GCM21BR72A104KA37L","Murata","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="100nF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000070","CAP","CAP 4.7nF 100V 10% X7R NFM 0603 AECQ","GCM188R72A472KA37D","Murata","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="4.7nF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000070","CAP","CAP 4.7nF 100V 10% X7R ST 0603 AECQ","GCJ188R72A472KA01D","Murata","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="4.7nF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="ST"});
            Add("4E22000071","CAP","CAP 1uF 100V 10% X7S ST 0805 AECQ","GCJ21BC72A105KE02L","Murata","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="1uF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7S", ["Failure Mode"]="ST"});
            Add("4E22000071","CAP","CAP 4.7uF 100V 10% X7R NFM 1210 AECQ","12101C475K4T2A","Kyocera","CAP-1210","CapacitorSymbol","CAP-1210",AECQ,SMD,new(){["Capacitance(F)"]="4.7uF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000072","CAP","CAP 1uF 50V 10% X7R ST 0805 AECQ","GCJ21BR71H105KA01L","Murata","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="1uF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="ST"});
            Add("4E22000073","CAP","CAP 100nF 100V 10% X7R NFM 0805 AECQ","GCM21BR72A104KA37L","Murata","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="100nF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000074","CAP","CAP 10nF 50V 10% X7R ST 0603 AECQ","GCE188R71H103KA01D","Murata","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="10nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="ST"});
            Add("4E22000075","CAP","CAP 47nF 100V 10% X7R ST 0805 AECQ","GCJ21BR72A473KA01L","Murata","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="47nF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="ST"});
            Add("4E22000075","CAP","CAP 47nF 100V 20% X8R ST 0805 AECQ","CGA4J3X8R2A473M125AE","TDK","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="47nF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="20%", ["Temperature Coefficient"]="X8R", ["Failure Mode"]="ST"});
            Add("4E22000075","CAP","CAP 47nF 100V 10% X7R ST 0805 STD","CL21B473KCFNNNE","Samsung","CAP-0805","CapacitorSymbol","CAP-0805",STD,SMD,new(){["Capacitance(F)"]="47nF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="ST"});
            Add("4E22000076","CAP","CAP 470pF 50V 10% X7R NFM 0603 AECQ","KEMC0603C471K5RAC7411","Kemet","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="470pF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000076","CAP","CAP 470pF 50V 10% X7R NFM 0603 AECQ","C0603C471K5RACAUTO7411","Kemet","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="470pF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000077","CAP","CAP 4.7nF 50V 10% X7R NFM 0603 AECQ","CL10B472KB85PNC","Samsung","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="4.7nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000078","CAP","CAP 220nF 50V 10% X7R NFM 0603 AECQ","CGA3E3X7R1H224K080AB","TDK","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="220nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000079","CAP","CAP 10nF 50V 10% X7R NFM 0603 AECQ","GCM188R71H103KA37D","Murata","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="10nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000080","CAP","CAP 470pF 50V 10% X7R NFM 0603 AECQ","C0603C471K5RACAUTO7411","Kemet","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="470pF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000081","CAP","CAP 4.7uF 50V 20% X7R NFM 1210 AECQ","UMK325B7475MMHP","Taiyo Yuden","CAP-1210","CapacitorSymbol","CAP-1210",AECQ,SMD,new(){["Capacitance(F)"]="4.7uF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="20%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000082","CAP","CAP 68nF 50V 10% X7R OM 0603 AECQ","CL10B683KB8WPNC","Samsung","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="68nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="OM"});
            Add("4E22000082","CAP","CAP 68nF 50V 10% X7R OM 0603 AECQ","CL10B683KB85PNC","Samsung","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="68nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="OM"});
            Add("4E22000083","CAP","CAP 100nF 100V 10% X7R FT 0805 AECQ","C0805X104K1RACAUTO","Kemet","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="100nF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="FT"});
            Add("4E22000084","CAP","CAP 100nF 100V 10% X7R FT 0805 AECQ","KEMC0805X104K1RACAUTO","Kemet","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="100nF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="FT"});
            Add("4E22000084","CAP","CAP 100nF 100V 10% X7R ST 0805 AECQ","GCJ21BR72A104KA01L","Murata","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="100nF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="ST"});
            Add("4E22000084","CAP","CAP 100nF 100V 10% X7R FT 0805 AECQ","C0805X104K1RACAUTO","Kemet","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="100nF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="FT"});
            Add("4E22000085","CAP","CAP 33nF 50V 10% X7R NFM 0603 AECQ","CGA3E2X7R1H333K080AA","TDK","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="33nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000086","CAP","CAP 4.7nF 50V 10% X7R NFM 0603 AECQ","CL10B472KB8NNNC","Samsung","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="4.7nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000086","CAP","CAP 4.7nF 50V 10% X7R OM 0603 AECQ","SMSCL10B472KB85PNC","Samsung","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="4.7nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="OM"});
            Add("4E22000086","CAP","CAP 4.7nF 50V 10% X7R OM 0603 AECQ","CL10B472KB8WPNC","Samsung","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="4.7nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="OM"});
            Add("4E22000087","CAP","CAP 1uF 35V 10% ? NFM 1206 AECQ","T498A105K035ATE10K","Kemet","CAP-1206","CapacitorSymbol","CAP-1206",AECQ,SMD,new(){["Capacitance(F)"]="1uF", ["Voltage(V)"]="35V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="?", ["Failure Mode"]="NFM"});
            Add("4E22000088","CAP","CAP 100nF 100V 10% X7R FT 1206 AECQ","12061C104K4Z2A","Kyocera","CAP-1206","CapacitorSymbol","CAP-1206",AECQ,SMD,new(){["Capacitance(F)"]="100nF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="FT"});
            Add("4E22000089","CAP","CAP 100pF 50V 5% C0G NFM 0603 AECQ","CGA3E2C0G1H101J080AA","TDK","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="100pF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="5%", ["Temperature Coefficient"]="C0G", ["Failure Mode"]="NFM"});
            Add("4E22000090","CAP","CAP 100nF 50V 10% X7R NFM 0603 AECQ","CGA3E2X7R1H104K080AA","TDK","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="100nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000091","CAP","CAP 100nF 50V 10% X8L NFM 0603 AECQ","06035L104K4T2A","Kyocera","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="100nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X8L", ["Failure Mode"]="NFM"});
            Add("4E22000092","CAP","CAP 10nF 50V 10% X7R NFM 0603 AECQ","MUR GCM188R71H103KA37D","Murata","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="10nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000093","CAP","CAP 1nF 50V 10% X7R NFM 0603 AECQ","CGA3E2X7R1H102K080AA","TDK","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="1nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000094","CAP","CAP 2.2nF 50V 10% X7R NFM 0805 AECQ","C0805C222K5RAC3123","Kemet","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="2.2nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000095","CAP","CAP 10nF 50V 10% X7R NFM 1206 AECQ","C1206C103K5RAC3123","Kemet","CAP-1206","CapacitorSymbol","CAP-1206",AECQ,SMD,new(){["Capacitance(F)"]="10nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000096","CAP","CAP 4.7nF 50V 10% X7R OM 0603 AECQ","CL10B472KB85PNC","Samsung","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="4.7nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="OM"});
            Add("4E22000096","CAP","CAP 4.7nF 50V 10% X7R NFM 0603 AECQ","C0603C472K5RACAUTO","Kemet","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="4.7nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000097","CAP","CAP 2.2nF 50V 10% X7R NFM 0805 AECQ","C0805C222K5RAC3123","Kemet","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="2.2nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000098","CAP","CAP 1uF 50V 10% X7R OM 1206 AECQ","C1206F105K5RACAUTO","Kemet","CAP-1206","CapacitorSymbol","CAP-1206",AECQ,SMD,new(){["Capacitance(F)"]="1uF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="OM"});
            Add("4E22000099","CAP","CAP 2.2uF 100V 20% X7S NFM 1206 AECQ","CGA5L3X7S2A225M160AB","TDK","CAP-1206","CapacitorSymbol","CAP-1206",AECQ,SMD,new(){["Capacitance(F)"]="2.2uF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="20%", ["Temperature Coefficient"]="X7S", ["Failure Mode"]="NFM"});
            Add("4E22000100","CAP","CAP 10nF 50V 10% X7R NFM 0603 AECQ","AC0603KRX7R9BB103","Yageo","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="10nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000101","CAP","CAP 1uF 50V 10% X7R NFM 0805 AECQ","GCM21BR71H105KA03L","Murata","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="1uF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000101","CAP","CAP 1uF 50V 10% X7R NFM 0805 AECQ","CL21B105KBFVPNE","Samsung","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="1uF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000102","CAP","CAP 1nF 50V 10% X7R OM 0603 AECQ","CL10B102KB85PNL","Samsung","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="1nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="OM"});
            Add("4E22000102","CAP","CAP 1nF 50V 10% X7R OM 0603 AECQ","AC0603KRX7R9BB102","Yageo","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="1nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="OM"});
            Add("4E22000103","CAP","CAP 1uF 100V 10% X7S NFM 0805 AECQ","GCM21BC72A105KE36L","Murata","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="1uF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7S", ["Failure Mode"]="NFM"});
            Add("4E22000104","CAP","CAP 10nF 50V 10% X7R OM 0603 AECQ","CL10B103KB8WPNC","Samsung","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="10nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="OM"});
            Add("4E22000104","CAP","CAP 10nF 50V 10% X7R OM 0603 AECQ","CL10B103KB85PNL","Samsung","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="10nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="OM"});
            Add("4E22000105","CAP","CAP 4.7nF 100V 10% X7R OM 0603 AECQ","CL10B472KC8WPNC","Samsung","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="4.7nF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="OM"});
            Add("4E22000105","CAP","CAP 4.7nF 100V 10% X7R NFM 0603 AECQ","CGA3E2X7R2A472K080AA","TDK","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="4.7nF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000106","CAP","CAP 1nF 100V 10% X7R OM 0603 AECQ","CL10B102KC8WPNC","Samsung","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="1nF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="OM"});
            Add("4E22000106","CAP","CAP 1nF 100V 20% X7R NFM 0603 AECQ","CGA3E2X7R2A102M080AA","TDK","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="1nF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="20%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000107","CAP","CAP 1uF 100V 10% X7R NFM 0805 AECQ","08051C105K4T2A","Kyocera","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="1uF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000107","CAP","CAP 1uF 100V 10% X7S NFM 0805 AECQ","GCM21BC72A105KE36L","Murata","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="1uF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7S", ["Failure Mode"]="NFM"});
            Add("4E22000108","CAP","CAP 10nF 50V 10% X7R OM 0603 AECQ","CL10B103KB8WPNC","Samsung","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="10nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="OM"});
            Add("4E22000108","CAP","CAP 10nF 50V 10% X7R NFM 0603 AECQ","C0603C103K5RACAUTO","Kemet","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="10nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000109","CAP","CAP 4.7uF 50V 10% X7R NFM 1206 STD","CL31B475KBHNNNE- NOT AC","Samsung","CAP-1206","CapacitorSymbol","CAP-1206",STD,SMD,new(){["Capacitance(F)"]="4.7uF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000109","CAP","CAP 4.7uF 50V 10% X7R NFM 1206 AECQ","CL31B475KBHVPNE","Samsung","CAP-1206","CapacitorSymbol","CAP-1206",AECQ,SMD,new(){["Capacitance(F)"]="4.7uF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000109","CAP","CAP 4.7uF 50V 10% X7R NFM 1206 AECQ","12065C475K4T2A","Kyocera","CAP-1206","CapacitorSymbol","CAP-1206",AECQ,SMD,new(){["Capacitance(F)"]="4.7uF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000110","CAP","CAP 1uF 100V 10% X7S NFM 0805 AECQ","GCM21BC72A105KE36L","Murata","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="1uF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7S", ["Failure Mode"]="NFM"});
            Add("4E22000112","CAP","CAP 4.7uF 50V 10% X7R ST 0805 AECQ","CGA4J1X7R1H475K125AE","TDK","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="4.7uF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="ST"});
            Add("4E22000113","CAP","CAP 1uF 25V 10% X5R NFM 0402 STD","GRM155R61E105KA12D","Murata","CAP-0402","CapacitorSymbol","CAP-0402",STD,SMD,new(){["Capacitance(F)"]="1uF", ["Voltage(V)"]="25V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X5R", ["Failure Mode"]="NFM"});
            Add("4E22000113","CAP","CAP 1uF 25V 10% X5R NFM 0402 AECQ","GRT155R61E105KE01D","Murata","CAP-0402","CapacitorSymbol","CAP-0402",AECQ,SMD,new(){["Capacitance(F)"]="1uF", ["Voltage(V)"]="25V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X5R", ["Failure Mode"]="NFM"});
            Add("4E22000114","CAP","CAP 100nF 50V 10% X7R NFM 0402 AECQ","GCM155R71H104KE02D","Murata","CAP-0402","CapacitorSymbol","CAP-0402",AECQ,SMD,new(){["Capacitance(F)"]="100nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000114","CAP","CAP 100nF 50V 10% X7R NFM 0402 AECQ","AC0402KRX7R9BB104","Yageo","CAP-0402","CapacitorSymbol","CAP-0402",AECQ,SMD,new(){["Capacitance(F)"]="100nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000115","CAP","CAP 4.7uF 16V 10% X7R NFM 0805 AECQ","GCM21BR71C475KA73L","Murata","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="4.7uF", ["Voltage(V)"]="16V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000115","CAP","CAP 4.7uF 16V 10% X7R NFM 0805 AECQ","CL21B475KOQVPNE","Samsung","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="4.7uF", ["Voltage(V)"]="16V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000116","CAP","CAP 100nF 50V 10% X7R NFM 0603 AECQ","AC0603KRX7R9BB104","Yageo","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="100nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000116","CAP","CAP 1nF 100V 10% X7R OM 0603 AECQ","CL10B102KC8WPNC","Samsung","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="1nF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="OM"});
            Add("4E22000117","CAP","CAP 12pF 50V 5% C0G NFM 0603 AECQ","C0603C120J5GACAUTO","Kemet","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="12pF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="5%", ["Temperature Coefficient"]="C0G", ["Failure Mode"]="NFM"});
            Add("4E22000118","CAP","CAP 10uF 50V 10% X7S NFM 1206 AECQ","CL31Y106KBKVPNE","Samsung","CAP-1206","CapacitorSymbol","CAP-1206",AECQ,SMD,new(){["Capacitance(F)"]="10uF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7S", ["Failure Mode"]="NFM"});
            Add("4E22000118","CAP","CAP 10uF 50V 10% X7R NFM 1206 AECQ","CGA5L1X7R1H106K160AC","TDK","CAP-1206","CapacitorSymbol","CAP-1206",AECQ,SMD,new(){["Capacitance(F)"]="10uF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000118","CAP","CAP 10uF 50V 10% X7R NFM 1206 STD","CL31B106KBHNNNE (O'SIZ)","Samsung","CAP-1206","CapacitorSymbol","CAP-1206",STD,SMD,new(){["Capacitance(F)"]="10uF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000119","CAP","CAP 220pF 50V 10% C0G NFM 0603 AECQ","GCM1885C1H221JA16D","Murata","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="220pF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="C0G", ["Failure Mode"]="NFM"});
            Add("4E22000119","CAP","CAP 220pF 50V 10% X7R OM 0603 AECQ","CL10B221KB8WPNC","Samsung","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="220pF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="OM"});
            Add("4E22000120","CAP","CAP 220uF 16V 20%   Radial, Can AECQ","UCD1C221MCL6GS","Nichicon","Radial, Can","CapacitorSymbol","Radial, Can",AECQ,SMD,new(){["Capacitance(F)"]="220uF", ["Voltage(V)"]="16V", ["Tolerance(%)"]="20%"});
            Add("4E22000121","CAP","CAP 27pF 50V 5% C0G NFM 0603 AECQ","CL10C270JB81PNC","Samsung","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="27pF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="5%", ["Temperature Coefficient"]="C0G", ["Failure Mode"]="NFM"});
            Add("4E22000122","CAP","CAP 4.7nF 50V 10% X7R NFM 0603 AECQ","C0603C472K5RAC3121","Kemet","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="4.7nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000122","CAP","CAP 4.7nF 100V 10% X7R OM 0603 AECQ","CL10B472KC8WPNC","Samsung","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="4.7nF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="OM"});
            Add("4E22000123","CAP","CAP 22uF 63V 20%   Radial, Can AECQ","GYB1J220MCQ1GS","Nichicon","Radial, Can","CapacitorSymbol","Radial, Can",AECQ,SMD,new(){["Capacitance(F)"]="22uF", ["Voltage(V)"]="63V", ["Tolerance(%)"]="20%"});
            Add("4E22000124","CAP","CAP 22uF 16V 10% X7R NFM 1210 AECQ","SMSCL32B226KOJVPNE","Samsung","CAP-1210","CapacitorSymbol","CAP-1210",AECQ,SMD,new(){["Capacitance(F)"]="22uF", ["Voltage(V)"]="16V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000125","CAP","CAP 10uF 50V 10% X7S ST 1210 AECQ","GCJ32EC71H106KA01L","Murata","CAP-1210","CapacitorSymbol","CAP-1210",AECQ,SMD,new(){["Capacitance(F)"]="10uF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7S", ["Failure Mode"]="ST"});
            Add("4E22000126","CAP","CAP 10nF 50V 10% X7R NFM 0603 STD","CL10B103KB8NNNC","Samsung","CAP-0603","CapacitorSymbol","CAP-0603",STD,SMD,new(){["Capacitance(F)"]="10nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000127","CAP","CAP 47nF 50V 10% X7R ST 0805 AECQ","AS0805KRX7R9BB473","Yageo","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="47nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="ST"});
            Add("4E22000127","CAP","CAP 47nF 50V 10% X7R NFM 0805 AECQ","GCD21BR71H473KA01L","Murata","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="47nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000127","CAP","CAP 47nF 50V 10% X7R ST 0805 AECQ","GCE21BR71H473KA01L","Murata","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="47nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="ST"});
            Add("4E22000127","CAP","CAP 47nF 50V 10% X7R ST 0805 AECQ","AS0805KRX7R9BB473","Yageo","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="47nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="ST"});
            Add("4E22000130","CAP","CAP 4.7nF 100V 10% X7R ST 0603 AECQ","GCJ188R72A472KA01","Murata","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="4.7nF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="ST"});
            Add("4E22000131","CAP","CAP 1nF 50V 10% X7R OM 0603 AECQ","CL10B102KB8WPNC","Samsung","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="1nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="OM"});
            Add("4E22000133","CAP","CAP 2.2uF 100V 10% X7R FT 1206 AECQ","12061C225K4Z2A","Kyocera","CAP-1206","CapacitorSymbol","CAP-1206",AECQ,SMD,new(){["Capacitance(F)"]="2.2uF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="FT"});
            Add("4E22000135","CAP","CAP 470pF 50V 10% X7R NFM 0402 AECQ","CL05B471KB5VPNC","Samsung","CAP-0402","CapacitorSymbol","CAP-0402",AECQ,SMD,new(){["Capacitance(F)"]="470pF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000135","CAP","CAP 470pF 50V 10% X7R NFM 0402 AECQ","GCM155R71H471KA37D","Murata","CAP-0402","CapacitorSymbol","CAP-0402",AECQ,SMD,new(){["Capacitance(F)"]="470pF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000136","CAP","CAP 10nF 50V 10% X7R NFM 0402 AECQ","CL05B103KB5VPNC","Samsung","CAP-0402","CapacitorSymbol","CAP-0402",AECQ,SMD,new(){["Capacitance(F)"]="10nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000136","CAP","CAP 1nF 100V 10% X7R ST 0603 AECQ","GCJ188R72A102KA01J","Murata","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="1nF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="ST"});
            Add("4E22000136","CAP","CAP 1nF 100V 10% X7R OM 0603 AECQ","CL10B102KC8WPNC","Samsung","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="1nF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="OM"});
            Add("4E22000136","CAP","CAP 10nF 50V 10% X7R NFM 0402 AECQ","C0402C103K5RACAUTOMLCC","Kemet","CAP-0402","CapacitorSymbol","CAP-0402",AECQ,SMD,new(){["Capacitance(F)"]="10nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000136","CAP","CAP 10nF 50V 10% X7R NFM 0402 AECQ","GCM155R71H103KA55D","Murata","CAP-0402","CapacitorSymbol","CAP-0402",AECQ,SMD,new(){["Capacitance(F)"]="10nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000136","CAP","CAP 10nF 50V 10% X7R NFM 0402 AECQ","CL05B103KB5VPNC","Samsung","CAP-0402","CapacitorSymbol","CAP-0402",AECQ,SMD,new(){["Capacitance(F)"]="10nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000137","CAP","CAP 10nF 50V 10% X7R FT 1206 AECQ","C1206X103K5RACAUTO","Kemet","CAP-1206","CapacitorSymbol","CAP-1206",AECQ,SMD,new(){["Capacitance(F)"]="10nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="FT"});
            Add("4E22000137","CAP","CAP 100nF 100V 10% X7R OM 0805 AECQ","CL21B104KCFWPNE","Samsung","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="100nF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="OM"});
            Add("4E22000139","CAP","CAP 10nF 50V 10% X7R ST 0603 AECQ","CGA3E2X7R1H103K080AE","TDK","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="10nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="ST"});
            Add("4E22000140","CAP","CAP 5.6nF 10V 1% X8R NFM 0603 AECQ","C0603C562F8HACAUTO","Kemet","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="5.6nF", ["Voltage(V)"]="10V", ["Tolerance(%)"]="1%", ["Temperature Coefficient"]="X8R", ["Failure Mode"]="NFM"});
            Add("4E22000141","CAP","CAP 10uF 50V 10% X7S ST 1210 AECQ","GCJ32EC71H106KA01L","Murata","CAP-1210","CapacitorSymbol","CAP-1210",AECQ,SMD,new(){["Capacitance(F)"]="10uF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7S", ["Failure Mode"]="ST"});
            Add("4E22000143","CAP","CAP 1uF 100V 10% X7S ST 0805 AECQ","GCJ21BC72A105KE02L","Murata","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="1uF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7S", ["Failure Mode"]="ST"});
            Add("4E22000146","CAP","CAP 6.8nF 50V 10% X7R NFM 0603 AECQ","C0603C682K5RACAUTO","Kemet","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="6.8nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000147","CAP","CAP 47pF 50V 5% NP0  NFM 0805 AECQ","08055A470J4T2A","Kyocera","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="47pF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="5%", ["Temperature Coefficient"]="NP0", ["Failure Mode"]="NFM"});
            Add("4E22000148","CAP","CAP 22uF 16V 20% X5R NFM 0805 AECQ","GRT21BR61C226ME13L","Murata","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="22uF", ["Voltage(V)"]="16V", ["Tolerance(%)"]="20%", ["Temperature Coefficient"]="X5R", ["Failure Mode"]="NFM"});
            Add("4E22000149","CAP","CAP 27pF 50V 5% C0G OM 0805 AECQ","CL21C270JB61PNC","Samsung","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="27pF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="5%", ["Temperature Coefficient"]="C0G", ["Failure Mode"]="OM"});
            Add("4E22000150","CAP","CAP 5.6nF 50V 10% X7R NFM 0805 AECQ","08055C562K4T2A","Kyocera","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="5.6nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000151","CAP","CAP 100pF 50V 5% C0G NFM 0805 AECQ","CL21C101JB61PNC","Samsung","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="100pF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="5%", ["Temperature Coefficient"]="C0G", ["Failure Mode"]="NFM"});
            Add("4E22000154","CAP","CAP 220nF 50V 20% X7R FT 1206 STD","1206Y0500224MXTE03","Knowles Syfer","CAP-1206","CapacitorSymbol","CAP-1206",STD,SMD,new(){["Capacitance(F)"]="220nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="20%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="FT"});
            Add("4E22000155","CAP","CAP 68pF 50V 10% X7R NFM 0603 STD","06035C680KAT2A","Kyocera","CAP-0603","CapacitorSymbol","CAP-0603",STD,SMD,new(){["Capacitance(F)"]="68pF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000156","CAP","CAP 12nF 100V 5% X7R NFM 0603 AECQ","C0603C123J1RACAUTO","Kemet","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="12nF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="5%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000158","CAP","CAP 470nF 50V 10% X7R NFM 0805 AECQ","CL21B474KBFVPNE","Samsung","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="470nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000159","CAP","CAP 4.7pF 100V 0.53% C0G NFM 0603 AECQ","GCM1885C2A4R7CA16D","Murata","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="4.7pF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="0.53%", ["Temperature Coefficient"]="C0G", ["Failure Mode"]="NFM"});
            Add("4E22000160","CAP","CAP 470nF 50V 10% X7R NFM 0805 AECQ","GCM21BR71H474KA55L","Murata","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="470nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000160","CAP","CAP 470nF 50V 10% X7R NFM 0603 AECQ","CGA3E3X7R1H474K080AB","TDK","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="470nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000161","CAP","CAP 4.7uF 50V 10% X7R ST 0805 AECQ","CGA4J1X7R1H475K125AE","TDK","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="4.7uF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="ST"});
            Add("4E22000161","CAP","CAP 4.7uF 50V 10% X7S NFM 0805 AECQ","GRT21BC71H475KE13L","Murata","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="4.7uF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7S", ["Failure Mode"]="NFM"});
            Add("4E22000161","CAP","CAP 4.7uF 50V 10% X7S NFM 0805 AECQ","CL21Y475KBBVPJE","Samsung","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="4.7uF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7S", ["Failure Mode"]="NFM"});
            Add("4E22000161","CAP","CAP 4.7uF 50V 10% X5R NFM 0805 AECQ","GRT21BR61H475KE13K","Murata","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="4.7uF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X5R", ["Failure Mode"]="NFM"});
            Add("4E22000162","CAP","CAP 470nF 50V 10% X7R NFM 0805 AECQ","CL21B474KBFVPNE","Samsung","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="470nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000163","CAP","CAP 5.6nF 50V 5% C0G NFM 0603 AECQ","C0603C562J5GACAUTO","Kemet","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="5.6nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="5%", ["Temperature Coefficient"]="C0G", ["Failure Mode"]="NFM"});
            Add("4E22000164","CAP","CAP 2.7nF 50V 5% C0G NFM 0603 AECQ","GCM1885C1H272JA16D","Murata","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="2.7nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="5%", ["Temperature Coefficient"]="C0G", ["Failure Mode"]="NFM"});
            Add("4E22000165","CAP","CAP 47nF 50V 10% X7R NFM 1206 AECQ","C1206C473K5RECAUTO","Kemet","CAP-1206","CapacitorSymbol","CAP-1206",AECQ,SMD,new(){["Capacitance(F)"]="47nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000166","CAP","CAP 4.7nF 50V 10% X7R NFM 0402 AECQ","CL05B472KB5VPNC","Samsung","CAP-0402","CapacitorSymbol","CAP-0402",AECQ,SMD,new(){["Capacitance(F)"]="4.7nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000167","CAP","CAP 47pF 50V 10% X8R NFM 0603 AECQ","C0603C470K5HACAUTO","Kemet","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="47pF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X8R", ["Failure Mode"]="NFM"});
            Add("4E22000168","CAP","CAP 1uF 50V 20% X7R ST 0805 AECQ","CGA4J3X7R1H105M125AE","TDK","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="1uF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="20%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="ST"});
            Add("4E22000169","CAP","CAP 220nF 100V 10% X7R FT 0805 STD","C0805X224K1REC7210","Kemet","CAP-0805","CapacitorSymbol","CAP-0805",STD,SMD,new(){["Capacitance(F)"]="220nF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="FT"});
            Add("4E22000169","CAP","CAP 220nF 100V 10% X7R FT 0805 AECQ","KAF21KR72A224KU","Kyocera","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="220nF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="FT"});
            Add("4E22000170","CAP","CAP 100nF 100V 10% X7R OM 0805 AECQ","CL21B104KCFWPNE","Samsung","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="100nF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="OM"});
            Add("4E22000171","CAP","CAP 470nF 100V 10% X7R NFM 0805 AECQ","KAM21KR72A474KU","Kyocera","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="470nF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000171","CAP","CAP 470nF 100V 10% X7R FT 0805 AECQ","08051C474K4Z2A","Kyocera","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="470nF", ["Voltage(V)"]="100V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="FT"});
            Add("4E22000172","CAP","CAP 330pF 50V 5% C0G NFM 0805 AECQ","CL21C331JB61PNC","Samsung","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="330pF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="5%", ["Temperature Coefficient"]="C0G", ["Failure Mode"]="NFM"});
            Add("4E22000172","CAP","CAP 330pF 50V 5% NP0  NFM 0805 AECQ","AC0805JRNPO9BN331","Yageo","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="330pF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="5%", ["Temperature Coefficient"]="NP0", ["Failure Mode"]="NFM"});
            Add("4E22000172","CAP","CAP 330pF 50V 5% C0G NFM 0805 AECQ","C0805C331J5GAUTO","Kemet","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="330pF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="5%", ["Temperature Coefficient"]="C0G", ["Failure Mode"]="NFM"});
            Add("4E22000174","CAP","CAP 4.7uF 50V 10% X7R ST 1206 AECQ","CGA5L3X7R1H475K160AE","TDK","CAP-1206","CapacitorSymbol","CAP-1206",AECQ,SMD,new(){["Capacitance(F)"]="4.7uF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="ST"});
            Add("4E22000177","CAP","CAP 1uF 50V 10% X7T NFM 0603 AECQ","GCM188D71H105KE36J","Murata","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="1uF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7T", ["Failure Mode"]="NFM"});
            Add("4E22000177","CAP","CAP 1uF 50V 10% X7R NFM 0603 AECQ","CL10B105KB9VPJC","Samsung","CAP-0603","CapacitorSymbol","CAP-0603",AECQ,SMD,new(){["Capacitance(F)"]="1uF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            Add("4E22000178","CAP","CAP 2.2uF 50V 10% X7R ST 0805 AECQ","CGA4J3X7R1H225K125AE","TDK","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="2.2uF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="ST"});
            Add("4E22000179","CAP","CAP 4.7nF 50V 10% X7R NFM 0805 AECQ","08055C472K4T4A","Kyocera","CAP-0805","CapacitorSymbol","CAP-0805",AECQ,SMD,new(){["Capacitance(F)"]="4.7nF", ["Voltage(V)"]="50V", ["Tolerance(%)"]="10%", ["Temperature Coefficient"]="X7R", ["Failure Mode"]="NFM"});
            //  INDUCTORS 
            Add("4E02000002","IND","IND 56uH 2.3A 20% 100mOhm 12.5x12.5x6(Shielded) AECQ","SRR1260A-560M","Bourns","IND-12.5x12.5x6(Shielded)","InductorSymbol","IND-12.5x12.5x6(Shielded)",AECQ,SMD,new(){["Inductance(H)"]="56uH", ["Irms Current(A)"]="2.3A", ["Tolerance(%)"]="20%", ["Max DCR(Ohm)"]="100mOhm"});
            Add("4E02000009","IND","IND 47uH 2.9A 20% 164mOhm 7.4x7x4.8(Shielded) AECQ","ETQP4M470KFM","Panasonic","IND-7.4x7x4.8(Shielded)","InductorSymbol","IND-7.4x7x4.8(Shielded)",AECQ,SMD,new(){["Inductance(H)"]="47uH", ["Irms Current(A)"]="2.9A", ["Tolerance(%)"]="20%", ["Max DCR(Ohm)"]="164mOhm"});
            Add("4E02000011","IND","IND 4.7uH 5.5A 20% 40mOhm 7.3x6.6x2.8(Shielded) AECQ","PA4341.472NLT","Pulse","IND-7.3x6.6x2.8(Shielded)","InductorSymbol","IND-7.3x6.6x2.8(Shielded)",AECQ,SMD,new(){["Inductance(H)"]="4.7uH", ["Irms Current(A)"]="5.5A", ["Tolerance(%)"]="20%", ["Max DCR(Ohm)"]="40mOhm"});
            //  FERRITE BEADS 
            Add("4E02000001","FB","FB 120Ohm@100MHz 3.5A 20mOhm 1206 AECQ","BLM31PG121SZ1L","Murata","FB-1206","FerriteBeadSymbol","FB-1206",AECQ,SMD,new(){["Impedance(Ohm)@Frequency(Hz)"]="120Ohm@100MHz", ["Max Rated Current(A)"]="3.5A", ["Max DCR(Ohm)"]="20mOhm"});
            Add("4E02000001","FB","FB 120Ohm@100MHz 6A 9mOhm 1206 AECQ","BLM31KN121SH1L","Murata","FB-1206","FerriteBeadSymbol","FB-1206",AECQ,SMD,new(){["Impedance(Ohm)@Frequency(Hz)"]="120Ohm@100MHz", ["Max Rated Current(A)"]="6A", ["Max DCR(Ohm)"]="9mOhm"});
            Add("4E02000003","FB","FB 600Ohm@100MHz 1A 150mOhm 0603 AECQ","BLM18KG601BH1D","Murata","FB-0603","FerriteBeadSymbol","FB-0603",AECQ,SMD,new(){["Impedance(Ohm)@Frequency(Hz)"]="600Ohm@100MHz", ["Max Rated Current(A)"]="1A", ["Max DCR(Ohm)"]="150mOhm"});
            Add("4E02000004","FB","FB 1KOhm@100MHz 1A 200mOhm 0603 AECQ","BLM18KG102SH1D","Murata","FB-0603","FerriteBeadSymbol","FB-0603",AECQ,SMD,new(){["Impedance(Ohm)@Frequency(Hz)"]="1KOhm@100MHz", ["Max Rated Current(A)"]="1A", ["Max DCR(Ohm)"]="200mOhm"});
            Add("4E02000005","FB","FB 600Ohm@100MHz 0.2A 900mOhm 1206 AECQ","BLM31AJ601SH1L","Murata","FB-1206","FerriteBeadSymbol","FB-1206",AECQ,SMD,new(){["Impedance(Ohm)@Frequency(Hz)"]="600Ohm@100MHz", ["Max Rated Current(A)"]="0.2A", ["Max DCR(Ohm)"]="900mOhm"});
            Add("4E02000005","FB","FB 600Ohm@100MHz 0.2A 900mOhm 1206 STD","BLM31AJ601SN1L","Murata","FB-1206","FerriteBeadSymbol","FB-1206",STD,SMD,new(){["Impedance(Ohm)@Frequency(Hz)"]="600Ohm@100MHz", ["Max Rated Current(A)"]="0.2A", ["Max DCR(Ohm)"]="900mOhm"});
            Add("4E02000005","FB","FB 600Ohm@100MHz 0.1A 350mOhm 1206(4-Array) STD","BLA31AG601SN4D","Murata","FB-1206(4-Array)","FerriteBead(4-Array)Symbol","FB-1206(4-Array)",STD,SMD,new(){["Impedance(Ohm)@Frequency(Hz)"]="600Ohm@100MHz", ["Max Rated Current(A)"]="0.1A", ["Max DCR(Ohm)"]="350mOhm"});
            Add("4E02000008","FB","FB 1KOhm@100MHz 0.6A 270mOhm 0805 AECQ","BLM21AG102SH1D","Murata","FB-0805","FerriteBeadSymbol","FB-0805",AECQ,SMD,new(){["Impedance(Ohm)@Frequency(Hz)"]="1KOhm@100MHz", ["Max Rated Current(A)"]="0.6A", ["Max DCR(Ohm)"]="270mOhm"});
            Add("4E02000012","FB","FB 1KOhm@100MHz 0.25A 400mOhm 0805 STD","BLM21BD102SN1D","Murata","FB-0805","FerriteBeadSymbol","FB-0805",STD,SMD,new(){["Impedance(Ohm)@Frequency(Hz)"]="1KOhm@100MHz", ["Max Rated Current(A)"]="0.25A", ["Max DCR(Ohm)"]="400mOhm"});
            Add("4E02000012","FB","FB 1KOhm@100MHz 0.25A 400mOhm 0805 AECQ","BLM21BD102SZ1D","Murata","FB-0805","FerriteBeadSymbol","FB-0805",AECQ,SMD,new(){["Impedance(Ohm)@Frequency(Hz)"]="1KOhm@100MHz", ["Max Rated Current(A)"]="0.25A", ["Max DCR(Ohm)"]="400mOhm"});
            Add("4E02000023","FB","FB 1KOhm@100MHz 0.45A 500mOhm 0603 AECQ","BLM18AG102BH1D","Murata","FB-0603","FerriteBeadSymbol","FB-0603",AECQ,SMD,new(){["Impedance(Ohm)@Frequency(Hz)"]="1KOhm@100MHz", ["Max Rated Current(A)"]="0.45A", ["Max DCR(Ohm)"]="500mOhm"});
            Add("4E02000033","FB","FB 120Ohm@100MHz 4A 9mOhm 1206 AECQ","BLM31KN121BH1L","Murata","FB-1206","FerriteBeadSymbol","FB-1206",AECQ,SMD,new(){["Impedance(Ohm)@Frequency(Hz)"]="120Ohm@100MHz", ["Max Rated Current(A)"]="4A", ["Max DCR(Ohm)"]="9mOhm"});
            Add("4E02000037","FB","FB 600Ohm@100MHz 2A 100mOhm 0805 STD","MFBM1V2012-601-R","Eaton","FB-0805","FerriteBeadSymbol","FB-0805",STD,SMD,new(){["Impedance(Ohm)@Frequency(Hz)"]="600Ohm@100MHz", ["Max Rated Current(A)"]="2A", ["Max DCR(Ohm)"]="100mOhm"});
            Add("4E02000020","FB","FB 120Ohm@100MHz 2.5A 35mOhm 0805 AECQ","782853121","Würth","FB-0805","FerriteBeadSymbol","FB-0805",AECQ,SMD,new(){["Impedance(Ohm)@Frequency(Hz)"]="120Ohm@100MHz", ["Max Rated Current(A)"]="2.5A", ["Max DCR(Ohm)"]="35mOhm"});
            Add("4E02000023","FB","FB 1KOhm@100MHz 0.2A 600mOhm 0603 AECQ","74279266","Würth","FB-0603","FerriteBeadSymbol","FB-0603",AECQ,SMD,new(){["Impedance(Ohm)@Frequency(Hz)"]="1KOhm@100MHz", ["Max Rated Current(A)"]="0.2A", ["Max DCR(Ohm)"]="600mOhm"});
            Add("4E02000047","FB","FB 1.5KOhm@100MHz 1A 300mOhm 0805 AECQ","782853162","Würth","FB-0805","FerriteBeadSymbol","FB-0805",AECQ,SMD,new(){["Impedance(Ohm)@Frequency(Hz)"]="1.5KOhm@100MHz", ["Max Rated Current(A)"]="1A", ["Max DCR(Ohm)"]="300mOhm"});
            //  CMC 
            Add("4E02000001","CMC","CMC 100uH@100KHz 5KOhm@10MHz 0.15A 4 PC Pad(ACT45B) AECQ","744235101","Würth","4 PC Pad(ACT45B)","CMC_Symbol","4 PC Pad(ACT45B)",AECQ,SMD,new(){["Inductance(H)@Frequency(Hz)"]="100uH@100KHz", ["Impedance(Ohm)@Frequency(Hz)"]="5KOhm@10MHz", ["Max Rated Current(A)"]="0.15A", ["Max DCR(Ohm)"]="2Ohm"});
            //  DIODE RECTIFIER 
            Add("4E42000001","DIO-REC","DIO-REC 600V 1.3V 1.5A DO-214AA(SMB) AECQ","RS2JHE3","Vishay","DO-214AA(SMB)","Diode_Symbol_CathodeLeft","DO-214AA(SMB)",AECQ,SMD,new(){["Max Peak Reverse Voltage(V)"]="600V", ["Max Forward Voltage(V)"]="1.3V", ["Max Forward Current(A)"]="1.5A"});
            Add("4E42000003","DIO-REC","DIO-REC 400V 1.1V 1A DO-214AC(SMA) STD","S1GE3","Vishay","DO-214AC(SMA)","Diode_Symbol_CathodeLeft","DO-214AC(SMA)",STD,SMD,new(){["Max Peak Reverse Voltage(V)"]="400V", ["Max Forward Voltage(V)"]="1.1V", ["Max Forward Current(A)"]="1A"});
            Add("4E42000003","DIO-REC","DIO-REC 400V 1.1V 1A DO-214AC(SMA) AECQ","S1GHE3","Vishay","DO-214AC(SMA)","Diode_Symbol_CathodeLeft","DO-214AC(SMA)",AECQ,SMD,new(){["Max Peak Reverse Voltage(V)"]="400V", ["Max Forward Voltage(V)"]="1.1V", ["Max Forward Current(A)"]="1A"});
            Add("4E42000004","DIO-REC","DIO-REC 600V 1.1V 1A DO-214AC(SMA) STD","S1J","Onsemi","DO-214AC(SMA)","Diode_Symbol_CathodeLeft","DO-214AC(SMA)",STD,SMD,new(){["Max Peak Reverse Voltage(V)"]="600V", ["Max Forward Voltage(V)"]="1.1V", ["Max Forward Current(A)"]="1A"});
            Add("4E42000004","DIO-REC","DIO-REC 600V 1.1V 1A DO-214AC(SMA) AECQ","S1JH","Taiwan Semiconductor","DO-214AC(SMA)","Diode_Symbol_CathodeLeft","DO-214AC(SMA)",AECQ,SMD,new(){["Max Peak Reverse Voltage(V)"]="600V", ["Max Forward Voltage(V)"]="1.1V", ["Max Forward Current(A)"]="1A"});
            Add("4E42000004","DIO-REC","DIO-REC 600V 1.1V 1A DO-214AC(SMA) AECQ","S1JHE3","Vishay","DO-214AC(SMA)","Diode_Symbol_CathodeLeft","DO-214AC(SMA)",AECQ,SMD,new(){["Max Peak Reverse Voltage(V)"]="600V", ["Max Forward Voltage(V)"]="1.1V", ["Max Forward Current(A)"]="1A"});
            Add("4E42000004","DIO-REC","DIO-REC 600V 1V 1A DO-214AC(SMA) AECQ","US1JH","Taiwan Semiconductor","DO-214AC(SMA)","Diode_Symbol_CathodeLeft","DO-214AC(SMA)",AECQ,SMD,new(){["Max Peak Reverse Voltage(V)"]="600V", ["Max Forward Voltage(V)"]="1V", ["Max Forward Current(A)"]="1A"});
            Add("4E42000005","DIO-REC","DIO-REC 600V 1.1V 1A SOD-323EP STD","S1JHE","Onsemi","SOD-323EP","Diode_Symbol_CathodeLeft","SOD-323EP",STD,SMD,new(){["Max Peak Reverse Voltage(V)"]="600V", ["Max Forward Voltage(V)"]="1.1V", ["Max Forward Current(A)"]="1A"});
            Add("4E42000007","DIO-REC","DIO-REC 1000V 1.1V 1A DO-214AC(SMA) STD","S1M-E3","Vishay","DO-214AC(SMA)","Diode_Symbol_CathodeLeft","DO-214AC(SMA)",STD,SMD,new(){["Max Peak Reverse Voltage(V)"]="1000V", ["Max Forward Voltage(V)"]="1.1V", ["Max Forward Current(A)"]="1A"});
            Add("4E42000007","DIO-REC","DIO-REC 1000V 1.1V 1A DO-214AC(SMA) AECQ","S1MHE3","Vishay","DO-214AC(SMA)","Diode_Symbol_CathodeLeft","DO-214AC(SMA)",AECQ,SMD,new(){["Max Peak Reverse Voltage(V)"]="1000V", ["Max Forward Voltage(V)"]="1.1V", ["Max Forward Current(A)"]="1A"});
            Add("4E42000012","DIO-REC","DIO-REC 600V 1.15V 3A DO-214AA(SMB) AECQ","NRVS3JB","Onsemi","DO-214AA(SMB)","Diode_Symbol","DO-214AA(SMB)",AECQ,SMD,new(){["Max Peak Reverse Voltage(V)"]="600V", ["Max Forward Voltage(V)"]="1.15V", ["Max Forward Current(A)"]="3A"});
            Add("4E42000012","DIO-REC","DIO-REC 600V 1.15V 3A DO-214AA(SMB) STD","S3JB","Onsemi","DO-214AA(SMB)","Diode_Symbol_CathodeLeft","DO-214AA(SMB)",STD,SMD,new(){["Max Peak Reverse Voltage(V)"]="600V", ["Max Forward Voltage(V)"]="1.15V", ["Max Forward Current(A)"]="3A"});
            Add("4E42000014","DIO-REC","DIO-REC 600V 1.1V 1A DO-214AC(SMA) AECQ","S1JHE3","Vishay","DO-214AC(SMA)","Diode_Symbol_CathodeLeft","DO-214AC(SMA)",AECQ,SMD,new(){["Max Peak Reverse Voltage(V)"]="600V", ["Max Forward Voltage(V)"]="1.1V", ["Max Forward Current(A)"]="1A"});
            Add("4E42000015","DIO-REC","DIO-REC 800V 1.1V 1A DO-214AC(SMA) AECQ","S1KHE3_A/I","Vishay","DO-214AC(SMA)","Diode_Symbol_CathodeLeft","DO-214AC(SMA)",AECQ,SMD,new(){["Max Peak Reverse Voltage(V)"]="800V", ["Max Forward Voltage(V)"]="1.1V", ["Max Forward Current(A)"]="1A"});
            Add("4E42000094","DIO-REC","DIO-REC 600V 1.1V 1A DO-214AC(SMA) AECQ","S1JHE3_A/I","Vishay","DO-214AC(SMA)","Diode_Symbol_CathodeLeft","DO-214AC(SMA)",AECQ,SMD,new(){["Max Peak Reverse Voltage(V)"]="600V", ["Max Forward Voltage(V)"]="1.1V", ["Max Forward Current(A)"]="1A"});
            Add("4E42000027","DIO-REC","DIO-REC 600V 1.1V 1A DO-214AC(SMA) AECQ","S1JHE3","Vishay","DO-214AC(SMA)","Diode_Symbol_CathodeLeft","DO-214AC(SMA)",AECQ,SMD,new(){["Max Peak Reverse Voltage(V)"]="600V", ["Max Forward Voltage(V)"]="1.1V", ["Max Forward Current(A)"]="1A"});
            Add("4E42000027","DIO-REC","DIO-REC 600V 1.1V 1A DO-214AC(SMA) AECQ","S1JHE3_A/H","Vishay","DO-214AC(SMA)","Diode_Symbol_CathodeLeft","DO-214AC(SMA)",AECQ,SMD,new(){["Max Peak Reverse Voltage(V)"]="600V", ["Max Forward Voltage(V)"]="1.1V", ["Max Forward Current(A)"]="1A"});
            Add("4E42000027","DIO-REC","DIO-REC 600V 1.1V 1A DO-214AC(SMA) AECQ","S1JHE3_A/I","Vishay","DO-214AC(SMA)","Diode_Symbol_CathodeLeft","DO-214AC(SMA)",AECQ,SMD,new(){["Max Peak Reverse Voltage(V)"]="600V", ["Max Forward Voltage(V)"]="1.1V", ["Max Forward Current(A)"]="1A"});
            Add("4E42000024","DIO-REC","DIO-REC 200V 1.15V 0.5A DO-219AB(SMF) AECQ","RS07D-HM3-08","Vishay","DO-219AB(SMF)","Diode_Symbol","DO-219AB(SMF)",AECQ,SMD,new(){["Max Peak Reverse Voltage(V)"]="200V", ["Max Forward Voltage(V)"]="1.15V", ["Max Forward Current(A)"]="0.5A"});
            Add("4E42000024","DIO-REC","DIO-REC 200V 1.15V 0.5A DO-219AB(SMF) AECQ","RS07D-HM3","Vishay","DO-219AB(SMF)","Diode_Symbol","DO-219AB(SMF)",AECQ,SMD,new(){["Max Peak Reverse Voltage(V)"]="200V", ["Max Forward Voltage(V)"]="1.15V", ["Max Forward Current(A)"]="0.5A"});
            Add("4E42000078","DIO-REC","DIO-REC 200V 0.9V 3A DO-214AB(SMC) STD","ES3D-E3/57T","Vishay","DO-214AB(SMC)","","DO-214AB(SMC)",STD,SMD,new(){["Max Peak Reverse Voltage(V)"]="200V", ["Max Forward Voltage(V)"]="0.9V", ["Max Forward Current(A)"]="3A"});
            Add("4E42000078","DIO-REC","DIO-REC 200V 0.9V 3A DO-214AB(SMC) STD","ES3D-E3/9AT","Vishay","DO-214AB(SMC)","","DO-214AB(SMC)",STD,SMD,new(){["Max Peak Reverse Voltage(V)"]="200V", ["Max Forward Voltage(V)"]="0.9V", ["Max Forward Current(A)"]="3A"});
            Add("4E42000066","DIO-REC","DIO-REC 400V 0.98V 1A SOD-123FL AECQ","RR268MM-600TFTR","Rohm","SOD-123FL","","SOD-123FL",AECQ,SMD,new(){["Max Peak Reverse Voltage(V)"]="400V", ["Max Forward Voltage(V)"]="0.98V", ["Max Forward Current(A)"]="1A"});
            Add("4E42000065","DIO-REC","DIO-REC 400V 1.1V 0.7A SOD-123FL AECQ","RR264MM-400TFTR","Rohm","SOD-123FL","","SOD-123FL",AECQ,SMD,new(){["Max Peak Reverse Voltage(V)"]="400V", ["Max Forward Voltage(V)"]="1.1V", ["Max Forward Current(A)"]="0.7A"});
            //  DIODE SWITCHING 
            Add("4E42000009","DIO-SW","DIO-SW 250V 1V 0.2A SOT-23 AECQ","BAS21-AU","Panjit International Inc.","SOT-23","Diode_Single_Symbol_3pin","SOT-23",AECQ,SMD,new(){["Max Peak Reverse Voltage(V)"]="250V", ["Max Forward Voltage(V)"]="1V", ["Max Forward Current(A)"]="0.2A"});
            Add("4E42000009","DIO-SW","DIO-SW 200V 1.25V 0.4A SOT-23 AECQ","BAS21Q","Diodes","SOT-23","Diode_Single_Symbol_3pin","SOT-23",AECQ,SMD,new(){["Max Peak Reverse Voltage(V)"]="200V", ["Max Forward Voltage(V)"]="1.25V", ["Max Forward Current(A)"]="0.4A"});
            Add("4E42000009","DIO-SW","DIO-SW 200V 1.25V 0.2A SOT-23 STD","BAS21,215","Nexperia","SOT-23","","SOT-23",STD,SMD,new(){["Max Peak Reverse Voltage(V)"]="200V", ["Max Forward Voltage(V)"]="1.25V", ["Max Forward Current(A)"]="0.2A"});
            Add("4E42000009","DIO-SW","DIO-SW 250V 1.25V 0.2A SOT-23 AECQ","SBAS21LT","Onsemi","SOT-23","Diode_Single_Symbol_3pin","SOT-23",AECQ,SMD,new(){["Max Peak Reverse Voltage(V)"]="250V", ["Max Forward Voltage(V)"]="1.25V", ["Max Forward Current(A)"]="0.2A"});
            Add("4E42000017","DIO-SW","DIO-SW 75V 1.25V 0.125A SOT-23 STD","BAV170,215","Nexperia","SOT-23","Diode_Common_Cathode_Symbol_3pin","SOT-23",STD,SMD,new(){["Max Peak Reverse Voltage(V)"]="75V", ["Max Forward Voltage(V)"]="1.25V", ["Max Forward Current(A)"]="0.125A"});
            Add("4E42000017","DIO-SW","DIO-SW 85V 1.25V 0.15A SOT-23 AECQ","BAV170Q-7-F","Diodes","SOT-23","Diode_Common_Cathode_Symbol_3pin","SOT-23",AECQ,SMD,new(){["Max Peak Reverse Voltage(V)"]="85V", ["Max Forward Voltage(V)"]="1.25V", ["Max Forward Current(A)"]="0.15A"});
            Add("4E42000017","DIO-SW","DIO-SW 75V 1.25V 0.32A DFN1010D-3 (SOT1215) AECQ","BAV170QAZ","NXP","DFN1010D-3 (SOT1215)","Diode_Common_Cathode_Symbol_3pin","DFN1010D-3 (SOT1215)",AECQ,SMD,new(){["Max Peak Reverse Voltage(V)"]="75V", ["Max Forward Voltage(V)"]="1.25V", ["Max Forward Current(A)"]="0.32A"});
            Add("4E42000017","DIO-SW","DIO-SW 70V 1.1V 0.215A SOT-23 STD","PMBD6100,215","NXP","SOT-23","Diode_Common_Cathode_Symbol_3pin","SOT-23",STD,SMD,new(){["Max Peak Reverse Voltage(V)"]="70V", ["Max Forward Voltage(V)"]="1.1V", ["Max Forward Current(A)"]="0.215A"});
            Add("4E42000017","DIO-SW","DIO-SW 75V 1.25V 0.215A SOT-23 AECQ","BAV170-QR","Nexperia","SOT-23","Diode_Common_Cathode_Symbol_3pin","SOT-23",AECQ,SMD,new(){["Max Peak Reverse Voltage(V)"]="75V", ["Max Forward Voltage(V)"]="1.25V", ["Max Forward Current(A)"]="0.215A"});
            Add("4E42000017","DIO-SW","DIO-SW 100V 1.25V 0.2A SOT-23 AECQ","SBAS16LT1G","Onsemi","SOT-23","Diode_Single_Symbol_3pin","SOT-23",AECQ,SMD,new(){["Max Peak Reverse Voltage(V)"]="100V", ["Max Forward Voltage(V)"]="1.25V", ["Max Forward Current(A)"]="0.2A"});
            Add("4E42000017","DIO-SW","DIO-SW 75V 1.25V 0.32A DFN1010D-3 (SOT1215) AECQ","BAV170QAZ","Nexperia","DFN1010D-3 (SOT1215)","Diode_Common_Cathode_Symbol_3pin","DFN1010D-3 (SOT1215)",AECQ,SMD,new(){["Max Peak Reverse Voltage(V)"]="75V", ["Max Forward Voltage(V)"]="1.25V", ["Max Forward Current(A)"]="0.32A"});
            Add("4E42000017","DIO-SW","DIO-SW 75V 1.25V 0.32A DFN1010D-3 (SOT1215) AECQ","BAV170QA","Nexperia","DFN1010D-3 (SOT1215)","Diode_Common_Cathode_Symbol_3pin","DFN1010D-3 (SOT1215)",AECQ,SMD,new(){["Max Peak Reverse Voltage(V)"]="75V", ["Max Forward Voltage(V)"]="1.25V", ["Max Forward Current(A)"]="0.32A"});
            Add("4E42000021","DIO-SW","DIO-SW 100V 1.25V 0.215A SOT-23 STD","BAS16","Nexperia","SOT-23","Diode_Single_Symbol_3pin","SOT-23",STD,SMD,new(){["Max Peak Reverse Voltage(V)"]="100V", ["Max Forward Voltage(V)"]="1.25V", ["Max Forward Current(A)"]="0.215A"});
            Add("4E42000021","DIO-SW","DIO-SW 100V 1.25V 0.215A SOT-23 STD","BAS16, 215","Nexperia","SOT-23","Diode_Single_Symbol_3pin","SOT-23",STD,SMD,new(){["Max Peak Reverse Voltage(V)"]="100V", ["Max Forward Voltage(V)"]="1.25V", ["Max Forward Current(A)"]="0.215A"});
            Add("4E42000021","DIO-SW","DIO-SW 100V 1.25V 0.215A SOT-23 AECQ","BAS16-QR","Nexperia","SOT-23","Diode_Single_Symbol_3pin","SOT-23",AECQ,SMD,new(){["Max Peak Reverse Voltage(V)"]="100V", ["Max Forward Voltage(V)"]="1.25V", ["Max Forward Current(A)"]="0.215A"});
            Add("4E42000054","DIO-SW","DIO-SW 100V 1.25V 0.215A SOD-123F STD","BAS16H,115","Nexperia","SOD-123F","","SOD-123F",STD,SMD,new(){["Max Peak Reverse Voltage(V)"]="100V", ["Max Forward Voltage(V)"]="1.25V", ["Max Forward Current(A)"]="0.215A"});
            //  DIODE ZENER 
            Add("4E42000022","DIO-ZEN","DIO-ZEN 6.8V 2% SOT-23 STD","BZX84-B6V8","Nexperia","SOT-23","Diode_Zener_Single_Symbol_3pin","SOT-23",STD,SMD,new(){["Nom Zener Voltage(V)"]="6.8V", ["Tolerance(%)"]="2%"});
            Add("4E42000022","DIO-ZEN","DIO-ZEN 6.8V 2% SOT-23 AECQ","BZX84-B6V8-Q","Nexperia","SOT-23","","SOT-23",AECQ,SMD,new(){["Nom Zener Voltage(V)"]="6.8V", ["Tolerance(%)"]="2%"});
            Add("4E42000022","DIO-ZEN","DIO-ZEN 6.8V 2% SOT-23 AECQ","BZX84-B6V8-QR","Nexperia","SOT-23","Diode_Zener_Common_Anode_Symbol_3pin","SOT-23",AECQ,SMD,new(){["Nom Zener Voltage(V)"]="6.8V", ["Tolerance(%)"]="2%"});
            Add("4E42000033","DIO-ZEN","DIO-ZEN 16V 2% SOD523 AECQ","BZX585-B16-QX","Nexperia","SOD523","","SOD523",AECQ,SMD,new(){["Nom Zener Voltage(V)"]="16V", ["Tolerance(%)"]="2%"});
            Add("4E42000033","DIO-ZEN","DIO-ZEN 16V 2% SOD523 STD","BZX585-B16","Nexperia","SOD523","","SOD523",STD,SMD,new(){["Nom Zener Voltage(V)"]="16V", ["Tolerance(%)"]="2%"});
            Add("4E42000033","DIO-ZEN","DIO-ZEN 16V 2% SOD-123 STD","BZT52-B16T-HT","Nexperia","SOD-123","","SOD-123",STD,SMD,new(){["Nom Zener Voltage(V)"]="16V", ["Tolerance(%)"]="2%"});
            Add("4E42000033","DIO-ZEN","DIO-ZEN 16V 2% SOD523 STD","BZX585-B16,115","Nexperia","SOD523","","SOD523",STD,SMD,new(){["Nom Zener Voltage(V)"]="16V", ["Tolerance(%)"]="2%"});
            Add("4E42000042","DIO-ZEN","DIO-ZEN 6.2V 2% SOT-23 AECQ","BZX84B6V2HE3-TP","MCC","SOT-23","","SOT-23",AECQ,SMD,new(){["Nom Zener Voltage(V)"]="6.2V", ["Tolerance(%)"]="2%"});
            Add("4E42000042","DIO-ZEN","DIO-ZEN 6.2V 5% SOT-23 AECQ","BZX84C6V2-HE3-08","Vishay","SOT-23","","SOT-23",AECQ,SMD,new(){["Nom Zener Voltage(V)"]="6.2V", ["Tolerance(%)"]="5%"});
            Add("4E42000042","DIO-ZEN","DIO-ZEN 6.2V 2% SOT-23 STD","BZX84B6V2-E3-08","Vishay","SOT-23","","SOT-23",STD,SMD,new(){["Nom Zener Voltage(V)"]="6.2V", ["Tolerance(%)"]="2%"});
            Add("4E42000043","DIO-ZEN","DIO-ZEN 4.7V 5% SOT-23 AECQ","BZX84C4V7-HE3_A-08","Vishay","SOT-23","Diode_Zener_Single_Symbol_3pin","SOT-23",AECQ,SMD,new(){["Nom Zener Voltage(V)"]="4.7V", ["Tolerance(%)"]="5%"});
            Add("4E42000043","DIO-ZEN","DIO-ZEN 4.7V 5% SOT-23 STD","BZX84C4V7-E3-08","Vishay","SOT-23","Diode_Zener_Single_Symbol_3pin","SOT-23",STD,SMD,new(){["Nom Zener Voltage(V)"]="4.7V", ["Tolerance(%)"]="5%"});
            Add("4E42000043","DIO-ZEN","DIO-ZEN 4.7V 6.38% SOT-23 AECQ","BZX84C4V7Q-7-F","Diodes","SOT-23","Diode_Zener_Single_Symbol_3pin","SOT-23",AECQ,SMD,new(){["Nom Zener Voltage(V)"]="4.7V", ["Tolerance(%)"]="6.38%"});
            Add("4E42000043","DIO-ZEN","DIO-ZEN 4.7V 5% SOT-23 AECQ","BZX84C4V7-HE3-08","Vishay","SOT-23","","SOT-23",AECQ,SMD,new(){["Nom Zener Voltage(V)"]="4.7V", ["Tolerance(%)"]="5%"});
            Add("4E42000043","DIO-ZEN","DIO-ZEN 4.7V 5% SOT-23 STD","BZX84C4V7-E3-08","Vishay","SOT-23","","SOT-23",STD,SMD,new(){["Nom Zener Voltage(V)"]="4.7V", ["Tolerance(%)"]="5%"});
            Add("4E42000043","DIO-ZEN","DIO-ZEN 4.7V 5% SOT-23 AECQ","BZX84C4V7Q-7-F","Diodes","SOT-23","","SOT-23",AECQ,SMD,new(){["Nom Zener Voltage(V)"]="4.7V", ["Tolerance(%)"]="5%"});
            Add("4E42000051","DIO-ZEN","DIO-ZEN 12V 1% SOT-23 AECQ","BZX84-A12-QR","Nexperia","SOT-23","","SOT-23",AECQ,SMD,new(){["Nom Zener Voltage(V)"]="12V", ["Tolerance(%)"]="1%"});
            Add("4E42000051","DIO-ZEN","DIO-ZEN 12V 1% SOT-23 STD","BZX84-A12,215","Nexperia","SOT-23","","SOT-23",STD,SMD,new(){["Nom Zener Voltage(V)"]="12V", ["Tolerance(%)"]="1%"});
            Add("4E42000053","DIO-ZEN","DIO-ZEN 18V 5% SOD-123F STD","BZT52H-C18,115","Nexperia","SOD-123F","","SOD-123F",STD,SMD,new(){["Nom Zener Voltage(V)"]="18V", ["Tolerance(%)"]="5%"});
            Add("4E42000059","DIO-ZEN","DIO-ZEN 3.9V 5% SOT-23 STD","BZX84-C3V9,215","Nexperia","SOT-23","","SOT-23",STD,SMD,new(){["Nom Zener Voltage(V)"]="3.9V", ["Tolerance(%)"]="5%"});
            Add("4E42000059","DIO-ZEN","DIO-ZEN 3.9V 5% SOT-23 AECQ","BZX84-C3V9-QR","Nexperia","SOT-23","","SOT-23",AECQ,SMD,new(){["Nom Zener Voltage(V)"]="3.9V", ["Tolerance(%)"]="5%"});
            Add("4E42000059","DIO-ZEN","DIO-ZEN 3.9V 5% SOT-23 AECQ","BZX84-C3V9-Q","Nexperia","SOT-23","","SOT-23",AECQ,SMD,new(){["Nom Zener Voltage(V)"]="3.9V", ["Tolerance(%)"]="5%"});
            Add("4E42000069","DIO-ZEN","DIO-ZEN 12V 5% SOD-323 AECQ","BZT52C12SQ-7-F","Diodes","SOD-323","","SOD-323",AECQ,SMD,new(){["Nom Zener Voltage(V)"]="12V", ["Tolerance(%)"]="5%"});
            Add("4E42000070","DIO-ZEN","DIO-ZEN 5.1V 2% SOD-323 AECQ","PDZ5.1B-Q","Nexperia","SOD-323","","SOD-323",AECQ,SMD,new(){["Nom Zener Voltage(V)"]="5.1V", ["Tolerance(%)"]="2%"});
            Add("4E42000070","DIO-ZEN","DIO-ZEN 5.1V 2% SOD-323 STD","PZD5.1B","Nexperia","SOD-323","","SOD-323",STD,SMD,new(){["Nom Zener Voltage(V)"]="5.1V", ["Tolerance(%)"]="2%"});
            Add("4E42000070","DIO-ZEN","DIO-ZEN 5.1V 2% SOD-323 STD","PDZ5.1B,115","Nexperia","SOD-323","","SOD-323",STD,SMD,new(){["Nom Zener Voltage(V)"]="5.1V", ["Tolerance(%)"]="2%"});
            Add("4E42000088","DIO-ZEN","DIO-ZEN 3.6V 5% SOD-323F AECQ","BZX84J-C3V6,115","Nexperia","SOD-323F","Diode_Zener_Single_2pin","SOD-323F",AECQ,SMD,new(){["Nom Zener Voltage(V)"]="3.6V", ["Tolerance(%)"]="5%"});
            Add("4E42000089","DIO-ZEN","DIO-ZEN 15V 5% SOD-323F AECQ","BZX84J-C15,115","Nexperia","SOD-323F","Diode_Zener_Single_2pin","SOD-323F",AECQ,SMD,new(){["Nom Zener Voltage(V)"]="15V", ["Tolerance(%)"]="5%"});
            Add("4E42000095","DIO-ZEN","DIO-ZEN 5.6V 5% SOD-323F AECQ","BZX84J-C5V6,115","Nexperia","SOD-323F","Diode_Zener_Single_2pin","SOD-323F",AECQ,SMD,new(){["Nom Zener Voltage(V)"]="5.6V", ["Tolerance(%)"]="5%"});
            Add("4E42000116","DIO-ZEN","DIO-ZEN 4.7V 2% SOD-323F AECQ","BZX84J-B4V7,115","Nexperia","SOD-323F","Diode_Zener_Single_2pin","SOD-323F",AECQ,SMD,new(){["Nom Zener Voltage(V)"]="4.7V", ["Tolerance(%)"]="2%"});
            Add("4E42000112","DIO-ZEN","DIO-ZEN 4.3V 7% SOT-23 AECQ","BZX84-C4V3-QVL","Nexperia","SOT-23","Diode_Zener_Single_Symbol_3pin","SOT-23",AECQ,SMD,new(){["Nom Zener Voltage(V)"]="4.3V", ["Tolerance(%)"]="7%"});
            Add("4E42000146","DIO-ZEN","DIO-ZEN 4.3V 7% SOT-23 AECQ","SZBZX84C4V3LT1G","Onsemi","SOT-23","Diode_Zener_Single_Symbol_3pin","SOT-23",AECQ,SMD,new(){["Nom Zener Voltage(V)"]="4.3V", ["Tolerance(%)"]="7%"});
            Add("4E42000132","DIO-ZEN","DIO-ZEN 5.1V 2% SOD-323 AECQ","SZMM3Z5V1ST1G","Onsemi","SOD-323","Diode_Zener_Single_2pin","SOD-323",AECQ,SMD,new(){["Nom Zener Voltage(V)"]="5.1V", ["Tolerance(%)"]="2%"});
            //  DIODE SCHOTTKY 
            Add("4E42000013","DIO-SCH","DIO-SCH 60V 3A DO-214AC(SMA) AECQ","B360A","Diodes","DO-214AC(SMA)","Diode_Schottky_Symbol","DO-214AC(SMA)",AECQ,SMD,new(){["Max Reverse Voltage(V)"]="60V", ["Max Forward Voltage(V)"]="0.7V", ["Avg Forward Current(V)"]="3A"});
            Add("4E42000013","DIO-SCH","DIO-SCH 60V 4A SMA-FL AECQ","NRVBAF360T3G","Onsemi","SMA-FL","Diode_Schottky_Symbol","SMA-FL",AECQ,SMD,new(){["Max Reverse Voltage(V)"]="60V", ["Max Forward Voltage(V)"]="0.63V", ["Avg Forward Current(V)"]="4A"});
            Add("4E42000020","DIO-SCH","DIO-SCH 60V 1A SOD-123W AECQ","PMEG6010ER-Q","Nexperia","SOD-123W","Diode_Schottky_Symbol","SOD-123W",AECQ,SMD,new(){["Max Reverse Voltage(V)"]="60V", ["Max Forward Voltage(V)"]="0.53V", ["Avg Forward Current(V)"]="1A"});
            Add("4E42000020","DIO-SCH","DIO-SCH 60V 1A SOD-123W STD","PMEG6010ER","Nexperia","SOD-123W","Diode_Schottky_Symbol_CathodeLeft","SOD-123W",STD,SMD,new(){["Max Reverse Voltage(V)"]="60V", ["Max Forward Voltage(V)"]="0.53V", ["Avg Forward Current(V)"]="1A"});
            Add("4E42000020","DIO-SCH","DIO-SCH 60V 1A SOD-123W STD","PMEG6010ER,115","Nexperia","SOD-123W","Diode_Schottky_Symbol_CathodeLeft","SOD-123W",STD,SMD,new(){["Max Reverse Voltage(V)"]="60V", ["Max Forward Voltage(V)"]="0.53V", ["Avg Forward Current(V)"]="1A"});
            Add("4E42000020","DIO-SCH","DIO-SCH 60V 1A SOD-123W AECQ","PMEG6010ER-QX","Nexperia","SOD-123W","Diode_Schottky_Symbol_CathodeLeft","SOD-123W",AECQ,SMD,new(){["Max Reverse Voltage(V)"]="60V", ["Max Forward Voltage(V)"]="0.53V", ["Avg Forward Current(V)"]="1A"});
            Add("4E42000025","DIO-SCH","DIO-SCH 60V 1A SOD-123W STD","PMEG6010ER","Nexperia","SOD-123W","Diode_Schottky_Symbol_CathodeLeft","SOD-123W",STD,SMD,new(){["Max Reverse Voltage(V)"]="60V", ["Max Forward Voltage(V)"]="0.53V", ["Avg Forward Current(V)"]="1A"});
            Add("4E42000040","DIO-SCH","DIO-SCH 30V 1A SOD-323F STD","PMEG6010CEJ,115","Nexperia","SOD-323F","","SOD-323F",STD,SMD,new(){["Max Reverse Voltage(V)"]="30V", ["Max Forward Voltage(V)"]="0.52V", ["Avg Forward Current(V)"]="1A"});
            Add("4E42000050","DIO-SCH","DIO-SCH 60V 1A SOD-123W AECQ","PMEG6010ELR-Q","Nexperia","SOD-123W","Diode_Schottky_Symbol","SOD-123W",AECQ,SMD,new(){["Max Reverse Voltage(V)"]="60V", ["Max Forward Voltage(V)"]="0.66V", ["Avg Forward Current(V)"]="1A"});
            Add("4E42000050","DIO-SCH","DIO-SCH 60V 1A SOD-123W STD","PMEG6010ELR","Nexperia","SOD-123W","Diode_Schottky_Symbol","SOD-123W",STD,SMD,new(){["Max Reverse Voltage(V)"]="60V", ["Max Forward Voltage(V)"]="0.66V", ["Avg Forward Current(V)"]="1A"});
            Add("4E42000067","DIO-SCH","DIO-SCH 100V 1A SOD-123W STD","PMEG10010ELRX","Nexperia","SOD-123W","","SOD-123W",STD,SMD,new(){["Max Reverse Voltage(V)"]="100V", ["Max Forward Voltage(V)"]="0.77V", ["Avg Forward Current(V)"]="1A"});
            Add("4E42000067","DIO-SCH","DIO-SCH 100V 1A SOD-123W AECQ","PMEG10010ELR-Q","Nexperia","SOD-123W","","SOD-123W",AECQ,SMD,new(){["Max Reverse Voltage(V)"]="100V", ["Max Forward Voltage(V)"]="0.77V", ["Avg Forward Current(V)"]="1A"});
            Add("4E42000067","DIO-SCH","DIO-SCH 100V 1A SOD-123W AECQ","PMEG10010ELR-QX","Nexperia","SOD-123W","","SOD-123W",AECQ,SMD,new(){["Max Reverse Voltage(V)"]="100V", ["Max Forward Voltage(V)"]="0.77V", ["Avg Forward Current(V)"]="1A"});
            Add("4E42000068","DIO-SCH","DIO-SCH 40V 0.1A SOD-523 AECQ","RB511SM-40FHT2R","Rohm","SOD-523","","SOD-523",AECQ,SMD,new(){["Max Reverse Voltage(V)"]="40V", ["Max Forward Voltage(V)"]="0.77V", ["Avg Forward Current(V)"]="0.1A"});
            Add("4E42000068","DIO-SCH","DIO-SCH 40V 0.1A SOD-523 AECQ","RB510SM-40FHT2R","Rohm","SOD-523","","SOD-523",AECQ,SMD,new(){["Max Reverse Voltage(V)"]="40V", ["Max Forward Voltage(V)"]="0.85V", ["Avg Forward Current(V)"]="0.1A"});
            Add("4E42000141","DIO-SCH","DIO-SCH 60V 1A DO-219AB(SMF) AECQ","SS1FH6HM3/H","Diodes","DO-219AB(SMF)","","DO-219AB(SMF)",AECQ,SMD,new(){["Max Reverse Voltage(V)"]="60V", ["Max Forward Voltage(V)"]="0.61V", ["Avg Forward Current(V)"]="1A"});
            Add("4E42000143","DIO-SCH","DIO-SCH 100V 1A SOD-123W AECQ","SS1H10LWH","Taiwan Semiconductor","SOD-123W","","SOD-123W",AECQ,SMD,new(){["Max Reverse Voltage(V)"]="100V", ["Max Forward Voltage(V)"]="0.8V", ["Avg Forward Current(V)"]="1A"});
            //  DIODE TVS 
            Add("4E42000048","DIO-TVS","DIO-TVS 54V 400W 4.6A Uni-Polar SOD-123W AECQ","PTVS54VS1UR-Q","Nexperia","SOD-123W","Diode_TVS_Bidirectional_Symbol","SOD-123W",AECQ,SMD,new(){["Reverse Standoff Voltage(V)"]="54V", ["Peak Pulse Power(W)"]="400W", ["Max Peak Pulse Current(A)"]="4.6A", ["Polarization"]="Uni-Polar"});
            Add("4E42000002","DIO-TVS","DIO-TVS 39V 600W 11.3A Bi-Polar DO-214AA(SMB) AECQ","P6SMB39CA-Q","Bourns Inc.","DO-214AA(SMB)","Diode_TVS_Bidirectional_Symbol","DO-214AA(SMB)",AECQ,SMD,new(){["Reverse Standoff Voltage(V)"]="39V", ["Peak Pulse Power(W)"]="600W", ["Max Peak Pulse Current(A)"]="11.3A", ["Polarization"]="Bi-Polar"});
            Add("4E42000002","DIO-TVS","DIO-TVS 39V 600W 11.1A Bi-Polar DO-214AA(SMB) AECQ","SM6T39CAHE3","Vishay","DO-214AA(SMB)","Diode_TVS_Bidirectional_Symbol","DO-214AA(SMB)",AECQ,SMD,new(){["Reverse Standoff Voltage(V)"]="39V", ["Peak Pulse Power(W)"]="600W", ["Max Peak Pulse Current(A)"]="11.1A", ["Polarization"]="Bi-Polar"});
            Add("4E42000002","DIO-TVS","DIO-TVS 39V 600W 11.1A Bi-Polar DO-214AA(SMB) STD","SM6T39CA-E3","Vishay","DO-214AA(SMB)","Diode_TVS_Bidirectional_Symbol","DO-214AA(SMB)",STD,SMD,new(){["Reverse Standoff Voltage(V)"]="39V", ["Peak Pulse Power(W)"]="600W", ["Max Peak Pulse Current(A)"]="11.1A", ["Polarization"]="Bi-Polar"});
            Add("4E42000002","DIO-TVS","DIO-TVS 39V 600W 11.1A Bi-Polar DO-214AA(SMB) AECQ","SM6T39CAY","STMicroelectronics","DO-214AA(SMB)","Diode_TVS_Bidirectional_Symbol","DO-214AA(SMB)",AECQ,SMD,new(){["Reverse Standoff Voltage(V)"]="39V", ["Peak Pulse Power(W)"]="600W", ["Max Peak Pulse Current(A)"]="11.1A", ["Polarization"]="Bi-Polar"});
            Add("4E42000008","DIO-TVS","DIO-TVS 36V 600W 12A Bi-Polar DO-214AA(SMB) STD","SM6T36CA","STMicroelectronics","DO-214AA(SMB)","Diode_TVS_Bidirectional_Symbol","DO-214AA(SMB)",STD,SMD,new(){["Reverse Standoff Voltage(V)"]="36V", ["Peak Pulse Power(W)"]="600W", ["Max Peak Pulse Current(A)"]="12A", ["Polarization"]="Bi-Polar"});
            Add("4E42000008","DIO-TVS","DIO-TVS 36V 600W 10.3A Bi-Polar DO-214AA(SMB) AECQ","SMBJ36CAQ","Diodes","DO-214AA(SMB)","Diode_TVS_Bidirectional_Symbol","DO-214AA(SMB)",AECQ,SMD,new(){["Reverse Standoff Voltage(V)"]="36V", ["Peak Pulse Power(W)"]="600W", ["Max Peak Pulse Current(A)"]="10.3A", ["Polarization"]="Bi-Polar"});
            Add("4E42000010","DIO-TVS","DIO-TVS 30V 600W 14.5A Bi-Polar DO-214AA(SMB) AECQ","SM6T30CAH","Vishay","DO-214AA(SMB)","Diode_TVS_Bidirectional_Symbol","DO-214AA(SMB)",AECQ,SMD,new(){["Reverse Standoff Voltage(V)"]="30V", ["Peak Pulse Power(W)"]="600W", ["Max Peak Pulse Current(A)"]="14.5A", ["Polarization"]="Bi-Polar"});
            Add("4E42000010","DIO-TVS","DIO-TVS 30V 600W 14.5A Bi-Polar DO-214AA(SMB) AECQ","SM6T30CAY","STMicroelectronics","DO-214AA(SMB)","Diode_TVS_Bidirectional_Symbol","DO-214AA(SMB)",AECQ,SMD,new(){["Reverse Standoff Voltage(V)"]="30V", ["Peak Pulse Power(W)"]="600W", ["Max Peak Pulse Current(A)"]="14.5A", ["Polarization"]="Bi-Polar"});
            Add("4E42000010","DIO-TVS","DIO-TVS 30V 600W 14.5A Bi-Polar DO-214AA(SMB) AECQ","SM6T30CAHE3","Vishay","DO-214AA(SMB)","Diode_TVS_Bidirectional_Symbol","DO-214AA(SMB)",AECQ,SMD,new(){["Reverse Standoff Voltage(V)"]="30V", ["Peak Pulse Power(W)"]="600W", ["Max Peak Pulse Current(A)"]="14.5A", ["Polarization"]="Bi-Polar"});
            Add("4E42000011","DIO-TVS","DIO-TVS 24V 600W 18A Bi-Polar DO-214AA(SMB) AECQ","SM6T24CAHE3","Vishay","DO-214AA(SMB)","Diode_TVS_Bidirectional_Symbol","DO-214AA(SMB)",AECQ,SMD,new(){["Reverse Standoff Voltage(V)"]="24V", ["Peak Pulse Power(W)"]="600W", ["Max Peak Pulse Current(A)"]="18A", ["Polarization"]="Bi-Polar"});
            Add("4E42000011","DIO-TVS","DIO-TVS 24V 600W 18A Bi-Polar DO-214AA(SMB) AECQ","SM6T24CAY","STMicroelectronics","DO-214AA(SMB)","Diode_TVS_Bidirectional_Symbol","DO-214AA(SMB)",AECQ,SMD,new(){["Reverse Standoff Voltage(V)"]="24V", ["Peak Pulse Power(W)"]="600W", ["Max Peak Pulse Current(A)"]="18A", ["Polarization"]="Bi-Polar"});
            Add("4E42000011","DIO-TVS","DIO-TVS 24V 600W 18A Bi-Polar DO-214AA(SMB) STD","SM6T24CA-E3","Vishay","DO-214AA(SMB)","Diode_TVS_Bidirectional_Symbol","DO-214AA(SMB)",STD,SMD,new(){["Reverse Standoff Voltage(V)"]="24V", ["Peak Pulse Power(W)"]="600W", ["Max Peak Pulse Current(A)"]="18A", ["Polarization"]="Bi-Polar"});
            Add("4E42000016","DIO-TVS","DIO-TVS 30V 600W 14.5A Bi-Polar DO-214AA(SMB) AECQ","SM6T30CAHE3","Vishay","DO-214AA(SMB)","Diode_TVS_Bidirectional_Symbol","DO-214AA(SMB)",AECQ,SMD,new(){["Reverse Standoff Voltage(V)"]="30V", ["Peak Pulse Power(W)"]="600W", ["Max Peak Pulse Current(A)"]="14.5A", ["Polarization"]="Bi-Polar"});
            Add("4E42000019","DIO-TVS","DIO-TVS 33V 600W 10.6A Bi-Polar DO-214AA(SMB) STD","SMBJ33CA","Taiwan Semiconductor","DO-214AA(SMB)","Diode_TVS_Bidirectional_Symbol","DO-214AA(SMB)",STD,SMD,new(){["Reverse Standoff Voltage(V)"]="33V", ["Peak Pulse Power(W)"]="600W", ["Max Peak Pulse Current(A)"]="10.6A", ["Polarization"]="Bi-Polar"});
            Add("4E42000019","DIO-TVS","DIO-TVS 33V 600W 13.1A Bi-Polar DO-214AA(SMB) AECQ","SM6T33CAHE3","Vishay","DO-214AA(SMB)","Diode_TVS_Bidirectional_Symbol","DO-214AA(SMB)",AECQ,SMD,new(){["Reverse Standoff Voltage(V)"]="33V", ["Peak Pulse Power(W)"]="600W", ["Max Peak Pulse Current(A)"]="13.1A", ["Polarization"]="Bi-Polar"});
            Add("4E42000019","DIO-TVS","DIO-TVS 33V 600W 11.3A Bi-Polar DO-214AA(SMB) AECQ","SMBJ33CAHE3_A/I","Vishay","DO-214AA(SMB)","Diode_TVS_Bidirectional_Symbol","DO-214AA(SMB)",AECQ,SMD,new(){["Reverse Standoff Voltage(V)"]="33V", ["Peak Pulse Power(W)"]="600W", ["Max Peak Pulse Current(A)"]="11.3A", ["Polarization"]="Bi-Polar"});
            Add("4E42000019","DIO-TVS","DIO-TVS 33V 600W 11.3A Bi-Polar DO-214AA(SMB) AECQ","SMBJ33CAHE3_B/I","Vishay","DO-214AA(SMB)","Diode_TVS_Bidirectional_Symbol","DO-214AA(SMB)",AECQ,SMD,new(){["Reverse Standoff Voltage(V)"]="33V", ["Peak Pulse Power(W)"]="600W", ["Max Peak Pulse Current(A)"]="11.3A", ["Polarization"]="Bi-Polar"});
            Add("4E42000019","DIO-TVS","DIO-TVS 33V 600W 11.3A Bi-Polar DO-214AA(SMB) AECQ","SMBJ33CAHE3_B/H","Vishay","DO-214AA(SMB)","Diode_TVS_Bidirectional_Symbol","DO-214AA(SMB)",AECQ,SMD,new(){["Reverse Standoff Voltage(V)"]="33V", ["Peak Pulse Power(W)"]="600W", ["Max Peak Pulse Current(A)"]="11.3A", ["Polarization"]="Bi-Polar"});
            Add("4E42000026","DIO-TVS","DIO-TVS 36V 600W 12A Bi-Polar DO-214AA(SMB) STD","SM6T36CA","STMicroelectronics","DO-214AA(SMB)","Diode_TVS_Bidirectional_Symbol","DO-214AA(SMB)",STD,SMD,new(){["Reverse Standoff Voltage(V)"]="36V", ["Peak Pulse Power(W)"]="600W", ["Max Peak Pulse Current(A)"]="12A", ["Polarization"]="Bi-Polar"});
            Add("4E42000028","DIO-TVS","DIO-TVS 30V 600W 14.5A Bi-Polar DO-214AA(SMB) STD","SM6T30CA-E3/52","Vishay","DO-214AA(SMB)","Diode_TVS_Bidirectional_Symbol","DO-214AA(SMB)",STD,SMD,new(){["Reverse Standoff Voltage(V)"]="30V", ["Peak Pulse Power(W)"]="600W", ["Max Peak Pulse Current(A)"]="14.5A", ["Polarization"]="Bi-Polar"});
            Add("4E42000028","DIO-TVS","DIO-TVS 30V 600W 14.5A Bi-Polar DO-214AA(SMB) AECQ","SM6T30CAY","STMicroelectronics","DO-214AA(SMB)","Diode_TVS_Bidirectional_Symbol","DO-214AA(SMB)",AECQ,SMD,new(){["Reverse Standoff Voltage(V)"]="30V", ["Peak Pulse Power(W)"]="600W", ["Max Peak Pulse Current(A)"]="14.5A", ["Polarization"]="Bi-Polar"});
            Add("4E42000028","DIO-TVS","DIO-TVS 30V 600W 14.5A Bi-Polar DO-214AA(SMB) AECQ","SM6T30CAHE3","Vishay","DO-214AA(SMB)","Diode_TVS_Bidirectional_Symbol","DO-214AA(SMB)",AECQ,SMD,new(){["Reverse Standoff Voltage(V)"]="30V", ["Peak Pulse Power(W)"]="600W", ["Max Peak Pulse Current(A)"]="14.5A", ["Polarization"]="Bi-Polar"});
            Add("4E42000030","DIO-TVS","DIO-TVS 24V 600W 18A Uni-Polar DO-214AA(SMB) STD","SM6T24A-E3/52","Vishay","DO-214AA(SMB)","Diode_TVS_Unidirectional_Symbol_CathodeLeft","DO-214AA(SMB)",STD,SMD,new(){["Reverse Standoff Voltage(V)"]="24V", ["Peak Pulse Power(W)"]="600W", ["Max Peak Pulse Current(A)"]="18A", ["Polarization"]="Uni-Polar"});
            Add("4E42000030","DIO-TVS","DIO-TVS 24V 600W 18A Uni-Polar DO-214AA(SMB) AECQ","SM6T24A-HE3/52","Vishay","DO-214AA(SMB)","Diode_TVS_Unidirectional_Symbol_CathodeLeft","DO-214AA(SMB)",AECQ,SMD,new(){["Reverse Standoff Voltage(V)"]="24V", ["Peak Pulse Power(W)"]="600W", ["Max Peak Pulse Current(A)"]="18A", ["Polarization"]="Uni-Polar"});
            Add("4E42000030","DIO-TVS","DIO-TVS 24V 600W 18A Uni-Polar DO-214AA(SMB) AECQ","SM6T24AY","STMicroelectronics","DO-214AA(SMB)","Diode_TVS_Unidirectional_Symbol_CathodeLeft","DO-214AA(SMB)",AECQ,SMD,new(){["Reverse Standoff Voltage(V)"]="24V", ["Peak Pulse Power(W)"]="600W", ["Max Peak Pulse Current(A)"]="18A", ["Polarization"]="Uni-Polar"});
            Add("4E42000030","DIO-TVS","DIO-TVS 24V 600W 18A Uni-Polar DO-214AA(SMB) STD","SM6T24A-E3/52","Vishay","DO-214AA(SMB)","Diode_TVS_Unidirectional_Symbol_CathodeLeft","DO-214AA(SMB)",STD,SMD,new(){["Reverse Standoff Voltage(V)"]="24V", ["Peak Pulse Power(W)"]="600W", ["Max Peak Pulse Current(A)"]="18A", ["Polarization"]="Uni-Polar"});
            Add("4E42000031","DIO-TVS","DIO-TVS 24V 600W 18A Uni-Polar DO-214AA(SMB) STD","SM6T24A-E3/52","Vishay","DO-214AA(SMB)","Diode_TVS_Unidirectional_Symbol_CathodeLeft","DO-214AA(SMB)",STD,SMD,new(){["Reverse Standoff Voltage(V)"]="24V", ["Peak Pulse Power(W)"]="600W", ["Max Peak Pulse Current(A)"]="18A", ["Polarization"]="Uni-Polar"});
            Add("4E42000087","DIO-TVS","DIO-TVS 42V 600W 10.3A Bi-Polar DO-214AA(SMB) AECQ","SM6T42CAY","STMicroelectronics","DO-214AA(SMB)","Diode_TVS_Bidirectional_Symbol","DO-214AA(SMB)",AECQ,SMD,new(){["Reverse Standoff Voltage(V)"]="42V", ["Peak Pulse Power(W)"]="600W", ["Max Peak Pulse Current(A)"]="10.3A", ["Polarization"]="Bi-Polar"});
            Add("4E42000056","DIO-TVS","DIO-TVS 82V 400W 3.5A Bi-Polar DO-214AC(SMA) AECQ","SM4T82CAY","STMicroelectronics","DO-214AC(SMA)","","DO-214AC(SMA)",AECQ,SMD,new(){["Reverse Standoff Voltage(V)"]="82V", ["Peak Pulse Power(W)"]="400W", ["Max Peak Pulse Current(A)"]="3.5A", ["Polarization"]="Bi-Polar"});
            Add("4E42000055","DIO-TVS","DIO-TVS 33V 400W 8.8A Bi-Polar DO-214AC(SMA) AECQ","SM4T33CAY","STMicroelectronics","DO-214AC(SMA)","","DO-214AC(SMA)",AECQ,SMD,new(){["Reverse Standoff Voltage(V)"]="33V", ["Peak Pulse Power(W)"]="400W", ["Max Peak Pulse Current(A)"]="8.8A", ["Polarization"]="Bi-Polar"});
            Add("4E42000032","DIO-TVS","DIO-TVS 26V 600W 14.3A Bi-Polar DO-214AA(SMB) AECQ","P6SMBJ26CA-AQ","Diotec Semiconductor","DO-214AA(SMB)","","DO-214AA(SMB)",AECQ,SMD,new(){["Reverse Standoff Voltage(V)"]="26V", ["Peak Pulse Power(W)"]="600W", ["Max Peak Pulse Current(A)"]="14.3A", ["Polarization"]="Bi-Polar"});
            Add("4E42000032","DIO-TVS","DIO-TVS 30V 600W 14.5A Bi-Polar DO-214AA(SMB) AECQ","P6SMB30CAHE3_A/H","Vishay","DO-214AA(SMB)","","DO-214AA(SMB)",AECQ,SMD,new(){["Reverse Standoff Voltage(V)"]="30V", ["Peak Pulse Power(W)"]="600W", ["Max Peak Pulse Current(A)"]="14.5A", ["Polarization"]="Bi-Polar"});
            Add("4E42000032","DIO-TVS","DIO-TVS 30V 600W 14.7A Bi-Polar DO-214AA(SMB) AECQ","P6SMB30CA-Q","Bourns Inc.","DO-214AA(SMB)","","DO-214AA(SMB)",AECQ,SMD,new(){["Reverse Standoff Voltage(V)"]="30V", ["Peak Pulse Power(W)"]="600W", ["Max Peak Pulse Current(A)"]="14.7A", ["Polarization"]="Bi-Polar"});
            Add("4E42000032","DIO-TVS","DIO-TVS 26V 600W 14.3A Bi-Polar DO-214AA(SMB) STD","P6SMBJ26CA","Diotec Semiconductor","DO-214AA(SMB)","","DO-214AA(SMB)",STD,SMD,new(){["Reverse Standoff Voltage(V)"]="26V", ["Peak Pulse Power(W)"]="600W", ["Max Peak Pulse Current(A)"]="14.3A", ["Polarization"]="Bi-Polar"});
            Add("4E42000032","DIO-TVS","DIO-TVS 26V 600W 14.2A Bi-Polar DO-214AA(SMB) AECQ","SMBJ26CAHE3-TP","MCC","DO-214AA(SMB)","","DO-214AA(SMB)",AECQ,SMD,new(){["Reverse Standoff Voltage(V)"]="26V", ["Peak Pulse Power(W)"]="600W", ["Max Peak Pulse Current(A)"]="14.2A", ["Polarization"]="Bi-Polar"});
            Add("4E42000032","DIO-TVS","DIO-TVS 30V 600W 14.5A Bi-Polar DO-214AA(SMB) AECQ","SM6T30CA-HE3/52","Vishay","DO-214AA(SMB)","","DO-214AA(SMB)",AECQ,SMD,new(){["Reverse Standoff Voltage(V)"]="30V", ["Peak Pulse Power(W)"]="600W", ["Max Peak Pulse Current(A)"]="14.5A", ["Polarization"]="Bi-Polar"});
            Add("4E42000032","DIO-TVS","DIO-TVS 30V 600W 14.5A Bi-Polar DO-214AA(SMB) AECQ","P6SMB30CAHE3_B/H","Vishay","DO-214AA(SMB)","","DO-214AA(SMB)",AECQ,SMD,new(){["Reverse Standoff Voltage(V)"]="30V", ["Peak Pulse Power(W)"]="600W", ["Max Peak Pulse Current(A)"]="14.5A", ["Polarization"]="Bi-Polar"});
            Add("4E42000049","DIO-TVS","DIO-TVS 33V 1500W 29A Bi-Polar DO-214AB(SMC) STD","SMCJ33CA-TR","STMicroelectronics","DO-214AB(SMC)","","DO-214AB(SMC)",STD,SMD,new(){["Reverse Standoff Voltage(V)"]="33V", ["Peak Pulse Power(W)"]="1500W", ["Max Peak Pulse Current(A)"]="29A", ["Polarization"]="Bi-Polar"});
            Add("4E42000049","DIO-TVS","DIO-TVS 33V 1500W 28.1A Bi-Polar DO-214AB(SMC) STD","SMCJ33CA-E3/57T","Vishay","DO-214AB(SMC)","","DO-214AB(SMC)",STD,SMD,new(){["Reverse Standoff Voltage(V)"]="33V", ["Peak Pulse Power(W)"]="1500W", ["Max Peak Pulse Current(A)"]="28.1A", ["Polarization"]="Bi-Polar"});
            Add("4E42000049","DIO-TVS","DIO-TVS 33V 1500W 33A Bi-Polar DO-214AB(SMC) AECQ","SM15T33CAY","STMicroelectronics","DO-214AB(SMC)","","DO-214AB(SMC)",AECQ,SMD,new(){["Reverse Standoff Voltage(V)"]="33V", ["Peak Pulse Power(W)"]="1500W", ["Max Peak Pulse Current(A)"]="33A", ["Polarization"]="Bi-Polar"});
            Add("4E42000049","DIO-TVS","DIO-TVS 33V 1500W 28.1A Bi-Polar DO-214AB(SMC) AECQ","SMCJ33CAHE3/57T","Vishay","DO-214AB(SMC)","","DO-214AB(SMC)",AECQ,SMD,new(){["Reverse Standoff Voltage(V)"]="33V", ["Peak Pulse Power(W)"]="1500W", ["Max Peak Pulse Current(A)"]="28.1A", ["Polarization"]="Bi-Polar"});
            Add("4E42000071","DIO-TVS","DIO-TVS 33V 400W 7.5A Bi-Polar DO-214AC(SMA) AECQ","SMAJ33CAQ-13-F","Diodes","DO-214AC(SMA)","","DO-214AC(SMA)",AECQ,SMD,new(){["Reverse Standoff Voltage(V)"]="33V", ["Peak Pulse Power(W)"]="400W", ["Max Peak Pulse Current(A)"]="7.5A", ["Polarization"]="Bi-Polar"});
            Add("4E42000076","DIO-TVS","DIO-TVS 70V 400W 3.5A Bi-Polar DO-214AC(SMA) AECQ","SMAJ70CAH","Taiwan Semiconductor","DO-214AC(SMA)","","DO-214AC(SMA)",AECQ,SMD,new(){["Reverse Standoff Voltage(V)"]="70V", ["Peak Pulse Power(W)"]="400W", ["Max Peak Pulse Current(A)"]="3.5A", ["Polarization"]="Bi-Polar"});
            Add("4E42000076","DIO-TVS","DIO-TVS 70V 400W 3.5A Bi-Polar DO-214AC(SMA) STD","SMAJ70CA","Littelfuse","DO-214AC(SMA)","","DO-214AC(SMA)",STD,SMD,new(){["Reverse Standoff Voltage(V)"]="70V", ["Peak Pulse Power(W)"]="400W", ["Max Peak Pulse Current(A)"]="3.5A", ["Polarization"]="Bi-Polar"});
            Add("4E42000144","DIO-TVS","DIO-TVS 54V 400W 4.6A Uni-Polar SOD-123FL AECQ","SM4F54AHE3-TP","MCC","SOD-123FL","","SOD-123FL",AECQ,SMD,new(){["Reverse Standoff Voltage(V)"]="54V", ["Peak Pulse Power(W)"]="400W", ["Max Peak Pulse Current(A)"]="4.6A", ["Polarization"]="Uni-Polar"});
            //  NTC 
            Add("","NTC","NTC 10KOhm 5% 3940K 0805 AECQ","NTCS0805E3103JHT","Vishay","NTC-0805","","NTC-0805",AECQ,SMD,new(){["Resistance@25°C(Ohm)"]="10KOhm", ["Resistance Tolerance(%)"]="5%", ["Beta(K)"]="3940K", ["Beta Tolerance(%)"]="1%"});
            //  PTC 
            //  LED 
            Add("4E52000150","LED","LED KW3.CGLNM3.TK T6T9-6L07L0-JC AECQ","KW3.CGLNM3.TK","Osram","","LED_Symbol","",AECQ,SMD,new(){["Nom Forward Voltage(V)"]="9.47V", ["Nom Forward Current(A)"]="1A", ["BIN"]="T6T9-6L07L0-JC"});
            Add("4E52000100","LED","LED KS.DMLN31.23 GZHY-J3T3-68 AECQ","KS.DMLN31.23","Osram","","LED_Symbol","",AECQ,SMD,new(){["Nom Forward Voltage(V)"]="2.3V", ["Nom Forward Current(A)"]="0.2A", ["BIN"]="GZHY-J3T3-68"});
            Add("4E52000101","LED","LED KY.DMLN31.23 HYJX-46-J3T3 AECQ","KY.DMLN31.23","Osram","","LED_Symbol","",AECQ,SMD,new(){["Nom Forward Voltage(V)"]="2.35V", ["Nom Forward Current(A)"]="0.2A", ["BIN"]="HYJX-46-J3T3"});
            Add("4E52000127","LED","LED KW.DMLN33.SG 7J7K-8E8G-ebvF46fcbB46 AECQ","KW.DMLN33.SG","Osram","","","",AECQ,SMD,new(){["Nom Forward Voltage(V)"]="3V", ["Nom Forward Current(A)"]="0.2A", ["BIN"]="7J7K-8E8G-ebvF46fcbB46"});
            Add("4E52000115","LED","LED LS.E67F BACA-68-3A4A AECQ","LS.E67F","Osram","LS E67F_Fprt","LED_Symbol","LS E67F_Fprt",AECQ,SMD,new(){["Nom Forward Voltage(V)"]="2.15V", ["Nom Forward Current(A)"]="0.05A", ["BIN"]="BACA-68-3A4A"});
            Add("4E52000074","LED","LED NFSA172CT P13d22P14-L3- AECQ","NFSA172CT","Nichia","","","",AECQ,SMD,new(){["Nom Forward Voltage(V)"]="3V", ["Nom Forward Current(A)"]="0.15A", ["BIN"]="P13d22P14-L3-"});
            Add("4E52000007","LED","LED NFSA172CT P13d22P14-L3- AECQ","NFSA172CT","Nichia","","","",AECQ,SMD,new(){["Nom Forward Voltage(V)"]="3V", ["Nom Forward Current(A)"]="0.15A", ["BIN"]="P13d22P14-L3-"});
            Add("4E52000020","LED","LED LUW.CEUN.CE 5L7L-HNJN-8E8F AECQ","LUW.CEUN.CE","Osram","","","",AECQ,SMD,new(){["Nom Forward Voltage(V)"]="2.97V", ["Nom Forward Current(A)"]="0.35A", ["BIN"]="5L7L-HNJN-8E8F"});
            Add("4E52000009","LED","LED KY.DMLS31.FY 6L7L-5F-8E8G AECQ","KY.DMLS31.FY","Osram","","","",AECQ,SMD,new(){["Nom Forward Voltage(V)"]="2.95V", ["Nom Forward Current(A)"]="0.6A", ["BIN"]="6L7L-5F-8E8G"});
            Add("4E52000010","LED","LED LS.E67F BABB-1-3A4B AECQ","LS.E67F","Osram","","","",AECQ,SMD,new(){["Nom Forward Voltage(V)"]="2.15V", ["Nom Forward Current(A)"]="0.05A", ["BIN"]="BABB-1-3A4B"});
            Add("4E12345678","LED","LED KY DDLN31.FY LED 5H6H7H8H AECQ","KY DDLN31.FY LED","Osram","KY DDLN31.FY LED","LED_Symbol","KY DDLN31.FY LED",AECQ,SMD,new(){["Nom Forward Voltage(V)"]="3.25V", ["Nom Forward Current(A)"]="0.15A", ["BIN"]="5H6H7H8H"});
            Add("4E52000015","LED","LED LA.A67F  AA-AB AECQ","LA.A67F","Osram","LA.A67F","LED_Symbol","LA.A67F",AECQ,SMD,new(){["Nom Forward Voltage(V)"]="2V", ["Nom Forward Current(A)"]="0.03A", ["BIN"]="AA-AB"});
            Add("4E52000107","LED","LED LA G6SP.02 8E-5F AECQ","LA G6SP.02","Osram","LA G6SP.02 LED","LED_Symbol","LA G6SP.02 LED",AECQ,SMD,new(){["Nom Forward Voltage(V)"]="2.25V", ["Nom Forward Current(A)"]="0.14A", ["BIN"]="8E-5F"});
            Add("4E52000023","LED","LED LY G6SP CA-DB AECQ","LY G6SP","Osram","LY.G6SP LED","LED_Symbol","LY.G6SP LED",AECQ,SMD,new(){["Nom Forward Voltage(V)"]="2.25V", ["Nom Forward Current(A)"]="0.14A", ["BIN"]="CA-DB"});
            Add("4E52000024","LED","LED LWG6SP EA-FA AECQ","LWG6SP","Osram","LW.G6SP LED","LED_Symbol","LW.G6SP LED",AECQ,SMD,new(){["Nom Forward Voltage(V)"]="3.30V", ["Nom Forward Current(A)"]="0.14A", ["BIN"]="EA-FA"});
            Add("","LED","LED LY G6SP.02 8D5F-56 AECQ","LY G6SP.02","Osram","LY.G6SP LED","LED_Symbol","LY.G6SP LED",AECQ,SMD,new(){["Nom Forward Voltage(V)"]="2.25V", ["Nom Forward Current(A)"]="0.14A", ["BIN"]="8D5F-56"});
            Add("","LED","LED LA G6SP.02 8E-6F AECQ","LA G6SP.02","Osram","LA G6SP.02 LED","LED_Symbol","LA G6SP.02 LED",AECQ,SMD,new(){["Nom Forward Voltage(V)"]="2.25V", ["Nom Forward Current(A)"]="0.14A", ["BIN"]="8E-6F"});
            Add("","LED","LED KW DMLQ33.SG 6L8LF-ebxD46ebzB46-8E8G AECQ","KW DMLQ33.SG","Osram","","LED_Symbol","",AECQ,SMD,new(){["Nom Forward Voltage(V)"]="2.95V", ["Nom Forward Current(A)"]="0.35A", ["BIN"]="6L8LF-ebxD46ebzB46-8E8G"});
            Add("","LED","LED KR DMLQ32.23 6K8K-46-J3T3 AECQ","KR DMLQ32.23","Osram","","LED_Symbol","",AECQ,SMD,new(){["Nom Forward Voltage(V)"]="2.25V", ["Nom Forward Current(A)"]="0.35A", ["BIN"]="6K8K-46-J3T3"});
            //  TRANSISTORS 
            Add("4E72000015","TRAN","TRAN NPN 65V 0.1A 0.25W SOT-23 AECQ","BC846-Q","Nexperia","SOT-23","NPN_Transistor_Symbol","SOT-23",AECQ,SMD,new(){["Type"]="NPN", ["Max Collector-Emitter Voltage(V)"]="65V", ["Max Collector Current(I)"]="0.1A", ["Max Base-Emitter Voltage(V)"]="6V", ["Power(W)"]="0.25W"});
            Add("4E72000025","TRAN","TRAN NPN 50V 0.1A 0.25W SOT-23 AECQ","PDTC114YT-Q","Nexperia","SOT-23","NPN_Transistor_Symbol","SOT-23",AECQ,SMD,new(){["Type"]="NPN", ["Max Collector-Emitter Voltage(V)"]="50V", ["Max Collector Current(I)"]="0.1A", ["Max Base-Emitter Voltage(V)"]="6V", ["Power(W)"]="0.25W"});
            Add("4E72000016","TRAN","TRAN PNP 40V 0.2A 0.225W SOT-23 AECQ","SMMBT3906LT1G","Onsemi","SOT-23","PNP_Transistor_Symbol","SOT-23",AECQ,SMD,new(){["Type"]="PNP", ["Max Collector-Emitter Voltage(V)"]="40V", ["Max Collector Current(I)"]="0.2A", ["Max Base-Emitter Voltage(V)"]="5V", ["Power(W)"]="0.225W"});
            Add("4E72000007","TRAN","TRAN NPN 45V 0.5A 0.31W SOT-23 AECQ","BC817-40Q-13-F","Diodes","SOT-23","NPN_Transistor_Symbol","SOT-23",AECQ,SMD,new(){["Type"]="NPN", ["Max Collector-Emitter Voltage(V)"]="45V", ["Max Collector Current(I)"]="0.5A", ["Max Base-Emitter Voltage(V)"]="5V", ["Power(W)"]="0.31W"});
            Add("4E72000014","TRAN","TRAN NPN 80V 1A 1.5W SOT223 STD","BCP56T1G","Onsemi","SOT223","NPN_Transistor_Symbol","SOT223",STD,SMD,new(){["Type"]="NPN", ["Max Collector-Emitter Voltage(V)"]="80V", ["Max Collector Current(I)"]="1A", ["Max Base-Emitter Voltage(V)"]="5V", ["Power(W)"]="1.5W"});
            Add("4E72000014","TRAN","TRAN NPN 80V 1A 1.5W SOT223 AECQ","SBCP56T1G","Onsemi","SOT223","NPN_Transistor_Symbol","SOT223",AECQ,SMD,new(){["Type"]="NPN", ["Max Collector-Emitter Voltage(V)"]="80V", ["Max Collector Current(I)"]="1A", ["Max Base-Emitter Voltage(V)"]="5V", ["Power(W)"]="1.5W"});
            Add("4E72000010","TRAN","TRAN NPN 80V 1A 0.6W SOT223 AECQ","BCP56-16T-QF","Nexperia","SOT223","NPN_Transistor_Symbol","SOT223",AECQ,SMD,new(){["Type"]="NPN", ["Max Collector-Emitter Voltage(V)"]="80V", ["Max Collector Current(I)"]="1A", ["Max Base-Emitter Voltage(V)"]="5V", ["Power(W)"]="0.6W"});
            Add("4E72000012","TRAN","TRAN NPN 45V 1A 0.5W SOT89 AECQ","BCX54-16-QX","Nexperia","SOT89","NPN_Transistor_Symbol","SOT89",AECQ,SMD,new(){["Type"]="NPN", ["Max Collector-Emitter Voltage(V)"]="45V", ["Max Collector Current(I)"]="1A", ["Max Base-Emitter Voltage(V)"]="5V", ["Power(W)"]="0.5W"});
            Add("4E72000011","TRAN","TRAN NPN 65V 0.1A 0.25W SOT-23 AECQ","BC846-QR","Nexperia","SOT-23","NPN_Transistor_Symbol","SOT-23",AECQ,SMD,new(){["Type"]="NPN", ["Max Collector-Emitter Voltage(V)"]="65V", ["Max Collector Current(I)"]="0.1A", ["Max Base-Emitter Voltage(V)"]="6V", ["Power(W)"]="0.25W"});
            Add("4E72000011","TRAN","TRAN NPN 65V 0.1A 0.25W SOT-23 STD","BC846,215","Nexperia","SOT-23","NPN_Transistor_Symbol","SOT-23",STD,SMD,new(){["Type"]="NPN", ["Max Collector-Emitter Voltage(V)"]="65V", ["Max Collector Current(I)"]="0.1A", ["Max Base-Emitter Voltage(V)"]="6V", ["Power(W)"]="0.25W"});
            Add("4E72000022","TRAN","TRAN NPN 45V 1A 0.5W SOT89 AECQ","BCX54-QX","Nexperia","SOT89","NPN_Transistor_Symbol","SOT89",AECQ,SMD,new(){["Type"]="NPN", ["Max Collector-Emitter Voltage(V)"]="45V", ["Max Collector Current(I)"]="1A", ["Max Base-Emitter Voltage(V)"]="5V", ["Power(W)"]="0.5W"});
            Add("4E72000004","TRAN","TRAN PNP 50V 0.1A 0.187W SOT−363 AECQ","SMUN5111DW1T1G","Onsemi","SOT−363","PNP_Transistor_Symbol","SOT−363",AECQ,SMD,new(){["Type"]="PNP", ["Max Collector-Emitter Voltage(V)"]="50V", ["Max Collector Current(I)"]="0.1A", ["Max Base-Emitter Voltage(V)"]="6V", ["Power(W)"]="0.187W"});
            Add("4E72000033","TRAN","TRAN NPN 80V 0.5A 0.46W SC-70 AECQ","SMMBTA06WT3G","Onsemi","SC-70","NPN_Transistor_Symbol","SC-70",AECQ,SMD,new(){["Type"]="NPN", ["Max Collector-Emitter Voltage(V)"]="80V", ["Max Collector Current(I)"]="0.5A", ["Max Base-Emitter Voltage(V)"]="4V", ["Power(W)"]="0.46W"});
            Add("4E72000005","TRAN","TRAN NPN 50V 0.1A 0.2W SOT−363 AECQ","DDC114EUQ-7-F","Diodes","SOT−363","NPN_Transistor_Symbol","SOT−363",AECQ,SMD,new(){["Type"]="NPN", ["Max Collector-Emitter Voltage(V)"]="50V", ["Max Collector Current(I)"]="0.1A", ["Max Base-Emitter Voltage(V)"]="5V", ["Power(W)"]="0.2W"});
            Add("4E72000006","TRAN","TRAN PNP 50V 0.1A 0.202W SOT-23 AECQ","SMMUN2111LT1G","Onsemi","SOT-23","PNP_Transistor_Symbol","SOT-23",AECQ,SMD,new(){["Type"]="PNP", ["Max Collector-Emitter Voltage(V)"]="50V", ["Max Collector Current(I)"]="0.1A", ["Max Base-Emitter Voltage(V)"]="6V", ["Power(W)"]="0.202W"});
            Add("4E72000006","TRAN","TRAN PNP 50V 0.1A 0.23W SOT-23 AECQ","SMUN2111T1G","Onsemi","SOT-23","PNP_Transistor_Symbol","SOT-23",AECQ,SMD,new(){["Type"]="PNP", ["Max Collector-Emitter Voltage(V)"]="50V", ["Max Collector Current(I)"]="0.1A", ["Max Base-Emitter Voltage(V)"]="2.5V", ["Power(W)"]="0.23W"});
            Add("4E72000029","TRAN","TRAN PNP 60V 1A 0.65W SOT223 AECQ","BCP52-16,115","Nexperia","SOT223","PNP_Transistor_Symbol","SOT223",AECQ,SMD,new(){["Type"]="PNP", ["Max Collector-Emitter Voltage(V)"]="60V", ["Max Collector Current(I)"]="1A", ["Max Base-Emitter Voltage(V)"]="5V", ["Power(W)"]="0.65W"});
            Add("4E72000023","TRAN","TRAN NPN&PNP 50V 0.1A 0.25W SOT−363 AECQ","BCR35PNH6327XTSA1","Infineon","SOT−363","","SOT−363",AECQ,SMD,new(){["Type"]="NPN&PNP", ["Max Collector-Emitter Voltage(V)"]="50V", ["Max Collector Current(I)"]="0.1A", ["Max Base-Emitter Voltage(V)"]="6V", ["Power(W)"]="0.25W"});
            Add("4E72000013","TRAN","TRAN NPN 45V 0.5A 0.46W SC-70 STD","BC817-40W","Onsemi","SC-70","NPN_Transistor_Symbol","SC-70",STD,SMD,new(){["Type"]="NPN", ["Max Collector-Emitter Voltage(V)"]="45V", ["Max Collector Current(I)"]="0.5A", ["Max Base-Emitter Voltage(V)"]="5V", ["Power(W)"]="0.46W"});
            Add("4E72000013","TRAN","TRAN NPN 45V 0.5A 0.2W SC-70 STD","BC817-40W,135","Nexperia","SC-70","NPN_Transistor_Symbol","SC-70",STD,SMD,new(){["Type"]="NPN", ["Max Collector-Emitter Voltage(V)"]="45V", ["Max Collector Current(I)"]="0.5A", ["Max Base-Emitter Voltage(V)"]="5V", ["Power(W)"]="0.2W"});
            Add("4E72000030","TRAN","TRAN PNP 45V 0.5A 0.25W SOT-23 STD","BC807,215","Nexperia","SOT-23","PNP_Transistor_Symbol","SOT-23",STD,SMD,new(){["Type"]="PNP", ["Max Collector-Emitter Voltage(V)"]="45V", ["Max Collector Current(I)"]="0.5A", ["Max Base-Emitter Voltage(V)"]="5V", ["Power(W)"]="0.25W"});
            Add("4E72000030","TRAN","TRAN PNP 45V 0.5A 0.25W SOT-23 AECQ","BC807-QR","Nexperia","SOT-23","PNP_Transistor_Symbol","SOT-23",AECQ,SMD,new(){["Type"]="PNP", ["Max Collector-Emitter Voltage(V)"]="45V", ["Max Collector Current(I)"]="0.5A", ["Max Base-Emitter Voltage(V)"]="5V", ["Power(W)"]="0.25W"});
            Add("4E42000052","TRAN","TRAN NPN&PNP 50V 0.07A 0.2W SOT−363 AECQ","DCX114YUQ-7-F","Diodes","SOT−363","","SOT−363",AECQ,SMD,new(){["Type"]="NPN&PNP", ["Max Collector-Emitter Voltage(V)"]="50V", ["Max Collector Current(I)"]="0.07A", ["Max Base-Emitter Voltage(V)"]="5V", ["Power(W)"]="0.2W"});
            Add("4E72000002","TRAN","TRAN NPN 80V 1A 0.5W SOT89 STD","BCX56-16,135","Nexperia","SOT89","NPN_Transistor_Symbol","SOT89",STD,SMD,new(){["Type"]="NPN", ["Max Collector-Emitter Voltage(V)"]="80V", ["Max Collector Current(I)"]="1A", ["Max Base-Emitter Voltage(V)"]="5V", ["Power(W)"]="0.5W"});
            Add("4E72000002","TRAN","TRAN NPN 80V 1A 0.5W SOT89 STD","BCX56-16,115","Nexperia","SOT89","NPN_Transistor_Symbol","SOT89",STD,SMD,new(){["Type"]="NPN", ["Max Collector-Emitter Voltage(V)"]="80V", ["Max Collector Current(I)"]="1A", ["Max Base-Emitter Voltage(V)"]="5V", ["Power(W)"]="0.5W"});
            Add("4E72000002","TRAN","TRAN NPN 80V 1A 0.5W SOT89 STD","BCX56-16","Nexperia","SOT89","NPN_Transistor_Symbol","SOT89",STD,SMD,new(){["Type"]="NPN", ["Max Collector-Emitter Voltage(V)"]="80V", ["Max Collector Current(I)"]="1A", ["Max Base-Emitter Voltage(V)"]="5V", ["Power(W)"]="0.5W"});
            Add("4E72000001","TRAN","TRAN NPN 45V 0.5A 0.25W SOT-23 STD","BC817,215","Nexperia","SOT-23","NPN_Transistor_Symbol","SOT-23",STD,SMD,new(){["Type"]="NPN", ["Max Collector-Emitter Voltage(V)"]="45V", ["Max Collector Current(I)"]="0.5A", ["Max Base-Emitter Voltage(V)"]="5V", ["Power(W)"]="0.25W"});
            Add("4E72000001","TRAN","TRAN NPN 45V 0.5A 0.25W SOT-23 AECQ","BC817-QR","Nexperia","SOT-23","NPN_Transistor_Symbol","SOT-23",AECQ,SMD,new(){["Type"]="NPN", ["Max Collector-Emitter Voltage(V)"]="45V", ["Max Collector Current(I)"]="0.5A", ["Max Base-Emitter Voltage(V)"]="5V", ["Power(W)"]="0.25W"});
            Add("4E72000001","TRAN","TRAN NPN 45V 0.5A 0.25W SOT-23 STD","BC817-16","Nexperia","SOT-23","NPN_Transistor_Symbol","SOT-23",STD,SMD,new(){["Type"]="NPN", ["Max Collector-Emitter Voltage(V)"]="45V", ["Max Collector Current(I)"]="0.5A", ["Max Base-Emitter Voltage(V)"]="5V", ["Power(W)"]="0.25W"});
            Add("4E72000001","TRAN","TRAN NPN 45V 0.5A 0.25W SOT-23 STD","BC817-16,215","Nexperia","SOT-23","NPN_Transistor_Symbol","SOT-23",STD,SMD,new(){["Type"]="NPN", ["Max Collector-Emitter Voltage(V)"]="45V", ["Max Collector Current(I)"]="0.5A", ["Max Base-Emitter Voltage(V)"]="5V", ["Power(W)"]="0.25W"});
            Add("4E72000001","TRAN","TRAN NPN 45V 0.5A 0.25W SOT-23 STD","BC817-16","Nexperia","SOT-23","NPN_Transistor_Symbol","SOT-23",STD,SMD,new(){["Type"]="NPN", ["Max Collector-Emitter Voltage(V)"]="45V", ["Max Collector Current(I)"]="0.5A", ["Max Base-Emitter Voltage(V)"]="5V", ["Power(W)"]="0.25W"});
            Add("4E72000039","TRAN","TRAN NPN&PNP 60V 0.6A 0.25W SOT−363 AECQ","PMBT2907AYSX","Nexperia","SOT−363","","SOT−363",AECQ,SMD,new(){["Type"]="NPN&PNP", ["Max Collector-Emitter Voltage(V)"]="60V", ["Max Collector Current(I)"]="0.6A", ["Max Base-Emitter Voltage(V)"]="5V", ["Power(W)"]="0.25W"});
            Add("4E72000042","TRAN","TRAN NPN&PNP 45V 0.5A 0.37W SOT−363 STD","BC817DS,115","Nexperia","SOT−363","","SOT−363",STD,SMD,new(){["Type"]="NPN&PNP", ["Max Collector-Emitter Voltage(V)"]="45V", ["Max Collector Current(I)"]="0.5A", ["Max Base-Emitter Voltage(V)"]="5V", ["Power(W)"]="0.37W"});
            Add("4E72000069","TRAN","TRAN NPN 45V 0.5A 0.31W SOT-23 AECQ","BC817-25Q","Diodes","SOT-23","NPN_Transistor_Symbol","SOT-23",AECQ,SMD,new(){["Type"]="NPN", ["Max Collector-Emitter Voltage(V)"]="45V", ["Max Collector Current(I)"]="0.5A", ["Max Base-Emitter Voltage(V)"]="5V", ["Power(W)"]="0.31W"});
            Add("4E72000068","TRAN","TRAN PNP 60V 0.6A 0.31W SOT-23 AECQ","MMBT2907AQ","Diodes","SOT-23","PNP_Transistor_Symbol","SOT-23",AECQ,SMD,new(){["Type"]="PNP", ["Max Collector-Emitter Voltage(V)"]="60V", ["Max Collector Current(I)"]="0.6A", ["Max Base-Emitter Voltage(V)"]="6V", ["Power(W)"]="0.31W"});
            Add("4E72000065","TRAN","TRAN NPN 45V 0.5A 0.31W SOT-23 AECQ","BC817-16Q-7-F","Diodes","SOT-23","NPN_Transistor_Symbol","SOT-23",AECQ,SMD,new(){["Type"]="NPN", ["Max Collector-Emitter Voltage(V)"]="45V", ["Max Collector Current(I)"]="0.5A", ["Max Base-Emitter Voltage(V)"]="5V", ["Power(W)"]="0.31W"});
            Add("4E72000053","TRAN","TRAN NPN 65V 0.1A 0.225W SOT-23 AECQ","BC846AQ-7-F","Diodes","SOT-23","NPN_Transistor_Symbol","SOT-23",AECQ,SMD,new(){["Type"]="NPN", ["Max Collector-Emitter Voltage(V)"]="65V", ["Max Collector Current(I)"]="0.1A", ["Max Base-Emitter Voltage(V)"]="6V", ["Power(W)"]="0.225W"});
            Add("4E72000061","TRAN","TRAN PNP 80V 1A 2W SOT223 AECQ","BCP5316QTA","Diodes","SOT223","","SOT223",AECQ,SMD,new(){["Type"]="PNP", ["Max Collector-Emitter Voltage(V)"]="80V", ["Max Collector Current(I)"]="1A", ["Max Base-Emitter Voltage(V)"]="5V", ["Power(W)"]="2W"});
            Add("4E72000063","TRAN","TRAN PNP 45V 0.5A 0.31W SOT23-3 AECQ","BC807-40Q-7-F","Diodes","SOT23-3","","SOT23-3",AECQ,SMD,new(){["Type"]="PNP", ["Max Collector-Emitter Voltage(V)"]="45V", ["Max Collector Current(I)"]="0.5A", ["Max Base-Emitter Voltage(V)"]="5V", ["Power(W)"]="0.31W"});
            Add("4E72000073","TRAN","TRAN NPN 65V 0.1A 0.25W SOT-23 AECQ","BC846A-AQ","Diotec","SOT-23","NPN_Transistor_Symbol","SOT-23",AECQ,SMD,new(){["Type"]="NPN", ["Max Collector-Emitter Voltage(V)"]="65V", ["Max Collector Current(I)"]="0.1A", ["Max Base-Emitter Voltage(V)"]="6V", ["Power(W)"]="0.25W"});
            //  MOSFETS 
            Add("4E72000032","MOS","MOS P-Ch 40V 13A 3.2W WDFN-8 AECQ","NVTFWS9D6P04M8LTAG","Onsemi","WDFN-8","","WDFN-8",AECQ,SMD,new(){["Type"]="P-Ch", ["Max Drain-Source Voltage(V)"]="40V", ["Max Drain Current(I)"]="13A", ["Max Gate-Source Voltage(V)"]="20V", ["Power(W)"]="3.2W"});
            Add("4E72000034","MOS","MOS N-Ch 40V 100A 3.5W PowerDI3333-8 AECQ","DMTH43M7LFGQ","Diodes","PowerDI3333-8","","PowerDI3333-8",AECQ,SMD,new(){["Type"]="N-Ch", ["Max Drain-Source Voltage(V)"]="40V", ["Max Drain Current(I)"]="100A", ["Max Gate-Source Voltage(V)"]="20V", ["Power(W)"]="3.5W"});
            Add("4E72000034","MOS","MOS N-Ch 40V 100A 2.62W PowerDI3333-8 AECQ","DMTH43M8LFGQ","Diodes","PowerDI3333-8","","PowerDI3333-8",AECQ,SMD,new(){["Type"]="N-Ch", ["Max Drain-Source Voltage(V)"]="40V", ["Max Drain Current(I)"]="100A", ["Max Gate-Source Voltage(V)"]="20V", ["Power(W)"]="2.62W"});
            Add("4E72000031","MOS","MOS N-Ch 40V 9.4A 2.9W WDFN-8 AECQ","NVTFS015N04CTAG","Onsemi","WDFN-8","","WDFN-8",AECQ,SMD,new(){["Type"]="N-Ch", ["Max Drain-Source Voltage(V)"]="40V", ["Max Drain Current(I)"]="9.4A", ["Max Gate-Source Voltage(V)"]="20V", ["Power(W)"]="2.9W"});
            Add("4E72000038","MOS","MOS N-Ch 100V 4A 0.8W U-DFN2020-6 (Type F) AECQ","DMT10H072LFDFQ-7","Diodes","U-DFN2020-6 (Type F)","","U-DFN2020-6 (Type F)",AECQ,SMD,new(){["Type"]="N-Ch", ["Max Drain-Source Voltage(V)"]="100V", ["Max Drain Current(I)"]="4A", ["Max Gate-Source Voltage(V)"]="20V", ["Power(W)"]="0.8W"});
            Add("4E72000019","MOS","MOS N-Ch 55V 0.54A 0.36W SOT-23 AECQ","BSS670S2LH6327XTSA1","Infineon","SOT-23","","SOT-23",AECQ,SMD,new(){["Type"]="N-Ch", ["Max Drain-Source Voltage(V)"]="55V", ["Max Drain Current(I)"]="0.54A", ["Max Gate-Source Voltage(V)"]="20V", ["Power(W)"]="0.36W"});
            Add("4E72000009","MOS","MOS N-Ch 60V 3A 1.2W SOT-23 AECQ","YJL03N06AQ","Yangjie","SOT-23","","SOT-23",AECQ,SMD,new(){["Type"]="N-Ch", ["Max Drain-Source Voltage(V)"]="60V", ["Max Drain Current(I)"]="3A", ["Max Gate-Source Voltage(V)"]="20V", ["Power(W)"]="1.2W"});
            Add("4E42000018","MOS","MOS N-Ch 60V 2A 0.8W SOT-23 STD","DMN6075S","Diodes","SOT-23","","SOT-23",STD,SMD,new(){["Type"]="N-Ch", ["Max Drain-Source Voltage(V)"]="60V", ["Max Drain Current(I)"]="2A", ["Max Gate-Source Voltage(V)"]="20V", ["Power(W)"]="0.8W"});
            Add("4E42000018","MOS","MOS N-Ch 60V 2A 0.8W SOT-23 STD","DMN6075S-7","Diodes","SOT-23","","SOT-23",STD,SMD,new(){["Type"]="N-Ch", ["Max Drain-Source Voltage(V)"]="60V", ["Max Drain Current(I)"]="2A", ["Max Gate-Source Voltage(V)"]="20V", ["Power(W)"]="0.8W"});
            Add("4E42000018","MOS","MOS N-Ch 60V 2A 0.8W SOT-23 STD","DMN6075S-07","Diodes","SOT-23","","SOT-23",STD,SMD,new(){["Type"]="N-Ch", ["Max Drain-Source Voltage(V)"]="60V", ["Max Drain Current(I)"]="2A", ["Max Gate-Source Voltage(V)"]="20V", ["Power(W)"]="0.8W"});
            Add("4E42000018","MOS","MOS N-Ch 60V 2A 0.8W SOT-23 AECQ","DMN6075SQ-13","Diodes","SOT-23","","SOT-23",AECQ,SMD,new(){["Type"]="N-Ch", ["Max Drain-Source Voltage(V)"]="60V", ["Max Drain Current(I)"]="2A", ["Max Gate-Source Voltage(V)"]="20V", ["Power(W)"]="0.8W"});
            Add("4E42000018","MOS","MOS N-Ch 60V 2A 0.8W SOT-23 AECQ","DMN6075SQ-7","Diodes","SOT-23","","SOT-23",AECQ,SMD,new(){["Type"]="N-Ch", ["Max Drain-Source Voltage(V)"]="60V", ["Max Drain Current(I)"]="2A", ["Max Gate-Source Voltage(V)"]="20V", ["Power(W)"]="0.8W"});
            //  DC-DC CONVERTERS 
            Add("","DC-DC","DC-DC Buck 40V 2A SO-8EP AECQ","AP64200Q","Diodes","SO-8EP","AP64200Q_Symbol","SO-8EP",AECQ,SMD,new(){["Type"]="Buck", ["Max Output Voltage(V)"]="40V", ["Max Output Current(I)"]="2A"});
            Add("4E02000027","DC-DC","DC-DC Buck-Boost 25V - TQFN-EP(24-SW) AECQ","MAX20048ATGB/VY+","Analog Devices","TQFN-EP(24-SW)","","TQFN-EP(24-SW)",AECQ,SMD,new(){["Type"]="Buck-Boost", ["Max Output Voltage(V)"]="25V", ["Max Output Current(I)"]="-"});
            //  LDO 
            Add("4E32000013","LDO","LDO Adjustable 37V 1.5A D2PAK-3_Ftprnt AECQ","NCV317BD2TR4G","Onsemi","D2PAK-3_Ftprnt","NCV317BD2TR4G_Symboll","D2PAK-3_Ftprnt",AECQ,SMD,new(){["Output Type"]="Adjustable", ["Max Output Voltage(V)"]="37V", ["Max Output Current(I)"]="1.5A"});
            //  LED DRIVERS 
            Add("4E32000018","LED-DRV","LED-DRV Linear 1-Ch 180mA PG-TSDSO-14 AECQ","TLD2142-1EP","Infineon","PG-TSDSO-14","TLD2142-1EP_Symbol","PG-TSDSO-14",AECQ,SMD,new(){["Type"]="Linear", ["Total Channel Count"]="1-Ch", ["Output Current/Channel(A)"]="180mA"});
            Add("4E30000001","LED-DRV","LED-DRV Linear 3-Ch 150mA SOIC16N-EP AECQ","E522.80","Elmos","SOIC16N-EP","E522.80_Symbol","SOIC16N-EP",AECQ,SMD,new(){["Type"]="Linear", ["Total Channel Count"]="3-Ch", ["Output Current/Channel(A)"]="150mA"});
            Add("4E32000034","LED-DRV","LED-DRV Linear 3-Ch 150mA TFDSO-16 AECQ","TLD2382-3ET","Infineon","TFDSO-16","TLD2382-3ET_Symbol","TFDSO-16",AECQ,SMD,new(){["Type"]="Linear", ["Total Channel Count"]="3-Ch", ["Output Current/Channel(A)"]="150mA"});
            Add("","LED-DRV","LED-DRV Linear 3-Ch 150mA ETSSOP-16 AECQ","IS32LT3152B","Lumissil","ETSSOP-16","IS32LT3152_Symbol","ETSSOP-16",AECQ,SMD,new(){["Type"]="Linear", ["Total Channel Count"]="3-Ch", ["Output Current/Channel(A)"]="150mA"});
            Add("","LED-DRV","LED-DRV Linear 1-Ch 100mA SOT-26 AECQ","BCR430UW6-7","Diodes","SOT-26","BCR430UW6_Symbol","SOT-26",AECQ,SMD,new(){["Type"]="Linear", ["Total Channel Count"]="1-Ch", ["Output Current/Channel(A)"]="100mA"});
            Add("","LED-DRV","LED-DRV Linear 3-Ch 150mA eTSSOP-16_Fprt AECQ","IS32LT3144A","Lumissil","eTSSOP-16_Fprt","IS32LT3144A_Symbol","eTSSOP-16_Fprt",AECQ,SMD,new(){["Type"]="Linear", ["Total Channel Count"]="3-Ch", ["Output Current/Channel(A)"]="150mA"});
            //  MICROCONTROLLERS 
            Add("","MCU","MCU ARM-M3 32-Bit 72MHz 64KB Flash LQFP-48 STD","STM32F103C8T6","STM","LQFP-48","STM32F103C8T6_Symbol","LQFP-48",STD,SMD,new(){["Core Processor"]="ARM-M3", ["Core Size(Bit)"]="32-Bit", ["Clock Speed(Hz)"]="72MHz", ["Program Memory Size(B)"]="64KB", ["Program Memory Type"]="Flash", ["ASIL Category"]="NA"});
            Add("","MCU","MCU PIC 16-Bit 32Mhz 3.5KB Flash SOIC-8 STD","PIC16F15213T-E/SNVAO","Microchip","SOIC-8","PIC16F15213-E/SN_Symbol","SOIC-8",STD,SMD,new(){["Core Processor"]="PIC", ["Core Size(Bit)"]="16-Bit", ["Clock Speed(Hz)"]="32Mhz", ["Program Memory Size(B)"]="3.5KB", ["Program Memory Type"]="Flash", ["ASIL Category"]="NA"});
            //  OTHER IC 
            Add("","OTHER-IC","MOTOR DRIVER DRV8825PWP HTSSOP-28 STD","DRV8825PWP","Texas Instruments","HTSSOP-28","DRV8825PWP_Symbol","HTSSOP-28",STD,SMD,null);
            Add("","OTHER-IC","Digital Temperature Sensor TMP112AIDRLR SOT-563 STD","TMP112AIDRLR","Texas Instruments","SOT-563","TMP112_Symbol","SOT-563",STD,SMD,null);
            Add("4E02000014","OTHER-IC","Quad 2-input multiplexer 74HC157PW-Q100 SOT403-1 AECQ","74HC157PW-Q100","Nexperia","SOT403-1","","SOT403-1",AECQ,SMD,null);
            Add("4E02000049","OTHER-IC","Quad 2-input multiplexer SN74HCS157PW-Q1 TSSOP-16 AECQ","SN74HCS157PW-Q1","Texas Instruments","TSSOP-16","","TSSOP-16",AECQ,SMD,null);
            //  CONNECTORS 
            Add("4E80000022","CON","CON Male 2-Pin 2mm Natural AECQ","502352-0200","Molex","","","",AECQ,SMD,new(){["Gender"]="Male", ["Pin Count"]="2", ["Pitch(mm)"]="2mm", ["Color"]="Natural"});
            Add("4E80000095","CON","CON Male 8-Pin 2mm Black AECQ","502352-0801","Molex","","","",AECQ,SMD,new(){["Gender"]="Male", ["Pin Count"]="8", ["Pitch(mm)"]="2mm", ["Color"]="Black"});
            Add("4E80000096","CON","CON Male 2-Pin 2mm Black AECQ","502352-0201","Molex","","","",AECQ,SMD,new(){["Gender"]="Male", ["Pin Count"]="2", ["Pitch(mm)"]="2mm", ["Color"]="Black"});
            Add("4E80000097","CON","CON Male 8-Pin 2mm Black AECQ","560020-0821","Molex","","","",AECQ,SMD,new(){["Gender"]="Male", ["Pin Count"]="8", ["Pitch(mm)"]="2mm", ["Color"]="Black"});
            Add("4E80000027","CON","CON Male 4-Pin 2mm Natural AECQ","502352-0400","Molex","5023520400-Footprint-1","Molex-4Pin","5023520400-Footprint-1",AECQ,SMD,new(){["Gender"]="Male", ["Pin Count"]="4", ["Pitch(mm)"]="2mm", ["Color"]="Natural"});
            Add("","CON","CON Male 2-Pin 2.5mm Natural STD","MFMP 7238 - 202 - 060 - 960 - 000 - 00","Stocko","MFMP 7238 - 202 - W60","CON Male 2Pin 2.5mm","MFMP 7238 - 202 - W60",STD,SMD,new(){["Gender"]="Male", ["Pin Count"]="2", ["Pitch(mm)"]="2.5mm", ["Color"]="Natural"});
            //  TERMINALS 
            Add("","TERM","TERM Male BM:CuZn PM:Sn AECQ","Special Design","","FCL_Terminal_1X4","FCL_Terminal_1X4","FCL_Terminal_1X4",AECQ,THT,new(){["Gender"]="Male", ["Base Material"]="CuZn", ["Plating Material"]="Sn"});
            //  PADs 
            Add("","PAD","PAD Size:(1.5x4mm,Rectangular) Hole:(1.2x1.2mm,Rectangular) THT","","","","","",STD,THT,new(){["PAD X-Size(mm)"]="1.5mm", ["PAD Y-Size(mm)"]="4mm", ["Hole X-Size(mm)"]="1.2mm", ["Hole Y-Size(mm)"]="1.2mm", ["PAD Shape"]="Rectangular", ["Hole Shape"]="Rectangular"});
            Add("","PAD","PAD Size:(1.5x5mm,Rectangular) SMD","","","FCL_Cable_Pad_1X4","FCL_Cable_Pad_1X4","FCL_Cable_Pad_1X4",STD,SMD,new(){["PAD X-Size(mm)"]="1.5mm", ["PAD Y-Size(mm)"]="5mm", ["Hole X-Size(mm)"]="NA", ["Hole Y-Size(mm)"]="NA", ["PAD Shape"]="Rectangular", ["Hole Shape"]="NA"});

            Save();
        }

        private static void Save()
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(DataFile)!);

                var customCats = CategoryFields.ToDictionary(
                    kvp => kvp.Key,
                    kvp => kvp.Value.Select(f => new CategoryField { Key = f.Key, Label = f.Label }).ToList()
                );

                var json = JsonSerializer.Serialize(new SaveData
                {
                    Components       = _components,
                    NextId           = _nextId,
                    CustomCategories = customCats,
                }, new JsonSerializerOptions { WriteIndented = true });

                File.WriteAllText(DataFile, json);
            }
            catch { }
        }

        public static string ExportCsv(string categoryFilter = "")
        {
            var recs = string.IsNullOrEmpty(categoryFilter)
                ? _components
                : _components.Where(c => c.Category == categoryFilter).ToList();

            var sb = new System.Text.StringBuilder();
            sb.AppendLine("ID,Category,Description,Part Number,Manufacturer,Footprint,Rating,Process,Library Ref,Footprint Ref,Status");
            foreach (var r in recs)
                sb.AppendLine($"{r.Id},{r.Category},{Esc(r.Description)},{Esc(r.PartNumber)},{Esc(r.Manufacturer)},{r.Footprint},{r.Rating},{r.Process},{r.LibraryRef},{r.FootprintRef},{r.Status}");
            return sb.ToString();
        }

        public static int ImportCsv(string csvContent)
        {
            int count = 0;
            var lines = csvContent.Split('\n', StringSplitOptions.RemoveEmptyEntries);
            if (lines.Length < 2) return 0;
            var headers = lines[0].Split(',').Select(h => h.Trim()).ToArray();
            for (int i = 1; i < lines.Length; i++)
            {
                var vals = lines[i].Split(',');
                if (vals.Length < 2) continue;
                var dict = new Dictionary<string, string>();
                for (int j = 0; j < headers.Length && j < vals.Length; j++)
                    dict[headers[j]] = vals[j].Trim();

                var c = new Component
                {
                    IsCustom     = true,
                    Status       = ComponentStatus.Pending,
                    Category     = dict.GetValueOrDefault("Category", ""),
                    Description  = dict.GetValueOrDefault("Description", ""),
                    PartNumber   = dict.GetValueOrDefault("Part Number", ""),
                    Manufacturer = dict.GetValueOrDefault("Manufacturer", ""),
                    Footprint    = dict.GetValueOrDefault("Footprint", ""),
                };
                if (Enum.TryParse<ComponentRating>(dict.GetValueOrDefault("Rating", "STD"), out var r)) c.Rating = r;
                if (Enum.TryParse<ComponentProcess>(dict.GetValueOrDefault("Process", "SMD"), out var p)) c.Process = p;
                Add(c);
                count++;
            }
            return count;
        }

        private static string Esc(string s) => s?.Contains(',') == true ? $"\"{s}\"" : s ?? "";

        private class SaveData
        {
            public List<Component>                              Components       { get; set; } = new();
            public int                                          NextId           { get; set; } = 1;
            public Dictionary<string, List<CategoryField>>?    CustomCategories { get; set; }
        }
    }
}